"""Web page for the pipeline: upload result workbooks, fill the truth and score / compare in the browser.

    python kvc_web.py [--port 8000] [--no-browser]

Runs on this computer only (127.0.0.1) and needs nothing but Python. It works on the same versions.json,
truth/truth.xlsx, config.json and reports/ as the command line, so the two can be mixed freely. Uploaded workbooks
are sorted by the model written in them and kept in uploads/<version>/; uploading more workbooks of a version
adds their videos to it.
"""
import argparse
import contextlib
import io
import json
import math
import re
import shutil
import threading
import traceback
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import parse_qs, unquote, urlsplit

from . import loader, report
from .cli import _refresh_truth, cmd_remove, compare_versions, score_version
from .config import REPORTS_DIR, ROOT, TRUTH_PATH, load_config
from .loader import entry_sources, load_folder, load_registry, load_version, load_versions, natural, save_registry
from .observations import pct
from .scoring import evaluate, grade
from .truth import TruthRoll, build_truth, is_labelled, load_truth, match_rolls, write_truth
from .xlsx_io import fnum

UPLOADS_DIR = ROOT / "uploads"
INCOMING = UPLOADS_DIR / "_incoming"
PAGE = Path(__file__).with_name("webapp.html")
MAX_UPLOAD = 200 * 1024 * 1024
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
LOCK = threading.Lock()
LEVELS = {"ALERT": 0, "WARN": 1, "INFO": 2}
_sig = {}


# ---------------------------------------------------------------- helpers
def _folders(entries):
    return {f for e in entries for f, _ in entry_sources(e)}


def _fresh():
    """Drop the cached workbooks of a folder whose files changed since they were loaded."""
    for f in _folders(load_registry()):
        p = Path(f)
        sig = tuple(sorted((x.name, x.stat().st_mtime_ns, x.stat().st_size) for x in p.glob("*.xlsx"))) if p.is_dir() else None
        if _sig.get(f) != sig:
            _sig[f] = sig
            loader._folder_cache.pop(f, None)


def _quiet(fn, *args):
    """Run a command-line function and return what it printed."""
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        fn(*args)
    return buf.getvalue().strip()


def _versions_ok():
    """Every registered version that can still be loaded."""
    out = []
    for e in load_registry():
        try:
            out.append(load_version(e))
        except SystemExit:
            pass
    return out


def _truth():
    try:
        return load_truth()
    except SystemExit as e:
        return {}, [str(e)]
    except Exception as e:
        return {}, [f"truth.xlsx could not be read ({e})"]


def _prune(before):
    """Delete upload folders that no registered version uses any more."""
    for f in before - _folders(load_registry()):
        loader._folder_cache.pop(f, None)
        if Path(f).parent == UPLOADS_DIR:
            shutil.rmtree(f, ignore_errors=True)


def _cell(v):
    if isinstance(v, tuple):
        return {"v": _cell(v[0]), "s": v[1]}
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    return v


def _sheet(sh):
    return {"name": sh["name"], "header_rows": sh.get("header_rows", 0), "rows": [[_cell(c) for c in r] for r in sh["rows"]]}


def _obs(obs):
    return [{"level": lv, "text": t} for lv, t in sorted(obs, key=lambda x: LEVELS[x[0]])]


def _card(res, cfg):
    m, ver = res["overall"], res["version"]
    return {"name": ver.name, "label": ver.label, "scores": [
        {"title": title, "what": what, "score": m["scores"][key]["score"], "grade": grade(m["scores"][key]["score"], cfg)}
        for key, title, what in report.SCORES]}


# ---------------------------------------------------------------- api
def api_state(_=None):
    cfg = load_config()
    truth, issues = _truth()
    versions = []
    for e in load_registry():
        folders = [f for f, _ in entry_sources(e)]
        row = {"name": e["name"], "folder": "\n".join(folders),
               "uploaded": sum(Path(f).parent == UPLOADS_DIR for f in folders), "folders": len(folders)}
        try:
            v = load_version(e)
            m = evaluate(v, truth, cfg)["overall"]
            row.update(det=v.det, rec=v.rec, notes=v.notes, videos=m["videos"], rolls=m["rolls_reported"],
                       labelled=m["ply_n"], warnings=v.warnings, ran=sorted(v.videos, key=natural),
                       scores={k: m["scores"][k]["score"] for k in ("overall", "ocr", "det")},
                       grade=grade(m["scores"]["overall"]["score"], cfg))
        except SystemExit as x:
            row["error"] = str(x)
        versions.append(row)
    names = sorted({x for v in versions for x in v.get("ran", [])}, key=natural)
    videos = [{"name": x, "rolls": len(truth.get(x, [])), "labelled": sum(1 for t in truth.get(x, []) if t.ply),
               "scored": is_labelled(truth.get(x, []))} for x in names]
    return {"versions": versions, "videos": videos, "issues": issues, "root": str(ROOT),
            "truth": {"videos": len(truth), "rolls": sum(len(v) for v in truth.values()),
                      "labelled": sum(1 for v in truth.values() if is_labelled(v))}}


def api_upload(query, body):
    batch, i = query.get("batch", [""])[0], query.get("i", [""])[0]
    name = Path(query.get("name", [""])[0].replace("\\", "/")).name
    if not re.fullmatch(r"[A-Za-z0-9_-]{1,40}", batch) or not re.fullmatch(r"\d{1,5}", i):
        raise SystemExit("Bad upload id.")
    if not name.lower().endswith(".xlsx") or name.startswith("~$"):
        raise SystemExit(f"{name}: not an .xlsx workbook.")
    if not body.startswith(b"PK"):
        raise SystemExit(f"{name}: not a valid .xlsx file.")
    folder = INCOMING / batch / i      # one folder per file: two models' workbooks of a video have the same name
    folder.mkdir(parents=True, exist_ok=True)
    (folder / name).write_bytes(body)
    return {"saved": name}


def _default_name(tag):
    return tag if tag.startswith(("rec", "det")) else f"rec_{tag}"


def api_add(body):
    """Sort the uploaded workbooks by the model(s) written in them and add each group to its version."""
    batch = str(body.get("batch", ""))
    inc = INCOMING / batch
    if not re.fullmatch(r"[A-Za-z0-9_-]{1,40}", batch) or not inc.is_dir():
        raise SystemExit("No workbooks were uploaded.")
    try:
        groups, loose, skipped = {}, [], []
        for d in sorted(inc.iterdir(), key=lambda d: int(d.name)):
            f = next(d.glob("*.xlsx"))
            try:
                data = load_folder(d)
            except SystemExit:
                skipped.append(f"{f.name}: not a result workbook (no Frames / Rolls or Summary sheet)")
                continue
            tags = tuple(data["tags"]) if len(data["tags"]) == len(data["models"]) else ()
            if tags:
                groups.setdefault(tags, []).append((f, data))
            else:
                loose.append((f, data))
        for f, data in loose:   # a workbook that does not name its model(s): only clear when there is one group
            same = [k for k in groups if len(k) == len(data["models"])]
            if len(same) == 1:
                groups[same[0]].append((f, data))
            else:
                skipped.append(f"{f.name}: cannot tell which model it belongs to; upload it together with that "
                               "model's workbooks only")
        if not groups:
            raise SystemExit("None of the uploaded files is a result workbook.\n" + "\n".join(skipped))
        given = [x.strip() for x in str(body.get("names") or "").split(",") if x.strip()]
        if given:
            if len(groups) > 1:
                raise SystemExit("These workbooks come from different models ("
                                 + "; ".join(" + ".join(k) for k in groups) + "), so they are sorted into versions by "
                                 "the model written in them. Leave the name empty, or upload one model at a time to name it.")
            tags = next(iter(groups))
            if len(given) != len(tags):
                raise SystemExit(f"These workbooks hold {len(tags)} model(s) side by side ({', '.join(tags)}): "
                                 f"give {len(tags)} name(s) separated by a comma, or leave the name empty.")
            if len(set(given)) != len(given) or not all(re.fullmatch(r"[A-Za-z0-9_.-]{1,40}", x) for x in given):
                raise SystemExit("Version names must be different and use only letters, digits, _ . -")
        notes = str(body.get("notes") or "").strip()
        entries = load_registry()
        by_name = {e["name"]: e for e in entries}
        added = []
        for tags, items in groups.items():
            names = given or [re.sub(r"[^A-Za-z0-9_.-]+", "_", _default_name(t))[:40] for t in tags]
            folder = UPLOADS_DIR / "+".join(names)
            folder.mkdir(parents=True, exist_ok=True)
            videos = []
            for f, data in items:
                f.replace(folder / f.name)
                if f.stem in data["models"][0]["videos"]:
                    videos.append(f.stem)
            loader._folder_cache.pop(str(folder), None)
            data = load_folder(folder)
            if len(data["models"]) != len(names):
                raise SystemExit(f"{folder.name}: the workbooks kept for this version do not all have the same layout "
                                 f"({len(data['models'])} model(s) side by side, expected {len(names)}).")
            for b, name in enumerate(names):
                e, m = by_name.get(name), data["models"][b]
                new = e is None
                if new:
                    e = by_name[name] = {"name": name, "folder": str(folder), "block": b, "det": m["det"],
                                         "rec": m["rec"], "notes": notes}
                    entries.append(e)
                else:
                    if (str(folder), b) not in entry_sources(e):
                        e.setdefault("more", []).append({"folder": str(folder), "block": b})
                    e["notes"] = notes or e.get("notes", "")
                added.append({"name": name, "new": new, "videos": sorted(set(videos), key=natural),
                              "files": len(items), "model": " + ".join(x for x in (m["det"], m["rec"]) if x)})
        save_registry(entries)
    finally:
        shutil.rmtree(inc, ignore_errors=True)
    try:
        log = _quiet(_refresh_truth, load_config(), True)
    except SystemExit as e:
        log = f"The truth sheet was not updated: {e}"
    return {"added": added, "skipped": skipped, "log": log}


def api_remove(body):
    before = _folders(load_registry())
    log = _quiet(cmd_remove, SimpleNamespace(names=[str(body.get("name", ""))]))
    _prune(before)
    return {"log": log}


def api_truth_get(_=None):
    cfg = load_config()
    versions = _versions_ok()
    existing, issues = _truth()
    truth, _, _ = build_truth(versions, existing, cfg)   # also lists rolls not written to the sheet yet
    tol = cfg["roll_match_tolerance_s"]
    videos = []
    for video in sorted(truth, key=natural):
        trs = truth[video]
        ref = [{id(t): r for t, r in match_rolls(trs, v.videos[video].rolls, tol) if t is not None}
               if video in v.videos else None for v in versions]
        rows = []
        for t in trs:
            got = [None if m is None else [] if m.get(id(t)) is None else
                   [m[id(t)].ply, m[id(t)].rng, m[id(t)].status] for m in ref]
            rows.append({"n": t.n, "t": t.t, "real": "" if t.real is None else "yes" if t.real else "no",
                         "ply": t.ply, "rng": t.rng, "notes": t.notes, "ref": got})
        videos.append({"video": video, "rows": rows})
    return {"versions": [v.name for v in versions], "videos": videos, "issues": issues}


def api_truth_save(body):
    truth = {}
    for row in body.get("rows", []):
        video = str(row.get("video") or "").strip()
        if not video:
            continue
        rows, n, real = truth.setdefault(video, []), fnum(row.get("n")), row.get("real")
        rows.append(TruthRoll(video=video, n=int(n) if n is not None else len(rows) + 1, t=fnum(row.get("t")),
                              real=True if real == "yes" else False if real == "no" else None,
                              ply=str(row.get("ply") or "").strip(), rng=str(row.get("rng") or "").replace(" ", ""),
                              notes=str(row.get("notes") or "").strip()))
    cfg, versions = load_config(), _versions_ok()
    truth, _, _ = build_truth(versions, truth, cfg)
    write_truth(truth, versions, cfg)
    return {"rolls": sum(len(v) for v in truth.values()), "labelled": sum(1 for v in truth.values() if is_labelled(v))}


def api_truth_upload(body):
    if not body.startswith(b"PK"):
        raise SystemExit("That is not an .xlsx file.")
    TRUTH_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = TRUTH_PATH.with_name("_upload.xlsx")
    tmp.write_bytes(body)
    try:
        truth, _ = load_truth(tmp)     # refuses a workbook that is not a truth sheet
        if not truth:
            raise SystemExit("That workbook has no truth rows (a 'Video' and 'True ply' header with rows under it).")
        if TRUTH_PATH.exists():
            bdir = TRUTH_PATH.parent / "backups"
            bdir.mkdir(exist_ok=True)
            shutil.copy2(TRUTH_PATH, bdir / f"truth_{datetime.now():%Y%m%d_%H%M%S}_before_upload.xlsx")
        tmp.replace(TRUTH_PATH)
    finally:
        tmp.unlink(missing_ok=True)
    cfg, versions = load_config(), _versions_ok()
    truth, issues = load_truth()
    truth, _, _ = build_truth(versions, truth, cfg)
    write_truth(truth, versions, cfg)      # adds the rolls and reference columns of the registered versions
    return {"rolls": sum(len(v) for v in truth.values()), "issues": issues}


def api_run(body):
    names = [str(x) for x in body.get("versions") or []]
    if not names:
        raise SystemExit("Pick at least one version.")
    cfg = load_config()
    truth, issues = load_truth()
    versions = load_versions(names)
    videos = body.get("videos")      # a list of exact video names, a 'a*,b*' pattern, or nothing = all
    videos = ([str(x) for x in videos] if isinstance(videos, list) else str(videos or "").strip()) or None
    if len(versions) == 1:
        run = score_version(versions[0], cfg, truth, issues, videos)
        res = run["res"]
        m, ver = res["overall"], res["version"]
        lines = [f"Scored on {m['ply_n']} labelled roll(s) in {m['videos_scored']} video(s); this version ran "
                 f"{m['videos']} videos, {m['rolls_reported']} rolls."]
        if m["ply_ci"]:
            lines.append(f"Ply accuracy {m['ply_right']}/{m['ply_n']} = {pct(m['ply_right'] / m['ply_n'])}, "
                         f"95% interval {pct(m['ply_ci'][0])} to {pct(m['ply_ci'][1])}.")
        breakdown = []
        for key, title, _ in report.SCORES:
            sc = m["scores"][key]
            if sc["parts"]:
                breakdown.append({"title": title, "score": sc["score"],
                                  "left_out": [report.LABELS[k] for k in sc.get("left_out", [])],
                                  "parts": [{"label": report.LABELS[k], "result": pct(val, 1), "weight": round(w, 1),
                                             "points": round(pts, 1), "detail": report.detail(k, m)}
                                            for k, val, w, pts in sc["parts"]]})
        sessions = [["Session", "Videos", "Rolls", "Ply right", "Range exact", "Overall"]]
        for g in sorted((g for g, gm in res["by_group"].items() if gm["ply_n"]), key=natural):
            gm = res["by_group"][g]
            s = gm["scores"]["overall"]["score"]
            sessions.append([g, gm["videos_scored"], gm["truth_rolls"], f"{gm['ply_right']}/{gm['ply_n']}",
                             f"{gm['range_exact']}/{gm['range_n']}" if gm["range_n"] else "-",
                             None if s is None else round(s, 1)])
        extra = [{"name": "Sessions", "header_rows": 1, "rows": sessions}] if len(sessions) > 2 else []
        out = {"kind": "score", "title": ver.label, "lines": lines, "cards": [_card(res, cfg)], "breakdown": breakdown,
               "groups": [], "extra": extra}
    else:
        run = compare_versions(versions, cfg, truth, issues, videos, own=bool(body.get("own")))
        results = run["results"]
        matrix = []
        for row in report._matrix(results, cfg):
            if row is None:
                matrix.append(None)
                continue
            label, vals, better, fmt = row
            best = report._best(vals, better)
            matrix.append({"label": label, "cells": [{"text": "-" if v is None else fmt(v), "best": i in best}
                                                     for i, v in enumerate(vals)]})
        groups = [{"title": f"{name}: its own alerts", "items": _obs([x for x in vo if x[0] in ("ALERT", "WARN")])}
                  for name, vo in run["per"].items()]
        out = {"kind": "compare", "title": "  vs  ".join(r["version"].name for r in results), "lines": [run["scope"]],
               "cards": [_card(r, cfg) for r in results], "matrix": matrix,
               "groups": [g for g in groups if g["items"]], "extra": []}
    out.update(observations=_obs(run["obs"]), text=run["text"], report=run["out"].name,
               sheets=out.pop("extra") + [_sheet(s) for s in run["sheets"] if s["name"] not in ("Summary", "Observations")])
    return out


GET = {"/api/state": api_state, "/api/truth": api_truth_get}
POST = {"/api/add": api_add, "/api/remove": api_remove, "/api/truth": api_truth_save, "/api/run": api_run}


# ---------------------------------------------------------------- http
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):   # keep the console for errors only
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8", download=None):
        if not isinstance(body, bytes):
            body = json.dumps(body).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if download:
            self.send_header("Content-Disposition", f'attachment; filename="{download}"')
        self.end_headers()
        self.wfile.write(body)

    def _local(self):
        """Only this computer's own page may talk to the server (no other site, no other host name)."""
        host = (self.headers.get("Host") or "").rsplit(":", 1)[0]
        return host in ("127.0.0.1", "localhost")

    def _api(self, fn, *args):
        with LOCK:
            try:
                _fresh()
                self._send(200, fn(*args))
            except SystemExit as e:
                self._send(400, {"error": str(e)})
            except PermissionError as e:
                self._send(400, {"error": f"Cannot write {e.filename or 'a file'}: it is open in another program "
                                          "(Excel?). Close it and try again."})
            except Exception as e:
                traceback.print_exc()
                self._send(500, {"error": f"{type(e).__name__}: {e}"})

    def do_GET(self):
        if not self._local():
            return self._send(403, {"error": "forbidden"})
        path = unquote(urlsplit(self.path).path)
        if path == "/":
            self._send(200, PAGE.read_bytes(), "text/html; charset=utf-8")
        elif path in GET:
            self._api(GET[path])
        elif path == "/truth.xlsx" and TRUTH_PATH.exists():
            self._send(200, TRUTH_PATH.read_bytes(), XLSX, "truth.xlsx")
        elif path.startswith("/reports/"):
            f = (REPORTS_DIR / path[len("/reports/"):]).resolve()
            if REPORTS_DIR.resolve() not in f.parents or not f.is_file() or f.suffix not in (".xlsx", ".txt"):
                return self._send(404, {"error": "not found"})
            self._send(200, f.read_bytes(), XLSX if f.suffix == ".xlsx" else "text/plain; charset=utf-8",
                       f"{f.parent.name}{f.suffix}" if f.suffix == ".xlsx" else None)
        else:
            self._send(404, {"error": "not found"})

    def do_POST(self):
        if not self._local() or self.headers.get("X-KVC") != "1":
            return self._send(403, {"error": "forbidden"})
        u = urlsplit(self.path)
        size = int(self.headers.get("Content-Length") or 0)
        if size > MAX_UPLOAD:
            return self._send(413, {"error": "file too large"})
        body = self.rfile.read(size)
        if u.path == "/api/upload":
            self._api(api_upload, parse_qs(u.query), body)
        elif u.path == "/api/truth/upload":
            self._api(api_truth_upload, body)
        elif u.path in POST:
            try:
                data = json.loads(body or b"{}")
            except ValueError:
                return self._send(400, {"error": "bad request"})
            self._api(POST[u.path], data)
        else:
            self._send(404, {"error": "not found"})


def main(argv=None):
    p = argparse.ArgumentParser(prog="kvc_web", description="Web page of the kitting version comparison pipeline "
                                "(this computer only).")
    p.add_argument("--port", type=int, default=8000)
    p.add_argument("--no-browser", action="store_true", help="do not open the page in the browser")
    a = p.parse_args(argv)
    try:
        srv = ThreadingHTTPServer(("127.0.0.1", a.port), Handler)
    except OSError as e:
        raise SystemExit(f"Cannot use port {a.port} ({e}). Try another one:  python kvc_web.py --port {a.port + 1}")
    url = f"http://127.0.0.1:{a.port}/"
    print(f"Kitting version comparison is running at {url}\nKeep this window open; press Ctrl+C to stop.")
    if not a.no_browser:
        webbrowser.open(url)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
