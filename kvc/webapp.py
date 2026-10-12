"""Web page for the pipeline: drop the result sheets of different model versions for the same video, get the comparison.

    python kvc_web.py [--port 8000] [--no-browser]

Runs on this computer only (127.0.0.1) and needs nothing but Python. The uploaded sheets are read and thrown away:
nothing is registered and no report is written. The only thing kept is the truth you type (truth/truth.xlsx), so
the true ply / range of a video is entered once. Scoring settings come from config.json.

Next to each roll the page shows snapshots of its labels, cut from the run's video (see snapshots.py); they only
live in the page and are gone when it is refreshed or other sheets are compared.
"""
import argparse
import base64
import io
import json
import math
import re
import tempfile
import threading
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from . import report
from .cli import compare_versions, score_version
from .config import load_config
from .loader import Version, load_folder, natural
from .observations import pct
from .scoring import grade
from .snapshots import SnapshotError, run_snapshots
from .truth import TruthRoll, build_truth, is_labelled, load_truth, match_rolls, write_truth
from .xlsx_io import fnum, write_workbook

PAGE = Path(__file__).with_name("webapp.html")
MAX_BODY = 300 * 1024 * 1024
LOCK = threading.Lock()
LEVELS = {"ALERT": 0, "WARN": 1, "INFO": 2}


# ---------------------------------------------------------------- reading the uploaded sheets
def _version_name(tag):
    name = tag if tag.startswith(("rec", "det")) else f"rec_{tag}"
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", name)[:40]


def read_sheets(files, tmp):
    """Uploaded workbooks -> ([Version], [files left out and why]). A version is named after the model written in the
    sheet; a second sheet of the same model and video becomes <name>_2 so two runs can be compared too."""
    versions, used, summaries, skipped = [], {}, [], []
    for i, f in enumerate(files):
        name = Path(str(f.get("name") or "").replace("\\", "/")).name
        try:
            raw = base64.b64decode(f.get("data") or "")
        except ValueError:
            raw = b""
        if not name.lower().endswith(".xlsx") or not raw.startswith(b"PK"):
            skipped.append(f"{name}: not an .xlsx workbook")
            continue
        d = tmp / str(i)
        d.mkdir()
        (d / name).write_bytes(raw)
        try:
            data = load_folder(d)
        except SystemExit:
            skipped.append(f"{name}: not a result sheet (no Frames / Rolls or Summary sheet)")
            continue
        n, video = len(data["models"]), Path(name).stem
        tags = data["tags"] if len(data["tags"]) == n else [f"model{b + 1}" for b in range(n)]
        if video not in data["models"][0]["videos"]:
            summaries.append((name, tags, data))
            continue
        for b, tag in enumerate(tags):
            m, base, k = data["models"][b], _version_name(tag), 1
            while True:
                vname = base if k == 1 else f"{base}_{k}"
                ver = next((v for v in versions if v.name == vname), None)
                if ver is None:
                    ver = Version(name=vname, folder="", kind=data["kind"], block=b, tag=tag, det=m["det"], rec=m["rec"])
                    versions.append(ver)
                if video not in ver.videos:
                    break
                k += 1
            ver.videos[video] = m["videos"][video]
            ver.warnings += [w for w in data["warnings"] if w not in ver.warnings]
            used.setdefault(vname, []).append(name)
    for name, tags, data in summaries:      # a summary workbook only supplies the frame counts and speeds
        hit = False
        for ver in versions:
            if ver.tag in tags and tags.index(ver.tag) == ver.block:
                for video, run in ver.videos.items():
                    st = data["models"][ver.block]["videos"].get(video)
                    if st is not None:
                        hit = True
                        for k, v in st.stats.items():
                            run.stats.setdefault(k, v)
        if not hit:
            skipped.append(f"{name}: summary workbook, but it has no numbers for the uploaded video sheets")
    for ver in versions:
        ver.folder = "uploaded: " + ", ".join(used[ver.name])
    return versions, skipped


# ---------------------------------------------------------------- truth
def _saved_truth():
    try:
        return load_truth()
    except SystemExit as e:
        return {}, [str(e)]
    except Exception as e:
        return {}, [f"truth.xlsx could not be read ({e})"]


def _edited(rows, videos):
    """Truth rows typed on the page -> {video: [TruthRoll]}, only for compared videos the page sent rows for."""
    out = {}
    for row in rows:
        video = str(row.get("video") or "").strip()
        if video not in videos:
            continue
        n, real = fnum(row.get("n")), row.get("real")
        out.setdefault(video, []).append(TruthRoll(
            video=video, n=int(n) if n is not None else len(out[video]) + 1, t=fnum(row.get("t")),
            real=True if real == "yes" else False if real == "no" else None, ply=str(row.get("ply") or "").strip(),
            rng=str(row.get("rng") or "").replace(" ", ""), notes=str(row.get("notes") or "").strip()))
    return out


def _truth_table(truth, versions, videos, cfg):
    tol, out = cfg["roll_match_tolerance_s"], []
    for video in sorted(videos, key=natural):
        trs = truth.get(video, [])
        ref = [{id(t): r for t, r in match_rolls(trs, v.videos[video].rolls, tol) if t is not None}
               if video in v.videos else None for v in versions]
        out.append({"video": video, "rows": [
            {"n": t.n, "t": t.t, "real": "" if t.real is None else "yes" if t.real else "no", "ply": t.ply,
             "rng": t.rng, "notes": t.notes,
             "ref": [None if m is None else [] if m.get(id(t)) is None else
                     [m[id(t)].ply, m[id(t)].rng, m[id(t)].status, m[id(t)].track] for m in ref]} for t in trs]})
    return out


# ---------------------------------------------------------------- result for the page
def _cell(v):
    if isinstance(v, tuple):
        return {"v": _cell(v[0]), "s": v[1]}
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    return v


def _obs(obs):
    return [{"level": lv, "text": t} for lv, t in sorted(obs, key=lambda x: LEVELS[x[0]])]


def _card(res, cfg):
    m, ver = res["overall"], res["version"]
    return {"name": ver.name, "label": ver.label, "scores": [
        {"title": title, "what": what, "score": m["scores"][key]["score"], "grade": grade(m["scores"][key]["score"], cfg)}
        for key, title, what in report.SCORES]}


def _score_result(run, cfg):
    res = run["res"]
    m, ver = res["overall"], res["version"]
    lines = [f"Scored on {m['ply_n']} labelled roll(s) in {m['videos_scored']} video(s); the sheets hold "
             f"{m['videos']} video(s), {m['rolls_reported']} rolls."]
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
    return {"kind": "score", "title": ver.label, "lines": lines, "cards": [_card(res, cfg)], "breakdown": breakdown,
            "groups": []}


def _compare_result(run, cfg):
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
    return {"kind": "compare", "title": "  vs  ".join(r["version"].name for r in results), "lines": [run["scope"]],
            "cards": [_card(r, cfg) for r in results], "matrix": matrix, "groups": [g for g in groups if g["items"]]}


def _uploaded(body):
    """The sheets of a request -> (versions, files left out, videos every version has, settings)."""
    files = body.get("files") or []
    if not files:
        raise SystemExit("Choose the sheets to compare first.")
    cfg = load_config()
    with tempfile.TemporaryDirectory() as tmp:      # the sheets only live here while they are read
        versions, skipped = read_sheets(files, Path(tmp))
    if not versions:
        raise SystemExit("None of the files is a result sheet of a video.\n" + "\n".join(skipped))
    videos = set.intersection(*(set(v.videos) for v in versions))
    if not videos:
        raise SystemExit("These sheets are not of the same video, so they cannot be compared:\n"
                         + "\n".join(f"  {v.name}: {', '.join(sorted(v.videos, key=natural))}" for v in versions)
                         + "\nUpload the sheet of the same video for each model version.")
    return versions, skipped, videos, cfg


def api_snapshots(body):
    """Pictures of every roll of the compared videos, keyed 'video|version index|track' like the truth table's
    reference columns. A roll is shown as the first version that found it saw it."""
    versions, _, videos, cfg = _uploaded(body)
    truth, _, _ = build_truth(versions, _saved_truth()[0], cfg)
    want = {}       # (version index, video) -> roll tracks
    for video in videos:
        trs = truth.get(video, [])
        maps = [{id(t): r for t, r in match_rolls(trs, v.videos[video].rolls, cfg["roll_match_tolerance_s"])
                 if t is not None and r is not None} for v in versions]
        for t in trs:
            vi = next((i for i, m in enumerate(maps) if id(t) in m), None)
            if vi is not None:
                want.setdefault((vi, video), set()).add(maps[vi][id(t)].track)
    snaps, problems = {}, []
    for (vi, video), tracks in sorted(want.items()):
        log = versions[vi].videos[video].stats.get("log")
        try:
            if not log:
                raise SnapshotError("its sheet does not name the run's log folder (Run info sheet)")
            for track, shot in run_snapshots(cfg["runs_root"], log, tracks).items():
                snaps[f"{video}|{vi}|{track}"] = shot
        except SnapshotError as e:
            problems.append(f"{video} ({versions[vi].name}): {e}")
    return {"snaps": snaps, "problems": problems}


def api_compare(body):
    versions, skipped, videos, cfg = _uploaded(body)
    truth, issues = _saved_truth()
    rows = body.get("truth")
    if rows is not None:        # truth typed on the page: keep it for next time, other videos stay as they are
        truth.update(_edited(rows, videos))
        truth, _, _ = build_truth(versions, truth, cfg)
        write_truth(truth, versions, cfg)
        truth, issues = load_truth()
    else:
        truth, _, _ = build_truth(versions, truth, cfg)
    issues = [i for i in issues if any(v in i for v in videos)]
    if len(versions) == 1:
        run = score_version(versions[0], cfg, truth, issues, save=False)
        out = _score_result(run, cfg)
    else:
        run = compare_versions(versions, cfg, truth, issues, save=False)
        out = _compare_result(run, cfg)
    book = io.BytesIO()
    write_workbook(book, run["sheets"])
    out.update(observations=_obs(run["obs"]), text=run["text"],
               xlsx=base64.b64encode(book.getvalue()).decode("ascii"),
               xlsx_name=re.sub(r"[^A-Za-z0-9_.-]+", "_", f"{out['kind']}_{'_vs_'.join(v.name for v in versions)}")[:120] + ".xlsx",
               sheets=[{"name": s["name"], "header_rows": s.get("header_rows", 0),
                        "rows": [[_cell(c) for c in r] for r in s["rows"]]}
                       for s in run["sheets"] if s["name"] not in ("Summary", "Observations")])
    return {"versions": [{"name": v.name, "label": v.label, "files": v.folder[len("uploaded: "):],
                          "videos": sorted(v.videos, key=natural)} for v in versions],
            "skipped": skipped, "videos": sorted(videos, key=natural), "saved": rows is not None,
            "labelled": sum(1 for v in videos if is_labelled(truth.get(v, []))),
            "truth": _truth_table(truth, versions, videos, cfg), "result": out}


# ---------------------------------------------------------------- http
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):   # keep the console for errors only
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        if not isinstance(body, bytes):
            body = json.dumps(body).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _local(self):
        """Only this computer's own page may talk to the server (no other site, no other host name)."""
        host = (self.headers.get("Host") or "").rsplit(":", 1)[0]
        return host in ("127.0.0.1", "localhost")

    def do_GET(self):
        if not self._local():
            return self._send(403, {"error": "forbidden"})
        if self.path.split("?")[0] == "/":
            self._send(200, PAGE.read_bytes(), "text/html; charset=utf-8")
        else:
            self._send(404, {"error": "not found"})

    def do_POST(self):
        if not self._local() or self.headers.get("X-KVC") != "1":
            return self._send(403, {"error": "forbidden"})
        size = int(self.headers.get("Content-Length") or 0)
        api = {"/api/compare": api_compare, "/api/snapshots": api_snapshots}.get(self.path)
        if api is None or size > MAX_BODY:
            return self._send(404 if size <= MAX_BODY else 413, {"error": "not found" if size <= MAX_BODY else "too large"})
        try:
            body = json.loads(self.rfile.read(size) or b"{}")
        except ValueError:
            return self._send(400, {"error": "bad request"})
        with LOCK:
            try:
                self._send(200, api(body))
            except SystemExit as e:
                self._send(400, {"error": str(e)})
            except PermissionError as e:
                self._send(400, {"error": f"Cannot write {e.filename or 'a file'}: it is open in another program "
                                          "(Excel?). Close it and try again."})
            except Exception as e:
                traceback.print_exc()
                self._send(500, {"error": f"{type(e).__name__}: {e}"})


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
