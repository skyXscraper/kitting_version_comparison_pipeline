"""Command line of the kitting version comparison pipeline. See README.md, or: python kvc.py --help"""
import argparse
import fnmatch
import itertools
import re
import sys
from datetime import datetime
from pathlib import Path

from . import report
from .config import CONFIG_PATH, REPORTS_DIR, TRUTH_PATH, load_config, write_default_config
from .loader import load_folder, load_registry, load_versions, natural, save_registry
from .observations import observe, observe_compare
from .scoring import evaluate, grade, head_to_head
from .truth import build_truth, is_labelled, load_truth, write_truth
from .xlsx_io import write_workbook


def _filter(videos, pattern):
    if not pattern:
        return set(videos)
    pats = [p.strip() for p in pattern.split(",") if p.strip()]
    keep = {v for v in videos if any(fnmatch.fnmatch(v.lower(), p.lower()) for p in pats)}
    if not keep:
        raise SystemExit(f"--videos {pattern}: matches none of the {len(videos)} videos (e.g. {sorted(videos)[:3]})")
    return keep


def _save(kind, name, text, sheets):
    out = REPORTS_DIR / f"{datetime.now():%Y-%m-%d_%H-%M-%S}_{kind}_{re.sub(r'[^A-Za-z0-9_.-]+', '_', name)[:80]}"
    out.mkdir(parents=True, exist_ok=True)
    (out / "report.txt").write_text(text + "\n", encoding="utf-8")
    write_workbook(out / "report.xlsx", sheets)
    return out


def _refresh_truth(cfg, quiet=False):
    versions = load_versions()
    existing, issues = load_truth()
    truth, added, harvested = build_truth(versions, existing, cfg)
    write_truth(truth, versions, cfg)
    rolls = sum(len(v) for v in truth.values())
    done = sum(1 for v in truth.values() if is_labelled(v))
    print(f"Truth sheet: {TRUTH_PATH}")
    print(f"  {rolls} rolls in {len(truth)} videos; {done} video(s) labelled, {len(truth) - done} still empty"
          + (f"; {added} new roll row(s) added" if added else "")
          + (f"; {harvested} true value(s) copied from the result workbooks" if harvested else ""))
    for i in issues:
        print(f"  ! {i}")
    if not quiet:
        print("  Fill the blue cells (True ply / True range), save, close Excel, then run score / compare.")


def cmd_add(a):
    folder = Path(a.folder).expanduser().resolve()
    data = load_folder(folder)
    n = len(data["models"])
    if a.names:
        names = [x.strip() for x in a.names.split(",")]
        if len(names) != n:
            raise SystemExit(f"--names: this folder holds {n} model(s) ({', '.join(data['tags'])}), give {n} name(s)")
    elif a.name and n == 1:
        names = [a.name]
    elif a.name:
        raise SystemExit(f"This is a side-by-side folder with {n} models ({', '.join(data['tags'])}): "
                         "use --names first,second instead of --name")
    else:
        names = [t if t.startswith(("rec", "det")) else f"rec_{t}" for t in data["tags"]]
    entries = load_registry()
    taken = {e["name"] for e in entries}
    clash = [x for x in names if x in taken]
    if clash and not a.replace:
        raise SystemExit(f"Version name(s) already registered: {', '.join(clash)}\n"
                         "Pick another with --name / --names, or pass --replace to point the name at this folder.")
    entries = [e for e in entries if e["name"] not in names]
    for b, name in enumerate(names):
        m = data["models"][b]
        entries.append({"name": name, "folder": str(folder), "block": b, "det": a.det or m["det"],
                        "rec": a.rec or m["rec"], "notes": a.notes or ""})
        rolls = sum(len(r.rolls) for r in m["videos"].values())
        print(f"Registered '{name}': {len(m['videos'])} videos, {rolls} rolls   "
              f"detector {a.det or m['det'] or '?'}, recognizer {a.rec or m['rec'] or '?'}")
    for w in data["warnings"]:
        print(f"  ! {w}")
    save_registry(entries)
    try:
        _refresh_truth(load_config(), quiet=True)
    except SystemExit as e:
        print(f"Version registered, but the truth sheet was not updated:\n{e}\nRun 'python kvc.py truth' afterwards.")


def cmd_remove(a):
    entries = load_registry()
    keep = [e for e in entries if e["name"] not in a.names]
    gone = [n for n in a.names if n not in {e["name"] for e in entries}]
    if gone:
        raise SystemExit(f"Not registered: {', '.join(gone)}")
    save_registry(keep)
    print(f"Removed {', '.join(a.names)} from the registry (result workbooks and truth are untouched).")


def cmd_list(a):
    cfg = load_config()
    truth, _ = load_truth()
    versions = load_versions()
    print(f"{'version':<18}{'videos':>7}{'rolls':>7}{'labelled':>10}{'overall':>9}{'OCR':>7}{'det':>7}   detector + recognizer / folder")
    for v in versions:
        m = evaluate(v, truth, cfg)["overall"]
        s = [m["scores"][k]["score"] for k in ("overall", "ocr", "det")]
        print(f"{v.name:<18}{m['videos']:>7}{m['rolls_reported']:>7}{m['ply_n']:>10}"
              + "".join(f"{'-' if x is None else f'{x:.1f}':>{w}}" for x, w in zip(s, (9, 7, 7)))
              + f"   {v.det or '?'} + {v.rec or '?'}")
        print(f"{'':<18}{v.folder}" + (f"   ({v.notes})" if v.notes else ""))
    print(f"\n'labelled' = rolls of this version's videos with a true ply in {TRUTH_PATH.name}. "
          "Scores here use each version's own videos;\nuse 'compare' for a like-for-like comparison on shared videos.")


def cmd_truth(a):
    _refresh_truth(load_config())


def score_version(ver, cfg, truth, issues, videos=None):
    """Score one version and save its report -> {res, obs, text, sheets, out}. Used by the command line and the web page."""
    res = evaluate(ver, truth, cfg, _filter(ver.videos, videos) if videos else None)
    obs = observe(res, cfg, issues)
    text = report.score_text(res, obs, cfg)
    sheets = report.score_sheets(res, obs, cfg)
    return {"res": res, "obs": obs, "text": text, "sheets": sheets, "out": _save("score", ver.name, text, sheets)}


def cmd_score(a):
    cfg = load_config()
    truth, issues = load_truth()
    for ver in load_versions(a.versions):
        run = score_version(ver, cfg, truth, issues, a.videos)
        print(run["text"])
        print(f"\n Saved: {run['out']}\\report.xlsx  (+ report.txt)\n")


def compare_versions(versions, cfg, truth, issues, videos=None, own=False, kind="compare"):
    """Compare versions and save the report -> {results, obs, per, scope, text, sheets, out}."""
    if len(versions) < 2:
        raise SystemExit("compare needs at least two versions (register more with 'add', or use 'score' for one).")
    sets = [_filter(v.videos, videos) if videos else set(v.videos) for v in versions]
    common = set.intersection(*sets)
    if not own and not common:
        raise SystemExit("These versions have no video in common, so they cannot be compared like for like.\n"
                         "Use --all-videos to score each on its own videos (less fair).")
    results = [evaluate(v, truth, cfg, s if own else common) for v, s in zip(versions, sets)]
    excluded = {} if own else {v.name: sorted(s - common, key=natural) for v, s in zip(versions, sets)}
    if own:
        scope = ("Each version scored on ITS OWN videos (" + ", ".join(f"{v.name} {len(s)}" for v, s in zip(versions, sets))
                 + "): not like for like where the video sets differ.")
    else:
        lab = sum(1 for v in common if is_labelled(truth.get(v, [])))
        scope = f"Like for like on the {len(common)} videos every version ran ({lab} of them labelled)."
    scored = [r for r in results if r["overall"]["scores"]["overall"]["score"] is not None]
    if len(results) <= 4:
        pairs = [(x, y, head_to_head(x, y)) for x, y in itertools.combinations(results, 2)]
    else:   # many versions: compare each against the best one only
        best = max(scored or results, key=lambda r: r["overall"]["scores"]["overall"]["score"] or 0)
        pairs = [(best, r, head_to_head(best, r)) for r in results if r is not best]
    obs = observe_compare(results, pairs, cfg, excluded, issues)
    per = {r["version"].name: observe(r, cfg, brief=True) for r in results}
    text = report.compare_text(results, obs, per, cfg, scope)
    sheets = report.compare_sheets(results, obs, per, pairs, cfg, scope)
    return {"results": results, "obs": obs, "per": per, "scope": scope, "text": text, "sheets": sheets,
            "out": _save(kind, "_vs_".join(v.name for v in versions), text, sheets)}


def cmd_compare(a, own_videos=False):
    cfg = load_config()
    truth, issues = load_truth()
    run = compare_versions(load_versions(a.versions or None), cfg, truth, issues, a.videos,
                           own=own_videos or a.all_videos, kind="rank" if own_videos else "compare")
    print(run["text"])
    print(f"\n Saved: {run['out']}\\report.xlsx  (+ report.txt)\n")


def cmd_rank(a):
    a.versions, a.all_videos = [], True
    cmd_compare(a, own_videos=True)


def cmd_config(a):
    write_default_config()
    print(f"Scoring settings: {CONFIG_PATH}\nEdit the weights / thresholds there; they apply to the next run.")


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    p = argparse.ArgumentParser(prog="kvc", description="Score and compare kitting OCR / detection versions from "
                                "their result workbooks.", epilog="Typical use:  add <folder>  ->  truth (fill the "
                                "sheet)  ->  score <version>  /  compare <v1> <v2> ...")
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("add", help="register a folder of result workbooks as a version (two versions for a side-by-side folder)")
    s.add_argument("folder")
    s.add_argument("--name", help="version name (single-model folder); default rec_<tag>")
    s.add_argument("--names", help="comma-separated names for a side-by-side folder, left model first")
    s.add_argument("--det", help="detector name, if the workbooks do not say")
    s.add_argument("--rec", help="recognizer name, if the workbooks do not say")
    s.add_argument("--notes", help="free text shown in reports (e.g. 'Hailo, Pi 5')")
    s.add_argument("--replace", action="store_true", help="overwrite a version name that is already registered")
    s.set_defaults(fn=cmd_add)
    s = sub.add_parser("remove", help="forget a registered version")
    s.add_argument("names", nargs="+")
    s.set_defaults(fn=cmd_remove)
    s = sub.add_parser("list", help="registered versions with their scores on their own videos")
    s.set_defaults(fn=cmd_list)
    s = sub.add_parser("truth", help="create / update truth/truth.xlsx with every roll of every registered version")
    s.set_defaults(fn=cmd_truth)
    s = sub.add_parser("score", help="score one version (or several, one report each)")
    s.add_argument("versions", nargs="+")
    s.add_argument("--videos", help="only these videos, wildcards allowed, comma-separated (e.g. oct_3_*,19_part2_*)")
    s.set_defaults(fn=cmd_score)
    s = sub.add_parser("compare", help="compare two or more versions on the videos they all ran (all versions if none named)")
    s.add_argument("versions", nargs="*")
    s.add_argument("--videos", help="only these videos, wildcards allowed, comma-separated")
    s.add_argument("--all-videos", action="store_true", help="score each version on its own videos instead of the shared ones")
    s.set_defaults(fn=cmd_compare)
    s = sub.add_parser("rank", help="leaderboard of every registered version, each on its own videos")
    s.add_argument("--videos", help="only these videos, wildcards allowed, comma-separated")
    s.set_defaults(fn=cmd_rank)
    s = sub.add_parser("config", help="write config.json so the scoring weights / thresholds can be edited")
    s.set_defaults(fn=cmd_config)
    a = p.parse_args(argv)
    a.fn(a)
