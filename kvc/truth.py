"""The truth sheet: one row per physical roll of every video, filled in by hand once and used for every version.

truth/truth.xlsx, sheet Truth:
  Video | Roll # | First seen (s) | Real roll? | True ply | True range | Notes | <what each version decided, for reference>
A roll is identified by its video and the time it first appears, so versions with different detectors (different
track ids / seq numbers) still line up with the same truth row.
"""
import shutil
from dataclasses import dataclass
from datetime import datetime

from .config import TRUTH_PATH
from .loader import natural
from .xlsx_io import S, cell, fnum, read_workbook, txt, write_workbook

COLS = ["Video", "Roll #", "First seen (s)", "Real roll?", "True ply", "True range", "Notes"]
EXTRA_ROWS = 200


@dataclass
class TruthRoll:
    video: str
    n: int = 0
    t: float = None
    real: bool = None       # None = not said (treated as a real roll when a true value is given)
    ply: str = ""
    rng: str = ""
    notes: str = ""

    @property
    def filled(self):
        return bool(self.ply or self.rng or self.real is not None)


def is_labelled(rows):
    """A video is scored once at least one of its rolls has a true value (or is marked 'not a roll')."""
    return any(r.ply or r.rng or r.real is False for r in rows)


def load_truth(path=TRUTH_PATH):
    """-> ({video: [TruthRoll]}, [problems found in the sheet])."""
    if not path.exists():
        return {}, []
    wb = read_workbook(path)
    rows = wb.get("Truth") or next(iter(wb.values()))
    h = next((i for i, r in enumerate(rows) if {"Video", "True ply"} <= {txt(v) for v in r}), None)
    if h is None:
        raise SystemExit(f"{path}: no header row with 'Video' and 'True ply' found")
    hdr = [txt(v) for v in rows[h]]
    ix = {c: hdr.index(c) if c in hdr else None for c in COLS}
    out, issues = {}, []
    for rn, row in enumerate(rows[h + 1:], h + 2):
        video = txt(cell(row, ix["Video"]))
        if not video:
            continue
        raw_range = cell(row, ix["True range"])
        if isinstance(raw_range, (int, float)):
            issues.append(f"truth.xlsx row {rn} ({video}): True range is the number {raw_range}, not text like "
                          "'50.4-68.4' - Excel may have turned it into a date/number; retype it with a leading '")
        real = txt(cell(row, ix["Real roll?"])).lower()
        ply = txt(cell(row, ix["True ply"]))
        n = fnum(cell(row, ix["Roll #"]))
        out.setdefault(video, []).append(TruthRoll(
            video=video, n=int(n) if n is not None else len(out.get(video, [])) + 1,
            t=fnum(cell(row, ix["First seen (s)"])),
            real=True if real.startswith("y") else False if real.startswith("n") else None,
            ply=ply, rng=txt(raw_range).replace(" ", ""), notes=txt(cell(row, ix["Notes"]))))
    for video, trs in out.items():
        if is_labelled(trs):
            blank = [t.n for t in trs if not t.filled]
            if blank:
                issues.append(f"truth.xlsx {video}: roll(s) {', '.join(map(str, blank))} left blank while other rolls "
                              "of this video are filled - they are skipped (fill them, or mark Real roll? = no)")
    return out, issues


def match_rolls(trows, rolls, tol):
    """Pair truth rows with a version's rolls by first-seen time (closest first, one to one, within tol seconds).
    -> [(truth row or None, roll or None)]"""
    cand = sorted((abs(t.t - r.first_seen), i, j) for i, t in enumerate(trows) for j, r in enumerate(rolls)
                  if t.t is not None and r.first_seen is not None and abs(t.t - r.first_seen) <= tol)
    ti, rj = {}, {}
    for _, i, j in cand:
        if i not in ti and j not in rj:
            ti[i], rj[j] = j, i
    # rows or rolls without a time: fall back to roll number == seq
    for i, t in enumerate(trows):
        if i in ti or t.t is not None:
            continue
        for j, r in enumerate(rolls):
            if j not in rj and r.seq == t.n:
                ti[i], rj[j] = j, i
                break
    for j, r in enumerate(rolls):
        if j in rj or r.first_seen is not None:
            continue
        for i, t in enumerate(trows):
            if i not in ti and t.n == r.seq:
                ti[i], rj[j] = j, i
                break
    pairs = [(t, rolls[ti[i]] if i in ti else None) for i, t in enumerate(trows)]
    pairs += [(None, r) for j, r in enumerate(rolls) if j not in rj]
    return pairs


def build_truth(versions, existing, cfg):
    """Merge the rolls of every version into the truth rows (never dropping what was typed).
    -> ({video: [TruthRoll]}, rows added, values taken from the result workbooks)"""
    tol = cfg["roll_match_tolerance_s"]
    truth = {v: list(rows) for v, rows in existing.items()}
    added = harvested = 0
    for ver in versions:
        for video, run in ver.videos.items():
            rows = truth.setdefault(video, [])
            for t, r in match_rolls(rows, run.rolls, tol):
                if r is None:
                    continue
                if t is None:
                    t = TruthRoll(video=video, n=r.seq, t=r.first_seen)
                    rows.append(t)
                    added += 1
                if r.true_ply and not t.ply:
                    t.ply, harvested = r.true_ply, harvested + 1
                if r.true_range and not t.rng:
                    t.rng, harvested = r.true_range.replace(" ", ""), harvested + 1
    for rows in truth.values():
        rows.sort(key=lambda t: (t.t is None, t.t if t.t is not None else 0, t.n))
        timed = all(t.t is not None for t in rows)
        for i, t in enumerate(rows, 1):
            if timed:
                t.n = i
    return truth, added, harvested


def write_truth(truth, versions, cfg, path=TRUTH_PATH):
    tol = cfg["roll_match_tolerance_s"]
    header = COLS + [f"{v.name} {c}" for v in versions for c in ("ply", "range", "status")]
    rows = [header]
    for video in sorted(truth, key=natural):
        trs = truth[video]
        ref = []
        for v in versions:
            run = v.videos.get(video)
            ref.append({id(t): r for t, r in match_rolls(trs, run.rolls, tol) if t is not None} if run else None)
        for t in trs:
            row = [video, t.n, t.t, S("" if t.real is None else "yes" if t.real else "no", "input"),
                   S(t.ply, "input"), S(t.rng, "input"), S(t.notes, "input")]
            for m in ref:
                r = m.get(id(t)) if m is not None else None
                row += ["", "", S("(video not run)", "dim")] if m is None else \
                    ["", "", S("(not found)", "dim")] if r is None else [r.ply, r.rng, r.status]
            rows.append(row)
    n = len(rows)
    for _ in range(EXTRA_ROWS):   # spare formatted rows for rolls that no version found
        rows.append([S("", "input"), S("", "input"), S("", "input")] + [S("", "input")] * 4)
    help_rows = [
        [S("How to fill the truth sheet", "title")],
        [],
        ["1. Each row of the Truth sheet is one roll of one video, in the order the rolls appear."],
        ["   'First seen (s)' is the video time at which the roll first appears - use it to find the roll in the video."],
        ["2. Type the real ply number in 'True ply' and the real range (e.g. 50.4-68.4) in 'True range'."],
        ["   Leave True range empty if the roll has no range label; the range is then not scored for that roll."],
        ["3. If a row is NOT a real roll (the detector tracked something else, or it is the same roll tracked a"],
        ["   second time), set 'Real roll?' to no."],
        ["4. If a real roll is missing from the list (no version found it), add a row at the bottom:"],
        ["   Video (exact name), First seen (s) (roughly when it appears), True ply, True range."],
        ["5. Fill EVERY roll of a video. A video is scored as soon as one of its rolls is filled; in a scored video"],
        ["   a roll that a version reports but that is not listed counts as a false roll, and a listed roll"],
        ["   that a version never reports counts as missed."],
        ["6. Videos with nothing filled are simply not scored yet."],
        [],
        ["The blue cells are the ones to type in. The columns on the right show what each version decided, for"],
        ["reference only - they are rewritten every time you run:   python kvc.py truth"],
        ["Running that command again never removes what you typed; a backup is kept in truth/backups."],
        ["Save and CLOSE this file before running the pipeline (Excel locks it while it is open)."],
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        bdir = path.parent / "backups"
        bdir.mkdir(exist_ok=True)
        shutil.copy2(path, bdir / f"truth_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    try:
        write_workbook(path, [
            {"name": "Truth", "rows": rows, "header_rows": 1, "freeze_cols": 3, "autofilter": True,
             "widths": [30, 7, 10, 9, 10, 14, 30] + [12, 14, 16] * len(versions),
             "validations": [(f"D2:D{n + EXTRA_ROWS}", "yes,no")]},
            {"name": "How to fill", "rows": help_rows, "widths": [120]}])
    except PermissionError:
        raise SystemExit(f"Cannot write {path}\nIt is open in Excel - close it and run the command again.")
