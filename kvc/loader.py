"""Load result workbooks into versions.

Two folder layouts are understood (both made by the RTracker&OCR log scripts):
  single   rec_v5_logs_v1 style        one model per workbook
  compare  compare_rec_v3_vs_v3_2_v1   two (or more) models side by side in each workbook -> one version per model
A folder holds one workbook per video (sheets Frames / Rolls / Run info) and optionally a summary workbook
(sheets Summary / Rolls) with the frame counts and speeds. Columns are found by their header text, not by position.
"""
import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from .config import VERSIONS_PATH
from .xlsx_io import cell, fnum, read_workbook, txt


@dataclass
class Roll:
    video: str
    seq: int
    track: str = ""
    status: str = ""
    ply: str = ""
    rng: str = ""
    bin: str = ""
    decided_at: float = None
    latency: float = None
    ply_rule: str = ""
    n_reads: int = 0
    ply_reads: list = field(default_factory=list)      # [(text, count)]
    range_reads: list = field(default_factory=list)
    true_ply: str = ""                                 # truth already typed in the workbook, if any
    true_range: str = ""

    @property
    def first_seen(self):
        if self.decided_at is None:
            return None
        return round(self.decided_at - (self.latency or 0.0), 2)


@dataclass
class Read:
    frame: int
    time: float
    det_conf: float
    track: str
    det_ms: float
    fld: str            # 'ply' / 'range' / '-'
    text: str
    score: float
    other: str
    other_score: float
    ocr_ms: float
    queue_ms: float
    total_ms: float
    vote: bool


@dataclass
class VideoRun:
    name: str
    rolls: list = field(default_factory=list)
    reads: list = field(default_factory=list)
    n_boxes: int = None
    stats: dict = field(default_factory=dict)   # frames fps det_mean det_median ocr_reads votes ocr_mean ... whole log


@dataclass
class Version:
    name: str
    folder: str
    kind: str = "single"
    block: int = 0
    tag: str = ""
    det: str = ""
    rec: str = ""
    notes: str = ""
    videos: dict = field(default_factory=dict)
    warnings: list = field(default_factory=list)

    @property
    def label(self):
        bits = [b for b in (self.det, self.rec) if b]
        return f"{self.name} ({' + '.join(bits)})" if bits else self.name


def natural(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def video_group(name, known=()):
    """Recording session of a video: 'oct_3_cam0_2' -> 'oct_3', '19_part2_cam0_19_part2_7' -> '19_part2'."""
    m = re.match(r"^(.*?)_cam\d", name)
    if m:
        return m.group(1)
    for g in sorted(known, key=len, reverse=True):
        if name.startswith(g + "_"):
            return g
    return "other"


def _hist(s):
    out = []
    for tok in txt(s).split():
        t, _, n = tok.rpartition("x")
        out.append((t, int(n)) if t and n.isdigit() else (tok, 1))
    return out


def _find(rows, *must):
    for i, r in enumerate(rows):
        vals = {txt(v) for v in r}
        if all(m in vals for m in must):
            return i
    return None


def _tag(s):
    return re.sub(r"^rec\s+", "", txt(s)).strip()


def _blocks(hdr, start_name, stop_names):
    """Column spans of the repeated per-model blocks: [(start, end)], and the index where the shared tail starts."""
    starts = [i for i, v in enumerate(hdr) if v == start_name]
    if not starts:
        return [], len(hdr)
    tail = min([i for i, v in enumerate(hdr) if v in stop_names and i > starts[-1]] or [len(hdr)])
    return list(zip(starts, starts[1:] + [tail])), tail


def parse_rolls(rows, default_video=None):
    """Rolls sheet -> (tags, [rolls of model 0, rolls of model 1, ...])."""
    h = _find(rows, "Seq", "Status")
    if h is None:
        return [], []
    hdr = [txt(v) for v in rows[h]]
    group = [txt(v) for v in rows[h - 1]] if h else []
    spans, tail = _blocks(hdr, "Track", ("Same ply", "True ply"))
    lead = hdr[:spans[0][0]]
    iv = lead.index("Video") if "Video" in lead else None
    iseq = lead.index("Seq")
    itp = hdr.index("True ply") if "True ply" in hdr else None
    itr = hdr.index("True range") if "True range" in hdr else None
    tags = [_tag(cell(group, a)) for a, _ in spans]
    if not all(tags):
        tail_tags = [v[:-4] for v in hdr[tail:] if v.endswith(" ply") and v != "True ply" and v != "Same ply"]
        tags = tail_tags if len(tail_tags) == len(spans) else [t or f"model{i + 1}" for i, t in enumerate(tags)]
    out = [[] for _ in spans]
    for row in rows[h + 1:]:
        seq = fnum(cell(row, iseq))
        if seq is None:
            continue
        video = txt(cell(row, iv)) if iv is not None else default_video
        for b, (a, e) in enumerate(spans):
            col = {hdr[i]: cell(row, i) for i in range(a, e)}
            if not txt(col.get("Status")) and col.get("Track") is None:
                continue
            out[b].append(Roll(
                video=video, seq=int(seq), track=txt(col.get("Track")), status=txt(col.get("Status")),
                ply=txt(col.get("Ply")), rng=txt(col.get("Range")), bin=txt(col.get("Bin")),
                decided_at=fnum(col.get("Decided at (s)")), latency=fnum(col.get("Decided after first seen (s)")),
                ply_rule=txt(col.get("Ply rule")), n_reads=int(fnum(col.get("OCR reads")) or 0),
                ply_reads=_hist(col.get("Ply reads")), range_reads=_hist(col.get("Range reads")),
                true_ply=txt(cell(row, itp)), true_range=txt(cell(row, itr))))
    return tags, out


def parse_frames(rows):
    """Frames sheet -> (tags, [reads of model 0, ...], number of text boxes)."""
    h = _find(rows, "Det #", "Frame")
    if h is None:
        return [], [], None
    hdr = [txt(v) for v in rows[h]]
    group = [txt(v) for v in rows[h - 1]] if h else []
    spans, _ = _blocks(hdr, "Field", ("Same read", "True ply"))
    if not spans:
        return [], [], None
    lead = hdr[:spans[0][0]]
    li = {name: i for i, name in enumerate(lead)}
    det_cols = [i for i, v in enumerate(lead) if v.startswith("Detection (ms)")]
    tags = [_tag(cell(group, a)) for a, _ in spans]
    out, n = [[] for _ in spans], 0
    for row in rows[h + 1:]:
        if cell(row, li.get("Det #")) is None:
            continue
        n += 1
        for b, (a, e) in enumerate(spans):
            col = {hdr[i]: cell(row, i) for i in range(a, e)}
            if col.get("Field") is None:
                continue
            read, other = txt(col.get("Read")), txt(col.get("Other way up"))
            dc = det_cols[b] if b < len(det_cols) else det_cols[0] if det_cols else None
            out[b].append(Read(
                frame=int(fnum(cell(row, li.get("Frame"))) or 0), time=fnum(cell(row, li.get("Video time (s)"))),
                det_conf=fnum(cell(row, li.get("Det conf"))), track=txt(cell(row, li.get("Roll (track)"))),
                det_ms=fnum(cell(row, dc)), fld=txt(col.get("Field")),
                text="" if read == "(empty)" else read, score=fnum(col.get("Score")),
                other="" if other == "(empty)" else other, other_score=fnum(col.get("Other score")),
                ocr_ms=fnum(col.get("OCR (ms)")), queue_ms=fnum(col.get("Queue (ms)")),
                total_ms=fnum(col.get("Total (ms)")), vote=txt(col.get("Vote")) == "yes"))
    return tags, out, n


def parse_run_info(rows):
    """Run info sheet -> [{det, rec, log}] one per model section."""
    lines = [txt(cell(r, 0)) for r in rows]
    sections, cur = [], None
    for line in lines:
        m = re.match(r"^rec (\S+): test_logs/(\S+)", line)
        if m:
            cur = {"log": m.group(2), "det": "", "rec": ""}
            sections.append(cur)
            continue
        if cur is None:
            cur = {"log": "", "det": "", "rec": ""}
            sections.append(cur)
        if line.startswith("model"):
            d = re.search(r"models[\\/]+([^\\/\s]+)[\\/]", line)
            r = re.search(r"recognizer\s+(\S+)", line)
            cur["det"] = d.group(1) if d else ""
            cur["rec"] = re.split(r"[\\/]", r.group(1))[-1] if r else ""
    return sections


_SUMMARY_COLS = {"Text boxes": "boxes", "Detection mean (ms)": "det_mean", "Detection median (ms)": "det_median",
                 "fps processed": "fps", "OCR reads": "ocr_reads", "Counted as votes": "votes",
                 "OCR mean (ms)": "ocr_mean", "OCR median (ms)": "ocr_median", "Crop total mean (ms)": "total_mean",
                 "Whole video": "whole", "Run (test_logs)": "log"}


def parse_summary(rows):
    """Summary sheet -> [ {video: stats} per model ], target fps or None."""
    h = _find(rows, "Video", "Text boxes")
    if h is None:
        return [], None
    hdr = [txt(v) for v in rows[h]]
    spans, _ = _blocks(hdr, "Detection mean (ms)", ("Same read (of boxes read by both)", "Rolls with truth"))
    if not spans:
        return [], None
    single = len(spans) == 1
    lead_end = spans[0][0]
    target = None
    out = [{} for _ in spans]
    for row in rows[h + 1:]:
        video = txt(cell(row, 0))
        if not video or video.startswith("Total ("):
            continue
        shared = {}
        for i in range(lead_end):
            if hdr[i].startswith("Frames"):
                shared["frames"] = fnum(cell(row, i))
                m = re.search(r"\((\d+(?:\.\d+)?) fps\)", hdr[i])
                target = float(m.group(1)) if m else target
            elif hdr[i] in _SUMMARY_COLS:
                v = cell(row, i)
                shared[_SUMMARY_COLS[hdr[i]]] = txt(v) if _SUMMARY_COLS[hdr[i]] in ("whole", "log") else fnum(v)
        for b, (a, e) in enumerate(spans):
            st = dict(shared)
            # single layout: the whole row belongs to the one model; compare layout: only its own span
            for i in (range(lead_end, len(hdr)) if single else range(a, e)):
                if hdr[i] in _SUMMARY_COLS:
                    v = cell(row, i)
                    key = _SUMMARY_COLS[hdr[i]]
                    st[key] = txt(v) if key in ("whole", "log") else fnum(v)
            out[b][video] = st
    return out, target


def load_folder(folder):
    """-> {'kind', 'tags', 'models': [{'videos': {name: VideoRun}, 'det', 'rec'}], 'warnings', 'target_fps'}."""
    folder = Path(folder)
    if not folder.is_dir():
        raise SystemExit(f"Folder not found: {folder}")
    models, tags, warnings, summaries = [], [], [], []
    for p in sorted(folder.glob("*.xlsx"), key=lambda p: natural(p.stem)):
        if p.name.startswith("~$"):
            continue
        try:
            wb = read_workbook(p)
        except Exception as e:  # a half-written or locked file should not stop the whole load
            warnings.append(f"{p.name}: could not be read ({e})")
            continue
        if "Frames" in wb and "Rolls" in wb:
            rtags, rolls = parse_rolls(wb["Rolls"], p.stem)
            ftags, reads, n_boxes = parse_frames(wb["Frames"])
            if not rolls and not reads:
                warnings.append(f"{p.name}: no Rolls / Frames table found")
                continue
            n = max(len(rolls), len(reads))
            if models and n != len(models):
                warnings.append(f"{p.name}: has {n} model block(s), the folder has {len(models)}; skipped")
                continue
            if not models:
                models = [{"videos": {}, "det": "", "rec": ""} for _ in range(n)]
                tags = [t for t in (ftags if all(ftags) and len(ftags) == n else rtags)] or [f"model{i + 1}" for i in range(n)]
            info = parse_run_info(wb.get("Run info", []))
            for b in range(n):
                run = VideoRun(p.stem, rolls[b] if b < len(rolls) else [], reads[b] if b < len(reads) else [], n_boxes)
                if b < len(info):
                    run.stats["log"] = info[b]["log"]
                    models[b]["det"] = models[b]["det"] or info[b]["det"]
                    models[b]["rec"] = models[b]["rec"] or info[b]["rec"]
                models[b]["videos"][p.stem] = run
        elif "Summary" in wb:
            summaries.append((p.name, wb))
        else:
            warnings.append(f"{p.name}: not a result workbook (no Frames/Rolls or Summary sheet); ignored")
    target = None
    for name, wb in summaries:
        stats, t = parse_summary(wb["Summary"])
        target = t or target
        if not models and stats:
            models = [{"videos": {}, "det": "", "rec": ""} for _ in stats]
        if stats and len(stats) != len(models):
            warnings.append(f"{name}: summary has {len(stats)} model block(s), expected {len(models)}; ignored")
            continue
        stags, srolls = parse_rolls(wb.get("Rolls", []))
        tags = tags or stags
        for b, per_video in enumerate(stats):
            for video, st in per_video.items():
                run = models[b]["videos"].get(video)
                if run is None:   # video only present in the summary workbook: rolls but no per-read rows
                    run = models[b]["videos"][video] = VideoRun(video)
                    if b < len(srolls):
                        run.rolls = [r for r in srolls[b] if r.video == video]
                    warnings.append(f"{video}: only in {name} (no per-video workbook), read-level detail missing")
                log = run.stats.get("log")
                run.stats.update({k: v for k, v in st.items() if v not in (None, "")})
                if log and not run.stats.get("log"):
                    run.stats["log"] = log
    if not models:
        raise SystemExit(f"No result workbooks found in {folder}\n(expected one .xlsx per video with Frames and Rolls "
                         "sheets, like rec_v5_logs_v1 or compare_rec_v3_vs_v3_2_v1)")
    return {"kind": "single" if len(models) == 1 else "compare", "tags": tags, "models": models,
            "warnings": warnings, "target_fps": target}


# ---------------------------------------------------------------- registry
def load_registry():
    if not VERSIONS_PATH.exists():
        return []
    return json.loads(VERSIONS_PATH.read_text(encoding="utf-8")).get("versions", [])


def save_registry(entries):
    VERSIONS_PATH.write_text(json.dumps({"versions": entries}, indent=2), encoding="utf-8")


_folder_cache = {}


def load_version(entry):
    key = entry["folder"]
    if key not in _folder_cache:
        _folder_cache[key] = load_folder(key)
    data = _folder_cache[key]
    b = entry.get("block", 0)
    if b >= len(data["models"]):
        raise SystemExit(f"Version '{entry['name']}': model block {b + 1} no longer exists in {key}")
    m = data["models"][b]
    return Version(name=entry["name"], folder=key, kind=data["kind"], block=b,
                   tag=data["tags"][b] if b < len(data["tags"]) else "", det=entry.get("det") or m["det"],
                   rec=entry.get("rec") or m["rec"], notes=entry.get("notes", ""), videos=m["videos"],
                   warnings=list(data["warnings"]))


def load_versions(names=None):
    entries = load_registry()
    if not entries:
        raise SystemExit("No versions registered yet. Add one first:\n  python kvc.py add <folder with the result workbooks>")
    by_name = {e["name"]: e for e in entries}
    if names:
        missing = [n for n in names if n not in by_name]
        if missing:
            raise SystemExit(f"Unknown version(s): {', '.join(missing)}\nRegistered: {', '.join(by_name)}")
        entries = [by_name[n] for n in names]
    return [load_version(e) for e in entries]
