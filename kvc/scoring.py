"""Score a version against the truth sheet.

Every roll becomes a record (kind: real / false / ignored / unlabelled); records are then aggregated into metrics,
0..1 components, and the three scores (overall, OCR, detection). See README.md > Scoring for what each part means.
"""
import math
import statistics
from collections import Counter, defaultdict

from .loader import natural, video_group
from .truth import is_labelled, match_rolls


def lev(a, b):
    """Edit distance: characters to insert / delete / replace to turn a into b."""
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _credit(read, true, table):
    """1 for an exact read, the partial credit of the table for a close one (by characters off), else 0."""
    read, true = read.replace(" ", ""), true.replace(" ", "")
    if not read:
        return 0.0
    if read == true:
        return 1.0
    return float(table.get(str(lev(read, true)), 0.0))


def range_credit(read, true, cfg):
    return _credit(read, true, cfg["range_partial_credit"])


def ply_credit(read, true, cfg):
    return _credit(read, true, cfg["ply_partial_credit"])


def lin(x, good, bad):
    """1 at or below good, 0 at or above bad, linear in between; None stays None."""
    if x is None:
        return None
    return max(0.0, min(1.0, (bad - x) / (bad - good)))


def wilson(k, n, z=1.96):
    """95% interval of a proportion k/n -> (low, high)."""
    if not n:
        return None
    p, d = k / n, 1 + z * z / n
    c, h = (p + z * z / (2 * n)) / d, z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return max(0.0, c - h), min(1.0, c + h)


def _read_stats(rec, run, by_track, cfg):
    """Raw recognizer quality on one found roll, before any voting."""
    r, tp, tr = rec["roll"], rec["true_ply"], rec["true_range"]
    s = {"ply_reads": 0, "ply_reads_right": 0, "ply_reads_credit": 0.0, "range_reads": 0, "range_reads_credit": 0.0, "flip_right": 0,
         "conf_right": [], "conf_wrong": [], "wrong_reads": Counter()}
    if run.reads:
        for rd in by_track.get(r.track, []):
            if rd.fld == "ply" and tp:
                s["ply_reads"] += 1
                s["ply_reads_credit"] += ply_credit(rd.text, tp, cfg)
                if rd.text == tp:
                    s["ply_reads_right"] += 1
                    if rd.score is not None:
                        s["conf_right"].append(rd.score)
                else:
                    s["wrong_reads"][rd.text] += 1
                    if rd.score is not None:
                        s["conf_wrong"].append(rd.score)
                    s["flip_right"] += rd.other == tp
            elif rd.fld == "range" and tr:
                s["range_reads"] += 1
                s["range_reads_credit"] += range_credit(rd.text, tr, cfg)
                s["flip_right"] += rd.text != tr and rd.other.replace(" ", "") == tr
    else:   # no per-read rows: use the vote histograms of the Rolls sheet
        if tp:
            for text, n in r.ply_reads:
                s["ply_reads"] += n
                s["ply_reads_credit"] += n * ply_credit(text, tp, cfg)
                if text == tp:
                    s["ply_reads_right"] += n
                else:
                    s["wrong_reads"][text] += n
        if tr:
            for text, n in r.range_reads:
                s["range_reads"] += n
                s["range_reads_credit"] += n * range_credit(text, tr, cfg)
    rec.update(s)


def build_records(ver, truth, cfg, videos=None):
    tol = cfg["roll_match_tolerance_s"]
    recs, runs = [], {}
    for video in sorted(ver.videos, key=natural):
        if videos is not None and video not in videos:
            continue
        run = runs[video] = ver.videos[video]
        trows = truth.get(video, [])
        labelled = is_labelled(trows)
        by_track = defaultdict(list)
        for rd in run.reads:
            by_track[rd.track].append(rd)
        for t, r in match_rolls(trows, run.rolls, tol):
            if not labelled:
                kind = "unlabelled"
            elif t is None or t.real is False:
                # a re-tracked roll the run itself flagged as 'duplicate' was handled, so it is not a false roll
                kind = "false" if r is not None and r.status != "duplicate" else None
            elif not t.filled:
                kind = "ignored"
            else:
                kind = "real"
            if kind is None:
                continue
            rec = {"video": video, "kind": kind, "n": t.n if t else None,
                   "t": t.t if t and t.t is not None else r.first_seen if r else None,
                   "true_ply": t.ply if t and kind == "real" else "", "true_range": t.rng if t and kind == "real" else "",
                   "roll": r, "ply_verdict": None, "ply_credit": None, "ply_close": False, "range_credit": None,
                   "notes": t.notes if t else ""}
            if kind == "real":
                if rec["true_ply"]:
                    rec["ply_verdict"] = ("missed" if r is None else "right" if r.ply == rec["true_ply"] else
                                          "no_read" if not r.ply else "wrong_pick" if r.status == "picked" else "wrong")
                    rec["ply_credit"] = 0.0 if r is None else ply_credit(r.ply, rec["true_ply"], cfg)
                    rec["ply_close"] = 0 < rec["ply_credit"] < 1     # wrong, but only just: partial credit
                if rec["true_range"]:
                    rec["range_credit"] = 0.0 if r is None else range_credit(r.rng, rec["true_range"], cfg)
                if r is not None:
                    _read_stats(rec, run, by_track, cfg)
            recs.append(rec)
    return recs, runs


def _wmean(pairs):
    pairs = [(v, w) for v, w in pairs if v is not None and w]
    return sum(v * w for v, w in pairs) / sum(w for _, w in pairs) if pairs else None


def _ratio(a, b):
    return a / b if b else None


def aggregate(recs, runs, cfg):
    """Metrics + 0..1 components + scores for any set of records (a whole version, one video, one session)."""
    pen = cfg["wrong_pick_penalty"]
    real = [r for r in recs if r["kind"] == "real"]
    found = [r for r in real if r["roll"] is not None]
    false = [r for r in recs if r["kind"] == "false"]
    ply = [r for r in real if r["ply_verdict"]]
    ply_found = [r for r in ply if r["roll"] is not None]
    rng = [r for r in real if r["range_credit"] is not None]
    rng_found = [r for r in rng if r["roll"] is not None]
    v = Counter(r["ply_verdict"] for r in ply)
    all_rolls = [r for run in runs.values() for r in run.rolls]
    lat = [r.latency for r in all_rolls if r.latency is not None]
    m = {
        "videos": len(runs), "videos_scored": len({r["video"] for r in real}), "rolls_reported": len(all_rolls),
        "status": Counter(r.status for r in all_rolls),
        "truth_rolls": len(real), "found": len(found), "missed": len(real) - len(found), "false": len(false),
        "ply_n": len(ply), "ply_right": v["right"], "ply_wrong_pick": v["wrong_pick"], "ply_wrong": v["wrong"],
        "ply_no_read": v["no_read"], "ply_missed": v["missed"], "ply_ci": wilson(v["right"], len(ply)),
        "ply_close": sum(r["ply_close"] for r in ply),
        "ply_penalized": sum(r["ply_verdict"] == "wrong_pick" and not r["ply_close"] for r in ply),
        "ply_reads_credit": sum(r.get("ply_reads_credit", 0.0) for r in found),
        "range_n": len(rng), "range_exact": sum(r["range_credit"] == 1.0 for r in rng),
        "range_partial": sum(0 < r["range_credit"] < 1 for r in rng),
        "ply_reads": sum(r.get("ply_reads", 0) for r in found),
        "ply_reads_right": sum(r.get("ply_reads_right", 0) for r in found),
        "range_reads": sum(r.get("range_reads", 0) for r in found),
        "range_reads_credit": sum(r.get("range_reads_credit", 0.0) for r in found),
        "latency_median": statistics.median(lat) if lat else None, "latency_max": max(lat) if lat else None,
        "fps": _wmean((run.stats.get("fps"), 1) for run in runs.values()),
        "det_ms": _wmean((run.stats.get("det_mean"), run.stats.get("frames") or 1) for run in runs.values()),
        "ocr_ms": _wmean((run.stats.get("ocr_mean"), run.stats.get("ocr_reads") or 1) for run in runs.values()),
        "ocr_reads": sum(len(run.reads) or int(run.stats.get("ocr_reads") or 0) for run in runs.values()),
        "boxes": sum(run.n_boxes or int(run.stats.get("boxes") or 0) for run in runs.values()),
    }
    m["pen"] = pen
    m["covered"] = sum(r["roll"].n_reads >= cfg["min_reads_for_coverage"] for r in found)
    reported = m["found"] + m["false"]
    rec_, prec = _ratio(m["found"], m["truth_rolls"]), _ratio(m["found"], reported)
    credit = lambda rows: max(0.0, (sum(r["ply_credit"] for r in rows) - pen * sum(
        r["ply_verdict"] == "wrong_pick" and not r["ply_close"] for r in rows)) / len(rows)) if rows else None
    m["comp"] = {
        "ply_decision": credit(ply),
        "ply_on_found": credit(ply_found),
        "range_decision": _ratio(sum(r["range_credit"] for r in rng), len(rng)),
        "range_on_found": _ratio(sum(r["range_credit"] for r in rng_found), len(rng_found)),
        "roll_recall": rec_, "roll_precision": prec,
        "roll_f1": None if rec_ is None or prec is None else 0.0 if rec_ + prec == 0 else 2 * rec_ * prec / (rec_ + prec),
        "label_coverage": _ratio(m["covered"], len(found)),
        "ply_read_acc": _ratio(m["ply_reads_credit"], m["ply_reads"]),
        "range_read_acc": _ratio(m["range_reads_credit"], m["range_reads"]),
        "latency": lin(m["latency_median"], cfg["latency_good_s"], cfg["latency_bad_s"]),
        "throughput": None if m["fps"] is None else min(1.0, m["fps"] / cfg["target_fps"]),
        "det_speed": lin(m["det_ms"], cfg["det_ms_good"], cfg["det_ms_bad"]),
        "ocr_speed": lin(m["ocr_ms"], cfg["ocr_ms_good"], cfg["ocr_ms_bad"]),
    }
    m["scores"] = {}
    for name, need in (("overall", "ply_decision"), ("ocr", "ply_on_found"), ("det", "roll_recall")):
        weights = cfg[f"{name}_weights"]
        have = {k: w for k, w in weights.items() if m["comp"].get(k) is not None and w}
        if m["comp"].get(need) is None or not have:     # no truth for the part that matters: no score
            m["scores"][name] = {"score": None, "parts": []}
            continue
        scale = 100.0 / sum(have.values())
        parts = [(k, m["comp"][k], w * scale, m["comp"][k] * w * scale) for k, w in have.items()]
        m["scores"][name] = {"score": sum(p[3] for p in parts), "parts": parts,
                             "left_out": [k for k in weights if k not in have]}
    return m


def grade(score, cfg):
    if score is None:
        return "-"
    for g, lo in sorted(cfg["grades"].items(), key=lambda kv: -kv[1]):
        if score >= lo:
            return g
    return "E"


def evaluate(ver, truth, cfg, videos=None):
    recs, runs = build_records(ver, truth, cfg, videos)
    groups = {video_group(v) for v in runs} - {"other"}
    res = {"version": ver, "recs": recs, "runs": runs, "overall": aggregate(recs, runs, cfg),
           "by_video": {}, "by_group": {}}
    by_group = defaultdict(list)
    for video, run in runs.items():
        res["by_video"][video] = aggregate([r for r in recs if r["video"] == video], {video: run}, cfg)
        by_group[video_group(video, groups)].append(video)
    for g, vids in by_group.items():
        res["by_group"][g] = aggregate([r for r in recs if r["video"] in vids], {v: runs[v] for v in vids}, cfg)
    return res


def head_to_head(ra, rb):
    """Paired comparison of two results on the truth rolls both were scored on."""
    a = {(r["video"], r["n"]): r for r in ra["recs"] if r["ply_verdict"]}
    b = {(r["video"], r["n"]): r for r in rb["recs"] if r["ply_verdict"]}
    keys = sorted(set(a) & set(b), key=lambda k: (natural(k[0]), k[1]))
    ok = lambda r: r["ply_verdict"] == "right"
    only_a = [k for k in keys if ok(a[k]) and not ok(b[k])]
    only_b = [k for k in keys if ok(b[k]) and not ok(a[k])]
    n = len(only_a) + len(only_b)
    # exact two-sided sign test on the rolls where the two versions differ
    p = min(1.0, 2 * sum(math.comb(n, i) for i in range(min(len(only_a), len(only_b)) + 1)) / 2 ** n) if n else None
    ua = {(r["video"], r["n"]): r["roll"] for r in ra["recs"] if r["n"] is not None and r["roll"] is not None}
    ub = {(r["video"], r["n"]): r["roll"] for r in rb["recs"] if r["n"] is not None and r["roll"] is not None}
    both = set(ua) & set(ub)
    return {"n": len(keys), "both": sum(ok(a[k]) and ok(b[k]) for k in keys), "only_a": only_a, "only_b": only_b,
            "neither": sum(not ok(a[k]) and not ok(b[k]) for k in keys), "p": p, "a": a, "b": b,
            "same_ply": sum(ua[k].ply == ub[k].ply for k in both), "same_n": len(both)}
