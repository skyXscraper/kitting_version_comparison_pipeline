"""Turn the numbers into plain-language observations: what is abnormal, and where the errors come from.

Each observation is (level, text) with level ALERT (wrong outcome on the line), WARN (worth a look) or INFO.
"""
import re
import statistics
from collections import Counter

RANGE_RE = re.compile(r"^(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)$")
MAX_LIST = 8


def pct(x, digits=0):
    return "-" if x is None else f"{100 * x:.{digits}f}%"


def _where(rec):
    return f"{rec['video']} roll {rec['n']}" if rec["n"] is not None else f"{rec['video']} seq {rec['roll'].seq}"


def _top(hist, k=4):
    return ", ".join(f"'{t}' x{n}" for t, n in sorted(hist, key=lambda x: -x[1])[:k]) or "none"


def _listed(items):
    items = list(items)
    more = f" ... and {len(items) - MAX_LIST} more (see the Rolls sheet)" if len(items) > MAX_LIST else ""
    return "; ".join(items[:MAX_LIST]) + more


def malformed_ranges(rolls):
    bad = []
    for r in rolls:
        if not r.rng:
            continue
        m = RANGE_RE.match(r.rng)
        if not m or float(m.group(1)) >= float(m.group(2)):
            bad.append(r)
    return bad


def observe(res, cfg, truth_issues=(), brief=False):
    o = []
    add = lambda level, text: o.append((level, text))
    m, ver, recs, runs = res["overall"], res["version"], res["recs"], res["runs"]
    real = [r for r in recs if r["kind"] == "real"]
    found = [r for r in real if r["roll"] is not None]

    # ---- how much the score can be trusted
    if not m["ply_n"]:
        add("ALERT", "No true ply values for this version's videos yet, so there is no accuracy score. "
                     "Run 'python kvc.py truth', fill truth/truth.xlsx, then score again.")
    elif m["ply_n"] < cfg["min_rolls_reliable"] and not brief:
        lo, hi = m["ply_ci"]
        add("WARN", f"Only {m['ply_n']} labelled roll(s): the score is shaky. With this few, the true ply accuracy "
                    f"could be anywhere from {pct(lo)} to {pct(hi)} (95% interval). Label at least "
                    f"{cfg['min_rolls_reliable']} rolls before trusting small differences.")
    if m["ply_n"] and m["videos"] > m["videos_scored"]:
        add("INFO", f"{m['videos'] - m['videos_scored']} of {m['videos']} videos have no truth yet and are not in "
                    "the accuracy score (speed numbers use all videos).")
    for t in truth_issues:
        add("WARN", t)
    for w in ver.warnings:
        add("INFO", f"Input: {w}")
    dup = [r for run in runs.values() for r in run.rolls if r.status == "duplicate"]
    if dup:
        add("INFO", f"{len(dup)} roll(s) were flagged 'duplicate' by the run itself (same ply as an earlier roll): "
                    + _listed(f"{r.video} seq {r.seq} ply {r.ply}" for r in dup)
                    + ". If one is a re-tracked roll, mark it Real roll? = no in the truth sheet; it then counts "
                      "neither for nor against the detector.")
    partial = [f"{v} ({run.stats['whole']})" for v, run in runs.items() if str(run.stats.get("whole", "")).startswith("no")]
    if partial:
        add("WARN", f"Partial run(s), the video was not processed to the end: {_listed(partial)}. "
                    "Rolls late in these videos may be counted as missed.")

    # ---- wrong outcomes
    wp = [r for r in real if r["ply_verdict"] == "wrong_pick"]
    if wp:
        add("ALERT", f"{len(wp)} roll(s) were PICKED with a wrong ply (would go to a wrong bin): " + _listed(
            f"{_where(r)} decided {r['roll'].ply} / true {r['true_ply']} -> bin {r['roll'].bin or '?'}"
            + (" (close: partial credit)" if r["ply_close"] else "") for r in wp))
    bad = [r for r in found if r["ply_verdict"] in ("wrong_pick", "wrong", "no_read")]
    outvoted = [r for r in bad if r.get("ply_reads_right")]
    never = [r for r in bad if not r.get("ply_reads_right")]
    if outvoted:
        add("WARN", f"{len(outvoted)} roll(s) where the true ply WAS read but the decision went elsewhere "
                    "(voting / decision rule problem, not the recognizer): " + _listed(
                        f"{_where(r)} true {r['true_ply']} read {r['ply_reads_right']}/{r['ply_reads']} times, "
                        f"decided '{r['roll'].ply or 'nothing'}' ({r['roll'].ply_rule})" for r in outvoted))
    if never:
        add("WARN", f"{len(never)} roll(s) where the true ply was NEVER read in any crop (recognizer or crop "
                    "quality problem): " + _listed(
                        f"{_where(r)} true {r['true_ply']}, {r['ply_reads']} ply reads, most common "
                        f"{_top(r['wrong_reads'].items(), 3)}" for r in never))
    missed = [r for r in real if r["roll"] is None]
    if missed:
        add("ALERT", f"{len(missed)} true roll(s) never reported (detector / tracker missed them, or the first-seen "
                     f"time in the truth sheet is more than {cfg['roll_match_tolerance_s']:g} s off): "
                     + _listed(f"{_where(r)} at ~{r['t']} s" for r in missed))
    false = [r for r in recs if r["kind"] == "false"]
    if false:
        acted = sum(r["roll"].status == "picked" for r in false)
        add("WARN", f"{len(false)} reported roll(s) are not real rolls according to the truth sheet"
                    + (f", {acted} of them were even PICKED into a bin" if acted else "") + ": " + _listed(
                        f"{r['video']} seq {r['roll'].seq} at {r['roll'].first_seen} s ({r['roll'].status}"
                        f"{', ply ' + r['roll'].ply if r['roll'].ply else ''})" for r in false))

    # ---- where the errors concentrate
    acc = m["comp"]["ply_decision"]
    groups = {g: gm for g, gm in res["by_group"].items() if gm["ply_n"] >= 3}
    if acc is not None and len(groups) > 1:
        weak = [f"{g} {gm['ply_right']}/{gm['ply_n']}" for g, gm in groups.items()
                if gm["ply_right"] / gm["ply_n"] <= m["ply_right"] / m["ply_n"] - 0.25]
        if weak:
            add("WARN", f"Recording session(s) clearly below the overall ply accuracy ({m['ply_right']}/{m['ply_n']}): "
                        f"{', '.join(weak)}. Look at lighting / camera angle / label type there.")
    wrong = Counter()
    extra = trunc = 0
    for r in found:
        tp = r["true_ply"]
        for text, n in r.get("wrong_reads", {}).items():
            wrong[text] += n
            if tp and text and tp in text and len(text) > len(tp):
                extra += n
            elif tp and text and text in tp:
                trunc += n
    nw = sum(wrong.values())
    if nw:
        add("INFO", f"Single ply reads: {m['ply_reads_right']} right of {m['ply_reads']} ({pct(m['comp']['ply_read_acc'])}). "
                    f"Most common wrong reads: {_top(wrong.items(), 6)}.")
        if extra / nw >= 0.2:
            add("WARN", f"{pct(extra / nw)} of the wrong ply reads contain the true ply plus extra characters "
                        "(e.g. 108 read as 1081): the recognizer is adding characters - check crop margins / "
                        "neighbouring text in the crop.")
        if trunc / nw >= 0.2:
            add("WARN", f"{pct(trunc / nw)} of the wrong ply reads are a piece of the true ply (digits dropped, e.g. "
                        "235 read as 23): partial crops, blur or the label entering / leaving the frame.")
        flips = sum(r.get("flip_right", 0) for r in found)
        if flips / nw >= 0.05:
            add("WARN", f"{flips} wrong reads would have been right the OTHER way up: the upright / flipped choice "
                        "is picking the wrong orientation.")
    cr = [s for r in found for s in r.get("conf_right", [])]
    cw = [s for r in found for s in r.get("conf_wrong", [])]
    if cw:
        hi = sum(s >= cfg["high_conf"] for s in cw) / len(cw)
        line = f"Read confidence: wrong reads average {statistics.mean(cw):.2f}" + \
               (f", right reads {statistics.mean(cr):.2f}" if cr else ", there are no right reads to compare with")
        if hi >= 0.2:
            add("WARN", f"{line}. {pct(hi)} of the wrong reads have a score >= {cfg['high_conf']}: the model is "
                        "confidently wrong, so a score threshold will not filter these out.")
        elif cr and statistics.mean(cr) - statistics.mean(cw) < 0.05:
            add("WARN", f"{line}. The score barely separates right from wrong reads, so it is not a useful filter.")
        else:
            add("INFO", line + ".")

    # ---- range
    rolls = [r for run in runs.values() for r in run.rolls]
    mal = malformed_ranges(rolls)
    with_range = [r for r in rolls if r.rng]
    if mal:
        add("WARN", f"{len(mal)} of {len(with_range)} decided ranges are not a valid 'low-high' range "
                    f"(e.g. {', '.join(repr(r.rng) for r in mal[:5])}). No truth is needed to see these are wrong.")
    if m["range_n"]:
        none = sum(1 for r in found if r["range_credit"] is not None and not r["roll"].rng)
        add("INFO", f"Range: {m['range_exact']} exact and {m['range_partial']} close of {m['range_n']} rolls with a "
                    f"true range; {none} had no range decided at all.")

    # ---- speed and load
    st = m["status"]
    if rolls:
        add("INFO", "Roll outcomes over all videos: " + ", ".join(f"{k or '(blank)'} {n}" for k, n in st.most_common())
            + f" (of {len(rolls)}).")
    slow = sorted(((run.stats["fps"], v) for v, run in runs.items()
                   if run.stats.get("fps") is not None and run.stats["fps"] < 0.9 * cfg["target_fps"]))
    if slow:
        add("WARN", f"Frame rate fell below 90% of the {cfg['target_fps']:g} fps target in {len(slow)} video(s) "
                    f"(frames were processed late): " + _listed(f"{v} {f:g} fps" for f, v in slow))
    ocr = {v: run.stats["ocr_mean"] for v, run in runs.items() if run.stats.get("ocr_mean") is not None}
    if len(ocr) >= 4:
        med = statistics.median(ocr.values())
        out = [f"{v} {x:g} ms" for v, x in sorted(ocr.items(), key=lambda kv: -kv[1]) if x > 1.6 * med]
        if out:
            add("WARN", f"OCR time per crop is unusually high (median over videos {med:.1f} ms) in: {_listed(out)}. "
                        "Usually machine load during that run rather than the model - rerun before comparing speed.")
    lat = sorted(((r.latency, f"{r.video} seq {r.seq} {r.latency:g} s") for r in rolls
                  if r.latency is not None and r.latency >= cfg["latency_bad_s"]), reverse=True)
    if lat:
        add("WARN", f"{len(lat)} roll(s) took {cfg['latency_bad_s']:g} s or more to decide: " + _listed(x[1] for x in lat))
    queue = [rd.queue_ms for run in runs.values() for rd in run.reads if rd.queue_ms is not None]
    if queue and max(queue) >= 200:
        add("WARN", f"OCR queue backed up: a crop waited {max(queue):.0f} ms (mean {statistics.mean(queue):.1f} ms). "
                    "The recognizer is not keeping up with the detector at peak.")
    votes, reads = sum(rd.vote for run in runs.values() for rd in run.reads), sum(len(run.reads) for run in runs.values())
    if reads and votes / reads < 0.8:
        add("INFO", f"Only {pct(votes / reads)} of the OCR reads counted as votes; the rest were neither a ply nor a range.")
    return o


def observe_compare(results, pairs, cfg, excluded, truth_issues=()):
    """results: evaluated on the same videos; pairs: [(a, b, head_to_head)]; excluded: {version: [videos left out]}."""
    o = []
    add = lambda level, text: o.append((level, text))
    names = [r["version"].name for r in results]
    scored = [r for r in results if r["overall"]["scores"]["overall"]["score"] is not None]
    if not scored:
        add("ALERT", "None of the compared videos has truth yet, so only speed and roll outcomes can be compared. "
                     "Run 'python kvc.py truth', fill truth/truth.xlsx, then compare again.")
    else:
        rank = sorted(scored, key=lambda r: -r["overall"]["scores"]["overall"]["score"])
        best = rank[0]
        line = " > ".join(f"{r['version'].name} {r['overall']['scores']['overall']['score']:.1f}" for r in rank)
        add("INFO", f"Overall ranking: {line}.")
        n = best["overall"]["ply_n"]
        if n < cfg["min_rolls_reliable"]:
            add("WARN", f"The ranking rests on only {n} labelled roll(s). One roll changes the ply accuracy by "
                        f"{pct(1 / n)}; label at least {cfg['min_rolls_reliable']} before calling a winner.")
        for name in ("ocr", "det"):
            sc = [(r["overall"]["scores"][name]["score"], r["version"].name) for r in scored
                  if r["overall"]["scores"][name]["score"] is not None]
            if len(sc) > 1 and max(sc)[0] - min(sc)[0] < 1:
                add("INFO", f"{'OCR' if name == 'ocr' else 'Detection'} scores are within 1 point of each other "
                            f"({', '.join(f'{v} {s:.1f}' for s, v in sc)}): no real difference there.")
    for a, b, h in pairs:
        an, bn = a["version"].name, b["version"].name
        if h["same_n"]:
            add("INFO", f"{an} vs {bn}: same ply decided on {h['same_ply']} of {h['same_n']} rolls both reported "
                        f"({pct(h['same_ply'] / h['same_n'])} agreement, truth not needed for this).")
        if not h["n"]:
            continue
        txt = (f"{an} vs {bn} on the {h['n']} labelled rolls both ran: both right {h['both']}, only {an} right "
               f"{len(h['only_a'])}, only {bn} right {len(h['only_b'])}, both wrong {h['neither']}.")
        if h["p"] is None:
            txt += " They never disagree on right / wrong."
        elif h["p"] <= 0.05:
            txt += f" The difference is statistically reliable (sign test p = {h['p']:.3f})."
        else:
            txt += (f" This is NOT enough to call one better (sign test p = {h['p']:.2f}; it needs to be <= 0.05) - "
                    "more labelled rolls are needed.")
        add("INFO", txt)
        for keys, win, lose in ((h["only_a"], an, bn), (h["only_b"], bn, an)):
            if keys:
                side = h["b"] if lose == bn else h["a"]
                add("WARN", f"Rolls {win} gets right and {lose} gets wrong: " + _listed(
                    f"{k[0]} roll {k[1]} (true {side[k]['true_ply']}, {lose} decided "
                    f"'{side[k]['roll'].ply if side[k]['roll'] else 'roll missed'}')" for k in keys))
    # the same detector should report the same rolls: flag when it does not
    counts = {r["version"].name: (r["overall"]["found"], r["overall"]["false"], r["overall"]["rolls_reported"])
              for r in results}
    if len({c[2] for c in counts.values()}) > 1:
        add("INFO", "Rolls reported on the compared videos differ: "
                    + ", ".join(f"{n} {c[2]}" for n, c in counts.items())
                    + ". With the same detector this should be equal - different detectors or run-to-run tracking noise.")
    for key, label, unit, tol in (("det_ms", "Detection time", "ms", 0.25), ("ocr_ms", "OCR time per crop", "ms", 0.25),
                                  ("latency_median", "Median time to decide", "s", 0.25)):
        vals = [(r["overall"][key], r["version"].name) for r in results if r["overall"][key] is not None]
        if len(vals) > 1 and min(vals)[0] > 0 and max(vals)[0] / min(vals)[0] - 1 >= tol:
            add("INFO", f"{label}: " + ", ".join(f"{n} {v:.1f} {unit}" for v, n in sorted(vals))
                + f" ({pct(max(vals)[0] / min(vals)[0] - 1)} slower at worst). Speeds are only comparable when the "
                  "runs were made on the same machine under the same load.")
    for name, vids in excluded.items():
        if vids:
            add("INFO", f"{name}: {len(vids)} of its videos are not in this comparison because another version "
                        f"did not run them ({_listed(vids)}). Use --all-videos to score each version on its own set.")
    for t in truth_issues:
        add("WARN", t)
    return o
