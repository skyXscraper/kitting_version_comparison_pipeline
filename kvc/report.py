"""Console text and Excel workbooks for a single-version score and for a comparison of versions."""
import textwrap

from .loader import natural, video_group
from .observations import pct
from .scoring import grade
from .xlsx_io import S

LABELS = {
    "ply_decision": "Ply decision (of all true rolls)",
    "ply_on_found": "Ply decision (of rolls the detector found)",
    "range_decision": "Range decided right (of all true rolls)",
    "range_on_found": "Range decided right (of found rolls)",
    "roll_f1": "Rolls found without false rolls (F1)",
    "roll_recall": "True rolls found",
    "roll_precision": "Reported rolls that are real",
    "label_coverage": "Found rolls with enough label reads",
    "ply_read_acc": "Single ply reads right (before voting)",
    "range_read_acc": "Single range reads right (before voting)",
    "latency": "Time to decide a roll",
    "throughput": "Frame rate kept",
    "det_speed": "Detection time per frame",
    "ocr_speed": "OCR time per crop",
}
SCORES = (("overall", "OVERALL", "end to end: right ply, right range, rolls found, speed"),
          ("ocr", "OCR", "recognizer: judged only on the rolls the detector found"),
          ("det", "DETECTION", "detector: rolls found, false rolls, labels covered, speed"))
VERDICT = {"right": ("right", "good"), "wrong_pick": ("WRONG, picked", "bad"), "wrong": ("wrong", "bad"),
           "no_read": ("no ply decided", "warn"), "missed": ("roll missed", "bad")}


def num(x, d=1, unit=""):
    return "-" if x is None else f"{x:.{d}f}{unit}"


def detail(k, m):
    """The raw counts behind a component."""
    if k in ("ply_decision", "ply_on_found"):
        s = f"{m['ply_right']} right of {m['ply_n'] - (m['ply_missed'] if k == 'ply_on_found' else 0)}"
        bits = [f"{m[x]} {t}" for x, t in (("ply_wrong_pick", "wrong + picked"), ("ply_wrong", "wrong"),
                                           ("ply_no_read", "no ply"), ("ply_missed", "roll missed")) if m[x]
                and not (k == "ply_on_found" and x == "ply_missed")]
        if m["ply_wrong_pick"] and m["pen"]:
            bits.append(f"each wrong pick also costs {m['pen']:g} of a roll")
        return s + (f" ({', '.join(bits)})" if bits else "")
    if k in ("range_decision", "range_on_found"):
        return f"{m['range_exact']} exact, {m['range_partial']} close of {m['range_n']} with a true range"
    if k in ("roll_f1", "roll_recall", "roll_precision"):
        return f"{m['found']} of {m['truth_rolls']} true rolls found, {m['false']} false roll(s)"
    if k == "label_coverage":
        return f"{m['covered']} of {m['found']} found rolls"
    if k == "ply_read_acc":
        return f"{m['ply_reads_right']} of {m['ply_reads']} reads"
    if k == "range_read_acc":
        return f"credit {m['range_reads_credit']:.1f} of {m['range_reads']} reads"
    if k == "latency":
        return f"median {num(m['latency_median'], 2, ' s')}, worst {num(m['latency_max'], 1, ' s')}"
    if k == "throughput":
        return f"{num(m['fps'], 1)} fps processed"
    if k == "det_speed":
        return f"mean {num(m['det_ms'], 1, ' ms')}"
    if k == "ocr_speed":
        return f"mean {num(m['ocr_ms'], 1, ' ms')}"
    return ""


def _wrap(level, text, width=110):
    head = f"  [{level}] "
    return textwrap.fill(text, width=width, initial_indent=head, subsequent_indent=" " * len(head))


def _obs_text(obs):
    order = {"ALERT": 0, "WARN": 1, "INFO": 2}
    return [_wrap(lv, t) for lv, t in sorted(obs, key=lambda x: order[x[0]])] or ["  (nothing unusual found)"]


def _obs_rows(obs):
    order = {"ALERT": 0, "WARN": 1, "INFO": 2}
    style = {"ALERT": "bad", "WARN": "warn", "INFO": None}
    return [[S(lv, style[lv]) if style[lv] else lv, S(t, "wrap")] for lv, t in sorted(obs, key=lambda x: order[x[0]])]


def _scoring_rows(cfg):
    rows = [["Setting", "Value", "Meaning"]]
    for name, title in (("overall_weights", "OVERALL score"), ("ocr_weights", "OCR score"), ("det_weights", "DETECTION score")):
        rows.append([S(title, "bold"), "", "weight = points out of 100; a part with no data is left out and the rest rescaled"])
        rows += [[f"   {LABELS[k]}", w, k] for k, w in cfg[name].items()]
    rows.append([S("Rules", "bold")])
    rows += [["wrong_pick_penalty", cfg["wrong_pick_penalty"], "a roll picked with a wrong ply loses this much of a roll on top of getting no credit"],
             ["range_partial_credit", str(cfg["range_partial_credit"]), "range credit by characters off: exact = 1"],
             ["roll_match_tolerance_s", cfg["roll_match_tolerance_s"], "a version's roll matches a truth roll when first seen within this many seconds"],
             ["min_reads_for_coverage", cfg["min_reads_for_coverage"], "OCR reads a found roll needs to count as 'labels covered'"],
             ["target_fps", cfg["target_fps"], "frame rate the run should keep"],
             ["latency good / bad (s)", f"{cfg['latency_good_s']} / {cfg['latency_bad_s']}", "full marks at good, zero at bad"],
             ["detection good / bad (ms)", f"{cfg['det_ms_good']} / {cfg['det_ms_bad']}", "full marks at good, zero at bad"],
             ["OCR good / bad (ms)", f"{cfg['ocr_ms_good']} / {cfg['ocr_ms_bad']}", "full marks at good, zero at bad"],
             ["grades", str(cfg["grades"]), "lowest score for each grade; below D is E"]]
    return {"name": "Scoring", "rows": rows, "header_rows": 1, "widths": [46, 22, 90]}


def _rec_cells(rec):
    """Decision of one version on one roll: ply, verdict, range, credit, status, reads."""
    r = rec["roll"] if rec else None
    if rec is None:
        return [S("(video not run)", "dim")] + [""] * 6
    v = VERDICT.get(rec["ply_verdict"])
    verdict = S(*v) if v else S("false roll", "warn") if rec["kind"] == "false" else ""
    if r is None:
        return ["", verdict, "", rec["range_credit"], S("(not found)", "dim"), "", ""]
    credit = rec["range_credit"]
    return [r.ply, verdict, r.rng,
            "" if credit is None else S(credit, "good" if credit == 1 else "warn" if credit else "bad"),
            r.status, f"{rec.get('ply_reads_right', 0)}/{rec.get('ply_reads', 0)}" if rec["true_ply"] else r.n_reads,
            r.latency]


# ================================================================ single version
def score_text(res, obs, cfg):
    m, ver = res["overall"], res["version"]
    L = ["=" * 100, f" KITTING VERSION SCORE   {ver.label}", "=" * 100, f" folder   {ver.folder}"]
    if ver.notes:
        L.append(f" notes    {ver.notes}")
    L.append(f" scored on {m['ply_n']} labelled roll(s) in {m['videos_scored']} video(s)   "
             f"(this version ran {m['videos']} videos, {m['rolls_reported']} rolls)")
    L.append("")
    for key, title, what in SCORES:
        s = m["scores"][key]["score"]
        L.append(f"   {title:<10} {num(s):>6} / 100   grade {grade(s, cfg)}     {what}"
                 if s is not None else f"   {title:<10}      -           (needs truth)     {what}")
    for key, title, _ in SCORES:
        sc = m["scores"][key]
        if not sc["parts"]:
            continue
        L += ["", f" {title} SCORE = {sc['score']:.1f}", f"   {'part':<48}{'result':>8}{'weight':>8}{'points':>8}   detail"]
        for k, val, w, pts in sc["parts"]:
            L.append(f"   {LABELS[k]:<48}{pct(val):>8}{w:>8.1f}{pts:>8.1f}   {detail(k, m)}")
        if sc.get("left_out"):
            L.append(f"   (no data for: {', '.join(LABELS[k] for k in sc['left_out'])}; weights rescaled to 100)")
    if m["ply_ci"]:
        L += ["", f" ply accuracy {m['ply_right']}/{m['ply_n']} = {pct(m['ply_right'] / m['ply_n'])}, "
                  f"95% interval {pct(m['ply_ci'][0])} to {pct(m['ply_ci'][1])}"]
    groups = {g: gm for g, gm in res["by_group"].items() if gm["ply_n"]}
    if len(groups) > 1:
        L += ["", " BY RECORDING SESSION", f"   {'session':<16}{'videos':>7}{'rolls':>7}{'ply right':>11}{'range exact':>13}{'overall':>9}"]
        for g in sorted(groups, key=natural):
            gm = groups[g]
            L.append(f"   {g:<16}{gm['videos_scored']:>7}{gm['truth_rolls']:>7}{str(gm['ply_right']) + '/' + str(gm['ply_n']):>11}"
                     f"{(str(gm['range_exact']) + '/' + str(gm['range_n'])) if gm['range_n'] else '-':>13}"
                     f"{num(gm['scores']['overall']['score']):>9}")
    L += ["", " OBSERVATIONS"] + _obs_text(obs)
    return "\n".join(L)


def _video_rows(res, cfg):
    rows = [["Video", "Session", "Labelled", "Whole video", "True rolls", "Found", "Missed", "False rolls", "Ply right",
             "Ply labelled", "Range exact", "Range labelled", "Ply reads right", "Ply reads", "fps", "Detection (ms)",
             "OCR (ms)", "Median time to decide (s)", "Overall", "OCR", "Detection", "Run (test_logs)"]]
    groups = {video_group(v) for v in res["runs"]} - {"other"}
    for v in sorted(res["runs"], key=natural):
        m, st = res["by_video"][v], res["runs"][v].stats
        lab = m["truth_rolls"] > 0 or m["false"] > 0
        sc = [m["scores"][k]["score"] for k in ("overall", "ocr", "det")]
        rows.append([v, video_group(v, groups), "yes" if lab else S("no", "dim"), st.get("whole", "")]
                    + ([m["truth_rolls"], m["found"], S(m["missed"], "bad") if m["missed"] else 0,
                        S(m["false"], "warn") if m["false"] else 0, m["ply_right"], m["ply_n"], m["range_exact"],
                        m["range_n"], m["ply_reads_right"], m["ply_reads"]] if lab else [""] * 10)
                    + [m["fps"], m["det_ms"], m["ocr_ms"], m["latency_median"]]
                    + [None if s is None else round(s, 1) for s in sc] + [st.get("log", "")])
    return rows


def score_sheets(res, obs, cfg):
    m, ver = res["overall"], res["version"]
    summary = [[S(f"Kitting version score: {ver.label}", "title")], [f"Folder: {ver.folder}"],
               [f"Scored on {m['ply_n']} labelled roll(s) in {m['videos_scored']} video(s); the version ran "
                f"{m['videos']} videos, {m['rolls_reported']} rolls."], [],
               [S("Score", "header"), S("Points / 100", "header"), S("Grade", "header"), S("What it measures", "header")]]
    for key, title, what in SCORES:
        s = m["scores"][key]["score"]
        summary.append([S(title, "bold"), None if s is None else round(s, 1), grade(s, cfg), what])
    for key, title, _ in SCORES:
        sc = m["scores"][key]
        if not sc["parts"]:
            continue
        summary += [[], [S(f"{title} score breakdown", "bold")],
                    [S(h, "header") for h in ("Part", "Result", "Weight", "Points", "Detail")]]
        summary += [[LABELS[k], pct(val, 1), round(w, 1), round(pts, 1), detail(k, m)] for k, val, w, pts in sc["parts"]]
    summary += [[], [S("Key numbers", "bold")], [S(h, "header") for h in ("Number", "Value")],
                ["True rolls labelled", m["truth_rolls"]], ["Rolls found", m["found"]], ["Rolls missed", m["missed"]],
                ["False rolls", m["false"]], ["Ply right", f"{m['ply_right']}/{m['ply_n']}"],
                ["Ply accuracy 95% interval", "-" if not m["ply_ci"] else f"{pct(m['ply_ci'][0])} to {pct(m['ply_ci'][1])}"],
                ["Picked with a wrong ply", m["ply_wrong_pick"]],
                ["Range exact / close / labelled", f"{m['range_exact']} / {m['range_partial']} / {m['range_n']}"],
                ["Text boxes detected", m["boxes"]], ["OCR reads", m["ocr_reads"]],
                ["fps processed", m["fps"]], ["Detection mean (ms)", m["det_ms"]], ["OCR mean (ms)", m["ocr_ms"]],
                ["Median time to decide (s)", m["latency_median"]]]
    rolls = [["Video", "Roll #", "First seen (s)", "Kind", "True ply", "True range", "Decided ply", "Ply verdict",
              "Decided range", "Range credit", "Status", "Ply reads right / all", "Time to decide (s)", "Bin", "Ply rule",
              "Seq", "Track", "True ply seen in reads", "Ply reads", "Range reads", "Notes"]]
    for rec in res["recs"]:
        r = rec["roll"]
        seen = "" if not rec["true_ply"] or r is None else "yes" if rec.get("ply_reads_right") else S("never", "bad")
        rolls.append([rec["video"], rec["n"], rec["t"], rec["kind"], rec["true_ply"], rec["true_range"]] + _rec_cells(rec)
                     + ([r.bin, r.ply_rule, r.seq, r.track, seen, " ".join(f"{t}x{n}" for t, n in r.ply_reads),
                         " ".join(f"{t}x{n}" for t, n in r.range_reads)] if r else [""] * 7) + [rec["notes"]])
    wrong = {}
    for rec in res["recs"]:
        for text, n in rec.get("wrong_reads", {}).items():
            wrong[(rec["true_ply"], text)] = wrong.get((rec["true_ply"], text), 0) + n
    wr = [["True ply", "Read as", "Times", "Kind"]]
    for (tp, text), n in sorted(wrong.items(), key=lambda kv: -kv[1]):
        kind = "extra characters" if tp in text and len(text) > len(tp) else "digits dropped" if text and text in tp \
            else "empty" if not text else "different"
        wr.append([S(tp, "text"), S(text, "text"), n, kind])
    return [{"name": "Summary", "rows": summary, "widths": [46, 14, 10, 12, 70]},
            {"name": "Observations", "rows": [["Level", "Observation"]] + _obs_rows(obs), "header_rows": 1, "widths": [9, 150]},
            {"name": "Rolls", "rows": rolls, "header_rows": 1, "freeze_cols": 2, "autofilter": True,
             "widths": [28, 7, 10, 11, 9, 13, 10, 15, 13, 8, 16, 11, 9, 11, 19, 6, 7, 10, 45, 45, 30]},
            {"name": "Videos", "rows": _video_rows(res, cfg), "header_rows": 1, "freeze_cols": 1, "autofilter": True,
             "widths": [28, 11, 9, 12] + [8] * 10 + [7, 10, 9, 10, 8, 8, 9, 34]},
            {"name": "Wrong ply reads", "rows": wr, "header_rows": 1, "autofilter": True, "widths": [10, 14, 8, 18]},
            _scoring_rows(cfg)]


# ================================================================ comparison
def _matrix(results, cfg):
    """Rows of (label, [value per version], better: 'high' / 'low' / None, formatter)."""
    ms = [r["overall"] for r in results]
    rows = [(f"{t} score", [m["scores"][k]["score"] for m in ms], "high", lambda x: num(x, 1)) for k, t, _ in SCORES]
    rows.append(("Grade (overall)", [grade(m["scores"]["overall"]["score"], cfg) for m in ms], None, str))
    rows.append(None)
    for k in ("ply_decision", "ply_on_found", "range_decision", "roll_recall", "roll_precision", "label_coverage",
              "ply_read_acc", "range_read_acc"):
        rows.append((LABELS[k], [m["comp"][k] for m in ms], "high", lambda x: pct(x, 1)))
    rows.append(None)
    rows.append(("Ply right / labelled", [f"{m['ply_right']}/{m['ply_n']}" for m in ms], None, str))
    rows.append(("Picked with a wrong ply", [m["ply_wrong_pick"] for m in ms], "low", str))
    rows.append(("Rolls missed", [m["missed"] for m in ms], "low", str))
    rows.append(("False rolls", [m["false"] for m in ms], "low", str))
    rows.append(("Range exact / labelled", [f"{m['range_exact']}/{m['range_n']}" for m in ms], None, str))
    rows.append(None)
    rows.append(("fps processed", [m["fps"] for m in ms], "high", lambda x: num(x, 1)))
    rows.append(("Detection mean (ms)", [m["det_ms"] for m in ms], "low", lambda x: num(x, 1)))
    rows.append(("OCR mean (ms)", [m["ocr_ms"] for m in ms], "low", lambda x: num(x, 1)))
    rows.append(("Median time to decide (s)", [m["latency_median"] for m in ms], "low", lambda x: num(x, 2)))
    rows.append(("Rolls reported", [m["rolls_reported"] for m in ms], None, str))
    return rows


def _best(vals, better):
    """Indexes holding the single best value, or empty when tied / not comparable."""
    nums = [v for v in vals if isinstance(v, (int, float))]
    if not better or len(nums) < 2 or len(set(nums)) == 1:
        return set()
    top = max(nums) if better == "high" else min(nums)
    return {i for i, v in enumerate(vals) if v == top}


def compare_text(results, obs, per_version_obs, cfg, scope):
    names = [r["version"].name for r in results]
    w = max(14, max(len(n) for n in names) + 2)
    L = ["=" * 100, f" KITTING VERSION COMPARISON   {'  vs  '.join(names)}", "=" * 100]
    for r in results:
        L.append(f"   {r['version'].label}")
    L += [f" {scope}", "", f"   {'':<48}" + "".join(f"{n:>{w}}" for n in names)]
    for row in _matrix(results, cfg):
        if row is None:
            L.append("")
            continue
        label, vals, better, fmt = row
        best = _best(vals, better)
        L.append(f"   {label:<48}" + "".join(f"{('* ' if i in best else '') + ('-' if v is None else fmt(v)):>{w}}"
                                             for i, v in enumerate(vals)))
    L += ["", "   * = best of the compared versions", "", " OBSERVATIONS"] + _obs_text(obs)
    for name, vo in per_version_obs.items():
        vo = [x for x in vo if x[0] in ("ALERT", "WARN")]
        if vo:
            L += ["", f" {name}: its own alerts"] + _obs_text(vo)
    return "\n".join(L)


def compare_sheets(results, obs, per_version_obs, pairs, cfg, scope):
    names = [r["version"].name for r in results]
    summary = [[S("Kitting version comparison: " + "  vs  ".join(names), "title")]]
    summary += [[r["version"].label + f"   [{r['version'].folder}]"] for r in results]
    summary += [[scope], [], [S("", "header")] + [S(n, "header") for n in names]]
    for row in _matrix(results, cfg):
        if row is None:
            summary.append([])
            continue
        label, vals, better, fmt = row
        best = _best(vals, better)
        summary.append([label] + [S("-" if v is None else fmt(v), "good") if i in best else ("-" if v is None else fmt(v))
                                  for i, v in enumerate(vals)])
    summary += [[], [S("Green = best of the compared versions. Scores are out of 100; see the Scoring sheet.", "dim")]]

    orows = [["Scope", "Level", "Observation"]] + [["comparison"] + r for r in _obs_rows(obs)]
    for name, vo in per_version_obs.items():
        orows += [[name] + r for r in _obs_rows(vo)]

    h2h = [["Version A", "Version B", "Labelled rolls both ran", "Both right", "Only A right", "Only B right",
            "Both wrong", "Sign test p", "Reliable difference (p <= 0.05)", "Same ply decided", "Rolls both reported"]]
    for a, b, h in pairs:
        h2h.append([a["version"].name, b["version"].name, h["n"], h["both"], len(h["only_a"]), len(h["only_b"]),
                    h["neither"], None if h["p"] is None else round(h["p"], 4),
                    "-" if h["p"] is None else S("yes", "good") if h["p"] <= 0.05 else "no", h["same_ply"], h["same_n"]])

    # every roll side by side: truth rolls first (keyed by video + roll #), then each version's unmatched rolls
    per = [{} for _ in results]
    keys = {}
    for i, r in enumerate(results):
        for rec in r["recs"]:
            k = (rec["video"], rec["n"] if rec["n"] is not None else f"seq {rec['roll'].seq} ({names[i]})")
            per[i][k] = rec
            keys.setdefault(k, rec)
    block = ["ply", "ply verdict", "range", "range credit", "status", "ply reads right / all", "time to decide (s)"]
    rolls = [["Video", "Roll #", "First seen (s)", "True ply", "True range", "Verdicts"]
             + [f"{n} {c}" for n in names for c in block]]
    def order(k):
        return (natural(k[0]), 0, k[1]) if isinstance(k[1], int) else (natural(k[0]), 1, 0)
    for k in sorted(keys, key=order):
        base = keys[k]
        verdicts = [per[i][k]["ply_verdict"] if k in per[i] else None for i in range(len(results))]
        rights = [v == "right" for v in verdicts if v]
        mix = "" if not rights else S("all right", "good") if all(rights) else S("all wrong", "bad") \
            if not any(rights) else S("differ", "warn")
        row = [k[0], k[1], base["t"], base["true_ply"], base["true_range"], mix]
        for i, r in enumerate(results):
            row += _rec_cells(per[i].get(k)) if k[0] in r["runs"] else _rec_cells(None)
        rolls.append(row)

    videos = [["Video", "True rolls"] + [f"{n} {c}" for n in names for c in ("ply right", "missed", "false", "overall", "OCR (ms)")]]
    all_videos = sorted({v for r in results for v in r["runs"]}, key=natural)
    for v in all_videos:
        ms = [r["by_video"].get(v) for r in results]
        row = [v, max((m["truth_rolls"] for m in ms if m), default=0)]
        for m in ms:
            if m is None:
                row += [S("(not run)", "dim"), "", "", "", ""]
                continue
            s = m["scores"]["overall"]["score"]
            row += [f"{m['ply_right']}/{m['ply_n']}" if m["ply_n"] else "", m["missed"] or "", m["false"] or "",
                    None if s is None else round(s, 1), m["ocr_ms"]]
        videos.append(row)
    return [{"name": "Summary", "rows": summary, "widths": [50] + [18] * len(names)},
            {"name": "Observations", "rows": orows, "header_rows": 1, "widths": [16, 9, 150]},
            {"name": "Head to head", "rows": h2h, "header_rows": 1, "widths": [16, 16, 12, 10, 11, 11, 10, 10, 16, 12, 12]},
            {"name": "Rolls", "rows": rolls, "header_rows": 1, "freeze_cols": 2, "autofilter": True,
             "widths": [28, 14, 10, 9, 13, 10] + [9, 15, 13, 8, 16, 11, 9] * len(names)},
            {"name": "Videos", "rows": videos, "header_rows": 1, "freeze_cols": 1, "autofilter": True,
             "widths": [28, 9] + [10, 8, 7, 9, 9] * len(names)},
            _scoring_rows(cfg)]
