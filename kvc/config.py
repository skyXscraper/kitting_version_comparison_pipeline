"""Scoring weights and thresholds. Defaults live here; config.json (next to kvc.py) overrides any of them."""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / "config.json"
VERSIONS_PATH = ROOT / "versions.json"
TRUTH_PATH = ROOT / "truth" / "truth.xlsx"
REPORTS_DIR = ROOT / "reports"

DEFAULTS = {
    # ---- the three scores: component -> weight (points out of 100). A component with no data is left out and
    # the remaining weights are rescaled to 100.
    "overall_weights": {"ply_decision": 45, "range_decision": 15, "roll_f1": 15, "ply_read_acc": 10,
                        "range_read_acc": 5, "latency": 5, "throughput": 5},
    "ocr_weights": {"ply_on_found": 55, "range_on_found": 20, "ply_read_acc": 15, "range_read_acc": 5,
                    "ocr_speed": 5},
    "det_weights": {"roll_recall": 40, "roll_precision": 25, "label_coverage": 20, "det_speed": 15},
    # ---- ply: a roll PICKED with a ply that is not even close goes to a wrong bin, so it costs this much of a
    # roll on top of getting no credit (0 = a wrong ply is no worse than an unreadable one).
    "wrong_pick_penalty": 0.25,
    # ---- close is not entirely wrong: exact = 1.0, otherwise credit by the number of characters to change
    # (edit distance: a digit replaced, added or dropped). Ply: 1 off still counts; range: up to 3 off.
    "ply_partial_credit": {"1": 0.5},
    "range_partial_credit": {"1": 0.75, "2": 0.5, "3": 0.25},
    # ---- matching a version's rolls to the truth rolls: first-seen times at most this far apart (s).
    "roll_match_tolerance_s": 6.0,
    # ---- a found roll counts as "labels covered" when it got at least this many OCR reads.
    "min_reads_for_coverage": 3,
    # ---- speed: full marks at "good", zero at "bad", linear in between.
    "target_fps": 15,
    "latency_good_s": 2.0, "latency_bad_s": 10.0,       # decided after first seen (median)
    "det_ms_good": 35.0, "det_ms_bad": 100.0,           # detection mean per frame
    "ocr_ms_good": 40.0, "ocr_ms_bad": 150.0,           # OCR mean per crop
    # ---- observations
    "min_rolls_reliable": 20,                           # fewer labelled rolls than this -> score flagged as shaky
    "high_conf": 0.9,                                   # a wrong read at or above this score is "confidently wrong"
    "grades": {"A": 90, "B": 75, "C": 60, "D": 40},
    # ---- web page snapshots: the folder of the tracker runs, the one that holds test_logs\ and the videos. The
    # pictures of each roll's labels are cut from the video with the boxes the run logged there.
    "runs_root": "",
}


def load_config():
    cfg = json.loads(json.dumps(DEFAULTS))
    if CONFIG_PATH.exists():
        user = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        for k, v in user.items():
            if k.startswith("_"):
                continue
            if k not in cfg:
                raise SystemExit(f"config.json: unknown setting '{k}'")
            cfg[k] = v
    return cfg


def write_default_config():
    if not CONFIG_PATH.exists():
        doc = {"_help": "Edit any value. Delete a line to go back to its default. See README.md > Scoring."}
        doc.update(DEFAULTS)
        CONFIG_PATH.write_text(json.dumps(doc, indent=2), encoding="utf-8")
