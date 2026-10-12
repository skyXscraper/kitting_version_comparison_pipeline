"""Snapshots of what a run detected, cut out of its video, so the true ply / range can be read off the picture.

Nothing is detected again: a run of the RTracker&OCR script already wrote every box it found (frame, rotated box,
roll track) to test_logs/<run>/detections.csv, and run_info.json names the video. For each roll track this module
takes the sharpest frames that have a text box on the roll and returns the straightened label crops plus one
picture of the whole roll, as in-memory JPEG data URLs. Nothing is written to disk.

Needs opencv-python and numpy (already there on a machine that runs the tracker).
"""
import base64
import csv
import json
import math
from pathlib import Path

CANDIDATES = 8      # frames looked at per roll (the ones with the most confident text boxes)
KEEP = 4            # sharpest of those shown
MAX_CROPS = 8       # label crops shown per roll
CROP_HEIGHT = 110   # label crops are enlarged to about this many pixels high
ROLL_WIDTH = 760    # the roll picture is enlarged (up to 2x) to at most this width


class SnapshotError(Exception):
    pass


class _Video:
    """Frames of a video by index; walks forward when the next one is near, seeks otherwise."""
    def __init__(self, cv2, path):
        self.cv2, self.cap, self.pos = cv2, cv2.VideoCapture(str(path)), 0
        if not self.cap.isOpened():
            raise SnapshotError(f"the video {path} could not be opened")
        self.fps = self.cap.get(cv2.CAP_PROP_FPS)

    def frame(self, idx):
        if idx < self.pos or idx - self.pos > 120:
            self.cap.set(self.cv2.CAP_PROP_POS_FRAMES, idx)
            self.pos = idx
        while self.pos < idx:
            if not self.cap.grab():
                return None
            self.pos += 1
        ok, img = self.cap.read()
        self.pos += 1
        return img if ok else None

    def close(self):
        self.cap.release()


def _cut(cv2, np, frame, rect, pad):
    """A rotated box (cx, cy, w, h, theta in radians) cut out straight, the same way the run cut its OCR crops."""
    cx, cy, w, h, t = rect
    W, H = max(8, int(w + 2 * pad)), max(8, int(h + 2 * pad))
    ux, uy, vx, vy = math.cos(t), math.sin(t), -math.sin(t), math.cos(t)
    M = np.float32([[ux, vx, cx - W / 2 * ux - H / 2 * vx], [uy, vy, cy - W / 2 * uy - H / 2 * vy]])
    return cv2.warpAffine(frame, M, (W, H), flags=cv2.INTER_LINEAR | cv2.WARP_INVERSE_MAP,
                          borderMode=cv2.BORDER_REPLICATE)


def _url(cv2, img, scale):
    if scale != 1:
        img = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC if scale > 1 else cv2.INTER_AREA)
    ok, buf = cv2.imencode(".jpg", img, [cv2.IMWRITE_JPEG_QUALITY, 88])
    return "data:image/jpeg;base64," + base64.b64encode(buf.tobytes()).decode("ascii") if ok else None


def run_snapshots(root, log, tracks):
    """Snapshots of the given roll tracks of one run -> {track: {"roll": url or None, "crops": [{img, t, conf}]}}."""
    try:
        import cv2
        import numpy as np
    except ImportError:
        raise SnapshotError("snapshots need OpenCV:  pip install opencv-python")
    if not root:
        raise SnapshotError("set \"runs_root\" in config.json to the folder that holds test_logs and the videos")
    root = Path(root)
    run = root / "test_logs" / log
    if not (run / "detections.csv").is_file() or not (run / "run_info.json").is_file():
        raise SnapshotError(f"the log folder of the run is not there ({run}); it holds the boxes the run detected")
    info = json.loads((run / "run_info.json").read_text(encoding="utf-8"))
    source = Path(str(info.get("source", "")).replace("\\", "/"))
    video_path = source if source.is_absolute() else root / source
    if not video_path.is_file():
        raise SnapshotError(f"the video of the run is not there ({video_path})")

    wanted = {str(t) for t in tracks if str(t)}
    texts, rolls = {t: {} for t in wanted}, {t: {} for t in wanted}       # track -> frame -> boxes
    with open(run / "detections.csv", newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            track = (row.get("roll_track_id") or "").strip()
            if track not in wanted:
                continue
            try:
                frame, conf = int(row["frame"]), float(row["conf"])
                rect = tuple(float(row[k]) for k in ("cx", "cy", "w", "h", "theta"))
            except (KeyError, ValueError):
                continue
            if row.get("class") == "roll":
                rolls[track][frame] = rect
            else:
                texts[track].setdefault(frame, []).append((conf, rect))

    need = {}       # frame -> tracks that want it
    for t in wanted:
        frames = sorted(texts[t], key=lambda fr: -max(c for c, _ in texts[t][fr]))[:CANDIDATES]
        if not frames and rolls[t]:     # a roll with no text box at all: show the roll halfway through its track
            frames = [sorted(rolls[t])[len(rolls[t]) // 2]]
        for fr in frames:
            need.setdefault(fr, []).append(t)

    video = _Video(cv2, video_path)
    rate = float(info.get("processing_fps") or 15)
    step = (video.fps or rate) / rate       # the run read every step-th frame of the file
    seen = {t: [] for t in wanted}          # track -> [(sharpness, frame, [(conf, crop)], roll picture)]
    try:
        for fr in sorted(need):
            img = video.frame(math.ceil((fr - 1) * step - 1e-9))
            if img is None:
                continue
            for t in need[fr]:
                crops, sharp = [], 0.0
                for conf, rect in sorted(texts[t].get(fr, []), key=lambda b: (b[1][1], b[1][0])):
                    tight = cv2.cvtColor(_cut(cv2, np, img, rect, 0), cv2.COLOR_BGR2GRAY)
                    sharp = max(sharp, float(cv2.Laplacian(tight, cv2.CV_64F).var()))
                    crops.append((conf, _cut(cv2, np, img, rect, 6 + 0.3 * min(rect[2], rect[3]))))
                roll = _cut(cv2, np, img, rolls[t][fr], 12) if fr in rolls[t] else None
                seen[t].append((sharp, fr, crops, roll))
    finally:
        video.close()

    out = {}
    for t, shots in seen.items():
        if not shots:
            continue
        best = sorted(sorted(shots, key=lambda s: -s[0])[:KEEP], key=lambda s: s[1])
        sharpest = max(best, key=lambda s: s[0])
        crops = []
        for _, fr, cs, _ in best:
            for conf, c in cs:
                url = _url(cv2, c, min(4.0, max(1.0, CROP_HEIGHT / c.shape[0])))
                if url:
                    crops.append({"img": url, "t": round((fr - 1) / rate, 2), "conf": round(conf, 2)})
        roll = sharpest[3]
        out[t] = {"crops": crops[:MAX_CROPS], "t": round((sharpest[1] - 1) / rate, 2),
                  "roll": _url(cv2, roll, min(2.0, ROLL_WIDTH / roll.shape[1])) if roll is not None else None}
    return out
