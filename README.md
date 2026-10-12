# Kitting version comparison pipeline

Scores and compares versions of the kitting OCR (recognizer) and detection models from the Excel result
workbooks that the RTracker&OCR log scripts write. You give the true ply and range of each roll once; every
version is then scored against the same truth.

Needs only Python 3 (no extra packages). Run everything from this folder.

## Quick start

```
python kvc.py add "C:\Users\Tanvi\Documents\RTracker&OCR\rec_v5_logs_v1"     1. register a version
python kvc.py truth                                                          2. make / update the truth sheet
        ... open truth\truth.xlsx, fill the blue cells, save, close Excel ...
python kvc.py score rec_v5                                                   3. score one version
python kvc.py compare rec_v4 rec_v5                                          4. compare versions
```

`kvc.bat` is a shortcut, so `kvc score rec_v5` works too.

## Web page

```
python kvc_web.py          (or double-click kvc_web.bat)
```

Opens `http://127.0.0.1:8000` in the browser; keep the console window open while you use it (`--port N` for
another port). One page, nothing to register:

1. Drop the result sheets to compare: one `.xlsx` per model version, all of the **same video** (the sheets
   usually have the same file name in different folders, so choose again to add the next one). A side-by-side
   sheet counts as two versions. One sheet alone gives that version's score.
2. Press **Compare**. Each version is named after the model written in its sheet (`rec v5` -> `rec_v5`; a second
   sheet of the same model becomes `rec_v5_2`).
3. Type the true ply / range of the video's rolls in the blue cells and press **Save truth & score again**.

The scores, side-by-side table, observations and per-roll tables are shown on the page; *Download as Excel* gives
the same report as a workbook. The uploaded sheets are only read: they are not kept, nothing is added to
`versions.json` and no report folder is written. The one thing kept is the truth you type (`truth	ruth.xlsx`,
shared with the command line), so a video's true values are entered once.

Speed parts of the score (frame rate, detection / OCR time) come from the summary workbook; drop it in together
with the sheets to get them, otherwise the score is accuracy only and the weights are rescaled.

The page is served to this computer only.

## Commands

| Command | What it does |
|---|---|
| `add <folder>` | Registers a folder of result workbooks as a version and adds its rolls to the truth sheet. A side-by-side folder (two models per workbook) becomes two versions. Options: `--name`, `--names a,b` (side-by-side), `--det`, `--rec`, `--notes "..."`, `--replace`, `--append` (add this folder's videos to a version that already exists). |
| `truth` | Creates / updates `truth\truth.xlsx` with every roll of every registered version. Never removes what you typed; the previous file is copied to `truth\backups`. |
| `score <version> [...]` | Scores one version: three scores, the breakdown, per-session table, observations. |
| `compare <v1> <v2> [...]` | Compares versions **on the videos they all ran** (like for like). No names = all versions. `--all-videos` scores each on its own videos instead. |
| `rank` | Leaderboard of every registered version, each on its own videos. |
| `list` | Registered versions, how many of their rolls are labelled, and their scores. |
| `remove <version>` | Forgets a version (workbooks and truth are untouched). |
| `config` | Writes `config.json` so the weights and thresholds can be edited. |

`score`, `compare` and `rank` accept `--videos "oct_3_*,19_part2_*"` to limit the videos (wildcards allowed).

Every `score` / `compare` / `rank` prints the report and saves it to `reports\<date>_<kind>_<names>\`:
`report.txt` (what was printed) and `report.xlsx` (Summary, Observations, Rolls, Videos, plus Wrong ply reads
for a single version and Head to head for a comparison, and the Scoring settings used).

## Input: what a version folder must contain

One `.xlsx` per video with the sheets **Frames**, **Rolls** and **Run info**, and optionally a summary workbook
(sheets **Summary** and **Rolls**) that supplies frame counts and speeds. These are the folders written by:

- `rec_v5_logs_v1\make_logs_xlsx_v1.py` — one model per workbook (one version)
- `compare_rec_v3_vs_v3_2_v1\make_compare_xlsx_v1.py` — two models side by side (two versions)

Columns are found by their header text, so adding or reordering columns does not break the loader; renaming
the headers the pipeline relies on (`Seq`, `Track`, `Status`, `Ply`, `Range`, `Decided at (s)`,
`Decided after first seen (s)`, `Field`, `Read`, ...) does. The video name is the workbook file name.

The detector and recognizer names are read from the Run info sheet; override them with `--det` / `--rec`.

## The truth sheet

`truth\truth.xlsx` has one row per physical roll: `Video | Roll # | First seen (s) | Real roll? | True ply |
True range | Notes`, followed by what each version decided (for reference only).

- Fill **True ply** and **True range** for **every roll of a video**. Leave True range empty if the roll has
  no range label; range is then not scored for that roll.
- Set **Real roll? = no** for a row that is not a roll (or is the same roll tracked a second time).
- If no version found a roll, **add a row** at the bottom: video name, roughly when it first appears, true values.
- A video is scored as soon as one of its rolls is filled. In a scored video, a roll a version reports that is
  not in the sheet is a **false roll**, and a listed roll a version never reports is **missed**.
- Videos with nothing filled are not scored yet (their speed numbers are still used).

A version's roll is matched to a truth row by the time the roll is first seen (within 6 s), not by its seq or
track id, so versions with different detectors still line up with the same truth rows.

True values already typed into the result workbooks (True ply / True range columns) are copied into the
truth sheet automatically.

## Scoring

Three scores, each 0-100 (grades: A >= 90, B >= 75, C >= 60, D >= 40, else E).

**OVERALL** — the whole system, end to end.

| Part | Weight | Meaning |
|---|---|---|
| Ply decision (of all true rolls) | 45 | Rolls whose decided ply equals the true ply. A missed roll scores 0. A roll *picked* with a wrong ply also loses 0.25 of a roll, because it goes to a wrong bin. |
| Range decided right | 15 | Exact = 1, one character off = 0.5, two off = 0.25, else 0. |
| Rolls found without false rolls (F1) | 15 | Balance of true rolls found and reported rolls that are real. |
| Single ply reads right | 10 | Share of individual crop reads equal to the true ply, before any voting. |
| Single range reads right | 5 | Same for range reads (with the partial credit above). |
| Time to decide a roll | 5 | Median seconds from first seen to decided: full marks <= 2 s, zero >= 10 s. |
| Frame rate kept | 5 | fps processed / 15 fps target. |

**OCR** — the recognizer only, judged on the rolls the detector found (so a detector miss does not count
against it): ply decision 55, range 20, single ply reads 15, single range reads 5, OCR time per crop 5.

**DETECTION** — the detector / tracker only: true rolls found 40, reported rolls that are real 25, found
rolls with at least 3 label reads 20, detection time per frame 15.

A part with no data (for example no true ranges yet) is left out and the other weights are rescaled to 100;
the report says so. All weights and thresholds are in `config.json` (`python kvc.py config`).

### Reading the result

- **Sample size matters.** The report prints the 95% interval of the ply accuracy and warns below 20
  labelled rolls. A comparison also runs a sign test on the rolls where two versions differ and says whether
  the difference is statistically reliable.
- **Compare like for like.** `compare` uses only the videos every version ran. `rank` and `list` do not.
- **Speed is machine-dependent.** Times come from the run that produced the workbooks (laptop GPU/CPU), not
  from the Hailo / Pi. Compare speed only between runs made on the same machine under the same load.

## Observations

Each report lists what looks abnormal, most serious first:

- **ALERT** — rolls picked with a wrong ply (wrong bin), true rolls never reported, no truth yet.
- **WARN** — true ply was read but outvoted (decision-rule problem) vs. never read (recognizer problem),
  false rolls, a weak recording session, extra / dropped characters in reads, wrong orientation chosen,
  confidently wrong reads, malformed decided ranges, frame rate drops, slow decisions, OCR queue backlog,
  partial runs, problems in the truth sheet.
- **INFO** — most common wrong reads, roll outcome mix, coverage of the truth, and for comparisons the
  ranking, agreement between versions, and the rolls where they differ.

## Files

```
kvc.py / kvc.bat     entry point (command line)
kvc_web.py / .bat    entry point (web page)
kvc\                 the code (xlsx_io, loader, truth, scoring, observations, report, cli, webapp)
versions.json        registered versions (written by add / remove)
config.json          scoring weights and thresholds
truth\truth.xlsx     the truth sheet;  truth\backups\  earlier copies
reports\             one folder per run
```
