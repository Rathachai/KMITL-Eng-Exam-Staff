# AGENTS.md

Guidance for AI coding agents (and humans) working on **KMITL-Eng-Exam-Staff**: a tool that assigns exam invigilators ("ผู้คุมสอบ") for the Faculty of Engineering, KMITL, and produces the Excel reports staff need. User-facing docs are in `README.md` (Thai); this file is for contributors.

## Quick start

```bash
pip install -r requirements.txt          # pandas, numpy, openpyxl, flask
python3 gen-exam-room.py                 # assign staff + build all reports  (Windows: py)
python3 gen-exam-room.py report-only     # rebuild reports from the latest run, keep the assignment
python3 gen-exam-room.py report-only 2026-10-17_14-26   # ...from a specific run
python3 app.py                           # web UI on http://localhost:8089 (no login)
python3 anonymize.py                     # write de-identified copies of exam-data/ to exam-data-anon/
```

There is no test suite. Verify a change by running `gen-exam-room.py` and reading the `สรุปตรวจสอบ` sheet of `for-cheking.xlsx` (10 pass/fail checks), plus the console lines `master match: N / M` and `unmatched: []`.

## Layout

| File | Role |
|---|---|
| `config.py` | The only file end users edit: paths, `EXAM_MONTHS`, `EXTYPE`, payment settings |
| `common.py` | Shared constants/functions, `process_exam_report`, run-folder + state helpers (`save_state`, `load_state`, `out_path`, `find_run_dir`) |
| `exam-code.py` | The scheduling engine: read `exam-data/`, assign invigilators greedily, `save_state(...)` |
| `gen-report-{subjects,room,staff,building,pickup,check,payment}.py` | One report each; read the pickled state, write one `.xlsx` |
| `gen-exam-room.py` | CLI orchestrator (full run, or `report-only [run]`) |
| `app.py` | Flask web UI: upload → queue → progress → zip download → regenerate |
| `anonymize.py` | Replaces personal names with codes like `คอม-อจ-001` |
| `exam-data/` | Input: `Exam-Master-Table.xlsx` + the exam schedule `ร*.xlsx` |
| `reports/<yyyy-mm-dd_HH-MM>/` | One folder per run; flat (no sub-folders); git-ignored |

Data flow: `exam-data/` → `exam-code.py` → `reports/<run>/exam-state.pkl` → each `gen-report-*.py` → `reports/<run>/*.xlsx`. State keys: `dfX` (exam rows with `staff`, `staff_index`, `reason`), `dfST` (staff master with `load`, `free`, `ex_dt`), `dfslot`, `dtlist`, `extype`, `dfoff` (leave dates).

## Conventions

- **Paths**: always `os.path.join` / `out_path()`; never hard-code `/` or absolute paths (Windows must keep working). Reports are written only via `out_path('file.xlsx')`.
- **Run folders**: `yyyy-mm-dd_HH-MM`, collision suffix `-2`, `-3`. `gen-report-*.py` take the run name as `argv[1]`, defaulting to the newest run that has `exam-state.pkl`.
- **Settings** end users may change go in `config.py`, not in the scripts.
- **Thai text and fonts**: Excel output uses `TH Sarabun New`; keep Thai strings in UTF-8; sentence case, no terminal punctuation on labels.
- **pandas**: pinned `>=1.4,<3`. Code relies on pandas-1.x behaviour (e.g. `fillna(inplace=True)` on a column). `for-cheking.xlsx` is written with `engine='openpyxl'` (the default writer may be xlsxwriter on some machines, which lacks the formatting API used).
- **Randomness**: the greedy engine shuffles; two runs differ. Compare engines/runs through the check sheet, not by diffing output.
- **Python**: written against 3.9 (`list[tuple]` annotations, no `match`).

## Gotchas (learned the hard way)

1. **Name matching is string equality.** Staff are keyed by `first last` lower-cased with single spaces. Anything compared against them (`staff_cannot`, `fix_staff`, schedule teachers) must be normalised the same way, otherwise it silently fails to match — this once let three people be scheduled on their leave days. `exam-code.py` prints a `WARNING` for `staff_cannot` names missing from the master; read it.
2. **Role strings**: `dfST['role']` contains `'อาจารย์ (ประเภทพิเศษ) '` with a trailing space, and `dfslot.index` does not. `main_role` is only set for อาจารย์/ผู้บริหาร/หัวหน้าภาค (teacher) and เจ้าหน้าที่/หัวหน้างาน (support); special-type lecturers have `NaN` and load 0.
3. **"เพิ่มแถว" rows** (`Remark` set) are extra rooms assigned after the normal rows; they are part of `dfX` in the final state.
4. **Executive quota** (`ผู้บริหาร`) is `min(number of rows they teach, role quota)` and executives are only ever assigned to their own subjects.
5. **Support-staff rule**: each department keeps at least one support person off duty in every time slot (`TTB`).
6. `anon-mapping.csv` (real name → code) and `reports/` are git-ignored on purpose.
7. A Windows/Excel lock file `~$*.xlsx` in a run folder just means the workbook is open; it is not output.

## Privacy

`exam-data/` historically held real names of staff (personal data). Do not commit new real-name data. Use `python3 anonymize.py` and publish only `exam-data-anon/` content; never commit `anon-mapping.csv`. Names remain in the git history of earlier commits; making the repository private or rewriting history is a separate, manual decision (do not force-push on your own).

## Working with the environment

- Do **not** `pip install` heavy or unrelated packages into the user's global/anaconda Python; one such install (OR-Tools) upgraded numpy/pandas/protobuf and broke other tools (see ADR-007). Use a virtual environment for experiments.
- Don't delete run folders or `exam-data/` files without being asked.
- Commit only when asked.

---

# Architecture Decision Records

Format: Context → Decision → Consequences. Newest decisions are appended; superseded ones stay with their status.

## ADR-001 — Split the single script into config / common / engine / one script per report
*Status: accepted · 2026-10-17*

**Context.** `exam-code.py` was one 2,000-line notebook-style file mixing scheduling (lines 1–996) with six Excel exports. It only ran cell-by-cell (functions were defined after first use) and hard-coded a Google Drive path.
**Decision.** Keep `exam-code.py` for scheduling only; one `gen-report-<name>.py` per report; `config.py` for user-editable settings; `common.py` for shared constants/functions; hand over data with a pickled state.
**Consequences.** Reports can be regenerated without re-running the (random) assignment; each report is small and independently editable; forward-reference bugs disappeared. The pickle couples scripts to the state's key names.

## ADR-002 — One output folder per run, flat, with a timestamp name
*Status: accepted · 2026-10-17*

**Context.** Re-running overwrote earlier results, and the greedy assignment is random, so an earlier good result could not be recovered.
**Decision.** Every full run creates `reports/yyyy-mm-dd_HH-MM/` containing `exam-state.pkl` and all reports; no deeper sub-folders. `report-only` reuses a run folder.
**Consequences.** History is kept and comparable; the web UI treats each folder as a record. Disk usage grows; folders are never auto-deleted.

## ADR-003 — Detect mid/final from exam months
*Status: accepted · 2026-10-18*

**Context.** `extype` was edited by hand and was easy to forget; it changes quotas and report titles.
**Decision.** `config.EXAM_MONTHS = {'mid': [1,2,8,9], 'final': [3,4,10,11]}`; `detect_extype` uses the majority month of the schedule's dates. `config.EXTYPE` overrides it. An unknown month raises an error.
**Consequences.** No manual switch in the normal case; an unusual schedule (e.g. June) fails loudly instead of silently picking a type.

## ADR-004 — Web UI as a local Flask app, no login, queued subprocesses
*Status: accepted · 2026-10-18*

**Context.** Non-technical staff need to upload files and download reports; the existing code is Python/pandas. A desktop app (PyInstaller/Electron) would add per-OS packaging and heavy binaries.
**Decision.** `app.py` (Flask, port 8089, no login). Each run is a record backed by its folder plus `status.json`; a single worker thread runs jobs one at a time as subprocesses of the existing scripts; uploaded files are stored in the run folder as `input-*.xlsx` and passed via `EXAM_MASTER_PATH` / `EXAM_SCHEDULE_FILES`; the zip excludes the pickle and inputs. Binds to `127.0.0.1` unless `EXAM_HOST` is set.
**Consequences.** Almost no new logic; works on macOS and Windows. With no authentication, exposing it on a LAN exposes staff data, so it is local-only by default. Jobs are serial (the assignment is random and writes files), and an interrupted run is marked as such on restart.

## ADR-005 — De-identify data with role/department codes, not hashes
*Status: accepted · 2026-10-17*

**Context.** The Excel inputs contain real staff names and were pushed to GitHub. A hash would be opaque and would break the Thai-name conventions the code relies on.
**Decision.** `anonymize.py` rewrites every cell using a name → code map (`<dept>-<role>-<nnn>`, e.g. `คอม-อจ-001`; people not in the master become `นอก-อจ-nnn`), keeps titles, lowers ASCII department codes (the code matches names case-insensitively), clears workbook author metadata, and writes to `exam-data-anon/`. The mapping goes to a git-ignored CSV.
**Consequences.** The full pipeline still runs on anonymised data with identical match rates. It does not remove indirect identifiers (department + role + subjects) and does not clean git history.

## ADR-006 — Normalise leave-date keys (`staff_cannot`)
*Status: accepted · 2026-10-17*

**Context.** Leave dates were matched as the raw string `fullname + ' ' + cannot`. Trailing spaces, doubled spaces and capital letters in the master meant 20 leave entries (3 people) never matched, and those people were scheduled on their leave days.
**Decision.** Build the lookup from lower-cased, single-spaced names and stripped slots; warn about names absent from the master. Report 9 in the check workbook verifies the result independently.
**Consequences.** Leave days are respected. The warning surfaces data-entry mistakes (one name in the current master is still unmatched).

## ADR-007 — Evaluated ASP (clingo) and OR-Tools; kept the greedy engine
*Status: accepted (alternatives rejected) · 2026-10-18 – 2026-11-01*

**Context.** The greedy engine is order-dependent and random, and the check workbook shows several unmet goals (own-subject coverage, executive quotas, per-person quotas). Declarative solvers looked promising.
**Decision.**
- *ASP/clingo* was implemented end to end (facts from the same data, hard rules for slots/quotas/leave/language/executives/room coverage, weighted preferences, results written in the same state format). In 120 s it reached roughly the same quality as greedy (own-subject ≈45 %, own-department ≈64 %) — it found a good first model with domain heuristics but did not improve afterwards — and it excluded the extra "เพิ่มแถว" rows. It was removed because it brought no clear gain.
- *OR-Tools (CP-SAT)* was dropped before any model code was written: installing it into the anaconda base upgraded numpy to 2.x, pandas to 2.x and protobuf to 6.x and broke other installed tooling. numpy/pandas/protobuf were restored by hand (protobuf to a best-guess version).
**Consequences.** One engine to maintain. If a better optimiser is revisited, do it in a dedicated virtual environment, warm-start from the greedy result, and compare through `สรุปตรวจสอบ`. Capacity is tight (≈1,498 slots for 1,440 rows including extras), so "meet every quota exactly" cannot be a goal once extra rows are excluded.

## ADR-008 — Independent check workbook with explicit pass/fail criteria
*Status: accepted · 2026-10-18*

**Context.** There was no way to judge a schedule other than eyeballing it.
**Decision.** `gen-report-check.py` adds a summary sheet plus detail sheets `ck1`–`ck10` to `for-cheking.xlsx`: every row assigned; every room has a teacher/executive; Inter subjects have a lang=en person or teacher; own-subject coverage per subject (≥ 50 %, `MAJORITY`); own-department coverage; no double booking; distribution by `group_dept` and day; over/under quota; leave conflicts; executive quotas. It recomputes everything from the state rather than trusting the engine's own bookkeeping.
**Consequences.** Schedules are comparable across runs/engines. Several checks currently fail by design of the data (e.g. executives, under-quota) and are signals, not bugs; thresholds are judgement calls and live at the top of the script.

## ADR-009 — Keep all user-visible reports in Excel with Thai print layouts
*Status: accepted · 2026-10-17*

**Context.** Staff print these sheets (A4 landscape, TH Sarabun New) and sign them (exam paper pick-up/hand-over, payment evidence).
**Decision.** Reports are `.xlsx` built with openpyxl and explicit print setup; `for-pickup.xlsx` (รับ-ส่งข้อสอบ) was added as its own report script.
**Consequences.** Output is print-ready but layout code is long and repetitive across reports; shared helpers could be factored into `common.py` later.
