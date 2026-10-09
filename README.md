# PM Qualification Activity Tracker

A local Python tool that turns IBM Maximo exports into annual PM (preventive maintenance) compliance workbooks. It shows, for every technician and qualification, whether each annual cycle has a completed PM, is coming due, or needs requalification.

Run it with a double-click. A normal monthly run needs no coding and no terminal.

## Contents

- [How it works](#how-it-works)
- [Quick start](#quick-start)
- [Input files](#input-files)
- [Output workbooks](#output-workbooks)
- [Business rules](#business-rules)
- [Configuration](#configuration)
- [Project structure](#project-structure)
- [Diagnostic scripts](#diagnostic-scripts)
- [Privacy](#privacy)
- [Troubleshooting](#troubleshooting)

## How it works

```text
Qualification reports ─┐
Task history reports ──┼─► match by SSO + qualification ─► annual cycles ─► Excel workbooks
Animation roster ──────┤                                       │
Line Schedule ─────────┘                              Data/cycle_history.csv
                                                      (permanent evidence)
```

1. Reads one qualification report per tracked qualification and takes each technician's original qualification date.
2. Builds annual cycles from that date, up to today.
3. Reads every worksheet in every task report and keeps only `Status = COMP` rows.
4. Matches tasks to technicians by SSO and qualification code. Names are display-only.
5. Assigns each task to the one cycle that contains its completion date.
6. Restores completed cycles from saved history if a task has dropped out of the rolling one-year export.
7. Adds each technician's shift from the Line Schedule.
8. Writes one workbook per venue, plus one for Animation.

## Quick start

**Requirements:** Windows, Python 3.14, pandas 3.0.6, openpyxl 3.1.5.

### One-time setup

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### Monthly run

1. Download the current Maximo reports as `.xlsx`.
2. Put qualification reports in `Input/Qualification/`.
3. Put task history reports in `Input/Tasks/`.
4. Close every Excel workbook, including the previous outputs.
5. Double-click **`Run_Qualification_Tracker.bat`**.

The runner works in two steps:

1. Moves the previous workbooks to `Archive/<timestamp>/Output/`. If any file is locked, it stops and puts everything back.
2. Generates the new workbooks and opens `Output/`.

Your input files are never modified.

To run the generator directly:

```powershell
.\.venv\Scripts\python.exe Scripts\report_generator.py
```

## Input files

| Input | Location | Notes |
|---|---|---|
| Qualification reports | `Input/Qualification/` | One file per qualification. The filename must contain exactly one code, e.g. `CUMOD1 Qualification.xlsx`. The sheet is `List of Labor`. |
| Task history reports | `Input/Tasks/` | Any filename. Can hold one or several qualifications, and a qualification can span several files. Every worksheet is read. |
| Animation roster | `Input/Manual roster/Animation Roster.xlsx` | Columns `Name` and `ID` (the SSO). |
| Line Schedule | Folder set in `Config/settings.json` | Exactly one `.xlsx` in that folder. |

**Qualification report columns:** `Labor`, `Name`, `Effective Date`, `Status`.

**Task report columns:** `Task Num`, `Status`, `Changed Date`, `Change By`, `Name`, `Qualification`. The task description is read from the first column.

**Line Schedule sheets:** `1st Shift`, `2nd Shift`, `3rd Shift Unchained`, `3rd Shift Wolf`. Data starts on row 4. Column A is the name and column D is the SSO.

Qualification is never inferred from Work Location. Files with `MOD0` in the name are ignored.

## Output workbooks

Written to `Output/`:

- `Wolf Qualification Tracker.xlsx`
- `Darkmoor Qualification Tracker.xlsx`
- `Unchained Qualification Tracker.xlsx`
- `Animation Qualification Tracker.xlsx`

Animation roster technicians appear **only** in the Animation workbook, which evaluates UCMOD4 only. They are removed from the venue workbooks.

| Sheet | Contents |
|---|---|
| Dashboard | Counts per cycle result and the processing date |
| PM Status | One row per technician, qualification and cycle |
| Exceptions | Task Data Missing cycles, non-ACTIVE qualifications, and technicians whose shift is not found or in conflict |
| Roster Exceptions | Animation workbook only: blank, duplicate or unmatched roster SSOs and name mismatches |
| Run Summary | Record counts, qualifications and disciplines |

**PM Status columns:** Venue, Shift, Discipline, Qualification, SSO, Technician Name, Qualification Date, Qualification Status, Cycle Number, Cycle Start Date, Cycle Due Date, Task Data Available, PM Completion Date, Task Description, Task Number, Days Remaining, Cycle Result.

**Shift** is `1st`, `2nd` or `3rd`. It is `Shift Not Found` if the technician is not on the schedule, and `Shift Conflict` if they appear on more than one shift.

## Business rules

### Valid PM

A task counts only if all of these are true:

- `Status` is `COMP`.
- The task's `Change By` SSO matches the qualification's `Labor` SSO.
- The task's qualification code matches.
- The completion date falls inside the cycle, inclusive of both boundaries.

Any single qualifying task counts. A complete work order is not required. If a cycle has several valid tasks, the latest is shown.

### Annual cycles

Cycles stay anchored to the original qualification date. Cycle 1 is generated, and every later cycle once it has started.

```text
Qualification date: 2025-06-03
Cycle 1: 2025-06-03 → 2026-06-03
Cycle 2: 2026-06-03 → 2027-06-03
Cycle 3: 2027-06-03 → 2028-06-03
```

A PM counts only in its own cycle. It never carries forward or repairs a missed cycle.

### Cycle result

| Condition | Result |
|---|---|
| Valid PM exists | Good / Completed |
| No PM, past the due date | Requalification Needed |
| No PM, 30 days or fewer remaining | Due Soon |
| No PM, more than 30 days remaining | Good |
| No task report for that qualification | Task Data Missing |

### Cycle history

Maximo task exports cover a rolling year, so old evidence ages out. Completed cycles are saved to `Data/cycle_history.csv`. On later runs a saved completion is restored (shown as `Historical Evidence`) unless a newer one is found in the source data. Only `Good / Completed` cycles are stored. Every other result is recalculated on each run. The file is saved only after all four workbooks are created successfully.

## Configuration

**`Config/qualification_mapping.csv`** defines what is tracked. Add areas, venues and qualifications here rather than in code.

| Venue | Qualification | Discipline | Tracked |
|---|---|---|---|
| Darkmoor | DUMOD1 | Effects | Yes |
| Wolf | CUMOD1 / CUMOD2 / CUMOD3 | RVs / Controls / Track | Yes |
| Unchained | UCMOD1 / UCMOD2 / UCMOD3 / UCMOD4 | RVs / Controls / Track / Animation | Yes |
| All | DUMOD0 / CUMOD0 / UCMOD0 | Safety | No |

**`Config/settings.json`** points to the Line Schedule folder:

```json
{
  "line_schedule_folder": "C:\\path\\to\\Line Schedule"
}
```

Set this to a folder on your machine that holds the schedule workbook.

## Project structure

```text
Qual_Tracker/
├── Run_Qualification_Tracker.bat   Double-click runner
├── Config/                         Qualification mapping and settings
├── Input/
│   ├── Qualification/              Maximo qualification reports
│   ├── Tasks/                      Maximo task history reports
│   └── Manual roster/              Animation roster
├── Scripts/
│   ├── report_generator.py         Main entry point: builds the workbooks
│   ├── cycle_processor.py          Reads inputs, builds cycles, assigns tasks
│   ├── cycle_history.py            Saves and restores completed-cycle evidence
│   ├── line_schedule_reader.py     Reads technician shifts
│   └── archive_outputs.py          Archives previous workbooks
├── Output/                         Generated workbooks
├── Archive/                        Timestamped copies of earlier outputs
├── Data/                           Permanent cycle history
├── Logs/                           Diagnostic CSVs
├── Documentation/                  Project record and coding-agent instructions
└── Tests/
```

## Diagnostic scripts

Run any of these with `.\.venv\Scripts\python.exe Scripts\<name>.py`:

| Script | Purpose |
|---|---|
| `input_scanner.py` | Lists input files and validates folders and configuration |
| `task_report_inspector.py` | Shows which qualifications each task report covers |
| `qualification_reader.py` | Validates the qualification reports |
| `completed_task_reader.py` | Reads and summarizes completed tasks |
| `cycle_processor.py` | Writes cycle validation CSVs to `Logs/` |
| `line_schedule_reader.py` | Validates the Line Schedule and reports shift conflicts |

## Privacy

This project handles employee names, SSOs, qualification dates and maintenance activity. These are never committed:

- `Input/Qualification/*` and `Input/Tasks/*`
- `Input/Manual roster/*`
- `Output/*`, `Archive/*`, `Logs/*`, `Data/*`
- `.venv/`

Use synthetic data in tests, issues and examples.

## Troubleshooting

| Message | Fix |
|---|---|
| A previous output workbook is open or locked | Close all four output workbooks, then run again. |
| Missing qualification reports: … | Add the file named in the message to `Input/Qualification/`. |
| qualification code was not found in filename | Rename the file so it contains one approved code. |
| More than one Line Schedule workbook was found | Leave exactly one `.xlsx` in the schedule folder. |
| Line Schedule folder not found | Fix `line_schedule_folder` in `Config/settings.json`. |
| Task Data Missing in a report | There is no task report covering that qualification. Add one. |
| Project Python environment was not found | Complete the one-time setup above. |
