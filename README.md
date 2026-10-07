# PM Qualification Activity Tracker

Local Python application for processing IBM Maximo qualification and task-status history exports into annual PM compliance reports.

## What It Does

The tracker:

- Reads all current qualification reports.
- Reads every worksheet from all rolling one-year task reports.
- Keeps only completed tasks with `Status = COMP`.
- Matches technicians by SSO and qualification.
- Creates annual cycles tied permanently to the original qualification date.
- Assigns each PM task only to the cycle in which it was completed.
- Displays the latest valid task in each cycle.
- Calculates Good, Good / Completed, Due Soon, Requalification Needed, or Task Data Missing.
- Generates one Excel workbook per venue.

## Generated Reports

```text
Output/Wolf Qualification Tracker.xlsx
Output/Darkmoor Qualification Tracker.xlsx
Output/Unchained Qualification Tracker.xlsx
```

Each workbook contains:

- Dashboard
- PM Status
- Exceptions
- Run Summary

## Tracked Qualifications

### Darkmoor

- DUMOD1: Effects

### Wolf

- CUMOD1: RVs
- CUMOD2: Controls
- CUMOD3: Track

### Unchained

- UCMOD1: RVs
- UCMOD2: Controls
- UCMOD3: Track
- UCMOD4: Animation

The following Mod 0 safety qualifications are excluded:

```text
DUMOD0
CUMOD0
UCMOD0
```

## Business Rules

A PM task counts when:

- `Status = COMP`
- Task SSO matches qualification Labor/SSO
- Task qualification matches the technician qualification
- Completion date falls inside the applicable annual cycle

Cycles remain tied to the original qualification Effective Date.

Example:

```text
Qualification Date: 2025-06-03
Cycle 1: 2025-06-03 through 2026-06-03
Cycle 2: 2026-06-03 through 2027-06-03
Cycle 3: 2027-06-03 through 2028-06-03
```

A task counts only in its own cycle. It does not carry into another cycle.

Status rules:

```text
Valid PM exists                         Good / Completed
No PM and deadline passed               Requalification Needed
No PM and 30 or fewer days remain       Due Soon
No PM and more than 30 days remain      Good
Task report unavailable                 Task Data Missing
```

## Repository Structure

```text
Qual Tracker/
├── Input/
│   ├── Qualification/
│   └── Tasks/
├── Output/
├── Archive/
├── Config/
├── Scripts/
├── Tests/
├── Documentation/
├── Logs/
├── .gitignore
├── AGENTS.md
├── CLAUDE.md
├── README.md
├── requirements.txt
└── Run Qualification Tracker.bat
```

Real Maximo reports, generated workbooks, logs, archives, and the virtual environment are excluded from Git.

## Requirements

- Windows
- Python 3.14.x
- pandas 3.0.6
- openpyxl 3.1.5

## Initial Setup

From the repository root:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Verify:

```powershell
.\.venv\Scripts\python.exe -c "import pandas; import openpyxl; print('Python environment is ready')"
```

## Input Files

### Qualification reports

Place all qualification reports in:

```text
Input/Qualification
```

Qualification filenames must contain the exact qualification code:

```text
DUMOD1 Qualification.xlsx
CUMOD1 Qualification.xlsx
CUMOD2 Qualification.xlsx
CUMOD3 Qualification.xlsx
UCMOD1 Qualification.xlsx
UCMOD2 Qualification.xlsx
UCMOD3 Qualification.xlsx
UCMOD4 Qualification.xlsx
```

Do not infer qualification from Work Location.

### Task reports

Place all rolling one-year task-history reports in:

```text
Input/Tasks
```

Task filenames can remain as downloaded from Maximo. The tracker uses the internal `Qualification` column.

A task file can contain one or several qualifications. Multiple files for the same venue are supported.

Close all source workbooks before running the tracker so temporary `~$` files are not created.

## Run the Current Generator

```powershell
.\.venv\Scripts\python.exe Scripts\report_generator.py
```

Generated workbooks appear in:

```text
Output
```

The final target is to run the same process by double-clicking:

```text
Run Qualification Tracker.bat
```

## Diagnostic Scripts

Scan input files:

```powershell
.\.venv\Scripts\python.exe Scripts\input_scanner.py
```

Inspect task qualification coverage:

```powershell
.\.venv\Scripts\python.exe Scripts\task_report_inspector.py
```

Validate qualification reports:

```powershell
.\.venv\Scripts\python.exe Scripts\qualification_reader.py
```

Read completed tasks:

```powershell
.\.venv\Scripts\python.exe Scripts\completed_task_reader.py
```

Generate cycle validation CSV files:

```powershell
.\.venv\Scripts\python.exe Scripts\cycle_processor.py
```

## Verified DUMOD1 Reference

With the full-year DUMOD1 report and processing date October 7, 2026:

```text
Total Cycles: 59
Good / Completed: 27
Good: 16
Due Soon: 3
Requalification Needed: 13
Task Data Missing: 0
```

Python source validation confirmed a Cycle 2 task that the earlier Power Query prototype missed. Python is the authoritative processing path.

## Privacy

This project handles employee names, SSOs, qualification dates, and task activity.

Never commit:

- Maximo input reports
- Generated output workbooks
- Logs or validation CSV files
- Archives
- Employee names or SSOs in tests
- `.venv`

Use synthetic data for tests and examples.

## Configuration

Qualification mappings are stored in:

```text
Config/qualification_mapping.csv
```

Add new areas, venues, qualifications, and disciplines through configuration rather than hardcoding them in processing logic.

## Current Status

Completed:

- Repository and virtual environment
- Input scanning
- Multi-sheet Maximo task inspection
- Qualification reader
- Completed-task reader
- Annual cycle processor
- Separate venue report generation
- Missing-task-data protection
- DUMOD1 source validation
- Excel compatibility fix

Remaining:

- Obtain UCMOD1 task report
- Complete Unchained validation
- Commit final report generator
- Add synthetic automated tests
- Add automatic archiving
- Finish the BAT runner
- Complete final documentation and release testing

## Agent Instructions

Coding agents should read:

```text
AGENTS.md
CLAUDE.md
.github/copilot-instructions.md
```

`AGENTS.md` is the authoritative repository instruction file.
