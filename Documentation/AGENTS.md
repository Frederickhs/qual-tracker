# AGENTS.md

## Purpose

This file provides repository-wide instructions for coding agents such as OpenAI Codex, GitHub Copilot coding agent, Claude Code, and similar tools working on the PM Qualification Activity Tracker.

## Project Summary

The PM Qualification Activity Tracker processes unedited IBM Maximo qualification and task-status history exports. It generates one Excel workbook per venue:

- `Wolf Qualification Tracker.xlsx`
- `Darkmoor Qualification Tracker.xlsx`
- `Unchained Qualification Tracker.xlsx`

The normal user workflow must remain:

1. Download current Maximo reports as XLSX.
2. Place all qualification reports in `Input/Qualification`.
3. Place all rolling one-year task reports in `Input/Tasks`.
4. Double-click `Run Qualification Tracker.bat`.
5. Review the generated workbooks in `Output`.

The final user must not need to edit Python, use VS Code, use Power Query, or enter terminal commands for a normal monthly run.

## Repository Location and Runtime

Expected local repository:

```text
G:\Dev\Qual_Tracker
```

Expected runtime:

- Windows
- Python 3.14.x through the local `.venv`
- pandas 3.0.6
- openpyxl 3.1.5

Always run project scripts with:

```powershell
.\.venv\Scripts\python.exe Scripts\<script_name>.py
```

Do not rely on the global `python` command. The installed launcher may be available only as `py`.

## Privacy and Data Handling

The project processes employee names, SSOs, qualification dates, and maintenance activity.

Never commit, upload, or place real employee data in source-controlled files.

These locations must remain excluded from Git:

```text
Input/Qualification/*
Input/Tasks/*
Output/*
Archive/*
Logs/*
.venv/
```

Do not add real Maximo files, generated workbooks, logs, CSV validation outputs, employee names, or SSOs to tests, issues, documentation, examples, commits, or pull requests.

Use synthetic identifiers in tests and documentation.

## Authoritative Business Rules

### Tracked qualifications

```text
Dark Universe | Darkmoor  | DUMOD1 | Effects
Dark Universe | Wolf      | CUMOD1 | RVs
Dark Universe | Wolf      | CUMOD2 | Controls
Dark Universe | Wolf      | CUMOD3 | Track
Dark Universe | Unchained | UCMOD1 | RVs
Dark Universe | Unchained | UCMOD2 | Controls
Dark Universe | Unchained | UCMOD3 | Track
Dark Universe | Unchained | UCMOD4 | Animation
```

### Excluded qualifications

All Mod 0 safety qualifications are excluded:

```text
DUMOD0
CUMOD0
UCMOD0
```

The configuration file is authoritative:

```text
Config/qualification_mapping.csv
```

Do not duplicate venue or discipline mappings throughout the code. Read configuration instead.

### Matching rule

Match qualification records to completed task records using both:

```text
Qualification Labor/SSO = Task Change By/SSO
Qualification code = Task Qualification
```

Names are display fields only and must not be used as the primary key.

### Valid PM activity

A task counts when:

1. `Status` is exactly `COMP` after trimming and standardizing case.
2. SSO matches.
3. Qualification matches.
4. Completion date falls inside the applicable annual cycle.

Any individual qualifying task counts. A complete work order is not required.

### Annual cycles

Cycles remain permanently anchored to the original qualification Effective Date.

Example:

```text
Original qualification date: 2025-06-03
Cycle 1: 2025-06-03 through 2026-06-03
Cycle 2: 2026-06-03 through 2027-06-03
Cycle 3: 2027-06-03 through 2028-06-03
```

Generate Cycle 1 and every cycle that has started as of the processing date. Do not generate a future cycle before its start date.

Cycle boundaries are inclusive.

A PM counts only in the cycle containing its completion date. It must not carry into a later cycle or repair a missed earlier cycle.

If multiple valid tasks exist in one cycle, display the latest valid task.

### Status logic

```text
Valid PM exists                         -> Good / Completed
No PM and Days Remaining < 0            -> Requalification Needed
No PM and Days Remaining <= 30          -> Due Soon
No PM and Days Remaining > 30           -> Good
Task report unavailable for qualification -> Task Data Missing
```

Do not label missing task coverage as Good, Due Soon, or Requalification Needed.

### Inactive qualification records

Preserve inactive qualification records for audit visibility. Do not silently remove them. Their presence should be visible in Exceptions or equivalent review output.

## Maximo Input Rules

### Qualification reports

Qualification files do not reliably identify the qualification inside the workbook. The filename must contain exactly one approved qualification code, for example:

```text
CUMOD1 Qualification.xlsx
UCMOD4 Qualification.xlsx
```

Expected worksheet:

```text
List of Labor
```

Expected fields include:

```text
Labor
Name
Effective Date
Expiration Date
Status
Work Site
Work Location
```

Do not infer qualification from Work Location.

### Task reports

Task filenames do not matter. Identify task records using the internal `Qualification` column.

One task workbook may contain one or several qualifications. A qualification may also be split across multiple task files.

Expected fields include:

```text
Task Num
Status
Changed Date
Change By
Name
Qualification
```

The task description is normally stored in the first column.

Maximo XLSX reports can have incorrect worksheet dimensions. When using openpyxl in read-only mode, call:

```python
worksheet.reset_dimensions()
```

before iterating rows.

Maximo workbooks may show the harmless warning that a workbook contains no default style. Suppress only that specific warning. Do not suppress unrelated warnings globally.

## Current Scripts

```text
Scripts/input_scanner.py
Scripts/task_report_inspector.py
Scripts/qualification_reader.py
Scripts/completed_task_reader.py
Scripts/cycle_processor.py
Scripts/report_generator.py
```

Roles:

- `input_scanner.py`: Lists available input files and validates configuration/folders.
- `task_report_inspector.py`: Inspects task workbook worksheets, row counts, COMP counts, and qualification coverage.
- `qualification_reader.py`: Reads and validates all qualification reports.
- `completed_task_reader.py`: Collects COMP task rows and summarizes task coverage.
- `cycle_processor.py`: Generates cycles, assigns tasks, and calculates statuses.
- `report_generator.py`: Generates one formatted workbook per venue.

Prefer extracting reusable logic into modules rather than duplicating readers across scripts. Preserve existing behavior while refactoring.

## Output Requirements

Generate these files in `Output`:

```text
Wolf Qualification Tracker.xlsx
Darkmoor Qualification Tracker.xlsx
Unchained Qualification Tracker.xlsx
```

Each workbook should contain:

```text
Dashboard
PM Status
Exceptions
Run Summary
```

The PM Status worksheet should include:

```text
Venue
Discipline
Qualification
SSO
Technician Name
Qualification Date
Qualification Status
Cycle Number
Cycle Start Date
Cycle Due Date
Task Data Available
PM Completion Date
Task Description
Task Number
Days Remaining
Cycle Result
```

Status colors:

```text
Good / Completed       Green
Good                   Blue
Due Soon               Yellow
Requalification Needed Red
Task Data Missing      Purple
```

Do not use openpyxl `Table` objects in generated reports unless a compatibility test proves they open correctly in desktop Excel. A previous implementation produced corrupted workbooks because Excel removed generated Table and AutoFilter features. The current safe approach is normal worksheet ranges plus:

```python
worksheet.auto_filter.ref = worksheet.dimensions
```

Generated workbooks must open without Excel repair warnings.

## Verified Reference Results

With the verified full-year DUMOD1 task report and processing date 2026-10-07, Python produced:

```text
Total Cycles: 59
Good / Completed: 27
Good: 16
Due Soon: 3
Requalification Needed: 13
Task Data Missing: 0
```

Verified source example:

```text
Synthetic description of reference:
SSO 703146043
Qualification DUMOD1
Cycle 2 start 2026-06-03
Task T22746641
COMP date 2026-10-06 07:35
Expected result Good / Completed
```

Do not place this real SSO in new public documentation or tests. It is retained here only because this repository instruction file is intended for the private project. Prefer replacing it with a synthetic regression fixture before broader distribution.

## Current Task Coverage

At the latest saved checkpoint:

```text
DUMOD1 available
CUMOD1 available
CUMOD2 available
CUMOD3 available
UCMOD2 available
UCMOD3 available
UCMOD4 available
UCMOD1 missing or still downloading
```

The processor must continue successfully when one qualification lacks task data and mark those cycles `Task Data Missing`.

## Automated Archive Requirement

Automatic archiving is planned but not yet implemented.

Preferred behavior after a fully successful run:

```text
Archive/YYYY-MM-DD_HHMM/
â”œâ”€â”€ Qualification/
â”œâ”€â”€ Tasks/
â””â”€â”€ Output/
```

Copy input and output files. Do not move source files in the first implementation.

Archive only after:

- All requested output workbooks are created successfully.
- No critical processing errors occurred.
- Workbook compatibility checks pass.

If task data is incomplete, generate reports with `Task Data Missing`, but do not treat the run as a complete monthly archive unless the user explicitly allows partial archives.

## BAT Runner Requirement

The final `Run Qualification Tracker.bat` should:

1. Change to the repository root.
2. Verify `.venv\Scripts\python.exe` exists.
3. Run `Scripts\report_generator.py`.
4. Stop and display a clear error if processing fails.
5. Open the `Output` folder after success.
6. Keep the window open long enough for the user to read the result.

Do not require an EXE.

## Testing Requirements

Before changing cycle logic, create or update tests for:

- Task on cycle start date.
- Task on cycle due date.
- Task one day before cycle start.
- Task one day after cycle due date.
- Multiple valid tasks in one cycle, latest retained.
- PM in Cycle 1 not carrying into Cycle 2.
- PM in Cycle 2 not correcting Cycle 1.
- No task data available for a qualification.
- Inactive qualification record.
- Duplicate task rows across files.
- Qualification filename missing a code.
- Multiple codes in one qualification filename.
- Leap-day qualification date.
- Future cycle not created before its start.

Use synthetic records only.

## Agent Working Style

The project owner prefers clear, simple instructions grouped into meaningful checkpoints.

Do not pause for routine saves, tiny edits, or obvious confirmations. Group minor actions together. Stop at meaningful verification points such as:

- Environment setup completed.
- Input validation completed.
- Script execution completed.
- Output workbook reviewed.
- Accuracy test passed.
- Commit ready.

Do not assume missing business rules. Ask before changing behavior.

When making code changes:

1. Explain the planned change briefly.
2. Make the smallest safe change.
3. Run the relevant script or tests.
4. Check generated output compatibility.
5. Summarize results and remaining concerns.
6. Suggest a concise Git commit message.

## Git Rules

Repository:

```text
https://github.com/Frederickhs/qual-tracker.git
```

Default branch:

```text
main
```

Never force-push unless explicitly instructed.

Do not commit generated output or personal data.

Before committing:

```powershell
git status
git diff --check
```

Use concise commit messages such as:

```text
Add completed task reader
Add annual qualification cycle processor
Fix Excel output compatibility
Add automatic archive workflow
```

