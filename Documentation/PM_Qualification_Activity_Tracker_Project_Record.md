# PM Qualification Activity Tracker

## Detailed Project Record and Technical Plan

**Project owner:** Frederick Herrera  
**Primary source:** IBM Maximo  
**Repository:** `https://github.com/Frederickhs/qual-tracker.git`  
**Local repository:** `G:\Dev\Qual_Tracker`  
**Default branch:** `main`  
**Current implementation:** Local Python with Excel output  
**Updated:** October 7, 2026

---

## 1. Purpose

The tracker determines whether each technician with a tracked qualification completed qualifying preventive-maintenance activity during every applicable annual qualification cycle.

The project replaces manual Ctrl+F searching through large Maximo task-history workbooks.

The final user workflow is intended to be:

1. Download unedited Maximo qualification reports as XLSX.
2. Download rolling one-year Maximo task-status history reports as XLSX.
3. Put qualification files in `Input\Qualification`.
4. Put task files in `Input\Tasks`.
5. Double-click `Run Qualification Tracker.bat`.
6. Open the generated venue workbooks from `Output`.

No hosting, Power BI, EXE packaging, or website is required.

---

## 2. Why Excel Was Built First

The original proof of concept was created in Excel and Power Query for DUMOD1. That work established and validated:

- Maximo report structure.
- SSO matching.
- COMP-only task logic.
- Qualification-cycle behavior.
- Multiple annual cycles.
- Status labels and warning threshold.
- Dashboard needs.
- Required task evidence fields.
- Archive workflow.

The Excel prototype was useful as a business-rules prototype, but Python became the preferred final implementation because the final user experience should be folder-based and double-click driven.

Python also proved better at handling thousands of Maximo worksheets and directly selecting the latest valid task per cycle.

---

## 3. Qualification Mapping

The authoritative configuration is:

```text
Config\qualification_mapping.csv
```

Mappings:

```text
Area          Venue       Qualification Discipline Tracked
Dark Universe Darkmoor    DUMOD0        Safety     No
Dark Universe Darkmoor    DUMOD1        Effects    Yes
Dark Universe Wolf        CUMOD0        Safety     No
Dark Universe Wolf        CUMOD1        RVs        Yes
Dark Universe Wolf        CUMOD2        Controls   Yes
Dark Universe Wolf        CUMOD3        Track      Yes
Dark Universe Unchained   UCMOD0        Safety     No
Dark Universe Unchained   UCMOD1        RVs        Yes
Dark Universe Unchained   UCMOD2        Controls   Yes
Dark Universe Unchained   UCMOD3        Track      Yes
Dark Universe Unchained   UCMOD4        Animation  Yes
```

All Mod 0 safety qualifications are excluded from PM tracking.

---

## 4. Business Rules

### Qualification identity

Qualification reports do not contain a dependable qualification field. The qualification report filename must contain the exact qualification code, for example:

```text
DUMOD1 Qualification.xlsx
CUMOD2 Qualification.xlsx
UCMOD4 Qualification.xlsx
```

The program must not infer qualification from Work Location.

### Task identity

Task-report filenames do not matter. The program identifies task records using the internal `Qualification` column.

One task workbook can contain several qualifications. Qualifications can also be split across several task workbooks.

### Master matching key

```text
Qualification report Labor = Task report Change By
Qualification code = Qualification code
```

SSO and Qualification together identify a technician qualification. Names are display-only fields.

### Valid PM

A task is valid when:

- Status is `COMP`.
- SSO matches.
- Qualification matches.
- Changed Date falls inside the cycle.

Any single qualifying task meets the requirement.

### Annual cycles

Cycles are permanently anchored to the original Effective Date.

Example:

```text
Effective Date: 2025-06-03
Cycle 1: 2025-06-03 to 2026-06-03
Cycle 2: 2026-06-03 to 2027-06-03
Cycle 3: 2027-06-03 to 2028-06-03
```

The report creates all cycles that have started by the run date. It does not create future cycles early.

Dates on both boundaries count.

A PM belongs only to the cycle containing its completion date. PM activity never carries forward or backward between cycles.

If a cycle contains several valid tasks, the latest valid task is displayed.

### Status calculation

```text
Valid PM exists                   Good / Completed
No PM, due date passed            Requalification Needed
No PM, 30 or fewer days remain    Due Soon
No PM, more than 30 days remain   Good
Task source unavailable           Task Data Missing
```

The Task Data Missing state protects the report from falsely labeling records when a qualification task export has not arrived.

---

## 5. Repository Setup

Repository structure:

```text
G:\Dev\Qual_Tracker\
â”œâ”€â”€ .venv\
â”œâ”€â”€ Input\
â”‚   â”œâ”€â”€ Qualification\
â”‚   â””â”€â”€ Tasks\
â”œâ”€â”€ Output\
â”œâ”€â”€ Archive\
â”œâ”€â”€ Config\
â”‚   â””â”€â”€ qualification_mapping.csv
â”œâ”€â”€ Scripts\
â”œâ”€â”€ Tests\
â”œâ”€â”€ Documentation\
â”œâ”€â”€ Logs\
â”œâ”€â”€ .gitignore
â”œâ”€â”€ AGENTS.md
â”œâ”€â”€ CLAUDE.md
â”œâ”€â”€ README.md
â”œâ”€â”€ requirements.txt
â””â”€â”€ Run Qualification Tracker.bat
```

Real Maximo inputs, generated outputs, logs, archives, virtual environments, Python cache, and Excel temporary files are excluded from Git.

A Git safe-directory exception was required because the G: file system does not record ownership:

```powershell
git config --global --add safe.directory "G:/Dev/Qual Tracker"
```

---

## 6. Python Environment

Verified environment:

```text
Python 3.14.8
pandas 3.0.6
openpyxl 3.1.5
Git 2.54.0.windows.1
```

Project environment:

```text
G:\Dev\Qual_Tracker\.venv
```

Setup commands:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Scripts are run through the local environment:

```powershell
.\.venv\Scripts\python.exe Scripts\report_generator.py
```

---

## 7. Scripts Built

### `input_scanner.py`

- Validates required folders and configuration.
- Lists qualification and task reports.
- Ignores temporary `~$` files.
- Reports file counts and sizes.

### `task_report_inspector.py`

- Inspects every worksheet in every task workbook.
- Counts worksheets, task rows, and COMP rows.
- Lists qualification coverage per file.
- Uses `worksheet.reset_dimensions()` because Maximo XLSX worksheet dimensions are unreliable.

### `qualification_reader.py`

- Reads all eight qualification reports.
- Extracts qualification codes from filenames.
- Maps Area, Venue, and Discipline through configuration.
- Counts active/inactive records.
- Validates Effective Dates.
- Detects duplicate SSO/Qualification rows.

Verified qualification totals:

```text
Reports: 8
Qualification records: 334
Unique technicians: 79
Duplicate SSO/Qualification records: 0
Invalid dates: 0
```

### `completed_task_reader.py`

- Reads all task reports.
- Processes all worksheets.
- Collects only COMP task rows.
- Reports qualification coverage and missing task qualifications.
- Detects exact duplicate task rows.

A test with three task files processed:

```text
5,736 worksheets
45,972 COMP rows
75 task technicians
0 exact duplicate rows
0 invalid completion dates
35.3 seconds
```

### `cycle_processor.py`

- Combines qualification records.
- Generates all active annual cycles.
- Matches tasks by SSO and Qualification.
- Assigns tasks to the correct cycle.
- Retains the latest valid task.
- Calculates Days Remaining and Cycle Result.
- Creates local validation and summary CSV files in Logs.

A test with available inputs processed:

```text
334 qualification records
508 annual cycle records
45,972 completed task rows
79 technicians
45.5 seconds
```

### `report_generator.py`

- Reuses cycle processing functions.
- Generates one workbook per venue.
- Adds Dashboard, PM Status, Exceptions, and Run Summary.
- Applies status colors and worksheet filters.

Generated outputs:

```text
Output\Wolf Qualification Tracker.xlsx
Output\Darkmoor Qualification Tracker.xlsx
Output\Unchained Qualification Tracker.xlsx
```

First tested row counts:

```text
Wolf: 174 cycles
Darkmoor: 59 cycles
Unchained: 275 cycles
```

---

## 8. Maximo Workbook Technical Findings

Maximo task workbooks can contain thousands of worksheets.

Examples processed:

```text
DUMOD1 workbook: 327 worksheets, 13,077 task rows, 3,134 COMP rows
Wolf workbook: 2,625 worksheets, 104,989 task rows, 20,562 COMP rows
Unchained UCMOD3/UCMOD4 workbook: 2,784 worksheets, 111,339 task rows, 22,276 COMP rows
```

In openpyxl read-only mode, the Maximo workbook initially appeared to contain only Column A. Calling:

```python
worksheet.reset_dimensions()
```

corrected the worksheet dimensions and exposed all columns.

The harmless missing-default-style warning is suppressed specifically.

---

## 9. Excel Compatibility Issue and Fix

The initial `report_generator.py` used openpyxl Table objects. Desktop Excel opened the output with a repair warning and removed:

```text
AutoFilter from /xl/tables/table1.xml
Table from /xl/tables/table1.xml
AutoFilter from /xl/tables/table2.xml
Table from /xl/tables/table2.xml
```

The fix was:

- Remove `Table` and `TableStyleInfo` imports.
- Remove generated Table objects.
- Use normal worksheet ranges.
- Use `worksheet.auto_filter.ref = worksheet.dimensions`.

After regeneration, workbooks opened normally without repair warnings.

---

## 10. Verified DUMOD1 Results

The full-year DUMOD1 task report produced:

```text
Total Cycles: 59
Good / Completed: 27
Good: 16
Due Soon: 3
Requalification Needed: 13
Task Data Missing: 0
```

A direct source validation confirmed task `T22746641`:

```text
SSO: 703146043
Qualification: DUMOD1
Status: COMP
Changed Date: 2026-10-06 07:35:29
Cycle: Cycle 2, 2026-06-03 through 2027-06-03
Expected result: Good / Completed
```

The Python result was correct. The earlier Power Query prototype missed this and other valid full-year tasks, likely because its sort-plus-remove-duplicates approach did not reliably retain the correct row.

Python is now the authoritative processing path.

---

## 11. Current Input Status

All eight qualification reports are present.

Latest saved task coverage:

```text
DUMOD1 present
CUMOD1 present
CUMOD2 present
CUMOD3 present
UCMOD2 present
UCMOD3 present
UCMOD4 present
UCMOD1 still downloading or problematic
```

UCMOD1 records must display Task Data Missing until the task report becomes available.

---

## 12. Output Workbook Design

Three separate venue workbooks are required.

Each workbook contains:

### Dashboard

- Total Cycles
- Good / Completed
- Good
- Due Soon
- Requalification Needed
- Task Data Missing
- Processing Date
- Cycle rule note

### PM Status

Fields:

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

### Exceptions

Shows Task Data Missing and inactive qualification records.

### Run Summary

Shows venue, processing date, qualification count, cycle count, unique technicians, qualifications, and disciplines.

---

## 13. Planned Automatic Archive

Automatic archiving is not yet implemented.

Planned archive structure:

```text
Archive\YYYY-MM-DD_HHMM\
â”œâ”€â”€ Qualification\
â”œâ”€â”€ Tasks\
â””â”€â”€ Output\
```

The first implementation should copy, not move, source and output files.

Archive only after a successful run. A partial run with Task Data Missing may generate useful reports, but should not be marked as a complete monthly archive unless explicitly allowed.

---

## 14. Planned BAT Runner

The final BAT file should:

1. Change to the repository folder.
2. Confirm `.venv\Scripts\python.exe` exists.
3. Run `Scripts\report_generator.py`.
4. Display clear progress and errors.
5. Stop on failure.
6. Open `Output` after success.
7. Keep the command window visible until acknowledged.

No EXE is required.

---

## 15. Expansion Beyond Dark Universe

The processor should remain configuration-driven.

If another area uses the same Maximo structure and annual-cycle rules, expansion should primarily require:

1. Add Area, Venue, Qualification, Discipline, and Tracked rows to configuration.
2. Add corresponding qualification reports.
3. Add task reports.
4. Allow output filenames to be created dynamically by venue.
5. Test the new area with source validation.

Avoid Dark Universe-specific logic in the processing engine.

---

## 16. Remaining Work

1. Obtain UCMOD1 task report.
2. Confirm all eight tracked task qualifications are present.
3. Regenerate all output workbooks.
4. Confirm Unchained Task Data Missing becomes zero.
5. Commit the final compatible `report_generator.py`.
6. Create synthetic automated tests.
7. Refactor duplicated reader logic into reusable modules if safe.
8. Implement automatic archive.
9. Finish `Run Qualification Tracker.bat`.
10. Update README and user operating instructions.
11. Perform final accuracy checks for Wolf and Unchained.
12. Tag a stable release after acceptance testing.

---

## 17. Preferred Working Style

Instructions should be simple and step-by-step, but routine saves and tiny actions should be grouped into meaningful checkpoints.

Do not assume missing business rules. Ask before changing logic.

