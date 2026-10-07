# CLAUDE.md

## Mission

Help develop and maintain the PM Qualification Activity Tracker in this private repository. Follow `AGENTS.md` as the authoritative repository instruction file. This file adds Claude Code-specific working guidance.

## Start Every Session

1. Read `AGENTS.md`.
2. Read `README.md`.
3. Inspect `git status`.
4. Inspect `Config/qualification_mapping.csv`.
5. Identify which task qualifications are currently available without opening or committing private data.
6. Do not modify files until the requested task and current state are understood.

## Project Commands

Use the project virtual environment:

```powershell
.\.venv\Scripts\python.exe Scripts\input_scanner.py
.\.venv\Scripts\python.exe Scripts\task_report_inspector.py
.\.venv\Scripts\python.exe Scripts\qualification_reader.py
.\.venv\Scripts\python.exe Scripts\completed_task_reader.py
.\.venv\Scripts\python.exe Scripts\cycle_processor.py
.\.venv\Scripts\python.exe Scripts\report_generator.py
```

Never assume `python` is on PATH.

## Privacy

Treat all files in Input, Output, Archive, and Logs as private employee data. Do not quote names, SSOs, dates, or task history in chat, commits, documentation, or test fixtures unless explicitly required for a private validation step.

Do not use real data to create repository test files. Use synthetic fixtures.

## Coding Priorities

1. Correct annual-cycle logic.
2. Protection against missing task data.
3. Clear user-facing error messages.
4. Excel compatibility.
5. Fast processing of thousands of worksheets.
6. Configuration-driven expansion to other areas.
7. Simple double-click monthly operation.

## Non-Negotiable Logic

- Match SSO plus Qualification.
- Count only COMP task rows.
- Anchor cycles to the original Effective Date permanently.
- Include cycle boundary dates.
- Do not carry PM activity between cycles.
- Show the latest valid task within each cycle.
- Use a 30-day Due Soon threshold.
- Use Task Data Missing when qualification task coverage is unavailable.
- Exclude Mod 0 safety qualifications.

## Maximo Workbook Handling

Always call:

```python
worksheet.reset_dimensions()
```

when reading Maximo worksheets through openpyxl read-only mode.

Task workbooks may contain thousands of worksheets and several qualifications. Qualification workbooks require a qualification code in the filename.

## Excel Generation Warning

Do not reintroduce openpyxl Table objects without a desktop Excel compatibility test. Previous Table/AutoFilter objects caused Excel to repair generated workbooks. Use ordinary cells, formatting, conditional formatting, freeze panes, and worksheet AutoFilter.

## Change Workflow

For each requested change:

1. State the change in one short paragraph.
2. Inspect related code before editing.
3. Modify the minimum number of files.
4. Run targeted scripts or tests.
5. Confirm generated workbooks open without repair warnings.
6. Report exact counts or timings when relevant.
7. Provide a proposed Git commit message.

## Communication Style

Use clear, direct language. Group routine actions into practical checkpoints. Do not stop after every save, file creation, or tiny command. Ask before changing a business rule or output meaning.
