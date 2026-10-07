# GitHub Copilot Repository Instructions

Follow `AGENTS.md` for complete repository rules.

## Project Intent

Build a local Python tracker that reads IBM Maximo qualification and rolling one-year task reports, calculates annual PM compliance cycles, and generates separate Excel workbooks for Wolf, Darkmoor, and Unchained.

## Mandatory Rules

- Use the local `.venv` interpreter.
- Never commit files from Input, Output, Archive, Logs, or `.venv`.
- Use synthetic test data only.
- Match task activity by SSO and Qualification.
- Keep only Status = COMP.
- Anchor every cycle to the original qualification Effective Date.
- Do not carry PM completion between annual cycles.
- Keep the latest valid PM task inside each cycle.
- Use Due Soon at 30 days.
- Use Task Data Missing if a qualification has no task-report coverage.
- Exclude DUMOD0, CUMOD0, and UCMOD0.
- Read mappings from `Config/qualification_mapping.csv`.
- Call `worksheet.reset_dimensions()` when reading Maximo XLSX worksheets with openpyxl.
- Do not use openpyxl Table objects in generated workbooks unless desktop Excel compatibility is verified.

## Expected Outputs

```text
Output/Wolf Qualification Tracker.xlsx
Output/Darkmoor Qualification Tracker.xlsx
Output/Unchained Qualification Tracker.xlsx
```

Each workbook should contain Dashboard, PM Status, Exceptions, and Run Summary.

## Development Approach

Prefer reusable functions, type hints where practical, clear error messages, and synthetic automated tests. Preserve current verified DUMOD1 behavior before refactoring.
