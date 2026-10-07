# PM Qualification Activity Tracker

Python-based tracker that processes IBM Maximo qualification and task-history reports.

## Planned Workflow

1. Place all current qualification reports in `Input/Qualification`.
2. Place all rolling one-year task reports in `Input/Tasks`.
3. Double-click `Run Qualification Tracker.bat`.
4. Review the generated Excel workbooks in `Output`.

## Planned Outputs

- Wolf Qualification Tracker.xlsx
- Darkmoor Qualification Tracker.xlsx
- Unchained Qualification Tracker.xlsx

## Venue Mapping

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

## Excluded Qualifications

- DUMOD0
- CUMOD0
- UCMOD0

## Privacy

Real Maximo reports, generated outputs, archives, logs, technician names, and SSOs are excluded from Git.
