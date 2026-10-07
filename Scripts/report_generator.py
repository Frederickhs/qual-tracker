from pathlib import Path
import sys
import time

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill

from cycle_processor import (
    TODAY,
    assign_tasks_to_cycles,
    build_cycles,
    load_mapping,
    read_qualifications,
    read_tasks,
)


PROJECT_ROOT = Path(__file__).resolve().parent.parent
OUTPUT_FOLDER = PROJECT_ROOT / "Output"

OUTPUT_FILES = {
    "Wolf": "Wolf Qualification Tracker.xlsx",
    "Darkmoor": "Darkmoor Qualification Tracker.xlsx",
    "Unchained": "Unchained Qualification Tracker.xlsx",
}

STATUS_STYLES = {
    "Good / Completed": ("C6EFCE", "006100"),
    "Good": ("DDEBF7", "1F4E78"),
    "Due Soon": ("FFF2CC", "7F6000"),
    "Requalification Needed": ("F4CCCC", "9C0006"),
    "Task Data Missing": ("E4DFEC", "5F497A"),
}

DISPLAY_COLUMNS = [
    "Venue",
    "Discipline",
    "Qualification",
    "SSO",
    "Technician Name",
    "Qualification Date",
    "Qualification Status",
    "Cycle Number",
    "Cycle Start Date",
    "Cycle Due Date",
    "Task Data Available",
    "PM Completion Date",
    "Task Description",
    "Task Number",
    "Days Remaining",
    "Cycle Result",
]


def write_dataframe(worksheet, dataframe):
    worksheet.append(list(dataframe.columns))

    for row in dataframe.itertuples(index=False, name=None):
        worksheet.append(list(row))


def format_table_sheet(worksheet, table_name):
    worksheet.freeze_panes = "A2"

    header_fill = PatternFill("solid", fgColor="548235")
    header_font = Font(color="FFFFFF", bold=True)

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    widths = {
        "A": 15,
        "B": 15,
        "C": 15,
        "D": 14,
        "E": 31,
        "F": 18,
        "G": 20,
        "H": 14,
        "I": 18,
        "J": 18,
        "K": 19,
        "L": 21,
        "M": 44,
        "N": 16,
        "O": 16,
        "P": 26,
    }

    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width

    for row in worksheet.iter_rows(min_row=2):
        row[3].number_format = "@"
        row[5].number_format = "m/d/yyyy"
        row[8].number_format = "m/d/yyyy"
        row[9].number_format = "m/d/yyyy"
        row[11].number_format = "m/d/yyyy h:mm AM/PM"
        row[12].alignment = Alignment(wrap_text=True, vertical="top")

    if worksheet.max_row >= 2:
        worksheet.auto_filter.ref = worksheet.dimensions

        status_range = f"P2:P{worksheet.max_row}"
        for status, (fill_color, font_color) in STATUS_STYLES.items():
            worksheet.conditional_formatting.add(
                status_range,
                FormulaRule(
                    formula=[f'$P2="{status}"'],
                    fill=PatternFill("solid", fgColor=fill_color),
                    font=Font(color=font_color, bold=True),
                ),
            )


def create_dashboard(workbook, venue_data, venue):
    worksheet = workbook.active
    worksheet.title = "Dashboard"
    worksheet.sheet_view.showGridLines = False

    worksheet["A1"] = "PM Qualification Activity Tracker"
    worksheet["A1"].font = Font(size=20, bold=True, color="FFFFFF")
    worksheet["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    worksheet.merge_cells("A1:D1")

    worksheet["A2"] = f"{venue} Annual PM Compliance"
    worksheet["A2"].font = Font(size=12, bold=True)
    worksheet.merge_cells("A2:D2")

    worksheet["A4"] = "Metric"
    worksheet["B4"] = "Count"

    metrics = [
        ("Total Cycles", len(venue_data)),
        (
            "Good / Completed",
            int((venue_data["Cycle Result"] == "Good / Completed").sum()),
        ),
        ("Good", int((venue_data["Cycle Result"] == "Good").sum())),
        ("Due Soon", int((venue_data["Cycle Result"] == "Due Soon").sum())),
        (
            "Requalification Needed",
            int((venue_data["Cycle Result"] == "Requalification Needed").sum()),
        ),
        (
            "Task Data Missing",
            int((venue_data["Cycle Result"] == "Task Data Missing").sum()),
        ),
    ]

    for row_number, (label, value) in enumerate(metrics, start=5):
        worksheet.cell(row=row_number, column=1, value=label)
        worksheet.cell(row=row_number, column=2, value=value)
        worksheet.cell(row=row_number, column=1).font = Font(bold=True)
        worksheet.cell(row=row_number, column=2).font = Font(bold=True, size=14)

        if label in STATUS_STYLES:
            fill_color, font_color = STATUS_STYLES[label]
            worksheet.cell(row=row_number, column=1).fill = PatternFill(
                "solid", fgColor=fill_color
            )
            worksheet.cell(row=row_number, column=1).font = Font(
                bold=True, color=font_color
            )

    worksheet["A12"] = "Processing Date"
    worksheet["B12"] = TODAY.to_pydatetime()
    worksheet["B12"].number_format = "m/d/yyyy"

    worksheet["A14"] = (
        "Each PM counts only for the annual cycle in which it was completed."
    )
    worksheet["A14"].font = Font(italic=True, color="666666")
    worksheet.merge_cells("A14:F14")

    worksheet.column_dimensions["A"].width = 32
    worksheet.column_dimensions["B"].width = 18
    worksheet.column_dimensions["C"].width = 18
    worksheet.column_dimensions["D"].width = 18


def create_exceptions_sheet(workbook, venue_data):
    worksheet = workbook.create_sheet("Exceptions")

    exceptions = venue_data[
        (venue_data["Cycle Result"] == "Task Data Missing")
        | (venue_data["Qualification Status"] != "ACTIVE")
    ][DISPLAY_COLUMNS].copy()

    if exceptions.empty:
        worksheet["A1"] = "No exceptions found."
        worksheet["A1"].font = Font(bold=True)
        return

    write_dataframe(worksheet, exceptions)
    format_table_sheet(worksheet, "ExceptionsTable")


def create_run_summary(workbook, venue_data, venue):
    worksheet = workbook.create_sheet("Run Summary")
    worksheet.sheet_view.showGridLines = False

    rows = [
        ("Venue", venue),
        ("Processing Date", TODAY.to_pydatetime()),
        ("Qualification Records", venue_data[["SSO", "Qualification"]].drop_duplicates().shape[0]),
        ("Annual Cycle Records", len(venue_data)),
        ("Unique Technicians", venue_data["SSO"].nunique()),
        (
            "Qualifications",
            ", ".join(sorted(venue_data["Qualification"].unique())),
        ),
        (
            "Disciplines",
            ", ".join(sorted(venue_data["Discipline"].unique())),
        ),
    ]

    for row_number, (label, value) in enumerate(rows, start=1):
        worksheet.cell(row=row_number, column=1, value=label)
        worksheet.cell(row=row_number, column=2, value=value)
        worksheet.cell(row=row_number, column=1).font = Font(bold=True)

    worksheet["B2"].number_format = "m/d/yyyy"
    worksheet.column_dimensions["A"].width = 28
    worksheet.column_dimensions["B"].width = 60


def create_venue_workbook(results, venue, output_path):
    venue_data = results[results["Venue"] == venue].copy()

    workbook = Workbook()
    create_dashboard(workbook, venue_data, venue)

    status_sheet = workbook.create_sheet("PM Status")
    status_data = venue_data[DISPLAY_COLUMNS].copy()
    write_dataframe(status_sheet, status_data)
    format_table_sheet(status_sheet, "PMStatusTable")

    create_exceptions_sheet(workbook, venue_data)
    create_run_summary(workbook, venue_data, venue)

    workbook.save(output_path)


def main():
    started = time.perf_counter()

    print("=" * 70)
    print("PM Qualification Activity Tracker - Excel Report Generator")
    print("=" * 70)

    try:
        mapping = load_mapping()
        qualifications = read_qualifications(mapping)
        tasks, qualifications_with_tasks = read_tasks()
        cycles = build_cycles(qualifications)
        results = assign_tasks_to_cycles(
            cycles,
            tasks,
            qualifications_with_tasks,
        )
    except Exception as error:
        print(f"ERROR: {error}")
        return 1

    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    for venue, filename in OUTPUT_FILES.items():
        output_path = OUTPUT_FOLDER / filename
        create_venue_workbook(results, venue, output_path)
        venue_count = int((results["Venue"] == venue).sum())
        print(f"Created: {filename} ({venue_count:,} cycle rows)")

    print()
    print(f"Output folder: {OUTPUT_FOLDER}")
    print(f"Total processing time: {time.perf_counter() - started:,.1f} seconds")
    print("Excel report generation completed successfully.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
