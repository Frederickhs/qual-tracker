from pathlib import Path
import re
import sys
import warnings

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent.parent
QUALIFICATION_FOLDER = PROJECT_ROOT / "Input" / "Qualification"
CONFIG_FILE = PROJECT_ROOT / "Config" / "qualification_mapping.csv"
LOG_FOLDER = PROJECT_ROOT / "Logs"
SUMMARY_FILE = LOG_FOLDER / "qualification_summary.csv"

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
REQUIRED_REPORT_COLUMNS = {
    "Labor",
    "Name",
    "Effective Date",
    "Status",
}
REQUIRED_CONFIG_COLUMNS = {
    "Area",
    "Venue",
    "Qualification",
    "Discipline",
    "Tracked",
}

warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style*",
    category=UserWarning,
)


def get_qualification_files():
    files = [
    file_path
    for file_path in files
    if "MOD0" not in file_path.stem.upper()
]
    return sorted(
        [
            file_path
            for file_path in QUALIFICATION_FOLDER.iterdir()
            if file_path.is_file()
            and not file_path.name.startswith("~$")
            and file_path.suffix.lower() in SUPPORTED_EXTENSIONS
        ],
        key=lambda file_path: file_path.name.lower(),
    )


def qualification_from_filename(file_path):
    filename = file_path.stem.upper()

    for qualification in [
        "DUMOD1",
        "CUMOD1",
        "CUMOD2",
        "CUMOD3",
        "UCMOD1",
        "UCMOD2",
        "UCMOD3",
        "UCMOD4",
    ]:
        if qualification in filename:
            return qualification

    return None


def load_mapping():
    mapping = pd.read_csv(CONFIG_FILE, dtype=str).fillna("")
    mapping.columns = [str(column).strip() for column in mapping.columns]

    missing_columns = REQUIRED_CONFIG_COLUMNS.difference(mapping.columns)
    if missing_columns:
        missing = ", ".join(sorted(missing_columns))
        raise ValueError(f"Configuration is missing columns: {missing}")

    for column in REQUIRED_CONFIG_COLUMNS:
        mapping[column] = mapping[column].astype(str).str.strip()

    mapping["Qualification"] = mapping["Qualification"].str.upper()
    mapping["Tracked"] = mapping["Tracked"].str.lower()
    return mapping


def read_qualification_report(file_path, qualification, mapping_row):
    workbook = pd.ExcelFile(file_path, engine="openpyxl")

    if "List of Labor" in workbook.sheet_names:
        sheet_name = "List of Labor"
    elif len(workbook.sheet_names) == 1:
        sheet_name = workbook.sheet_names[0]
    else:
        available = ", ".join(workbook.sheet_names)
        raise ValueError(
            f"Expected 'List of Labor' worksheet. Available worksheets: {available}"
        )

    report = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        engine="openpyxl",
        dtype={"Labor": "string"},
    )
    report.columns = [str(column).strip() for column in report.columns]

    missing_columns = REQUIRED_REPORT_COLUMNS.difference(report.columns)
    if missing_columns:
        missing = ", ".join(sorted(missing_columns))
        raise ValueError(f"Missing required columns: {missing}")

    report = report.dropna(how="all").copy()
    report["Labor"] = report["Labor"].astype("string").str.strip()
    report["Name"] = report["Name"].astype("string").str.strip()
    report["Status"] = report["Status"].astype("string").str.strip().str.upper()
    report["Effective Date"] = pd.to_datetime(
    report["Effective Date"],
    format="mixed",
    errors="coerce",
)

    report["Qualification"] = qualification
    report["Area"] = mapping_row["Area"]
    report["Venue"] = mapping_row["Venue"]
    report["Discipline"] = mapping_row["Discipline"]
    report["Source File"] = file_path.name

    return report


def main():
    print("=" * 70)
    print("PM Qualification Activity Tracker - Qualification Reader")
    print("=" * 70)

    if not CONFIG_FILE.exists():
        print(f"ERROR: Configuration file not found: {CONFIG_FILE}")
        return 1

    files = get_qualification_files()
    print(f"Qualification reports found: {len(files)}")

    if not files:
        print("ERROR: No qualification reports were found.")
        return 1

    try:
        mapping = load_mapping()
    except Exception as error:
        print(f"ERROR: {error}")
        return 1

    tracked_mapping = mapping[mapping["Tracked"] == "yes"].copy()
    tracked_codes = set(tracked_mapping["Qualification"])

    reports = []
    errors = []
    detected_codes = []

    for file_path in files:
        qualification = qualification_from_filename(file_path)

        if qualification is None:
            errors.append(
                f"{file_path.name}: qualification code was not found in filename"
            )
            continue

        detected_codes.append(qualification)

        matching_rows = tracked_mapping[
            tracked_mapping["Qualification"] == qualification
        ]

        if matching_rows.empty:
            errors.append(
                f"{file_path.name}: {qualification} is not an approved tracked qualification"
            )
            continue

        if len(matching_rows) > 1:
            errors.append(
                f"{file_path.name}: duplicate configuration rows for {qualification}"
            )
            continue

        mapping_row = matching_rows.iloc[0]

        try:
            report = read_qualification_report(
                file_path,
                qualification,
                mapping_row,
            )
            reports.append(report)
            active_count = int((report["Status"] == "ACTIVE").sum())
            inactive_count = int((report["Status"] == "INACTIVE").sum())
            invalid_dates = int(report["Effective Date"].isna().sum())

            print()
            print(f"{file_path.name}")
            print(f"  Qualification: {qualification}")
            print(f"  Venue: {mapping_row['Venue']}")
            print(f"  Discipline: {mapping_row['Discipline']}")
            print(f"  Records: {len(report):,}")
            print(f"  Active: {active_count:,}")
            print(f"  Inactive: {inactive_count:,}")
            print(f"  Invalid qualification dates: {invalid_dates:,}")

        except Exception as error:
            errors.append(f"{file_path.name}: {error}")

    duplicate_input_codes = sorted(
        {
            code
            for code in detected_codes
            if detected_codes.count(code) > 1
        }
    )

    missing_tracked_codes = sorted(tracked_codes.difference(detected_codes))

    if duplicate_input_codes:
        errors.append(
            "Duplicate qualification report codes found: "
            + ", ".join(duplicate_input_codes)
        )

    if missing_tracked_codes:
        errors.append(
            "Missing tracked qualification reports: "
            + ", ".join(missing_tracked_codes)
        )

    print()
    print("Summary")
    print("-------")
    print(f"Reports successfully read: {len(reports)}")

    if reports:
        combined = pd.concat(reports, ignore_index=True)
        duplicate_keys = combined.duplicated(
            subset=["Labor", "Qualification"], keep=False
        )

        print(f"Qualification records combined: {len(combined):,}")
        print(f"Unique technicians: {combined['Labor'].nunique(dropna=True):,}")
        print(
            "Duplicate SSO and qualification rows: "
            f"{int(duplicate_keys.sum()):,}"
        )

        summary = (
            combined.groupby(
                ["Area", "Venue", "Qualification", "Discipline"],
                dropna=False,
            )
            .agg(
                Records=("Labor", "size"),
                Unique_Technicians=("Labor", "nunique"),
                Active=("Status", lambda values: int((values == "ACTIVE").sum())),
                Inactive=("Status", lambda values: int((values == "INACTIVE").sum())),
                Invalid_Dates=("Effective Date", lambda values: int(values.isna().sum())),
            )
            .reset_index()
        )

        LOG_FOLDER.mkdir(parents=True, exist_ok=True)
        summary.to_csv(SUMMARY_FILE, index=False)
        print(f"Summary file created: {SUMMARY_FILE}")

    if errors:
        print()
        print("ERRORS")
        print("------")
        for error in errors:
            print(f"- {error}")
        return 1

    print()
    print("Qualification report reading completed successfully.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
