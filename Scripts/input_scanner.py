from pathlib import Path
import sys

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent.parent
QUALIFICATION_FOLDER = PROJECT_ROOT / "Input" / "Qualification"
TASK_FOLDER = PROJECT_ROOT / "Input" / "Tasks"
CONFIG_FILE = PROJECT_ROOT / "Config" / "qualification_mapping.csv"

SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}


def readable_size(size_bytes: int) -> str:
    size_mb = size_bytes / (1024 * 1024)

    if size_mb >= 1:
        return f"{size_mb:,.1f} MB"

    size_kb = size_bytes / 1024
    return f"{size_kb:,.1f} KB"


def find_excel_files(folder: Path) :
    files = []

    for file_path in folder.iterdir():
        if not file_path.is_file():
            continue

        if file_path.name.startswith("~$"):
            continue

        if file_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue

        files.append(file_path)

    return sorted(files, key=lambda path: path.name.lower())


def display_files(title: str, files: list[Path]) -> None:
    print()
    print(title)
    print("-" * len(title))

    if not files:
        print("No Excel reports found.")
        return

    for number, file_path in enumerate(files, start=1):
        print(
            f"{number}. {file_path.name} "
            f"({readable_size(file_path.stat().st_size)})"
        )


def main() -> int:
    print("=" * 55)
    print("PM Qualification Activity Tracker - Input Scanner")
    print("=" * 55)

    required_folders = [
        QUALIFICATION_FOLDER,
        TASK_FOLDER,
        PROJECT_ROOT / "Output",
        PROJECT_ROOT / "Archive",
        PROJECT_ROOT / "Logs",
    ]

    missing_folders = [
        folder for folder in required_folders if not folder.exists()
    ]

    if missing_folders:
        print()
        print("ERROR: Required folders are missing:")

        for folder in missing_folders:
            print(f"  - {folder}")

        return 1

    if not CONFIG_FILE.exists():
        print()
        print("ERROR: Configuration file not found:")
        print(f"  {CONFIG_FILE}")
        return 1

    mapping = pd.read_csv(CONFIG_FILE)

    required_columns = {
        "Area",
        "Venue",
        "Qualification",
        "Discipline",
        "Tracked",
    }

    missing_columns = required_columns.difference(mapping.columns)

    if missing_columns:
        print()
        print("ERROR: Configuration columns are missing:")

        for column in sorted(missing_columns):
            print(f"  - {column}")

        return 1

    tracked = mapping[
        mapping["Tracked"].astype(str).str.strip().str.lower() == "yes"
    ]

    excluded = mapping[
        mapping["Tracked"].astype(str).str.strip().str.lower() != "yes"
    ]

    qualification_files = find_excel_files(QUALIFICATION_FOLDER)
    task_files = find_excel_files(TASK_FOLDER)

    print()
    print(f"Project folder: {PROJECT_ROOT}")
    print(f"Configuration records: {len(mapping)}")
    print(f"Tracked qualifications: {len(tracked)}")
    print(f"Excluded qualifications: {len(excluded)}")

    display_files("Qualification Reports", qualification_files)
    display_files("Task Reports", task_files)

    print()
    print("Summary")
    print("-------")
    print(f"Qualification reports found: {len(qualification_files)}")
    print(f"Task reports found: {len(task_files)}")

    if not qualification_files:
        print("WARNING: No qualification reports are currently available.")

    if not task_files:
        print("WARNING: No task reports are currently available.")

    print()
    print("Input scan completed successfully.")
    return 0


if __name__ == "__main__":
    sys.exit(main())