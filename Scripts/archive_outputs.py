from datetime import datetime
from pathlib import Path
import shutil
import sys


PROJECT_ROOT = Path(__file__).resolve().parent.parent
OUTPUT_FOLDER = PROJECT_ROOT / "Output"
ARCHIVE_FOLDER = PROJECT_ROOT / "Archive"

OUTPUT_FILES = [
    "Wolf Qualification Tracker.xlsx",
    "Darkmoor Qualification Tracker.xlsx",
    "Unchained Qualification Tracker.xlsx",
    "Animation Qualification Tracker.xlsx",
]


def existing_output_files():
    return [
        OUTPUT_FOLDER / filename
        for filename in OUTPUT_FILES
        if (OUTPUT_FOLDER / filename).exists()
    ]


def create_unique_archive_folder():
    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    archive_run = ARCHIVE_FOLDER / timestamp / "Output"

    counter = 1
    while archive_run.exists():
        archive_run = ARCHIVE_FOLDER / f"{timestamp}_{counter}" / "Output"
        counter += 1

    return archive_run


def move_with_rollback(files, destination):
    moved = []

    try:
        destination.mkdir(parents=True, exist_ok=False)

        for source in files:
            target = destination / source.name
            shutil.move(str(source), str(target))
            moved.append((source, target))

    except Exception:
        for original, archived in reversed(moved):
            if archived.exists() and not original.exists():
                shutil.move(str(archived), str(original))

        try:
            if destination.exists() and not any(destination.iterdir()):
                destination.rmdir()
                parent = destination.parent
                if parent.exists() and not any(parent.iterdir()):
                    parent.rmdir()
        except OSError:
            pass

        raise


def main():
    print("Checking for previous output reports...")

    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    ARCHIVE_FOLDER.mkdir(parents=True, exist_ok=True)

    files = existing_output_files()

    if not files:
        print("No previous tracker reports were found. Nothing to archive.")
        return 0

    archive_output = create_unique_archive_folder()

    try:
        move_with_rollback(files, archive_output)
    except PermissionError:
        print()
        print("ERROR: A previous output workbook is open or locked.")
        print(
            "Close all Wolf, Darkmoor, Unchained, and Animation "
            "output workbooks, then run again."
        )
        return 1
    except Exception as error:
        print()
        print(f"ERROR: Previous output reports could not be archived: {error}")
        print("Any files already moved were returned to Output when possible.")
        return 1

    print(f"Archived {len(files)} previous report(s) to:")
    print(f"  {archive_output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
