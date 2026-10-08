from pathlib import Path
from datetime import datetime

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_FOLDER = PROJECT_ROOT / "Data"
HISTORY_FILE = DATA_FOLDER / "cycle_history.csv"

KEY_COLUMNS = [
    "SSO",
    "Qualification",
    "Cycle Start Date",
    "Cycle Due Date",
]

HISTORY_COLUMNS = [
    "SSO",
    "Qualification",
    "Cycle Number",
    "Cycle Start Date",
    "Cycle Due Date",
    "PM Completion Date",
    "Task Number",
    "Task Description",
    "Task Source File",
    "First Recorded Date",
    "Last Confirmed Date",
]


def _normalize_dates(dataframe):
    result = dataframe.copy()

    for column in ["Cycle Start Date", "Cycle Due Date"]:
        result[column] = pd.to_datetime(
            result[column], format="mixed", errors="coerce"
        ).dt.normalize()

    result["PM Completion Date"] = pd.to_datetime(
        result["PM Completion Date"], format="mixed", errors="coerce"
    )

    for column in ["First Recorded Date", "Last Confirmed Date"]:
        if column in result.columns:
            result[column] = pd.to_datetime(
                result[column], format="mixed", errors="coerce"
            ).dt.normalize()

    result["SSO"] = result["SSO"].astype(str).str.strip()
    result["Qualification"] = (
        result["Qualification"].astype(str).str.strip().str.upper()
    )
    return result


def load_cycle_history():
    if not HISTORY_FILE.exists():
        return pd.DataFrame(columns=HISTORY_COLUMNS)

    history = pd.read_csv(HISTORY_FILE, dtype={"SSO": str})

    missing = set(HISTORY_COLUMNS).difference(history.columns)
    if missing:
        raise ValueError(
            "Cycle history is missing columns: "
            + ", ".join(sorted(missing))
        )

    history = _normalize_dates(history[HISTORY_COLUMNS])

    if history[KEY_COLUMNS].isna().any().any():
        raise ValueError(
            "Cycle history contains an invalid SSO, qualification, "
            "cycle start date, or cycle due date."
        )

    duplicate_keys = history.duplicated(subset=KEY_COLUMNS, keep=False)
    if duplicate_keys.any():
        raise ValueError(
            "Cycle history contains duplicate cycle keys. "
            "Review Data/cycle_history.csv before continuing."
        )

    return history


def _key_from_record(record):
    return (
        str(record["SSO"]).strip(),
        str(record["Qualification"]).strip().upper(),
        pd.Timestamp(record["Cycle Start Date"]).normalize(),
        pd.Timestamp(record["Cycle Due Date"]).normalize(),
    )


def apply_cycle_history(results):
    """
    Restore historical completed-cycle evidence and prepare the next history state.

    Only Good / Completed cycles are stored permanently. Temporary statuses are
    always recalculated from current qualification dates, task inputs, and today.
    """
    current = results.copy()
    current["SSO"] = current["SSO"].astype(str).str.strip()
    current["Qualification"] = (
        current["Qualification"].astype(str).str.strip().str.upper()
    )
    current["Cycle Start Date"] = pd.to_datetime(
        current["Cycle Start Date"], format="mixed", errors="coerce"
    ).dt.normalize()
    current["Cycle Due Date"] = pd.to_datetime(
        current["Cycle Due Date"], format="mixed", errors="coerce"
    ).dt.normalize()
    current["PM Completion Date"] = pd.to_datetime(
        current["PM Completion Date"], format="mixed", errors="coerce"
    )

    history = load_cycle_history()
    history_by_key = {
        _key_from_record(record): record
        for record in history.to_dict("records")
    }

    source_completed_keys = set()
    restored_count = 0
    newer_source_count = 0

    for index, row in current.iterrows():
        key = _key_from_record(row)
        current_completed = (
            row["Cycle Result"] == "Good / Completed"
            and pd.notna(row["PM Completion Date"])
        )

        if current_completed:
            source_completed_keys.add(key)

        historical = history_by_key.get(key)
        if historical is None:
            continue

        historical_date = historical["PM Completion Date"]
        current_date = row["PM Completion Date"]

        use_history = (
            not current_completed
            or pd.isna(current_date)
            or historical_date > current_date
        )

        if use_history:
            current.at[index, "PM Completion Date"] = historical_date
            current.at[index, "Task Number"] = historical["Task Number"]
            current.at[index, "Task Description"] = historical[
                "Task Description"
            ]
            current.at[index, "Task Source File"] = historical[
                "Task Source File"
            ]
            current.at[index, "Task Data Available"] = "Historical Evidence"
            current.at[index, "Cycle Result"] = "Good / Completed"
            restored_count += 1
        elif current_completed and current_date > historical_date:
            newer_source_count += 1

    run_date = pd.Timestamp.today().normalize()
    existing_first_recorded = {
        key: record["First Recorded Date"]
        for key, record in history_by_key.items()
    }
    existing_last_confirmed = {
        key: record["Last Confirmed Date"]
        for key, record in history_by_key.items()
    }

    completed = current[
        (current["Cycle Result"] == "Good / Completed")
        & current["PM Completion Date"].notna()
    ].copy()

    history_rows = []
    for record in completed.to_dict("records"):
        key = _key_from_record(record)
        first_recorded = existing_first_recorded.get(key, run_date)

        if key in source_completed_keys:
            last_confirmed = run_date
        else:
            last_confirmed = existing_last_confirmed.get(key, first_recorded)

        history_rows.append(
            {
                "SSO": str(record["SSO"]).strip(),
                "Qualification": str(record["Qualification"]).strip().upper(),
                "Cycle Number": int(record["Cycle Number"]),
                "Cycle Start Date": record["Cycle Start Date"],
                "Cycle Due Date": record["Cycle Due Date"],
                "PM Completion Date": record["PM Completion Date"],
                "Task Number": record.get("Task Number", ""),
                "Task Description": record.get("Task Description", ""),
                "Task Source File": record.get("Task Source File", ""),
                "First Recorded Date": first_recorded,
                "Last Confirmed Date": last_confirmed,
            }
        )

    updated_history = pd.DataFrame(history_rows, columns=HISTORY_COLUMNS)
    updated_history = updated_history.sort_values(
        ["Qualification", "SSO", "Cycle Start Date"], kind="stable"
    )

    stats = {
        "History records loaded": len(history),
        "Completed cycles in updated history": len(updated_history),
        "Cycles restored from history": restored_count,
        "Cycles updated by newer source task": newer_source_count,
    }

    return current, updated_history, stats


def save_cycle_history(history):
    """Save history atomically only after report generation succeeds."""
    DATA_FOLDER.mkdir(parents=True, exist_ok=True)
    temporary_file = DATA_FOLDER / "cycle_history.tmp"

    history_to_save = history.copy()
    history_to_save.to_csv(
        temporary_file,
        index=False,
        date_format="%Y-%m-%d %H:%M:%S",
    )

    temporary_file.replace(HISTORY_FILE)
