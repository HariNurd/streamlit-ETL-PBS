from datetime import datetime
from pathlib import Path

import pandas as pd


ERROR_COLUMNS = [
    "job_id",
    "source_file",
    "stage",
    "error_type",
    "error_message",
    "timestamp",
]


def normalize_error_rows(errors, job_id=None, stage="extract", error_type="ProcessingError"):
    rows = []
    if errors is None:
        return pd.DataFrame(columns=ERROR_COLUMNS)

    records = errors.to_dict("records") if isinstance(errors, pd.DataFrame) else list(errors or [])
    timestamp = datetime.now().isoformat(timespec="seconds")
    for record in records:
        if not isinstance(record, dict):
            continue
        rows.append(
            {
                "job_id": record.get("job_id") or job_id,
                "source_file": record.get("source_file") or record.get("file") or record.get("source_file_name"),
                "stage": record.get("stage") or stage,
                "error_type": record.get("error_type") or error_type,
                "error_message": record.get("error_message") or record.get("error"),
                "timestamp": record.get("timestamp") or record.get("created_at") or timestamp,
            }
        )

    return pd.DataFrame(rows, columns=ERROR_COLUMNS)


def write_error_report(errors, output_path):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    errors_df = normalize_error_rows(errors)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        errors_df.to_excel(writer, sheet_name="Error_Report", index=False)
    return output_path


def append_data_quality_sheet(workbook_path, data_quality_df, sheet_name="Data_Quality"):
    workbook_path = Path(workbook_path)
    if data_quality_df is None or data_quality_df.empty or not workbook_path.exists():
        return workbook_path

    with pd.ExcelWriter(workbook_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        data_quality_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return workbook_path
