from datetime import datetime

import pandas as pd


LINEAGE_COLUMNS = [
    "job_id",
    "source_file",
    "source_file_hash",
    "processed_at",
    "parser_name",
    "parser_version",
    "loaded_from_cache",
]

BANK_LINEAGE_COLUMNS = LINEAGE_COLUMNS + ["bank_name", "input_format"]


def current_processed_at():
    return datetime.now().isoformat(timespec="seconds")


def add_lineage_columns(
    df,
    job_id,
    source_file,
    source_file_hash,
    processed_at,
    parser_name,
    parser_version,
    loaded_from_cache=False,
    extra=None,
):
    if df is None:
        df = pd.DataFrame()

    result = df.copy()
    result["job_id"] = job_id
    result["source_file"] = source_file
    result["source_file_hash"] = source_file_hash
    result["processed_at"] = processed_at
    result["parser_name"] = parser_name
    result["parser_version"] = parser_version
    result["loaded_from_cache"] = loaded_from_cache

    for key, value in (extra or {}).items():
        result[key] = value

    return result


def without_lineage_columns(df):
    if df is None:
        return df
    return df.drop(columns=[column for column in BANK_LINEAGE_COLUMNS if column in df.columns])
