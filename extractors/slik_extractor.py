from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

import pandas as pd

from extract_slik_text_to_excel import (
    FACILITY_DATA_COLUMNS,
    OUTPUT_COLUMNS,
    count_real_facility_rows,
    deduplicate_facilities,
    extract_slik_facilities,
    parse_slik_pdf,
)


SLIK_PARSER_NAME = "extract_slik_text_to_excel.extract_slik_facilities"
SLIK_PARSER_VERSION = "1.2.0"


def _add_metadata(df, metadata):
    if metadata:
        for column, value in metadata.items():
            df[column] = value
    return df


def _parse_slik_worker(pdf_file):
    pdf_file = Path(pdf_file)
    try:
        df = parse_slik_pdf(pdf_file)
        return {
            "file": str(pdf_file),
            "records": df.to_dict("records"),
            "columns": list(df.columns),
            "rows_found": count_real_facility_rows(df),
            "error": None,
        }
    except Exception as exc:
        return {
            "file": str(pdf_file),
            "records": [],
            "columns": OUTPUT_COLUMNS,
            "rows_found": 0,
            "error": str(exc),
        }


def extract_slik_facilities_parallel(pdf_files, max_workers=4, progress_callback=None, row_metadata_callback=None):
    pdf_files = [Path(pdf_file) for pdf_file in pdf_files or []]
    all_results = []
    errors = []
    completed = 0

    if not pdf_files:
        return pd.DataFrame(columns=FACILITY_DATA_COLUMNS), 0, pd.DataFrame(columns=["file", "error"])

    with ProcessPoolExecutor(max_workers=max(1, int(max_workers or 1))) as executor:
        future_map = {executor.submit(_parse_slik_worker, str(pdf_file)): pdf_file for pdf_file in pdf_files}
        for future in as_completed(future_map):
            pdf_file = future_map[future]
            completed += 1
            try:
                result = future.result()
            except Exception as exc:
                result = {
                    "file": str(pdf_file),
                    "records": [],
                    "columns": OUTPUT_COLUMNS,
                    "rows_found": 0,
                    "error": str(exc),
                }

            if result["error"]:
                errors.append({"file": pdf_file.name, "error": result["error"]})
                if progress_callback:
                    progress_callback(
                        processed=completed,
                        total=len(pdf_files),
                        file=pdf_file,
                        rows_found=0,
                        error=result["error"],
                    )
                continue

            df = pd.DataFrame(result["records"], columns=result["columns"])
            if row_metadata_callback:
                df = _add_metadata(df, row_metadata_callback(pdf_file) or {})
            if not df.empty:
                all_results.append(df)

            if progress_callback:
                progress_callback(
                    processed=completed,
                    total=len(pdf_files),
                    file=pdf_file,
                    rows_found=result["rows_found"],
                    error=None,
                )

    errors_df = pd.DataFrame(errors, columns=["file", "error"])
    if not all_results:
        return pd.DataFrame(columns=FACILITY_DATA_COLUMNS), 0, errors_df

    facilities_df = pd.concat(all_results, ignore_index=True)
    before_dedupe_count = len(facilities_df)
    facilities_df = deduplicate_facilities(facilities_df)
    duplicate_count = before_dedupe_count - len(facilities_df)

    return facilities_df, duplicate_count, errors_df


def extract_facilities(
    pdf_files,
    progress_callback=None,
    row_metadata_callback=None,
    use_multiprocessing=False,
    max_workers=4,
):
    if use_multiprocessing:
        return extract_slik_facilities_parallel(
            pdf_files,
            max_workers=max_workers,
            progress_callback=progress_callback,
            row_metadata_callback=row_metadata_callback,
        )

    return extract_slik_facilities(
        pdf_files,
        progress_callback=progress_callback,
        row_metadata_callback=row_metadata_callback,
    )
