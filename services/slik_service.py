from pathlib import Path

import pandas as pd

from excel_dashboard import prepare_dashboard_data
from extract_slik_text_to_excel import deduplicate_facilities, export_slik_excel_dashboard
from extractors.slik_extractor import (
    SLIK_PARSER_NAME,
    SLIK_PARSER_VERSION,
    extract_facilities,
)
from services.cache_service import load_cached_dataframe, save_cached_dataframe, cache_path
from services.error_report_service import normalize_error_rows, write_error_report
from services.etl_storage_service import persist_slik_layers
from services.logging_service import log_event
from storage.file_hash import compute_file_hash
from storage.lineage import add_lineage_columns, current_processed_at, without_lineage_columns
from storage.sqlite_metadata import (
    add_data_quality_result,
    add_error_log,
    add_job_file_record,
    create_job,
    find_successful_file_by_hash,
    update_job_status,
)
from transformers.slik_transformer import build_slik_dwh_df, build_slik_marts
from validators.slik_validator import validate_slik_facilities


def _final_status(success_count, error_count):
    if error_count and success_count:
        return "partial_success"
    if error_count:
        return "failed"
    return "success"


def _record_storage_errors(job_id, storage_result):
    for error in storage_result.get("errors", []):
        add_error_log(
            job_id=job_id,
            source_file_name=None,
            stage="storage",
            error_type=error.get("target"),
            error_message=error.get("error"),
        )


def _option(options, key, default=None):
    return (options or {}).get(key, default)


def _metadata_for_file(job_id, pdf_file, source_hash, processed_at, loaded_from_cache=False):
    return {
        "job_id": job_id,
        "source_file": Path(pdf_file).name,
        "source_file_hash": source_hash,
        "processed_at": processed_at,
        "parser_name": SLIK_PARSER_NAME,
        "parser_version": SLIK_PARSER_VERSION,
        "loaded_from_cache": loaded_from_cache,
    }


def _record_dq_results(job_id, dq_df):
    if dq_df is None or dq_df.empty:
        return
    for row in dq_df.to_dict("records"):
        add_data_quality_result(
            job_id=job_id,
            source_file_name=row.get("source_file"),
            table_name=row.get("table_name"),
            rule_name=row.get("rule_name"),
            severity=row.get("severity"),
            column_name=row.get("column_name"),
            issue_count=1,
            message=row.get("message"),
        )


def build_slik_data(pdf_files, progress_callback=None, etl_options=None):
    pdf_files = [Path(pdf_file) for pdf_file in pdf_files or []]
    use_cache = bool(_option(etl_options, "use_cache", True))
    incremental_mode = bool(_option(etl_options, "incremental_processing", True))
    force_reprocess = bool(_option(etl_options, "force_reprocess", False))
    use_multiprocessing = bool(_option(etl_options, "use_multiprocessing", False))
    max_workers = int(_option(etl_options, "max_workers", 4) or 4)
    cleanup_enabled = any(_option(etl_options, key, False) for key in ("clean_temp", "clean_uploads", "clean_cache"))

    job_id = create_job(
        "slik_pdf",
        input_file_count=len(pdf_files),
        notes=(
            f"SLIK PDF ETL layers; cache={use_cache}; incremental={incremental_mode}; "
            f"force_reprocess={force_reprocess}; multiprocessing={use_multiprocessing}"
        ),
    )
    log_event("job_started", job_id=job_id, job_type="slik_pdf", input_files=len(pdf_files))
    processed_at = current_processed_at()
    source_hashes = {}
    for pdf_file in pdf_files:
        try:
            source_hashes[str(pdf_file)] = compute_file_hash(pdf_file)
            log_event("file_hash_computed", job_id=job_id, source_file=pdf_file.name)
        except Exception as exc:
            source_hashes[str(pdf_file)] = None
            add_error_log(
                job_id=job_id,
                source_file_name=pdf_file.name,
                stage="lineage",
                error_type="FileHashError",
                error_message=str(exc),
            )
            log_event("file_hash_failed", job_id=job_id, source_file=pdf_file.name, error=exc)

    file_results = {}
    cached_frames = []
    cache_status = {}
    parse_files = []

    for pdf_file in pdf_files:
        source_hash = source_hashes.get(str(pdf_file))
        should_try_cache = bool(
            use_cache
            and incremental_mode
            and not force_reprocess
            and source_hash
            and find_successful_file_by_hash(
                source_hash,
                SLIK_PARSER_NAME,
                SLIK_PARSER_VERSION,
                document_type="slik_pdf",
            )
        )
        if should_try_cache:
            try:
                cached_df, cached_path = load_cached_dataframe("slik", SLIK_PARSER_VERSION, source_hash)
                if cached_df is not None:
                    cached_df = without_lineage_columns(cached_df)
                    cached_df = add_lineage_columns(
                        cached_df,
                        job_id=job_id,
                        source_file=pdf_file.name,
                        source_file_hash=source_hash,
                        processed_at=processed_at,
                        parser_name=SLIK_PARSER_NAME,
                        parser_version=SLIK_PARSER_VERSION,
                        loaded_from_cache=True,
                    )
                    cached_frames.append(cached_df)
                    cache_status[pdf_file.name] = {
                        "loaded_from_cache": True,
                        "cache_path": cached_path,
                        "row_count": len(cached_df),
                    }
                    file_results[str(pdf_file)] = {"row_count": len(cached_df), "error": None}
                    log_event("cache_hit", job_id=job_id, source_file=pdf_file.name, cache_path=cached_path)
                    continue
            except Exception as exc:
                add_error_log(job_id, pdf_file.name, "cache", "CacheLoadError", str(exc))
                log_event("cache_load_failed", job_id=job_id, source_file=pdf_file.name, error=exc)

        cache_status[pdf_file.name] = {
            "loaded_from_cache": False,
            "cache_path": cache_path("slik", SLIK_PARSER_VERSION, source_hash) if source_hash else None,
            "row_count": 0,
        }
        parse_files.append(pdf_file)
        log_event("cache_miss", job_id=job_id, source_file=pdf_file.name)

    def row_metadata_callback(pdf_file):
        pdf_file = Path(pdf_file)
        return _metadata_for_file(
            job_id,
            pdf_file,
            source_hashes.get(str(pdf_file)),
            processed_at,
            loaded_from_cache=False,
        )

    def etl_progress_callback(processed, total, file, rows_found=0, error=None):
        file_path = Path(file)
        file_results[str(file_path)] = {
            "row_count": rows_found,
            "error": error,
        }
        if progress_callback:
            progress_callback(
                processed=processed,
                total=total,
                file=file,
                rows_found=rows_found,
                error=error,
            )

    parsed_facilities_df = pd.DataFrame()
    parsed_duplicate_count = 0
    errors_df = pd.DataFrame(columns=["file", "error"])
    if parse_files:
        try:
            parsed_facilities_df, parsed_duplicate_count, errors_df = extract_facilities(
                parse_files,
                progress_callback=etl_progress_callback,
                row_metadata_callback=row_metadata_callback,
                use_multiprocessing=use_multiprocessing,
                max_workers=max_workers,
            )
        except Exception as exc:
            if use_multiprocessing:
                log_event("multiprocessing_failed_fallback_sequential", job_id=job_id, error=exc)
                parsed_facilities_df, parsed_duplicate_count, errors_df = extract_facilities(
                    parse_files,
                    progress_callback=etl_progress_callback,
                    row_metadata_callback=row_metadata_callback,
                    use_multiprocessing=False,
                    max_workers=max_workers,
                )
            else:
                raise

    all_frames = [frame for frame in [*cached_frames, parsed_facilities_df] if frame is not None and not frame.empty]
    if all_frames:
        combined_df = pd.concat(all_frames, ignore_index=True)
        before_dedupe_count = len(combined_df)
        facilities_df = deduplicate_facilities(combined_df)
        duplicate_count = parsed_duplicate_count + (before_dedupe_count - len(facilities_df))
    else:
        facilities_df = pd.DataFrame(columns=list(pd.DataFrame().columns))
        duplicate_count = parsed_duplicate_count

    if use_cache and not parsed_facilities_df.empty:
        for source_file, group in parsed_facilities_df.groupby("source_file", dropna=False):
            source_hash = None
            source_hash_values = group.get("source_file_hash")
            if source_hash_values is not None and not source_hash_values.dropna().empty:
                source_hash = source_hash_values.dropna().iloc[0]
            if not source_hash:
                continue
            try:
                cache_df = without_lineage_columns(group)
                saved_path = save_cached_dataframe(cache_df, "slik", SLIK_PARSER_VERSION, source_hash)
                cache_status[str(source_file)] = {
                    "loaded_from_cache": False,
                    "cache_path": saved_path,
                    "row_count": len(group),
                }
                log_event("cache_saved", job_id=job_id, source_file=source_file, cache_path=saved_path)
            except Exception as exc:
                add_error_log(job_id, source_file, "cache", "CacheSaveError", str(exc))
                log_event("cache_save_failed", job_id=job_id, source_file=source_file, error=exc)

    clean_df = prepare_dashboard_data(facilities_df)
    dwh_df = build_slik_dwh_df(facilities_df, clean_df)
    marts = build_slik_marts(dwh_df)
    data_quality_df = validate_slik_facilities(facilities_df)
    _record_dq_results(job_id, data_quality_df)
    log_event("validation_summary", job_id=job_id, issue_count=len(data_quality_df))
    storage_result = persist_slik_layers(
        facilities_df,
        dwh_df,
        marts,
        options=etl_options,
    )
    _record_storage_errors(job_id, storage_result)

    errors_by_file = {
        row["file"]: row["error"]
        for row in errors_df.to_dict("records")
    }
    row_count_by_file = {}
    if not facilities_df.empty and "source_file" in facilities_df.columns:
        row_count_by_file = facilities_df.groupby("source_file").size().to_dict()

    for pdf_file in pdf_files:
        progress_result = file_results.get(str(pdf_file), {})
        error_message = errors_by_file.get(pdf_file.name) or progress_result.get("error")
        file_cache_status = cache_status.get(pdf_file.name, {})
        add_job_file_record(
            job_id=job_id,
            source_file_name=pdf_file.name,
            source_file_path=pdf_file,
            source_file_hash=source_hashes.get(str(pdf_file)),
            document_type="slik_pdf",
            parser_name=SLIK_PARSER_NAME,
            parser_version=SLIK_PARSER_VERSION,
            status="error" if error_message else "success",
            row_count=row_count_by_file.get(pdf_file.name, progress_result.get("row_count", 0)),
            error_message=error_message,
            processed_at=processed_at,
            loaded_from_cache=file_cache_status.get("loaded_from_cache", False),
            cache_path=file_cache_status.get("cache_path"),
        )
        if error_message:
            add_error_log(
                job_id=job_id,
                source_file_name=pdf_file.name,
                stage="extract",
                error_type="SlikParseError",
                error_message=error_message,
            )

    error_count = len(errors_df)
    success_count = max(len(pdf_files) - error_count, 0)
    error_report_path = None
    if error_count:
        error_report_path = Path("outputs") / f"error_report_{job_id}.xlsx"
        write_error_report(
            normalize_error_rows(errors_df, job_id=job_id, stage="extract", error_type="SlikParseError"),
            error_report_path,
        )

    cache_used = any(status.get("loaded_from_cache") for status in cache_status.values())
    update_job_status(
        job_id,
        _final_status(success_count, error_count),
        input_file_count=len(pdf_files),
        success_file_count=success_count,
        error_file_count=error_count,
        output_path=storage_result["parquet_paths"][0] if storage_result["parquet_paths"] else None,
        notes="SLIK ETL completed",
        parser_version=SLIK_PARSER_VERSION,
        cache_used=cache_used,
        incremental_mode=incremental_mode,
        multiprocessing_used=use_multiprocessing,
        max_workers=max_workers if use_multiprocessing else None,
        cleanup_enabled=cleanup_enabled,
    )
    log_event(
        "job_finished",
        job_id=job_id,
        status=_final_status(success_count, error_count),
        success_files=success_count,
        error_files=error_count,
        cache_used=cache_used,
    )

    return {
        "job_id": job_id,
        "facilities_df": facilities_df,
        "clean_df": clean_df,
        "dwh_df": dwh_df,
        "marts": marts,
        "data_quality_df": data_quality_df,
        "duplicate_count": duplicate_count,
        "errors_df": errors_df,
        "error_report_path": str(error_report_path) if error_report_path else None,
        "storage": storage_result,
        "cache_used": cache_used,
    }


def export_dashboard_workbook(
    facilities_df,
    output_file,
    include_excel_dashboard=True,
    include_angsuran=True,
    data_quality_df=None,
    include_data_quality=False,
):
    return export_slik_excel_dashboard(
        facilities_df,
        output_file,
        include_excel_dashboard=include_excel_dashboard,
        include_angsuran=include_angsuran,
        data_quality_df=data_quality_df,
        include_data_quality=include_data_quality,
    )
