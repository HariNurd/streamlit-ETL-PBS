from pathlib import Path

from services.error_report_service import normalize_error_rows, write_error_report
from extractors.bank_statement_extractor import (
    BANK_PDF_PARSERS,
    DKI_TXT_PARSER_NAME,
    DKI_TXT_PARSER_VERSION,
    extract_dki_txt_transactions,
    extract_pdf_transactions,
)
from services.etl_storage_service import persist_bank_statement_layers
from services.logging_service import log_event
from storage.file_hash import compute_file_hash
from storage.lineage import current_processed_at
from storage.sqlite_metadata import (
    add_data_quality_result,
    add_error_log,
    add_job_file_record,
    create_job,
    update_job_status,
)
from transformers.bank_statement_transformer import (
    build_bank_statement_dwh_df,
    build_bank_statement_marts,
)
from validators.bank_statement_validator import validate_bank_transactions


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


def build_bank_pdf_etl(file_items, etl_options=None, progress_callback=None):
    file_items = [
        {"path": Path(item["path"]), "bank": item["bank"]}
        for item in file_items or []
    ]
    job_id = create_job(
        "bank_statement_pdf",
        input_file_count=len(file_items),
        notes=(
            "PDF bank statement ETL layers; "
            f"cache={_option(etl_options, 'use_cache', True)}; "
            f"incremental={_option(etl_options, 'incremental_processing', True)}; "
            f"force_reprocess={_option(etl_options, 'force_reprocess', False)}"
        ),
    )
    log_event("job_started", job_id=job_id, job_type="bank_statement_pdf", input_files=len(file_items))
    processed_at = current_processed_at()

    result = extract_pdf_transactions(
        file_items,
        job_id=job_id,
        progress_callback=progress_callback,
        use_cache=bool(_option(etl_options, "use_cache", True)),
        incremental_processing=bool(_option(etl_options, "incremental_processing", True)),
        force_reprocess=bool(_option(etl_options, "force_reprocess", False)),
    )
    errors_df = result["errors_df"]
    error_by_file = {
        row["file"]: row["error"]
        for row in errors_df.to_dict("records")
    }
    rows_by_file = {}
    transactions_df = result["transactions_df"]
    file_statuses = result.get("file_statuses", {})
    if not transactions_df.empty and "source_file" in transactions_df.columns:
        rows_by_file = transactions_df.groupby("source_file").size().to_dict()

    for item in file_items:
        path = item["path"]
        bank_name = item["bank"]
        parser_name = BANK_PDF_PARSERS[bank_name][2] if bank_name in BANK_PDF_PARSERS else None
        parser_version = BANK_PDF_PARSERS[bank_name][3] if bank_name in BANK_PDF_PARSERS else None
        error_message = error_by_file.get(path.name)
        file_status = file_statuses.get(path.name, {})
        try:
            source_hash = file_status.get("source_file_hash") or compute_file_hash(path)
        except Exception as exc:
            source_hash = None
            error_message = error_message or str(exc)

        add_job_file_record(
            job_id=job_id,
            source_file_name=path.name,
            source_file_path=path,
            source_file_hash=source_hash,
            document_type="bank_statement_pdf",
            parser_name=parser_name,
            parser_version=parser_version,
            status="error" if error_message else "success",
            row_count=rows_by_file.get(path.name, 0),
            error_message=error_message,
            processed_at=processed_at,
            loaded_from_cache=file_status.get("loaded_from_cache", False),
            cache_path=file_status.get("cache_path"),
        )
        if error_message:
            add_error_log(
                job_id=job_id,
                source_file_name=path.name,
                stage="extract",
                error_type="BankStatementPdfParseError",
                error_message=error_message,
            )

    dwh_df = build_bank_statement_dwh_df(transactions_df)
    marts = build_bank_statement_marts(dwh_df)
    data_quality_df = validate_bank_transactions(transactions_df, result["summaries_df"])
    _record_dq_results(job_id, data_quality_df)
    log_event("validation_summary", job_id=job_id, issue_count=len(data_quality_df))
    storage_result = persist_bank_statement_layers(
        transactions_df,
        dwh_df,
        marts,
        options=etl_options,
    )
    _record_storage_errors(job_id, storage_result)

    error_count = len(errors_df)
    success_count = max(len(file_items) - error_count, 0)
    error_report_path = None
    if error_count:
        error_report_path = Path("outputs") / f"error_report_{job_id}.xlsx"
        write_error_report(
            normalize_error_rows(errors_df, job_id=job_id, stage="extract", error_type="BankStatementPdfParseError"),
            error_report_path,
        )
    cache_used = any(status.get("loaded_from_cache") for status in file_statuses.values())
    update_job_status(
        job_id,
        _final_status(success_count, error_count),
        input_file_count=len(file_items),
        success_file_count=success_count,
        error_file_count=error_count,
        output_path=storage_result["parquet_paths"][0] if storage_result["parquet_paths"] else None,
        notes="Bank statement PDF ETL completed",
        parser_version=",".join(sorted({BANK_PDF_PARSERS[item["bank"]][3] for item in file_items if item["bank"] in BANK_PDF_PARSERS})),
        cache_used=cache_used,
        incremental_mode=bool(_option(etl_options, "incremental_processing", True)),
        cleanup_enabled=any(_option(etl_options, key, False) for key in ("clean_temp", "clean_uploads", "clean_cache")),
    )
    log_event("job_finished", job_id=job_id, status=_final_status(success_count, error_count), cache_used=cache_used)

    return {
        "job_id": job_id,
        "transactions_df": transactions_df,
        "summaries_df": result["summaries_df"],
        "errors_df": errors_df,
        "data_quality_df": data_quality_df,
        "error_report_path": str(error_report_path) if error_report_path else None,
        "dwh_df": dwh_df,
        "marts": marts,
        "storage": storage_result,
        "cache_used": cache_used,
    }


def build_dki_txt_etl(txt_files, etl_options=None, progress_callback=None):
    txt_files = [Path(txt_file) for txt_file in txt_files or []]
    job_id = create_job(
        "bank_statement_txt",
        input_file_count=len(txt_files),
        notes=(
            "DKI TXT bank statement ETL layers; "
            f"cache={_option(etl_options, 'use_cache', True)}; "
            f"incremental={_option(etl_options, 'incremental_processing', True)}; "
            f"force_reprocess={_option(etl_options, 'force_reprocess', False)}"
        ),
    )
    log_event("job_started", job_id=job_id, job_type="bank_statement_txt", input_files=len(txt_files))
    processed_at = current_processed_at()

    result = extract_dki_txt_transactions(
        txt_files,
        job_id=job_id,
        progress_callback=progress_callback,
        use_cache=bool(_option(etl_options, "use_cache", True)),
        incremental_processing=bool(_option(etl_options, "incremental_processing", True)),
        force_reprocess=bool(_option(etl_options, "force_reprocess", False)),
    )
    errors_df = result["errors_df"]
    error_by_file = {
        row["file"]: row["error"]
        for row in errors_df.to_dict("records")
    }
    transactions_df = result["transactions_df"]
    file_statuses = result.get("file_statuses", {})
    rows_by_file = {}
    if not transactions_df.empty and "source_file" in transactions_df.columns:
        rows_by_file = transactions_df.groupby("source_file").size().to_dict()

    for path in txt_files:
        error_message = error_by_file.get(path.name)
        file_status = file_statuses.get(path.name, {})
        try:
            source_hash = file_status.get("source_file_hash") or compute_file_hash(path)
        except Exception as exc:
            source_hash = None
            error_message = error_message or str(exc)

        add_job_file_record(
            job_id=job_id,
            source_file_name=path.name,
            source_file_path=path,
            source_file_hash=source_hash,
            document_type="bank_statement_txt",
            parser_name=DKI_TXT_PARSER_NAME,
            parser_version=DKI_TXT_PARSER_VERSION,
            status="error" if error_message else "success",
            row_count=rows_by_file.get(path.name, 0),
            error_message=error_message,
            processed_at=processed_at,
            loaded_from_cache=file_status.get("loaded_from_cache", False),
            cache_path=file_status.get("cache_path"),
        )
        if error_message:
            add_error_log(
                job_id=job_id,
                source_file_name=path.name,
                stage="extract",
                error_type="BankStatementTxtParseError",
                error_message=error_message,
            )

    dwh_df = build_bank_statement_dwh_df(transactions_df)
    marts = build_bank_statement_marts(dwh_df)
    data_quality_df = validate_bank_transactions(transactions_df, result["summaries_df"])
    _record_dq_results(job_id, data_quality_df)
    log_event("validation_summary", job_id=job_id, issue_count=len(data_quality_df))
    storage_result = persist_bank_statement_layers(
        transactions_df,
        dwh_df,
        marts,
        options=etl_options,
    )
    _record_storage_errors(job_id, storage_result)

    error_count = len(errors_df)
    success_count = max(len(txt_files) - error_count, 0)
    error_report_path = None
    if error_count:
        error_report_path = Path("outputs") / f"error_report_{job_id}.xlsx"
        write_error_report(
            normalize_error_rows(errors_df, job_id=job_id, stage="extract", error_type="BankStatementTxtParseError"),
            error_report_path,
        )
    cache_used = any(status.get("loaded_from_cache") for status in file_statuses.values())
    update_job_status(
        job_id,
        _final_status(success_count, error_count),
        input_file_count=len(txt_files),
        success_file_count=success_count,
        error_file_count=error_count,
        output_path=storage_result["parquet_paths"][0] if storage_result["parquet_paths"] else None,
        notes="DKI TXT ETL completed",
        parser_version=DKI_TXT_PARSER_VERSION,
        cache_used=cache_used,
        incremental_mode=bool(_option(etl_options, "incremental_processing", True)),
        cleanup_enabled=any(_option(etl_options, key, False) for key in ("clean_temp", "clean_uploads", "clean_cache")),
    )
    log_event("job_finished", job_id=job_id, status=_final_status(success_count, error_count), cache_used=cache_used)

    return {
        "job_id": job_id,
        "transactions_df": transactions_df,
        "summaries_df": result["summaries_df"],
        "errors_df": errors_df,
        "data_quality_df": data_quality_df,
        "error_report_path": str(error_report_path) if error_report_path else None,
        "dwh_df": dwh_df,
        "marts": marts,
        "storage": storage_result,
        "cache_used": cache_used,
    }
