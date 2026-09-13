from importlib import import_module
from pathlib import Path

import pandas as pd

from services.cache_service import cache_path, load_cached_dataframe, save_cached_dataframe
from storage.file_hash import compute_file_hash
from storage.lineage import add_lineage_columns, current_processed_at, without_lineage_columns
from storage.sqlite_metadata import find_successful_file_by_hash


BANK_PDF_PARSERS = {
    "BCA": ("convert_mutasi_bca", "process_pdf", "convert_mutasi_bca.process_pdf", "2.0.0"),
    "BNI": ("convert_mutasi_bni", "process_pdf", "convert_mutasi_bni.process_pdf", "2.0.0"),
    "DKI": ("convert_mutasi_dki", "process_pdf", "convert_mutasi_dki.process_pdf", "2.0.0"),
    "Mandiri": ("convert_mutasi_mandiri", "process_pdf", "convert_mutasi_mandiri.process_pdf", "2.0.0"),
    "BRI": ("convert_mutasi_BRI", "process_pdf", "convert_mutasi_BRI.process_pdf", "2.0.1"),
}

DKI_TXT_PARSER_NAME = "convert_mutasi_dki_txt.parse_txt_file"
DKI_TXT_PARSER_VERSION = "2.0.0"


def _empty_errors():
    return pd.DataFrame(columns=["file", "bank", "input_format", "error"])


def _get_pdf_parser(bank_name):
    if bank_name not in BANK_PDF_PARSERS:
        raise ValueError(f"Bank belum didukung untuk ETL: {bank_name}")

    module_name, function_name, parser_name, parser_version = BANK_PDF_PARSERS[bank_name]
    module = import_module(module_name)
    return getattr(module, function_name), parser_name, parser_version


def extract_pdf_transactions(
    file_items,
    job_id=None,
    progress_callback=None,
    use_cache=True,
    incremental_processing=True,
    force_reprocess=False,
):
    """
    Extract transaction dataframes from existing per-bank PDF parsers.

    file_items accepts dictionaries with:
    - path: source PDF path
    - bank: detected or selected bank name
    """
    file_items = list(file_items or [])
    processed_at = current_processed_at()
    transaction_frames = []
    summary_frames = []
    errors = []
    file_statuses = {}

    for index, item in enumerate(file_items, start=1):
        pdf_path = Path(item["path"])
        bank_name = item["bank"]
        try:
            parser, parser_name, parser_version = _get_pdf_parser(bank_name)
            source_hash = compute_file_hash(pdf_path)
            cache_file = cache_path("bank_pdf", parser_version, source_hash)
            cache_allowed = bool(
                use_cache
                and incremental_processing
                and not force_reprocess
                and find_successful_file_by_hash(
                    source_hash,
                    parser_name,
                    parser_version,
                    document_type="bank_statement_pdf",
                )
            )
            if cache_allowed:
                try:
                    cached_df, cached_path = load_cached_dataframe("bank_pdf", parser_version, source_hash)
                    if cached_df is not None:
                        cached_df = add_lineage_columns(
                            without_lineage_columns(cached_df),
                            job_id=job_id,
                            source_file=pdf_path.name,
                            source_file_hash=source_hash,
                            processed_at=processed_at,
                            parser_name=parser_name,
                            parser_version=parser_version,
                            loaded_from_cache=True,
                            extra={"bank_name": bank_name, "input_format": "PDF"},
                        )
                        if not cached_df.empty:
                            transaction_frames.append(cached_df)
                        file_statuses[pdf_path.name] = {
                            "source_file_hash": source_hash,
                            "loaded_from_cache": True,
                            "cache_path": cached_path,
                            "row_count": len(cached_df),
                        }
                        if progress_callback:
                            progress_callback(
                                processed=index,
                                total=len(file_items),
                                file=pdf_path,
                                bank=bank_name,
                                rows_found=len(cached_df),
                                error=None,
                            )
                        continue
                except Exception:
                    pass

            transactions_df, summary_df = parser(pdf_path)

            extra = {"bank_name": bank_name, "input_format": "PDF"}
            transactions_df = add_lineage_columns(
                transactions_df,
                job_id=job_id,
                source_file=pdf_path.name,
                source_file_hash=source_hash,
                processed_at=processed_at,
                parser_name=parser_name,
                parser_version=parser_version,
                loaded_from_cache=False,
                extra=extra,
            )
            summary_df = add_lineage_columns(
                summary_df,
                job_id=job_id,
                source_file=pdf_path.name,
                source_file_hash=source_hash,
                processed_at=processed_at,
                parser_name=parser_name,
                parser_version=parser_version,
                loaded_from_cache=False,
                extra=extra,
            )
            try:
                save_cached_dataframe(
                    without_lineage_columns(transactions_df),
                    "bank_pdf",
                    parser_version,
                    source_hash,
                )
            except Exception:
                pass

            if not transactions_df.empty:
                transaction_frames.append(transactions_df)
            if not summary_df.empty:
                summary_frames.append(summary_df)

            if progress_callback:
                progress_callback(
                    processed=index,
                    total=len(file_items),
                    file=pdf_path,
                    bank=bank_name,
                    rows_found=len(transactions_df),
                    error=None,
                )
            file_statuses[pdf_path.name] = {
                "source_file_hash": source_hash,
                "loaded_from_cache": False,
                "cache_path": cache_file,
                "row_count": len(transactions_df),
            }
        except Exception as exc:
            errors.append(
                {
                    "file": pdf_path.name,
                    "bank": bank_name,
                    "input_format": "PDF",
                    "error": str(exc),
                }
            )
            if progress_callback:
                progress_callback(
                    processed=index,
                    total=len(file_items),
                    file=pdf_path,
                    bank=bank_name,
                    rows_found=0,
                    error=str(exc),
                )
            file_statuses[pdf_path.name] = {
                "source_file_hash": None,
                "loaded_from_cache": False,
                "cache_path": None,
                "row_count": 0,
            }

    return {
        "transactions_df": (
            pd.concat(transaction_frames, ignore_index=True)
            if transaction_frames
            else pd.DataFrame()
        ),
        "summaries_df": (
            pd.concat(summary_frames, ignore_index=True)
            if summary_frames
            else pd.DataFrame()
        ),
        "errors_df": pd.DataFrame(errors, columns=_empty_errors().columns),
        "file_statuses": file_statuses,
    }


def extract_dki_txt_transactions(
    txt_files,
    job_id=None,
    progress_callback=None,
    use_cache=True,
    incremental_processing=True,
    force_reprocess=False,
):
    from convert_mutasi_dki_txt import parse_txt_file, reconcile_transactions_with_balance

    txt_files = [Path(txt_file) for txt_file in txt_files or []]
    processed_at = current_processed_at()
    transaction_frames = []
    summary_rows = []
    errors = []
    file_statuses = {}

    for index, txt_file in enumerate(txt_files, start=1):
        try:
            source_hash = compute_file_hash(txt_file)
            cache_file = cache_path("bank_txt", DKI_TXT_PARSER_VERSION, source_hash)
            cache_allowed = bool(
                use_cache
                and incremental_processing
                and not force_reprocess
                and find_successful_file_by_hash(
                    source_hash,
                    DKI_TXT_PARSER_NAME,
                    DKI_TXT_PARSER_VERSION,
                    document_type="bank_statement_txt",
                )
            )
            if cache_allowed:
                try:
                    cached_df, cached_path = load_cached_dataframe("bank_txt", DKI_TXT_PARSER_VERSION, source_hash)
                    if cached_df is not None:
                        cached_df = add_lineage_columns(
                            without_lineage_columns(cached_df),
                            job_id=job_id,
                            source_file=txt_file.name,
                            source_file_hash=source_hash,
                            processed_at=processed_at,
                            parser_name=DKI_TXT_PARSER_NAME,
                            parser_version=DKI_TXT_PARSER_VERSION,
                            loaded_from_cache=True,
                            extra={"bank_name": "DKI", "input_format": "TXT"},
                        )
                        if not cached_df.empty:
                            transaction_frames.append(cached_df)
                        file_statuses[txt_file.name] = {
                            "source_file_hash": source_hash,
                            "loaded_from_cache": True,
                            "cache_path": cached_path,
                            "row_count": len(cached_df),
                        }
                        if progress_callback:
                            progress_callback(
                                processed=index,
                                total=len(txt_files),
                                file=txt_file,
                                bank="DKI",
                                rows_found=len(cached_df),
                                error=None,
                            )
                        continue
                except Exception:
                    pass

            df_raw, accounts = parse_txt_file(txt_file)
            reconciled_frames = []
            if not df_raw.empty:
                for account, account_df in df_raw.groupby("Account", sort=True):
                    account_info = accounts.get(account, {"saldo_awal": pd.NA, "saldo_akhir": pd.NA})
                    reconciled_frames.append(
                        reconcile_transactions_with_balance(account_df, account_info)
                    )
            transactions_df = (
                pd.concat(reconciled_frames, ignore_index=True)
                if reconciled_frames
                else pd.DataFrame(columns=df_raw.columns)
            )
            transactions_df = add_lineage_columns(
                transactions_df,
                job_id=job_id,
                source_file=txt_file.name,
                source_file_hash=source_hash,
                processed_at=processed_at,
                parser_name=DKI_TXT_PARSER_NAME,
                parser_version=DKI_TXT_PARSER_VERSION,
                loaded_from_cache=False,
                extra={"bank_name": "DKI", "input_format": "TXT"},
            )
            try:
                save_cached_dataframe(
                    without_lineage_columns(transactions_df),
                    "bank_txt",
                    DKI_TXT_PARSER_VERSION,
                    source_hash,
                )
            except Exception:
                pass

            for account, account_info in accounts.items():
                parser_diagnostics = [
                    issue.get("message")
                    for issue in account_info.get("validation_issues", [])
                    if issue.get("message")
                ]
                validation_messages = [
                    issue.get("message")
                    for issue in account_info.get("validation_issues", [])
                    if issue.get("message") and issue.get("severity") != "INFO"
                ]
                summary_rows.append(
                    {
                        "Account": account,
                        "Saldo Awal": account_info.get("saldo_awal", pd.NA),
                        "Saldo Akhir": account_info.get("saldo_akhir", pd.NA),
                        "Periode Awal": account_info.get("period_start"),
                        "Periode Akhir": account_info.get("period_end"),
                        "Statement Blocks": account_info.get("statement_block_count", 1),
                        "Duplicates Removed": account_info.get("duplicates_removed", 0),
                        "Parser Diagnostics": " | ".join(parser_diagnostics),
                        "Validation Issues": " | ".join(validation_messages),
                        "source_file": txt_file.name,
                        "source_file_hash": source_hash,
                        "processed_at": processed_at,
                        "parser_name": DKI_TXT_PARSER_NAME,
                        "parser_version": DKI_TXT_PARSER_VERSION,
                        "job_id": job_id,
                        "bank_name": "DKI",
                        "input_format": "TXT",
                    }
                )

            if not transactions_df.empty:
                transaction_frames.append(transactions_df)

            if progress_callback:
                progress_callback(
                    processed=index,
                    total=len(txt_files),
                    file=txt_file,
                    bank="DKI",
                    rows_found=len(transactions_df),
                    error=None,
                )
            file_statuses[txt_file.name] = {
                "source_file_hash": source_hash,
                "loaded_from_cache": False,
                "cache_path": cache_file,
                "row_count": len(transactions_df),
            }
        except Exception as exc:
            errors.append(
                {
                    "file": txt_file.name,
                    "bank": "DKI",
                    "input_format": "TXT",
                    "error": str(exc),
                }
            )
            if progress_callback:
                progress_callback(
                    processed=index,
                    total=len(txt_files),
                    file=txt_file,
                    bank="DKI",
                    rows_found=0,
                    error=str(exc),
                )
            file_statuses[txt_file.name] = {
                "source_file_hash": None,
                "loaded_from_cache": False,
                "cache_path": None,
                "row_count": 0,
            }

    return {
        "transactions_df": (
            pd.concat(transaction_frames, ignore_index=True)
            if transaction_frames
            else pd.DataFrame()
        ),
        "summaries_df": pd.DataFrame(summary_rows),
        "errors_df": pd.DataFrame(errors, columns=_empty_errors().columns),
        "file_statuses": file_statuses,
    }
