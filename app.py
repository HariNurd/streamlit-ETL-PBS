import json
import uuid
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from converters.bank_detector import detect_bank_from_pdf
from dashboards.streamlit_slik_dashboard import render_slik_dashboard
from services.bank_statement_etl_service import build_bank_pdf_etl, build_dki_txt_etl
from services.etl_storage_service import normalize_etl_options
from services.error_report_service import (
    append_data_quality_sheet,
    normalize_error_rows,
    write_error_report,
)
from services.file_utils import (
    cleanup_paths,
    cleanup_working_dirs,
    combine_cleanup_summaries,
    create_session_dir,
    log_cleanup_summary,
    safe_filename,
    save_uploaded_files,
    zip_files,
)
from services.slik_service import build_slik_data, export_dashboard_workbook
from storage.lineage import without_lineage_columns
from storage.paths import ensure_project_dirs
from storage.sqlite_metadata import add_cleanup_log, get_job_history


APP_DIRS = {
    "uploads": Path("uploads"),
    "outputs": Path("outputs"),
    "cache": Path("cache"),
    "temp": Path("temp"),
}

PDF_SUFFIXES = {".pdf"}
TXT_SUFFIXES = {".txt"}


def _init_dirs():
    ensure_project_dirs()

    for folder in APP_DIRS.values():
        folder.mkdir(parents=True, exist_ok=True)

    if "session_id" not in st.session_state:
        st.session_state.session_id = uuid.uuid4().hex[:12]

    session_dirs = {}
    for name, folder in APP_DIRS.items():
        session_dir = folder / st.session_state.session_id
        session_dir.mkdir(parents=True, exist_ok=True)
        session_dirs[name] = session_dir
    return session_dirs


def _timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _download_file(file_path, label, key, mime):
    file_path = Path(file_path)
    if not file_path.exists():
        return

    st.download_button(
        label,
        data=file_path.read_bytes(),
        file_name=file_path.name,
        mime=mime,
        key=key,
    )


def _collect_folder_files(folder_path, suffixes):
    folder = Path(str(folder_path or "")).expanduser()
    if not folder.exists():
        return [], "Folder tidak ditemukan."
    if not folder.is_dir():
        return [], "Path bukan folder."
    files = sorted(path for path in folder.rglob("*") if path.is_file() and path.suffix.lower() in suffixes)
    if not files:
        return [], "Tidak ada file yang didukung di folder ini."
    return files, None


def _folder_source_ui(key_prefix, suffixes):
    folder_path = st.text_input("Path folder", key=f"{key_prefix}_folder_path")
    scan_clicked = st.button("Scan folder", key=f"{key_prefix}_scan_folder")
    if scan_clicked:
        files, error = _collect_folder_files(folder_path, suffixes)
        st.session_state[f"{key_prefix}_folder_files"] = [str(path) for path in files]
        st.session_state[f"{key_prefix}_folder_error"] = error

    files = [Path(path) for path in st.session_state.get(f"{key_prefix}_folder_files", [])]
    error = st.session_state.get(f"{key_prefix}_folder_error")
    if folder_path and not files and not error:
        files, error = _collect_folder_files(folder_path, suffixes)

    if error:
        st.warning(error)
    elif files:
        st.info(f"{len(files)} file ditemukan.")

    return files


def _show_error_report_download(error_report_path, key):
    if error_report_path:
        _download_file(
            error_report_path,
            "Download Error Report",
            key,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def _write_legacy_error_report(error_rows, output_dir, report_name):
    if not error_rows:
        return None
    output_path = Path(output_dir) / safe_filename(report_name)
    return write_error_report(
        normalize_error_rows(error_rows, stage="extract", error_type="ProcessingError"),
        output_path,
    )


def _show_data_quality(data_quality_df):
    if data_quality_df is not None and not data_quality_df.empty:
        with st.expander("Data Quality", expanded=False):
            st.dataframe(data_quality_df, use_container_width=True)


def _append_data_quality_to_outputs(output_files, data_quality_df, enabled):
    if not enabled or data_quality_df is None or data_quality_df.empty:
        return
    for output_file in output_files or []:
        try:
            append_data_quality_sheet(output_file, data_quality_df)
        except Exception:
            continue


def _render_cleanup_summary(summary):
    if not summary:
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Deleted files", summary.get("deleted_files", 0))
    c2.metric("Deleted folders", summary.get("deleted_dirs", 0))
    c3.metric("Skipped files", summary.get("skipped_files", 0))
    c4.metric("Cleanup errors", len(summary.get("errors", [])))

    targets = summary.get("targets", {})
    if targets:
        st.dataframe(
            pd.DataFrame(
                [
                    {
                        "target_folder": folder,
                        "deleted_files": item.get("deleted_files", 0),
                        "deleted_dirs": item.get("deleted_dirs", 0),
                        "skipped_files": item.get("skipped_files", 0),
                        "error_count": len(item.get("errors", [])),
                    }
                    for folder, item in targets.items()
                ]
            ),
            use_container_width=True,
        )

    if summary.get("errors"):
        st.warning("Sebagian file tidak dibersihkan.")
        st.dataframe(pd.DataFrame(summary["errors"]), use_container_width=True)


def _show_cleanup_result(result):
    if not result:
        return
    with st.expander("Cleanup Summary", expanded=bool(result.get("errors"))):
        _render_cleanup_summary(result)


def _record_cleanup_audit(job_id, cleanup_type, summary):
    if not summary:
        return

    log_cleanup_summary(summary, cleanup_type=cleanup_type, job_id=job_id)
    targets = summary.get("targets") or {"unknown": summary}
    for target_folder, target_summary in targets.items():
        try:
            add_cleanup_log(
                job_id=job_id,
                cleanup_type=cleanup_type,
                target_folder=target_folder,
                deleted_files=target_summary.get("deleted_files", 0),
                deleted_dirs=target_summary.get("deleted_dirs", 0),
                skipped_files=target_summary.get("skipped_files", 0),
                error_count=len(target_summary.get("errors", [])),
                errors_json=json.dumps(target_summary.get("errors", []), ensure_ascii=False),
            )
        except Exception as exc:
            log_cleanup_summary(
                {
                    "deleted_files": 0,
                    "deleted_dirs": 0,
                    "skipped_files": 1,
                    "errors": [{"path": target_folder, "error": f"cleanup audit failed: {exc}"}],
                    "targets": {},
                },
                cleanup_type=f"{cleanup_type}_audit_error",
                job_id=job_id,
            )


def _failed_file_names(*error_sources):
    names = set()
    for error_source in error_sources:
        if error_source is None:
            continue
        if isinstance(error_source, pd.DataFrame):
            records = error_source.to_dict("records")
        else:
            records = error_source

        for record in records or []:
            if not isinstance(record, dict):
                continue
            value = record.get("file") or record.get("source_file") or record.get("source_file_name")
            if not value:
                continue
            names.add(Path(str(value)).name)
            for part in str(value).split(","):
                part = part.strip()
                if part:
                    names.add(Path(part).name)
    return names


def _storage_saved(result):
    if not result:
        return False
    storage = result.get("storage", {})
    if storage.get("errors"):
        return False
    return bool(storage.get("parquet_paths") or storage.get("duckdb_tables"))


def _cleanup_enabled(options):
    return any(options.get(key) for key in ("clean_temp", "clean_uploads", "clean_cache"))


def _auto_cleanup_after_job(
    job_id,
    cleanup_type,
    options,
    upload_paths=None,
    temp_dirs=None,
    failed_file_names=None,
    allow_cleanup=False,
):
    if not _cleanup_enabled(options) or not allow_cleanup:
        return None

    failed_file_names = set(failed_file_names or [])
    older_than_hours = options.get("older_than_hours")
    keep_failed_files = options.get("keep_failed_files", True)

    upload_delete_paths = []
    if options.get("clean_uploads", True):
        for upload_path in upload_paths or []:
            if keep_failed_files and Path(upload_path).name in failed_file_names:
                continue
            upload_delete_paths.append(upload_path)

    temp_delete_dirs = []
    if options.get("clean_temp", True):
        if not (keep_failed_files and failed_file_names):
            temp_delete_dirs = list(temp_dirs or [])

    summaries = []
    if upload_delete_paths or temp_delete_dirs:
        summaries.append(
            cleanup_paths(
                file_paths=upload_delete_paths,
                dir_paths=temp_delete_dirs,
                keep_gitkeep=True,
                older_than_hours=older_than_hours,
            )
        )

    if options.get("clean_cache", False):
        summaries.append(
            cleanup_working_dirs(
                clean_temp=False,
                clean_uploads=False,
                clean_cache=True,
                older_than_hours=older_than_hours,
            )
        )

    summary = combine_cleanup_summaries(*summaries)
    _record_cleanup_audit(job_id, cleanup_type, summary)
    return summary


def _etl_options_ui(key_prefix):
    with st.expander("Advanced ETL Settings"):
        use_cache = st.checkbox("Gunakan cache file hash", value=True, key=f"{key_prefix}_use_cache")
        incremental_processing = st.checkbox(
            "Proses hanya file baru/berubah",
            value=True,
            key=f"{key_prefix}_incremental_processing",
        )
        force_reprocess = st.checkbox(
            "Paksa proses ulang semua file",
            value=False,
            key=f"{key_prefix}_force_reprocess",
        )
        use_multiprocessing = st.checkbox(
            "Gunakan multiprocessing",
            value=False,
            key=f"{key_prefix}_use_multiprocessing",
        )
        max_workers = st.number_input(
            "Jumlah worker",
            min_value=1,
            max_value=16,
            value=4,
            step=1,
            key=f"{key_prefix}_max_workers",
        )
        add_data_quality_sheet = st.checkbox(
            "Tambahkan sheet Data Quality",
            value=True,
            key=f"{key_prefix}_add_data_quality_sheet",
        )
        save_staging = st.checkbox("Simpan staging Parquet", value=True, key=f"{key_prefix}_save_staging")
        save_dwh = st.checkbox("Simpan DWH Parquet", value=True, key=f"{key_prefix}_save_dwh")
        save_mart = st.checkbox("Simpan mart Parquet", value=True, key=f"{key_prefix}_save_mart")
        load_duckdb = st.checkbox("Load ke DuckDB", value=True, key=f"{key_prefix}_load_duckdb")
        st.divider()
        clean_temp = st.checkbox(
            "Hapus file temporary setelah proses berhasil",
            value=True,
            key=f"{key_prefix}_clean_temp",
        )
        clean_uploads = st.checkbox(
            "Hapus uploaded source files setelah tersimpan ke staging/DWH",
            value=True,
            key=f"{key_prefix}_clean_uploads",
        )
        clean_cache = st.checkbox(
            "Hapus cache setelah proses berhasil",
            value=False,
            key=f"{key_prefix}_clean_cache",
        )
        keep_failed_files = st.checkbox(
            "Tetap simpan file yang gagal diproses",
            value=True,
            key=f"{key_prefix}_keep_failed_files",
        )
        older_than_hours_value = st.number_input(
            "Hapus hanya file lebih lama dari X jam",
            min_value=0,
            value=0,
            step=1,
            key=f"{key_prefix}_cleanup_older_than_hours",
            help="Isi 0 untuk membersihkan tanpa batas umur file.",
        )
        older_than_hours = int(older_than_hours_value) if older_than_hours_value else None

        if st.button("Bersihkan Temp/Uploads/Cache Sekarang", key=f"{key_prefix}_manual_cleanup"):
            manual_summary = cleanup_working_dirs(
                clean_temp=clean_temp,
                clean_uploads=clean_uploads,
                clean_cache=clean_cache,
                older_than_hours=older_than_hours,
            )
            _record_cleanup_audit(None, f"{key_prefix}_manual", manual_summary)
            st.session_state[f"{key_prefix}_manual_cleanup_result"] = manual_summary

        manual_result = st.session_state.get(f"{key_prefix}_manual_cleanup_result")
        if manual_result:
            _render_cleanup_summary(manual_result)

    options = normalize_etl_options(
        {
            "save_staging": save_staging,
            "save_dwh": save_dwh,
            "save_mart": save_mart,
            "load_duckdb": load_duckdb,
        }
    )
    options.update(
        {
            "clean_temp": clean_temp,
            "clean_uploads": clean_uploads,
            "clean_cache": clean_cache,
            "keep_failed_files": keep_failed_files,
            "older_than_hours": older_than_hours,
            "use_cache": use_cache,
            "incremental_processing": incremental_processing and not force_reprocess,
            "force_reprocess": force_reprocess,
            "use_multiprocessing": use_multiprocessing,
            "max_workers": int(max_workers),
            "add_data_quality_sheet": add_data_quality_sheet,
        }
    )
    return options


def _etl_enabled(options):
    return any(bool((options or {}).get(key)) for key in ("save_staging", "save_dwh", "save_mart", "load_duckdb"))


def _show_etl_result(result):
    if not result:
        return

    storage = result.get("storage", {})
    errors_df = result.get("errors_df")
    has_storage_errors = bool(storage.get("errors"))
    has_parse_errors = errors_df is not None and not errors_df.empty

    with st.expander("ETL Storage", expanded=has_storage_errors or has_parse_errors):
        job_id = result.get("job_id")
        if job_id:
            st.caption(f"Job ID: {job_id}")

        parquet_paths = storage.get("parquet_paths") or []
        duckdb_tables = storage.get("duckdb_tables") or []
        if parquet_paths:
            st.dataframe(pd.DataFrame({"parquet_path": parquet_paths}), use_container_width=True)
        if duckdb_tables:
            st.dataframe(pd.DataFrame({"duckdb_table": duckdb_tables}), use_container_width=True)
        if has_parse_errors:
            st.warning("Sebagian ekstraksi ETL gagal. Excel utama tetap mengikuti proses converter.")
            st.dataframe(errors_df, use_container_width=True)
        if has_storage_errors:
            st.warning("Sebagian penyimpanan ETL gagal.")
            st.dataframe(pd.DataFrame(storage["errors"]), use_container_width=True)
        if not parquet_paths and not duckdb_tables and not has_storage_errors and not has_parse_errors:
            st.info("Tidak ada penyimpanan ETL yang aktif untuk proses ini.")


def _get_pdf_converter(bank):
    if bank == "BCA":
        from converters.bca import convert_pdf
    elif bank == "BNI":
        from converters.bni import convert_pdf
    elif bank == "DKI":
        from converters.dki import convert_pdf
    elif bank == "Mandiri":
        from converters.mandiri import convert_pdf
    elif bank == "BRI":
        from converters.bri import convert_pdf
    else:
        raise ValueError(f"Bank belum didukung: {bank}")

    return convert_pdf


def _get_pdf_batch_converter(bank):
    if bank == "BCA":
        from convert_mutasi_bca import run_folder
    elif bank == "BNI":
        from convert_mutasi_bni import run_folder
    elif bank == "DKI":
        from convert_mutasi_dki import run_folder
    elif bank == "Mandiri":
        from convert_mutasi_mandiri import run_folder
    elif bank == "BRI":
        from convert_mutasi_BRI import run_folder
    else:
        raise ValueError(f"Bank belum didukung: {bank}")

    return run_folder


def _looks_like_java_tabula_issue(message):
    lowered = str(message).lower()
    return any(token in lowered for token in ("java", "jvm", "tabula", "jpype"))


def _prepare_bank_batch_folder(pdf_paths, output_parent, bank):
    batch_dir = output_parent / safe_filename(bank)
    batch_dir.mkdir(parents=True, exist_ok=True)

    for pdf_path in pdf_paths:
        target = batch_dir / Path(pdf_path).name
        target.write_bytes(Path(pdf_path).read_bytes())

    return batch_dir


def _show_java_warning_if_needed(exc):
    message = str(exc)
    if _looks_like_java_tabula_issue(message):
        st.warning(
            "Converter PDF rekening koran menggunakan tabula-py. "
            "Pastikan Java sudah terpasang dan bisa diakses dari terminal."
        )
    st.error(message)


def page_slik(session_dirs):
    st.header("SLIK / iDeb Converter")

    source_mode = st.radio("Sumber file SLIK", ["Upload file", "Input folder lokal"], horizontal=True, key="slik_source_mode")
    uploaded_files = []
    folder_pdf_paths = []
    if source_mode == "Upload file":
        uploaded_files = st.file_uploader(
            "Upload PDF SLIK",
            type=["pdf"],
            accept_multiple_files=True,
            key="slik_upload",
        )
    else:
        folder_pdf_paths = _folder_source_ui("slik", PDF_SUFFIXES)

    etl_options = _etl_options_ui("slik")
    can_process = bool(uploaded_files) if source_mode == "Upload file" else bool(folder_pdf_paths)

    if st.button("Proses SLIK", type="primary", disabled=not can_process):
        if source_mode == "Upload file":
            upload_dir = create_session_dir("slik_upload", session_dirs["uploads"])
            pdf_paths = save_uploaded_files(uploaded_files, upload_dir)
            cleanup_upload_paths = pdf_paths
        else:
            pdf_paths = folder_pdf_paths
            cleanup_upload_paths = []

        progress_bar = st.progress(0)
        status = st.empty()

        def progress_callback(processed, total, file, rows_found=0, error=None):
            progress_bar.progress(processed / total if total else 0)
            file_name = Path(file).name
            if error:
                status.warning(f"{processed}/{total} gagal: {file_name}")
            else:
                status.write(f"{processed}/{total} diproses: {file_name} ({rows_found} fasilitas)")

        with st.spinner("Memproses PDF SLIK..."):
            result = build_slik_data(
                pdf_paths,
                progress_callback=progress_callback,
                etl_options=etl_options,
            )

        st.session_state.slik_pdf_files = [path.name for path in pdf_paths]
        st.session_state.slik_facilities_df = result["facilities_df"]
        st.session_state.slik_clean_df = result["clean_df"]
        st.session_state.slik_duplicate_count = result["duplicate_count"]
        st.session_state.slik_errors_df = result["errors_df"]
        st.session_state.slik_data_quality_df = result["data_quality_df"]
        st.session_state.slik_error_report_path = result.get("error_report_path")
        st.session_state.slik_etl_result = {
            "job_id": result["job_id"],
            "storage": result["storage"],
            "errors_df": result["errors_df"],
        }
        slik_failed_names = _failed_file_names(result["errors_df"])
        slik_success_count = max(len(pdf_paths) - len(slik_failed_names), 0)
        st.session_state.slik_cleanup_result = _auto_cleanup_after_job(
            job_id=result["job_id"],
            cleanup_type="slik_auto",
            options=etl_options,
            upload_paths=cleanup_upload_paths,
            temp_dirs=[],
            failed_file_names=slik_failed_names,
            allow_cleanup=slik_success_count > 0 and _storage_saved(result),
        )
        st.session_state.slik_output_file = None
        status.success("Proses SLIK selesai.")

    facilities_df = st.session_state.get("slik_facilities_df")
    clean_df = st.session_state.get("slik_clean_df")
    errors_df = st.session_state.get("slik_errors_df", pd.DataFrame(columns=["file", "error"]))

    if facilities_df is None:
        st.info("Upload satu atau beberapa PDF SLIK, lalu klik Proses SLIK.")
        return

    file_count = len(st.session_state.get("slik_pdf_files", []))
    duplicate_count = st.session_state.get("slik_duplicate_count", 0)
    error_count = len(errors_df) if errors_df is not None else 0
    facility_count = len(clean_df) if clean_df is not None else len(facilities_df)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Jumlah file diproses", file_count)
    c2.metric("Jumlah fasilitas ditemukan", facility_count)
    c3.metric("Jumlah duplikat dihapus", duplicate_count)
    c4.metric("Jumlah file error", error_count)
    _show_etl_result(st.session_state.get("slik_etl_result"))
    _show_cleanup_result(st.session_state.get("slik_cleanup_result"))
    _show_error_report_download(st.session_state.get("slik_error_report_path"), "download_slik_error_report")
    _show_data_quality(st.session_state.get("slik_data_quality_df"))

    if errors_df is not None and not errors_df.empty:
        st.warning("Sebagian file gagal diproses. File lain tetap diproses.")
        st.dataframe(errors_df, use_container_width=True)

    if facilities_df.empty:
        if error_count and error_count == file_count:
            st.error("Semua file gagal diproses.")
        else:
            st.info("Tidak ada fasilitas aktif/dihapusbukukan/lunas bermasalah yang ditemukan.")
        return

    st.subheader("Preview Data Fasilitas")
    st.dataframe(without_lineage_columns(facilities_df), use_container_width=True, height=360)

    render_slik_dashboard(clean_df, errors_df)

    st.subheader("Excel Dashboard")
    if st.button("Generate Excel Dashboard", type="primary"):
        output_file = session_dirs["outputs"] / safe_filename(f"slik_dashboard_{_timestamp()}.xlsx")
        with st.spinner("Membuat workbook Excel..."):
            export_dashboard_workbook(
                facilities_df,
                output_file,
                data_quality_df=st.session_state.get("slik_data_quality_df"),
                include_data_quality=bool(etl_options.get("add_data_quality_sheet", True)),
            )
        st.session_state.slik_output_file = str(output_file)
        st.success(f"Workbook selesai dibuat: {output_file.name}")

    output_file = st.session_state.get("slik_output_file")
    if output_file:
        _download_file(
            output_file,
            "Download Excel",
            "download_slik_excel",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def page_rekening_koran(session_dirs):
    st.header("Rekening Koran Converter")
    st.warning("Converter PDF rekening koran menggunakan tabula-py. Java wajib terpasang.")

    jenis_file = st.radio("Jenis File Rekening Koran", ["PDF", "TXT"], horizontal=True)
    etl_options = _etl_options_ui(f"rekening_{jenis_file.lower()}")

    if jenis_file == "PDF":
        source_mode = st.radio("Sumber file PDF", ["Upload file", "Input folder lokal"], horizontal=True, key="rekening_pdf_source_mode")
        uploaded_pdf = []
        folder_pdf_paths = []
        if source_mode == "Upload file":
            uploaded_pdf = st.file_uploader(
                "Upload PDF Rekening Koran",
                type=["pdf"],
                accept_multiple_files=True,
                key="rekening_pdf",
            )
        else:
            folder_pdf_paths = _folder_source_ui("rekening_pdf", PDF_SUFFIXES)
        selected_bank = st.selectbox("Pilih Bank", ["Auto Detect", "BCA", "BNI", "DKI", "Mandiri", "BRI"])
        can_process_pdf = bool(uploaded_pdf) if source_mode == "Upload file" else bool(folder_pdf_paths)

        if st.button("Convert PDF", type="primary", disabled=not can_process_pdf):
            output_dir = create_session_dir("rekening_pdf_output", session_dirs["outputs"])
            if source_mode == "Upload file":
                upload_dir = create_session_dir("rekening_pdf", session_dirs["uploads"])
                pdf_paths = save_uploaded_files(uploaded_pdf, upload_dir)
                cleanup_upload_paths = pdf_paths
            else:
                pdf_paths = folder_pdf_paths
                cleanup_upload_paths = []

            progress_bar = st.progress(0)
            status = st.empty()
            output_files = []
            result_rows = []
            error_rows = []
            etl_file_items = []
            etl_result = None

            bank_groups = {}
            detection_rows = []
            for index, pdf_path in enumerate(pdf_paths, start=1):
                bank = selected_bank
                try:
                    if bank == "Auto Detect":
                        bank = detect_bank_from_pdf(pdf_path)
                        if not bank:
                            raise ValueError("Auto detect gagal. Silakan pilih bank secara manual.")
                    bank_groups.setdefault(bank, []).append(pdf_path)
                    etl_file_items.append({"path": pdf_path, "bank": bank})
                    detection_rows.append({"file": pdf_path.name, "bank": bank})
                except Exception as exc:
                    error_rows.append(
                        {
                            "file": pdf_path.name,
                            "bank": bank or selected_bank,
                            "error": str(exc),
                        }
                    )
                    status.warning(f"{index}/{len(pdf_paths)} gagal: {pdf_path.name}")
                finally:
                    progress_bar.progress(index / (len(pdf_paths) * 2) if pdf_paths else 0)

            total_batches = len(bank_groups)
            completed_batches = 0
            batch_parent = create_session_dir("rekening_pdf_batch", session_dirs["temp"])
            for bank, bank_pdf_paths in sorted(bank_groups.items()):
                try:
                    batch_input_dir = _prepare_bank_batch_folder(bank_pdf_paths, batch_parent, bank)
                    batch_output_dir = output_dir / safe_filename(bank)
                    batch_converter = _get_pdf_batch_converter(bank)

                    status.write(
                        f"Membuat workbook gabungan {bank}: "
                        f"{len(bank_pdf_paths)} file PDF"
                    )
                    written_files = batch_converter(
                        batch_input_dir,
                        batch_output_dir,
                        preserve_relative_folders=True,
                    )
                    for written_file in written_files:
                        written_file = Path(written_file)
                        output_files.append(written_file)
                        result_rows.append(
                            {
                                "bank": bank,
                                "workbook": written_file.name,
                                "path": str(written_file),
                                "status": "Berhasil",
                            }
                        )
                except Exception as exc:
                    error_rows.append(
                        {
                            "file": ", ".join(path.name for path in bank_pdf_paths),
                            "bank": bank,
                            "error": str(exc),
                        }
                    )
                    status.warning(f"Batch {bank} gagal.")
                finally:
                    completed_batches += 1
                    if total_batches:
                        progress_bar.progress(0.5 + (completed_batches / total_batches) * 0.5)

            st.session_state.rekening_pdf_outputs = [str(path) for path in output_files]
            st.session_state.rekening_pdf_results_df = pd.DataFrame(
                result_rows,
                columns=["bank", "workbook", "path", "status"],
            )
            st.session_state.rekening_pdf_errors_df = pd.DataFrame(
                error_rows,
                columns=["file", "bank", "error"],
            )
            st.session_state.rekening_pdf_detection_df = pd.DataFrame(
                detection_rows,
                columns=["file", "bank"],
            )
            st.session_state.rekening_pdf_zip = None
            st.session_state.rekening_pdf_etl_result = None
            st.session_state.rekening_pdf_cleanup_result = None
            st.session_state.rekening_pdf_data_quality_df = None
            st.session_state.rekening_pdf_error_report_path = None

            if etl_file_items and _etl_enabled(etl_options):
                try:
                    with st.spinner("Menyimpan data ETL rekening koran PDF..."):
                        etl_result = build_bank_pdf_etl(
                            etl_file_items,
                            etl_options=etl_options,
                        )
                    st.session_state.rekening_pdf_etl_result = {
                        "job_id": etl_result["job_id"],
                        "storage": etl_result["storage"],
                        "errors_df": etl_result["errors_df"],
                    }
                    st.session_state.rekening_pdf_data_quality_df = etl_result.get("data_quality_df")
                    st.session_state.rekening_pdf_error_report_path = etl_result.get("error_report_path")
                except Exception as exc:
                    st.session_state.rekening_pdf_etl_result = {
                        "storage": {"errors": [{"target": "bank_statement_pdf_etl", "error": str(exc)}]},
                    }

            if output_files and etl_result:
                _append_data_quality_to_outputs(
                    output_files,
                    etl_result.get("data_quality_df"),
                    bool(etl_options.get("add_data_quality_sheet", True)),
                )

            if error_rows and not st.session_state.get("rekening_pdf_error_report_path"):
                legacy_report = _write_legacy_error_report(
                    error_rows,
                    session_dirs["outputs"],
                    f"error_report_rekening_pdf_{_timestamp()}.xlsx",
                )
                st.session_state.rekening_pdf_error_report_path = str(legacy_report) if legacy_report else None

            pdf_failed_names = _failed_file_names(
                error_rows,
                etl_result["errors_df"] if etl_result else None,
            )
            st.session_state.rekening_pdf_cleanup_result = _auto_cleanup_after_job(
                job_id=etl_result["job_id"] if etl_result else None,
                cleanup_type="bank_statement_pdf_auto",
                options=etl_options,
                upload_paths=cleanup_upload_paths,
                temp_dirs=[batch_parent],
                failed_file_names=pdf_failed_names,
                allow_cleanup=bool(output_files) or _storage_saved(etl_result),
            )

            if output_files:
                status.success(f"{len(output_files)} workbook selesai dibuat.")
            else:
                status.error("Tidak ada workbook dibuat.")

            if any(_looks_like_java_tabula_issue(row["error"]) for row in error_rows):
                st.warning(
                    "Sebagian error tampaknya terkait Java/tabula-py. "
                    "Pastikan Java sudah terpasang dan bisa diakses dari terminal."
                )

        results_df = st.session_state.get(
            "rekening_pdf_results_df",
            pd.DataFrame(columns=["bank", "workbook", "path", "status"]),
        )
        errors_df = st.session_state.get(
            "rekening_pdf_errors_df",
            pd.DataFrame(columns=["file", "bank", "error"]),
        )
        detection_df = st.session_state.get(
            "rekening_pdf_detection_df",
            pd.DataFrame(columns=["file", "bank"]),
        )
        output_files = [Path(path) for path in st.session_state.get("rekening_pdf_outputs", [])]

        if detection_df is not None and not detection_df.empty and selected_bank == "Auto Detect":
            st.subheader("Hasil Deteksi Bank")
            st.dataframe(detection_df, use_container_width=True)

        if not results_df.empty:
            st.subheader("Workbook Dibuat")
            st.dataframe(results_df, use_container_width=True)

        if errors_df is not None and not errors_df.empty:
            st.warning("Sebagian file gagal diproses. File lain tetap diproses.")
            st.dataframe(errors_df, use_container_width=True)

        _show_etl_result(st.session_state.get("rekening_pdf_etl_result"))
        _show_cleanup_result(st.session_state.get("rekening_pdf_cleanup_result"))
        _show_error_report_download(st.session_state.get("rekening_pdf_error_report_path"), "download_rekening_pdf_error_report")
        _show_data_quality(st.session_state.get("rekening_pdf_data_quality_df"))

        if output_files:
            if len(output_files) == 1:
                _download_file(
                    output_files[0],
                    "Download Workbook",
                    "download_rekening_pdf_single",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                zip_path = st.session_state.get("rekening_pdf_zip")
                if not zip_path:
                    zip_path = session_dirs["outputs"] / safe_filename(f"rekening_pdf_workbooks_{_timestamp()}.zip")
                    zip_files(output_files, zip_path)
                    st.session_state.rekening_pdf_zip = str(zip_path)
                _download_file(
                    zip_path,
                    "Download Semua sebagai ZIP",
                    "download_rekening_pdf_zip",
                    "application/zip",
                )

    else:
        txt_converter = st.selectbox("Pilih TXT Converter", ["DKI / Bank Jakarta TXT"])
        source_mode = st.radio("Sumber file TXT", ["Upload file", "Input folder lokal"], horizontal=True, key="rekening_txt_source_mode")
        uploaded_txt = []
        folder_txt_paths = []
        if source_mode == "Upload file":
            uploaded_txt = st.file_uploader(
                "Upload TXT Bank DKI",
                type=["txt"],
                accept_multiple_files=True,
                key="rekening_txt",
            )
        else:
            folder_txt_paths = _folder_source_ui("rekening_txt", TXT_SUFFIXES)
        can_process_txt = bool(uploaded_txt) if source_mode == "Upload file" else bool(folder_txt_paths)

        if st.button("Convert TXT DKI", type="primary", disabled=not can_process_txt):
            from converters.dki_txt import convert_dki_txt

            output_dir = create_session_dir("dki_txt_output", session_dirs["outputs"])
            if source_mode == "Upload file":
                upload_dir = create_session_dir("dki_txt", session_dirs["uploads"])
                txt_paths = save_uploaded_files(uploaded_txt, upload_dir)
                input_path = upload_dir
                cleanup_upload_paths = txt_paths
            else:
                txt_paths = folder_txt_paths
                input_path = Path(st.session_state.get("rekening_txt_folder_path", "")).expanduser()
                cleanup_upload_paths = []

            try:
                etl_result = None
                with st.spinner(f"Mengonversi {txt_converter}..."):
                    output_files = convert_dki_txt(input_path, output_dir)
                st.session_state.dki_txt_outputs = [str(path) for path in output_files]
                st.session_state.dki_txt_zip = None
                st.session_state.dki_txt_cleanup_result = None
                st.session_state.dki_txt_data_quality_df = None
                st.session_state.dki_txt_error_report_path = None
                if output_files:
                    st.success(f"{len(output_files)} workbook selesai dibuat.")
                else:
                    st.warning("Tidak ada workbook dibuat.")

                st.session_state.dki_txt_etl_result = None
                if txt_paths and _etl_enabled(etl_options):
                    try:
                        with st.spinner("Menyimpan data ETL rekening koran TXT..."):
                            etl_result = build_dki_txt_etl(
                                txt_paths,
                                etl_options=etl_options,
                            )
                        st.session_state.dki_txt_etl_result = {
                            "job_id": etl_result["job_id"],
                            "storage": etl_result["storage"],
                            "errors_df": etl_result["errors_df"],
                        }
                        st.session_state.dki_txt_data_quality_df = etl_result.get("data_quality_df")
                        st.session_state.dki_txt_error_report_path = etl_result.get("error_report_path")
                    except Exception as exc:
                        st.session_state.dki_txt_etl_result = {
                            "storage": {"errors": [{"target": "bank_statement_txt_etl", "error": str(exc)}]},
                        }

                if output_files and etl_result:
                    _append_data_quality_to_outputs(
                        output_files,
                        etl_result.get("data_quality_df"),
                        bool(etl_options.get("add_data_quality_sheet", True)),
                    )

                txt_failed_names = _failed_file_names(etl_result["errors_df"] if etl_result else None)
                st.session_state.dki_txt_cleanup_result = _auto_cleanup_after_job(
                    job_id=etl_result["job_id"] if etl_result else None,
                    cleanup_type="bank_statement_txt_auto",
                    options=etl_options,
                    upload_paths=cleanup_upload_paths,
                    temp_dirs=[],
                    failed_file_names=txt_failed_names,
                    allow_cleanup=bool(output_files) or _storage_saved(etl_result),
                )
            except Exception as exc:
                st.session_state.dki_txt_outputs = []
                st.session_state.dki_txt_zip = None
                st.session_state.dki_txt_etl_result = None
                st.session_state.dki_txt_cleanup_result = None
                st.session_state.dki_txt_data_quality_df = None
                st.session_state.dki_txt_error_report_path = None
                st.error(str(exc))

        output_files = [Path(path) for path in st.session_state.get("dki_txt_outputs", [])]
        if output_files:
            st.subheader("Workbook Dibuat")
            st.dataframe(
                pd.DataFrame({"file": [path.name for path in output_files], "path": [str(path) for path in output_files]}),
                use_container_width=True,
            )

            if len(output_files) == 1:
                _download_file(
                    output_files[0],
                    "Download Workbook",
                    "download_dki_txt_single",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                zip_path = st.session_state.get("dki_txt_zip")
                if not zip_path:
                    zip_path = session_dirs["outputs"] / safe_filename(f"dki_txt_workbooks_{_timestamp()}.zip")
                    zip_files(output_files, zip_path)
                    st.session_state.dki_txt_zip = str(zip_path)
                _download_file(
                    zip_path,
                    "Download Semua sebagai ZIP",
                    "download_dki_txt_zip",
                    "application/zip",
                )

        _show_etl_result(st.session_state.get("dki_txt_etl_result"))
        _show_cleanup_result(st.session_state.get("dki_txt_cleanup_result"))
        _show_error_report_download(st.session_state.get("dki_txt_error_report_path"), "download_dki_txt_error_report")
        _show_data_quality(st.session_state.get("dki_txt_data_quality_df"))


def page_about():
    st.header("Tentang / Instruksi")
    st.write(
        "Platform lokal ini membungkus script converter yang sudah ada. "
        "Parsing tetap memakai fungsi lama, sementara Streamlit hanya menangani upload, progress, preview, dan download."
    )
    st.markdown(
        """
- Jalankan aplikasi dengan `streamlit run app.py`.
- Gunakan menu `SLIK / iDeb Converter` untuk upload banyak PDF SLIK, preview dashboard, lalu generate Excel saat dibutuhkan.
- Gunakan menu `Rekening Koran Converter` untuk PDF BCA/BNI/DKI/Mandiri/BRI atau TXT DKI / Bank Jakarta. Multi-upload PDF akan dibuat sebagai workbook gabungan per rekening dan tahun.
- Untuk converter PDF rekening koran, pastikan Java tersedia karena tabula-py membutuhkannya.
- Folder kerja lokal: `uploads/`, `outputs/`, `cache/`, dan `temp/`.
"""
    )


def page_job_history():
    st.header("Riwayat Proses")
    try:
        history_df = get_job_history(limit=100)
    except Exception as exc:
        st.error(f"Gagal membaca riwayat proses: {exc}")
        return

    if history_df.empty:
        st.info("Belum ada riwayat proses.")
        return

    preferred_columns = [
        "started_at",
        "job_type",
        "status",
        "duration_seconds",
        "input_file_count",
        "success_file_count",
        "error_file_count",
        "parser_version",
        "cache_used",
        "incremental_mode",
        "multiprocessing_used",
        "max_workers",
        "output_path",
        "job_id",
    ]
    columns = [column for column in preferred_columns if column in history_df.columns]
    st.dataframe(history_df[columns], use_container_width=True)


def main():
    st.set_page_config(page_title="PDF to Excel Platform", layout="wide")
    session_dirs = _init_dirs()

    st.sidebar.title("PDF to Excel")
    menu = st.sidebar.radio(
        "Pilih Menu",
        ["SLIK / iDeb Converter", "Rekening Koran Converter", "Riwayat Proses", "Tentang / Instruksi"],
    )

    if menu == "SLIK / iDeb Converter":
        page_slik(session_dirs)
    elif menu == "Rekening Koran Converter":
        page_rekening_koran(session_dirs)
    elif menu == "Riwayat Proses":
        page_job_history()
    else:
        page_about()


if __name__ == "__main__":
    main()
