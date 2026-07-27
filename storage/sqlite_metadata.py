import sqlite3
import uuid
from datetime import datetime
from pathlib import Path

import pandas as pd


DEFAULT_METADATA_DB = "data/metadata/etl_metadata.sqlite"


JOB_EXTRA_COLUMNS = {
    "parser_version": "TEXT",
    "cache_used": "INTEGER DEFAULT 0",
    "incremental_mode": "INTEGER DEFAULT 0",
    "multiprocessing_used": "INTEGER DEFAULT 0",
    "max_workers": "INTEGER",
    "cleanup_enabled": "INTEGER DEFAULT 0",
}

JOB_FILE_EXTRA_COLUMNS = {
    "loaded_from_cache": "INTEGER DEFAULT 0",
    "cache_path": "TEXT",
}


def _now():
    return datetime.now().isoformat(timespec="seconds")


def _connect(db_path=DEFAULT_METADATA_DB):
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init_metadata_db(db_path=DEFAULT_METADATA_DB):
    with _connect(db_path) as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS jobs (
                job_id TEXT PRIMARY KEY,
                job_type TEXT NOT NULL,
                status TEXT NOT NULL,
                started_at TEXT NOT NULL,
                finished_at TEXT,
                duration_seconds REAL,
                input_file_count INTEGER DEFAULT 0,
                success_file_count INTEGER DEFAULT 0,
                error_file_count INTEGER DEFAULT 0,
                output_path TEXT,
                notes TEXT
            );

            CREATE TABLE IF NOT EXISTS job_files (
                job_file_id TEXT PRIMARY KEY,
                job_id TEXT NOT NULL,
                source_file_name TEXT,
                source_file_path TEXT,
                source_file_hash TEXT,
                document_type TEXT,
                parser_name TEXT,
                parser_version TEXT,
                status TEXT,
                row_count INTEGER,
                error_message TEXT,
                processed_at TEXT,
                FOREIGN KEY(job_id) REFERENCES jobs(job_id)
            );

            CREATE TABLE IF NOT EXISTS error_logs (
                error_id TEXT PRIMARY KEY,
                job_id TEXT,
                source_file_name TEXT,
                stage TEXT,
                error_type TEXT,
                error_message TEXT,
                created_at TEXT,
                FOREIGN KEY(job_id) REFERENCES jobs(job_id)
            );

            CREATE TABLE IF NOT EXISTS data_quality_results (
                dq_id TEXT PRIMARY KEY,
                job_id TEXT,
                source_file_name TEXT,
                table_name TEXT,
                rule_name TEXT,
                severity TEXT,
                column_name TEXT,
                issue_count INTEGER,
                message TEXT,
                created_at TEXT,
                FOREIGN KEY(job_id) REFERENCES jobs(job_id)
            );

            CREATE TABLE IF NOT EXISTS cleanup_logs (
                cleanup_id TEXT PRIMARY KEY,
                job_id TEXT,
                cleanup_type TEXT,
                target_folder TEXT,
                deleted_files INTEGER DEFAULT 0,
                deleted_dirs INTEGER DEFAULT 0,
                skipped_files INTEGER DEFAULT 0,
                error_count INTEGER DEFAULT 0,
                errors_json TEXT,
                created_at TEXT,
                FOREIGN KEY(job_id) REFERENCES jobs(job_id)
            );
            """
        )
        _ensure_columns(conn, "jobs", JOB_EXTRA_COLUMNS)
        _ensure_columns(conn, "job_files", JOB_FILE_EXTRA_COLUMNS)


def _ensure_columns(conn, table_name, columns):
    existing = {
        row["name"]
        for row in conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    }
    for column_name, column_type in columns.items():
        if column_name not in existing:
            conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}")


def create_job(job_type, input_file_count=0, notes=None, db_path=DEFAULT_METADATA_DB):
    init_metadata_db(db_path)
    job_id = uuid.uuid4().hex
    with _connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO jobs (
                job_id, job_type, status, started_at, input_file_count, notes
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (job_id, job_type, "running", _now(), int(input_file_count or 0), notes),
        )
    return job_id


def update_job_status(
    job_id,
    status,
    input_file_count=None,
    success_file_count=None,
    error_file_count=None,
    output_path=None,
    notes=None,
    parser_version=None,
    cache_used=None,
    incremental_mode=None,
    multiprocessing_used=None,
    max_workers=None,
    cleanup_enabled=None,
    db_path=DEFAULT_METADATA_DB,
):
    init_metadata_db(db_path)
    finished_at = _now() if str(status).lower() != "running" else None

    with _connect(db_path) as conn:
        row = conn.execute(
            "SELECT started_at FROM jobs WHERE job_id = ?",
            (job_id,),
        ).fetchone()
        duration_seconds = None
        if row and finished_at:
            started_at = datetime.fromisoformat(row["started_at"])
            duration_seconds = (datetime.fromisoformat(finished_at) - started_at).total_seconds()

        assignments = ["status = ?"]
        values = [status]
        if finished_at:
            assignments.extend(["finished_at = ?", "duration_seconds = ?"])
            values.extend([finished_at, duration_seconds])
        if input_file_count is not None:
            assignments.append("input_file_count = ?")
            values.append(int(input_file_count))
        if success_file_count is not None:
            assignments.append("success_file_count = ?")
            values.append(int(success_file_count))
        if error_file_count is not None:
            assignments.append("error_file_count = ?")
            values.append(int(error_file_count))
        if output_path is not None:
            assignments.append("output_path = ?")
            values.append(str(output_path))
        if notes is not None:
            assignments.append("notes = ?")
            values.append(notes)
        if parser_version is not None:
            assignments.append("parser_version = ?")
            values.append(parser_version)
        if cache_used is not None:
            assignments.append("cache_used = ?")
            values.append(1 if cache_used else 0)
        if incremental_mode is not None:
            assignments.append("incremental_mode = ?")
            values.append(1 if incremental_mode else 0)
        if multiprocessing_used is not None:
            assignments.append("multiprocessing_used = ?")
            values.append(1 if multiprocessing_used else 0)
        if max_workers is not None:
            assignments.append("max_workers = ?")
            values.append(int(max_workers))
        if cleanup_enabled is not None:
            assignments.append("cleanup_enabled = ?")
            values.append(1 if cleanup_enabled else 0)

        values.append(job_id)
        conn.execute(
            f"UPDATE jobs SET {', '.join(assignments)} WHERE job_id = ?",
            values,
        )


def add_job_file_record(
    job_id,
    source_file_name,
    source_file_path,
    source_file_hash,
    document_type,
    parser_name,
    parser_version,
    status,
    row_count=0,
    error_message=None,
    processed_at=None,
    loaded_from_cache=False,
    cache_path=None,
    db_path=DEFAULT_METADATA_DB,
):
    init_metadata_db(db_path)
    with _connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO job_files (
                job_file_id, job_id, source_file_name, source_file_path,
                source_file_hash, document_type, parser_name, parser_version,
                status, row_count, error_message, processed_at,
                loaded_from_cache, cache_path
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                uuid.uuid4().hex,
                job_id,
                source_file_name,
                str(source_file_path) if source_file_path is not None else None,
                source_file_hash,
                document_type,
                parser_name,
                parser_version,
                status,
                int(row_count or 0),
                error_message,
                processed_at or _now(),
                1 if loaded_from_cache else 0,
                str(cache_path) if cache_path is not None else None,
            ),
        )


def add_error_log(
    job_id,
    source_file_name,
    stage,
    error_type,
    error_message,
    db_path=DEFAULT_METADATA_DB,
):
    init_metadata_db(db_path)
    with _connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO error_logs (
                error_id, job_id, source_file_name, stage, error_type,
                error_message, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                uuid.uuid4().hex,
                job_id,
                source_file_name,
                stage,
                error_type,
                str(error_message),
                _now(),
            ),
        )


def add_data_quality_result(
    job_id,
    source_file_name,
    table_name,
    rule_name,
    severity,
    column_name,
    issue_count,
    message,
    db_path=DEFAULT_METADATA_DB,
):
    init_metadata_db(db_path)
    with _connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO data_quality_results (
                dq_id, job_id, source_file_name, table_name, rule_name,
                severity, column_name, issue_count, message, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                uuid.uuid4().hex,
                job_id,
                source_file_name,
                table_name,
                rule_name,
                severity,
                column_name,
                int(issue_count or 0),
                message,
                _now(),
            ),
        )


def add_cleanup_log(
    job_id,
    cleanup_type,
    target_folder,
    deleted_files=0,
    deleted_dirs=0,
    skipped_files=0,
    error_count=0,
    errors_json=None,
    db_path=DEFAULT_METADATA_DB,
):
    init_metadata_db(db_path)
    with _connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO cleanup_logs (
                cleanup_id, job_id, cleanup_type, target_folder,
                deleted_files, deleted_dirs, skipped_files, error_count,
                errors_json, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                uuid.uuid4().hex,
                job_id,
                cleanup_type,
                target_folder,
                int(deleted_files or 0),
                int(deleted_dirs or 0),
                int(skipped_files or 0),
                int(error_count or 0),
                errors_json,
                _now(),
            ),
        )


def find_successful_file_by_hash(
    source_file_hash,
    parser_name,
    parser_version,
    document_type=None,
    db_path=DEFAULT_METADATA_DB,
):
    init_metadata_db(db_path)
    conditions = [
        "source_file_hash = ?",
        "parser_name = ?",
        "parser_version = ?",
        "status = 'success'",
    ]
    values = [source_file_hash, parser_name, parser_version]
    if document_type:
        conditions.append("document_type = ?")
        values.append(document_type)

    with _connect(db_path) as conn:
        row = conn.execute(
            f"""
            SELECT *
            FROM job_files
            WHERE {' AND '.join(conditions)}
            ORDER BY processed_at DESC
            LIMIT 1
            """,
            values,
        ).fetchone()
    return dict(row) if row else None


def get_job_history(limit=50, db_path=DEFAULT_METADATA_DB):
    init_metadata_db(db_path)
    query = """
        SELECT *
        FROM jobs
        ORDER BY started_at DESC
    """
    params = []
    if limit is not None:
        query += " LIMIT ?"
        params.append(int(limit))

    with _connect(db_path) as conn:
        rows = conn.execute(query, params).fetchall()
    return pd.DataFrame([dict(row) for row in rows])
