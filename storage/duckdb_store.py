import re
from pathlib import Path

import pandas as pd


DEFAULT_DUCKDB_PATH = "data/warehouse.duckdb"


def _require_duckdb():
    try:
        import duckdb
    except ImportError as exc:
        raise RuntimeError(
            "duckdb is required for analytical storage. Install dependencies from requirements.txt."
        ) from exc
    return duckdb


def _table_name(table_name):
    cleaned = re.sub(r"[^A-Za-z0-9_]+", "_", str(table_name)).strip("_").lower()
    if not cleaned:
        raise ValueError("Nama table DuckDB kosong.")
    if cleaned[0].isdigit():
        cleaned = f"t_{cleaned}"
    return cleaned


def _quote_identifier(table_name):
    return f'"{_table_name(table_name)}"'


def _parquet_sql_path(parquet_path):
    return Path(parquet_path).as_posix().replace("'", "''")


def get_duckdb_connection(db_path=DEFAULT_DUCKDB_PATH):
    duckdb = _require_duckdb()
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    return duckdb.connect(str(db_path))


def register_or_load_parquet(table_name, parquet_path, db_path=DEFAULT_DUCKDB_PATH):
    conn = get_duckdb_connection(db_path)
    try:
        quoted_table = _quote_identifier(table_name)
        parquet_sql_path = _parquet_sql_path(parquet_path)
        conn.execute(
            f"CREATE OR REPLACE VIEW {quoted_table} AS "
            f"SELECT * FROM read_parquet('{parquet_sql_path}')"
        )
        return _table_name(table_name)
    finally:
        conn.close()


def write_df_to_duckdb(df, table_name, mode="replace", db_path=DEFAULT_DUCKDB_PATH):
    if df is None:
        df = pd.DataFrame()

    mode = str(mode).lower()
    if mode not in {"replace", "append"}:
        raise ValueError("mode harus 'replace' atau 'append'.")

    conn = get_duckdb_connection(db_path)
    try:
        normalized_name = _table_name(table_name)
        quoted_table = _quote_identifier(normalized_name)
        conn.register("_etl_df_to_write", df)

        if mode == "replace":
            conn.execute(f"CREATE OR REPLACE TABLE {quoted_table} AS SELECT * FROM _etl_df_to_write")
        else:
            table_exists = conn.execute(
                """
                SELECT COUNT(*)
                FROM information_schema.tables
                WHERE lower(table_name) = lower(?)
                """,
                [normalized_name],
            ).fetchone()[0]
            if table_exists:
                conn.execute(f"INSERT INTO {quoted_table} SELECT * FROM _etl_df_to_write")
            else:
                conn.execute(f"CREATE TABLE {quoted_table} AS SELECT * FROM _etl_df_to_write")

        return normalized_name
    finally:
        try:
            conn.unregister("_etl_df_to_write")
        except Exception:
            pass
        conn.close()


def query_duckdb(sql, db_path=DEFAULT_DUCKDB_PATH):
    conn = get_duckdb_connection(db_path)
    try:
        return conn.execute(sql).fetchdf()
    finally:
        conn.close()
