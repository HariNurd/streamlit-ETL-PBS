from pathlib import Path

import pandas as pd


CACHE_ROOT = Path("cache/staging")


def _require_parquet_engine():
    try:
        import pyarrow  # noqa: F401
    except ImportError as exc:
        raise RuntimeError(
            "pyarrow is required for cache Parquet files. Install dependencies from requirements.txt."
        ) from exc


def _safe_parser_version(parser_version):
    return str(parser_version).replace("/", "_").replace("\\", "_").replace(":", "_")


def cache_path(document_type, parser_version, source_file_hash):
    return (
        CACHE_ROOT
        / str(document_type)
        / _safe_parser_version(parser_version)
        / f"{source_file_hash}.parquet"
    )


def load_cached_dataframe(document_type, parser_version, source_file_hash):
    path = cache_path(document_type, parser_version, source_file_hash)
    if not path.exists():
        return None, path
    _require_parquet_engine()
    return pd.read_parquet(path, engine="pyarrow"), path


def save_cached_dataframe(df, document_type, parser_version, source_file_hash):
    _require_parquet_engine()
    path = cache_path(document_type, parser_version, source_file_hash)
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_parquet(path, index=False, engine="pyarrow")
    return path
