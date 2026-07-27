import re
from pathlib import Path


VALID_LAYERS = {"staging", "dwh", "mart"}


def _require_parquet_engine():
    try:
        import pyarrow  # noqa: F401
    except ImportError as exc:
        raise RuntimeError(
            "pyarrow is required to write Parquet files. Install dependencies from requirements.txt."
        ) from exc


def sanitize_dataset_name(dataset_name):
    cleaned = re.sub(r"[^A-Za-z0-9_]+", "_", str(dataset_name)).strip("_").lower()
    return cleaned or "dataset"


def layer_path(layer, dataset_name, base_dir="data"):
    if layer not in VALID_LAYERS:
        raise ValueError(f"Layer tidak didukung: {layer}")
    return Path(base_dir) / layer / f"{sanitize_dataset_name(dataset_name)}.parquet"


def write_df_to_parquet(df, parquet_path, index=False):
    _require_parquet_engine()
    parquet_path = Path(parquet_path)
    parquet_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_parquet(parquet_path, index=index, engine="pyarrow")
    return parquet_path


def write_layer_dataframe(df, layer, dataset_name, base_dir="data", index=False):
    return write_df_to_parquet(df, layer_path(layer, dataset_name, base_dir=base_dir), index=index)
