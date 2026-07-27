from loaders.parquet_store import write_layer_dataframe
from storage.duckdb_store import register_or_load_parquet, write_df_to_duckdb


DEFAULT_ETL_OPTIONS = {
    "save_staging": True,
    "save_dwh": True,
    "save_mart": True,
    "load_duckdb": True,
}


def normalize_etl_options(options=None):
    normalized = DEFAULT_ETL_OPTIONS.copy()
    normalized.update(options or {})
    return normalized


def _record_error(errors, target, exc):
    errors.append({"target": target, "error": str(exc)})


def _write_parquet_layer(results, errors, layer, dataset_name, df):
    try:
        path = write_layer_dataframe(df, layer, dataset_name)
        results["parquet_paths"].append(str(path))
        return path
    except Exception as exc:
        _record_error(errors, f"{layer}.{dataset_name}", exc)
        return None


def _write_duckdb_table(results, errors, table_name, df):
    try:
        table = write_df_to_duckdb(df, table_name, mode="replace")
        results["duckdb_tables"].append(table)
    except Exception as exc:
        _record_error(errors, f"duckdb.{table_name}", exc)


def _register_parquet(results, errors, table_name, parquet_path):
    try:
        table = register_or_load_parquet(table_name, parquet_path)
        if table not in results["duckdb_tables"]:
            results["duckdb_tables"].append(table)
    except Exception as exc:
        _record_error(errors, f"duckdb.{table_name}", exc)


def persist_slik_layers(facilities_df, dwh_df, marts=None, options=None):
    options = normalize_etl_options(options)
    results = {"parquet_paths": [], "duckdb_tables": [], "errors": []}
    parquet_paths = {}

    if options["save_staging"]:
        parquet_paths["stg_slik_facilities"] = _write_parquet_layer(
            results,
            results["errors"],
            "staging",
            "stg_slik_facilities",
            facilities_df,
        )

    if options["save_dwh"]:
        parquet_paths["dwh_slik_facilities"] = _write_parquet_layer(
            results,
            results["errors"],
            "dwh",
            "dwh_slik_facilities",
            dwh_df,
        )

    if options["save_mart"]:
        for dataset_name, mart_df in (marts or {}).items():
            parquet_paths[dataset_name] = _write_parquet_layer(
                results,
                results["errors"],
                "mart",
                dataset_name,
                mart_df,
            )

    if options["load_duckdb"]:
        if options["save_staging"] or options["save_dwh"] or options["save_mart"]:
            for table_name, parquet_path in parquet_paths.items():
                if parquet_path is not None:
                    _register_parquet(results, results["errors"], table_name, parquet_path)
        else:
            _write_duckdb_table(results, results["errors"], "stg_slik_facilities", facilities_df)
            _write_duckdb_table(results, results["errors"], "dwh_slik_facilities", dwh_df)
            for table_name, mart_df in (marts or {}).items():
                _write_duckdb_table(results, results["errors"], table_name, mart_df)

    return results


def persist_bank_statement_layers(transactions_df, dwh_df, marts=None, options=None):
    options = normalize_etl_options(options)
    results = {"parquet_paths": [], "duckdb_tables": [], "errors": []}
    parquet_paths = {}

    if options["save_staging"]:
        parquet_paths["stg_bank_transactions"] = _write_parquet_layer(
            results,
            results["errors"],
            "staging",
            "stg_bank_transactions",
            transactions_df,
        )

    if options["save_dwh"]:
        parquet_paths["dwh_bank_transactions"] = _write_parquet_layer(
            results,
            results["errors"],
            "dwh",
            "dwh_bank_transactions",
            dwh_df,
        )

    if options["save_mart"]:
        for dataset_name, mart_df in (marts or {}).items():
            parquet_paths[dataset_name] = _write_parquet_layer(
                results,
                results["errors"],
                "mart",
                dataset_name,
                mart_df,
            )

    if options["load_duckdb"]:
        if options["save_staging"] or options["save_dwh"] or options["save_mart"]:
            for table_name, parquet_path in parquet_paths.items():
                if parquet_path is not None:
                    _register_parquet(results, results["errors"], table_name, parquet_path)
        else:
            _write_duckdb_table(results, results["errors"], "stg_bank_transactions", transactions_df)
            _write_duckdb_table(results, results["errors"], "dwh_bank_transactions", dwh_df)
            for table_name, mart_df in (marts or {}).items():
                _write_duckdb_table(results, results["errors"], table_name, mart_df)

    return results
