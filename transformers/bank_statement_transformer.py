from decimal import Decimal

import pandas as pd


def _to_float(value):
    if value is None or pd.isna(value):
        return 0.0
    if isinstance(value, Decimal):
        return float(value)
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def build_bank_statement_dwh_df(transactions_df):
    if transactions_df is None:
        return pd.DataFrame()

    dwh_df = transactions_df.copy()
    for column in ["DB", "CR", "Saldo", "Amount"]:
        if column in dwh_df.columns:
            dwh_df[f"{column}_numeric"] = dwh_df[column].apply(_to_float)
    return dwh_df


def build_bank_statement_marts(dwh_df):
    if dwh_df is None or dwh_df.empty:
        return {"mart_bank_statement_summary": pd.DataFrame()}

    df = dwh_df.copy()
    for source_column in ["DB", "CR"]:
        numeric_column = f"{source_column}_numeric"
        if numeric_column not in df.columns:
            df[numeric_column] = df[source_column].apply(_to_float) if source_column in df.columns else 0.0

    group_columns = [
        column
        for column in ["job_id", "bank_name", "input_format", "source_file"]
        if column in df.columns
    ]
    if not group_columns:
        return {"mart_bank_statement_summary": pd.DataFrame()}

    summary = (
        df.groupby(group_columns, dropna=False)
        .agg(
            transaction_count=("source_file", "size"),
            total_db=("DB_numeric", "sum"),
            total_cr=("CR_numeric", "sum"),
        )
        .reset_index()
    )
    return {"mart_bank_statement_summary": summary}
