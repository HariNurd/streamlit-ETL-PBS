from decimal import Decimal

import pandas as pd


ISSUE_COLUMNS = [
    "source_file",
    "table_name",
    "rule_name",
    "severity",
    "column_name",
    "row_number",
    "message",
]


def _text(value):
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _amount(value):
    if value is None or pd.isna(value):
        return None
    if isinstance(value, Decimal):
        return value
    try:
        return Decimal(str(value).replace(",", ""))
    except Exception:
        return None


def _add_issue(issues, row, row_number, rule_name, severity, column_name, message):
    issues.append(
        {
            "source_file": row.get("source_file"),
            "table_name": "stg_bank_transactions",
            "rule_name": rule_name,
            "severity": severity,
            "column_name": column_name,
            "row_number": row_number,
            "message": message,
        }
    )


def validate_bank_transactions(transactions_df, summaries_df=None):
    issues = []
    if transactions_df is None or transactions_df.empty:
        return pd.DataFrame(columns=ISSUE_COLUMNS)

    df = transactions_df.reset_index(drop=True)
    for idx, row in df.iterrows():
        row_number = idx + 2
        if "Tanggal" in row and not _text(row.get("Tanggal")):
            _add_issue(issues, row, row_number, "tanggal_required", "ERROR", "Tanggal", "Tanggal tidak boleh kosong.")
        if "Keterangan" in row and not _text(row.get("Keterangan")):
            _add_issue(issues, row, row_number, "keterangan_required", "ERROR", "Keterangan", "Keterangan tidak boleh kosong.")
        db = _amount(row.get("DB"))
        cr = _amount(row.get("CR"))
        if db is None and cr is None:
            _add_issue(issues, row, row_number, "db_or_cr_required", "ERROR", "DB/CR", "Minimal salah satu dari DB atau CR harus terisi.")
        if "Saldo" in row and _text(row.get("Saldo")) and _amount(row.get("Saldo")) is None:
            _add_issue(issues, row, row_number, "saldo_numeric", "ERROR", "Saldo", "Saldo harus numerik jika terisi.")
        if _text(row.get("ValidationIssue")):
            _add_issue(
                issues,
                row,
                row_number,
                "dki_txt_transaction_validation",
                "WARNING",
                "ValidationIssue",
                _text(row.get("ValidationIssue")),
            )

        joined = " ".join(_text(row.get(col)) for col in ["Tanggal", "Keterangan", "DB", "CR", "Saldo"])
        if "TANGGAL" in joined.upper() and "KETERANGAN" in joined.upper():
            _add_issue(issues, row, row_number, "header_row_removed", "WARNING", None, "Baris header berulang masih ada di transaksi.")

    if summaries_df is not None and not summaries_df.empty and "source_file" in df.columns:
        # Lightweight reconciliation: only run when summary columns are available in a normalized form.
        group_columns = ["source_file"]
        if "Account" in df.columns and "Account" in summaries_df.columns:
            group_columns.append("Account")

        for group_key, group in df.groupby(group_columns, dropna=False):
            if not isinstance(group_key, tuple):
                group_key = (group_key,)
            source_file = group_key[0]
            account = group_key[1] if len(group_key) > 1 else None
            summary = summaries_df[summaries_df.get("source_file", pd.Series(dtype=str)) == source_file]
            if account is not None:
                summary = summary[summary["Account"] == account]
            if summary.empty:
                continue
            saldo_awal_cols = [col for col in summary.columns if str(col).lower() in {"saldo awal", "saldo_awal", "saldo awal (rp)"}]
            saldo_akhir_cols = [col for col in summary.columns if str(col).lower() in {"saldo akhir", "saldo_akhir", "saldo (rp)"}]
            if not saldo_awal_cols or not saldo_akhir_cols:
                continue
            saldo_awal = _amount(summary[saldo_awal_cols[0]].dropna().iloc[0]) if not summary[saldo_awal_cols[0]].dropna().empty else None
            saldo_akhir = _amount(summary[saldo_akhir_cols[0]].dropna().iloc[-1]) if not summary[saldo_akhir_cols[0]].dropna().empty else None
            if saldo_awal is None or saldo_akhir is None:
                continue
            total_db = sum((_amount(value) or Decimal("0")) for value in group.get("DB", pd.Series(dtype=object)))
            total_cr = sum((_amount(value) or Decimal("0")) for value in group.get("CR", pd.Series(dtype=object)))
            expected = saldo_awal + total_cr - total_db
            if abs(expected - saldo_akhir) > Decimal("1"):
                issues.append(
                    {
                        "source_file": source_file,
                        "table_name": "stg_bank_transactions",
                        "rule_name": "reconciliation_balance",
                        "severity": "WARNING",
                        "column_name": "Saldo",
                        "row_number": None,
                        "message": f"Saldo Awal + CR - DB = {expected}, tidak sama dengan Saldo Akhir {saldo_akhir}.",
                    }
                )

        if "Validation Issues" in summaries_df.columns:
            for _, summary_row in summaries_df.iterrows():
                message = _text(summary_row.get("Validation Issues"))
                if not message:
                    continue
                issues.append(
                    {
                        "source_file": summary_row.get("source_file"),
                        "table_name": "stg_bank_statement_summary",
                        "rule_name": "dki_txt_parser_diagnostics",
                        "severity": "WARNING",
                        "column_name": "Validation Issues",
                        "row_number": None,
                        "message": message,
                    }
                )

    return pd.DataFrame(issues, columns=ISSUE_COLUMNS)
