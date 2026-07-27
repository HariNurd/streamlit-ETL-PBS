import re

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

FALLBACK_KETERANGAN = "Tidak ada Fasilitas Aktif"
IDENTITY_COLUMN = "NIK/NPWP"


def _text(value):
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _num(value):
    return pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]


def _is_fallback(row):
    return _text(row.get("Keterangan")).casefold() == FALLBACK_KETERANGAN.casefold()


def _add_issue(issues, row, row_number, rule_name, severity, column_name, message):
    issues.append(
        {
            "source_file": row.get("source_file"),
            "table_name": "stg_slik_facilities",
            "rule_name": rule_name,
            "severity": severity,
            "column_name": column_name,
            "row_number": row_number,
            "message": message,
        }
    )


def validate_slik_facilities(facilities_df):
    issues = []
    if facilities_df is None or facilities_df.empty:
        return pd.DataFrame(columns=ISSUE_COLUMNS)

    for idx, row in facilities_df.reset_index(drop=True).iterrows():
        row_number = idx + 2
        fallback = _is_fallback(row)

        if not fallback and not _text(row.get("Nama")):
            _add_issue(issues, row, row_number, "nama_required", "ERROR", "Nama", "Nama tidak boleh kosong.")

        identity = _text(row.get(IDENTITY_COLUMN)) or _text(row.get("NIK"))
        if identity and not re.fullmatch(r"\d{15,16}", identity):
            _add_issue(
                issues,
                row,
                row_number,
                "identity_format",
                "WARNING",
                IDENTITY_COLUMN,
                "NIK/NPWP harus kosong atau berisi 15-16 digit angka.",
            )

        if not fallback and not _text(row.get("Bank")):
            _add_issue(issues, row, row_number, "bank_required", "ERROR", "Bank", "Bank tidak boleh kosong.")

        kol = _num(row.get("Kol."))
        if not pd.isna(kol) and int(kol) not in {1, 2, 3, 4, 5}:
            _add_issue(issues, row, row_number, "kol_range", "ERROR", "Kol.", "Kol. harus bernilai 1 sampai 5.")

        plafond = _num(row.get("Plafond (Rp)"))
        baki_debet = _num(row.get("Baki Debet (Rp)"))
        if not pd.isna(plafond) and plafond < 0:
            _add_issue(issues, row, row_number, "plafond_non_negative", "ERROR", "Plafond (Rp)", "Plafond tidak boleh negatif.")
        if not pd.isna(baki_debet) and baki_debet < 0:
            _add_issue(issues, row, row_number, "baki_debet_non_negative", "ERROR", "Baki Debet (Rp)", "Baki Debet tidak boleh negatif.")
        if not pd.isna(plafond) and not pd.isna(baki_debet) and plafond >= 0 and baki_debet > plafond:
            _add_issue(issues, row, row_number, "baki_debet_over_plafond", "WARNING", "Baki Debet (Rp)", "Baki Debet melebihi Plafond.")

        rate = _num(row.get("Rate"))
        if not pd.isna(rate):
            rate_pct = rate * 100 if 0 <= rate <= 1 else rate
            if rate_pct < 0 or rate_pct > 100:
                _add_issue(issues, row, row_number, "rate_range", "WARNING", "Rate", "Rate harus di antara 0 dan 100%.")

        jatuh_tempo = _text(row.get("Jatuh Tempo"))
        if jatuh_tempo and pd.isna(pd.to_datetime(jatuh_tempo, errors="coerce", dayfirst=True)):
            _add_issue(issues, row, row_number, "jatuh_tempo_date", "WARNING", "Jatuh Tempo", "Jatuh Tempo tidak terbaca sebagai tanggal.")

        if fallback and (not pd.isna(plafond) and plafond != 0 or not pd.isna(baki_debet) and baki_debet != 0):
            _add_issue(issues, row, row_number, "fallback_exposure", "ERROR", "Keterangan", "Fallback 'Tidak ada fasilitas' tidak boleh punya exposure.")

    return pd.DataFrame(issues, columns=ISSUE_COLUMNS)
