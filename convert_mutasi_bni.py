import re
import argparse
import sys
from decimal import Decimal, InvalidOperation
from functools import lru_cache
from pathlib import Path

import fitz
import pandas as pd
import tabula
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# =========================
# PDF CONFIG
# =========================
PDF_AREA = [125, 15, 760, 830] # [top, left, bottom, right]
PDF_COLUMNS = [105, 250, 610, 700] 

BNI_DIR = Path("pdf_file") / "BNI Taplus Muda"
EXCEL_DIR = Path("excel_file") / "BNI Taplus Muda"

MONTH_ALIASES = {
    "JAN": ("Jan", 1),
    "JANUARI": ("Jan", 1),
    "FEB": ("Feb", 2),
    "FEBRUARI": ("Feb", 2),
    "MAR": ("Mar", 3),
    "MARET": ("Mar", 3),
    "APR": ("Apr", 4),
    "APRIL": ("Apr", 4),
    "MEI": ("May", 5),
    "MAY": ("May", 5),
    "JUN": ("Jun", 6),
    "JUNI": ("Jun", 6),
    "JUL": ("Jul", 7),
    "JULI": ("Jul", 7),
    "AGU": ("Aug", 8),
    "AGS": ("Aug", 8),
    "AUG": ("Aug", 8),
    "AGUSTUS": ("Aug", 8),
    "SEP": ("Sep", 9),
    "SEPT": ("Sep", 9),
    "SEPTEMBER": ("Sep", 9),
    "OKT": ("Oct", 10),
    "OCT": ("Oct", 10),
    "OKTOBER": ("Oct", 10),
    "NOV": ("Nov", 11),
    "NOVEMBER": ("Nov", 11),
    "DES": ("Dec", 12),
    "DEC": ("Dec", 12),
    "DESEMBER": ("Dec", 12),
}

DATE_MONTHS = {name.upper(): order for name, order in [
    ("Jan", 1), ("Feb", 2), ("Mar", 3), ("Apr", 4), ("May", 5), ("Mei", 5),
    ("Jun", 6), ("Jul", 7), ("Aug", 8), ("Agu", 8), ("Sep", 9), ("Oct", 10),
    ("Okt", 10), ("Nov", 11), ("Dec", 12), ("Des", 12),
]}

TRANSACTION_COLUMNS = ["Tanggal", "Jam", "Keterangan", "Mutasi", "DB", "CR", "Saldo"]
SUMMARY_COLUMNS = ["Keterangan", "Amount"]


# =========================
# HELPER FUNCTIONS
# =========================
def clean_text(value):
    if pd.isna(value):
        return ""

    text = str(value).replace("\r", " ").replace("\n", " ").strip()
    return re.sub(r"\s+", " ", text)


def append_text(base, extra):
    base = clean_text(base)
    extra = clean_text(extra)

    if not extra:
        return base
    if not base:
        return extra

    return f"{base} {extra}"


def parse_number(value):
    """
    Convert money text into Decimal without dropping cents.
    Examples: '+4,000,000.00', '-96,600', '6,119,110'
    """
    text = clean_text(value).upper()

    if text in ("", "NAN"):
        return pd.NA

    text = text.replace(",", "").replace("+", "").strip()

    try:
        return Decimal(text)
    except (InvalidOperation, ValueError):
        return pd.NA


def excel_number(value):
    if pd.isna(value):
        return None
    if isinstance(value, Decimal):
        return float(value)
    return value


def dataframe_for_excel(df):
    result = df.copy()
    for col in result.columns:
        result[col] = result[col].apply(excel_number)
    return result


def extract_amount_text(value):
    text = clean_text(value)
    matches = re.findall(r"[+-]?\d[\d,]*(?:\.\d+)?", text)
    return matches[-1] if matches else ""


def is_date(text):
    """
    Example:
    02 Mar 2026
    """
    text = clean_text(text)
    return bool(re.match(r"^\d{2}\s+[A-Za-z]{3}\s+\d{4}$", text))


def date_to_day_month(text):
    text = clean_text(text)
    match = re.match(r"^(\d{2})\s+([A-Za-z]{3})\s+\d{4}$", text)
    if not match:
        return text

    month = DATE_MONTHS.get(match.group(2).upper())
    if not month:
        return text

    return f"{match.group(1)}/{month:02d}"


def is_time(text):
    """
    Example:
    16:53:00 WIB
    """
    text = clean_text(text)
    return bool(re.match(r"^\d{2}:\d{2}:\d{2}\s+WIB$", text))


def extract_nominal_saldo(text):
    """
    Extract nominal and saldo from text ending with:
    -96,600 3,930,095
    +4,000,000 7,716,795
    """
    text = clean_text(text)

    pattern = r"([+-]?\d[\d,]*(?:\.\d+)?)\s+(\d[\d,]*(?:\.\d+)?)$"
    match = re.search(pattern, text)

    if not match:
        return text, "", ""

    nominal = match.group(1)
    saldo = match.group(2)
    remaining_text = text[:match.start()].strip()

    return remaining_text, nominal, saldo


def split_db_cr(nominal):
    text = clean_text(nominal)
    number = parse_number(nominal)

    if pd.isna(number):
        return pd.Series([pd.NA, pd.NA], index=["DB", "CR"])

    if text.startswith("+"):
        return pd.Series([pd.NA, abs(number)], index=["DB", "CR"])

    if text:
        return pd.Series([abs(number), pd.NA], index=["DB", "CR"])

    return pd.Series([pd.NA, pd.NA], index=["DB", "CR"])


def is_money_line(text, signed=None):
    text = clean_text(text)
    if signed is True:
        return bool(re.fullmatch(r"[+-]\d[\d,]*(?:\.\d{1,2})?", text))
    if signed is False:
        return bool(re.fullmatch(r"\d[\d,]*(?:\.\d{1,2})?", text))
    return bool(re.fullmatch(r"[+-]?\d[\d,]*(?:\.\d{1,2})?", text))


# =========================
# READ PDF
# =========================
@lru_cache(maxsize=None)
def read_pdf_text_for_metadata(pdf_file):
    pdf_file = Path(pdf_file)
    text_parts = []

    try:
        with fitz.open(pdf_file) as doc:
            for page in doc:
                text_parts.append(page.get_text("text"))
    except Exception:
        return ""

    return "\n".join(text_parts)


def read_pdf_text_pages(pdf_file, max_pages=2):
    pdf_file = Path(pdf_file)
    text_parts = []

    try:
        with fitz.open(pdf_file) as doc:
            for page_index in range(min(max_pages, len(doc))):
                text_parts.append(doc[page_index].get_text("text"))
    except Exception:
        return ""

    return "\n".join(text_parts)


def detect_bni_format(pdf_file):
    """
    Return:
    - "old" for the Indonesian Laporan Mutasi Rekening format
    - "new" for the English TRANSACTION INQUIRY format
    - None if unknown
    """
    text = normalize_metadata_text(read_pdf_text_pages(pdf_file, max_pages=2))
    if not text:
        return None

    new_markers = [
        "TRANSACTION INQUIRY",
        "ACCOUNT STATEMENT",
        "ACCOUNT INFORMATION",
        "POST DATE",
        "POSTING DATE",
        "EFFECTIVE DATE",
        "TRANSACTION DESCRIPTION",
        "DB/CR",
        "TOTAL DEBIT",
        "TOTAL CREDIT",
    ]
    old_markers = [
        "LAPORAN MUTASI REKENING",
        "TANGGAL & WAKTU",
        "RINCIAN TRANSAKSI",
        "NOMINAL (IDR)",
        "SALDO (IDR)",
    ]

    new_score = sum(1 for marker in new_markers if marker in text)
    old_score = sum(1 for marker in old_markers if marker in text)

    if new_score >= 3 and new_score > old_score:
        return "new"
    if old_score >= 3 and old_score >= new_score:
        return "old"
    return None


def pdf_text_lines(pdf_file):
    text = read_pdf_text_for_metadata(Path(pdf_file))
    return [clean_text(line) for line in text.splitlines() if clean_text(line)]


def normalize_metadata_text(text):
    return re.sub(r"\s+", " ", clean_text(text)).upper()


def extract_account_from_text(text):
    normalized = normalize_metadata_text(text)
    match = re.search(r"TAPLUS\s+MUDA\s*-\s*(\d{6,})", normalized)
    if match:
        return match.group(1)

    match = re.search(r"\bNO\.?\s*REKENING\s*:?\s*(\d{6,})", normalized)
    if match:
        return match.group(1)

    match = re.search(r"\bACCOUNT\s*:?\s*(\d{6,})\b", normalized)
    if match:
        return match.group(1)

    match = re.search(r"\b(\d{6,})\s*/\s*[A-Z]", normalized)
    if match:
        return match.group(1)

    return None


def extract_period_from_text(text):
    normalized = normalize_metadata_text(text)
    month_pattern = "|".join(sorted(MONTH_ALIASES, key=len, reverse=True))

    match = re.search(
        rf"PERIODE\s*:?\s*\d{{1,2}}\s*-\s*\d{{1,2}}\s+({month_pattern})\s+(20\d{{2}})",
        normalized,
    )
    if not match:
        match = re.search(rf"\b({month_pattern})\s+(20\d{{2}})\b", normalized)

    if not match:
        match = re.search(rf"\b\d{{1,2}}[-/\s]+({month_pattern})[-/\s]+(20\d{{2}})\b", normalized)

    if not match:
        match = re.search(rf"\b\d{{1,2}}[-/\s]+({month_pattern})[-/\s]+(\d{{2}})\b", normalized)

    if match:
        month_key, year = match.groups()
        if len(year) == 2:
            year = f"20{year}"
        month_name, month_order = MONTH_ALIASES[month_key]
        return f"{month_name}-{year[-2:]}", month_order, year

    return None


def extract_period_from_filename(pdf_file):
    stem = pdf_file.stem.upper()
    year_match = re.search(r"(20\d{2})", stem)
    year = year_match.group(1) if year_match else None

    month_label = None
    month_order = 99
    for token in re.split(r"[^A-Z]+", stem):
        if token in MONTH_ALIASES:
            month_name, month_order = MONTH_ALIASES[token]
            month_label = f"{month_name}-{year[-2:]}" if year else month_name
            break

    if year is None and re.fullmatch(r"20\d{2}", pdf_file.parent.name):
        year = pdf_file.parent.name
        if month_label and "-" not in month_label:
            month_label = f"{month_label}-{year[-2:]}"

    return month_label or pdf_file.stem[:20], month_order, year or "Unknown"


def extract_pdf_metadata(pdf_file):
    pdf_file = Path(pdf_file)
    text = read_pdf_text_for_metadata(pdf_file)
    month_label, month_order, year = extract_period_from_text(text) or extract_period_from_filename(pdf_file)

    account = extract_account_from_text(text)
    if account is None:
        match = re.search(r"(?:BNI[_\s-]*)?(\d{6,})", pdf_file.stem, flags=re.IGNORECASE)
        account = match.group(1) if match else pdf_file.stem.split("_")[0]

    return {
        "account": account,
        "month_label": month_label,
        "month_order": month_order,
        "year": str(year),
    }


def build_summary_from_text(pdf_file):
    lines = pdf_text_lines(pdf_file)
    summary = []
    labels = {
        "SALDO AWAL": "Saldo Awal",
        "TOTAL PEMASUKAN": "Total Pemasukan",
        "TOTAL PENGELUARAN": "Total Pengeluaran",
        "SALDO AKHIR": "Saldo Akhir",
    }

    first_table_idx = next((idx for idx, line in enumerate(lines) if line.upper() == "TANGGAL & WAKTU"), len(lines))
    header_lines = lines[:first_table_idx]

    for idx, line in enumerate(header_lines):
        key = line.upper()
        if key not in labels:
            continue

        amount = ""
        for next_line in header_lines[idx + 1:idx + 4]:
            if is_money_line(next_line):
                amount = next_line
                break

        parsed = parse_number(amount)
        summary.append({
            "Keterangan": labels[key],
            "Amount": abs(parsed) if not pd.isna(parsed) else pd.NA,
        })

    df_summary = pd.DataFrame(summary, columns=["Keterangan", "Amount"])
    return df_summary.drop_duplicates(subset=["Keterangan"], keep="first").reset_index(drop=True)


def parse_transactions_from_text(pdf_file):
    lines = pdf_text_lines(pdf_file)
    records = []
    current = None
    in_table = False
    skip_next_balance = False

    def finish_current():
        nonlocal current
        if current and current["Mutasi"] and not pd.isna(parse_number(current["Saldo"])):
            records.append(current)
        current = None

    for line in lines:
        upper = line.upper()

        if upper == "TANGGAL & WAKTU":
            in_table = True
            continue
        if not in_table:
            continue
        if upper == "INFORMASI LAINNYA":
            finish_current()
            break

        if upper in {
            "RINCIAN TRANSAKSI",
            "NOMINAL (IDR)",
            "SALDO (IDR)",
            "LAPORAN MUTASI REKENING",
        }:
            continue
        if re.fullmatch(r"\d+\s+dari\s+\d+", line, flags=re.IGNORECASE):
            continue
        if upper.startswith("PERIODE:") or upper.startswith("PT BANK NEGARA") or "PESERTA PENJAMINAN" in upper:
            continue

        if upper in {"SALDO AWAL", "SALDO AKHIR"}:
            finish_current()
            skip_next_balance = True
            continue
        if skip_next_balance and is_money_line(line, signed=False):
            skip_next_balance = False
            continue
        skip_next_balance = False

        if is_date(line):
            finish_current()
            current = {
                "Tanggal": line,
                "Jam": "",
                "Keterangan": "",
                "Mutasi": "",
                "Saldo": "",
            }
            continue

        if current is None:
            continue

        if is_time(line):
            current["Jam"] = line
        elif (
            (is_money_line(line, signed=True) or (current["Keterangan"] and is_money_line(line, signed=False)))
            and current["Mutasi"] == ""
        ):
            current["Mutasi"] = line
        elif is_money_line(line, signed=False) and current["Mutasi"] and current["Saldo"] == "":
            current["Saldo"] = line
            finish_current()
        else:
            current["Keterangan"] = append_text(current["Keterangan"], line)

    finish_current()
    return pd.DataFrame(records, columns=["Tanggal", "Jam", "Keterangan", "Mutasi", "Saldo"])


def is_new_date(text):
    text = clean_text(text)
    return bool(re.fullmatch(r"\d{2}/\d{2}/\d{4}", text))


def is_new_time(text):
    text = clean_text(text)
    return bool(re.fullmatch(r"\d{2}\.\d{2}\.\d{2}", text))


def new_date_to_day_month(text):
    text = clean_text(text)
    match = re.fullmatch(r"(\d{2})/(\d{2})/\d{4}", text)
    if not match:
        return text
    return f"{match.group(1)}/{match.group(2)}"


def normalize_new_time(text):
    text = clean_text(text)
    return text.replace(".", ":") if is_new_time(text) else text


def empty_transactions_dataframe():
    return pd.DataFrame(columns=TRANSACTION_COLUMNS)


def group_pdf_words_into_lines(words, y_tolerance=3):
    lines = []
    for word in sorted(words, key=lambda item: (item[1], item[0])):
        if not lines or abs(word[1] - lines[-1]["y"]) > y_tolerance:
            lines.append({"y": word[1], "words": [word]})
            continue

        lines[-1]["words"].append(word)
        lines[-1]["y"] = (lines[-1]["y"] + word[1]) / 2

    return [sorted(line["words"], key=lambda item: item[0]) for line in lines]


def words_text(words):
    return clean_text(" ".join(word[4] for word in sorted(words, key=lambda item: item[0])))


def words_in_column(words, min_x, max_x):
    return [word for word in words if min_x <= word[0] < max_x]


def find_line_word(words, min_x, max_x, predicate):
    for word in sorted(words_in_column(words, min_x, max_x), key=lambda item: item[0]):
        text = clean_text(word[4])
        if predicate(text):
            return text
    return ""


def is_new_table_header(words):
    text = words_text(words).upper()
    return (
        "POST DATE" in text
        and "DESCRIPTION" in text
        and "DB/CR" in text
        and "BALANCE" in text
    )


def extract_new_summary_from_pdf(pdf_file):
    pdf_file = Path(pdf_file)
    summary = []
    labels = {
        "BEGINNING BALANCE": "Saldo Awal",
        "TOTAL DEBIT": "Total Pengeluaran",
        "TOTAL CREDIT": "Total Pemasukan",
    }

    try:
        with fitz.open(pdf_file) as doc:
            if len(doc) == 0:
                return pd.DataFrame(columns=SUMMARY_COLUMNS)
            lines = group_pdf_words_into_lines(doc[0].get_text("words"))
    except Exception:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    for words in lines:
        text = words_text(words).upper()
        label = next((value for marker, value in labels.items() if marker in text), None)
        if not label:
            continue

        amount_text = find_line_word(
            words,
            210,
            340,
            lambda value: is_money_line(value, signed=None),
        )
        amount = parse_number(amount_text)
        summary.append({
            "Keterangan": label,
            "Amount": abs(amount) if not pd.isna(amount) else pd.NA,
        })

    return pd.DataFrame(summary, columns=SUMMARY_COLUMNS).drop_duplicates(
        subset=["Keterangan"],
        keep="first",
    ).reset_index(drop=True)


def add_or_update_summary_value(df_summary, label, amount):
    df_summary = df_summary.copy()
    amount = abs(amount) if not pd.isna(amount) else pd.NA

    if df_summary.empty:
        return pd.DataFrame([{"Keterangan": label, "Amount": amount}], columns=SUMMARY_COLUMNS)

    mask = df_summary["Keterangan"].str.upper().eq(label.upper())
    if mask.any():
        df_summary.loc[mask, "Amount"] = amount
    else:
        df_summary = pd.concat(
            [df_summary, pd.DataFrame([{"Keterangan": label, "Amount": amount}])],
            ignore_index=True,
        )

    return df_summary[SUMMARY_COLUMNS]


def is_account_statement_format(pdf_file):
    text = normalize_metadata_text(read_pdf_text_pages(pdf_file, max_pages=2))
    return (
        "ACCOUNT STATEMENT" in text
        and "POSTING DATE" in text
        and "EFFECTIVE DATE" in text
        and "TRANSACTION DESCRIPTION" in text
    )


def is_account_statement_datetime(text):
    return bool(re.fullmatch(r"\d{2}/\d{2}/\d{4}\s+\d{2}\.\d{2}\.\d{2}", clean_text(text)))


def is_account_statement_money(text):
    text = clean_text(text)
    if "," not in text and "." not in text:
        return False
    return is_money_line(text, signed=None)


def split_account_statement_datetime(text):
    text = clean_text(text)
    match = re.fullmatch(r"(\d{2})/(\d{2})/\d{4}\s+(\d{2})\.(\d{2})\.(\d{2})", text)
    if not match:
        return text, ""
    return f"{match.group(1)}/{match.group(2)}", f"{match.group(3)}:{match.group(4)}:{match.group(5)}"


def is_account_statement_header_line(line):
    upper = clean_text(line).upper()
    return upper in {
        "ACCOUNT STATEMENT",
        "ACCOUNT INFORMATION",
        "POSTING DATE",
        "EFFECTIVE DATE",
        "TRANSACTION DESCRIPTION",
        "AMOUNT",
        "BALANCE",
        "ACCOUNT NO.",
        "ACCOUNT TYPE",
        "CURRENT",
        "PERIOD",
        "PAGE",
        "DB/CR",
        "BRANCH",
        "JOURNAL",
        ":",
        "-",
    }


def extract_account_statement_summary_from_text(pdf_file):
    lines = pdf_text_lines(pdf_file)
    summary_values = {}

    for idx, line in enumerate(lines):
        upper = line.upper()
        if "LEDGER BALANCE" in upper:
            for next_line in lines[idx + 1:idx + 5]:
                if is_account_statement_money(next_line):
                    summary_values["Saldo Awal"] = abs(parse_number(next_line))
                    break

    ending_idx = next(
        (idx for idx, line in enumerate(lines) if line.upper().startswith("ENDING BALANCE")),
        None,
    )
    if ending_idx is not None:
        tail = lines[ending_idx:ending_idx + 12]
        money_values = [
            parse_number(line)
            for line in tail
            if is_account_statement_money(line) and not pd.isna(parse_number(line))
        ]
        if len(money_values) >= 3:
            summary_values["Total Pengeluaran"] = abs(money_values[0])
            summary_values["Total Pemasukan"] = abs(money_values[1])
            summary_values["Saldo Akhir"] = abs(money_values[-1])

    return pd.DataFrame(
        [
            {"Keterangan": label, "Amount": summary_values.get(label, pd.NA)}
            for label in ["Saldo Awal", "Total Pengeluaran", "Total Pemasukan", "Saldo Akhir"]
            if label in summary_values
        ],
        columns=SUMMARY_COLUMNS,
    )


def parse_bni_account_statement_format(pdf_file):
    records = []
    current = None
    lines = pdf_text_lines(pdf_file)
    df_summary = extract_account_statement_summary_from_text(pdf_file)

    def current_is_complete():
        return (
            current is not None
            and current["AmountCandidates"]
            and clean_text(current["DbCr"])
            and not pd.isna(parse_number(current["Saldo"]))
        )

    def finish_current():
        nonlocal current
        if current is None:
            return

        saldo = parse_number(current["Saldo"])
        if pd.isna(saldo):
            saldo = summary_value(df_summary, "Saldo Akhir")

        amount_candidates = [value for value in current["AmountCandidates"] if not pd.isna(value)]
        amount = max((abs(value) for value in amount_candidates), default=pd.NA)

        if pd.isna(amount) or pd.isna(saldo):
            current = None
            return

        db_cr = clean_text(current["DbCr"]).upper()
        db_amount = pd.NA
        cr_amount = pd.NA
        if db_cr == "D":
            db_amount = amount
        elif db_cr in {"K", "C"}:
            cr_amount = amount

        records.append({
            "Tanggal": current["Tanggal"],
            "Jam": current["Jam"],
            "Keterangan": clean_text(current["Keterangan"]),
            "Mutasi": amount,
            "DB": db_amount,
            "CR": cr_amount,
            "Saldo": saldo,
        })
        current = None

    for line in lines:
        line = clean_text(line)
        upper = line.upper()
        if not line:
            continue

        if upper.startswith("ENDING BALANCE") or upper.startswith("TOTAL DEBET") or upper.startswith("TOTAL CREDIT"):
            finish_current()
            break

        if is_account_statement_header_line(line):
            continue

        if is_account_statement_datetime(line):
            if current_is_complete():
                finish_current()

            tanggal, jam = split_account_statement_datetime(line)
            if current is None:
                current = {
                    "Tanggal": tanggal,
                    "Jam": jam,
                    "Keterangan": "",
                    "AmountCandidates": [],
                    "DbCr": "",
                    "Saldo": "",
                    "SeenEffectiveDate": False,
                }
            elif not current["SeenEffectiveDate"]:
                current["SeenEffectiveDate"] = True
            continue

        if current is None:
            continue

        if clean_text(current["DbCr"]) and not current["Saldo"] and is_account_statement_money(line):
            current["Saldo"] = line
            continue

        if upper in {"D", "K", "C"}:
            current["DbCr"] = upper
            continue

        if not current["DbCr"] and is_account_statement_money(line):
            parsed = parse_number(line)
            if not pd.isna(parsed):
                current["AmountCandidates"].append(parsed)
            continue

        if not current["DbCr"]:
            current["Keterangan"] = append_text(current["Keterangan"], line)

    finish_current()

    df_final = pd.DataFrame(records, columns=TRANSACTION_COLUMNS)

    if not df_final.empty:
        if pd.isna(summary_value(df_summary, "Saldo Awal")):
            first_balance = df_final["Saldo"].iloc[0]
            first_db = df_final["DB"].iloc[0]
            first_cr = df_final["CR"].iloc[0]
            if not pd.isna(first_db):
                df_summary = add_or_update_summary_value(df_summary, "Saldo Awal", first_balance + first_db)
            elif not pd.isna(first_cr):
                df_summary = add_or_update_summary_value(df_summary, "Saldo Awal", first_balance - first_cr)
        if pd.isna(summary_value(df_summary, "Total Pengeluaran")):
            df_summary = add_or_update_summary_value(df_summary, "Total Pengeluaran", sum_amount(df_final, "DB"))
        if pd.isna(summary_value(df_summary, "Total Pemasukan")):
            df_summary = add_or_update_summary_value(df_summary, "Total Pemasukan", sum_amount(df_final, "CR"))
        if pd.isna(summary_value(df_summary, "Saldo Akhir")):
            df_summary = add_or_update_summary_value(df_summary, "Saldo Akhir", df_final["Saldo"].iloc[-1])

    return df_final, df_summary


def parse_bni_transaction_inquiry_format(pdf_file):
    pdf_file = Path(pdf_file)
    records = []
    current = None

    def add_line_to_current(words):
        nonlocal current
        if current is None:
            return

        time_text = find_line_word(words, 55, 130, is_new_time)
        if time_text:
            current["Jam"] = normalize_new_time(time_text)

        amount_words = [
            word for word in words_in_column(words, 570, 670)
            if is_money_line(word[4], signed=None)
        ]
        description_words = [
            word for word in words_in_column(words, 300, 590)
            if word not in amount_words
        ]

        description = words_text(description_words)
        amount = words_text(amount_words)
        db_cr = words_text(words_in_column(words, 670, 720))
        balance = words_text(words_in_column(words, 720, 830))

        if description:
            current["Keterangan"] = append_text(current["Keterangan"], description)
        if amount and not current["AmountText"]:
            current["AmountText"] = amount
        if db_cr and not current["DbCr"]:
            current["DbCr"] = db_cr
        if balance and not current["BalanceText"]:
            current["BalanceText"] = balance

    def finish_current():
        nonlocal current
        if current is None:
            return

        amount = parse_number(current["AmountText"])
        balance = parse_number(current["BalanceText"])
        if pd.isna(amount) or pd.isna(balance):
            current = None
            return

        amount = abs(amount)
        db_cr = clean_text(current["DbCr"]).upper()
        db_amount = pd.NA
        cr_amount = pd.NA

        if db_cr == "D":
            db_amount = amount
        elif db_cr == "C":
            cr_amount = amount
        elif clean_text(current["AmountText"]).startswith("-"):
            db_amount = amount
        elif clean_text(current["AmountText"]).startswith("+"):
            cr_amount = amount

        records.append({
            "Tanggal": new_date_to_day_month(current["Tanggal"]),
            "Jam": current["Jam"],
            "Keterangan": clean_text(current["Keterangan"]),
            "Mutasi": amount,
            "DB": db_amount,
            "CR": cr_amount,
            "Saldo": balance,
        })
        current = None

    with fitz.open(pdf_file) as doc:
        for page in doc:
            in_table = False
            for words in group_pdf_words_into_lines(page.get_text("words")):
                if is_new_table_header(words):
                    in_table = True
                    continue
                if not in_table:
                    continue

                transaction_no = find_line_word(words, 25, 55, lambda value: bool(re.fullmatch(r"\d+", value)))
                transaction_date = find_line_word(words, 55, 130, is_new_date)

                if transaction_no and transaction_date:
                    finish_current()
                    current = {
                        "No": transaction_no,
                        "Tanggal": transaction_date,
                        "Jam": "",
                        "Keterangan": "",
                        "AmountText": "",
                        "DbCr": "",
                        "BalanceText": "",
                    }
                    add_line_to_current(words)
                    continue

                add_line_to_current(words)

    finish_current()

    df_final = pd.DataFrame(records, columns=TRANSACTION_COLUMNS)
    df_summary = extract_new_summary_from_pdf(pdf_file)

    if not df_final.empty:
        df_summary = add_or_update_summary_value(df_summary, "Saldo Akhir", df_final["Saldo"].iloc[-1])
        if pd.isna(summary_value(df_summary, "Total Pengeluaran")):
            df_summary = add_or_update_summary_value(df_summary, "Total Pengeluaran", sum_amount(df_final, "DB"))
        if pd.isna(summary_value(df_summary, "Total Pemasukan")):
            df_summary = add_or_update_summary_value(df_summary, "Total Pemasukan", sum_amount(df_final, "CR"))

    return df_final, df_summary


def parse_bni_new_format(pdf_file):
    if is_account_statement_format(pdf_file):
        return parse_bni_account_statement_format(pdf_file)
    return parse_bni_transaction_inquiry_format(pdf_file)


def read_pdf_table(pdf_file):
    dfs = tabula.read_pdf(
        str(pdf_file),
        pages="all",
        stream=True,
        guess=False,
        area=PDF_AREA,
        columns=PDF_COLUMNS,
        pandas_options={"header": None},
        multiple_tables=False,
        force_subprocess=True,
        encoding="cp1252",
    )

    if not dfs:
        raise ValueError("No table found. Try adjusting PDF_AREA / PDF_COLUMNS.")

    df = pd.concat(dfs, ignore_index=True)

    print(f"Jumlah kolom hasil Tabula: {df.shape[1]}")

    rows = []

    for _, row in df.iterrows():
        values = [clean_text(v) for v in row.tolist()]
        values = values + [""] * (5 - len(values))

        tanggal_waktu = values[0]
        rincian = ""
        nominal = ""
        saldo = ""

        if df.shape[1] == 3:
            # usually: tanggal/waktu | rincian | nominal+saldo
            tanggal_waktu = values[0]
            rincian = values[1]
            right_text = values[2]

            remaining, extracted_nominal, extracted_saldo = extract_nominal_saldo(right_text)

            if remaining:
                rincian = append_text(rincian, remaining)

            nominal = extracted_nominal
            saldo = extracted_saldo

        elif df.shape[1] == 4:
            # usually: tanggal/waktu | rincian | nominal | saldo
            tanggal_waktu = values[0]
            rincian = values[1]
            nominal = values[2]
            saldo = values[3]

        else:
            # if Tabula returns 5+ columns:
            # tanggal | waktu | rincian | nominal | saldo
            tanggal_waktu = append_text(values[0], values[1])
            rincian = values[2]
            nominal = values[3]
            saldo = values[4]

        rows.append({
            "TanggalWaktu": tanggal_waktu,
            "Rincian": rincian,
            "Nominal": nominal,
            "Saldo": saldo,
        })

    result = pd.DataFrame(rows, columns=["TanggalWaktu", "Rincian", "Nominal", "Saldo"])

    for col in result.columns:
        result[col] = result[col].apply(clean_text)

    return result


# =========================
# CLEANING
# =========================
def remove_garbage_rows(df):
    garbage_regex = (
        r"Laporan Mutasi Rekening|"
        r"Periode:|"
        r"Otoritas Jasa Keuangan|"
        r"peserta penjaminan|"
        r"Tanggal & Waktu|"
        r"Rincian Transaksi|"
        r"Nominal \(IDR\)|"
        r"Saldo \(IDR\)|"
        r"Informasi Lainnya|"
        r"Apabila terdapat kesalahan|"
        r"BNI dapat sewaktu|"
        r"Dokumen ini dibuat|"
        r"^\d+\s+dari\s+\d+$"
    )

    def is_garbage(row):
        joined = clean_text(" ".join(row.astype(str).tolist()))

        if not joined:
            return True

        return bool(re.search(garbage_regex, joined, flags=re.IGNORECASE))

    cleaned = df[~df.apply(is_garbage, axis=1)].reset_index(drop=True)

    footer_start = None
    for idx, row in cleaned.iterrows():
        joined = clean_text(" ".join(row.astype(str).tolist()))
        normalized = re.sub(r"\s+", "", joined).upper()
        if "INFORMASILAINNYA" in normalized or joined.upper().startswith("1.APABILA"):
            footer_start = idx
            break

    if footer_start is not None:
        cleaned = cleaned.iloc[:footer_start].copy()

    return cleaned.reset_index(drop=True)


def extract_overview_amounts(text):
    tokens = re.findall(r"[+-]?\d[\d,]*(?:\.\d+)?", clean_text(text))
    amounts = []

    for token in tokens:
        if (
            amounts
            and re.fullmatch(r"\d", token)
            and re.search(r",\d{2}$", amounts[-1])
        ):
            amounts[-1] = f"{amounts[-1]}{token}"
        else:
            amounts.append(token)

    return amounts


def split_summary_rows(df):
    summary_rows = []
    drop_indexes = set()

    keywords = [
        "Saldo Awal",
        "Total Pemasukan",
        "Total Pengeluaran",
        "Saldo Akhir",
    ]

    for i in range(len(df)):
        row = df.iloc[i]

        cells = [
            clean_text(row["TanggalWaktu"]),
            clean_text(row["Rincian"]),
            clean_text(row["Nominal"]),
            clean_text(row["Saldo"]),
        ]

        joined = clean_text(" ".join([c for c in cells if c]))
        normalized_joined = re.sub(r"\s+", "", joined).upper()

        if (
            "SALDOAWAL" in normalized_joined
            and "TOTALPEMASUKAN" in normalized_joined
            and "TOTALPENGELUARAN" in normalized_joined
            and "SALDOAKHIR" in normalized_joined
            and i + 1 < len(df)
        ):
            next_joined = clean_text(" ".join(df.iloc[i + 1].astype(str).tolist()))
            amounts = extract_overview_amounts(next_joined)
            if len(amounts) >= 4:
                overview = [
                    ("Saldo Awal", amounts[0]),
                    ("Total Pemasukan", amounts[1]),
                    ("Total Pengeluaran", amounts[2]),
                    ("Saldo Akhir", amounts[3]),
                ]
                for label, amount in overview:
                    summary_rows.append({
                        "Keterangan": label,
                        "Amount": abs(parse_number(amount)) if not pd.isna(parse_number(amount)) else pd.NA,
                    })
                drop_indexes.update({i, i + 1})
                continue

        for keyword in keywords:
            if keyword.upper() in joined.upper():
                amount = ""

                # case 1: keyword and amount in the same row
                same_row_text = joined.upper().replace(keyword.upper(), "").strip()
                if same_row_text:
                    amount = same_row_text
                    if pd.isna(parse_number(amount)):
                        amount = extract_amount_text(same_row_text)

                # case 2: amount is in another column on same row
                for cell in cells:
                    if cell != keyword and not pd.isna(parse_number(cell)):
                        amount = cell
                        break

                # case 3: amount is in the next few rows
                if not amount:
                    for j in range(i + 1, min(i + 4, len(df))):
                        next_joined = clean_text(" ".join(df.iloc[j].astype(str).tolist()))
                        parsed = parse_number(extract_amount_text(next_joined))

                        if not pd.isna(parsed):
                            amount = extract_amount_text(next_joined)
                            drop_indexes.add(j)
                            break

                if not amount:
                    amount = extract_amount_text(joined)

                summary_rows.append({
                    "Keterangan": keyword,
                    "Amount": abs(parse_number(amount)) if not pd.isna(parse_number(amount)) else pd.NA,
                })

                drop_indexes.add(i)
                break

    df_summary = pd.DataFrame(summary_rows, columns=["Keterangan", "Amount"])

    # remove duplicate summary rows, keep first occurrence
    df_summary = df_summary.drop_duplicates(subset=["Keterangan"], keep="first").reset_index(drop=True)

    df_trans = df.drop(index=list(drop_indexes), errors="ignore").reset_index(drop=True)

    # remove table saldo awal / saldo akhir rows if still left
    def is_table_saldo_row(row):
        joined = clean_text(" ".join(row.astype(str).tolist()))
        return bool(
            re.match(
                r"^Saldo Awal\s+[\d,+-]+$|^Saldo Akhir\s+[\d,+-]+$",
                joined,
                flags=re.IGNORECASE
            )
        )

    df_trans = df_trans[~df_trans.apply(is_table_saldo_row, axis=1)].reset_index(drop=True)

    return df_trans, df_summary


def merge_transactions(df):
    """
    BNI format:
    02 Mar 2026
    16:53:00 WIB
    Pembayaran Qris
    ALGO J942 AFM RS POLRI JAKARTA TIMURID  -96,600  3,930,095
    """
    records = []
    current = None

    for _, row in df.iterrows():
        tanggal_waktu = clean_text(row["TanggalWaktu"])
        rincian = clean_text(row["Rincian"])
        nominal = clean_text(row["Nominal"])
        saldo = clean_text(row["Saldo"])

        joined = clean_text(" ".join([tanggal_waktu, rincian, nominal, saldo]))

        if not joined:
            continue

        # start new transaction
        if is_date(tanggal_waktu):
            if current is not None:
                records.append(current)

            current = {
                "Tanggal": tanggal_waktu,
                "Jam": "",
                "Keterangan": "",
                "Mutasi": "",
                "Saldo": "",
            }

            if rincian:
                current["Keterangan"] = append_text(current["Keterangan"], rincian)

            if nominal:
                current["Mutasi"] = nominal

            if saldo:
                current["Saldo"] = saldo

            continue

        if current is None:
            continue

        # time row
        if is_time(tanggal_waktu):
            current["Jam"] = tanggal_waktu

            if rincian:
                current["Keterangan"] = append_text(current["Keterangan"], rincian)

            if nominal:
                current["Mutasi"] = nominal

            if saldo:
                current["Saldo"] = saldo

            continue

        # continuation row
        if tanggal_waktu:
            current["Keterangan"] = append_text(current["Keterangan"], tanggal_waktu)

        if rincian:
            current["Keterangan"] = append_text(current["Keterangan"], rincian)

        if nominal:
            if current["Mutasi"] == "":
                current["Mutasi"] = nominal
            else:
                current["Keterangan"] = append_text(current["Keterangan"], nominal)

        if saldo:
            if current["Saldo"] == "":
                current["Saldo"] = saldo
            else:
                current["Keterangan"] = append_text(current["Keterangan"], saldo)

    if current is not None:
        records.append(current)

    return pd.DataFrame(records, columns=["Tanggal", "Jam", "Keterangan", "Mutasi", "Saldo"])


def finalize_transactions(df):
    df = df.copy()

    for col in ["Tanggal", "Jam", "Keterangan", "Mutasi", "Saldo"]:
        if col not in df.columns:
            df[col] = ""

    if df.empty:
        return empty_transactions_dataframe()

    df[["DB", "CR"]] = df["Mutasi"].apply(split_db_cr)
    df["Mutasi"] = df["Mutasi"].apply(
        lambda value: abs(parse_number(value)) if not pd.isna(parse_number(value)) else pd.NA
    )
    df["Saldo"] = df["Saldo"].apply(parse_number)
    df["Tanggal"] = df["Tanggal"].apply(date_to_day_month)
    df["Jam"] = df["Jam"].apply(clean_text)
    df["Keterangan"] = df["Keterangan"].apply(clean_text)

    df = df[TRANSACTION_COLUMNS].copy()
    df = df[~df.apply(is_saldo_awal_transaction, axis=1)].reset_index(drop=True)

    return df


def is_saldo_awal_transaction(row):
    description = clean_text(row.get("Keterangan", "")).upper()
    return (
        "SALDO AWAL" in description
        and pd.isna(row.get("DB", pd.NA))
        and pd.isna(row.get("CR", pd.NA))
    )


def summary_value(df_summary, label):
    if df_summary.empty:
        return pd.NA

    mask = df_summary["Keterangan"].str.upper().eq(label.upper())
    if not mask.any():
        return pd.NA

    return df_summary.loc[mask, "Amount"].iloc[0]


def reconcile_transactions_with_balance(clean_df, df_summary):
    if clean_df.empty:
        return clean_df

    result = clean_df.copy()
    last_balance = summary_value(df_summary, "Saldo Awal")

    for idx, row in result.iterrows():
        current_balance = row["Saldo"]
        if pd.isna(current_balance) or pd.isna(last_balance):
            last_balance = current_balance
            continue

        delta = current_balance - last_balance
        if delta < 0:
            result.at[idx, "DB"] = abs(delta)
            result.at[idx, "CR"] = pd.NA
        elif delta > 0:
            result.at[idx, "DB"] = pd.NA
            result.at[idx, "CR"] = delta

        last_balance = current_balance

    return result


def report_extraction_balance(clean_df, df_summary, pdf_file):
    extracted_db = sum(clean_df["DB"].dropna(), Decimal("0.00"))
    extracted_cr = sum(clean_df["CR"].dropna(), Decimal("0.00"))
    summary_db = summary_value(df_summary, "Total Pengeluaran")
    summary_cr = summary_value(df_summary, "Total Pemasukan")

    mismatches = []
    if not pd.isna(summary_db) and extracted_db != summary_db:
        mismatches.append(f"DB extracted {extracted_db} but PDF summary {summary_db}")
    if not pd.isna(summary_cr) and extracted_cr != summary_cr:
        mismatches.append(f"CR extracted {extracted_cr} but PDF summary {summary_cr}")

    if mismatches:
        print(f"PERINGATAN: {Path(pdf_file).name}: " + "; ".join(mismatches))


# =========================
# EXPORT
# =========================
def auto_fit_columns(sheet):
    for col_idx, col_cells in enumerate(sheet.columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_length = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in col_cells
        )
        sheet.column_dimensions[col_letter].width = max_length + 2


def safe_sheet_name(name):
    cleaned = re.sub(r"[\[\]\:\*\?\/\\]", "_", clean_text(name)).strip()
    return cleaned[:31] or "Sheet"


def unique_sheet_name(base_name, used_names):
    base_name = safe_sheet_name(base_name)
    candidate = base_name
    counter = 2

    while candidate.lower() in used_names:
        suffix = f"_{counter}"
        candidate = f"{base_name[:31 - len(suffix)]}{suffix}"
        counter += 1

    used_names.add(candidate.lower())
    return candidate


def extract_month_year(pdf_file):
    metadata = extract_pdf_metadata(pdf_file)
    return metadata["month_label"], metadata["month_order"], metadata["year"]


def extract_account_number(pdf_file):
    return extract_pdf_metadata(pdf_file)["account"]


def sum_amount(df, column):
    values = df[column].dropna()
    if values.empty:
        return pd.NA
    return sum(values, Decimal("0.00"))


def count_amount(df, column):
    return int(df[column].dropna().count())


def sum_by_description(df, amount_column, pattern):
    if df.empty:
        return pd.NA

    mask = df["Keterangan"].str.upper().str.contains(pattern, regex=True, na=False)
    values = df.loc[mask, amount_column].dropna()

    if values.empty:
        return pd.NA
    return sum(values, Decimal("0.00"))


def build_month_summary(clean_df, df_summary, pdf_file, metadata=None):
    metadata = metadata or extract_pdf_metadata(pdf_file)

    return {
        "Source File": pdf_file.name,
        "Account": metadata["account"],
        "Year": metadata["year"],
        "MonthOrder": metadata["month_order"],
        "Bulan": metadata["month_label"],
        "Mutasi Debet Nominal (Rp)": summary_value(df_summary, "Total Pengeluaran"),
        "Mutasi Debet Frek": count_amount(clean_df, "DB"),
        "Mutasi Kredit Nominal (Rp)": summary_value(df_summary, "Total Pemasukan"),
        "Mutasi Kredit Frek": count_amount(clean_df, "CR"),
        "Saldo (Rp)": summary_value(df_summary, "Saldo Akhir"),
        "Saldo Awal (Rp)": summary_value(df_summary, "Saldo Awal"),
        "Adm": sum_by_description(clean_df, "DB", r"\bADM\b|ADMIN|BIAYA\s+ADMIN"),
        "Pajak": sum_by_description(clean_df, "DB", r"\bPPH\b|PAJAK|TAX"),
        "Bunga": sum_by_description(clean_df, "CR", r"BUNGA|INTEREST"),
        "Saldo Min": sum_by_description(clean_df, "DB", r"SALDO\s+MIN"),
        "JaGir": sum_by_description(clean_df, "CR", r"JASA\s+GIRO|JAGIR"),
    }


def first_day_label(pdf_file, clean_df):
    _, month_order, _ = extract_month_year(pdf_file)
    if month_order != 99:
        return f"01/{month_order:02d}"

    if not clean_df.empty:
        match = re.match(r"^\d{1,2}/\d{1,2}", clean_text(clean_df["Tanggal"].iloc[0]))
        if match:
            return match.group(0)

    return ""


def build_monthly_display_dataframe(clean_df, df_summary, pdf_file):
    saldo_awal = summary_value(df_summary, "Saldo Awal")
    opening_row = pd.DataFrame([{
        "Tanggal": first_day_label(pdf_file, clean_df),
        "Jam": "",
        "Keterangan": "SALDO AWAL",
        "Mutasi": pd.NA,
        "DB": pd.NA,
        "CR": pd.NA,
        "Saldo": saldo_awal,
    }], columns=TRANSACTION_COLUMNS)

    clean_df = clean_df.copy()
    for col in TRANSACTION_COLUMNS:
        if col not in clean_df.columns:
            clean_df[col] = pd.NA

    return pd.concat([opening_row, clean_df[TRANSACTION_COLUMNS]], ignore_index=True)


def add_monthly_sheet_totals(sheet):
    first_data_row = 2
    last_data_row = sheet.max_row
    total_row = last_data_row + 3
    freq_row = total_row + 1
    average_row = total_row + 2

    for row_num, label in {
        total_row: "Total",
        freq_row: "Frekuensi",
        average_row: "Rata2",
    }.items():
        sheet.cell(row=row_num, column=3, value=label)

    for col_num in (4, 5, 6):
        col_letter = get_column_letter(col_num)
        sheet.cell(row=total_row, column=col_num, value=f"=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})")
        sheet.cell(row=freq_row, column=col_num, value=f"=COUNT({col_letter}{first_data_row}:{col_letter}{last_data_row})")
        sheet.cell(row=average_row, column=col_num, value=f"=IFERROR(AVERAGE({col_letter}{first_data_row}:{col_letter}{last_data_row}),0)")


def style_monthly_sheet(sheet):
    header_fill = PatternFill("solid", fgColor="BFBFBF")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    summary_start = sheet.max_row - 2

    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False

    for col_letter, width in {
        "A": 10,
        "B": 14,
        "C": 95,
        "D": 16,
        "E": 16,
        "F": 16,
        "G": 16,
    }.items():
        sheet.column_dimensions[col_letter].width = width

    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=7):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

            if cell.row == 1:
                cell.fill = header_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif cell.column in (4, 5, 6, 7):
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif cell.column == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for row_num in range(summary_start, sheet.max_row + 1):
        label_cell = sheet.cell(row=row_num, column=3)
        label_cell.fill = header_fill
        label_cell.font = Font(bold=True)
        label_cell.alignment = Alignment(horizontal="center", vertical="center")

        for col_num in (4, 5, 6):
            value_cell = sheet.cell(row=row_num, column=col_num)
            value_cell.number_format = "#,##0.00" if row_num != summary_start + 1 else "0"
            value_cell.alignment = Alignment(horizontal="right", vertical="center")

    for row_num in range(1, sheet.max_row + 1):
        sheet.row_dimensions[row_num].height = 18


def write_monthly_transaction_sheet(writer, clean_df, df_summary, pdf_file, sheet_name):
    display_df = build_monthly_display_dataframe(clean_df, df_summary, pdf_file)
    display_df = dataframe_for_excel(display_df)
    display_df.to_excel(writer, sheet_name=sheet_name, index=False)
    sheet = writer.sheets[sheet_name]
    add_monthly_sheet_totals(sheet)
    style_monthly_sheet(sheet)


def write_summary_sheet(writer, summary_rows):
    workbook = writer.book
    sheet = workbook.create_sheet("Summary", 0)
    writer.sheets["Summary"] = sheet

    headers_top = [
        "No", "Bulan", "Mutasi Debet", None, "Mutasi Kredit", None,
        "Saldo", "Saldo Awal", "Adm", "Pajak", "Bunga", "Saldo Min", "JaGir",
    ]
    headers_bottom = [
        "No", "Bulan", "Nominal (Rp)", "Frek", "Nominal (Rp)", "Frek",
        "(Rp)", "(Rp)", "Adm", "Pajak", "Bunga", "Saldo Min", "JaGir",
    ]

    for col_num, value in enumerate(headers_top, start=1):
        if value is not None:
            sheet.cell(row=1, column=col_num, value=value)
    for col_num, value in enumerate(headers_bottom, start=1):
        sheet.cell(row=2, column=col_num, value=value)

    sheet.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    sheet.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)
    sheet.merge_cells(start_row=1, start_column=3, end_row=1, end_column=4)
    sheet.merge_cells(start_row=1, start_column=5, end_row=1, end_column=6)
    for col_num in range(7, 14):
        sheet.merge_cells(start_row=1, start_column=col_num, end_row=2, end_column=col_num)

    sorted_rows = sorted(summary_rows, key=lambda row: (row["Year"], row["MonthOrder"], row["Source File"]))
    for row_idx, row in enumerate(sorted_rows, start=3):
        values = [
            row_idx - 2,
            row["Bulan"],
            row["Mutasi Debet Nominal (Rp)"],
            row["Mutasi Debet Frek"],
            row["Mutasi Kredit Nominal (Rp)"],
            row["Mutasi Kredit Frek"],
            row["Saldo (Rp)"],
            row["Saldo Awal (Rp)"],
            row["Adm"],
            row["Pajak"],
            row["Bunga"],
            row["Saldo Min"],
            row["JaGir"],
        ]
        for col_num, value in enumerate(values, start=1):
            sheet.cell(row=row_idx, column=col_num, value=excel_number(value))

    total_row = len(sorted_rows) + 3
    average_row = total_row + 1
    sheet.cell(row=total_row, column=2, value="Total")
    sheet.cell(row=average_row, column=2, value="Rata-Rata")

    data_start = 3
    data_end = total_row - 1
    if data_end >= data_start:
        for col_num in [3, 4, 5, 6, 9, 10, 11, 12, 13]:
            col_letter = get_column_letter(col_num)
            sheet.cell(row=total_row, column=col_num, value=f"=SUM({col_letter}{data_start}:{col_letter}{data_end})")
        for col_num in [3, 4, 5, 6, 7]:
            col_letter = get_column_letter(col_num)
            sheet.cell(row=average_row, column=col_num, value=f"=AVERAGE({col_letter}{data_start}:{col_letter}{data_end})")

    style_summary_sheet(sheet)
    auto_fit_columns(sheet)


def style_summary_sheet(sheet):
    header_fill = PatternFill("solid", fgColor="BFBFBF")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in sheet.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            if cell.row in (1, 2):
                cell.fill = header_fill
                cell.font = Font(bold=True)
            elif cell.row >= 3 and cell.column in (3, 5, 7, 8, 9, 10, 11, 12, 13):
                cell.number_format = "#,##0.00"
            elif cell.row >= 3 and cell.column in (4, 6):
                cell.number_format = "0"

    for row_num in (sheet.max_row - 1, sheet.max_row):
        for col_num in range(1, sheet.max_column + 1):
            sheet.cell(row=row_num, column=col_num).font = Font(bold=True)


def parse_bni_old_format(pdf_file):
    print("Membaca teks PDF BNI format lama...")
    df_summary = build_summary_from_text(pdf_file)

    print("Mengambil transaksi asli dari isi PDF...")
    df_merged = parse_transactions_from_text(pdf_file)

    print("Finalisasi DB/CR dan saldo...")
    df_final = finalize_transactions(df_merged)
    df_final = reconcile_transactions_with_balance(df_final, df_summary)
    report_extraction_balance(df_final, df_summary, pdf_file)

    return df_final, df_summary


def amount_is_present(value):
    if pd.isna(value):
        return False
    if isinstance(value, Decimal):
        return True
    return not pd.isna(parse_number(value))


def validate_bni_transactions(df):
    if df is None or df.empty:
        raise ValueError("No BNI transactions were extracted.")

    missing_columns = [col for col in TRANSACTION_COLUMNS if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing transaction columns: {', '.join(missing_columns)}")

    if df["Tanggal"].apply(clean_text).eq("").all():
        raise ValueError("Tanggal is empty for all extracted transactions.")

    if df["Keterangan"].apply(clean_text).eq("").all():
        raise ValueError("Keterangan is empty for all extracted transactions.")

    amount_mask = df[["Mutasi", "DB", "CR"]].apply(
        lambda row: any(amount_is_present(value) for value in row),
        axis=1,
    )
    if not amount_mask.all():
        raise ValueError("Some extracted transactions do not contain Mutasi, DB, or CR amounts.")

    invalid_saldo = [
        value for value in df["Saldo"]
        if not pd.isna(value) and not isinstance(value, Decimal) and pd.isna(parse_number(value))
    ]
    if invalid_saldo:
        raise ValueError("Saldo contains non-numeric values.")

    header_regex = (
        r"TANGGAL\s*&\s*WAKTU|"
        r"POST\s+DATE|"
        r"POSTING\s+DATE|"
        r"EFFECTIVE\s+DATE|"
        r"RINCIAN\s+TRANSAKSI|"
        r"TRANSACTION\s+DESCRIPTION|"
        r"NOMINAL\s+\(IDR\)|"
        r"SALDO\s+\(IDR\)|"
        r"\bDB/CR\b"
    )
    header_rows = df.apply(
        lambda row: bool(re.search(
            header_regex,
            clean_text(f"{row.get('Tanggal', '')} {row.get('Keterangan', '')}"),
            flags=re.IGNORECASE,
        )),
        axis=1,
    )
    if header_rows.any():
        raise ValueError("Header rows remain in transaction output.")

    return True


def parser_error_message(errors):
    detail = "; ".join(errors)
    message = (
        "BNI PDF format not recognized or table extraction failed. "
        "Please check whether the PDF is text-based and not scanned. "
        "OCR is not currently supported."
    )

    if re.search(r"\b(java|jvm|tabula|jpype)\b", detail, flags=re.IGNORECASE):
        message += " Java is required for tabula-py based extraction."

    if detail:
        message += f" Details: {detail}"

    return message


def parse_bni_with_fallback(pdf_file):
    pdf_file = Path(pdf_file)
    format_type = detect_bni_format(pdf_file)
    print(f"Format BNI terdeteksi: {format_type or 'unknown'}")

    if format_type == "old":
        parser_order = [("old", parse_bni_old_format), ("new", parse_bni_new_format)]
    elif format_type == "new":
        parser_order = [("new", parse_bni_new_format), ("old", parse_bni_old_format)]
    else:
        parser_order = [("old", parse_bni_old_format), ("new", parse_bni_new_format)]

    errors = []
    for parser_name, parser in parser_order:
        try:
            df_final, df_summary = parser(pdf_file)
            validate_bni_transactions(df_final)
            return df_final, df_summary
        except Exception as exc:
            errors.append(f"{parser_name}: {exc}")

    raise ValueError(parser_error_message(errors))


def process_pdf(pdf_file):
    return parse_bni_with_fallback(pdf_file)


def export_to_excel(df_final, df_summary, output_file, pdf_file=None):
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        if pdf_file is None:
            dataframe_for_excel(df_final).to_excel(writer, sheet_name="Transaksi", index=False)
            style_monthly_sheet(writer.sheets["Transaksi"])
        else:
            write_monthly_transaction_sheet(writer, df_final, df_summary, pdf_file, "Transaksi")
        dataframe_for_excel(df_summary).to_excel(writer, sheet_name="Summary", index=False)

        auto_fit_columns(writer.sheets["Summary"])


def export_year_workbook(extracted_files, output_file):
    output_file.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        used_names = set()
        write_summary_sheet(writer, [item["summary_row"] for item in extracted_files])

        for item in extracted_files:
            sheet_name = unique_sheet_name(item["sheet_name"], used_names)
            write_monthly_transaction_sheet(
                writer,
                item["transactions"],
                item["summary"],
                item["pdf_file"],
                sheet_name,
            )


def discover_pdf_groups(input_path):
    input_path = Path(input_path)
    if input_path.is_file():
        return {input_path.parent: [input_path]}

    pdf_files = sorted(input_path.rglob("*.pdf"))
    groups = {}
    for pdf_file in pdf_files:
        groups.setdefault(pdf_file.parent, []).append(pdf_file)
    return groups


def discover_pdf_groups_by_content(input_path):
    input_path = Path(input_path)
    pdf_files = [input_path] if input_path.is_file() else sorted(input_path.rglob("*.pdf"))
    groups = {}

    for pdf_file in pdf_files:
        metadata = extract_pdf_metadata(pdf_file)
        key = (metadata["account"], metadata["year"])
        if key not in groups:
            groups[key] = {
                "account": metadata["account"],
                "year": metadata["year"],
                "folder": metadata["year"],
                "files": [],
            }
        groups[key]["files"].append({"path": pdf_file, "metadata": metadata})

    return groups


def ensure_xlsx_name(name):
    name = clean_text(name) or "{account}_{year}.xlsx"
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name


def output_name_for_group(group_info, output_name_template=None, force_unique=False):
    account = group_info["account"]
    year = group_info["year"]
    folder_name = group_info.get("folder", year)

    if not output_name_template:
        return f"{account}_{year}.xlsx"

    output_name = ensure_xlsx_name(output_name_template)
    has_year_or_folder = any(token in output_name for token in ("{year}", "{folder}"))
    try:
        output_name = output_name.format(account=account, year=year, folder=folder_name)
    except (KeyError, ValueError) as exc:
        raise ValueError("Output name template only supports {account}, {year}, and {folder}.") from exc

    if force_unique and not has_year_or_folder:
        output_path = Path(output_name)
        output_name = f"{output_path.stem}_{year}{output_path.suffix}"

    return output_name


def convert_pdf(pdf_file, output_file):
    """
    Convert one BNI PDF statement into one Excel workbook.

    The parser auto-detects legacy and new BNI statement layouts.
    """
    pdf_file = Path(pdf_file)
    output_file = Path(output_file)
    df_final, df_summary = process_pdf(pdf_file)
    print("Menulis ke Excel...")
    export_to_excel(df_final, df_summary, output_file, pdf_file)
    print(f"Data successfully written to {output_file}")
    return output_file


def run_single_file(pdf_file, output_file):
    convert_pdf(pdf_file, output_file)


def run_folder(input_path, output_dir, output_name_template=None, preserve_relative_folders=True):
    groups = discover_pdf_groups_by_content(input_path)
    written_files = []

    if not groups:
        raise FileNotFoundError(f"No PDF files found in folder: {input_path}")

    for _, group_info in sorted(groups.items(), key=lambda item: (item[0][1], item[0][0])):
        file_items = group_info["files"]
        print(
            f"\nMemproses rekening {group_info['account']} tahun {group_info['year']} "
            f"({len(file_items)} file PDF)"
        )
        extracted_files = []

        for file_item in sorted(file_items, key=lambda item: (item["metadata"]["month_order"], item["path"].name)):
            pdf_file = file_item["path"]
            metadata = file_item["metadata"]
            print(f"Memproses file: {pdf_file.name}")
            df_final, df_summary = process_pdf(pdf_file)

            extracted_files.append({
                "pdf_file": pdf_file,
                "transactions": df_final,
                "summary": df_summary,
                "summary_row": build_month_summary(df_final, df_summary, pdf_file, metadata),
                "sheet_name": metadata["month_label"].replace("-", "_"),
                "month_order": metadata["month_order"],
                "year": metadata["year"],
            })

        relative_folder = Path(group_info["year"]) if preserve_relative_folders else Path()
        output_folder = output_dir / relative_folder
        output_file = output_folder / output_name_for_group(
            group_info,
            output_name_template,
            force_unique=len(groups) > 1,
        )
        export_year_workbook(extracted_files, output_file)
        written_files.append(output_file)
        print(f"Workbook selesai dibuat: {output_file}")

    return written_files


def choose_batch_options_with_gui():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, simpledialog
    except ImportError:
        return None, None, None

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    messagebox.showinfo(
        "BNI Taplus Muda PDF to Excel",
        "Choose the BNI Taplus Muda folder that contains your monthly PDF files. Subfolders will be included.",
        parent=root,
    )
    input_dir = filedialog.askdirectory(
        title="Choose BNI Taplus Muda PDF folder",
        initialdir=str(BNI_DIR.resolve()) if BNI_DIR.exists() else str(Path.cwd()),
        parent=root,
    )
    if not input_dir:
        root.destroy()
        return None, None, None

    messagebox.showinfo(
        "BNI Taplus Muda PDF to Excel",
        "Choose the folder where the Excel output should be saved.",
        parent=root,
    )
    output_dir = filedialog.askdirectory(
        title="Choose Excel output folder",
        initialdir=str(EXCEL_DIR.resolve()) if EXCEL_DIR.exists() else str(Path.cwd()),
        parent=root,
    )
    if not output_dir:
        root.destroy()
        return None, None, None

    output_name = simpledialog.askstring(
        "BNI Taplus Muda PDF to Excel",
        "Enter Excel filename pattern:\nUse {account}, {year}, or {folder} if needed.",
        initialvalue="{account}_{year}.xlsx",
        parent=root,
    )
    if output_name is None:
        root.destroy()
        return None, None, None

    root.destroy()
    return Path(input_dir), Path(output_dir), ensure_xlsx_name(output_name)


def choose_batch_options_with_console():
    input_dir = input("Choose input folder path: ").strip()
    output_dir = input("Choose output folder path: ").strip()
    output_name = input("Excel filename pattern [{account}_{year}.xlsx]: ").strip()
    return Path(input_dir), Path(output_dir), ensure_xlsx_name(output_name or "{account}_{year}.xlsx")


def show_gui_result(message, is_error=False):
    try:
        import tkinter as tk
        from tkinter import messagebox
    except ImportError:
        return

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    if is_error:
        messagebox.showerror("BNI Taplus Muda PDF to Excel", message, parent=root)
    else:
        messagebox.showinfo("BNI Taplus Muda PDF to Excel", message, parent=root)

    root.destroy()


def parse_args():
    parser = argparse.ArgumentParser(description="Convert BNI Taplus Muda mutasi PDF files to Excel.")
    parser.add_argument(
        "input",
        nargs="?",
        default=None,
        help="PDF filename without extension, PDF file path, or folder path. No arguments opens folder dialogs.",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=str(EXCEL_DIR),
        help="Output folder for generated Excel files. Default: excel_file/BNI Taplus Muda",
    )
    parser.add_argument(
        "--output-name",
        default=None,
        help="Workbook filename pattern. Supports {account}, {year}, {folder}. Default: {account}_{year}.xlsx",
    )
    parser.add_argument(
        "--flat-output",
        action="store_true",
        help="Save all yearly workbooks directly in output-dir instead of mirroring input subfolders.",
    )
    return parser.parse_args()


def resolve_input_path(input_value):
    input_path = Path(input_value)

    if input_path.exists():
        return input_path

    candidate_pdf = Path("pdf_file") / f"{input_path}.pdf"
    candidate_folder = Path("pdf_file") / input_path
    if candidate_pdf.exists():
        return candidate_pdf
    if candidate_folder.exists():
        return candidate_folder

    raise FileNotFoundError(f"Input not found: {input_path}")


# =========================
# MAIN
# =========================
def main():
    args = parse_args()
    use_gui = len(sys.argv) == 1

    if use_gui:
        selected = choose_batch_options_with_gui()
        if selected == (None, None, None):
            selected = choose_batch_options_with_console()
        input_path, excel_dir, output_name_template = selected
    else:
        input_path = resolve_input_path(args.input or BNI_DIR)
        excel_dir = Path(args.output_dir)
        output_name_template = args.output_name

    excel_dir.mkdir(parents=True, exist_ok=True)

    if input_path.is_file():
        output_name = ensure_xlsx_name(output_name_template or f"{input_path.stem}.xlsx")
        try:
            output_name = output_name.format(
                account=extract_account_number(input_path),
                year=extract_month_year(input_path)[2],
                folder=input_path.parent.name,
            )
        except (KeyError, ValueError) as exc:
            raise ValueError("Output name template only supports {account}, {year}, and {folder}.") from exc
        run_single_file(input_path, excel_dir / output_name)
    else:
        written_files = run_folder(
            input_path,
            excel_dir,
            output_name_template,
            preserve_relative_folders=not getattr(args, "flat_output", False),
        )
        print("\nSemua proses selesai.")
        for output_file in written_files:
            print(f"- {output_file}")

        if use_gui:
            show_gui_result("Done.\n\nSaved files:\n" + "\n".join(str(path) for path in written_files))


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Error: {exc}")
        if len(sys.argv) == 1:
            show_gui_result(str(exc), is_error=True)
            sys.exit(1)
        raise
