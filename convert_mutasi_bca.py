from services.pdf_statement_adapter import extract_frames, summary_metrics
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


PDF_AREA = [250, 15, 800, 830]      # [top, left, bottom, right]
PDF_COLUMNS = [85, 450, 650, 760]   # split into: tanggal | keterangan+mutasi | saldo

GARBAGE_PATTERNS = [
    r"^TANGGAL$",
    r"^KETERANGAN$",
    r"^CBG$",
    r"^MUTASI$",
    r"^SALDO$",
    r"^TANGGAL\s+KETERANGAN\s+CBG\s+MUTASI\s+SALDO$",
    r"^TANGGAL\s+KETERANGAN\s+MUTASI\s+SALDO$",
    r"^CATATAN:?$",
    r"^\d+\s*/\s*\d+$",   # matches page markers like 1 / 4
    r"^BERSAMBUNG KE HALAMAN BERIKUT$",
]

SUMMARY_REGEX = r"^SALDO AWAL\s*:|^MUTASI CR\s*:|^MUTASI DB\s*:|^SALDO AKHIR\s*:"

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

SUMMARY_COLUMNS = [
    "No",
    "Bulan",
    "Mutasi Debet Nominal (Rp)",
    "Mutasi Debet Frek",
    "Mutasi Kredit Nominal (Rp)",
    "Mutasi Kredit Frek",
    "Saldo (Rp)",
    "Saldo Awal (Rp)",
    "Adm",
    "Pajak",
    "Bunga",
    "Saldo Min",
    "JaGir",
]

BCA_DIR = Path("pdf_file") / "BCA"
EXCEL_DIR = Path("excel_file")


def clean_text(value):
    """Normalize text: remove line breaks, trim spaces, collapse multiple spaces."""
    if pd.isna(value):
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ").strip()
    return re.sub(r"\s+", " ", text)


def parse_number(value):
    """
    Convert numeric text like '6,903,500.20' or '36,000.50 DB' into Decimal.
    Returns pd.NA if conversion fails.
    """
    text = str(value).strip().upper()

    if text in ("", "NAN"):
        return pd.NA

    text = text.replace("DB", "").replace("CR", "").replace(",", "").strip()

    try:
        return Decimal(text)
    except (InvalidOperation, ValueError):
        return pd.NA


def parse_frequency(value):
    """Convert frequency text into int without changing money parsing."""
    number = parse_number(value)
    if pd.isna(number):
        return pd.NA
    return int(number)


def parse_mutasi(value):
    """
    Split Mutasi into DB and CR.
    Example:
    - '36,000.00 DB' -> DB=36000, CR=<NA>
    - '6,903,500.00' -> DB=<NA>, CR=6903500
    """
    text = str(value).strip().upper()

    if text in ("", "NAN"):
        return pd.Series([pd.NA, pd.NA], index=["DB", "CR"])

    number = parse_number(text)
    if pd.isna(number):
        return pd.Series([pd.NA, pd.NA], index=["DB", "CR"])

    if "DB" in text:
        return pd.Series([number, pd.NA], index=["DB", "CR"])

    return pd.Series([pd.NA, number], index=["DB", "CR"])


def repair_mutasi_text(mutasi, keterangan):
    """
    Repair BCA rows where Tabula shifts the leading digit of a large amount
    into the description/CBG area.
    """
    mutasi_text = clean_text(mutasi)
    description = clean_text(keterangan)

    if not mutasi_text:
        return mutasi_text

    upper_mutasi = mutasi_text.upper()
    direction = " DB" if "DB" in upper_mutasi else ""
    amount_part = re.sub(r"\s+DB\b", "", mutasi_text, flags=re.IGNORECASE)

    if "TARIKAN TUNAI" in description.upper():
        suffix_match = re.search(r"\b0407\s+([1-9])\b", description)
        if suffix_match:
            prefix = suffix_match.group(1)
            if re.match(r"^\d{2},\d{3},\d{3}\.\d{2}(?:\s+DB)?$", upper_mutasi):
                return f"{prefix}{amount_part}{direction}"
            if re.match(r"^0{2},\d{3},\d{3}\.\d{2}(?:\s+DB)?$", upper_mutasi):
                return f"{prefix}{amount_part}{direction}"

    embedded_amounts = re.findall(r"\b\d{7,}(?:\.\d{2})\b", description)
    if embedded_amounts and re.match(r"^0[\d,]*\.\d{2}(?:\s+DB)?$", upper_mutasi):
        return f"{embedded_amounts[-1]}{direction}"

    if re.match(r"^0{2},\d{3},\d{3}\.\d{2}(?:\s+DB)?$", upper_mutasi):
        small_tokens = re.findall(r"\b\d{1,3}\b", description)
        if small_tokens:
            return f"{small_tokens[0]}{amount_part}{direction}"

    return mutasi_text


def is_garbage_row(row, garbage_patterns):
    """Detect repeated headers, page markers, and empty rows."""
    values = [str(v).strip() for v in row.tolist() if pd.notna(v) and str(v).strip()]
    joined = " ".join(values).strip()

    if not joined:
        return True

    return any(re.fullmatch(pattern, joined, flags=re.IGNORECASE) for pattern in garbage_patterns)


def append_text(base, extra):
    """Append extra text to base with a single space."""
    base = (base or "").strip()
    extra = (extra or "").strip()

    if not extra:
        return base
    if not base:
        return extra
    return f"{base} {extra}"


def read_pdf_table(pdf_file):
    """Read all pages from PDF using Tabula and combine them into one dataframe."""
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
    )

    if not dfs:
        raise ValueError("No table found. Try adjusting area/columns slightly.")

    print(f"Proses membaca PDF selesai, ditemukan {len(dfs)} potongan tabel. Menggabungkan data...")

    df = pd.concat(dfs, ignore_index=True)
    return normalize_tabula_rows(df)


def normalize_tabula_rows(df):
    """Normalize Tabula output into Tanggal/Keterangan/Mutasi/Saldo columns."""
    rows = []

    for _, row in df.iterrows():
        values = [clean_text(value) for value in row.tolist()]
        values = values + [""] * (4 - len(values))

        if len(row.tolist()) <= 3:
            tanggal, combined, right = values[:3]
            parsed = split_combined_transaction_text(tanggal, combined, right)
        else:
            parsed = normalize_legacy_tabula_row(values)

        rows.append(parsed)

    return pd.DataFrame(rows, columns=["Tanggal", "Keterangan", "Mutasi", "Saldo"])


def split_combined_transaction_text(tanggal, combined, right):
    """Split rows where Tabula keeps keterangan and mutasi in one column."""
    right = clean_text(right)
    combined = clean_text(combined)

    saldo = right
    mutasi = ""
    keterangan = combined

    if right.upper().startswith("DB "):
        saldo = right[3:].strip()

    amount_match = re.search(r"(\d[\d,]*\.\d{2})(?:\s+(DB))?$", combined, flags=re.IGNORECASE)
    if amount_match:
        amount = amount_match.group(1)
        marker = amount_match.group(2) or ("DB" if right.upper().startswith("DB ") else "")
        mutasi = f"{amount} {marker}".strip()
        keterangan = combined[:amount_match.start()].strip()

    return {
        "Tanggal": tanggal,
        "Keterangan": keterangan,
        "Mutasi": mutasi,
        "Saldo": saldo,
    }


def normalize_legacy_tabula_row(values):
    """Fallback for older 4-column Tabula output."""
    return {
        "Tanggal": values[0],
        "Keterangan": values[1],
        "Mutasi": values[2],
        "Saldo": values[3],
    }


def remove_garbage_rows(df):
    """Remove repeated headers, page markers, and other obvious non-transaction rows."""
    mask = df.apply(lambda row: is_garbage_row(row, GARBAGE_PATTERNS), axis=1)
    return df[~mask].reset_index(drop=True)


def split_summary_rows(df):
    """Separate transaction rows from summary rows at the bottom."""
    summary_mask = (
        df["Tanggal"].str.strip().eq("") &
        df["Keterangan"].str.upper().str.contains(SUMMARY_REGEX, regex=True, na=False)
    )

    print(f"Terdeteksi {summary_mask.sum()} baris summary. Memisahkan transaksi dan summary...")

    df_summary = df[summary_mask].reset_index(drop=True)
    df_trans = df[~summary_mask].reset_index(drop=True)

    return df_trans, df_summary


def merge_continuation_rows(df_trans):
    """
    Merge multiline transaction rows.
    If Tanggal is empty, the row is treated as continuation of previous row.
    """
    merged_rows = []
    current = None

    for _, row in df_trans.iterrows():
        tanggal = row["Tanggal"].strip()
        keterangan = row["Keterangan"].strip()
        mutasi = row["Mutasi"].strip()
        saldo = row["Saldo"].strip()

        if tanggal:
            if current is not None:
                merged_rows.append(current)

            current = {
                "Tanggal": tanggal,
                "Keterangan": keterangan,
                "Mutasi": mutasi,
                "Saldo": saldo,
            }
        else:
            if current is None:
                continue

            if keterangan:
                current["Keterangan"] = append_text(current["Keterangan"], keterangan)

            if mutasi:
                if current["Mutasi"] == "":
                    current["Mutasi"] = mutasi
                else:
                    current["Keterangan"] = append_text(current["Keterangan"], mutasi)

            if saldo:
                if current["Saldo"] == "":
                    current["Saldo"] = saldo
                else:
                    current["Keterangan"] = append_text(current["Keterangan"], saldo)

    if current is not None:
        merged_rows.append(current)

    return pd.DataFrame(merged_rows, columns=["Tanggal", "Keterangan", "Mutasi", "Saldo"])


def build_summary_dataframe(df_summary):
    """Convert summary rows into structured summary dataframe."""
    summary_clean = []

    for _, row in df_summary.iterrows():
        text = " ".join([row["Keterangan"], row["Mutasi"], row["Saldo"]]).strip()
        text = re.sub(r"\s+", " ", text)
        summary_clean.append(text)

    parsed_rows = []

    for row in summary_clean:
        parts = row.split(":", 1)

        if len(parts) == 2:
            label = parts[0].strip()
            values = parts[1].strip().split()

            amount = values[0] if len(values) > 0 else ""
            freq = values[1] if len(values) > 1 else ""

            parsed_rows.append({
                "Keterangan": label,
                "Amount": parse_number(amount),
                "Frekuensi": parse_frequency(freq),
            })

    return pd.DataFrame(parsed_rows, columns=["Keterangan", "Amount", "Frekuensi"])


def summary_value(df_summary_final, label, column="Amount"):
    """Return a value from the PDF summary table by label."""
    if df_summary_final.empty:
        return pd.NA

    mask = df_summary_final["Keterangan"].str.upper().eq(label.upper())
    if not mask.any():
        return pd.NA

    return df_summary_final.loc[mask, column].iloc[0]


def finalize_transactions(clean_df):
    """Create final transaction dataframe with DB, CR, and cleaned saldo."""
    clean_df["Mutasi"] = clean_df.apply(
        lambda row: repair_mutasi_text(row["Mutasi"], row["Keterangan"]),
        axis=1,
    )
    clean_df[["DB", "CR"]] = clean_df["Mutasi"].apply(parse_mutasi)
    clean_df["Saldo"] = clean_df["Saldo"].apply(parse_number)

    clean_df = clean_df.drop(columns=["Mutasi"])
    clean_df = clean_df[["Tanggal", "Keterangan", "DB", "CR", "Saldo"]]
    clean_df = clean_df[~clean_df.apply(is_saldo_awal_transaction, axis=1)]
    clean_df = clean_df.reset_index(drop=True)

    return clean_df


def is_saldo_awal_transaction(row):
    """Detect SALDO AWAL rows extracted from the PDF body."""
    description = clean_text(row.get("Keterangan", "")).upper()
    db_value = row.get("DB", pd.NA)
    cr_value = row.get("CR", pd.NA)
    return "SALDO AWAL" in description and pd.isna(db_value) and pd.isna(cr_value)


@lru_cache(maxsize=None)
def read_pdf_text_for_metadata(pdf_file):
    """Read PDF text for metadata extraction without relying on the filename."""
    pdf_file = Path(pdf_file)
    text_parts = []

    try:
        with fitz.open(pdf_file) as doc:
            for page in doc:
                text_parts.append(page.get_text("text"))
    except Exception:
        return ""

    return "\n".join(text_parts)


def normalize_metadata_text(text):
    """Make PDF header text easier to search across line breaks."""
    return re.sub(r"\s+", " ", clean_text(text)).upper()


def extract_account_from_text(text):
    """Extract account number from a NO. REKENING label in the source PDF."""
    normalized = normalize_metadata_text(text)
    match = re.search(r"NO\.?\s*REKENING\s*:?\s*(\d{6,})", normalized)
    if match:
        return match.group(1)

    lines = [clean_text(line) for line in text.splitlines()]
    for idx, line in enumerate(lines):
        if re.fullmatch(r"NO\.?\s*REKENING", line, flags=re.IGNORECASE):
            for next_line in lines[idx + 1:idx + 5]:
                number_match = re.search(r"\d{6,}", next_line)
                if number_match:
                    return number_match.group(0)

    return None


def extract_period_from_text(text):
    """Extract month/year from labels like PERIODE : JANUARI 2026."""
    normalized = normalize_metadata_text(text)
    month_pattern = "|".join(sorted(MONTH_ALIASES, key=len, reverse=True))
    match = re.search(rf"PERIODE\s*:?\s*({month_pattern})\s+(20\d{{2}})", normalized)
    if match:
        month_key, year = match.groups()
        month_name, month_order = MONTH_ALIASES[month_key]
        return f"{month_name}-{year[-2:]}", month_order, year

    generic_match = re.search(rf"\b({month_pattern})\s+(20\d{{2}})\b", normalized)
    if generic_match:
        month_key, year = generic_match.groups()
        month_name, month_order = MONTH_ALIASES[month_key]
        return f"{month_name}-{year[-2:]}", month_order, year

    return None


def extract_period_from_filename(pdf_file):
    """Fallback month/year extraction from <account>_<month>_<year>.pdf names."""
    stem = pdf_file.stem.upper()
    year_match = re.search(r"(20\d{2})", stem)
    year = year_match.group(1) if year_match else None

    month_label = None
    month_order = 99
    for token in re.split(r"[^A-Z]+", stem):
        if token in MONTH_ALIASES:
            month_name, month_order = MONTH_ALIASES[token]
            if year:
                month_label = f"{month_name}-{year[-2:]}"
            else:
                month_label = month_name
            break

    if year is None and re.fullmatch(r"20\d{2}", pdf_file.parent.name):
        year = pdf_file.parent.name
        if month_label and "-" not in month_label:
            month_label = f"{month_label}-{year[-2:]}"

    if month_label is None:
        month_label = pdf_file.stem[:20]

    return month_label, month_order, year or "Unknown"


def extract_pdf_metadata(pdf_file):
    """Extract account/month/year from PDF content, with filename as fallback."""
    pdf_file = Path(pdf_file)
    text = read_pdf_text_for_metadata(pdf_file)
    period = extract_period_from_text(text) or extract_period_from_filename(pdf_file)

    account = extract_account_from_text(text)
    if account is None:
        account_match = re.match(r"(\d+)", pdf_file.stem)
        account = account_match.group(1) if account_match else pdf_file.stem.split("_")[0]

    month_label, month_order, year = period
    return {
        "account": account,
        "month_label": month_label,
        "month_order": month_order,
        "year": str(year),
    }


def extract_month_year(pdf_file):
    """Return month label, month order, and year from PDF content."""
    metadata = extract_pdf_metadata(pdf_file)
    return metadata["month_label"], metadata["month_order"], metadata["year"]


def extract_account_number(pdf_file):
    """Return account number from PDF content."""
    return extract_pdf_metadata(pdf_file)["account"]


def safe_sheet_name(name):
    """Create an Excel-safe worksheet name."""
    cleaned = re.sub(r"[\[\]\:\*\?\/\\]", "_", name).strip()
    return cleaned[:31] or "Sheet"


def unique_sheet_name(base_name, used_names):
    """Avoid duplicate sheet names inside a workbook."""
    base_name = safe_sheet_name(base_name)
    candidate = base_name
    counter = 2

    while candidate.lower() in used_names:
        suffix = f"_{counter}"
        candidate = f"{base_name[:31 - len(suffix)]}{suffix}"
        counter += 1

    used_names.add(candidate.lower())
    return candidate


def sum_by_description(clean_df, amount_column, pattern):
    """Sum DB/CR values for rows whose description matches a keyword pattern."""
    if clean_df.empty:
        return pd.NA

    mask = clean_df["Keterangan"].str.upper().str.contains(pattern, regex=True, na=False)
    values = clean_df.loc[mask, amount_column].dropna()

    if values.empty:
        return pd.NA
    return sum(values, Decimal("0.00"))


def build_month_summary(clean_df, df_summary_final, pdf_file, metadata=None):
    """Build one row for the yearly recap sheet."""
    metadata = metadata or extract_pdf_metadata(pdf_file)
    month_label = metadata["month_label"]
    month_order = metadata["month_order"]
    year = metadata["year"]

    row = {
        "Source File": pdf_file.name,
        "Account": metadata["account"],
        "Year": year,
        "MonthOrder": month_order,
        "Bulan": month_label,
        "Mutasi Debet Nominal (Rp)": summary_value(df_summary_final, "MUTASI DB", "Amount"),
        "Mutasi Debet Frek": summary_value(df_summary_final, "MUTASI DB", "Frekuensi"),
        "Mutasi Kredit Nominal (Rp)": summary_value(df_summary_final, "MUTASI CR", "Amount"),
        "Mutasi Kredit Frek": summary_value(df_summary_final, "MUTASI CR", "Frekuensi"),
        "Saldo (Rp)": summary_value(df_summary_final, "SALDO AKHIR", "Amount"),
        "Saldo Awal (Rp)": summary_value(df_summary_final, "SALDO AWAL", "Amount"),
        **summary_metrics(clean_df),
    }

    return row


def process_pdf(pdf_file):
    """Parse through the validated shared PDF entry point."""
    return extract_frames(pdf_file, "BCA")


def reconcile_transactions_with_balance(clean_df, df_summary_final):
    """Use saldo movements to repair amount/direction when Tabula split text badly."""
    if clean_df.empty:
        return clean_df

    result = clean_df.copy()
    last_balance = summary_value(df_summary_final, "SALDO AWAL", "Amount")
    pending_without_balance = 0

    for idx, row in result.iterrows():
        current_balance = row["Saldo"]

        if pd.isna(current_balance):
            pending_without_balance += 1
            continue

        if not pd.isna(last_balance) and pending_without_balance == 0:
            delta = current_balance - last_balance
            if delta < 0:
                result.at[idx, "DB"] = abs(delta)
                result.at[idx, "CR"] = pd.NA
            elif delta > 0:
                result.at[idx, "DB"] = pd.NA
                result.at[idx, "CR"] = delta

        last_balance = current_balance
        pending_without_balance = 0

    return result


def report_extraction_balance(clean_df, df_summary_final, pdf_file):
    """Print a warning if extracted rows do not match the PDF recap."""
    extracted_db = sum(clean_df["DB"].dropna(), Decimal("0.00"))
    extracted_cr = sum(clean_df["CR"].dropna(), Decimal("0.00"))
    summary_db = summary_value(df_summary_final, "MUTASI DB", "Amount")
    summary_cr = summary_value(df_summary_final, "MUTASI CR", "Amount")

    mismatches = []
    if not pd.isna(summary_db) and extracted_db != summary_db:
        mismatches.append(f"DB extracted {extracted_db} but PDF summary {summary_db}")
    if not pd.isna(summary_cr) and extracted_cr != summary_cr:
        mismatches.append(f"CR extracted {extracted_cr} but PDF summary {summary_cr}")

    if mismatches:
        print(f"PERINGATAN: {Path(pdf_file).name}: " + "; ".join(mismatches))


def first_day_label(pdf_file, clean_df):
    """Return a DD/MM date label for the added opening-balance row."""
    _, month_order, _ = extract_month_year(pdf_file)
    if month_order != 99:
        return f"01/{month_order:02d}"

    if not clean_df.empty:
        first_date = clean_text(clean_df["Tanggal"].iloc[0])
        match = re.match(r"^\d{1,2}/\d{1,2}", first_date)
        if match:
            return match.group(0)

    return ""


def build_monthly_display_dataframe(clean_df, df_summary_final, pdf_file):
    """Add the opening balance row used by the monthly sheet style."""
    saldo_awal = summary_value(df_summary_final, "SALDO AWAL", "Amount")
    opening_row = pd.DataFrame([{
        "Tanggal": first_day_label(pdf_file, clean_df),
        "Keterangan": "SALDO AWAL",
        "DB": pd.NA,
        "CR": pd.NA,
        "Saldo": saldo_awal,
    }])

    return pd.concat([opening_row, clean_df], ignore_index=True)


def auto_fit_columns(sheet):
    """Auto-fit Excel column widths based on content length."""
    for col_idx, col_cells in enumerate(sheet.columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_length = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in col_cells
        )
        sheet.column_dimensions[col_letter].width = max_length + 2


def write_monthly_transaction_sheet(writer, clean_df, df_summary_final, pdf_file, sheet_name):
    """Write one monthly transaction sheet and apply the sample style."""
    display_df = build_monthly_display_dataframe(clean_df, df_summary_final, pdf_file)
    display_df = dataframe_for_excel(display_df)
    display_df.to_excel(writer, sheet_name=sheet_name, index=False)
    sheet = writer.sheets[sheet_name]
    add_monthly_sheet_totals(sheet, df_summary_final)
    style_monthly_sheet(sheet)


def excel_number(value):
    """Convert Decimal/pandas values into Excel-friendly numbers."""
    if pd.isna(value):
        return None
    if isinstance(value, Decimal):
        return float(value)
    return value


def dataframe_for_excel(df):
    """Convert Decimal values to numeric Excel values while preserving blanks."""
    result = df.copy()
    for col in result.columns:
        result[col] = result[col].apply(excel_number)
    return result


def add_monthly_sheet_totals(sheet, df_summary_final):
    """Append totals calculated from visible monthly transaction rows."""
    first_data_row = 2
    last_data_row = sheet.max_row
    total_row = last_data_row + 3
    freq_row = total_row + 1
    average_row = total_row + 2

    labels = {
        total_row: "Total",
        freq_row: "Frekuensi",
        average_row: "Rata2",
    }

    for row_num, label in labels.items():
        sheet.merge_cells(start_row=row_num, start_column=2, end_row=row_num, end_column=2)
        sheet.cell(row=row_num, column=2, value=label)

    for col_num in (3, 4):
        col_letter = get_column_letter(col_num)
        extracted_total_cell = f"{col_letter}{total_row}"

        sheet.cell(row=total_row, column=col_num, value=f"=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})")
        sheet.cell(row=freq_row, column=col_num, value=f"=COUNT({col_letter}{first_data_row}:{col_letter}{last_data_row})")
        sheet.cell(row=average_row, column=col_num, value=f"=IFERROR({extracted_total_cell}/{col_letter}{freq_row},0)")


def style_monthly_sheet(sheet):
    """Apply the monthly sheet style from the reference image."""
    header_fill = PatternFill("solid", fgColor="BFBFBF")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    summary_start = sheet.max_row - 2

    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False

    widths = {
        "A": 10,
        "B": 95,
        "C": 16,
        "D": 16,
        "E": 16,
    }
    for col_letter, width in widths.items():
        sheet.column_dimensions[col_letter].width = width

    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=5):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

            if cell.row == 1:
                cell.fill = header_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif cell.column in (3, 4, 5):
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif cell.column == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for row_num in range(summary_start, sheet.max_row + 1):
        label_cell = sheet.cell(row=row_num, column=2)
        label_cell.fill = header_fill
        label_cell.font = Font(bold=True)
        label_cell.alignment = Alignment(horizontal="center", vertical="center")

        for col_num in (3, 4):
            value_cell = sheet.cell(row=row_num, column=col_num)
            if row_num == summary_start + 1:
                value_cell.number_format = "0"
                value_cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                value_cell.number_format = "#,##0.00"
                value_cell.alignment = Alignment(horizontal="right", vertical="center")

    for row_num in range(1, sheet.max_row + 1):
        sheet.row_dimensions[row_num].height = 18


def write_summary_sheet(writer, summary_rows):
    """Write the yearly recap sheet with the table structure from the sample image."""
    workbook = writer.book
    sheet = workbook.create_sheet("Summary", 0)
    writer.sheets["Summary"] = sheet

    headers_top = [
        "No",
        "Bulan",
        "Mutasi Debet",
        None,
        "Mutasi Kredit",
        None,
        "Saldo",
        "Saldo Awal",
        "Adm",
        "Pajak",
        "Bunga",
        "Saldo Min",
        "JaGir",
    ]
    headers_bottom = [
        "No",
        "Bulan",
        "Nominal (Rp)",
        "Frek",
        "Nominal (Rp)",
        "Frek",
        "(Rp)",
        "(Rp)",
        "Adm",
        "Pajak",
        "Bunga",
        "Saldo Min",
        "JaGir",
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

    sum_columns = [3, 4, 5, 6, 9, 10, 11, 12, 13]
    avg_columns = [3, 4, 5, 6, 7]
    data_start = 3
    data_end = total_row - 1

    if data_end >= data_start:
        for col_num in sum_columns:
            col_letter = get_column_letter(col_num)
            sheet.cell(row=total_row, column=col_num, value=f"=SUM({col_letter}{data_start}:{col_letter}{data_end})")
        for col_num in avg_columns:
            col_letter = get_column_letter(col_num)
            sheet.cell(row=average_row, column=col_num, value=f"=AVERAGE({col_letter}{data_start}:{col_letter}{data_end})")

    style_summary_sheet(sheet)
    auto_fit_columns(sheet)


def style_summary_sheet(sheet):
    """Apply simple Excel styling to the yearly recap."""
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
                cell.number_format = '#,##0.00'
            elif cell.row >= 3 and cell.column in (4, 6):
                cell.number_format = '0'

    for row_num in (sheet.max_row - 1, sheet.max_row):
        for col_num in range(1, sheet.max_column + 1):
            sheet.cell(row=row_num, column=col_num).font = Font(bold=True)


def export_to_excel(df_final, df_summary, output_file, pdf_file=None):
    output_file = Path(output_file)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        if pdf_file is None:
            dataframe_for_excel(df_final).to_excel(writer, sheet_name="Transaksi", index=False)
            dataframe_for_excel(df_summary).to_excel(writer, sheet_name="Summary", index=False)
        else:
            pdf_file = Path(pdf_file)
            write_summary_sheet(writer, [build_month_summary(df_final, df_summary, pdf_file)])
            write_monthly_transaction_sheet(writer, df_final, df_summary, pdf_file, "Transaksi")


def export_year_workbook(extracted_files, output_file):
    """Write one workbook for one folder/year: Summary + one sheet per PDF."""
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
    """Return PDF groups keyed by folder. A file input stays as single-file mode."""
    input_path = Path(input_path)

    if input_path.is_file():
        return {input_path.parent: [input_path]}

    pdf_files = sorted(input_path.rglob("*.pdf"))
    groups = {}

    for pdf_file in pdf_files:
        groups.setdefault(pdf_file.parent, []).append(pdf_file)

    return groups


def ensure_xlsx_name(name):
    """Normalize a user-provided workbook name."""
    name = clean_text(name) or "{account}_{year}.xlsx"
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name


def output_name_for_group(group_info, output_name_template=None, force_unique=False):
    """Create yearly workbook names using the account and statement year."""
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
        raise ValueError(
            "Output name template only supports {account}, {year}, and {folder} placeholders."
        ) from exc

    if force_unique and not has_year_or_folder:
        output_path = Path(output_name)
        output_name = f"{output_path.stem}_{year}{output_path.suffix}"

    return output_name


def convert_pdf(pdf_file, output_file):
    pdf_file, output_file = Path(pdf_file), Path(output_file)
    transactions, summary = process_pdf(pdf_file)
    export_to_excel(transactions, summary, output_file, pdf_file)
    return output_file


def run_single_file(pdf_file, output_file):
    clean_df, df_summary_final = process_pdf(pdf_file)

    print("Proses pembersihan data selesai. Menulis ke Excel...")
    output_file.parent.mkdir(parents=True, exist_ok=True)
    export_to_excel(clean_df, df_summary_final, output_file, pdf_file)
    print(f"Data successfully written to {output_file}")


def discover_pdf_groups_by_content(input_path):
    """Group PDFs by account/year discovered from PDF content."""
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
            clean_df, df_summary_final = process_pdf(pdf_file)
            month_label = metadata["month_label"]
            month_order = metadata["month_order"]
            year = metadata["year"]
            sheet_hint = f"{month_label.replace('-', '_')}"

            extracted_files.append({
                "pdf_file": pdf_file,
                "transactions": clean_df,
                "summary": df_summary_final,
                "summary_row": build_month_summary(clean_df, df_summary_final, pdf_file, metadata),
                "sheet_name": sheet_hint,
                "month_order": month_order,
                "year": year,
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
    """Ask for input folder, output folder, and workbook name pattern."""
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, simpledialog
    except ImportError:
        return None, None, None

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    messagebox.showinfo(
        "BCA Mutasi PDF to Excel",
        "Choose the BCA folder that contains your monthly PDF files. Subfolders will be included.",
        parent=root,
    )
    input_dir = filedialog.askdirectory(
        title="Choose BCA PDF folder",
        initialdir=str(BCA_DIR.resolve()) if BCA_DIR.exists() else str(Path.cwd()),
        parent=root,
    )
    if not input_dir:
        root.destroy()
        return None, None, None

    messagebox.showinfo(
        "BCA Mutasi PDF to Excel",
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
        "BCA Mutasi PDF to Excel",
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
    """Fallback prompts when Tkinter is unavailable."""
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
        messagebox.showerror("BCA Mutasi PDF to Excel", message, parent=root)
    else:
        messagebox.showinfo("BCA Mutasi PDF to Excel", message, parent=root)

    root.destroy()


def parse_args():
    parser = argparse.ArgumentParser(description="Convert BCA mutasi PDF files to Excel.")
    parser.add_argument(
        "input",
        nargs="?",
        help="PDF filename without extension, PDF file path, or folder path. No arguments opens folder dialogs.",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=str(EXCEL_DIR),
        help="Output folder for generated Excel files. Default: excel_file",
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
    pdf_dir = Path("pdf_file")
    input_path = Path(input_value)

    if input_path.exists():
        return input_path

    candidate_pdf = pdf_dir / f"{input_path}.pdf"
    candidate_folder = pdf_dir / input_path
    if candidate_pdf.exists():
        return candidate_pdf
    if candidate_folder.exists():
        return candidate_folder

    raise FileNotFoundError(f"Input not found: {input_path}")


def main():
    args = parse_args()
    use_gui = len(sys.argv) == 1

    if use_gui:
        selected = choose_batch_options_with_gui()
        if selected == (None, None, None):
            selected = choose_batch_options_with_console()
        input_path, excel_dir, output_name_template = selected
    else:
        input_path = resolve_input_path(args.input)
        excel_dir = Path(args.output_dir)
        output_name_template = args.output_name

    excel_dir.mkdir(parents=True, exist_ok=True)

    if not input_path.exists():
        raise FileNotFoundError(f"Input not found: {input_path}")

    if input_path.is_file():
        output_name = ensure_xlsx_name(output_name_template or f"{input_path.stem}.xlsx")
        try:
            output_name = output_name.format(
                account=extract_account_number(input_path),
                year=extract_month_year(input_path)[2],
                folder=input_path.parent.name,
            )
        except (KeyError, ValueError) as exc:
            raise ValueError(
                "Output name template only supports {account}, {year}, and {folder} placeholders."
            ) from exc
        output_file = excel_dir / output_name
        run_single_file(input_path, output_file)
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
            show_gui_result(
                "Done.\n\nSaved files:\n" + "\n".join(str(path) for path in written_files)
            )

if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Error: {exc}")
        if len(sys.argv) == 1:
            show_gui_result(str(exc), is_error=True)
            sys.exit(1)
        raise
