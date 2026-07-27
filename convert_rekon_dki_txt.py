import re
from pathlib import Path

import pandas as pd


# =========================================================
# HELPER
# =========================================================
def clean_text(value):
    if pd.isna(value):
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ").strip()
    return re.sub(r"\s+", " ", text)


def parse_number(value):
    """
    Convert:
    - '25,000.00 DB' -> 25000
    - '150,000.00'   -> 150000
    - ''             -> pd.NA
    """
    text = str(value).strip().upper()

    if text in ("", "NAN", "NONE"):
        return pd.NA

    text = text.replace("DB", "").replace(",", "").strip()

    try:
        number = float(text)
        if number.is_integer():
            return int(number)
        return number
    except ValueError:
        return pd.NA


def parse_mutasi(value):
    """
    Pisahkan mutasi menjadi DB / CR.
    Jika ada DB -> debit
    Jika tidak ada DB -> kredit
    """
    text = str(value).strip().upper()

    if text in ("", "NAN", "NONE"):
        return pd.Series([pd.NA, pd.NA], index=["DB", "CR"])

    number = parse_number(text)
    if pd.isna(number):
        return pd.Series([pd.NA, pd.NA], index=["DB", "CR"])

    if "DB" in text:
        return pd.Series([number, pd.NA], index=["DB", "CR"])

    return pd.Series([pd.NA, number], index=["DB", "CR"])


def append_text(base, extra):
    base = clean_text(base)
    extra = clean_text(extra)

    if not extra:
        return base
    if not base:
        return extra
    return f"{base} {extra}"


def auto_fit_columns(sheet):
    for col_cells in sheet.columns:
        col_letter = col_cells[0].column_letter
        max_length = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in col_cells
        )
        sheet.column_dimensions[col_letter].width = max_length + 2


# =========================================================
# PARSER
# =========================================================
DATE_PATTERN = r"\d{1,2}/\d{2}/\d{2}"
AMOUNT_PATTERN = r"[\d,]+\.\d{2}(?:\s+DB)?"

# pola baris transaksi utama:
# tanggal | keterangan | kolom3 | mutasi | saldo berjalan
TX_PATTERN = re.compile(
    rf"^\s*(?P<Tanggal>{DATE_PATTERN})\s+"
    rf"(?P<body>.*?)\s+"
    rf"(?P<Mutasi>{AMOUNT_PATTERN})"
    rf"(?:\s+(?P<Saldo_Berjalan>{AMOUNT_PATTERN}))?\s*$"
)

ACCOUNT_PATTERN = re.compile(r"No\.\s*Rekening\s*:\s*(\d+)", flags=re.IGNORECASE)
SALDO_AWAL_PATTERN = re.compile(r"SALDO AWAL", flags=re.IGNORECASE)
SALDO_AKHIR_PATTERN = re.compile(r"SALDO AKHIR", flags=re.IGNORECASE)
PINDAHAN_PATTERN = re.compile(r"PINDAHAN", flags=re.IGNORECASE)


def split_body_keterangan_kol3(body):
    """
    Pisahkan body menjadi Keterangan dan Kolom3.
    Asumsi kolom3 biasanya token terakhir tanpa spasi, mis:
    DC16, JAKOM, USERSP2D, NEWCMS, dst.

    Jika body tidak bisa dipisah dengan aman, semua masuk ke Keterangan.
    """
    body = clean_text(body)

    if not body:
        return "", ""

    parts = body.rsplit(None, 1)
    if len(parts) == 2:
        left, right = parts[0], parts[1]

        # kolom3 umumnya kode tanpa spasi
        if re.fullmatch(r"[A-Z0-9/\-]+", right, flags=re.IGNORECASE):
            return left, right

    return body, ""


def is_metadata_line(line):
    text = clean_text(line).upper()

    if not text:
        return True

    metadata_keywords = [
        "NO. REKENING",
        "NAMA NASABAH",
        "KUASA",
        "ALAMAT",
        "PLAFOND",
        "PERIODE TGL.",
        "HALAMAN KE",
        "BUNGA",
        "BUNGA OD",
        "DENDA",
        "TANGGAL BUKA",
        "TANGGAL AKHIR",
    ]

    if any(keyword in text for keyword in metadata_keywords):
        return True

    if set(text) == {"*"}:
        return True

    return False


def is_summary_like_line(line):
    text = clean_text(line).upper()
    return (
        bool(SALDO_AWAL_PATTERN.search(text))
        or bool(SALDO_AKHIR_PATTERN.search(text))
        or bool(PINDAHAN_PATTERN.search(text))
    )


def parse_txt_file(txt_file):
    transactions = []
    summaries = []

    current_account = None
    current_tx = None
    current_saldo_awal = pd.NA
    current_saldo_akhir = pd.NA

    with open(txt_file, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()

    for raw_line in lines:
        line = raw_line.rstrip("\n")
        text = clean_text(line)

        # skip kosong
        if not text:
            continue

        # detect rekening baru
        m_acc = ACCOUNT_PATTERN.search(line)
        if m_acc:
            # simpan transaksi sebelumnya jika masih ada
            if current_tx is not None:
                transactions.append(current_tx)
                current_tx = None

            # jika pindah rekening, simpan summary rekening sebelumnya
            if current_account is not None:
                summaries.append({
                    "No_Rekening": current_account,
                    "Saldo_Awal": current_saldo_awal,
                    "Saldo_Akhir": current_saldo_akhir,
                })

            current_account = m_acc.group(1)
            current_saldo_awal = pd.NA
            current_saldo_akhir = pd.NA
            continue

        # skip metadata/header
        if is_metadata_line(line):
            continue

        # saldo awal
        if SALDO_AWAL_PATTERN.search(line):
            amounts = re.findall(r"[\d,]+\.\d{2}", line)
            if amounts:
                current_saldo_awal = parse_number(amounts[-1])
            continue

        # saldo akhir
        if SALDO_AKHIR_PATTERN.search(line):
            amounts = re.findall(r"[\d,]+\.\d{2}(?:\s+DB)?", line)
            if amounts:
                current_saldo_akhir = parse_saldo(amounts[-1])
            continue

        # baris pindahan bukan transaksi
        if PINDAHAN_PATTERN.search(line):
            continue

        # coba parse transaksi utama
        m_tx = TX_PATTERN.match(line)
        if m_tx:
            if current_tx is not None:
                transactions.append(current_tx)

            tanggal = clean_text(m_tx.group("Tanggal"))
            body = clean_text(m_tx.group("body"))
            mutasi = clean_text(m_tx.group("Mutasi"))
            saldo_berjalan = clean_text(m_tx.group("Saldo_Berjalan") or "")

            keterangan, kolom3 = split_body_keterangan_kol3(body)

            current_tx = {
                "No_Rekening": current_account,
                "Tanggal": tanggal,
                "Keterangan": keterangan,
                "Kolom3": kolom3,
                "Mutasi": mutasi,
                "Saldo_Berjalan": saldo_berjalan,
            }
            continue

        # continuation row: gabungkan ke keterangan transaksi sebelumnya
        if current_tx is not None:
            upper_text = text.upper()

            # hindari metadata/summary yang nyasar
            if not is_summary_like_line(text):
                current_tx["Keterangan"] = append_text(current_tx["Keterangan"], text)

    # append terakhir
    if current_tx is not None:
        transactions.append(current_tx)

    if current_account is not None:
        summaries.append({
            "No_Rekening": current_account,
            "Saldo_Awal": current_saldo_awal,
            "Saldo_Akhir": current_saldo_akhir,
        })

    df_tx = pd.DataFrame(
        transactions,
        columns=["No_Rekening", "Tanggal", "Keterangan", "Kolom3", "Mutasi", "Saldo_Berjalan"]
    )

    df_sum = pd.DataFrame(
        summaries,
        columns=["No_Rekening", "Saldo_Awal", "Saldo_Akhir"]
    )

    return df_tx, df_sum

def parse_saldo(value):
    text = str(value).strip().upper()

    if text in ("", "NAN"):
        return pd.NA

    is_debit = "DB" in text

    text = text.replace("DB", "").replace(",", "").strip()

    try:
        number = float(text)
        if number.is_integer():
            number = int(number)

        return -number if is_debit else number
    except ValueError:
        return pd.NA

# =========================================================
# FINALIZE
# =========================================================
def finalize_transactions(df_tx):
    df_tx = df_tx.copy()

    df_tx[["DB", "CR"]] = df_tx["Mutasi"].apply(parse_mutasi)
    df_tx["Saldo_Berjalan"] = df_tx["Saldo_Berjalan"].apply(parse_saldo)

    df_tx = df_tx.drop(columns=["Mutasi"])
    df_tx = df_tx[
        ["No_Rekening", "Tanggal", "Keterangan", "Kolom3", "DB", "CR", "Saldo_Berjalan"]
    ]

    return df_tx


def build_summary(df_tx, df_sum):
    """
    Bangun summary per rekening:
    - No_Rekening
    - Saldo_Awal
    - Total_DB
    - Total_CR
    - Saldo_Akhir
    """
    df_group = (
        df_tx.groupby("No_Rekening", dropna=False)
        .agg(
            Total_DB=("DB", lambda s: s.dropna().sum() if not s.dropna().empty else pd.NA),
            Total_CR=("CR", lambda s: s.dropna().sum() if not s.dropna().empty else pd.NA),
        )
        .reset_index()
    )

    df_summary = df_sum.merge(df_group, on="No_Rekening", how="left")

    df_summary = df_summary[
        ["No_Rekening", "Saldo_Awal", "Total_DB", "Total_CR", "Saldo_Akhir"]
    ]

    return df_summary


# =========================================================
# EXPORT
# =========================================================
def export_to_excel(df_tx, df_summary, output_file):
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_tx.to_excel(writer, sheet_name="Transaksi", index=False)
        df_summary.to_excel(writer, sheet_name="Summary", index=False)

        auto_fit_columns(writer.sheets["Transaksi"])
        auto_fit_columns(writer.sheets["Summary"])


# =========================================================
# MAIN
# =========================================================
def main():
    file_name = input("Enter the TXT filename without extension: ").strip()

    txt_dir = Path("txt_file")
    excel_dir = Path("excel_file")
    excel_dir.mkdir(parents=True, exist_ok=True)

    txt_file = txt_dir / f"{file_name}.txt"
    output_file = excel_dir / f"{file_name}.xlsx"

    if not txt_file.exists():
        raise FileNotFoundError(f"TXT file not found: {txt_file}")

    print("Membaca dan mem-parse file TXT...")
    df_tx_raw, df_sum_raw = parse_txt_file(txt_file)

    print("Membersihkan dan memfinalisasi transaksi...")
    df_tx = finalize_transactions(df_tx_raw)
    df_tx["Saldo_Berjalan"] = df_tx["Saldo_Berjalan"].apply(parse_saldo)

    print("Menyusun summary...")
    df_summary = build_summary(df_tx, df_sum_raw)

    print("Menulis hasil ke Excel...")
    export_to_excel(df_tx, df_summary, output_file)

    print(f"Data successfully written to {output_file}")


if __name__ == "__main__":
    main()