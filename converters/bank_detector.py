import re
from pathlib import Path

import fitz


def _normalize_text(text):
    return re.sub(r"\s+", " ", text.upper()).strip()


def _score_markers(text, markers):
    score = 0
    for marker, weight in markers:
        if isinstance(marker, str):
            if marker in text:
                score += weight
        elif marker.search(text):
            score += weight
    return score


def detect_bank_from_pdf(pdf_file):
    """
    Return "BCA", "BNI", "DKI", "Mandiri", "BRI", or None based on document-level text markers.

    Bank names can appear inside transaction descriptions, so this intentionally
    scores statement/header markers instead of checking a raw "BCA"/"BNI" token.
    """
    pdf_file = Path(pdf_file)
    text_parts = []

    with fitz.open(pdf_file) as document:
        for page_index in range(min(2, len(document))):
            text_parts.append(document[page_index].get_text("text"))

    text = _normalize_text("\n".join(text_parts))
    filename = _normalize_text(pdf_file.name)

    scores = {
        "BCA": _score_markers(
            text,
            [
                ("BANK CENTRAL ASIA", 10),
                ("REKENING TAHAPAN", 8),
                ("BCA BERHAK", 7),
                ("TANGGAL KETERANGAN CBG MUTASI SALDO", 5),
                ("CBG MUTASI SALDO", 4),
                ("NO. REKENING : HALAMAN : PERIODE", 3),
            ],
        ),
        "BNI": _score_markers(
            text,
            [
                ("PT BANK NEGARA INDONESIA", 12),
                ("BANK NEGARA INDONESIA", 10),
                ("TRANSACTION INQUIRY", 8),
                ("ACCOUNT STATEMENT", 6),
                ("ACCOUNT INFORMATION", 5),
                ("POSTING DATE", 4),
                ("EFFECTIVE DATE", 4),
                ("TRANSACTION DESCRIPTION", 4),
                ("POST DATE BRANCH JOURNAL NO. DESCRIPTION AMOUNT DB/CR BALANCE", 7),
                ("TOTAL DEBIT", 4),
                ("TOTAL CREDIT", 4),
                ("TAPLUS", 6),
                ("TAPLUS MUDA", 6),
                ("TANGGAL & WAKTU", 4),
                ("RINCIAN TRANSAKSI", 4),
            ],
        ),
        "DKI": _score_markers(
            text,
            [
                ("BANK JAKARTA", 10),
                ("BANK DKI", 10),
                ("TABUNGAN MONAS", 8),
                ("TAB MONAS", 7),
                (re.compile(r"\bDKI\s*-\s*\d{6,}\b"), 8),
                ("E-STATEMENT RINGKASAN", 5),
                ("TRANSAKSI MASUK TRANSAKSI KELUAR", 4),
                ("JAKONE", 4),
            ],
        ),
        "Mandiri": _score_markers(
            text,
            [
                ("KOPRA BY MANDIRI", 12),
                ("BANK MANDIRI", 10),
                ("BMRIIDJA", 8),
                ("ACCOUNT STATEMENT SUMMARY", 5),
                ("ACCOUNT NO. ACCOUNT NAME ALIAS PERIOD CURRENCY BRANCH", 5),
                ("TOTAL AMOUNT DEBITED", 4),
                ("TOTAL AMOUNT CREDITED", 4),
            ],
        ),
        "BRI": _score_markers(
            text,
            [
                ("BANK RAKYAT INDONESIA", 12),
                ("BRINIDJA", 8),
                ("GIRO UMUM-IDR", 7),
                ("UNIT KERJA BUSINESS UNIT", 6),
                ("NO. REKENING ACCOUNT NO", 5),
                ("TANGGAL TRANSAKSI TRANSACTION DATE", 5),
                ("URAIAN TRANSAKSI TRANSACTION DESCRIPTION", 5),
            ],
        ),
    }

    if "BCA" in filename:
        scores["BCA"] += 2
    if "BNI" in filename:
        scores["BNI"] += 2
    if "JAKOM" in filename or "BANK JAKARTA" in filename:
        scores["DKI"] += 2
    if "MANDIRI" in filename:
        scores["Mandiri"] += 3
    if re.search(r"\bBRI\b", filename):
        scores["BRI"] += 4

    best_bank, best_score = max(scores.items(), key=lambda item: item[1])
    if best_score < 4:
        return None

    tied = [bank for bank, score in scores.items() if score == best_score]
    if len(tied) > 1:
        return None

    return best_bank
