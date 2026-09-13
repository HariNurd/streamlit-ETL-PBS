"""Extract supported bank statement labels locally; validate before writing.

Run from the project root with: env/Scripts/python.exe generate_labels.py
Requires pypdf and cryptography. Existing labels are never overwritten.
"""

import json
import re
from decimal import Decimal
from pathlib import Path

from pypdf import PdfReader


ROOT = Path(__file__).resolve().parent
MONEY = re.compile(r"(?<![\d.,])\d[\d,]*\.\d{2}(?!\d)")
DATE = re.compile(r"^\s*(\d{2}/\d{2})\s+")


def normalized_money(value):
    return format(Decimal(value.replace(",", "")), ".2f")


def compact(value):
    return " ".join(value.split())


def extract_statement(path, bank=None):
    """Return validated structured data and page count; reject unsupported PDFs.

    ``bank`` optionally checks the caller's bank selection against the layout.
    The existing BCA JSON schema is retained for compatibility.
    """
    from statement_pdf import detect_layout, extract_other_statement

    path = Path(path)
    pages = [p.extract_text(extraction_mode="layout") for p in PdfReader(path).pages]
    if not pages or not all(p.strip() for p in pages):
        raise ValueError(f"{path.name}: missing text layer; manual transcription required")
    detected = detect_layout(pages)
    expected = {"Bank Jakarta": "DKI", "Jakarta": "DKI"}.get(bank, bank)
    if expected and expected.upper() != detected.upper():
        raise ValueError(f"{path.name}: selected {bank}, but statement layout is {detected}")
    if detected != "BCA":
        return extract_other_statement(path, pages, detected), len(pages)
    return _extract_bca(path, pages), len(pages)


def _extract_bca(path, pages):
    text = "\n".join(pages)

    def field(pattern):
        match = re.search(pattern, text, re.MULTILINE)
        if not match:
            raise ValueError(f"{path.name}: missing field {pattern}")
        return compact(match.group(1))

    account_type = re.search(r"^\s*(REKENING (?:TAHAPAN|GIRO))\s*$", pages[0], re.MULTILINE)
    if "BCA" not in text or account_type is None:
        raise ValueError(
            f"{path.name}: unsupported statement format; "
            "expected a BCA REKENING TAHAPAN or REKENING GIRO statement"
        )
    parsed = {
        "bank_name": "BCA",
        "account_type": account_type.group(1),
        "account_number": field(r"NO\. REKENING\s*:\s*(\d+)"),
        "account_holder": field(r"^\s*(.*?)\s+NO\. REKENING\s*:"),
        "statement_period": field(r"PERIODE\s*:\s*([^\r\n]+)"),
        "currency": field(r"MATA UANG\s*:\s*(\w+)"),
    }
    summary = {}
    for printed, key in [("SALDO AWAL", "opening_balance"),
                         ("MUTASI CR", "credit_amount_total"),
                         ("MUTASI DB", "debit_amount_total"),
                         ("SALDO AKHIR", "closing_balance")]:
        matches = re.findall(printed + r"\s*:\s*([\d,]+\.\d{2})(?:[^\S\n]+(\d+))?", text)
        if len(matches) != 1:
            raise ValueError(f"{path.name}: expected one {printed} summary")
        amount, count = matches[0]
        summary[key] = normalized_money(amount)
        if printed.startswith("MUTASI"):
            if not count:
                raise ValueError(f"{path.name}: missing transaction count")
            summary["credit_transaction_count" if printed.endswith("CR") else "debit_transaction_count"] = int(count)
    # Match the key order of the existing labels.
    parsed["summary"] = {k: summary[k] for k in (
        "opening_balance", "credit_amount_total", "debit_amount_total", "closing_balance",
        "credit_transaction_count", "debit_transaction_count")}

    transactions = []
    opening = []
    for page in pages:
        lines = page.splitlines()
        headers = [(i, line) for i, line in enumerate(lines)
                   if all(s in line for s in ("TANGGAL", "KETERANGAN", "CBG", "MUTASI", "SALDO"))]
        if len(headers) != 1:
            raise ValueError(f"{path.name}: missing or ambiguous transaction table")
        header_index, header = headers[0]
        amount_column = header.index("CBG")
        for line in lines[header_index + 1:]:
            if "Bersambung ke halaman" in line or re.search(r"SALDO AWAL\s*:", line):
                break
            if not line.strip():
                continue
            date = DATE.match(line)
            if date:
                amounts = list(MONEY.finditer(line, amount_column))
                if "SALDO AWAL" in line:
                    if len(amounts) != 1:
                        raise ValueError(f"{path.name}: ambiguous opening balance row")
                    opening.append(normalized_money(amounts[0].group()))
                    continue
                if len(amounts) not in (1, 2):
                    raise ValueError(f"{path.name}: ambiguous transaction amounts: {line}")
                amount = amounts[0]
                suffix = line[amount.end():].lstrip()
                debit = suffix.startswith("DB")
                description = compact(line[date.end():amount.start()])
                transactions.append({
                    "date": date.group(1),
                    "description": description,
                    "debit": normalized_money(amount.group()) if debit else None,
                    "credit": None if debit else normalized_money(amount.group()),
                    "balance": normalized_money(amounts[1].group()) if len(amounts) == 2 else None,
                })
            else:
                if not transactions:
                    raise ValueError(f"{path.name}: orphan continuation line: {line}")
                transactions[-1]["description"] += " " + compact(line)

    if opening != [summary["opening_balance"]]:
        raise ValueError(f"{path.name}: opening balance row does not match summary")
    running = Decimal(summary["opening_balance"])
    for index, tx in enumerate(transactions, 1):
        running += Decimal(tx["credit"] or "0") - Decimal(tx["debit"] or "0")
        if tx["balance"] is not None and running != Decimal(tx["balance"]):
            raise ValueError(f"{path.name}: transaction {index} printed balance mismatch")
    for side in ("credit", "debit"):
        selected = [t[side] for t in transactions if t[side] is not None]
        if len(selected) != summary[f"{side}_transaction_count"]:
            raise ValueError(f"{path.name}: {side} transaction count mismatch")
        if sum(map(Decimal, selected), Decimal(0)) != Decimal(summary[f"{side}_amount_total"]):
            raise ValueError(f"{path.name}: {side} total mismatch")
    if running != Decimal(summary["closing_balance"]):
        raise ValueError(f"{path.name}: closing balance mismatch")
    parsed["transactions"] = transactions
    return {"gt_parse": parsed}


def main():
    sources = sorted((ROOT / "pdf").rglob("*.pdf"))
    if not sources:
        raise ValueError("No statements found under pdf/")
    # Parse and validate every source before creating any label files.
    results = [(p, *extract_statement(p)) for p in sources]
    key_dir = ROOT / "key"
    key_dir.mkdir(exist_ok=True)
    for source, label, _ in results:
        destination = (key_dir / source.relative_to(ROOT / "pdf")).with_suffix(".json")
        if destination.exists() and json.loads(destination.read_text(encoding="utf-8")) != label:
            raise FileExistsError(f"Refusing to replace differing label: {destination}")
    for source, label, pages in results:
        destination = (key_dir / source.relative_to(ROOT / "pdf")).with_suffix(".json")
        destination.parent.mkdir(parents=True, exist_ok=True)
        if not destination.exists():
            destination.write_text(json.dumps(label, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        transaction_count = sum(len(part['transactions']) for part in label.get('statements', [label['gt_parse']]))
        print(f"{destination.name}: {pages} page(s), {transaction_count} transactions; validated")
    print(f"Validated {len(results)} statements and their JSON labels.")


if __name__ == "__main__":
    main()
