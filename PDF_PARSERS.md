# PDF Parser Documentation

This document explains the bank statement PDF extraction interface, supported layouts, validation rules, and Excel integration.

The five bank PDF converters use `generate_labels.extract_statement()` as their shared entry point. Streamlit remains the interface, and the bank converters retain their Excel reporting responsibilities. The SLIK and TXT extraction workflows are separate and were not changed by this PDF migration.

## Supported layouts

| Bank | Supported text-based statement layout | Specific handling |
| --- | --- | --- |
| BCA | Rekening Tahapan and Rekening Giro | Reads transaction columns and wrapped descriptions; validates printed counts, totals, and balances. |
| BNI | Laporan Mutasi Rekening, including the supplied Taplus Muda statements | Reads dates, times, transaction descriptions, amounts, and balances. |
| BRI | Laporan Transaksi Finansial | Reads debit/credit columns and supports recap amounts printed on the page after their headings. |
| Mandiri | Laporan Rekening Koran and Kopra Account Statement | Reads transaction text in PDF content order so wrapped remarks remain attached to their transaction. |
| Bank Jakarta / Bank DKI | JakOne E-Statement | Supports multiple accounts and uses unchanged printed balances to resolve supported same-day ordering differences. |

Support is based on these layouts, not simply on the bank name. Scanned PDFs without a usable text layer require a separate OCR workflow. Unknown layouts and unreadable PDFs produce errors rather than silently falling back to the previous amount-repair parsers.

## Extraction flow

```text
Streamlit or a bank converter
  -> bank converter's process_pdf()
  -> services.pdf_statement_adapter.extract_frames()
  -> generate_labels.extract_statement()
     -> read PDF text and identify the layout
     -> parse metadata, printed recap, and transactions
     -> validate numeric consistency
  -> adapt structured data to transaction and summary dataframes
  -> Excel export and optional ETL processing
```

BCA's layout reader is in `generate_labels.py`. The additional bank readers are in `statement_pdf.py`.

## Python interface

Use the extraction function when you need structured data without writing an Excel file:

```python
from pathlib import Path
from generate_labels import extract_statement

label, page_count = extract_statement(Path("pdf/example.pdf"), bank="BCA")

statements = label.get("statements", [label["gt_parse"]])
for statement in statements:
    print(statement["bank_name"], statement["account_number"])
    print(statement["summary"])
    print(len(statement["transactions"]))
```

The `bank` argument is optional. Supported selections are `BCA`, `BNI`, `BRI`, `Mandiri`, and `DKI`; `Bank Jakarta` and `Jakarta` are also accepted aliases for DKI. Supplying a bank checks that the detected layout matches the selection. It does not force a different parser to accept an unsupported layout.

The function returns `(label, page_count)` and does not write files. Parsing or validation failures raise exceptions; callers should report them as extraction failures.

### Structured data

For a single-account PDF, `label["gt_parse"]` contains the statement:

| Field | Meaning |
| --- | --- |
| `bank_name` | Detected bank identifier. |
| `account_number` | Account number as a string, preserving leading zeros. |
| `account_type` | Product or account type identified by the reader. |
| `statement_period` | Period text extracted from the statement. |
| `currency` | Statement currency. |
| `summary` | Opening/closing balances, debit/credit totals, and transaction counts. |
| `transactions` | Ordered transaction records. |

Additional metadata varies by reader; BCA also extracts `account_holder`. Consumers should not assume that every bank supplies all additional fields.

Common transaction fields are `date`, `description`, `debit`, `credit`, and `balance`. Dates use `DD/MM`; the statement period supplies the year context. Non-BCA readers additionally provide `posting_datetime` and `source_page`.

Amounts in the structured result are decimal strings such as `"1250.00"`. An unused debit/credit side is `None`. Where BCA does not print a transaction balance, `balance` remains `None`. Calculated running balances are used for validation without replacing those missing printed values.

The summary contains:

```text
opening_balance
debit_amount_total
credit_amount_total
closing_balance
debit_transaction_count
credit_transaction_count
```

Counts are checked against printed counts where available. Otherwise, they are calculated from the extracted transactions; their presence in the result does not imply that the PDF printed them.

### Multiple accounts in Jakarta PDFs

A multi-account Jakarta PDF returns an additional `label["statements"]` list containing every account. `label["gt_parse"]` remains the first account for backward compatibility. **Iterate `statements` when present**, as shown in the example, to avoid omitting another account.

Jakarta records include `account_number` and `source_row`. The reader first preserves the printed sequence when it reconciles. For supported same-day ordering differences, it uses unchanged transaction amounts and printed balances to establish the sequence. It rejects a sequence that cannot be resolved. Original timestamps and source row numbers remain available even when output order changes.

## Numeric validation

Money calculations use `Decimal`. Depending on the fields available in the layout, the readers check:

- The printed opening balance and closing balance.
- Debit and credit totals against the printed recap.
- Debit and credit counts where printed.
- Each available printed running balance against accumulated transactions.
- The relationship `opening balance + credits - debits = closing balance`.

These checks run during extraction, independently of the optional ETL and `Data_Quality` settings. They do not change amounts or directions to force a statement to reconcile.

An error can reflect an unsupported layout, a parsing problem, or inconsistent source figures. Compare the reported values with the PDF before concluding that the statement itself is wrong. Numeric reconciliation does not establish the accuracy of every description or its accounting category.

### BRI recap across pages

Some BRI PDFs place the recap headings at the bottom of one page and the four amounts on the next page. `bri_recap()` searches between the recap heading and the following `Terbilang` section. It requires exactly one row containing only the four recap amounts, excluding page numbers, timestamps, and unrelated text.

This replaced the earlier seven-line search that failed on a supplied May 2026 BRI statement. The extracted recap still passes the same numeric validation.

## Dataframe and Excel integration

`services.pdf_statement_adapter.extract_frames(path, bank)` returns transaction and summary dataframes in the formats expected by the existing converters.

The common transaction columns are `Tanggal`, `Keterangan`, `DB`, `CR`, and `Saldo`. The adapter converts amount strings into Decimal values and uses pandas missing values where appropriate. Additional columns retain bank-specific information such as BNI transaction times, Mandiri posting dates, and Jakarta account identifiers and source row numbers.

The bank converters' `process_pdf()` functions call this adapter. Their `convert_pdf()` and `run_folder()` functions handle Excel export. Jakarta's export keeps account summaries and transaction sheets separate, including when one PDF contains multiple accounts.

## Summary classification

The adapter's `summary_metrics()` classifies explicitly described amounts into `Adm`, `Pajak`, `Bunga`, `Saldo Min`, and `JaGir`. Single-file and batch/yearly Summary sheets use the shared rules.

Classification is best-effort. A reference to "fee" or "admin" does not reliably identify the payment's purpose or who charged it. The known Mandiri principal-transfer case containing `Transfer Fee` is excluded, and unlabeled amounts remain unclassified.

Category amounts are already included in the debit/credit totals. Do not add them again. Read [PDF Summary Classification](PDF_SUMMARY_CLASSIFICATION.md) for the current rules and limitations.

## Dependencies and cache versions

Install dependencies from the project root:

```powershell
python -m pip install -r requirements.txt
```

The shared reader requires `pypdf` and `cryptography`. The adapter and Excel layers also use pandas and openpyxl. The primary bank PDF extraction path does not invoke Java/Tabula; retained legacy Tabula helpers require Java if called directly.

PDF parser versions in `extractors/bank_statement_extractor.py` are:

| Bank | Version |
| --- | --- |
| BCA, BNI, DKI, Mandiri | `2.0.0` |
| BRI | `2.0.1` |

These versions invalidate older parsed cache entries. BRI's patch version includes the recap page-break fix.

## Verification

Run the automated regression suite:

```powershell
python -m unittest discover -s tests -v
```

To exercise local PDFs through bank detection, parsing, and Excel export:

```powershell
python scripts/verify_pdf_parsers.py
```

The verification script reads PDFs recursively under `pdf/` and writes per-file results and workbooks under `outputs/pdf_parser_validation/`. Review `results.json` for accepted and rejected files. The generated data stays local and is not required by the synthetic regression tests.

Recorded checks from this migration:

- Initial full-folder run: 125 of 128 PDFs exported successfully, covering 4,885 reconciled transactions. Three files required review because their printed recap figures were inconsistent.
- After the BRI page-break fix: all 52 available BRI PDFs parsed and validated, and 24 automated tests passed.

These are separate, overlapping sample runs. They do not establish support for every statement layout or scanned PDFs.

## Implementation files

| File | Responsibility |
| --- | --- |
| [generate_labels.py](generate_labels.py) | Shared extraction entry point and BCA reader. |
| [statement_pdf.py](statement_pdf.py) | Additional bank readers, BRI recap handling, and numeric validation helpers. |
| [pdf_statement_adapter.py](services/pdf_statement_adapter.py) | Dataframe adaptation and Summary classification. |
| [bank_statement_extractor.py](extractors/bank_statement_extractor.py) | ETL parser registration, versions, caching, and lineage. |
| [test_pdf_statement_parsers.py](tests/test_pdf_statement_parsers.py) | Synthetic regression tests. |
| [verify_pdf_parsers.py](scripts/verify_pdf_parsers.py) | Local PDF and Excel verification. |
