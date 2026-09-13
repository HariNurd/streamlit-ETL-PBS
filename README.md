# Financial Document ETL & Reporting Platform

Local Streamlit ETL application for converting Indonesian financial documents into Excel reports and local analytical datasets.

The app currently supports:

- SLIK / iDeb PDF extraction
- Bank statement PDF conversion for BCA, BNI, DKI, BRI, and Mandiri
- Bank statement TXT conversion for DKI / Bank Jakarta
- Streamlit dashboard previews
- Excel report output
- Local Parquet, DuckDB, and SQLite ETL storage

The project is intended for local/internal use and should be run on localhost.

## Current ETL Flow

```text
Uploaded files or local folder
  -> document-specific parser / converter
     Bank PDFs: generate_labels.extract_statement()
                -> bank-specific parsing and numeric validation
                -> dataframe adapter
     SLIK PDFs and DKI TXT: their existing extraction workflows
  -> Excel output and Streamlit preview
  -> optional ETL processing
     -> staging dataframe, validation and lineage
     -> Parquet staging / DWH / mart
     -> DuckDB analytical views/tables
     -> SQLite job metadata and audit logs
```

Excel remains the publish/reporting output. Parquet and DuckDB are the local analytical storage layers.

## Main Features

### SLIK / iDeb

- Upload one or more PDF files.
- Process a local folder of SLIK PDF files.
- Extract credit facility data.
- Add row-level source lineage:
  - `job_id`
  - `source_file`
  - `source_file_hash`
  - `processed_at`
  - `parser_name`
  - `parser_version`
  - `loaded_from_cache`
- Optional file hash cache and incremental processing.
- Optional multiprocessing for larger SLIK batches.
- Streamlit dashboard preview.
- Excel dashboard workbook with optional `Data_Quality` sheet.

### Bank Statement PDF

- Upload one or more PDFs.
- Process a local folder of PDFs.
- Auto-detect or manually select bank.
- All five PDF converters use `generate_labels.extract_statement()` through `services/pdf_statement_adapter.py`.
- Streamlit remains the interface, with single-file and batch/yearly Excel reports.
- Bank Jakarta PDFs with multiple accounts are exported with separate account balances and summaries.
- BRI recap headings and amounts can span a page break.
- ETL wrappers store transaction data with lineage when enabled.

Supported **text-based statement layouts**:

| Bank | Layout |
| --- | --- |
| BCA | Rekening Tahapan and Rekening Giro |
| BNI | Laporan Mutasi Rekening, including the supplied Taplus Muda statements |
| BRI | Laporan Transaksi Finansial |
| Mandiri | Laporan Rekening Koran and Kopra Account Statement |
| Bank Jakarta / Bank DKI | JakOne E-Statement |

The PDF readers use Decimal arithmetic and check printed debit/credit totals, opening/closing balances, running balances, and transaction counts where available. Inconsistent statements, incomplete transactions, and unsupported layouts raise errors before Excel export. Amounts are not silently changed to make a statement reconcile. Scanned PDFs without a text layer require a separate OCR workflow.

Bank names alone do not guarantee support for every statement format. See [PDF parser documentation](PDF_PARSERS.md) for the extraction interface and layout details.

#### Summary categories

The `Summary` sheet includes **Adm**, **Pajak**, **Bunga**, **Saldo Min**, and **JaGir** when transaction descriptions provide matching evidence. Classification is shared across the five PDF converters and applies to single-file and batch/yearly exports.

These are **best-effort classifications**, not verified accounting categories. Words such as "admin", "fee", and "interest" may refer to different kinds of payments. The rules exclude the known Mandiri principal-transfer descriptions containing `Transfer Fee`, but other ambiguities can remain. Unlabeled transactions are not assigned a category based only on their amount.

Classification does not change transaction amounts or debit/credit totals. Category amounts are already included in those totals and must not be added again. See [Summary classification: meaning and limitations](PDF_SUMMARY_CLASSIFICATION.md).

### Bank Statement TXT

- Upload one or more TXT files.
- Process a local folder of TXT files.
- Current TXT converter:
  - DKI / Bank Jakarta TXT
- Existing account/year workbook output is preserved.
- ETL wrappers store transaction data with lineage when enabled.

The shared PDF parser and summary changes do not modify the TXT parser or its classification logic.

### Local Analytical Storage

The app can write:

- `data/staging/*.parquet`
- `data/dwh/*.parquet`
- `data/mart/*.parquet`
- `data/warehouse.duckdb`
- `data/metadata/etl_metadata.sqlite`

DuckDB is used for local analytical querying. SQLite is used for job metadata, file audit records, error logs, data quality results, and cleanup logs.

### Data Quality

Validators are included for:

- SLIK facilities
- Bank statement transactions

Validation issues can be shown in Streamlit and optionally added to generated Excel workbooks as `Data_Quality`.

Bank PDF parsing also performs mandatory numeric validation before returning transactions, independently of the optional ETL and `Data_Quality` settings. A printed recap mismatch may indicate inconsistent figures in the source document; review the reported values against the PDF. Successful reconciliation establishes numeric consistency, not the accuracy of description-based categories.

### Error Reports

Batch errors are captured and can be downloaded as Excel reports from Streamlit.

Suggested output pattern:

```text
outputs/error_report_<job_id>.xlsx
```

### Safe Cleanup

The app includes safe cleanup utilities for working files.

Cleanable folders:

- `temp/`
- `uploads/`
- `cache/` only when explicitly selected

Protected folders:

- `data/`
- `outputs/`
- `logs/`
- source code folders
- virtual environments
- project root

Cleanup does not follow symlinks and keeps failed source files by default.

## Streamlit Pages

- `SLIK / iDeb Converter`
- `Rekening Koran Converter`
- `Riwayat Proses`
- `Tentang / Instruksi`

## Advanced ETL Settings

The Streamlit app exposes advanced settings for:

- file hash cache
- incremental processing
- force reprocess
- SLIK multiprocessing
- worker count
- staging / DWH / mart Parquet writes
- DuckDB loading
- Data Quality sheet
- temp/uploads/cache cleanup
- manual cleanup

Recommended defaults:

- cache: on
- incremental processing: on
- force reprocess: off
- multiprocessing: off
- staging / DWH / mart: on
- DuckDB: on
- clean temp/uploads after success: on
- clean cache: off

## Project Structure

```text
project/
|-- app.py
|-- requirements.txt
|-- README.md
|-- PDF_PARSERS.md
|-- PDF_SUMMARY_CLASSIFICATION.md
|-- Streamlit Project Structure.txt
|
|-- converters/
|-- extractors/
|-- transformers/
|-- validators/
|-- loaders/
|-- services/
|-- storage/
|-- dashboards/
|-- scripts/
|-- tests/
|
|-- data/
|   |-- raw/
|   |-- staging/
|   |-- dwh/
|   |-- mart/
|   `-- metadata/
|
|-- logs/
|-- uploads/
|-- outputs/
|-- cache/
|-- temp/
|
|-- extract_slik_text_to_excel.py
|-- excel_dashboard.py
|-- generate_labels.py
|-- statement_pdf.py
|-- convert_mutasi_bca.py
|-- convert_mutasi_bni.py
|-- convert_mutasi_dki.py
|-- convert_mutasi_BRI.py
|-- convert_mutasi_mandiri.py
`-- convert_mutasi_dki_txt.py
```

Layer responsibilities:

| Folder | Responsibility |
|---|---|
| `extractors/` | ETL extraction wrappers, source lineage, and parser cache integration |
| `transformers/` | DWH and mart dataframe transformations |
| `validators/` | Data quality rules |
| `loaders/` | Parquet loaders |
| `storage/` | DuckDB, SQLite, lineage, hash, and path helpers |
| `services/` | Streamlit orchestration, PDF dataframe adapter and summary classification, cache, cleanup, logging, reports |
| `dashboards/` | Streamlit dashboard rendering |
| `converters/` | Existing bank converter wrappers |
| `scripts/` | Local PDF verification, documentation export, and repository checks |
| `tests/` | Automated regression tests using synthetic data |
| `data/` | Local analytical storage layers |
| `outputs/` | Generated Excel/ZIP/error reports, validation results, and documentation PDFs |
| `uploads/`, `temp/`, `cache/` | Local working folders |

## How to Run

Create and activate a virtual environment:

```powershell
python -m venv env
.\env\Scripts\activate
```

Install dependencies:

```powershell
pip install -r requirements.txt
```

Run the app:

```powershell
streamlit run app.py
```

Open:

```text
http://localhost:8501
```

## Dependencies

Main dependencies are listed in `requirements.txt`.

Notes:

- The primary bank PDF readers use `pypdf` and `cryptography`; install them through `requirements.txt`.
- The current shared bank PDF extraction path does not invoke Java/Tabula. Retained legacy Tabula helpers require Java if called directly; their Python dependencies remain listed in `requirements.txt`.
- SLIK PDF extraction uses PyMuPDF.
- DuckDB and PyArrow are required for local analytical storage and Parquet cache/layers.
- SQLite, hashlib, multiprocessing, pathlib, datetime, uuid, and logging are Python standard library modules.

## Generated Files

Generated files are intentionally ignored by Git:

- uploaded source PDFs/TXTs
- session temp files
- cache files
- Excel/ZIP outputs
- generated documentation PDFs and local validation reports
- Parquet files
- DuckDB databases
- SQLite metadata databases
- logs
- build/dist artifacts

The layer folders can be kept in the repo with `.gitkeep` files.

## Testing Checklist

Run the automated regression suite:

```powershell
python -m unittest discover -s tests -v
```

To check your local source PDFs, place them under `pdf/` (bank/year subfolders are supported), then run:

```powershell
python scripts/verify_pdf_parsers.py
```

The script checks bank detection, parsing, Excel export, and Summary category values. It writes per-file results to `outputs/pdf_parser_validation/results.json` and workbooks under `outputs/pdf_parser_validation/workbooks/`. Inspect the results for rejected files. Source PDFs and generated artifacts remain local and are not required by the automated regression suite.

Recorded validation runs:

- Initial migration: 125 of 128 PDFs exported successfully, covering 4,885 reconciled transactions. Three PDFs were rejected for inconsistent printed recaps.
- After the BRI recap page-break fix: all 52 available BRI PDFs parsed and validated; 24 automated tests passed.

These are separate, overlapping sample runs, not a guarantee for unseen layouts or a combined file count.

Recommended manual checks:

- SLIK single upload
- SLIK folder batch
- SLIK batch with one broken file
- SLIK cache hit after repeated run
- SLIK multiprocessing on/off
- SLIK Excel dashboard with `Data_Quality`
- BCA / BNI / DKI / BRI / Mandiri PDF conversion
- Bank Jakarta PDF containing multiple accounts
- BRI PDF with recap headings and amounts on different pages
- inconsistent PDF rejected with a useful validation message
- Summary categories checked against the underlying transaction descriptions
- DKI TXT upload and folder mode
- error report download
- job appears in `Riwayat Proses`
- Parquet files remain in `data/staging`, `data/dwh`, and `data/mart`
- DuckDB can be queried
- cleanup only removes `temp/`, `uploads/`, and optionally `cache/`

## Documentation PDFs

To generate the project change summary and a PDF version of the classification guide:

```powershell
python scripts/export_project_docs.py
```

The documents are saved as `outputs/project_documentation/PROJECT_CHANGES_SUMMARY.pdf` and `outputs/project_documentation/PDF_SUMMARY_CLASSIFICATION.pdf`. The change summary records this PDF parser migration and its follow-up fixes; its content is maintained in the export script. The classification PDF is rendered from `PDF_SUMMARY_CLASSIFICATION.md`.

## Security Notes

This project processes sensitive financial data.

Recommended precautions:

- Run only on localhost.
- Do not expose Streamlit directly to the public internet.
- Do not commit uploaded PDFs, TXT files, Excel outputs, cache files, Parquet datasets, DuckDB databases, SQLite metadata databases, or logs.
- Use only synthetic or fully anonymized fixtures in tests and documentation.
- Do not publish a ZIP or archive of the whole working directory; ignored files can still contain customer data.
- Keep cache cleanup off unless rerun performance is not important.
- Review generated reports before sharing.

Before publishing or pushing changes, run:

```powershell
python -B scripts/check_public_repo.py
python -B scripts/check_public_repo.py --history
```

Both checks must pass. `.gitignore` prevents normal additions but does not remove data from existing commits. If the history check fails, publish from a new clean repository or rewrite and verify every public branch and tag before changing repository visibility.

## Status

Current status:

- The five bank PDF converters share a validated extraction entry point while retaining Streamlit and Excel reporting.
- PDF summaries include description-based fee, tax, and interest categories; TXT workflows remain unchanged.
- PDF cache versions are 2.0.0 for BCA, BNI, DKI, and Mandiri, and 2.0.1 for BRI after the recap page-break fix. Older cached parser results are invalidated by version.
- ETL layers, local analytical storage, metadata audit, caching, incremental controls, folder mode, error reports, validation, cleanup, logging, and job history are implemented.
- Multiprocessing is available for SLIK and is off by default.
