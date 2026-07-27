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
  -> legacy parser / converter
  -> staging dataframe
  -> validation and lineage
  -> Parquet staging / DWH / mart
  -> DuckDB analytical views/tables
  -> SQLite job metadata and audit logs
  -> Excel output and Streamlit preview
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
- Supported banks:
  - BCA
  - BNI
  - DKI
  - BRI
  - Mandiri
- Existing Excel workbook output is preserved.
- ETL wrappers store transaction data with lineage when enabled.

### Bank Statement TXT

- Upload one or more TXT files.
- Process a local folder of TXT files.
- Current TXT converter:
  - DKI / Bank Jakarta TXT
- Existing account/year workbook output is preserved.
- ETL wrappers store transaction data with lineage when enabled.

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
| `extractors/` | Wrapper extraction functions around legacy parsers |
| `transformers/` | DWH and mart dataframe transformations |
| `validators/` | Data quality rules |
| `loaders/` | Parquet loaders |
| `storage/` | DuckDB, SQLite, lineage, hash, and path helpers |
| `services/` | Streamlit orchestration, cache, cleanup, logging, reports |
| `dashboards/` | Streamlit dashboard rendering |
| `converters/` | Existing bank converter wrappers |
| `data/` | Local analytical storage layers |
| `outputs/` | Generated Excel/ZIP/error report files |
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

- PDF bank statement converters using `tabula-py` require Java.
- SLIK PDF extraction uses PyMuPDF.
- DuckDB and PyArrow are required for local analytical storage and Parquet cache/layers.
- SQLite, hashlib, multiprocessing, pathlib, datetime, uuid, and logging are Python standard library modules.

## Generated Files

Generated files are intentionally ignored by Git:

- uploaded source PDFs/TXTs
- session temp files
- cache files
- Excel/ZIP outputs
- Parquet files
- DuckDB databases
- SQLite metadata databases
- logs
- build/dist artifacts

The layer folders can be kept in the repo with `.gitkeep` files.

## Testing Checklist

Recommended manual checks:

- SLIK single upload
- SLIK folder batch
- SLIK batch with one broken file
- SLIK cache hit after repeated run
- SLIK multiprocessing on/off
- SLIK Excel dashboard with `Data_Quality`
- BCA / BNI / DKI / BRI / Mandiri PDF conversion
- DKI TXT upload and folder mode
- error report download
- job appears in `Riwayat Proses`
- Parquet files remain in `data/staging`, `data/dwh`, and `data/mart`
- DuckDB can be queried
- cleanup only removes `temp/`, `uploads/`, and optionally `cache/`

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

- Existing converter and Excel output workflows are preserved.
- ETL layers, local analytical storage, metadata audit, caching, incremental controls, folder mode, error reports, validation, cleanup, logging, and job history are implemented.
- Multiprocessing is available for SLIK and is off by default.
