# SLIK Bank Garansi Change Log

Date: 2026-07-20

## Purpose

Capture the SLIK converter changes made to support `Garansi Yang Diberikan` / bank garansi facilities separately from normal `Kredit/Pembiayaan` facilities.

## Source Files Updated

### `extract_slik_text_to_excel.py`

- Added `Garansi Yang Diberikan` as a supported facility start label.
- Added facility type constants:
  - `KREDIT/PEMBIAYAAN` represented internally as `Kredit/Pembiayaan`.
  - Bank garansi represented internally as `Garansi Yang Diberikan`.
- Added internal detail columns for facility-specific data.
- Added a dedicated bank garansi output schema with columns such as:
  - `No Rekening`
  - `Jenis Garansi`
  - `Tujuan Garansi`
  - `Nama Yang Dijamin`
  - `Nominal (Rp)`
  - `Setoran Jaminan (Rp)`
  - `Tanggal Diterbitkan`
  - `No Akad Awal`
  - `No Akad Akhir`
- Added guarantee-specific field labels to prevent parser values from bleeding into the next label.
- Added `facility_type_from_block()` to identify the current facility type from each parsed block.
- Updated `parse_facility_block()` so bank garansi rows keep their specific fields while still contributing to the existing normalized `SLIK` sheet.
- Added `build_bank_garansi_output_df()`, `style_bank_garansi_sheet()`, and `write_bank_garansi_sheet()`.
- Updated Excel export so:
  - Existing `SLIK` sheet remains backward compatible.
  - New `Bank_Garansi` sheet is written when bank garansi rows exist.
  - `Angsuran` excludes bank garansi rows because guarantees should not generate loan installment schedules.
- Updated deduplication to include `Jenis Fasilitas`, `No Rekening`, and `No Akad Akhir` when available.
- Renamed the main identity column from `NIK` to `NIK/NPWP`.
- Added NPWP extraction for company debtors when a valid NIK is not available.
- Set `Penggunaan` to `Bank Garansi` for bank garansi rows in the main `SLIK` sheet.

### `extractors/slik_extractor.py`

- Imported `FACILITY_DATA_COLUMNS` for the expanded parser output schema.
- Bumped `SLIK_PARSER_VERSION` from `1.0.0` to `1.1.0`.
- Bumped `SLIK_PARSER_VERSION` again from `1.1.0` to `1.2.0` after the `NIK/NPWP` schema and NPWP extraction update.
- Updated empty dataframe returns to preserve the expanded facility schema.

### `excel_dashboard.py`

- Updated dashboard source columns from `NIK` to `NIK/NPWP`.
- Kept backward compatibility for older dataframes that still contain `NIK`.

### `dashboards/streamlit_slik_dashboard.py`

- Updated the dashboard filter label and column reference from `NIK` to `NIK/NPWP`.

### `validators/slik_validator.py`

- Updated identity validation from NIK-only to `NIK/NPWP`.
- Allows empty identity or 15-16 digit numeric identity values.

## Verification Performed

Test file:

- `synthetic_slik_company.pdf`

Results:

- Direct parser returned 3 synthetic sample rows.
- All 3 rows were identified as `Garansi Yang Diberikan`.
- Main `SLIK` sheet used `NIK/NPWP` and filled the company NPWP.
- Bank garansi rows used `Penggunaan = Bank Garansi`.
- Existing `SLIK` sheet showed 3 active rows instead of the fallback row.
- New `Bank_Garansi` sheet was created with 3 detailed rows.
- `Angsuran` did not generate schedules for bank garansi rows.
- Data quality result had 0 issues.
- Service path used parser version `1.2.0` and did not reuse stale earlier-version cache.
- Import validation passed with bytecode writing disabled.

Verification workbook:

- `temp/synthetic_slik_bank_garansi_check.xlsx`

## Cache Note

The old incorrect fallback result was cached under parser version `1.0.0`. The version bump to `1.1.0` made the app parse this PDF fresh instead of loading the stale cached fallback row.

Parser version `1.2.0` is used for the later identity-column and NPWP extraction update.
