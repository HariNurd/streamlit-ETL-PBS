# SLIK Extractor Workflow

Date: 2026-07-20

This document explains how the SLIK extractor works in plain language. It is written so a non-programmer can understand where data comes from, why rows appear or disappear, and which part of the script to change for common updates.

## Files Involved

The SLIK converter is spread across a few files, but most extraction logic is in one file.

| File | Role |
| --- | --- |
| `extract_slik_text_to_excel.py` | Main parser and Excel writer. This is where PDF text is read, facilities are detected, rows are built, and workbook sheets are written. |
| `extractors/slik_extractor.py` | Wrapper around the main parser. Handles sequential or multiprocessing extraction and stores the parser version used by cache. |
| `services/slik_service.py` | App service layer. Handles cache, metadata, validation, storage, and calls the extractor. |
| `excel_dashboard.py` | Builds `Data_Clean` and dashboard sheets from extracted SLIK rows. |
| `validators/slik_validator.py` | Checks extracted rows for data quality issues. |
| `app.py` | Streamlit user interface. Uploads files, starts processing, previews dashboard, and exports workbooks. |

For most SLIK extraction changes, start with:

```text
extract_slik_text_to_excel.py
```

If the parser behavior changes, also update:

```text
extractors/slik_extractor.py
```

because that file contains the parser cache version.

## Big Picture

The converter follows this flow:

```text
PDF file
  |
  v
Read PDF text lines
  |
  v
Find debtor identity and report number
  |
  v
Split text into facility blocks
  |
  v
Parse each facility block into one row
  |
  v
Keep only important facilities
  |
  v
Deduplicate rows
  |
  v
Validate data quality
  |
  v
Write Excel workbook
```

The current workbook can contain these SLIK-related sheets:

| Sheet | Meaning |
| --- | --- |
| `SLIK` | Main backward-compatible summary sheet. Keeps the existing column layout. |
| `Bank_Garansi` | Detail sheet for `Garansi Yang Diberikan` / bank guarantee facilities. |
| `Data_Clean` | Cleaned data used by the dashboard. |
| `Dashboard_Summary` | KPI summary. |
| `Dashboard_Risk` | Risk tables and charts. |
| `Dashboard_Maturity_Rate` | Maturity and rate analysis. |
| `Angsuran` | Installment schedule for credit facilities only. Bank garansi rows are excluded. |
| `Data_Quality` | Optional sheet listing validation issues. |

## Step 1: Read PDF Text

Main functions:

```text
read_pdf_page_lines()
read_pdf_lines()
clean_text()
clean_lines()
```

What happens:

1. The script opens the PDF using PyMuPDF / `fitz`.
2. Every page is converted into plain text.
3. The text is split into lines.
4. Extra spaces, line breaks, and blank lines are cleaned.

Important idea:

The parser does not read the PDF visually like a human. It reads a long list of text lines in the order returned by the PDF library. If a label appears in a strange order, the parser follows that order.

Example extracted text can look like:

```text
Garansi Yang Diberikan
Tanggal Jatuh Tempo
17 Juni 2026
Tujuan Garansi
Lainnya
No Rekening
<ACCOUNT_NUMBER>
Kualitas
1 - Lancar
...
Kondisi
Fasilitas Aktif
```

## Step 2: Get Debtor Information

Main functions:

```text
extract_nama_debitur()
extract_nik_debitur()
extract_npwp_debitur()
extract_debitur_identity()
extract_nomor_laporan()
```

What happens:

1. The first page is searched for debtor information.
2. The debtor name is cleaned and checked against blocklists.
3. The identity value is extracted into `NIK/NPWP`.
4. The parser tries to find NIK first. If no valid NIK is found, it tries to find NPWP.
5. The report number is found using a pattern like:

```text
<REPORT_NUMBER>
```

Why blocklists exist:

PDF text often includes headers like `NOMOR LAPORAN`, `DATA POKOK DEBITUR`, or city names. The parser blocks these from being mistaken as debtor names.

Relevant constants:

```text
DEBTOR_NAME_EXACT_BLOCKLIST
DEBTOR_NAME_BLOCKED_PHRASES
DEBTOR_NAME_BLOCKED_TOKENS
```

Change these only if debtor names are being extracted incorrectly.

Identity output rule:

```text
Individual debtor with NIK    -> NIK/NPWP = NIK
Company debtor with NPWP      -> NIK/NPWP = NPWP
No identity found             -> NIK/NPWP is blank
```

## Step 3: Decide Where Facility Blocks Start

Main constant:

```text
FACILITY_START_LABELS
```

Current facility starts:

```text
Kredit/Pembiayaan
Garansi Yang Diberikan
```

Main function:

```text
split_facility_blocks()
```

What happens:

1. The script scans all PDF text lines.
2. Every line matching `FACILITY_START_LABELS` becomes the start of a facility block.
3. Each block continues until the next facility start or until an end label is found.

Example:

```text
Garansi Yang Diberikan   <- block starts here
Tanggal Jatuh Tempo
17 Juni 2026
...
Kondisi
Fasilitas Aktif
...
Agunan                   <- block ends before this
Penjamin
```

## Step 4: Decide Where Facility Blocks End

Main constant:

```text
SECTION_END_LABELS
```

Current section end labels include:

```text
Agunan
Penjamin
Irrecovable L/C
Surat Berharga
Fasilitas Lain
Nomor Laporan
Operator
```

Plain-language meaning:

These words tell the parser, "The current facility detail section is finished. Stop reading this block."

Important note:

`Garansi Yang Diberikan` is now a start label, not an end label. This is necessary because bank guarantee rows begin with that label.

## Step 5: Identify Facility Type

Main constants:

```text
CREDIT_FACILITY_TYPE = "Kredit/Pembiayaan"
BANK_GARANSI_FACILITY_TYPE = "Garansi Yang Diberikan"
```

Main function:

```text
facility_type_from_block()
```

What happens:

The first line of a block tells the parser what kind of facility it is.

Example:

```text
Kredit/Pembiayaan
```

means normal credit / financing.

```text
Garansi Yang Diberikan
```

means bank garansi.

## Step 6: Read Values After Labels

Main constants:

```text
FIELD_LABELS
SECTION_END_LABELS
```

Main functions:

```text
is_label()
next_values_after_label()
value_after_label()
```

Plain-language rule:

When the parser sees a label, it collects the next lines as the value until it reaches another known label.

Example:

```text
No Rekening
<ACCOUNT_NUMBER>
Kualitas
1 - Lancar
```

The parser reads:

```text
No Rekening = <ACCOUNT_NUMBER>
```

Then it stops because `Kualitas` is another known label.

Why `FIELD_LABELS` is important:

If a label is missing from `FIELD_LABELS`, the parser may accidentally read too much.

Bad example:

```text
Kondisi = Fasilitas Aktif Nama Yang Dijamin PERUSAHAAN CONTOH
```

Good example:

```text
Kondisi = Fasilitas Aktif
Nama Yang Dijamin = PERUSAHAAN CONTOH
```

To fix this kind of problem, add the missing label to `FIELD_LABELS`.

## Step 7: Decide Whether a Facility Should Be Kept

Main function:

```text
parse_facility_block()
```

Helper functions:

```text
is_active_condition()
is_dihapusbukukan_condition()
is_hapus_tagih_condition()
extract_kolektibilitas()
```

The parser keeps a facility if at least one of these is true:

| Condition | Meaning |
| --- | --- |
| `Kondisi` is active | The facility is active. |
| `Kondisi` contains `Dihapusbukukan` | The facility is written off. |
| `Kondisi` contains `Hapus Tagih` | The facility is charged off. |
| `Kol.` is not `1` | The collectibility is worse than normal. |

The parser skips rows like:

```text
Kondisi = Lunas
Kol. = 1
```

because they are closed and normal.

## Step 8: Build One Row Per Facility

Main function:

```text
parse_facility_block()
```

For every kept facility, the parser creates one row.

The main `SLIK` sheet uses these columns:

```text
Nama
NIK/NPWP
Bank
Penggunaan
Plafond (Rp)
Baki Debet (Rp)
Kol.
Awal
Jatuh Tempo
Rate
No Laporan
Keterangan
```

Bank garansi rows also get internal detail fields:

```text
Jenis Fasilitas
No Rekening
No Akad Awal
No Akad Akhir
Jenis Garansi
Tujuan Garansi
Nama Yang Dijamin
Nominal (Rp)
Setoran Jaminan (Rp)
Tanggal Diterbitkan
Tanggal Akad Akhir
```

These internal fields are used to write the separate `Bank_Garansi` sheet.

For bank garansi rows, the main `SLIK` sheet sets:

```text
Penggunaan = Bank Garansi
```

This keeps the old `SLIK` sheet readable even though bank garansi does not have the same `Jenis Penggunaan` field as normal credit / financing.

## Step 9: Fallback Row If Nothing Is Found

Main function:

```text
build_no_matching_facility_row()
```

If the parser finds debtor information and report number, but no active/problem facilities, it creates one fallback row:

```text
Tidak ada Fasilitas Aktif
```

This is why an Excel file can still have one row even when no real facility is found.

Important helper:

```text
count_real_facility_rows()
```

This counts real rows and ignores fallback rows.

## Step 10: Deduplicate Rows

Main function:

```text
deduplicate_facilities()
```

Plain-language meaning:

If the same facility appears more than once, keep only one copy.

The dedupe check includes fields like:

```text
Nama
NIK/NPWP
Jenis Fasilitas
No Rekening
No Akad Akhir
Bank
Plafond (Rp)
Baki Debet (Rp)
Kol.
Jatuh Tempo
Keterangan
```

Why this matters:

Bank garansi rows may have similar amounts and dates, so `No Rekening` and `No Akad Akhir` help prevent different guarantees from being merged incorrectly.

## Step 11: Cache and Parser Version

Main file:

```text
extractors/slik_extractor.py
```

Important constant:

```text
SLIK_PARSER_VERSION = "1.2.0"
```

Plain-language meaning:

The app can cache parser results. Cache is based on:

```text
PDF file hash + parser version
```

If parser logic changes but the version is not changed, the app may reuse old results.

Rule of thumb:

If extraction logic changes, bump the version.

Example:

```text
1.2.0 -> 1.2.1
```

or for bigger behavior changes:

```text
1.2.0 -> 1.3.0
```

## Step 12: Validation

Main file:

```text
validators/slik_validator.py
```

Main function:

```text
validate_slik_facilities()
```

Checks include:

| Check | Meaning |
| --- | --- |
| `Nama` required | Real facility rows must have a debtor name. |
| `Bank` required | Real facility rows must have a bank. |
| `Kol.` range | `Kol.` should be 1 to 5. |
| Negative amounts | Plafond and baki debet should not be negative. |
| Baki debet over plafond | Warning if outstanding amount exceeds plafond. |
| Rate range | Rate should be between 0 and 100 percent. |
| Date parse | Jatuh tempo should be readable as a date. |
| Fallback exposure | Fallback rows should not contain exposure. |

Validation results may appear in the `Data_Quality` sheet.

## Step 13: Write Excel Workbook

Main function:

```text
export_slik_excel_dashboard()
```

What it writes:

1. `SLIK`
2. `Bank_Garansi`, only if bank garansi rows exist
3. Dashboard sheets, if enabled
4. `Angsuran`, if enabled
5. `Data_Quality`, if enabled and issues exist

Important behavior:

Bank garansi rows are included in the main `SLIK` sheet and `Bank_Garansi` sheet.

Bank garansi rows are excluded from `Angsuran`.

Reason:

`Angsuran` means loan installment schedule. Bank guarantees are contingent facilities, not normal amortizing loans.

## Important Lists And When To Change Them

### `FACILITY_START_LABELS`

Change this when a new kind of facility should be parsed as its own block.

Example:

```text
Kredit/Pembiayaan
Garansi Yang Diberikan
```

Add a new label here only if that label begins a real facility detail section.

### `SECTION_END_LABELS`

Change this when a facility block keeps reading too far.

Example symptom:

One facility row accidentally includes text from `Agunan`, `Penjamin`, or footer sections.

Add the stopping word to `SECTION_END_LABELS`.

### `FIELD_LABELS`

Change this when one value captures text from the next field.

Example symptom:

```text
Kondisi = Fasilitas Aktif Nama Yang Dijamin PERUSAHAAN CONTOH
```

Fix:

Add `Nama Yang Dijamin` to `FIELD_LABELS`.

### `OUTPUT_COLUMNS`

Change this only if the main `SLIK` sheet should have a new column.

Be careful:

Many dashboards and existing users may expect the current `SLIK` layout.

### `BANK_GARANSI_COLUMNS`

Change this when the `Bank_Garansi` sheet needs a new detail column.

This is safer than changing `OUTPUT_COLUMNS` because it affects only the bank garansi detail sheet.

### `FACILITY_DETAIL_COLUMNS`

Change this when a parsed value needs to be stored internally before export.

Usually, if you add a column to `BANK_GARANSI_COLUMNS`, you may also need to add it here and fill it inside `parse_facility_block()`.

## Common Change Recipes

### Recipe 1: A Value Reads Too Much Text

Symptom:

```text
Kondisi = Fasilitas Aktif Nama Yang Dijamin PERUSAHAAN CONTOH
```

Likely cause:

`Nama Yang Dijamin` is not recognized as a label.

Fix:

1. Open `extract_slik_text_to_excel.py`.
2. Find `FIELD_LABELS`.
3. Add the missing label exactly as it appears in the PDF text.
4. Re-run the sample PDF.

### Recipe 2: A Facility Type Is Missing

Symptom:

The PDF clearly has a section, but parser returns fallback:

```text
Tidak ada Fasilitas Aktif
```

Likely cause:

The section start is not listed in `FACILITY_START_LABELS`.

Fix:

1. Inspect raw PDF text.
2. Find the exact first line of the facility section.
3. Add that line to `FACILITY_START_LABELS`.
4. If the new type has special columns, add a separate detail sheet like `Bank_Garansi`.
5. Bump `SLIK_PARSER_VERSION`.

### Recipe 3: A Facility Block Includes Footer Text

Symptom:

Fields contain `Nomor Laporan`, `Operator`, or other footer text.

Likely cause:

The parser does not know where the facility section ends.

Fix:

1. Open `extract_slik_text_to_excel.py`.
2. Find `SECTION_END_LABELS`.
3. Add the footer or section label that should stop the block.
4. Re-run the sample PDF.

### Recipe 4: A New Column Is Needed In `Bank_Garansi`

Steps:

1. Add the column name to `BANK_GARANSI_COLUMNS`.
2. If the value is not already stored internally, add it to `FACILITY_DETAIL_COLUMNS`.
3. In `parse_facility_block()`, fill that new field from `value_after_label()`.
4. If it is a money field, update `style_bank_garansi_sheet()` so Excel formats it as currency.
5. Re-run the sample PDF and inspect the `Bank_Garansi` sheet.

### Recipe 5: Old Results Keep Appearing After A Fix

Symptom:

You changed the parser, but Streamlit output still looks unchanged.

Likely cause:

Cache is being reused.

Fix options:

1. Bump `SLIK_PARSER_VERSION` in `extractors/slik_extractor.py`.
2. Disable cache in the app options.
3. Clear the SLIK cache.

Preferred option for parser behavior changes:

```text
Bump SLIK_PARSER_VERSION
```

## How To Test A Change

Use a known PDF:

```text
synthetic_slik_company.pdf
```

Expected result after the bank garansi update:

| Check | Expected |
| --- | --- |
| Direct parser rows | 3 synthetic sample rows |
| Facility type | `Garansi Yang Diberikan` |
| `SLIK` sheet | 3 active rows plus total |
| `Bank_Garansi` sheet | 3 detail rows |
| `Angsuran` sheet | No installment rows for these bank garansi facilities |
| Data quality | 0 issues |
| Identity column | `NIK/NPWP` should contain company NPWP when NIK is not available |
| Bank garansi usage | `Penggunaan` should be `Bank Garansi` |
| Cache | Should use parser version `1.2.0` or newer |

Useful commands:

```powershell
env\Scripts\python.exe -B -c "import extract_slik_text_to_excel, extractors.slik_extractor, services.slik_service; print('imports ok')"
```

```powershell
env\Scripts\python.exe -B -c "from extract_slik_text_to_excel import parse_slik_pdf; df = parse_slik_pdf('synthetic_slik_company.pdf'); print(df[['Nama','NIK/NPWP','Jenis Fasilitas','Bank','Penggunaan','No Rekening','Plafond (Rp)','Keterangan']])"
```

## Safety Checklist Before Finishing A Parser Change

Before saying a SLIK parser change is done, check:

| Question | Why it matters |
| --- | --- |
| Did direct parsing return the expected rows? | Confirms the extractor logic works. |
| Did the Streamlit/service path avoid stale cache? | Confirms users will see the new result. |
| Did `Data_Quality` stay clean? | Confirms row values are valid. |
| Did the workbook sheets appear as expected? | Confirms export behavior works. |
| Did `SLIK_PARSER_VERSION` change if parser behavior changed? | Prevents old cached results. |
| Did existing `SLIK` columns stay compatible? | Prevents breaking old reports. |

## Quick Mental Model

Think of the extractor like a form reader:

1. It turns the PDF into a long checklist of text lines.
2. It looks for section titles like `Kredit/Pembiayaan` or `Garansi Yang Diberikan`.
3. It treats each section as one possible facility.
4. It reads fields by looking for labels like `Kondisi`, `Plafon`, and `No Rekening`.
5. It keeps the facility only if it is active, written off, charged off, or has non-normal collectibility.
6. It writes normal shared fields to `SLIK`.
7. It writes bank guarantee-specific fields to `Bank_Garansi`.

When something is wrong, first ask:

```text
Did the parser find the block?
Did it stop the block in the right place?
Did it recognize the labels inside the block?
Did it keep or skip the row for the right reason?
Did cache hide the new result?
```
