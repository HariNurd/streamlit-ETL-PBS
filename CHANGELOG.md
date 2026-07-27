# Changelog

## 2026-07-24

### Fixed

- Changed the Bank DKI TXT parser to preserve each repeated account/page occurrence as a separate statement block.
- Reconciled balances independently inside each statement block so a later-period balance cannot overwrite an earlier-period transaction.
- Kept explicitly parsed transaction amounts and DB/CR directions as the source of truth. Balance differences now validate or fill missing balances instead of silently replacing source amounts.
- Read `PINDAHAN` balances as the opening balance for continuation-page blocks.
- Removed exact transaction duplicates only when they occur across different statement blocks. Duplicate occurrence counts are preserved, so legitimate same-day repeated transactions are not collapsed into one row.
- Retained potentially conflicting overlap rows and reported them for review instead of deleting them.
- Updated ETL reconciliation validation to evaluate each account separately instead of combining all accounts in one source file.

### Added

- Added a `Data_Quality` worksheet to generated Bank DKI TXT workbooks. It reports multiple blocks, removed exact duplicates, balance mismatches, and overlap conflicts.
- Added statement-block, source-line, inferred-balance, and validation metadata to parsed DKI TXT transaction rows for traceability.
- Added DKI TXT parser diagnostics to ETL summary and data-quality results.
- Added regression tests covering nonchronological overlapping blocks, exact duplicate removal, repeated-transaction counts, source-value preservation, page carry balances, and generated workbook values.

### Changed

- Increased the DKI TXT ETL parser/cache version from `1.0.0` to `2.0.0` so cached results created by the old reconciliation logic are not reused.
- The original TXT input is never edited; duplicate cleanup happens only in parsed data and generated outputs.
- Changed the DKI TXT `Bunga` summary to classify interest expense from debit (`DB`) transactions instead of credit (`CR`) transactions.
- Added `BNG` as an accepted `Bunga` description keyword and `PROVISI` as an accepted `Adm` description keyword.
