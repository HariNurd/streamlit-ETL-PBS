The converter creates five separate summary categories: `Adm`, `Pajak`, `Bunga`, `Saldo Min`, and `JaGir`. There is no combined “Bunga Adm” category.

### Processing flow

1. Each TXT transaction is parsed into:

   - `Keterangan`
   - `Amount`
   - `DB` if the mutation contains the `DB` marker
   - `CR` if it does not contain `DB`

   This happens in [convert_mutasi_dki_txt.py](convert_mutasi_dki_txt.py).

2. Transactions are grouped by account, year, and month.

3. For each month, `sum_by_description()`:

   - Searches `Keterangan` using a case-insensitive regular expression.
   - Takes either the `DB` or `CR` column.
   - Ignores empty amounts.
   - Sums every matching transaction.
   - Returns `0.00` if nothing matches.

   See `sum_by_description()` and `build_month_summary()` in [convert_mutasi_dki_txt.py](convert_mutasi_dki_txt.py).

### Current classification rules

| Summary column | Amount used | Description patterns |
|---|---:|---|
| `Adm` | DB only | `ADM`, `ADMIN`, `BIAYA ADMIN`, or `FEE` |
| `Pajak` | DB only | `PPH`, `PAJAK`, or `TAX` |
| `Bunga` | CR only | `BUNGA` or `INTEREST` |
| `Saldo Min` | DB only | `SALDO MIN` |
| `JaGir` | CR only | `JASA GIRO` or `JAGIR` |

Synthetic examples:

- `BIAYA ADMINISTRASI 10,000.00 DB` → `Adm`
- `PAJAK JAGIR 2,000.00 DB` → `Pajak`, not `JaGir`
- `JAGIR BULANAN 8,000.00 CR` → `JaGir`
- `BUNGA BULANAN 50,000.00 CR` → `Bunga`
- `BEBAN BUNGA PRK 1,000.00 DB` → not included in `Bunga`, because `Bunga` only sums CR
- `SALDO MIN 5,000.00 DB` → `Saldo Min`

### Important limitations

- The `Bunga` column represents interest-related credits only. It does not include debit interest expenses such as `BEBAN BUNGA PRK`.
- `BNG ...` is not recognized as `Bunga` because the rule only searches for `BUNGA` or `INTEREST`.
- `BYADM ...` may not be recognized as `Adm`, because `ADM` must appear as a separate word. `BIAYA ADMINISTRASI` and `ADM ...` are recognized.
- The account-header interest rate—such as `Bunga: 12.00%`—is not used. Only transaction descriptions and DB/CR amounts are considered.
- Categories are not mutually exclusive. For example, a DB description containing both `FEE` and `PAJAK` would be counted in both `Adm` and `Pajak`.
- These categories do not affect the transaction totals or balances. They are additional monthly analytical summaries.

Use synthetic or fully anonymized fixtures when validating category counts. Production statements must remain outside Git.
