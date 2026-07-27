import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from convert_mutasi_dki_txt import (
    build_month_summary,
    parse_txt_file,
    reconcile_transactions_with_balance,
    run_path,
)


SYNTHETIC_ACCOUNT = "00000"

OVERLAPPING_BLOCKS_TXT = f"""\
No. Rekening : {SYNTHETIC_ACCOUNT}
Periode Tgl. :  1/01/25 To 31/01/25
SALDO AWAL .......... 1,000.00
 2/01/25 DUPLICATE PAYMENT REF25 400.00 DB 600.00
SALDO AKHIR ......... 600.00

No. Rekening : {SYNTHETIC_ACCOUNT}
Periode Tgl. :  1/01/24 To 31/01/25
SALDO AWAL .......... 2,000.00
 2/01/24 TARGET DEBIT REF24 500.00 DB 1,500.00
 2/01/24 CLEAR BALANCE REF24 1,500.00 DB
31/12/24 FUNDING REF24 1,000.00 1,000.00
 2/01/25 DUPLICATE PAYMENT REF25 400.00 DB 600.00
SALDO AKHIR ......... 600.00
"""


class DkiTxtBlockAwareTests(unittest.TestCase):
    def write_txt(self, folder, content, name="statement.txt"):
        path = Path(folder) / name
        path.write_text(content, encoding="utf-8")
        return path

    def test_summary_classifies_interest_expense_and_provision_from_debits(self):
        transactions = pd.DataFrame([
            {
                "Keterangan": "BEBAN BUNGA PRK",
                "DB": Decimal("100.00"),
                "CR": pd.NA,
                "Saldo": Decimal("900.00"),
            },
            {
                "Keterangan": "BNG PINJAMAN",
                "DB": Decimal("75.00"),
                "CR": pd.NA,
                "Saldo": Decimal("825.00"),
            },
            {
                "Keterangan": "INTEREST EXPENSE",
                "DB": Decimal("25.00"),
                "CR": pd.NA,
                "Saldo": Decimal("800.00"),
            },
            {
                "Keterangan": "BUNGA CREDIT",
                "DB": pd.NA,
                "CR": Decimal("999.00"),
                "Saldo": Decimal("1799.00"),
            },
            {
                "Keterangan": "BAYAR PROVISI PRK",
                "DB": Decimal("300.00"),
                "CR": pd.NA,
                "Saldo": Decimal("1499.00"),
            },
            {
                "Keterangan": "BIAYA ADMINISTRASI",
                "DB": Decimal("10.00"),
                "CR": pd.NA,
                "Saldo": Decimal("1489.00"),
            },
        ])
        month_info = {
            "year": "2024",
            "month_order": 1,
            "month_label": "Jan-24",
            "opening": Decimal("1000.00"),
            "transactions": transactions,
        }

        summary = build_month_summary(
            "123",
            month_info,
            Path("statement.txt"),
        )

        self.assertEqual(summary["Bunga"], Decimal("200.00"))
        self.assertEqual(summary["Adm"], Decimal("310.00"))

    def test_overlapping_blocks_reconcile_independently_and_deduplicate(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            txt_path = self.write_txt(temp_dir, OVERLAPPING_BLOCKS_TXT)
            raw_df, accounts = parse_txt_file(txt_path)
            account_info = accounts[SYNTHETIC_ACCOUNT]
            cleaned = reconcile_transactions_with_balance(
                raw_df[raw_df["Account"] == SYNTHETIC_ACCOUNT],
                account_info,
            )

        target = cleaned[
            (cleaned["Tanggal"] == "2/01/24")
            & (cleaned["Keterangan"] == "TARGET DEBIT")
        ].iloc[0]
        duplicate_2025 = cleaned[
            (cleaned["Tanggal"] == "2/01/25")
            & (cleaned["Keterangan"] == "DUPLICATE PAYMENT")
        ]

        self.assertEqual(account_info["statement_block_count"], 2)
        self.assertEqual(account_info["duplicates_removed"], 1)
        self.assertEqual(target["DB"], Decimal("500.00"))
        self.assertTrue(pd.isna(target["CR"]))
        self.assertEqual(target["Saldo"], Decimal("1500.00"))
        self.assertEqual(len(duplicate_2025), 1)

    def test_balance_mismatch_preserves_explicit_source_amount(self):
        content = """\
No. Rekening : 123
Periode Tgl. :  1/01/24 To 31/01/24
SALDO AWAL .......... 1,000.00
 2/01/24 SOURCE DEBIT REF 100.00 DB 950.00
SALDO AKHIR ......... 950.00
"""
        with tempfile.TemporaryDirectory() as temp_dir:
            txt_path = self.write_txt(temp_dir, content)
            raw_df, accounts = parse_txt_file(txt_path)
            account_info = accounts["123"]
            cleaned = reconcile_transactions_with_balance(raw_df, account_info)

        row = cleaned.iloc[0]
        self.assertEqual(row["DB"], Decimal("100.00"))
        self.assertTrue(pd.isna(row["CR"]))
        self.assertIn("Nilai DB/CR sumber tidak diubah", row["ValidationIssue"])
        self.assertTrue(
            any(
                issue["issue_type"] == "balance_mismatch"
                for issue in account_info["validation_issues"]
            )
        )

    def test_multiset_deduplication_keeps_legitimate_occurrence_count(self):
        repeated_rows = """\
No. Rekening : 456
Periode Tgl. :  1/01/24 To 31/01/24
SALDO AWAL .......... 300.00
 2/01/24 SAME DAY FEE REF 100.00 DB 200.00
 2/01/24 SAME DAY FEE REF 100.00 DB 200.00
SALDO AKHIR ......... 200.00
"""
        content = repeated_rows + "\n" + repeated_rows
        with tempfile.TemporaryDirectory() as temp_dir:
            txt_path = self.write_txt(temp_dir, content)
            raw_df, accounts = parse_txt_file(txt_path)
            account_info = accounts["456"]
            cleaned = reconcile_transactions_with_balance(raw_df, account_info)

        self.assertEqual(len(cleaned), 2)
        self.assertEqual(account_info["duplicates_removed"], 2)

    def test_page_carry_balance_starts_a_new_block_chain(self):
        content = """\
No. Rekening : 789
Periode Tgl. :  1/01/24 To 31/01/24
SALDO AWAL .......... 1,000.00
 2/01/24 PAGE ONE DEBIT REF1 100.00 DB 900.00

No. Rekening : 789
Periode Tgl. :  1/01/24 To 31/01/24
PINDAHAN ............ 900.00
 3/01/24 PAGE TWO CREDIT REF2 100.00
SALDO AKHIR ......... 1,000.00
"""
        with tempfile.TemporaryDirectory() as temp_dir:
            txt_path = self.write_txt(temp_dir, content)
            raw_df, accounts = parse_txt_file(txt_path)
            account_info = accounts["789"]
            cleaned = reconcile_transactions_with_balance(raw_df, account_info)

        page_two = cleaned[cleaned["Tanggal"] == "3/01/24"].iloc[0]
        self.assertEqual(account_info["statement_block_count"], 2)
        self.assertEqual(account_info["saldo_akhir"], Decimal("1000.00"))
        self.assertEqual(page_two["Saldo"], Decimal("1000.00"))
        self.assertFalse(
            any(
                issue["issue_type"] == "balance_mismatch"
                for issue in account_info["validation_issues"]
            )
        )

    def test_generated_workbook_contains_correct_debit_and_quality_report(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            txt_path = self.write_txt(temp_path, OVERLAPPING_BLOCKS_TXT)
            output_files = run_path(txt_path, temp_path / "outputs")
            workbook_path = next(
                path
                for path in output_files
                if path.name == f"{SYNTHETIC_ACCOUNT}_2024.xlsx"
            )
            workbook = load_workbook(workbook_path, data_only=False, read_only=True)
            try:
                january = workbook["Jan_24"]
                quality = workbook["Data_Quality"]
                quality_messages = [
                    quality.cell(row=row, column=6).value
                    for row in range(2, quality.max_row + 1)
                ]

                self.assertEqual(january["C3"].value, 500)
                self.assertIsNone(january["D3"].value)
                self.assertEqual(january["E3"].value, 1500)
                self.assertTrue(
                    any("duplikat identik antarblok" in message for message in quality_messages)
                )
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
