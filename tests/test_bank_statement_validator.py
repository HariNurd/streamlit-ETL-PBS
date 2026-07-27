import unittest

import pandas as pd

from extractors.bank_statement_extractor import DKI_TXT_PARSER_VERSION
from validators.bank_statement_validator import validate_bank_transactions


class BankStatementValidatorTests(unittest.TestCase):
    def test_reconciliation_is_scoped_to_each_account(self):
        transactions = pd.DataFrame([
            {
                "source_file": "multi-account.txt",
                "Account": "A",
                "Tanggal": "02/01/24",
                "Keterangan": "DEBIT",
                "DB": 100,
                "CR": pd.NA,
                "Saldo": 900,
            },
            {
                "source_file": "multi-account.txt",
                "Account": "B",
                "Tanggal": "02/01/24",
                "Keterangan": "CREDIT",
                "DB": pd.NA,
                "CR": 100,
                "Saldo": 2100,
            },
        ])
        summaries = pd.DataFrame([
            {
                "source_file": "multi-account.txt",
                "Account": "A",
                "Saldo Awal": 1000,
                "Saldo Akhir": 900,
            },
            {
                "source_file": "multi-account.txt",
                "Account": "B",
                "Saldo Awal": 2000,
                "Saldo Akhir": 2100,
            },
        ])

        issues = validate_bank_transactions(transactions, summaries)

        self.assertNotIn("reconciliation_balance", set(issues["rule_name"]))

    def test_parser_warning_from_summary_is_exposed(self):
        transactions = pd.DataFrame([
            {
                "source_file": "warning.txt",
                "Account": "A",
                "Tanggal": "02/01/24",
                "Keterangan": "DEBIT",
                "DB": 100,
                "CR": pd.NA,
                "Saldo": 900,
            }
        ])
        summaries = pd.DataFrame([
            {
                "source_file": "warning.txt",
                "Account": "A",
                "Saldo Awal": 1000,
                "Saldo Akhir": 900,
                "Validation Issues": "Potential overlap conflict.",
            }
        ])

        issues = validate_bank_transactions(transactions, summaries)

        parser_issues = issues[
            issues["rule_name"] == "dki_txt_parser_diagnostics"
        ]
        self.assertEqual(len(parser_issues), 1)
        self.assertEqual(
            parser_issues.iloc[0]["message"],
            "Potential overlap conflict.",
        )

    def test_parser_version_invalidates_old_cache(self):
        self.assertEqual(DKI_TXT_PARSER_VERSION, "2.0.0")


if __name__ == "__main__":
    unittest.main()
