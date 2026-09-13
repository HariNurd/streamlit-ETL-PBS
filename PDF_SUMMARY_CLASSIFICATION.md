# PDF summary classification: meaning and limitations

The PDF converters extract transactions and then classify some amounts into the Summary sheet's **Adm**, **Pajak**, **Bunga**, **Saldo Min**, and **JaGir** columns.

These categories are **best-effort classifications based on transaction descriptions**, not verified accounting classifications. Keeping the current rules is reasonable when these columns are used for review and convenience.

## Why classification can be ambiguous

A description containing "admin", "fee", or "interest" does not always establish what the payment represents or who charged it. The words may describe a bank charge, a service charge associated with a transfer, or part of a payment reference supplied by the sender.

For example, Mandiri can append `Transfer Fee` to an `MCM InhouseTrf` description even when the row contains the principal transfer amount. Counting the entire row as an administration fee would be incorrect. The current rules exclude this known case, but other ambiguous descriptions may still occur.

Some statements also contain transactions with blank descriptions. A small debit is not enough evidence to label it an administration fee, and a small credit is not enough evidence to label it interest.

## Current classification rules

The PDF classification logic is implemented in `summary_metrics()` in `services/pdf_statement_adapter.py`.

| Summary column | Evidence used by the current rules |
| --- | --- |
| Adm | Debit descriptions containing ADM, admin, provisi, or fee, or starting with Biaya; excludes the known Mandiri principal-transfer case described above. |
| Pajak | Debit descriptions containing PPH, Pajak, or Tax. |
| Bunga | Explicit debit interest charges, or credited interest when the account is not explicitly identified as Giro. |
| Saldo Min | Debit descriptions explicitly mentioning Saldo Min. |
| JaGir | Credits explicitly mentioning Jasa Giro or JaGir, or interest credits on an explicitly identified Giro account. |

Each transaction is assigned to at most one of these categories. Transactions without matching evidence remain unclassified, and categories without matches remain blank. A blank category does not prove that no such charges occurred.

The rules do not establish whether a fee was charged by the bank, another service provider, or a transaction counterparty. They can both include amounts that do not belong in a category and miss amounts that do.

## What classification does and does not change

Classification only determines which amounts are included in the additional Summary categories. It does not change the extracted transaction amounts, debit/credit directions, or balances. It also does not change the overall debit and credit totals.

The category amounts are already included in the transaction totals. **Do not add Adm, Pajak, Bunga, Saldo Min, or JaGir to those totals again.** Doing so would double-count the amounts.

Numeric reconciliation is a separate check. A statement can reconcile correctly while a description-based category is inaccurate. Reconciliation checks the arithmetic; it does not establish the economic purpose of a transaction.

## How to use the summary

Use the category columns as a starting point for review. Check the underlying descriptions and supporting records before treating the values as verified bank charges, interest income, tax, or accounting entries.

For the current PDF-to-Excel workflow, the existing rules can remain in place with these limitations documented. Add targeted rules when a recurring, clearly understood description justifies them, rather than assuming every new use of "fee" or "admin" has the same meaning.

This note documents the PDF summary behavior. It does not change the parser or the TXT workflow.
