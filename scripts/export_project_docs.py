"""Render the project change summary and classification Markdown as local PDFs."""
import html
import re
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'outputs' / 'project_documentation'

CHANGE_SUMMARY = """# Project changes: PDF bank statements to Excel

This update improves the PDF extraction layer while keeping Streamlit as the interface and the existing Excel export workflow. It covers BCA, BNI, BRI, Mandiri, and Bank Jakarta/Bank DKI. The TXT parsers and their classification logic were not edited.

## One shared parsing entry point

All five PDF converters now call `generate_labels.extract_statement()` through `services/pdf_statement_adapter.py`. BCA retains its validated Tahapan/Giro reader. The additional bank-specific readers live in `statement_pdf.py` and return structured data that the adapter converts into the existing dataframe formats.

| Bank | Supported layouts and handling |
| --- | --- |
| BCA | Text-based Rekening Tahapan and Giro. Replaces the app's former Tabula extraction path. |
| BNI | Laporan Mutasi Rekening, including the supplied Taplus Muda statements. |
| BRI | Laporan Transaksi Finansial, including summaries split across pages. |
| Mandiri | Laporan Rekening Koran and Kopra Account Statement. Reads transaction text in PDF content order to preserve wrapped descriptions. |
| Bank Jakarta / DKI | JakOne E-Statement, including multiple accounts and same-day transactions printed out of balance order. |

## Validation and data integrity

Amounts use Decimal arithmetic. The readers check printed opening and closing balances, debit/credit totals, transaction counts where available, and printed running balances. Unsupported layouts, missing text layers, incomplete rows, and inconsistent statements produce errors before Excel export.

The new PDF path does not silently repair transaction amounts to force balances to agree. Jakarta transactions can be reordered when their unchanged printed balances establish the sequence; original timestamps and PDF row numbers remain available. Multiple accounts are kept separate in the Jakarta export.

## Summary sheet categories

Both single-file and batch/yearly exports populate Adm, Pajak, Bunga, Saldo Min, and JaGir from explicit descriptions. Giro interest credits are included in JaGir when the account is identified as Giro. Mandiri principal transfers containing the incidental phrase Transfer Fee are excluded from Adm.

Classification remains best-effort: descriptions do not reliably identify who charged a fee or its accounting purpose. Unlabeled transactions remain unclassified. Category amounts are already included in debit/credit totals and must not be added to those totals again. The companion classification PDF explains the rules and limitations in detail.

## BRI page-break bug fix

A later BRI sample exposed a layout case absent from the earlier sample run: recap headings were at the bottom of page 3, while their four amounts were on page 4. The reader's seven-line search missed those amounts and raised an ambiguous recap error.

The BRI reader now searches between the recap heading and the Terbilang section, requiring exactly one row containing the four summary amounts. This supports page breaks without weakening numeric validation. The affected statement exported successfully with 79 reconciled transactions.

## Verification results

These results describe separate test runs; their file counts overlap and should not be added together.

| Verification | Result |
| --- | --- |
| Initial full-folder parsing and Excel export | 125 of 128 PDFs exported; 4,885 transactions reconciled. |
| Files requiring source review in that run | One BNI and two Jakarta samples: their printed recap figures are internally inconsistent. |
| BRI regression after the page-break fix | All 52 available BRI PDFs parsed and validated. |
| Automated regression suite after the fix | 24 tests passed, including the existing TXT tests. |
| Batch export checks | One local sample per bank, including separate exports for the two accounts in a Jakarta PDF. |

The generated Excel workbooks were reopened to check Summary columns and values. The affected BRI page break and the three inconsistent source recaps were also visually inspected. These checks establish consistency on the supplied text-based layouts; they are not exhaustive manual verification of every description or support for scanned PDFs and unseen layouts.

## Files and dependencies

Modified existing files:

- `generate_labels.py`
- `convert_mutasi_bca.py`, `convert_mutasi_bni.py`, `convert_mutasi_BRI.py`, `convert_mutasi_mandiri.py`, `convert_mutasi_dki.py`
- `converters/bank_detector.py`
- `extractors/bank_statement_extractor.py`
- `requirements.txt`

Added parsing, testing, and documentation files:

- `statement_pdf.py`
- `services/pdf_statement_adapter.py`
- `tests/test_pdf_statement_parsers.py`
- `scripts/verify_pdf_parsers.py`
- `PDF_PARSERS.md` and `PDF_SUMMARY_CLASSIFICATION.md`

`scripts/compare_bca_parsers.py` was added for the original comparison. `scripts/export_project_docs.py` generates these documentation PDFs. Neither is required to run the app.

The PDF reader requires pypdf and cryptography, now listed in requirements.txt. PDF cache versions were raised to 2.0.0; BRI was subsequently raised to 2.0.1 for the page-break fix. The TXT parser version and app.py were unchanged.

## Local verification commands

Run the regression suite from the project root:

`env/Scripts/python.exe -m unittest discover -s tests -v`

Validate the local PDFs and generate Excel outputs:

`env/Scripts/python.exe scripts/verify_pdf_parsers.py`

Generated workbooks and reports stay under outputs/. Source statements and generated data are kept local; they are not source-code test fixtures.
"""


def inline(value):
    value = html.escape(value)
    value = re.sub(r'`([^`]+)`', r'<span class="code">\1</span>', value)
    return re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', value)


def markdown_html(source):
    lines = source.splitlines()
    blocks = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        if line.startswith('#'):
            level = len(line) - len(line.lstrip('#'))
            blocks.append(f'<h{level}>{inline(line[level:].strip())}</h{level}>')
            i += 1
        elif line.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
                if not all(re.fullmatch(r':?-+:?', c) for c in cells):
                    tag = 'th' if not rows else 'td'
                    rows.append('<tr>' + ''.join(f'<{tag}>{inline(c)}</{tag}>' for c in cells) + '</tr>')
                i += 1
            blocks.append('<table>' + ''.join(rows) + '</table>')
        elif line.startswith('- '):
            items = []
            while i < len(lines) and lines[i].strip().startswith('- '):
                items.append('<li>' + inline(lines[i].strip()[2:]) + '</li>')
                i += 1
            blocks.append('<ul>' + ''.join(items) + '</ul>')
        else:
            paragraph = []
            while i < len(lines) and lines[i].strip():
                paragraph.append(lines[i].strip())
                i += 1
            blocks.append('<p>' + inline(' '.join(paragraph)) + '</p>')
    return '\n'.join(blocks)


CSS = """
body { font-family: sans-serif; font-size: 10.5pt; line-height: 1.4; color: #253447; }
h1 { font-size: 23pt; line-height: 1.15; color: #153a56; margin: 0 0 18pt; }
h2 { font-size: 13pt; color: #153a56; margin: 16pt 0 7pt; page-break-after: avoid; }
p { margin: 0 0 9pt; page-break-inside: avoid; }
table { border-collapse: collapse; width: 100%; margin: 8pt 0 12pt; font-size: 9.5pt; }
th { background-color: transparent; color: #153a56; text-align: left; }
th, td { border: none; padding: 7pt; vertical-align: top; }
tr { page-break-inside: avoid; }
.code { font-family: monospace; font-size: 9pt; color: #153a56; }
ul { margin-top: 3pt; margin-bottom: 10pt; padding-left: 17pt; page-break-inside: avoid; }
li { margin-bottom: 4pt; }
"""


def render(source, filename, title):
    story = fitz.Story(html=markdown_html(source), user_css=CSS)
    temporary = OUT / ('_' + filename)
    writer = fitz.DocumentWriter(str(temporary))
    page_rect = fitz.paper_rect('a4')
    content_rect = fitz.Rect(48, 62, page_rect.width - 48, page_rect.height - 52)
    more = True
    count = 0
    while more:
        count += 1
        if count > 20:
            raise RuntimeError('Document pagination did not finish')
        device = writer.begin_page(page_rect)
        more, _ = story.place(content_rect)
        story.draw(device)
        writer.end_page()
    writer.close()
    del writer  # Release the Windows file handle before removing the intermediate PDF.
    with fitz.open(temporary) as document:
        document.set_metadata({'title': title, 'subject': 'Streamlit PDF-to-Excel project documentation',
                               'author': 'Project documentation'})
        for i, page in enumerate(document, 1):
            page.insert_text((48, 32), 'PDF TO EXCEL  /  PROJECT DOCUMENTATION', fontsize=8, color=(0.35, 0.43, 0.50))
            page.draw_line((48, page_rect.height - 37), (page_rect.width - 48, page_rect.height - 37),
                           color=(0.78, 0.82, 0.86), width=0.5)
            page.insert_text((48, page_rect.height - 23), title, fontsize=7.5, color=(0.35, 0.43, 0.50))
            page.insert_text((page_rect.width - 90, page_rect.height - 23), f'{i} / {len(document)}', fontsize=8)
        document.save(OUT / filename, garbage=4, deflate=True)
    temporary.unlink()
    print(f'{filename}: {count} pages')


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / 'PROJECT_CHANGES_SUMMARY.md').write_text(CHANGE_SUMMARY, encoding='utf-8')
    render(CHANGE_SUMMARY, 'PROJECT_CHANGES_SUMMARY.pdf', 'Project changes summary')
    render((ROOT / 'PDF_SUMMARY_CLASSIFICATION.md').read_text(encoding='utf-8'),
           'PDF_SUMMARY_CLASSIFICATION.pdf', 'Summary classification: meaning and limitations')


if __name__ == '__main__':
    main()
