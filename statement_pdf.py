"""Layout-specific readers for generate_labels. No amount inference or OCR."""
import re
from datetime import datetime
from decimal import Decimal
from itertools import groupby

MONEY = re.compile(r"(?<![\d.,])[+-]?\d[\d,]*\.\d{2}(?!\d)")
INTEGER_MONEY = re.compile(r"[+-]?\d[\d,]*(?:\.\d{2})?")


def compact(text):
    return ' '.join(text.split())


def amount(text):
    return format(Decimal(text.replace(',', '')), '.2f')


def field(pattern, text):
    match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE)
    if not match:
        raise ValueError(f'Missing statement field: {pattern}')
    return compact(match.group(1))


def detect_layout(pages):
    text = '\n'.join(pages)
    if re.search(r'REKENING (TAHAPAN|GIRO)', text) and 'CBG' in text and 'MUTASI' in text:
        return 'BCA'
    if 'LAPORAN TRANSAKSI FINANSIAL' in text and 'Total Transaksi Debet' in text:
        return 'BRI'
    if 'Laporan Mutasi Rekening' in text and 'Rincian Transaksi' in text:
        return 'BNI'
    if 'Posting Date' in text and 'Total Amount Debited' in text:
        return 'Mandiri'
    if 'E-Statement' in text and 'Transaksi Masuk' in text and 'Transaksi Keluar' in text:
        return 'DKI'
    raise ValueError('Unsupported PDF statement layout; manual review required')


def recap_row(text, header, keys, integer=False):
    lines = text.splitlines()
    idx = next((i for i, line in enumerate(lines) if all(s in line for s in header)), None)
    if idx is None:
        raise ValueError('Missing printed statement recap')
    for line in lines[idx + 1:idx + 8]:
        tokens = (INTEGER_MONEY if integer else MONEY).findall(line)
        if len(tokens) == len(keys):
            return dict(zip(keys, (amount(v.lstrip('+-')) for v in tokens)))
    raise ValueError('Ambiguous printed statement recap amounts')


def bri_recap(text):
    """Read the BRI recap even when a page break separates labels and values.

    Bound the search by the recap heading and the following Terbilang section.
    Require exactly one row containing only the four recap amounts, so page
    numbers, timestamps, and transaction amounts cannot become summary values.
    """
    lines = text.splitlines()
    headers = [i for i, line in enumerate(lines) if all(label in line for label in
               ['Saldo Awal', 'Total Transaksi Debet', 'Total Transaksi Kredit', 'Saldo Akhir'])]
    if len(headers) != 1:
        raise ValueError('Missing or ambiguous BRI recap heading')
    start = headers[0] + 1
    end = next((i for i in range(start, len(lines))
                if re.match(r'\s*Terbilang\b', lines[i], re.IGNORECASE)), None)
    if end is None:
        raise ValueError('Missing BRI recap end marker (Terbilang)')
    candidates = []
    for line in lines[start:end]:
        values = MONEY.findall(line)
        if len(values) == 4 and not MONEY.sub('', line).strip():
            candidates.append(values)
    if len(candidates) != 1:
        raise ValueError('Ambiguous printed BRI recap amounts')
    return dict(zip(['opening_balance', 'debit_amount_total', 'credit_amount_total', 'closing_balance'],
                    map(amount, candidates[0])))


def metadata_and_summary(text, bank):
    if bank == 'BRI':
        meta = dict(account_number=field(r'No\. Rekening\s*:\s*(\d+)', text),
                    statement_period=field(r'Periode Transaksi\s*:\s*(\d{2}/\d{2}/\d{2}\s*-\s*\d{2}/\d{2}/\d{2})', text),
                    account_type=field(r'Nama Produk\s*:\s*(.*?)\s{3,}', text),
                    currency=field(r'Valuta\s*:\s*(\w+)', text))
        summary = bri_recap(text)
    elif bank == 'BNI':
        meta = dict(account_number=field(r'TAPLUS.*?-\s*(\d+)', text),
                    account_type='TAPLUS MUDA', currency=field(r'Mata Uang:\s*(\w+)', text),
                    statement_period=field(r'Periode:\s*([^\r\n]+)', text))
        summary = recap_row(text, ['Saldo Awal','Total Pemasukan','Total Pengeluaran','Saldo Akhir'],
                            ['opening_balance','credit_amount_total','debit_amount_total','closing_balance'], True)
    elif bank == 'Mandiri':
        if re.search(r'Account No\.\s+Account Name', text):
            meta = dict(account_number=field(r'Account No\.\s+Account Name[^\n]*\n\s*(\d+)',text),
                        account_type='Account Statement',currency='IDR',
                        statement_period=field(r'Period\s+Currency\s+Branch\s*\n\s*([^\n]+?\d{4})',text))
            summary = {}
            for label,keys in [(['Opening Balance','No. of Debit','Total Amount Debited'],
                               ['opening_balance','debit_transaction_count','debit_amount_total']),
                              (['Closing Balance','No. of Credit','Total Amount Credited'],
                               ['closing_balance','credit_transaction_count','credit_amount_total'])]:
                values=recap_row(text,label,keys,True)
                values[keys[1]]=int(Decimal(values[keys[1]]))
                summary.update(values)
            return dict(bank_name=bank, **meta, summary=summary)
        meta = dict(account_number=field(r'Account No\.?\s+(\d+)', text),
                    account_type='Account Statement', currency=field(r'^Currency\s+(\w+)', text),
                    statement_period=field(r'^Period\s+([^\r\n]+)', text))
        summary = {key: amount(field(label + r'\s+([\d,]+\.\d{2})', text)) for label,key in [
            ('Opening Balance','opening_balance'),('Closing Balance','closing_balance'),
            ('Total Amount Debited','debit_amount_total'),('Total Amount Credited','credit_amount_total')]}
        for side in ['Debit','Credit']:
            summary[side.lower() + '_transaction_count'] = int(field(r'No\.? of ' + side + r'\s+(\d+)', text))
    else:
        meta = dict(account_number=field(r'DKI\s*-\s*(\d+)', text), account_type='Tabungan', currency='IDR',
                    statement_period=field(r'\b((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+20\d{2})', text))
        summary = recap_row(text, ['Saldo Awal','Transaksi Masuk','Transaksi Keluar','Saldo Akhir'],
                            ['opening_balance','credit_amount_total','debit_amount_total','closing_balance'])
    return dict(bank_name=bank, **meta, summary=summary)


def make_transaction(date, description, db, cr, balance, **extra):
    db, cr = Decimal(db), Decimal(cr)
    if db < 0 or cr < 0 or (db != 0 and cr != 0):
        raise ValueError('Ambiguous debit/credit amounts')
    return dict(date=date, description=compact(description), debit=format(db,'.2f') if db else None,
                credit=format(cr,'.2f') if cr else None, balance=amount(balance), **extra)


def read_bri(pages):
    rows = []
    date_re = re.compile(r'^\s*(\d{2}/\d{2}/\d{2} \d{2}:\d{2}:\d{2})\s+')
    for page_no,page in enumerate(pages,1):
        active = False
        desc_start = None
        desc_end = None
        for line in page.splitlines():
            if all(s in line for s in ['Tanggal Transaksi','Uraian Transaksi','Teller']):
                active = True
                desc_end = line.index('Teller') - 2
                continue
            if not active:
                continue
            if any(s in line for s in ['Saldo Awal','Created By','IBIZ_', 'Terbilang']):
                active = False
                continue
            if 'Transaction Date' in line or not line.strip():
                continue
            date = date_re.match(line)
            if date:
                values = list(MONEY.finditer(line))
                if len(values) != 3:
                    raise ValueError(f'BRI page {page_no}: expected debit, credit and balance')
                desc_start = date.end()
                # Teller has its own column, separate from the description.
                description = line[date.end():values[0].start()]
                description = re.sub(r'\s{2,}[A-Z0-9]{4,10}\s*$', '', description)
                stamp = datetime.strptime(date.group(1),'%d/%m/%y %H:%M:%S')
                rows.append(make_transaction(stamp.strftime('%d/%m'), description,
                    *(amount(v.group()) for v in values), posting_datetime=stamp.isoformat(), source_page=page_no))
            elif rows and desc_start is not None:
                continuation = compact(line)
                if continuation:
                    rows[-1]['description'] = compact(rows[-1]['description'] + ' ' + continuation)
    return rows


def read_bni(pages):
    rows = []
    date_re = re.compile(r'^\s*(\d{2} [A-Za-z]{3} \d{4})\s+')
    for page_no,page in enumerate(pages,1):
        active = False
        for line in page.splitlines():
            if 'Tanggal & Waktu' in line and 'Nominal (IDR)' in line:
                active = True
                desc_end = line.index('Nominal (IDR)') - 2
                continue
            if not active:
                continue
            if 'PT Bank Negara' in line or 'Informasi Lainnya' in line or 'Saldo Akhir' in line:
                active = False
                continue
            if not line.strip() or 'Saldo Awal' in line:
                continue
            date = date_re.match(line)
            if date:
                values = re.search(r'([+-]?\d[\d,]*(?:\.\d{2})?)\s+([\d,]+(?:\.\d{2})?)\s*$', line)
                if not values:
                    raise ValueError(f'BNI page {page_no}: missing signed amount or balance')
                nominal = Decimal(values.group(1).replace(',',''))
                stamp = datetime.strptime(date.group(1),'%d %b %Y')
                rows.append(make_transaction(stamp.strftime('%d/%m'),line[date.end():values.start()],
                    max(-nominal,0),max(nominal,0),values.group(2), posting_datetime=stamp.isoformat(), source_page=page_no))
            elif rows:
                time = re.match(r'^\s*(\d{2}:\d{2}:\d{2})\s+WIB\s*',line)
                if time:
                    rows[-1]['posting_datetime'] = rows[-1]['posting_datetime'][:11] + time.group(1)
                    line = line[time.end():desc_end]
                else:
                    line = line[:desc_end]
                rows[-1]['description'] = compact(rows[-1]['description']+' '+line)
    return rows


def read_mandiri(pages):
    """Use PDF content order: remarks can begin above the visual date baseline."""
    rows = []
    current = None
    old_date = re.compile(r'^(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:)(\d{2})?$')
    new_date = re.compile(r'^(\d{2} [A-Za-z]{3} \d{4}),$')
    for page_no, page in enumerate(pages, 1):
        for raw in page.splitlines():
            line = raw.strip()
            date = old_date.match(line) or new_date.match(line)
            if date:
                if current is not None:
                    raise ValueError('Mandiri transaction missing amounts before next date')
                modern = bool(new_date.match(line))
                current = {'stamp': date.group(1) + (' ' if modern else ''),
                           'time': '' if modern else date.group(2) or '',
                           'format': '%d %b %Y %H:%M:%S' if modern else '%d/%m/%Y %H:%M:%S',
                           'description': [], 'page': page_no}
                continue
            if current is None or not line:
                continue
            if not current['time'] and re.fullmatch(r'\d{2}(?::\d{2}:\d{2})?', line):
                current['time'] = line
                continue
            values = re.search(r'(?:^|\s)([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$', line)
            if values:
                current['description'].append(line[:values.start()].rstrip().removesuffix('-'))
                stamp = datetime.strptime(current['stamp'] + current['time'], current['format'])
                rows.append(make_transaction(stamp.strftime('%d/%m'), ' '.join(current['description']),
                            *(amount(v) for v in values.groups()), posting_datetime=stamp.isoformat(),
                            source_page=current['page']))
                current = None
            else:
                current['description'].append(line)
    if current is not None:
        raise ValueError('Mandiri final transaction missing amounts')
    return rows


def read_dki(pages):
    rows=[]
    account=None
    date_re=re.compile(r'^\s*(\d{2} [A-Za-z]{3} \d{2} \d{2}:\d{2})\s+')
    for page_no,page in enumerate(pages,1):
        active=False
        for line in page.splitlines():
            account_match=re.match(r'^\s*TAB(?:UNGAN)?\s+MONAS.*?-\s*(\d+)\s*$',line)
            if account_match:
                account=account_match.group(1);active=False
                continue
            if 'DISCLAIMER' in line:
                active=False
                continue
            date=date_re.match(line)
            if date:
                active=True
                match=re.search(r'\b(DEBIT|KREDIT)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$',line)
                if not match:
                    raise ValueError(f'DKI page {page_no}: missing direction/amount/balance')
                stamp=datetime.strptime(date.group(1),'%d %b %y %H:%M')
                rows.append(make_transaction(stamp.strftime('%d/%m'),line[date.end():match.start()],
                    amount(match.group(2)) if match.group(1)=='DEBIT' else '0',
                    amount(match.group(2)) if match.group(1)=='KREDIT' else '0',match.group(3),
                    posting_datetime=stamp.isoformat(),source_page=page_no,source_row=len(rows)+1,account_number=account))
            elif active and line.strip():
                rows[-1]['description']=compact(rows[-1]['description']+' '+line)
    return rows


def order_same_timestamp(rows, opening):
    """Resolve Jakarta's same-day posting order using unchanged printed balances.

    Interest can print at 01:43 but post after afternoon transactions.
    Source row and timestamp remain available for audit.
    """
    ordered=[]
    balance=Decimal(opening)
    for _,group in groupby(rows,key=lambda r:r['posting_datetime'][:10]):
        pending=list(group)
        trial=balance
        for row in pending:
            trial+=Decimal(row['credit'] or '0')-Decimal(row['debit'] or '0')
            if trial!=Decimal(row['balance']):break
        else:
            ordered.extend(pending);balance=trial
            continue
        while pending:
            candidates=[r for r in pending if balance+Decimal(r['credit'] or '0')-Decimal(r['debit'] or '0')==Decimal(r['balance'])]
            if len(candidates)!=1:
                raise ValueError(f'Cannot reconcile printed transactions on {pending[0]["date"]} from balance {balance}; source statement requires review')
            row=candidates[0]; pending.remove(row); ordered.append(row); balance=Decimal(row['balance'])
    return ordered


def validate_statement(parsed):
    summary=parsed['summary']; rows=parsed['transactions']
    running=Decimal(summary['opening_balance'])
    for idx,row in enumerate(rows,1):
        if row['debit'] is None and row['credit'] is None:
            raise ValueError(f'Transaction {idx}: missing debit/credit')
        running += Decimal(row['credit'] or '0')-Decimal(row['debit'] or '0')
        if row['balance'] is not None and running != Decimal(row['balance']):
            raise ValueError(f'Transaction {idx}: printed balance mismatch ({running} != {row["balance"]})')
    for side in ['debit','credit']:
        values=[Decimal(r[side]) for r in rows if r[side] is not None]
        if sum(values,Decimal(0)) != Decimal(summary[side+'_amount_total']):
            raise ValueError(f'{side} total mismatch: extracted {sum(values,Decimal(0))}, printed {summary[side+"_amount_total"]}')
        key=side+'_transaction_count'
        if key in summary and len(values)!=summary[key]:
            raise ValueError(f'{side} transaction count mismatch')
        summary.setdefault(key,len(values))
    if running != Decimal(summary['closing_balance']):
        raise ValueError('Closing balance mismatch')


def validate_recap(summary):
    expected=Decimal(summary['opening_balance'])+Decimal(summary['credit_amount_total'])-Decimal(summary['debit_amount_total'])
    if expected!=Decimal(summary['closing_balance']):
        raise ValueError(f'Printed recap is inconsistent: opening + credits - debits = {expected}, '
                         f'but printed closing balance is {summary["closing_balance"]}; source statement requires review')


def extract_other_statement(path,pages,bank):
    try:
        parsed=metadata_and_summary('\n'.join(pages),bank)
        if bank!='DKI':validate_recap(parsed['summary'])
        if bank == 'Mandiri':
            from pypdf import PdfReader
            transaction_pages = [p.extract_text() for p in PdfReader(path).pages]
        else:
            transaction_pages = pages
        rows={'BRI':read_bri,'BNI':read_bni,'Mandiri':read_mandiri,'DKI':read_dki}[bank](transaction_pages)
        if bank=='DKI':
            statements=[]
            lines=pages[0].splitlines()
            for i,line in enumerate(lines):
                if not re.match(r'^\s*TAB(?:UNGAN)?\s+MONAS',line):continue
                values=MONEY.findall(line)
                if len(values)!=4:raise ValueError('Ambiguous DKI account recap')
                account=field(r'-\s*(\d+)',lines[i+1])
                part={**parsed,'account_number':account,'account_type':compact(line[:MONEY.search(line).start()]),
                      'summary':dict(zip(['opening_balance','credit_amount_total','debit_amount_total','closing_balance'],map(amount,values)))}
                validate_recap(part['summary'])
                part['transactions']=order_same_timestamp([r for r in rows if r['account_number']==account],part['summary']['opening_balance'])
                validate_statement(part);statements.append(part)
            if not statements or sum(len(s['transactions']) for s in statements)!=len(rows):
                raise ValueError('Unmatched DKI account or transaction')
            if len(statements)>1:
                return {'gt_parse':statements[0],'statements':statements}
            return {'gt_parse':statements[0]}
        parsed['transactions']=rows
        validate_statement(parsed)
        return {'gt_parse':parsed}
    except ValueError as exc:
        raise ValueError(f'{path.name}: {exc}') from exc
