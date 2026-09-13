"""Adapt validated PDF labels to the existing Excel/dataframe interfaces."""
import re
from decimal import Decimal

import pandas as pd

from generate_labels import extract_statement


def summary_metrics(df):
    """Classify explicit descriptions; never infer fees from a small amount.

    Based on the TXT summary categories, with PDF-specific fee descriptions.
    Mandiri appends 'Transfer Fee' to principal transfers; those are not fees.
    Credit interest on Giro accounts is jasa giro; savings interest stays Bunga.
    """
    totals={key:None for key in ['Adm','Pajak','Bunga','Saldo Min','JaGir']}
    giro='GIRO' in df.attrs.get('account_type','').upper()
    for row in df.to_dict('records'):
        text=str(row.get('Keterangan','')).upper()
        db=row.get('DB');cr=row.get('CR')
        category=None;value=None
        if db is not None and not pd.isna(db):
            value=db
            if re.search(r'\bPPH\b|PAJAK|\bTAX\b',text):category='Pajak'
            elif re.search(r'SALDO\s+MIN',text):category='Saldo Min'
            elif re.search(r'BUNGA|INTEREST|\bBNG\b',text):category='Bunga'
            elif re.search(r'\bADM\b|ADMIN|\bPROVISI\b|FEE|^BIAYA(?:\s|$)',text):
                if not ('MCM INHOUSETRF' in text and 'TRANSFER FEE' in text):category='Adm'
        elif cr is not None and not pd.isna(cr):
            value=cr
            if re.search(r'JASA\s*GIRO|JAGIR',text):category='JaGir'
            elif re.search(r'BUNGA|INTEREST|\bBNG\b',text):category='JaGir' if giro else 'Bunga'
        if category:
            totals[category]=(totals[category] or Decimal('0'))+Decimal(str(value))
    return {key:value if value is not None else pd.NA for key,value in totals.items()}


def statement_frames(parsed):
    bank=parsed['bank_name']; summary=parsed['summary']
    rows=[]
    for tx in parsed['transactions']:
        row={'Tanggal':tx['date'],'Keterangan':tx['description'],
             'DB':Decimal(tx['debit']) if tx['debit'] is not None else pd.NA,
             'CR':Decimal(tx['credit']) if tx['credit'] is not None else pd.NA,
             'Saldo':Decimal(tx['balance']) if tx['balance'] is not None else pd.NA}
        if bank=='Mandiri':row['PostingDate']=pd.Timestamp(tx['posting_datetime'])
        if bank=='BNI':
            row['Jam']=tx['posting_datetime'][11:]+' WIB'
            row['Mutasi']=Decimal(tx['credit']) if tx['credit'] is not None else -Decimal(tx['debit'])
        if bank=='DKI':
            row['Account']=parsed['account_number']
            row['TanggalJam']=tx['posting_datetime']
            row['source_pdf_row']=tx['source_row']
        rows.append(row)
    columns=['Tanggal','Keterangan','DB','CR','Saldo']
    if bank=='Mandiri':columns+=['PostingDate']
    if bank=='BNI':columns=['Tanggal','Jam','Keterangan','Mutasi','DB','CR','Saldo']
    if bank=='DKI':columns+=['Account','TanggalJam','source_pdf_row']
    df=pd.DataFrame(rows,columns=columns)
    df.attrs['account_type']=parsed.get('account_type','')
    metrics=summary_metrics(df)
    if bank=='DKI':
        recap=pd.DataFrame([{'Rekening':parsed['account_number'],'Account':parsed['account_number'],
            'Saldo_Awal':Decimal(summary['opening_balance']),'Transaksi_Masuk':Decimal(summary['credit_amount_total']),
            'Transaksi_Keluar':Decimal(summary['debit_amount_total']),'Saldo_Akhir':Decimal(summary['closing_balance']),
            **metrics}])
    else:
        labels=['SALDO AWAL','MUTASI CR','MUTASI DB','SALDO AKHIR'] if bank=='BCA' else ['Saldo Awal','Total Pemasukan','Total Pengeluaran','Saldo Akhir']
        keys=['opening_balance','credit_amount_total','debit_amount_total','closing_balance']
        recap_rows=[]
        for label,key in zip(labels,keys):
            entry={'Keterangan':label,'Amount':Decimal(summary[key])}
            if bank=='BCA':
                entry['Frekuensi']=summary.get(key.replace('amount_total','transaction_count'),pd.NA) if 'amount_total' in key else pd.NA
            recap_rows.append(entry)
        recap_rows.extend({'Keterangan':key,'Amount':value} for key,value in metrics.items())
        recap=pd.DataFrame(recap_rows)
    # Normalized fields let the ETL validator reconcile the bank-specific recap.
    recap['Saldo Awal']=Decimal(summary['opening_balance'])
    recap['Saldo Akhir']=Decimal(summary['closing_balance'])
    return df,recap


def extract_frames(pdf_file,bank):
    label,_=extract_statement(pdf_file,bank=bank)
    statements=label.get('statements',[label['gt_parse']])
    frames=[statement_frames(part) for part in statements]
    if len(frames)==1:return frames[0]
    return tuple(pd.concat([frame[i] for frame in frames],ignore_index=True) for i in [0,1])
