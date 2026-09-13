"""Local comparison; writes aggregate evidence, never modifies source PDFs."""
import contextlib
import io
import json
import re
import sys
import time
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import pandas as pd
from pypdf import PdfReader
from generate_labels import extract_statement
from convert_mutasi_bca import process_pdf
from validators.bank_statement_validator import validate_bank_transactions


def money(v):
    return None if v is None or pd.isna(v) else format(Decimal(str(v)), '.2f')


def main():
    results = []
    for path in sorted((ROOT / 'pdf').glob('*.pdf')):
        result = {'file': path.name}
        reader = PdfReader(path)
        result['pages'] = len(reader.pages)
        text = '\n'.join(p.extract_text(extraction_mode='layout') for p in reader.pages)
        result['account_type'] = next((s for s in ['REKENING TAHAPAN', 'REKENING GIRO'] if s in text), 'other')
        # Independent printed recap extraction, not either parser's output.
        summary = {}
        for label in ['SALDO AWAL', 'MUTASI CR', 'MUTASI DB', 'SALDO AKHIR']:
            matches = re.findall(label + r'\s*:\s*([\d,]+\.\d{2})(?:[^\S\n]+(\d+))?', text)
            if len(matches) == 1:
                amount, count = matches[0]
                summary[label] = (Decimal(amount.replace(',', '')), int(count) if count else None)
        outputs = {}
        for name in ['labels', 'app']:
            start = time.perf_counter()
            log = io.StringIO()
            try:
                with contextlib.redirect_stdout(log):
                    if name == 'labels':
                        parsed, _ = extract_statement(path)
                        rows = parsed['gt_parse']['transactions']
                    else:
                        df, recap = process_pdf(path)
                        rows = [{'date': r['Tanggal'], 'description': r['Keterangan'],
                                 'debit': money(r['DB']), 'credit': money(r['CR']),
                                 'balance': money(r['Saldo'])} for r in df.to_dict('records')]
                        df['source_file'] = path.name
                        recap['source_file'] = path.name
                        issues = validate_bank_transactions(df, recap)
                        result['app_validator_issues'] = issues['rule_name'].value_counts().to_dict()
                        result['app_summary_rows'] = len(recap)
                checks = {}
                for side, label in [('debit', 'MUTASI DB'), ('credit', 'MUTASI CR')]:
                    values = [Decimal(r[side]) for r in rows if r[side] is not None]
                    checks[side + '_count_matches'] = len(values) == summary.get(label, (None,None))[1]
                    checks[side + '_total_matches'] = sum(values, Decimal(0)) == summary.get(label, (None,None))[0]
                running = summary.get('SALDO AWAL', (None,None))[0]
                bad_balances = []
                if running is not None:
                    for i, row in enumerate(rows, 1):
                        running += Decimal(row['credit'] or '0') - Decimal(row['debit'] or '0')
                        if row['balance'] is not None and running != Decimal(row['balance']):
                            bad_balances.append(i)
                    checks['closing_balance_matches'] = running == summary.get('SALDO AKHIR',(None,None))[0]
                checks['printed_balance_mismatches'] = bad_balances
                result[name] = {'rows':len(rows), 'checks':checks}
                outputs[name] = rows
            except Exception as exc:
                result[name] = {'error':str(exc)}
            result[name]['seconds'] = round(time.perf_counter()-start, 3)
            result[name]['warnings'] = [s for s in log.getvalue().splitlines() if 'PERINGATAN' in s]
        if len(outputs) == 2:
            a,b = outputs['labels'],outputs['app']
            diffs = {key: [i for i,(x,y) in enumerate(zip(a,b),1) if x[key]!=y[key]]
                     for key in ['date','debit','credit','balance','description']}
            result['row_differences'] = diffs
            result['description_examples'] = [{'row':i,'labels':a[i-1]['description'],'app':b[i-1]['description']}
                                               for i in diffs['description'][:3]]
        results.append(result)
        out = ROOT / 'outputs' / 'parser_comparison'
        out.mkdir(parents=True,exist_ok=True)
        (out / 'results.json').write_text(json.dumps(results,indent=2),encoding='utf-8')
        print(path.name, {k: {a:b for a,b in result[k].items() if a!='warnings'} for k in ['labels','app']}, flush=True)
    print('Saved outputs/parser_comparison/results.json',flush=True)


if __name__ == '__main__':
    main()
