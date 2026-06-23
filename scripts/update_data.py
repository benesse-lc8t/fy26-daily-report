#!/usr/bin/env python3
"""
Update data.json actualRows + actuals from a new FY26 Excel file.

Usage:
    python3 scripts/update_data.py <path-to-excel>

Example:
    python3 scripts/update_data.py '/root/.claude/uploads/xxx/FY26______claude.xlsx'
"""

import json, re, sys, openpyxl
from datetime import datetime, timezone
from collections import defaultdict

DATA_JSON = '/home/user/fy26-daily-report/data.json'

BIZ_NAME_MAP = {'學習': '學習商品', '數位經典': '數位典藏'}
SKIP_BIZ = {'', '0', '事業', '事業別', '總計', '(空白)', '(全部)'}
MONTHS = ['4月','5月','6月','7月','8月','9月','10月','11月','12月','1月','2月','3月']

SHEET_RE      = re.compile(r'^(\d{2})(\d{2})\s*(WEB&KOL|EC|TM)$')
SHEET_RE_DIGI = re.compile(r'^(\d{2})(\d{2})WEB數位典藏$')

SHEET_CONFIG = {
    'WEB&KOL': {'header_row':2, 'biz':0,'code':1,'date':2,'channel':3,'qty':4,'amt':5,
                'channel_map':{'網站':'WEB','團購':'KOL'}},
    'EC':      {'header_row':2, 'biz':0,'code':1,'date':2,'qty':3,'amt':4, 'fixed':'EC'},
    'TM':      {'header_row':2, 'biz':0,'code':1,'date':2,'qty':4,'amt':6, 'fixed':'TM'},
}

def normalize_biz(name):
    t = str(name).strip()
    return BIZ_NAME_MAP.get(t, t)

def num(v):
    if v is None or v == '': return 0.0
    try: return float(v)
    except: return 0.0

def parse_date(v):
    if v is None: return None
    if isinstance(v, datetime): return v.strftime('%Y-%m-%d')
    s = str(v).strip()
    for fmt in ('%Y/%m/%d', '%Y-%m-%d', '%m/%d/%Y'):
        try: return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except: pass
    return None

def main(excel_path):
    print(f'Loading {excel_path} …')
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    actual_rows = []

    for sheet_name in wb.sheetnames:
        sname = sheet_name.strip()

        # --- WEB數位典藏 sheets: <YYMM>WEB數位典藏 ---
        md = SHEET_RE_DIGI.match(sname)
        if md:
            mm = md.group(2)
            month_label = f'{int(mm)}月'
            ws = wb[sheet_name]
            all_rows = list(ws.iter_rows(values_only=True))
            kept = skip_zero = 0
            # header row 0; data from row 1
            # cols: 0=seq, 1=date, 2=起訂年月, 3=管道代碼, 4=方案, 5=單價, 6=qty, 7=訂購金額, 8=銷售額(amt), 9=稅
            for row in all_rows[1:]:
                if not row: continue
                amt = num(row[8])
                if amt == 0: skip_zero += 1; continue
                qty = num(row[6])
                actual_rows.append({
                    'month': month_label, 'biz': '數位典藏', 'ch': 'WEB',
                    'code': '', 'name': '',
                    'qty': qty, 'amt': round(amt, 2),
                    'date': parse_date(row[1]),
                })
                kept += 1
            print(f'  {sheet_name}: {kept} rows kept, {skip_zero} zero-amt skipped')
            continue

        # --- Standard sheets: WEB&KOL / EC / TM ---
        m = SHEET_RE.match(sname)
        if not m: continue
        yy, mm, suffix = m.group(1), m.group(2), m.group(3)
        month_label = f'{int(mm)}月'
        cfg = SHEET_CONFIG[suffix]
        ws = wb[sheet_name]
        all_rows = list(ws.iter_rows(values_only=True))
        kept = skip_zero = 0

        for row in all_rows[cfg['header_row']+1:]:
            if not row: continue
            raw_biz = row[cfg['biz']]
            raw_trim = '' if raw_biz is None else str(raw_biz).strip()
            if raw_trim in SKIP_BIZ: continue
            biz = normalize_biz(raw_trim)

            qty = num(row[cfg['qty']])
            amt = num(row[cfg['amt']])
            if amt == 0: skip_zero += 1; continue

            if 'channel_map' in cfg:
                ch_raw = row[cfg['channel']]
                ch_key = '' if ch_raw is None else str(ch_raw).strip()
                channel = cfg['channel_map'].get(ch_key)
                if not channel: continue
            else:
                channel = cfg['fixed']

            code = '' if row[cfg['code']] is None else str(row[cfg['code']]).strip()
            actual_rows.append({
                'month': month_label, 'biz': biz, 'ch': channel,
                'code': code, 'name': '',
                'qty': qty, 'amt': round(amt, 2),
                'date': parse_date(row[cfg['date']]),
            })
            kept += 1

        print(f'  {sheet_name}: {kept} rows kept, {skip_zero} zero-amt skipped')

    print(f'Total: {len(actual_rows)} actualRows')

    # Build actuals aggregate
    actuals = {m: {'amount': defaultdict(lambda: defaultdict(float)),
                   'qty':    defaultdict(lambda: defaultdict(float))} for m in MONTHS}
    for r in actual_rows:
        if r['month'] not in actuals: continue
        actuals[r['month']]['amount'][r['biz']][r['ch']] += r['amt']
        actuals[r['month']]['qty'][r['biz']][r['ch']] += r['qty']
    actuals_plain = {
        m: {
            'amount': {b: dict(ch) for b, ch in actuals[m]['amount'].items()},
            'qty':    {b: dict(ch) for b, ch in actuals[m]['qty'].items()},
        } for m in MONTHS
    }

    # Print month summary
    for m in MONTHS:
        if actuals_plain[m]['amount']:
            totals = {b: round(sum(v.values())/10000) for b, v in actuals_plain[m]['amount'].items()}
            print(f'  {m}: {totals} 萬')

    # Update data.json
    with open(DATA_JSON) as f:
        data = json.load(f)
    data['actualRows'] = actual_rows
    data['actuals'] = actuals_plain
    data['generatedAt'] = datetime.now(timezone.utc).isoformat()
    with open(DATA_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    print(f'data.json updated ({len(json.dumps(data, ensure_ascii=False))//1024} KB)')

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python3 scripts/update_data.py <excel-path>')
        sys.exit(1)
    main(sys.argv[1])
