#!/usr/bin/env python3
"""
Update data.json with KOL檔期規劃表 (kolSchedule) and refresh master.products
from a combined Excel file that contains both:
  - sheet '銷貨單品號主檔'
  - sheet 'KOL檔期規劃表'

Usage:
    python3 scripts/update_kol.py <path-to-excel>

Notes:
  * kolSchedule entries: {ym, kol, biz, start (ISO), end (ISO)}
  * 賊賊 has blank 事業別 in source → forced to 'Mirafeel' (per business owner).
  * A KOL whose 檔期 spans a month boundary can appear twice (once per 年月);
    such duplicates (same kol+start+end+biz) are de-duplicated to one entry.
  * Attribution of sales to each campaign is computed live in the browser
    from actualRows (ch=KOL, matching biz, date within [start,end]).
"""

import json, re, sys, openpyxl
from datetime import datetime, timezone
from update_master import build_products

import os
DATA_JSON = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data.json')
KOL_SHEET   = 'KOL檔期規劃表'
MASTER_SHEET = '銷貨單品號主檔'

# 事業別が空欄の KOL を補完（業務指示）
BIZ_OVERRIDE = {'賊賊': 'Mirafeel'}

DATE_SPLIT_RE = re.compile(r'\s*[~～]\s*')

def parse_date(s):
    s = str(s).strip()
    for fmt in ('%Y/%m/%d', '%Y-%m-%d'):
        try: return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except: pass
    return None

def parse_schedule(ws):
    rows = list(ws.iter_rows(values_only=True))
    seen = set()
    out = []
    skipped = []
    for row in rows[1:]:  # skip header
        if not row or row[0] is None: continue
        ym  = str(row[0]).strip()
        kol = (str(row[1]).strip() if row[1] else '')
        rng = (str(row[2]).strip() if row[2] else '')
        biz = (str(row[3]).strip() if row[3] else '')
        if not kol or not rng: continue
        if not biz:
            biz = BIZ_OVERRIDE.get(kol, '')
        parts = DATE_SPLIT_RE.split(rng)
        if len(parts) != 2:
            skipped.append((ym, kol, rng, '檔期格式非「起~迄」'))
            continue
        start, end = parse_date(parts[0]), parse_date(parts[1])
        if not start or not end:
            bad = parts[0] if not start else parts[1]
            skipped.append((ym, kol, rng, f'日期無法解析：{bad!r}'))
            continue
        key = (kol, start, end, biz)
        if key in seen: continue   # 跨月重複去重
        seen.add(key)
        out.append({'ym': ym, 'kol': kol, 'biz': biz, 'start': start, 'end': end})
    out.sort(key=lambda e: (e['start'], e['biz']))
    return out, skipped

def parse_master(ws):
    rows = list(ws.iter_rows(values_only=True))
    products = {}
    for row in rows[1:]:
        if not row or not row[0]: continue
        code = str(row[0]).strip()
        if not code: continue
        products[code] = {
            'name':      (row[1] or '').strip() if row[1] else '',
            'bizMajor':  (row[2] or '').strip() if row[2] else '',
            'bizMinor':  (row[3] or '').strip() if row[3] else '',
            'category':  (row[4] or '').strip() if len(row) > 4 and row[4] else '',
        }
    return products

def main(excel_path):
    print(f'Loading {excel_path} …')
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)

    if KOL_SHEET not in wb.sheetnames:
        print(f'ERROR: sheet "{KOL_SHEET}" not found. Sheets: {wb.sheetnames}')
        sys.exit(1)

    schedule, skipped = parse_schedule(wb[KOL_SHEET])
    print(f'Parsed {len(schedule)} KOL campaigns:')
    for e in schedule:
        print(f"  {e['start']}~{e['end']}  {e['biz']:<10} {e['kol']}")
    if skipped:
        # 靜默丟棄會讓來源表的打字錯誤變成「檔期憑空消失」，故一律顯示
        print(f'  ⚠ 已略過 {len(skipped)} 列（來源資料有誤，請修正後重跑）：')
        for ym, kol, rng, why in skipped:
            print(f'      {ym} {kol}「{rng}」— {why}')
    no_biz = [e for e in schedule if not e['biz']]
    if no_biz:
        print(f'  ⚠ 事業別空白 {len(no_biz)} 筆（無法歸因業績）：'
              + '、'.join(e['kol'] for e in no_biz))

    with open(DATA_JSON) as f:
        data = json.load(f)

    data['kolSchedule'] = schedule

    if MASTER_SHEET in wb.sheetnames:
        # 與 update_master.py 共用同一套規則（只留有實績製編／舊品名保留／
        # NAME_OVERRIDES）。先前這裡直接塞入全品目錄，會使 data.json 肥大
        # 並洗掉人工補的品名。
        data['master']['products'] = build_products(parse_master(wb[MASTER_SHEET]), data)

    data['generatedAt'] = datetime.now(timezone.utc).isoformat()
    with open(DATA_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    print(f'data.json updated ({len(json.dumps(data, ensure_ascii=False))//1024} KB)')

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python3 scripts/update_kol.py <excel-path>')
        sys.exit(1)
    main(sys.argv[1])
