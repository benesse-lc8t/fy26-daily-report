#!/usr/bin/env python3
"""
Update data.json master.products from a 商品主檔 Excel file.

Robust to format changes:
  * sheet name: tries several known names, else the first sheet whose header
    contains a product-code column.
  * columns: detected by header keywords (not fixed positions), so it works
    for both the old (製品編號/商品簡稱/事業大類/事業中類/類別) and the new
    (商品編號/商品名稱/事業別/商品大類/商品中類) layouts.

To keep data.json lean, only products whose code actually appears in
actualRows (plus manual NAME_OVERRIDES) are kept — the dashboard only ever
looks up names for products that have sales.

Usage:
    python3 scripts/update_master.py <path-to-excel>
"""

import json, sys, openpyxl

import os
DATA_JSON = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data.json')

SHEET_CANDIDATES = ['銷貨單品號主檔', '銷貨單-商品主檔', '商品主檔']

# 各欄位可能的標題（依序比對）
COL_CODE  = ['商品編號', '製品編號', '製品編号', '品號']
COL_NAME  = ['商品名稱', '商品簡稱', '品名']
COL_MAJOR = ['商品大類', '事業大類']
COL_MINOR = ['事業別', '事業中類']
COL_CAT   = ['商品中類', '類別']

# 主檔缺漏／需修正的品名（人工維護，每次更新自動補回）
NAME_OVERRIDES = {
    '20203C601': {'name': '動物百科探索組', 'bizMajor': '周邊商品', 'bizMinor': '學習周邊', 'category': ''},
}

def find_col(header, keywords):
    for i, h in enumerate(header):
        if h and str(h).strip() in keywords:
            return i
    return None

def pick_sheet(wb):
    for name in SHEET_CANDIDATES:
        if name in wb.sheetnames:
            return wb[name]
    # fallback：找 header 含「編號/品號」的分頁
    for name in wb.sheetnames:
        ws = wb[name]
        header = next(ws.iter_rows(values_only=True), None)
        if header and find_col(header, COL_CODE) is not None:
            return ws
    return None

def cell(row, idx):
    if idx is None or idx >= len(row) or row[idx] is None:
        return ''
    return str(row[idx]).strip()

def main(excel_path):
    print(f'Loading {excel_path} …')
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    ws = pick_sheet(wb)
    if ws is None:
        print(f'ERROR: 找不到商品主檔分頁。Sheets: {wb.sheetnames}')
        sys.exit(1)

    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    c_code  = find_col(header, COL_CODE)
    c_name  = find_col(header, COL_NAME)
    c_major = find_col(header, COL_MAJOR)
    c_minor = find_col(header, COL_MINOR)
    c_cat   = find_col(header, COL_CAT)
    if c_code is None or c_name is None:
        print(f'ERROR: 無法辨識製編/品名欄。header={header}')
        sys.exit(1)
    print(f'欄位對應：code={c_code} name={c_name} major={c_major} minor={c_minor} cat={c_cat}')

    full = {}
    for row in rows[1:]:
        if not row: continue
        code = cell(row, c_code)
        if not code: continue
        full[code] = {
            'name': cell(row, c_name),
            'bizMajor': cell(row, c_major),
            'bizMinor': cell(row, c_minor),
            'category': cell(row, c_cat),
        }
    print(f'主檔解析：{len(full)} 筆全品目錄')

    with open(DATA_JSON) as f:
        data = json.load(f)

    # 只保留「有實績」的製編（＋人工補品名），避免 data.json 肥大
    sold_codes = {r['code'] for r in data.get('actualRows', []) if r.get('code')}
    products = {c: full[c] for c in sold_codes if c in full}
    for code, info in NAME_OVERRIDES.items():
        products[code] = info

    old_count = len(data['master']['products'])
    data['master']['products'] = products
    with open(DATA_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    print(f'master.products：{old_count} -> {len(products)} 筆（有實績 {len(sold_codes)} 個製編中命中 {len(products)-len(NAME_OVERRIDES)}）')

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python3 scripts/update_master.py <excel-path>')
        sys.exit(1)
    main(sys.argv[1])
