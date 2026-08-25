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

# 暫定品名（人工維護）：僅在「主檔查無此製編」時才補，主檔一旦收錄即以主檔為準。
# 用於尚未建檔、名稱待確認的新製編 —— 放這裡才不會永久蓋掉日後的正式品名。
NAME_FALLBACKS = {
    # 8/24 首次出現，名稱與事業歸屬待業務確認；暫掛學習周邊
    '20250FF01': {'name': '20250FF01（品名待補）', 'bizMajor': '周邊商品', 'bizMinor': '學習周邊', 'category': ''},
}

# 主檔缺漏／需修正的品名（人工維護，每次更新自動補回）
# 注意：這裡的值會蓋過主檔。名稱未確定者請放 NAME_FALLBACKS。
NAME_OVERRIDES = {
    '20203C601': {'name': '動物百科探索組', 'bizMajor': '周邊商品', 'bizMinor': '學習周邊', 'category': ''},
    '20240FJ00': {'name': '互動遊戲組',     'bizMajor': '周邊商品', 'bizMinor': '學習周邊', 'category': ''},
    '20201C610': {'name': '2010月號寶寶版交通工具拼圖書', 'bizMajor': '周邊商品', 'bizMinor': '學習周邊', 'category': ''},
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

def build_products(full, data):
    """由主檔全品目錄 full 建出 data.json 用的 master.products。

    三道規則（update_master.py / update_kol.py 共用，避免兩邊行為不一致）：
      1. 只保留「有實績」的製編 —— 避免 data.json 肥大
      2. 舊品名保留 —— 主檔改版會移除停售舊品，但這些製編仍有 FY26 實績；
         actualRows 的 name 一律為空、主檔是品名唯一來源，查無即顯示空白品名，
         故沿用既有品名：品名只會被補齊／更新，不會憑空消失
      3. NAME_FALLBACKS —— 暫定品名，僅在主檔查無時補上（不蓋過主檔）
      4. NAME_OVERRIDES —— 需修正的品名，一律蓋過主檔
    """
    sold_codes = {r['code'] for r in data.get('actualRows', []) if r.get('code')}
    prev = data['master']['products']
    products = {c: full[c] for c in sold_codes if c in full}
    carried = {c: prev[c] for c in sold_codes if c not in full and c in prev}
    products.update(carried)
    # 暫定品名：主檔查無才補（不蓋過主檔）
    fallback_used = {c: i for c, i in NAME_FALLBACKS.items() if c not in full}
    products.update(fallback_used)
    # 修正品名：一律蓋過主檔
    for code, info in NAME_OVERRIDES.items():
        products[code] = info

    print(f'master.products：{len(prev)} -> {len(products)} 筆'
          f'（有實績 {len(sold_codes)} 個製編中命中 {len(products)-len(NAME_OVERRIDES)-len(carried)-len(fallback_used)}）')
    if fallback_used:
        print(f'  暫定品名 {len(fallback_used)} 筆（主檔尚未收錄，待正式品名覆蓋）：'
              + '、'.join(f'{c}={i["name"]}' for c, i in sorted(fallback_used.items())))
    if carried:
        print(f'  舊品名保留 {len(carried)} 筆（新主檔已無此製編，沿用既有品名）：'
              + '、'.join(f"{c}={carried[c]['name']}" for c in sorted(carried)[:5])
              + ('…' if len(carried) > 5 else ''))
    missing = sorted(c for c in sold_codes if c not in products)
    if missing:
        print(f'  ⚠ 仍無品名 {len(missing)} 筆：{missing}')
    return products

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

    products = build_products(full, data)
    data['master']['products'] = products
    with open(DATA_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python3 scripts/update_master.py <excel-path>')
        sys.exit(1)
    main(sys.argv[1])
