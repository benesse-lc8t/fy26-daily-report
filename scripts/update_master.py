#!/usr/bin/env python3
"""
Update data.json master.products from a new 銷貨單品號主檔 Excel file.

Usage:
    python3 scripts/update_master.py <path-to-excel>
"""

import json, sys, openpyxl

DATA_JSON = '/home/user/fy26-daily-report/data.json'
SHEET_NAME = '銷貨單品號主檔'

# 主檔缺漏／需修正的品名（人工維護，每次更新自動補回）
NAME_OVERRIDES = {
    '20203C601': {'name': '動物百科探索組', 'bizMajor': '周邊商品', 'bizMinor': '學習周邊', 'category': ''},
}

def main(excel_path):
    print(f'Loading {excel_path} …')
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    if SHEET_NAME not in wb.sheetnames:
        print(f'ERROR: sheet "{SHEET_NAME}" not found. Sheets: {wb.sheetnames}')
        sys.exit(1)

    ws = wb[SHEET_NAME]
    rows = list(ws.iter_rows(values_only=True))
    # header row 0: 製品編號, 商品簡稱, 事業大類, 事業中類, 類別, ...
    products = {}
    for row in rows[1:]:
        if not row or not row[0]: continue
        code = str(row[0]).strip()
        if not code: continue
        name      = (row[1] or '').strip() if row[1] else ''
        biz_major = (row[2] or '').strip() if row[2] else ''
        biz_minor = (row[3] or '').strip() if row[3] else ''
        category  = (row[4] or '').strip() if row[4] else ''
        products[code] = {
            'name': name, 'bizMajor': biz_major,
            'bizMinor': biz_minor, 'category': category,
        }

    for code, info in NAME_OVERRIDES.items():
        products[code] = info
    print(f'Parsed {len(products)} products (含 {len(NAME_OVERRIDES)} 筆人工補品名)')

    with open(DATA_JSON) as f:
        data = json.load(f)
    old_count = len(data['master']['products'])
    data['master']['products'] = products

    with open(DATA_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))

    print(f'master.products updated: {old_count} -> {len(products)}')

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python3 scripts/update_master.py <excel-path>')
        sys.exit(1)
    main(sys.argv[1])
