#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把「FY26 期中目標」Excel 寫入 data.json 的 midTarget 鍵。

用法：
    python3 scripts/update_midterm.py <期中目標.xlsx> [--version 2607]

只寫 data.json 的 `midTarget`，其餘鍵完全不動；
`update_data.py` 亦只覆寫 actualRows / actuals / holidayInfo / generatedAt，
兩支腳本互不干擾，每日更新不會洗掉期中目標。

── 通路對應（已用 4~7 月實績逐月驗證，誤差 < 200 元）──────────────
    WEB   ← shopline-WEB ＋ shopline-生活館
    EC    ← 蝦皮-生活館 ＋ 蝦皮-學習館
    TM    ← CC ＋ 0800
    KOL   ← KOL
    經代銷 ← 經代銷
  舞台劇商售／樂園商售／其他管道 **不納入** ——
  儀表板的實績本來就不含這三個通路，納入目標會讓達成率結構性失真。

── 注意 ────────────────────────────────────────────────
  期中目標檔會把「已結算月份」直接填入實績（2607 版為 4~7 月）。
  這些月份的「對期中目標達成率」必然是 100.0%，屬期中計畫的本質而非錯誤。
  腳本會自動偵測並記錄在 midTarget.actualMonths，供前端加註。
"""
import json
import os
import re
import sys
from datetime import datetime, timezone

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_JSON = os.path.join(ROOT, 'data.json')

MONTHS = ['4月', '5月', '6月', '7月', '8月', '9月',
          '10月', '11月', '12月', '1月', '2月', '3月']

# 儀表板事業別（事業中項 → data.json 的 biz）
BIZ_KEEP = {'學習商品', '學習周邊', '生活周邊', 'Mirafeel', '數位典藏'}
# 需略過的彙總列
BIZ_SKIP = {'周邊', '學習+周邊+Mirafeel', '事業中項', None, ''}

# 儀表板通路 ← 期中目標檔通路
CH_MAP = {
    'WEB':   ['shopline-WEB', 'shopline-生活館'],
    'EC':    ['蝦皮-生活館', '蝦皮-學習館'],
    'TM':    ['CC', '0800'],
    'KOL':   ['KOL'],
    '經代銷': ['經代銷'],
}
SRC_TO_CH = {src: ch for ch, srcs in CH_MAP.items() for src in srcs}
# 刻意不納入（儀表板實績不含這些通路）
CH_EXCLUDED = ['舞台劇商售', '樂園商售', '其他管道']

SHEET_AMOUNT = '事業+管道商售金額(未稅)'
SHEET_QTY = '事業+管道套數'


def find_header_row(ws, max_scan=30):
    """找出含『事業 / 管道 / 4月』的標題列，回傳 (列號, 事業欄, 管道欄, 4月欄)。"""
    for r, row in enumerate(ws.iter_rows(max_row=max_scan, max_col=25, values_only=True), 1):
        vals = [str(v).strip() if v is not None else '' for v in row]
        if '管道' in vals and '4月' in vals:
            biz_i = vals.index('事業中項') if '事業中項' in vals else None
            ch_i = vals.index('管道')
            m_i = vals.index('4月')
            if biz_i is not None:
                return r, biz_i, ch_i, m_i
    raise SystemExit(f'找不到標題列（需含「事業中項」「管道」「4月」）：{ws.title}')


def parse_sheet(wb, sheet_name):
    """回傳 ({biz: {ch: {month: value}}}, 略過統計)。"""
    if sheet_name not in wb.sheetnames:
        raise SystemExit(f'找不到分頁「{sheet_name}」')
    ws = wb[sheet_name]
    hr, bi, ci, mi = find_header_row(ws)

    out, skipped, seen_src = {}, {}, set()
    for row in ws.iter_rows(min_row=hr + 1, max_row=ws.max_row,
                            max_col=mi + 12, values_only=True):
        biz = (str(row[bi]).strip() if row[bi] is not None else '')
        src = (str(row[ci]).strip() if row[ci] is not None else '')
        if not biz or not src:
            continue
        if biz in BIZ_SKIP or biz not in BIZ_KEEP:
            continue
        seen_src.add(src)
        if src in ('全體', '管道'):          # 小計列
            continue
        ch = SRC_TO_CH.get(src)
        if ch is None:
            skipped[src] = skipped.get(src, 0) + 1
            continue
        bucket = out.setdefault(biz, {}).setdefault(ch, {})
        for k, m in enumerate(MONTHS):
            v = row[mi + k]
            if isinstance(v, (int, float)):
                bucket[m] = round(bucket.get(m, 0) + float(v), 2)
            else:
                bucket.setdefault(m, 0)
    return out, skipped, seen_src


def detect_actual_months(mid_amount, data):
    """偵測哪些月份的期中目標＝實績（誤差 < 200 元視為相同）。

    只比對期中目標涵蓋的事業別與通路 —— 實績裡可能出現期中檔沒有的事業
    （例：5、6 月的「舞台劇」），直接比總額會誤判。
    """
    actuals = data.get('actuals') or {}
    chs_keep = set(CH_MAP)
    hit = []
    for m in MONTHS:
        md = actuals.get(m)
        if not md:
            continue
        a = sum(v for b, chs in md['amount'].items() if b in BIZ_KEEP
                for c, v in chs.items() if c in chs_keep)
        t = sum(chs.get(m, 0) for biz in mid_amount.values() for chs in biz.values())
        if a > 0 and abs(a - t) < 200:
            hit.append(m)
    return hit


def total_of(d):
    return sum(v for biz in d.values() for ch in biz.values() for v in ch.values())


def main(path, version=None):
    print(f'Loading {path} …')
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)

    amount, sk_a, src_a = parse_sheet(wb, SHEET_AMOUNT)
    qty, sk_q, _ = parse_sheet(wb, SHEET_QTY)

    print(f'\n解析結果：')
    for label, d in (('金額', amount), ('套數', qty)):
        n = sum(len(chs) for chs in d.values())
        print(f'  {label}：{len(d)} 個事業別 × 合計 {n} 組 (事業×通路)')
    print(f'  金額全年合計：{total_of(amount)/10000:,.0f} 萬')
    print(f'  套數全年合計：{total_of(qty):,.0f} 套')

    unmapped = sorted(set(sk_a) | set(sk_q))
    if unmapped:
        print(f'\n  未納入的通路（刻意排除，儀表板實績亦不含）：{", ".join(unmapped)}')
    unexpected = [s for s in unmapped if s not in CH_EXCLUDED]
    if unexpected:
        print(f'  *** 警告：出現預期外的通路 {unexpected}，請確認對應規則是否需更新 ***')

    with open(DATA_JSON) as f:
        data = json.load(f)

    actual_months = detect_actual_months(amount, data)
    if actual_months:
        print(f'\n  已偵測到「期中目標＝實績」的月份：{"、".join(actual_months)}')
        print(f'  （這些月份的對期中目標達成率必然是 100.0%）')

    if version is None:
        m = re.search(r'(26\d{2})', os.path.basename(path))
        version = m.group(1) if m else 'unknown'

    data['midTarget'] = {
        'version': version,
        'sourceFile': os.path.basename(path),
        'generatedAt': datetime.now(timezone.utc).isoformat(),
        'channelMap': CH_MAP,
        'excludedChannels': CH_EXCLUDED,
        'actualMonths': actual_months,
        'amount': amount,
        'qty': qty,
    }

    with open(DATA_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    size = len(json.dumps(data, ensure_ascii=False)) // 1024
    print(f'\ndata.json updated ({size} KB) — midTarget 版本 {version}')
    print('（actualRows / actuals / holidayInfo / generatedAt 完全未動）')


if __name__ == '__main__':
    args = [a for a in sys.argv[1:]]
    ver = None
    if '--version' in args:
        i = args.index('--version')
        ver = args[i + 1]
        del args[i:i + 2]
    if not args:
        print('Usage: python3 scripts/update_midterm.py <期中目標.xlsx> [--version 2607]')
        sys.exit(1)
    main(args[0], ver)
