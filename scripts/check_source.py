#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_source.py — 送檔前的「來源檔完整性自檢」（在使用者 Windows 筆電執行）

用途
  RefreshAll 有時會跑不完（列數最多的「2610WEB&KOL」樞紐表最常被切掉），
  造成某些分頁停留在前一天。這種缺漏「不會報錯」，數字看起來也合理，
  推上儀表板後會變成「某天業績崩盤」的假象。
  2026-10-19／20／21 連續三天踩到，每次都要重跑、重傳，來回 10〜20 分鐘。
  本腳本在送檔前 5 秒內把這件事擋下來。

用法（擇一）
  1) 雙擊同資料夾的  check_source.bat
  2) 命令列：  python check_source.py
     指定檔案：python check_source.py "D:\\path\\FY26商售實績明細(claude).xlsx"
     指定日期：python check_source.py --date 2026-10-20

預設行為
  · 未指定檔案時，讀 daily_run.local.json 的 master_xlsx
  · 未指定日期時，期望值＝「昨天」（週一則期望上週五，因為週末也有 WEB／EC 的單，
    所以實際上仍用昨天；見 expected_date 說明）
  · 回傳碼 0＝全部通過、1＝有分頁沒跟上、2＝設定或檔案問題

相依套件
  pip install openpyxl
"""

import sys, os, json, datetime, collections, argparse

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(HERE, 'daily_run.local.json')

# 每日實績檔裡「當月」會用到的分頁字首（月份字首在執行時動態組出來）
SHEET_SUFFIXES = ['WEB&KOL', 'EC', 'TM', 'WEB數位典藏']

GREEN = '\033[92m'; RED = '\033[91m'; YEL = '\033[93m'; OFF = '\033[0m'
if os.name == 'nt':
    try:
        import colorama; colorama.just_fix_windows_console()
    except Exception:
        GREEN = RED = YEL = OFF = ''


def parse_date(v):
    """Excel 的日期可能是 datetime、date，或 '2026/10/1'、'2026-10-01 00:00:00' 這類字串。"""
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str):
        p = v.strip()[:10].replace('/', '-').split('-')
        if len(p) == 3:
            try:
                return datetime.date(int(p[0]), int(p[1]), int(p[2]))
            except ValueError:
                return None
    return None


def fy_month_prefix(d):
    """2026-10-20 → '2610'（分頁字首）。"""
    return '%02d%02d' % (d.year % 100, d.month)


def sheet_last_date(ws):
    """回傳 (總列數, 最後日期, {日期: 列數})。找不到日期欄時回 (0, None, {})。"""
    rows = list(ws.iter_rows(values_only=True))
    hdr = None
    for i, r in enumerate(rows[:8]):
        if r and any(isinstance(x, str) and '日期' in x for x in r if x):
            hdr = i
            break
    if hdr is None:
        return 0, None, {}
    j = next(k for k, x in enumerate(rows[hdr]) if isinstance(x, str) and '日期' in x)
    per = collections.Counter()
    for r in rows[hdr + 1:]:
        if j < len(r):
            d = parse_date(r[j])
            if d:
                per[d] += 1
    if not per:
        return 0, None, {}
    return sum(per.values()), max(per), per


def main():
    ap = argparse.ArgumentParser(description='每日實績檔的分頁完整性自檢')
    ap.add_argument('xlsx', nargs='?', help='實績明細檔路徑（省略則讀 daily_run.local.json）')
    ap.add_argument('--date', help='期望的最後日期 YYYY-MM-DD（省略＝昨天）')
    args = ap.parse_args()

    # ── 檔案路徑 ──
    path = args.xlsx
    if not path:
        if not os.path.exists(CONFIG):
            print(f'{RED}找不到設定檔 {CONFIG}{OFF}')
            print('請把檔案路徑當參數傳入，或先建立 daily_run.local.json。')
            return 2
        with open(CONFIG, encoding='utf-8') as f:
            path = json.load(f).get('master_xlsx')
    if not path or not os.path.exists(path):
        print(f'{RED}找不到實績檔：{path}{OFF}')
        return 2

    # ── 期望日期 ──
    if args.date:
        want = parse_date(args.date)
    else:
        want = datetime.date.today() - datetime.timedelta(days=1)
    prefix = fy_month_prefix(want)

    mtime = datetime.datetime.fromtimestamp(os.path.getmtime(path))
    print(f'檔案：{os.path.basename(path)}')
    print(f'存檔時間：{mtime:%Y-%m-%d %H:%M}')
    print(f'期望最後日期：{want}（{prefix} 分頁）')
    print('-' * 66)

    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    found = {}
    for sn in wb.sheetnames:
        s = sn.strip()
        if not s.startswith(prefix):
            continue
        n, mx, per = sheet_last_date(wb[sn])
        found[s] = (n, mx, per)

    ok = True
    missing_sheets = []
    for suf in SHEET_SUFFIXES:
        name = prefix + suf
        hit = next((k for k in found if k.replace(' ', '') == name.replace(' ', '')), None)
        if hit is None:
            print(f'{YEL}？ {name:<22} 找不到這個分頁{OFF}')
            missing_sheets.append(name)
            continue
        n, mx, per = found[hit]
        if mx is None:
            print(f'{RED}✗ {hit:<22} 讀不到日期欄{OFF}')
            ok = False
        elif mx >= want:
            print(f'{GREEN}✓ {hit:<22} 列數 {n:>5}　最後 {mx}　當日 {per.get(want, 0)} 列{OFF}')
        else:
            lag = (want - mx).days
            print(f'{RED}✗ {hit:<22} 列數 {n:>5}　最後 {mx}　落後 {lag} 天 ← 請重跑這個分頁{OFF}')
            ok = False
    wb.close()

    print('-' * 66)
    if ok and not missing_sheets:
        print(f'{GREEN}全部分頁都到 {want}，可以送出。{OFF}')
        return 0
    if missing_sheets and ok:
        print(f'{YEL}有分頁找不到（可能是分頁名稱改了），請人工確認。{OFF}')
        return 1
    print(f'{RED}有分頁沒跟上。請在 Excel 重跑 RefreshAll（或對該分頁的樞紐表按右鍵→重新整理），{OFF}')
    print(f'{RED}存檔後再執行一次本檢查。{OFF}')
    return 1


if __name__ == '__main__':
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(130)
