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

import os
DATA_JSON = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data.json')

BIZ_NAME_MAP = {'學習': '學習商品', '數位經典': '數位典藏'}
SKIP_BIZ = {'', '0', '事業', '事業別', '總計', '(空白)', '(全部)'}
MONTHS = ['4月','5月','6月','7月','8月','9月','10月','11月','12月','1月','2月','3月']

SHEET_RE      = re.compile(r'^(\d{2})(\d{2})\s*(WEB&KOL|EC|TM|經代銷)$')
SHEET_RE_DIGI = re.compile(r'^(\d{2})(\d{2})WEB數位典藏$')

SHEET_CONFIG = {
    'WEB&KOL': {'header_row':2, 'biz':0,'code':1,'date':2,'channel':3,'qty':4,'amt':5,
                'channel_map':{'網站':'WEB','團購':'KOL'}},
    'EC':      {'header_row':2, 'biz':0,'code':1,'date':2,'qty':3,'amt':4, 'fixed':'EC'},
    'TM':      {'header_row':2, 'biz':0,'code':1,'date':2,'qty':4,'amt':6, 'fixed':'TM'},
    '經代銷':   {'header_row':2, 'biz':0,'code':1,'date':2,'qty':4,'amt':5, 'fixed':'經代銷'},
}

# 國定假日表：欄位 [年度, 日期, 假日名稱, 星期]
HOLIDAY_SHEET = '國定假日表'
FY26_START = '2026-04-01'
FY26_END   = '2027-03-31'
WEEKDAY_CH = ['一', '二', '三', '四', '五', '六', '日']  # Mon..Sun

# 停班停課（颱風假等，非國定假日；手動維護，每次更新自動併入）
# refs：各日期各自對應之佐證新聞連結，供事後查證停班停課正確性（一日一則）
EXTRA_SUSPENSIONS = [
    {'date': '2026-07-10', 'name': '巴威颱風・台北市停班停課', 'category': '停班停課',
     'refs': ['https://news.pts.org.tw/article/816893']},
    {'date': '2026-07-11', 'name': '巴威颱風・台北市停班停課', 'category': '停班停課',
     'refs': ['https://news.pts.org.tw/article/817086']},
]

# ══════════════════════════════════════════════════════════════════
# 預購（尚未出貨）→ 不計入實績
# ------------------------------------------------------------------
# 商品停售後開放預購，但要等入庫才出貨。來源 Excel 會把預購單掛在
# 「訂單日」，若直接計入實績等於把尚未發生的收入提前認列。
# 這裡在解析階段就把符合規則的列分流到 preorderRows，
# 不進 actualRows / actuals，但完整保留供「預購狀況」分頁追蹤。
#
# ⚠ 出貨後的認列：來源 Excel **不會**以出貨日重新開列（已向業務端確認），
#   訂單列會一直掛在訂單日。因此出貨完成後必須手動把
#   `recognizeMonth` 設為認列月份（例：'11月' 或 '12月'），
#   屆時這些列會改以該月計入實績。在那之前維持 None＝不認列。
PREORDER_RULES = [
    {
        'id': 'dianduo-tujian',
        'title': '巧虎雙語點讀圖鑑',
        'codes': ['20250FE00', '20250FXBH', '20250FE03'],
        # ⚠ 通路很重要：各通路把訂單寫進每日 Excel 的「時點」不同。
        #   EC / WEB / KOL：以「訂單日」開列 —— 尚未出貨就會出現在每日明細，
        #                   故訂單日落在預購期內者＝預購，必須排除。
        #   TM           ：以「發貨日」開列 —— 預購單在「未發送」狀態下
        #                   根本不會進到每日明細（已於 2026-10-01 核對確認：
        #                   預購檔有 TM 9/29 4 套，actualRows 中 9/22 之後 0 列）。
        #                   因此 TM 一旦出現在每日明細就代表「已出貨」＝真實營收，
        #                   **絕對不能排除**。11 月底入庫出貨時發貨日極可能落在
        #                   11/20~11/26（在本區間內），納入 TM 會把已出貨的營收誤殺。
        'channels': ['WEB', 'KOL', 'EC', '經代銷'],
        # 分冊（20250FE03，381 元／件）另外開一組通路：使用者 2026-10-02 指示
        # 比照排除，而它唯一一列就是 TM 9/30。分冊單價低、數量零星，
        # 11 月誤殺的金額風險極小，故只對這個製編放行 TM；
        # 套組（FE00／FXBH）的 TM 保護維持不變。
        'channelsByCode': {'20250FE03': ['WEB', 'KOL', 'EC', '經代銷', 'TM']},
        # 已人工確認過、同意排除的 TM 列截止日。晚於這天的 TM 列會觸發警報，
        # 因為 TM 出現在每日明細＝已出貨，代表入庫出貨已經開始。
        'tmAcknowledgedThrough': '2026-09-30',
        'start': '2026-09-29',
        'end':   '2026-11-26',
        'recognizeMonth': None,          # ← 出貨認列後改為 '11月' 或 '12月'
        'timeline': [
            {'date': '2026-09-21', 'label': '陸續於各管道停售'},
            {'date': '2026-09-29', 'label': '開始預購'},
            {'date': '2026-11-26', 'label': '預購結束'},
            {'date': '2026-12-01', 'label': '重新上架'},
        ],
        'note': '「巧虎雙語點讀圖鑑」9/21 陸續於各管道停售，9/29(二)–11/26(四) 進行預購，'
                '12/1(二) 重新上架。預購須待 11 月底入庫後才出貨，故預購金額不計入實績。'
                '分冊（20250FE03）亦比照辦理。',
    },
]


def preorder_match(code, date_str, channel=None):
    """回傳符合的預購規則；不符則 None。

    規則若有指定 channels，只有該些通路才算預購（見 PREORDER_RULES 內註解：
    以發貨日開列的通路，進到每日明細時已出貨，不可排除）。
    """
    if not code or not date_str:
        return None
    for rule in PREORDER_RULES:
        if code not in rule['codes']:
            continue
        if not (rule['start'] <= date_str <= rule['end']):
            continue
        chs = (rule.get('channelsByCode') or {}).get(code) or rule.get('channels')
        if chs and channel not in chs:
            continue
        return rule
    return None


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

def build_holiday_info(wb):
    """國定假日表（FY26 期間）＋停班停課 → [{date, weekday, name, category}] 昇順。"""
    items = {}
    if HOLIDAY_SHEET in wb.sheetnames:
        for row in list(wb[HOLIDAY_SHEET].iter_rows(values_only=True))[1:]:
            if not row or not isinstance(row[1], datetime): continue
            iso = row[1].strftime('%Y-%m-%d')
            if iso < FY26_START or iso > FY26_END: continue
            name = str(row[2]).strip() if row[2] else ''
            items[iso] = {'date': iso, 'weekday': WEEKDAY_CH[row[1].weekday()],
                          'name': name, 'category': '國定假日'}
    for e in EXTRA_SUSPENSIONS:
        d = datetime.strptime(e['date'], '%Y-%m-%d')
        items[e['date']] = {'date': e['date'], 'weekday': WEEKDAY_CH[d.weekday()],
                            'name': e['name'], 'category': e['category'],
                            'refs': e.get('refs', [])}
    return [items[k] for k in sorted(items)]

def main(excel_path):
    print(f'Loading {excel_path} …')
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    actual_rows = []
    preorder_rows = []      # 預購（未出貨）→ 不計入實績

    for sheet_name in wb.sheetnames:
        sname = sheet_name.strip()

        # --- WEB數位典藏 sheets: <YYMM>WEB數位典藏 ---
        md = SHEET_RE_DIGI.match(sname)
        if md:
            mm = md.group(2)
            month_label = f'{int(mm)}月'
            ws = wb[sheet_name]
            all_rows = list(ws.iter_rows(values_only=True))
            # 佔位（假）資料判定：整張分頁無任何有效日期 → 略過
            if not any(parse_date(r[1]) for r in all_rows[1:] if r and len(r) > 1):
                print(f'  {sheet_name}: 略過（無有效日期，判定為佔位假資料）')
                continue
            kept = skip_zero = 0
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
        # 佔位（假）資料判定：整張分頁無任何有效日期 → 略過
        dcol = cfg['date']
        if not any(parse_date(r[dcol]) for r in all_rows[cfg['header_row']+1:] if r and len(r) > dcol):
            print(f'  {sheet_name}: 略過（無有效日期，判定為佔位假資料）')
            continue
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
            dstr = parse_date(row[cfg['date']])
            rec = {
                'month': month_label, 'biz': biz, 'ch': channel,
                'code': code, 'name': '',
                'qty': qty, 'amt': round(amt, 2),
                'date': dstr,
            }
            rule = preorder_match(code, dstr, channel)
            if rule and not rule.get('recognizeMonth'):
                rec['preorderId'] = rule['id']
                preorder_rows.append(rec)          # 預購未出貨 → 不進實績
                continue
            if rule and rule.get('recognizeMonth'):
                rec['preorderId'] = rule['id']     # 已出貨 → 改以認列月計入實績
                rec['orderMonth'] = rec['month']
                rec['month'] = rule['recognizeMonth']
            actual_rows.append(rec)
            kept += 1

        print(f'  {sheet_name}: {kept} rows kept, {skip_zero} zero-amt skipped')

    print(f'Total: {len(actual_rows)} actualRows')
    if preorder_rows:
        pq = sum(r['qty'] for r in preorder_rows)
        pa = sum(r['amt'] for r in preorder_rows)
        print(f'預購（不計入實績）: {len(preorder_rows)} 列 / {pq:.0f} 件 / {pa:,.0f} 元'
              f'  最後日期 {max(r["date"] for r in preorder_rows)}')
        # ⚠ 出貨開始的早期警報：TM 以發貨日開列，TM 列一旦被排除就代表
        #   「已經出貨的營收被擋掉了」。套組（FE00／FXBH）的 TM 不在規則內，
        #   所以這裡只會抓到分冊；但若未來放寬，這道警報就是最後一關。
        #   看到它＝該去把對應規則的 recognizeMonth 設成出貨月份。
        for rule in PREORDER_RULES:
            if rule.get('recognizeMonth'):
                continue
            ack = rule.get('tmAcknowledgedThrough') or ''
            tm = [r for r in preorder_rows
                  if r.get('preorderId') == rule['id'] and r['ch'] == 'TM'
                  and r['date'] > ack]
            if tm:
                print()
                print('  ' + '!' * 68)
                print(f'  !! 規則「{rule["id"]}」排除了 {len(tm)} 列**新的** TM 資料'
                      f'（{sum(r["qty"] for r in tm):.0f} 件 / {sum(r["amt"] for r in tm):,.0f} 元）')
                print('  !! TM 以發貨日開列 ——「出現在每日明細」＝「已經出貨」。')
                print('  !! 若這些是已出貨的營收，請立刻把該規則的 recognizeMonth')
                print(f'  !! 設為出貨月份（目前 None），否則這筆營收會被永久少算。')
                print('  !! 最後日期：' + max(r['date'] for r in tm))
                print('  ' + '!' * 68)

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
    data['preorderRows'] = preorder_rows
    data['preorderRules'] = [
        {k: v for k, v in rule.items() if k != 'codes'} | {'codes': rule['codes']}
        for rule in PREORDER_RULES
    ]
    holiday_info = build_holiday_info(wb)
    data['holidayInfo'] = holiday_info
    print(f'holidayInfo: {len(holiday_info)} 筆（含 {len(EXTRA_SUSPENSIONS)} 停班停課）')
    data['generatedAt'] = datetime.now(timezone.utc).isoformat()
    with open(DATA_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    print(f'data.json updated ({len(json.dumps(data, ensure_ascii=False))//1024} KB)')

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python3 scripts/update_data.py <excel-path>')
        sys.exit(1)
    main(sys.argv[1])
