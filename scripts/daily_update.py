#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
daily_update.py — 每日更新的機械步驟（使用者在 Excel 跑完 RefreshAll、存檔之後）

  python scripts/daily_update.py prepare   自檢 → 備份 → 兩支 ETL → 六項必查 → validate_html，輸出摘要
  python scripts/daily_update.py publish   commit → push main → 等 GitHub Pages 部署 → 確認線上 data.json
  python scripts/daily_update.py hold      來源檔不完整時：把 data.json 存到 local/hold 分支，main 還原

prepare 不會 commit 也不會 push。回傳碼：
  0 ＝ 六項必查全過、可以 publish
  1 ＝ 有停止條件（印出原因），由人判斷：重跑來源檔就 hold，確認無誤才 publish --force
  2 ＝ 設定、檔案或腳本錯誤
  3 ＝ 資料沒有任何變動（不需發布）

來源檔路徑：命令列 --actual／--ledger，或 scripts/daily_run.local.json 的 master_xlsx／ledger_xlsx。
來源檔一律先複製到暫存資料夾再讀：Excel 或 OneDrive 鎖住原檔時 open() 會被拒，但 CopyFile 讀得到
已存檔的版本（2026-10-22 實測與原檔結果一致）。

摘要（JSON）寫在 %TEMP%/fy26_daily/<時間戳>/summary.json，並更新 %TEMP%/fy26_daily/latest.json。
只放筆數、日期與差異明細，不放推估或缺口：commit 訊息由它產生，而 repo 是公開的。
"""

import os, sys, json, shutil, subprocess, argparse, datetime, tempfile, collections, time, urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
DATA_JSON = os.path.join(REPO, 'data.json')
CONFIG = os.path.join(HERE, 'daily_run.local.json')
WORK_ROOT = os.path.join(tempfile.gettempdir(), 'fy26_daily')
LATEST = os.path.join(WORK_ROOT, 'latest.json')

GH_API = 'https://api.github.com/repos/benesse-lc8t/fy26-daily-report/actions/runs?per_page=10'
LIVE_DATA = 'https://benesse-lc8t.github.io/fy26-daily-report/data.json'

MONTHS = ['4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月']
TOTAL_BIZ = ['學習商品', '學習周邊', '生活周邊', 'Mirafeel']   # 全社口徑（index.html 的 BIZ_LIST_TOTAL）
NODATE_EXPECTED = 11          # 4 月 WEB 數位典藏的真實實績
RETRO_ROW_LIMIT = 20000       # 回溯調整：單一（日期×通路×事業×製編）超過 ±2 萬就停（使用者 2026-10-22 定）
RETRO_DAY_LIMIT = 50000       #           單日淨額超過 ±5 萬就停
UNCHANGED_KEYS = ['version', 'budget', 'master', 'workingDays', 'holidays', 'fy25', 'kolSchedule',
                  'midTarget', 'preorderRules', 'holidayInfo']
TRAILER = 'Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>'


# ───────────────────────── 共用 ─────────────────────────
def env_utf8():
    return dict(os.environ, PYTHONUTF8='1', PYTHONIOENCODING='utf-8')


def run(cmd, cwd=REPO):
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True, encoding='utf-8', errors='replace', env=env_utf8())
    return p.returncode, (p.stdout or '') + (p.stderr or '')


def git(*args):
    return run(['git', *args])


def load_json(p):
    with open(p, encoding='utf-8') as f:
        return json.load(f)


def md(iso):
    return f'{int(iso[5:7])}/{int(iso[8:10])}' if iso else '—'


def lock_safe_copy(src, dst_dir):
    """原檔被 Excel／OneDrive 鎖住時 open() 會 PermissionError，改用 Windows CopyFile。"""
    dst = os.path.join(dst_dir, os.path.basename(src))   # 保留原檔名（台帳腳本會記 sourceFile）
    try:
        shutil.copy2(src, dst)
        return dst, False
    except PermissionError:
        pass
    if os.name == 'nt':
        import ctypes
        if ctypes.windll.kernel32.CopyFileW(src, dst, False):
            return dst, True
        rc, out = run(['powershell', '-NoProfile', '-Command',
                       f"Copy-Item -LiteralPath '{src}' -Destination '{dst}' -Force"])
        if rc == 0 and os.path.exists(dst):
            return dst, True
    raise PermissionError(f'讀不到來源檔（可能正被 Excel 開啟）：{src}')


# ───────────────────────── 六項必查 ─────────────────────────
def last_date(d):
    return max((r['date'] for r in d['actualRows'] if r.get('date')), default='')


def agg(rows, key):
    out = collections.defaultdict(lambda: [0.0, 0.0, 0])
    for r in rows:
        a = out[key(r)]
        a[0] += r['qty']; a[1] += r['amt']; a[2] += 1
    return out


def diff_rows(prev_rows, new_rows, key):
    p, n = agg(prev_rows, key), agg(new_rows, key)
    res = []
    for k in sorted(set(p) | set(n)):
        a, b = p.get(k, [0, 0, 0]), n.get(k, [0, 0, 0])
        if abs(a[1] - b[1]) > 0.005 or a[0] != b[0]:
            res.append({'key': list(k) if isinstance(k, tuple) else k,
                        'qty': [a[0], b[0]], 'amt': [round(a[1], 2), round(b[1], 2)], 'rows': [a[2], b[2]]})
    return res


def month_total(d, m, field='amount'):
    return sum(sum(chs.values()) for chs in d['actuals'].get(m, {}).get(field, {}).values())


def missing_codes(d):
    pr = d['master']['products']
    return sorted({r['code'] for r in d['actualRows'] if r['code'] and r['code'] not in pr})


def six_checks(prev, new):
    stops, items = [], {}
    pl, nl = last_date(prev), last_date(new)
    names = new['master']['products']

    # 1 筆數與最後日期
    items['1_rows'] = {'prev': len(prev['actualRows']), 'new': len(new['actualRows']), 'prevLast': pl, 'newLast': nl}
    if nl < pl:
        stops.append(f'最後日期倒退：{pl} → {nl}')

    # 2 逐日差異（新增日以外＝回溯）
    by_day = diff_rows(prev['actualRows'], new['actualRows'], lambda r: r.get('date') or '(無日期)')
    detail = diff_rows(prev['actualRows'], new['actualRows'],
                       lambda r: (r.get('date') or '', r['ch'], r['biz'], r['code']))
    retro = [x for x in detail if x['key'][0] <= pl]
    for x in retro:
        x['name'] = names.get(x['key'][3], {}).get('name', '')
        if abs(x['amt'][1] - x['amt'][0]) > RETRO_ROW_LIMIT:
            stops.append(f"回溯單筆超過 ±2 萬：{x['key'][0]} {x['key'][1]} {x['key'][3]} {x['amt'][0]:,.0f}→{x['amt'][1]:,.0f}")
    for x in by_day:
        if x['key'] <= pl and abs(x['amt'][1] - x['amt'][0]) > RETRO_DAY_LIMIT:
            stops.append(f"回溯單日淨額超過 ±5 萬：{x['key']} {x['amt'][1] - x['amt'][0]:+,.0f}")
    new_days = []
    for dt in sorted({r['date'] for r in new['actualRows'] if r.get('date') and r['date'] > pl}):
        byc = collections.defaultdict(float); digi = 0.0; qty = 0.0
        for r in new['actualRows']:
            if r.get('date') != dt:
                continue
            if r['biz'] in TOTAL_BIZ:
                byc[r['ch']] += r['amt']; qty += r['qty']
            else:
                digi += r['amt']
        new_days.append({'date': dt, 'total': round(sum(byc.values())), 'qty': qty,
                         'byCh': {k: round(v) for k, v in sorted(byc.items(), key=lambda x: -x[1])}, 'digi': round(digi)})
    items['2_daily'] = {'byDay': by_day, 'retro': retro, 'newDays': new_days}

    # 3 前月彙總（4 月〜上月）
    cur = f"{int(nl[5:7])}月" if nl else None
    prior = MONTHS[:MONTHS.index(cur)] if cur in MONTHS else []
    changed = []
    for m in prior:
        a, b = month_total(prev, m), month_total(new, m)
        qa, qb = month_total(prev, m, 'qty'), month_total(new, m, 'qty')
        if abs(a - b) > 0.5 or qa != qb:
            changed.append({'month': m, 'amt': [round(a), round(b)], 'qty': [qa, qb]})
    items['3_prior'] = {'months': prior, 'changed': changed}
    if changed:
        stops.append('前月彙總有變動：' + '、'.join(c['month'] for c in changed))

    # 4 無日期列
    nd = collections.Counter(f"{r['month']}/{r['ch']}/{r['biz']}" for r in new['actualRows'] if not r.get('date'))
    items['4_nodate'] = {'count': sum(nd.values()), 'detail': dict(nd)}
    if sum(nd.values()) != NODATE_EXPECTED or set(nd) != {'4月/WEB/數位典藏'}:
        stops.append(f'無日期列異常：{dict(nd)}（應為 4月 WEB 數位典藏 11 筆）')

    # 5 缺品名製編
    mp, mn = missing_codes(prev), missing_codes(new)
    fresh = sorted(set(mn) - set(mp))
    items['5_missing'] = {'prev': mp, 'new': mn, 'fresh': fresh}
    if fresh:
        stops.append(f'新的缺品名製編：{fresh}（請使用者提供商品主檔後跑 update_master.py）')

    # 6 kolSchedule 與其他鍵
    changed_keys = [k for k in UNCHANGED_KEYS if prev.get(k) != new.get(k)]
    items['6_keys'] = {'changed': changed_keys}
    if changed_keys:
        stops.append(f'不應變動的鍵有變動：{changed_keys}')

    # 預購（不計入實績）
    pre = diff_rows(prev['preorderRows'], new['preorderRows'], lambda r: (r['date'], r['ch'], r['code']))
    L0, L1 = prev.get('preorderLedger') or {}, new.get('preorderLedger') or {}
    items['preorder'] = {
        'dailyDiff': pre,
        'dailyTotal': [sum(r['qty'] for r in prev['preorderRows']), sum(r['qty'] for r in new['preorderRows'])],
        'ledgerQty': [L0.get('totalQty'), L1.get('totalQty')],
        'ledgerAmt': [L0.get('totalAmt'), L1.get('totalAmt')],
        'ledgerByCh': [L0.get('byChannelQty'), L1.get('byChannelQty')],
        'ledgerLast': (L1.get('rows') or [{}])[-1].get('date'),
    }
    return items, stops


# ───────────────────────── prepare ─────────────────────────
def cmd_prepare(a):
    cfg = load_json(CONFIG) if os.path.exists(CONFIG) else {}
    actual = a.actual or cfg.get('master_xlsx')
    ledger = a.ledger or cfg.get('ledger_xlsx')
    for label, p in (('實績檔', actual), ('台帳', ledger)):
        if not p or not os.path.exists(p):
            print(f'✗ 找不到{label}：{p}'); return 2
    want = a.date or (datetime.date.today() - datetime.timedelta(days=1)).isoformat()

    # git：必須在 main、工作區乾淨、與 origin 同步
    git('fetch', '-q', 'origin')
    _, br = git('rev-parse', '--abbrev-ref', 'HEAD')
    _, dirty = git('status', '--porcelain', '--untracked-files=no')
    if br.strip() != 'main':
        print(f'✗ 目前在分支 {br.strip()}，請先切回 main'); return 2
    if dirty.strip():
        print('✗ 工作區有未提交的變更：\n' + dirty); return 2
    rc, out = git('merge', '--ff-only', 'origin/main')
    if rc:
        print('✗ 無法快轉到 origin/main：\n' + out); return 2

    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    work = os.path.join(WORK_ROOT, stamp); src_dir = os.path.join(work, 'src')
    os.makedirs(src_dir, exist_ok=True)
    s = {'status': None, 'stops': [], 'warnings': [], 'workDir': work, 'expectedDate': want,
         'sources': {}, 'headBefore': git('rev-parse', '--short', 'HEAD')[1].strip()}

    copies = {}
    for key, p in (('actual', actual), ('ledger', ledger)):
        cp, locked = lock_safe_copy(p, src_dir)
        copies[key] = cp
        s['sources'][key] = {'path': p, 'mtime': datetime.datetime.fromtimestamp(os.path.getmtime(p)).isoformat(timespec='minutes'),
                             'lockedCopy': locked}

    # ① 來源檔自檢
    rc, out = run([sys.executable, os.path.join(HERE, 'check_source.py'), copies['actual'], '--date', want])
    s['checkSource'] = {'exit': rc, 'output': out.strip()}
    print(out.strip())
    if rc != 0:
        s['stops'].append(f'check_source 未通過（exit {rc}）：有分頁沒有到 {want}')
        s['status'] = 'stop'
        return finish(s, work)

    # ② 備份 → ③ ETL
    prev_path = os.path.join(work, 'prev.json')
    shutil.copy2(DATA_JSON, prev_path); s['prevJson'] = prev_path
    for key, script in (('updateData', 'update_data.py'), ('updateLedger', 'update_preorder_ledger.py')):
        rc, out = run([sys.executable, os.path.join(HERE, script), copies['actual' if key == 'updateData' else 'ledger']])
        s[key] = {'exit': rc, 'output': out.strip()[-6000:]}
        if rc != 0:
            print(out); s['stops'].append(f'{script} 失敗（exit {rc}）'); s['status'] = 'error'
            return finish(s, work)
        if '!!!!' in out:
            block = [ln.strip() for ln in out.splitlines() if ln.strip().startswith('!!')]
            s['warnings'].extend(block)
            s['stops'].append(f'{script} 印出 !!!! 警報（見 warnings）')

    prev, new = load_json(prev_path), load_json(DATA_JSON)
    s['checks'], stops = six_checks(prev, new)
    s['stops'].extend(stops)

    # ④ validate_html
    rc, out = run([sys.executable, os.path.join(HERE, 'validate_html.py')])
    s['validate'] = {'exit': rc, 'tail': out.strip().splitlines()[-1] if out.strip() else ''}
    if rc != 0:
        s['stops'].append('validate_html.py 未通過'); print(out)

    # ETL 每次都會改寫 generatedAt；除此之外完全相同＝同一份來源檔重跑，還原即可
    def strip_gen(d):
        d = dict(d); d.pop('generatedAt', None)
        if isinstance(d.get('preorderLedger'), dict):
            d['preorderLedger'] = {k: v for k, v in d['preorderLedger'].items() if k != 'generatedAt'}
        return d
    if strip_gen(prev) == strip_gen(new):
        git('checkout', '--', 'data.json')
    _, dirty = git('status', '--porcelain', '--untracked-files=no')
    s['dirty'] = dirty.split('\n') if dirty.strip() else []
    if not s['dirty']:
        s['status'] = 'nochange'
    else:
        s['status'] = 'stop' if s['stops'] else 'ok'
    return finish(s, work)


def finish(s, work):
    path = os.path.join(work, 'summary.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(s, f, ensure_ascii=False, indent=1)
    os.makedirs(WORK_ROOT, exist_ok=True)
    with open(LATEST, 'w', encoding='utf-8') as f:
        json.dump({'summary': path}, f)
    print_summary(s)
    print(f'\n摘要：{path}')
    return {'ok': 0, 'stop': 1, 'error': 2, 'nochange': 3}[s['status']]


def print_summary(s):
    print('\n' + '=' * 66)
    c = s.get('checks')
    if c:
        r = c['1_rows']
        print(f"1 筆數 {r['prev']:,} → {r['new']:,}（{r['new'] - r['prev']:+d}）　最後日期 {r['prevLast']} → {r['newLast']}")
        for d in c['2_daily']['newDays']:
            chs = '／'.join(f'{k} {v / 1e4:.1f}' for k, v in d['byCh'].items())
            print(f"  新增 {d['date']}：全社 {d['total'] / 1e4:.1f} 萬（{chs}）數位典藏 {d['digi'] / 1e4:.2f} 萬")
        retro = c['2_daily']['retro']
        print(f"2 回溯調整 {len(retro)} 筆")
        for x in retro:
            k = x['key']
            print(f"    {k[0]} {k[1]:4s} {k[2]} {k[3]} {x['name'][:20]}  件 {x['qty'][0]:g}→{x['qty'][1]:g}  {x['amt'][0]:,.0f}→{x['amt'][1]:,.0f}")
        print(f"3 前月彙總：{'未變' if not c['3_prior']['changed'] else c['3_prior']['changed']}")
        print(f"4 無日期列：{c['4_nodate']['count']} 筆 {c['4_nodate']['detail']}")
        print(f"5 缺品名製編：{len(c['5_missing']['new'])} 個，新出現 {c['5_missing']['fresh'] or '無'}")
        print(f"6 不應變動的鍵：{'全部未變' if not c['6_keys']['changed'] else c['6_keys']['changed']}")
        p = c['preorder']
        print(f"預購 每日檔排除 {p['dailyTotal'][0]:g} → {p['dailyTotal'][1]:g} 件；台帳 {p['ledgerQty'][0]} → {p['ledgerQty'][1]} 套（最後 {p['ledgerLast']}）")
        for x in p['dailyDiff']:
            print(f"    {x['key'][0]} {x['key'][1]} {x['key'][2]} 件 {x['qty'][0]:g}→{x['qty'][1]:g}")
    if s.get('validate'):
        print(f"validate_html：{s['validate']['tail']}")
    for w in s['warnings']:
        print('  ' + w)
    label = {'ok': '✓ 可以 publish', 'stop': '✗ 停止：需要人工判斷', 'error': '✗ 錯誤', 'nochange': '－ 資料沒有變動'}[s['status']]
    print(f'\n狀態：{label}')
    for x in s['stops']:
        print(f'  · {x}')


# ───────────────────────── publish ─────────────────────────
def latest_summary(path=None):
    if not path:
        path = load_json(LATEST)['summary']
    return load_json(path), path


def commit_message(s):
    c = s['checks']; r = c['1_rows']; p = c['preorder']
    nd = c['2_daily']['newDays']
    span = md(nd[0]['date']) + (f'〜{md(nd[-1]["date"])}' if len(nd) > 1 else '') if nd else md(r['newLast'])
    head = f"data: {span} 實績進帳（{r['prev']:,}→{r['new']:,}）"
    if p['ledgerLast']:
        head += f"＋預購台帳更新至 {md(p['ledgerLast'])}"
    body = [f"check_source 四分頁到 {md(s['expectedDate'])}（exit 0）。", '',
            f"- actualRows {r['prev']:,} → {r['new']:,}，最後日期 {md(r['prevLast'])} → {md(r['newLast'])}"]
    retro = c['2_daily']['retro']
    body.append(f"- 回溯調整 {len(retro)} 筆" + ('' if not retro else '：' + '；'.join(
        f"{md(x['key'][0])} {x['key'][1]} {x['key'][3]} {x['qty'][0]:g}→{x['qty'][1]:g} 件" for x in retro[:5])))
    body.append(f"- 預購（每日檔排除）{p['dailyTotal'][0]:g} → {p['dailyTotal'][1]:g} 件；台帳 {p['ledgerQty'][0]:g} → {p['ledgerQty'][1]:g} 套")
    body.append(f"- 六項必查：前月彙總未變、無日期列 {c['4_nodate']['count']}、缺品名 {len(c['5_missing']['new'])}、kolSchedule 未變")
    body.append(f"- validate_html.py {s['validate']['tail']}")
    return head + '\n\n' + '\n'.join(body) + '\n\n' + TRAILER + '\n'


def cmd_publish(a):
    s, path = latest_summary(a.summary)
    if s['status'] != 'ok' and not a.force:
        print(f"✗ 摘要狀態是 {s['status']}，不發布。原因：{s['stops']}\n  確認無誤後才可加 --force。"); return 1
    _, dirty = git('status', '--porcelain', '--untracked-files=no')
    files = {ln[3:].strip() for ln in dirty.splitlines() if ln.strip()}
    if files != {'data.json'}:
        print(f'✗ 工作區變更不只 data.json：{sorted(files)}'); return 2
    msg = commit_message(s)
    git('add', 'data.json')
    p = subprocess.run(['git', 'commit', '-q', '-F', '-'], cwd=REPO, input=msg, text=True, encoding='utf-8',
                       capture_output=True, env=env_utf8())
    if p.returncode:
        print('✗ commit 失敗：' + p.stdout + p.stderr); return 2
    sha = git('rev-parse', 'HEAD')[1].strip()
    print(f'commit {sha[:7]}：{msg.splitlines()[0]}')
    rc, out = git('push', 'origin', 'main')
    if rc:
        print('✗ push 失敗：\n' + out); return 2
    print('push 完成，等待部署…')

    local = load_json(DATA_JSON)
    deadline = time.time() + 300; state = 'pending'
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(GH_API, timeout=20) as f:
                runs = json.load(f)['workflow_runs']
            hit = [r for r in runs if r['head_sha'] == sha and r['name'] == 'Deploy to GitHub Pages']
            if hit:
                state = f"{hit[0]['status']} {hit[0]['conclusion']}"
                if hit[0]['status'] == 'completed':
                    break
        except Exception as e:  # 網路暫時失敗就再等
            state = f'查詢失敗 {e}'
        time.sleep(10)
    print(f'部署：{state}')
    ok = state == 'completed success'
    if ok:
        time.sleep(5)
        try:
            with urllib.request.urlopen(f'{LIVE_DATA}?_={int(time.time())}', timeout=30) as f:
                live = json.load(f)
            same = live.get('generatedAt') == local.get('generatedAt') and len(live['actualRows']) == len(local['actualRows'])
            print(f"線上 data.json：generatedAt {live.get('generatedAt')}　{len(live['actualRows']):,} 筆　{'與本機一致' if same else '與本機不一致！'}")
            ok = same
        except Exception as e:
            print(f'線上 data.json 讀取失敗：{e}'); ok = False
    s.update({'published': {'sha': sha[:7], 'deploy': state, 'liveVerified': ok}})
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(s, f, ensure_ascii=False, indent=1)
    return 0 if ok else 1


# ───────────────────────── hold ─────────────────────────
def cmd_hold(a):
    s, _ = latest_summary(a.summary)
    _, dirty = git('status', '--porcelain', '--untracked-files=no')
    if not dirty.strip():
        print('沒有可保留的變更。'); return 0
    reason = '；'.join(s.get('stops') or ['手動保留'])[:300]
    for step in (('checkout', '-B', 'local/hold'), ('add', 'data.json')):
        rc, out = git(*step)
        if rc:
            print(out); return 2
    p = subprocess.run(['git', 'commit', '-q', '-F', '-'], cwd=REPO, text=True, encoding='utf-8', capture_output=True,
                       input=f"hold: 來源檔不完整，暫不發布（期望 {s.get('expectedDate')}）\n\n{reason}\n\n{TRAILER}\n", env=env_utf8())
    rc, out = git('checkout', 'main')
    print(f"已存到 local/hold（{git('rev-parse', '--short', 'local/hold')[1].strip()}），main 的 data.json 已還原。" if not rc else out)
    return 0 if not (p.returncode or rc) else 2


def main():
    ap = argparse.ArgumentParser(description='FY26 日報每日更新')
    sub = ap.add_subparsers(dest='cmd', required=True)
    p1 = sub.add_parser('prepare'); p1.add_argument('--actual'); p1.add_argument('--ledger'); p1.add_argument('--date')
    p2 = sub.add_parser('publish'); p2.add_argument('--summary'); p2.add_argument('--force', action='store_true')
    p3 = sub.add_parser('hold'); p3.add_argument('--summary')
    a = ap.parse_args()
    return {'prepare': cmd_prepare, 'publish': cmd_publish, 'hold': cmd_hold}[a.cmd](a)


if __name__ == '__main__':
    sys.exit(main())
