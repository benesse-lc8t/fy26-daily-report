#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
daily_run.py — 每日集計「一鍵自動化」（在使用者 Windows 筆電執行）

流程：
  ① 解密蝦皮密碼檔（msoffcrypto）→ 產生無密碼副本，供主檔 Power Query 讀取
  ② 開啟主檔 Excel、RefreshAll（CRM cube＋樞紐表＋Power Query）、存檔、關閉（pywin32）
  ③ 跑 update_data.py → 更新 data.json
  ④ git add / commit / push 到 main → GitHub Pages 自動部署

設定：把「機器相關路徑」與「蝦皮密碼」填在同資料夾的  daily_run.local.json
      （此檔已被 .gitignore 擋掉，不會上傳 GitHub）。範本見 daily_run.local.example.json

第一次先安裝相依套件（在筆電命令列）：
      pip install msoffcrypto-tool pywin32 openpyxl

注意：② 需要 Windows + 已安裝 Excel（M365）。若尚未把主檔改成 Power Query 直接讀
      蝦皮/Shopline 檔，② 只會刷新 CRM cube 與樞紐；蝦皮/Shopline 明細仍需人工貼上。
"""

import os, sys, json, glob, subprocess, datetime, traceback

HERE        = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(HERE, 'daily_run.local.json')


def log(msg):
    print(f'[{datetime.datetime.now():%H:%M:%S}] {msg}', flush=True)


def load_config():
    if not os.path.exists(CONFIG_PATH):
        log(f'找不到設定檔：{CONFIG_PATH}')
        log('請把 daily_run.local.example.json 複製成 daily_run.local.json 並填入你的設定。')
        sys.exit(1)
    with open(CONFIG_PATH, encoding='utf-8') as f:
        return json.load(f)


def step_decrypt(cfg):
    """① 蝦皮密碼檔解密（取 inbox 內最新符合檔名者）。"""
    pattern = cfg.get('shopee_encrypted_glob')
    if not pattern:
        log('① （略過蝦皮解密：未設定 shopee_encrypted_glob）')
        return
    files = sorted(glob.glob(pattern), key=os.path.getmtime)
    if not files:
        log(f'① 找不到蝦皮檔：{pattern}')
        sys.exit(1)
    src = files[-1]
    dst = cfg['shopee_decrypted_path']
    import msoffcrypto
    with open(src, 'rb') as fin:
        office = msoffcrypto.OfficeFile(fin)
        office.load_key(password=cfg['shopee_password'])
        with open(dst, 'wb') as fout:
            office.decrypt(fout)
    log(f'① 蝦皮解密完成：{os.path.basename(src)} → {dst}')


def step_refresh_excel(cfg):
    """② 開主檔 → RefreshAll → 等查詢完成 → 存檔關閉。"""
    master = os.path.abspath(cfg['master_xlsx'])
    import win32com.client  # pywin32
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(master, UpdateLinks=0)
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()   # 等 Power Query / cube 刷新完成
        wb.Save()
        wb.Close(SaveChanges=True)
        log(f'② Excel RefreshAll 完成並存檔：{master}')
    finally:
        excel.Quit()


def step_update_data(cfg):
    """③ 跑 update_data.py（產生 data.json）。"""
    script = os.path.join(HERE, 'update_data.py')
    r = subprocess.run([sys.executable, script, os.path.abspath(cfg['master_xlsx'])])
    if r.returncode != 0:
        log('③ update_data.py 失敗')
        sys.exit(1)
    log('③ data.json 更新完成')


def _git(repo, *args):
    return subprocess.run(['git', *args], cwd=repo)


def step_git_push(cfg):
    """④ 拉最新 → add/commit → push 到 main。"""
    repo = cfg['repo_dir']
    _git(repo, 'fetch', 'origin', 'main')
    _git(repo, 'add', 'data.json')
    if subprocess.run(['git', 'diff', '--cached', '--quiet'], cwd=repo).returncode == 0:
        log('④ data.json 無變更，略過推送')
        return
    today = datetime.date.today().isoformat()
    _git(repo, 'commit', '-m', f'data: daily auto-update {today}')
    if _git(repo, 'push', 'origin', 'HEAD:main').returncode != 0:
        log('④ git push 失敗（檢查 git 權杖是否設定）')
        sys.exit(1)
    log('④ 已推送到 main → GitHub Pages 將自動部署')


def main():
    cfg = load_config()
    try:
        step_decrypt(cfg)
        step_refresh_excel(cfg)
        step_update_data(cfg)
        step_git_push(cfg)
        log('✅ 全部完成')
    except SystemExit:
        raise
    except Exception:
        log('❌ 發生錯誤：')
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
