# 交接文件 — FY26 商售日報儀表板（雲端 session → Claude Code 本機版）

> **新 session 請先讀這份，再讀 `OPERATIONS.md`。**
> 本檔只寫「repo 裡沒有的東西」：本機環境、使用者偏好、目前狀態、行為準則。
> 技術細節（解析規則、儀表板功能、停售預購規則）一律看 `OPERATIONS.md`。
>
> 交接時點：**2026-10-21**（資料已更新至 10/20，commit `87a4416`）

---

## 0. 三十秒摘要

每天使用者把兩個 Excel 更新好，你把資料灌進 `data.json`、驗證、推 `main`，
GitHub Pages 自動部署，公司各部門看儀表板。
**最重要的一條規則：來源檔不完整時不要推上線。**（兩週內已救過三次）

---

## 1. 使用者

**Doris（王詩雅）**，倍樂生台北分公司・總經理室經營戰略組。
職務是行銷數據分析：每月銷售數據彙整、異常偵測、趨勢報告、公司課題整合。

### 溝通偏好（務必遵守）
- **一律繁體中文**
- **先給結論**，再給理由；條列、有結構
- **能用表格就用表格**
- 開頭不要客套話、贅詞；**不要用破折號（em dash）**
- 專業術語在文末加一行白話說明
- 她要的不只是整理資料，更要**讀出數字背後的訊號**、主動提假設、指出風險
- 她卡住時，幫她把語意和邏輯整理出來
- 她說的內容當作「當下限定的資訊」，不要過度斷定

### 她的立場
你是她的**工作夥伴**，不是執行工具。發現問題要主動講，包含「這個指標沒有意義」
「這個數字不能這樣比」「這件事做不到」這類逆耳的結論。她會接受，而且需要。

---

## 2. 本機環境（Claude Code 本機版專屬）

### 檔案位置（Windows）
```
C:\Users\014235\mail.benesse.com.tw\財務相關-連結confluence - 01.連結confluence\FY26\★每月更新PL分析作業(重要文件夾)\N.商售集計(每日更新)\
```
這個資料夾裡每天會更新的兩個檔：

| 檔名 | 用途 | 餵給哪支腳本 |
|---|---|---|
| `FY26商售實績明細(claude).xlsx` | 每日實績明細 | `update_data.py` |
| `▲FY26雙語點讀圖鑑預購統計(2609-2611月).xlsx` | 預購台帳 | `update_preorder_ledger.py` |

> 這是 OneDrive／SharePoint 同步資料夾。**本機版讀得到**（雲端版讀不到，這也是改用本機版的原因）。
> 注意：若檔案正被 Excel 開啟，可能讀到鎖定狀態，請使用者先關閉 Excel。

### GitHub repo

| 項目 | 值 |
|---|---|
| **Repo** | **https://github.com/benesse-lc8t/fy26-daily-report** |
| Clone | `git clone https://github.com/benesse-lc8t/fy26-daily-report.git` |
| 預設分支 | `main` ← **推這裡才會部署** |
| 開發分支 | `claude/gifted-shannon-kqjcrs`（雲端 session 用的，本機版可自己開新的） |
| **線上儀表板** | **https://benesse-lc8t.github.io/fy26-daily-report/** |
| 部署設定 | `.github/workflows/deploy.yml`（push 到 `main` 即自動部署，約 1〜2 分鐘） |
| 部署狀態 | GitHub Actions → 「Deploy to GitHub Pages」 |

**第一次設定（在使用者筆電）**
```bash
cd D:\dashboard                 # 放哪裡都行，記得填進 daily_run.local.json
git clone https://github.com/benesse-lc8t/fy26-daily-report.git
cd fy26-daily-report
pip install openpyxl
python scripts/check_source.py --help     # 確認能跑
```

**推送規則**
```bash
git add -A
git commit -m "data: ..."
git push origin <目前分支>:main          # 一定要推到 main 才會部署
```
- 本機版需要使用者的 GitHub 認證（Git Credential Manager 或 PAT，第一次 push 時會要求登入）
- commit 顯示「Unverified」是正常的（無簽章金鑰），不影響部署與資料
- **不要開 Pull Request**，除非使用者明講要

### 本機 repo 路徑
使用者的筆電上需要有一份 clone。路徑請第一次跟她確認後記進 `scripts/daily_run.local.json`。

### 相依套件
```
pip install openpyxl msoffcrypto-tool pywin32
```
（`pywin32` 只有要做 Excel RefreshAll 自動化時才需要）

---

## 3. 每日流程（目前的 SOP）

| # | 誰 | 動作 |
|---|---|---|
| 1 | 使用者 | Excel 跑 RefreshAll → 存檔 |
| 2 | 使用者 | 雙擊 `scripts\check_source.bat`（或你代跑）確認四個分頁都到昨天 |
| 3 | 使用者 | 在 chat 說「更新」 |
| 4 | **你** | 讀那兩個檔 → 跑 ETL → 六項必查 → 驗證 → commit → 推 `main` → 確認部署 → 回報 |

**本機版的優勢**：第 2、3 步可以合併 —— 使用者說一聲，你自己跑 `check_source.py`，
綠燈就直接往下做，紅燈就告訴她哪個分頁要重跑。

### 指令
```bash
python scripts/check_source.py "<實績檔完整路徑>"          # 先自檢，exit 0 才往下
cp data.json /tmp/prev.json                                 # 留底供比對
python scripts/update_data.py "<實績檔完整路徑>"
python scripts/update_preorder_ledger.py "<台帳完整路徑>"
python scripts/validate_html.py                             # 11 項靜態檢查
git add -A && git commit -m "..." && git push origin <branch>:main
```

---

## 4. 每次更新必查的六項（不可省略）

| # | 項目 | 判準 |
|---|---|---|
| 1 | 筆數與最後日期 | 變化是否合理 |
| 2 | **逐日差異** | 新增日期以外的變動要**逐筆說明**（取消／改列日期／補登） |
| 3 | 前月彙總未變 | 4月〜上月合計若變動，**先停下來確認** |
| 4 | 無日期列 | 應恆為 11 筆（4月 WEB 數位典藏），其餘月份為 0 |
| 5 | 缺品名製編 | 有新製編要請使用者提供商品主檔 |
| 6 | kolSchedule | 無 KOL 檔期表更新時應完全一致 |

> 第 2 項是重點。**不要只報「金額 +X 萬」**，要列出哪一天、哪個通路、哪個商品、差幾件幾元。
> 使用者靠這個判斷是真實退單還是系統問題。

---

## 5. 絕對不要做的事（都是踩過的坑）

### 🔴 來源檔不完整時不要推上線
RefreshAll 常常跑不完，某些分頁會停在前一天。**這種缺漏不會報錯、數字看起來也合理**，
推上去會變成「某天業績崩盤」的假象，而儀表板是總經理和各部門在看的。

實例（2026-10 連三天）：
| 日期 | 缺的分頁 | 若推上去會顯示 | 實際值 |
|---|---|---|---|
| 10/19 | `2610WEB&KOL`（缺 3 天） | 週六 0.8 萬、週日 0.1 萬 | 4.6 萬、5.5 萬 |
| 10/20 | `2610WEB&KOL`（缺 1 天） | 週一 7.2 萬 | 15.9 萬 |
| 10/21 | `2610WEB數位典藏`（缺 1 天） | 影響小但數字不對 | — |

**正確做法**：commit 到開發分支保留進度，**不推 `main`**，把證據列給使用者請她重跑。
使用者已明確表示偏好這個做法（2026-10-19 當面確認過）。

### 🔴 小額金額不要用 `fmtAmt()`
`fmtAmt()` 是 `Math.round(n/10000)`，單日 × 單一管道常常不到 1 萬會被抹成 0。
已踩過兩次（手偶 `Math.floor`、前日業績 EC 整條消失）。
**單日或單一細分維度一律用 `fmtAmtFine()`。**

### 🔴 口徑改動要先問
例如「數位典藏要不要納入全社」（目前排除，見第 7 節）。
這類決策會動到所有歷史數字，**不要自己決定**。

### 🔴 不要擴充 TM 三指標到其他事業
`index.html` 的 TM 推案結構三指標**鎖定學習商品**，原因寫在程式碼註解與
`OPERATIONS.md`。其他事業要嘛門檻數學上不可達、要嘛合購判定會誤判。

---

## 6. 兩個已知的資料源特性（不是 bug）

### 蝦皮訂單狀態會來回跳
同一筆單會「2 → 1 → 2」反覆變動，不是取消，是來源檔對訂單狀態的暫時性判定。
觀察到的穩定期約 **T+2**。月底結算建議 **11/3 再確認一次 10 月的蝦皮數字**。

### 台帳的 WEB 欄晚一天
預購台帳的 WEB 欄固定隔天才補上（蝦皮、TM 當天就有）。
所以台帳總數會比每日檔少，**明天會自動對上**，不是錯誤。

### TM 的時點不同（設計如此）
TM 以**發貨日**開列，預購未出貨不進每日明細，所以台帳 TM 永遠多於每日檔。
儀表板的 `poTaken()` 因此是：WEB／KOL／EC 取每日檔、TM 取台帳。

---

## 7. 目前待辦

| # | 事項 | 球在誰 | 備註 |
|---|---|---|---|
| 1 | 台帳「金額小計」公式漏算 TM | 使用者 | |
| 2 | 蝦皮訂單狀態判定（哪些狀態計入／幾天穩定） | 使用者 → 資訊部門 | |
| 3 | 四個製編正式品名與事業歸屬 | 使用者 → 製作部門 | `20244C901`／`20244C910`／`20250FF01`／`20260FXAC`，目前暫掛「（品名待補）／學習周邊」 |
| 4 | 預購出貨後設 `recognizeMonth` | **你**，約 11 月底 | 143 套／約 33 萬要認列到正確月份，**要主動提醒使用者** |
| 5 | **數位典藏是否納入「全社」** | **等使用者決定** | 目前 `BIZ_LIST_TOTAL` 排除它（當初是 β 版）。現在 FY26 累計 85 萬、有正式期中目標，前提可能過時 |
| 6 | 補完 `daily_run.py`（一鍵到底） | 評估中 | 目前缺：完整性檢查、`update_preorder_ledger.py`、`validate_html.py` |
| 7 | 兩個 Excel 檔 git-tracked 且公開 | 未決 | `FY25商售實績明細.xlsx`、`FY26商售實績明細(claude).xlsx` 會隨 GitHub Pages 公開，已多次提醒 |

---

## 8. 目前狀態快照（2026-10-21）

| 項目 | 值 |
|---|---|
| `main` HEAD | `87a4416` |
| 開發分支 | `claude/gifted-shannon-kqjcrs` |
| `actualRows` | 10,387 筆 |
| 最後實績日 | 2026-10-20 |
| 預購台帳 | 139 套／317,582 元（儀表板顯示 143 套，差額說明見第 6 節） |
| 線上網址 | https://benesse-lc8t.github.io/fy26-daily-report/ |

### 10 月落點（截至 10/20，全社口徑不含數位典藏）
| 項目 | 值 |
|---|---|
| 累計 | 251.1 萬 |
| 期中目標 | 447.7 萬 |
| 對時間進度達成率 | 86.3% |
| **月末預估** | **382 萬（區間 363〜434）** |
| 對期中目標 | **85.3%，缺口 66 萬** |
| 剩餘工作日 | 6 天（22、23、27、28、29、30） |

**已建議使用者提早向上層預警**，並指出三個可動的槓桿：蝦皮檔期、TM 合購檔期、WEB 高單價品曝光。

---

## 9. 驗證與發布的標準動作

```bash
python scripts/validate_html.py     # 11 項，全過才可 commit
```
改過 HTML／JS／CSS 時，另需實機驗證（Playwright + 本機 http-server）。
雲端環境的既有測試腳本在對話脈絡裡，本機版可重新寫；
最低要求是：**22 個分頁無 JS 例外、無 console error、1440／834／390px 無水平溢出。**

### 發布
```bash
git push origin <branch>:main    # 推 main 才會部署，約 1〜2 分鐘
```
- Repo：https://github.com/benesse-lc8t/fy26-daily-report
- 部署狀態：GitHub Actions → 「Deploy to GitHub Pages」
  （`gh api repos/benesse-lc8t/fy26-daily-report/actions/runs` 可查，若本機有 gh CLI）
- 線上網址：https://benesse-lc8t.github.io/fy26-daily-report/
- commit 顯示「Unverified」是正常的（無簽章金鑰），不影響部署
- **推完要確認部署成功再跟使用者說完成**

### 回報格式（使用者習慣的）
1. 更新內容（前後對照表）
2. 六項必查結果（表格）
3. 回溯調整逐筆
4. **異常與風險的判讀**（這部分她最重視）
5. 驗證紀錄
6. commit 編號 ＋ 部署結果

---

## 10. 延伸閱讀（repo 內）

| 檔案 | 內容 |
|---|---|
| `OPERATIONS.md` | **技術細節總表**：解析規則、儀表板功能、停售預購、三基準比較、版面規範 |
| `ANALYSIS_NOTES.md` | 歸因方法、各通路基準線、KOL 檔期分級、月末推估倍率 |
| `AUTOMATION_PLAN.md` | 自動化規劃 |
| `scripts/check_source.py` | 送檔前自檢（四個分頁完整性） |
| `scripts/daily_run.py` | 一鍵自動化骨架（尚未補完） |
