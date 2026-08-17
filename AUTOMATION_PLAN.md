# 集計流程自動化 — 規劃與現況

## 目標
把「每日集計 → 更新儀表板」從半自動推進到全自動。儀表板 repo＝`benesse-lc8t/fy26-daily-report`，部署於 GitHub Pages（從 `main` 分支）。

## 使用者的每日實際流程
1. 同事從**蝦皮、Shopline**下載特定期間訂單報表。
2. 原數據 Excel 含三來源：**CRM(cube)、蝦皮、Shopline**。
3. **蝦皮報表有密碼**，目前人工輸入開啟。
4. 主檔 Excel：更新 CRM cube → 複製蝦皮/Shopline 明細 → 更新樞紐表 → 完成。
5. 提供檔案 → 更新 html → push。

## 已確認的環境條件
1. **執行機器＝使用者筆電**（排程時筆電須開機登入；或做成「點一下捷徑」）；檔案放公司內部 FTP。
2. **蝦皮無 API**（只能人工下載＋自動解密）；**Shopline 有 API**。
3. **CRM cube** 已設定「開檔自動 RefreshAll、關檔」→ 可程式驅動。
4. 有 Power Automate 與 M365；**主語言用 Python**。

## 已定案架構（路線 A：先自動化後段）
一支 `scripts/daily_run.py`（跑在筆電）串起：
`① 解密蝦皮檔 → ② 開主檔 RefreshAll(cube＋樞紐)存檔關閉 → ③ update_data.py 產生 data.json → ④ git push 到 main → Pages 自動部署`
→ **每天人工只剩「下載蝦皮報表丟固定資料夾 + 點一下」。**

## 已交付（本 repo 內）
- `scripts/daily_run.py`：一鍵自動化骨架（解密／Excel刷新／解析／推送）。
- `scripts/daily_run.local.example.json`：設定範本；複製成 `daily_run.local.json` 填入路徑與蝦皮密碼（已 `.gitignore`，不上傳）。
- `scripts/update_data.py`／`update_master.py`／`update_kol.py`：已改為**路徑相對於 repo**，可在 Windows 筆電執行。

## 筆電端一次性設定
1. `pip install msoffcrypto-tool pywin32 openpyxl`
2. `git clone` repo 到筆電，設定 **git 權杖**（讓 push 能自動跑）。
3. 複製並填好 `scripts/daily_run.local.json`。
4. 建立「點一下捷徑」或 Windows 工作排程器（排程時筆電須開機）。

## 兩個前置（讓自動化更完整）
- **主檔改用 Power Query 直接讀報表**（去掉手動複製貼上蝦皮/Shopline 明細）；蝦皮檔先解密成無密碼副本供 Power Query 讀。
- 安全原則：**密碼、API 金鑰不進 GitHub**，只存在筆電本機設定檔。

## 長期（路線 B，之後評估）
Shopline 走 API 直接拉單 →（若 CRM 也能查）跳過 Excel，端到端「排程→API→data.json→push」。真正瓶頸在「資料來源自動化」與「筆電須常開」（未來可換常開機器/雲端）。

## 待辦（球在使用者）
提供 `daily_run.local.json` 的實際值後即可上機測試：
1. 主檔 Excel 路徑　2. 蝦皮檔名規則＋密碼　3. inbox 資料夾　4. repo 位置　5. 觸發方式（排程幾點／手動）
