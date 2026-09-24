# FY26 商售業績日報儀表板 — 操作說明（交接用）

單頁式儀表板：`index.html`（內嵌 JS + Chart.js）讀取 `data.json` 呈現。

## 新 session 起手式
1. 讀本檔（每日流程與解析規則）。
2. 讀 `ANALYSIS_NOTES.md`（累積的分析框架與基準線）—— 要做歸因、寫月報、
   回答「為什麼」時才需要；純粹的每日更新可略過。
3. 之後即可依既有模式作業：使用者上傳當日 FY26 Excel → 跑 `update_data.py`
   → 驗證 → commit → 推 `main` → 確認部署。

> `pip install openpyxl -q`：執行環境偶爾會重建，套件與先前上傳的檔案都會消失，
> 報 ModuleNotFoundError 時裝一次即可。舊的上傳檔案不會留存，需要商品主檔或
> KOL 檔期表時要請使用者重新上傳。

## 對外網址
- **https://benesse-lc8t.github.io/fy26-daily-report/**（GitHub Pages，從 `main` 分支自動部署）
- 部署設定：`.github/workflows/deploy.yml`（push 到 `main` 即自動部署，約 1–2 分鐘）
- `data.json` 讀取已加 `?_=<timestamp>` 快取破除，一般重整即最新。

## 每日流程（使用者上傳最新 FY26 Excel 後）
```bash
cp data.json /tmp/prev.json                        # 先留一份供比對
python3 scripts/update_data.py <上傳的 FY26 Excel 路徑>
# → 比對 prev.json：筆數、最後日期、逐日增減、前月彙總是否未變、
#   無日期列數、缺品名製編、kolSchedule 是否變動
git add data.json
git commit -m "data: update actualRows (...)"
git push origin <目前分支>:main      # 一定要推到 main 才會部署
```
> 相依套件：`openpyxl`。若報 ModuleNotFoundError → `pip install openpyxl -q`。

### 每次更新必查的六項
1. **筆數與最後日期**的變化是否合理
2. **逐日差異** —— 新增日期之外的回溯調整要逐筆說明（取消／改列日期／補齊）
3. **前月彙總未變** —— 4月~上月的合計若變動，代表來源檔有問題，先停下來確認
4. **無日期列** —— 應恆為 11 筆（4月 WEB 數位典藏的真實實績），其餘月份應為 0
5. **缺品名製編** —— 有新製編時請使用者提供商品主檔，再跑 `update_master.py`
6. **kolSchedule** —— 檔數與內容是否如預期（無 KOL 檔期表更新時應完全一致）

## 修改 HTML／JS／CSS 後必跑
```bash
python3 scripts/validate_html.py       # 11 項靜態檢查，全過才可 commit
```
檢查項目：script 標籤對齊、`node --check` 語法、關鍵變數宣告順序、CDN 來源
（只允許 cdnjs.cloudflare.com）、重複 let/const、render 函式是否有呼叫點、
DOM 引用殘留、新功能格式（斜線零須關閉）、本機 fallback 檔案、data.json 結構。

需要實機驗證時（Playwright + 本機 http-server）：
```bash
npx http-server -p 8899 -s -c-1 .      # 另開背景執行
PLAYWRIGHT_BROWSERS_PATH=/opt/pw-browsers node <e2e腳本>
```
> 本執行環境的網路政策會擋掉 `cdnjs.cloudflare.com` 與 `github.io`，
> 測試時 console 會出現 `ERR_TUNNEL_CONNECTION_FAILED`（Chart.js 會自動退回
> 本機 `chart.umd.js`）。**這是沙箱限制，不是網站問題**，實機不會發生。

## 其他更新腳本
- **商品主檔**：`python3 scripts/update_master.py <商品主檔.xlsx>`
  自動偵測分頁與欄位標題（相容新舊格式）；只保留「有實績」的製編品名以維持 data.json 精簡。
- **KOL 檔期**：`python3 scripts/update_kol.py <含 KOL檔期規劃表 的檔案>`

## `update_data.py` 解析規則重點
- 解析分頁：`<YYMM>WEB&KOL`、`<YYMM>EC`、`<YYMM>TM`、`<YYMM>WEB數位典藏`、`<YYMM>經代銷`
- **跳過「整張無有效日期」的分頁**（＝未來月份的佔位假資料，例如尚未有實績的 8/9 月）
- 事業名正規化：`學習→學習商品`、`數位經典→數位典藏`；事業別**取自 Excel 的「事業」欄**（不是主檔）
- 略過小計列（事業欄為 0/總計等）與未稅金額 0 的列
- 同時建 `holidayInfo`（國定假日表＋人工登錄的停班停課，如颱風假）

## 儀表板功能備忘
- 預設月份依「今天實際月份」自動切換。
- 身份客製化：`?role=` 連結 → `gm / strategy / finance / product / marketing / tmsales / other`。
- 分頁：主儀表板、前日業績、業績卡片、月別推移、週別趨勢、對目標、AOV、TOP20、手偶、
  事業別(學習商品/學習周邊/生活周邊/Mirafeel/數位典藏)、通路別、KOL檔期、國定假日／停班課。
- 達成率配色規則：**≥100% 綠、<100% 紅**（無黃色）；圖表達成率折線維持紅色、只有數字標籤依此變色。
- **月累積預估趨勢**（前日業績分頁下方）：月份／管道別／事業別三個下拉選單，
  以「月末倍率法」推估全月（保守／基準／樂觀／本月實測四情境）。
  倍率由已完結月份動態計算，資料更新後自動重算，無須手動維護。
- 數字字型：全站統一 `--font-ui`，`font-feature-settings` 已關閉斜線零
  （`"zero" 0`）避免 0 被誤讀為 Q；勿再引入 monospace 字型。

## 已知事項
- commit 在 GitHub 顯示「Unverified」＝執行環境無簽章金鑰，**不影響部署與資料**，可忽略。
- GitHub Pages 部署狀態：Actions → workflow「Deploy to GitHub Pages」。

## 分析相關
歸因方法、各通路基準線、KOL 檔期分級、月末推估倍率等，見 **`ANALYSIS_NOTES.md`**。

## 尚在討論、未落地（不在 repo 內，需靠對話脈絡）
- **集計製作流程自動化**（`daily_run.py`：解密蝦皮檔→Excel RefreshAll→update_data.py→git push，跑在使用者筆電）。
- **一封信全覽 + 各角色連結** 的 email 自動化。
