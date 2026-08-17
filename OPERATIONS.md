# FY26 商售業績日報儀表板 — 操作說明（交接用）

單頁式儀表板：`index.html`（內嵌 JS + Chart.js）讀取 `data.json` 呈現。

## 對外網址
- **https://benesse-lc8t.github.io/fy26-daily-report/**（GitHub Pages，從 `main` 分支自動部署）
- 部署設定：`.github/workflows/deploy.yml`（push 到 `main` 即自動部署，約 1–2 分鐘）
- `data.json` 讀取已加 `?_=<timestamp>` 快取破除，一般重整即最新。

## 每日流程（使用者上傳最新 FY26 Excel 後）
```bash
python3 scripts/update_data.py <上傳的 FY26 Excel 路徑>
git add data.json
git commit -m "data: update actualRows (...)"
git push origin <目前分支>:main      # 一定要推到 main 才會部署
```
> 相依套件：`openpyxl`。若報 ModuleNotFoundError → `pip install openpyxl -q`。

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

## 已知事項
- commit 在 GitHub 顯示「Unverified」＝執行環境無簽章金鑰，**不影響部署與資料**，可忽略。
- GitHub Pages 部署狀態：Actions → workflow「Deploy to GitHub Pages」。

## 尚在討論、未落地（不在 repo 內，需靠對話脈絡）
- **集計製作流程自動化**（`daily_run.py`：解密蝦皮檔→Excel RefreshAll→update_data.py→git push，跑在使用者筆電）。
- **一封信全覽 + 各角色連結** 的 email 自動化。
