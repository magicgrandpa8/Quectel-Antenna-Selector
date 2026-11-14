# Quectel 天線選型 · TELEC 認證 PWA — Claude Code 專案說明

公開網站：https://magicgrandpa8.github.io/Quectel-Antenna-Selector/ （GitHub Pages，push 到 main 後約 1～2 分鐘自動部署）
維護者：王柏寒 Andy Wang（andy.wang@pactw.com），以繁體中文溝通。

## 最重要的規則
- **這是公開 repo**：不可 commit 公司內部 Excel（`.xlsx`，已列在 `.gitignore`）或任何內部資料。Excel 放在 repo 以外的資料夾，或 repo 內的 `work/`（已忽略）。
- **README 不寫版本更新紀錄**：不要新增「更新紀錄」或任何版本歷史；只更新開頭「目前版本：**vX.Y.Z**」一行。
- 每次修改前先 `git pull`；修改後要測試再 commit。commit 訊息用英文、簡短，例如 `v3.26.0: data update Rev26.11`。
- 版本號要同步四處：`version.json`、`sw.js` 的 `CACHE_VERSION`、`index.html` 內 DATA 的 `"version"`（由 build 產生）、README 的「目前版本」。資料更新用次版號 +1，介面小修用修訂號 +1。
- `sw.js` 一定和其他檔案在同一次 commit 推送（避免使用者拿到不完整的版本）。
- **每個版本都打 git tag**：發版 push 完成後，在該版本最後一個 commit 建立 annotated tag 並推送，用 tag 區分版本、避免混淆：
  `git tag -a vX.Y.Z -m "vX.Y.Z: <摘要>"`，接著 `git push origin vX.Y.Z`。版本紀錄以 git tag / commit 為準，不寫在 README。
  查詢某版內容：`git show vX.Y.Z:version.json`；比較兩版：`git diff vA.B.C vX.Y.Z --stat`。

## 專案結構
- `template.html`：介面範本（HTML/CSS/JS、四種語言字串、使用說明內容）。修改介面時改這裡，再 build。
- `build_data.py`：讀 Excel → 產生 `index.html`、`photos.json`、`version.json`、`sw.js`。
- `data/datasheet_links.csv`：每個料號的 OneDrive「任何人可檢視」datasheet 連結。
- `photo_datasheet.json`：從 datasheet 擷取的產品照片；`data/photo_datasheet_exclude.csv`：不採用的料號。
- `tools/`：版本更新與連結工具；完整流程見 `docs/RELEASE_PLAYBOOK.md`。

## 環境（Windows）
```powershell
py -m pip install openpyxl pillow pymupdf playwright pypdf
py -m playwright install chromium
```
文件中的 `python3` 在 Windows 請用 `py`；`/mnt/user-data/...` 的路徑換成本機路徑。

## 常用指令
```powershell
# 只改介面 (template.html) 後重新建置：需要目前版本的 Excel
py build_data.py --version 3.26.0 --xlsx "<Excel 路徑>"
# 匯出四種語言使用說明 PDF (guides/)
py tools/export_guides.py
# 本機預覽
py -m http.server 8765    # 開 http://localhost:8765/
```

## 新版 Excel（例如 Rev26.11）
依 `docs/RELEASE_PLAYBOOK.md`（或技能 quectel-antenna-selector-release）：
1. `py tools/release_prepare.py --xlsx <新 Excel> --out <輸出資料夾> --kit <kit.zip> --json work/release.json`
   → 比對 PRO 頁面、寫入 Histories（直接改 XML，保留照片）、輸出標準檔名
   `Quectel Antenna Product and TELEC Certification_V3.71_20261029_Rev26.10.xlsx` 格式的 Excel。
2. 在本機直接執行 kit 內的 PowerShell（複製新 datasheet 到客戶分享資料夾、建立並以未登入驗證「任何人」連結）。
   本機 OneDrive 路徑寫在技能 quectel-antenna-selector-release 中（不寫在公開 repo）。
3. `py tools/release_finalize.py --xlsx <輸出 Excel> --version <新版> --links <CSV> --pdf-dir <新 PDF 資料夾> --diff work/release.json`
4. 逐張檢查新照片（EVB、包裝圖等加入排除清單）→ `py tools/export_guides.py` → 測試 → commit → push。

## 測試重點
天線總數、新天線可搜尋、Datasheet 按鈕開啟正確連結、比較（最多 6 款）、手機篩選面板（右上「清除」、底部「Go (數量)」）、縮圖完整顯示、版本資訊顯示新版本與資料來源檔名、無 JS 錯誤。
