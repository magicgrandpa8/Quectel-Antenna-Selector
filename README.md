# Quectel 天線選型 · TELEC 認證 PWA

目前版本：**v3.23.3**
資料來源：`Quectel_Antenna_Product_and_TELEC_Certification_*.xlsx`（單一 Excel）

## v3.23.3 圖片檢查
- 已檢查 319 款天線圖片與全部 318 份連結文件；更新 235 款圖片，保留 84 款較合適的現有圖。
- 完整保留產品、線材、接頭及正反面視圖，加入白邊；縮圖上限 256px / q88、放大圖 1200px / q94，僅縮小不放大。
- 放大檢視取消強制放大至原圖兩倍，保留 v3.23.2 的手機兩欄三列規格排版。
- 各 OC 檢查結果與來源記錄於 `data/photo_review.csv`。
- YENA00L5AH 無 datasheet；YCGS010AA 封面是 EVB；YFGA010E3AM 連結文件實際為 YFGA225E3AM，這三款保留原圖。

## 功能
- **TELEC 認證模組選單**：依類別分組（LPWA、LTE-Cat1 BIS、LTE-Cat1/4、5G/LTE-A、Auto、Smart、Short Range），顯示認證日期與已認證天線數；選定後只列出與該模組一起取得 TELEC 認證的天線
- 清單每列顯示「TELEC ×N」已認證模組數、JP 相容標示、尺寸、Datasheet 按鈕
- 主動式 GNSS 天線顯示 LNA 增益；IPEX MHF 與 RG174 / RG174LL 線材：連接器、線材型號、線長（IPEX 另含線徑）合併描述
- 詳細資料列出所有認證模組（依類別、附認證日期）、射頻效能、完整規格
- 快速選單（由上到下）：Tech → Japan Market（Yes / No）→ TELEC 認證模組（僅 Japan Market = Yes 時出現，可多選：只列出全部通過認證的天線）→ GNSS 天線類型（僅 Tech = GNSS 時出現：Passive / Active / Dipole）→ Product Type → Form（型態）；側欄：安裝、連接器、IP、埠數、銷售區域、尺寸上限
- 篩選：僅顯示有 TELEC 認證；日本市場相容改由 Japan Market 下拉選單篩選
- 「移遠天線Agent」按鈕：位於頁面標題右側，開啟 https://gemini.google.com/gem/54870e5583f8（`template.html` 的 `AGENT_URL`）
- 介面語言：中文 / English / 日本語 / 한국어（首次開啟預設為英文，之後記住使用者的選擇；預設值為 `template.html` 的 `DEFAULT_LANG`）
- 深色 / 淺色、版本資訊、強制更新
- 完全離線、可安裝為 App（iOS / Android / Windows / macOS）

## 使用說明（自動同步）
右上角「使用說明」按鈕會依**目前選擇的語言**即時產生使用說明（中 / 英 / 日 / 韓），可「列印 / 存成 PDF」。

- 使用說明內建在 `index.html`（`index.html?guide=1&lang=xx`），離線也能開啟。
- iPhone 畫面是工具本身在示範狀態下的**即時畫面**（`?embed=…`），不是截圖；介面改動會自動反映。
- 版本號、Excel 版次、天線與模組數量、範例料號都取自目前資料。範例料號會自動挑選仍存在的料號。
- 因此每次用 `build_data.py` 更新工具，使用說明就會同步更新，不需另外處理。
- **新增功能時**：使用說明的文字在 `template.html` 的 `GUIDE` 物件（四種語言），需隨功能一併補上說明。

匯出 PDF 檔（選用，例如要寄送或上傳）：
```bash
pip install playwright && python -m playwright install chromium
python3 tools/export_guides.py            # 輸出到 guides/，四種語言
```

## Excel 讀取規則
| 工作表 | 用途 |
|---|---|
| `Antenna vs Cert` | 天線規格（`Compatible with Japan Market` 之前的欄位）＋ 其後各模組欄的 ☑ 認證矩陣；第 2 列為統計列，資料自第 3 列起 |
| `Name` | 模組名稱、Cert Date、Category（合併儲存格自動向下填補） |
| `Histories` | 取最新一筆 Version 與 Updated Date 顯示於版本資訊 |

不讀取：veryHidden 的各模組原始認證工作表、`Histories` 的人員姓名。
Build 時會比對各模組認證數與 Excel 統計列，並警告矩陣中不在 `Name` 的模組。

## 安裝為 App
| 平台 | 瀏覽器 | 步驟 |
|---|---|---|
| iPhone / iPad | **Safari**（必須） | 分享 → 加入主畫面 → 新增 |
| Android | Chrome / Edge | App 內「安裝 App」或選單 ⋮ → 安裝應用程式 |
| Windows | Edge / Chrome | App 內「安裝 App」或網址列安裝圖示 |

安裝後在有網路時開啟一次，之後即可完全離線使用。

## 手動指定產品照片（最優先）
Excel 照片錯誤或太模糊時，把正確的照片放進 `photo_overrides/` 資料夾，**檔名用料號**
（例如 `YC0001CA.png`，支援 png / jpg / jpeg / webp），重新產生工具即可。
這裡的照片優先於 Excel 與 `--photo-xlsx`（若該料號有 datasheet 照片，則以 datasheet 為準），縮圖與放大檢視都會使用。
檔名找不到對應料號時，build 會顯示警告。

## Datasheet 連結（OneDrive）
`data/datasheet_links.csv`（欄位：`PN, Type, Datasheet_File, OneDrive_URL`）是每個料號的 datasheet 對照表，
連結為 OneDrive 個別檔案分享連結（任何人可檢視、不需登入）。build 時：
- 各天線的 Datasheet 按鈕改用 `OneDrive_URL`，於新分頁開啟（取代 Excel 中的連結）。
- `OneDrive_URL` 空白（例如 YENA00L5AH 沒有 datasheet）或不在表中的料號，不顯示 Datasheet 按鈕。
- 只接受 `https://` 連結；表中有、工具中沒有的料號會顯示警告。
- Service Worker 不快取外部連結；離線時按 Datasheet 會提示需要網路。

新增或更新 datasheet：在 OneDrive 建立分享連結，更新 CSV 對應列，重新 build 即可。

## 從 datasheet 擷取清晰照片（建議）
擷取工具使用本機的 datasheet PDF（資料夾或 .zip），檔名含料號最佳（沒有時會讀 PDF 第 1 頁文字找料號）：
```bash
pip install pymupdf pillow
python3 tools/extract_datasheet_photos.py <PDF 資料夾或 .zip> --only-blurry
```
- 自動略過 logo、各頁重複的頁首圖、滿版背景，取第 1～3 頁中的產品照片；透明背景補白底並裁白邊。
- 結果存入 `photo_datasheet.json`（單一檔案，WebP、最長邊 800px）；之後執行 `build_data.py` 即會使用。
- 不想採用 datasheet 照片的料號（例如封面是評估板），列在 `data/photo_datasheet_exclude.csv`（`PN, Reason`），擷取與 build 都會略過。
- 請先檢查 `photo_datasheet_review.png`（擷取結果對照圖）與 `photo_datasheet_report.csv`（每份 PDF 的處理結果）。
- 照片優先順序：**datasheet 照片** > `photo_overrides/` 手動照片 > `--photo-xlsx` > TELEC Excel。

## 較清晰的產品照片（選用）
TELEC Excel 中約一半的照片解析度偏低。若其他 Excel（例如 JP 天線表）有同一料號更清晰的照片，
build 時加上 `--photo-xlsx`（可重複指定），清晰 2 倍以上的照片會自動替換（縮圖與放大檢視一起換）；
其他 Excel 只用來取照片，不讀取任何資料。

```bash
python3 build_data.py --version X.Y.Z --xlsx <TELEC Excel> \
  --photo-xlsx Quectel_Antenna_Product_for_JP_202604__-_update.xlsx
```

## 發佈新版本（Excel 每季更新時）
```bash
pip install openpyxl pillow
python3 build_data.py --version 3.2.1 \
  --xlsx ../Quectel_Antenna_Product_and_TELEC_Certification_V3_x_xxxxxxxx_RevXX_Qx_for_CU.xlsx
```
`--app-url` 可指定工具正式網址（預設為 GitHub Pages 網址），用於安裝說明的「開啟網頁 / 複製連結」。
會同時產生 `index.html`、`version.json`，並自動更新 `sw.js` 的 `CACHE_VERSION`。
發佈後，使用者有網路時**第一次開啟**就會載入新版（等待網路上限 3.5 秒，逾時或離線則使用快取）；App 從背景切回時也會重新檢查新版本。

接著 commit / push：`index.html`、`version.json`、`sw.js`（有修改時再加上 `template.html`）。
App 內不顯示更新紀錄；版本變更請記錄在本 README 的「更新紀錄」。

版本號：X 大改版 ／ Y 新功能 ／ Z 僅更新 Excel 資料或修正。

## 檔案說明
| 檔案 | 用途 | 部署需要 |
|---|---|---|
| `index.html` | App 本體（內嵌資料、圖片、字型） | ✅ |
| `manifest.webmanifest`、`icon-*.png`、`apple-touch-icon.png`、`screenshot-*.png` | 安裝資訊與圖示 | ✅ |
| `sw.js` | 離線快取 | ✅ |
| `photos.json` | 產品照片（放大檢視用，原始解析度，全部打包成單一檔案，build 時自動產生） | ✅ |
| `version.json` | 新版偵測 | ✅ |
| `template.html`、`build_data.py`、`fonts/` | 產生 `index.html` 的原始碼 | 建議放在 repo |
| `data/photo_datasheet_exclude.csv` | 不採用 datasheet 照片的料號與原因 | ✅ |
| `data/datasheet_links.csv` | Datasheet 連結對照表（PN → OneDrive 分享連結） | ✅ |
| `photo_datasheet.json` | 從 datasheet 擷取的產品照片（`tools/extract_datasheet_photos.py` 產生） | ✅（有擷取時） |
| `photo_overrides/` | 手動指定的產品照片（檔名為料號） | 建議放在 repo |
| `tools/export_guides.py` | 將內建使用說明匯出為 PDF（選用） | 建議放在 repo |
| `guides/*.pdf` | 已匯出的使用說明 PDF（四種語言） | 選用 |

## 注意事項
- 不要 commit 原始 `.xlsx`。
- Datasheet 連結不顯示在畫面上，但存在於 `index.html` 原始碼；公開 repo 時任何人都能看到這些 SharePoint 網址。
- 頻段預設邊界為一般公開配置，僅供初篩。

## 更新紀錄
- **v3.23.3**
