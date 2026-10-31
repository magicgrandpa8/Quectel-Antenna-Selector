# Quectel 天線選型 · TELEC 認證 PWA

目前版本：**v3.25.3**
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
- 「移遠天線Agent」按鈕：位於頁面標題右側，開啟飛書 Agent https://applink.feishu.cn/T9ao3zlnY4j5（`template.html` 的 `AGENT_URL`）
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

## 版本更新（新版 Excel，例如 Rev26.11）

完整流程見 **[docs/RELEASE_PLAYBOOK.md](docs/RELEASE_PLAYBOOK.md)**。摘要：

1. `python3 tools/release_prepare.py --xlsx <新版 Excel> --out <輸出 Excel> --kit <Windows 程式 zip> --json <比對結果>`
   比對 PRO 頁面（新增／刪除 OC，分有／無 TELEC；原有 OC 無 → 有 TELEC），寫入 Histories，並產生 Windows 一鍵程式。
2. 在 Windows 執行一鍵程式：複製新 datasheet 到客戶分享資料夾、建立「任何人可檢視」連結、收集新 PDF 與連結 CSV。
3. `python3 tools/release_finalize.py --xlsx <輸出 Excel> --version <版本> --links <CSV> --pdf-dir <PDF 資料夾> --diff <比對結果>`
   合併連結、擷取照片、建置，並檢查每款新天線；再執行 `python3 tools/export_guides.py` 匯出使用說明。

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
| `tools/create_and_verify_all_links.ps1`、`tools/apply_datasheet_links.py` | 全部 datasheet 連結建立與未登入驗證、產生對照 Excel | ✅ |
| `tools/release_prepare.py`、`tools/release_finalize.py` | 版本更新自動化（見 docs/RELEASE_PLAYBOOK.md） | ✅ |
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
- **v3.25.3**：移遠天線Agent按鈕改回飛書 Agent（https://applink.feishu.cn/T9ao3zlnY4j5）
- **v3.25.2**：328 份 datasheet 連結以未登入方式全部驗證可開啟（連結不變）；新增 tools/create_and_verify_all_links.ps1、tools/apply_datasheet_links.py；版本更新流程加入「未登入驗證」，NG 的連結不會放上工具
- **v3.25.1**：資料來源 Excel 改用標準檔名（Quectel Antenna Product and TELEC Certification_V3.71_20261029_Rev26.10.xlsx）；新增版本更新自動化（tools/release_prepare.py、tools/release_finalize.py、docs/RELEASE_PLAYBOOK.md）
- **v3.25.0**：資料更新為 Rev26.10（PRO-26.10，330 款天線，新增 11 款、刪除 0 款；EG21-G 新增 6 筆認證；4 款 GNSS 天線安裝方式更正為 Screw；新增 11 款天線的 datasheet 照片與 Datasheet 連結）。build_data.py 支援新版 Excel 結構（自動採用最新的 PRO-YY.MM 頁面、ModuleList、Histories 取最新一筆），照片來源僅為 Excel 時沿用上一版較清晰的照片
- **v3.24.0**：電腦版天線清單新增顯示支援頻率範圍（產品名稱下方，多技術天線以「｜」分隔）
- **v3.23.4**：天線比較上限由 3 款改為 6 款（比較視窗加寬，4～6 款時可左右捲動，欄位名稱固定在左側）
- **v3.23.3**：手機版規格欄位中，IP 等級固定排在右欄，版面更整齊
- **v3.23.0**：強化 Liquid Glass（液態玻璃）效果：動態流動的色彩桌布、隨光源移動的鏡面高光（電腦跟隨滑鼠、手機自動流動）、更強的折射與邊緣光；版本資訊新增「製作者：王柏寒 andy.wang@pactw.com」（email 可點擊寄信）

- **v3.22.1**：移遠天線Agent按鈕改連結至Gemini天線Agent；同步更新離線快取版本。

- **v3.22.0**：移除主畫面的「強制更新」按鈕（改在版本資訊視窗中更新，有新版本時版本按鈕顯示紅點）；標題下方說明改為「用篩選條件找出合適的天線…」；使用說明 01 章改為「篩選條件說明」
- **v3.21.2**：再修正手機深色模式時最上方狀態列為淺色（載入前即套用主題、theme-color 標籤只更新不刪除、加入 color-scheme）
- **v3.21.1**：修正手機深色模式時最上方狀態列仍為淺色；修正手機「篩選」按鈕文字未置中
- **v3.21.0**：新增 127 款 datasheet 官方照片（合計 175 款）；photo_datasheet.json 改存 WebP（15 MB → 2.5 MB）；App 放大照片改為 720px；新增 datasheet 照片排除清單（YCGS010AA）
- **v3.20.0**：Datasheet 按鈕改用 OneDrive 個別檔案分享連結（`data/datasheet_links.csv`），移除 repo 內 PDF 設計；48 款天線改用 datasheet 中的官方產品照片
- **v3.19.1**：datasheet 照片改為最優先（取代手動照片）；YC0001CA 改用 datasheet 照片
- **v3.19.0**：支援 `datasheets/` 資料夾（Datasheet 按鈕連到 repo 中的 PDF、開啟過的可離線閱讀）；修正 datasheet 照片擷取（透明背景、裝飾色條、重複出現的產品照）；5 款天線改用 datasheet 中的清晰照片；使用說明 PDF 改為固定檔名
- **v3.18.0**：手機版結果優先（篩選改為底部面板＋浮動「篩選」按鈕）；天線比較（最多 3 款、差異標色、只看差異、共同認證模組）；TELEC 認證模組可多選（共同認證）
- **v3.17.0**：介面改為 Liquid Glass（液態玻璃）風格；移除所有 AI 修飾照片；新增 datasheet 產品照片擷取工具
- **v3.16.0**：153 張模糊照片（< 200px）以 AI 超解析與平滑處理；放大檢視時標示「此照片經 AI 修飾」
- **v3.15.1**：新增 `photo_overrides/` 手動指定照片；YC0001CA、YC0002AA、YC0011AA 改用正確且較清晰的照片
- **v3.15.0**：19 款天線改用 JP 天線表中更清晰的照片（`--photo-xlsx`）
- **v3.14.1**：放大檢視用照片改為打包成單一檔案 `photos.json`（原本 `photos/` 資料夾有 320 個檔案，超過 GitHub 網頁一次上傳 100 個檔案的上限）
- **v3.14.0**：所有天線照片可點擊放大檢視（清單與詳細資料皆可），使用原始解析度照片並可離線檢視
- **v3.13.0**：電腦版（寬度 > 820px）版面整理：篩選集中左欄（搜尋 → 下拉選單 → 更多條件）、頁首精簡為語言／使用說明／「⋯」、清單每列改為一行關鍵值；手機與平板維持原版面，頁首按鈕依螢幕寬度平均分配；標題下方說明改為「使用下拉式選單篩選天線或是跟移遠天線Agent說明您的需求」
- **v3.12.0**：淺色模式改為無印良品風格配色；App 圖示（iOS / Android / Windows）改為 Q logo；標題下方說明改為「請用下拉式選單進行篩選天線，或是直接使用移遠天線Agent」
- **v3.11.0**：清單每列的規格欄位新增 IP Rating（IP 等級）
- **v3.10.1**：修正發佈新版後需開啟兩次才會更新的問題（開啟時優先載入網路最新版）；App 從背景切回時重新檢查新版本
- **v3.10.0**：頁面標題右側新增「移遠天線Agent」連結按鈕；使用說明同步加入說明
- **v3.9.0**：新增「型態 Form」下拉選單（Product Type 下方，取代側欄外型篩選）；右上角新增「使用說明」按鈕，依目前語言即時產生使用說明並可存成 PDF，隨工具更新自動同步
- **v3.8.0**：下拉選單改為單欄由上往下排列；移除 TELEC 認證模組的「選填」說明（顯示條件不變）
- **v3.7.0**：Type 改名為 Tech 並移到最上方；TELEC 認證模組改為僅在 Japan Market = Yes 時出現；Product Type 移到最下方
- **v3.6.0**：RG174 線材的連接器與線長合併描述；TELEC 認證模組加註「選填」；下拉選單順序調整；天線類型改為僅在 Type = GNSS 時出現
- **v3.5.0**：移除頻段尺、自訂頻段與頻譜條；主動式 GNSS 天線顯示 LNA 增益；IPEX MHF 連接器與線材（線徑、線長）合併描述；手機版射頻效能表格改為左右捲動
- **v3.4.0**：新增 Japan Market 下拉選單（Yes / No），取代側欄「僅顯示日本市場相容」勾選；舊設定自動轉換
- **v3.3.0**：首次開啟的預設語言改為英文（不再依裝置語言自動選擇）；韓文介面改為以詞為單位換行
- **v3.2.0**：安裝說明新增「開啟網頁」「複製連結」按鈕；修正手機版頁尾長檔名造成頁面過寬、iOS 自動縮小的問題；修正手機版詳細資料視窗內容被截斷
- **v3.1.0**：新增日文、韓文介面；版本資訊不再顯示更新紀錄
- **v3.0.0**：改用 TELEC 認證 Excel 單一資料來源；移除 Global / 日本分頁；新增 TELEC 認證模組選單與認證資訊
- **v2.3.0**：尺寸顯示、排序與篩選
- **v2.2.0**：Datasheet 按鈕
- **v2.1.0**：完全離線、安裝 App
- **v2.0.0**：強制更新、深色模式、中英文、版本資訊、下拉選單
- **v1.0.0**：初版
