# 版本更新流程（例如 Rev26.10 → Rev26.11）

只要提供**新版 Excel**，例如 `Quectel_Antenna_Product_and_TELEC_Certification_…_Rev26_11.xlsx`，就能依照這份流程完成所有更新。本流程整理自 Rev26.10 的實際更新經驗（2026-10-31）。

## 總覽

| 階段 | 誰 | 做什麼 | 產出 |
|---|---|---|---|
| 0 | Claude | 從 GitHub 下載最新版本，作為修改基礎 | 工作檔案 |
| 1 | Claude | `tools/release_prepare.py`：比對 PRO 頁面、寫入 Histories、產生 Windows 一鍵程式 | 更新後的 Excel、`release_RevYY.MM.json`、Windows 一鍵程式 zip |
| 2 | 使用者 | 在電腦上執行一鍵程式：複製新 datasheet、建立分享連結、收集檔案 | 桌面「RevYY.MM_for_Claude」資料夾：新 PDF＋連結 CSV |
| 3 | Claude | `tools/release_finalize.py`：合併連結、擷取照片、建置；匯出使用說明；打包 | PWA zip（完整 repo 檔案） |
| 4 | 使用者 | 把 zip 解壓縮後上傳到 GitHub | 工具上線 |

沒有新增天線時，可以略過階段 2，階段 3 只需要建置。

---

## 階段 0：同步 GitHub（每次修改前都要做）

```bash
git clone --depth 1 https://github.com/magicgrandpa8/Quectel-Antenna-Selector.git repo
```
- 以 GitHub 上的版本為準，因為使用者可能直接在 GitHub 上修改過，例如 Agent 連結。
- 如果使用者另外提供 zip，先比對它和 GitHub 的差異，以使用者指定的為準。
- **Excel 原始檔不在 repo 中**（公司內部資料），由使用者每次提供。

環境準備：
```bash
pip install pymupdf openpyxl pillow playwright --break-system-packages
python -m playwright install chromium
```

## 階段 1：比對與 Histories

```bash
python3 tools/release_prepare.py --xlsx <新版 Excel> --out <輸出資料夾> \
    --kit RevYY.MM_datasheet程式.zip --json release_RevYY.MM.json
```
**輸出檔名格式（固定）**：`Quectel Antenna Product and TELEC Certification_V3.71_20261029_Rev26.10.xlsx`
- 名稱部分保留空白；只有「名稱、版本、日期、版次」之間用底線 `_` 分隔。
- **不要把空白或點全部換成底線**：上傳的檔名常被系統改成 `V3_7_1`、`Rev26_10`，`--out` 指定資料夾時，程式會自動還原為 `V3.71`、`Rev26.10`（版次取自 PRO 頁面）。
- 建置工具時請使用這個標準檔名的 Excel，因為版本資訊視窗的「資料來源」會顯示檔名。
程式會自動完成以下工作：
1. **找出比對對象**：取 Excel 中最新的兩個 `PRO-YY.MM` 頁面，例如 PRO-26.11 和 PRO-26.10。
2. **判斷 TELEC 認證**：第 61 欄起是各模組的 TELEC 認證矩陣，任何一欄有標記，就算「有 TELEC」。
3. **比對結果**：
   - 新增、刪除的 OC，各自分成有 TELEC 和無 TELEC 兩組。
   - 原有 OC 從「無 TELEC → 有 TELEC」（寫入 Histories）。
   - 「有 TELEC → 無 TELEC」、原有天線新增的模組認證、TELEC 天線總數（只列在報告中）。
4. **交叉檢查**：找出模組工作表有列出、但 PRO 矩陣沒有標記的 **PRO 天線**。不在 PRO 清單裡的舊型號、-JP 客製版本會自動忽略。
5. **寫入 Histories**：寫在版次對應的那一列，例如 PRO-26.11 對應 Rev26.11；找不到時取日期最新的一列。
   - **欄位**：G～K 欄依序為 Added OC (TELEC)、Added OC (No TELEC)、Removed OC (TELEC)、Removed OC (No TELEC)、Existing OC: No TELEC → TELEC。
   - **表頭**：已有表頭就沿用；沒有就自動新增，並替所有資料列補上框線。
   - **內容**：OC 一行一個，沒有就填「—」；該列 A～F 欄改為靠上對齊。
6. **產生 Windows 一鍵程式**：只在有新增 OC 時產生，見階段 2。

**務必回報使用者**：比對摘要、交叉檢查結果，以及 Histories 寫在第幾列。

## 階段 2：使用者在電腦上執行一鍵程式

把階段 1 產生的 zip 交給使用者。使用者解壓縮後點兩下 `Run_datasheet_RevYY.MM.bat`，程式會：
1. **尋找 datasheet**：在 `C:\Users\gogol\OneDrive - 金碗國際股份有限公司\Antenna service\01 Antenna Datasheet`（含子資料夾）依料號尋找。
   - **選檔規則**：英文版（Datasheet）優先於中文規格書；正式版優先於 Preliminary；版本號取最新。
   - **相似料號**：比對時要求料號前後是 `_` 或 `&`，所以 YECM001L1AH 不會誤選到 YECM001L1AH-C。
2. **複製到客戶分享資料夾**：`Customer\Antenna FAE\0.For Customer\Quectel_Antenna_Selector\V3.7_20260916_Rev26.Q3`。已有的檔案會略過，不覆蓋。
3. **建立分享連結**：用使用者自己的帳號登入 Microsoft Graph，為每份檔案建立「**任何人可檢視**」連結。第一次執行會自動安裝登入模組。
4. **收集檔案**：把新 PDF 和 `datasheet_links_RevYY.MM.csv` 放到桌面的「RevYY.MM_for_Claude」資料夾。

使用者把該資料夾內的**所有檔案上傳到對話中**。

**為什麼要使用者執行**：
- **連接器限制**：Claude 的 Microsoft 365 連接器只有唯讀權限（Files.Read），無法複製檔案或建立分享連結，也讀不到 PDF 裡的圖片（只能讀文字）。
- **瀏覽器不一定可用**：Chrome 擴充功能不一定有連線，所以不能依賴網頁操作。

## 階段 3：完成更新

```bash
python3 tools/release_finalize.py --xlsx <階段 1 輸出的 Excel> --version 3.26.0 \
    --links datasheet_links_RevYY.MM.csv --pdf-dir <上傳的 PDF 資料夾> --diff release_RevYY.MM.json
python3 tools/export_guides.py          # 匯出四種語言的使用說明 PDF (會自己啟動網頁伺服器)
```
`release_finalize.py` 會依序完成：
1. **合併連結**：把連結合併進 `data/datasheet_links.csv`。同料號以新連結取代，空白連結不會覆蓋既有連結。
2. **擷取照片**：用 `tools/extract_datasheet_photos.py` 從 PDF 封面擷取照片。
3. **建置工具**：執行 `build_data.py`，自動採用最新的 PRO 頁面、ModuleList 和 Histories 的最新一筆。
4. **檢查新天線**：逐款列出 datasheet 照片、照片像素、Datasheet 按鈕，並回報缺少的項目。

**必做的人工檢查**：
- **照片對照圖**：產生新照片和原本 Excel 照片的對照圖，**逐張確認是同一個產品**。要特別留意評估板（EVB）、包裝圖、印字晶片；錯誤的料號要加入 `data/photo_datasheet_exclude.csv`。
- **資料差異**：和上一版工具的資料比對，向使用者回報認證組合數、規格欄位的變化。例如 Rev26.10 有 EG21-G 新增 6 筆認證，以及 4 款 GNSS 天線的安裝方式從 Buckle 改為 Screw。
- **功能測試**：天線總數、新天線可以搜尋、Datasheet 按鈕會開啟正確連結、比較功能、版本資訊顯示新版次。

**打包**：把 repo 的完整檔案（約 30 個）打包成 zip，排除 `tools/models/`、`datasheets/`。上傳前要和 GitHub 比對檔案清單，確認沒有需要刪除的檔案。

---

## 經驗與注意事項（Rev26.10 實際踩過的坑）

| 問題 | 正確做法 |
|---|---|
| openpyxl 存檔可能遺失 Excel 中的產品照片（約 500 張）與格式 | Histories 一律**直接修改工作表 XML**（`release_prepare.py` 已處理） |
| 重新封裝 xlsx 時沿用原本的 ZipInfo，會造成 CRC 錯誤 | 每個項目建立新的 ZipInfo（已處理） |
| 讀取時把工作表起始列算錯，造成列號偏移 | 以儲存格本身的列號為準（已處理） |
| Histories 由舊到新排列，最新一筆在最下面 | build 與 prepare 都取「最新日期或最後一筆」（已處理） |
| Excel 結構從 Antenna vs Cert / Name 改成 PRO-YY.MM / ModuleList | `build_data.py` 自動偵測，新舊格式都支援 |
| 模組工作表比 PRO 矩陣多出許多 OC | 多出來的大多不是 PRO 天線（舊型號、-JP、筆誤），只需要回報 PRO 天線的不一致 |
| 原有天線新增模組認證（例如 EG21-G）不等於「無 → 有 TELEC」 | 前者只列在報告中，Histories 的 K 欄只記錄 TELEC 狀態真正改變的天線 |
| 沒有提供 JP 天線表時，先前換過的清晰照片可能退回模糊 | `build_data.py` 會沿用上一版較清晰的照片（只限照片來源為 Excel 的天線） |
| OneDrive 搜尋索引常常查不到新檔案 | 以讀取資料夾清單為準（每頁 200 筆，依檔名排序） |
| 沙箱環境可能被重置 | 一切以 GitHub 為準；Excel 由使用者重新提供；套件需要重新安裝 |
