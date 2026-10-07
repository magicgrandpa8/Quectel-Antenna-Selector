# Quectel 天線選型 · TELEC 認證 PWA

目前版本：**v3.0.0**
資料來源：`Quectel_Antenna_Product_and_TELEC_Certification_*.xlsx`（單一 Excel）

## 功能
- **TELEC 認證模組選單**：依類別分組（LPWA、LTE-Cat1 BIS、LTE-Cat1/4、5G/LTE-A、Auto、Smart、Short Range），顯示認證日期與已認證天線數；選定後只列出與該模組一起取得 TELEC 認證的天線
- 清單每列顯示「TELEC ×N」已認證模組數、JP 相容標示、尺寸、Datasheet 按鈕
- 詳細資料列出所有認證模組（依類別、附認證日期）、射頻效能、完整規格
- 頻譜尺頻段篩選（預設頻段 + 自訂 MHz）
- 快速選單：Type / Product Type / Antenna Type；側欄：外型、安裝、連接器、IP、埠數、銷售區域、尺寸上限
- 篩選：僅顯示有 TELEC 認證、僅顯示日本市場相容
- 中文 / English、深色 / 淺色、版本資訊、強制更新
- 完全離線、可安裝為 App（iOS / Android / Windows / macOS）

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

## 發佈新版本（Excel 每季更新時）
```bash
pip install openpyxl pillow
python3 build_data.py --version 3.0.1 \
  --xlsx ../Quectel_Antenna_Product_and_TELEC_Certification_V3_x_xxxxxxxx_RevXX_Qx_for_CU.xlsx
```
會同時產生 `index.html`、`version.json`，並自動更新 `sw.js` 的 `CACHE_VERSION`。
在 `template.html` 的 `CHANGELOG` 最上方加入說明後再執行一次，commit / push：
`index.html`、`version.json`、`sw.js`、`template.html`。

版本號：X 大改版 ／ Y 新功能 ／ Z 僅更新 Excel 資料或修正。

## 檔案說明
| 檔案 | 用途 | 部署需要 |
|---|---|---|
| `index.html` | App 本體（內嵌資料、圖片、字型） | ✅ |
| `manifest.webmanifest`、`icon-*.png`、`apple-touch-icon.png`、`screenshot-*.png` | 安裝資訊與圖示 | ✅ |
| `sw.js` | 離線快取 | ✅ |
| `version.json` | 新版偵測 | ✅ |
| `template.html`、`build_data.py`、`fonts/` | 產生 `index.html` 的原始碼 | 建議放在 repo |

## 注意事項
- 不要 commit 原始 `.xlsx`。
- Datasheet 連結不顯示在畫面上，但存在於 `index.html` 原始碼；公開 repo 時任何人都能看到這些 SharePoint 網址。
- 頻段預設邊界為一般公開配置，僅供初篩。

## 更新紀錄
- **v3.0.0**：改用 TELEC 認證 Excel 單一資料來源；移除 Global / 日本分頁；新增 TELEC 認證模組選單與認證資訊
- **v2.3.0**：尺寸顯示、排序與篩選
- **v2.2.0**：Datasheet 按鈕
- **v2.1.0**：完全離線、安裝 App
- **v2.0.0**：強制更新、深色模式、中英文、版本資訊、下拉選單
- **v1.0.0**：初版
