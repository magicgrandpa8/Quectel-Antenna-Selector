# Quectel 天線選型 PWA

目前版本：**v2.3.0**

兩個頁面（同一個 App，hash 路由）：
- `index.html#/global`：Global 選型（Selector「PRO」+「Non-PRO」）
- `index.html#/japan`：日本市場選型（JP 天線表 + Selector「PRO-JP」模組相容勾選）

## 功能
- 頻譜尺頻段篩選（預設頻段 + 自訂 MHz 範圍）
- 快速下拉選單：Type / Product Type / Antenna Type（日本頁另含 Quectel 模組）
- 側欄篩選：外型、安裝方式、連接器、防護等級等
- 中文 / English 切換、深色 / 淺色模式切換
- 版本資訊視窗（App 版本、伺服器版本、資料建置時間、更新紀錄）
- 強制更新：清除 Service Worker 與快取後重新載入；偵測到伺服器新版時按鈕會出現紅點
- 尺寸 Dimensions (mm)：清單每列顯示；可依尺寸排序、以「最長邊 ≤ X mm」篩選
- Datasheet 按鈕：清單每一列與詳細規格皆可開啟，畫面不顯示網址（需網路與公司帳號登入）
- 完全離線：不使用任何外部網域資源（字型、資料、圖片全部內嵌），首次開啟即快取全部檔案
- 可安裝為 App：iOS / Android / Windows / macOS；App 內有「安裝 App」按鈕與各平台說明
- 離線狀態標示；強制更新會先確認伺服器可連線，離線時不會清除快取

## 安裝為 App
| 平台 | 瀏覽器 | 步驟 |
|---|---|---|
| iPhone / iPad | **Safari**（必須） | 分享 → 加入主畫面 → 新增 |
| Android | Chrome / Edge / Samsung Internet | 點 App 內「安裝 App」，或選單 ⋮ → 安裝應用程式 |
| Windows | Edge / Chrome | 點 App 內「安裝 App」，或網址列右側安裝圖示；安裝後可從開始功能表開啟、釘選到工作列 |
| macOS | Chrome / Edge；Safari 17+ | Chrome/Edge 同 Windows；Safari：檔案 → 加入 Dock |

安裝後**在有網路時開啟一次**（右上角版本資訊中「離線快取」顯示已啟用），之後即可完全離線使用。

**iOS 注意**：若 App 連續數週未開啟，iOS 可能清除離線資料，屆時連網開啟一次即可恢復。

## 部署 (GitHub Pages)
將本資料夾所有檔案放在 repo 根目錄 → Settings → Pages → Deploy from a branch → `main` / `(root)`。

## 發佈新版本
```bash
pip install openpyxl pillow
python3 build_data.py --version 2.1.0 \
  --global-xlsx ../Quectel_Antenna_Product_Selector_Vx_x.xlsx \
  --jp-xlsx     ../Quectel_Antenna_Product_for_JP_xxxx.xlsx
```
此指令會同時：
- 產生 `index.html`（內嵌資料、版本號、建置時間）
- 產生 `version.json`（App 用來偵測新版）
- 自動將 `sw.js` 的 `CACHE_VERSION` 改為 `antenna-selector-v2.1.0`

接著在 `template.html` 的 `CHANGELOG` 最上方加入新版說明，重新執行一次上面的指令，再 commit / push：
`index.html`、`version.json`、`sw.js`、`template.html`。

版本號規則（X.Y.Z）：
- **X**：介面或資料結構大改
- **Y**：新增功能
- **Z**：只更新 Excel 資料或修正錯誤

## 檔案說明
| 檔案 | 用途 | 部署時需要 |
|---|---|---|
| `index.html` | App 本體（內嵌資料、圖片、字型） | ✅ |
| `manifest.webmanifest`、`icon-*.png`、`apple-touch-icon.png`、`screenshot-*.png` | 安裝資訊與圖示 | ✅ |
| `sw.js` | Service Worker（離線快取） | ✅ |
| `version.json` | 新版偵測 | ✅ |
| `template.html`、`build_data.py`、`fonts/` | 產生 `index.html` 用的原始碼 | 建議一併放在 repo |

## 注意事項
- 不要 commit 原始 `.xlsx`。
- 未匯入「Buy&Sell/Self-Production」欄。
- Datasheet 連結取自 Excel：Global 用 Selector「PRO」的 Datasheet link，日本頁用 JP 表的 Datasheet Link，兩邊互相補齊。
  沒有連結的料號不顯示按鈕；也可在 `template.html` 設定 `DATASHEET_URL`（含 `{oc}`）作為備援樣板。
- **連結雖不顯示在畫面上，但仍存在於 `index.html` 原始碼中。** 若 repo / 網站為公開，任何人都能看到這些 SharePoint 網址（檔案本身仍需公司帳號登入才能開啟）。
- Excel 需保留公式快取值（以 Excel 存檔過即可）。
- 頻段預設邊界為一般公開配置，僅供初篩。

## 更新紀錄
- **v2.3.0**：清單顯示尺寸 Dimensions (mm)、依尺寸排序、尺寸上限篩選、尺寸格式統一
- **v2.2.0**：新增 Datasheet 按鈕（清單與詳細規格，不顯示網址）
- **v2.1.0**：完全離線（移除外部字型、全部檔案預先快取）、安裝 App 按鈕與各平台說明、離線狀態標示、強制更新離線保護、永久儲存
- **v2.0.0**：強制更新按鈕與新版偵測、深色/淺色切換、中英文切換、版本資訊、Type / Product Type / Antenna Type 選單
- **v1.0.0**：初版
