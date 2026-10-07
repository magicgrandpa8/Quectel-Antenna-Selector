# Quectel 天線選型 PWA

兩個頁面（同一個 App，hash 路由）：
- `index.html#/global`：Global 選型（Selector V3.7「PRO」+「Non-PRO」，371 項）
- `index.html#/japan`：日本市場選型（JP 天線表 + Selector「PRO-JP」模組相容勾選，91 項、47 個模組）

## 部署
將整個資料夾放到任一 HTTPS 靜態主機（IIS、Nginx、SharePoint 以外的內網 Web Server、GitHub Pages 等）。
Service Worker 只在 HTTPS 或 localhost 下啟用；啟用後可離線使用並可「安裝為 App」。

本機測試：
```bash
python3 -m http.server 8080   # 開啟 http://localhost:8080/index.html
```

## 更新資料
```bash
pip install openpyxl pillow
python3 build_data.py \
  --global-xlsx Quectel_Antenna_Product_Selector_V3_7_xxxxxxxx.xlsx \
  --jp-xlsx     Quectel_Antenna_Product_for_JP_xxxxxx.xlsx \
  --template template.html --out index.html
```
接著把 `sw.js` 的 `CACHE_VERSION` 加 1，使用者下次開啟即取得新版。

## 注意事項
- 刻意未匯入「Buy&Sell/Self-Production」欄與內部 SharePoint/OneDrive datasheet 連結。
  若要顯示 datasheet 按鈕，於 `template.html` 設定 `DATASHEET_URL`（含 `{oc}` 佔位符）後重新 build。
- Excel 需保留公式快取值（以 Excel 存檔過即可）；build 讀取的是 cached values。
- 頻段預設邊界為一般公開配置，僅供初篩。
