# Quectel 天線選型 PWA / Quectel Antenna Selector PWA

目前版本：**v3.27.6**  
Current version: **v3.27.6**

資料來源：`Quectel_Antenna_Product_and_TELEC_Certification_*.xlsx`（單一 Excel）  
Data source: `Quectel_Antenna_Product_and_TELEC_Certification_*.xlsx` (single Excel workbook)

## 線材與接頭（Cable & Connector）

- 工具上方可切換「天線｜線材與接頭」兩個頁面；兩頁的篩選條件分開保存，天線頁面的資料與規則不受影響。  
  Switch between **Antennas** and **Cables & Connectors** at the top of the app. Each page saves its filters separately, so the antenna data and filtering rules are unaffected.
- 線材資料來源為獨立的 Excel（例如 `Quectel_Antenna_Product_Selector_V3_71_20261019.xlsx`）的「Cable&Connector」工作表，建置時以 `--cable-xlsx` 指定。  
  Cable data comes from the `Cable&Connector` worksheet in a separate Excel workbook, such as `Quectel_Antenna_Product_Selector_V3_71_20261019.xlsx`. Specify the workbook at build time with `--cable-xlsx`.
- 建置後線材資料（含縮圖與放大照片）存於 `data/cables.json`；之後只更新天線時不需再提供線材 Excel，會自動沿用。  
  The build saves cable data, including thumbnails and zoomable photos, in `data/cables.json`. Antenna-only updates automatically reuse this data without requiring the cable workbook again.
- 線材的 Datasheet 按鈕只使用 `data/datasheet_links.csv` 中「任何人可檢視」且經未登入驗證的連結；沒有連結時不顯示按鈕。  
  Cable **Datasheet** buttons use only links in `data/datasheet_links.csv` that are available to anyone and have been verified while signed out. If there is no link, the button is hidden.

## 功能 / Features

- **TELEC 認證模組選單**：依類別分組（LPWA、LTE-Cat1 BIS、LTE-Cat1/4、5G/LTE-A、Auto、Smart、Short Range），顯示認證日期與已認證天線數；選定後只列出與該模組一起取得 TELEC 認證的天線  
  **TELEC-certified module menu** groups modules by category (LPWA, LTE-Cat1 BIS, LTE-Cat1/4, 5G/LTE-A, Auto, Smart, and Short Range) and shows certification dates and certified antenna counts. Selecting a module shows only antennas certified with that module.
- 清單每列顯示「TELEC ×N」已認證模組數、JP 相容標示、尺寸、Datasheet 按鈕  
  Each result shows the number of certified modules as **TELEC ×N**, a Japan-market compatibility indicator, dimensions, and a Datasheet button.
- 主動式 GNSS 天線顯示 LNA 增益；IPEX MHF 與 RG174 / RG174LL 線材：連接器、線材型號、線長（IPEX 另含線徑）合併描述  
  Active GNSS antennas show LNA gain. For IPEX MHF and RG174 / RG174LL cables, the connector, cable type, and length appear together. IPEX entries also show cable diameter.
- 詳細資料列出所有認證模組（依類別、附認證日期）、射頻效能、完整規格  
  The detail view shows all certified modules grouped by category, along with certification dates, RF performance, and full specifications.
- 快速選單（由上到下）：Tech → Japan Market（Yes / No）→ TELEC 認證模組（僅 Japan Market = Yes 時出現，可多選：只列出全部通過認證的天線）→ GNSS 天線類型（僅 Tech = GNSS 時出現：Passive / Active / Dipole）→ Product Type → Form（型態）；側欄：安裝、連接器、IP、埠數、銷售區域、尺寸上限  
  Quick filters, from top to bottom, are **Tech** → **Japan Market** (Yes / No) → **TELEC-certified modules** (shown only when Japan Market is Yes, with multiple selections requiring certification with every selected module) → **GNSS antenna type** (Passive / Active / Dipole, shown only when Tech is GNSS) → **Product Type** → **Form**. Sidebar filters include mounting, connector, IP rating, port count, sales region, and maximum dimensions.
- 篩選：僅顯示有 TELEC 認證；日本市場相容改由 Japan Market 下拉選單篩選  
  Use the TELEC filter to show only certified antennas. Use the **Japan Market** dropdown to filter by Japan-market compatibility.
- 「移遠天線Agent」按鈕：位於頁面標題右側，開啟飛書 Agent [https://applink.feishu.cn/T9ao3zlnY4j5](https://applink.feishu.cn/T9ao3zlnY4j5)（`template.html` 的 `AGENT_URL`）  
  The **Quectel Antenna Agent** button to the right of the page title opens the [Feishu Agent](https://applink.feishu.cn/T9ao3zlnY4j5). Its URL is defined by `AGENT_URL` in `template.html`.
- 介面語言：中文 / English / 日本語 / 한국어（首次開啟預設為英文，之後記住使用者的選擇；預設值為 `template.html` 的 `DEFAULT_LANG`）  
  The interface is available in Traditional Chinese, English, Japanese, and Korean. It defaults to English on first launch and remembers the user's choice. The default is set by `DEFAULT_LANG` in `template.html`.
- 深色 / 淺色、版本資訊、強制更新  
  Dark and light modes, version information, and a manual update option.
- 完全離線、可安裝為 App（iOS / Android / Windows / macOS）  
  Works fully offline and can be installed as an app on iOS, Android, Windows, and macOS.
