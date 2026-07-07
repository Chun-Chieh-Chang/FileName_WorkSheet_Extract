# 品檢報表統計 ETL Pipeline

自動化 ETL 管線，從 `RawData/` 目錄中的原始 Excel 品檢記錄提取、轉換、彙總資料，產出年度品檢統計 Excel 報表，並驅動 McKinsey 風格互動式儀表板（React SPA）與靜態 HTML 分析報告。

## 功能

- **多來源資料彙整**：自動掃描 RawData 目錄，辨識 7 大類 QC 檢驗記錄
  - 原物料進料品檢 (QC10002-R02)
  - QIP 尺寸檢驗 (QC10004-R02)
  - 裝配對樣巡檢 (QC10006-R01)
  - 半成品品檢 (QC10006-R02)
  - 完成品品檢 (QC10007-R01)
  - 零組件入庫品檢 (QC10007-R03)
  - 出貨檢驗 (QC10008-R02)
- **智能分類統計**：按月份、品項/類別自動統計筆數
- **空白樣板守衛**：自動識別並跳過批號為空值的空白預備樣板檔，防止未來月份幻象數據
- **McKinsey 互動儀表板**：React SPA，用戶可任意勾選品項、以月份為 X 軸生成堆疊柱狀圖，支援跨年度對比模式
- **跨年度對比**：可勾選多個年份進行數據比較，圖表 X 軸顯示年份，每個項目 1 根柱子
- **月份篩選**：下拉選單可篩選特定月份，圖表、KPI、數據表格即時更新
- **多檔案匯入**：支援一次選擇多個年份的 `.xlsx` 統計檔
- **品檢對照提取與過濾工具**：瀏覽器端 ETL，可匯入本地資料夾即時生成品檢對照 Excel。新增「檔案路徑」相對路徑顯示，新增「原因說明」顯示工作表是否被納入 ETL 計算的具體原因，並為提取表格的所有欄位配備如同 Excel 的獨立下拉篩選器及一鍵重設篩選功能，助於在海量品檢表中極速篩選與定位。
- **獨立報表匯出**：一次產出 16 個獨立 QC 表單 Excel 檔案
- **HTML 分析報告**：含統計圖表（橫向長條圖）之年報與跨年比較導覽頁
- **多年度支援**：2015–2027+（自動偵測年份，兼容年份後綴資料夾如 `原料-2020`）

## 使用方式

```bash
# 安裝依賴（僅首次）
npm install

# 執行 ETL（產出 DataExtract/ 目錄下的 Excel + HTML 報告）
node etl_pipeline.cjs 2025     # 僅處理 2025
node etl_pipeline.cjs 2026     # 僅處理 2026
node etl_pipeline.cjs all      # 同時處理 2025 + 2026 + 比較報告

# 啟動互動式儀表板（開發模式）
npm run dev

# 部署生產環境（輸出至 dist/）
npm run build
```

## ETL 執行流程

```
RawData/<year>/**/*.xlsx
        │
        ▼ etl_pipeline.cjs
        │  Step 1: Scan RawData/ → 統計各類別月份筆數
        │  Step 2: Write Summary Excel → DataExtract/<year>品檢報表統計.xlsx
        │          (並自動複製至 public/DataExtract/ 供 SPA 使用)
        │  Step 3: Spawn generate_styled_reports.cjs → <year>品檢報表分析.html
        ▼
DataExtract/<year>品檢報表統計.xlsx   ← React SPA 自動載入
<year>品檢報表分析.html              ← 靜態 HTML 年報
品檢報表比較分析.html                ← 跨年比較導覽頁（all 模式）
```

## 目錄結構

```
├── etl_pipeline.cjs              # ETL 主程式（Scan → Excel → HTML）
├── generate_styled_reports.cjs   # HTML 報告產生器（由 ETL 呼叫）
├── src/                          # React SPA 互動儀表板
│   ├── App.jsx                   # 主介面（McKinsey Dashboard + Extractor）
│   ├── index.css                 # 全域樣式（Outfit + Noto Sans TC 字型）
│   └── utils/
│       ├── db.js                 # QC 表單編號對照表（localStorage）
│       ├── excelParser.js        # 瀏覽器端 Excel 解析工具
│       └── browserETL.js         # 瀏覽器端 ETL：匯入資料夾即時生成品檢 Excel
├── public/DataExtract/           # 部署用：預建統計 Excel（GitHub Pages）
├── scratch/
│   └── cleanup.cjs               # 本地 MECE 清理工具（清 DataExtract、dist）
├── RawData/                      # 原始品檢 Excel 輸入檔（gitignored）
│   ├── 2025/
│   └── 2026/
├── DataExtract/                  # 產出 Excel（gitignored，公開版在 public/）
├── DEV_LOG.md                    # 開發日誌（含 RCA + CAPA）
├── package.json
└── .gitignore
```

## 技術棧

- **Node.js** + **SheetJS (xlsx)** — Excel 讀寫（伺服器端 ETL 與瀏覽器端）
- **React + Vite** — 前端 SPA 儀表板
- **Chart.js** — 互動式圖表（堆疊柱狀圖）
- **Google Fonts (Outfit + Noto Sans TC)** — 高級中英文字型
- **GitHub Pages** + **GitHub Actions** — 自動部署

## 輸出 Excel 報表結構

| 工作表 | 欄位結構 |
|---|---|
| 原物料品檢(QC10002-R02) | 月份 + 原料 + B膠/收縮膜/色粉/空白包裝袋/空白感壓紙/塑膠袋/塑膠袋40*50/紙箱/過濾網連蓋/標籤/射出D + 小計 |
| QIP(QC10004-R02) | 押出/射出 Setup + 押出/射出 巡檢（4欄 + 4欄 並列） |
| 裝配對樣巡檢(QC10006-R01) | 月份 + 裝配巡檢 |
| 半成品品檢(QC10006-R02) | 月份 + 裝配C + BD + Biometrix + MPS + Vivus |
| 完成品品檢(QC10007-R01 R02) | 月份 + Biometrix + MarMed + Saxon + Vivus |
| 零組件入庫品檢(QC10007-R03) | 月份 + Tubing + 射出 + 射出A + 射出C + 射出D(組件) + 裝配A + 裝配B + 裝配C |
| 出貨檢驗(QC10008-R02) | 月份 + ICU + 其他 |