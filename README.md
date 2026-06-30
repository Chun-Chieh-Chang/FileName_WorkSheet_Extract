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
- **單一統合 layout**：原物料工作表採 14 欄單表（原料 + 12 子物料 + 小計）
- **彙總表**：3×3 並排網格，彙整全部檢驗類別的每月趨勢
- **McKinsey 互動儀表板**：React SPA，用戶可任意勾選品項、以月份為 X 軸生成堆疊柱狀圖
- **HTML 分析報告**：含統計圖表（橫向長條圖）之年報與跨年比較導覽頁
- **多年度支援**：2025、2026

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
│   ├── index.css                 # 全域樣式
│   └── utils/
│       ├── db.js                 # QC 表單編號對照表（localStorage）
│       └── excelParser.js        # 瀏覽器端 Excel 解析工具
├── public/DataExtract/           # 部署用：預建統計 Excel（GitHub Pages）
├── RawData/                      # 原始品檢 Excel 輸入檔（gitignored）
│   ├── 2025/
│   └── 2026/
├── DataExtract/                  # 產出 Excel（gitignored，公開版在 public/）
├── DEV_LOG.md                    # 開發日誌
├── package.json
└── .gitignore
```

## 技術棧

- **Node.js** + **SheetJS (xlsx)** — Excel 讀寫
- **React + Vite** — 前端 SPA 儀表板
- **Chart.js** — 互動式圖表（堆疊柱狀圖）
- **GitHub Pages** + **GitHub Actions** — 自動部署

## 輸出 Excel 報表結構

| 工作表 | 欄位結構 |
|---|---|
| 品檢地圖 | 各 QC 類別的 Sheet 索引 |
| 原物料品檢(QC10002-R02) | 月份 + 原料 + B膠/收縮膜/色粉/空白包裝袋/空白感壓紙/塑膠袋/塑膠袋40\*50/紙箱/過濾網連蓋/標籤/射出D + 小計（14欄單表） |
| QIP(QC10004-R02) | 押出/射出 Setup + 押出/射出 巡檢（4欄 + 4欄 並列） |
| 裝配對樣巡檢(QC10006-R01) | 月份 + 裝配巡檢（2欄） |
| 半成品品檢(QC10006-R02) | 月份 + 裝配C + BD + Biometrix + MPS + Vivus（6欄） |
| 完成品品檢(QC10007-R01 R02) | 月份 + Biometrix + MarMed + Saxon + Vivus（5欄） |
| 零組件入庫品檢(QC10007-R03) | 月份 + Tubing + 射出 + 射出A + 射出C + 射出D(組件) + 裝配A + 裝配B + 裝配C（9欄） |
| 出貨檢驗(QC10008-R02) | 月份 + ICU + 其他（3欄） |
| 彙總表 | 3×3 並排網格，全類別年度彙總 |
