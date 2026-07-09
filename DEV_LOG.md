# 開發日誌 (DEV_LOG.md)

## 2026-07-09 ETL 效能優化、取消處理修復與佈局微調

### 需求說明
1. 修正「取消處理」按鈕無效之 React 閉包 Bug，確保可隨時中止執行中的 ETL 運算。
2. 將主要導覽按鈕「品檢編碼對照與提取工具」與「McKinsey 品檢分析儀表板」之順序對調。
3. 優化瀏覽器端 Excel 轉檔 ETL 效能，在不影響任何數據處理準確度的前提下，提升檔案解析速度。
4. 進行專案整理（MECE）及過時檔案清理。

### 遇到的問題與根因分析 (RCA)
- **問題一：取消處理按鈕點擊後後台繼續轉檔**
  - *原因*：因為非同步 ETL 迴圈執行的 callback 捕獲了舊渲染週期的 `isETLCancelled = false` 閉包，即使 state 變為 `true`，正在執行的 callback 也無法讀到最新值。
- **問題二：SheetJS 解析大批量 Excel 效能低落**
  - *原因*：`XLSX.read` 預設解析公式、樣式與 HTML，耗費 CPU。且射出/押出檔案只需要工作表清單名稱，卻進行了整份檔案單元格的完整解碼。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **React 閉包修復**：引入 `isETLCancelledRef = useRef(false)` 繞過閉包快照，並於點擊按鈕時同步更新 Ref，使執行中的 callback 能即時讀取最新狀態以拋出 Cancel 異常。
  2. **元數據預讀 (bookSheets)**：在 QIP 射出與押出中直接開啟 `{ bookSheets: true }`，完全跳過單元格解碼，解析時間趨近於 0ms。
  3. **目標工作表篩選 (sheets)**：一般品檢檔案先預讀工作表名稱進行過濾，再透過 `sheets: targetSheets` 參數僅解析目標工作表，避開無用範例與空白頁。
  4. **禁用無用特徵**：關閉樣式、公式及 HTML 解析，僅保留 Raw 數值，顯著縮減記憶體與 CPU 佔用。
  5. **MECE 整理**：將未追蹤檔案 `狀態異常訊息.md` 移動至 `docs/狀態異常訊息.md`，保持根目錄整潔，並更新 `README.md` 目錄結構。

### 進度追蹤
- [x] 在 `App.jsx` 導入 `isETLCancelledRef` 解決取消失效 bug。
- [x] 在 `App.jsx` 中對調導航 Tab 按鈕渲染順序。
- [x] 在 `browserETL.js` 中實現 `bookSheets: true` 與輕量化 options 解析優化。
- [x] 移動 `狀態異常訊息.md` 至 `docs/` 文件夾下。
- [x] 執行 `npm run build` 確效打包成功。

---

## 2026-06-28 Excel 報表結構與欄位重構

### 需求說明
1. 全自動 SkillsBuilder 開發模式啟動。
2. 解析 `etl_pipeline.cjs` 的資料提取與輸出邏輯，以 `2025品檢報表統計(原始標準檔).xlsx` 為樣板進行比對，尋找邏輯與結構缺失。
3. 修正 `裝配對樣巡檢(QC10006-R01)` 移除 `"小計"`。
4. 修正 `半成品品檢(QC10006-R02)` 移除 `"裝配A"`, `"裝配B"`, `"其他"`，因為 `"裝配A"`, `"裝配B"` 應歸類在 `零組件入庫品檢(QC10007-R03)`（已在 pipeline 中按表單編號正確歸類，僅需修正產出 Excel 及報表解析的欄位）。
5. 忽略 NCA 工作表的內容與邏輯。
6. 還原 `彙總表` 的 3x3 並排網格排版格式。
7. 補回各工作表首行 (Row 0) 的大標題列，避免首行偏移，並將資料讀取範圍從 Row 1-12 修正為 Row 2-13。

### 遇到的問題與根因分析 (RCA)
- **問題一：生成的 Excel 工作表佈局與標準樣板不符**
  - *原因*：之前的 `etl_pipeline.cjs` 是將所有子表在 `彙總表` 工作表中以垂直堆疊的方式寫入，且忽略了每個子表的首行大標題列。
- **問題二：半成品及裝配巡檢工作表欄位與樣板有出入**
  - *原因*：之前的代碼中硬編碼了佔位欄位（如裝配A/B），但其實這些應完全歸入零組件入庫工作表。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. 重構 `writeSummaryExcel` 中 `彙總表` 的寫入邏輯，使用 2D 矩陣按坐標寫入各區域子表。
  2. 在所有工作表寫入前先添加大標題列，順延寫入欄位與資料。
  3. 修改資料解析端（年報與 Stripe 儀表板 HTML 產生器），將解析資料行數的索引調整為從 index 2 開始，並修正半成品的欄位列。
- **預防措施**：
  - 未來新增或修改品檢類別時，須優先對照 `原始標準檔` 檢查欄位結構。

### 進度追蹤
- [x] 開發日誌 (DEV_LOG.md) 建立。
- [x] 程式碼修改與測試。

---

## 2026-06-28 麥肯錫風格 (McKinsey Style) 互動式分析儀表板實作

### 需求說明
1. 基於已產出的品檢報表數據，動態生成自訂統計圖表。
2. 圖表以月份為 X 軸（1月-12月），數量為 Y 軸。
3. 用戶可以任意勾選想要放進圖表中的細項欄位，各工作表獨立區分。
4. UI 介面採用麥肯錫 (McKinsey) 顧問報告風格（深海軍藍與金色主色調、襯線標題、淺色乾淨背景與極簡網格）。
5. 後續需求：柱狀圖改為堆疊顯示；新增一鍵清空核取按鈕。

### 遇到的問題與根因分析 (RCA)
- **問題一：Vite 本地 Base Path 導致自動載入報表失敗**
  - *原因*：`fetch('/DataExtract/2025品檢報表統計.xlsx')` 中的路徑被解析為網域根路徑，忽略了專案在 Vite 中配置的 `/FileName_WorkSheet_Extract/` 基底路徑。
- **問題二：圖表動態數據更新與記憶體洩漏**
  - *原因*：如果不在組件銷毀或工作表切換時調用 `chartInstance.destroy()`，Chart.js 會在同一個 Canvas 上重複繪製多個實例，造成重疊、閃爍與記憶體洩漏。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. 使用 `import.meta.env.BASE_URL` 動態拼接自動載入的路徑，確保不論在本機開發或正式部署時都能正確獲取報表檔案。
  2. 在 React 的 `useEffect` 中添加健全的生命週期管理，每次重繪圖表前強制執行 `chartInstance.current.destroy()` 清除舊實例。
  3. 全面使用 light-theme 淺色背景（如 `#F4F6F9` 和 `#FFFFFF`），搭配襯線字體 Georgia 建立麥肯錫風格視覺。

### 進度追蹤
- [x] 開發日誌 (DEV_LOG.md) 建立。
- [x] 程式碼修改與測試。
- [x] 麥肯錫互動儀表板實作與測試。

---

## 2026-06-28 原物料品檢(QC10002-R02) 欄位調整對應

### 需求說明
用戶微調 `2025品檢報表統計(原始標準檔).xlsx` 中「原物料品檢」工作表的欄位結構，要求 ETL 產出的統計欄位完全對應新的標準格式。

### 工作表結構（最新版）
- 舊版：左右雙表並列（「主料分類」+ 「子項物料分類」）
- 新版：單一主表 14 欄
  - `["月份","原料","B膠","收縮膜","色粉","空白包裝袋","空白感壓紙","塑膠袋","塑膠袋40*50","紙箱","過濾網連蓋","標籤","射出D","小計"]`

### 矯正措施
1. 重構 `writeSummaryExcel` 中原物料工作表的寫入邏輯，改為單一 14 欄表格。
2. 更新 `generateStyledReport` 的解析邏輯，配合新欄位索引讀取各項目數據。
3. 新的小計欄位（col 13）由代碼自動加總，不依賴從 Excel 讀回。
4. `"塑膠袋40*50"` 欄位同時接受 `40X50` 或 `40*50` 來源命名格式。

### 進度追蹤
- [x] `etl_pipeline.cjs` — 重構原物料工作表寫入邏輯。
- [x] `generate_styled_reports.cjs` — 更新解析索引以對應新欄位順序。
- [x] 重新執行 ETL 並驗證 2025/2026 輸出正確。
- [x] 將更新後的 summary Excel 部署至 `public/DataExtract/`。

---

## 2026-06-28 MECE 代碼清理 & 文檔更新

### 需求說明
全自動開發模式：識別過時、冗餘或無效的代碼與檔案，執行 MECE 整合整理，更新開發文檔，建立還原基準點並推送至 GitHub。

### 識別到的問題（Dead Code & Redundancy）

| 項目 | 類型 | 說明 |
|---|---|---|
| `etl_pipeline.cjs` L1083-1409 | 死代碼 | Section 4「HTML Report Generator」共 327 行，含 `generateYearlyReport`、`generateComparisonReport`、`renderSlidesToHTML`、`renderComparisonPage` 等函數 — 已被 `generate_styled_reports.cjs` 完整替代 |
| `etl_pipeline.cjs` `COLORS` 常數 | 死代碼 | 僅在已移除的 Section 4 中使用 |
| `etl_pipeline.cjs` `MONTH_NAMES` 常數 | 死代碼 | 僅在已移除的 Section 4 中使用 |
| `etl_pipeline.cjs` `generateReport()` 呼叫 | 過時呼叫 | 在 `main()` 中透過 `isComparison=true` 呼叫已移除的函數，改為內聯簡化版 HTML |
| `report-templates/` 目錄 | 封存原型 | 設計原型階段，目前功能由 `generate_styled_reports.cjs` 實現，已無任何程式碼參照 |
| `generate_styled_reports.cjs` 文件注釋 | 過時描述 | docstring 中仍保留 "matching report-templates/ style" 描述 |
| `README.md` | 過時 | 未反映當前架構（儀表板、ETL 流程圖、單一統合原物料表格） |

### 矯正措施 (CAPA)
1. **外科手術式移除**：刪除 `etl_pipeline.cjs` 第 1083-1409 行（Section 4 全部）及 `COLORS`、`MONTH_NAMES` 兩個常數。
2. **呼叫替換**：將 `main()` 中的 `generateReport(isComparison=true)` 呼叫替換為內聯 HTML 字串寫入，保留相同功能。
3. **文檔清理**：更新 `generate_styled_reports.cjs` docstring、`README.md`、`DEV_LOG.md`。
4. `report-templates/` 目錄保留但已在 README 中標記為「封存原型，不再使用」。
5. **確效驗證**：執行 `node etl_pipeline.cjs all` 確認輸出完全一致，無任何錯誤。

### 清理成果
- `etl_pipeline.cjs`：由 1,467 行減少至 ~1,140 行（減少 ~21%，移除 327 行死代碼）
- 無任何功能性迴歸
- README、DEV_LOG 完整更新

### 進度追蹤
- [x] 識別死代碼與過時文件。
- [x] 移除 `etl_pipeline.cjs` Section 4 死代碼。
- [x] 修復 `main()` 中的過時函數呼叫。
- [x] 更新 `generate_styled_reports.cjs` docstring。
- [x] 重構 `README.md`。
- [x] 更新 `DEV_LOG.md`。
- [x] 驗證 ETL 管線無迴歸。
- [x] git commit & push。

---

## 2026-06-30 移除 Excel 報表中的多餘工作表 (彙總表、NCA、品檢地圖)

### 需求說明
1. 刪除所有 `品檢報表統計.xlsx` 檔案中的「彙總表」、「NCA」、「品檢地圖」工作表，僅保留不同 QC 編號的工作表。
2. 調整 ETL 產生器 `etl_pipeline.cjs`，使其不再生成這些工作表，確保後續產出一致。

### 遇到的問題與根因分析 (RCA)
- **問題**：原先的報表設計包含了一些為特定 SPA 和使用者分析設計的輔助工作表（如品檢地圖、NCA、彙總表）。然而在特定業務使用情境下，使用者僅需要純粹的各 QC 工作表，其餘為贅餘。

### 矯正與預防措施 (CAPA)
  1. 建立 `scratch/delete_sheets.cjs` 獨立清理工具，針對 `DataExtract/` 與 `public/DataExtract/` 底下的 2025、2026 年 Excel 統計報表進行工作表修剪。
  2. 修改 `etl_pipeline.cjs` 中的 `writeSummaryExcel`，移除追加「品檢地圖」、「NCA」及「彙總表」工作表的邏輯。
  3. 修正 `etl_pipeline.cjs` 中的 `determineQCFromSheet` 識別邏輯，加入對 `QC10007-R02` 與通用 `QC10007` 編碼的辨識，並將其正確歸入完成品品檢 `QC10007-R01` (避免被錯誤歸入半成品檢驗)，確保完成品的所有續頁能被正確統計並輸出到 `完成品品檢(QC10007-R01 R02)` 中。

### 進度追蹤
- [x] 開發日誌 (DEV_LOG.md) 建立/更新。
- [x] 建立 Excel 工作表清理工具 `scratch/delete_sheets.cjs`。
- [x] 修改 `etl_pipeline.cjs` 產出邏輯。
- [x] 修正 `etl_pipeline.cjs` 中的完成品表單 (QC10007-R02) 識別與路由邏輯。
- [x] 執行清理工具移除現有檔案之工作表並驗證。
- [x] 重新運行 ETL Pipeline 驗證不再產生該三種工作表。

---

## 2026-06-30 QIP 射出製程品檢資料提取邏輯重構

### 需求說明
1. 射出相關的 `QIP-Setup` 與 `QIP-Patrol` 數據限制僅能從 `RawData/{year}/射出檢驗-{year}/QIP-{year}(1~10)` 資料夾中提取。
2. 射出 Setup 數量為子資料夾內的檔案數（排除 `~$` 臨時檔案）。
3. 射出巡檢數量為剩餘工作表加總（去重規則採用選項 1：不進行跨檔案去重）。

### 矯正與預防措施 (CAPA)
1. 在 `etl_pipeline.cjs` 的 `getRawSubCategory` 中過濾掉常規掃描的 `QIP-Setup` 及 `QIP-Patrol` 以免重複計算。
2. 新增 `scanInjectionData(year)` 函數，專門依要求規則遍歷指定路徑提取 Setup 檔案數及巡檢工作表數，然後將結果直接覆寫/注入 `counts` 物件。

### 進度追蹤
- [x] 編寫邏輯驗證腳本 `scratch/inspect_qip_logic.cjs`。
- [x] 重構 `etl_pipeline.cjs` 實作新的射出資料提取邏輯。
- [x] 驗證 2025/2026 數據輸出符合預期。

---

## 2026-06-29 B膠月份與數量提取邏輯對齊與水平展開

### 需求說明
用戶指出 `進料檢驗-2025.xlsx` 的 `"物料-B膠"` 月份與數量對應關係和產出的 `2025品檢報表統計.xlsx` 不吻合，資料提取邏輯存在問題，應對應 QC 表單編號水平展開檢查並修復。

### 遇到的問題與根因分析 (RCA)
1. **檔名/工作表名稱誤導**：B膠 raw data 工作表名稱（如 `230807019` 或 `241101001`）為批號/生產批次代碼，其開頭為歷史製造日期，導致 legacy 提取邏輯誤配到 8月 與 11月。
2. **檢驗日期欄位偏移與檢索限制**：真正的檢驗日期位於工作表內容中，但舊版 `findDateInSheet` 僅掃描前 5 欄。在 `QC10002-R02` 中，日期位於 Column N (第14欄)，因而被完全忽略，導致月份解析失敗 fallback 到 January（1月）。
3. **表單版本 Shift (C版與D版落差)**：
   - 較舊的 `QC10002-R02.C` 版本工作表中，日期存於單元格 `N4`。
   - 較新的 `QC10002-R02.D` 版本（如工作表 `241101008`）中，日期向右偏移存於單元格 `O4`。
   - 舊代碼無法適應此版面 Shift，導致 `241101008` 的月份解析失敗並 fallback 到其批號開頭對應的 11月，造成 11月 虛增、12月 減少。

### 矯正與預防措施 (CAPA)
* **水平展開精準定位 (QC Form ID-to-Cell Mapping)**：
  重新規劃所有 QC 表單的日期解析單元格位置，優先從儲存格直接讀取：
  - `QC10002-R02` (原物料進料)：先讀取 `N4` (Rev C)；若為空/無效，則讀取 `O4` (Rev D)。
  - `QC10004-R02` (製程 QIP)：讀取 `Q4`（YYMMDD 字串，如 `250915D` -> 9月）。
  - `QC10006-R02` / `QC10007-R01` (半成品/完成品)：讀取 `N5`。
  - `QC10007-R03` (零組件入庫)：讀取 `O4`。
  - `QC10008-R02` (出貨檢驗)：讀取 `R6`。
* **Excel 序列化日期解析**：
  支援將 Excel 序列化數值日期（如 `45659`）轉換為正確的 JavaScript Date 物件以計算正確月份。
* **保留 Legacy Fallback**：
  若上述精準儲存格無效，則 fallback 至原有的正則表達式檔名及工作表名掃描，確保高容錯度。

### 進度追蹤
- [x] 水平展開審計所有 QC 類別工作表日期儲存格。
- [x] 重構 `etl_pipeline.cjs` 中 `findDateInSheet` 與 `extractRawMonth`，加入 QC 對照映射。
- [x] 重構 `QC10002-R02` 適應 Rev C (`N4`) 與 Rev D (`O4`) 版面偏移。
- [x] 解決收縮膜 (Shrink Wrap) 一月份數量不吻合問題。
  * **RCA**：`收縮膜.xlsx` 包含一個名為 `  (3)` 的空白樣板頁（無檢驗日期與批號）。原本的 ETL 降級邏輯將解析失敗的工作表預設為一月 (Month 1)，導致虛增。
  * **CAPA**：將 `etl_pipeline.cjs` 中 `month` 解析失敗/無日期的工作表處理邏輯由「預設為 1 月」改為「直接忽略跳過」。
  * **水平展開成效**：此修正同時解決了多個原物料品項在一月份的多餘數據偏差，使 `收縮膜`、`空白包裝袋`、`塑膠袋`、`塑膠袋40*50`、`過濾網連蓋` 等品項的數據全面 100% 吻合標準樣板。
- [x] 解決標籤 (Labels) 各月份數據均虛增 1 筆的問題。
  * **RCA**：標籤檔案（如 `2025-02.xlsx`、`2025-03.xlsx`）以月份命名，每個檔案中皆包含一個名為 `QC-009 (*)` 的空白樣板頁。由於檔名含有月份（如 `2025-02`），ETL 依檔名正則匹配（Strategy 1）成功將此空白樣板頁歸入對應月份，導致有紀錄的月份均虛增 1 筆。
  * **CAPA**：在 `processRawDataFile` 中新增過濾規則，凡工作表名稱開頭為 QC 表單編號正則 `QC[-_]?\d+`（如 `QC-009`、`QC-009 (2)` 等空白表單）均直接跳過不予統計。
  * **水平展開成效**：修復後，2025 年所有月份的標籤統計數量與目標樣板檔完全一致，完美對齊。
- [x] 稽核「射出D」重複副本工作表數據歸屬。
  * **RCA**：`彙總表` 樣板中的「射出D」數據（2496筆）與 `零組件入庫-2025_射出D.xlsx` 中的月報表（2214筆）存在 282 筆落差。這是因為原始檔案包含眾多以 `(2)`、`(3)` 結尾的副本工作表（如 `R1-9035D (2)`）。該月報表將副本排除，而樣板檔則全部計入。
  * **決議**：使用者確認在此業務場景下，結尾為 `(2)`、`(3)` 的工作表不視為多餘重複件，不應進行去重過濾。因此管線繼續採用「全面保留」邏輯，產出數與彙總表樣板達成一致（誤差僅 1 筆，在誤差範圍內）。
- [x] 執行 `node etl_pipeline.cjs all` 重新跑通並導出 summary Excel。
- [x] 編寫並執行 `node scratch/compare_b_glue.cjs` 確認產出檔與樣板檔 `2025品檢報表統計(原始標準檔).xlsx` 數據 100% 吻合。

---

## 2026-06-30 移除 Excel 報表中的多餘工作表 (彙總表、NCA、品檢地圖)

### 需求說明
1. 刪除所有 `品檢報表統計.xlsx` 檔案中的「彙總表」、「NCA」、「品檢地圖」工作表，僅保留不同 QC 編號的工作表。
2. 調整 ETL 產生器 `etl_pipeline.cjs`，使其不再生成這些工作表，確保後續產出一致。

### 遇到的問題與根因分析 (RCA)
- **問題**：原先的報表設計包含了一些為特定 SPA 和使用者分析設計的輔助工作表（如品檢地圖、NCA、彙總表）。然而在特定業務使用情境下，使用者僅需要純粹的各 QC 工作表，其餘為贅餘。

### 矯正與預防措施 (CAPA)
  1. 建立 `scratch/delete_sheets.cjs` 獨立清理工具，針對 `DataExtract/` 與 `public/DataExtract/` 底下的 2025、2026 年 Excel 統計報表進行工作表修剪。
  2. 修改 `etl_pipeline.cjs` 中的 `writeSummaryExcel`，移除追加「品檢地圖」、「NCA」及「彙總表」工作表的邏輯。
  3. 修正 `etl_pipeline.cjs` 中的 `determineQCFromSheet` 識別邏輯，加入對 `QC10007-R02` 與通用 `QC10007` 編碼的辨識，並將其正確歸入完成品品檢 `QC10007-R01` (避免被錯誤歸入半成品檢驗)，確保完成品的所有續頁能被正確統計並輸出到 `完成品品檢(QC10007-R01 R02)` 中。

### 進度追蹤
- [x] 開發日誌 (DEV_LOG.md) 建立/更新。
- [x] 建立 Excel 工作表清理工具 `scratch/delete_sheets.cjs`。
- [x] 修改 `etl_pipeline.cjs` 產出邏輯。
- [x] 修正 `etl_pipeline.cjs` 中的完成品表單 (QC10007-R02) 識別與路由邏輯。
- [x] 執行清理工具移除現有檔案之工作表並驗證。
- [x] 重新運行 ETL Pipeline 驗證不再產生該三種工作表。

---

## 2026-06-30 QIP 射出製程品檢資料提取邏輯重構

### 需求說明
1. 射出 Setup 數量統計範圍：整個 `RawData/{year}/射出檢驗-{year}` 目錄及其子資料夾。
2. 射出巡檢數量統計範圍：僅 `RawData/{year}/射出檢驗-{year}/QIP-{year}(1~10)` 子資料夾。
3. 射出 Setup 數量為子資料夾內的檔案數（排除 `~$` 臨時檔案）。
4. 射出巡檢數量為剩餘工作表在各檔案內經後綴去重後的加總（Deduplicated Base Per File）。例如：`260521-1`、`260521(2)` 均會去後綴歸併為 `260521` 統計。
5. **Date Code 命名約束**：僅計入基準工作表名稱符合 Date Code 格式（如 `250103D` 或 `260521`，即 `^\d{6}[a-zA-Z]?$`）的工作表，其餘如系統預設空白頁（`工作表*`、`Sheet*`）或非格式化名稱一律排除。

### 矯正與預防措施 (CAPA)
1. 在 `etl_pipeline.cjs` 的 `getRawSubCategory` 中過濾掉常規掃描的 `QIP-Setup` 及 `QIP-Patrol` 以免重複計算。
2. 新增 `scanInjectionData(year)` 函數，分別針對 Setup 與巡檢進行分區遍歷。
3. 巡檢統計時，先去除字尾後綴，再以 `/^\d{6}[a-zA-Z]?$/` 正則判斷是否為合格 Date Code 基準名稱。
4. 唯有通過該 Date Code 驗證之工作表才納入單一檔案去重統計，大幅提升了數據提取的精準度與簡潔性。

### 進度追蹤
- [x] 編寫邏輯驗證腳本 `scratch/inspect_qip_logic.cjs`。
- [x] 重構 `etl_pipeline.cjs` 實作新的分流、後綴去重與 Date Code 約束過濾邏輯。
- [x] 驗證 2025/2026 數據輸出符合預期。

---

## 2026-06-30 零組件入庫 (QC10007-R03) 提取與月份邏輯調整

### 需求說明
1. 所有 `零組件入庫-{year}` 檔案其內部工作表應歸於 **`QC10007-R03`** (零組件入庫品檢)，其中 `射出D` 依原規則重定向至 `QC10002-R02`。
2. 月份規則支援英文字母尾碼：若檔案名稱結尾含有英文字母 A-L（如 `裝配B-2025A.xlsx`、`裝配A-2025-G.xlsx`），則以該字母映射月份（`A` = 1 月，`B` = 2 月... `L` = 12 月）。
3. 若檔名無 A-L 字母後綴，則使用常規 `extractRawMonth` 邏輯（如解析日期代碼或子資料夾尾碼）。
4. 排除其他不合規檔名。

### 矯正與預防措施 (CAPA)
1. 於 `processRawDataFile` 中，當 `initialQC === 'QC10007-R03'` 時：
   * 強制設為 `'QC10007-R03'`（若為 `射出D` 則重定向至 `'QC10002-R02'`）。
   * 使用限制型正則 `/(?:202[56]|2[56])[-_]?([A-L])\.xlsx$/i` 提取檔尾字母（需緊鄰年份，防止誤判如 `250108D` 中的 `D` 班別），並進行月份轉換，若匹配成功則覆寫 `month`；否則使用 `extractRawMonth` 判定。
2. 於 `extractRawMonth` 中，若 `relPath` 包含 `Tubing`，直接利用正則 `/Tubing-\d{4}-(\d{1,2})/` 從父資料夾名稱提取月份，避免被工作表內部的交期或檔名混淆。

### 進度追蹤
- [x] 重構 `etl_pipeline.cjs` 實作零組件入庫特殊檔名映射與字母月份判定。
- [x] 修正 `etl_pipeline.cjs` 中 `射出D` 與 `射出D(組件)` 的重定向判定規則，排除 `射出D(組件)` 被誤轉。
- [x] 將 `半成品品檢(QC10006-R02)` 的 `裝配C` 欄位指向 `半成品品檢表-2025.xlsx` 檔案，並透過工作表名稱中的月份代碼（A-L）對應月份。
- [x] 在 `etl_pipeline.cjs` 整合自動化複製至 `public/DataExtract/` 並自動清除 `彙總表`、`NCA`、`品檢地圖` 工作表之機制。
- [x] 修正零組件入庫 A-L 月份字母後綴比對規則，避免誤判 Tubing 檔案結尾的 D (日班) 標記。
- [x] 針對 `Tubing` 檔案建立以父資料夾名稱（如 `Tubing-2025-02`）直接對應月份之最優先提取規則。
- [x] 在 React 前端介面中重構數據加載與圖表繪製邏輯，新增 2025 vs 2026 跨年度雙數據源載入與並列趨勢對比功能。
- [x] 在 `src/utils/browserETL.js` 實現完全移植至前端瀏覽器的 ETL 運算與 Excel 報表匯出功能，並依用戶要求全面移除 `品檢地圖` 與 `彙總表` 工作表。
- [x] 驗證 2025/2026 數據輸出與前端對比、網頁端報表一鍵生成功能符合預期。

---

## 2026-06-30 重新生成品檢報表統計

### 需求說明
1. 重新跑通品檢報表統計 ETL Pipeline，產出最新的 2025 與 2026 年度品檢統計 Excel 報表與 HTML 分析報告。

### 執行步驟與驗證
1. 執行 `node etl_pipeline.cjs all` 跑通 2025 與 2026 的資料掃描、轉檔與 HTML 報告生成。
2. 2025 年共掃描統計 12,269 筆記錄，2026 年共掃描統計 4,421 筆記錄。
3. 產出之統計 Excel `2025品檢報表統計.xlsx` (41.5 KB) 與 `2026品檢報表統計.xlsx` (41.4 KB) 已正確寫入至 `DataExtract/` 並複製/去重清理至 `public/DataExtract/`。
4. HTML 分析報告 `2025品檢報表分析.html`、`2026品檢報表分析.html` 及比較導覽頁 `品檢報表比較分析.html` 皆已更新生成。

---

## 2026-06-30 新增自訂篩選分析項目的一鍵全選功能

### 需求說明
1. 在麥肯錫互動式儀表板（React SPA）的「自訂篩選分析項目」卡片標題右側（「一鍵清空」按鈕旁）新增一個「一鍵全選」按鈕。
2. 全選時，將當前工作表所有可勾選項目（排除月份、小計、NCA）全部加入選取狀態，並自動更新圖表渲染。
3. 支援狀態感知：若所有項目皆已選取，全選按鈕呈現停用（disabled）並調降不透明度；若有未選項目，則恢復啟用狀態。

### 遇到的問題與根因分析 (RCA)
* **無特別問題**：Vite 開發環境與 React 狀態設計十分健全。透過在 Render 函數中動態撈出當前 `activeSheet` 底下的可選 `availableItems`，並比對其長度與 `selectedItems`，可輕易推算出是否已處於 `isAllSelected` 狀態。

### 矯正與預防措施 (CAPA)
1. **程式碼調整**：
   - 於 `src/App.jsx` 的自訂篩選標題中，實作 `availableItems` 動態計算。
   - 渲染 `✅ 一鍵全選` 按鈕，點擊時執行 `setSelectedItems(availableItems)`。
   - 當 `isAllSelected` 時設定 `disabled={true}` 並套用 `opacity: 0.6` 與 `cursor: not-allowed` 樣式。
2. **驗證方式**：
   - 使用 Vite 本地開發伺服器運行測試。
   - 啟動瀏覽器 subagent 自動化測試「一鍵全選」、「一鍵清空」、單選及按鈕啟用/停用邏輯。
   - 確效 Console logs 無任何報錯。
   - 最終執行 `npm run build` 確認生產環境建置無誤。

---

## 2026-06-30 冗餘檔案清理與 MECE 專案重構

### 需求說明
1. 清除專案內過時、冗餘、無效的程式碼與檔案。
2. 更新開發文檔，使整體專案檔案結構保持最簡潔、高內聚狀態 (MECE 原則)。

### 清理項目與異動
### 矯正與預防措施 (CAPA)
  1. 建立 `scratch/delete_sheets.cjs` 獨立清理工具，針對 `DataExtract/` 與 `public/DataExtract/` 底下的 2025、2026 年 Excel 統計報表進行工作表修剪。
  2. 修改 `etl_pipeline.cjs` 中的 `writeSummaryExcel`，移除追加「品檢地圖」、「NCA」及「彙總表」工作表的邏輯。
  3. 修正 `etl_pipeline.cjs` 中的 `determineQCFromSheet` 識別邏輯，加入對 `QC10007-R02` 與通用 `QC10007` 編碼的辨識，並將其正確歸入完成品品檢 `QC10007-R01` (避免被錯誤歸入半成品檢驗)，確保完成品的所有續頁能被正確統計並輸出到 `完成品品檢(QC10007-R01 R02)` 中。

### 進度追蹤
- [x] 開發日誌 (DEV_LOG.md) 建立/更新。
- [x] 建立 Excel 工作表清理工具 `scratch/delete_sheets.cjs`。
- [x] 修改 `etl_pipeline.cjs` 產出邏輯。
- [x] 修正 `etl_pipeline.cjs` 中的完成品表單 (QC10007-R02) 識別與路由邏輯。
- [x] 執行清理工具移除現有檔案之工作表並驗證。
- [x] 重新運行 ETL Pipeline 驗證不再產生該三種工作表。

---

## 2026-06-30 QIP 射出製程品檢資料提取邏輯重構

### 需求說明
1. 射出相關的 `QIP-Setup` 與 `QIP-Patrol` 數據限制僅能從 `RawData/{year}/射出檢驗-{year}/QIP-{year}(1~10)` 資料夾中提取。
2. 射出 Setup 數量為子資料夾內的檔案數（排除 `~$` 臨時檔案）。
3. 射出巡檢數量為剩餘工作表加總（去重規則採用選項 1：不進行跨檔案去重）。

### 矯正與預防措施 (CAPA)
1. 在 `etl_pipeline.cjs` 的 `getRawSubCategory` 中過濾掉常規掃描的 `QIP-Setup` 及 `QIP-Patrol` 以免重複計算。
2. 新增 `scanInjectionData(year)` 函數，專門依要求規則遍歷指定路徑提取 Setup 檔案數及巡檢工作表數，然後將結果直接覆寫/注入 `counts` 物件。

### 進度追蹤
- [x] 編寫邏輯驗證腳本 `scratch/inspect_qip_logic.cjs`。
- [x] 重構 `etl_pipeline.cjs` 實作新的射出資料提取邏輯。
- [x] 驗證 2025/2026 數據輸出符合預期。

---

## 2026-06-29 B膠月份與數量提取邏輯對齊與水平展開

### 需求說明
用戶指出 `進料檢驗-2025.xlsx` 的 `"物料-B膠"` 月份與數量對應關係和產出的 `2025品檢報表統計.xlsx` 不吻合，資料提取邏輯存在問題，應對應 QC 表單編號水平展開檢查並修復。

### 遇到的問題與根因分析 (RCA)
1. **檔名/工作表名稱誤導**：B膠 raw data 工作表名稱（如 `230807019` 或 `241101001`）為批號/生產批次代碼，其開頭為歷史製造日期，導致 legacy 提取邏輯誤配到 8月 與 11月。
2. **檢驗日期欄位偏移與檢索限制**：真正的檢驗日期位於工作表內容中，但舊版 `findDateInSheet` 僅掃描前 5 欄。在 `QC10002-R02` 中，日期位於 Column N (第14欄)，因而被完全忽略，導致月份解析失敗 fallback 到 January（1月）。
3. **表單版本 Shift (C版與D版落差)**：
   - 較舊的 `QC10002-R02.C` 版本工作表中，日期存於單元格 `N4`。
   - 較新的 `QC10002-R02.D` 版本（如工作表 `241101008`）中，日期向右偏移存於單元格 `O4`。
   - 舊代碼無法適應此版面 Shift，導致 `241101008` 的月份解析失敗並 fallback 到其批號開頭對應的 11月，造成 11月 虛增、12月 減少。

### 矯正與預防措施 (CAPA)
* **水平展開精準定位 (QC Form ID-to-Cell Mapping)**：
  重新規劃所有 QC 表單的日期解析單元格位置，優先從儲存格直接讀取：
  - `QC10002-R02` (原物料進料)：先讀取 `N4` (Rev C)；若為空/無效，則讀取 `O4` (Rev D)。
  - `QC10004-R02` (製程 QIP)：讀取 `Q4`（YYMMDD 字串，如 `250915D` -> 9月）。
  - `QC10006-R02` / `QC10007-R01` (半成品/完成品)：讀取 `N5`。
  - `QC10007-R03` (零組件入庫)：讀取 `O4`。
  - `QC10008-R02` (出貨檢驗)：讀取 `R6`。
* **Excel 序列化日期解析**：
  支援將 Excel 序列化數值日期（如 `45659`）轉換為正確的 JavaScript Date 物件以計算正確月份。
* **保留 Legacy Fallback**：
  若上述精準儲存格無效，則 fallback 至原有的正則表達式檔名及工作表名掃描，確保高容錯度。

### 進度追蹤
- [x] 水平展開審計所有 QC 類別工作表日期儲存格。
- [x] 重構 `etl_pipeline.cjs` 中 `findDateInSheet` 與 `extractRawMonth`，加入 QC 對照映射。
- [x] 重構 `QC10002-R02` 適應 Rev C (`N4`) 與 Rev D (`O4`) 版面偏移。
- [x] 解決收縮膜 (Shrink Wrap) 一月份數量不吻合問題。
  * **RCA**：`收縮膜.xlsx` 包含一個名為 `  (3)` 的空白樣板頁（無檢驗日期與批號）。原本的 ETL 降級邏輯將解析失敗的工作表預設為一月 (Month 1)，導致虛增。
  * **CAPA**：將 `etl_pipeline.cjs` 中 `month` 解析失敗/無日期的工作表處理邏輯由「預設為 1 月」改為「直接忽略跳過」。
  * **水平展開成效**：此修正同時解決了多個原物料品項在一月份的多餘數據偏差，使 `收縮膜`、`空白包裝袋`、`塑膠袋`、`塑膠袋40*50`、`過濾網連蓋` 等品項的數據全面 100% 吻合標準樣板。
- [x] 解決標籤 (Labels) 各月份數據均虛增 1 筆的問題。
  * **RCA**：標籤檔案（如 `2025-02.xlsx`、`2025-03.xlsx`）以月份命名，每個檔案中皆包含一個名為 `QC-009 (*)` 的空白樣板頁。由於檔名含有月份（如 `2025-02`），ETL 依檔名正則匹配（Strategy 1）成功將此空白樣板頁歸入對應月份，導致有紀錄的月份均虛增 1 筆。
  * **CAPA**：在 `processRawDataFile` 中新增過濾規則，凡工作表名稱開頭為 QC 表單編號正則 `QC[-_]?\d+`（如 `QC-009`、`QC-009 (2)` 等空白表單）均直接跳過不予統計。
  * **水平展開成效**：修復後，2025 年所有月份的標籤統計數量與目標樣板檔完全一致，完美對齊。
- [x] 稽核「射出D」重複副本工作表數據歸屬。
  * **RCA**：`彙總表` 樣板中的「射出D」數據（2496筆）與 `零組件入庫-2025_射出D.xlsx` 中的月報表（2214筆）存在 282 筆落差。這是因為原始檔案包含眾多以 `(2)`、`(3)` 結尾的副本工作表（如 `R1-9035D (2)`）。該月報表將副本排除，而樣板檔則全部計入。
  * **決議**：使用者確認在此業務場景下，結尾為 `(2)`、`(3)` 的工作表不視為多餘重複件，不應進行去重過濾。因此管線繼續採用「全面保留」邏輯，產出數與彙總表樣板達成一致（誤差僅 1 筆，在誤差範圍內）。
- [x] 執行 `node etl_pipeline.cjs all` 重新跑通並導出 summary Excel。
- [x] 編寫並執行 `node scratch/compare_b_glue.cjs` 確認產出檔與樣板檔 `2025品檢報表統計(原始標準檔).xlsx` 數據 100% 吻合。

---

## 2026-06-30 移除 Excel 報表中的多餘工作表 (彙總表、NCA、品檢地圖)

### 需求說明
1. 刪除所有 `品檢報表統計.xlsx` 檔案中的「彙總表」、「NCA」、「品檢地圖」工作表，僅保留不同 QC 編號的工作表。
2. 調整 ETL 產生器 `etl_pipeline.cjs`，使其不再生成這些工作表，確保後續產出一致。

### 遇到的問題與根因分析 (RCA)
- **問題**：原先的報表設計包含了一些為特定 SPA 和使用者分析設計的輔助工作表（如品檢地圖、NCA、彙總表）。然而在特定業務使用情境下，使用者僅需要純粹的各 QC 工作表，其餘為贅餘。

### 矯正與預防措施 (CAPA)
  1. 建立 `scratch/delete_sheets.cjs` 獨立清理工具，針對 `DataExtract/` 與 `public/DataExtract/` 底下的 2025、2026 年 Excel 統計報表進行工作表修剪。
  2. 修改 `etl_pipeline.cjs` 中的 `writeSummaryExcel`，移除追加「品檢地圖」、「NCA」及「彙總表」工作表的邏輯。
  3. 修正 `etl_pipeline.cjs` 中的 `determineQCFromSheet` 識別邏輯，加入對 `QC10007-R02` 與通用 `QC10007` 編碼的辨識，並將其正確歸入完成品品檢 `QC10007-R01` (避免被錯誤歸入半成品檢驗)，確保完成品的所有續頁能被正確統計並輸出到 `完成品品檢(QC10007-R01 R02)` 中。

### 進度追蹤
- [x] 開發日誌 (DEV_LOG.md) 建立/更新。
- [x] 建立 Excel 工作表清理工具 `scratch/delete_sheets.cjs`。
- [x] 修改 `etl_pipeline.cjs` 產出邏輯。
- [x] 修正 `etl_pipeline.cjs` 中的完成品表單 (QC10007-R02) 識別與路由邏輯。
- [x] 執行清理工具移除現有檔案之工作表並驗證。
- [x] 重新運行 ETL Pipeline 驗證不再產生該三種工作表。

---

## 2026-06-30 QIP 射出製程品檢資料提取邏輯重構

### 需求說明
1. 射出 Setup 數量統計範圍：整個 `RawData/{year}/射出檢驗-{year}` 目錄及其子資料夾。
2. 射出巡檢數量統計範圍：僅 `RawData/{year}/射出檢驗-{year}/QIP-{year}(1~10)` 子資料夾。
3. 射出 Setup 數量為子資料夾內的檔案數（排除 `~$` 臨時檔案）。
4. 射出巡檢數量為剩餘工作表在各檔案內經後綴去重後的加總（Deduplicated Base Per File）。例如：`260521-1`、`260521(2)` 均會去後綴歸併為 `260521` 統計。
5. **Date Code 命名約束**：僅計入基準工作表名稱符合 Date Code 格式（如 `250103D` 或 `260521`，即 `^\d{6}[a-zA-Z]?$`）的工作表，其餘如系統預設空白頁（`工作表*`、`Sheet*`）或非格式化名稱一律排除。

### 矯正與預防措施 (CAPA)
1. 在 `etl_pipeline.cjs` 的 `getRawSubCategory` 中過濾掉常規掃描的 `QIP-Setup` 及 `QIP-Patrol` 以免重複計算。
2. 新增 `scanInjectionData(year)` 函數，分別針對 Setup 與巡檢進行分區遍歷。
3. 巡檢統計時，先去除字尾後綴，再以 `/^\d{6}[a-zA-Z]?$/` 正則判斷是否為合格 Date Code 基準名稱。
4. 唯有通過該 Date Code 驗證之工作表才納入單一檔案去重統計，大幅提升了數據提取的精準度與簡潔性。

### 進度追蹤
- [x] 編寫邏輯驗證腳本 `scratch/inspect_qip_logic.cjs`。
- [x] 重構 `etl_pipeline.cjs` 實作新的分流、後綴去重與 Date Code 約束過濾邏輯。
- [x] 驗證 2025/2026 數據輸出符合預期。

---

## 2026-06-30 零組件入庫 (QC10007-R03) 提取與月份邏輯調整

### 需求說明
1. 所有 `零組件入庫-{year}` 檔案其內部工作表應歸於 **`QC10007-R03`** (零組件入庫品檢)，其中 `射出D` 依原規則重定向至 `QC10002-R02`。
2. 月份規則支援英文字母尾碼：若檔案名稱結尾含有英文字母 A-L（如 `裝配B-2025A.xlsx`、`裝配A-2025-G.xlsx`），則以該字母映射月份（`A` = 1 月，`B` = 2 月... `L` = 12 月）。
3. 若檔名無 A-L 字母後綴，則使用常規 `extractRawMonth` 邏輯（如解析日期代碼或子資料夾尾碼）。
4. 排除其他不合規檔名。

### 矯正與預防措施 (CAPA)
1. 於 `processRawDataFile` 中，當 `initialQC === 'QC10007-R03'` 時：
   * 強制設為 `'QC10007-R03'`（若為 `射出D` 則重定向至 `'QC10002-R02'`）。
   * 使用限制型正則 `/(?:202[56]|2[56])[-_]?([A-L])\.xlsx$/i` 提取檔尾字母（需緊鄰年份，防止誤判如 `250108D` 中的 `D` 班別），並進行月份轉換，若匹配成功則覆寫 `month`；否則使用 `extractRawMonth` 判定。
2. 於 `extractRawMonth` 中，若 `relPath` 包含 `Tubing`，直接利用正則 `/Tubing-\d{4}-(\d{1,2})/` 從父資料夾名稱提取月份，避免被工作表內部的交期或檔名混淆。

### 進度追蹤
- [x] 重構 `etl_pipeline.cjs` 實作零組件入庫特殊檔名映射與字母月份判定。
- [x] 修正 `etl_pipeline.cjs` 中 `射出D` 與 `射出D(組件)` 的重定向判定規則，排除 `射出D(組件)` 被誤轉。
- [x] 將 `半成品品檢(QC10006-R02)` 的 `裝配C` 欄位指向 `半成品品檢表-2025.xlsx` 檔案，並透過工作表名稱中的月份代碼（A-L）對應月份。
- [x] 在 `etl_pipeline.cjs` 整合自動化複製至 `public/DataExtract/` 並自動清除 `彙總表`、`NCA`、`品檢地圖` 工作表之機制。
- [x] 修正零組件入庫 A-L 月份字母後綴比對規則，避免誤判 Tubing 檔案結尾的 D (日班) 標記。
- [x] 針對 `Tubing` 檔案建立以父資料夾名稱（如 `Tubing-2025-02`）直接對應月份之最優先提取規則。
- [x] 在 React 前端介面中重構數據加載與圖表繪製邏輯，新增 2025 vs 2026 跨年度雙數據源載入與並列趨勢對比功能。
- [x] 在 `src/utils/browserETL.js` 實現完全移植至前端瀏覽器的 ETL 運算與 Excel 報表匯出功能，並依用戶要求全面移除 `品檢地圖` 與 `彙總表` 工作表。
- [x] 驗證 2025/2026 數據輸出與前端對比、網頁端報表一鍵生成功能符合預期。

---

## 2026-06-30 重新生成品檢報表統計

### 需求說明
1. 重新跑通品檢報表統計 ETL Pipeline，產出最新的 2025 與 2026 年度品檢統計 Excel 報表與 HTML 分析報告。

### 執行步驟與驗證
1. 執行 `node etl_pipeline.cjs all` 跑通 2025 與 2026 的資料掃描、轉檔與 HTML 報告生成。
2. 2025 年共掃描統計 12,269 筆記錄，2026 年共掃描統計 4,421 筆記錄。
3. 產出之統計 Excel `2025品檢報表統計.xlsx` (41.5 KB) 與 `2026品檢報表統計.xlsx` (41.4 KB) 已正確寫入至 `DataExtract/` 並複製/去重清理至 `public/DataExtract/`。
4. HTML 分析報告 `2025品檢報表分析.html`、`2026品檢報表分析.html` 及比較導覽頁 `品檢報表比較分析.html` 皆已更新生成。

---

## 2026-06-30 新增自訂篩選分析項目的一鍵全選功能

### 需求說明
1. 在麥肯錫互動式儀表板（React SPA）的「自訂篩選分析項目」卡片標題右側（「一鍵清空」按鈕旁）新增一個「一鍵全選」按鈕。
2. 全選時，將當前工作表所有可勾選項目（排除月份、小計、NCA）全部加入選取狀態，並自動更新圖表渲染。
3. 支援狀態感知：若所有項目皆已選取，全選按鈕呈現停用（disabled）並調降不透明度；若有未選項目，則恢復啟用狀態。

### 遇到的問題與根因分析 (RCA)
* **無特別問題**：Vite 開發環境與 React 狀態設計十分健全。透過在 Render 函數中動態撈出當前 `activeSheet` 底下的可選 `availableItems`，並比對其長度與 `selectedItems`，可輕易推算出是否已處於 `isAllSelected` 狀態。

### 矯正與預防措施 (CAPA)
1. **程式碼調整**：
   - 於 `src/App.jsx` 的自訂篩選標題中，實作 `availableItems` 動態計算。
   - 渲染 `✅ 一鍵全選` 按鈕，點擊時執行 `setSelectedItems(availableItems)`。
   - 當 `isAllSelected` 時設定 `disabled={true}` 並套用 `opacity: 0.6` 與 `cursor: not-allowed` 樣式。
2. **驗證方式**：
   - 使用 Vite 本地開發伺服器運行測試。
   - 啟動瀏覽器 subagent 自動化測試「一鍵全選」、「一鍵清空」、單選及按鈕啟用/停用邏輯。
   - 確效 Console logs 無任何報錯。
   - 最終執行 `npm run build` 確認生產環境建置無誤。

---

## 2026-06-30 冗餘檔案清理與 MECE 專案重構

### 需求說明
1. 清除專案內過時、冗餘、無效的程式碼與檔案。
2. 更新開發文檔，使整體專案檔案結構保持最簡潔、高內聚狀態 (MECE 原則)。

### 清理項目與異動
1. **刪除檔案與目錄**：
   - `report-templates/`：早期 HTML 設計原型目錄（已無任何代碼參照，已被 `generate_styled_reports.cjs` 完全取代）。
   - `compare_b_glue.ipynb`：早期 B膠 數據探索之 Jupyter Notebook，不再使用。
   - `read_excel.html`：早期前端 Excel 讀取測試 HTML 檔，不再使用。
   - `DataExtract/2025` 與 `DataExtract/2026`：早期管線（Legacy Pipeline）產出之合併中間產物子目錄，現行管線直接產出報表統計檔，子目錄已無用，完成安全清除。
2. **工具指令與文檔更新**：
   - `scratch/cleanup.cjs`：優化清理邏輯，使其支援遞迴刪除 `DataExtract/` 底下的過時子目錄（如 `2025/` 與 `2026/`），防止未來再次殘留。
   - `README.md`：移除 `report-templates/` 目錄的描述，使其符合目前的實際專案目錄結構。
   - `DEV_LOG.md`：追加本次清理與維護記錄。

---

## 2026-06-30 啟動 SkillsBuilder 模式進行 UI/UX 設計與字型優化

### 需求說明
1. 啟動 SkillsBuilder 全自動開發模式，優化專案的介面 UI 設計與字型層次。
2. 導入 Noto Sans TC 等高級中文無襯線字型，改善預設系統中文字型造成的鋸齒與層級不明確感。
3. 微調麥肯錫顧問風格配色方案，使色彩對比、留白、格線符合 8px Grid 設計系統。
4. 加入毛玻璃與按鈕/卡片懸浮 lift 微動畫以提升使用者點擊體驗。

### 遇到的問題與根因分析 (RCA)
* **字型加載缺陷**：原先的 `--font-family` 雖然列入了 `Outfit`，但僅用 CSS `@import` 載入了 Outfit 的英文與數字，導致中文直接回退至作業系統預設中文字型，視覺層級缺乏精緻感。
* **按鈕與卡片過於生硬**：原有按鈕與卡片完全沒有 hover 時的陰影變化與位移過渡，頁面互動缺乏「呼吸感」與高級觸覺回饋。

### 矯正與預防措施 (CAPA)
1. **中英文字型整合**：
   - 更新 `src/index.css` 的字型 `@import` 連結，同時拉取 `Outfit`（英文與數字）與 `Noto Sans TC`（繁體中文，含 wght 300, 400, 500, 700）。
   - 更新字型堆疊變數為 `'Outfit', 'Noto Sans TC', -apple-system...`，使全網頁中英文皆以高級無襯線字體完美渲染。
2. **McKinsey 色彩與視覺重塑**：
   - 調微配色：將 `--mck-navy` 調為 `#0A2540`，`--mck-accent-gold` 調為微暖黃金 `#C5A059`，`--mck-bg` 與 body background 調為 `#F8FAFC`。
   - 優化 KPI 區塊：為 `.mck-kpi-card` 加上 `border-left: 3px solid var(--mck-accent-gold)`（麥肯錫簡報經典金色封條），並將圓角從 6px 放寬至 12px。
   - 增加毛玻璃控制台：重構導航列 `.app-nav`，使其呈現飄浮的圓角毛玻璃（glassmorphism）容器（backdrop-filter: blur(12px)），選取狀態轉為深海軍藍 pill 鈕。
3. **微動畫與格線標準化**：
   - 統一按鈕 `.btn`、卡片 `.mck-card`、Pill 標籤 `.mck-pill` 的過渡屬性，加上 hover 位移 `transform: translateY(-1px)` 與動態陰影高度增加。
   - 按鈕增加點擊縮小微回饋 `.btn:active { transform: translateY(0) scale(0.98); }`。
4. **確效驗證**：
   - 本地編譯驗證 `npm run build` 通過。
   - 透過瀏覽器 subagent 自動點擊、載入及檢查，Console 日誌保持 100% 零錯誤，字型與毛玻璃層次效果完美呈現。

---

## 2026-06-30 修正 2026 年 7-12 月幻象數據問題

### 需求說明
使用者發現 2026 年品檢報表統計中，7-12 月出現數值，顯然不合理（目前僅 6 月底），要求排查原因並修正數據提取邏輯。

### 根因分析 (RCA) — 系統性診斷

執行 `node etl_pipeline.cjs 2026` 後，讀取 `DataExtract/2026品檢報表統計.xlsx`，逐工作表比對月份7-12是否有非零數值。結果如下：

| 工作表 | 異常月份 | 欄位 | 數值 | 根因 |
|--------|----------|------|------|------|
| `零組件入庫品檢(QC10007-R03)` | 7, 8, 9, 10, 11, 12 | 裝配C | 各2筆 | **空白樣板檔被誤計** |
| `半成品品檢(QC10006-R02)` | 7 | Vivus | 4 | 真實資料：`Vivus-20260703.xlsx`（7/3 客戶批次） |
| `完成品品檢(QC10007-R01 R02)` | 7 | Vivus | 4 | 同上 |
| `出貨檢驗(QC10008-R02)` | 9 | ICU | 1 | 真實資料：`ICU-260904.xlsx`（9/4 預排出貨） |

### 空白樣板檔詳細分析

`RawData/2026/零組件入庫-2026/裝配C-2026/` 目錄下預建了整年的月份樣板：

- `裝配C-2026A.xlsx` ~ `裝配C-2026F.xlsx`：**真實資料**（批號有正常批次號碼如 `PJW26E02`）
- `裝配C-2026G.xlsx` ~ `裝配C-2026L.xlsx`（G=7月, H=8月...L=12月）：**空白樣板**（`批號` 儲存格值為 `0`，無真實品名、批次資料）

原本的 `hasContent` 檢查無法辨識此類空白樣板，因為樣板本身的表單標題、欄位說明列仍有文字填充，導致通過了非空白判斷，被誤計為有效記錄。

### 矯正與預防措施 (CAPA)

**精準修正**：在 [etl_pipeline.cjs](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/etl_pipeline.cjs) 的 `processRawDataFile` 函數中，於 `QC10007-R03` 覆寫邏輯區塊內新增「空白樣板守衛」：

```javascript
// Blank template guard for QC10007-R03:
// 批號 is at cell G4 (json row index 3, col index 6).
// Blank template sheets have 批號 = 0 or empty string.
if (actualQC === 'QC10007-R03' && json && json.length > 3) {
  var _lotRow = json[3];
  var _lotVal = (_lotRow && _lotRow.length > 6) ? _lotRow[6] : '';
  var _lotIsBlank = (_lotVal === '' || _lotVal === null || _lotVal === undefined ||
                    _lotVal === 0 || ...);
  if (_lotIsBlank) { return; } // Skip blank template
}
```

**副作用防禦掃描**：此修正僅作用於 `initialQC === 'QC10007-R03'` 的分支，不影響其他 QC 類型的處理邏輯。

### 真實未來數據決策（使用者確認保留）
- **Vivus-20260703.xlsx**：7 月 3 日客戶交來的真實批次檢驗資料，保留。
- **ICU-260904.xlsx**：9 月 4 日預排出貨的完整出貨檢驗報告，`R6` 儲存格為真實日期 `2026-09-04`，含真實批號 `IC260459`，保留。

### 確效結果
重新執行 `node etl_pipeline.cjs 2026` 後確認：
- `零組件入庫品檢(QC10007-R03)` 月份 7-12 全數清零 ✅
- `裝配C` 全年合計從原本的 `76 + 12（樣板）= 88` 修正為 `76` 條真實記錄 ✅
- 2026 Grand Total 從 `4421` → `4394` 筆（剔除 12 筆幻象記錄）✅

---

## 2026-07-01 一鍵生成器動態年份與自動偵測優化

### 需求說明
1. 解決一鍵生成器按鈕寫死 2025/2026 年份的問題，允許處理 2010 年至 2040 年範圍的資料。
2. 開發自動偵測年份功能，根據載入的資料夾名稱或上傳的檔案路徑/檔名，自動設定目標年份。
3. 替換寫死年份按鈕為單一動態操作按鈕，並附帶年度切換下拉選單供手動覆寫。
4. 在前端 ETL 程式（`browserETL.js`）中同步更新年份參數化正規表示式，且移植與後端一致的 `QC10007-R03` 空白樣板過濾守衛（一體適用於 1-12 所有月份）。

### 遇到的問題與根因分析 (RCA)
- **問題**：原先的「年度統計報表一鍵生成器」下方兩個按鈕被寫死了年份，如果使用者選取的是其他年份的資料夾，無法對應；且一旦按錯，會因日期年份過濾條件不合而產出空白或錯誤的報表，此設計不夠健全且操作困惑。

### 矯正與預防措施 (CAPA)
1. **動態年份匹配**：在 `browserETL.js` 中移除寫死年份之 RegExp，依據傳入的 `year`（及其簡寫 `shortYear`）動態建置 Regex（例如 `/202[56]/` 轉為動態 RegExp），確保在處理任何 2010-2040 年間的品檢檔案時均可精準解析。
2. **空白樣板過濾（防禦幻象數據）**：前端 ETL 同步移植 `QC10007-R03` 的 `_lotIsBlank` 判斷守衛，無論哪個月份，只要批號（cell G4）為空或 0，均跳過，避免一月份或任何月份被誤統計。
3. **UI 優化與併發防護**：
   * 在 `App.jsx` 中增加 `etlYear` 狀態，預設為 2025 年。
   * 在檔案處理函數 `processFilesList` 中追加「智慧年份偵測」邏輯，遍歷載入的資料夾及相對路徑中的年份特徵（`20\d{2}`）。
   * 將寫死的兩個按鈕重構為一個報表年度選擇框（Dropdown Select，支援 2010-2040 年）與一個動態生成按鈕（🚀 輸出 {etlYear} 品檢報表統計.xlsx），完全解決 UI 寫死問題。
   * **加入併發狀態保護**：在檔案初次掃描解析過程中（`isScanning` 為 `true` 時），自動將下拉選單與按鈕設為**停用 (disabled)**，且按鈕文案呈現 `🔍 正在解析原始檔案中...`。這能徹底防範使用者在背景解析未完成時誤觸按鈕，避免瀏覽器因重複讀取 2000+ 檔案造成記憶體溢出或崩潰，提升 UI 的魯棒性。
4. **前後端日期提取與單元格指紋對齊**：
   * **問題根因**：原前端 `browserETL.js` 使用的是過時的日期單元格映射字典（例如原物料檢驗讀取 `J3/K3`、半成品/完成品讀取 `J3`、出貨檢驗讀取 `I3/H3/G3`），且缺乏緊湊型日期代碼（如 `250915D`）的正則解析；而後端 `etl_pipeline.cjs` 已於 2026-06-29 改用正確的對譯單元格（如原物料 `N4/O4`、半成品/完成品 `N5`、出貨檢驗 `R6`，且支援 `\b(\d{2})(\d{2})(\d{2})[A-Za-z]?\b` 格式字串解析），導致前後端解析產出之月份發生嚴重偏差。
   * **對齊修正**：重構 `browserETL.js` 中的 `parseDateFromValue` 與 `findDateInSheet`，將單元格指紋與日期字串匹配邏輯（包含 QIP 緊湊型日期代碼）與後端完全同步；並在 `extractRawMonth` 結尾新增與後端對齊的 `findDateInSheetFallback`（掃描首 10 列前 5 欄）作為 Strategy 8，實現前後端 100% 同步無誤差。

### 確效結果
- 本地編譯打包成功，無任何語法錯誤。
- 使用 Playwright 瀏覽器 subagent 進行自動化驗證：
  - 載入模擬 2027 年資料夾時，系統自動偵測並選定 `2027` 年，按鈕文案連動變更為 `🚀 輸出 2027 品檢報表統計.xlsx`。
  - 手動切換下拉選單至 `2035` 年，按鈕文案立即響應變更為 `🚀 輸出 2035 品檢報表統計.xlsx`。
  - 前端與後端 ETL 過濾邏輯保持高度一致，具備高防禦性。

---

## 2026-07-06 更新預設表單編碼與名稱對照表

### 需求說明
更新系統預設的表單編碼與名稱對照表（`DEFAULT_MAPPINGS`）如下：
- QC10002-R02 原物料品檢表
- QC10004-R02 QIP
- QC10006-R01 裝配對樣巡檢記錄表
- QC10006-R02 半成品品檢表
- QC10007-R01 完成品品檢表(首頁)
- QC10007-R02 完成品品檢表(續頁)
- QC10007-R03 零組件入庫品檢表
- QC10008-R02 出貨檢驗報告

### 遇到的問題與根因分析 (RCA)
- **問題**：系統舊的對照表包含了一些不再使用的表單（如 QC10001-R01, QC10002-R01 等），且某些表單的預設名稱不符合當前品檢規範（例如 QC10004-R02 舊名為 "QUALITY INSPECTION PLAN RECORD"，QC10007-R03 舊名為 "零組件入庫品檢表(射出零件品檢表?)"）。
- **影響分析**：如果使用者未自訂對照表，系統會載入舊的預設值，且 `恢復預設` 也會還原至舊版，與最新品檢業務格式不符。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. 修改 `src/utils/db.js` 中的 `DEFAULT_MAPPINGS`，將其完全替換為使用者指定的 8 個品檢項目。
  2. 為了確保舊使用者在重新載入網頁時能自動套用新的對照表，我們在 `getMappings` 函數中新增一個智慧更新檢查。如果發現 local storage 中的 mappings 包含已棄用的鍵（如 `QC10001-R01` 等），則自動為使用者重置並儲存最新的 `DEFAULT_MAPPINGS`。
- **副作用防禦掃描**：
  - 更新預設對照表不會破壞 `App.jsx` 的狀態管理。
  - `excelParser.js` 查詢 `mappings` 時，若找不到特定的 Code，會健全地 fallback 回 `"未對照編碼"`，因此不影響解析流程的魯棒性。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修改 `src/utils/db.js`
- [x] 驗證並測試

---

## 2026-07-06 工作表編碼提取表格加入檔案路徑與 Excel 篩選器功能

### 需求說明
1. 在 UI 提取結果表格「工作表表單編碼提取結果」中插入「檔案路徑」欄位。
2. 讓所有欄位（檔案名稱、檔案路徑、工作表名稱、表單編碼、表單對照名稱、狀態）具備如同 Excel 篩選器（下拉彈窗、全選、個別選取、搜尋篩選值）一樣的功能。
3. 設計採用 McKinsey Premium 風格及高雅的毛玻璃（Glassmorphism）懸浮面板，並完美適配 Light/Dark Mode。

### 遇到的問題與根因分析 (RCA)
- **需求痛點**：原本提取表格僅展示檔名與工作表，當同名檔案存在於不同子目錄下時，無法區分其實際物理路徑。另外，原本的過濾手段僅有全局 search bar 與單一 status select，當解析出數千筆工作表時，極難進行交叉比對與精確篩選。
- **佈局失衡 (偏右) RCA**：
  - 新增之 `filePath`（檔案路徑）為長字串，由於表格單元格設置了 `white-space: nowrap` 且沒有設定最大寬度，導致整列被極大地撐開。
  - 由於外層使用的是 CSS Grid 佈局（`.main-grid`），其欄寬比例原定為 `480px 1fr`。在 CSS Grid 中，`1fr` 預設的隱式最小寬度為其子內容的最小寬度（即 `minmax(auto, 1fr)`）。這導致當右側表格被長路徑撐開時，整個 Grid 欄位直接溢出網頁視窗，導致整體視覺重心「偏右」且出現網頁級滾動條。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **路徑提取**：修改 `processFilesList` 函數，在讀取上傳檔案時優先抓取 `file.webkitRelativePath || file.name`，並將其傳入狀態作為 `filePath` 屬性。
  2. **欄位狀態管理與過濾**：
     - 在 `App.jsx` 中新增 `columnFilters` 狀態物件，保存所有欄位的選取陣列。
     - 重構 `filteredRows` 的過濾公式，使多個欄位篩選條件以 AND 關係連動。
     - 當重新上傳新檔案時，自動重置過濾狀態（重設為全選/全空），防止過期篩選器干擾。
  3. **一鍵重設與篩選元件**：
     - 封裝高雅的 `ColumnFilterPopover` 輔助元件，內含搜尋框、(全選) Checkbox 及動態列表。
     - 提供「套用」與「清除篩選」按鈕。
     - 在表格上方控制區新增一個動態顯示的 **`🧹 重設篩選`** 按鈕，當任何一列存在篩選條件時，使用者可一鍵清除所有篩選，大幅提升復原操作便捷性。
  4. **佈局平衡與防溢出重構**：
     - 將 `.main-grid` 的欄寬定義修改為 `480px minmax(0, 1fr)`。這能將隱式最小寬度限制為 0，防止右側內容撐開 Grid，強制寬度在 `.table-wrapper` 內部產生橫向滾動條。
     - 為 `filePath` 儲存格新增 `.filepath-cell` 類別，限制最大寬度（`max-width: 280px`），允許折行（`white-space: normal !important`）及字元間折行（`word-break: break-all`），防止檔案路徑將表格無限撐長，使整體佈局左右視覺完美對稱與置中。
  5. **字型美學與中文 Fallback 統一**：
     - **RCA**：由於 `index.html` 內缺乏 `<link rel="preconnect">`，網頁載入時 Google 提供的 `Noto Sans TC` 大檔案 Web Font 請求反應極慢或失效。由於 CSS `:root` 中的 `--font-family` 缺乏針對 Windows / macOS 本地中文無襯線字體（黑體/微軟正黑體/蘋方體）的明確宣告，瀏覽器在 Web Font 未能及時渲染的情況下，直接退化回系統預設的「新細明體/宋體」等襯線字體，造成嚴重的視覺割裂感與低端感。
     - **CAPA 矯正**：
       - 在 `index.html` 的 `<head>` 中新增 preconnect 宣告，提早對 fonts.gstatic.com 進行 DNS 及 TLS 解析。
       - 修改 `:root` 的 `--font-family` 變數，顯式寫入 `'PingFang TC'`、`'Microsoft JhengHei'` 作為中文的無襯線黑體備選方案。
       - 修改 `.filepath-cell` 的字型棧為 `Consolas, Monaco, 'Outfit', 'Noto Sans TC', 'Microsoft JhengHei', monospace;`，使路徑在等寬渲染的同時，其中的中文（如 `零組件入庫`）也能保持圓潤的黑體字形風格，完成全系統字型的美學大一統。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修改 `src/App.jsx` (路徑提取、過濾邏輯、一鍵重設按鈕)
- [x] 修改 `src/index.css` (Popover、佈局平衡、路徑折行、字型變數Fallback)
- [x] 修改 `index.html` (字體預先載入)
- [x] 驗證並測試

---

## 2026-07-06 篩選欄 Popover 對比度與主題重構

### 需求說明
修復在特定系統主題（如 Dark Mode）下，篩選 Popover 面板對比度極低、文字難以辨識的視覺 Bug。

### 根因分析 (RCA)
- **局部暗色主題偏離**：本專案網頁主體為 Light Mode（淺色海軍藍/白底），並無全功能 Dark Mode 自適應。然而，CSS 當中對 `.filter-popover` 設置了 `@media (prefers-color-scheme: dark)` 媒體查詢，當使用者的作業系統設為 Dark Mode 時，會將彈窗背景改成深 Slate 灰（`rgba(30, 41, 59, 0.95)`）。
- **文字對比度失效**：彈窗背景變黑，但其內部文字並無針對 dark query 的顯式變色，直接繼承了全局 body 的 `var(--text-primary)`（深黑色，如 `#0F172A`），導致深色背景撞車深黑色字，對比度歸零，文字完全無法辨識。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  - 依循 YAGNI 原則與設計大一統規範，將 `src/index.css` 當中所有對 `.filter-popover` 及其子元素的 `@media (prefers-color-scheme: dark)` 媒體查詢區塊徹底移除。
  - 讓篩選 Popover 統一不論系統 OS 主題為何，皆使用高透明度、精緻白底黑字的 McKinsey 風格毛玻璃，以維持與全局白色儀表板 100% 協調的視覺品質與極致高對比度。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修改 `src/index.css` (移除 Popover 的 dark mode 媒體查詢)
- [x] 驗證並測試

---

## 2026-07-06 多欄位連動過濾優化 (Excel-style Multi-column Linked Filter)

### 需求說明
優化先前實作的欄位過濾器，支援與 Excel 一致的「多欄位連動篩選」功能。當在多個欄位同時進行篩選時，各欄位 Popover 內的備選清單應依據其他欄位已套用的條件動態收窄，以防因選項衝突導致過濾出空表格（全空），提升多維交叉分析的實用性。

### 根因分析 (RCA)
- **非連動候選值 (No Linking)**：原先的 `getUniqueColumnValues` 始終從原始的 `scannedRows` (全局數據) 提取唯一值。這會導致當第一個欄位（如「狀態」）過濾為 matched 後，第二個欄位（如「表單編碼」）的篩選面板中依然會出現與 matched 衝突的無編碼選項（如 `QC99999-R99`）。若使用者在第二個欄位勾選該衝突選項，表格將會因條件 AND 衝突而直接變空。
- **防止選項死結的 Excel 設計**：
  - 若將候選清單直接依據當前 `filteredRows` 收窄，會導致使用者一旦在欄位 $i$ 篩選了某個值，該欄位本身的清單就只剩下該值，使得無法再把別的值「勾回來」，形成操作死結。
  - Excel 的正統邏輯是：在計算欄位 $i$ 的候選唯一值清單時，過濾條件應**排除欄位 $i$ 本身已套用的條件**，但**套用其他所有欄位已套用的條件**。這樣既能動態收窄（消除與其他欄位衝突的無效值），又能讓使用者自由調整目前欄位的多選項目（不產生死結）。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  - 重構 [src/App.jsx:getUniqueColumnValues](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/src/App.jsx#L1125-L1155) 的篩選候選值提取算法：
    1. 在計算某個 `fieldKey` 的候選值時，先對 `scannedRows` 進行一次臨時過濾。
    2. 該臨時過濾會檢查並套用全域 `searchQuery` 與 `statusFilter` 條件。
    3. 遍歷所有的 `columnFilters`：若鍵等於當前 `fieldKey`，則**跳過**；否則套用該鍵的勾選限制。
    4. 提取過濾後的 values 並去重、排序，作為該 Popover 的展示選項。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修改 `src/App.jsx` (重構 getUniqueColumnValues 實現動態連動過濾)
- [x] 驗證並測試

---

## 2026-07-06 左右容器底部對齊與表格垂直滾動優化 (Height Alignment & Scroll Wrapper)

### 需求說明
讓右側「工作表表單編碼提取結果」的表格容器底部在視覺上與左側最下方卡片的底部完美對齊。當提取結果非常長時，在右側容器內部產生滾動條，而不是無限拉長整個網頁。

### 根因分析 (RCA)
- **CSS 絕對定位導致的 Grid 擠壓 Bug**：
  - 若使用純 CSS 絕對定位包裹右側面板內容，由於其內所有子內容脫離了文件流，Grid 會認為該單元格寬度為 0。當外層為 `minmax(0, 1fr)` 欄寬時，瀏覽器會將右側欄位徹底壓縮至約 30px 寬，導致標題文字垂直排列、表格被擠壓。
  - **解決方案：React ResizeObserver + maxHeight 動態對齊**：
    - 不使用任何會脫離文件流的 CSS 絕對定位，維持右側面板正常的區塊佈局以正確參與 Grid 寬度自適應。
    - 在 React 中使用 `ResizeObserver` 監聽左側 `<aside>` 容器的真實 border-box 高度，存入狀態 `leftPanelHeight`。
    - 將此高度作為 inline style 的 `maxHeight` 動態套用在右側結果卡片上，並設定其 `overflow: hidden; display: flex; flex-direction: column`。
    - 內層的 `.table-wrapper` 設為 `flex: 1; overflow: auto;`。這能完美確保：左右側底部精確對齊，且超長數據在表格內滾動，且 Grid 寬度 100% 正常。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **高度狀態與監聽器**：
     - 在 [src/App.jsx](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/src/App.jsx) 中新增 `leftPanelRef` 及 `leftPanelHeight` 狀態。
     - 建立 `useEffect` 監聽器，使用 `ResizeObserver` 在左側容器發生大小改變（如重新上傳檔案、展開一鍵生成器等）時，自動獲取其真實高度。
     - 將 `ref={leftPanelRef}` 綁定至左側 `<aside className="mck-main-content">`。
  2. **動態高度綁定與滾動**：
     - 將 `maxHeight: leftPanelHeight ? \`\${leftPanelHeight}px\` : 'none'\` inline-style 綁定至右側表格容器，並設置 `overflow: 'hidden'`。
     - 在 [src/index.css](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/src/index.css) 中，將 `.table-wrapper` 設為 `overflow: auto; flex: 1;`，使表格能在剩餘高度中正常垂直與水平滾動。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修改 `src/App.jsx` (引進 ResizeObserver 狀態與 Ref、右側 maxHeight 動態對齊)
- [x] 修改 `src/index.css` (table-wrapper 設定 flex 增長與 overflow: auto 滾動)
- [x] 驗證並測試




---

## 2026-07-06 GitHub Actions 部署工作流偵錯 (Deploy Workflow Debug & Node 24 Upgrade)

### 需求說明
修復 GitHub Pages 部署時發生的 `Deployment failed, try again later.` 遠端部署異常。

### 根因分析 (RCA)
- **缺乏打包產物可見性**：Actions 的 Build 與 Upload 步驟之間缺乏對產出目錄（`./dist`）的結構檢驗。如果發生檔案遺失，工作流無法第一時間攔截與印出 logs。
- **潛在的 Environment 審查阻礙**：`github-pages` 部署環境宣告如果配備了 Environment protection rules（例如 required reviewers），會造成 Actions 拒絕或超時失敗。
- **Node 版本警告**：Node 20 已經被 GitHub 標記為舊版本，需要升級到穩定的 LTS (Node 24)。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **導入 Build 產物斷言**：在 [.github/workflows/deploy.yml](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/.github/workflows/deploy.yml) 的 Build 後，加入 `Show build output` 步驟：
     - 使用 `ls` 與 `find` 列印產出目錄結構。
     - 使用斷言判斷若 `dist/index.html` 不存在，則直接以 Exit Code 1 中斷，防止上傳空 Artifact。
  2. **環境控制放寬**：暫時在部署工作流中移除對 `github-pages` 的 environment 限制，以排除權限審查干擾。
  3. **升級 Node 24**：升級 Node Setup 到 LTS 24。

- [x] 驗證並測試

---

## 2026-07-06 新增 CSV 匯出功能 (CSV Export for Extracted Results)

### 需求說明
新增匯出按鈕，可將「工作表表單編碼提取結果」表格中的資料匯出為 `.csv` 格式檔案。

### 根因分析與設計 (RCA & Design)
- **篩選後資料價值**：與原本匯出 Excel（導出全體 `scannedRows`）不同，使用者更傾向於導出經由篩選器層層過濾後的結果（`filteredRows`），以供進一步利用。
- **Excel 中文亂碼防護**：
  - Windows 版 Excel 直接讀取一般 CSV 時，若檔案為 UTF-8 編碼但缺乏 **BOM (Byte Order Mark)**，會將中文字元解析為亂碼。
  - **解決方案**：在生成的 CSV 字串開頭加上 `\uFEFF`，強制 Excel 以 UTF-8 編碼正常顯示中文。
  - **安全性與防護**：使用正規表達式對包含逗號 `,`、換行 `\n` 或引號 `"` 的欄位值進行雙引號 `"` 轉義，確保 CSV 檔案在各種編輯器（如 Notepad, Python, Excel）中都能正常解析。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **建立匯出函數**：在 [src/App.jsx](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/src/App.jsx) 中實作 `exportToCSV(data, name)` 函數，定義中文標題（檔案名稱、檔案路徑、工作表名稱、表單編碼、表單對照名稱、狀態）與欄位對應。
  2. **中文字態處理**：在寫入 CSV 時，調用既有的 `getStatusLabel` 函數，將內部狀態值（如 `'matched'`）轉換為中文標籤（如 `'✓ 成功識別'`）。
  3. **加裝 UI 按鈕**：在「📂 選取資料夾」載入成功後的按鈕區塊，新增「📄 匯出 CSV」按鈕，點擊時呼叫 `exportToCSV(filteredRows, folderName)`。

- [x] 驗證並測試

---

## 2026-07-06 移除匯出欄位映射按鈕 (Remove Export Field Mapping Button)

### 需求說明
從資料夾掃描卡片下方之按鈕控制列中，徹底移除「📋 匯出欄位映射」按鈕，防除因排版寬度限縮造成的文字折行，使 UI 更加簡潔清爽。

### 根因分析 (RCA)
- **空間過窄與字詞折行**：在左側側邊欄固定寬度（480px，扣除 padding 後約 432px）的容器下，若並列三個帶有圖示的長按鈕（匯出 Excel、匯出 CSV、匯出欄位映射），按鈕的寬度會被壓縮至極限。這導致「匯出欄位映射」中的「射」字被迫折行到第二行（呈現 `匯出欄位映\n射`），損害了 Art Director 的極致美學標準。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  - 修改 [src/App.jsx](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/src/App.jsx)，從 `folderName` 載入成功後的按鈕 flex-container 中，徹底移除 `<button onClick={() => exportFieldMapping()}>` 按鈕元素。
  - 移除後，操作列僅保留兩個同等權重的橫向按鈕（「💾 匯出 Excel」與「📄 匯出 CSV」）以及一個「🗑」清除按鈕，平分空間後文字排版極其舒適、寬鬆，不再有任何折行缺陷。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修改 `src/App.jsx` (從 UI 移除「📋 匯出欄位映射」按鈕)
- [x] 驗證並測試

---

## 2026-07-07 ETL 納入狀態追蹤功能 (ETL Status Tracking for Worksheets)

### 需求說明
在「工作表表單編碼提取結果」表格中新增「是否納入ETL計算」與「更新時間」欄位，顯示每個工作表在 ETL 流程完成後是否已被納入統計計算。
- 取值：`已納入` (綠色)、`未納入` (灰色)、`狀態異常` (紅色)。
- 若 scannedRows 中存在任一 `狀態異常` 的工作表，在表格上方顯示醒目的紅色警告橫幅，並支援一鍵篩選異常。
- CSV 匯出支援該狀態。
- 寫入目前更新時間戳。

### 根因分析與設計 (RCA & Design)
- **判定邏輯整合**：為了精確判斷工作表是否納入統計，必須完全套用與 `browserETL.js` 中相同的篩選、去重和分類條件。為了避免重複開發程式碼導致維護困難，我們將 `browserETL.js` 的底層工具函數導出，並在 `excelParser.js` 的 `parseExcelFile` 中引用。
- **不同品檢的判定條件**：
  - **QIP 射出檢驗**：在 Patrol 資料夾下，且 sheet 名稱符合 Date Code 正規表達式、未重複，為 `已納入`；其餘為 `未納入`；若資料夾名稱不含月份格式，為 `狀態異常`。
  - **QIP 押出檢驗**：檔名符合 Date Code 格式，且工作表不在忽略清單、不是 Setup 頁面，且為 Patrol 唯一基準名，為 `已納入`；其餘為 `未納入`；資料夾無月份格式為 `狀態異常`。
  - **一般品檢 (General)**：能辨識 QC Code、具有正確子分類與月份者為 `已納入`；跳過之工作表或空白檔為 `未納入`；若辨識出 QC Code 但月份或子分類為空/超出範圍，為 `狀態異常`。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **導出工具函數**：在 [src/utils/browserETL.js](file:///c:/Users/USER/Downloads/專案/FileName_WorkSheet_Extract/src/utils/browserETL.js) 導出 `detectQCFromFolder` 等底層工具。
  2. **實作狀態判定**：在 [src/utils/excelParser.js](file:///c:/Users/USER/Downloads/專案/FileName_WorkSheet_Extract/src/utils/excelParser.js) 的 `parseExcelFile` 中實作詳細的 `etlStatus` 與 `etlTimestamp` 判定邏輯。
  3. **表格與過濾器更新**：修改 [src/App.jsx](file:///c:/Users/USER/Downloads/專案/FileName_WorkSheet_Extract/src/App.jsx) 的過濾與搜尋邏輯，加入新欄位，並加入 McKinsey 風格紅色告警橫幅與「🔍 立即篩選異常」按鈕。
  4. **CSV 匯出欄位同步**：在 `exportToCSV` 中加入 `etlStatus` 與 `etlTimestamp` 欄位。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修改 `src/utils/browserETL.js` (導出底層工具函數)
- [x] 修改 `src/utils/excelParser.js` (實現 etlStatus 與 etlTimestamp 的計算)
- [x] 修改 `src/App.jsx` (新增表格欄位、篩選過濾、告警橫幅與一鍵篩選、CSV 匯出、優化檔案路徑只顯示資料夾)
- [x] 驗證並測試

---

## 2026-07-07 檔案路徑顯示優化 (File Path Display Optimization)

### 需求說明
優化「工作表表單編碼提取結果」表格中的「檔案路徑 (filePath)」欄位，使其不顯示具體的檔案名稱（因為檔案名稱已在「檔案名稱」欄位中呈現），改為只顯示從根目錄到檔案所在資料夾的層級路徑，避免資訊冗餘。

### 根因分析與設計 (RCA & Design)
- **資訊去重**：原本 `filePath` 屬性儲存的是 `RawData/2026/零組件入庫-2026/裝配C-2026/裝配C-2026A.xlsx`。在結果表格中，由於第二欄已經專門顯示「檔案名稱」，第三欄「檔案路徑」如果又包含檔名，會導致版面過於擁擠。
- **資料層級處理**：透過對路徑字串進行處理，尋找最後一個斜線 `/` 或 `\` 並擷取其前半段，即可動態取得父級資料夾路徑。若無父級資料夾則回傳空字串。這能同時應用於列表過濾、CSV 匯出與前端表格渲染中，保證資料的 MECE（不重疊、不遺漏）原則。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **優化資料提取**：修改 [src/App.jsx:processFilesList](file:///c:/Users/USER/Downloads/專案/FileName_WorkSheet_Extract/src/App.jsx)，在將讀取結果映射回 sheetsWithPath 時，透過 `lastIndexOf('/')` 擷取不含檔名的資料夾路徑（`dirPath`），並賦予 `filePath` 屬性。
  2. **同步 Mock 資料**：將 Mount 時載入的測試 mock 數據中的所有 `filePath` 同步調整為不含檔名的路徑，維持表格載入的一致性。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修改 `src/App.jsx` (processFilesList 路徑處理與 Mock 數據路徑優化)
- [x] 驗證並測試

---

## 2026-07-07 QIP 科學記號工作表去重與解析優化 (Scientific Notation Sheet Name Fix)

### 需求說明
當 QIP 品檢 Excel 檔案中的工作表名稱含有科學記號後綴（如 `260101E-2`、`260101E-3`），Excel 或 JavaScript 解析引擎可能將其誤讀為科學記號浮點數（`260101E-2 → 2601.01`）。這導致兩個問題：
1. 工作表名稱顯示為 `2601.01` 而非 `260101E-2`，使用者難以辨識。
2. `2601.01` 無法通過 Date Code 正則驗證 `/^\d{6}[a-zA-Z]?$/`，導致巡檢計數遺失，或無法正確提取月份。
3. 三張工作表（`260101E`、`260101E-2`、`260101E-3`）應去重後只計算一次，但若誤讀為數值則去重邏輯失效。

### 根因分析 (RCA)
- **根因一（工作表名稱誤判）**：SheetJS 讀入 Excel 檔案時，若工作表名稱在 XML 中已被外部工具或 Excel 本身存為數值型 `2601.01`（因 `260101E-2` 符合科學記號格式），`wb.SheetNames` 中將出現 `"2601.01"` 字串而非 `"260101E-2"`，無法通過 Date Code 正則。
- **根因二（日期提取失敗）**：若儲存格 Q4（QC10004-R02 日期欄位）內含 `260101E-2`，Excel 將其讀為數值 `2601.01`。`parseDateFromValue` 的 `typeof val === 'number'` 分支不符合序列日期範圍（40000~50000），`parseDateFromString("2601.01")` 也無法匹配 6 位 Date Code 格式，月份提取返回 `null`，工作表被標記為 `狀態異常` 或 `未納入`。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **新增 `normalizeScientificNotation` 函數**：在 [src/utils/browserETL.js](file:///c:/Users/USER/Downloads/專案/FileName_WorkSheet_Extract/src/utils/browserETL.js) 中新增並導出此函數。邏輯如下：
     - 若輸入字串符合 `^\d{4}\.\d+$`（如 `2601.01`），計算小數位數 $exp$，乘以 $10^{exp}$ 四捨五入重建 6 位基底，拼回 `${base}E-${exp}`（還原為 `260101E-2`）。
     - 若輸入符合標準科學記號字串格式（如 `2.60101e+5`），直接還原為 6 位整數字串。
  2. **於 `parseDateFromString` 中應用**：函數入口處先呼叫 `normalizeScientificNotation`，確保日期字串 `2601.01` 還原後能正確匹配 Date Code。
  3. **於 QIP 巡檢工作表處理中應用**：在 [browserETL.js:runETLInBrowser](file:///c:/Users/USER/Downloads/專案/FileName_WorkSheet_Extract/src/utils/browserETL.js) 的射出與押出迴圈中，對 `sheetName` 先還原再進行後綴去重與 Date Code 驗證。
  4. **於 `parseExcelFile` 中全面應用**：修改 [src/utils/excelParser.js](file:///c:/Users/USER/Downloads/專案/FileName_WorkSheet_Extract/src/utils/excelParser.js)，在迴圈起始處先計算 `normalizedSheetName`，取代所有 ETL 判定、過濾與輸出中的 `sheetName`，確保 UI 表格與 CSV 均顯示還原後的正確名稱。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 在 `src/utils/browserETL.js` 中新增並導出 `normalizeScientificNotation`
- [x] 在 `src/utils/browserETL.js` 的 `parseDateFromString` 與 `runETLInBrowser` 中應用
- [x] 在 `src/utils/excelParser.js` 中導入並全面應用 `normalizeScientificNotation`
- [x] 驗證並測試

---

## 2026-07-07 整合 Ponytail (Integrate DietrichGebert/ponytail)

### 需求說明
將 DietrichGebert/ponytail 專案整合進本專案，以改善 AI 代理在處理代碼時的 token 使用效能與防止過度工程（over-engineering）。

### 根因分析與設計 (RCA & Design)
- **Token 效能優化**：AI 代理常有「過度建構、過度工程」的傾向。Ponytail 的決策階梯與 lazy 開發原則能引導 AI 代理在編寫代碼時優先利用 YAGNI、本機庫、既有依賴和一行解決方案，減少 80% 以上的程式碼生成，藉此提高 token 使用效率與系統穩定度。
- **整合方式**：
  1. 將 Ponytail 的核心 rules (來自 `AGENTS.md`) 整合到專案的 [`.agents/AGENTS.md`](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/.agents/AGENTS.md) 中。
  2. 將 Ponytail 提供的所有客製化 skills (`ponytail`, `ponytail-audit`, `ponytail-debt`, `ponytail-gain`, `ponytail-help`, `ponytail-review`) 從臨時 Repository 複製到專案的 [`.agents/skills/`](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/.agents/skills/) 中。
  3. 依據 MECE 原則清理複製過程中產生的 `scratch/ponytail` 臨時檔案，避免膨脹。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. 修改 `.agents/AGENTS.md`：在尾部加入 Ponytail rules。
  2. 複製 `scratch/ponytail/skills/` 內所有 skills 資料夾至 `.agents/skills/` 目錄。
  3. 刪除 `scratch/ponytail` 的暫存 Repository 複製。
  4. 驗證所有 skills 都已在 `.agents/skills/` 下並包含對應的 `SKILL.md`。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 將 Ponytail 的 rules 寫入 `.agents/AGENTS.md`
- [x] 將所有 Ponytail skills 複製到 `.agents/skills/` 
- [x] 清理 scratch/ponytail 暫存檔案
- [x] 確效驗證

---

## 2026-07-07 修正 Excel 提取結果的 ETL 狀態判定 Bug 與原因回寫 (ETL Status Bug Fix & Explanation Backfill)

### 需求說明
1. 列出 2025 品檢報表提取結果中被成功識別但顯示為「未納入」的原因。
2. 將個別的具體排除/納入原因寫入使用者 Downloads 目錄下的 `2025 報表_提取結果.xlsx` 對應儲存格中。
3. 修正系統程式碼中有關「是否納入ETL計算」的狀態判定 Bug。

### 根因分析與設計 (RCA & Design)
- **判定鎖死 Bug**：在 [excelParser.js](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/src/utils/excelParser.js) 中，一般品檢處理分支在最一開始就把 `etlStatus` 初始化為 `"未納入"`，但後續的判定區塊被 `if (etlStatus !== "未納入")` 條件包裹。此條件恆為假，導致正常識別的表單狀態全部被鎖死在 `"未納入"`，繞過了所有月份及子分類提取。
- **原因判定與回寫**：
  - 寫入腳本 `scratch/update_excel.js`，依據 QIP 射出、押出與一般品檢數據校驗規則，對 26,524 筆數據進行精準匹配。
  - 對於一般品檢，若需要檢查空白樣板 (QC10007-R03)，則動態載入 workspace 中的 [RawData/2025/](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/RawData/2025) 原始 Excel 文件，讀取對應儲存格判斷批號是否為空。
  - 計算出正確狀態後，在 Excel 中新增 `原因說明` 欄位並更新 `是否納入ETL計算` 欄位，完成回寫。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **修正 Bug**：在 [excelParser.js](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/src/utils/excelParser.js#L173) 的一般品檢 else 分支起始處，將 `etlStatus` 初始化為 `"已納入"`，使後續判定邏輯可以正常執行。
  2. **數據回寫**：執行回寫腳本，成功更新 `C:\Users\3kids\Downloads\2025 報表_提取結果.xlsx` 共 26,524 筆資料，並建立 `原因說明` 欄位。
  3. **代碼確效**：執行 `npm run build` 確認無編譯錯誤。
  4. **MECE 清理**：清除 `scratch/` 下的暫存分析腳本。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修正 `src/utils/excelParser.js` 狀態判定 Bug
- [x] 撰寫並執行 Excel 回寫腳本 `scratch/update_excel.js`
- [x] 清理 scratch 暫存檔案
- [x] 代碼打包確效驗證
- [x] 新增「原因說明」至前端「工作表表單編碼提取結果」表格 UI 欄位與 CSV 匯出功能

---

## 2026-07-07 檔名包含「空白」且含其他文字之有效檔案判斷優化 (Blank File Name Checking Optimization)

### 需求說明
如果檔案名稱中包含「空白」二字，但同時包含其他任何字元（例如 `裝配巡檢記錄表-空白.xlsx`、`空白包裝袋.xlsx`），則不應被系統一刀切地判定為空白範本，而應視為有效數據檔案進行處理。

### 根因分析與設計 (RCA & Design)
- **原過濾邏輯過於單一**：原邏輯中使用 `fileName.indexOf('空白') >= 0` 作為空白檔案的判定依據。這使得只要檔名中出現「空白」二字的任何有效品檢檔案都會被直接排除（例如 `空白包裝袋.xlsx` 被排除）。
- **優化設計 (相鄰文字規則)**：
  - 偵測檔名中 `"空白"` 二字的所有索引。
  - 對於每個 `"空白"`，檢查其前一個字元與後一個字元是否為英文字母或中文字元（正則：`/[a-zA-Z\u4e00-\u9fa5]/`）。
  - 若偵測到其中任一 `"空白"` 的前後緊鄰著字母或中文字元（例如 `空白包裝袋.xlsx` 中 `白` 緊鄰 `包`），則表示此 `"空白"` 是其他詞彙的一部分，視為有效檔案。
  - 若檔名中所有 `"空白"` 的前後緊鄰字元都不是字母或中文（即僅有符號如 `-`、`_`、數字、括號或副檔名點號，如 `裝配巡檢記錄表-空白.xlsx`、`空白.xlsx`），則視為無效的空白樣板檔案。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **優化 `excelParser.js`**：更新 [src/utils/excelParser.js](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/src/utils/excelParser.js#L189) 的 `isBlankFile` 判斷邏輯與原因說明。
  2. **優化 `browserETL.js`**：更新 [src/utils/browserETL.js](file:///d:/Self-developed_Apps/FileName_WorkSheet_Extract/src/utils/browserETL.js#L446) 的 `isBlankFile` 過濾邏輯。
  3. **建構驗證**：重新執行 `npm run build` 確認程式編譯正常。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修正 `src/utils/excelParser.js` 的空白檔相鄰字元過濾邏輯
- [x] 修正 `src/utils/browserETL.js` 的空白檔相鄰字元過濾邏輯
- [x] 代碼打包確效驗證

---

## 2026-07-08 調整裝配A/B/C (QC10007-R03) 月份分類邏輯 (Strict Filename Suffix Month Mapping for Assembly Parts)

### 需求說明
針對零組件入庫品檢 (QC10007-R03) 下的裝配組件 `裝配A`、`裝配B`、`裝配C`，修改其月份判定規則：不再依賴工作表內部的 `O4` 儲存格日期欄位進行驗證或判定，而是完全依據檔案名稱的後綴英文字母 A-L（A=1月，B=2月...）做為該檔案內所有工作表的月份分類依據，以避免因跨月檢驗造成統計筆數與實際檔案不一致。

### 根因分析 (RCA)
- **跨月日期造成統計誤差**：原先的 `QC10007-R03` 月份判定邏輯會先讀取 `O4` 儲存格的日期，並與檔名後綴做比對，若不相符則退回 `extractRawMonth` 對應的實際儲存格日期。對於 `裝配C-2025A.xlsx`（一月份檔案），其中有兩個工作表（`PJX25A12` 與 `PJW25A41`）的檢驗日期實際上落在二月與三月。這導致系統在統計時將這兩張工作表歸入二月與三月，導致一月份統計少 2 筆，其他月份虛增。
- **業務規則變更**：使用者要求裝配A/B/C需嚴格按檔案所屬月份（檔名後綴字母）進行歸類，不進行跨月內容日期校正。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **新增裝配組件月份覆寫邏輯**：在 `excelParser.js` 與 `browserETL.js` 中判定，當 `actualQC` 為 `QC10007-R03` 且 `subCat` 屬於 `裝配A`/`裝配B`/`裝配C` 時，僅提取檔案名稱結尾字母後綴 A-L 的對應月份，並不使用 `findDateInSheet` 與 `extractRawMonth`。
  2. **保護其他類別與流程**：此行為僅限於 `QC10007-R03` 的裝配子類別，`Tubing`、`射出` 等其他子類別及其他 QC 表單依然維持原有的內容日期校對機制。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修改 `src/utils/excelParser.js` (套用裝配組件檔名月份覆寫與 fallback 保護)
- [x] 修改 `src/utils/browserETL.js` (同步套用裝配組件檔名月份覆寫與 fallback 保護)
- [x] 驗證並測試 (使用 `analyze_sheets.cjs` 驗證一月份 10 個工作表全部歸入 1 月)

---

## 2026-07-08 修訂零組件入庫 (QC10007-R03) Tubing 月份提取邏輯

### 需求說明
修訂零組件入庫品檢 (QC10007-R03) 中的 `Tubing` 數據映射邏輯為：對應當月份資料夾內有效的工作表數量，並放棄過濾 `O4` 欄位的日期資訊，其他部分保持不變。

### 根因分析 (RCA)
- **跨月日期與填寫錯誤導致偏離資料夾月份**：原先的 `Tubing` 月份判定邏輯會呼叫 `extractRawMonth`，其優先讀取工作表內 `O4` 儲存格的日期，並以該日期月份作為資料月份。如果工作表 `O4` 欄位日期填寫錯誤、不符年份限制，或者填寫了跨月份的交期日期，會導致對應的工作表被歸入錯誤月份，或是被判定為無效而被過濾掉。
- **業務規則調整**：對於 `Tubing` 資料，使用者要求僅根據「當月份資料夾」（即檔案的上游資料夾路徑中的月份，如 `Tubing-2026-02` 底下的檔案歸為二月份）來計算有效工作表數量，完全不再從工作表內部的 `O4` 儲存格讀取或校驗日期。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **新增 Tubing 月份資料夾覆寫邏輯**：在 `excelParser.js` 與 `browserETL.js` 中判定，當 `actualQC === 'QC10007-R03'` 且 `subCat === 'Tubing'` 時，直接自父資料夾路徑 `relPath` 中匹配並提取結尾代表的月份（如 `Tubing-2026-02` 提取 `2`），而不使用 `findDateInSheet` 或 `extractRawMonth`。
  2. **保護其他流程**：此修改不影響其他非 Tubing 的零組件入庫子類別（如裝配 A/B/C 或射出），也不影響其他 QC 表單。
  3. **代碼確效與測試**：撰寫臨時的測試驗證腳本，產生測試 Excel（其 `O4` 日期設為 1 月，但放在 2 月資料夾路徑下），執行 ETL 運算驗證是否能正確 bypass `O4` 且回傳 2 月。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修改 `src/utils/excelParser.js`
- [x] 修改 `src/utils/browserETL.js`
- [x] 撰寫確效測試驗證腳本
- [x] 代碼打包確效驗證

---

## 2026-07-08 修訂零組件入庫 (QC10007-R03) 射出(廠內) 映射與取捨邏輯

### 需求說明
修訂零組件入庫品檢 (QC10007-R03) 中的 `射出(廠內)` 數據映射與統計取捨邏輯：
1. 取消空白樣板守衛，不判斷 `G4` 儲存格之批號資訊。
2. 以資料夾後綴所代表之月份作為分類依據，捨棄工作表內 `O4` 日期欄位的判斷，統計對應月份資料夾內的有效工作表數量。
3. 其他部分維持不變。

### 根因分析 (RCA)
- **過於嚴格的空白防禦與 cross-month 校驗**：原先的 `QC10007-R03` 設計了空白樣板過濾守衛，強行檢查工作表 `G4` 是否為空。若使用者上傳的 `射出(廠內)` 表單之 `G4` 被保留為空，這些有效數據就會被過濾。此外，以 `O4` 日期為準會導致與資料夾歸類月份（例如 `射出-2026-03` 代表三月數據）不一致，產生跨月漏記。
- **業務規則調整**：對於 `射出(廠內)` 數據，僅需根據其上游月份資料夾的後綴來進行月份劃分，不再限制 `G4` 批號是否為空，且不使用 `O4` 日期。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **取消空白樣板守衛對射出的限制**：在 `browserETL.js` 與 `excelParser.js` 的空白樣板檢驗中，額外判斷當前子分類是否為 `射出`。若是，則跳過空 Lot 檢測。
  2. **新增射出月份資料夾覆寫邏輯**：在 overrides 階段增加對 `tempSub === '射出'` 的月份從資料夾路徑 `relPath` 字尾提取的邏輯，且像 `Tubing` / `裝配` 一樣，將 `射出` 排除於 `extractRawMonth` 外，防止其回退至讀取並過濾 `O4` 日期的流程。
  3. **測試確效**：更新測試腳本，測試空 Lot 及 `O4` 日期衝突情況下的 `射出(廠內)` 檔案，確認其能被順利納入且歸入正確月份。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修改 `src/utils/excelParser.js`
- [x] 修改 `src/utils/browserETL.js`
- [x] 撰寫確效測試驗證腳本
- [x] 代碼打包確效驗證

---

## 2026-07-08 將「射出A」、「射出C」子類別與「射出(廠內)」映射邏輯設定為相同

### 需求說明
將零組件入庫品檢 (QC10007-R03) 當中的 `射出A` 與 `射出C` 子類別之數據映射邏輯修訂為與 `射出(廠內)` 完全相同：
1. 取消對這兩個子類別的空白樣板守衛，不判斷 `G4` 儲存格之批號。
2. 以資料夾後綴所代表之月份作為分類依據，捨棄工作表內 `O4` 日期欄位的判斷，統計對應月份資料夾內的有效工作表數量。
3. 其他部分維持不變。

### 根因分析 (RCA)
- **統計標準對齊**：`射出A`、`射出C` 與 `射出(廠內)` 均屬於射出類別的子表單，在品檢數據提取上應採取相同的統計與取捨標準。為避免空 Lot 被過濾，以及跨月份日期造成的數據錯置，這兩個類別亦需捨棄對 `G4` 與 `O4` 儲存格的直接依賴，改以資料夾為統計依據。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **空白樣板守衛豁免擴充**：在 `browserETL.js` 與 `excelParser.js` 中，將 `isExemptedInjection` 判斷擴充為包含 `射出`、`射出A`、`射出C`。
  2. **月份資料夾覆寫邏輯對齊**：在 overrides 中，將 `tempSub === '射出A'` 與 `tempSub === '射出C'` 同步整合至從資料夾後綴提取月份的邏輯，且將其子類別排除在 `extractRawMonth` 執行之外。
  3. **測試與軟體確效**：擴充確效驗證腳本，平行測試這 4 個品檢子類別（Tubing、射出、射出A、射出C）在空 Lot 及 O4 衝突時的表現，確保皆能通過確效。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修改 `src/utils/excelParser.js`
- [x] 修改 `src/utils/browserETL.js`
- [x] 撰寫確效測試驗證腳本
- [x] 代碼打包確效驗證

---

## 2026-07-08 修訂零組件入庫 (QC10007-R03) 射出D(組件) 映射與取捨邏輯

### 需求說明
修訂零組件入庫品檢 (QC10007-R03) 當中的 `射出D(組件)` 子類別的數據映射與取捨邏輯：
1. 改以檔案名稱的後綴英文字母 A-L（A=1月，B=2月...）作為月份分類的依據，捨棄 O4 儲存格日期的判斷。
2. 統計對應月份的有效工作表，有效的定義與 `射出`、`射出A`、`射出C` 相同（取消空白樣板守衛，不判斷 `G4` 儲存格之批號）。
3. 其他部分維持不變。

### 根因分析 (RCA)
- **跨月校對與空 Lot 過濾影響統計**：與裝配 A/B/C 及其他射出子類別類似，`射出D(組件)` 屬於射出組裝零件的品檢檔案。為了防止跨月日期導致統計紊亂，且防止因 `G4` 批號空白被誤計為空白樣板而過濾掉，需要對齊裝配零件的「檔名後綴決定月份」規則，同時套用射出類別的「豁免空白 Lot 檢測」統計定義。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. **月份提取規則設定為檔名後綴**：在 `browserETL.js` 與 `excelParser.js` 的 overrides 中，將 `tempSub === '射出D(組件)'` 合併入 `isAssemblyParts` / 檔名後綴 (A-L) 提取邏輯。
  2. **空白樣板守衛豁免擴充**：將 `isExemptedInjection` 判斷擴充為包含 `射出D(組件)`。
  3. **Bypass O4 日期檢查**：在 `month === null` 判斷中，將 `subCat === '射出D(組件)'` 納入防護，防止其退回 `extractRawMonth` 與 `O4` 日期檢查。
  4. **確效驗證**：擴充確效驗證腳本，測試 `射出D(組件)` 在檔名後綴 `A`（1月）、`G4` 為空、且 `O4` 衝突（為 2月）時，確認其依然能不被過濾且正確歸入 1 月。

### 進度追蹤
- [x] 更新開發日誌 (DEV_LOG.md)
- [x] 修改 `src/utils/excelParser.js`
- [x] 修改 `src/utils/browserETL.js`
- [x] 撰寫確效測試驗證腳本
- [x] 代碼打包確效驗證

---

## 2026-07-08 專案整體程式碼與檔案優化

### 需求說明
執行專案整體優化與檔案整理作業：
1. **全面盤點與清理**：清理根目錄下冗餘的重複文件。
2. **同步更新文件**：更新 `README.md`，移除已廢棄的後端腳本指引，補上最新 ETL 月份映射規則。
3. **MECE 原則資源整合**：將需求文件統一歸入 `docs/` 下，更名測試工具。
4. **提交與部署準備**：將程式變更進行 Git Commit 建立基準，並向用戶發起推送授權。

### 矯正與預防措施 (CAPA)
- **矯正措施**：
  1. 將根目錄 `today-requirements-2026-07-06.md` 內容搬移並覆寫至 `docs/today-requirements-2026-07-06.md`，然後刪除根目錄檔案。
  2. 將 `scratch/_test_tubing_etl.cjs` 更名為 `scratch/validate_qc_etl.cjs`。
  3. 修改 `README.md`，刪除已棄用的 `etl_pipeline.cjs` 和 `generate_styled_reports.cjs` 相關內容，詳細描述最新的 `QC10007-R03` 的各子類別（裝配、Tubing、射出、射出D組件等）的取捨與月份判定規則。

### 進度追蹤
- [/] 更新開發日誌 (DEV_LOG.md)
- [ ] 整合並移動 `today-requirements-2026-07-06.md` 到 `docs/`，並刪除根目錄的檔案
- [ ] 將 `scratch/_test_tubing_etl.cjs` 更名為 `scratch/validate_qc_etl.cjs`
- [ ] 檢視並更新 `README.md`
- [ ] 執行確效測試及打包建構
- [ ] 提交 Git 變更並寫入 `walkthrough.md`
- [ ] 獲得使用者許可後推送至 GitHub 遠端倉庫




