# 品管報表統計工具 - 現代化網頁應用整合計畫

本專案旨在將原有的 VBA 巨集工具（透過遍歷 Excel 檔案提取工作表及 QC 表單編碼）重構成一個**一體化、極致美觀的現代化單頁網頁應用 (SPA)**。

## 1. 現狀診斷與分析

### 原有 VBA 機制 (`檔名與工作表提取.txt`)
1. **運作流程**：
   - 透過 `Application.FileDialog(msoFileDialogFolderPicker)` 彈窗選擇資料夾。
   - 讀取資料夾內所有 `.xls*` 檔案。
   - 逐一打開 Excel 檔案，遍歷所有工作表。
   - 在每個工作表的 `UsedRange` 中以搜尋格式 `QC?????-R??` 定位。
   - 以正則表達式 `QC\d{5}-R\d{2}` 提取精確編碼，並與硬編碼字典（表單名稱對照表）比對。
   - 將「檔案名稱」、「工作表名稱」、「表單編碼」、「表單名稱」寫入新建的工作表中。

2. **痛點分析**：
   - **相依性高**：必須安裝 Microsoft Excel，且需手動開啟並允許啟用巨集（安全性警告風險）。
   - **維護困難**：QC 表單編碼與名稱對照表（`dictMap`）硬編碼在 VBA 程式碼內。當有新表單時，一般使用者無法自行修改，必須開啟 VBA 編輯器修改代碼。
   - **UI 違和處**：巨集執行時進度不透明（僅有關閉畫面更新），無即時過濾、搜尋與圖表分析功能，結果僅以靜態表格呈現，缺乏現代化美感。
   - **跨平台限制**：無法在 macOS、Linux、手機或平版上運作。

---

## 2. 解決方案：現代化單頁 Web 應用 (SPA)

我們建議使用 **Vite + React + Vanilla CSS** 打造一個完全運行於瀏覽器客戶端（Client-side）的品管統計工具。所有檔案解析均在本地瀏覽器完成，具備高安全性（檔案不需上傳至伺服器）與跨平台運行優勢。

### 核心特色：
1. **多重導入方式**：支援拖放（Drag & Drop）多個 Excel 檔案，或**直接選擇/拖放整個資料夾**進行掃描。
2. **極致 UI/UX 視覺設計 (Color Master Palette - 淺色背景)**：
   - 採用 `#F9FAFB` (Cool Gray 50) 作為 Base Background，並搭配 `#FFFFFF` (Pure White) 卡片區塊，符合色彩大師 Light Mode 規範。
   - 文字使用 `#111827` (Gray 900) 提供高對比度極致閱讀體驗。
   - 品牌輔助色使用 `#3B82F6` (Royal Blue) 做為主要 CTA。
   - 完美對比的字體與間距，極佳的微動態 Hover 與點擊反饋。

3. **動態對照表編輯器**：
   - 使用者可在介面上直接新增、修改或刪除 QC 編碼與名稱，並自動儲存至 LocalStorage（不需碰觸任何代碼，亦可匯入/匯出 JSON 對照檔）。
4. **一體化操作面板**：
   - **Dashboard 統計看板**：視覺化展示已識別表單類型比例、成功率、解析速度與總表單數。
   - **即時互動數據表**：支援模糊搜尋、按檔案或 QC 編碼過濾、一鍵切換「未匹配編碼」或「未識別警告」。
   - **數據導出**：支援將統計結果一鍵導出為乾淨的 Excel (.xlsx) 或 CSV 檔案。

---

## 3. 開發實作方案

### 3.1 專案結構 (MECE 規劃)
我們將在目前工作區下建立 Web App 結構：
```text
f:\Self-developed_Apps\FileName_WorkSheet_Extract/
├── package.json               # npm 專案配置
├── index.html                 # 應用入口
├── vite.config.js             # Vite 配置
├── src/
│   ├── main.jsx               # React Entry
│   ├── App.jsx                # 一體化操作介面主程式
│   ├── index.css              # 全域 Design Tokens 與毛玻璃樣式定義
│   ├── components/            # 模組化組件 (Dashboard, Table, MappingEditor, FileLoader)
│   └── utils/
│       ├── excelParser.js     # SheetJS 封裝，處理前端 Excel 解析與 QC 提取
│       └── db.js              # LocalStorage 封裝，管理 QC 表單對照表
└── DataExtract/               # 測試用 Excel 資料夾 (維持不變)
```

### 3.2 關鍵技術選型
- **開發框架**：Vite + React (超快速開發與熱重載)。
- **Excel 解析**：`xlsx` (SheetJS) 庫。用於讀取檔案、提取工作表、搜尋 UsedRange 內容。
- **正則與提取規則**：與原 VBA 對齊的 `/QC\d{5}-R\d{2}/i` 邏輯。

---

## 4. 使用者審查與開放問題

> [!IMPORTANT]
> **請檢視以下設計考量並給予反饋：**
> 1. **資料夾讀取限制**：
>    在瀏覽器中讀取資料夾需要使用 `<input type="file" webkitdirectory directory>` 屬性，使用者在選擇資料夾時瀏覽器會跳出「確認讀取資料夾內所有檔案」的安全性提示。您是否能接受此瀏覽器原生提示？
> 2. **本機獨立執行**：
>    您希望此工具能直接在本機雙擊 `index.html` 離線打開（無須執行 node 伺服器），還是能接受透過執行 `npm run dev` 啟動本機伺服器瀏覽？（若需要雙擊 HTML 直接離線使用，我們將採用無建置步驟的單檔案 HTML (React via CDN) 或是透過打包工具輸出單一離線 HTML 文件）。
> 3. **統計圖表需求**：
>    是否需要額外的品管圖表（例如各類表單隨時間/月份的統計趨勢），還是基本的總量與識別比例看板即可？

---

## 5. 驗證計畫

### 自動化確效
- 使用測試腳本自動導入 `DataExtract/` 下的 16 個 Excel 檔案，比對提取出的 (File, Sheet, Code, Name) 數量與 VBA 執行的結果是否 100% 一致。

### 手動驗證
- 測試大體積 Excel 檔案載入時的 UI 載入動畫（Spinner）與流暢度。
- 測試動態對照表新增一筆 "QC10009-R01 測試表單"，並驗證掃描結果是否能即時正確對應。
