import React, { useState, useEffect, useRef } from 'react';
import { getMappings, saveMappings, resetMappings } from './utils/db';
import { parseExcelFile, exportToExcel, parseSummaryExcel } from './utils/excelParser';

const MONTH_LABELS = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"];

const MCK_COLORS = [
  '#002D62', // McKinsey Navy
  '#D4AF37', // Consulting Gold
  '#4A5568', // Slate Gray
  '#0EA5E9', // Sky Blue
  '#10B981', // Emerald Green
  '#8B5CF6', // Royal Purple
  '#EF4444', // Coral Red
  '#EC4899', // Magenta Pink
  '#F59E0B', // Warm Amber
];

function App() {
  // Navigation State
  const [activeTab, setActiveTab] = useState("dashboard"); // "dashboard" or "extractor"

  // ==========================================
  // STATE FOR TAB 1: McKinsey Dashboard
  // ==========================================
  const [summaryData, setSummaryData] = useState(null);
  const [activeSheet, setActiveSheet] = useState("");
  const [selectedItems, setSelectedItems] = useState([]); // Array of {name, idx}
  const [summaryFileName, setSummaryFileName] = useState("");
  const [dashboardInsights, setDashboardInsights] = useState({ total: 0, peakMonth: "", peakVal: 0, avg: 0 });

  const chartRef = useRef(null);
  const chartInstance = useRef(null);

  // ==========================================
  // STATE FOR TAB 2: Mappings & Extractor
  // ==========================================
  const [mappings, setMappings] = useState({});
  const [scannedRows, setScannedRows] = useState([]);
  const [isScanning, setIsScanning] = useState(false);
  const [folderName, setFolderName] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [newCode, setNewCode] = useState("");
  const [newName, setNewName] = useState("");

  const folderInputRef = useRef(null);
  const fileInputRef = useRef(null);

  // Load mappings on mount & Auto-load 2025 report if available
  useEffect(() => {
    setMappings(getMappings());

    fetch(import.meta.env.BASE_URL + 'DataExtract/2025品檢報表統計.xlsx')
      .then(res => {
        if (res.ok) return res.blob();
        throw new Error('Not found');
      })
      .then(blob => {
        const file = new File([blob], "2025品檢報表統計.xlsx");
        handleLoadSummaryFile(file);
      })
      .catch(err => {
        console.log("Auto-load of pre-generated 2025 summary Excel not available (can upload manually).");
      });
  }, []);

  // Update mappings helper
  const updateMappings = (newMappings) => {
    setMappings(newMappings);
    saveMappings(newMappings);
  };

  // Sync localStorage
  const handleAddMapping = (e) => {
    e.preventDefault();
    if (!newCode || !newName) return;
    const formattedCode = newCode.trim().toUpperCase();
    if (!/^QC\d{5}-R\d{2}$/i.test(formattedCode)) {
      alert("表單編碼格式必須為 QCxxxxx-Rxx (例如 QC10001-R01)！");
      return;
    }
    const updated = { ...mappings, [formattedCode]: newName.trim() };
    updateMappings(updated);
    setNewCode("");
    setNewName("");
  };

  const handleDeleteMapping = (code) => {
    const updated = { ...mappings };
    delete updated[code];
    updateMappings(updated);
  };

  const handleResetMappings = () => {
    if (window.confirm("確定要將對照表恢復為系統預設值嗎？自訂對照將會遺失。")) {
      const reset = resetMappings();
      setMappings(reset);
    }
  };

  const handleImportMappings = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const json = JSON.parse(event.target.result);
        const merged = { ...mappings, ...json };
        updateMappings(merged);
        alert("匯入對照表成功！");
      } catch (err) {
        alert("檔案格式錯誤，請確保匯入正確的 JSON 檔案！");
      }
    };
    reader.readAsText(file);
    e.target.value = null;
  };

  const handleExportMappings = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(mappings, null, 2));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute("download", "qc_mappings_export.json");
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
  };

  // Process folder extraction
  const processFilesList = async (files, nameOfFolder) => {
    if (files.length === 0) return;
    setIsScanning(true);
    setFolderName(nameOfFolder || "個別檔案統計");
    
    const excelFiles = Array.from(files).filter(f => 
      f.name.endsWith('.xlsx') || f.name.endsWith('.xls') || f.name.endsWith('.xlsm')
    );
    
    if (excelFiles.length === 0) {
      alert("選取的對象中沒有可支援的 Excel 檔案！");
      setIsScanning(false);
      return;
    }
    
    const results = [];
    for (let file of excelFiles) {
      try {
        const fileRes = await parseExcelFile(file, mappings);
        if (fileRes.success) {
          results.push(...fileRes.sheets);
        } else {
          results.push({
            fileName: file.name,
            sheetName: "N/A",
            foundCode: "錯誤",
            foundName: fileRes.error || "讀取失敗",
            status: "error"
          });
        }
      } catch (e) {
        results.push({
          fileName: file.name,
          sheetName: "N/A",
          foundCode: "錯誤",
          foundName: "開啟失敗",
          status: "error"
        });
      }
    }
    
    setScannedRows(results);
    setIsScanning(false);
  };

  const handleFolderChange = (e) => {
    if (e.target.files.length === 0) return;
    const firstFile = e.target.files[0];
    const path = firstFile.webkitRelativePath || "";
    const folder = path.split('/')[0] || "資料夾夾檔案統計";
    processFilesList(e.target.files, folder);
  };

  const handleFilesChange = (e) => {
    processFilesList(e.target.files, "個別選取檔案");
  };

  const handleDragOver = (e) => {
    e.preventDefault();
  };

  const handleDrop = (e) => {
    e.preventDefault();
    if (e.dataTransfer.files) {
      processFilesList(e.dataTransfer.files, "拖曳檔案統計");
    }
  };

  // ==========================================
  // MCKINSEY SUMMARY EXCEL PARSER HANDLERS
  // ==========================================
  const handleLoadSummaryFile = async (file) => {
    try {
      const data = await parseSummaryExcel(file);
      setSummaryFileName(file.name);
      setSummaryData(data);
      
      // Auto select first sheet (excluding 品檢地圖 if possible)
      const sheets = Object.keys(data);
      const initialSheet = sheets.find(s => s !== "品檢地圖") || sheets[0] || "";
      setActiveSheet(initialSheet);
    } catch (err) {
      console.error(err);
      alert("解析品檢彙總 Excel 失敗，請確保上傳的是正確的報表統計檔。");
    }
  };

  const handleSummaryFileChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      handleLoadSummaryFile(file);
    }
  };

  // Handle active sheet changing
  useEffect(() => {
    if (!summaryData || !activeSheet) return;
    
    const rows = summaryData[activeSheet];
    if (!rows || rows.length < 2) {
      setSelectedItems([]);
      return;
    }
    
    // Row 1 contains column headers
    const headerRow = rows[1] || [];
    
    // Exclude columns like Month (月份), Grand Total (小計), NCA or empty spacers
    const available = [];
    headerRow.forEach((h, idx) => {
      const name = String(h || '').trim();
      if (name && name !== '月份' && name !== '小計' && name !== 'NCA') {
        available.push({ name, idx });
      }
    });
    
    // Auto select first 3 items by default to initialize the plot beautifully
    setSelectedItems(available.slice(0, 3));
  }, [activeSheet, summaryData]);

  // Handle plotting Chart.js
  useEffect(() => {
    if (!summaryData || !activeSheet || selectedItems.length === 0 || !chartRef.current) {
      if (chartInstance.current) {
        chartInstance.current.destroy();
        chartInstance.current = null;
      }
      return;
    }

    const rows = summaryData[activeSheet];
    const ctx = chartRef.current.getContext('2d');

    // Create datasets
    const datasets = selectedItems.map((item, idx) => {
      // Monthly data is at rows 2 to 13 (index 2-13)
      const dataPoints = [];
      for (let m = 2; m <= 13; m++) {
        const val = rows[m] ? Number(rows[m][item.idx]) || 0 : 0;
        dataPoints.push(val);
      }

      return {
        label: item.name,
        data: dataPoints,
        borderColor: MCK_COLORS[idx % MCK_COLORS.length],
        backgroundColor: MCK_COLORS[idx % MCK_COLORS.length] + '15', // light transparent fill
        borderWidth: 2.5,
        tension: 0.1, // slightly curved lines
        pointRadius: 4,
        pointHoverRadius: 6,
      };
    });

    if (chartInstance.current) {
      chartInstance.current.destroy();
    }

    // McK style: Minimal grid, clean legend and title
    chartInstance.current = new window.Chart(ctx, {
      type: 'line',
      data: {
        labels: MONTH_LABELS,
        datasets: datasets
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'top',
            labels: {
              font: {
                family: 'Outfit, sans-serif',
                size: 12
              },
              usePointStyle: true,
              boxWidth: 8
            }
          },
          tooltip: {
            padding: 12,
            backgroundColor: '#0A192F',
            titleFont: { family: 'Outfit, sans-serif', weight: 'bold' },
            bodyFont: { family: 'Outfit, sans-serif' }
          }
        },
        scales: {
          x: {
            grid: {
              display: false // McKinsey style hides X vertical gridlines
            },
            ticks: {
              color: '#5A6A85',
              font: { family: 'Outfit, sans-serif', size: 11 }
            }
          },
          y: {
            beginAtZero: true,
            grid: {
              color: '#F1F5F9' // light grey thin gridlines
            },
            ticks: {
              color: '#5A6A85',
              font: { family: 'Outfit, sans-serif', size: 11 }
            }
          }
        }
      }
    });

    // Calculate executive insights
    let totalSum = 0;
    let peakValue = 0;
    let peakM = "";
    
    // Sum month by month
    for (let m = 0; m < 12; m++) {
      let monthlySum = 0;
      selectedItems.forEach(item => {
        const val = rows[m + 2] ? Number(rows[m + 2][item.idx]) || 0 : 0;
        monthlySum += val;
      });
      totalSum += monthlySum;
      if (monthlySum > peakValue) {
        peakValue = monthlySum;
        peakM = MONTH_LABELS[m];
      }
    }
    
    setDashboardInsights({
      total: totalSum,
      peakMonth: peakM || "N/A",
      peakVal: peakValue,
      avg: Math.round(totalSum / 12)
    });

  }, [selectedItems, activeSheet, summaryData]);

  // Toggle dynamic selection pill
  const handleTogglePill = (item) => {
    const isSelected = selectedItems.some(i => i.idx === item.idx);
    if (isSelected) {
      setSelectedItems(selectedItems.filter(i => i.idx !== item.idx));
    } else {
      setSelectedItems([...selectedItems, item]);
    }
  };

  // Render Table content for Dashboard
  const getDashboardTableData = () => {
    if (!summaryData || !activeSheet || selectedItems.length === 0) return { columns: [], rows: [] };
    const rows = summaryData[activeSheet];
    
    const columns = ["月份", ...selectedItems.map(i => i.name)];
    const tableRows = [];
    
    // Monthly rows (indices 2 to 13)
    for (let m = 2; m <= 13; m++) {
      const monthRow = rows[m];
      const rowArr = [monthRow ? monthRow[0] : `${m-1}月`];
      selectedItems.forEach(item => {
        rowArr.push(monthRow ? Number(monthRow[item.idx]) || 0 : 0);
      });
      tableRows.push(rowArr);
    }

    // Totals row (Row 14)
    const totalsRow = rows[14];
    const totalArr = ["小計"];
    selectedItems.forEach(item => {
      totalArr.push(totalsRow ? Number(totalsRow[item.idx]) || 0 : 0);
    });
    tableRows.push(totalArr);

    return { columns, rows: tableRows };
  };

  const { columns: dbCols, rows: dbRows } = getDashboardTableData();

  // Search/Filters for Tab 2
  const filteredRows = scannedRows.filter(row => {
    const matchesSearch = 
      row.fileName.toLowerCase().includes(searchQuery.toLowerCase()) ||
      row.sheetName.toLowerCase().includes(searchQuery.toLowerCase()) ||
      row.foundCode.toLowerCase().includes(searchQuery.toLowerCase()) ||
      row.foundName.toLowerCase().includes(searchQuery.toLowerCase());
      
    if (statusFilter === "all") return matchesSearch;
    return matchesSearch && row.status === statusFilter;
  });

  return (
    <div className="app-container">
      {/* App Header */}
      <header className="app-header">
        <div className="app-title-group">
          <h1 className="mck-serif-title">Mouldex 品管報表統計系統</h1>
          <p>Mouldex QC Record Extraction & Dynamic McKinsey Dashboard</p>
        </div>
        <span className="badge-version">v2026.06.28</span>
      </header>

      {/* Main Tab Mode Toggle */}
      <nav className="app-nav">
        <button 
          className={`nav-tab ${activeTab === 'dashboard' ? 'active' : ''}`}
          onClick={() => setActiveTab('dashboard')}
        >
          📊 McKinsey 品檢分析儀表板
        </button>
        <button 
          className={`nav-tab ${activeTab === 'extractor' ? 'active' : ''}`}
          onClick={() => setActiveTab('extractor')}
        >
          ⚙️ 品檢編碼對照與提取工具
        </button>
      </nav>

      {/* Tab 1: McKinsey Dashboard */}
      {activeTab === 'dashboard' && (
        <div className="mck-main-content">
          {/* Uploader Card */}
          <div className="mck-card" style={{ backgroundColor: '#F8FAFC' }}>
            <div className="mck-card-header" style={{ marginBottom: '12px' }}>
              <div>
                <h2 className="mck-card-title">匯入年度報表統計檔</h2>
                <div className="mck-card-subtitle">請上傳由 ETL Pipeline 產出的 2025/2026 品檢報表統計.xlsx 檔案</div>
              </div>
              {summaryFileName && (
                <div style={{ fontSize: '13px', color: 'var(--color-success)', fontWeight: '600' }}>
                  ✓ 已載入: {summaryFileName}
                </div>
              )}
            </div>
            
            <div style={{ display: 'flex', gap: '16px', alignItems: 'center' }}>
              <input 
                type="file" 
                accept=".xlsx" 
                onChange={handleSummaryFileChange} 
                id="summary-file-input" 
                style={{ display: 'none' }}
              />
              <label 
                htmlFor="summary-file-input" 
                className="btn btn-primary"
                style={{ minHeight: '40px', padding: '8px 16px', fontSize: '13px' }}
              >
                📁 選擇報表檔案...
              </label>
              <span style={{ fontSize: '13px', color: 'var(--mck-slate)' }}>
                或在本地執行 <code>node etl_pipeline.cjs all</code> 產出後，拖入檔案進行互動解讀。
              </span>
            </div>
          </div>

          {summaryData ? (
            <div className="mck-dashboard">
              {/* Sidebar Tabs */}
              <aside className="mck-sidebar">
                <div style={{ fontSize: '11px', fontWeight: '700', textTransform: 'uppercase', letterSpacing: '0.6px', color: 'var(--mck-slate)', padding: '0 8px 8px 8px', borderBottom: '1px solid var(--mck-border)' }}>
                  品檢報表工作表
                </div>
                {Object.keys(summaryData).map((sheetName) => (
                  <button
                    key={sheetName}
                    className={`mck-sheet-btn ${activeSheet === sheetName ? 'active' : ''}`}
                    onClick={() => setActiveSheet(sheetName)}
                  >
                    📝 {sheetName}
                  </button>
                ))}
              </aside>

              {/* Main Analysis Panels */}
              <div className="mck-main-content">
                {/* Checkbox pills */}
                <div className="mck-card">
                  <div className="mck-card-header">
                    <div>
                      <h2 className="mck-card-title">自訂篩選分析項目</h2>
                      <div className="mck-card-subtitle">點選下方標籤以新增或移除圖表分析的統計項目</div>
                    </div>
                  </div>
                  <div className="mck-pill-container">
                    {(() => {
                      const rows = summaryData[activeSheet];
                      if (!rows || rows.length < 2) return <p style={{ fontSize: '13px', color: 'var(--mck-slate)' }}>該工作表無有效數據欄位</p>;
                      const headerRow = rows[1] || [];
                      const items = [];
                      headerRow.forEach((h, idx) => {
                        const name = String(h || '').trim();
                        if (name && name !== '月份' && name !== '小計' && name !== 'NCA') {
                          items.push({ name, idx });
                        }
                      });
                      return items.map((item) => {
                        const active = selectedItems.some(i => i.idx === item.idx);
                        return (
                          <div 
                            key={item.idx} 
                            className={`mck-pill ${active ? 'active' : ''}`}
                            onClick={() => handleTogglePill(item)}
                          >
                            <span className="mck-pill-checkbox"></span>
                            {item.name}
                          </div>
                        );
                      });
                    })()}
                  </div>
                </div>

                {/* Plot Panel */}
                <div className="mck-card">
                  <div className="mck-card-header">
                    <div>
                      <h2 className="mck-card-title">{activeSheet} 數據趨勢</h2>
                      <div className="mck-card-subtitle">以月份為 X 軸，數量為 Y 軸呈現趨勢</div>
                    </div>
                  </div>
                  {selectedItems.length > 0 ? (
                    <div className="mck-chart-container">
                      <canvas ref={chartRef}></canvas>
                    </div>
                  ) : (
                    <div style={{ textAlign: 'center', padding: '64px 0', color: 'var(--mck-slate)' }}>
                      ⚠️ 請在上方選擇至少一個分析項目以繪製圖表。
                    </div>
                  )}
                </div>

                {/* Executive Insights Deck */}
                {selectedItems.length > 0 && (
                  <div className="mck-kpi-deck">
                    <div className="mck-kpi-card">
                      <div className="mck-kpi-label">篩選項目總筆數</div>
                      <div className="mck-kpi-val">{dashboardInsights.total.toLocaleString()}</div>
                      <div className="mck-kpi-desc">已選取子項累計</div>
                    </div>
                    <div className="mck-kpi-card">
                      <div className="mck-kpi-label">月均量</div>
                      <div className="mck-kpi-val">{dashboardInsights.avg.toLocaleString()}</div>
                      <div className="mck-kpi-desc">月平均品檢次數</div>
                    </div>
                    <div className="mck-kpi-card">
                      <div className="mck-kpi-label">高峰月份</div>
                      <div className="mck-kpi-val">{dashboardInsights.peakVal.toLocaleString()}</div>
                      <div className="mck-kpi-desc">最高峰出現在 {dashboardInsights.peakMonth}</div>
                    </div>
                  </div>
                )}

                {/* Data Table Card */}
                {selectedItems.length > 0 && dbRows.length > 0 && (
                  <div className="mck-card">
                    <div className="mck-card-header">
                      <h2 className="mck-card-title">數據明細報表</h2>
                    </div>
                    <div className="mck-table-wrapper">
                      <table className="mck-table">
                        <thead>
                          <tr>
                            {dbCols.map((col, idx) => (
                              <th key={idx} style={{ textAlign: idx === 0 ? 'left' : 'right' }}>{col}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {dbRows.map((row, rIdx) => {
                            const isTotal = row[0] === '小計';
                            return (
                              <tr key={rIdx} className={isTotal ? 'total-row' : ''}>
                                <td style={{ fontWeight: isTotal ? '700' : 'normal' }}>{row[0]}</td>
                                {row.slice(1).map((val, cIdx) => (
                                  <td key={cIdx} style={{ textAlign: 'right' }}>
                                    {val.toLocaleString()}
                                  </td>
                                ))}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

              </div>
            </div>
          ) : (
            <div className="panel-card" style={{ padding: '64px', textAlign: 'center' }}>
              <span style={{ fontSize: '48px' }}>📂</span>
              <h3 style={{ margin: '16px 0 8px 0', fontSize: '18px', fontWeight: '600' }}>尚未載入品檢統計報表</h3>
              <p style={{ color: 'var(--mck-slate)', maxWidth: '500px', margin: '0 auto 24px auto', fontSize: '14px' }}>
                請點擊上方按鈕，載入生成的 <code>2025品檢報表統計.xlsx</code>。載入後，系統將依麥肯錫風格為您呈現自訂動態製圖與數據分析。
              </p>
            </div>
          )}
        </div>
      )}

      {/* Tab 2: Original Extraction & Mappings Extractor */}
      {activeTab === 'extractor' && (
        <div className="main-grid">
          {/* Left Panel: Files Scanning controls */}
          <aside className="mck-main-content">
            {/* Folder scanner card */}
            <section className="panel-card">
              <h2 className="card-title">📁 品管檔案夾掃描</h2>
              
              <div 
                className="upload-zone"
                onDragOver={handleDragOver}
                onDrop={handleDrop}
              >
                <div className="upload-icon">📥</div>
                <p style={{ fontWeight: 500 }}>拖曳檔案或資料夾至此</p>
                <p style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>支援批次讀取多個 Excel 檔案</p>
                
                <div className="upload-btn-group">
                  <button 
                    className="btn btn-secondary" 
                    style={{ flex: 1 }}
                    onClick={() => folderInputRef.current.click()}
                  >
                    選取資料夾
                  </button>
                  <button 
                    className="btn btn-secondary" 
                    style={{ flex: 1 }}
                    onClick={() => fileInputRef.current.click()}
                  >
                    選擇多檔案
                  </button>
                </div>

                <input 
                  type="file" 
                  ref={folderInputRef} 
                  webkitdirectory="true" 
                  directory="true" 
                  multiple 
                  onChange={handleFolderChange} 
                  style={{ display: 'none' }} 
                />
                <input 
                  type="file" 
                  ref={fileInputRef} 
                  multiple 
                  onChange={handleFilesChange} 
                  style={{ display: 'none' }} 
                />
              </div>

              {scannedRows.length > 0 && (
                <div style={{ marginTop: '16px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  <div style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
                    已加載: <strong>{folderName}</strong>
                  </div>
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <button 
                      className="btn btn-primary" 
                      style={{ flex: 1, minHeight: '36px' }}
                      onClick={() => exportToExcel(scannedRows, folderName)}
                    >
                      💾 匯出 Excel
                    </button>
                    <button 
                      className="btn btn-danger btn-icon" 
                      onClick={() => setScannedRows([])}
                      title="清除結果"
                    >
                      🗑
                    </button>
                  </div>
                </div>
              )}
            </section>

            {/* Mappings Editor */}
            <section className="panel-card">
              <h2 className="card-title">⚙️ 表單編碼與名稱對照表</h2>
              
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '12px' }}>
                <button className="btn btn-secondary" style={{ fontSize: '12px', minHeight: '32px', padding: '4px 8px' }} onClick={handleExportMappings}>
                  📤 匯出 JSON
                </button>
                <label className="btn btn-secondary" style={{ fontSize: '12px', minHeight: '32px', padding: '4px 8px', cursor: 'pointer' }}>
                  📥 匯入 JSON
                  <input type="file" accept=".json" onChange={handleImportMappings} style={{ display: 'none' }} />
                </label>
                <button className="btn btn-danger" style={{ fontSize: '12px', minHeight: '32px', padding: '4px 8px' }} onClick={handleResetMappings}>
                  🔄 恢復預設
                </button>
              </div>

              <div className="mapping-list">
                {Object.keys(mappings).map((code) => (
                  <div key={code} className="mapping-item">
                    <span className="mapping-input mapping-code">{code}</span>
                    <input 
                      type="text" 
                      className="mapping-input mapping-name" 
                      value={mappings[code]} 
                      onChange={(e) => {
                        const updated = { ...mappings, [code]: e.target.value };
                        updateMappings(updated);
                      }} 
                    />
                    <button 
                      className="btn btn-danger" 
                      style={{ minHeight: '32px', width: '32px', padding: 0 }}
                      onClick={() => handleDeleteMapping(code)}
                    >
                      ✕
                    </button>
                  </div>
                ))}
              </div>

              <form onSubmit={handleAddMapping} style={{ borderTop: '1px solid var(--color-border)', paddingTop: '16px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                <div style={{ fontSize: '12px', fontWeight: 600, color: 'var(--text-secondary)' }}>新增對照關係：</div>
                <div style={{ display: 'flex', gap: '8px' }}>
                  <input 
                    type="text" 
                    className="mapping-input" 
                    placeholder="QC10006-R02" 
                    value={newCode} 
                    onChange={(e) => setNewCode(e.target.value)} 
                    style={{ width: '120px' }}
                  />
                  <input 
                    type="text" 
                    className="mapping-input" 
                    placeholder="對應表單名稱" 
                    value={newName} 
                    onChange={(e) => setNewName(e.target.value)} 
                    style={{ flex: 1 }}
                  />
                </div>
                <button type="submit" className="btn btn-primary" style={{ width: '100%', minHeight: '36px' }}>
                  ＋ 新增至對照表
                </button>
              </form>
            </section>
          </aside>

          {/* Right Panel: Result data table */}
          <section className="panel-card" style={{ display: 'flex', flexDirection: 'column', minHeight: '500px' }}>
            <div className="table-controls">
              <h2 className="card-title" style={{ margin: 0 }}>📊 工作表表單編碼提取結果</h2>
              
              <div className="filter-group">
                <input 
                  type="text" 
                  className="search-input" 
                  placeholder="搜尋結果..." 
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                />
                
                <select 
                  className="filter-select"
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value)}
                >
                  <option value="all">所有狀態</option>
                  <option value="matched">成功識別</option>
                  <option value="unmatched">缺對照</option>
                  <option value="none">無編碼</option>
                  <option value="error">讀取錯誤</option>
                </select>
              </div>
            </div>

            {isScanning ? (
              <div className="empty-state" style={{ padding: '128px 0' }}>
                <div className="upload-icon" style={{ animation: 'spin 2s linear infinite' }}>🔄</div>
                <p style={{ fontWeight: 500 }}>正在解析 Excel 報表，請稍候...</p>
                <p style={{ color: 'var(--text-secondary)', fontSize: '13px' }}>這通常只需要幾秒鐘，依檔案數量及大小與定</p>
              </div>
            ) : filteredRows.length > 0 ? (
              <div className="table-wrapper">
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>序號</th>
                      <th>檔案名稱</th>
                      <th>工作表名稱</th>
                      <th>表單編碼</th>
                      <th>表單對照名稱</th>
                      <th>狀態</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRows.map((row, idx) => (
                      <tr key={idx}>
                        <td style={{ color: 'var(--text-secondary)' }}>{idx + 1}</td>
                        <td style={{ fontWeight: 500 }}>{row.fileName}</td>
                        <td>{row.sheetName}</td>
                        <td style={{ fontFamily: 'monospace', fontWeight: 600 }}>{row.foundCode}</td>
                        <td>{row.foundName}</td>
                        <td>
                          {row.status === 'matched' && (
                            <span className="status-badge status-matched">✓ 成功識別</span>
                          )}
                          {row.status === 'unmatched' && (
                            <span className="status-badge status-unmatched">⚠ 缺對照</span>
                          )}
                          {row.status === 'none' && (
                            <span className="status-badge status-none">無編碼</span>
                          )}
                          {row.status === 'error' && (
                            <span className="status-badge status-error">✗ 讀取錯誤</span>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="empty-state" style={{ flex: 1, display: 'flex', justifyContent: 'center' }}>
                <span className="empty-icon">📈</span>
                <h3>暫無統計資料</h3>
                <p>請由左側批次選取或拖入 Excel 檔案，即可在此處展開提取與編碼比對結果。</p>
              </div>
            )}
          </section>
        </div>
      )}

      {/* Footer */}
      <footer style={{
        marginTop: '24px',
        paddingTop: '16px',
        borderTop: '1px solid var(--mck-border)',
        textAlign: 'center',
        color: 'var(--mck-slate)',
        fontSize: '13px'
      }}>
        Developed by Wesley Chang @Mouldex, 2026.
      </footer>
    </div>
  );
}

export default App;
