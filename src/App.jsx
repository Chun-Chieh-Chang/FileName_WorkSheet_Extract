import React, { useState, useEffect, useRef } from 'react';
import { getMappings, saveMappings, resetMappings } from './utils/db';
import { parseExcelFile, exportToExcel } from './utils/excelParser';

function App() {
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

  // Load mappings on mount
  useEffect(() => {
    setMappings(getMappings());
  }, []);

  // Sync mappings to localStorage when updated
  const updateMappings = (newMappings) => {
    setMappings(newMappings);
    saveMappings(newMappings);
  };

  // Add new QC Code mapping
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

  // Delete mapping
  const handleDeleteMapping = (code) => {
    const updated = { ...mappings };
    delete updated[code];
    updateMappings(updated);
  };

  // Reset mappings to default
  const handleResetMappings = () => {
    if (window.confirm("確定要將對照表恢復為系統預設值嗎？自訂對照將會遺失。")) {
      const reset = resetMappings();
      setMappings(reset);
    }
  };

  // Import mappings from JSON file
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
    e.target.value = null; // Clear
  };

  // Export mappings as JSON
  const handleExportMappings = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(mappings, null, 2));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute("download", "qc_mappings_export.json");
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
  };

  // Process files
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

    const rows = [];
    for (let file of excelFiles) {
      const result = await parseExcelFile(file, mappings);
      if (result.success) {
        rows.push(...result.sheets);
      } else {
        rows.push({
          fileName: result.fileName,
          sheetName: "檔案無法開啟",
          foundCode: "錯誤",
          foundName: result.error,
          status: "error"
        });
      }
    }
    
    setScannedRows(rows);
    setIsScanning(false);
  };

  // Handle Folder selection
  const handleFolderChange = (e) => {
    const files = e.target.files;
    if (files.length === 0) return;
    
    // Get folder name from path of first file
    let nameOfFolder = "品管報表統計";
    if (files[0].webkitRelativePath) {
      nameOfFolder = files[0].webkitRelativePath.split('/')[0];
    }
    
    processFilesList(files, nameOfFolder);
  };

  // Handle individual files selection
  const handleFileChange = (e) => {
    const files = e.target.files;
    processFilesList(files, "已選取檔案");
  };

  // Clear Scan results
  const handleClearResults = () => {
    setScannedRows([]);
    setFolderName("");
  };

  // Export extraction table to Excel
  const handleExportTable = () => {
    if (scannedRows.length === 0) return;
    exportToExcel(scannedRows, folderName);
  };

  // Statistics calculation
  const totalFiles = Array.from(new Set(scannedRows.map(r => r.fileName))).length;
  const totalSheets = scannedRows.length;
  const matchedCount = scannedRows.filter(r => r.status === "matched").length;
  const unmatchedCount = scannedRows.filter(r => r.status === "unmatched").length;
  const noneCount = scannedRows.filter(r => r.status === "none").length;
  const errorCount = scannedRows.filter(r => r.status === "error").length;
  const matchRate = totalSheets > 0 ? ((matchedCount / totalSheets) * 100).toFixed(1) : "0.0";

  // Filter and Search rows
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
      {/* Header */}
      <header className="app-header">
        <div className="app-title-group">
          <h1>📊 品管報表統計與提取工具</h1>
          <p>自動掃描 Excel 檔案，分析工作表並智能提取 QC 表單編碼與名稱</p>
        </div>
        <span className="badge-version">v2.0 Web Edition</span>
      </header>

      {/* Main Content Dashboard */}
      <div className="main-grid">
        
        {/* Left Side: Controls & Mapping Editor */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
          
          {/* File loader Card */}
          <section className="panel-card">
            <h2 className="card-title">📁 選擇數據源</h2>
            <div className="upload-zone" onClick={() => folderInputRef.current?.click()}>
              <span className="upload-icon">📂</span>
              <p style={{ fontWeight: 500 }}>點擊此處選擇整包「品管資料夾」</p>
              <p style={{ color: 'var(--text-secondary)', fontSize: '13px' }}>支援自動讀取子資料夾內的所有 Excel 檔案</p>
              <input 
                type="file" 
                ref={folderInputRef}
                style={{ display: 'none' }}
                webkitdirectory="" 
                directory="" 
                multiple
                onChange={handleFolderChange}
              />
            </div>
            
            <div style={{ margin: '16px 0', textAlign: 'center', color: 'var(--text-secondary)' }}>—— 或 ——</div>
            
            <div className="upload-zone" style={{ padding: '24px 16px' }} onClick={() => fileInputRef.current?.click()}>
              <span className="upload-icon" style={{ fontSize: '24px' }}>📄</span>
              <p style={{ fontWeight: 500, fontSize: '13px' }}>點擊選取多個 Excel 檔案</p>
              <input 
                type="file" 
                ref={fileInputRef}
                style={{ display: 'none' }}
                multiple
                accept=".xlsx,.xls,.xlsm"
                onChange={handleFileChange}
              />
            </div>
          </section>

          {/* Mapping Editor Card */}
          <section className="panel-card">
            <h2 className="card-title">⚙️ QC 表單名稱對照表</h2>
            
            <div className="mapping-list">
              {Object.keys(mappings).map(code => (
                <div className="mapping-item" key={code}>
                  <span className="mapping-input mapping-code">{code}</span>
                  <input 
                    type="text" 
                    className="mapping-input mapping-name" 
                    value={mappings[code]} 
                    onChange={(e) => {
                      const val = e.target.value;
                      updateMappings({ ...mappings, [code]: val });
                    }}
                  />
                  <button 
                    className="btn btn-secondary btn-icon" 
                    onClick={() => handleDeleteMapping(code)}
                    title="刪除此對照"
                  >
                    🗑️
                  </button>
                </div>
              ))}
              {Object.keys(mappings).length === 0 && (
                <div style={{ color: 'var(--text-secondary)', textAlign: 'center', padding: '16px 0' }}>目前無對照項目</div>
              )}
            </div>

            {/* Add New Mapping Form */}
            <form onSubmit={handleAddMapping} style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginTop: '16px' }}>
              <div style={{ display: 'flex', gap: '8px' }}>
                <input 
                  type="text" 
                  placeholder="QC10001-R01" 
                  className="mapping-input" 
                  style={{ width: '130px' }}
                  value={newCode}
                  onChange={(e) => setNewCode(e.target.value)}
                  required
                />
                <input 
                  type="text" 
                  placeholder="表單中文名稱" 
                  className="mapping-input" 
                  style={{ flex: 1 }}
                  value={newName}
                  onChange={(e) => setNewName(e.target.value)}
                  required
                />
              </div>
              <button type="submit" className="btn btn-primary" style={{ width: '100%' }}>➕ 新增對照項目</button>
            </form>

            {/* Mapping Backup Controls */}
            <div className="upload-btn-group" style={{ marginTop: '12px' }}>
              <button className="btn btn-secondary" style={{ flex: 1, minHeight: '38px' }} onClick={handleExportMappings}>📤 備份</button>
              <button className="btn btn-secondary" style={{ flex: 1, minHeight: '38px', position: 'relative' }} onClick={() => document.getElementById('mappings-import-input').click()}>
                📥 還原
                <input 
                  id="mappings-import-input"
                  type="file" 
                  style={{ display: 'none' }}
                  accept=".json"
                  onChange={handleImportMappings}
                />
              </button>
              <button className="btn btn-danger btn-icon" style={{ minHeight: '38px', width: '38px' }} onClick={handleResetMappings} title="恢復系統預設">🔄</button>
            </div>
          </section>

        </div>

        {/* Right Side: Dashboard Stats & Results Table */}
        <div style={{ display: 'flex', flexDirection: 'column' }}>
          
          {/* Dashboard Stats */}
          <div className="dashboard-row">
            <div className="widget-card">
              <div className="widget-icon" style={{ backgroundColor: '#E0F2FE', color: '#0284C7' }}>📁</div>
              <div className="widget-info">
                <span className="widget-label">總掃描檔案</span>
                <span className="widget-val">{totalFiles} 個</span>
              </div>
            </div>
            <div className="widget-card">
              <div className="widget-icon" style={{ backgroundColor: '#F3E8FF', color: '#7C3AED' }}>📄</div>
              <div className="widget-info">
                <span className="widget-label">總工作表</span>
                <span className="widget-val">{totalSheets} 個</span>
              </div>
            </div>
            <div className="widget-card">
              <div className="widget-icon" style={{ backgroundColor: 'var(--color-success-bg)', color: 'var(--color-success)' }}>✅</div>
              <div className="widget-info">
                <span className="widget-label">成功識別表單</span>
                <span className="widget-val">{matchedCount} 個</span>
              </div>
            </div>
            <div className="widget-card">
              <div className="widget-icon" style={{ backgroundColor: '#FEF3C7', color: '#D97706' }}>📊</div>
              <div className="widget-info">
                <span className="widget-label">識別覆蓋率</span>
                <span className="widget-val">{matchRate} %</span>
              </div>
            </div>
          </div>

          {/* Results Table Section */}
          <section className="panel-card" style={{ flex: 1 }}>
            
            {/* Table Control Bar */}
            <div className="table-controls">
              <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                <h3 style={{ fontSize: '16px', fontWeight: 600 }}>
                  {folderName ? `📂 目前統計來源：${folderName}` : "📋 提取結果列表"}
                </h3>
                <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
                  共找到 {filteredRows.length} 筆項目 {statusFilter !== 'all' && `(篩選: ${statusFilter})`}
                </span>
              </div>
              
              <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap' }}>
                <input 
                  type="text" 
                  placeholder="搜尋檔案、工作表或表單編碼..." 
                  className="search-input"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                />
                
                <select 
                  className="filter-select"
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value)}
                >
                  <option value="all">🔍 顯示所有狀態</option>
                  <option value="matched">🟢 成功識別對照</option>
                  <option value="unmatched">🟡 偵測編碼但無對照</option>
                  <option value="none">⚪ 無 QC 表單編碼</option>
                  <option value="error">🔴 檔案讀取失敗</option>
                </select>
                
                {scannedRows.length > 0 && (
                  <>
                    <button className="btn btn-primary" onClick={handleExportTable}>⬇️ 導出統計 Excel</button>
                    <button className="btn btn-secondary" onClick={handleClearResults}>🗑️ 清除</button>
                  </>
                )}
              </div>
            </div>

            {/* Table Display */}
            {isScanning ? (
              <div className="empty-state">
                <div style={{
                  border: '4px solid #E5E7EB',
                  borderTop: '4px solid var(--accent-brand)',
                  borderRadius: '50%',
                  width: '40px',
                  height: '40px',
                  animation: 'spin 1s linear infinite'
                }} />
                <style>{`
                  @keyframes spin {
                    0% { transform: rotate(0deg); }
                    100% { transform: rotate(360deg); }
                  }
                `}</style>
                <p style={{ fontWeight: 500 }}>正在解析 Excel 報表，請稍候...</p>
                <p style={{ color: 'var(--text-secondary)', fontSize: '13px' }}>這通常只需要幾秒鐘，依檔案大小及數量而定</p>
              </div>
            ) : scannedRows.length > 0 ? (
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
                    {filteredRows.length === 0 && (
                      <tr>
                        <td colSpan="6" style={{ textAlign: 'center', padding: '32px', color: 'var(--text-secondary)' }}>
                          找不到符合搜尋或篩選條件的報表項目
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="empty-state">
                <span className="empty-icon">📈</span>
                <h3>暫無統計資料</h3>
                <p>請由左側拖入或選擇品管 Excel 檔案，即可在此處展開統計分析與結果提取</p>
              </div>
            )}

          </section>

        </div>

      </div>

      {/* Footer */}
      <footer style={{
        marginTop: '24px',
        paddingTop: '16px',
        borderTop: '1px solid var(--color-border)',
        textAlign: 'center',
        color: 'var(--text-secondary)',
        fontSize: '13px'
      }}>
        Developed by Wesley Chang @Mouldex, 2026.
      </footer>
    </div>
  );
}

export default App;
