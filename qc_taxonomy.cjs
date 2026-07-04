/**
 * QC 標籤表單與資料夾/檔案名稱對應圖譜
 * 
 * 用途：讓 ETL 可以根據目錄結構或檔案名稱靈活識別 QC 表單類型
 * 
 * 使用方式：
 *   node qc_taxonomy.cjs                    # 顯示完整圖譜
 *   node qc_taxonomy.cjs QC10002-R02        # 查詢特定 QC 碼的對應名稱
 *   node qc_taxonomy.cjs "進料檢驗"          # 查詢特定名稱對應的 QC 碼
 */

// ============================================================
// QC 標籤定義
// ============================================================

var QC_LABELS = {
  'QC10002-R02': {
    name: '原物料/配件進料品檢表',
    category: '進料檢驗',
    description: '原物料進料檢驗紀錄'
  },
  'QC10004-R02': {
    name: 'QIP尺寸檢驗紀錄表',
    category: '製程檢驗',
    description: '射出/押出製程 QIP 尺寸檢驗'
  },
  'QC10006-R01': {
    name: '裝配對樣巡檢記錄表',
    category: '裝配巡檢',
    description: '裝配過程對樣巡檢'
  },
  'QC10006-R02': {
    name: '半成品檢驗記錄表',
    category: '半成品檢驗',
    description: '半成品入庫前檢驗'
  },
  'QC10007-R01': {
    name: '完成品裝配品檢紀錄表',
    category: '完成品檢驗',
    description: '最終成品出貨前檢驗'
  },
  'QC10007-R03': {
    name: '零組件入庫品檢表',
    category: '零組件檢驗',
    description: '零組件（射出/押出/裝配）入庫檢驗'
  },
  'QC10008-R02': {
    name: '出貨檢驗報告',
    category: '出貨檢驗',
    description: '成品出貨前最終檢驗'
  }
};

// ============================================================
// 資料夾/檔案名稱 → QC 碼 對應表
// ============================================================

var NAME_TO_QC = {
  // 進料檢驗
  '進料檢驗': 'QC10002-R02',
  '原物料品檢': 'QC10002-R02',
  '原料': 'QC10002-R02',
  '物料': 'QC10002-R02',
  
  // 製程檢驗 (QIP)
  'QIP尺寸檢驗': 'QC10004-R02',
  '射出檢驗': 'QC10004-R02',
  '押出檢驗': 'QC10004-R02',
  '射出': 'QC10004-R02',
  '押出': 'QC10004-R02',
  
  // 裝配巡檢
  '裝配巡檢': 'QC10006-R01',
  '裝配對樣': 'QC10006-R01',
  
  // 半成品檢驗
  '半成品品檢表': 'QC10006-R02',
  '裝配檢驗': 'QC10006-R02',
  '半成品': 'QC10006-R02',
  
  // 完成品檢驗
  '完成品品檢': 'QC10007-R01',
  '完成品': 'QC10007-R01',
  
  // 零組件入庫
  '零組件入庫': 'QC10007-R03',
  '射出D': 'QC10002-R02',  // 特殊：射出D 重定向至原物料
  'Tubing': 'QC10007-R03',
  '裝配A': 'QC10007-R03',
  '裝配B': 'QC10007-R03',
  '裝配C': 'QC10007-R03',
  
  // 出貨檢驗
  '出貨檢驗': 'QC10008-R02',
  'ICU': 'QC10008-R02',
  '其他': 'QC10008-R02'
};

// ============================================================
// QC 碼 → 資料夾/檔案名稱 對應表
// ============================================================

var QC_TO_NAMES = {};
for (var name in NAME_TO_QC) {
  var qc = NAME_TO_QC[name];
  if (!QC_TO_NAMES[qc]) QC_TO_NAMES[qc] = [];
  QC_TO_NAMES[qc].push(name);
}

// ============================================================
// 查詢函數
// ============================================================

function queryByName(keyword) {
  var qc = NAME_TO_QC[keyword];
  if (qc) {
    console.log('\n【名稱查詢】');
    console.log('  關鍵字: "' + keyword + '"');
    console.log('  對應 QC 碼: ' + qc);
    if (QC_LABELS[qc]) {
      console.log('  表單名稱: ' + QC_LABELS[qc].name);
      console.log('  類別: ' + QC_LABELS[qc].category);
      console.log('  說明: ' + QC_LABELS[qc].description);
    }
    return true;
  }
  return false;
}

function queryByQC(qcCode) {
  qcCode = qcCode.toUpperCase();
  if (QC_LABELS[qcCode]) {
    console.log('\n【QC 碼查詢】');
    console.log('  QC 碼: ' + qcCode);
    console.log('  表單名稱: ' + QC_LABELS[qcCode].name);
    console.log('  類別: ' + QC_LABELS[qcCode].category);
    console.log('  說明: ' + QC_LABELS[qcCode].description);
    
    if (QC_TO_NAMES[qcCode]) {
      console.log('\n  對應資料夾/檔案名稱：');
      QC_TO_NAMES[qcCode].forEach(function(name) {
        console.log('    - ' + name);
      });
    }
    return true;
  }
  return false;
}

// ============================================================
// 主程式
// ============================================================

var args = process.argv.slice(2);

if (args.length === 0) {
  // 顯示完整圖譜
  console.log('╔══════════════════════════════════════════════════════════╗');
  console.log('║        QC 標籤表單與資料夾/檔案名稱對應圖譜             ║');
  console.log('╚══════════════════════════════════════════════════════════╝\n');
  
  console.log('【QC 標籤定義】\n');
  for (var qc in QC_LABELS) {
    console.log('  ' + qc + ':');
    console.log('    表單名稱: ' + QC_LABELS[qc].name);
    console.log('    類別: ' + QC_LABELS[qc].category);
    console.log('    說明: ' + QC_LABELS[qc].description);
    if (QC_TO_NAMES[qc]) {
      console.log('    對應名稱: ' + QC_TO_NAMES[qc].join(', '));
    }
    console.log('');
  }
  
  console.log('【完整對應表】\n');
  console.log('  名稱 → QC 碼:');
  for (var name in NAME_TO_QC) {
    console.log('    "' + name + '" → ' + NAME_TO_QC[name]);
  }
  
} else {
  var query = args[0];
  
  // 嘗試作為 QC 碼查詢
  if (/^QC\d{5}-R\d{2}$/i.test(query)) {
    if (!queryByQC(query)) {
      console.log('找不到 QC 碼: ' + query);
    }
  }
  // 嘗試作為名稱查詢
  else if (queryByName(query)) {
    // 已顯示結果
  }
  // 都找不到，列出所有包含該關鍵字的對應
  else {
    console.log('\n找不到精確匹配，列出包含 "' + query + '" 的對應：\n');
    var found = false;
    for (var name in NAME_TO_QC) {
      if (name.indexOf(query) >= 0) {
        console.log('  "' + name + '" → ' + NAME_TO_QC[name]);
        found = true;
      }
    }
    if (!found) {
      console.log('  (無)');
    }
  }
}
