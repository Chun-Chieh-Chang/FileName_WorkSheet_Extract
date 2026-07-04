/**
 * 數據路徑圖譜建置與 QC 標籤分類工具
 * 
 * 用途：
 * 1. 遍歷用戶提供的原始數據路徑，建立結構化圖譜
 * 2. 自動識別 QC 標籤並分類檔案
 * 3. 檢核圖譜與業務邏輯匹配性
 * 4. 觸發異常修訂提示
 * 
 * 使用方式：
 *   node path_taxonomy.cjs "F:\2026 報表\進料檢驗-2026"
 *   node path_taxonomy.cjs "F:\2026 報表\進料檢驗-2026" --verbose
 *   node path_taxonomy.cjs "F:\2026 報表\進料檢驗-2026" --json
 */

const fs = require('fs');
const path = require('path');

// ============================================================
// QC 標籤對應表
// ============================================================

const QC_LABELS = {
  'QC10002-R02': {
    name: '原物料/配件進料品檢表',
    category: '進料檢驗',
    expectedSubFolders: ['原料', '物料'],
    description: '原物料進料檢驗紀錄'
  },
  'QC10004-R02': {
    name: 'QIP尺寸檢驗紀錄表',
    category: '製程檢驗',
    expectedSubFolders: ['射出', '押出', 'QIP尺寸檢驗'],
    description: '射出/押出製程 QIP 尺寸檢驗'
  },
  'QC10006-R01': {
    name: '裝配對樣巡檢記錄表',
    category: '裝配巡檢',
    expectedSubFolders: ['裝配巡檢'],
    description: '裝配過程對樣巡檢'
  },
  'QC10006-R02': {
    name: '半成品檢驗記錄表',
    category: '半成品檢驗',
    expectedSubFolders: ['半成品', '裝配檢驗'],
    description: '半成品入庫前檢驗'
  },
  'QC10007-R01': {
    name: '完成品裝配品檢紀錄表',
    category: '完成品檢驗',
    expectedSubFolders: ['完成品'],
    description: '最終成品出貨前檢驗'
  },
  'QC10007-R03': {
    name: '零組件入庫品檢表',
    category: '零組件檢驗',
    expectedSubFolders: ['零組件入庫', 'Tubing', '裝配A', '裝配B', '裝配C'],
    description: '零組件（射出/押出/裝配）入庫檢驗'
  },
  'QC10008-R02': {
    name: '出貨檢驗報告',
    category: '出貨檢驗',
    expectedSubFolders: ['出貨檢驗', 'ICU'],
    description: '成品出貨前最終檢驗'
  }
};

// 檔案副檔名與 QC 標籤的對應
const FILENAME_TO_QC = {
  '原料': 'QC10002-R02',
  '物料': 'QC10002-R02',
  '射出': 'QC10004-R02',
  '押出': 'QC10004-R02',
  'QIP': 'QC10004-R02',
  '裝配巡檢': 'QC10006-R01',
  '半成品': 'QC10006-R02',
  '裝配檢驗': 'QC10006-R02',
  '完成品': 'QC10007-R01',
  '零組件入庫': 'QC10007-R03',
  'Tubing': 'QC10007-R03',
  '出貨檢驗': 'QC10008-R02',
  'ICU': 'QC10008-R02'
};

// ============================================================
// 路徑圖譜建置
// ============================================================

function buildPathTree(rootDir, baseDepth = 0) {
  const tree = {
    name: path.basename(rootDir),
    path: rootDir,
    type: 'directory',
    children: [],
    depth: baseDepth
  };

  let entries;
  try {
    entries = fs.readdirSync(rootDir, { withFileTypes: true });
  } catch (e) {
    console.error(`錯誤：無法讀取目錄 "${rootDir}"`);
    console.error(`  原因：${e.message}`);
    process.exit(1);
  }

  // 排序：目錄在前，檔案在後
  entries.sort((a, b) => {
    if (a.isDirectory() && !b.isDirectory()) return -1;
    if (!a.isDirectory() && b.isDirectory()) return 1;
    return a.name.localeCompare(b.name);
  });

  entries.forEach(entry => {
    // 跳過隱藏檔案和目錄
    if (entry.name.startsWith('.')) return;

    const fullPath = path.join(rootDir, entry.name);
    
    if (entry.isDirectory()) {
      tree.children.push(buildPathTree(fullPath, baseDepth + 1));
    } else if (entry.isFile()) {
      const ext = path.extname(entry.name).toLowerCase();
      const stat = entry.stats || fs.statSync(fullPath);
      
      tree.children.push({
        name: entry.name,
        path: fullPath,
        type: 'file',
        extension: ext,
        size: stat.size,
        depth: baseDepth + 1,
        parentPath: rootDir
      });
    }
  });

  return tree;
}

// ============================================================
// QC 標籤識別
// ============================================================

function detectQCFromFilename(filename) {
  const nameWithoutExt = path.basename(filename, path.extname(filename));
  
  // 檢查檔案名稱是否包含 QC 標籤
  for (const qcCode in QC_LABELS) {
    if (nameWithoutExt.indexOf(qcCode) >= 0) {
      return qcCode;
    }
  }
  
  // 檢查父目錄名稱
  return null;
}

function detectQCFromFolderPath(folderPath) {
  const folderName = path.basename(folderPath);
  
  // 檢查文件名與 QC 標籤的對應
  if (FILENAME_TO_QC[folderName]) {
    return FILENAME_TO_QC[folderName];
  }
  
  // 檢查是否包含 QC 標籤
  for (const qcCode in QC_LABELS) {
    if (folderName.indexOf(qcCode) >= 0) {
      return qcCode;
    }
  }
  
  return null;
}

function assignQCLabels(tree) {
  const qcFileMap = {};
  const orphanFiles = [];
  
  function traverse(node, parentPath = '') {
    node.parentPath = parentPath;
    
    if (node.type === 'file') {
      // 嘗試從檔案名稱識別 QC 標籤
      let qcLabel = detectQCFromFilename(node.name);
      
      // 如果檔案名稱無法識別，嘗試從路徑識別
      if (!qcLabel && node.parentPath) {
        qcLabel = detectQCFromFolderPath(node.parentPath);
      }
      
      node.qcLabel = qcLabel;
      
      if (qcLabel) {
        if (!qcFileMap[qcLabel]) {
          qcFileMap[qcLabel] = [];
        }
        qcFileMap[qcLabel].push(node);
      } else {
        orphanFiles.push(node);
      }
    } else if (node.type === 'directory') {
      node.children.forEach(child => traverse(child, node.path));
    }
  }
  
  traverse(tree);
  
  return { qcFileMap, orphanFiles };
}

// ============================================================
// 路徑回溯
// ============================================================

function buildPathChain(fileNode) {
  const chain = [];
  let currentPath = fileNode.path;
  const rootPath = fileNode.path.split(path.sep).slice(0, -fileNode.depth).join(path.sep);
  
  // 從檔案所在目錄開始回溯到根目錄
  let dir = path.dirname(currentPath);
  while (dir && dir !== path.parse(dir).root) {
    chain.unshift({
      path: dir,
      name: path.basename(dir),
      level: chain.length
    });
    
    if (dir === rootPath || dir === path.parse(rootPath).root) {
      break;
    }
    dir = path.dirname(dir);
  }
  
  return chain;
}

// ============================================================
// 圖譜與業務邏輯匹配性檢核
// ============================================================

function validatePathTaxonomy(tree, qcFileMap, orphanFiles) {
  const issues = [];
  
  // 檢核 1：檢查每個 QC 標籤下的子目錄是否符合預期
  for (const qcCode in qcFileMap) {
    const expectedFolders = QC_LABELS[qcCode]?.expectedSubFolders;
    if (!expectedFolders) continue;
    
    const actualFolders = new Set();
    qcFileMap[qcCode].forEach(file => {
      const parentDir = path.basename(file.parentPath || path.dirname(file.path));
      actualFolders.add(parentDir);
    });
    
    actualFolders.forEach(folder => {
      if (!expectedFolders.includes(folder)) {
        issues.push({
          type: 'UNEXPECTED_FOLDER',
          severity: 'WARNING',
          qcLabel: qcCode,
          folder: folder,
          message: `QC 標籤 ${qcCode} 下發現未預期的資料夾: "${folder}"`,
          expected: expectedFolders.join(', '),
          suggestion: `請確認資料夾 "${folder}" 是否應歸類為 QC ${qcCode}`
        });
      }
    });
  }
  
  // 檢核 2：檢查孤兒檔案
  if (orphanFiles.length > 0) {
    issues.push({
      type: 'ORPHAN_FILES',
      severity: 'ERROR',
      count: orphanFiles.length,
      files: orphanFiles.map(f => f.path),
      message: `發現 ${orphanFiles.length} 個未分類檔案（無 QC 標籤）`,
      suggestion: '請檢查這些檔案是否應歸類到某個 QC 標籤，或更新對應表'
    });
  }
  
  // 檢核 3：檢查路徑層級深度
  function checkDepth(node, currentDepth = 0) {
    if (node.type === 'file') {
      if (currentDepth > 4) { // 預期最大深度為 4 層
        issues.push({
          type: 'DEPTH_EXCEEDED',
          severity: 'WARNING',
          file: node.path,
          depth: currentDepth,
          message: `檔案路徑層級過深 (${currentDepth} 層)，預期不超過 4 層`,
          suggestion: '請簡化目錄結構'
        });
      }
    }
    
    if (node.children) {
      node.children.forEach(child => checkDepth(child, currentDepth + 1));
    }
  }
  
  checkDepth(tree);
  
  return issues;
}

// ============================================================
// 圖譜視覺化輸出
// ============================================================

function printTree(tree, indent = 0, maxDepth = 3, verbose = false) {
  const prefix = '  '.repeat(indent);
  const icon = tree.type === 'directory' ? '📁' : '📄';
  
  if (indent <= maxDepth) {
    console.log(`${prefix}${icon} ${tree.name}`);
  }
  
  if (tree.children && indent < maxDepth) {
    tree.children.forEach(child => printTree(child, indent + 1, maxDepth, verbose));
  }
}

function printQCStats(qcFileMap, orphanFiles) {
  console.log('\n╔══════════════════════════════════════════════════════════╗');
  console.log('║              QC 標籤檔案統計                              ║');
  console.log('╚══════════════════════════════════════════════════════════╝\n');
  
  for (const qcCode in qcFileMap) {
    const label = QC_LABELS[qcCode];
    const files = qcFileMap[qcCode];
    
    console.log(`  ${qcCode}: ${files.length} 個檔案`);
    if (label) {
      console.log(`    表單名稱: ${label.name}`);
      console.log(`    類別: ${label.category}`);
    }
    
    // 顯示子目錄
    const subFolders = new Set(files.map(f => path.basename(path.dirname(f.path))));
    if (subFolders.size > 0) {
      console.log(`    子目錄: ${[...subFolders].join(', ')}`);
    }
    console.log('');
  }
  
  if (orphanFiles.length > 0) {
    console.log(`  ⚠️  未分類檔案: ${orphanFiles.length} 個`);
    orphanFiles.slice(0, 5).forEach(f => {
      console.log(`      - ${f.path}`);
    });
    if (orphanFiles.length > 5) {
      console.log(`      ... 還有 ${orphanFiles.length - 5} 個`);
    }
    console.log('');
  }
}

function printIssues(issues) {
  if (issues.length === 0) {
    console.log('\n✅ 所有檢核項目均通過，無異常！');
    return;
  }
  
  console.log('\n╔══════════════════════════════════════════════════════════╗');
  console.log('║              異常修訂提示                                ║');
  console.log('╚══════════════════════════════════════════════════════════╝\n');
  
  issues.forEach((issue, index) => {
    const severityIcon = issue.severity === 'ERROR' ? '❌' : '⚠️';
    
    console.log(`[${index + 1}] ${severityIcon} ${issue.type}`);
    console.log(`    訊息: ${issue.message}`);
    
    if (issue.qcLabel) {
      console.log(`    QC 標籤: ${issue.qcLabel}`);
    }
    
    if (issue.file) {
      console.log(`    檔案路徑: ${issue.file}`);
    }
    
    if (issue.expected) {
      console.log(`    標準規範: ${issue.expected}`);
    }
    
    if (issue.suggestion) {
      console.log(`    修正建議: ${issue.suggestion}`);
    }
    
    console.log('');
  });
  
  console.log(`總計: ${issues.filter(i => i.severity === 'ERROR').length} 個錯誤, ${issues.filter(i => i.severity === 'WARNING').length} 個警告`);
}

// ============================================================
// JSON 輸出模式
// ============================================================

function outputJSON(tree, qcFileMap, orphanFiles, issues) {
  const output = {
    metadata: {
      rootPath: tree.path,
      totalFiles: tree.children.reduce((count, node) => {
        if (node.type === 'file') return count + 1;
        if (node.children) return count + node.children.reduce((c, n) => n.type === 'file' ? c + 1 : c, 0);
        return count;
      }, 0),
      totalDirectories: tree.children.length,
      generatedAt: new Date().toISOString()
    },
    taxonomy: {
      qcLabels: {},
      orphanFiles: orphanFiles.map(f => f.path)
    },
    validation: {
      issues: issues,
      passed: issues.filter(i => i.severity === 'ERROR').length === 0
    }
  };
  
  for (const qcCode in qcFileMap) {
    output.taxonomy.qcLabels[qcCode] = {
      label: QC_LABELS[qcCode],
      fileCount: qcFileMap[qcCode].length,
      files: qcFileMap[qcCode].map(f => ({
        path: f.path,
        name: f.name,
        size: f.size,
        pathChain: buildPathChain(f)
      }))
    };
  }
  
  console.log(JSON.stringify(output, null, 2));
}

// ============================================================
// 主程式
// ============================================================

function main() {
  const args = process.argv.slice(2);
  
  if (args.length === 0) {
    console.log('用法: node path_taxonomy.cjs <數據路徑> [選項]');
    console.log('');
    console.log('參數:');
    console.log('  <數據路徑>    原始數據根目錄的絕對路徑');
    console.log('');
    console.log('選項:');
    console.log('  --verbose     顯示詳細輸出（完整路徑圖譜）');
    console.log('  --json        以 JSON 格式輸出');
    console.log('  --help        顯示幫助訊息');
    console.log('');
    console.log('範例:');
    console.log('  node path_taxonomy.cjs "F:\\2026 報表\\進料檢驗-2026"');
    console.log('  node path_taxonomy.cjs "F:\\2026 報表\\進料檢驗-2026" --verbose');
    console.log('  node path_taxonomy.cjs "F:\\2026 報表\\進料檢驗-2026" --json');
    process.exit(0);
  }
  
  const rootDir = args[0];
  const verbose = args.includes('--verbose');
  const jsonMode = args.includes('--json');
  
  // 驗證路徑是否存在
  if (!fs.existsSync(rootDir)) {
    console.error(`錯誤：目錄不存在 "${rootDir}"`);
    process.exit(1);
  }
  
  console.log('╔══════════════════════════════════════════════════════════╗');
  console.log('║        數據路徑圖譜建置與 QC 標籤分類工具                ║');
  console.log('╚══════════════════════════════════════════════════════════╝\n');
  
  console.log(`數據路徑: ${rootDir}`);
  console.log('');
  
  // Step 1: 建置路徑圖譜
  console.log('【Step 1】建置路徑圖譜...');
  const startTime = Date.now();
  const tree = buildPathTree(rootDir);
  const buildTime = Date.now() - startTime;
  
  console.log(`  ✓ 圖譜建置完成（${buildTime}ms）`);
  if (!verbose) {
    console.log('  （使用 --verbose 查看完整路徑圖譜）');
  }
  
  // Step 2: 分配 QC 標籤
  console.log('\n【Step 2】識別 QC 標籤並分類檔案...');
  const { qcFileMap, orphanFiles } = assignQCLabels(tree);
  console.log(`  ✓ 已分類 ${Object.values(qcFileMap).reduce((sum, arr) => sum + arr.length, 0)} 個檔案`);
  console.log(`  ✓ 發現 ${orphanFiles.length} 個未分類檔案`);
  
  // Step 3: 檢核匹配性
  console.log('\n【Step 3】檢核圖譜與業務邏輯匹配性...');
  const issues = validatePathTaxonomy(tree, qcFileMap, orphanFiles);
  console.log(`  ✓ 檢核完成，發現 ${issues.length} 個問題`);
  
  // Step 4: 輸出結果
  if (jsonMode) {
    console.log('\n【Step 4】以 JSON 格式輸出...');
    outputJSON(tree, qcFileMap, orphanFiles, issues);
  } else {
    console.log('\n【Step 4】輸出統計結果...');
    printTree(tree, 0, verbose ? 10 : 3, verbose);
    printQCStats(qcFileMap, orphanFiles);
    printIssues(issues);
  }
  
  console.log('\n完成！');
}

main();
