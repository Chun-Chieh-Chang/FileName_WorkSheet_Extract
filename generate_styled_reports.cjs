/**
 * Generate styled HTML reports matching report-templates/ style
 * Usage: node generate_styled_reports.cjs [year]
 *   e.g., node generate_styled_reports.cjs 2025
 *         node generate_styled_reports.cjs all
 */

var XLSX = require('xlsx');
var fs = require('fs');

var MONTHS = [1,2,3,4,5,6,7,8,9,10,11,12];
var COLORS = ['#3B82F6','#10B981','#F59E0B','#EF4444','#8B5CF6','#EC4899','#06B6D4','#84CC16','#F97316','#6366F1'];

function totalOf(obj) {
  var total = 0;
  for (var m = 1; m <= 12; m++) {
    total += obj[m] || 0;
  }
  return total;
}

function getSheet(wb, name) {
  var s = wb.Sheets[name];
  return s ? XLSX.utils.sheet_to_json(s, {header:1, defval:''}) : [];
}

function generateStyledReport(summaryFile, year, outFile) {
  console.log('  Reading: ' + summaryFile);
  var wb = XLSX.readFile(summaryFile);
  
  // Extract data from summary sheets
  var rawSheet = getSheet(wb, '原物料品檢(QC10002-R02)');
  var semiSheet = getSheet(wb, '半成品品檢(QC10006-R02)');
  var finSheet = getSheet(wb, '完成品品檢(QC10007-R01 R02)');
  var partsSheet = getSheet(wb, '零組件入庫品檢(QC10007-R03)');
  var shipSheet = getSheet(wb, '出貨檢驗(QC10008-R02)');
  
  // Parse monthly data
  var rawMain = {}, rawSub = {};
  if (rawSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = rawSheet[i];
      if (!r) continue;
      var m = i - 1;
      rawMain[m] = (r[1] !== '' && r[1] !== undefined) ? Number(r[1]) || 0 : 0;
      // Index 13 is the grand total "小計" for the new consolidated sheet
      var monthTotal = (r[13] !== '' && r[13] !== undefined) ? Number(r[13]) || 0 : 0;
      rawSub[m] = monthTotal - rawMain[m];
    }
  }
  
  var semiAssyA = {}, semiAssyB = {}, semiAssyC = {};
  var semiBD = {}, semiBM = {}, semiMP = {}, semiVV = {}, semiOther = {};
  if (semiSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = semiSheet[i];
      if (!r) continue;
      var m = i - 1;
      semiAssyA[m] = 0;
      semiAssyB[m] = 0;
      semiAssyC[m] = (r[1] !== '' && r[1] !== undefined) ? Number(r[1]) || 0 : 0;
      semiBD[m] = (r[2] !== '' && r[2] !== undefined) ? Number(r[2]) || 0 : 0;
      semiBM[m] = (r[3] !== '' && r[3] !== undefined) ? Number(r[3]) || 0 : 0;
      semiMP[m] = (r[4] !== '' && r[4] !== undefined) ? Number(r[4]) || 0 : 0;
      semiVV[m] = (r[5] !== '' && r[5] !== undefined) ? Number(r[5]) || 0 : 0;
      semiOther[m] = 0;
    }
  }
  
  var finBM = {}, finMM = {}, finSX = {}, finVV = {};
  if (finSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = finSheet[i];
      if (!r) continue;
      var m = i - 1;
      finBM[m] = (r[1] !== '' && r[1] !== undefined) ? Number(r[1]) || 0 : 0;
      finMM[m] = (r[2] !== '' && r[2] !== undefined) ? Number(r[2]) || 0 : 0;
      finSX[m] = (r[3] !== '' && r[3] !== undefined) ? Number(r[3]) || 0 : 0;
      finVV[m] = (r[4] !== '' && r[4] !== undefined) ? Number(r[4]) || 0 : 0;
    }
  }
  
  var pTub = {}, pInj = {}, pInjA = {}, pInjC = {};
  var pInjD = {}, pAsmA = {}, pAsmB = {}, pAsmC = {};
  if (partsSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = partsSheet[i];
      if (!r) continue;
      var m = i - 1;
      pTub[m] = (r[1] !== '' && r[1] !== undefined) ? Number(r[1]) || 0 : 0;
      pInj[m] = (r[2] !== '' && r[2] !== undefined) ? Number(r[2]) || 0 : 0;
      pInjA[m] = (r[3] !== '' && r[3] !== undefined) ? Number(r[3]) || 0 : 0;
      pInjC[m] = (r[4] !== '' && r[4] !== undefined) ? Number(r[4]) || 0 : 0;
      pInjD[m] = (r[5] !== '' && r[5] !== undefined) ? Number(r[5]) || 0 : 0;
      pAsmA[m] = (r[6] !== '' && r[6] !== undefined) ? Number(r[6]) || 0 : 0;
      pAsmB[m] = (r[7] !== '' && r[7] !== undefined) ? Number(r[7]) || 0 : 0;
      pAsmC[m] = (r[8] !== '' && r[8] !== undefined) ? Number(r[8]) || 0 : 0;
    }
  }
  
  var sICU = {}, sOth = {};
  if (shipSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = shipSheet[i];
      if (!r) continue;
      var m = i - 1;
      sICU[m] = (r[1] !== '' && r[1] !== undefined) ? Number(r[1]) || 0 : 0;
      sOth[m] = (r[2] !== '' && r[2] !== undefined) ? Number(r[2]) || 0 : 0;
    }
  }
  
  // Calculate category totals
  var rawMainTotal = totalOf(rawMain);
  var rawSubTotal = totalOf(rawSub);
  var semiTotal = totalOf(semiAssyA) + totalOf(semiAssyB) + totalOf(semiAssyC) + totalOf(semiBD) + totalOf(semiBM) + totalOf(semiMP) + totalOf(semiVV) + totalOf(semiOther);
  var finTotal = totalOf(finBM) + totalOf(finMM) + totalOf(finSX) + totalOf(finVV);
  var partsTotal = totalOf(pTub) + totalOf(pInj) + totalOf(pInjA) + totalOf(pInjC) + totalOf(pInjD) + totalOf(pAsmA) + totalOf(pAsmB) + totalOf(pAsmC);
  var shipTotal = totalOf(sICU) + totalOf(sOth);
  
  // Get QIP data from the summary Excel
  var qipSheet = getSheet(wb, 'QIP(QC10004-R02)');
  var injSetup = {}, extSetup = {}, injPatrol = {}, extPatrol = {};
  if (qipSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = qipSheet[i];
      if (!r) continue;
      var m = i - 1;
      extSetup[m] = (r[1] !== '' && r[1] !== undefined) ? Number(r[1]) || 0 : 0;
      injSetup[m] = (r[2] !== '' && r[2] !== undefined) ? Number(r[2]) || 0 : 0;
      extPatrol[m] = (r[6] !== '' && r[6] !== undefined) ? Number(r[6]) || 0 : 0;
      injPatrol[m] = (r[7] !== '' && r[7] !== undefined) ? Number(r[7]) || 0 : 0;
    }
  }
  
  // Get assembly patrol data
  var asmSheet = getSheet(wb, '裝配對樣巡檢(QC10006-R01)');
  var patrolAsm = {};
  if (asmSheet.length > 1) {
    for (var i = 2; i <= 13; i++) {
      var r = asmSheet[i];
      var m = i - 1;
      if (r && r[1] !== '' && r[1] !== undefined) {
        patrolAsm[m] = Number(r[1]) || 0;
      }
    }
  }
  
  // Calculate grand total
  var grandTotal = rawMainTotal + rawSubTotal + totalOf(injSetup) + totalOf(injPatrol) + totalOf(patrolAsm) + semiTotal + finTotal + partsTotal + shipTotal;
  
  // Build category data
  var catData = [
    {label:'原物料進料', total:rawMainTotal + rawSubTotal},
    {label:'射出Setup', total:totalOf(injSetup)},
    {label:'射出巡檢', total:totalOf(injPatrol)},
    {label:'裝配巡檢', total:totalOf(patrolAsm)},
    {label:'半成品', total:semiTotal},
    {label:'完成品', total:finTotal},
    {label:'零組件入庫', total:partsTotal},
    {label:'出貨檢驗', total:shipTotal},
  ];
  
  // Sort by total descending
  catData.sort(function(a,b){return b.total - a.total;});
  
  var catLabels = JSON.stringify(catData.map(function(c){return c.label;}));
  var catTotals = JSON.stringify(catData.map(function(c){return c.total;}));
  var colors = JSON.stringify(COLORS.slice(0, 8));
  
  var catTotalsArr = catData.map(function(c){return c.total;}); var maxCatIdx = catTotalsArr.indexOf(Math.max.apply(null, catTotalsArr));
  var maxCat = catData[maxCatIdx].label;
  var maxCount = catData[maxCatIdx].total;
  
  // Build category details HTML
  var catDetailsHtml = '';
  catData.forEach(function(c, i) {
    var pct = grandTotal > 0 ? (c.total / grandTotal * 100).toFixed(1) : 0;
    var color = COLORS[i % COLORS.length];
    catDetailsHtml += '<li><span class="cat"><span class="dot" style="background:' + color + '"></span>' + c.label + '</span><span><span class="val">' + c.total.toLocaleString() + '</span><span class="pct">(' + pct + '%)</span></span></li>\n';
  });
  
  // Build detail table HTML
  var detailTableHtml = '';
  detailTableHtml += '<tr><td>原物料進料</td><td class="num">' + (rawMainTotal + rawSubTotal).toLocaleString() + '</td><td class="num">' + ((rawMainTotal + rawSubTotal) / grandTotal * 100).toFixed(1) + '%</td><td>原料 ' + rawMainTotal.toLocaleString() + ' + 物料 ' + rawSubTotal.toLocaleString() + '</td></tr>\n';
  detailTableHtml += '<tr><td>零組件入庫</td><td class="num">' + partsTotal.toLocaleString() + '</td><td class="num">' + (partsTotal / grandTotal * 100).toFixed(1) + '%</td><td>射出(廠內) ' + totalOf(pInj).toLocaleString() + ' + Tubing ' + totalOf(pTub).toLocaleString() + '</td></tr>\n';
  detailTableHtml += '<tr><td>出貨檢驗</td><td class="num">' + shipTotal.toLocaleString() + '</td><td class="num">' + (shipTotal / grandTotal * 100).toFixed(1) + '%</td><td>ICU ' + totalOf(sICU).toLocaleString() + ' + 其他 ' + totalOf(sOth).toLocaleString() + '</td></tr>\n';
  detailTableHtml += '<tr><td>射出 Setup</td><td class="num">' + totalOf(injSetup).toLocaleString() + '</td><td class="num">' + (totalOf(injSetup) / grandTotal * 100).toFixed(1) + '%</td><td>月均 ' + Math.round(totalOf(injSetup)/12).toLocaleString() + ' 筆</td></tr>\n';
  detailTableHtml += '<tr><td>裝配巡檢</td><td class="num">' + totalOf(patrolAsm).toLocaleString() + '</td><td class="num">' + (totalOf(patrolAsm) / grandTotal * 100).toFixed(1) + '%</td><td>月均 ' + Math.round(totalOf(patrolAsm)/12).toLocaleString() + ' 筆</td></tr>\n';
  detailTableHtml += '<tr><td>半成品</td><td class="num">' + semiTotal.toLocaleString() + '</td><td class="num">' + (semiTotal / grandTotal * 100).toFixed(1) + '%</td><td>裝配 ' + (totalOf(semiAssyA) + totalOf(semiAssyB) + totalOf(semiAssyC)).toLocaleString() + ' + BD ' + totalOf(semiBD).toLocaleString() + ' + 其他</td></tr>\n';
  detailTableHtml += '<tr><td>完成品</td><td class="num">' + finTotal.toLocaleString() + '</td><td class="num">' + (finTotal / grandTotal * 100).toFixed(1) + '%</td><td>Vivus ' + totalOf(finVV).toLocaleString() + ' + Biometrix ' + totalOf(finBM).toLocaleString() + '</td></tr>\n';
  
  // Generate HTML
  var html = '<!DOCTYPE html>\n' +
'<html lang="zh-TW">\n' +
'<head>\n' +
'<meta charset="UTF-8">\n' +
'<meta name="viewport" content="width=device-width, initial-scale=1.0">\n' +
'<title>' + year + ' 品檢報表分析 — Stripe 儀表板 (直式)</title>\n' +
'<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">\n' +
'<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"><\/script>\n' +
'<style>\n' +
'*{margin:0;padding:0;box-sizing:border-box}\n' +
'body{background:#F1F5F9;font-family:Inter,"Noto Sans TC",sans-serif;color:#0F172A;padding:24px 32px;}\n' +
'.page{width:100%;max-width:1000px;margin:0 auto;display:flex;flex-direction:column;gap:12px}\n' +
'.header{display:flex;justify-content:space-between;align-items:center;flex-shrink:0}\n' +
'.header h1{font-size:18px;font-weight:700;letter-spacing:-0.3px;color:#0F172A}\n' +
'.header h1 span{color:#64748B;font-weight:400;font-size:13px;margin-left:10px}\n' +
'.header .badge{background:#E0E7FF;color:#4338CA;font-size:11px;font-weight:600;padding:3px 10px;border-radius:20px}\n' +
'.kpi{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;flex-shrink:0}\n' +
'.kpi-card{background:#fff;border-radius:8px;padding:10px 16px;box-shadow:0 1px 2px rgba(0,0,0,.05);display:flex;align-items:center;justify-content:space-between}\n' +
'.kpi-card .kl{font-size:11px;font-weight:500;color:#64748B;text-transform:uppercase;letter-spacing:.4px}\n' +
'.kpi-card .kv{font-size:22px;font-weight:700;color:#0F172A}\n' +
'.kpi-card .kd{font-size:11px;font-weight:500;color:#64748B}\n' +
'.main{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:4px}\n' +
'.card{background:#fff;border-radius:8px;padding:14px;box-shadow:0 1px 2px rgba(0,0,0,.05);display:flex;flex-direction:column}\n' +
'.card-title{font-size:12px;font-weight:600;color:#1E293B;margin-bottom:8px;display:flex;justify-content:space-between;align-items:center;flex-shrink:0}\n' +
'.card-title .tag{font-size:10px;font-weight:500;color:#64748B;background:#F1F5F9;padding:1px 6px;border-radius:4px}\n' +
'.chart-wrap{flex:1;min-height:0;position:relative}\n' +
'.insight-list{list-style:none;flex:1;overflow-y:auto}\n' +
'.insight-list li{display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #F1F5F9;font-size:12px}\n' +
'.insight-list li:last-child{border:none}\n' +
'.insight-list .cat{font-weight:500;color:#334155;display:flex;align-items:center;gap:6px}\n' +
'.insight-list .cat .dot{width:6px;height:6px;border-radius:50%;flex-shrink:0}\n' +
'.insight-list .val{font-weight:600;color:#0F172A}\n' +
'.insight-list .pct{color:#94A3B8;margin-left:4px;font-size:11px}\n' +
'table{width:100%;border-collapse:collapse;font-size:11px}\n' +
'thead th{background:#F8FAFC;color:#64748B;font-weight:600;padding:6px 8px;text-align:left;border-bottom:2px solid #E2E8F0;white-space:nowrap;font-size:10px;text-transform:uppercase;letter-spacing:.3px}\n' +
'tbody td{padding:5px 8px;border-bottom:1px solid #F1F5F9;color:#334155}\n' +
'tbody tr:hover td{background:#F8FAFC}\n' +
'tbody .num{font-weight:600;color:#0F172A;text-align:right;font-variant-numeric:tabular-nums}\n' +
'.bar-mini{height:3px;background:#E2E8F0;border-radius:2px;overflow:hidden}\n' +
'.bar-mini .f{height:100%;border-radius:2px}\n' +
'.table-wrap{flex:1;overflow-y:auto;min-height:0}\n' +
'<\/style>\n' +
'</head>\n' +
'<body>\n' +
'<div class="page">\n' +
'  <div class="header">\n' +
'    <h1>' + year + ' 品檢報表分析 <span>Dashboard</span></h1>\n' +
'    <span class="badge">全年彙總</span>\n' +
'  </div>\n' +
'  <div class="kpi">\n' +
'    <div class="kpi-card"><div><div class="kl">總筆數</div><div class="kv">' + grandTotal.toLocaleString() + '</div></div><div class="kd">' + year + ' 全年</div></div>\n' +
'    <div class="kpi-card"><div><div class="kl">最大類別</div><div class="kv">' + maxCat + '</div></div><div class="kd">' + maxCount.toLocaleString() + ' 筆</div></div>\n' +
'    <div class="kpi-card"><div><div class="kl">月均量</div><div class="kv">' + Math.round(grandTotal/12).toLocaleString() + '</div></div><div class="kd">最高 8月 558</div></div>\n' +
'    <div class="kpi-card"><div><div class="kl">類別總數</div><div class="kv">8</div></div><div class="kd">8 大類別</div></div>\n' +
'  </div>\n' +
'  <div class="main">\n' +
'    <div class="card">\n' +
'      <div class="card-title">' + year + ' 各類別總量 <span class="tag">水平長條圖</span></div>\n' +
'      <div class="chart-wrap"><canvas id="c0"><\/canvas></div>\n' +
'    </div>\n' +
'    <div class="card">\n' +
'      <div class="card-title">類別分布 <span class="tag">佔比</span></div>\n' +
'      <ul class="insight-list">\n' +
catDetailsHtml +
'      </ul>\n' +
'    </div>\n' +
'    <div class="card">\n' +
'      <div class="card-title">細項明細 <span class="tag">備註</span></div>\n' +
'      <div class="table-wrap">\n' +
'        <table>\n' +
'          <thead><tr><th>類別</th><th>數量</th><th>佔比</th><th>備註</th></tr></thead>\n' +
'          <tbody>\n' +
detailTableHtml +
'          </tbody>\n' +
'        </table>\n' +
'      </div>\n' +
'    </div>\n' +
'  </div>\n' +
'</div>\n' +
'<script>\n' +
'new Chart(document.getElementById("c0"),{type:"bar",data:{labels:' + catLabels + ',datasets:[{label:"' + year + '",data:' + catTotals + ',backgroundColor:' + colors + ',borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,grid:{color:"#F1F5F9"},ticks:{font:{size:10}}},y:{grid:{display:false},ticks:{font:{size:11}}}}}});\\n' +
'<\/script>\n' +
'</body>\n' +
'</html>';

  fs.writeFileSync(outFile, html, 'utf8');
  console.log('  Styled HTML report: ' + outFile);
}

// Main
var years = process.argv[2] === 'all' ? [2025, 2026] : [parseInt(process.argv[2], 10)];
years.forEach(function(year) {
  var summaryFile = 'DataExtract/' + year + '品檢報表統計.xlsx';
  var outFile = year + '品檢報表分析.html';
  if (fs.existsSync(summaryFile)) {
    generateStyledReport(summaryFile, year, outFile);
  } else {
    console.log('  Summary file not found: ' + summaryFile);
  }
});

console.log('All done!');