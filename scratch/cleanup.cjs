const fs = require('fs');
const path = require('path');

const rootDir = path.join(__dirname, '..');

console.log('=== Starting Local MECE Cleanup ===\n');

// 1. Clean scratch folder
const scratchDir = __dirname;
const scratchKeep = ['generate_all_sheets_report.cjs', 'cleanup.cjs'];
console.log('1. Cleaning scratch folder...');
fs.readdirSync(scratchDir).forEach(file => {
  if (scratchKeep.includes(file)) return;
  const full = path.join(scratchDir, file);
  if (fs.statSync(full).isFile()) {
    try {
      fs.unlinkSync(full);
      console.log(`   Removed: scratch/${file}`);
    } catch (e) {
      console.log(`   Failed to remove scratch/${file}: ${e.message}`);
    }
  }
});

// 2. Clean DataExtract folder
const dataExtractDir = path.join(rootDir, 'DataExtract');
const dataExtractKeep = ['2025品檢報表統計.xlsx', '2026品檢報表統計.xlsx'];
if (fs.existsSync(dataExtractDir)) {
  console.log('\n2. Cleaning DataExtract folder...');
  fs.readdirSync(dataExtractDir).forEach(file => {
    if (dataExtractKeep.includes(file)) return;
    const full = path.join(dataExtractDir, file);
    if (fs.statSync(full).isFile()) {
      try {
        fs.unlinkSync(full);
        console.log(`   Removed: DataExtract/${file}`);
      } catch (e) {
        console.log(`   Failed to remove DataExtract/${file}: ${e.message}`);
      }
    }
  });
}

// 3. Remove dist/ folder (Vite build output cache)
const distDir = path.join(rootDir, 'dist');
if (fs.existsSync(distDir)) {
  console.log('\n3. Cleaning stale build directory (dist/)...');
  try {
    fs.rmSync(distDir, { recursive: true, force: true });
    console.log('   Removed: dist/ directory');
  } catch (e) {
    console.log(`   Failed to remove dist/: ${e.message}`);
  }
}

// 4. Remove temporary sync database files in root
const syncDbFile = path.join(rootDir, 'sync.ffs_db');
if (fs.existsSync(syncDbFile)) {
  console.log('\n4. Cleaning local sync database files...');
  try {
    fs.unlinkSync(syncDbFile);
    console.log('   Removed: sync.ffs_db');
  } catch (e) {
    console.log(`   Failed to remove sync.ffs_db: ${e.message}`);
  }
}

console.log('\n=== MECE Cleanup Finished Successfully ===');
