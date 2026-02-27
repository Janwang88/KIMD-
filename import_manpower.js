/**
 * 一次性导入脚本：将「2026组装外借人力统计.xlsx」历史数据导入 SQLite
 * 运行方式: node scripts/import_manpower.js
 */
const sqlite3 = require('better-sqlite3');
const XLSX = require('xlsx');
const path = require('path');

const DB_PATH = path.join(__dirname, '../data/reviews.db');
const EXCEL_PATH = path.join(__dirname, '../源表/2026组装外借人力统计.xlsx');

const db = new sqlite3(DB_PATH);
db.pragma('journal_mode = WAL');

// Excel序列号 → YYYY-MM-DD
function excelDateToStr(n) {
    if (!n || typeof n !== 'number') return null;
    const ms = Math.round((n - 25569) * 86400 * 1000);
    const d = new Date(ms);
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
}

// Excel时间小数 → HH:MM
function excelTimeToStr(t) {
    const frac = parseFloat(t);
    if (isNaN(frac)) return null;
    const mins = Math.round(frac * 1440);
    return String(Math.floor(mins / 60) % 24).padStart(2, '0') + ':' + String(mins % 60).padStart(2, '0');
}

console.log('读取 Excel 文件...');
const wb = XLSX.readFile(EXCEL_PATH);
const ws = wb.Sheets['组装外借人力'];
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

// 过滤有效行（日期是数字，姓名不为空）
const validRows = rows.slice(1).filter(r => typeof r[0] === 'number' && String(r[3]).trim() !== '');
console.log(`有效数据行: ${validRows.length} 条`);

// 检查是否已有数据，避免重复导入
const existing = db.prepare('SELECT COUNT(*) as c FROM outsource_manpower').get();
if (existing.c > 0) {
    console.log(`⚠️  数据库中已有 ${existing.c} 条记录，跳过导入以防重复。`);
    console.log('如需重新导入，请先执行: DELETE FROM outsource_manpower;');
    process.exit(0);
}

const ins = db.prepare(`
  INSERT INTO outsource_manpower 
  (work_date, work_order, project_name, worker_name, worker_level, start_time, end_time, hours, supplier, shift, manager)
  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
`);

const runAll = db.transaction((recs) => {
    let ok = 0;
    for (const r of recs) {
        ins.run(
            excelDateToStr(r[0]),
            String(r[1]).trim(),
            String(r[2]).trim(),
            String(r[3]).trim(),
            String(r[4]).replace(/\n/g, '').trim() || '大工',
            excelTimeToStr(r[5]),
            excelTimeToStr(r[6]),
            parseFloat(r[7]) || 0,
            String(r[8]).trim() || null,
            String(r[10]).trim() || null,
            String(r[11]).trim() || null
        );
        ok++;
    }
    return ok;
});

console.log('开始写入数据库...');
const count = runAll(validRows);
console.log(`✅ 导入完成！共写入 ${count} 条记录`);

// 汇总验证
console.log('\n--- 按外协单位汇总 ---');
db.prepare('SELECT supplier, COUNT(*) as n, ROUND(SUM(hours),1) as h FROM outsource_manpower GROUP BY supplier').all()
    .forEach(x => console.log(` ${x.supplier}: ${x.n} 条, 合计 ${x.h} 工时`));

console.log('\n--- 日期范围 ---');
const r = db.prepare('SELECT MIN(work_date) as mn, MAX(work_date) as mx FROM outsource_manpower').get();
console.log(` ${r.mn} ~ ${r.mx}`);

console.log('\n--- 涉及工单数 ---');
const wc = db.prepare('SELECT COUNT(DISTINCT work_order) as c FROM outsource_manpower').get();
console.log(` 共 ${wc.c} 个工单`);
