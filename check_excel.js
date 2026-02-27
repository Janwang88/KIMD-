const XLSX = require('xlsx');
const filePath = '/Users/wangjian/Documents/New project/插件管理/插件管理/data/生产排程/生产排程+2026-02-21.xlsx';
const workbook = XLSX.readFile(filePath);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

const headerRow = rows[0]; // Assuming headers are on the first row
console.log('Column index 6 (G) header:', headerRow[6]);
console.log('Column index 0 (A) header:', headerRow[0]);

const targetWOs = ['N.E-005662601M070055', 'N.E-005662601M070056', 'N.E-005662601M070057'];

rows.forEach((row, idx) => {
    const wo = String(row[0]).trim(); // First column is workOrderNo
    if (targetWOs.includes(wo)) {
        console.log(`WO: ${wo}, Column G (index 6): ${row[6]}, Full row length: ${row.length}`);
    }
});
