const XLSX = require('xlsx');
const filePath = '/Users/wangjian/Documents/New project/插件管理/插件管理/data/生产排程/生产排程+2026-02-21.xlsx';
const workbook = XLSX.readFile(filePath);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

const keywords = ['工单数量', '订单数量', '数量', 'Qty', 'Quantity'];
const headerRow = rows[0];

const matches = [];
headerRow.forEach((cell, idx) => {
    if (cell && keywords.some(k => String(cell).includes(k))) {
        matches.push({ index: idx, name: cell, letter: String.fromCharCode(65 + idx % 26) + (idx >= 26 ? Math.floor(idx / 26) : '') });
    }
});

console.log('Matches for orderQty keywords:', JSON.stringify(matches, null, 2));

const targetWO = 'N.E-005662601M070055';
rows.forEach((row, idx) => {
    if (String(row[0]).trim() === targetWO) {
        console.log(`WO: ${targetWO}`);
        matches.forEach(m => {
            console.log(`Column ${m.index} (${m.name}): ${row[m.index]}`);
        });
    }
});
