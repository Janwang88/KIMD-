const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const filePath = path.join(__dirname, 'Data/物料明细/N.E-005662601M070055_20260217.xlsx');

if (!fs.existsSync(filePath)) {
    console.error('File not found');
    process.exit(1);
}

const workbook = XLSX.readFile(filePath, { cellDates: true });
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

const partField = '料号';
const procPrefix = '7.';
const stdDeliveredFields = ['手工收料时间', '入库时间'];
const procDeliveredFields = ['收料时间', '收料时间2', '手工收料时间', '入库时间'];

const parseDate = (val) => {
    if (val === null || val === undefined || val === '') return null;
    if (val instanceof Date && !isNaN(val.getTime())) return val;
    if (typeof val === 'number') {
        const utcMs = Math.round((val - 25569) * 86400 * 1000);
        const d = new Date(utcMs);
        return isNaN(d.getTime()) ? null : d;
    }
    const s = String(val).trim();
    if (!s || s.toLowerCase() === 'nan' || s.toLowerCase() === 'null') return null;
    const normalized = s.replace(/\./g, '-').replace(/\//g, '-');
    const d = new Date(normalized);
    return isNaN(d.getTime()) ? null : d;
};

const hasAnyValue = (row, fields) => fields.some((f) => parseDate(row[f]));

const partAgg = new Map();
rows.forEach((r) => {
    const part = (r[partField] || '').toString().trim();
    if (!part) return;

    if (!partAgg.has(part)) {
        partAgg.set(part, {
            isProc: part.startsWith(procPrefix),
            totalRows: 0,
            deliveredRows: 0,
            rows: []
        });
    }
    const agg = partAgg.get(part);

    agg.totalRows += 1;
    const isDelivered = agg.isProc ? hasAnyValue(r, procDeliveredFields) : hasAnyValue(r, stdDeliveredFields);

    if (isDelivered) {
        agg.deliveredRows += 1;
    }

    agg.rows.push({
        isDelivered,
        recTime: r['收料时间'],
        inStock: r['入库时间']
    });
});

let procCount = 0;
let procDelivered = 0;
let stdCount = 0;
let stdDelivered = 0;
let undeliveredList = [];

for (const [part, agg] of partAgg) {
    const isFullyDelivered = agg.deliveredRows >= agg.totalRows;

    if (agg.isProc) {
        procCount++;
        if (isFullyDelivered) procDelivered++;
        else {
            undeliveredList.push({ part, type: '加工件', total: agg.totalRows, delivered: agg.deliveredRows });
        }
    } else {
        stdCount++;
        if (isFullyDelivered) stdDelivered++;
        else {
            undeliveredList.push({ part, type: '标准件', total: agg.totalRows, delivered: agg.deliveredRows });
        }
    }
}

console.log(`Proc: Total ${procCount}, Delivered ${procDelivered}, Undelivered ${procCount - procDelivered}`);
console.log(`Std:  Total ${stdCount}, Delivered ${stdDelivered}, Undelivered ${stdCount - stdDelivered}`);
console.log('--- Undelivered Details ---');
console.log(JSON.stringify(undeliveredList.slice(0, 10), null, 2));
