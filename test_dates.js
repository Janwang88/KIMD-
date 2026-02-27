const XLSX = require('xlsx');
const filePath = 'data/物料明细/N.E-005662601M070055_20260223.xlsx';
const workbook = XLSX.readFile(filePath); // NO cellDates:true
const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: '' });

const parseDate = (val) => {
    if (!val) return null;
    if (typeof val === 'number') {
        const utcMs = Math.round((val - 25569) * 86400 * 1000);
        const offset = new Date().getTimezoneOffset() * 60000;
        return new Date(utcMs + offset);
    }
    return new Date(String(val).replace(/\./g, '-').replace(/\//g, '-'));
};

const pickEarliest = (dates) => {
    const list = dates.filter(Boolean);
    if (!list.length) return null;
    return list.reduce((min, cur) => (cur.getTime() < min.getTime() ? cur : min), list[0]);
};

let map = new Map();
rows.forEach(r => {
    const p = String(r['料号'] || '').trim();
    if (!p) return;
    if (p.startsWith('7.')) return;

    if (!map.has(p)) map.set(p, { o: null, r: null });
    const item = map.get(p);

    const od = pickEarliest(['制单日期', '创建时间'].map(k => parseDate(r[k])));
    if (od) item.o = item.o ? pickEarliest([item.o, od]) : od;

    const rd = pickEarliest(['收料时间', '收料时间2', '手工收料时间', '入库时间'].map(k => parseDate(r[k])));
    if (rd) item.r = item.r ? pickEarliest([item.r, rd]) : rd;
});

let ok = 0, ng = 0;
for (let [p, v] of map) {
    if (v.o && v.r) {
        let sd = new Date(v.o);
        if (sd.getHours() >= 15 && (sd.getHours() > 15 || sd.getMinutes() > 0 || sd.getSeconds() > 0)) sd.setDate(sd.getDate() + 1);
        sd.setHours(0, 0, 0, 0);

        let ed = new Date(v.r);
        ed.setHours(0, 0, 0, 0);
        let d = Math.ceil((ed - sd) / 86400000) + 1;
        let ad = d > 0 ? d : 1;
        if (ad <= 10) ok++; else ng++;
    }
}
console.log('OK:', ok, 'NG:', ng);
