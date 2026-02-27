const helper = require('./utils/excelHelper');
const XLSX = require('xlsx');
const wb = XLSX.readFile('data/物料明细/N.E-005662601M070055_20260223.xlsx', { cellDates: true });
const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });

const parseDate = (val) => {
    if (!val) return null;
    if (val instanceof Date) {
        if (isNaN(val.getTime())) return null;
        return val;
    }
    if (typeof val === 'number') {
        const utcMs = Math.round((val - 25569) * 86400 * 1000);
        return new Date(utcMs);
    }
    const s = String(val).trim();
    if (!s || s.toLowerCase() === 'nan' || s.toLowerCase() === 'null') return null;
    const normalized = s.replace(/\./g, '-').replace(/\//g, '-');
    const d = new Date(normalized);
    return isNaN(d.getTime()) ? null : d;
};

const pickEarliest = (dates) => {
    const list = dates.filter(Boolean);
    if (!list.length) return null;
    return list.reduce((min, cur) => (cur.getTime() < min.getTime() ? cur : min), list[0]);
};

const procPrefix = '7.';
let map = new Map();

rows.forEach(r => {
    const p = String(r['料号'] || '').trim();
    if (!p) return;
    if (p.startsWith(procPrefix)) return; // std only

    if (!map.has(p)) {
        map.set(p, { o: null, r: null });
    }
    const item = map.get(p);

    const od = pickEarliest(['制单日期', '创建时间'].map(k => parseDate(r[k])));
    if (od) {
        item.o = item.o ? pickEarliest([item.o, od]) : od;
    }

    const rd = pickEarliest(['收料时间', '收料时间2', '手工收料时间', '入库时间'].map(k => parseDate(r[k])));
    if (rd) {
        item.r = item.r ? pickEarliest([item.r, rd]) : rd;
    }
});

let ng = [];
let ok = [];

for (let [p, v] of map) {
    if (v.o && v.r) {
        let sd = new Date(v.o);
        if (sd.getHours() >= 15 && (sd.getHours() > 15 || sd.getMinutes() > 0 || sd.getSeconds() > 0)) {
            sd.setDate(sd.getDate() + 1);
        }
        sd.setHours(0, 0, 0, 0);

        let ed = new Date(v.r);
        ed.setHours(0, 0, 0, 0);

        let diff = ed.getTime() - sd.getTime();
        let d = Math.ceil(diff / 86400000) + 1;
        let ad = d > 0 ? d : 1;

        if (ad > 10) {
            ng.push({ p, act: ad, o: v.o.toLocaleString(), r: v.r.toLocaleString() });
        } else {
            ok.push({ p, act: ad });
        }
    }
}

console.log('NG items:', ng.length);
ng.sort((a, b) => b.act - a.act).forEach(x => {
    console.log(`${x.p} | Days: ${x.act} | Order: ${x.o} | Receipt: ${x.r}`);
});
