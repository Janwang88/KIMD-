const fs = require('fs');
const path = require('path');
const sqlite3 = require('better-sqlite3');

const csvPath = '/Users/wangjian/Documents/New project/插件管理/插件管理/源表/2026组装外借人力统计.csv';
const dbPath = '/Users/wangjian/Documents/New project/插件管理/插件管理/data/reviews.db';

function importData() {
    const db = new sqlite3(dbPath);

    // 确保表存在
    db.exec(`
        CREATE TABLE IF NOT EXISTS outsource_manpower (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            work_date TEXT,
            work_order TEXT,
            project_name TEXT,
            worker_name TEXT,
            worker_level TEXT,
            start_time TEXT,
            end_time TEXT,
            hours REAL,
            supplier TEXT,
            shift TEXT,
            manager TEXT,
            content TEXT,
            remark TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    `);

    // 清空旧数据以防止重复导入（根据用户需求，重新导入整个表）
    db.exec('DELETE FROM outsource_manpower');

    const fileContent = fs.readFileSync(csvPath, 'utf-8');

    const lines = fileContent.split(/\r?\n/);
    const records = lines.map(line => {
        return line.split(',').map(cell => cell.replace(/^"(.*)"$/, '$1').trim());
    });

    const insert = db.prepare(`
        INSERT INTO outsource_manpower (
            work_date, work_order, project_name, worker_name, worker_level, 
            start_time, end_time, hours, supplier, shift, manager, content, remark
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    const batchInsert = db.transaction((rows) => {
        for (const row of rows) insert.run(row);
    });

    let dataRows = [];
    let count = 0;

    for (let i = 0; i < records.length; i++) {
        const row = records[i];
        if (row.length < 5) continue;

        const dateStr = (row[0] || '').trim();
        const workOrder = (row[1] || '').trim();

        const isDate = /\d+\/\d+/.test(dateStr);
        const isWorkOrder = workOrder.startsWith('N.E') || workOrder.startsWith('R.E');

        if (isDate && isWorkOrder) {
            const val = (idx) => (row[idx] || '').toString().trim();
            const hours = parseFloat(val(7)) || 0;

            dataRows.push([
                val(0),  // work_date
                val(1),  // work_order
                val(2),  // project_name
                val(3),  // worker_name
                val(4),  // worker_level
                val(5),  // start_time
                val(6),  // end_time
                hours,   // hours
                val(8),  // supplier
                val(10), // shift
                val(11), // manager
                val(12), // content
                val(13)  // remark
            ]);
            count++;
        }
    }

    if (dataRows.length > 0) {
        batchInsert(dataRows);
        console.log(`成功导入 ${count} 条数据。`);
    } else {
        console.log('未找到有效数据行。');
    }
}

importData();
