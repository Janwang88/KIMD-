const XLSX = require('xlsx');

// Mock data: Column A is "MO-TEST-001", Column B is "Task-001", etc.
const data = [
    ["工单号", "任务单", "项目名称", "接单时间", "产品名称"], // Header row (simulated, though we force index 0)
    ["MO-TEST-001", "Task-001", "Test Project", "2026-02-17", "Test Product"]
];

// Logic extracted from server.js (modified version)
function testImport(rawRows) {
    let headerRowIndex = -1;
    let colMap = {
        workOrderNo: -1,
        taskNo: -1,
        projectName: -1,
        orderDate: -1,
        productName: -1
    };

    const keywords = {
        workOrderNo: ['工单号', 'MO单', '生产订单'],
        taskNo: ['任务单', '计划单', '排程单'],
        projectName: ['项目名称', '项目', 'Project'],
        orderDate: ['接单', '接单时间', '下单日期', '日期', 'Date', '制单日期'],
        productName: ['产品名称', '物料名称', '产品']
    };

    for (let i = 0; i < Math.min(rawRows.length, 10); i++) {
        const row = rawRows[i].map(c => String(c).trim());
        let matchCount = 0;
        const newMap = { ...colMap };

        row.forEach((cell, idx) => {
            if (!cell) return;
            // if (keywords.workOrderNo.some(k => cell.includes(k))) { newMap.workOrderNo = idx; matchCount++; } // 强制使用第一列
            if (keywords.taskNo.some(k => cell.includes(k))) { newMap.taskNo = idx; matchCount++; }
            else if (keywords.projectName.some(k => cell.includes(k))) { newMap.projectName = idx; matchCount++; }
            else if (keywords.orderDate.some(k => cell.includes(k))) { newMap.orderDate = idx; matchCount++; }
            else if (keywords.productName.some(k => cell.includes(k))) { newMap.productName = idx; matchCount++; }
        });

        // 强制指定 Column A (Index 0) 为工单号
        newMap.workOrderNo = 0;

        if (matchCount >= 2) {
            headerRowIndex = i;
            colMap = newMap;
            break;
        }
    }

    console.log("Header Found:", headerRowIndex !== -1);
    console.log("Column Map:", colMap);

    if (headerRowIndex === -1) return [];

    const result = [];
    for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
        const row = rawRows[i];
        const getVal = (idx) => (idx !== -1 && row[idx] !== undefined) ? row[idx] : '';

        const workOrderNo = String(getVal(colMap.workOrderNo)).trim();
        result.push({ workOrderNo });
    }
    return result;
}

const table = testImport(data);
console.log("Result:", table);

if (table.length > 0 && table[0].workOrderNo === "MO-TEST-001") {
    console.log("SUCCESS: Work Order No mapped correctly to Column A.");
} else {
    console.log("FAILURE: Work Order No NOT mapped correctly.");
}
