const fs = require('fs');
const os = require('os');
const path = require('path');
const XLSX = require('xlsx');

function findLatestExcel(pattern = '物料') {
    const downloads = path.join(os.homedir(), 'Downloads');
    if (!fs.existsSync(downloads)) return null;
    try {
        // ★ 先判断 pattern 是否为空，避免 null/undefined.toLowerCase() 崩溃
        const hasPattern = pattern && (Array.isArray(pattern) ? pattern.length > 0 : true);
        const lowerPatterns = hasPattern
            ? (Array.isArray(pattern) ? pattern : [pattern]).map(p => (p || '').toLowerCase()).filter(Boolean)
            : [];

        const files = fs.readdirSync(downloads)
            .filter(name => {
                const lowerName = name.toLowerCase();
                return lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls');
            })
            .filter(name => !name.startsWith('~$'))
            .filter(name => {
                if (!hasPattern || lowerPatterns.length === 0) return true;
                const lowerName = name.toLowerCase();
                return lowerPatterns.some(p => lowerName.includes(p));
            })
            .map(name => {
                const full = path.join(downloads, name);
                try {
                    const stat = fs.statSync(full);
                    return { name, full, mtime: stat.mtimeMs };
                } catch (e) { return null; }
            })
            .filter(Boolean)
            .sort((a, b) => b.mtime - a.mtime);

        // --- 增强诊断日志 ---
        if (files.length === 0) {
            // 如果没找到匹配的 Excel，打印 Downloads 下最近的 5 个任意文件，帮用户查看到底下到哪了
            try {
                const anyFiles = fs.readdirSync(downloads)
                    .map(n => {
                        const f = path.join(downloads, n);
                        const s = fs.statSync(f);
                        return { n, m: s.mtimeMs };
                    })
                    .sort((a, b) => b.m - a.m)
                    .slice(0, 5);
                console.log(`[findLatestExcel] 未找到匹配Excel。Downloads 最近 5 个文件: ${anyFiles.map(f => `${f.n}(${new Date(f.m).toLocaleTimeString()})`).join(', ')}`);
            } catch (e) { }
        }

        return files[0] || null;
    } catch (e) {
        console.error('[findLatestExcel] 异常:', e.message);
        return null;
    }
}


async function waitForNewExcel(pattern, sinceMs, timeoutMs = 60000) {
    const start = Date.now();
    let lastCandidate = null;
    let loopCount = 0;
    while (Date.now() - start < timeoutMs) {
        loopCount++;
        const latest = findLatestExcel(pattern);

        // 每10次循环（约10秒）打印一次诊断日志
        if (loopCount % 10 === 1) {
            if (latest) {
                console.log(`[ExcelWait] 诊断 #${loopCount}: 最新文件="${latest.name}", 修改时间=${new Date(latest.mtime).toLocaleString()}, 基准时间=${new Date(sinceMs).toLocaleString()}, 差值=${((latest.mtime - sinceMs) / 1000).toFixed(1)}s`);
            } else {
                console.log(`[ExcelWait] 诊断 #${loopCount}: Downloads目录无匹配Excel文件 (pattern=${JSON.stringify(pattern)})`);
            }
        }

        // 增加 5 秒的容错缓冲 (sinceMs - 5000)，防止下载瞬间与请求发起的微小误差导致错过文件
        if (latest && latest.mtime > (sinceMs - 5000)) {
            console.log(`[ExcelWait] 检测到匹配文件: ${latest.name}, 修改时间: ${new Date(latest.mtime).toLocaleString()}, 比对基准: ${new Date(sinceMs).toLocaleString()}`);
            try {
                const currentSize = fs.statSync(latest.full).size;
                // 简单稳定性检测：同一个文件连续两次检测到且大小不变
                if (lastCandidate && lastCandidate.full === latest.full && lastCandidate.size === currentSize) {
                    console.log(`[ExcelWait] 文件读取稳定，开始解析: ${latest.name}`);
                    return latest;
                }
                lastCandidate = {
                    full: latest.full,
                    mtime: latest.mtime,
                    size: currentSize
                };
            } catch (statErr) {
                console.warn(`[ExcelWait] 读取文件状态失败: ${latest.name}`, statErr.message);
                lastCandidate = null;
            }
        }
        await new Promise(r => setTimeout(r, 1000));
    }
    console.log(`[ExcelWait] 超时退出，共等待 ${((Date.now() - start) / 1000).toFixed(0)}s`);
    return null;
}

function computeStatsFromExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    // 动态侦测核心列名
    const allHeaders = rows.length > 0 ? Object.keys(rows[0]) : [];
    const findField = (candidates) => candidates.find(c => allHeaders.includes(c));

    const partField = findField(['料号', '物料编码', '物料代码', '编码', 'PartNo', 'Part No', 'Material No']) || '料号';
    const procPrefix = '7.';
    const stdDeliveredFields = ['手工收料时间', '入库时间'];
    const procDeliveredFields = ['收料时间', '收料时间2', '手工收料时间', '入库时间'];
    const receiptFieldsForOnTime = ['收料时间', '收料时间2', '手工收料时间', '入库时间'];
    const iqcReceiptFields = ['收料时间', '收料时间2', '手工收料时间'];
    const inStockFields = ['入库时间'];
    const dueDateFieldCandidates = ['PMC需求时间', 'PMC需求日期'];

    const parseDate = (val) => {
        if (val === null || val === undefined || val === '') return null;
        if (val instanceof Date && !isNaN(val.getTime())) return val;
        if (typeof val === 'number') {
            const utcMs = Math.round((val - 25569) * 86400 * 1000);
            const offsetMs = new Date().getTimezoneOffset() * 60 * 1000;
            const d = new Date(utcMs + offsetMs);
            return isNaN(d.getTime()) ? null : d;
        }
        const s = String(val).trim();
        if (!s || s.toLowerCase() === 'nan' || s.toLowerCase() === 'null') return null;
        const normalized = s.replace(/\./g, '-').replace(/\//g, '-');
        const d = new Date(normalized);
        return isNaN(d.getTime()) ? null : d;
    };

    const startOfDay = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const hasAnyValue = (row, fields) => fields.some((f) => parseDate(row[f]));
    const pickEarliest = (dates) => {
        const list = dates.filter(Boolean);
        if (!list.length) return null;
        return list.reduce((min, cur) => (cur.getTime() < min.getTime() ? cur : min), list[0]);
    };

    const partAgg = new Map();
    let stdOnTimeOk = 0;
    let stdOnTimeNg = 0;
    let stdOnTimeChecked = 0;

    let procOnTimeOk = 0;
    let procOnTimeNg = 0;
    let procOnTimeChecked = 0;

    rows.forEach((r) => {
        const part = (r[partField] || '').toString().trim();
        if (!part) return;
        const isProc = part.startsWith(procPrefix);
        if (!partAgg.has(part)) {
            partAgg.set(part, {
                isProc: isProc,
                totalRows: 0,
                deliveredRows: 0,
                hasIqcReceipt: false,
                hasInStock: false,
                receiptEarliest: null,
                dueEarliest: null,
                orderDateEarliest: null,
                name: '',
                model: '',
                qty: 0,
                purchaseReplyDate: null
            });
        }
        const agg = partAgg.get(part);

        // Extract info
        if (!agg.name) {
            const nameFields = ['名称', '物料名称', '品名', 'Name'];
            const f = nameFields.find(k => r[k]);
            if (f) agg.name = String(r[f]).trim();
        }
        if (!agg.model) {
            const modelFields = ['规格型号', '规格', '型号', 'Model', 'Specification'];
            const f = modelFields.find(k => r[k]);
            if (f) agg.model = String(r[f]).trim();
        }

        // Accumulate Quantity
        const qtyFields = ['PMC下单数量', '数量', '计划数量', '订单数量', '需求数量', 'Qty', 'Quantity'];
        const qf = qtyFields.find(k => r[k] !== undefined && r[k] !== '');
        if (qf) {
            const val = parseFloat(r[qf]);
            if (!isNaN(val)) agg.qty += val;
        }

        agg.totalRows += 1;

        // Check Delivery (Row level)
        const isDelivered = agg.isProc ? hasAnyValue(r, procDeliveredFields) : hasAnyValue(r, stdDeliveredFields);
        if (isDelivered) {
            agg.deliveredRows += 1;
        }

        if (hasAnyValue(r, iqcReceiptFields)) agg.hasIqcReceipt = true;
        if (hasAnyValue(r, inStockFields)) agg.hasInStock = true;

        // 采购回复到料日期：取 Y/Z/AA 三列（采购回复到料日期1/2/3）的最大值
        const purchaseReplyFields = ['采购回复到料日期1', '采购回复到料日期2', '采购回复到料日期3'];
        const rowPurchaseReplyDate = purchaseReplyFields
            .map(f => parseDate(r[f]))
            .filter(d => d !== null)
            .reduce((max, d) => (!max || d > max ? d : max), null);
        if (rowPurchaseReplyDate) {
            agg.purchaseReplyDate = (!agg.purchaseReplyDate || rowPurchaseReplyDate > agg.purchaseReplyDate)
                ? rowPurchaseReplyDate
                : agg.purchaseReplyDate;
        }

        const rowReceiptEarliest = pickEarliest(receiptFieldsForOnTime.map((f) => parseDate(r[f])));
        if (rowReceiptEarliest) {
            agg.receiptEarliest = agg.receiptEarliest
                ? pickEarliest([agg.receiptEarliest, rowReceiptEarliest])
                : rowReceiptEarliest;
        }

        const orderDateFields = ['制单日期', '创建时间'];
        const rowOrderDate = pickEarliest(orderDateFields.map((f) => parseDate(r[f])));
        if (rowOrderDate) {
            agg.orderDateEarliest = agg.orderDateEarliest
                ? pickEarliest([agg.orderDateEarliest, rowOrderDate])
                : rowOrderDate;
        }

        const dueField = dueDateFieldCandidates.find((k) => Object.prototype.hasOwnProperty.call(r, k));
        const dueVal = dueField ? parseDate(r[dueField]) : null;
        if (dueVal) {
            agg.dueEarliest = agg.dueEarliest ? pickEarliest([agg.dueEarliest, dueVal]) : dueVal;
        }

        // On Time Calculation Logic
        let isOk = false;
        if (rowReceiptEarliest && dueVal) {
            const receiptDay = startOfDay(rowReceiptEarliest).getTime();
            const dueDay = startOfDay(dueVal).getTime();
            if (receiptDay <= dueDay) {
                isOk = true;
            }
        }

        if (isProc) {
            procOnTimeChecked += 1;
            if (isOk) procOnTimeOk += 1;
            else procOnTimeNg += 1;
        } else {
            stdOnTimeChecked += 1;
            if (isOk) stdOnTimeOk += 1;
            else stdOnTimeNg += 1;
        }
    });

    // 智能提取项目名称 (从前 15 行或特定字段中搜寻)
    let projectName = '';
    // 重新读取以获取完整原始行 (包含表头之前的行)
    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    for (let i = 0; i < Math.min(rawRows.length, 15); i++) {
        const rowValues = rawRows[i].map(v => String(v || '').trim());
        const pIdx = rowValues.findIndex(v => v.includes('项目名称') || v === '项目');
        if (pIdx !== -1 && rowValues[pIdx + 1]) {
            projectName = rowValues[pIdx + 1];
            break;
        }
        const selfMatch = rowValues.find(v => v.includes('项目名称:') || v.includes('项目名称：'));
        if (selfMatch) {
            const m = selfMatch.match(/项目名称[:：]\s*(.*)/);
            if (m && m[1]) {
                projectName = m[1].trim();
                break;
            }
        }
    }

    let stdCycleOk = 0;
    let stdCycleNg = 0;
    let procCycleOk = 0;
    let procCycleNg = 0;
    let stdUnRows = 0;
    let procUnRows = 0;
    let stdTotalDays = 0;
    let procTotalDays = 0;

    rows.forEach(r => {
        const part = String(r[partField] || '').trim();
        if (!part) return;
        const isProc = part.startsWith(procPrefix);

        // 使用多字段探测确保日期被读取
        const rowOrderDate = pickEarliest(['制单日期', '创建时间', '采购订单日期'].map(k => parseDate(r[k])));
        const rowReceiptEarliest = pickEarliest(['收料时间', '收料时间2', '手工收料时间', '入库时间', '实计到料时间'].map(k => parseDate(r[k])));

        if (!rowReceiptEarliest) {
            if (isProc) procUnRows++; else stdUnRows++;
        } else if (rowOrderDate) {
            const startDay = new Date(rowOrderDate);
            // 15:00 逻辑
            if (startDay.getHours() >= 15 && (startDay.getHours() > 15 || startDay.getMinutes() > 0 || startDay.getSeconds() > 0)) {
                startDay.setDate(startDay.getDate() + 1);
            }
            startDay.setHours(0, 0, 0, 0);

            const endDay = new Date(rowReceiptEarliest);
            endDay.setHours(0, 0, 0, 0);

            const days = Math.round((endDay - startDay) / 86400000);
            const actualDays = days >= 0 ? days : 0;

            if (isProc) {
                procTotalDays += actualDays;
                if (actualDays <= 7) procCycleOk++; else procCycleNg++;
            } else {
                stdTotalDays += actualDays;
                if (actualDays <= 10) stdCycleOk++; else stdCycleNg++;
            }
        } else {
            if (isProc) procUnRows++; else stdUnRows++;
        }
    });

    // 统计工单对应的订单总数 (G列：PMC下单数量)
    const woQtyMap = new Map();
    const woKeyCandidates = ['工单', '工单号', '工单编号', 'WorkOrder'];
    const qtyKeyCandidates = ['PMC下单数量', '数量', '计划数量', '订单数量', '需求数量', 'Qty'];

    rows.forEach(r => {
        const woKey = woKeyCandidates.find(k => r[k]);
        const qKey = qtyKeyCandidates.find(k => r[k] !== undefined && r[k] !== '');
        if (woKey && qKey) {
            const wo = String(r[woKey]).trim();
            const qty = parseFloat(r[qKey]) || 0;
            if (wo && !woQtyMap.has(wo)) {
                woQtyMap.set(wo, qty);
            }
        }
    });
    const totalOrderQty = Array.from(woQtyMap.values()).reduce((sum, q) => sum + q, 0);

    let stdCount = 0; // Items
    let procCount = 0; // Items
    let stdRows = 0; // Rows
    let procRows = 0; // Rows
    let stdDelivered = 0;
    let procDelivered = 0;
    let pendingIqc = 0;

    const undeliveredList = [];

    let firstMilestones = {
        assemblyStart: null,
        assemblyEnd: null,
        debugStart: null,
        debugEnd: null,
        shipStart: null
    };

    rows.forEach(r => {
        if (!firstMilestones.assemblyStart && (r['组装计划开始'] || r['AU'])) firstMilestones.assemblyStart = r['组装计划开始'] || r['AU'];
        if (!firstMilestones.assemblyEnd && (r['组装计划结束'] || r['AV'])) firstMilestones.assemblyEnd = r['组装计划结束'] || r['AV'];
        // Note: the prompt asks for '工程调试计划结束' twice for AW and AX.
        // Usually AW is start, AX is end. We try both '工程调试计划开始' and AW for debugStart, and 'AX' and '工程调试计划结束' for debugEnd
        if (!firstMilestones.debugStart && (r['工程调试计划开始'] || r['AW'] || r['工程调试计划时间'] || r['调试开始'])) firstMilestones.debugStart = r['工程调试计划开始'] || r['AW'] || r['调试开始'];
        if (!firstMilestones.debugEnd && (r['工程调试计划结束'] || r['AX'] || r['调试结束'])) firstMilestones.debugEnd = r['工程调试计划结束'] || r['AX'] || r['调试结束'];
        if (!firstMilestones.shipStart && (r['出货计划开始'] || r['BI'] || r['出货时间'])) firstMilestones.shipStart = r['出货计划开始'] || r['BI'] || r['出货时间'];

        // If headers are missing, we can try to get them by column index if using sheet_to_json({header:1}) 
        // but since {defval: ''} uses row 1 as headers, they should be accessible by header names.
    });

    // Formatting date helpers
    const formatDateObj = (val) => {
        const d = parseDate(val);
        if (!d) return '-';
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    };

    // Calculate item counts
    for (const [part, agg] of partAgg) {
        // Condition: Fully Delivered means all rows are delivered
        const isFullyDelivered = agg.deliveredRows >= agg.totalRows;

        if (agg.isProc) {
            procCount += 1;
            if (isFullyDelivered) {
                procDelivered += 1;
            } else {
                // Collect undelivered processed part info
                undeliveredList.push({
                    partNo: part,
                    name: agg.name || '',
                    model: agg.model || '',
                    qty: agg.qty || 0,
                    type: agg.isProc ? '加工件' : '标准件',
                    deliveredRows: agg.deliveredRows,
                    totalRows: agg.totalRows,
                    actualDays: agg.actualDays || null,
                    purchaseReplyDate: agg.purchaseReplyDate ? formatDateObj(agg.purchaseReplyDate) : null
                });
            }
        } else {
            stdCount += 1;
            if (isFullyDelivered) {
                stdDelivered += 1;
            } else {
                // Collect undelivered standard part info
                undeliveredList.push({
                    partNo: part,
                    name: agg.name || '',
                    model: agg.model || '',
                    qty: agg.qty || 0,
                    type: agg.isProc ? '加工件' : '标准件',
                    deliveredRows: agg.deliveredRows,
                    totalRows: agg.totalRows,
                    actualDays: agg.actualDays || null,
                    purchaseReplyDate: agg.purchaseReplyDate ? formatDateObj(agg.purchaseReplyDate) : null
                });
            }
        }

        if (agg.hasIqcReceipt && !agg.hasInStock) {
            pendingIqc += 1;
        }
    }

    // Calculate row counts (iterate through all rows again or can be done in first pass)
    rows.forEach((r) => {
        const part = (r[partField] || '').toString().trim();
        if (!part) return;
        if (part.startsWith(procPrefix)) {
            procRows += 1;
        } else {
            stdRows += 1;
        }
    });

    return {
        rows: rows.length,
        uniqueTotal: partAgg.size,
        projectName,
        stdTotal: stdCount,   // Item count
        procTotal: procCount, // Item count
        stdRows: stdRows,     // Row count
        procRows: procRows,   // Row count
        stdDelivered,
        procDelivered,
        stdUndelivered: stdCount - stdDelivered,
        procUndelivered: procCount - procDelivered,
        undeliveredList,      // Array of detailed info
        stdOnTimeOk,
        stdOnTimeNg,
        stdOnTimeChecked,
        stdOnTimeRate: stdOnTimeChecked > 0 ? Number(((stdOnTimeOk / stdOnTimeChecked) * 100).toFixed(2)) : null,
        procOnTimeOk,
        procOnTimeNg,
        procOnTimeChecked,
        procOnTimeRate: procOnTimeChecked > 0 ? Number(((procOnTimeOk / procOnTimeChecked) * 100).toFixed(2)) : null,
        pendingIqc,
        totalOrderQty: totalOrderQty, // 返回汇总的订单数量
        milestones: {
            assemblyStart: formatDateObj(firstMilestones.assemblyStart),
            assemblyEnd: formatDateObj(firstMilestones.assemblyEnd),
            debugStart: formatDateObj(firstMilestones.debugStart),
            debugEnd: formatDateObj(firstMilestones.debugEnd),
            shipStart: formatDateObj(firstMilestones.shipStart)
        },
        cycleStats: {
            stdOk: stdCycleOk,
            stdNg: stdCycleNg,
            stdUn: stdUnRows,
            stdAvg: (stdCycleOk + stdCycleNg) > 0 ? Number((stdTotalDays / (stdCycleOk + stdCycleNg)).toFixed(1)) : 0,
            procOk: procCycleOk,
            procNg: procCycleNg,
            procUn: procUnRows,
            procAvg: (procCycleOk + procCycleNg) > 0 ? Number((procTotalDays / (procCycleOk + procCycleNg)).toFixed(1)) : 0
        }
    };
}

// 从生产排程 Excel 中提取特定工单的时间节点
function extractMilestonesFromSchedule(schedulePath, workOrder) {
    if (!fs.existsSync(schedulePath)) return null;
    const workbook = XLSX.readFile(schedulePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!rows || rows.length < 2) return null;

    const header = rows[0].map(c => String(c).trim());
    const findIdx = (names) => names.map(n => header.indexOf(n)).find(i => i !== -1);

    const colMap = {
        workOrder: findIdx(['工单号', '生产任务单号', 'MO单']),
        taskNo: findIdx(['任务单', '计划单', '排程单', '生产任务单号']),
        assemblyStart: findIdx(['组装计划开始', '装配开始']),
        assemblyEnd: findIdx(['组装计划结束', '装配结束']),
        debugStart: findIdx(['工程调试计划开始', '调试开始']),
        debugEnd: findIdx(['工程调试计划结束', '调试结束', '入库计划开始']),
        shipStart: findIdx(['出货计划开始', '发货日期', '实际出货'])
    };

    // 格式化日期辅助函数
    const parseExcelDate = (val) => {
        if (!val) return null;
        if (typeof val === 'number') {
            const date_info = new Date(Math.round((val - 25569) * 86400 * 1000));
            return `${date_info.getUTCFullYear()}-${String(date_info.getUTCMonth() + 1).padStart(2, '0')}-${String(date_info.getUTCDate()).padStart(2, '0')}`;
        }
        const s = String(val).trim();
        const match = s.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
        if (match) return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
        return s;
    };

    const targetWo = String(workOrder).trim();
    // 查找匹配的行
    const matchedRow = rows.find((r, idx) => {
        if (idx === 0) return false;
        const woVal = colMap.workOrder !== undefined ? String(r[colMap.workOrder]).trim() : '';
        const taskVal = colMap.taskNo !== undefined ? String(r[colMap.taskNo]).trim() : '';
        return woVal === targetWo || taskVal === targetWo || (woVal && targetWo.includes(woVal)) || (taskVal && targetWo.includes(taskVal));
    });

    if (!matchedRow) return null;

    return {
        assemblyStart: parseExcelDate(matchedRow[colMap.assemblyStart]) || '-',
        assemblyEnd: parseExcelDate(matchedRow[colMap.assemblyEnd]) || '-',
        debugStart: parseExcelDate(matchedRow[colMap.debugStart]) || '-',
        debugEnd: parseExcelDate(matchedRow[colMap.debugEnd]) || '-',
        shipStart: parseExcelDate(matchedRow[colMap.shipStart]) || '-'
    };
}

module.exports = {
    findLatestExcel,
    waitForNewExcel,
    computeStatsFromExcel,
    extractMilestonesFromSchedule
};
