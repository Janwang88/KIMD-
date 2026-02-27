const express = require('express');
const fetch = require('node-fetch');
const https = require('https');
const path = require('path');
const fs = require('fs');
const os = require('os');
const XLSX = require('xlsx');

const app = express();
const PORT = 3000;

// 创建一个忽略SSL证书验证的agent（仅用于开发环境）
const httpsAgent = new https.Agent({
    rejectUnauthorized: false
});

const BASE_URL = 'https://chajian.kimd.cn:9999';
const BASE_IP_URL = 'https://153.35.130.19:9999';
const DATA_DIR = path.join(__dirname, 'data');
const CACHED_EXCEL = path.join(DATA_DIR, 'material_progress.xlsx');

if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
}

// 子目录定义
const SUB_DIRS = {
    schedule: path.join(DATA_DIR, '生产排程'),
    material: path.join(DATA_DIR, '物料明细'),
    hours: path.join(DATA_DIR, '工时统计')
};

// 创建子目录
Object.values(SUB_DIRS).forEach(dir => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// Helper: Move file to target directory with date/name
function moveFileToDir(sourcePath, targetDir, newName) {
    const destPath = path.join(targetDir, newName);
    if (fs.existsSync(destPath)) {
        try { fs.unlinkSync(destPath); } catch (e) { }
    }
    try {
        fs.renameSync(sourcePath, destPath);
        return destPath;
    } catch (e) {
        try {
            fs.copyFileSync(sourcePath, destPath);
            fs.unlinkSync(sourcePath);
            return destPath;
        } catch (e2) {
            console.error('Move file error:', e2);
            return null;
        }
    }
}

async function fetchWithTimeout(url, options, timeoutMs = 20000) {
    return await Promise.race([
        fetch(url, options),
        new Promise((_, reject) => {
            setTimeout(() => reject(new Error('TIMEOUT')), timeoutMs);
        })
    ]);
}

async function fetchWithRetry(url, options, retries = 2) {
    let lastErr = null;
    for (let i = 0; i <= retries; i++) {
        try {
            return await fetchWithTimeout(url, options);
        } catch (err) {
            lastErr = err;
            // 简单退避
            await new Promise(r => setTimeout(r, 500 * (i + 1)));
        }
    }
    throw lastErr;
}

async function fetchWithFallback(path, options) {
    const url = `${BASE_URL}${path}`;
    try {
        return await fetchWithRetry(url, options);
    } catch (err) {
        // DNS 解析失败或超时，尝试 IP 直连
        if (err && (err.code === 'ENOTFOUND' || err.message === 'TIMEOUT' || err.code === 'ECONNRESET')) {
            const headers = Object.assign({}, options?.headers || {});
            // 使用 IP 直连时，补充 Host 头保持与域名一致
            if (!headers['Host']) {
                headers['Host'] = 'chajian.kimd.cn:9999';
            }
            const ipUrl = `${BASE_IP_URL}${path}`;
            return await fetchWithRetry(ipUrl, { ...options, headers });
        }
        throw err;
    }
}

// 中间件
// 中间件
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));
app.use((req, res, next) => {
    // 允许 KIMD 页面直接调用本地接口
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') {
        return res.sendStatus(204);
    }
    next();
});
app.use(express.static(path.join(__dirname, 'public')));

// 存储token
let authToken = null;
let authCookie = null;

function findLatestExcel(pattern = '物料') {
    const downloads = path.join(os.homedir(), 'Downloads');
    if (!fs.existsSync(downloads)) return null;
    try {
        const patterns = Array.isArray(pattern) ? pattern : [pattern];
        const files = fs.readdirSync(downloads)
            .filter(name => name.toLowerCase().endsWith('.xlsx'))
            .filter(name => !name.startsWith('~$'))
            .filter(name => !name.includes(' (1)') && !name.includes(' - 副本'))
            .filter(name => {
                if (!pattern) return true;
                return patterns.some(p => name.includes(p));
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
        return files[0] || null;
    } catch (e) {
        return null;
    }
}

function computeStatsFromExcel(filePath) {
    const workbook = XLSX.readFile(filePath, { cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const partField = '料号';
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
            // Excel serial date (day 1 = 1900-01-01, with Excel leap bug)
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

    const startOfDay = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const hasAnyValue = (row, fields) => fields.some((f) => parseDate(row[f]));
    const pickEarliest = (dates) => {
        const list = dates.filter(Boolean);
        if (!list.length) return null;
        return list.reduce((min, cur) => (cur.getTime() < min.getTime() ? cur : min), list[0]);
    };

    let rowOnTimeOk = 0;
    let rowOnTimeNg = 0;
    let rowOnTimeChecked = 0;

    const partAgg = new Map();
    rows.forEach((r) => {
        const part = (r[partField] || '').toString().trim();
        if (!part) return;
        if (!partAgg.has(part)) {
            partAgg.set(part, {
                isProc: part.startsWith(procPrefix),
                totalRows: 0,
                deliveredRows: 0,
                hasIqcReceipt: false,
                hasInStock: false,
                receiptEarliest: null,
                dueEarliest: null,
                name: '',
                model: '',
                qty: 0
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

        const rowReceiptEarliest = pickEarliest(receiptFieldsForOnTime.map((f) => parseDate(r[f])));
        if (rowReceiptEarliest) {
            agg.receiptEarliest = agg.receiptEarliest
                ? pickEarliest([agg.receiptEarliest, rowReceiptEarliest])
                : rowReceiptEarliest;
        }

        const dueField = dueDateFieldCandidates.find((k) => Object.prototype.hasOwnProperty.call(r, k));
        const dueVal = dueField ? parseDate(r[dueField]) : null;
        if (dueVal) {
            agg.dueEarliest = agg.dueEarliest ? pickEarliest([agg.dueEarliest, dueVal]) : dueVal;
        }

        // On Time Calculation Logic
        rowOnTimeChecked += 1;
        let isOk = false;
        if (rowReceiptEarliest && dueVal) {
            const receiptDay = startOfDay(rowReceiptEarliest).getTime();
            const dueDay = startOfDay(dueVal).getTime();
            if (receiptDay <= dueDay) {
                isOk = true;
            }
        }
        if (isOk) {
            rowOnTimeOk += 1;
        } else {
            rowOnTimeNg += 1;
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
    let onTimeOk = rowOnTimeOk;
    let onTimeNg = rowOnTimeNg;
    let pendingIqc = 0;
    let onTimeChecked = rowOnTimeChecked;

    const undeliveredList = [];

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
                    type: '加工件',
                    deliveredRows: agg.deliveredRows,
                    totalRows: agg.totalRows
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
                    type: '标准件',
                    deliveredRows: agg.deliveredRows,
                    totalRows: agg.totalRows
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
        stdTotal: stdCount,   // Item count
        procTotal: procCount, // Item count
        stdRows: stdRows,     // Row count
        procRows: procRows,   // Row count
        stdDelivered,
        procDelivered,
        stdUndelivered: stdCount - stdDelivered,
        procUndelivered: procCount - procDelivered,
        undeliveredList,      // Array of detailed info
        onTimeOk,
        onTimeNg,
        onTimeChecked,
        pendingIqc,
        onTimeRate: onTimeChecked > 0 ? Number(((onTimeOk / onTimeChecked) * 100).toFixed(2)) : null,
        totalOrderQty: totalOrderQty // 返回汇总的订单数量
    };
}

async function waitForNewExcel(pattern, sinceMs, timeoutMs = 60000) {
    const start = Date.now();
    let lastCandidate = null;
    while (Date.now() - start < timeoutMs) {
        const latest = findLatestExcel(pattern);
        if (latest && latest.mtime > sinceMs) {
            // 简单稳定性检测：同一个文件连续两次检测到且大小不变
            if (lastCandidate && lastCandidate.full === latest.full && lastCandidate.size === fs.statSync(latest.full).size) {
                return latest;
            }
            lastCandidate = {
                full: latest.full,
                mtime: latest.mtime,
                size: fs.statSync(latest.full).size
            };
        }
        await new Promise(r => setTimeout(r, 1000));
    }
    return null;
}

// 手动设置Cookie（用于跳过登录）
app.post('/api/set-cookie', async (req, res) => {
    try {
        const { cookie, userInfo } = req.body || {};

        if (cookie && typeof cookie === 'string') {
            authCookie = cookie.includes('userInfo=') ? cookie : `userInfo=${cookie}`;
        } else if (userInfo && typeof userInfo === 'string') {
            authCookie = `userInfo=${userInfo}`;
        } else {
            return res.json({ success: false, message: '未提供cookie或userInfo' });
        }

        return res.json({ success: true });
    } catch (error) {
        return res.status(500).json({ success: false, message: error.message });
    }
});

// 登录接口
app.post('/api/login', async (req, res) => {
    try {
        const { userCode, password, company } = req.body;

        const attempts = [];
        const companyValue = (company || '').trim();
        const pwd = (password || '').trim();
        const pwdVariants = Array.from(new Set([
            pwd,
            // 兼容输入 kimd123456 / Kimd123456 的差异
            pwd ? (pwd.charAt(0).toUpperCase() + pwd.slice(1)) : '',
            pwd.toLowerCase()
        ].filter(Boolean)));

        // 兼容不同后端字段大小写 + 是否包含公司字段
        const pushAttempts = (userKey, pwdKey, companyKey) => {
            pwdVariants.forEach((pv) => {
                if (companyValue) {
                    attempts.push({
                        [userKey]: userCode,
                        [pwdKey]: pv,
                        [companyKey]: companyValue
                    });
                }
                attempts.push({
                    [userKey]: userCode,
                    [pwdKey]: pv
                });
            });
        };

        pushAttempts('UserCode', 'Password', 'Company');
        pushAttempts('userCode', 'password', 'company');

        const payloadVariants = [];
        attempts.forEach((a) => {
            // 1) JSON body
            payloadVariants.push({
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(a)
            });
            // 2) form-urlencoded body（部分旧接口只认表单）
            const form = new URLSearchParams();
            Object.entries(a).forEach(([k, v]) => form.append(k, String(v ?? '')));
            payloadVariants.push({
                headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
                body: form.toString()
            });
        });

        let lastData = null;

        for (const pv of payloadVariants) {
            const response = await fetchWithFallback('/SYSLogin', {
                method: 'POST',
                headers: pv.headers,
                body: pv.body,
                agent: httpsAgent
            });

            const text = await response.text();
            let data = null;
            try {
                data = JSON.parse(text);
            } catch (e) {
                lastData = { message: text?.slice?.(0, 300) || '非JSON响应' };
                continue;
            }
            lastData = data;

            if (data.error === '000000' && data.data) {
                authToken = data.data.Token;
                authCookie = `userInfo=${encodeURIComponent(JSON.stringify({
                    UserCode: data.data.UserCode,
                    UserName: data.data.UserName,
                    LoginType: '0',
                    Token: data.data.Token,
                    HeadPortraitUrl: ''
                }))}`;

                return res.json({
                    success: true,
                    data: {
                        userName: data.data.UserName,
                        token: data.data.Token
                    }
                });
            }
        }

        res.json({
            success: false,
            message: (lastData && (lastData.msg || lastData.message)) || '登录失败',
            raw: lastData
        });
    } catch (error) {
        console.error('Login error:', error);
        res.status(500).json({
            success: false,
            message: '服务器错误: ' + error.message
        });
    }
});

// 从导出的Excel计算统计
app.post('/api/excel-material-stats', async (req, res) => {
    try {
        const { pattern } = req.body || {};
        const latest = findLatestExcel(pattern || '物料');
        if (!latest) {
            return res.json({ success: false, message: '未找到最新的Excel文件（请先在KIMD导出）' });
        }
        fs.copyFileSync(latest.full, CACHED_EXCEL);
        const stats = computeStatsFromExcel(CACHED_EXCEL);
        return res.json({ success: true, file: latest.name, filePath: latest.full, stats });
    } catch (err) {
        return res.status(500).json({ success: false, message: err.message });
    }
});

// 清空Excel缓存
app.post('/api/excel-clear-cache', async (_req, res) => {
    try {
        if (fs.existsSync(CACHED_EXCEL)) fs.unlinkSync(CACHED_EXCEL);
        return res.json({ success: true });
    } catch (err) {
        return res.status(500).json({ success: false, message: err.message });
    }
});

// 关注工单日志
app.post('/api/watchlist-log', async (req, res) => {
    try {
        const { action, workOrder, project } = req.body || {};
        const ts = new Date().toISOString();
        const line = JSON.stringify({
            ts,
            action: action || 'unknown',
            workOrder: workOrder || '',
            project: project || ''
        });
        const logPath = path.join(DATA_DIR, 'watchlist.log');
        fs.appendFileSync(logPath, line + '\n', 'utf8');
        return res.json({ success: true });
    } catch (err) {
        return res.status(500).json({ success: false, message: err.message });
    }
});

// 等待新导出的Excel并统计
// 等待新导出的Excel并统计
// 等待新导出的Excel并统计 (物料明细)
app.post('/api/excel-wait-stats', async (req, res) => {
    try {
        const { pattern, since, timeoutMs, workOrder } = req.body || {};
        const sinceMs = typeof since === 'number' ? since : Date.now();

        // 智能构造监控模式：如果提供了工单号，将其加入匹配列表，以防止默认文件名不包含“物料”
        let searchPattern = pattern || ['物料', 'Material', 'Export', '进度', '清单'];
        if (workOrder) {
            if (!Array.isArray(searchPattern)) searchPattern = [searchPattern];
            const woPart = workOrder.split(/[, \n]/)[0].trim(); // 取第一个工单号作为核心特征
            if (woPart && !searchPattern.includes(woPart)) {
                searchPattern.push(woPart);
            }
        }

        const latest = await waitForNewExcel(searchPattern, sinceMs, timeoutMs || 120000);
        if (!latest) {
            return res.json({ success: false, message: '等待超时，未检测到新导出的Excel' });
        }

        let targetPath = latest.full;
        let savedAs = null;

        if (workOrder) {
            const safeName = workOrder.replace(/[\\/:*?"<>|]/g, '_');
            const now = new Date();
            const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
            const newFileName = `${safeName}_${dateStr}.xlsx`;

            // Move to '物料明细' subfolder
            const newPath = moveFileToDir(latest.full, SUB_DIRS.material, newFileName);

            if (newPath) {
                targetPath = newPath;
                latest.name = newFileName;
                savedAs = newFileName;
            }
        }

        const stats = computeStatsFromExcel(targetPath);

        return res.json({
            success: true,
            file: path.basename(targetPath),
            filePath: targetPath,
            savedAs: savedAs,
            stats
        });
    } catch (err) {
        return res.status(500).json({ success: false, message: err.message });
    }
});

// 实际工时统计逻辑
function computeHoursStatsFromExcel(filePath) {
    if (!fs.existsSync(filePath)) return null;

    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    // 使用 header:1 获取二维数组，方便处理
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    if (!rows || rows.length < 2) return { totalHours: 0, details: [] };

    // 查找表头行 (通常在第一行，但也可能在第二行，或者有合并单元格)
    // 关键字：'工序' or 'Process', '实际工时' or 'Actual Hours'
    let headerRowIndex = -1;
    let colMap = {};

    for (let i = 0; i < Math.min(rows.length, 10); i++) {
        const row = rows[i];
        const strRow = row.map(c => (c || '').toString().trim());
        if (strRow.includes('工艺名称') || strRow.includes('工序名称')) {
            headerRowIndex = i;
            // Map columns
            strRow.forEach((col, idx) => {
                if (col.includes('工艺名称') || col.includes('工序名称')) colMap.process = idx;
                else if (col.includes('组别名称')) colMap.group = idx;
                else if (col.includes('工位名称')) colMap.station = idx;
                // Metrics
                else if (col === '生产计划用时') colMap.planHours = idx;
                else if (col === 'KIMD工时' || col.includes('KIMD')) colMap.kimdHours = idx;
                else if (col === '外协工时') colMap.outsourceHours = idx;
                else if (col === '实际工时') colMap.actualHours = idx;
                else if (col.includes('生产工时') && colMap.actualHours === undefined) colMap.actualHours = idx;
            });
            console.log(`[HoursStats] Found header at row ${i}:`, colMap);
            break;
        }
    }

    if (headerRowIndex === -1) {
        console.warn('[HoursStats] No header found with "工艺名称" or "工序名称"');
        return { error: '无法识别表头，请确认Excel包含“工艺名称”或“工序名称”列' };
    }

    // 如果没有找到实际工时列，尝试用 plan/kimd 代替检查，或者允许宽容模式
    if (colMap.actualHours === undefined && colMap.planHours === undefined && colMap.kimdHours === undefined) {
        console.warn('[HoursStats] No hours columns found (actual/plan/kimd)');
        return { error: '无法识别工时列，请确认Excel包含“实际工时”、“生产计划用时”或“KIMD工时”' };
    }

    let stats = {
        totalHours: 0,
        // Assembly
        assembly: { plan: 0, kimd: 0, outsource: 0, total: 0, processes: [] },
        // Wiring
        wiring: { plan: 0, kimd: 0, outsource: 0, total: 0, processes: [] },
        // Mixed
        mixed: { plan: 0, kimd: 0, outsource: 0, total: 0, processes: [] },
        details: []
    };

    const RULES = {
        mixed: new Set(['项目管理', '领料', '上线准备', '总装', '清洁', '打包']),
        assembly: new Set(['组装-返工', '模组组装', '整机接气', '出货']),
        wiring: new Set(['接线-返工', '电控配线', '整机接线', '通电通气'])
    };

    // 遍历数据行
    for (let i = headerRowIndex + 1; i < rows.length; i++) {
        const row = rows[i];
        const process = (row[colMap.process] || '').toString().trim();

        // Skip header repetition or empty
        if (!process || process === '工艺名称' || process === '合计' || process === '总计') continue;

        // Categorize (Exact Match)
        let target = null;
        if (RULES.mixed.has(process)) target = stats.mixed;
        else if (RULES.assembly.has(process)) target = stats.assembly;
        else if (RULES.wiring.has(process)) target = stats.wiring;

        if (target) {
            // Extract Values
            const valPlan = row[colMap.planHours];
            const valKimd = row[colMap.kimdHours];
            const valOut = row[colMap.outsourceHours];
            const valProd = row[colMap.productionHours];

            const plan = parseFloat(valPlan || 0) || 0;
            const kimd = parseFloat(valKimd || 0) || 0;
            const outsource = parseFloat(valOut || 0) || 0;
            const production = parseFloat(valProd || 0) || 0;

            // Logic Update: Total Actual = Production (M) + KIMD (Mgmt) + Outsource
            const actual = production + kimd + outsource;

            console.log(`[HoursDebug] Row ${i} matched '${process}': Prod=${production}(${valProd}) KIMD=${kimd}(${valKimd}) Out=${outsource}(${valOut}) -> Actual=${actual}`);

            target.plan += plan;
            target.kimd += kimd;
            target.outsource += outsource;
            target.total += actual;
            if (!target.processes.includes(process)) target.processes.push(process);
            stats.totalHours += actual;
        } else {
            console.log(`[HoursDebug] Row ${i} skipped '${process}' - Not in RULES`);
        }
    }

    return stats;
}



// API: Wait for Actual Hours Excel
app.post('/api/hours-wait-stats', async (req, res) => {
    try {
        const { pattern, timeoutMs, workOrder } = req.body || {};
        const sinceMs = Date.now() - (1000 * 60 * 5);
        const searchPattern = pattern || ['工时', 'Actual', 'Export'];

        const latest = await waitForNewExcel(searchPattern, sinceMs, timeoutMs || 120000);

        if (!latest) {
            return res.json({ success: false, message: '等待超时，未检测到导出的工时Excel' });
        }

        let targetPath = latest.full;
        let savedAs = null;

        if (workOrder) {
            const safeName = workOrder.replace(/[\\/:*?"<>|]/g, '_');
            const now = new Date();
            const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
            const newFileName = `${safeName}_工时明细_${dateStr}.xlsx`;

            // Move to '工时统计' subfolder
            const newPath = moveFileToDir(latest.full, SUB_DIRS.hours, newFileName);

            if (newPath) {
                targetPath = newPath;
                latest.name = newFileName;
                savedAs = newFileName;
            }
        }

        const stats = computeHoursStatsFromExcel(targetPath);

        if (stats && stats.error) {
            return res.json({ success: false, message: stats.error, file: path.basename(targetPath) });
        }

        return res.json({
            success: true,
            file: path.basename(targetPath),
            filePath: targetPath,
            savedAs: savedAs,
            stats
        });

    } catch (err) {
        return res.status(500).json({ success: false, message: err.message });
    }
});

// 获取生产投料单明细
app.post('/api/ppbom-entry', async (req, res) => {
    try {
        const { workOrderNo, pageIndex = 1, pageSize = 100 } = req.body;

        if (!authToken && !authCookie) {
            return res.status(401).json({
                success: false,
                message: '请先登录'
            });
        }

        // 构建过滤条件
        const filters = [];

        if (workOrderNo) {
            filters.push({
                field: 'FProjectNumber',
                op: 'contains',
                value: workOrderNo
            });
        }

        // 默认过滤条件：计划数量大于0，状态不等于结案
        filters.push({
            field: 'FQty',
            op: 'gt',
            value: '0'
        });
        filters.push({
            field: 'FICMOStatus',
            op: 'neq',
            value: '结案'
        });

        const requestBody = {
            tableId: 'ppbomEntry',
            pageIndex: pageIndex,
            pageSize: pageSize,
            filters: filters,
            sorts: []
        };

        const response = await fetchWithFallback('/Action/WMS.SYS.Table.GetTableContent', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Cookie': authCookie
            },
            body: JSON.stringify(requestBody),
            agent: httpsAgent
        });

        const data = await response.json();

        if (data.error === '000000') {
            res.json({
                success: true,
                data: data.data,
                total: data.total || data.data.length
            });
        } else {
            res.json({
                success: false,
                message: data.msg || '查询失败',
                rawError: data.error
            });
        }
    } catch (error) {
        console.error('Query error:', error);
        res.status(500).json({
            success: false,
            message: '服务器错误: ' + error.message
        });
    }
});

// 获取生产排程数据
app.post('/api/production-schedule', async (req, res) => {
    try {
        const { workOrderNo, pageIndex = 1, pageSize = 50 } = req.body;

        if (!authToken && !authCookie) {
            return res.status(401).json({
                success: false,
                message: '请先登录'
            });
        }

        // 生产排程页面实际请求是表单提交 + tableCode
        // 这里尽量对齐真实请求格式（发送一个被URL编码的JSON字符串）
        const filterDic = {
            page: Number(pageIndex) || 1,
            pageSize: Number(pageSize) || 50,
            orderKey: 'SCPriority ASC, BillDate ASC',
            filter: '',
            detail1FilterStr: '',
            detail2FilterStr: '',
            detail3FilterStr: '',
            detail4FilterStr: '',
            detail5FilterStr: ''
        };

        // 如果需要按工单号过滤，尝试拼接过滤条件（字段名可能需要再确认）
        if (workOrderNo) {
            filterDic.filter = `AND FProjectNumber like '%${workOrderNo}%'`;
        }

        const payload = {
            filterDic,
            tableCode: 'SC_ProductionScheduling'
        };

        async function postBody(body) {
            const resp = await fetchWithFallback('/Action/WMS.SYS.Table.GetTableContent', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                    'Cookie': authCookie,
                    'Origin': BASE_URL,
                    'Referer': `${BASE_URL}/`
                },
                body,
                agent: httpsAgent
            });
            const text = await resp.text();
            try {
                return JSON.parse(text);
            } catch (e) {
                return { error: 'NON_JSON', message: '响应不是JSON', rawText: text.slice(0, 500) };
            }
        }

        // 先尝试发送“原始JSON字符串”（不编码）
        let data = await postBody(JSON.stringify(payload));
        // 如果返回提示 JSON 基元错误，再尝试 URL 编码后的字符串
        if (data && (data.msg || data.message) && /JSON/.test(data.msg || data.message)) {
            data = await postBody(encodeURIComponent(JSON.stringify(payload)));
        }

        if (data.error === '000000') {
            res.json({
                success: true,
                data: data.data,
                total: data.total || data.data.length
            });
        } else {
            res.json({
                success: false,
                message: data.msg || '查询失败',
                rawError: data.error
            });
        }
    } catch (error) {
        console.error('Query error:', error);
        res.status(500).json({
            success: false,
            message: '服务器错误: ' + error.message
        });
    }
});

// 获取物料执行情况明细（全量）
app.post('/api/material-progress', async (req, res) => {
    try {
        const { workOrder, pageSize = 200 } = req.body || {};

        if (!authToken && !authCookie) {
            return res.status(401).json({
                success: false,
                message: '请先登录'
            });
        }

        if (!workOrder) {
            return res.status(400).json({
                success: false,
                message: '缺少工单号'
            });
        }

        const tableCode = 'V_WMS_MaterialProgres';
        const filterRaw = ` AND (1=1 OR 工单 LIKE '%${workOrder}%')`;

        async function postBody(body) {
            const resp = await fetchWithFallback('/Action/WMS.SYS.Table.GetTableContent', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                    'Cookie': authCookie,
                    'Origin': BASE_URL,
                    'Referer': `${BASE_URL}/`
                },
                body,
                agent: httpsAgent
            });
            return resp.json();
        }

        const rows = [];
        let page = 1;
        let safety = 0;
        let total = null;

        while (safety < 20) {
            safety += 1;
            const filterDic = {
                pageSize: Number(pageSize) || 200,
                page: page,
                orderKey: '制单日期 DESC',
                // 实际请求中 filter 字段是 URL 编码后的字符串
                filter: encodeURIComponent(filterRaw),
                detail1FilterStr: '',
                detail2FilterStr: '',
                detail3FilterStr: '',
                detail4FilterStr: '',
                detail5FilterStr: ''
            };

            let data = await postBody(JSON.stringify({ filterDic, tableCode }));
            if (data && (data.msg || data.message) && /JSON/.test(data.msg || data.message)) {
                data = await postBody(encodeURIComponent(JSON.stringify({ filterDic, tableCode })));
            }

            if (!data || data.error !== '000000') {
                return res.json({
                    success: false,
                    message: data?.msg || data?.message || '查询失败',
                    raw: data || null
                });
            }

            const list = Array.isArray(data.data) ? data.data : [];
            if (typeof data.total === 'number') {
                total = data.total;
            }
            rows.push(...list);

            if (list.length < filterDic.pageSize) break;
            if (total !== null && rows.length >= total) break;
            page += 1;
        }

        return res.json({
            success: true,
            data: rows
        });
    } catch (error) {
        console.error('material-progress error:', error);
        res.status(500).json({
            success: false,
            message: '服务器错误: ' + error.message
        });
    }
});

// 代理API请求
app.post('/api/proxy', async (req, res) => {
    try {
        const { endpoint, body } = req.body;

        const response = await fetchWithFallback(endpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Cookie': authCookie || ''
            },
            body: JSON.stringify(body),
            agent: httpsAgent
        });

        const data = await response.json();
        res.json(data);
    } catch (error) {
        console.error('Proxy error:', error);
        res.status(500).json({
            success: false,
            message: '代理请求失败: ' + error.message
        });
    }
});

// 导入工单数据（来自 Excel 上传）
app.post('/api/import-work-orders', async (req, res) => {
    try {
        const { fileContent } = req.body; // Base64 string
        if (!fileContent) {
            return res.status(400).json({ success: false, message: '未接收到文件内容' });
        }

        // 解析 Excel
        const workbook = XLSX.read(fileContent, { type: 'base64' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        // 转换为二维数组，便于处理表头
        const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        if (!rawRows || rawRows.length === 0) {
            return res.status(400).json({ success: false, message: 'Excel 文件为空' });
        }

        // 1. 定位表头行（寻找包含关键字的行）
        let headerRowIndex = -1;
        let colMap = {
            workOrderNo: -1,
            taskNo: -1,
            projectName: -1,
            orderDate: -1,
            productName: -1,
            orderQty: -1
        };

        const keywords = {
            workOrderNo: ['工单号', 'MO单', '生产订单'],
            taskNo: ['任务单', '计划单', '排程单'],
            projectName: ['项目名称', '项目', 'Project'],
            orderDate: ['接单时间', '下单日期', '日期', 'Date', '制单日期'],
            productName: ['产品名称', '物料名称', '产品'],
            orderQty: ['工单数量', '订单数量', '数量', 'Qty', 'Quantity']
        };

        // 扫描前10行寻找表头
        for (let i = 0; i < Math.min(rawRows.length, 10); i++) {
            const row = rawRows[i].map(c => String(c).trim());

            // 检查这一行是否包含至少两个我们关心的关键词
            let matchCount = 0;
            const newMap = { ...colMap };

            row.forEach((cell, idx) => {
                if (!cell) return;
                // if (keywords.workOrderNo.some(k => cell.includes(k))) { newMap.workOrderNo = idx; matchCount++; } // 强制使用第一列
                if (keywords.taskNo.some(k => cell.includes(k))) { newMap.taskNo = idx; matchCount++; }
                else if (keywords.projectName.some(k => cell.includes(k))) { newMap.projectName = idx; matchCount++; }
                else if (keywords.orderDate.some(k => cell.includes(k))) { newMap.orderDate = idx; matchCount++; }
                else if (keywords.productName.some(k => cell.includes(k))) { newMap.productName = idx; matchCount++; }
                else if (keywords.orderQty.some(k => cell.includes(k))) {
                    // 优先选择“工单数量”或“订单数量”这种精准匹配，且不被后面出现的“领料数量”等覆盖
                    if (newMap.orderQty === -1 || cell === '工单数量' || cell === '订单数量') {
                        newMap.orderQty = idx;
                        matchCount++;
                    }
                }
            });

            // 强制指定 Column A (Index 0) 为工单号
            newMap.workOrderNo = 0;

            if (matchCount >= 2) {
                headerRowIndex = i;
                colMap = newMap;
                break;
            }
        }

        if (headerRowIndex === -1) {
            // 尝试使用第一行作为默认表头
            headerRowIndex = 0;
            // 简单的按位置猜一下？不，还是让前端提示失败比较好。
            // 但为了健壮性，我们可以尝试再次扫描完整匹配
            // 这里简化处理：如果没有找到表头，返回错误
            return res.status(400).json({ success: false, message: '无法识别表头，请确保 Excel 包含“工单号”、“接单时间”等列' });
        }

        console.log(`[Import] 找到表头在第 ${headerRowIndex + 1} 行:`, colMap);

        // 2. 提取数据
        const data = [];
        const seen = new Set();

        // Excel 日期处理工具
        const parseExcelDate = (val) => {
            if (!val) return '';
            if (val instanceof Date) {
                return val.toISOString().split('T')[0];
            }
            if (typeof val === 'number') {
                // Excel serial date -> JS Date
                const utc_days = Math.floor(val - 25569);
                const utc_value = utc_days * 86400;
                const date_info = new Date(utc_value * 1000);
                // 处理时区偏差
                const date = new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate());
                return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
            }
            // 字符串处理
            const str = String(val).trim();
            // 2025/1/1 or 2025-1-1
            const match = str.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
            if (match) {
                return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
            }
            return str; // 原样返回
        };

        for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
            const row = rawRows[i];
            // 获取各列数据
            const getVal = (idx) => (idx !== -1 && row[idx] !== undefined) ? row[idx] : '';

            const rawWorkOrder = getVal(colMap.workOrderNo);
            const rawTaskNo = getVal(colMap.taskNo);

            // 必须要有工单号
            if (!rawWorkOrder && !rawTaskNo) continue;

            const workOrderNo = String(rawWorkOrder || rawTaskNo).trim(); // 优先用工单列，没有则用任务单
            const taskNo = String(rawTaskNo || rawWorkOrder).trim();      // 任务单号，如果没有则用工单号

            // 去重
            if (seen.has(taskNo)) continue;
            seen.add(taskNo);

            const projectName = String(getVal(colMap.projectName)).trim();
            const rawDate = getVal(colMap.orderDate);
            const orderDate = parseExcelDate(rawDate);
            const orderQty = parseFloat(getVal(colMap.orderQty)) || 0;

            data.push({
                workOrderNo,
                taskNo,
                projectName,
                orderDate,
                orderQty
            });
        }

        const workOrdersFile = path.join(SUB_DIRS.schedule, 'work_orders.json');
        const storageData = {
            data: data,
            updateTime: new Date().toISOString(),
            source: 'excel_import'
        };

        fs.writeFileSync(workOrdersFile, JSON.stringify(storageData, null, 2), 'utf8');

        console.log(`[Import] 成功导入 ${data.length} 条数据`);
        return res.json({ success: true, count: data.length });

    } catch (error) {
        console.error('Import error:', error);
        res.status(500).json({ success: false, message: '导入失败: ' + error.message });
    }
});

// 监听生产排程 Excel 导出并自动导入
app.post('/api/watch-schedule-export', async (req, res) => {
    try {
        const { since, timeoutMs } = req.body || {};
        const sinceMs = typeof since === 'number' ? since : Date.now();

        // 监听包含 '生产排程' 或 'Production' 的文件
        const latest = await waitForNewExcel('排程', sinceMs, timeoutMs || 120000);

        if (!latest) {
            return res.json({ success: false, message: '等待超时，未检测到新导出的 Excel' });
        }

        // -------------------------------------------------
        // 新增逻辑: 重命名并移动文件到 Data 目录
        // -------------------------------------------------
        const now = new Date();
        const dateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
        const newFilename = `生产排程+${dateStr}.xlsx`;

        // Move to '生产排程' subfolder
        const destPath = moveFileToDir(latest.full, SUB_DIRS.schedule, newFilename);

        if (!destPath) {
            console.error('Failed to move schedule file');
        }

        const fileToRead = destPath || latest.full;

        // 读取归档后的文件内容
        const fileBuffer = fs.readFileSync(fileToRead); // 读新文件

        // 解析 Excel
        const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        if (!rawRows || rawRows.length === 0) {
            return res.json({ success: false, message: 'Excel 文件为空' });
        }

        // 1. 定位表头行
        let headerRowIndex = -1;
        let colMap = {
            workOrderNo: -1,
            taskNo: -1,
            projectName: -1,
            orderDate: -1,
            productName: -1,
            orderQty: -1
        };

        const keywords = {
            workOrderNo: ['工单号', 'MO单', '生产订单'],
            taskNo: ['任务单', '计划单', '排程单'],
            projectName: ['项目名称', '项目', 'Project'],
            // 增加 "接单" 匹配
            orderDate: ['接单', '接单时间', '下单日期', '日期', 'Date', '制单日期'],
            productName: ['产品名称', '物料名称', '产品'],
            orderQty: ['工单数量', '订单数量', '数量', 'Qty', 'Quantity']
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
                else if (keywords.orderQty.some(k => cell.includes(k))) {
                    // 优先选择“工单数量”或“订单数量”，防止被后续“领料数量/应发数量”覆盖
                    if (newMap.orderQty === -1 || cell === '工单数量' || cell === '订单数量') {
                        newMap.orderQty = idx;
                        matchCount++;
                    }
                }
            });

            // 强制指定 Column A (Index 0) 为工单号
            newMap.workOrderNo = 0;

            if (matchCount >= 2) {
                headerRowIndex = i;
                colMap = newMap;
                break;
            }
        }

        if (headerRowIndex === -1) {
            headerRowIndex = 0; // 兜底
        }

        // 强力修正：用户反馈“接单时间”在 L 列 (Index 11)
        // 如果自动识别没找到，或者找到的位置不对，强制检查 Index 11
        if (colMap.orderDate === -1) {
            // 检查 Index 11 是否有值
            const headerRow = rawRows[headerRowIndex] || [];
            if (headerRow[11]) {
                console.log(`[AutoImport] 自动识别未找到接单时间，强制使用 L 列 (Index 11): ${headerRow[11]}`);
                colMap.orderDate = 11;
            }
        } else if (colMap.orderDate !== 11) {
            // 如果自动识别到了其他列，但 Index 11 看起来更像“接单”？
            const headerRow = rawRows[headerRowIndex] || [];
            if (headerRow[11] && (String(headerRow[11]).includes('接单') || String(headerRow[11]).includes('日期'))) {
                console.log(`[AutoImport] 自动识别列 (${colMap.orderDate}) 可能不准确，优先使用 L 列 (Index 11): ${headerRow[11]}`);
                colMap.orderDate = 11;
            }
        }

        console.log(`[AutoImport] 表头行: ${headerRowIndex}, 列映射:`, colMap);

        // 2. 提取数据
        const data = [];
        const seen = new Set();

        const parseExcelDate = (val) => {
            if (!val) return '';
            if (val instanceof Date) return val.toISOString().split('T')[0];
            if (typeof val === 'number') {
                const date_info = new Date(Math.round((val - 25569) * 86400 * 1000));
                return `${date_info.getFullYear()}-${String(date_info.getMonth() + 1).padStart(2, '0')}-${String(date_info.getDate()).padStart(2, '0')}`;
            }
            const str = String(val).trim();
            const match = str.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
            if (match) return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
            return str;
        };

        for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
            const row = rawRows[i];
            const getVal = (idx) => (idx !== -1 && row[idx] !== undefined) ? row[idx] : '';

            const rawWorkOrder = getVal(colMap.workOrderNo);
            const rawTaskNo = getVal(colMap.taskNo);

            if (!rawWorkOrder && !rawTaskNo) continue;

            const workOrderNo = String(rawWorkOrder || rawTaskNo).trim();
            const taskNo = String(rawTaskNo || rawWorkOrder).trim();

            if (seen.has(taskNo)) continue;
            seen.add(taskNo);

            const projectName = String(getVal(colMap.projectName)).trim();
            const rawDate = getVal(colMap.orderDate);
            const orderDate = parseExcelDate(rawDate);
            const orderQty = parseFloat(getVal(colMap.orderQty)) || 0;

            data.push({
                workOrderNo,
                taskNo,
                projectName,
                orderDate,
                orderQty
            });
        }

        const workOrdersFile = path.join(SUB_DIRS.schedule, 'work_orders.json');
        const storageData = {
            data: data,
            updateTime: new Date().toISOString(),
            source: 'excel_auto_import',
            sourceFile: newFilename // 记录来源文件
        };

        fs.writeFileSync(workOrdersFile, JSON.stringify(storageData, null, 2), 'utf8');

        console.log(`[AutoImport] 自动抓取并导入 ${data.length} 条数据，归档为: ${newFilename}`);

        return res.json({
            success: true,
            count: data.length,
            file: newFilename
        });

    } catch (error) {
        console.error('WatchExport error:', error);
        res.status(500).json({ success: false, message: '自动导入失败: ' + error.message });
    }
});

// 同步工单数据（来自篡改猴脚本）
app.post('/api/sync-work-orders', async (req, res) => {
    try {
        const { data } = req.body;
        if (!Array.isArray(data)) {
            return res.status(400).json({ success: false, message: '数据格式错误: 必须是数组' });
        }

        const workOrdersFile = path.join(SUB_DIRS.schedule, 'work_orders.json');

        // 保存数据
        const storageData = {
            data: data,
            updateTime: new Date().toISOString(),
            source: 'tampermonkey'
        };

        fs.writeFileSync(workOrdersFile, JSON.stringify(storageData, null, 2), 'utf8');

        console.log(`[Sync] 收到 ${data.length} 条工单数据`);

        return res.json({
            success: true,
            count: data.length,
            message: `成功同步 ${data.length} 条工单数据`
        });
    } catch (error) {
        console.error('Sync error:', error);
        return res.status(500).json({ success: false, message: error.message });
    }
});

// 获取已同步的工单数据
app.get('/api/work-orders', async (req, res) => {
    try {
        const workOrdersFile = path.join(SUB_DIRS.schedule, 'work_orders.json');

        if (fs.existsSync(workOrdersFile)) {
            const content = fs.readFileSync(workOrdersFile, 'utf8');
            const data = JSON.parse(content);
            return res.json({ success: true, ...data });
        } else {
            return res.json({ success: true, data: [], updateTime: null });
        }
    } catch (error) {
        return res.status(500).json({ success: false, message: error.message });
    }
});

// 启动服务器
app.listen(PORT, () => {
    console.log(`\n========================================`);
    console.log(`  物料执行情况查询工具`);
    console.log(`  服务已启动: http://localhost:${PORT}`);
    console.log(`========================================\n`);
});
