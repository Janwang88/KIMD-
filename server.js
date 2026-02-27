const express = require('express');
const path = require('path');
const fs = require('fs');
const os = require('os');
const XLSX = require('xlsx');
const sqlite3 = require('better-sqlite3');

// 导入外部模块
const { fetchWithFallback, httpsAgent } = require('./services/kimdProxy');
const { findLatestExcel, waitForNewExcel, computeStatsFromExcel, extractMilestonesFromSchedule } = require('./utils/excelHelper');
const { startDailyCleanup } = require('./utils/garbageCollector');

const app = express();
const PORT = 3000;

const BASE_URL = 'https://chajian.kimd.cn:9999';
const BASE_IP_URL = 'https://153.35.130.19:9999';
const DATA_DIR = path.join(__dirname, 'data');
const CACHED_EXCEL = path.join(DATA_DIR, 'material_progress.xlsx');

if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
}

// 建立并初始化 SQLite 数据库连接
const db = new sqlite3(path.join(DATA_DIR, 'reviews.db'));
db.pragma('journal_mode = WAL');
db.exec('CREATE TABLE IF NOT EXISTS reviews (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id TEXT, content TEXT NOT NULL, milestone TEXT, created_at DATETIME DEFAULT CURRENT_TIMESTAMP);');

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

// 启动每天自动清理任务 (保留 30 天)
startDailyCleanup(DATA_DIR, 30);

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
        const pwd = (password || '').trim();
        const code = (userCode || '').trim();

        // 1. 本地账户数据中心优先授权拦截
        const usersDbPath = path.join(__dirname, 'data', 'users.json');
        if (fs.existsSync(usersDbPath)) {
            try {
                const localUsers = JSON.parse(fs.readFileSync(usersDbPath, 'utf8') || '[]');
                const foundLocal = localUsers.find(u => u.userCode === code && u.password === pwd);
                if (foundLocal) {
                    return res.json({
                        success: true,
                        data: {
                            userName: foundLocal.userName || foundLocal.userCode,
                            token: 'LOCAL_TOKEN_' + Date.now(),
                            role: foundLocal.role || 'user',
                            department: foundLocal.department || ''
                        }
                    });
                }
            } catch (e) {
                console.error('读取本地授权库失败:', e);
            }
        }

        // 2. 如果本地没有，再兜底尝试原版的远程探测 (KIMD SYSLogin)
        const attempts = [];
        const companyValue = (company || '').trim();
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

        pushAttempts('userCode', 'password', 'company');

        const payloadVariants = [];
        // 缩减无用的 x-www-form-urlencoded，防止对面 C# 因为格式抛出解析基元错误
        attempts.forEach((a) => {
            payloadVariants.push({
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(a)
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

        // 扩充监控词列表，提高文件捕获成功率
        let searchPattern = pattern || ['物料', 'Material', 'Export', '进度', '清单', '工单', 'PPBOM'];
        if (workOrder) {
            if (!Array.isArray(searchPattern)) searchPattern = [searchPattern];
            const woPart = workOrder.split(/[, \n]/)[0].trim();
            if (woPart && !searchPattern.includes(woPart)) {
                searchPattern.push(woPart);
            }
        }

        console.log(`[ExcelWait] 监控已启动: 工单=${workOrder || '未知'}, 特征词=[${searchPattern.join(', ')}], 起始时间=${new Date(sinceMs).toLocaleString()}`);

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

        let stats = null;
        try {
            console.log(`[Stats] 开始解析 Excel: ${targetPath}`);
            stats = computeStatsFromExcel(targetPath);
            console.log(`[Stats] 解析完成: ${targetPath}, 行数: ${stats ? stats.rows : 0}`);
        } catch (parseErr) {
            console.error('[Stats] Excel 解析崩溃:', parseErr);
            return res.json({ success: false, message: 'Excel 文件解析失败，可能格式不兼容: ' + parseErr.message });
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

    const RULES = {
        mixed: ['项目管理', '领料', '上线准备', '总装', '清洁', '打包'],
        assembly: ['组装-返工', '模组组装', '整机接气', '出货'],
        wiring: ['接线-返工', '电控配线', '整机接线', '通电通气']
    };

    const stats = {
        totalHours: 0,
        assembly: { plan: 0, kimd: 0, outsource: 0, total: 0, processes: [...RULES.assembly], processBreakdown: {} },
        wiring: { plan: 0, kimd: 0, outsource: 0, total: 0, processes: [...RULES.wiring], processBreakdown: {} },
        mixed: { plan: 0, kimd: 0, outsource: 0, total: 0, processes: [...RULES.mixed], processBreakdown: {} },
        details: []
    };

    const ruleSets = {
        mixed: new Set(RULES.mixed),
        assembly: new Set(RULES.assembly),
        wiring: new Set(RULES.wiring)
    };

    let projectName = '';
    // 探测项目名称
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
        const row = rows[i];
        const strRow = row.map(c => (c || '').toString().trim());
        const pIdx = strRow.findIndex(c => c.includes('项目名称') || c === '项目' || c === '项目名称:');
        if (pIdx !== -1 && row[pIdx + 1]) {
            projectName = row[pIdx + 1].toString().trim();
            break;
        }
        // 或者字段就在当前列之后
        const pFieldIdx = strRow.findIndex(c => c.includes('项目名称:'));
        if (pFieldIdx !== -1) {
            const match = strRow[pFieldIdx].match(/项目名称[:：]\s*(.*)/);
            if (match) projectName = match[1].trim();
        }
    }

    // 遍历数据行
    for (let i = headerRowIndex + 1; i < rows.length; i++) {
        const row = rows[i];
        const process = (row[colMap.process] || '').toString().trim();

        if (!process || process === '工艺名称' || process === '合计' || process === '总计') continue;

        let target = null;
        if (ruleSets.mixed.has(process)) target = stats.mixed;
        else if (ruleSets.assembly.has(process)) target = stats.assembly;
        else if (ruleSets.wiring.has(process)) target = stats.wiring;

        if (target) {
            const valPlan = row[colMap.planHours];
            const valKimd = row[colMap.kimdHours];
            const valOut = row[colMap.outsourceHours];
            const valActual = row[colMap.actualHours];

            const plan = parseFloat(valPlan || 0) || 0;
            const kimd = parseFloat(valKimd || 0) || 0;
            const outsource = parseFloat(valOut || 0) || 0;
            const reportedActual = parseFloat(valActual || 0) || 0;

            // 逻辑优化：如果有明确的“实际工时”列且不为0，以此为准；否则取 kimd + outsource
            const actual = reportedActual > 0 ? reportedActual : (kimd + outsource);

            target.plan += plan;
            target.kimd += kimd;
            target.outsource += outsource;
            target.total += actual;

            // 确保 processBreakdown 初始化
            target.processBreakdown[process] = target.processBreakdown[process] || { kimd: 0, outsource: 0 };
            target.processBreakdown[process].kimd += kimd;
            target.processBreakdown[process].outsource += outsource;
            stats.totalHours += actual;
        } else {
            console.log(`[HoursDebug] Row ${i} skipped '${process}' - Not in RULES`);
        }
    }
    stats.projectName = projectName;
    return stats;
}



// API: Wait for Actual Hours Excel
app.post('/api/hours-wait-stats', async (req, res) => {
    try {
        const { pattern, timeoutMs, workOrder, since } = req.body || {};
        // ★ 优先使用前端传入的精确时间戳；若未传则默认当前时刻（保守策略，不会误捕获旧文件）
        const sinceMs = (typeof since === 'number' && !isNaN(since)) ? since : Date.now();
        // ★ 工时文件名格式由KIMD系统决定，不固定包含"工时"等关键词
        // 因此不用文件名过滤（传null），只靠 sinceMs 时间戳来识别"这次"导出的新文件
        // sinceMs 是打开KIMD工时页面之前的精确时刻，足以唯一定位本次下载
        const searchPattern = null;

        console.log(`[HoursWait] 监控启动: 工单=${workOrder || '未知'}, 起始时间=${new Date(sinceMs).toLocaleString()}, 策略=时间戳匹配(不限文件名)`);

        const latest = await waitForNewExcel(searchPattern, sinceMs, timeoutMs || 180000);

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
            orderQty: -1,
            assemblyStart: -1,
            assemblyEnd: -1,
            debugStart: -1,
            debugEnd: -1,
            shipStart: -1
        };

        const keywords = {
            workOrderNo: ['工单号', 'MO单', '生产订单'],
            taskNo: ['任务单', '计划单', '排程单'],
            projectName: ['项目名称', '项目', 'Project'],
            orderDate: ['接单时间', '下单日期', '日期', 'Date', '制单日期'],
            productName: ['产品名称', '物料名称', '产品'],
            orderQty: ['工单数量', '订单数量', '数量', 'Qty', 'Quantity'],
            assemblyStart: ['组装计划开始', '装配开始'],
            assemblyEnd: ['组装计划结束', '装配结束'],
            debugStart: ['工程调试计划开始', '调试开始'],
            debugEnd: ['工程调试计划结束', '调试结束', '入库计划开始'],
            shipStart: ['出货计划开始', '发货日期', '实际出货']
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
                else if (keywords.assemblyStart.some(k => cell.includes(k))) {
                    if (newMap.assemblyStart === -1 || cell === '组装计划开始') { newMap.assemblyStart = idx; matchCount++; }
                }
                else if (keywords.assemblyEnd.some(k => cell.includes(k))) {
                    if (newMap.assemblyEnd === -1 || cell === '组装计划结束') { newMap.assemblyEnd = idx; matchCount++; }
                }
                else if (keywords.debugStart.some(k => cell.includes(k))) {
                    if (newMap.debugStart === -1 || cell === '工程调试计划开始') { newMap.debugStart = idx; matchCount++; }
                }
                else if (keywords.debugEnd.some(k => cell.includes(k))) {
                    if (newMap.debugEnd === -1 || cell === '工程调试计划结束') { newMap.debugEnd = idx; matchCount++; }
                }
                else if (keywords.shipStart.some(k => cell.includes(k))) {
                    if (newMap.shipStart === -1 || cell === '出货计划开始' || cell === '实际出货') { newMap.shipStart = idx; matchCount++; }
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
                const date_info = new Date(Math.round((val - 25569) * 86400 * 1000));
                return `${date_info.getUTCFullYear()}-${String(date_info.getUTCMonth() + 1).padStart(2, '0')}-${String(date_info.getUTCDate()).padStart(2, '0')}`;
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

            const assemblyStart = parseExcelDate(getVal(colMap.assemblyStart)) || '-';
            const assemblyEnd = parseExcelDate(getVal(colMap.assemblyEnd)) || '-';
            const debugStart = parseExcelDate(getVal(colMap.debugStart)) || '-';
            const debugEnd = parseExcelDate(getVal(colMap.debugEnd)) || '-';
            const shipStart = parseExcelDate(getVal(colMap.shipStart)) || '-';

            data.push({
                workOrderNo,
                taskNo,
                projectName,
                orderDate,
                orderQty,
                assemblyStart,
                assemblyEnd,
                debugStart,
                debugEnd,
                shipStart
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
            orderQty: -1,
            assemblyStart: -1,
            assemblyEnd: -1,
            debugStart: -1,
            debugEnd: -1,
            shipStart: -1
        };

        const keywords = {
            workOrderNo: ['工单号', 'MO单', '生产订单'],
            taskNo: ['任务单', '计划单', '排程单'],
            projectName: ['项目名称', '项目', 'Project'],
            // 增加 "接单" 匹配
            orderDate: ['接单', '接单时间', '下单日期', '日期', 'Date', '制单日期'],
            productName: ['产品名称', '物料名称', '产品'],
            orderQty: ['工单数量', '订单数量', '数量', 'Qty', 'Quantity'],
            assemblyStart: ['组装计划开始', '装配开始'],
            assemblyEnd: ['组装计划结束', '装配结束'],
            debugStart: ['工程调试计划开始', '调试开始'],
            debugEnd: ['工程调试计划结束', '调试结束', '入库计划开始'],
            shipStart: ['出货计划开始', '发货日期', '实际出货']
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
                else if (keywords.assemblyStart.some(k => cell.includes(k))) {
                    if (newMap.assemblyStart === -1 || cell === '组装计划开始') { newMap.assemblyStart = idx; matchCount++; }
                }
                else if (keywords.assemblyEnd.some(k => cell.includes(k))) {
                    if (newMap.assemblyEnd === -1 || cell === '组装计划结束') { newMap.assemblyEnd = idx; matchCount++; }
                }
                else if (keywords.debugStart.some(k => cell.includes(k))) {
                    if (newMap.debugStart === -1 || cell === '工程调试计划开始') { newMap.debugStart = idx; matchCount++; }
                }
                else if (keywords.debugEnd.some(k => cell.includes(k))) {
                    if (newMap.debugEnd === -1 || cell === '工程调试计划结束') { newMap.debugEnd = idx; matchCount++; }
                }
                else if (keywords.shipStart.some(k => cell.includes(k))) {
                    if (newMap.shipStart === -1 || cell === '出货计划开始' || cell === '实际出货') { newMap.shipStart = idx; matchCount++; }
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
                return `${date_info.getUTCFullYear()}-${String(date_info.getUTCMonth() + 1).padStart(2, '0')}-${String(date_info.getUTCDate()).padStart(2, '0')}`;
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

            const assemblyStart = parseExcelDate(getVal(colMap.assemblyStart)) || '-';
            const assemblyEnd = parseExcelDate(getVal(colMap.assemblyEnd)) || '-';
            const debugStart = parseExcelDate(getVal(colMap.debugStart)) || '-';
            const debugEnd = parseExcelDate(getVal(colMap.debugEnd)) || '-';
            const shipStart = parseExcelDate(getVal(colMap.shipStart)) || '-';

            data.push({
                workOrderNo,
                taskNo,
                projectName,
                orderDate,
                orderQty,
                assemblyStart,
                assemblyEnd,
                debugStart,
                debugEnd,
                shipStart
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

// 获取特定工单的时间节点（从极速缓存提取）
app.get('/api/milestones', async (req, res) => {
    try {
        const { workOrder } = req.query;
        if (!workOrder) return res.status(400).json({ success: false, message: 'Missing workOrder' });

        const workOrdersFile = path.join(SUB_DIRS.schedule, 'work_orders.json');
        if (!fs.existsSync(workOrdersFile)) {
            return res.json({ success: false, message: '暂无缓存数据，无法显示时间节点' });
        }

        const content = fs.readFileSync(workOrdersFile, 'utf8');
        const data = JSON.parse(content);
        const list = data.data || [];

        const targetWo = String(workOrder).trim();
        const matched = list.find(item => item.workOrderNo === targetWo || item.taskNo === targetWo || (item.workOrderNo && targetWo.includes(item.workOrderNo)) || (item.taskNo && targetWo.includes(item.taskNo)));

        if (!matched) {
            return res.json({ success: false, message: 'WorkOrder not found in schedule cache' });
        }

        const milestones = {
            assemblyStart: matched.assemblyStart || '-',
            assemblyEnd: matched.assemblyEnd || '-',
            debugStart: matched.debugStart || '-',
            debugEnd: matched.debugEnd || '-',
            shipStart: matched.shipStart || '-'
        };

        return res.json({ success: true, milestones });
    } catch (error) {
        return res.status(500).json({ success: false, message: error.message });
    }
});

// ==========================================
// 外协工时对账账单 API
// ==========================================

const outsourceDataPath = path.join(__dirname, 'data', 'outsource_manpower.json');

// 获取所有或按工单查询本地外协工时账单
app.get('/api/outsource/records', (req, res) => {
    try {
        if (!fs.existsSync(outsourceDataPath)) {
            return res.json({ success: true, data: [] });
        }
        const data = JSON.parse(fs.readFileSync(outsourceDataPath, 'utf8') || '[]');
        const { workOrder } = req.query;
        let filteredData = data;

        if (workOrder) {
            const woUpper = workOrder.trim().toUpperCase();
            filteredData = data.filter(r => r.workOrder && r.workOrder.toUpperCase().includes(woUpper));
        }
        // 根据添加时间倒序
        filteredData.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
        res.json({ success: true, data: filteredData });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

// 新增一条外协工时账单记录
app.post('/api/outsource/add', express.json(), (req, res) => {
    try {
        const { date, workOrder, projectName, workerName, workerLevel, supplier, workHours, shiftType, workContent } = req.body;
        if (!workOrder || !workerName || !workHours) {
            return res.status(400).json({ success: false, message: '工单号、姓名、工时 为必填项' });
        }

        let data = [];
        if (fs.existsSync(outsourceDataPath)) {
            data = JSON.parse(fs.readFileSync(outsourceDataPath, 'utf8') || '[]');
        }

        const newRecord = {
            id: Date.now().toString() + Math.floor(Math.random() * 1000),
            date: date || new Date().toISOString().split('T')[0],
            workOrder: workOrder.trim(),
            projectName: projectName || '',
            workerName: workerName.trim(),
            workerLevel: workerLevel || '大工',
            supplier: supplier || '默认关联单位',
            workHours: parseFloat(workHours),
            shiftType: shiftType || '白班',
            workContent: workContent || '',
            createdAt: new Date().toISOString()
        };

        data.push(newRecord);
        fs.writeFileSync(outsourceDataPath, JSON.stringify(data, null, 2), 'utf8');

        res.json({ success: true, data: newRecord });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

// 删除单独的工时台账
app.delete('/api/outsource/records/:id', (req, res) => {
    try {
        if (!fs.existsSync(outsourceDataPath)) return res.json({ success: false, message: 'No Data' });
        let data = JSON.parse(fs.readFileSync(outsourceDataPath, 'utf8') || '[]');
        const targetId = req.params.id;
        const initLen = data.length;
        data = data.filter(r => r.id !== targetId);

        if (initLen === data.length) return res.status(404).json({ success: false, message: 'Not found' });
        fs.writeFileSync(outsourceDataPath, JSON.stringify(data, null, 2), 'utf8');

        res.json({ success: true, message: 'Deleted' });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

// ==========================================
// 评论及备忘点 (Reviews) 相关 API
// ==========================================

// 添加一条新的记录
app.post('/api/reviews', express.json(), (req, res) => {
    try {
        const { content, milestone, user_id } = req.body;
        if (!content) {
            return res.status(400).json({ success: false, message: '内容不能为空' });
        }

        const stmt = db.prepare('INSERT INTO reviews (content, milestone, user_id) VALUES (?, ?, ?)');
        const info = stmt.run(content, milestone || null, user_id || null);

        res.json({
            success: true,
            data: { id: info.lastInsertRowid, content, milestone, user_id }
        });
    } catch (error) {
        console.error('Add review error:', error);
        res.status(500).json({ success: false, message: '无法保存记录' });
    }
});

// 获取记录列表
app.get('/api/reviews', (req, res) => {
    try {
        const { milestone, limit = 50, offset = 0 } = req.query;
        let stmt;
        let rows;

        if (milestone) {
            stmt = db.prepare('SELECT * FROM reviews WHERE milestone = ? ORDER BY created_at DESC LIMIT ? OFFSET ?');
            rows = stmt.all(milestone, Number(limit), Number(offset));
        } else {
            stmt = db.prepare('SELECT * FROM reviews ORDER BY created_at DESC LIMIT ? OFFSET ?');
            rows = stmt.all(Number(limit), Number(offset));
        }

        res.json({ success: true, data: rows });
    } catch (error) {
        console.error('Get reviews error:', error);
        res.status(500).json({ success: false, message: '查询记录失败' });
    }
});

// 删除记录
app.delete('/api/reviews/:id', (req, res) => {
    try {
        const stmt = db.prepare('DELETE FROM reviews WHERE id = ?');
        const info = stmt.run(req.params.id);

        if (info.changes > 0) {
            res.json({ success: true, message: '记录已删除' });
        } else {
            res.status(404).json({ success: false, message: '未找到该记录' });
        }
    } catch (error) {
        res.status(500).json({ success: false, message: '删除失败' });
    }
});

// ==========================================
// 账号授权数据中心 API
// ==========================================
const usersDbPath = path.join(__dirname, 'data', 'users.json');

app.get('/api/auth/users', (req, res) => {
    try {
        if (!fs.existsSync(usersDbPath)) return res.json({ success: true, data: [] });
        const list = JSON.parse(fs.readFileSync(usersDbPath, 'utf8') || '[]');
        // 不要把所有人的明文密码返回给页面，脱敏处理
        const safeList = list.map(u => ({ userCode: u.userCode, userName: u.userName, role: u.role, department: u.department, createdAt: u.createdAt }));
        res.json({ success: true, data: safeList });
    } catch (error) {
        res.status(500).json({ success: false, message: error.message });
    }
});

app.post('/api/auth/addUser', express.json(), (req, res) => {
    try {
        const { userCode, password, userName, role, department } = req.body;
        if (!userCode) return res.status(400).json({ success: false, message: '账号必填' });

        let list = [];
        if (fs.existsSync(usersDbPath)) list = JSON.parse(fs.readFileSync(usersDbPath, 'utf8') || '[]');

        const existingIndex = list.findIndex(u => u.userCode === userCode);
        if (existingIndex >= 0) {
            // 更新编辑模式
            if (password) list[existingIndex].password = password;
            if (userName) list[existingIndex].userName = userName;
            if (role) list[existingIndex].role = role;
            if (department) list[existingIndex].department = department;
            fs.writeFileSync(usersDbPath, JSON.stringify(list, null, 2), 'utf8');
            return res.json({ success: true, message: '用户信息更新成功！' });
        }

        // 新增模式
        if (!password) return res.status(400).json({ success: false, message: '新增账号和密码必填' });
        list.push({
            userCode, password, userName: userName || userCode, role: role || 'user', department: department || '', createdAt: new Date().toISOString()
        });
        fs.writeFileSync(usersDbPath, JSON.stringify(list, null, 2), 'utf8');
        res.json({ success: true, message: '授权新用户成功！' });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

app.delete('/api/auth/users/:code', (req, res) => {
    try {
        const code = req.params.code;
        if (code === 'wangjian') return res.status(403).json({ success: false, message: '超级管理员不可被剥夺权限' });

        if (!fs.existsSync(usersDbPath)) return res.json({ success: true });
        let list = JSON.parse(fs.readFileSync(usersDbPath, 'utf8') || '[]');
        const originLen = list.length;
        list = list.filter(u => u.userCode !== code);

        if (list.length === originLen) return res.status(404).json({ success: false, message: '用户不存在' });

        fs.writeFileSync(usersDbPath, JSON.stringify(list, null, 2), 'utf8');
        res.json({ success: true, message: '已吊销授权' });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

app.get('/api/outsource-hours', (req, res) => {
    try {
        const { workOrder } = req.query;
        if (!workOrder) {
            return res.status(400).json({ success: false, message: '请提供工单号' });
        }

        const targets = workOrder.split(/[, \n]/).map(s => s.trim()).filter(Boolean);
        if (targets.length === 0) {
            return res.json({ success: true, total: 0 });
        }

        const placeholders = targets.map(() => '?').join(',');

        // 外协记录（用于本地外协列）
        const outsourceRows = db.prepare(`SELECT content, hours FROM outsource_manpower WHERE work_order IN (${placeholders}) AND (category = '外协' OR category IS NULL OR category = '')`).all(targets);

        // KIMD 记录（用于本地KIMD列）
        const kimdRows = db.prepare(`SELECT content, hours FROM outsource_manpower WHERE work_order IN (${placeholders}) AND category = 'KIMD'`).all(targets);

        const RULES = {
            mixed: new Set(['项目管理', '领料', '上线准备', '总装', '清洁', '打包']),
            assembly: new Set(['组装-返工', '模组组装', '整机接气', '出货']),
            wiring: new Set(['接线-返工', '电控配线', '整机接线', '通电通气'])
        };

        let stats = {
            total: 0,
            assembly: 0,
            mixed: 0,
            wiring: 0,
            uncategorized: 0,
            processBreakdown: {},     // 记录各个工序的外协总计
            kimdBreakdown: {},        // 记录各个工序的KIMD总计
            detailedProcessBreakdown: {} // 细化每个工序： { kimd: 0, outsource: { "大工":0, "中工":0, "小工":0, "total":0 } }
        };

        const initDetailedProcess = (process) => {
            if (!stats.detailedProcessBreakdown[process]) {
                stats.detailedProcessBreakdown[process] = {
                    kimd: 0,
                    outsource: { '大工': 0, '中工': 0, '小工': 0, 'total': 0 }
                };
            }
        };

        // 统计外协数据
        outsourceRows.forEach(row => {
            const h = parseFloat(row.hours) || 0;
            const process = (row.content || '').trim();
            const level = (row.worker_level || '').trim() || '大工'; // 默认无等级计为大工
            stats.total += h;

            if (RULES.mixed.has(process)) stats.mixed += h;
            else if (RULES.assembly.has(process)) stats.assembly += h;
            else if (RULES.wiring.has(process)) stats.wiring += h;
            else stats.uncategorized += h;

            if (process) {
                stats.processBreakdown[process] = (stats.processBreakdown[process] || 0) + h;

                initDetailedProcess(process);
                if (stats.detailedProcessBreakdown[process].outsource[level] !== undefined) {
                    stats.detailedProcessBreakdown[process].outsource[level] += h;
                }
                stats.detailedProcessBreakdown[process].outsource.total += h;
            }
        });

        // 统计 KIMD 数据（按工序汇总）
        kimdRows.forEach(row => {
            const h = parseFloat(row.hours) || 0;
            const process = (row.content || '').trim();
            if (process) {
                stats.kimdBreakdown[process] = (stats.kimdBreakdown[process] || 0) + h;

                initDetailedProcess(process);
                stats.detailedProcessBreakdown[process].kimd += h;
            }
        });

        res.json({ success: true, ...stats });
    } catch (e) {
        console.error('Error fetching outsource hours:', e);
        res.status(500).json({ success: false, message: e.message });
    }
});

// ====== 外借人力数据库操作接口 ======

// 获取所有/分页数据
app.get('/api/outsource-records', (req, res) => {
    try {
        const { keyword = '', searchDate = '', page = 1, pageSize = 100 } = req.query;
        console.log('[DEBUG] GET /api/outsource-records', req.query, 'page:', page, 'typeof page:', typeof page);

        let queryStr = `SELECT * FROM outsource_manpower WHERE 1=1`;
        let countStr = `SELECT COUNT(*) as total FROM outsource_manpower WHERE 1=1`;
        let params = [];

        if (keyword) {
            queryStr += ` AND (work_order LIKE ? OR worker_name LIKE ? OR project_name LIKE ?)`;
            countStr += ` AND (work_order LIKE ? OR worker_name LIKE ? OR project_name LIKE ?)`;
            const k = `%${keyword}%`;
            params.push(k, k, k);
        }

        if (searchDate) {
            queryStr += ` AND work_date = ?`;
            countStr += ` AND work_date = ?`;
            params.push(searchDate);
        }

        queryStr += ` ORDER BY id DESC LIMIT ? OFFSET ?`;

        const p = Number(page) || 1;
        const ps = Number(pageSize) || 20;
        const offset = (p - 1) * ps;

        console.log('[DEBUG] bind values:', params, ps, offset);

        const totalRow = db.prepare(countStr).get(...params);

        const dataParams = [...params, ps, offset];
        const data = db.prepare(queryStr).all(...dataParams);

        res.json({ success: true, total: totalRow.total, data });
    } catch (e) {
        console.error('[DEBUG] /api/outsource-records ERROR:', e);
        res.status(500).json({ success: false, message: e.message });
    }
});

// 获取推荐人员 (最近一天的名单)
app.get('/api/outsource-record/suggest', (req, res) => {
    try {
        const { workOrder, category, content } = req.query;
        console.log(`[SuggestAPI] WO: ${workOrder}, Category: ${category}, Content: ${content}`);

        if (!workOrder) {
            return res.status(400).json({ success: false, message: '请提供工单号' });
        }

        // 定义工艺分组映射 (使用英文 Key 避免编码匹配问题)
        const groups = {
            'assembly': ['组装-返工', '模组组装', '整机接气', '出货'],
            'mixed': ['项目管理', '领料', '上线准备', '总装', '清洁', '打包'],
            'wiring': ['接线-返工', '电控配线', '整机接线', '通电通气']
        };

        // 确定要搜索的工艺列表
        let contentList = [];
        const normalizedCategory = (category || '').toLowerCase();

        if (groups[normalizedCategory]) {
            contentList = groups[normalizedCategory];
        } else if (content) {
            contentList = [content];
        }

        console.log(`[SuggestAPI] Resolved contentList (Category: ${normalizedCategory}):`, contentList);

        // 严谨逻辑：如果没有指定分类且没有指定具体工艺，不要返回任何推荐
        if (contentList.length === 0) {
            console.log(`[SuggestAPI] No contentList resolved, returning empty.`);
            return res.json({ success: true, data: [] });
        }

        // 1. 先找到该工单+工艺最近的一个日期
        let sqlRecentDate = 'SELECT work_date FROM outsource_manpower WHERE work_order = ?';
        let params = [workOrder];
        if (contentList.length > 0) {
            const placeholders = contentList.map(() => '?').join(',');
            sqlRecentDate += ` AND content IN (${placeholders})`;
            params.push(...contentList);
        }
        sqlRecentDate += ' ORDER BY id DESC LIMIT 1';

        const lastRecord = db.prepare(sqlRecentDate).get(...params);
        if (!lastRecord) {
            return res.json({ success: true, data: [] });
        }

        // 2. 获取那一天的所有人员 (同样限定在工艺组内)
        let sqlStaff = 'SELECT DISTINCT worker_name, worker_level, supplier, category, content FROM outsource_manpower WHERE work_order = ? AND work_date = ?';
        let paramsStaff = [workOrder, lastRecord.work_date];
        if (contentList.length > 0) {
            const placeholders = contentList.map(() => '?').join(',');
            sqlStaff += ` AND content IN (${placeholders})`;
            paramsStaff.push(...contentList);
        }

        const staff = db.prepare(sqlStaff).all(...paramsStaff);
        res.json({ success: true, data: staff });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

// 新增数据
app.post('/api/outsource-record', express.json(), (req, res) => {
    try {
        const { work_date, work_order, project_name, worker_name, worker_level, start_time, end_time, hours, supplier, shift, manager, content, remark, remark1, category } = req.body;
        const insert = db.prepare(`
            INSERT INTO outsource_manpower (
                work_date, work_order, project_name, worker_name, worker_level, 
                start_time, end_time, hours, supplier, shift, manager, content, remark, remark1, category
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `);
        const info = insert.run(
            work_date || '', work_order || '', project_name || '', worker_name || '', worker_level || '',
            start_time || '', end_time || '', parseFloat(hours) || 0, supplier || '', shift || '', manager || '', content || '', remark || '', remark1 || '', category || '外协'
        );
        res.json({ success: true, message: '添加成功', id: info.lastInsertRowid });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

// 更新数据
app.put('/api/outsource-record/:id', express.json(), (req, res) => {
    try {
        const id = req.params.id;
        const { work_date, work_order, project_name, worker_name, worker_level, start_time, end_time, hours, supplier, shift, manager, content, remark, remark1, category } = req.body;
        const update = db.prepare(`
            UPDATE outsource_manpower SET 
                work_date = ?, work_order = ?, project_name = ?, worker_name = ?, worker_level = ?, 
                start_time = ?, end_time = ?, hours = ?, supplier = ?, shift = ?, manager = ?, content = ?, remark = ?, remark1 = ?, category = ?
            WHERE id = ?
        `);
        const info = update.run(
            work_date || '', work_order || '', project_name || '', worker_name || '', worker_level || '',
            start_time || '', end_time || '', parseFloat(hours) || 0, supplier || '', shift || '', manager || '', content || '', remark || '', remark1 || '', category || '外协',
            id
        );
        if (info.changes > 0) {
            res.json({ success: true, message: '更新成功' });
        } else {
            res.status(404).json({ success: false, message: '未找到对应记录' });
        }
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

// 批量删除数据
app.post('/api/outsource-record/batch-delete', express.json(), (req, res) => {
    try {
        const { ids } = req.body;
        if (!Array.isArray(ids) || ids.length === 0) {
            return res.status(400).json({ success: false, message: '请提供有效的ID列表' });
        }
        const placeholders = ids.map(() => '?').join(',');
        const info = db.prepare(`DELETE FROM outsource_manpower WHERE id IN (${placeholders})`).run(...ids);
        res.json({ success: true, message: `成功删除 ${info.changes} 条记录` });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

// 批量修改数据 (仅修改工艺和备注)
app.post('/api/outsource-record/batch-update', express.json(), (req, res) => {
    try {
        const { ids, content, remark1 } = req.body;
        if (!Array.isArray(ids) || ids.length === 0) {
            return res.status(400).json({ success: false, message: '请提供有效的ID列表' });
        }

        let updateParts = [];
        let params = [];
        if (content !== undefined) {
            updateParts.push('content = ?');
            params.push(content);
        }
        if (remark1 !== undefined) {
            updateParts.push('remark1 = ?');
            params.push(remark1);
        }

        if (updateParts.length === 0) {
            return res.status(400).json({ success: false, message: '没有提供需要更新的字段' });
        }

        const placeholders = ids.map(() => '?').join(',');
        params.push(...ids);

        const info = db.prepare(`
            UPDATE outsource_manpower 
            SET ${updateParts.join(', ')}
            WHERE id IN (${placeholders})
        `).run(...params);

        res.json({ success: true, message: `成功更新 ${info.changes} 条记录` });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

// 删除数据
app.delete('/api/outsource-record/:id', (req, res) => {
    try {
        const id = req.params.id;
        const info = db.prepare('DELETE FROM outsource_manpower WHERE id = ?').run(id);
        if (info.changes > 0) {
            res.json({ success: true, message: '删除成功' });
        } else {
            res.status(404).json({ success: false, message: '未找到对应记录' });
        }
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

// 批量导入外协数据 (支持事务)
app.post('/api/outsource-record/import', express.json({ limit: '50mb' }), (req, res) => {
    try {
        const { records } = req.body;
        if (!records || !Array.isArray(records)) {
            return res.status(400).json({ success: false, message: '无效的数据格式' });
        }

        const insert = db.prepare(`
            INSERT INTO outsource_manpower (
                work_date, work_order, project_name, worker_name,
                worker_level, hours, supplier, content, shift, remark1, category, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        `);

        // 使用事务以实现最佳性能和原子性
        const insertMany = db.transaction((rows) => {
            let count = 0;
            for (const row of rows) {
                // 如果必需的关键字段为空，可以选择跳过
                if (!row.work_order && !row.worker_name && !row.hours) continue;

                insert.run(
                    row.work_date || '',
                    row.work_order || '',
                    row.project_name || '',
                    row.worker_name || '',
                    row.worker_level || '',
                    parseFloat(row.hours) || 0,
                    row.supplier || '',
                    row.content || '',
                    row.shift || '',
                    row.remark1 || '',
                    row.category || '外协'
                );
                count++;
            }
            return count;
        });

        const successCount = insertMany(records);
        res.json({ success: true, message: `成功导入 ${successCount} 条记录`, count: successCount });

    } catch (error) {
        console.error('Import Error:', error);
        res.status(500).json({ success: false, message: '导入过程中发生错误: ' + error.message });
    }
});

// 启动服务器
app.listen(PORT, () => {
    console.log(`\n========================================`);
    console.log(`  物料执行情况查询工具`);
    console.log(`  服务已启动: http://localhost:${PORT}`);
    console.log(`========================================\n`);
});
