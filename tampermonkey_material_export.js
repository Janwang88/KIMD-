// ==UserScript==
// @name         KIMD 物料执行情况明细自动导出并统计 v5.5
// @namespace    kimd
// @version      5.5
// @match        https://chajian.kimd.cn:9999/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const STATS_URL = 'http://localhost:3000/material_stats.html';
    const log = (...args) => console.log('%c[KIMD]', 'color: #1890ff; font-weight: bold;', ...args);

    log('Script loaded! v5.5 Current URL:', location.href);

    // --- Common Helpers ---
    function normalize(t) { return (t || '').replace(/\s+/g, '').trim(); }

    function parseHashParams() {
        const hash = location.hash || '';
        const query = hash.includes('?') ? hash.slice(hash.indexOf('?') + 1) : '';
        return new URLSearchParams(query);
    }

    function parseParamStr() {
        const p = parseHashParams();
        const raw = p.get('parameterStr');
        if (!raw) return null;
        try { return JSON.parse(decodeURIComponent(raw)); } catch (e) { return null; }
    }

    function extractWorkOrderFromParamStr(obj) {
        try {
            const filter = obj?.filterDic?.filter || '';
            const decoded = decodeURIComponent(filter);
            let m = decoded.match(/工单\s*=\s*'([^']+)'/);
            if (m) return m[1];
            m = decoded.match(/工单\s+LIKE\s+'%([^%']+)%'/);
            return m ? m[1] : '';
        } catch (e) { return ''; }
    }

    function getWorkOrder() {
        const p = parseHashParams();
        const fromQuery = p.get('workOrder');
        if (fromQuery) return fromQuery.trim();
        const obj = parseParamStr();
        return extractWorkOrderFromParamStr(obj).trim();
    }

    function getProject() {
        const p = parseHashParams();
        return decodeURIComponent((p.get('project') || '').trim());
    }

    async function waitUntil(action, timeoutMs = 20000, intervalMs = 300) {
        const start = Date.now();
        while (Date.now() - start < timeoutMs) {
            try {
                const res = action();
                if (res) return res;
            } catch (e) { }
            await new Promise(r => setTimeout(r, intervalMs));
        }
        return null;
    }

    async function simulateTyping(input, text) {
        input.focus();
        await new Promise(r => setTimeout(r, 50));
        const valueSetter = Object.getOwnPropertyDescriptor(input, 'value')?.set;
        const prototype = Object.getPrototypeOf(input);
        const prototypeValueSetter = Object.getOwnPropertyDescriptor(prototype, 'value')?.set;
        const setValue = (val) => {
            if (valueSetter && valueSetter !== prototypeValueSetter) prototypeValueSetter.call(input, val);
            else if (valueSetter) valueSetter.call(input, val);
            else input.value = val;
            input.dispatchEvent(new Event('input', { bubbles: true }));
        };
        setValue('');
        for (let i = 0; i < text.length; i++) {
            setValue(input.value + text[i]);
            await new Promise(r => setTimeout(r, 10));
        }
        input.dispatchEvent(new Event('change', { bubbles: true }));
        input.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter', code: 'Enter', keyCode: 13 }));
        input.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'Enter', code: 'Enter', keyCode: 13 }));
        input.blur();
    }

    function clickEl(el) {
        if (!el) return false;
        el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        el.dispatchEvent(new MouseEvent('click', { bubbles: true }));
        return true;
    }

    // --- Material Progress Logic ---
    function isMaterialPage() {
        return (location.hash || '').includes('/wms/reportManage/materialProgres');
    }

    async function runMaterial() {
        const workOrder = getWorkOrder();
        if (!workOrder) return;
        const project = getProject();

        log('Start Material:', workOrder);

        // Find Input
        const findInput = () => Array.from(document.querySelectorAll('input')).find(i => (i.placeholder || '').includes('工单'));
        const input = await waitUntil(findInput, 10000);
        if (!input) return;

        // Type
        await simulateTyping(input, workOrder);
        await new Promise(r => setTimeout(r, 500));
        if (normalize(input.value) !== normalize(workOrder)) {
            input.value = workOrder;
            input.dispatchEvent(new Event('input', { bubbles: true }));
        }

        // Click Query
        const findQuery = () => Array.from(document.querySelectorAll('button, .el-button')).find(b => normalize(b.innerText).includes('查询') || normalize(b.innerText).includes('搜索'));
        const queryBtn = findQuery();
        if (queryBtn) clickEl(queryBtn);

        await new Promise(r => setTimeout(r, 3000));

        // Export Loop
        let exported = false;
        for (let i = 0; i < 5; i++) {
            const findExport = () => Array.from(document.querySelectorAll('button, .el-button')).find(b => normalize(b.innerText).includes('导出'));
            const exportBtn = findExport();
            if (exportBtn) clickEl(exportBtn);

            await new Promise(r => setTimeout(r, 1500));

            // Confirm Dialog (More robust check)
            const findConfirm = () => {
                // Try standard Element UI message box
                const wrappers = Array.from(document.querySelectorAll('.el-message-box__wrapper, .el-dialog__wrapper'))
                    .filter(w => w.style.display !== 'none' && !w.innerText.includes('display: none'));

                for (const w of wrappers) {
                    // Check for Primary button or button with text '确定'
                    const btns = Array.from(w.querySelectorAll('button, .el-button, .el-button--primary'));
                    const target = btns.find(b => {
                        const t = normalize(b.innerText);
                        return t === '确定' || t === 'OK' || t === '确认';
                    });
                    if (target) return target;
                }

                // Fallback: finding any button with '确定' that is visible and has high z-index
                const allBtns = Array.from(document.querySelectorAll('button, .el-button--primary'));
                return allBtns.find(b => {
                    return (normalize(b.innerText) === '确定' || normalize(b.innerText) === '确认') && b.offsetParent !== null;
                });
            };

            // Wait slightly longer for animation
            await new Promise(r => setTimeout(r, 500));
            const confirmBtn = await waitUntil(findConfirm, 5000, 500); // 5s wait

            if (confirmBtn) {
                log('Found confirm button, clicking...');
                clickEl(confirmBtn);
                exported = true;
                break;
            } else {
                log('No confirm button found in this attempt');
            }
        }

        // Notify
        const p = parseHashParams();
        if (p.get('auto') === 'true' && exported) {
            const div = document.createElement('div');
            div.innerHTML = `<div style="position:fixed;top:10px;right:10px;z-index:9999;background:#52c41a;color:white;padding:15px">✅ 原料下载触发</div>`;
            document.body.appendChild(div);
            if (window.opener) window.opener.postMessage({ type: 'kimd-download-triggered', payload: { workOrder } }, '*');
        } else if (exported) {
            openStatsAndNotify(workOrder, project);
        }
    }

    function openStatsAndNotify(workOrder, project) {
        const url = `${STATS_URL}?workOrder=${encodeURIComponent(workOrder)}&project=${encodeURIComponent(project || '')}&t=${Date.now()}`;
        const win = window.open(url, `kimd-stats-${workOrder}`);
        if (win) {
            let n = 0;
            const timer = setInterval(() => {
                n++;
                try { win.postMessage({ type: 'kimd-excel-wait', payload: { workOrder, project } }, 'http://localhost:3000'); } catch (e) { }
                if (n >= 8) clearInterval(timer);
            }, 600);
        }
    }

    // --- Actual Hours Logic ---
    function isHoursPage() { return (location.hash || '').includes('/sc/work/actualHour'); }

    async function runHours() {
        const p = parseHashParams();
        const auto = p.get('auto') === 'true';
        const targetWo = p.get('workOrder');
        if (!auto || !targetWo) return;

        log('Start Hours:', targetWo);

        // 1. Input Work Order
        const findInput = () => Array.from(document.querySelectorAll('input')).find(i => (i.placeholder || '').includes('工单'));
        const input = await waitUntil(findInput, 10000);
        if (!input) return;

        // ★ 关键修复：先等页面初始加载完毕，再输入工单号查询
        // 避免：脚本查询成功→初始慢加载数据晚到→覆盖了正确结果
        log('Waiting for page initial load to complete before querying...');
        await new Promise(r => setTimeout(r, 800)); // 短暂等一下让初始loading出现
        const waitForInitialLoad = async () => {
            const start = Date.now();
            let foundLoading = false;
            while (Date.now() - start < 20000) {
                const loading = document.querySelector('.el-loading-mask, .el-loading-spinner');
                if (loading && loading.style.display !== 'none' && getComputedStyle(loading).display !== 'none') {
                    foundLoading = true; // 找到了loading，继续等它消失
                } else if (foundLoading) {
                    break; // loading出现过又消失了，说明加载完成
                }
                await new Promise(r => setTimeout(r, 400));
            }
            // 兜底：若始终未出现loading元素（某些页面不显示spinner），至少固定等3秒，确保默认数据加载完
            if (!foundLoading) {
                log('No loading indicator found, using fallback wait of 3s...');
                await new Promise(r => setTimeout(r, 3000));
            } else {
                await new Promise(r => setTimeout(r, 800)); // loading消失后再等一点
            }
        };
        await waitForInitialLoad();
        log('Initial load done. Now inputting work order...');

        if (normalize(input.value) !== normalize(targetWo)) {
            await simulateTyping(input, targetWo);
            await new Promise(r => setTimeout(r, 500));
        }

        // 强制触发查询，确保系统刷新至最新工单数据
        const findQuery = () => Array.from(document.querySelectorAll('button, .el-button')).find(b => normalize(b.innerText).includes('查询'));
        const queryBtn = findQuery();
        if (queryBtn) {
            log('Clicking Query to force refresh...');
            clickEl(queryBtn);
            // 查询后等 loading 消失再找数据行
            await new Promise(r => setTimeout(r, 800));
            await waitForInitialLoad(); // 复用同一个等待函数
        }

        // 2. 遍历所有行，找到工单号匹配的那一行，而不只检查第一行
        // 这样可以避免页面数据未刷新或者多工单时抓到错误行
        const findDetailBtnAccurate = () => {
            const rows = Array.from(document.querySelectorAll('.el-table__row'));
            if (!rows.length) {
                log('Waiting for results rows...');
                return null;
            }

            for (const row of rows) {
                const rowText = normalize(row.innerText);
                if (rowText.includes(normalize(targetWo))) {
                    log('Match confirmed! Row text:', rowText.slice(0, 60));
                    const cells = row.querySelectorAll('td');
                    const lastCell = cells[cells.length - 1];
                    if (!lastCell) continue;
                    const btn = lastCell.querySelector('.el-button');
                    if (btn) return btn;
                }
            }

            log('No row matching WO found yet. Rows count:', rows.length,
                '| First row text:', rows[0] ? normalize(rows[0].innerText).slice(0, 60) : '(empty)');
            return null;
        };

        function showStatus(msg, color = '#6366f1') {
            const existing = document.getElementById('kimd-hours-status');
            if (existing) existing.remove();
            const div = document.createElement('div');
            div.id = 'kimd-hours-status';
            div.style.cssText = `position:fixed;top:10px;right:10px;z-index:9999;background:${color};color:white;padding:12px 18px;border-radius:8px;font-size:14px;font-weight:bold;max-width:320px;word-break:break-all;box-shadow:0 4px 12px rgba(0,0,0,0.3)`;
            div.textContent = msg;
            document.body.appendChild(div);
        }

        showStatus('⏳ 步骤2: 等待数据行...');
        const detailBtn = await waitUntil(findDetailBtnAccurate, 20000, 1000);
        if (!detailBtn) {
            showStatus('❌ 未找到工单行（20s超时），工单号：' + targetWo, '#ef4444');
            log('No matching data row found after 20s');
            return;
        }

        showStatus('✅ 步骤2: 找到行，点击详情...');
        clickEl(detailBtn);

        // 3. WAIT for Modal
        log('Waiting for modal...');
        showStatus('⏳ 步骤3: 等待弹窗...');
        await new Promise(r => setTimeout(r, 2500));

        const findExportInDialog = () => {
            const dialogs = Array.from(document.querySelectorAll('.el-dialog__wrapper')).filter(d => d.style.display !== 'none');
            if (!dialogs.length) return null;
            const topDialog = dialogs[dialogs.length - 1];
            const btns = Array.from(topDialog.querySelectorAll('button, .el-button'));
            return btns.find(b => normalize(b.innerText).includes('导出'));
        };

        const exportBtn = await waitUntil(findExportInDialog, 8000);
        if (!exportBtn) {
            showStatus('❌ 未找到导出按钮（8s超时）', '#ef4444');
            log('No export btn in dialog');
            return;
        }

        showStatus('✅ 步骤3: 找到导出按钮，点击...');
        clickEl(exportBtn);

        // 4. Confirm
        await new Promise(r => setTimeout(r, 1000));
        const findConfirm = () => {
            const wrappers = Array.from(document.querySelectorAll('.el-message-box__wrapper')).filter(w => w.style.display !== 'none');
            for (const w of wrappers) {
                const btn = Array.from(w.querySelectorAll('button, .el-button')).find(b => normalize(b.innerText) === '确定');
                if (btn) return btn;
            }
            return null;
        };
        const confirmBtn = await waitUntil(findConfirm, 3000);
        if (confirmBtn) {
            showStatus('✅ 步骤4: 点击确定，下载中...');
            clickEl(confirmBtn);
        } else {
            showStatus('⚠️ 未出现确认弹窗，可能已直接下载', '#f59e0b');
        }

        // 最终成功提示
        setTimeout(() => showStatus('✅ 工时下载触发完成！', '#10b981'), 1500);
    }

    async function main() {
        if (isMaterialPage()) await runMaterial();
        else if (isHoursPage()) await runHours();
    }

    setTimeout(main, 2000);
})();