// ==UserScript==
// @name         KIMD 生产排程数据同步助手
// @namespace    http://tampermonkey.net/
// @version      2.4
// @description  一键同步生产排程数据到本地物料查询工具（支持自动导出V2.4 - 超时优化版）
// @author       Antigravity
// @match        https://chajian.kimd.cn:9999/*
// @grant        GM_xmlhttpRequest
// @connect      localhost
// ==/UserScript==

(function () {
    'use strict';

    const API_URL = 'http://localhost:3000/api/sync-work-orders';
    let isSyncing = false;
    let autoExportTried = false;

    // 初始化
    function init() {
        console.log('KIMD 同步助手 V2.4 (超时优化版) 已加载');

        setTimeout(checkAutoExport, 1000);

        let lastUrl = location.href;
        new MutationObserver(() => {
            const url = location.href;
            if (url !== lastUrl) {
                lastUrl = url;
                setTimeout(() => {
                    checkPage();
                    checkAutoExport();
                }, 1000);
            }
        }).observe(document, { subtree: true, childList: true });

        setInterval(() => {
            checkPage();
            // 持续尝试直到成功
            if (location.href.includes('autoExport=true') && !autoExportTried) {
                tryClickExport();
            }
        }, 2000);

        checkPage();
    }

    function checkAutoExport() {
        if (autoExportTried) return;
        if (location.href.includes('autoExport=true')) {
            console.log('[AutoExport] 检测到自动导出请求...');
            tryClickExport();
        }
    }

    function cleanText(text) {
        return (text || '').replace(/\s+/g, '');
    }

    // 尝试点击导出按钮
    function tryClickExport() {
        console.log('[AutoExport] 正在搜索导出按钮...');

        const candidates = document.querySelectorAll('button, a, div[role="button"], span, div.btn');
        let targetBtn = null;

        for (const el of candidates) {
            if (el.offsetParent === null) continue;

            const txt = el.innerText || el.textContent || '';
            if (txt.includes('同步') || txt.includes('🐞')) continue;

            const clean = cleanText(txt);

            if ((clean === '导出' || (clean.includes('Export') && clean.length < 15)) && !el.disabled) {
                targetBtn = el;
                if (el.tagName.toLowerCase() === 'button') break;
            }
        }

        if (targetBtn) {
            console.log('[AutoExport] 找到导出按钮:', targetBtn);

            const originalBorder = targetBtn.style.border;
            const originalOutline = targetBtn.style.outline;

            targetBtn.style.outline = '4px solid #52c41a';
            targetBtn.style.zIndex = '9999';

            showToast('✅ 找到导出按钮，正在点击...');

            setTimeout(() => {
                targetBtn.click();
                setTimeout(() => {
                    targetBtn.style.border = originalBorder;
                    targetBtn.style.outline = originalOutline;
                }, 1500);

                autoExportTried = true;
                handleConfirmModal();
            }, 800);

            return true;
        } else {
            return false;
        }
    }

    function handleConfirmModal() {
        let attempts = 0;
        const checkModal = setInterval(() => {
            attempts++;
            const modalBtns = document.querySelectorAll('.ant-modal-confirm-btns button, .ant-modal-footer button, .el-message-box__btns button, button.ant-btn-primary');

            const confirmBtn = Array.from(modalBtns).find(b => {
                const txt = cleanText(b.innerText);
                return txt.includes('确') || txt.includes('OK') || txt.includes('是') || txt.includes('意');
            });

            if (confirmBtn) {
                console.log('[AutoExport] 找到确认按钮，点击...');
                confirmBtn.click();
                clearInterval(checkModal);
                showToast('正在自动导出数据 (请耐心等待)...');

                // 延长关闭时间，防止文件还在下载就关闭了
                setTimeout(() => {
                    if (window.opener && !window.opener.closed) {
                        window.close();
                    }
                }, 15000);
            }

            if (attempts > 10) clearInterval(checkModal);
        }, 500);
    }

    function checkPage() {
        if (location.hash.includes('productionScheduling')) {
            if (!document.getElementById('sync-btn')) {
                addSyncButton();
            }
            if (!document.getElementById('debug-btn')) {
                addDebugButton();
            }
        }
    }

    function addDebugButton() {
        const btn = document.createElement('button');
        btn.id = 'debug-btn';
        btn.innerText = '🐞 调试 (V2.4)';
        btn.style.cssText = `position: fixed; top: 60px; right: 220px; z-index: 9999; padding: 4px 10px; background: #faad14; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;`;
        btn.onclick = () => {
            const found = tryClickExport();
            if (!found) {
                alert('V2.4 仍然找不到按钮，请检查页面。\n(已启用强力去空格模式)');
            }
        };
        document.body.appendChild(btn);
    }

    function addSyncButton() {
        // ... previous code ...
        const btn = document.createElement('button');
        btn.id = 'sync-btn';
        btn.innerText = '手动同步到本地';
        btn.style.cssText = `position: fixed; top: 20px; right: 220px; z-index: 9999; padding: 8px 16px; background: #1890ff; color: white; border: none; border-radius: 4px; cursor: pointer; box-shadow: 0 2px 8px rgba(0,0,0,0.2); font-size: 14px; font-weight: bold;`;
        btn.onclick = performSync;
        document.body.appendChild(btn);
    }

    function performSync() {
        showToast('请使用上一页的“自动同步”按钮，效果更好！');
    }

    function showToast(message) {
        let toast = document.getElementById('kimd-toast');
        if (!toast) {
            toast = document.createElement('div');
            toast.id = 'kimd-toast';
            toast.style.cssText = `position: fixed; top: 100px; right: 50%; transform: translateX(50%); padding: 10px 20px; background: rgba(0,0,0,0.8); color: #fff; border-radius: 4px; z-index: 10001; font-size: 16px; pointer-events: none;`;
            document.body.appendChild(toast);
        }
        toast.innerText = message;
        toast.style.display = 'block';
        setTimeout(() => { toast.style.display = 'none'; }, 5000);
    }

    init();
})();
