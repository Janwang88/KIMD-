
const state = {
  rows: [],
  fields: [],
  waitingExcel: false,
  lastAppliedWorkOrder: '',
  currentStats: null,
  allWorkOrders: [] // 缓存的排程工单数据
};

// 获取所有工单排程数据，用于精准匹配工单数量
async function fetchAllWorkOrders() {
  try {
    const res = await fetch('/api/work-orders');
    const data = await res.json();
    if (data.success) {
      state.allWorkOrders = data.data || [];
      console.log(`[Stats] 已从后端同步 ${state.allWorkOrders.length} 条工单基础数据`);
    }
  } catch (e) {
    console.error('获取工单数据失败', e);
  }
}

function setStatus(text, loading = false) {
  const el = document.getElementById('status');
  const dot = document.querySelector('.status-dot');
  const heroStatus = document.getElementById('heroStatusText');

  el.textContent = text;
  if (heroStatus) heroStatus.textContent = text.split(' ')[0]; // 取简短文字

  if (loading) {
    el.parentElement.classList.add('status-loading');
    if (heroStatus) {
      heroStatus.style.color = 'var(--warning)';
      heroStatus.textContent = '统计中...';
    }
  } else {
    el.parentElement.classList.remove('status-loading');
    if (heroStatus) heroStatus.style.color = '#fff';
  }
}

function toggleSearchPanel() {
  const panel = document.getElementById('searchPanel');
  panel.classList.toggle('hidden');
}

function toggleAdvanced() {
  const panel = document.getElementById('hiddenConfig');
  panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
}

function handleEnter(e) {
  if (e.key === 'Enter') loadData();
}

function closeDetail() {
  document.getElementById('detailPanel').classList.add('hidden');
}

function setPageMeta(workOrder, project, woCount) {
  document.getElementById('heroWorkOrder').textContent = workOrder || '-';
  document.getElementById('heroProject').textContent = project || '未命名项目';
  document.getElementById('heroOrderQty').textContent = woCount || '-';

  const woSnippet = (workOrder || '').trim().slice(-7);
  document.title = workOrder ? `✅ ${woSnippet} ${project || ''}` : '物料执行统计';
}

function setSuccessState() {
  // 成功状态提示
}

async function autoSetCookie() {
  const cookie = localStorage.getItem('kimd_cookie');
  if (!cookie) return;
  try {
    await fetch('/api/set-cookie', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ cookie })
    });
  } catch (e) { }
}

function getQueryParam(name) {
  const params = new URLSearchParams(window.location.search);
  return params.get(name) || '';
}

// 从后台获取对应工单的排程时间节点（来自生产排程 Excel）
async function fetchMilestones(workOrder) {
  if (!workOrder) return;
  try {
    const res = await fetch(`/api/milestones?workOrder=${encodeURIComponent(workOrder)}`);
    const data = await res.json();
    if (data.success && data.milestones) {
      document.getElementById('milestoneStrip').style.display = 'flex';
      const ms = data.milestones;
      document.getElementById('msAssemblyStart').textContent = ms.assemblyStart || '-';
      document.getElementById('msAssemblyEnd').textContent = ms.assemblyEnd || '-';
      document.getElementById('msDebugStart').textContent = ms.debugStart || '-';
      document.getElementById('msDebugEnd').textContent = ms.debugEnd || '-';
      document.getElementById('msShipStart').textContent = ms.shipStart || '-';
    }
  } catch (e) {
    console.warn('[Milestones] Fetch failed:', e);
  }
}

function pickField(selectId, keys, prefer) {
  const sel = document.getElementById(selectId);
  sel.innerHTML = '<option value="">请选择字段</option>' + keys.map(k => `<option value="${k}">${k}</option>`).join('');
  const found = keys.find(k => prefer.some(p => k.includes(p) || k === p));
  if (found) sel.value = found;
}

function filterRowsByWorkOrder(rows, workOrder) {
  if (!workOrder) {
    state.lastFilterMode = 'none';
    return rows;
  }
  const keyCandidates = ['工单', '工单号', '工单编号', 'WorkOrder', 'workOrder'];
  const first = rows[0] || {};
  const key = keyCandidates.find(k => k in first) || '工单';
  const normalize = (v) => (v || '').toString().replace(/\s+/g, '').trim();
  const target = normalize(workOrder);
  const exact = rows.filter(r => normalize(r[key]) === target);
  state.lastFilterMode = exact.length ? 'exact' : 'none';
  return exact;
}

function compute() {
  const partField = document.getElementById('partField').value;
  const procPrefix = document.getElementById('procPrefix').value.trim() || '7.';
  const workOrder = document.getElementById('workOrder').value.trim();

  if (!partField) {
    setStatus('请选择料号字段');
    return;
  }

  let stdCount = 0;
  let procCount = 0;
  const filtered = filterRowsByWorkOrder(state.rows, workOrder);
  let emptyCount = 0;

  filtered.forEach(r => {
    const partNo = (r[partField] || '').toString().trim();
    if (!partNo) {
      emptyCount += 1;
      return;
    }

    const isProc = partNo.startsWith(procPrefix);
    if (isProc) {
      procCount += 1;
    } else {
      stdCount += 1;
    }
  });

  // Update basic counts
  document.getElementById('stdRows').textContent = stdCount; // Assuming row count for local compute
  document.getElementById('procRows').textContent = procCount;

  // Local compute doesn't use the full stats logic from server, 
  // primarily we rely on Excel wait for full stats.
  // We'll update just the basics here.

  const total = stdCount + procCount;
  setStatus(`前端预统计：${filtered.length} 行，建议使用 Excel 统计以获取准确交付数据。`);
}

// 手动重试：当文件已下载但等待超时时使用
function retryFetch() {
  const wo = (getQueryParam('workOrder') || document.getElementById('workOrder').value || '').trim();
  if (!wo) { setStatus('请先输入工单号'); return; }
  // 隐藏重试按钮
  const btn = document.getElementById('retryBtn');
  if (btn) btn.style.display = 'none';
  // 重置等待状态，用 10分钟前的时间窗口
  state.waitingExcel = false;
  triggerExcelWait(wo, false, Date.now() - 600000);
}

function loadData() {
  // For this tool, primary flow is triggering backend logic via waiting for Excel
  const wo = document.getElementById('workOrder').value;
  if (!wo) {
    setStatus('请输入工单号');
    return;
  }
  // Force export on manual click
  triggerExcelWait(wo, true);
}

// KIMD Base URL
const BASE_URL = 'https://chajian.kimd.cn:9999';

// State for KIMD Window
let kimdWindow = null;

function openKimdAutoExport(workOrder) {
  if (!workOrder) return;
  // Correct URL found in tampermonkey_material_export.js
  const url = `${BASE_URL}/#/wms/reportManage/materialProgres?auto=true&workOrder=${encodeURIComponent(workOrder)}`;
  kimdWindow = window.open(url, `kimd_export_${workOrder}`);
  setStatus('已请求 KIMD 导出数据，请在新窗口中保持登录...', true);
}

async function waitExcelAndApply(workOrder, opts = {}) {
  const sinceMs = typeof opts.since === 'number' ? opts.since : (Date.now() - 180000);
  setStatus('正在监控最新导出的 Excel 文件 (超时: 120秒)...', true);
  try {
    const res = await fetch('/api/excel-wait-stats', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      // Use multiple patterns to be robust
      body: JSON.stringify({ pattern: ['物料', 'Material', 'Export'], since: sinceMs, timeoutMs: 180000, workOrder })
    });
    const data = await res.json();
    if (!data.success) {
      setStatus(`等待超时，未检测到新Excel。如已下载，点击按钮重新获取`, false);
      // 显示重试按钮
      const retryBtn = document.getElementById('retryBtn');
      if (retryBtn) retryBtn.style.display = 'inline-block';
      return;
    }
    const currentProject = getQueryParam('project') || localStorage.getItem('last_project') || '';

    // Auto-close KIMD window on success before attempting fragile JS renders
    if (kimdWindow) {
      try {
        kimdWindow.close();
        kimdWindow = null;
      } catch (e) { console.error('Auto-close failed', e); }
    }

    applyStatsFromMessage({ workOrder, project: currentProject, stats: data.stats, file: data.savedAs || data.file });

  } catch (e) {
    setStatus(`系统错误：${e.message}`, false);
  } finally {
    state.waitingExcel = false;
  }
}

function triggerExcelWait(workOrder, shouldExport = true, since = null) {
  const wo = (workOrder || '').trim();
  if (!wo) return;

  // Always open KIMD if shouldExport is true, regardless of waiting state
  if (shouldExport) {
    openKimdAutoExport(wo);
  }

  if (state.waitingExcel) return;
  state.waitingExcel = true;

  // Use provided since, or default to 10 seconds ago (safe buffer)
  const sinceMs = typeof since === 'number' && !isNaN(since) ? since : (Date.now() - 10000);
  waitExcelAndApply(wo, { since: sinceMs });
}

function showUndelivered(type) {
  if (!state.currentStats || !state.currentStats.undeliveredList) return;

  const list = state.currentStats.undeliveredList.filter(item => {
    if (type === 'std') return item.type === '标准件';
    if (type === 'proc') return item.type === '加工件';
    return false;
  });

  const titleMap = { 'std': '标准件', 'proc': '加工件' };
  document.getElementById('detailTitle').textContent = `${titleMap[type]}未交货明细 (${list.length})`;

  const tbody = document.getElementById('detailBody');
  if (list.length === 0) {
    tbody.innerHTML = '<tr colspan="6" style="text-align:center; color: var(--text-muted); padding:32px;">🎉 全部已交货</td></tr>';
  } else {
    tbody.innerHTML = list.map((item, idx) => `
          <tr>
            <td>${idx + 1}</td>
            <td style="font-family:monospace; font-weight:500;">${item.partNo}</td>
            <td>${item.name || '-'}</td>
            <td>${item.model || '-'}</td>
            <td style="text-align:right; font-weight:600;">${item.qty || 0}</td>
            <td style="text-align:center; color:${item.purchaseReplyDate ? 'var(--warning)' : 'var(--text-muted)'}; font-weight:${item.purchaseReplyDate ? '600' : '400'};">${item.purchaseReplyDate || '未填写'}</td>
          </tr>
        `).join('');
  }

  document.getElementById('detailPanel').classList.remove('hidden');
  document.getElementById('detailPanel').scrollIntoView({ behavior: 'smooth' });
}

function updateDonutChart(rate, id = 'donutRing') {
  const ring = document.getElementById(id);
  if (!ring) return;
  const rateVal = parseFloat(rate) || 0;
  // dasharray: filled, gap. Circumference is approx 100.
  ring.setAttribute('stroke-dasharray', `${rateVal}, 100`);

  // Color based on rate
  if (rateVal >= 100) ring.setAttribute('stroke', '#10b981'); // Success
  else if (rateVal >= 80) ring.setAttribute('stroke', '#38bdf8'); // Primary (Sky Blue)
  else ring.setAttribute('stroke', '#f59e0b'); // Warning
}

function applyStatsFromMessage(payload) {
  console.log('[Stats] applyStatsFromMessage received:', payload);
  const { workOrder, rows, partField, receiptField, stats, file } = payload || {};
  if (stats && typeof stats === 'object') {
    state.currentStats = stats;
    const heroStatus = document.getElementById('heroStatusText');
    if (heroStatus) {
      heroStatus.textContent = '已完成';
      heroStatus.style.color = '#10b981'; // Success Green
    }

    if (payload && payload.project) localStorage.setItem('last_project', payload.project);

    const statsProject = (stats && stats.projectName) || '';
    const currentProject = statsProject || getQueryParam('project') || localStorage.getItem('last_project') || '';
    if (currentProject) localStorage.setItem('last_project', currentProject);

    // Update Meta
    if (workOrder) {
      setPageMeta(workOrder, currentProject, stats.totalOrderQty);
    }
  }

  // 先进行基础统计数据的渲染，防止页面白屏
  document.getElementById('stdRows').textContent = stats.stdRows ?? '-';
  document.getElementById('stdTotal').textContent = stats.stdTotal ?? '-';
  document.getElementById('procRows').textContent = stats.procRows ?? '-';
  document.getElementById('procTotal').textContent = stats.procTotal ?? '-';

  // 显示未交货：主显为"行数"(采用与周期统计完全一致的原始行级别未交货数量), 旁边带小字备注"款数"(基于归类的唯一料号总量)
  const stdUnRows = stats.cycleStats ? stats.cycleStats.stdUn : 0;
  const procUnRows = stats.cycleStats ? stats.cycleStats.procUn : 0;

  if (stats.stdUndelivered !== undefined) {
    document.getElementById('stdUndelivered').innerHTML = `${stdUnRows} <span style="font-size: 14px; font-weight: normal; color: var(--text-muted);">( ${stats.stdUndelivered} 款)</span>`;
  }

  if (stats.procUndelivered !== undefined) {
    document.getElementById('procUndelivered').innerHTML = `${procUnRows} <span style="font-size: 14px; font-weight: normal; color: var(--text-muted);">( ${stats.procUndelivered} 款)</span>`;
  }

  document.getElementById('pendingIqc').textContent = stats.pendingIqc ?? '-';

  if (stats.stdOnTimeChecked !== undefined) {
    document.getElementById('onTimeOkStd').textContent = stats.stdOnTimeOk ?? '-';
    document.getElementById('onTimeNgStd').textContent = stats.stdOnTimeNg ?? '-';
    const stdRate = stats.stdOnTimeRate !== null ? stats.stdOnTimeRate : 0;
    document.getElementById('onTimeRateStd').textContent = `${stdRate}%`;
    updateDonutChart(stdRate, 'donutRingStd');
  }

  if (stats.procOnTimeChecked !== undefined) {
    document.getElementById('onTimeOkProc').textContent = stats.procOnTimeOk ?? '-';
    document.getElementById('onTimeNgProc').textContent = stats.procOnTimeNg ?? '-';
    const procRate = stats.procOnTimeRate !== null ? stats.procOnTimeRate : 0;
    document.getElementById('onTimeRateProc').textContent = `${procRate}%`;
    updateDonutChart(procRate, 'donutRingProc');
  }
  document.getElementById('pendingIqc').textContent = stats.pendingIqc ?? '-';

  if (stats.cycleStats) {
    const sOk = stats.cycleStats.stdOk || 0;
    const sNg = stats.cycleStats.stdNg || 0;
    const sUn = stats.cycleStats.stdUn || 0;
    const sTotal = sOk + sNg + sUn;
    const sRate = sTotal > 0 ? (((sOk) / sTotal) * 100).toFixed(1) : 0;

    const pOk = stats.cycleStats.procOk || 0;
    const pNg = stats.cycleStats.procNg || 0;
    const pUn = stats.cycleStats.procUn || 0;
    const pTotal = pOk + pNg + pUn;
    const pRate = pTotal > 0 ? (((pOk) / pTotal) * 100).toFixed(1) : 0;

    document.getElementById('stdCycleOk').textContent = sOk;
    document.getElementById('stdCycleNg').textContent = sNg;
    document.getElementById('stdCycleUn').textContent = sUn;
    document.getElementById('stdCycleRate').textContent = `${sRate}%`;

    // UI 上可以给未交货也留个坑，或者通过文字补充说明
    document.getElementById('procCycleOk').textContent = pOk;
    document.getElementById('procCycleNg').textContent = pNg;
    document.getElementById('procCycleUn').textContent = pUn;
    document.getElementById('procCycleRate').textContent = `${pRate}%`;

    updateDonutChart(sRate, 'stdCycleRing');
    updateDonutChart(pRate, 'procCycleRing');
  }



  // --- 异步更新工单数量（G列），不阻塞主统计数据展示 ---
  async function getTrueOrderQty(woStr) {
    try {
      if (!state.allWorkOrders || state.allWorkOrders.length === 0) {
        await fetchAllWorkOrders();
      }

      const targets = (woStr || '').split(/[, \n]/).map(s => s.trim()).filter(Boolean);
      if (targets.length === 0) return 0;

      let total = 0;
      let matchedCount = 0;

      targets.forEach(wo => {
        const match = state.allWorkOrders.find(item => item.workOrderNo === wo || item.taskNo === wo);
        if (match && match.orderQty) {
          total += parseFloat(match.orderQty);
          matchedCount++;
        }
      });
      return matchedCount > 0 ? total : 0;
    } catch (e) {
      console.warn('获取真实工单量失败:', e);
      return 0;
    }
  }

  // 异步更新顶部 Meta 信息及 CSV 外协工时
  (async () => {
    let woCount = await getTrueOrderQty(workOrder);

    if (woCount <= 0) {
      woCount = stats.totalOrderQty || 1;
      if (!stats.totalOrderQty) {
        if (workOrder && workOrder.includes(',')) {
          woCount = workOrder.split(',').filter(Boolean).length;
        } else if (workOrder && workOrder.includes('\n')) {
          woCount = workOrder.split('\n').filter(Boolean).length;
        }
      }
    }
    setPageMeta(workOrder, payload.project || '', woCount);

    // 获取并展示 CSV 外协记录的总工时
    try {
      const res = await fetch(`/api/outsource-hours?workOrder=${encodeURIComponent(workOrder)}`);
      const data = await res.json();
      if (data.success) {
        state.csvOutsourceData = data;
        const outEl = document.getElementById('outsourceCsvTotal');
        if (outEl) {
          outEl.textContent = (data.total || 0).toFixed(1);
        }
        // 当重新加载新的 csv 时，尝试触发刷新一次已经展示出来的工时面板 Icon
        if (state.currentStats && state.currentStats.hoursStats) {
          renderHoursData(state.currentStats.hoursStats);
        }
      }
    } catch (e) {
      console.warn('获取CSV外协工时失败:', e);
    }

    const outEl = document.getElementById('outsourceCsvTotal');
    if (outEl) {
      outEl.parentElement.parentElement.style.display = 'flex';
      outEl.textContent = state.csvOutsourceTotal ? state.csvOutsourceTotal.toFixed(1) : "0.0";
    }

  })();

  const fileNote = file ? ` (文件: ${file})` : '';
  setStatus(`统计完成 ✅ ${fileNote}`, false);
  setSuccessState();

  const woSuffix = (workOrder || '').trim().slice(-7);
  const pjName = (payload.project || '').trim();
  document.title = `✅ ${woSuffix} ${pjName}`;

  if (window.opener) {
    try {
      window.opener.postMessage({
        type: 'STATS_BATCH_COMPLETED',
        workOrder: workOrder,
        project: payload.project,
        success: true
      }, '*');
    } catch (e) { console.error('Post message to opener failed', e); }
  }

  document.getElementById('detailPanel').classList.add('hidden');
  return;
}

// Fallback for raw rows (iframe usage)
if (Array.isArray(rows) && rows.length) {
  state.rows = rows;
  const keys = Object.keys(rows[0] || {});
  state.fields = keys;
  pickField('partField', keys, partField ? [partField] : ['料号', 'ItemNo']);
  pickField('receiptField', keys, receiptField ? [receiptField] : ['收料时间', 'ReceiptTime']);
  compute();
}

// --- Actual Hours Logic ---
function loadHoursData() {
  const wo = document.getElementById('workOrder').value;
  if (!wo) {
    setStatus('请先输入工单号', false);
    return;
  }

  // ★ 关键：在打开KIMD页面«前»记录时间戳，后端只匹配此时刻之后下载的文件
  const sinceMs = Date.now();

  // Open KIMD Hours Page
  const url = `${BASE_URL}/#/sc/work/actualHour?auto=true&workOrder=${encodeURIComponent(wo)}`;
  const win = window.open(url, `kimd_hours_${wo}`);

  setStatus('已请求工时数据，正在等待导出...', true);

  // Start waiting for file，传入精确的起始时间戳
  waitHoursExcelAndApply(wo, win, sinceMs);
}

async function waitHoursExcelAndApply(workOrder, winRef, sinceMs) {
  // ★ 使用调用方传入的精确时间戳，若未传则默认当前时刻（保险用）
  const since = (typeof sinceMs === 'number' && !isNaN(sinceMs)) ? sinceMs : Date.now();
  try {
    const res = await fetch('/api/hours-wait-stats', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        pattern: ['工时', 'Actual', 'Export'],
        timeoutMs: 180000,
        workOrder,
        since  // ★ 传入精确时间给后端，只匹配此刻之后的文件
      })
    });
    const data = await res.json();

    if (!data.success) {
      setStatus(`等待超时，未检测到新导出的Excel。如已下载请点击“重新获取”`, false);
      // 显示重试按钮
      const btn = document.getElementById('retryBtn');
      if (btn) btn.style.display = 'inline-block';
      return;
    }

    renderHoursData(data.stats);
    setStatus(`工时统计完成 ✅`, false);

    // Auto close KIMD window
    if (winRef) {
      try { winRef.close(); } catch (e) { }
    }

  } catch (e) {
    console.error(e);
    setStatus(`工时系统错误：${e.message}`, false);
  }
}

function renderHoursData(stats) {
  if (!stats) return;

  // 将工时数据存入状态，方便核对功能使用
  if (state.currentStats) {
    state.currentStats.hoursStats = stats;
  }

  // Helper to update a section
  const updateSection = (prefix, data) => {
    const total = data.total || 0;
    const plan = data.plan || 0;
    const kimd = data.kimd || 0;
    const out = data.outsource || 0;

    const rate = plan > 0 ? ((total / plan) * 100).toFixed(0) : (total > 0 ? 'Any' : '0');

    document.getElementById(prefix + 'Total').textContent = total.toFixed(1);
    document.getElementById(prefix + 'Plan').textContent = plan.toFixed(1);
    const rateEl = document.getElementById(prefix + 'Rate');
    if (rateEl) {
      rateEl.textContent = rate;
      // Color code rate
      if (plan > 0) {
        if (rate > 100) rateEl.style.color = 'var(--danger)'; // Over budget
        else if (rate > 80) rateEl.style.color = 'orange';  // Warning
        else rateEl.style.color = 'var(--success)'; // Good
      } else {
        rateEl.style.color = '#666';
      }
    }

    // 检查是否有数据差异，并在对应模块展示图标
    const prefixToDbKey = {
      'hoursAssembly': 'assembly',
      'hoursMixed': 'mixed',
      'hoursWiring': 'wiring'
    };
    const dbKey = prefixToDbKey[prefix];

    // 更新图标
    if (dbKey) {
      const iconEl = document.getElementById(prefix + 'StatusIcon');
      if (iconEl) {
        if (state.csvOutsourceData !== undefined) {
          const dbVal = state.csvOutsourceData[dbKey] || 0;
          const diff = Math.abs(dbVal - out);
          if (diff < 0.1) {
            iconEl.innerHTML = '✅';
            iconEl.style.animation = 'none';
          } else {
            iconEl.innerHTML = '⚠️';
            iconEl.style.color = 'var(--danger)';
            iconEl.style.animation = 'blink 1s infinite';
          }
        } else {
          iconEl.innerHTML = '';
        }
      }
    }

    // 渲染各个工艺的列表
    const detailsContainer = document.getElementById(prefix + 'Details');
    if (detailsContainer) {
      if ((data.processes || []).length > 0) {
        let html = `
          <table style="width: 100%; border-collapse: separate; border-spacing: 0 6px; font-size: 13px;">
            <thead>
              <tr>
                <th rowspan="2" style="border-bottom: 2px solid #e2e8f0; padding-bottom: 6px; width: 24%;"></th>
                <th colspan="2" style="text-align: center; font-weight: bold; color: #333; padding-bottom: 6px; width: 38%;">系统数据</th>
                <th colspan="2" style="text-align: center; font-weight: bold; color: #333; padding-bottom: 6px; width: 38%;">本地数据</th>
              </tr>
              <tr style="color: #475569; font-size: 11px;">
                <th style="font-weight: bold; text-align: center; padding: 4px 6px;">KIMD</th>
                <th style="font-weight: bold; text-align: center; padding: 4px 6px;">外协</th>
                <th style="font-weight: bold; text-align: center; padding: 4px 6px;">KIMD</th>
                <th style="font-weight: bold; text-align: center; padding: 4px 6px;">外协</th>
              </tr>
            </thead>
            <tbody>
        `;

        (data.processes || []).forEach(proc => {
          const kimdObj = (data.processBreakdown && data.processBreakdown[proc]) ? data.processBreakdown[proc] : { kimd: 0, outsource: 0 };
          const kimdVal = typeof kimdObj === 'number' ? 0 : (kimdObj.kimd || 0);
          const outsourceVal = typeof kimdObj === 'number' ? kimdObj : (kimdObj.outsource || 0);

          let dbValNum = 0;
          if (state.csvOutsourceData && state.csvOutsourceData.processBreakdown) {
            dbValNum = state.csvOutsourceData.processBreakdown[proc] || 0;
          }

          let dbKimdNum = 0;
          if (state.csvOutsourceData && state.csvOutsourceData.kimdBreakdown) {
            dbKimdNum = state.csvOutsourceData.kimdBreakdown[proc] || 0;
          }

          const formatVal = (val) => (val === 0 || Math.abs(val) < 0.01) ? '-' : val.toFixed(1);

          // 系统KIMD 与 本地KIMD 比较，有差异→红，一致→黑
          const kimdDiff = Math.abs(kimdVal - dbKimdNum);
          const kimdSysColor = kimdDiff >= 0.1 ? '#ef4444' : '#1e293b';

          // 系统外协 与 本地外协 比较，有差异→红，一致→黑
          const outsourceDiff = Math.abs(outsourceVal - dbValNum);
          const sysColor = outsourceDiff >= 0.1 ? '#ef4444' : '#1e293b';

          const localColor = '#1e293b';

          html += `
            <tr>
              <td style="padding: 6px 8px; font-weight: 700; color: #1e293b; text-align: left; background-color: #f1f5f9; border-radius: 4px; white-space: nowrap;">${proc}</td>
              <td style="padding: 6px 4px; text-align: center; font-weight: 700; color: ${kimdSysColor};">${formatVal(kimdVal)}</td>
              <td style="padding: 6px 4px; text-align: center; font-weight: 700; color: ${sysColor};">${formatVal(outsourceVal)}</td>
              <td style="padding: 6px 4px; text-align: center; font-weight: 700; color: ${localColor};">${formatVal(dbKimdNum)}</td>
              <td style="padding: 6px 4px; text-align: center; font-weight: 700; color: ${localColor};">${formatVal(dbValNum)}</td>
            </tr>
          `;
        });

        html += `
            </tbody>
          </table>
        `;
        detailsContainer.innerHTML = html;
      } else {
        detailsContainer.innerHTML = '';
      }
    }
  };

  updateSection('hoursAssembly', stats.assembly || {});
  updateSection('hoursMixed', stats.mixed || {});
  updateSection('hoursWiring', stats.wiring || {});
}

function openHoursComparison(section = 'assembly') {
  const dash = document.getElementById('hoursDashboard');
  const view = document.getElementById('hoursComparisonView');
  const content = document.getElementById('comparisonContent');

  if (!state.currentStats || !state.currentStats.hoursStats) {
    alert('请先加载工时数据');
    return;
  }

  // 1. 改变布局：保留左侧为当前 section 卡片，右侧为大详情板
  dash.style.display = 'grid';
  dash.style.gridTemplateColumns = '320px 1fr'; // 左边固定卡片大小，右边撑满剩余

  // 隐藏其他未选中的卡片，并将选中的卡片显示出来
  const sections = ['hoursAssemblySection', 'hoursMixedSection', 'hoursWiringSection'];
  sections.forEach(secId => {
    const el = document.getElementById(secId);
    if (!el) return;
    if (secId.toLowerCase().includes(section.toLowerCase())) {
      el.style.display = 'block';
    } else {
      el.style.display = 'none';
    }
  });

  // 把对比视图移动到 dash 容器里进行并排显示（原逻辑在 dash 下面）
  view.style.display = 'block';
  // 使 view 不再有上边距（因为它现在作为 grid 的右边格）
  view.style.marginTop = '0';
  dash.appendChild(view); // 将对比视图塞入 dashboard 形成第二列

  // 2. 获取所属大类的子工艺列表及总计对比
  const RULES = {
    'assembly': ['组装-返工', '模组组装', '整机接气', '出货'],
    'mixed': ['项目管理', '领料', '上线准备', '总装', '清洁', '打包'],
    'wiring': ['接线-返工', '电控配线', '整机接线', '通电通气']
  };
  const processes = RULES[section] || [];

  // 系统统计中该大类的总金额
  const kimdOutsourceTotal = state.currentStats.hoursStats[section] ? state.currentStats.hoursStats[section].outsource : 0;

  // 本地统计中该大类的总金额
  const dbData = state.csvOutsourceData || {};
  const dbTotal = dbData[section] || 0;
  const diffTotal = (dbTotal - kimdOutsourceTotal).toFixed(1);

  const titleMap = {
    'assembly': { name: '组装', color: '#ef4444' },
    'mixed': { name: '混合', color: '#3b82f6' },
    'wiring': { name: '接线', color: '#10b981' }
  };
  const sectionInfo = titleMap[section] || { name: '未知', color: '#cbd5e1' };

  document.getElementById('comparisonTitle').innerHTML = `
    <div style="display:flex; align-items:center; gap:12px; height:32px; width:100%; justify-content:space-between;">
      <div style="display:flex; align-items:center; gap:12px;">
        <span style="white-space: nowrap;">🔍 详细核对：${sectionInfo.name}工艺</span>
        <div style="display:flex; gap:8px;">
          <button onclick="changeComparison('assembly')" style="border:1px solid #ef4444; background:${section === 'assembly' ? '#fee2e2' : '#fff'}; color:#ef4444; padding:2px 8px; border-radius:4px; font-size:12px; cursor:pointer; white-space: nowrap;">组装</button>
          <button onclick="changeComparison('mixed')" style="border:1px solid #3b82f6; background:${section === 'mixed' ? '#dbeafe' : '#fff'}; color:#3b82f6; padding:2px 8px; border-radius:4px; font-size:12px; cursor:pointer; white-space: nowrap;">混合</button>
          <button onclick="changeComparison('wiring')" style="border:1px solid #10b981; background:${section === 'wiring' ? '#dcfce7' : '#fff'}; color:#10b981; padding:2px 8px; border-radius:4px; font-size:12px; cursor:pointer; white-space: nowrap;">接线</button>
        </div>
      </div>
      <div style="font-size:13px; font-weight:700; white-space: nowrap;">
        <span style="color:#64748b; margin-right:12px;">系统外协总计: <span style="color:#1e293b; font-size:15px;">${kimdOutsourceTotal.toFixed(1)}</span></span>
        <span style="color:#64748b; margin-right:12px;">本地外协总计: <span style="color:#1e293b; font-size:15px;">${dbTotal.toFixed(1)}</span></span>
        <span style="color:#64748b;">差异: <span style="color:${Math.abs(diffTotal) < 0.1 ? '#10b981' : '#ef4444'}; font-size:15px;">${diffTotal > 0 ? '+' : ''}${diffTotal}</span></span>
      </div>
    </div>
  `;

  // 3. 构建超级详细的对比 HTML，按子工艺行遍历
  let tbodyHtml = '';

  processes.forEach(proc => {
    // 系统 KIMD 和 外协工时 (来自网页/Excel 的各个工艺工时)
    let sysKimdStr = '-';
    let sysOutsourceStr = '-';
    let sysKimdVal = 0;
    let sysOutsourceVal = 0;

    // 正确的嵌套逻辑是 state.currentStats.hoursStats[section].processBreakdown[proc] = {kimd: 0, outsource: 0}
    if (state.currentStats.hoursStats[section] && state.currentStats.hoursStats[section].processBreakdown) {
      const pb = state.currentStats.hoursStats[section].processBreakdown[proc];
      if (pb) {
        sysKimdVal = pb.kimd || 0;
        sysKimdStr = sysKimdVal > 0 ? sysKimdVal.toFixed(1) : '-';

        sysOutsourceVal = pb.outsource || 0;
        sysOutsourceStr = sysOutsourceVal > 0 ? sysOutsourceVal.toFixed(1) : '-';
      }
    }

    // 本地拆解数据
    let localKimdStr = '-';
    let bigStr = '-';
    let midStr = '-';
    let smallStr = '-';
    let localTotalStr = '-';
    let localTotalVal = 0;

    if (dbData.detailedProcessBreakdown && dbData.detailedProcessBreakdown[proc]) {
      const d = dbData.detailedProcessBreakdown[proc];

      if (d.kimd > 0) localKimdStr = d.kimd.toFixed(1);
      if (d.outsource['大工'] > 0) bigStr = d.outsource['大工'].toFixed(1);
      if (d.outsource['中工'] > 0) midStr = d.outsource['中工'].toFixed(1);
      if (d.outsource['小工'] > 0) smallStr = d.outsource['小工'].toFixed(1);
      if (d.outsource.total > 0) localTotalStr = d.outsource.total.toFixed(1);

      localTotalVal = d.outsource.total || 0;
    }

    const rowDiff = (localTotalVal - sysOutsourceVal).toFixed(1);
    const diffColor = Math.abs(localTotalVal - sysOutsourceVal) < 0.1 ? 'var(--success)' : 'var(--danger)';
    const diffStr = rowDiff === '0.0' ? '-' : (rowDiff > 0 ? `+${rowDiff}` : rowDiff);

    tbodyHtml += `
      <tr>
        <td style="padding:10px 16px; font-weight:700; color:#1e293b; border-bottom:1px solid #f1f5f9; white-space:nowrap;">${proc}</td>
        <td style="padding:10px 16px; text-align:center; font-weight:800; color:#059669; border-bottom:1px solid #f1f5f9; border-left:1px dashed #e2e8f0; background:#f0fdf4;">${sysKimdStr}</td>
        <td style="padding:10px 16px; text-align:center; font-weight:800; color:#475569; border-bottom:1px solid #f1f5f9; border-right:1px dashed #e2e8f0; background:#f8fafc;">${sysOutsourceStr}</td>
        <td style="padding:10px 16px; text-align:center; font-weight:600; color:#10b981; border-bottom:1px solid #f1f5f9; background:#fbfeff;">${localKimdStr}</td>
        <td style="padding:10px 16px; text-align:center; font-weight:600; color:#0284c7; border-bottom:1px solid #f1f5f9; background:#f0f9ff border-left:1px dashed #f1f5f9;">${bigStr}</td>
        <td style="padding:10px 16px; text-align:center; font-weight:600; color:#0284c7; border-bottom:1px solid #f1f5f9; background:#f0f9ff border-left:1px dashed #f1f5f9;">${midStr}</td>
        <td style="padding:10px 16px; text-align:center; font-weight:600; color:#0284c7; border-bottom:1px solid #f1f5f9; background:#f0f9ff border-left:1px dashed #f1f5f9;">${smallStr}</td>
        <td style="padding:10px 16px; text-align:center; font-size:16px; font-weight:800; color:#0c4a6e; border-bottom:1px solid #f1f5f9; background:#e0f2fe; border-left:1px dashed #bae6fd;">${localTotalStr}</td>
        <td style="padding:10px 16px; text-align:center; font-weight:800; color:${diffColor}; border-bottom:1px solid #f1f5f9; border-left:1px dashed #e2e8f0;">${diffStr}</td>
      </tr>
    `;
  });

  content.innerHTML = `
    <table class="modern-table" style="margin-bottom:0; border-spacing:0; width:100%;">
      <thead>
        <tr style="background:#f8fafc;">
          <th rowspan="2" style="padding:10px 16px; border-bottom:1px solid #cbd5e1; width: 140px;">子工艺环节</th>
          <th colspan="2" style="padding:10px 16px; text-align:center; border-left:1px dashed #cbd5e1; border-right:1px dashed #cbd5e1; border-bottom:1px solid #cbd5e1; background:#f1f5f9; min-width:120px; white-space:nowrap;">
            系统数据 (总计 h)
          </th>
          <th rowspan="2" style="padding:10px 16px; text-align:center; background:#ebf8ff; border-bottom:1px dashed #bae6fd; white-space:nowrap;">本地<br>KIMD</th>
          <th colspan="4" style="padding:10px 16px; text-align:center; background:#e0f2fe; border-bottom:1px solid #bae6fd;">本地外协细分</th>
          <th rowspan="2" style="padding:10px 16px; text-align:center; border-left:1px dashed #cbd5e1; border-bottom:1px solid #cbd5e1; min-width:80px; white-space:nowrap;">外协差异<br><span style="font-size:11px; font-weight:normal;">(本地-系统)</span></th>
        </tr>
        <tr style="background:#f1f5f9; font-size:12px;">
          <th style="padding:6px 12px; text-align:center; font-weight:600; color:#059669; border-bottom:1px solid #cbd5e1; white-space:nowrap;">KIMD</th>
          <th style="padding:6px 12px; text-align:center; font-weight:600; color:#475569; border-bottom:1px solid #cbd5e1; white-space:nowrap;">外协</th>
          <th style="padding:6px 12px; text-align:center; font-weight:600; color:#0369a1; border-bottom:1px solid #cbd5e1; white-space:nowrap;">大工</th>
          <th style="padding:6px 12px; text-align:center; font-weight:600; color:#0369a1; border-bottom:1px solid #cbd5e1; white-space:nowrap;">中工</th>
          <th style="padding:6px 12px; text-align:center; font-weight:600; color:#0369a1; border-bottom:1px solid #cbd5e1; white-space:nowrap;">小工</th>
          <th style="padding:6px 12px; text-align:center; font-weight:800; color:#0c4a6e; background:#e0f2fe; border-bottom:1px solid #cbd5e1; border-left:1px dashed #bae6fd; white-space:nowrap;">合计</th>
        </tr>
      </thead>
      <tbody>
        ${tbodyHtml}
      </tbody>
    </table>
    <div style="padding:12px 20px; background:#fffaf0; border-top:1px solid #fbd38d; font-size:13px; color:#744210; display:flex; justify-content:space-between; align-items:center;">
      <div>
        <strong>💡 提示：</strong>本地的 KIMD 是专门供参考打卡的台账项，并未加总到外协合计里，这能帮您比对哪些行未录入系统中。差异列对应的是（本地外协合计 - 系统外协数据）。
      </div>
      <button class="btn-primary" onclick="closeComparison()" style="height:32px; font-size:13px; padding:0 24px; white-space: nowrap;">返回整体概览</button>
    </div>
  `;
}

// 供标签页直接切换的内部入口
window.changeComparison = function (section) {
  openHoursComparison(section);
}

function closeComparison() {
  const dash = document.getElementById('hoursDashboard');
  const view = document.getElementById('hoursComparisonView');

  // 将 view 从里边抽离出来放到后面 (因为之前 appendChild 放到了 grid 里)
  dash.parentNode.insertBefore(view, dash.nextSibling);

  // 恢复三分栏布局
  dash.style.gridTemplateColumns = 'repeat(3, 1fr)';

  // 恢复所有卡片的显示
  const sections = ['hoursAssemblySection', 'hoursMixedSection', 'hoursWiringSection'];
  sections.forEach(secId => {
    const el = document.getElementById(secId);
    if (el) el.style.display = 'block';
  });

  dash.style.display = 'grid';
  view.style.display = 'none';
  view.style.marginTop = '24px';
}

// Init
// --- 极简稳健初始化 ---
function bootstrap() {
  console.log('[Stats] Bootstrap started...');
  try {
    const params = new URLSearchParams(window.location.search);
    const woParam = params.get('workOrder');
    const sinceParam = params.get('since');

    console.log('[Stats] Detect workOrder:', woParam);

    const el = document.getElementById('heroWorkOrderInput') || document.getElementById('workOrder');
    if (woParam && el) {
      el.value = woParam;
      console.log('[Stats] Value set to input.');
    }

    if (typeof triggerExcelWait === 'function' && woParam) {
      const autoPopup = params.get('auto') === 'true';
      // 增加时间偏移补偿：如果 URL 提供了 since，额外提前 2 秒（2000ms），
      // 确保在页面刷新跳转期间下载的文件也能被监控捕获。
      const sinceVal = sinceParam ? (parseInt(sinceParam, 10) - 2000) : (Date.now() - 30000);
      console.log(`[Stats] Running triggerExcelWait with wo=${woParam}, since=${sinceVal}`);
      triggerExcelWait(woParam, autoPopup, sinceVal);
    } else if (!woParam) {
      const last = localStorage.getItem('last_work_order');
      if (last && el) el.value = last;
    }

    // 异步背景
    autoSetCookie();
    fetchAllWorkOrders();
    fetchMilestones(woParam || localStorage.getItem('last_work_order'));

    console.log('[Stats] Bootstrap success.');
  } catch (err) {
    console.error('[Stats] Bootstrap error:', err);
  }
}

// 强力轮询启动：最多尝试 10 次 (总计约 3s)，确保 DOM 彻底稳定且 ID 元素已挂载
let bootAttempts = 0;
const bootInterval = setInterval(() => {
  bootAttempts++;
  console.log(`[Stats] Boot attempt #${bootAttempts}...`);

  const el = document.getElementById('workOrder');
  if (el || bootAttempts > 10) {
    clearInterval(bootInterval);
    bootstrap();
  }
}, 200);

// 消息监听
window.addEventListener('message', (e) => {
  if (!e.data) return;
  if (e.data.type === 'STATS_DATA') applyStatsFromMessage(e.data.payload);
  else if (e.data.type === 'kimd-excel-wait') {
    const wo = e.data.payload.workOrder;
    if (wo && typeof triggerExcelWait === 'function') triggerExcelWait(wo, false);
  }
});

window.triggerHeroSearch = function () {
  const input = document.getElementById('workOrder');
  if (!input) return;
  const val = input.value.trim();
  if (!val) {
    alert("请输入工单号");
    return;
  }

  // update URL so reload keeps it
  const newUrl = window.location.protocol + "//" + window.location.host + window.location.pathname + '?workOrder=' + encodeURIComponent(val) + '&auto=true&since=' + Date.now();
  window.location.href = newUrl; // 此处直接强刷新跳转最稳，因为原来的触发可能需要页面刷新
}

// Also hook up Enter key
window.handleEnter = function (e) {
  if (e.key === 'Enter') {
    window.triggerHeroSearch();
  }
}

document.addEventListener('DOMContentLoaded', () => {
  const btn = document.querySelector('.btn-tool');
  if (btn) {
    btn.onclick = window.triggerHeroSearch;
  }
});

// ensure explicit binding regardless of html
document.addEventListener('DOMContentLoaded', () => {
  const btn = document.querySelector('button[onclick="window.triggerHeroSearch()"]');
  if (btn) {
    btn.onclick = window.triggerHeroSearch;
    btn.addEventListener('click', window.triggerHeroSearch, true);
  }
});

// Just in case, try attaching via ID
document.addEventListener('DOMContentLoaded', () => {
  const btnContainer = document.querySelector('.hero-dashboard .nav-tools');
  if (btnContainer) {
    btnContainer.onclick = function (e) {
      if (e.target.tagName === 'BUTTON' || e.target.closest('button')) {
        window.triggerHeroSearch();
      }
    };
  }
});
