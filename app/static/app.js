'use strict';

// ── 状态 ──────────────────────────────────────────────────────
let state = {
  fileId:     null,
  pdfName:    null,
  totalPages: 0,
  records:    [],   // 当前表格行数据
};

// ── DOM 引用 ──────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const dropZone      = $('drop-zone');
const fileInput     = $('file-input');
const secUpload     = $('section-upload');
const secProgress   = $('section-progress');
const secResult     = $('section-result');
const secDone       = $('section-done');
const progressBar   = $('progress-bar');
const progressText  = $('progress-text');
const progressSrc   = $('progress-source');
const progressLog   = $('progress-log');
const resultMeta    = $('result-meta');
const warningsBox   = $('warnings-box');
const tbody         = $('result-tbody');
const doneSummary   = $('done-summary');
const doneTbody     = $('done-tbody');

// ── 上传区交互 ────────────────────────────────────────────────
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) startDetect(file);
});

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) startDetect(fileInput.files[0]);
});

// ── 识别流程 ──────────────────────────────────────────────────
function startDetect(file) {
  if (!file.name.toLowerCase().endsWith('.pdf')) {
    showToast('请选择 PDF 文件');
    return;
  }

  // 重置状态
  state = { fileId: null, pdfName: file.name, totalPages: 0, records: [] };
  progressLog.innerHTML = '';
  progressBar.style.width = '0%';
  progressText.textContent = '正在准备…';
  progressSrc.textContent = '';

  show(secProgress);
  hide(secUpload, secResult, secDone);

  const formData = new FormData();
  formData.append('pdf', file);
  if ($('opt-no-qr').checked)   formData.append('no_qr', '1');
  if ($('opt-hires').checked)   formData.append('dpi', '300');

  const evtSource = new EventSource('/api/detect');   // placeholder; real: fetch + ReadableStream

  // 改用 fetch + ReadableStream 解析 SSE（EventSource 不支持 POST）
  evtSource.close();
  fetchSSE('/api/detect', formData);
}

function fetchSSE(url, formData) {
  fetch(url, { method: 'POST', body: formData })
    .then(res => {
      if (!res.ok) throw new Error(`服务器错误 ${res.status}`);
      const reader = res.body.getReader();
      const decoder = new TextDecoder();
      let buf = '';

      function pump() {
        return reader.read().then(({ done, value }) => {
          if (done) return;
          buf += decoder.decode(value, { stream: true });
          const lines = buf.split('\n');
          buf = lines.pop();
          for (const line of lines) {
            if (line.startsWith('data: ')) {
              try { handleSSE(JSON.parse(line.slice(6))); } catch {}
            }
          }
          return pump();
        });
      }
      return pump();
    })
    .catch(err => {
      showToast('识别失败：' + err.message);
      show(secUpload);
      hide(secProgress);
    });
}

function handleSSE(msg) {
  switch (msg.type) {
    case 'start':
      state.fileId     = msg.file_id;
      state.pdfName    = msg.pdf_name;
      state.totalPages = msg.total;
      progressText.textContent = `正在扫描第 0 / ${msg.total} 页`;
      break;

    case 'info':
      appendLog('ℹ ' + msg.message);
      break;

    case 'warn':
      appendLog('⚠ ' + msg.message);
      break;

    case 'page': {
      const pct = Math.round((msg.page / state.totalPages) * 100);
      progressBar.style.width = pct + '%';
      progressText.textContent = `正在扫描第 ${msg.page} / ${state.totalPages} 页`;

      if (msg.source === 'qr') {
        progressSrc.textContent = '🔵 二维码';
        appendLog(`第 ${msg.page} 页  🔵 QR  批号：${msg.batch_no}`);
      } else if (msg.source === 'ocr' && msg.conf) {
        const icon = msg.conf === '高' ? '🟢' : '🟡';
        progressSrc.textContent = `${icon} OCR ${msg.conf}置信`;
        if (msg.batch_no) appendLog(`第 ${msg.page} 页  ${icon} OCR  批号：${msg.batch_no}`);
      }
      break;
    }

    case 'done':
      state.fileId     = msg.file_id;
      state.records    = msg.records;
      state.totalPages = msg.total_pages;
      renderResult(msg);
      break;

    case 'error':
      showToast('错误：' + msg.message);
      show(secUpload);
      hide(secProgress);
      break;
  }
}

// ── 渲染结果表格 ──────────────────────────────────────────────
function renderResult(msg) {
  hide(secProgress);
  show(secResult);

  const qrCount  = msg.records.filter(r => r.source === 'qr').length;
  const ocrCount = msg.records.filter(r => r.source === 'ocr').length;
  const manCount = msg.records.filter(r => r.source === 'manual').length;
  const parts = [`共识别 ${msg.records.length} 份证明，PDF 共 ${msg.total_pages} 页`];
  if (qrCount)  parts.push(`🔵 QR ${qrCount}`);
  if (ocrCount) parts.push(`🟢 OCR ${ocrCount}`);
  if (manCount) parts.push(`🔴 手动 ${manCount}`);
  resultMeta.textContent = parts.join('　');

  // 警告
  if (msg.warnings && msg.warnings.length) {
    warningsBox.textContent = msg.warnings.join('\n');
    warningsBox.classList.add('show');
  } else {
    warningsBox.classList.remove('show');
  }

  tbody.innerHTML = '';
  state.records.forEach((r, idx) => addRow(r, idx));
}

function rowClass(r) {
  if (r.source === 'qr')     return 'row-qr';
  if (r.conf === '高')        return 'row-high';
  if (r.conf === '中')        return 'row-mid';
  return 'row-manual';
}

const SOURCE_BADGE = {
  qr:     '<span class="src-badge src-qr">🔵 QR</span>',
  ocr:    null,   // 由 conf 决定
  manual: '<span class="src-badge src-manual">🔴 手动</span>',
};

function sourceBadge(r) {
  if (r.source === 'qr')     return '<span class="src-badge src-qr">🔵 QR</span>';
  if (r.source === 'manual') return '<span class="src-badge src-manual">🔴 手动</span>';
  if (r.conf === '高')        return '<span class="src-badge src-ocr-high">🟢 OCR</span>';
  if (r.conf === '中')        return '<span class="src-badge src-ocr-mid">🟡 OCR?</span>';
  return '<span class="src-badge src-manual">🔴 手动</span>';
}

function addRow(r, idx) {
  const tr = document.createElement('tr');
  tr.dataset.idx = idx;
  tr.className   = rowClass(r);

  const pageCount = (r.end_page && r.start_page)
    ? r.end_page - r.start_page + 1 : '—';
  const abnormal = typeof pageCount === 'number' && (pageCount < 2 || pageCount > 4);

  tr.innerHTML = `
    <td style="text-align:center;color:#9ca3af">${idx + 1}</td>
    <td style="text-align:center">${sourceBadge(r)}</td>
    <td><input class="cell-input" data-field="batch_no"     value="${esc(r.batch_no || '')}"></td>
    <td><input class="cell-input" data-field="vaccine_name" value="${esc(r.vaccine_name || '')}"></td>
    <td><input class="cell-input" data-field="manufacturer" value="${esc(r.manufacturer || '')}"></td>
    <td><input class="cell-input" data-field="cert_no"      value="${esc(r.cert_no || '')}"></td>
    <td><input class="cell-input page-input" data-field="start_page" value="${r.start_page || ''}"></td>
    <td><input class="cell-input page-input" data-field="end_page"   value="${r.end_page || ''}"></td>
    <td class="page-count${abnormal ? ' abnormal' : ''}">${pageCount}</td>
    <td><button class="btn-del" title="删除此行">✕</button></td>
  `;

  // 页码变化时实时更新页数
  const startInput = tr.querySelector('[data-field="start_page"]');
  const endInput   = tr.querySelector('[data-field="end_page"]');
  const countCell  = tr.querySelector('.page-count');
  function updateCount() {
    const s = parseInt(startInput.value), e = parseInt(endInput.value);
    if (s && e && e >= s) {
      const c = e - s + 1;
      countCell.textContent = c;
      countCell.className = 'page-count' + (c < 2 || c > 4 ? ' abnormal' : '');
    } else {
      countCell.textContent = '—';
    }
    syncRowData(tr);
  }
  startInput.addEventListener('input', updateCount);
  endInput.addEventListener('input',   updateCount);

  // 其他字段变化时同步
  tr.querySelectorAll('.cell-input:not(.page-input)').forEach(input => {
    input.addEventListener('input', () => syncRowData(tr));
  });

  // 删除行
  tr.querySelector('.btn-del').addEventListener('click', () => {
    tr.remove();
    reindexRows();
  });

  tbody.appendChild(tr);
}

function syncRowData(tr) {
  const idx = parseInt(tr.dataset.idx);
  tr.querySelectorAll('.cell-input').forEach(input => {
    const field = input.dataset.field;
    let val = input.value.trim();
    if (field === 'start_page' || field === 'end_page') {
      state.records[idx][field] = val ? parseInt(val) : null;
    } else {
      state.records[idx][field] = val || null;
    }
  });
}

function reindexRows() {
  const rows = tbody.querySelectorAll('tr');
  rows.forEach((tr, i) => {
    tr.dataset.idx = i;
    tr.querySelector('td:first-child').textContent = i + 1;
  });
  state.records = Array.from(rows).map(tr => {
    const r = {};
    tr.querySelectorAll('.cell-input').forEach(inp => {
      const f = inp.dataset.field;
      r[f] = (f === 'start_page' || f === 'end_page')
        ? (inp.value ? parseInt(inp.value) : null)
        : (inp.value.trim() || null);
    });
    return r;
  });
}

// 添加空行
$('btn-add-row').addEventListener('click', () => {
  const r = { batch_no: '', vaccine_name: '', manufacturer: '',
               cert_no: '', start_page: null, end_page: null, conf: '' };
  state.records.push(r);
  addRow(r, state.records.length - 1);
});

// ── 拆分 ─────────────────────────────────────────────────────
$('btn-split').addEventListener('click', () => {
  // 收集最新表格数据
  reindexRows();

  // 简单前端校验
  let hasError = false;
  tbody.querySelectorAll('tr').forEach((tr, i) => {
    const r = state.records[i];
    if (!r || !r.batch_no) {
      tr.querySelector('[data-field="batch_no"]').style.borderColor = '#dc2626';
      hasError = true;
    }
    if (!r || !r.start_page || !r.end_page) {
      hasError = true;
    }
  });
  if (hasError) {
    showToast('请填写所有红色标注的必填项（批号、起始页、结束页）');
    return;
  }

  const btn = $('btn-split');
  btn.disabled    = true;
  btn.textContent = '拆分中…';

  doSplit(false);

  function doSplit(force) {
    fetch('/api/split', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file_id:  state.fileId,
        pdf_name: state.pdfName,
        records:  state.records,
        force:    force,
      }),
    })
      .then(res => res.json())
      .then(data => {
        if (data.already_split) {
          btn.disabled    = false;
          btn.textContent = '确认并拆分 ▶';
          if (confirm(`该文件已拆分过 ${data.prev_count} 份证明，是否覆盖重新拆分？`)) {
            btn.disabled    = true;
            btn.textContent = '拆分中…';
            doSplit(true);
          }
          return;
        }
        btn.disabled    = false;
        btn.textContent = '确认并拆分 ▶';
        if (data.error) {
          showToast('拆分失败：' + data.error);
          return;
        }
        renderDone(data);
      })
      .catch(err => {
        btn.disabled    = false;
        btn.textContent = '确认并拆分 ▶';
        showToast('请求失败：' + err.message);
      });
  }
});

// ── 渲染完成页 ────────────────────────────────────────────────
function renderDone(data) {
  hide(secResult);
  show(secDone);

  doneSummary.innerHTML =
    `共拆分 <strong>${data.count}</strong> 份证明` +
    (data.abnormal ? `，其中 <strong style="color:#b45309">${data.abnormal} 份</strong> 页数异常，请人工核查` : '，全部正常') + '。';

  const dl = $('btn-download');
  dl.href     = `/api/download/${state.fileId}`;
  dl.download = '';

  const xl = $('btn-excel');
  xl.href     = `/api/export_excel/${state.fileId}`;
  xl.download = '';

  doneTbody.innerHTML = '';
  (data.results || []).forEach(r => {
    const tr = document.createElement('tr');
    const ok = r.review_status === 'ok';
    tr.innerHTML = `
      <td>${esc(r.batch_no)}</td>
      <td>${r.pages}</td>
      <td>${r.page_count}</td>
      <td class="${ok ? 'status-ok' : 'status-warn'}">${ok ? '✓ 正常' : '⚠ 待复核'}</td>
    `;
    doneTbody.appendChild(tr);
  });
}

// ── 重新开始 ──────────────────────────────────────────────────
$('btn-restart').addEventListener('click', () => {
  state = { fileId: null, pdfName: null, totalPages: 0, records: [] };
  fileInput.value = '';
  tbody.innerHTML = '';
  doneTbody.innerHTML = '';
  progressLog.innerHTML = '';
  warningsBox.classList.remove('show');
  hide(secProgress, secResult, secDone);
  show(secUpload);
});

// ── 历史记录 ──────────────────────────────────────────────────
$('btn-history').addEventListener('click', () => {
  loadHistory();
  show($('drawer-overlay'), $('drawer-history'));
});
$('btn-close-history').addEventListener('click', closeHistory);
$('drawer-overlay').addEventListener('click', closeHistory);

function closeHistory() {
  hide($('drawer-overlay'), $('drawer-history'));
}

function loadHistory() {
  const htbody = $('history-tbody');
  htbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#9ca3af">加载中…</td></tr>';
  fetch('/api/records')
    .then(r => r.json())
    .then(rows => {
      if (!rows.length) {
        htbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#9ca3af">暂无记录</td></tr>';
        return;
      }
      htbody.innerHTML = rows.map(r => `
        <tr>
          <td>${esc(r.batch_no)}</td>
          <td>${esc(r.vaccine_name || '—')}</td>
          <td>${esc(r.cert_no || '—')}</td>
          <td>${r.page_count}</td>
          <td class="${r.review_status === 'ok' ? 'status-ok' : 'status-warn'}">
            ${r.review_status === 'ok' ? '正常' : '待复核'}
          </td>
          <td style="color:#9ca3af">${(r.split_time || '').slice(0, 10)}</td>
        </tr>
      `).join('');
    })
    .catch(() => {
      htbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#dc2626">加载失败</td></tr>';
    });
}

// ── 工具函数 ──────────────────────────────────────────────────
function show(...els)  { els.forEach(el => el && el.classList.remove('hidden')); }
function hide(...els)  { els.forEach(el => el && el.classList.add('hidden')); }
function esc(str)      { return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function appendLog(msg) {
  const line = document.createElement('div');
  line.textContent = msg;
  progressLog.appendChild(line);
  progressLog.scrollTop = progressLog.scrollHeight;
}

let _toastTimer;
function showToast(msg) {
  const t = $('toast');
  t.textContent = msg;
  t.classList.remove('hidden');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => t.classList.add('hidden'), 3500);
}
