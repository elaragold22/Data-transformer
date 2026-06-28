(function () {
  'use strict';

  const state = {
    mode: null,   // null | 'processing' | 'table' | 'text'
    fileName: '',
    headers: [],
    rows: [],
    text: '',
  };

  const PREVIEW_LIMIT = 200;

  const el = {
    inputScreen:      document.getElementById('input-screen'),
    dropzone:         document.getElementById('dropzone'),
    fileInput:        document.getElementById('file-input'),
    textInput:        document.getElementById('text-input'),
    btnLoadText:      document.getElementById('btn-load-text'),
    processingScreen: document.getElementById('processing-screen'),
    processingFile:   document.getElementById('processing-file'),
    progressBar:      document.getElementById('progress-bar'),
    progressPct:      document.getElementById('progress-pct'),
    previewWrap:      document.getElementById('preview-wrap'),
    tableHead:        document.getElementById('table-head'),
    tableBody:        document.getElementById('table-body'),
    previewNote:      document.getElementById('preview-note'),
    textPreviewWrap:  document.getElementById('text-preview-wrap'),
    textPreview:      document.getElementById('text-preview'),
    fileNameEl:       document.getElementById('file-name'),
    fileMeta:         document.getElementById('file-meta'),
    stats:            document.getElementById('stats'),
    statRows:         document.getElementById('stat-rows'),
    statCols:         document.getElementById('stat-cols'),
    readyPill:        document.getElementById('ready-pill'),
    btnDedupe:        document.getElementById('btn-dedupe'),
    btnCsv:           document.getElementById('btn-export-csv'),
    btnJson:          document.getElementById('btn-export-json'),
    btnXlsx:          document.getElementById('btn-export-xlsx'),
    btnTxt:           document.getElementById('btn-export-txt'),
    btnReset:         document.getElementById('btn-reset'),
    btnTheme:         document.getElementById('btn-theme'),
    toast:            document.getElementById('toast'),
  };

  /* =======================================================================
   * Theme toggle
   * ===================================================================== */
  el.btnTheme.addEventListener('click', () => {
    const isDark = document.documentElement.classList.toggle('dark');
    try { localStorage.setItem('theme', isDark ? 'dark' : 'light'); } catch (e) {}
  });

  /* =======================================================================
   * File intake — drag & drop + click to browse
   * ===================================================================== */
  el.dropzone.addEventListener('click', () => el.fileInput.click());

  el.fileInput.addEventListener('change', (e) => {
    const file = e.target.files && e.target.files[0];
    if (file) handleFile(file);
    el.fileInput.value = '';
  });

  el.btnLoadText.addEventListener('click', () => {
    const text = el.textInput.value;
    if (!text.trim()) { showToast('Please enter some text first.', true); return; }
    loadText(text);
  });

  ['dragenter', 'dragover'].forEach((evt) => {
    el.dropzone.addEventListener(evt, (e) => {
      e.preventDefault(); e.stopPropagation();
      el.dropzone.classList.add('border-brand-500', 'bg-brand-50');
    });
  });
  ['dragleave', 'drop'].forEach((evt) => {
    el.dropzone.addEventListener(evt, (e) => {
      e.preventDefault(); e.stopPropagation();
      el.dropzone.classList.remove('border-brand-500', 'bg-brand-50');
    });
  });
  el.dropzone.addEventListener('drop', (e) => {
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) handleFile(file);
  });

  ['dragover', 'drop'].forEach((evt) => {
    window.addEventListener(evt, (e) => e.preventDefault());
  });

  /* =======================================================================
   * Routing
   * ===================================================================== */
  function handleFile(file) {
    const name = file.name || 'file';
    const ext = name.split('.').pop().toLowerCase();
    if (ext === 'csv') {
      startProcessing(name);
      parseCSV(file, name);
    } else if (ext === 'xlsx' || ext === 'xls') {
      startProcessing(name);
      parseExcel(file, name);
    } else {
      showToast('Unsupported file type. Please use .csv, .xlsx, or .xls', true);
    }
  }

  /* =======================================================================
   * Processing progress bar
   * ===================================================================== */
  let progressTimer = null;
  let processingStart = 0;
  const MIN_PROCESSING_MS = 1400;

  function setProgress(p) {
    const v = Math.round(p);
    el.progressBar.style.width = v + '%';
    el.progressPct.textContent = v + '%';
  }

  function startProcessing(name) {
    processingStart = Date.now();
    el.processingFile.textContent = name;
    setProgress(0);
    setMode('processing');
    let p = 0;
    clearInterval(progressTimer);
    progressTimer = setInterval(() => {
      const inc = Math.max(1.5, (90 - p) * 0.14);
      p = Math.min(90, p + inc);
      setProgress(p);
      if (p >= 90) clearInterval(progressTimer);
    }, 90);
  }

  function completeProcessing(done) {
    clearInterval(progressTimer);
    setProgress(100);
    const elapsed = Date.now() - processingStart;
    const wait = Math.max(400, MIN_PROCESSING_MS - elapsed);
    setTimeout(done, wait);
  }

  function cancelProcessing() {
    clearInterval(progressTimer);
    setMode(null);
  }

  /* =======================================================================
   * Parsing
   * ===================================================================== */
  function parseCSV(file, name) {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: 'greedy',
      dynamicTyping: false,
      complete: (results) => {
        const headers = (results.meta && results.meta.fields) || [];
        const rows = results.data || [];
        if (!rows.length) { cancelProcessing(); showToast('That CSV appears to be empty.', true); return; }
        loadData(name, headers, rows);
      },
      error: (err) => {
        console.error(err);
        cancelProcessing();
        showToast('Could not read the CSV file.', true);
      },
    });
  }

  function parseExcel(file, name) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: false });

        if (!aoa.length) { cancelProcessing(); showToast('That worksheet appears to be empty.', true); return; }

        const headers = aoa[0].map((h, i) => String(h).trim() || `Column ${i + 1}`);
        const rows = aoa.slice(1).map((arr) => {
          const obj = {};
          headers.forEach((h, i) => { obj[h] = arr[i] !== undefined ? arr[i] : ''; });
          return obj;
        });

        if (!rows.length) { cancelProcessing(); showToast('No data rows found in that worksheet.', true); return; }

        loadData(name, headers, rows, workbook.SheetNames.length);
      } catch (err) {
        console.error(err);
        cancelProcessing();
        showToast('Could not read the Excel file.', true);
      }
    };
    reader.onerror = () => { cancelProcessing(); showToast('Failed to read the file.', true); };
    reader.readAsArrayBuffer(file);
  }

  /* =======================================================================
   * Load data
   * ===================================================================== */
  function loadData(name, headers, rows, sheetCount) {
    state.fileName = name;
    state.headers = headers;
    state.rows = rows;
    state.text = '';

    el.fileNameEl.textContent = name;
    const parts = [`${rows.length.toLocaleString()} rows`, `${headers.length} columns`];
    if (sheetCount && sheetCount > 1) parts.push(`sheet 1 of ${sheetCount}`);
    el.fileMeta.textContent = parts.join('  •  ');

    completeProcessing(() => {
      renderTable();
      setMode('table');
      showToast('File loaded successfully.');
    });
  }

  function loadText(text) {
    state.mode = 'text';
    state.fileName = 'prompt';
    state.headers = [];
    state.rows = [];
    state.text = text;

    el.fileNameEl.textContent = 'Text / Prompt';
    const lines = text.split(/\r\n|\r|\n/).length;
    el.fileMeta.textContent = `${text.length.toLocaleString()} characters  •  ${lines.toLocaleString()} lines`;

    el.textPreview.textContent = plainTextToMarkdown(text);
    setMode('text');
    showToast('Text loaded.');
  }

  /* =======================================================================
   * Render table
   * ===================================================================== */
  function renderTable() {
    const headHtml =
      '<tr>' +
      '<th class="sticky left-0 z-20 bg-slate-100 px-3 py-2.5 text-right text-slate-400">#</th>' +
      state.headers.map((h) =>
        `<th class="whitespace-nowrap border-b border-slate-200 px-4 py-2.5">${escapeHtml(h)}</th>`
      ).join('') +
      '</tr>';
    el.tableHead.innerHTML = headHtml;

    const view = state.rows.slice(0, PREVIEW_LIMIT);
    el.tableBody.innerHTML = view.map((row, idx) =>
      `<tr class="hover:bg-slate-50">` +
      `<td class="sticky left-0 z-10 bg-white px-3 py-2 text-right text-xs text-slate-400">${idx + 1}</td>` +
      state.headers.map((h) => {
        const v = row[h];
        return `<td class="whitespace-nowrap px-4 py-2 text-slate-700">${escapeHtml(v == null ? '' : String(v))}</td>`;
      }).join('') +
      `</tr>`
    ).join('');

    el.statRows.textContent = state.rows.length.toLocaleString();
    el.statCols.textContent = state.headers.length;
    el.stats.classList.remove('hidden');
    el.stats.classList.add('flex');

    el.previewNote.textContent = state.rows.length > PREVIEW_LIMIT
      ? `Showing first ${PREVIEW_LIMIT.toLocaleString()} of ${state.rows.length.toLocaleString()} rows`
      : `Showing all ${state.rows.length.toLocaleString()} rows`;
  }

  /* =======================================================================
   * Actions
   * ===================================================================== */
  el.btnDedupe.addEventListener('click', () => {
    if (!state.rows.length) return;
    const seen = new Set();
    const unique = [];
    for (const row of state.rows) {
      const key = state.headers.map((h) => String(row[h] ?? '')).join('\x00');
      if (!seen.has(key)) { seen.add(key); unique.push(row); }
    }
    const removed = state.rows.length - unique.length;
    state.rows = unique;
    renderTable();
    showToast(removed > 0
      ? `Removed ${removed.toLocaleString()} duplicate row${removed === 1 ? '' : 's'}.`
      : 'No duplicate rows found.');
  });

  el.btnCsv.addEventListener('click', () => {
    if (!state.rows.length) return;
    const csv = Papa.unparse({ fields: state.headers, data: state.rows }, { quotes: false });
    downloadBlob(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' }), baseName() + '.csv');
    showToast('CSV exported.');
  });

  el.btnJson.addEventListener('click', () => {
    if (!state.rows.length) return;
    downloadBlob(new Blob([JSON.stringify(state.rows, null, 2)], { type: 'application/json;charset=utf-8;' }), baseName() + '.json');
    showToast('JSON exported.');
  });

  el.btnXlsx.addEventListener('click', () => {
    if (!state.rows.length) return;
    const ws = XLSX.utils.json_to_sheet(state.rows, { header: state.headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    downloadBlob(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), baseName() + '.xlsx');
    showToast('Excel exported.');
  });

  /* Convert plain text / prompt into Markdown markup.
   * Heuristics: first short non-terminal line becomes an H1 title; "Label: value"
   * lines (1-3 word labels) become bold-label bullets; existing bullet and numbered
   * markers are normalized; everything else stays as paragraphs (newlines kept). */
  function plainTextToMarkdown(text) {
    const src = String(text == null ? '' : text).replace(/\r\n?/g, '\n');
    const lines = src.split('\n');
    const out = [];
    const prev = () => (out.length ? out[out.length - 1] : '');

    const KV     = /^([A-Za-z][\w&/()'’.-]*(?:[ ][A-Za-z0-9&/()'’.-]+){0,2}):[ \t]+(\S.*)$/;
    const BULLET = /^[-*•·]\s+(.*)$/;
    const NUM    = /^\d+[.)]\s+.*$/;
    const firstIdx = lines.findIndex((l) => l.trim() !== '');

    lines.forEach((raw, i) => {
      const line = raw.trim();
      if (!line) { if (prev() !== '') out.push(''); return; }

      // Optional H1: first content line, short, no terminal punctuation/colon, more text follows
      if (i === firstIdx &&
          line.split(/\s+/).length <= 8 &&
          !/[:.!?]$/.test(line) &&
          !KV.test(line) &&
          lines.slice(i + 1).some((l) => l.trim() !== '')) {
        out.push('# ' + line);
        return;
      }

      let m;
      if ((m = line.match(BULLET))) {
        if (prev() !== '' && !/^[-\d]/.test(prev())) out.push('');
        out.push('- ' + m[1].trim());
      } else if (NUM.test(line)) {
        if (prev() !== '' && !/^[-\d]/.test(prev())) out.push('');
        out.push(line);
      } else if ((m = line.match(KV))) {
        if (prev() !== '' && !prev().startsWith('- ')) out.push('');
        out.push(`- **${m[1].trim()}:** ${m[2].trim()}`);
      } else {
        if (prev().startsWith('- ') || /^\d+[.)]/.test(prev())) out.push('');
        out.push(line);
      }
    });

    return out.join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
  }

  el.btnTxt.addEventListener('click', () => {
    if (!state.text) return;
    const md = plainTextToMarkdown(state.text);
    downloadBlob(new Blob([md], { type: 'text/plain;charset=utf-8;' }), baseName() + '.txt');
    showToast('Markdown exported.');
  });

  el.btnReset.addEventListener('click', resetApp);

  function resetApp() {
    state.mode = null;
    state.fileName = '';
    state.headers = [];
    state.rows = [];
    state.text = '';
    el.tableHead.innerHTML = '';
    el.tableBody.innerHTML = '';
    el.textPreview.textContent = '';
    el.textInput.value = '';
    el.fileNameEl.textContent = 'No file loaded';
    el.fileMeta.textContent = 'Drop a CSV or Excel file to begin';
    el.stats.classList.add('hidden');
    el.stats.classList.remove('flex');
    setMode(null);
  }

  /* =======================================================================
   * Helpers
   * ===================================================================== */
  function setMode(mode) {
    const isTable = mode === 'table';
    const isText = mode === 'text';
    const isProcessing = mode === 'processing';
    const active = isTable || isText;
    const busy = active || isProcessing;

    el.inputScreen.classList.toggle('hidden', busy);
    el.inputScreen.classList.toggle('flex', !busy);

    el.processingScreen.classList.toggle('hidden', !isProcessing);
    el.processingScreen.classList.toggle('flex', isProcessing);

    el.previewWrap.classList.toggle('hidden', !isTable);
    el.previewWrap.classList.toggle('flex', isTable);

    el.textPreviewWrap.classList.toggle('hidden', !isText);
    el.textPreviewWrap.classList.toggle('flex', isText);

    el.readyPill.classList.toggle('hidden', !active);
    el.readyPill.classList.toggle('inline-flex', active);

    el.btnDedupe.disabled = !isTable;
    el.btnCsv.disabled    = !isTable;
    el.btnJson.disabled   = !isTable;
    el.btnXlsx.disabled   = !isTable;
    el.btnTxt.disabled    = !isText;   // Markdown (.txt) — text mode only
    el.btnReset.disabled  = !active;
  }

  function baseName() {
    const n = state.fileName || 'data';
    const dot = n.lastIndexOf('.');
    return (dot > 0 ? n.slice(0, dot) : n) + '_transformed';
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = Object.assign(document.createElement('a'), { href: url, download: filename });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  let toastTimer = null;
  function showToast(message, isError) {
    const icon = isError
      ? '<svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4 shrink-0 text-rose-300" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M12 9v3.75m9-.75a9 9 0 1 1-18 0 9 9 0 0 1 18 0Zm-9 3.75h.008v.008H12v-.008Z" /></svg>'
      : '<svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4 shrink-0 text-emerald-300" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12.75 11.25 15 15 9.75M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0Z" /></svg>';
    el.toast.innerHTML = icon + '<span>' + escapeHtml(message) + '</span>';
    el.toast.classList.toggle('bg-rose-600', !!isError);
    el.toast.classList.toggle('bg-slate-900', !isError);
    el.toast.classList.remove('opacity-0', 'translate-y-4');
    el.toast.classList.add('opacity-100', 'translate-y-0');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => {
      el.toast.classList.add('opacity-0', 'translate-y-4');
      el.toast.classList.remove('opacity-100', 'translate-y-0');
    }, 2600);
  }
})();
