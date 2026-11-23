// Workbook: screens, load of the WASM module, open, new, download, the dirty
// mark and file drop.

import {
  READY_TIMEOUT_MS, XLSX_MIME, el, screens, S, showStatus, clearStatus, rawCall, api, emit, confirmDialog,
} from './core.js';
import { render, updateSelectionUI, updateBars } from './grid.js';
import { showSheet, saveView } from './sheets.js';
import { endEdit } from './edit.js';

export function setState(state) {
  S.state = state;
  document.body.dataset.state = state;
  for (const key of Object.keys(screens)) screens[key].hidden = key !== state;
  if (state === 'start') el.backToEditor.hidden = !S.hasWorkbook;
}

function failLoad(detail) {
  if (S.state !== 'loading') return;
  el.loadErrorDetail.textContent = detail;
  setState('error');
}

export function boot() {
  if (typeof globalThis.Go !== 'function') {
    failLoad('wasm_exec.js is not available.');
    return;
  }
  if (!globalThis.WebAssembly || typeof WebAssembly.instantiateStreaming !== 'function') {
    failLoad('This browser does not support WebAssembly streaming.');
    return;
  }
  const go = new globalThis.Go();
  const timer = setTimeout(() => failLoad('The module did not become ready in time.'), READY_TIMEOUT_MS);
  globalThis.onWerkbookReady = () => {
    clearTimeout(timer);
    if (!globalThis.werkbook) {
      failLoad('The module did not set globalThis.werkbook.');
      return;
    }
    const v = rawCall('version', []);
    el.engineVersion.textContent = v.ok && v.version ? `werkbook ${v.version}` : '';
    if (S.state === 'loading') setState('start');
  };
  // The path is relative to the page, not to this module.
  WebAssembly.instantiateStreaming(fetch('werkbook.wasm'), go.importObject)
    .then((r) => go.run(r.instance))
    .then(() => {
      // go.run resolves only when the Go program stops.
      clearTimeout(timer);
      if (S.state === 'loading') failLoad('The module stopped before it was ready.');
      else showStatus('The spreadsheet engine stopped. Load the page again.', 'error');
    })
    .catch((err) => {
      clearTimeout(timer);
      const msg = err && err.message ? err.message : String(err);
      if (S.state === 'loading') failLoad(`werkbook.wasm did not load: ${msg}`);
      else showStatus(`The spreadsheet engine stopped: ${msg}`, 'error');
    });
}

// ------------------------------------------------------- workbook lifecycle

async function confirmDiscard() {
  if (!S.hasWorkbook || !S.dirty) return true;
  return confirmDialog(`"${S.fileName}" has unsaved changes. Discard the changes?`, 'Discard changes');
}

async function openFile(file) {
  if (!file) return;
  if (S.state !== 'start' && S.state !== 'editor') return;
  if (!/\.xlsx$/i.test(file.name)) {
    showStatus(`"${file.name}" is not an .xlsx file.`, 'error');
    return;
  }
  if (!(await confirmDiscard())) return;
  let bytes;
  try {
    bytes = new Uint8Array(await file.arrayBuffer());
  } catch (err) {
    showStatus(`Could not read "${file.name}": ${err && err.message ? err.message : err}`, 'error');
    return;
  }
  const res = api('open', bytes);
  if (!res.ok) {
    afterFailedReplace();
    return;
  }
  S.history = res.history || null;
  loadWorkbook(res.sheets, file.name);
}

async function newWorkbook() {
  if (!(await confirmDiscard())) return;
  const res = api('newWorkbook');
  if (!res.ok) {
    afterFailedReplace();
    return;
  }
  S.history = res.history || null;
  loadWorkbook(res.sheets, 'workbook.xlsx');
}

// The README does not say if the module keeps the old workbook when open
// fails, so ask the module.
function afterFailedReplace() {
  if (!S.hasWorkbook) return;
  const res = rawCall('sheets', []);
  if (res.ok && Array.isArray(res.sheets) && res.sheets.length) return;
  S.hasWorkbook = false;
  S.sheet = null;
  setState('start');
}

// loadWorkbook shows a workbook that the module opened. The caller stores
// the history state of the open or new result in S.history before the call.
export function loadWorkbook(sheets, fileName) {
  if (!Array.isArray(sheets) || !sheets.length) {
    showStatus('The workbook has no sheets.', 'error');
    afterFailedReplace();
    return;
  }
  clearStatus();
  endEdit();
  S.hasWorkbook = true;
  S.fileName = fileName;
  S.sheets = sheets;
  S.views.clear();
  S.sheet = null;
  setDirty(false);
  setState('editor');
  const first = sheets.find((s) => s.visible !== false) || sheets[0];
  showSheet(first.name);
  el.grid.focus({ preventScroll: true });
  emit('workbook');
  emit('history', S.history);
}

export function setDirty(dirty) {
  S.dirty = dirty;
  el.dirty.hidden = !dirty;
  el.fileName.textContent = S.fileName;
  el.fileName.title = S.fileName;
  document.title = `${dirty ? '● ' : ''}${S.fileName} – Werkbook`;
}

function download() {
  const res = api('save');
  if (!res.ok) return;
  const url = URL.createObjectURL(new Blob([res.data], { type: XLSX_MIME }));
  const a = document.createElement('a');
  a.href = url;
  a.download = S.fileName;
  a.dataset.downloadLink = '';
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
  setDirty(false);
  showStatus(`Downloaded ${S.fileName}.`, 'info');
}

// ------------------------------------------------- start screen and toolbar

const hasFiles = (e) => Boolean(e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files'));

function setDropActive(active) {
  const usable = active && (S.state === 'start' || S.state === 'editor');
  el.dropZone.classList.toggle('drag-over', usable && S.state === 'start');
  el.dropOverlay.hidden = !(usable && S.state === 'editor');
}

export function init() {
  el.openFile.addEventListener('click', () => el.fileInput.click());
  el.fileInput.addEventListener('change', () => {
    const file = el.fileInput.files && el.fileInput.files[0];
    el.fileInput.value = '';
    openFile(file);
  });
  el.newWorkbook.addEventListener('click', newWorkbook);
  el.backToEditor.addEventListener('click', () => {
    if (!S.hasWorkbook) return;
    setState('editor');
    render(true);
    updateSelectionUI();
    updateBars();
    el.grid.focus({ preventScroll: true });
  });
  el.openAnother.addEventListener('click', () => {
    // The workbook stays open until the user selects a different one.
    saveView();
    setState('start');
    el.openFile.focus();
  });
  el.download.addEventListener('click', download);
  el.retryLoad.addEventListener('click', () => location.reload());

  document.addEventListener('dragenter', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    S.dragDepth++;
    setDropActive(true);
  });
  document.addEventListener('dragover', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = S.state === 'start' || S.state === 'editor' ? 'copy' : 'none';
  });
  document.addEventListener('dragleave', (e) => {
    if (!hasFiles(e)) return;
    S.dragDepth = Math.max(0, S.dragDepth - 1);
    if (!S.dragDepth) setDropActive(false);
  });
  document.addEventListener('drop', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    S.dragDepth = 0;
    setDropActive(false);
    openFile(e.dataTransfer.files && e.dataTransfer.files[0]);
  });
}
