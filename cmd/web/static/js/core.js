// Werkbook web editor: shared base of all modules.
// Constants, DOM elements, state, helpers, the status area, the calls to the
// WASM module, the confirm dialog, the event bus and the key registry.
// The JavaScript API contract is in cmd/web/README.md. The module interface
// is in cmd/web/PLAN.md, section 6.
//
// This module imports no other module. A module must not call an imported
// function at load time, only in init() or later.

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;
export const DEF_ROW_H = 22;
export const DEF_COL_W = 88;
export const HEADER_H = 24;
export const MIN_ROWS = 100;
export const MIN_COLS = 26;
export const ROW_GROW = 200;
export const COL_GROW = 20;
export const ROW_BLOCK = 8; // the rendered range snaps to blocks, so a small scroll does not rebuild the DOM
export const COL_BLOCK = 4;
// Browsers clamp the size of one element (Firefox near 17.8 million pixels).
export const MAX_CANVAS_PX = 16000000;
export const MAX_FILL_CELLS = 200000;
export const MAX_COPY_CELLS = 2000000;
export const READY_TIMEOUT_MS = 30000;
export const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const $ = (id) => document.getElementById(id);

export const el = {
  loadErrorDetail: $('load-error-detail'),
  retryLoad: $('retry-load'),
  dropZone: $('drop-zone'),
  openFile: $('open-file'),
  newWorkbook: $('new-workbook'),
  fileInput: $('file-input'),
  backToEditor: $('back-to-editor'),
  engineVersion: $('engine-version'),
  fileName: $('file-name'),
  dirty: $('dirty-indicator'),
  openAnother: $('open-another'),
  download: $('download'),
  nameBox: $('name-box'),
  formula: $('formula-input'),
  grid: $('grid'),
  canvas: $('grid-canvas'),
  colHeaders: $('col-headers'),
  corner: $('corner'),
  rowHeaders: $('row-headers'),
  cells: $('cells'),
  overlay: $('overlay'),
  selRange: $('selection-range'),
  active: $('active-cell'),
  pointRange: $('point-range'),
  editor: $('cell-editor'),
  addSheet: $('add-sheet'),
  tabs: $('sheet-tabs'),
  status: $('status'),
  statusText: $('status-text'),
  statusDismiss: $('status-dismiss'),
  confirm: $('confirm-dialog'),
  confirmMessage: $('confirm-message'),
  confirmOk: $('confirm-ok'),
  confirmCancel: $('confirm-cancel'),
  dropOverlay: $('drop-overlay'),
  // Mount points for the feature modules (PLAN.md, section 6.3).
  toolbar: $('toolbar'),
  tbHistory: $('tb-history'),
  tbFormat: $('tb-format'),
  tbData: $('tb-data'),
  paneCorner: $('pane-corner'),
  paneTop: $('pane-top'),
  paneLeft: $('pane-left'),
  refLayer: $('ref-layer'),
  fillRange: $('fill-range'),
  fillHandle: $('fill-handle'),
  resizeGuide: $('resize-guide'),
  fnPopup: $('fn-popup'),
  fnHint: $('fn-hint'),
  contextMenu: $('context-menu'),
  findBar: $('find-bar'),
};

export const screens = {
  loading: $('loading'),
  error: $('load-error'),
  start: $('start-screen'),
  editor: $('editor'),
};

export const S = {
  state: 'loading',
  hasWorkbook: false,
  fileName: 'workbook.xlsx',
  dirty: false,
  sheets: [],
  sheet: null, // name of the sheet on screen
  views: new Map(), // sheet name -> scroll position and selection
  rows: null,
  cols: null,
  rowLimit: MAX_ROWS,
  colLimit: MAX_COLS,
  rowCount: MIN_ROWS,
  colCount: MIN_COLS,
  rw: 44, // row header width in pixels
  merges: [],
  anchor: { row: 1, col: 1 }, // the active cell
  focus: { row: 1, col: 1 }, // the opposite corner of the selected range
  tabOrigin: null,
  edit: null,
  rendered: null,
  cells: new Map(),
  drag: null,
  scrollRaf: 0,
  statusTimer: 0,
  confirmResolve: null,
  dragDepth: 0,
  history: null, // HistoryState of the last mutation result; null until the module reports one
  freeze: { rows: 0, cols: 0 }, // from sheetMeta of the sheet on screen
  warnings: [], // sheetMeta.warnings of the sheet on screen
  functions: null, // FunctionInfo[] from functions(); fnhelp.js reads it one time
};

export const darkQuery = window.matchMedia('(prefers-color-scheme: dark)');

// ------------------------------------------------------------------ helpers

export const clamp = (n, lo, hi) => Math.min(hi, Math.max(lo, n));
export const cellKey = (row, col) => row * (MAX_COLS + 1) + col;

export function colName(col) {
  let s = '';
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

export function colIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export const addrOf = (row, col) => colName(col) + row;

// parseRef accepts "B5", "B5:D9" and an optional sheet prefix ("Sheet2!A1", "'My sheet'!A1").
export function parseRef(text) {
  const m = /^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?$/.exec(text.trim());
  if (!m) return null;
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  const cA = colIndex(m[3]);
  const rA = Number(m[4]);
  const cB = m[5] ? colIndex(m[5]) : cA;
  const rB = m[6] ? Number(m[6]) : rA;
  if (Math.min(rA, rB) < 1 || Math.max(rA, rB) > MAX_ROWS || Math.max(cA, cB) > MAX_COLS) return null;
  return { sheet, row1: rA, col1: cA, row2: rB, col2: cB };
}

// An axis maps row or column numbers to pixel positions. Most rows and
// columns have the default size, so only the overrides are stored.
export function makeAxis(def, overrides, max) {
  const sizes = new Map();
  if (overrides && typeof overrides === 'object') {
    for (const k of Object.keys(overrides)) {
      const i = Number(k);
      const v = Number(overrides[k]);
      if (Number.isInteger(i) && i >= 1 && i <= max && v >= 0) sizes.set(i, v);
    }
  }
  const idx = [...sizes.keys()].sort((a, b) => a - b);
  const cum = [];
  let run = 0;
  for (const i of idx) {
    run += sizes.get(i) - def;
    cum.push(run);
  }
  const before = (i) => {
    let lo = 0;
    let hi = idx.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (idx[mid] < i) lo = mid + 1;
      else hi = mid;
    }
    return lo ? cum[lo - 1] : 0;
  };
  const offset = (i) => (i - 1) * def + before(i);
  return {
    offset,
    size: (i) => (sizes.has(i) ? sizes.get(i) : def),
    indexAt(px) {
      if (px <= 0) return 1;
      let lo = 1;
      let hi = max;
      while (lo < hi) {
        const mid = Math.floor((lo + hi + 1) / 2);
        if (offset(mid) <= px) lo = mid;
        else hi = mid - 1;
      }
      return lo;
    },
  };
}

export function luminance(hex) {
  const n = parseInt(hex.slice(1), 16);
  return (0.2126 * ((n >> 16) & 255) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255)) / 255;
}

export const isHexColor = (v) => typeof v === 'string' && /^#[0-9a-fA-F]{6}$/.test(v);

// ------------------------------------------------------------------- status

export function showStatus(message, kind) {
  el.statusText.textContent = message;
  el.status.dataset.kind = kind === 'error' ? 'error' : 'info';
  el.status.hidden = false;
  clearTimeout(S.statusTimer);
  // An error stays until the user dismisses it or a new message replaces it.
  if (kind !== 'error') S.statusTimer = setTimeout(clearStatus, 4000);
}

export function clearStatus() {
  clearTimeout(S.statusTimer);
  el.status.hidden = true;
  el.statusText.textContent = '';
}

// ---------------------------------------------------------------- event bus
//
// Events (PLAN.md, section 6.2): workbook, sheet, render, selection, write,
// history, edit:start, edit:end, edit:change.

const handlers = new Map();

// on adds a handler. Handlers of one event run in the sequence of their
// registration, which is the sequence of the init() calls in main.js.
export function on(event, handler) {
  if (!handlers.has(event)) handlers.set(event, []);
  handlers.get(event).push(handler);
}

// emit calls each handler of the event. An error in one handler does not
// stop the other handlers.
export function emit(event, data) {
  const list = handlers.get(event);
  if (!list) return;
  for (const handler of list.slice()) {
    try {
      handler(data);
    } catch (err) {
      console.error(`handler of "${event}" failed`, err);
    }
  }
}

// ---------------------------------------------------------------------- API

export function rawCall(name, args) {
  try {
    const fn = globalThis.werkbook && globalThis.werkbook[name];
    if (typeof fn !== 'function') return { ok: false, error: 'the function is not available' };
    const res = fn(...args);
    if (res && res.ok === true) return res;
    return { ok: false, error: res && res.error ? String(res.error) : 'unknown error' };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

// api calls the WASM module and shows each {ok:false} result in the status area.
export function api(name, ...args) {
  const res = rawCall(name, args);
  if (!res.ok) showStatus(`${name}: ${res.error}`, 'error');
  return res;
}

// mutate calls a contract function that changes the workbook. On success it
// stores res.history in S.history and emits 'write' (data: the result) and
// 'history' (data: S.history). grid.js sets the dirty mark and repaints on
// 'write'. All features must use mutate for writes, so that undo, the dirty
// mark and the repaint are the same for all.
export function mutate(name, ...args) {
  const res = api(name, ...args);
  if (!res.ok) return res;
  if (res.history) S.history = res.history;
  emit('write', res);
  emit('history', S.history);
  return res;
}

// ------------------------------------------------------------- key registry

const keyHandlers = { grid: [], edit: [] };

function comboMatches(combo, e) {
  const want = String(combo.key);
  const key = e.key || '';
  const same = want.length === 1 ? key.toLowerCase() === want.toLowerCase() : key === want;
  return (
    same &&
    Boolean(combo.mod) === (e.ctrlKey || e.metaKey) &&
    Boolean(combo.shift) === e.shiftKey &&
    Boolean(combo.alt) === e.altKey
  );
}

// registerKey adds a key handler.
// combo: { key, mod?: boolean (Ctrl or Cmd), shift?: boolean, alt?: boolean }.
// key is a KeyboardEvent.key value; one letter matches in both letter cases.
// A modifier that is absent must not be pressed.
// scope: 'grid' (the grid has the focus, no edit is open) or 'edit' (the
// cell editor or the formula bar has the focus, an edit is open).
// The handler gets the event. It returns false to let the next handler or
// the default behavior run.
export function registerKey(scope, combo, handler) {
  keyHandlers[scope].push({ combo, handler });
}

// runKey is for grid.js and edit.js only. It returns true when a handler
// used the event; the default behavior of the browser is then prevented.
export function runKey(scope, e) {
  for (const { combo, handler } of keyHandlers[scope]) {
    if (!comboMatches(combo, e)) continue;
    if (handler(e) === false) continue;
    e.preventDefault();
    return true;
  }
  return false;
}

// ----------------------------------------------------------- confirm dialog

export function confirmDialog(message, okLabel) {
  if (S.confirmResolve) S.confirmResolve(false);
  return new Promise((resolve) => {
    const previous = document.activeElement;
    el.confirmMessage.textContent = message;
    el.confirmOk.textContent = okLabel;
    el.confirm.hidden = false;
    el.confirmCancel.focus();
    S.confirmResolve = (value) => {
      S.confirmResolve = null;
      el.confirm.hidden = true;
      if (previous && document.contains(previous) && typeof previous.focus === 'function') previous.focus();
      resolve(value);
    };
  });
}

export function init() {
  el.confirmOk.addEventListener('click', () => S.confirmResolve && S.confirmResolve(true));
  el.confirmCancel.addEventListener('click', () => S.confirmResolve && S.confirmResolve(false));
  el.confirm.addEventListener('mousedown', (e) => {
    if (e.target === el.confirm && S.confirmResolve) S.confirmResolve(false);
  });
  el.confirm.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      if (S.confirmResolve) S.confirmResolve(false);
    } else if (e.key === 'Tab') {
      // Keep the focus in the dialog.
      e.preventDefault();
      (document.activeElement === el.confirmOk ? el.confirmCancel : el.confirmOk).focus();
    }
  });
  el.statusDismiss.addEventListener('click', clearStatus);
}
