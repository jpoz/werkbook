// Undo and redo UI (PLAN.md 4.1). Owner: fe-a.
//
// The history is in the bridge. This module shows its state (#undo, #redo,
// #grid[data-history-depth]) and applies the result of undo() and redo().

import { MAX_ROWS, MAX_COLS, el, S, clamp, showStatus, rawCall, mutate, on, emit, registerKey } from './core.js';
import { reloadMeta, setSelection } from './grid.js';
import { showSheet } from './sheets.js';

let undoBtn = null;
let redoBtn = null;

function makeButton(id, text, label) {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.id = id;
  btn.className = 'btn icon';
  btn.textContent = text;
  btn.title = label;
  btn.setAttribute('aria-label', label);
  btn.disabled = true;
  // The focus stays where it is, so a click does not move it out of the grid.
  btn.addEventListener('mousedown', (e) => e.preventDefault());
  return btn;
}

// update shows the history state in the buttons and on the grid.
function update() {
  const h = S.history;
  const canUndo = Boolean(h && h.canUndo);
  const canRedo = Boolean(h && h.canRedo);
  undoBtn.disabled = !canUndo;
  redoBtn.disabled = !canRedo;
  const undoTitle = canUndo && h.undoLabel ? `Undo ${h.undoLabel}` : 'Undo';
  const redoTitle = canRedo && h.redoLabel ? `Redo ${h.redoLabel}` : 'Redo';
  undoBtn.title = undoTitle;
  redoBtn.title = redoTitle;
  undoBtn.setAttribute('aria-label', undoTitle);
  redoBtn.setAttribute('aria-label', redoTitle);
  el.grid.dataset.historyDepth = h && Number.isFinite(h.depth) ? String(h.depth) : '0';
  if (h && h.cleared) {
    showStatus('The change was too large for the undo history. The undo history is now empty.', 'info');
  }
}

// step runs undo() or redo() and applies the HistoryResult: it shows the
// sheet of the step, selects its range and scrolls to it.
function step(name) {
  if (S.state !== 'editor' || !S.sheet || S.edit) return;
  // Nothing to undo or redo: do nothing, with no error message.
  const h = S.history;
  if (!h || (name === 'undo' ? !h.canUndo : !h.canRedo)) return;
  // mutate emits 'write' (grid.js sets the dirty mark and paints again) and 'history'.
  const res = mutate(name);
  if (!res.ok) return;
  if (Array.isArray(res.sheets) && res.sheets.length) S.sheets = res.sheets;
  const names = new Set(S.sheets.map((s) => s.name));
  // A step can add, remove or rename a sheet. Stored views of names that
  // are not in the workbook are of no use.
  for (const key of [...S.views.keys()]) if (!names.has(key)) S.views.delete(key);
  if (!names.has(S.sheet)) S.sheet = null;
  let target = names.has(res.sheet) ? res.sheet : S.sheet;
  if (!target) target = (S.sheets.find((s) => s.visible !== false) || S.sheets[0]).name;
  if (target !== S.sheet) showSheet(target);
  else reloadMeta(); // the layout of the sheet can be different after the step
  const sel = res.select;
  if (sel) {
    const anchor = {
      row: clamp(Math.min(sel.row1, sel.row2), 1, S.rowLimit),
      col: clamp(Math.min(sel.col1, sel.col2), 1, S.colLimit),
    };
    // The far corner keeps its true row and column, so a step on full
    // columns or full rows gives a full-column or full-row selection, also
    // when the grid shows fewer rows.
    const focus = {
      row: clamp(Math.max(sel.row1, sel.row2), 1, MAX_ROWS),
      col: clamp(Math.max(sel.col1, sel.col2), 1, MAX_COLS),
    };
    // Show the first cell; a large range does not move the view to its far corner.
    setSelection(anchor, anchor);
    if (focus.row !== anchor.row || focus.col !== anchor.col) setSelection(anchor, focus, { scroll: false });
  }
  el.grid.focus({ preventScroll: true });
}

export function undo() {
  step('undo');
}

export function redo() {
  step('redo');
}

export function init() {
  undoBtn = makeButton('undo', '↶', 'Undo');
  redoBtn = makeButton('redo', '↷', 'Redo');
  el.tbHistory.append(undoBtn, redoBtn);
  undoBtn.addEventListener('click', undo);
  redoBtn.addEventListener('click', redo);

  on('history', update);
  on('workbook', () => {
    // A module that has no history support gives no history field in its results.
    if (!S.history) {
      const res = rawCall('history', []);
      if (res.ok && res.history) {
        S.history = res.history;
        emit('history', S.history);
        return;
      }
    }
    update();
  });

  // The keys apply only when the grid has the focus. In the cell editor and
  // the formula bar, Ctrl+Z is the text undo of the browser.
  registerKey('grid', { key: 'z', mod: true }, undo);
  registerKey('grid', { key: 'y', mod: true }, redo);
  registerKey('grid', { key: 'z', mod: true, shift: true }, redo);
  update();
}
