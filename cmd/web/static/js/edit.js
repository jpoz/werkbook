// Editing: the in-cell editor, the formula bar, formula point mode, commit
// and cancel.

import { el, S, clamp, colName, addrOf, showStatus, mutate, on, emit, runKey } from './core.js';
import {
  render, rectOf, viewWidth, cellAt, mergeAt, updateBars, ensureExtent, scrollToCell, pointToCell,
  moveActive, enterMove, tabMove, dragTick,
} from './grid.js';
import { showSheet } from './sheets.js';

export function isReadOnly(cell) {
  return Boolean(cell && cell.spill && !cell.formula);
}

// activeInput is the input that the user edits in: the in-cell editor or the formula bar.
export function activeInput() {
  return S.edit && S.edit.source === 'bar' ? el.formula : el.editor;
}

export function formulaEditing() {
  return Boolean(S.edit) && activeInput().value.startsWith('=');
}

// editRect is the rectangle of the edited cell. Call it only while the
// sheet of the edited cell is on screen.
function editRect() {
  const ed = S.edit;
  const m = mergeAt(ed.row, ed.col);
  return m ? rectOf(m.row1, m.col1, m.row2, m.col2) : rectOf(ed.row, ed.col, ed.row, ed.col);
}

function placeEditor() {
  const a = editRect();
  el.editor.style.left = `${a.left - 1}px`;
  el.editor.style.top = `${a.top - 1}px`;
  el.editor.style.height = `${a.height + 1}px`;
  el.editor.hidden = false;
  sizeEditor();
}

function sizeEditor() {
  if (!S.edit || el.editor.hidden || S.sheet !== S.edit.sheet) return;
  const a = editRect();
  el.editor.style.width = `${a.width + 1}px`;
  if (el.editor.scrollWidth > el.editor.clientWidth) {
    const room = Math.max(a.width + 1, el.grid.scrollLeft + viewWidth() - a.left);
    el.editor.style.width = `${Math.min(room, el.editor.scrollWidth + 16)}px`;
  }
}

// startEdit opens the editor on the active cell. mode is 'type' (the user
// types a new value), 'edit' (F2 or double click) or 'bar' (formula bar).
export function startEdit(mode, initial) {
  if (S.edit) return true;
  if (S.state !== 'editor' || !S.sheet) return false;
  const { row, col } = S.anchor;
  const cell = cellAt(row, col);
  if (isReadOnly(cell)) {
    showStatus(`${addrOf(row, col)} is part of a spilled array and is read-only. Edit the cell that has the formula.`, 'error');
    return false;
  }
  const original = cell && cell.input != null ? String(cell.input) : '';
  const value = initial !== undefined ? initial : original;
  S.edit = {
    sheet: S.sheet,
    row,
    col,
    original,
    mode,
    source: mode === 'bar' ? 'bar' : 'cell',
    keyNav: mode === 'type', // the arrow keys point at cells (true) or move the caret (false)
    point: null, // the reference that point mode inserted, while the user has not typed
    sig: '', // text and caret at the last 'edit:change' event
  };
  ensureExtent(row, col);
  scrollToCell(row, col);
  render(false);
  el.editor.value = value;
  el.formula.value = value;
  el.grid.dataset.editing = 'true';
  placeEditor();
  if (mode !== 'bar') {
    el.editor.focus({ preventScroll: true });
    el.editor.setSelectionRange(value.length, value.length);
  }
  emit('edit:start');
  refreshPoint();
  return true;
}

export function endEdit() {
  if (!S.edit) return;
  S.edit = null;
  el.editor.hidden = true;
  delete el.grid.dataset.editing;
  refreshPoint();
  emit('edit:end');
}

// backToEditSheet shows the sheet of the edited cell again, if point mode
// went to a different sheet.
function backToEditSheet(ed) {
  if (S.sheet !== ed.sheet && S.sheets.some((s) => s.name === ed.sheet)) showSheet(ed.sheet);
}

export function cancelEdit() {
  const ed = S.edit;
  endEdit();
  if (ed) backToEditSheet(ed);
  updateBars();
  el.grid.focus({ preventScroll: true });
}

// setEditText replaces the text of the open edit in both inputs and sets
// the caret (or the selection) in the active input. It emits 'edit:change'.
// point is optional: { start, end } is the new position of the reference
// that point mode inserted, so that point mode stays valid for it (F4). With
// no point argument, the inserted reference is not tracked after the call.
export function setEditText(text, selStart, selEnd, point) {
  const ed = S.edit;
  if (!ed) return;
  const end = selEnd === undefined ? selStart : selEnd;
  el.editor.value = text;
  el.formula.value = text;
  activeInput().setSelectionRange(selStart, end);
  if (point && ed.point) {
    ed.point.start = point.start;
    ed.point.end = point.end;
    ed.point.text = text;
  } else {
    ed.point = null;
  }
  sizeEditor();
  refreshPoint();
}

// ------------------------------------------------------ formula point mode
//
// While the user edits a formula and the caret is at a position where a
// reference can go, a click, a drag or an arrow key on the grid puts the
// address into the formula. The selection of the sheet does not change.

const POINT_AFTER = '=(,:+-*/^&<>';

// pointValid reports whether the inserted reference is still the last
// change: no typing and no caret move after it.
function pointValid() {
  const ed = S.edit;
  if (!ed || !ed.point) return false;
  const inp = activeInput();
  return inp.value === ed.point.text && inp.selectionStart === ed.point.end && inp.selectionEnd === ed.point.end;
}

export function pointActive() {
  if (!S.edit) return false;
  const inp = activeInput();
  const text = inp.value;
  if (text[0] !== '=') return false;
  if (pointValid()) return true;
  const caret = inp.selectionStart;
  if (caret == null || caret < 1) return false;
  let quotes = 0;
  for (let i = 0; i < caret; i++) if (text[i] === '"') quotes++;
  if (quotes % 2) return false; // the caret is in a string literal
  let i = caret - 1;
  while (i > 0 && text[i] === ' ') i--;
  return POINT_AFTER.includes(text[i]);
}

function sheetPrefix(name) {
  // A name that looks like a cell address (for example "AB12") also needs quotes.
  const plain = /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name);
  return plain ? `${name}!` : `'${name.replace(/'/g, "''")}'!`;
}

function pointRefText(kind, a, f) {
  const prefix = S.sheet !== S.edit.sheet ? sheetPrefix(S.sheet) : '';
  const r1 = Math.min(a.row, f.row);
  const r2 = Math.max(a.row, f.row);
  const c1 = Math.min(a.col, f.col);
  const c2 = Math.max(a.col, f.col);
  if (kind === 'cols') return `${prefix}${colName(c1)}:${colName(c2)}`;
  if (kind === 'rows') return `${prefix}${r1}:${r2}`;
  if (r1 === r2 && c1 === c2) return prefix + addrOf(r1, c1);
  return `${prefix}${addrOf(r1, c1)}:${addrOf(r2, c2)}`;
}

// pointInsert puts a reference at the caret. If the last change was also
// an inserted reference, the new reference replaces it.
function pointInsert(kind, anchor, focus) {
  const ed = S.edit;
  const inp = activeInput();
  const text = inp.value;
  const valid = pointValid();
  const start = valid ? ed.point.start : inp.selectionStart;
  const end = valid ? ed.point.end : inp.selectionEnd;
  const ref = pointRefText(kind, anchor, focus);
  const next = text.slice(0, start) + ref + text.slice(end);
  const caret = start + ref.length;
  ed.point = { start, end: caret, text: next, sheet: S.sheet, kind, anchor: { ...anchor }, focus: { ...focus } };
  el.editor.value = next;
  el.formula.value = next;
  inp.setSelectionRange(caret, caret);
  sizeEditor();
  refreshPoint();
}

// drawPoint shows the dashed outline of the inserted reference.
function drawPoint() {
  const p = S.edit && S.edit.point;
  if (!p || p.sheet !== S.sheet || !S.rows) {
    el.pointRange.hidden = true;
    return;
  }
  let r1 = Math.min(p.anchor.row, p.focus.row);
  let r2 = Math.max(p.anchor.row, p.focus.row);
  let c1 = Math.min(p.anchor.col, p.focus.col);
  let c2 = Math.max(p.anchor.col, p.focus.col);
  if (p.kind === 'cols') {
    r1 = 1;
    r2 = S.rowCount;
  } else if (p.kind === 'rows') {
    c1 = 1;
    c2 = S.colCount;
  } else if (r1 === r2 && c1 === c2) {
    const m = mergeAt(r1, c1);
    if (m) {
      r2 = m.row2;
      c2 = m.col2;
    }
  }
  const r = rectOf(r1, c1, r2, c2);
  el.pointRange.style.cssText = `left:${r.left - 1}px;top:${r.top - 1}px;width:${r.width + 1}px;height:${r.height + 1}px;`;
  el.pointRange.dataset.ref = pointRefText(p.kind, p.anchor, p.focus);
  el.pointRange.hidden = false;
}

// refreshPoint brings the point mode mark and the outline up to date after
// each change of the text or the caret. It emits 'edit:change' when the
// text, the caret, the active input or the sheet on screen changed.
function refreshPoint() {
  const ed = S.edit;
  if (ed && ed.point && !pointValid()) ed.point = null;
  if (ed && pointActive()) el.grid.dataset.pointMode = 'true';
  else delete el.grid.dataset.pointMode;
  drawPoint();
  if (!ed) return;
  const inp = activeInput();
  const sig = `${ed.source}\n${S.sheet}\n${inp.selectionStart}\n${inp.selectionEnd}\n${inp.value}`;
  if (sig !== ed.sig) {
    ed.sig = sig;
    emit('edit:change');
  }
}

// pointMouseDown is the mousedown of the grid while point mode is active.
// grid.js calls it.
export function pointMouseDown(e) {
  const ed = S.edit;
  if (e.target.closest('#corner')) return;
  const prev = pointValid() && ed.point.sheet === S.sheet ? ed.point : null;
  let kind = 'cells';
  let hit;
  const colHeader = e.target.closest('[data-col-header]');
  const rowHeader = e.target.closest('[data-row-header]');
  if (colHeader) {
    kind = 'cols';
    hit = { row: 1, col: Number(colHeader.dataset.colHeader) };
  } else if (rowHeader) {
    kind = 'rows';
    hit = { row: Number(rowHeader.dataset.rowHeader), col: 1 };
  } else if (e.target.closest('#col-headers, #row-headers')) {
    return;
  } else {
    hit = pointToCell(e.clientX, e.clientY);
  }
  let anchor = hit;
  if (e.shiftKey && prev && prev.kind === kind) {
    anchor = prev.anchor;
  } else if (kind === 'cells') {
    const m = mergeAt(hit.row, hit.col);
    if (m) hit = anchor = { row: m.row1, col: m.col1 };
  }
  pointInsert(kind, anchor, hit);
  S.drag = { x: e.clientX, y: e.clientY, raf: requestAnimationFrame(dragTick), point: kind };
}

// pointDragUpdate is the drag step of point mode. grid.js calls it.
export function pointDragUpdate() {
  const ed = S.edit;
  if (!ed || !ed.point || !pointValid() || ed.point.sheet !== S.sheet) return;
  const kind = S.drag.point;
  const p = pointToCell(S.drag.x, S.drag.y);
  const a = ed.point.anchor;
  let next = p;
  if (kind === 'cols') next = { row: 1, col: p.col };
  else if (kind === 'rows') next = { row: p.row, col: 1 };
  else {
    const m = mergeAt(p.row, p.col);
    if (m && m === mergeAt(a.row, a.col)) next = a; // the pointer is still in the merged cell where the drag started
  }
  const f = ed.point.focus;
  if (next.row !== f.row || next.col !== f.col) pointInsert(kind, a, next);
}

// pointKey moves the inserted reference with an arrow key. With Shift it
// extends the reference to a range.
function pointKey(d, extend) {
  const ed = S.edit;
  const prev = pointValid() && ed.point.sheet === S.sheet && ed.point.kind === 'cells' ? ed.point : null;
  const base = prev ? prev.focus : S.sheet === ed.sheet ? { row: ed.row, col: ed.col } : S.anchor;
  const m = extend ? null : mergeAt(base.row, base.col);
  let row = base.row;
  let col = base.col;
  if (d[0] > 0) row = (m ? m.row2 : row) + 1;
  else if (d[0] < 0) row = (m ? m.row1 : row) - 1;
  if (d[1] > 0) col = (m ? m.col2 : col) + 1;
  else if (d[1] < 0) col = (m ? m.col1 : col) - 1;
  let next = { row: clamp(row, 1, S.rowLimit), col: clamp(col, 1, S.colLimit) };
  let anchor;
  if (extend && prev) {
    anchor = prev.anchor;
  } else {
    const nm = mergeAt(next.row, next.col);
    if (nm) next = { row: nm.row1, col: nm.col1 };
    anchor = next;
  }
  ensureExtent(next.row, next.col);
  scrollToCell(next.row, next.col);
  render(false);
  pointInsert('cells', anchor, next);
}

// pointShowSheet shows a different sheet while the formula edit stays open.
// sheets.js calls it for a click on a sheet tab during a formula edit.
export function pointShowSheet(name) {
  const ed = S.edit;
  if (!ed || name === S.sheet) return;
  if (name !== ed.sheet && ed.source !== 'bar') {
    // The in-cell editor is not on this sheet, so the edit continues in the
    // formula bar. Move the focus before the in-cell editor hides: a blur
    // to a different element commits the edit.
    const start = el.editor.selectionStart;
    const end = el.editor.selectionEnd;
    el.formula.value = el.editor.value;
    el.formula.focus({ preventScroll: true });
    el.formula.setSelectionRange(start, end);
    ed.source = 'bar';
  }
  showSheet(name, true);
}

// ---------------------------------------------------------- commit and keys

// commitEdit writes the edit. It returns false if the write failed. After a
// failed write the editor stays open, unless the commit came from a blur.
export function commitEdit(after, fromBlur) {
  const ed = S.edit;
  if (!ed) return true;
  const value = ed.source === 'bar' ? el.formula.value : el.editor.value;
  if (value !== ed.original) {
    // mutate emits 'write': grid.js sets the dirty mark and paints again.
    const res = mutate('setCell', ed.sheet, ed.row, ed.col, value);
    if (!res.ok) {
      if (fromBlur) {
        endEdit();
        backToEditSheet(ed);
        updateBars();
      }
      return false;
    }
  }
  endEdit();
  backToEditSheet(ed);
  updateBars();
  if (after) after();
  if (!fromBlur) el.grid.focus({ preventScroll: true });
  return true;
}

function onEditBlur(e) {
  if (!S.edit) return;
  const to = e.relatedTarget;
  if (to === el.editor || to === el.formula) return;
  commitEdit(null, true);
}

function onEditKey(e) {
  if (e.isComposing) return;
  // A registered key handler has priority over the default keys.
  if (S.edit && runKey('edit', e)) {
    e.stopPropagation();
    return;
  }
  const inBar = e.target === el.formula;
  if (e.key === 'Enter' || e.key === 'Tab') {
    e.preventDefault();
    e.stopPropagation();
    if (!S.edit) {
      if (!inBar || !startEdit('bar', el.formula.value)) {
        el.grid.focus({ preventScroll: true });
        return;
      }
      S.edit.original = null; // the value did not come from an input event, so always write it
    }
    S.edit.source = inBar ? 'bar' : 'cell';
    const back = e.shiftKey;
    commitEdit(e.key === 'Enter' ? () => enterMove(back) : () => tabMove(back), false);
  } else if (e.key === 'Escape') {
    e.preventDefault();
    e.stopPropagation();
    cancelEdit();
  } else if (e.key === 'F2' && S.edit) {
    // F2 changes what the arrow keys do: point at cells, or move the caret.
    e.preventDefault();
    e.stopPropagation();
    S.edit.keyNav = !S.edit.keyNav;
  } else if (S.edit && S.edit.keyNav && e.key.startsWith('Arrow') && !e.ctrlKey && !e.metaKey && !e.altKey) {
    const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
    if (!d) return;
    if (pointActive()) {
      e.preventDefault();
      e.stopPropagation();
      pointKey(d, e.shiftKey);
    } else if (!inBar && !e.shiftKey) {
      // The user typed a new value: the arrow keys commit and move, as Enter and Tab do.
      e.preventDefault();
      e.stopPropagation();
      commitEdit(() => moveActive(d[0], d[1]), false);
    }
  }
}

export function init() {
  // The outline of the inserted reference is drawn again after each render.
  on('render', () => {
    drawPoint();
    // With frozen panes, the place where the edited cell shows depends on
    // the scroll position.
    if (S.edit && !el.editor.hidden && S.sheet === S.edit.sheet && (S.freeze.rows || S.freeze.cols)) placeEditor();
  });
  // A sheet came on screen while an edit is open (point mode on a different
  // sheet). The in-cell editor shows only on the sheet that has the edited cell.
  on('sheet', () => {
    if (!S.edit) return;
    if (S.sheet === S.edit.sheet) placeEditor();
    else el.editor.hidden = true;
    refreshPoint();
  });

  el.editor.addEventListener('input', () => {
    if (!S.edit) return;
    S.edit.source = 'cell';
    S.edit.point = null;
    el.formula.value = el.editor.value;
    sizeEditor();
    refreshPoint();
  });
  // A click in the text makes the arrow keys move the caret.
  for (const input of [el.editor, el.formula]) {
    input.addEventListener('mousedown', () => {
      if (S.edit) S.edit.keyNav = false;
    });
    // Not all browsers send selectionchange for a caret move in an input.
    for (const type of ['keyup', 'mouseup']) {
      input.addEventListener(type, () => {
        if (S.edit) refreshPoint();
      });
    }
  }
  document.addEventListener('selectionchange', () => {
    if (S.edit) refreshPoint();
  });
  el.editor.addEventListener('focus', () => {
    if (S.edit) S.edit.source = 'cell';
  });
  el.editor.addEventListener('blur', onEditBlur);
  el.editor.addEventListener('keydown', onEditKey);

  el.formula.addEventListener('focus', () => {
    if (S.edit) {
      S.edit.source = 'bar';
      S.edit.mode = 'bar';
      return;
    }
    if (!startEdit('bar')) el.grid.focus({ preventScroll: true });
  });
  el.formula.addEventListener('input', () => {
    if (!S.edit && !startEdit('bar', el.formula.value)) {
      updateBars();
      return;
    }
    S.edit.source = 'bar';
    S.edit.point = null;
    el.editor.value = el.formula.value;
    sizeEditor();
    refreshPoint();
  });
  el.formula.addEventListener('blur', onEditBlur);
  el.formula.addEventListener('keydown', onEditKey);
}
