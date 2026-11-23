// Column and row resize (PLAN.md 4.5). Owner: fe-a.
//
// grid.js makes the handles in the headers ([data-col-resize],
// [data-row-resize]). This module does the drag, the fit of a column to its
// content and the reset of a row height. The sizes are written with
// setColWidth and setRowHeight; one drag is one undo step.

import { HEADER_H, el, S, clamp, rawCall, mutate } from './core.js';
import {
  reloadMeta, rectOf, viewHeight, viewWidth, sheetInfo, mergeAt, selRange, isFullRows, isFullCols,
  registerGridMouseDown,
} from './grid.js';

const MIN_COL = 12;
const MAX_COL = 2000;
const MIN_ROW = 8;
const MAX_ROW = 1000;
const FIT_MAX_CELLS = 5000;
const CELL_PAD = 12; // 5 px padding at each side, the grid line, and 1 px of room

let drag = null; // { axis: 'col' | 'row', index, start, size, origin }

// targets returns the columns (or rows) that get the new size: all selected
// full columns (rows) when the handle is in that selection, else only the
// column (row) of the handle.
function targets(axis, index) {
  const R = selRange();
  if (axis === 'col' && isFullCols() && index >= R.c1 && index <= R.c2) return [R.c1, Math.min(R.c2, S.colLimit)];
  if (axis === 'row' && isFullRows() && index >= R.r1 && index <= R.r2) return [R.r1, Math.min(R.r2, S.rowLimit)];
  return [index, index];
}

// write sets the size. px = 0 removes the custom size. The axes are built
// again from sheetMeta, because the limits of the grid can change.
function write(axis, index, px) {
  const [from, to] = targets(axis, index);
  const res = mutate(axis === 'col' ? 'setColWidth' : 'setRowHeight', S.sheet, from, to, px);
  if (res.ok) reloadMeta();
}

function showGuide() {
  const g = el.resizeGuide;
  if (drag.axis === 'col') {
    // rectOf gives the place where the column shows, also in a frozen pane.
    const x = rectOf(1, drag.index, 1, drag.index).left + drag.size;
    g.style.cssText = `left:${x - 1}px;top:${el.grid.scrollTop}px;width:1px;height:${viewHeight()}px;`;
  } else {
    const y = rectOf(drag.index, 1, drag.index, 1).top + drag.size;
    g.style.cssText = `left:${el.grid.scrollLeft}px;top:${y - 1}px;width:${viewWidth()}px;height:1px;`;
  }
  g.dataset.axis = drag.axis;
  g.dataset.size = String(drag.size);
  g.hidden = false;
}

function endDrag(apply) {
  if (!drag) return;
  const d = drag;
  drag = null;
  el.resizeGuide.hidden = true;
  delete el.resizeGuide.dataset.size;
  delete el.grid.dataset.resizing;
  if (apply && d.size !== d.origin) write(d.axis, d.index, d.size);
}

// onMouseDown gets the mousedown before the grid. A press on a handle does
// not start a selection and does not start point mode.
function onMouseDown(e) {
  const col = e.target.closest('[data-col-resize]');
  const row = col ? null : e.target.closest('[data-row-resize]');
  if (!col && !row) return false;
  e.preventDefault();
  // During an edit the handles are not active. The focus stays in the editor.
  if (S.edit || !S.sheet) return true;
  el.grid.focus({ preventScroll: true });
  const axis = col ? 'col' : 'row';
  const index = Number(col ? col.dataset.colResize : row.dataset.rowResize);
  const origin = axis === 'col' ? S.cols.size(index) : S.rows.size(index);
  drag = { axis, index, start: axis === 'col' ? e.clientX : e.clientY, size: origin, origin, moved: false };
  el.grid.dataset.resizing = axis;
  return true;
}

function onMouseMove(e) {
  if (!drag) return;
  const delta = (drag.axis === 'col' ? e.clientX : e.clientY) - drag.start;
  if (!drag.moved && Math.abs(delta) < 2) return; // a click or a double click, not a drag
  drag.moved = true;
  const raw = Math.round(drag.origin + delta);
  drag.size = drag.axis === 'col' ? clamp(raw, MIN_COL, MAX_COL) : clamp(raw, MIN_ROW, MAX_ROW);
  showGuide();
}

// ------------------------------------------------------------------- fit

let measureCtx = null;

// fitWidth returns the width in pixels that shows the content of a column,
// or 0 when the column has no content (the default width then applies).
// Only the browser can measure text, so this is the one place where layout
// logic is not in the bridge.
function fitWidth(col) {
  const info = sheetInfo();
  if (!info || info.maxRow < 1 || col > info.maxCol) return 0;
  const res = rawCall('getRange', [S.sheet, 1, col, info.maxRow, col]);
  if (!res.ok) return -1;
  if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d');
  if (!measureCtx) return -1;
  const style = getComputedStyle(el.cells);
  const size = style.fontSize || '13px';
  const family = style.fontFamily || 'sans-serif';
  let widest = 0;
  let count = 0;
  for (const cell of res.cells || []) {
    const text = cell.display == null ? '' : String(cell.display);
    if (!text) continue;
    if (mergeAt(cell.row, cell.col)) continue; // merged cells are ignored, as in Excel
    if (++count > FIT_MAX_CELLS) break;
    const st = cell.style || {};
    // An error cell shows in a heavier font (style.css, .t-error).
    const weight = st.bold ? '700' : cell.type === 'error' ? '600' : '400';
    measureCtx.font = `${st.italic ? 'italic ' : ''}${weight} ${size} ${family}`;
    widest = Math.max(widest, measureCtx.measureText(text).width);
  }
  if (!widest) return 0;
  return clamp(Math.ceil(widest) + CELL_PAD, MIN_COL, MAX_COL);
}

function onDoubleClick(e) {
  const col = e.target.closest('[data-col-resize]');
  const row = col ? null : e.target.closest('[data-row-resize]');
  if (!col && !row) return;
  e.preventDefault();
  e.stopPropagation();
  if (S.edit || !S.sheet) return;
  if (col) {
    const index = Number(col.dataset.colResize);
    const px = fitWidth(index);
    if (px >= 0) write('col', index, px);
  } else {
    // The grid shows one line of text, so the default height always fits.
    write('row', Number(row.dataset.rowResize), 0);
  }
}

export function init() {
  registerGridMouseDown(onMouseDown);
  document.addEventListener('mousemove', onMouseMove);
  document.addEventListener('mouseup', () => endDrag(true));
  document.addEventListener('keydown', (e) => {
    if (drag && e.key === 'Escape') endDrag(false);
  });
  el.colHeaders.addEventListener('dblclick', onDoubleClick);
  el.rowHeaders.addEventListener('dblclick', onDoubleClick);
}
