// Fill handle, Ctrl+D and Ctrl+R (PLAN.md 4.2). Owner: fe-a.
//
// The fill rules (pattern, formula shift, series) are in the bridge
// (cmd/web/README.md, `fill`). This module finds the source range and the
// destination range and shows the handle and the preview.

import { HEADER_H, el, S, addrOf, showStatus, rawCall, api, mutate, on, registerKey } from './core.js';
import {
  rectOf, selRange, setSelection, pointToCell, registerGridMouseDown,
} from './grid.js';

const HANDLE = 7; // size of the handle in pixels

let drag = null; // { src, dst, x, y, raf } while the user drags the handle

const toRange = (R) => ({ row1: R.r1, col1: R.c1, row2: R.r2, col2: R.c2 });

// placeHandle puts the handle at the bottom-right corner of the selection.
// It is hidden during an edit and when the corner is not in the grid.
function placeHandle() {
  if (S.state !== 'editor' || !S.sheet || !S.rows || S.edit) {
    el.fillHandle.hidden = true;
    return;
  }
  const R = selRange();
  if (R.r2 > S.rowCount || R.c2 > S.colCount) {
    el.fillHandle.hidden = true;
    return;
  }
  const r = rectOf(R.r1, R.c1, R.r2, R.c2);
  if (r.width <= 0 || r.height <= 0) {
    el.fillHandle.hidden = true;
    return;
  }
  const half = Math.floor(HANDLE / 2);
  el.fillHandle.style.cssText =
    `left:${r.left + r.width - half - 1}px;top:${r.top + r.height - half - 1}px;width:${HANDLE}px;height:${HANDLE}px;`;
  el.fillHandle.hidden = false;
}

function hidePreview() {
  el.fillRange.hidden = true;
  delete el.fillRange.dataset.range;
}

// destinationFor returns the range that a drag to the pointer position
// fills: the source, extended on one axis (the axis with the larger pointer
// distance). It returns null when the pointer is in the source range.
function destinationFor(src, clientX, clientY) {
  const p = pointToCell(clientX, clientY);
  const box = el.cells.getBoundingClientRect();
  const x = clientX - box.left;
  const y = clientY - box.top;
  const r = rectOf(src.r1, src.c1, src.r2, src.c2);
  const dx = x < r.left ? r.left - x : x > r.left + r.width ? x - (r.left + r.width) : 0;
  const dy = y < r.top ? r.top - y : y > r.top + r.height ? y - (r.top + r.height) : 0;
  const rowOut = p.row < src.r1 || p.row > src.r2;
  const colOut = p.col < src.c1 || p.col > src.c2;
  if (!rowOut && !colOut) return null;
  const vertical = rowOut && (!colOut || dy >= dx);
  const dst = { ...src };
  if (vertical) {
    if (p.row < src.r1) dst.r1 = p.row;
    else dst.r2 = p.row;
  } else if (p.col < src.c1) {
    dst.c1 = p.col;
  } else {
    dst.c2 = p.col;
  }
  return dst;
}

function showPreview() {
  const dst = drag && drag.dst;
  if (!dst) {
    hidePreview();
    return;
  }
  const r = rectOf(dst.r1, dst.c1, dst.r2, dst.c2);
  el.fillRange.style.cssText = `left:${r.left - 1}px;top:${r.top - 1}px;width:${r.width + 1}px;height:${r.height + 1}px;`;
  el.fillRange.dataset.range = `${addrOf(dst.r1, dst.c1)}:${addrOf(dst.r2, dst.c2)}`;
  el.fillRange.hidden = false;
}

function dragUpdate() {
  if (!drag) return;
  drag.dst = destinationFor(drag.src, drag.x, drag.y);
  showPreview();
}

// dragTick scrolls the grid while the pointer is out of the grid.
function dragTick() {
  if (!drag) return;
  const g = el.grid.getBoundingClientRect();
  let dx = 0;
  let dy = 0;
  if (drag.y > g.bottom) dy = 30;
  else if (drag.y < g.top + HEADER_H) dy = -30;
  if (drag.x > g.right) dx = 60;
  else if (drag.x < g.left + S.rw) dx = -60;
  if (dx || dy) {
    el.grid.scrollTop += dy;
    el.grid.scrollLeft += dx;
    dragUpdate();
  }
  drag.raf = requestAnimationFrame(dragTick);
}

// runFill calls the bridge and selects the destination range.
function runFill(src, dst, series, select) {
  const res = mutate('fill', S.sheet, toRange(src), toRange(dst), series);
  if (!res.ok) return false;
  if (select) {
    setSelection({ row: dst.r1, col: dst.c1 }, { row: dst.r2, col: dst.c2 }, { scroll: false });
  }
  if (res.skipped) {
    showStatus(`Fill: ${res.skipped} spilled ${res.skipped === 1 ? 'cell is' : 'cells are'} read-only and did not change.`, 'info');
  }
  return true;
}

function endDrag(apply) {
  if (!drag) return;
  const { src, dst } = drag;
  cancelAnimationFrame(drag.raf);
  drag = null;
  hidePreview();
  delete el.grid.dataset.filling;
  if (apply && dst) runFill(src, dst, true, true);
  placeHandle();
}

function onHandleMouseDown(e) {
  if (e.target !== el.fillHandle) return false;
  // The grid keeps the focus, and the grid does not start a selection.
  e.preventDefault();
  if (S.edit || !S.sheet) return true;
  el.grid.focus({ preventScroll: true });
  drag = { src: selRange(), dst: null, x: e.clientX, y: e.clientY, raf: 0 };
  drag.raf = requestAnimationFrame(dragTick);
  el.grid.dataset.filling = 'true';
  return true;
}

const hasContent = (cell) => Boolean(cell) && (cell.type !== 'empty' || cell.formula);

// blockEnd returns the last row of the data block in the given column that
// continues below the source range, or 0 when the column has no such block.
function blockEnd(col, src) {
  if (col < 1 || col > S.colLimit) return 0;
  const cur = rawCall('getRange', [S.sheet, src.r2, col, src.r2 + 1, col]);
  if (!cur.ok) return 0;
  const cells = cur.cells || [];
  const here = cells.find((c) => c.row === src.r2);
  const below = cells.find((c) => c.row === src.r2 + 1);
  if (!hasContent(here) || !hasContent(below)) return 0;
  // findEdge is a milestone 2 function. api shows a message when it is not built yet.
  const res = api('findEdge', S.sheet, src.r2, col, 1, 0);
  return res.ok && res.row > src.r2 ? res.row : 0;
}

// A double click on the handle fills down to the last row of the adjacent
// data block: the column at the left, and if that is empty, the column at
// the right.
function onHandleDoubleClick(e) {
  // The grid must not open the editor for this double click.
  e.stopPropagation();
  e.preventDefault();
  if (S.edit || !S.sheet) return;
  const src = selRange();
  const end = blockEnd(src.c1 - 1, src) || blockEnd(src.c2 + 1, src);
  if (!end) return;
  runFill(src, { ...src, r2: Math.min(end, S.rowLimit) }, true, true);
}

// Ctrl+D copies the top row of the selection into its other rows. For a
// selection of one row, it copies from the row above. Ctrl+R does the same
// with columns.
function fillDown() {
  const R = selRange();
  if (R.r2 > R.r1) runFill({ ...R, r2: R.r1 }, R, false, false);
  else if (R.r1 > 1) runFill({ ...R, r1: R.r1 - 1, r2: R.r1 - 1 }, { ...R, r1: R.r1 - 1 }, false, false);
}

function fillRight() {
  const R = selRange();
  if (R.c2 > R.c1) runFill({ ...R, c2: R.c1 }, R, false, false);
  else if (R.c1 > 1) runFill({ ...R, c1: R.c1 - 1, c2: R.c1 - 1 }, { ...R, c1: R.c1 - 1 }, false, false);
}

export function init() {
  registerGridMouseDown(onHandleMouseDown);
  el.fillHandle.addEventListener('dblclick', onHandleDoubleClick);
  document.addEventListener('mousemove', (e) => {
    if (!drag) return;
    drag.x = e.clientX;
    drag.y = e.clientY;
    dragUpdate();
  });
  document.addEventListener('mouseup', () => endDrag(true));
  document.addEventListener('keydown', (e) => {
    if (drag && e.key === 'Escape') endDrag(false);
  });

  for (const event of ['selection', 'render', 'sheet', 'edit:start', 'edit:end']) on(event, placeHandle);
  on('render', showPreview);

  // The handlers return nothing, so the browser default (bookmark, reload)
  // does not run while the grid has the focus.
  registerKey('grid', { key: 'd', mod: true }, fillDown);
  registerKey('grid', { key: 'r', mod: true }, fillRight);
}
