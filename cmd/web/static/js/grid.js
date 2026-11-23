// Grid: layout, render, headers, selection, scroll, the mouse and key
// handlers of the grid, and the name box.

import {
  MAX_ROWS, MAX_COLS, DEF_ROW_H, DEF_COL_W, HEADER_H, MIN_ROWS, MIN_COLS, ROW_GROW, COL_GROW,
  ROW_BLOCK, COL_BLOCK, MAX_CANVAS_PX,
  el, S, darkQuery, clamp, cellKey, colName, addrOf, parseRef, makeAxis, luminance, isHexColor,
  showStatus, rawCall, api, on, emit, runKey,
} from './core.js';
import { setDirty } from './workbook.js';
import { showSheet, refreshSheets } from './sheets.js';
import { startEdit, commitEdit, pointActive, pointMouseDown, pointDragUpdate, isReadOnly } from './edit.js';
import { clearSelection } from './clipboard.js';

const mouseDownHandlers = [];
const headerMenuHandlers = [];

// registerGridMouseDown adds a handler that gets the mousedown event (left
// button) before the grid. It returns true when it used the event; the grid
// then does nothing with the event.
export function registerGridMouseDown(handler) {
  mouseDownHandlers.push(handler);
}

// registerHeaderMenu adds a handler for the contextmenu event of a row
// header or a column header: handler({ kind: 'row' | 'col', index, event }).
// The handler calls event.preventDefault() when it shows a menu.
export function registerHeaderMenu(handler) {
  headerMenuHandlers.push(handler);
}

export function sheetInfo() {
  return S.sheets.find((s) => s.name === S.sheet) || null;
}

// ---------------------------------------------------------------- grid layout

export const viewHeight = () => Math.max(0, el.grid.clientHeight - HEADER_H);
export const viewWidth = () => Math.max(0, el.grid.clientWidth - S.rw);

// readMeta reads sheetMeta of the sheet on screen and builds the axes, the
// merged ranges and the grid limits from it. It does not paint.
export function readMeta() {
  const meta = api('sheetMeta', S.sheet);
  S.merges = [];
  if (meta.ok && Array.isArray(meta.merges)) {
    for (const m of meta.merges) {
      if (m && m.row1 >= 1 && m.col1 >= 1 && m.row2 >= m.row1 && m.col2 >= m.col1) S.merges.push(m);
    }
  }
  S.rows = makeAxis(DEF_ROW_H, meta.ok ? meta.rowHeights : null, MAX_ROWS);
  S.cols = makeAxis(DEF_COL_W, meta.ok ? meta.colWidths : null, MAX_COLS);
  S.rowLimit = S.rows.offset(MAX_ROWS + 1) <= MAX_CANVAS_PX ? MAX_ROWS : Math.max(MIN_ROWS, S.rows.indexAt(MAX_CANVAS_PX) - 1);
  S.colLimit = S.cols.offset(MAX_COLS + 1) <= MAX_CANVAS_PX ? MAX_COLS : Math.max(MIN_COLS, S.cols.indexAt(MAX_CANVAS_PX) - 1);
  const fr = meta.ok && meta.freeze ? meta.freeze : null;
  S.freeze = { rows: fr && fr.rows > 0 ? fr.rows : 0, cols: fr && fr.cols > 0 ? fr.cols : 0 };
  S.warnings = meta.ok && Array.isArray(meta.warnings) ? meta.warnings : [];
}

// reloadMeta reads sheetMeta again after a change of the layout (column
// width, row height, merged ranges, freeze) and paints the sheet again. The
// scroll position and the selection stay.
export function reloadMeta() {
  if (S.state !== 'editor' || !S.sheet) return;
  readMeta();
  S.rowCount = Math.min(S.rowCount, S.rowLimit);
  S.colCount = Math.min(S.colCount, S.colLimit);
  layout();
  render(true);
  updateSelectionUI();
  updateBars();
  emit('sheet', S.sheet);
}

export function layout() {
  S.rw = Math.max(44, String(S.rowCount).length * 8 + 18);
  const w = S.cols.offset(S.colCount + 1);
  const h = S.rows.offset(S.rowCount + 1);
  el.grid.style.setProperty('--rw', `${S.rw}px`);
  el.grid.style.setProperty('--hh', `${HEADER_H}px`);
  el.canvas.style.width = `${S.rw + w}px`;
  el.canvas.style.height = `${HEADER_H + h}px`;
  el.rowHeaders.style.height = `${h}px`;
  for (const layer of [el.cells, el.overlay]) {
    layer.style.width = `${w}px`;
    layer.style.height = `${h}px`;
  }
  // Frozen panes. The attributes show the freeze of the sheet; a pane is in
  // use only when its axis has frozen rows or columns.
  const fr = frozenRows();
  const fc = frozenCols();
  el.grid.dataset.frozenRows = String(S.freeze.rows);
  el.grid.dataset.frozenCols = String(S.freeze.cols);
  el.paneTop.hidden = !fr;
  el.paneLeft.hidden = !fc;
  el.paneCorner.hidden = !(fr && fc);
  S.rendered = null;
}

// ensureExtent grows the grid so that the given cell is in it.
export function ensureExtent(row, col) {
  let changed = false;
  if (row > S.rowCount - 20 && S.rowCount < S.rowLimit) {
    S.rowCount = Math.min(S.rowLimit, Math.max(S.rowCount, row + 50));
    changed = true;
  }
  if (col > S.colCount - 3 && S.colCount < S.colLimit) {
    S.colCount = Math.min(S.colLimit, Math.max(S.colCount, col + 10));
    changed = true;
  }
  if (changed) layout();
}

// growIfNeeded keeps the grid at least one screen of rows and columns
// larger than the viewport edge. render calls it, so it runs on each
// scroll, window resize and sheet change. The grid does not become smaller.
function growIfNeeded() {
  let changed = false;
  const vh = Math.max(viewHeight(), DEF_ROW_H);
  const wantRow = Math.min(S.rowLimit, S.rows.indexAt(el.grid.scrollTop + 2 * vh) + 1);
  if (wantRow > S.rowCount) {
    S.rowCount = Math.min(S.rowLimit, Math.max(wantRow, S.rowCount + ROW_GROW));
    changed = true;
  }
  const vw = Math.max(viewWidth(), DEF_COL_W);
  const wantCol = Math.min(S.colLimit, S.cols.indexAt(el.grid.scrollLeft + 2 * vw) + 1);
  if (wantCol > S.colCount) {
    S.colCount = Math.min(S.colLimit, Math.max(wantCol, S.colCount + COL_GROW));
    changed = true;
  }
  if (changed) layout();
}

export function visibleRange() {
  const top = el.grid.scrollTop;
  const left = el.grid.scrollLeft;
  const firstRow = S.rows.indexAt(top);
  const lastRow = S.rows.indexAt(top + viewHeight());
  const firstCol = S.cols.indexAt(left);
  const lastCol = S.cols.indexAt(left + viewWidth());
  return {
    r1: clamp(firstRow - ((firstRow - 1) % ROW_BLOCK) - ROW_BLOCK, 1, S.rowCount),
    r2: clamp(lastRow + (ROW_BLOCK - (lastRow % ROW_BLOCK)) + ROW_BLOCK, 1, S.rowCount),
    c1: clamp(firstCol - ((firstCol - 1) % COL_BLOCK) - COL_BLOCK, 1, S.colCount),
    c2: clamp(lastCol + (COL_BLOCK - (lastCol % COL_BLOCK)) + COL_BLOCK, 1, S.colCount),
  };
}

// ---------------------------------------------------------------- rendering

function makeCellNode(row, col, left, top, width, height, cell) {
  const node = document.createElement('div');
  node.className = 'cell';
  node.dataset.row = row;
  node.dataset.col = col;
  node.dataset.cell = addrOf(row, col);
  let css = `left:${left}px;top:${top}px;width:${width}px;height:${height}px;line-height:${Math.max(1, height - 1)}px;`;
  if (cell) {
    node.textContent = cell.display == null ? '' : String(cell.display);
    node.dataset.type = cell.type;
    if (cell.type === 'number') node.className = 'cell t-number';
    else if (cell.type === 'bool') node.className = 'cell t-bool';
    else if (cell.type === 'error') {
      node.className = 'cell t-error';
      node.title = `Error: ${cell.display}`;
    }
    if (cell.formula) node.dataset.formula = 'true';
    if (cell.spill) node.dataset.spill = 'true';
    const st = cell.style;
    if (st) {
      const fill = isHexColor(st.fill) ? st.fill : null;
      let color = isHexColor(st.color) ? st.color : null;
      if (fill) {
        css += `background-color:${fill};`;
        // A fill color comes from the file, so the default text color of the theme can be unreadable on it.
        if (!color && cell.type !== 'error') color = luminance(fill) > 0.5 ? '#1f2328' : '#f2f4f6';
      } else if (color && darkQuery.matches && luminance(color) < 0.4) {
        color = null; // dark text from the file is not readable on the dark sheet
      }
      if (color) css += `color:${color};`;
      if (st.bold) css += 'font-weight:700;';
      if (st.italic) css += 'font-style:italic;';
      if (st.underline) css += 'text-decoration:underline;';
      if (st.align === 'left' || st.align === 'center' || st.align === 'right') css += `text-align:${st.align};`;
    }
  }
  node.style.cssText = css;
  return node;
}

// render builds the cell nodes of the visible range. It returns true when
// it built them, and false when the range on screen did not change.
export function render(force) {
  if (S.state !== 'editor' || !S.sheet) return false;
  growIfNeeded();
  const R = visibleRange();
  const prev = S.rendered;
  if (!force && prev && prev.r1 === R.r1 && prev.r2 === R.r2 && prev.c1 === R.c1 && prev.c2 === R.c2) return false;

  // Frozen rows and columns. Only the part that fits in the view is built.
  const fr = frozenRows();
  const fc = frozenCols();
  R.fr = fr ? Math.min(fr, S.rows.indexAt(viewHeight()) + 1) : 0;
  R.fc = fc ? Math.min(fc, S.cols.indexAt(viewWidth()) + 1) : 0;
  S.rendered = R;

  // Segments of rows and columns that are built: [first, last, frozen].
  const rowSegs = [];
  const colSegs = [];
  if (R.fr) rowSegs.push([1, R.fr, true]);
  if (Math.max(R.r1, fr + 1) <= R.r2) rowSegs.push([Math.max(R.r1, fr + 1), R.r2, false]);
  if (R.fc) colSegs.push([1, R.fc, true]);
  if (Math.max(R.c1, fc + 1) <= R.c2) colSegs.push([Math.max(R.c1, fc + 1), R.c2, false]);
  const inSegs = (segs, i) => segs.some((s) => i >= s[0] && i <= s[1]);

  const res = api('getRange', S.sheet, R.r1, R.c1, R.r2, R.c2);
  const cells = new Map();
  const addCells = (list) => {
    if (Array.isArray(list)) for (const cell of list) cells.set(cellKey(cell.row, cell.col), cell);
  };
  if (res.ok) addCells(res.cells);
  const addRange = (r1, c1, r2, c2) => {
    const more = rawCall('getRange', [S.sheet, r1, c1, r2, c2]);
    if (more.ok) addCells(more.cells);
  };
  // The frozen panes can be out of the range R when the grid is scrolled.
  if (R.fr) addRange(1, R.c1, R.fr, R.c2);
  if (R.fc) addRange(R.r1, 1, R.r2, R.fc);
  if (R.fr && R.fc) addRange(1, 1, R.fr, R.fc);
  S.cells = cells;

  // Merged ranges: one node at the top-left cell covers the full range.
  const covered = new Set();
  const merges = [];
  for (const m of S.merges) {
    let hit = false;
    for (const rs of rowSegs) {
      const rA = Math.max(m.row1, rs[0]);
      const rB = Math.min(m.row2, rs[1]);
      if (rA > rB) continue;
      for (const cs of colSegs) {
        const cA = Math.max(m.col1, cs[0]);
        const cB = Math.min(m.col2, cs[1]);
        if (cA > cB) continue;
        hit = true;
        for (let r = rA; r <= rB; r++) {
          for (let c = cA; c <= cB; c++) covered.add(cellKey(r, c));
        }
      }
    }
    if (hit) merges.push(m);
  }

  // One fragment for each region. A cell node is in exactly one of #cells,
  // #pane-top, #pane-left and #pane-corner.
  const frags = {
    body: document.createDocumentFragment(),
    top: document.createDocumentFragment(),
    left: document.createDocumentFragment(),
    corner: document.createDocumentFragment(),
  };
  const region = (rowFrozen, colFrozen) => (rowFrozen ? (colFrozen ? 'corner' : 'top') : colFrozen ? 'left' : 'body');
  const markEdges = (node, lastRow, lastCol) => {
    // The line at the end of the frozen rows and columns.
    if (fr && lastRow === fr) node.classList.add('freeze-edge-row');
    if (fc && lastCol === fc) node.classList.add('freeze-edge-col');
  };

  for (const rs of rowSegs) {
    let y = S.rows.offset(rs[0]);
    for (let r = rs[0]; r <= rs[1]; r++) {
      const h = S.rows.size(r);
      if (h > 0) {
        for (const cs of colSegs) {
          const frag = frags[region(rs[2], cs[2])];
          let x = S.cols.offset(cs[0]);
          for (let c = cs[0]; c <= cs[1]; c++) {
            const w = S.cols.size(c);
            const key = cellKey(r, c);
            if (w > 0 && !covered.has(key)) {
              const node = makeCellNode(r, c, x, y, w, h, cells.get(key));
              markEdges(node, r, c);
              frag.append(node);
            }
            x += w;
          }
        }
      }
      y += h;
    }
  }
  for (const m of merges) {
    const rect = absRect(m.row1, m.col1, m.row2, m.col2);
    const rowFrozen = m.row1 <= fr;
    const colFrozen = m.col1 <= fc;
    // A merged range that crosses the freeze line is cut at the line.
    if (rowFrozen && m.row2 > fr) rect.height = S.rows.offset(fr + 1) - rect.top;
    if (colFrozen && m.col2 > fc) rect.width = S.cols.offset(fc + 1) - rect.left;
    if (rect.width <= 0 || rect.height <= 0) continue;
    let cell = cells.get(cellKey(m.row1, m.col1));
    if (!cell && !(inSegs(rowSegs, m.row1) && inSegs(colSegs, m.col1))) cell = fetchCell(m.row1, m.col1);
    const node = makeCellNode(m.row1, m.col1, rect.left, rect.top, rect.width, rect.height, cell);
    node.dataset.merge = `${addrOf(m.row1, m.col1)}:${addrOf(m.row2, m.col2)}`;
    markEdges(node, rowFrozen ? Math.min(m.row2, fr) : m.row2, colFrozen ? Math.min(m.col2, fc) : m.col2);
    frags[region(rowFrozen, colFrozen)].append(node);
  }
  el.cells.replaceChildren(frags.body);
  el.paneTop.replaceChildren(frags.top);
  el.paneLeft.replaceChildren(frags.left);
  el.paneCorner.replaceChildren(frags.corner);
  renderHeaders();
  // The overlays depend on the scroll position when a freeze is on.
  if (fr || fc) placeSelection();
  // edit.js draws the point-mode outline on this event.
  emit('render', R);
  return true;
}

export function renderHeaders() {
  const R = S.rendered;
  if (!R) return;
  const sel = selRange();
  const fr = frozenRows();
  const fc = frozenCols();

  // addCols makes the headers of columns c1..c2. base is the position of
  // column A in the parent node.
  const addCols = (parent, c1, c2, base) => {
    let x = base + S.cols.offset(c1);
    for (let c = c1; c <= c2; c++) {
      const w = S.cols.size(c);
      if (w > 0) {
        const node = document.createElement('div');
        node.className = c >= sel.c1 && c <= sel.c2 ? 'col-header sel' : 'col-header';
        node.dataset.colHeader = c;
        node.textContent = colName(c);
        node.style.cssText = `left:${x}px;width:${w}px;`;
        // The resize handle at the right edge. resize.js owns its behavior.
        const handle = document.createElement('div');
        handle.className = 'col-resize';
        handle.dataset.colResize = c;
        node.append(handle);
        parent.append(node);
      }
      x += w;
    }
  };
  const colFrag = document.createDocumentFragment();
  if (R.fc) {
    // The headers of the frozen columns are in a sticky node, so they stay in view.
    const frozen = document.createElement('div');
    frozen.className = 'frozen-heads';
    addCols(frozen, 1, R.fc, 0);
    colFrag.append(frozen);
  }
  colFrag.append(el.corner);
  if (Math.max(R.c1, fc + 1) <= R.c2) addCols(colFrag, Math.max(R.c1, fc + 1), R.c2, S.rw);
  el.colHeaders.replaceChildren(colFrag);

  const addRows = (parent, r1, r2) => {
    let y = S.rows.offset(r1);
    for (let r = r1; r <= r2; r++) {
      const h = S.rows.size(r);
      if (h > 0) {
        const node = document.createElement('div');
        node.className = r >= sel.r1 && r <= sel.r2 ? 'row-header sel' : 'row-header';
        node.dataset.rowHeader = r;
        node.textContent = r;
        node.style.cssText = `top:${y}px;height:${h}px;`;
        // The resize handle at the bottom edge.
        const handle = document.createElement('div');
        handle.className = 'row-resize';
        handle.dataset.rowResize = r;
        node.append(handle);
        parent.append(node);
      }
      y += h;
    }
  };
  const rowFrag = document.createDocumentFragment();
  if (R.fr) {
    const frozen = document.createElement('div');
    frozen.className = 'frozen-heads';
    addRows(frozen, 1, R.fr);
    rowFrag.append(frozen);
  }
  if (Math.max(R.r1, fr + 1) <= R.r2) addRows(rowFrag, Math.max(R.r1, fr + 1), R.r2);
  el.rowHeaders.replaceChildren(rowFrag);
}

// ------------------------------------------------------------ frozen panes
//
// Rows 1..S.freeze.rows and columns 1..S.freeze.cols stay in view. Their
// cells are in #pane-top, #pane-left and #pane-corner, which are sticky.
// All other cells are in #cells and scroll below the panes.
//
// #overlay scrolls with #cells. rectOf gives the rectangle where a range
// shows at this time: a frozen part moves with the scroll position, and a
// part that is below a pane is cut at the pane line. When a freeze is on,
// the overlays are thus placed again on each scroll.

export function frozenRows() {
  return Math.min(S.freeze.rows, Math.max(0, S.rowCount - 1));
}

export function frozenCols() {
  return Math.min(S.freeze.cols, Math.max(0, S.colCount - 1));
}

// absRect is the rectangle of a range in the coordinates of the sheet, with
// no frozen panes.
function absRect(r1, c1, r2, c2) {
  const left = S.cols.offset(c1);
  const top = S.rows.offset(r1);
  return { left, top, width: S.cols.offset(c2 + 1) - left, height: S.rows.offset(r2 + 1) - top };
}

// edgeOf maps the edge before row (or column) i to the position where it
// shows, for `frozen` frozen rows (columns) and the scroll position.
function edgeOf(axis, i, frozen, scroll) {
  if (!frozen) return axis.offset(i);
  if (i <= frozen + 1) return scroll + axis.offset(i);
  return Math.max(axis.offset(i), scroll + axis.offset(frozen + 1));
}

// rectOf is the rectangle where a range shows in #overlay. A range that is
// fully below a frozen pane has width 0 or height 0.
export function rectOf(r1, c1, r2, c2) {
  const fr = frozenRows();
  const fc = frozenCols();
  if (!fr && !fc) return absRect(r1, c1, r2, c2);
  const st = fr ? el.grid.scrollTop : 0;
  const sl = fc ? el.grid.scrollLeft : 0;
  const left = edgeOf(S.cols, c1, fc, sl);
  const top = edgeOf(S.rows, r1, fr, st);
  return {
    left,
    top,
    width: edgeOf(S.cols, c2 + 1, fc, sl) - left,
    height: edgeOf(S.rows, r2 + 1, fr, st) - top,
  };
}

export function fetchCell(row, col) {
  const res = rawCall('getRange', [S.sheet, row, col, row, col]);
  return res.ok && Array.isArray(res.cells) && res.cells.length ? res.cells[0] : null;
}

export function cellAt(row, col) {
  const R = S.rendered;
  // Cells of the frozen panes are in S.cells too, also out of the range R.
  const hit = R ? S.cells.get(cellKey(row, col)) : null;
  if (hit) return hit;
  if (R && row >= R.r1 && row <= R.r2 && col >= R.c1 && col <= R.c2) return S.cells.get(cellKey(row, col)) || null;
  return fetchCell(row, col);
}

// afterWrite runs on each 'write' event (a successful mutate): it sets the
// dirty mark, reads the sheet sizes again and paints the visible range again.
export function afterWrite(res) {
  // A call that reports that it changed no cell (for example a clear of
  // empty cells) does not make the workbook dirty.
  if (!(res && res.changed === 0)) setDirty(true);
  refreshSheets();
  // A sheet command (rename, delete) can leave S.sheet with a name that is
  // not in the workbook for a moment. The caller then shows a valid sheet.
  if (!sheetInfo()) return;
  render(true);
  updateSelectionUI();
  updateBars();
}

// usedClamp limits a range to the used range of the sheet. It returns null
// if the range has no cell in the used range.
export function usedClamp(R) {
  const info = sheetInfo();
  const maxRow = info ? info.maxRow : 0;
  const maxCol = info ? info.maxCol : 0;
  if (R.r1 > maxRow || R.c1 > maxCol) return null;
  return { r1: R.r1, c1: R.c1, r2: Math.min(R.r2, maxRow), c2: Math.min(R.c2, maxCol) };
}

// ---------------------------------------------------------------- selection

export function mergeAt(row, col) {
  for (const m of S.merges) {
    if (row >= m.row1 && row <= m.row2 && col >= m.col1 && col <= m.col2) return m;
  }
  return null;
}

// selRange returns the selected range. It grows to include each merged
// range that it touches.
export function selRange() {
  const R = {
    r1: Math.min(S.anchor.row, S.focus.row),
    c1: Math.min(S.anchor.col, S.focus.col),
    r2: Math.max(S.anchor.row, S.focus.row),
    c2: Math.max(S.anchor.col, S.focus.col),
  };
  let changed = S.merges.length > 0;
  while (changed) {
    changed = false;
    for (const m of S.merges) {
      if (m.row2 < R.r1 || m.row1 > R.r2 || m.col2 < R.c1 || m.col1 > R.c2) continue;
      if (m.row1 < R.r1) { R.r1 = m.row1; changed = true; }
      if (m.col1 < R.c1) { R.c1 = m.col1; changed = true; }
      if (m.row2 > R.r2) { R.r2 = m.row2; changed = true; }
      if (m.col2 > R.c2) { R.c2 = m.col2; changed = true; }
    }
  }
  return R;
}

export function isSingleSelection() {
  return S.anchor.row === S.focus.row && S.anchor.col === S.focus.col;
}

// isFullRows: the selection has all columns of its rows (a click on a row
// header, or select all).
export function isFullRows() {
  const R = selRange();
  return R.c1 === 1 && R.c2 >= MAX_COLS;
}

// isFullCols: the selection has all rows of its columns.
export function isFullCols() {
  const R = selRange();
  return R.r1 === 1 && R.r2 >= MAX_ROWS;
}

function activeRect() {
  const m = mergeAt(S.anchor.row, S.anchor.col);
  return m ? rectOf(m.row1, m.col1, m.row2, m.col2) : rectOf(S.anchor.row, S.anchor.col, S.anchor.row, S.anchor.col);
}

export function setSelection(anchor, focus, opts) {
  const o = opts || {};
  const m = mergeAt(anchor.row, anchor.col);
  const single = anchor.row === focus.row && anchor.col === focus.col;
  S.anchor = m ? { row: m.row1, col: m.col1 } : { row: anchor.row, col: anchor.col };
  S.focus = single ? { ...S.anchor } : { row: focus.row, col: focus.col };
  if (!o.keepTab) S.tabOrigin = null;
  if (o.scroll !== false) {
    const target = single ? S.anchor : S.focus;
    ensureExtent(target.row, target.col);
    scrollToCell(target.row, target.col);
  }
  render(false);
  updateSelectionUI();
  updateBars();
}

// placeSelection puts #active-cell and #selection-range at the place where
// the selection shows. A part that is below a frozen pane does not show.
function placeSelection() {
  const m = mergeAt(S.anchor.row, S.anchor.col);
  const R = selRange();
  const single = m
    ? R.r1 === m.row1 && R.c1 === m.col1 && R.r2 === m.row2 && R.c2 === m.col2
    : R.r1 === R.r2 && R.c1 === R.c2;
  const a = activeRect();
  el.active.hidden = a.width <= 0 || a.height <= 0;
  el.active.style.cssText = `left:${a.left - 1}px;top:${a.top - 1}px;width:${a.width + 1}px;height:${a.height + 1}px;`;
  el.active.dataset.cell = addrOf(S.anchor.row, S.anchor.col);
  const r = rectOf(R.r1, R.c1, R.r2, R.c2);
  el.selRange.hidden = single || r.width <= 0 || r.height <= 0;
  el.selRange.style.cssText = `left:${r.left - 1}px;top:${r.top - 1}px;width:${r.width + 1}px;height:${r.height + 1}px;`;
  // The address is the true selection: a full column is A1:A1048576, also
  // when the grid shows fewer rows.
  el.grid.dataset.selection = single
    ? addrOf(S.anchor.row, S.anchor.col)
    : `${addrOf(R.r1, R.c1)}:${addrOf(R.r2, R.c2)}`;
}

export function updateSelectionUI() {
  placeSelection();
  renderHeaders();
  emit('selection');
}

// scrollToCell scrolls the grid so that the cell is in view. only is
// optional: 'row' scrolls only on the row axis, 'col' only on the column
// axis. Frozen rows and columns are always in view; the other cells must
// show in the area that the frozen panes do not cover.
export function scrollToCell(row, col, only) {
  const m = mergeAt(row, col);
  const r1 = m ? m.row1 : row;
  const c1 = m ? m.col1 : col;
  const r2 = m ? m.row2 : row;
  const c2 = m ? m.col2 : col;
  const rect = absRect(r1, c1, r2, c2);
  const fr = frozenRows();
  const fc = frozenCols();
  const fh = fr ? S.rows.offset(fr + 1) : 0;
  const fw = fc ? S.cols.offset(fc + 1) : 0;
  const vh = viewHeight();
  const vw = viewWidth();
  let top = el.grid.scrollTop;
  let left = el.grid.scrollLeft;
  if (only !== 'col' && r2 > fr) {
    if (rect.top < top + fh || rect.height > vh - fh) top = rect.top - fh;
    else if (rect.top + rect.height > top + vh) top = rect.top + rect.height - vh;
  }
  if (only !== 'row' && c2 > fc) {
    if (rect.left < left + fw || rect.width > vw - fw) left = rect.left - fw;
    else if (rect.left + rect.width > left + vw) left = rect.left + rect.width - vw;
  }
  top = Math.max(0, top);
  left = Math.max(0, left);
  if (top !== el.grid.scrollTop) el.grid.scrollTop = top;
  if (left !== el.grid.scrollLeft) el.grid.scrollLeft = left;
}

export function updateBars() {
  if (S.edit || !S.sheet) return;
  if (document.activeElement !== el.nameBox) el.nameBox.value = addrOf(S.anchor.row, S.anchor.col);
  const cell = cellAt(S.anchor.row, S.anchor.col);
  el.formula.value = cell && cell.input != null ? String(cell.input) : '';
  el.formula.readOnly = isReadOnly(cell);
}

// moveActive moves the active cell and collapses the selection.
export function moveActive(dRow, dCol, opts) {
  const m = mergeAt(S.anchor.row, S.anchor.col);
  let row = S.anchor.row;
  let col = S.anchor.col;
  if (dRow > 0) row = (m ? m.row2 : row) + dRow;
  else if (dRow < 0) row = (m ? m.row1 : row) + dRow;
  if (dCol > 0) col = (m ? m.col2 : col) + dCol;
  else if (dCol < 0) col = (m ? m.col1 : col) + dCol;
  const next = { row: clamp(row, 1, S.rowLimit), col: clamp(col, 1, S.colLimit) };
  setSelection(next, next, opts);
}

export function extendSelection(dRow, dCol) {
  // Only the axis that moves changes. A full-column selection thus keeps
  // its true last row (and a full-row selection its true last column).
  const focus = {
    row: dRow ? clamp(S.focus.row + dRow, 1, S.rowLimit) : S.focus.row,
    col: dCol ? clamp(S.focus.col + dCol, 1, S.colLimit) : S.focus.col,
  };
  setSelection(S.anchor, focus, { scroll: false });
  ensureExtent(dRow ? focus.row : 1, dCol ? focus.col : 1);
  scrollToCell(focus.row, focus.col, dRow ? 'row' : 'col');
  render(false);
}

export function enterMove(back) {
  if (!back && S.tabOrigin !== null) {
    // After a sequence of Tab keys, Enter goes to the start column of the next row.
    const m = mergeAt(S.anchor.row, S.anchor.col);
    const next = { row: clamp((m ? m.row2 : S.anchor.row) + 1, 1, S.rowLimit), col: S.tabOrigin };
    setSelection(next, next);
  } else {
    moveActive(back ? -1 : 1, 0);
  }
}

export function tabMove(back) {
  if (S.tabOrigin === null) S.tabOrigin = S.anchor.col;
  moveActive(0, back ? -1 : 1, { keepTab: true });
}

export function gridHasFocus() {
  return S.state === 'editor' && !S.edit && !S.confirmResolve && document.activeElement === el.grid;
}

// ------------------------------------------------------- mouse and keyboard

export function pointToCell(clientX, clientY) {
  const rect = el.cells.getBoundingClientRect();
  let x = clientX - rect.left;
  let y = clientY - rect.top;
  // A point on a frozen pane: the pane does not scroll, so the position in
  // the view gives the row (column).
  const fr = frozenRows();
  const fc = frozenCols();
  if (fr && y - el.grid.scrollTop < S.rows.offset(fr + 1)) y -= el.grid.scrollTop;
  if (fc && x - el.grid.scrollLeft < S.cols.offset(fc + 1)) x -= el.grid.scrollLeft;
  return {
    row: clamp(S.rows.indexAt(y), 1, S.rowCount),
    col: clamp(S.cols.indexAt(x), 1, S.colCount),
  };
}

function dragUpdate() {
  if (!S.drag) return;
  if (S.drag.point) {
    pointDragUpdate();
    return;
  }
  const p = pointToCell(S.drag.x, S.drag.y);
  if (p.row !== S.focus.row || p.col !== S.focus.col) setSelection(S.anchor, p, { scroll: false });
}

// dragTick is the animation frame loop of a drag (selection, or point mode
// in edit.js). It scrolls the grid while the pointer is out of the grid.
export function dragTick() {
  if (!S.drag) return;
  const g = el.grid.getBoundingClientRect();
  let dx = 0;
  let dy = 0;
  if (S.drag.y > g.bottom) dy = 30;
  else if (S.drag.y < g.top + HEADER_H) dy = -30;
  if (S.drag.x > g.right) dx = 60;
  else if (S.drag.x < g.left + S.rw) dx = -60;
  if (dx || dy) {
    el.grid.scrollTop += dy;
    el.grid.scrollLeft += dx;
    growIfNeeded();
    render(false);
    dragUpdate();
  }
  S.drag.raf = requestAnimationFrame(dragTick);
}

function onGridMouseDown(e) {
  if (e.button !== 0) return;
  for (const handler of mouseDownHandlers) {
    if (handler(e) === true) return;
  }
  if (e.target === el.editor || e.target === el.grid) return;
  if (S.edit) {
    if (pointActive()) {
      // Formula point mode: the click gives a reference to the formula.
      // preventDefault keeps the focus and the caret in the editor.
      e.preventDefault();
      pointMouseDown(e);
      return;
    }
    const otherSheet = S.sheet !== S.edit.sheet;
    if (!commitEdit(null, false) || otherSheet) {
      // After a commit from a different sheet, the sheet of the edited cell
      // is on screen again, so the click position does not apply to it.
      e.preventDefault();
      return;
    }
  }
  e.preventDefault();
  el.grid.focus({ preventScroll: true });
  if (e.target.closest('#corner')) {
    setSelection({ row: 1, col: 1 }, { row: MAX_ROWS, col: MAX_COLS }, { scroll: false });
    return;
  }
  const colHeader = e.target.closest('[data-col-header]');
  if (colHeader) {
    const col = Number(colHeader.dataset.colHeader);
    setSelection({ row: 1, col }, { row: MAX_ROWS, col }, { scroll: false });
    return;
  }
  const rowHeader = e.target.closest('[data-row-header]');
  if (rowHeader) {
    const row = Number(rowHeader.dataset.rowHeader);
    setSelection({ row, col: 1 }, { row, col: MAX_COLS }, { scroll: false });
    return;
  }
  if (e.target.closest('#col-headers, #row-headers')) return;
  const p = pointToCell(e.clientX, e.clientY);
  if (e.shiftKey) setSelection(S.anchor, p);
  else setSelection(p, p);
  S.drag = { x: e.clientX, y: e.clientY, raf: requestAnimationFrame(dragTick) };
}

function onHeaderMenu(e) {
  if (!headerMenuHandlers.length || S.edit || !S.sheet) return;
  const colHeader = e.target.closest('[data-col-header]');
  const rowHeader = colHeader ? null : e.target.closest('[data-row-header]');
  if (!colHeader && !rowHeader) return;
  const data = colHeader
    ? { kind: 'col', index: Number(colHeader.dataset.colHeader), event: e }
    : { kind: 'row', index: Number(rowHeader.dataset.rowHeader), event: e };
  for (const handler of headerMenuHandlers) handler(data);
}

function onGridKey(e) {
  if (e.target !== el.grid || S.edit || S.confirmResolve || !S.sheet) return;
  // A registered key handler has priority over the default keys.
  if (runKey('grid', e)) return;
  const mod = e.ctrlKey || e.metaKey;
  const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
  if (arrows[e.key]) {
    if (mod || e.altKey) return;
    e.preventDefault();
    const d = arrows[e.key];
    if (e.shiftKey) extendSelection(d[0], d[1]);
    else moveActive(d[0], d[1]);
    return;
  }
  switch (e.key) {
    case 'Enter':
      e.preventDefault();
      enterMove(e.shiftKey);
      return;
    case 'Tab':
      e.preventDefault();
      tabMove(e.shiftKey);
      return;
    case 'F2':
      e.preventDefault();
      startEdit('edit');
      return;
    case 'Delete':
    case 'Backspace':
      e.preventDefault();
      clearSelection();
      return;
    case 'Escape':
      if (!isSingleSelection()) setSelection(S.anchor, S.anchor);
      return;
    case 'Home': {
      e.preventDefault();
      const next = mod ? { row: 1, col: 1 } : { row: S.anchor.row, col: 1 };
      setSelection(next, next);
      return;
    }
    case 'End': {
      if (!mod) return;
      e.preventDefault();
      const info = sheetInfo();
      const next = {
        row: clamp(info ? info.maxRow : 1, 1, S.rowLimit),
        col: clamp(info ? info.maxCol : 1, 1, S.colLimit),
      };
      setSelection(next, next);
      return;
    }
    case 'PageDown':
    case 'PageUp': {
      e.preventDefault();
      const page = Math.max(1, Math.floor(viewHeight() / DEF_ROW_H) - 1);
      moveActive(e.key === 'PageDown' ? page : -page, 0);
      return;
    }
    default:
  }
  if (mod && !e.altKey && (e.key === 'a' || e.key === 'A')) {
    e.preventDefault();
    const info = sheetInfo();
    setSelection(
      { row: 1, col: 1 },
      { row: Math.max(1, info ? info.maxRow : 1), col: Math.max(1, info ? info.maxCol : 1) },
      { scroll: false },
    );
    return;
  }
  if (!mod && !e.altKey && !e.isComposing && e.key.length === 1) {
    // The focus moves to the editor in this keydown, so the browser puts
    // the typed character into the editor.
    if (!startEdit('type', '')) e.preventDefault();
  }
}

function jumpTo(text) {
  const ref = parseRef(text);
  if (!ref) {
    showStatus(`"${text}" is not a cell address.`, 'error');
    return false;
  }
  if (ref.sheet !== undefined) {
    const target = S.sheets.find((s) => s.name.toLowerCase() === ref.sheet.toLowerCase());
    if (!target) {
      showStatus(`There is no sheet with the name "${ref.sheet}".`, 'error');
      return false;
    }
    if (target.name !== S.sheet) showSheet(target.name);
  }
  if (Math.max(ref.row1, ref.row2) > S.rowLimit || Math.max(ref.col1, ref.col2) > S.colLimit) {
    showStatus(`The grid can show rows 1 to ${S.rowLimit} and columns A to ${colName(S.colLimit)}.`, 'error');
    return false;
  }
  const anchor = { row: ref.row1, col: ref.col1 };
  const focus = { row: ref.row2, col: ref.col2 };
  // Show the active cell; a large range does not move the view to its far corner.
  setSelection(anchor, anchor);
  if (focus.row !== anchor.row || focus.col !== anchor.col) setSelection(anchor, focus, { scroll: false });
  return true;
}

export function init() {
  on('write', afterWrite);

  el.grid.addEventListener('mousedown', onGridMouseDown);
  el.grid.addEventListener('contextmenu', onHeaderMenu);
  document.addEventListener('mousemove', (e) => {
    if (!S.drag) return;
    S.drag.x = e.clientX;
    S.drag.y = e.clientY;
    dragUpdate();
  });
  document.addEventListener('mouseup', () => {
    if (!S.drag) return;
    cancelAnimationFrame(S.drag.raf);
    S.drag = null;
  });
  el.grid.addEventListener('dblclick', (e) => {
    if (S.edit || e.target === el.grid || e.target.closest('#col-headers, #row-headers')) return;
    startEdit('edit');
  });
  el.grid.addEventListener('keydown', onGridKey);
  el.grid.addEventListener('scroll', () => {
    if (S.scrollRaf) return;
    S.scrollRaf = requestAnimationFrame(() => {
      S.scrollRaf = 0;
      if (S.state !== 'editor' || !S.sheet) return;
      growIfNeeded();
      if (!render(false) && (frozenRows() || frozenCols())) {
        // The cell nodes did not change, but the overlays on and near the
        // frozen panes depend on the scroll position. The 'render' event
        // makes the other modules place their overlays again.
        placeSelection();
        emit('render', S.rendered);
      }
    });
  }, { passive: true });
  if (typeof ResizeObserver === 'function') {
    new ResizeObserver(() => render(false)).observe(el.grid);
  } else {
    window.addEventListener('resize', () => render(false));
  }
  const onScheme = () => render(true);
  if (darkQuery.addEventListener) darkQuery.addEventListener('change', onScheme);

  // Name box.
  el.nameBox.addEventListener('focus', () => el.nameBox.select());
  el.nameBox.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      if (jumpTo(el.nameBox.value)) el.grid.focus({ preventScroll: true });
      else el.nameBox.select();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      el.grid.focus({ preventScroll: true });
    }
  });
  el.nameBox.addEventListener('blur', () => {
    el.nameBox.value = addrOf(S.anchor.row, S.anchor.col);
  });
}
