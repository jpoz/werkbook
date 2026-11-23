// Ctrl+arrow and Ctrl+Shift+arrow (PLAN.md 4.7). Owner: fe-a.
//
// The rule (end of the data block, next data cell, edge of the sheet) is in
// the bridge: findEdge. This module sends the start cell and applies the
// result. Cmd+arrow on macOS is the same.

import { S, clamp, api, registerKey } from './core.js';
import { setSelection, ensureExtent, scrollToCell, render } from './grid.js';

const ARROWS = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };

// edge returns the cell that findEdge gives from a start cell, limited to
// the grid limits, or null when the call failed.
function edge(from, d) {
  const res = api('findEdge', S.sheet, from.row, from.col, d[0], d[1]);
  if (!res.ok) return null;
  return { row: clamp(res.row, 1, S.rowLimit), col: clamp(res.col, 1, S.colLimit) };
}

// available: a function that is not built yet is absent from the module.
// The keys then do nothing, with no error message for each key press.
const available = () => Boolean(globalThis.werkbook) && typeof globalThis.werkbook.findEdge === 'function';

function jump(d) {
  if (!available()) return false;
  const to = edge(S.anchor, d);
  if (to) setSelection(to, to);
  return true;
}

// extend moves the focus corner of the selection. The start cell is the
// focus cell.
function extend(d) {
  if (!available()) return false;
  const to = edge(S.focus, d);
  if (!to) return true;
  // Only the axis of the arrow changes. A full-column selection thus keeps
  // its true last row (and a full-row selection its true last column).
  if (!d[0]) to.row = S.focus.row;
  if (!d[1]) to.col = S.focus.col;
  setSelection(S.anchor, to, { scroll: false });
  ensureExtent(d[0] ? to.row : 1, d[1] ? to.col : 1);
  scrollToCell(to.row, to.col, d[0] ? 'row' : 'col');
  render(false);
  return true;
}

export function init() {
  for (const key of Object.keys(ARROWS)) {
    const d = ARROWS[key];
    registerKey('grid', { key, mod: true }, () => jump(d));
    registerKey('grid', { key, mod: true, shift: true }, () => extend(d));
  }
}
