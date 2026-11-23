// Freeze panes (PLAN.md 4.9). Owner: fe-a.
//
// The button #freeze-toggle. The freeze is in the workbook (setFreeze,
// sheetMeta.freeze) and is saved in the XLSX. grid.js shows the frozen rows
// and columns in the panes. While the engine has no setFreeze, the button is
// disabled.

import { el, S, showStatus, mutate, on } from './core.js';
import { reloadMeta } from './grid.js';
import { toolButton, engineHas } from './ui.js';

let button = null;

const isFrozen = () => S.freeze.rows > 0 || S.freeze.cols > 0;

function refresh() {
  if (!button) return;
  const frozen = isFrozen();
  button.dataset.state = frozen ? 'frozen' : 'none';
  button.textContent = frozen ? 'Unfreeze' : 'Freeze';
  const title = frozen
    ? 'Remove the freeze of rows and columns'
    : 'Freeze the rows above and the columns at the left of the active cell';
  button.title = title;
  button.setAttribute('aria-label', title);
  button.setAttribute('aria-pressed', frozen ? 'true' : 'false');
  button.disabled = S.state !== 'editor' || !S.sheet || Boolean(S.edit) || !engineHas('setFreeze');
}

// toggle freezes the rows above and the columns at the left of the active
// cell, or removes the freeze. One call is one undo step.
function toggle() {
  if (S.state !== 'editor' || !S.sheet || S.edit) return;
  let rows = 0;
  let cols = 0;
  if (!isFrozen()) {
    rows = S.anchor.row - 1;
    cols = S.anchor.col - 1;
    if (!rows && !cols) {
      showStatus('The active cell is A1, so there is nothing to freeze. Select the cell below the rows and at the right of the columns that must stay in view.', 'info');
      return;
    }
  }
  const res = mutate('setFreeze', S.sheet, rows, cols);
  if (!res.ok) return;
  if (rows || cols) {
    // The frozen rows and columns start at row 1 and column A. Show the
    // sheet from its start, so that the panes and the other cells agree.
    el.grid.scrollTop = 0;
    el.grid.scrollLeft = 0;
  }
  reloadMeta(); // emits 'sheet', which calls refresh
  el.grid.focus({ preventScroll: true });
}

export function init() {
  if (!el.tbData) return;
  button = toolButton(el.tbData, {
    id: 'freeze-toggle',
    label: 'Freeze',
    title: 'Freeze',
    onClick: toggle,
  });
  refresh();
  on('workbook', refresh);
  on('sheet', refresh);
  on('edit:start', refresh);
  on('edit:end', refresh);
}
