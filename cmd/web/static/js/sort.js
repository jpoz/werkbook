// Sort (PLAN.md 4.9). Owner: fe-b.
//
// The buttons #sort-asc and #sort-desc in #tb-data. The sort order and the
// move of the rows are done by the engine (currentRegion, sortRange). While
// the engine does not have these functions, the buttons are disabled.

import { S, el, api, mutate, on, showStatus } from './core.js';
import { selRange, isSingleSelection, usedClamp } from './grid.js';
import { toolButton, engineHas } from './ui.js';

const buttons = [];

const canSort = () =>
  S.state === 'editor' && Boolean(S.sheet) && !S.edit && engineHas('sortRange') && engineHas('currentRegion');

// sortTarget returns { range, header } for the sort, or null (a message was
// shown). A selection of more than one cell is the range, limited to the
// used range, with no header row. For one cell, the range is the data region
// around the active cell and the engine reports if it has a header row.
function sortTarget() {
  if (!isSingleSelection()) {
    const R = usedClamp(selRange());
    if (!R || S.anchor.col > R.c2) {
      showStatus('The selection has no data to sort.', 'info');
      return null;
    }
    return { range: { row1: R.r1, col1: R.c1, row2: R.r2, col2: R.c2 }, header: false };
  }
  const res = api('currentRegion', S.sheet, S.anchor.row, S.anchor.col);
  if (!res.ok || !res.range) return null;
  return { range: res.range, header: Boolean(res.header) };
}

// sort sorts the rows of the range by the column of the active cell. One
// call is one undo step.
function sort(desc) {
  if (!canSort()) return;
  const target = sortTarget();
  if (!target) return;
  const key = desc ? { col: S.anchor.col, desc: true } : { col: S.anchor.col };
  mutate('sortRange', S.sheet, target.range, [key], target.header);
}

function refresh() {
  const enabled = canSort();
  for (const btn of buttons) btn.disabled = !enabled;
}

export function init() {
  const mount = el.tbData;
  if (!mount) return;
  buttons.push(
    toolButton(mount, {
      id: 'sort-asc',
      label: 'A→Z',
      title: 'Sort ascending by the column of the active cell',
      ariaLabel: 'Sort ascending',
      onClick: () => sort(false),
    }),
    toolButton(mount, {
      id: 'sort-desc',
      label: 'Z→A',
      title: 'Sort descending by the column of the active cell',
      ariaLabel: 'Sort descending',
      onClick: () => sort(true),
    }),
  );
  refresh();
  on('workbook', refresh);
  on('sheet', refresh);
  on('edit:start', refresh);
  on('edit:end', refresh);
}
