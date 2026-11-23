// Insert and delete rows and columns (PLAN.md 4.6). Owner: fe-a.
//
// The bridge moves the cells and adjusts formulas, names, merged ranges and
// tables (insertRows, deleteRows, insertCols, deleteCols). This module has
// the header menu, the keys and the message about items that do not move.

import { MAX_ROWS, MAX_COLS, S, showStatus, mutate, on, registerKey } from './core.js';
import {
  reloadMeta, selRange, isFullRows, isFullCols, setSelection, registerHeaderMenu,
} from './grid.js';
import { openMenu, engineHas } from './ui.js';

// The sheets that already got the message about items that do not move.
const warned = new Set();

const plural = (n, word) => (n === 1 ? word : `${n} ${word}s`);

// run calls one structural function. After it, the layout of the sheet is
// read again. The selection stays at the same row and column numbers.
function run(name, at, count) {
  if (S.edit || !S.sheet) return;
  const sheet = S.sheet;
  const kinds = S.warnings.slice();
  const res = mutate(name, sheet, at, count);
  if (!res.ok) return;
  if (Array.isArray(res.sheets) && res.sheets.length) S.sheets = res.sheets;
  reloadMeta();
  // Passthrough items (conditional formats, data validation, hyperlinks,
  // drawings) stay at their old cell addresses in this version. Tell the
  // user one time for each sheet. A message about the undo history, from the
  // same call, has priority.
  const cleared = Boolean(res.history && res.history.cleared);
  if (kinds.length && !warned.has(sheet) && !cleared) {
    warned.add(sheet);
    showStatus(
      `This sheet has items that do not move with rows and columns in this version: ${kinds.join(', ')}. Check them after the change.`,
      'info',
    );
  }
}

// rowSpan and colSpan return the selected full rows (columns), or null.
// Select all is not a row selection and not a column selection.
function rowSpan() {
  if (!isFullRows() || isFullCols()) return null;
  const R = selRange();
  return { at: R.r1, count: R.r2 - R.r1 + 1 };
}

function colSpan() {
  if (!isFullCols() || isFullRows()) return null;
  const R = selRange();
  return { at: R.c1, count: R.c2 - R.c1 + 1 };
}

const COMMANDS = {
  'insert-rows-above': (s) => run('insertRows', s.at, s.count),
  'insert-rows-below': (s) => run('insertRows', s.at + s.count, s.count),
  'delete-rows': (s) => run('deleteRows', s.at, s.count),
  'insert-cols-left': (s) => run('insertCols', s.at, s.count),
  'insert-cols-right': (s) => run('insertCols', s.at + s.count, s.count),
  'delete-cols': (s) => run('deleteCols', s.at, s.count),
};

function menuItems(kind, span) {
  const n = span.count;
  const item = (cmd, label, fn) => ({
    cmd,
    label,
    // A function that is not built yet is absent from the module.
    disabled: !engineHas(fn),
    onSelect: () => COMMANDS[cmd](span),
  });
  if (kind === 'row') {
    return [
      item('insert-rows-above', `Insert ${plural(n, 'row')} above`, 'insertRows'),
      item('insert-rows-below', `Insert ${plural(n, 'row')} below`, 'insertRows'),
      { separator: true },
      item('delete-rows', n === 1 ? 'Delete row' : `Delete ${n} rows`, 'deleteRows'),
    ];
  }
  return [
    item('insert-cols-left', `Insert ${plural(n, 'column')} left`, 'insertCols'),
    item('insert-cols-right', `Insert ${plural(n, 'column')} right`, 'insertCols'),
    { separator: true },
    item('delete-cols', n === 1 ? 'Delete column' : `Delete ${n} columns`, 'deleteCols'),
  ];
}

// onHeaderMenu is the right click on a row header or a column header.
function onHeaderMenu({ kind, index, event }) {
  event.preventDefault();
  let span = kind === 'row' ? rowSpan() : colSpan();
  if (!span || index < span.at || index >= span.at + span.count) {
    // The header is not in the selection: select that row or column first.
    if (kind === 'row') setSelection({ row: index, col: 1 }, { row: index, col: MAX_COLS }, { scroll: false });
    else setSelection({ row: 1, col: index }, { row: MAX_ROWS, col: index }, { scroll: false });
    span = { at: index, count: 1 };
  }
  openMenu(menuItems(kind, span), event.clientX, event.clientY);
}

// The keys apply only when the selection is full rows or full columns. In
// all other cases they return false and the browser keeps its behavior.
function keyInsert() {
  const rows = rowSpan();
  const cols = rows ? null : colSpan();
  if (rows && engineHas('insertRows')) run('insertRows', rows.at, rows.count);
  else if (cols && engineHas('insertCols')) run('insertCols', cols.at, cols.count);
  else return false;
  return true;
}

function keyDelete() {
  const rows = rowSpan();
  const cols = rows ? null : colSpan();
  if (rows && engineHas('deleteRows')) run('deleteRows', rows.at, rows.count);
  else if (cols && engineHas('deleteCols')) run('deleteCols', cols.at, cols.count);
  else return false;
  return true;
}

export function init() {
  registerHeaderMenu(onHeaderMenu);
  on('workbook', () => warned.clear());

  // Ctrl+Shift+= gives the key "+" on most keyboard layouts and "=" on some.
  // Ctrl++ on the number pad has no Shift.
  registerKey('grid', { key: '+', mod: true, shift: true }, keyInsert);
  registerKey('grid', { key: '=', mod: true, shift: true }, keyInsert);
  registerKey('grid', { key: '+', mod: true }, keyInsert);
  registerKey('grid', { key: '-', mod: true }, keyDelete);
}
