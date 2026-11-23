// Sheets: show a sheet, read the sheet sizes again, and the sheet tabs
// (switch, add, rename, delete).

import { MIN_ROWS, MIN_COLS, el, S, clamp, showStatus, rawCall, mutate, emit, confirmDialog } from './core.js';
import { readMeta, layout, render, updateSelectionUI, updateBars, mergeAt, sheetInfo } from './grid.js';
import { endEdit, formulaEditing, pointShowSheet } from './edit.js';

export function saveView() {
  if (!S.sheet) return;
  S.views.set(S.sheet, {
    anchor: { ...S.anchor },
    focus: { ...S.focus },
    top: el.grid.scrollTop,
    left: el.grid.scrollLeft,
    rowCount: S.rowCount,
    colCount: S.colCount,
  });
}

// showSheet puts a sheet on screen. keepEdit is for formula point mode: the
// open edit stays open while the user looks at a different sheet.
export function showSheet(name, keepEdit) {
  if (!keepEdit) endEdit();
  if (S.sheet && S.sheet !== name) saveView();
  S.sheet = name;
  readMeta();

  const info = sheetInfo();
  const view = S.views.get(name);
  S.rowCount = clamp(Math.max(MIN_ROWS, (info ? info.maxRow : 0) + 50, view ? view.rowCount : 0), 1, S.rowLimit);
  S.colCount = clamp(Math.max(MIN_COLS, (info ? info.maxCol : 0) + 5, view ? view.colCount : 0), 1, S.colLimit);
  S.anchor = view ? { ...view.anchor } : { row: 1, col: 1 };
  S.focus = view ? { ...view.focus } : { row: 1, col: 1 };
  S.tabOrigin = null;
  const first = mergeAt(S.anchor.row, S.anchor.col);
  if (first) S.anchor = { row: first.row1, col: first.col1 };

  layout();
  el.grid.scrollTop = view ? view.top : 0;
  el.grid.scrollLeft = view ? view.left : 0;
  render(true);
  renderTabs();
  updateSelectionUI();
  updateBars();
  // edit.js puts the in-cell editor back (or hides it) on this event.
  emit('sheet', S.sheet);
  if (info && info.maxRow > S.rowLimit) {
    showStatus(`This sheet has ${info.maxRow} rows. The grid can show rows 1 to ${S.rowLimit}.`, 'info');
  }
}

// refreshSheets reads the sheet sizes again after a write.
export function refreshSheets() {
  const res = rawCall('sheets', []);
  if (!res.ok || !Array.isArray(res.sheets)) return;
  S.sheets = res.sheets;
  const info = sheetInfo();
  if (!info) return;
  const rows = Math.min(S.rowLimit, info.maxRow + 50);
  const cols = Math.min(S.colLimit, info.maxCol + 5);
  if (rows > S.rowCount || cols > S.colCount) {
    S.rowCount = Math.max(S.rowCount, rows);
    S.colCount = Math.max(S.colCount, cols);
    layout();
  }
}

export function renderTabs() {
  const frag = document.createDocumentFragment();
  for (const s of S.sheets) {
    const tab = document.createElement('div');
    const activeTab = s.name === S.sheet;
    tab.className = 'sheet-tab' + (s.visible === false ? ' sheet-hidden' : '');
    tab.dataset.sheetTab = s.name;
    tab.setAttribute('role', 'tab');
    tab.setAttribute('aria-selected', activeTab ? 'true' : 'false');
    tab.tabIndex = 0;
    tab.title = s.visible === false ? `${s.name} (hidden sheet). Double-click to rename.` : 'Double-click to rename.';
    const label = document.createElement('span');
    label.className = 'sheet-tab-label';
    label.textContent = s.name;
    tab.append(label);
    if (activeTab && S.sheets.length > 1) {
      const del = document.createElement('button');
      del.type = 'button';
      del.className = 'sheet-tab-delete';
      del.dataset.sheetDelete = s.name;
      del.title = 'Delete sheet';
      del.setAttribute('aria-label', `Delete sheet ${s.name}`);
      del.textContent = '×';
      tab.append(del);
    }
    frag.append(tab);
  }
  el.tabs.replaceChildren(frag);
  const current = el.tabs.querySelector('[aria-selected="true"]');
  if (current) current.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

function addSheet() {
  const names = new Set(S.sheets.map((s) => s.name.toLowerCase()));
  let n = S.sheets.length + 1;
  while (names.has(`sheet${n}`)) n++;
  const name = `Sheet${n}`;
  const res = mutate('addSheet', name);
  if (!res.ok) return;
  S.sheets = res.sheets;
  showSheet(name);
  el.grid.focus({ preventScroll: true });
}

async function deleteSheet(name) {
  const ok = await confirmDialog(`Delete the sheet "${name}"? You cannot undo this, and the undo history becomes empty.`, 'Delete sheet');
  if (!ok) return;
  const index = S.sheets.findIndex((s) => s.name === name);
  const res = mutate('deleteSheet', name);
  if (!res.ok) return;
  S.sheets = res.sheets;
  S.views.delete(name);
  if (S.sheet === name) {
    S.sheet = null;
    const next = S.sheets[Math.min(Math.max(index, 0), S.sheets.length - 1)];
    showSheet(next.name);
  } else {
    renderTabs();
    render(true);
    updateBars();
  }
  el.grid.focus({ preventScroll: true });
}

function startRename(tab) {
  const oldName = tab.dataset.sheetTab;
  const label = tab.querySelector('.sheet-tab-label');
  if (!label || tab.querySelector('.sheet-tab-rename')) return;
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'sheet-tab-rename';
  input.dataset.sheetRenameInput = oldName;
  input.value = oldName;
  input.spellcheck = false;
  input.setAttribute('aria-label', `New name for sheet ${oldName}`);
  label.replaceWith(input);
  let done = false;
  const finish = (commit) => {
    if (done) return;
    done = true;
    const newName = input.value.trim();
    if (commit && newName && newName !== oldName) {
      const res = mutate('renameSheet', oldName, newName);
      if (res.ok) {
        S.sheets = res.sheets;
        const view = S.views.get(oldName);
        if (view) {
          S.views.delete(oldName);
          S.views.set(newName, view);
        }
        if (S.sheet === oldName) S.sheet = newName;
      }
    }
    renderTabs();
    // Formulas that refer to the sheet can show a new text.
    render(true);
    updateBars();
  };
  input.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Enter') {
      e.preventDefault();
      finish(true);
      el.grid.focus({ preventScroll: true });
    } else if (e.key === 'Escape') {
      e.preventDefault();
      finish(false);
      el.grid.focus({ preventScroll: true });
    }
  });
  input.addEventListener('blur', () => finish(true));
  input.addEventListener('click', (e) => e.stopPropagation());
  input.addEventListener('dblclick', (e) => e.stopPropagation());
  input.focus();
  input.select();
}

export function init() {
  el.tabs.addEventListener('mousedown', (e) => {
    // During a formula edit the focus stays in the editor, so the edit does
    // not commit and the user can point at cells on a different sheet.
    if (formulaEditing() && e.target.closest('[data-sheet-tab]')) e.preventDefault();
  });
  el.tabs.addEventListener('click', (e) => {
    const del = e.target.closest('[data-sheet-delete]');
    const tab = e.target.closest('[data-sheet-tab]');
    if (formulaEditing()) {
      if (tab && !del && tab.dataset.sheetTab !== S.sheet) pointShowSheet(tab.dataset.sheetTab);
      return;
    }
    if (del) {
      deleteSheet(del.dataset.sheetDelete);
      return;
    }
    if (tab && tab.dataset.sheetTab !== S.sheet) {
      showSheet(tab.dataset.sheetTab);
      el.grid.focus({ preventScroll: true });
    }
  });
  el.tabs.addEventListener('dblclick', (e) => {
    if (S.edit) return;
    if (e.target.closest('[data-sheet-delete]')) return;
    const tab = e.target.closest('[data-sheet-tab]');
    if (tab) startRename(tab);
  });
  el.tabs.addEventListener('keydown', (e) => {
    const tab = e.target.closest('[data-sheet-tab]');
    if (!tab || e.target !== tab) return;
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      if (tab.dataset.sheetTab !== S.sheet) showSheet(tab.dataset.sheetTab);
      el.grid.focus({ preventScroll: true });
    } else if (e.key === 'F2') {
      e.preventDefault();
      startRename(tab);
    }
  });
  el.addSheet.addEventListener('click', addSheet);
}
