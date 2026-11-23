// Find and replace (PLAN.md 4.9). Owner: fe-b.
//
// The button #find-open, the key Ctrl+F and the panel #find-bar. The search
// and the replace are done by the engine (find, replace, replaceAll). While
// the engine has no find, the button is disabled and Ctrl+F is not used.

import { S, el, rawCall, mutate, on, registerKey, showStatus } from './core.js';
import { setSelection } from './grid.js';
import { toolButton, engineHas } from './ui.js';

const COUNT_LIMIT = 10000; // find counts at most this number of matches

const C = {}; // the controls of the bar
let built = false;

const usable = () => S.state === 'editor' && Boolean(S.sheet) && !S.edit && engineHas('find');
const isOpen = () => built && !el.findBar.hidden;

// options returns the FindOptions that the two checkboxes give. The engine
// rejects unknown fields, so only fields of the contract are sent.
function options(extra) {
  const o = { ...extra };
  if (C.matchCase.checked) o.matchCase = true;
  if (C.wholeCell.checked) o.wholeCell = true;
  return o;
}

function say(text, kind) {
  C.count.textContent = text;
  if (kind) C.count.dataset.kind = kind;
  else delete C.count.dataset.kind;
}

const plural = (n, one, many) => `${n} ${n === 1 ? one : many}`;

function queryText() {
  const q = C.find.value;
  if (q === '') say('Type the text to find.', 'info');
  return q;
}

// search calls find from the active cell. It shows an error in the bar and
// returns null when the call fails.
function search(query, backward) {
  const opts = options({ from: { row: S.anchor.row, col: S.anchor.col } });
  if (backward) opts.backward = true;
  const res = rawCall('find', [S.sheet, query, opts]);
  if (!res.ok) {
    say(res.error, 'error');
    return null;
  }
  return res;
}

function countText(res) {
  if (!res.found || !(res.count > 0)) return 'No matches';
  return res.count >= COUNT_LIMIT ? `${COUNT_LIMIT} or more matches` : plural(res.count, 'match', 'matches');
}

// showCount shows the number of matches after a change of the text or of an
// option. It does not move the selection.
function showCount() {
  const query = C.find.value;
  if (query === '' || !usable()) {
    say('');
    return;
  }
  const res = search(query, false);
  if (res) say(countText(res), res.found ? undefined : 'info');
}

// findNext selects the next match after the active cell (or the match
// before it). The focus stays in the bar, so Enter can find again.
function findNext(backward) {
  if (!usable()) return false;
  const query = queryText();
  if (query === '') return false;
  const res = search(query, backward);
  if (!res) return false;
  if (!res.found) {
    say(countText(res), 'info');
    return false;
  }
  if (res.row > S.rowLimit || res.col > S.colLimit) {
    say('The match is out of the rows and columns that the grid can show.', 'error');
    return false;
  }
  const cell = { row: res.row, col: res.col };
  setSelection(cell, cell);
  const n = countText(res);
  say(res.wrapped ? `${n}. The search started again from the ${backward ? 'end' : 'start'}.` : n);
  return true;
}

// replaceOne replaces in the active cell if it matches, then finds the next match.
function replaceOne() {
  if (!usable() || !engineHas('replace')) return;
  const query = queryText();
  if (query === '') return;
  const res = mutate('replace', S.sheet, S.anchor.row, S.anchor.col, query, C.replace.value, options());
  if (!res.ok) return;
  findNext(false);
}

// replaceAll is one undo step.
function replaceAll() {
  if (!usable() || !engineHas('replaceAll')) return;
  const query = queryText();
  if (query === '') return;
  const res = mutate('replaceAll', S.sheet, query, C.replace.value, options());
  if (!res.ok) return;
  let text = `Replaced the text in ${plural(res.replaced, 'cell', 'cells')}.`;
  if (res.skipped) text += ` ${plural(res.skipped, 'spilled cell is', 'spilled cells are')} read-only and did not change.`;
  say(text);
  showStatus(text, 'info');
}

function open() {
  if (!built || !usable()) return false;
  refresh();
  el.findBar.hidden = false;
  C.find.focus({ preventScroll: true });
  C.find.select();
  return true;
}

function close() {
  if (!isOpen()) return;
  el.findBar.hidden = true;
  say('');
  if (S.state === 'editor') el.grid.focus({ preventScroll: true });
}

function refresh() {
  if (!built) return;
  C.open.disabled = !usable();
  C.replaceOne.disabled = !engineHas('replace');
  C.replaceAll.disabled = !engineHas('replaceAll');
}

// ---------------------------------------------------------------------- build

function textInput(id, label) {
  const input = document.createElement('input');
  input.type = 'text';
  input.id = id;
  input.placeholder = label;
  input.spellcheck = false;
  input.autocomplete = 'off';
  input.setAttribute('aria-label', label);
  return input;
}

function button(id, label, ariaLabel, onClick) {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'btn';
  btn.id = id;
  btn.textContent = label;
  if (ariaLabel) {
    btn.setAttribute('aria-label', ariaLabel);
    btn.title = ariaLabel;
  }
  btn.addEventListener('click', onClick);
  return btn;
}

function checkbox(id, label) {
  const wrap = document.createElement('label');
  wrap.className = 'find-check';
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.id = id;
  wrap.append(input, label);
  return { wrap, input };
}

function row(...nodes) {
  const div = document.createElement('div');
  div.className = 'find-row';
  div.append(...nodes);
  return div;
}

function build() {
  const bar = el.findBar;
  if (!bar || !el.tbData) return false;
  C.open = toolButton(el.tbData, {
    id: 'find-open',
    label: 'Find',
    title: 'Find and replace (Ctrl+F)',
    onClick: () => (isOpen() ? close() : open()),
  });

  bar.setAttribute('role', 'search');
  bar.setAttribute('aria-label', 'Find and replace');
  C.find = textInput('find-input', 'Find');
  C.replace = textInput('replace-input', 'Replace with');
  const matchCase = checkbox('find-case', 'Match case');
  const wholeCell = checkbox('find-whole', 'Whole cell');
  C.matchCase = matchCase.input;
  C.wholeCell = wholeCell.input;
  C.replaceOne = button('replace-one', 'Replace', null, replaceOne);
  C.replaceAll = button('replace-all', 'Replace all', null, replaceAll);
  C.count = document.createElement('div');
  C.count.id = 'find-count';
  C.count.setAttribute('role', 'status');
  C.count.setAttribute('aria-live', 'polite');
  bar.replaceChildren(
    row(
      C.find,
      button('find-prev', '↑', 'Find previous', () => findNext(true)),
      button('find-next', '↓', 'Find next', () => findNext(false)),
      button('find-close', '×', 'Close', close),
    ),
    row(C.replace, C.replaceOne, C.replaceAll),
    row(matchCase.wrap, wholeCell.wrap),
    C.count,
  );

  bar.addEventListener('keydown', (e) => {
    if (e.isComposing) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    } else if (e.key === 'Enter' && e.target === C.find) {
      e.preventDefault();
      findNext(e.shiftKey);
    } else if (e.key === 'Enter' && e.target === C.replace) {
      e.preventDefault();
      replaceOne();
    } else if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && e.key.toLowerCase() === 'f') {
      // Ctrl+F in the bar goes back to the find input.
      e.preventDefault();
      C.find.focus({ preventScroll: true });
      C.find.select();
    }
  });
  // A new query or a new option gives a new number of matches.
  C.find.addEventListener('input', showCount);
  C.matchCase.addEventListener('change', showCount);
  C.wholeCell.addEventListener('change', showCount);
  return true;
}

export function init() {
  built = build();
  if (!built) return;
  refresh();
  on('workbook', () => {
    say('');
    refresh();
  });
  on('sheet', () => {
    say('');
    refresh();
  });
  on('edit:start', refresh);
  on('edit:end', refresh);

  // false while there is no find: the find of the browser then runs.
  registerKey('grid', { key: 'f', mod: true }, () => open());
}
