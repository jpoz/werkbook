// Function autocomplete and argument hints (PLAN.md 4.3). Owner: fe-b.
//
// The function list and the formula context come from the engine
// (functions, formulaContext). This module has no formula parser. While the
// engine does not have these functions, the module does nothing.

import { S, el, rawCall, on, registerKey } from './core.js';
import { activeInput, setEditText } from './edit.js';
import { placeBelow } from './ui.js';

// ------------------------------------------------------------- pure logic

// matchFunctions returns the records whose name starts with the prefix
// (no letter case), sorted by name.
export function matchFunctions(list, prefix) {
  if (!prefix || !Array.isArray(list)) return [];
  const p = String(prefix).toUpperCase();
  const out = list.filter((f) => f && typeof f.name === 'string' && f.name.toUpperCase().startsWith(p));
  out.sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
  return out;
}

// currentArg returns the index in info.args of the argument to show in bold,
// or -1. When the call has more arguments than the record and the last
// argument of the record repeats, the last argument stays current.
export function currentArg(info, argIndex) {
  const args = info && Array.isArray(info.args) ? info.args : [];
  if (!args.length || !(argIndex >= 0)) return -1;
  if (argIndex < args.length) return argIndex;
  return args[args.length - 1].repeat ? args.length - 1 : -1;
}

// acceptEdit replaces the prefix with "NAME(". If a "(" already follows the
// prefix, it is not written two times. The caret goes after the parenthesis.
export function acceptEdit(text, prefix, name) {
  const paren = text[prefix.end] === '(' ? 1 : 0;
  const head = `${text.slice(0, prefix.start)}${name}(`;
  return { text: head + text.slice(prefix.end + paren), caret: head.length };
}

// ------------------------------------------------------- engine data, cached

const memo = { text: null, caret: -1, ctx: null };

// editContext returns the FormulaContext of the open formula edit, or null
// (no edit, text that is not a formula, or formulaContext is not available).
// One engine call is made for each (text, caret) pair; fnhelp.js and
// refcolor.js share the result (PLAN.md risk R11).
export function editContext() {
  if (!S.edit) return null;
  const inp = activeInput();
  const text = inp.value;
  if (text[0] !== '=') return null;
  const caret = inp.selectionStart == null ? text.length : inp.selectionStart;
  if (memo.text === text && memo.caret === caret) return memo.ctx;
  // rawCall: a function that is not available must not show a message on each key.
  const res = rawCall('formulaContext', [text, caret]);
  memo.text = text;
  memo.caret = caret;
  memo.ctx = res.ok ? res : null;
  return memo.ctx;
}

let byName = null; // upper-case name -> FunctionInfo

// loadFunctions reads the function list one time. It tries again later when
// the engine does not have the function yet.
function loadFunctions() {
  if (byName) return true;
  const res = rawCall('functions', []);
  if (!res.ok || !Array.isArray(res.functions)) return false;
  S.functions = res.functions;
  byName = new Map();
  for (const f of res.functions) {
    if (f && typeof f.name === 'string') byName.set(f.name.toUpperCase(), f);
  }
  return true;
}

// ------------------------------------------------------------------- state

let popup = null;
let hint = null;

const H = {
  prefixText: null, // the upper-case prefix of the list on screen
  matches: [],
  index: 0,
  dismissedText: null, // the editor text at the time Escape closed the list
  hintKey: null,
};

const listOpen = () => Boolean(popup) && !popup.hidden;

// ---------------------------------------------------------------- the list

function selectRow(index) {
  H.index = Math.min(H.matches.length - 1, Math.max(0, index));
  const rows = popup.children;
  for (let i = 0; i < rows.length; i++) rows[i].setAttribute('aria-selected', i === H.index ? 'true' : 'false');
  const row = rows[H.index];
  if (row) row.scrollIntoView({ block: 'nearest' });
}

function showList(matches, prefixText) {
  const key = prefixText.toUpperCase();
  const same = listOpen() && H.prefixText === key;
  H.matches = matches;
  if (!same) {
    H.prefixText = key;
    const frag = document.createDocumentFragment();
    matches.forEach((f, i) => {
      const row = document.createElement('div');
      row.className = 'fn-row';
      row.id = `fn-opt-${i}`;
      row.setAttribute('role', 'option');
      row.dataset.fn = f.name;
      const name = document.createElement('span');
      name.className = 'fn-name';
      name.textContent = f.name;
      row.append(name);
      if (f.summary) {
        const summary = document.createElement('span');
        summary.className = 'fn-summary';
        summary.textContent = f.summary;
        row.append(summary);
      }
      frag.append(row);
    });
    popup.replaceChildren(frag);
    popup.hidden = false;
    popup.scrollTop = 0;
  }
  // A new list starts at the first row. A list that did not change keeps its row.
  selectRow(same ? H.index : 0);
}

function hideList() {
  if (!popup || popup.hidden) return;
  popup.hidden = true;
  popup.replaceChildren();
  H.prefixText = null;
  H.matches = [];
  H.index = 0;
}

// ---------------------------------------------------------------- the hint

function showHint(info, argIndex) {
  const key = `${info.name}:${argIndex}`;
  if (!hint.hidden && H.hintKey === key) return;
  H.hintKey = key;
  hint.dataset.fn = info.name;
  hint.dataset.argIndex = String(argIndex);
  const args = Array.isArray(info.args) ? info.args : [];
  const frag = document.createDocumentFragment();
  if (!args.length) {
    // No authored record ("NAME(...)") or a function with no arguments.
    frag.append(info.signature || `${info.name}(...)`);
  } else {
    const current = currentArg(info, argIndex);
    const name = document.createElement('span');
    name.className = 'fn-hint-name';
    name.textContent = info.name;
    frag.append(name, '(');
    args.forEach((arg, i) => {
      if (i) frag.append(', ');
      const span = document.createElement('span');
      span.className = i === current ? 'fn-arg current' : 'fn-arg';
      span.dataset.arg = String(i);
      span.textContent = arg.optional ? `[${arg.name}]` : arg.name;
      frag.append(span);
    });
    if (args.some((arg) => arg.repeat)) frag.append(', ...');
    frag.append(')');
  }
  hint.replaceChildren(frag);
  hint.title = info.summary || '';
  hint.hidden = false;
}

function hideHint() {
  if (!hint || hint.hidden) return;
  hint.hidden = true;
  hint.replaceChildren();
  H.hintKey = null;
  delete hint.dataset.fn;
  delete hint.dataset.argIndex;
}

// ------------------------------------------------------------ show and place

// place puts the list below the active input and the hint below the list.
// When the list had to go above the input, the hint goes above the list.
function place() {
  if (!S.edit) return;
  let anchor = activeInput().getBoundingClientRect();
  let prefer = 'below';
  if (listOpen()) {
    prefer = placeBelow(popup, anchor);
    anchor = popup.getBoundingClientRect();
  }
  if (!hint.hidden) placeBelow(hint, anchor, { prefer });
}

function closeAll() {
  hideList();
  hideHint();
}

function update() {
  if (!popup || !hint) return;
  const ctx = editContext();
  if (!ctx) {
    closeAll();
    return;
  }
  const inp = activeInput();
  const text = inp.value;
  if (H.dismissedText !== null && H.dismissedText !== text) H.dismissedText = null;

  let matches = [];
  const prefix = ctx.prefix;
  const collapsed = inp.selectionStart === inp.selectionEnd;
  if (collapsed && prefix && prefix.text && H.dismissedText === null && loadFunctions()) {
    matches = matchFunctions(S.functions, prefix.text);
  }
  if (matches.length) showList(matches, prefix.text);
  else hideList();

  const call = ctx.call;
  const info = call && call.name && loadFunctions() ? byName.get(String(call.name).toUpperCase()) : null;
  if (info) showHint(info, call.argIndex);
  else hideHint();

  place();
}

function reset() {
  closeAll();
  H.dismissedText = null;
  memo.text = null;
  memo.caret = -1;
  memo.ctx = null;
}

// -------------------------------------------------------------------- accept

function accept(name) {
  const ctx = editContext();
  if (!S.edit || !ctx || !ctx.prefix) return;
  const next = acceptEdit(activeInput().value, ctx.prefix, name);
  hideList();
  setEditText(next.text, next.caret, next.caret); // emits 'edit:change'
  update();
}

// ---------------------------------------------------------------------- keys
//
// runKey (core.js) prevents the default behavior when a handler does not
// return false. Each handler returns false when the list is closed, so that
// the usual key of the editor runs (point mode, commit, cancel).

function keyMove(step) {
  return () => {
    if (!listOpen()) return false;
    selectRow(H.index + step);
    return true;
  };
}

function keyAccept() {
  if (!listOpen()) return false;
  const f = H.matches[H.index];
  if (!f) return false;
  accept(f.name);
  return true;
}

function keyDismiss() {
  if (!listOpen()) return false;
  // The list stays closed until the text changes. The edit stays open.
  H.dismissedText = activeInput().value;
  hideList();
  place();
  return true;
}

// ---------------------------------------------------------------------- init

export function init() {
  popup = el.fnPopup;
  hint = el.fnHint;
  if (!popup || !hint) return;

  // A press on the list must not take the focus from the editor: a blur commits the edit.
  popup.addEventListener('mousedown', (e) => e.preventDefault());
  popup.addEventListener('click', (e) => {
    const row = e.target.closest('[data-fn]');
    if (row && popup.contains(row)) accept(row.dataset.fn);
  });
  hint.addEventListener('mousedown', (e) => e.preventDefault());

  on('workbook', loadFunctions);
  on('edit:start', update);
  // edit.js emits 'edit:change' for each change of the text, the caret, the
  // active input and the sheet on screen.
  on('edit:change', update);
  on('edit:end', reset);
  el.grid.addEventListener('scroll', () => {
    if (listOpen() || !hint.hidden) place();
  }, { passive: true });
  window.addEventListener('resize', () => {
    if (listOpen() || !hint.hidden) place();
  });

  registerKey('edit', { key: 'ArrowDown' }, keyMove(1));
  registerKey('edit', { key: 'ArrowUp' }, keyMove(-1));
  registerKey('edit', { key: 'Tab' }, keyAccept);
  registerKey('edit', { key: 'Escape' }, keyDismiss);
}
