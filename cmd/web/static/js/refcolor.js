// Colored reference outlines and the F4 absolute cycle (PLAN.md 4.4).
// Owner: fe-b.
//
// The references come from the engine (formulaContext). The cycle is done by
// the engine (cycleRef). This module has no formula parser. While the engine
// does not have these functions, the module does nothing.

import { S, el, api, on, registerKey } from './core.js';
import { rectOf, mergeAt } from './grid.js';
import { activeInput, setEditText } from './edit.js';
import { editContext } from './fnhelp.js';
import { engineHas } from './ui.js';

const COLORS = 6;

let layer = null;

// resolveSheet returns the name of the sheet that a reference refers to, or
// null when the workbook has no such sheet. "" is the sheet of the edited cell.
function resolveSheet(name) {
  if (!name) return S.edit ? S.edit.sheet : null;
  const lower = String(name).toLowerCase();
  const found = S.sheets.find((s) => s.name.toLowerCase() === lower);
  return found ? found.name : null;
}

// livePoint returns the reference that point mode inserted, while the text
// of the input is still the text of that insert.
function livePoint() {
  const p = S.edit && S.edit.point;
  if (!p) return null;
  return activeInput().value === p.text ? p : null;
}

// outlineRect returns the rectangle of a reference on the sheet on screen,
// limited to the grid extent, or null when no part of it is in the extent.
function outlineRect(ref) {
  const g = ref.range;
  if (!g) return null;
  let r1 = Math.min(g.row1, g.row2);
  let r2 = Math.max(g.row1, g.row2);
  let c1 = Math.min(g.col1, g.col2);
  let c2 = Math.max(g.col1, g.col2);
  if (!(r1 >= 1) || !(c1 >= 1) || r1 > S.rowCount || c1 > S.colCount) return null;
  if (ref.kind === 'cell') {
    // A reference to one cell of a merged range shows the full merged range.
    const m = mergeAt(r1, c1);
    if (m) {
      r1 = m.row1;
      c1 = m.col1;
      r2 = m.row2;
      c2 = m.col2;
    }
  }
  r2 = Math.min(r2, S.rowCount);
  c2 = Math.min(c2, S.colCount);
  const rect = rectOf(r1, c1, r2, c2);
  return rect.width > 0 && rect.height > 0 ? rect : null;
}

function clear() {
  if (layer.firstChild) layer.replaceChildren();
  delete el.pointRange.dataset.color;
}

function draw() {
  if (!layer) return;
  const ctx = S.edit && S.rows && S.sheet ? editContext() : null;
  const refs = ctx && Array.isArray(ctx.refs) ? ctx.refs : [];
  if (!refs.length) {
    clear();
    return;
  }
  const point = livePoint();
  let pointColor = null;
  const frag = document.createDocumentFragment();
  for (const ref of refs) {
    const color = ((Number(ref.color) % COLORS) + COLORS) % COLORS;
    if (point && ref.start === point.start && ref.end === point.end) {
      // #point-range shows this reference; it takes the color.
      pointColor = color;
      continue;
    }
    const sheet = resolveSheet(ref.sheet);
    if (sheet === null || sheet !== S.sheet) continue;
    const rect = outlineRect(ref);
    if (!rect) continue;
    const node = document.createElement('div');
    node.className = 'ref-outline';
    node.dataset.ref = ref.text;
    node.dataset.sheet = sheet;
    node.dataset.color = String(color);
    node.style.cssText = `left:${rect.left - 1}px;top:${rect.top - 1}px;width:${rect.width + 1}px;height:${rect.height + 1}px;`;
    frag.append(node);
  }
  layer.replaceChildren(frag);
  if (pointColor === null) delete el.pointRange.dataset.color;
  else el.pointRange.dataset.color = String(pointColor);
}

// cycle is the F4 command. It returns false (the key is not used) when the
// engine has no cycleRef or the text is not a formula.
function cycle() {
  if (!S.edit || !engineHas('cycleRef')) return false;
  const inp = activeInput();
  const text = inp.value;
  if (text[0] !== '=') return false;
  const selStart = inp.selectionStart;
  const selEnd = inp.selectionEnd;
  const res = api('cycleRef', text, selStart, selEnd);
  if (!res.ok || !res.changed || typeof res.text !== 'string') return true;
  // Point mode stays valid for a reference that point mode inserted: the
  // caret was at the end of that reference, and it is there again after the
  // cycle. setEditText moves the record of the inserted reference.
  const p = S.edit.point;
  const keep = p && p.text === text && selStart === p.end && selEnd === p.end && res.selStart === res.selEnd;
  setEditText(res.text, res.selStart, res.selEnd, keep ? { start: p.start, end: res.selEnd } : undefined);
  draw();
  return true;
}

export function init() {
  layer = el.refLayer;
  if (!layer) return;
  on('edit:start', draw);
  // edit.js emits 'edit:change' also when a different sheet comes on screen
  // during an edit.
  on('edit:change', draw);
  on('edit:end', draw);
  on('sheet', draw);
  on('render', draw); // the grid extent can grow; full-column outlines follow it

  registerKey('edit', { key: 'F4' }, cycle);
  if (/Mac|iPhone|iPad/.test(navigator.platform || '')) registerKey('edit', { key: 't', mod: true }, cycle);
}
