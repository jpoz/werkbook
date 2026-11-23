// Format controls (PLAN.md 4.8). Owner: fe-b.
//
// The toolbar group #tb-format, the keys Ctrl+B, Ctrl+I, Ctrl+U, and the
// state of the controls from the active cell. Each command is one setStyle
// call (one undo step). While the engine has no setStyle, the controls are
// disabled and the keys are not used.

import { S, el, mutate, on, registerKey, showStatus, isHexColor } from './core.js';
import { cellAt, selRange } from './grid.js';
import { toolButton, setPressed, engineHas } from './ui.js';

const PRESETS = [
  ['general', 'General'],
  ['number', 'Number'],
  ['currency', 'Currency'],
  ['percent', 'Percent'],
  ['date', 'Date'],
];
const ALIGNS = [
  ['left', 'Align left', '⇤'],
  ['center', 'Align center', '↔'],
  ['right', 'Align right', '⇥'],
];
const TOGGLES = [
  ['bold', 'Bold', 'B'],
  ['italic', 'Italic', 'I'],
  ['underline', 'Underline', 'U'],
];
const DEFAULT_FILL = '#ffffff';
const DEFAULT_COLOR = '#000000';
const CLAMP_MESSAGE = 'The format applies to the used range. Column and row formats are not supported yet.';

const C = { toggles: {}, aligns: {}, all: [] }; // the controls
let built = false;

const focusGrid = () => el.grid.focus({ preventScroll: true });

// activeStyle returns the CellStyle of the active cell ({} for the default style).
function activeStyle() {
  if (S.state !== 'editor' || !S.sheet) return {};
  const cell = cellAt(S.anchor.row, S.anchor.col);
  return (cell && cell.style) || {};
}

const canFormat = () => S.state === 'editor' && Boolean(S.sheet) && !S.edit && engineHas('setStyle');

// apply sends one patch for the selection. Full-row and full-column
// selections go to the engine as they are; the engine limits them to the
// used range and reports that with clamped.
function apply(patch) {
  if (!canFormat()) return false;
  const R = selRange();
  const res = mutate('setStyle', S.sheet, { row1: R.r1, col1: R.c1, row2: R.r2, col2: R.c2 }, patch);
  if (res.ok && res.clamped) showStatus(CLAMP_MESSAGE, 'info');
  refresh(); // after a failure the controls show the state of the cell again
  return res.ok;
}

// The active cell decides the direction of a toggle: if the active cell is
// bold, the command removes bold from the full selection (Excel behavior).
const toggle = (field) => apply({ [field]: !activeStyle()[field] });

// An align button that is already active sets the alignment back to general.
const align = (value) => apply({ align: activeStyle().align === value ? 'general' : value });

function refresh() {
  if (!built) return;
  const st = activeStyle();
  const enabled = canFormat();
  for (const [field] of TOGGLES) setPressed(C.toggles[field], Boolean(st[field]));
  for (const [value] of ALIGNS) setPressed(C.aligns[value], st.align === value);
  const numFmt = typeof st.numFmt === 'string' && st.numFmt ? st.numFmt : 'general';
  // "custom" is a state only: the option shows when the cell has a format that is not a preset.
  C.custom.hidden = numFmt !== 'custom';
  C.numFmt.value = numFmt;
  C.fill.value = isHexColor(st.fill) ? st.fill.toLowerCase() : DEFAULT_FILL;
  C.color.value = isHexColor(st.color) ? st.color.toLowerCase() : DEFAULT_COLOR;
  for (const node of C.all) node.disabled = !enabled;
}

function colorControl(mount, id, text, label, field) {
  const wrap = document.createElement('label');
  wrap.className = 'fmt-color';
  wrap.title = label;
  const tag = document.createElement('span');
  tag.textContent = text;
  const input = document.createElement('input');
  input.type = 'color';
  input.id = id;
  input.setAttribute('aria-label', label);
  // 'change' comes one time when the user closes the color picker, so one
  // color selection is one undo step.
  input.addEventListener('change', () => {
    if (isHexColor(input.value)) apply({ [field]: input.value.toUpperCase() });
    focusGrid();
  });
  wrap.append(tag, input);
  mount.append(wrap);
  const clearBtn = toolButton(mount, {
    id: `${id}-clear`,
    label: '×',
    title: `Remove the ${label.toLowerCase()}`,
    ariaLabel: `Remove the ${label.toLowerCase()}`,
    onClick: () => apply({ [field]: null }),
  });
  C.all.push(input, clearBtn);
  return input;
}

function build() {
  const mount = el.tbFormat;
  if (!mount) return false;
  for (const [field, label, text] of TOGGLES) {
    C.toggles[field] = toolButton(mount, {
      id: `fmt-${field}`,
      label: text,
      title: `${label} (Ctrl+${text})`,
      ariaLabel: label,
      pressed: false,
      onClick: () => toggle(field),
    });
    C.all.push(C.toggles[field]);
  }

  const select = document.createElement('select');
  select.id = 'fmt-numfmt';
  select.className = 'fmt-select';
  select.title = 'Number format';
  select.setAttribute('aria-label', 'Number format');
  for (const [value, label] of PRESETS) {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = label;
    select.append(opt);
  }
  C.custom = document.createElement('option');
  C.custom.value = 'custom';
  C.custom.textContent = 'Custom';
  C.custom.disabled = true;
  C.custom.hidden = true;
  select.append(C.custom);
  select.addEventListener('change', () => {
    if (select.value !== 'custom') apply({ numFmt: select.value });
    focusGrid();
  });
  mount.append(select);
  C.numFmt = select;
  C.all.push(select);

  C.fill = colorControl(mount, 'fmt-fill', 'Fill', 'Fill color', 'fill');
  C.color = colorControl(mount, 'fmt-color', 'Text', 'Font color', 'color');

  for (const [value, label, text] of ALIGNS) {
    C.aligns[value] = toolButton(mount, {
      id: `fmt-align-${value}`,
      label: text,
      title: label,
      ariaLabel: label,
      pressed: false,
      onClick: () => align(value),
    });
    C.all.push(C.aligns[value]);
  }
  return true;
}

// keyToggle is a grid key. It returns false while there is no setStyle, so
// the key is not used.
function keyToggle(field) {
  return () => {
    if (!canFormat()) return false;
    toggle(field);
    return true;
  };
}

export function init() {
  built = build();
  if (!built) return;
  refresh();
  // grid.js emits 'selection' after each selection change and after each
  // write (afterWrite calls updateSelectionUI), so the state follows both.
  on('selection', refresh);
  on('sheet', refresh);
  on('workbook', refresh);
  on('edit:start', refresh);
  on('edit:end', refresh);

  registerKey('grid', { key: 'b', mod: true }, keyToggle('bold'));
  registerKey('grid', { key: 'i', mod: true }, keyToggle('italic'));
  registerKey('grid', { key: 'u', mod: true }, keyToggle('underline'));
}
