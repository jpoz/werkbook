// Clipboard: copy, cut and paste through the clipboard events, and the clear
// of the selection. The bridge makes and reads the TSV text and keeps an
// internal clipboard, so that a paste in the editor keeps formulas and styles
// (cmd/web/README.md, milestone 1).

import { S, clamp, showStatus, api, mutate } from './core.js';
import { selRange, isSingleSelection, mergeAt, setSelection, gridHasFocus } from './grid.js';

const toRange = (R) => ({ row1: R.r1, col1: R.c1, row2: R.r2, col2: R.c2 });

function spillNote(n) {
  return `${n} spilled ${n === 1 ? 'cell is' : 'cells are'} read-only and did not change`;
}

// clearSelection removes the values and formulas of the selection. It is one
// bridge call and one undo step.
export function clearSelection() {
  const res = mutate('clearRange', S.sheet, toRange(selRange()));
  if (!res.ok) return;
  if (res.skipped) showStatus(`${spillNote(res.skipped)}.`, 'info');
}

// pasteDestination is the range that the user pasted to. One selected cell
// is one cell, also when it is a merged cell: the bridge then pastes the
// clipboard one time at that cell.
function pasteDestination() {
  if (isSingleSelection() || isOneMergedCell()) {
    return { row1: S.anchor.row, col1: S.anchor.col, row2: S.anchor.row, col2: S.anchor.col };
  }
  return toRange(selRange());
}

function isOneMergedCell() {
  const m = mergeAt(S.anchor.row, S.anchor.col);
  if (!m) return false;
  const R = selRange();
  return R.r1 === m.row1 && R.c1 === m.col1 && R.r2 === m.row2 && R.c2 === m.col2;
}

// The custom clipboard type that carries the id of the internal clipboard of
// the bridge (cmd/web/README.md, `copy` and `paste`).
const CLIP_TYPE = 'application/x-werkbook-clip';

// writeClipboard puts the text of a copy result on the system clipboard,
// and the id of the internal clipboard next to it. An id of "" (the bridge
// kept no internal clipboard) is not written: a paste then finds no such
// type and is a text paste.
function writeClipboard(e, res) {
  e.clipboardData.setData('text/plain', res.text == null ? '' : String(res.text));
  if (typeof res.id === 'string' && res.id !== '') {
    try {
      e.clipboardData.setData(CLIP_TYPE, res.id);
    } catch (err) {
      // The browser does not accept the custom type. A paste then has no id.
    }
  }
}

// readClipId returns the id for paste:
// - the id that a copy in an editor wrote (this tab or a different tab);
// - "" when the clipboard has no such type: the text is from a different program;
// - undefined when the browser cannot report the clipboard types. The
//   bridge then compares the text (the fallback of the README).
function readClipId(data) {
  let types;
  try {
    types = data.types;
  } catch (err) {
    return undefined;
  }
  if (!types || typeof types.length !== 'number') return undefined;
  if (!Array.from(types).includes(CLIP_TYPE)) return '';
  try {
    const id = data.getData(CLIP_TYPE);
    return typeof id === 'string' ? id : undefined;
  } catch (err) {
    return undefined;
  }
}

function pasteText(text, id) {
  if (!text) return;
  const args = ['paste', S.sheet, pasteDestination(), text];
  // An absent id is not sent, so the bridge sees it as absent.
  if (id !== undefined) args.push(id);
  const res = mutate(...args);
  if (!res.ok) return;
  const R = res.range;
  if (R) {
    const anchor = { row: clamp(R.row1, 1, S.rowLimit), col: clamp(R.col1, 1, S.colLimit) };
    const focus = { row: clamp(R.row2, 1, S.rowLimit), col: clamp(R.col2, 1, S.colLimit) };
    setSelection(anchor, focus, { scroll: false });
  }
  const notes = [];
  if (res.skipped) notes.push(spillNote(res.skipped));
  if (res.clipped) notes.push('the data did not fit in the sheet and was cut');
  if (notes.length) showStatus(`Paste: ${notes.join('; ')}.`, 'info');
}

export function init() {
  document.addEventListener('copy', (e) => {
    if (!gridHasFocus()) return;
    // The default copy would put the page selection on the clipboard.
    e.preventDefault();
    const res = api('copy', S.sheet, toRange(selRange()), false);
    if (!res.ok) return;
    writeClipboard(e, res);
  });
  document.addEventListener('cut', (e) => {
    if (!gridHasFocus()) return;
    e.preventDefault();
    // A cut clears the source cells at once, as one undo step.
    const res = mutate('copy', S.sheet, toRange(selRange()), true);
    if (!res.ok) return;
    writeClipboard(e, res);
  });
  document.addEventListener('paste', (e) => {
    if (!gridHasFocus()) return;
    e.preventDefault();
    pasteText(e.clipboardData.getData('text/plain'), readClipId(e.clipboardData));
  });
}
