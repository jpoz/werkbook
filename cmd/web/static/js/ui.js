// Shared helpers for popups, the context menu and toolbar buttons
// (PLAN.md 6.2). Owner: fe-b.
// This module has no imports, so each feature module can use it.

const VIEW_PAD = 4; // minimum distance between a popup and the viewport edge

const byId = (id) => document.getElementById(id);
const clamp = (n, lo, hi) => Math.min(hi, Math.max(lo, n));

// engineHas reports whether the WASM module has a contract function. A
// function that is not built yet is absent (PLAN.md 2.5). A feature uses this
// check to stay inactive, with no error message, until its function exists.
export function engineHas(name) {
  return Boolean(globalThis.werkbook) && typeof globalThis.werkbook[name] === 'function';
}

// ------------------------------------------------------------ popup placement

function viewport() {
  const root = document.documentElement;
  return { width: root.clientWidth, height: root.clientHeight };
}

// placeBelow puts a node below a rectangle (viewport coordinates, as
// getBoundingClientRect gives). The node gets position: fixed. If there is
// not sufficient room below and there is more room above, the node goes
// above. The node stays in the viewport.
//
// The node must not be hidden when this function runs: the function reads
// its size. opts: { gap?: number, prefer?: 'below' | 'above' }.
// The result is the side that was used: 'below' or 'above'.
export function placeBelow(node, anchorRect, opts) {
  const o = opts || {};
  const gap = typeof o.gap === 'number' ? o.gap : 2;
  node.style.position = 'fixed';
  node.style.left = '0px';
  node.style.top = '0px';
  const w = node.offsetWidth;
  const h = node.offsetHeight;
  const view = viewport();
  const roomBelow = view.height - anchorRect.bottom - gap - VIEW_PAD;
  const roomAbove = anchorRect.top - gap - VIEW_PAD;
  let side = o.prefer === 'above' ? 'above' : 'below';
  if (side === 'below' && h > roomBelow && roomAbove > roomBelow) side = 'above';
  else if (side === 'above' && h > roomAbove && roomBelow > roomAbove) side = 'below';
  const top = side === 'below' ? anchorRect.bottom + gap : anchorRect.top - gap - h;
  node.style.top = `${clamp(top, VIEW_PAD, Math.max(VIEW_PAD, view.height - h - VIEW_PAD))}px`;
  node.style.left = `${clamp(anchorRect.left, VIEW_PAD, Math.max(VIEW_PAD, view.width - w - VIEW_PAD))}px`;
  node.dataset.side = side;
  return side;
}

// placeAt puts a node with its top-left corner at a viewport point and keeps
// it in the viewport.
function placeAt(node, x, y) {
  node.style.position = 'fixed';
  node.style.left = '0px';
  node.style.top = '0px';
  const w = node.offsetWidth;
  const h = node.offsetHeight;
  const view = viewport();
  node.style.left = `${clamp(x, VIEW_PAD, Math.max(VIEW_PAD, view.width - w - VIEW_PAD))}px`;
  node.style.top = `${clamp(y, VIEW_PAD, Math.max(VIEW_PAD, view.height - h - VIEW_PAD))}px`;
}

// --------------------------------------------------------------- context menu

const menu = {
  node: null,
  bound: false,
  open: false,
  handlers: new Map(), // item node -> onSelect
  previous: null, // the element that had the focus before the menu opened
};

const menuItems = () => Array.from(menu.node.querySelectorAll('[data-cmd]:not([disabled])'));

function hideMenu(restoreFocus) {
  if (!menu.open) return;
  menu.open = false;
  menu.node.hidden = true;
  menu.node.replaceChildren();
  menu.handlers.clear();
  const previous = menu.previous;
  menu.previous = null;
  if (restoreFocus && previous && document.contains(previous) && typeof previous.focus === 'function') {
    previous.focus({ preventScroll: true });
  }
}

function focusMenuItem(step, from) {
  const list = menuItems();
  if (!list.length) return;
  let i;
  if (from === 'first') i = 0;
  else if (from === 'last') i = list.length - 1;
  else {
    const cur = list.indexOf(document.activeElement);
    i = cur < 0 ? (step > 0 ? 0 : list.length - 1) : (cur + step + list.length) % list.length;
  }
  list[i].focus({ preventScroll: true });
}

function bindMenu() {
  if (menu.bound) return;
  menu.bound = true;
  const node = menu.node;
  node.addEventListener('click', (e) => {
    const item = e.target.closest('[data-cmd]');
    if (!item || item.disabled || !node.contains(item)) return;
    const onSelect = menu.handlers.get(item);
    // Close first: the command can open a dialog or change the focus.
    hideMenu(true);
    if (typeof onSelect === 'function') onSelect(item.dataset.cmd);
  });
  node.addEventListener('keydown', (e) => {
    // The menu owns the keyboard while it is open.
    e.stopPropagation();
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      focusMenuItem(1);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      focusMenuItem(-1);
    } else if (e.key === 'Home') {
      e.preventDefault();
      focusMenuItem(0, 'first');
    } else if (e.key === 'End') {
      e.preventDefault();
      focusMenuItem(0, 'last');
    } else if (e.key === 'Escape' || e.key === 'Tab') {
      e.preventDefault();
      hideMenu(true);
    }
  });
  node.addEventListener('contextmenu', (e) => e.preventDefault());
  // A press out of the menu closes it. The press keeps its usual effect.
  document.addEventListener('mousedown', (e) => {
    if (menu.open && !node.contains(e.target)) hideMenu(false);
  }, true);
  document.addEventListener('scroll', (e) => {
    if (menu.open && !node.contains(e.target)) hideMenu(true);
  }, true);
  window.addEventListener('resize', () => hideMenu(true));
  window.addEventListener('blur', () => hideMenu(true));
}

// openMenu shows #context-menu at a viewport point (clientX, clientY).
// items: an array of
//   { cmd: string, label: string, onSelect?: (cmd) => void, disabled?: boolean }
//   { separator: true }
// Each item is a button with role="menuitem" and data-cmd. The menu takes the
// focus; closeMenu gives the focus back. It returns false when there is no
// #context-menu element or no item.
export function openMenu(items, x, y) {
  if (!menu.node) menu.node = byId('context-menu');
  if (!menu.node || !Array.isArray(items)) return false;
  bindMenu();
  const previous = menu.open ? menu.previous : document.activeElement;
  hideMenu(false);

  const frag = document.createDocumentFragment();
  let count = 0;
  for (const item of items) {
    if (!item) continue;
    if (item.separator) {
      const sep = document.createElement('div');
      sep.className = 'menu-sep';
      sep.setAttribute('role', 'separator');
      frag.append(sep);
      continue;
    }
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'menu-item';
    btn.setAttribute('role', 'menuitem');
    btn.tabIndex = -1;
    btn.dataset.cmd = String(item.cmd);
    btn.textContent = String(item.label);
    btn.disabled = Boolean(item.disabled);
    menu.handlers.set(btn, item.onSelect);
    frag.append(btn);
    count++;
  }
  if (!count) return false;

  menu.node.replaceChildren(frag);
  menu.node.hidden = false;
  menu.open = true;
  menu.previous = previous;
  placeAt(menu.node, x, y);
  const first = menuItems()[0];
  if (first) first.focus({ preventScroll: true });
  else {
    menu.node.tabIndex = -1;
    menu.node.focus({ preventScroll: true });
  }
  return true;
}

// closeMenu hides #context-menu and gives the focus back to the element
// that had it before.
export function closeMenu() {
  hideMenu(true);
}

export const menuIsOpen = () => menu.open;

// ------------------------------------------------------------ toolbar buttons

// toolButton makes a button in a toolbar mount and returns it.
// mount: an element or the id of an element.
// spec: { id, label, title?, ariaLabel?, onClick, pressed?: boolean }.
// When pressed is given, the button is a toggle button (aria-pressed).
// A mouse press on the button does not move the focus, so the grid (or the
// open editor) keeps the focus and the keyboard continues to work there.
export function toolButton(mount, spec) {
  const parent = typeof mount === 'string' ? byId(mount) : mount;
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'btn tool';
  if (spec.id) btn.id = spec.id;
  btn.textContent = spec.label == null ? '' : String(spec.label);
  if (spec.title) btn.title = spec.title;
  if (spec.ariaLabel) btn.setAttribute('aria-label', spec.ariaLabel);
  if (spec.pressed !== undefined) setPressed(btn, spec.pressed);
  btn.addEventListener('mousedown', (e) => e.preventDefault());
  if (typeof spec.onClick === 'function') btn.addEventListener('click', (e) => spec.onClick(e, btn));
  if (parent) parent.append(btn);
  return btn;
}

// setPressed sets the state of a toggle button.
export function setPressed(btn, pressed) {
  btn.setAttribute('aria-pressed', pressed ? 'true' : 'false');
}

export function init() {}
