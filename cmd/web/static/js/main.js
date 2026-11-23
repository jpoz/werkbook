// Werkbook web editor: entry point. Plain JavaScript ES modules, no
// dependencies, no build step. index.html loads only this module.
//
// The sequence of the init() calls is fixed. Handlers of one bus event and
// of one key run in this sequence.

import * as core from './core.js';
import * as workbook from './workbook.js';
import * as sheets from './sheets.js';
import * as grid from './grid.js';
import * as edit from './edit.js';
import * as clipboard from './clipboard.js';
import * as history from './history.js';
import * as fill from './fill.js';
import * as resize from './resize.js';
import * as nav from './nav.js';
import * as structure from './structure.js';
import * as freeze from './freeze.js';
import * as ui from './ui.js';
import * as fnhelp from './fnhelp.js';
import * as refcolor from './refcolor.js';
import * as format from './format.js';
import * as sort from './sort.js';
import * as find from './find.js';

const modules = [
  // Base modules.
  ['core', core],
  ['sheets', sheets],
  ['edit', edit],
  ['clipboard', clipboard],
  ['grid', grid],
  ['workbook', workbook],
  // Feature modules.
  ['ui', ui],
  ['history', history],
  ['fill', fill],
  ['resize', resize],
  ['nav', nav],
  ['structure', structure],
  ['freeze', freeze],
  ['fnhelp', fnhelp],
  ['refcolor', refcolor],
  ['format', format],
  ['sort', sort],
  ['find', find],
];

for (const [name, mod] of modules) {
  try {
    mod.init();
  } catch (err) {
    // A feature module that fails must not stop the editor.
    console.error(`init of ${name}.js failed`, err);
  }
}

workbook.boot();
