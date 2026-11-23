# cmd/web

`cmd/web` is a small web server that serves an in-browser XLSX editor. The
werkbook library runs in the browser as WebAssembly. The server only serves
static files. Workbook data does not leave the browser.

## Layout

| Path | Purpose |
| --- | --- |
| `main.go` | HTTP server. Serves the embedded `static/` directory. |
| `internal/bridge/` | Pure Go session logic (no `syscall/js`). Unit tested with the usual toolchain. |
| `wasm/` | `GOOS=js GOARCH=wasm` entry point. Thin `syscall/js` wrapper around `internal/bridge`. |
| `static/index.html`, `static/style.css`, `static/js/`, `static/css/` | The editor UI. Plain JavaScript (ES modules), no build step, no dependencies. |
| `static/werkbook.wasm`, `static/wasm_exec.js` | Build outputs. Not committed. |
| `smoke/smoke.mjs` | Node test of the JavaScript API contract against the built module. |
| `PLAN.md` | The plan of stage 2: design decisions, feature specifications, test plan. |

## Build and run

```bash
make web        # builds static/werkbook.wasm, copies wasm_exec.js, builds bin/werkbook-web
make web-run    # builds, then starts the server on http://localhost:8080
make web-smoke  # builds, then runs the Node smoke test of the contract (smoke/smoke.mjs)
```

Server flags: `-addr` (default `localhost:8080`) and `-dev <dir>` (serve the
static files from disk instead of the embedded copy).

## JavaScript API contract

The WASM module sets `globalThis.werkbook` and then calls
`globalThis.onWerkbookReady()` if that function is defined. The module holds
one workbook at a time.

All functions are synchronous. All rows and columns are 1-based integers.
Each function returns a plain object:

- Success: `{ ok: true, ...fields }`
- Failure: `{ ok: false, error: "message" }`

The functions do not throw for usual errors. A failure does not change the
workbook and does not change the undo history.

The contract is in five parts: the base functions in this section, and one
section for each milestone of stage 2. Where a later section changes a
function or a type of an earlier section, the later section is correct, and
the tables of this section show the final form.

### Function index

| Group | Functions | Section |
| --- | --- | --- |
| Workbook | `version`, `newWorkbook`, `open`, `save`, `sheets`, `addSheet`, `renameSheet`, `deleteSheet` | This section |
| Read | `getRange`, `sheetMeta` | This section |
| Edit cells | `setCell`, `setCells`, `clearRange`, `fill`, `copy`, `paste` | This section, milestone 1 |
| History | `history`, `undo`, `redo` | Milestone 1 |
| Formula editor | `functions`, `formulaContext`, `cycleRef` | Milestone 2 |
| Layout | `setColWidth`, `setRowHeight`, `setFreeze` | Milestone 2, milestone 4 |
| Navigation | `findEdge`, `currentRegion` | Milestone 2 |
| Structure | `insertRows`, `deleteRows`, `insertCols`, `deleteCols` | Milestone 3 |
| Format | `setStyle` | Milestone 3 |
| Data | `sortRange`, `find`, `replace`, `replaceAll` | Milestone 4 |

### Types

```ts
// maxRow and maxCol are the used range: the last row and column with a stored cell
// or a spilled value. A row that has only a custom height does not count.
type SheetInfo = { name: string; maxRow: number; maxCol: number; visible: boolean };

type Cell = {
  row: number;
  col: number;
  display: string;   // text to show in the grid (number format applied when possible)
  input: string;     // text to show in the editor: "=SUM(A1:A3)" for a formula, else the raw value
  type: "number" | "string" | "bool" | "error" | "empty";
  formula: boolean;  // true if the cell has a formula
  spill?: boolean;   // true if the value comes from a dynamic array spill (cell is read-only)
  style?: CellStyle; // omitted when the cell has the default style
};

type CellStyle = {
  bold?: boolean; italic?: boolean; underline?: boolean;
  color?: string;      // font color, "#RRGGBB"
  fill?: string;       // background color, "#RRGGBB"
  align?: "left" | "center" | "right";
  numFmt?: "number" | "currency" | "percent" | "date" | "custom"; // omitted for the general format (milestone 3)
};
```

`HistoryState` and `Range` are in the milestone 1 section.

### Functions

| Function | Result fields |
| --- | --- |
| `version()` | `version: string` |
| `newWorkbook()` | `sheets: SheetInfo[]` (one empty sheet), `history` (empty) |
| `open(bytes: Uint8Array)` | `sheets: SheetInfo[]`, `history` (empty). An `ArrayBuffer` is accepted too. |
| `sheets()` | `sheets: SheetInfo[]` |
| `getRange(sheet, row1, col1, row2, col2)` | `cells: Cell[]` (only non-empty cells, row-major order; an empty cell that has a visible `style` is included with `type: "empty"`; an empty cell with only a number format is not) |
| `sheetMeta(sheet)` | `colWidths: {[col: number]: number}` (pixels), `rowHeights: {[row: number]: number}` (pixels), `merges: {row1, col1, row2, col2}[]`, `warnings: string[]` (milestone 3), `freeze: {rows, cols}` (milestone 4) |
| `setCell(sheet, row, col, input: string)` | `history` |
| `setCells(sheet, edits: {row, col, input}[])` | `history` (one batch; one recalculation; one undo step; atomic) |
| `addSheet(name)` | `sheets: SheetInfo[]`, `history` |
| `renameSheet(oldName, newName)` | `sheets: SheetInfo[]`, `history` |
| `deleteSheet(name)` | `sheets: SheetInfo[]`, `history` (the call clears the history) |
| `save()` | `data: Uint8Array` (the XLSX file) |

### Input rules for `setCell` and `setCells`

- `""` clears the cell.
- Text that starts with `=` (and is longer than one character) is a formula.
- `TRUE` / `FALSE` (any letter case) is a boolean.
- Text that parses as a number is a number.
- Text that starts with `'` is a string without the apostrophe.
- All other text is a string.

After a write, formulas that depend on the cell can have new values. The UI
must call `getRange` again for the visible range.

## Contract additions: milestone 1 (undo, fill, clipboard)

A function that is not built yet is absent from `globalThis.werkbook`.
`make web-smoke` is the proof that a function is ready.

Rules for all functions in this section:

- A failure (`ok: false`) does not change the workbook and does not change the
  history. Each function is atomic.
- A `Range` argument is a plain object. The corners can be in any order. An
  unknown field or a missing field is an error. Ranges in results are
  normalized (`row1 <= row2`, `col1 <= col2`).
- After each successful mutation the UI reads `sheets()`, `sheetMeta()` (when
  the layout can change) and `getRange()` again.

### Types

```ts
type Range = { row1: number; col1: number; row2: number; col2: number }; // inclusive

type HistoryState = {
  canUndo: boolean;
  canRedo: boolean;
  undoLabel: string;  // "" when canUndo is false
  redoLabel: string;  // "" when canRedo is false
  depth: number;      // number of undo steps
  cleared?: boolean;  // true when this call cleared the history because the step was too large
};

type HistoryResult = {
  history: HistoryState;
  sheet: string;         // the sheet that the step changed (its name after the undo or redo)
  select: Range | null;  // the range to select; null for a step with no range (add sheet, rename sheet)
  sheets: SheetInfo[];
  label: string;         // the label of the step that was undone or redone
};
```

### Changes to functions that exist

| Function | Change |
| --- | --- |
| `setCell`, `setCells` | The result gains `history: HistoryState`. One call is one undo step. `setCells` is atomic: an invalid edit leaves all cells unchanged. |
| `addSheet`, `renameSheet` | The result gains `history`. One undo step each. |
| `deleteSheet` | The result gains `history`. The call clears the history (it cannot be undone), so the state is always empty. |
| `open`, `newWorkbook` | The result gains `history` (always empty). A failed `open` keeps the workbook and its history. |

All other results are unchanged.

### New functions

| Function | Result fields |
| --- | --- |
| `history()` | `history: HistoryState` |
| `undo()` | `HistoryResult`. Fails with `nothing to undo`. |
| `redo()` | `HistoryResult`. Fails with `nothing to redo`. |
| `clearRange(sheet, range: Range)` | `history`, `changed: number`, `skipped: number` (read-only spill cells). Values and formulas are removed; styles stay. |
| `fill(sheet, src: Range, dst: Range, series: boolean)` | `history`, `range: Range` (= `dst`), `changed: number`, `skipped: number` |
| `copy(sheet, range: Range, cut: boolean)` | `text: string` (TSV for the system clipboard), `id: string` (the id of the internal clipboard; `""` when `internal` is false), `cells: number` (non-empty cells in the text), `internal: boolean` (true when the bridge kept a clipboard snapshot), `history` (changes only when `cut` is true) |
| `paste(sheet, dst: Range, text: string, id?: string)` | `history`, `range: Range` (the range that was written), `changed: number`, `skipped: number`, `clipped: boolean`, `mode: "internal" \| "text"` |

### History

- One successful mutating call is one undo step. A call that changes nothing
  (for example `clearRange` on empty cells, or `setCell` with the value that
  the cell already has) makes no step and leaves the redo stack as it is.
- A new step clears the redo stack.
- `undo()` and `redo()` return the sheet and the range of the step, so that the
  UI can show the sheet, select the range and scroll to it.
- Labels: `Edit cell` (`setCell`), `Edit cells` (`setCells`), `Clear`, `Fill`,
  `Cut`, `Paste`, `Add sheet`, `Rename sheet`. The UI shows them as
  "Undo Fill".
- `open`, `newWorkbook` and `deleteSheet` clear the history.
- Spill results are not stored. An undo of a write that blocked a spill makes
  the anchor spill again.
- Limits: 100 steps; 500,000 changed cells in all steps together (the oldest
  steps are dropped first); 200,000 cells in one call (the call fails).

### `clearRange`

- Removes values and formulas. A cell that has a style stays as a styled empty
  cell.
- A spill follower is read-only: it is counted in `skipped`, unless its anchor
  is in the range too.
- A full-row or full-column range is valid; only cells that exist are touched.
- `changed` is the number of cells that had a value or a formula.

### `fill`

- `dst` must contain `src`, and must extend it on exactly one axis (down, up,
  right or left). `dst` equal to `src` changes nothing. A different shape
  fails.
- Each lane (a column for a vertical fill, a row for a horizontal fill) is
  filled independently:
  1. The source cells repeat as a pattern.
  2. A formula is copied with its relative references shifted by the distance
     from its source cell. A reference that leaves the sheet becomes `#REF!`.
  3. With `series: true`: if the lane has two or more source cells and all are
     number constants, the fill continues their linear trend (least squares;
     `1, 2` continues `3, 4`). If the lane has one source cell that is a
     number constant with a date format, each step adds 1 (up or left:
     subtracts 1). All other lanes copy.
  4. The style of the source cell is copied.
  5. A source cell that is a spill follower gives its value. Exception: when
     its anchor is in `src` too, the follower gives an empty cell, so that
     the copy of the anchor formula can spill.
- An empty source cell clears the destination cell.
- A destination cell that is a spill follower is skipped (`skipped`), unless
  its anchor is overwritten in the same call.
- `changed` is the number of destination cells that were written (the cells
  of `dst` that are not in `src` and were not skipped). `paste` counts in the
  same way.
- Fails when `src` or `dst` intersects a merged range, and when `dst` has more
  than 200,000 cells.
- Ctrl+D and Ctrl+R use `series: false`. For a selection of one row, the UI
  sends `src` = the cells above the selection and `dst` = `src` + selection.

### `copy` and `paste`

- `copy` returns TSV text: tab between cells, `\n` between rows, no line end
  after the last row. A cell text that has a tab, a line end or a quote is in
  quotes, with each quote doubled. A formula cell and a spill cell give their
  `display` text (the shown value). All other cells give their `input` text,
  so that a text paste reads back the same value and type: a number with the
  format `0%` gives `0.5`, not `50%`, and the text `007` gives `'007`.
- A full-row or full-column range is limited to the used range of the sheet.
- `copy` keeps an internal clipboard in the bridge: the source sheet, the
  source range, a snapshot of the cells, and the text. The snapshot does not
  change when the source cells change later. `internal` is false when the
  range has more than 200,000 non-empty cells; the text is still made. A range
  with more than 2,000,000 cells fails.
- `copy` with `cut: true` also clears the source cells (values and formulas;
  styles stay) as one undo step with the label `Cut`. Read-only spill
  followers are not cleared.
- `copy` returns an `id` for its internal clipboard. The UI writes it to the
  system clipboard next to the text, in the custom type
  `application/x-werkbook-clip`, and gives it back to `paste`:
  - `id` is the id of the internal clipboard: `mode` is `"internal"`. The
    text is not compared.
  - `id` is `""` (the system clipboard has no such type: the text is from a
    different program) or a different id (a different tab, an older copy):
    `mode` is `"text"`.
  - `id` is absent (`undefined` or `null`): the UI cannot read clipboard
    types. The bridge then compares `text` with the text of the internal
    clipboard (line ends normalized, one trailing line end ignored). This is
    the fallback only; a short text from a different program can be equal to
    the text of the internal clipboard.
- `mode: "internal"`: formulas are copied with their relative references
  shifted by the distance from the source cell, and styles are copied.
  `mode: "text"`: the text is read as TSV and each field goes through the
  input rules, as `setCells` does.
- An internal paste of a cut clipboard is a move: the formulas must read the
  same cells as before. The references do not shift. On a different sheet,
  each reference with no sheet name gets the name of the source sheet
  (`=A1*2` cut from `Sheet1` becomes `=Sheet1!A1*2`).
- An undo of the `Cut` step makes the clipboard a copy: the source cells are
  back, so a paste shifts the references. A redo makes it a cut again.
- Destination: if `dst` is a whole multiple of the clipboard size on both
  axes, the clipboard repeats to fill `dst`. If not, it is pasted one time at
  the top-left cell of `dst`.
- Data that does not fit on the sheet is cut; `clipped` is true.
- A paste to a different sheet keeps the reference text: a reference with no
  sheet name refers to the destination sheet.
- Internal paste: an empty clipboard cell clears the destination cell and
  gives it the style of the source. A clipboard that has a spill anchor and
  its followers pastes only the anchor formula. Followers with no anchor in
  the clipboard paste as values.
- Text paste: an empty field clears the destination cell. The styles of the
  destination stay. Text that is empty as a whole (also only a line end) is
  "nothing to paste": the call changes nothing, `changed` is 0 and `range` is
  the top-left cell of `dst`.
- A destination cell that is a spill follower is skipped (`skipped`), unless
  its anchor is overwritten in the same call.
- Merged ranges: merged cells have no cell-by-cell copy. When the copied
  range or the destination intersects a merged range, the paste uses the text
  path and returns `mode: "text"`: the values are pasted, formulas and styles
  are not. A text paste writes to the top-left cell of a merged range and
  skips the cells that the merge covers.
- A destination with more than 200,000 cells fails.
- `open`, `newWorkbook`, and `deleteSheet` of the source sheet clear the
  internal clipboard. A later paste of the same text is then a text paste.

## Contract additions: milestone 2 (formula help, layout, navigation)

The rules of milestone 1 apply: a failure changes nothing, object arguments
are strict, and each successful mutation is one undo step.

### Types

```ts
type FunctionInfo = {
  name: string;         // upper case
  signature: string;    // "SUMIF(range, criteria, [sum_range])"; "NAME(...)" when there is no authored record
  summary: string;      // one sentence; "" when there is no authored record
  args: { name: string; optional: boolean; repeat: boolean }[]; // [] when there is no authored record
  minArgs: number;      // the minimum that the engine accepts
  variadic: boolean;
};

type RefSpan = {
  start: number; end: number;   // offsets in the editor text (UTF-16 code units, "=" included)
  text: string;                 // the reference as written, with its sheet qualifier
  sheet: string;                // "" = the sheet of the edited cell
  sheetEnd?: string;            // only for a 3D reference (Sheet1:Sheet3!A1): the second sheet
  kind: "cell" | "range" | "cols" | "rows";
  range: Range;                 // normalized; full columns use rows 1..1048576; full rows use columns 1..16384
  color: number;                // 0-based index; equal sheet + range = equal index
};

type FormulaContext = {
  prefix: { start: number; end: number; text: string } | null; // a function-name prefix that ends at the caret
  call: { name: string; argIndex: number; nameStart: number } | null; // the innermost open function call at the caret
  refs: RefSpan[];
};
```

### New functions

| Function | Result fields |
| --- | --- |
| `functions()` | `functions: FunctionInfo[]` (all functions that the engine has, sorted by name) |
| `formulaContext(text: string, caret: number)` | `FormulaContext` |
| `cycleRef(text: string, selStart: number, selEnd: number)` | `text: string`, `selStart: number`, `selEnd: number`, `changed: boolean` |
| `setColWidth(sheet, col1, col2, px: number)` | `history`, `range: Range` (the full columns), `clamped: boolean` (always false) |
| `setRowHeight(sheet, row1, row2, px: number)` | `history`, `range: Range` (the full rows that the call covers), `clamped: boolean` |
| `findEdge(sheet, row, col, dRow, dCol)` | `row: number`, `col: number` |
| `currentRegion(sheet, row, col)` | `range: Range`, `header: boolean` |

### `functions`

- The list has one record for each function of the engine, and for the forms
  that the parser handles itself (`LET`, `LAMBDA`, `MAP`, `REDUCE`, `SCAN`,
  `BYROW`, `BYCOL`, `MAKEARRAY`). The UI calls it one time and keeps the list.
- For a function with an authored record, `minArgs` is the number of
  arguments of the signature that are not optional, and `variadic` is true
  when an argument repeats (`SUMIF`: 2, false; `SUM`: 1, true). For a function
  with no authored record the two values are an estimate from the engine
  registry; do not use them to mark a call as wrong.
- A signature has the form `NAME(arg, [optional], repeat1, [repeat2], ...)`.
  `args` is derived from the signature. An argument with `repeat: true` can
  occur again; for a hint, the last argument stays current when the user types
  more arguments than the list has.
- A function with no authored record has `signature: "NAME(...)"`,
  `summary: ""` and `args: []`.

### `formulaContext`

- `text` is the full editor text. `caret` is an offset in UTF-16 code units
  (the value of `selectionStart`). All offsets in the result have the same
  unit and count the leading `=`.
- The function never fails for a string and a number. Text that does not
  start with `=` gives `prefix: null, call: null, refs: []`. A caret out of
  the text is moved to the nearest end.
- `prefix`: the name that ends at the caret, when it can be the start of a
  function name: it starts with a letter and has only letters, digits, `.` and
  `_`. It is `null` when the caret is in a string, in the middle of a name
  (the next character continues the name), at the end of a name that already
  has its `(`, after a number, or after a sheet qualifier (`Sheet1!A`). `=A1` gives the prefix `A1`; the UI shows no list
  because no function name starts with it.
- `call`: the innermost function call that is open at the caret. `name` is in
  upper case. `nameStart` is the offset of the function name. `argIndex` is
  the 0-based index of the argument at the caret. Commas in strings, in array
  literals `{1,2}`, in nested calls and in parentheses that are not a call do
  not count. A caret in a text argument is in the call of that argument
  (`=VLOOKUP("ab` gives `VLOOKUP`, argument 0), so that the hint stays while
  the user types the text.
- `refs`: all complete references of the text, in text sequence. References
  in strings, structured references (`Table1[Col]`), names and references to
  other workbooks are not in the list. `=SUM(A1:` gives `A1`.
- `color`: references with the same sheet and the same range have the same
  index. Indexes are given in the sequence of first occurrence, with no upper
  limit; the UI takes the index modulo its number of colors. The sheet
  compare ignores letter case; `""` is a sheet of its own (the UI knows which
  sheet the edited cell is on).
- A 3D reference (`Sheet1:Sheet3!A1`) has `sheet` = the first sheet and
  `sheetEnd` = the second sheet. The UI draws its outline on each sheet of
  the span, in tab order.

### `cycleRef` (F4)

- `text`, `selStart` and `selEnd` are the editor text and its selection
  (UTF-16 code units, `=` included). The result has the new text and the new
  selection.
- With no selection (`selStart == selEnd`), the reference that the caret is
  in or immediately after (or immediately before) changes. With a selection,
  all references that the selection touches change together.
- The cycle is `A1` -> `$A$1` -> `A$1` -> `$A1` -> `A1`. A range changes both
  ends together, from the state of its first end. `A:A` <-> `$A:$A`.
  `1:3` <-> `$1:$3`.
- After the change with no selection, the caret is at the end of the
  reference. With a selection, the new selection goes from the start of the
  first changed reference to the end of the last.
- `changed` is false when there is no reference at the caret or the text does
  not start with `=`; `text` and the selection are then returned unchanged.

### `setColWidth` and `setRowHeight`

- `setColWidth(sheet, col1, col2, px)` gives all columns `col1..col2` the
  width `px` (CSS pixels). `px <= 0` removes the custom width (the column gets
  the default width). Else `px` is limited to 12..2000.
- `setRowHeight(sheet, row1, row2, px)` is the same for rows, with the limits
  8..1000. `px <= 0` removes the custom height.
- The numbers can be in any order. One call is one undo step (labels
  `Column width`, `Row height`). `undo()` returns `select` = the full columns
  or the full rows.
- All columns in one call are valid (`setColWidth(sheet, 1, 16384, px)`).
- All rows in one call are valid too:
  - A removal (`px <= 0`) changes only the rows that have a custom height. It
    has no limit.
  - A height for more than 200,000 rows (for example a selection of all rows)
    is limited to the used range of the sheet. The result then has
    `clamped: true` and `range` = the rows that got the height. The call
    fails only when the used range itself has more than 200,000 rows.
- A size change does not start a recalculation: the values of volatile
  formulas (`RAND`, `NOW`) stay.
- `sheetMeta` reports the same pixel number that was set: the size is stored
  in the XLSX as `(px - 5) / 7` character units for a width and `px * 0.75`
  points for a height.
- A call that changes no size makes no undo step.

### `findEdge` (Ctrl+arrow)

- Exactly one of `dRow`, `dCol` is `-1` or `1`; the other is `0`.
- From the start cell, in the direction of the arrow:
  1. If the start cell is not empty and the next cell is not empty: the
     result is the last cell of the run of non-empty cells.
  2. Else: the result is the next non-empty cell.
  3. If there is none: the result is the edge of the sheet (row 1 or
     1,048,576; column 1 or 16,384). The UI limits it to its grid.
- "Not empty" = a value, a formula (also when its result is `""`), or a spill
  follower with a value. A cell with only a style is empty.
- Merged ranges: a merged range counts as one cell, and it is non-empty when
  its top-left cell is non-empty. The scan starts from the edge of the merged
  range of the start cell. A result in a merged range is the top-left cell of
  that range.

## Contract additions: milestone 3 (structure and format)

### Types

```ts
// Addition to CellStyle (in getRange results).
type CellStyle = {
  bold?: boolean; italic?: boolean; underline?: boolean;
  color?: string; fill?: string;
  align?: "left" | "center" | "right";
  numFmt?: NumFmtPreset | "custom"; // omitted for the general format
};

type NumFmtPreset = "general" | "number" | "currency" | "percent" | "date";

type StylePatch = {            // only the fields that are present change
  bold?: boolean; italic?: boolean; underline?: boolean;
  color?: string | null;       // "#RRGGBB"; null removes the font color
  fill?: string | null;        // "#RRGGBB"; null removes the fill
  align?: "left" | "center" | "right" | "general";
  numFmt?: NumFmtPreset;
};
```

### Changes to functions that exist

| Function | Change |
| --- | --- |
| `getRange` | `Cell.style` can have `numFmt`. A cell whose only style is a number format now has a `style` (before, `style` was omitted for it). An empty cell with only a number format is still not in the result. |
| `sheetMeta` | The result gains `warnings: string[]` (see "Structure"). It is `[]` for most sheets. |

### New functions

| Function | Result fields |
| --- | --- |
| `insertRows(sheet, row, count)` | `history`, `sheets: SheetInfo[]`. The new empty rows are `row .. row+count-1`. |
| `deleteRows(sheet, row, count)` | `history`, `sheets`. Removes the rows `row .. row+count-1`. |
| `insertCols(sheet, col, count)` | `history`, `sheets` |
| `deleteCols(sheet, col, count)` | `history`, `sheets` |
| `setStyle(sheet, range: Range, patch: StylePatch)` | `history`, `changed: number`, `range: Range` (the range that was formatted), `clamped: boolean` |

### Structure: insert and delete of rows and columns

- One call is one undo step (labels `Insert rows`, `Delete rows`,
  `Insert columns`, `Delete columns`). `undo()` restores the workbook exactly:
  formulas that became `#REF!`, removed cells, merged ranges, tables, defined
  names and sizes. `undo()` and `redo()` return `select` = the full rows or
  columns of the call.
- "Insert below" and "insert right" are the same call with `row` (or `col`)
  one more than the last selected row (column).
- New rows and columns are empty and have no style and no custom size.
- What moves with the rows or columns: cells, styles, row heights, hidden
  flags, column widths, merged ranges, tables, defined names, spill ranges.
- Formulas in all sheets and defined names are adjusted, as in Excel:

  | Reference | Insert `count` at `at` | Delete `at .. at+count-1` |
  | --- | --- | --- |
  | Cell | At or after `at`: moves by `count` | In the band: `#REF!`. After the band: moves back |
  | Range | Start at or after `at`: both ends move. `at` in the range: the range grows | All in the band: `#REF!`. Part in the band: the range shrinks. After the band: moves back |
  | Absolute (`$A$1`) | Adjusted as a relative reference is | Same |
  | Full column on a row edit, full row on a column edit | No change | No change |
  | No sheet name | Adjusted only in formulas on the edited sheet | Same |

- A deleted reference becomes `#REF!` with no sheet name (`=Sheet2!A5`
  becomes `=#REF!`). Other text of the formula does not change.
- Merged ranges move, grow or shrink with the range rule. A merged range that
  shrinks to one cell or to nothing is removed.
- Tables move, grow or shrink. A column inserted in a table gets a unique
  name (`Column1`, ...). A table that is fully in a deleted band is removed.
- Errors (the workbook and the history do not change):
  - `row`/`col` or `count` out of range (`count` must be 1 or more, and
    `row + count - 1` must be on the sheet).
  - An insert that would push a non-empty cell, a merged range or a table off
    the sheet.
  - A delete that removes the header row of a table but not the full table
    (`cannot delete the header row of table "X"`).
  - A delete that removes all data rows of a table but not the full table
    (`cannot delete all data rows of table "X"`). Excel permits this and
    keeps one empty data row; this stage does not.
- A structured reference to a table column that a column delete removed
  (`Tbl[Qty]`, `Tbl[[#This Row],[Qty]]`) becomes `#REF!`.
- Limits of this stage, reported in `sheetMeta.warnings`: conditional formats,
  data validation, hyperlinks, filters, drawings and charts are kept in the
  file as they are and do **not** move with rows and columns. `warnings` has
  one short text for each such kind that the sheet has, for example
  `"conditional formats"`. The UI shows one information message at the first
  structural edit of a sheet that has warnings.
- A step that is too large for the history (more than 500,000 removed cells
  and changed formulas) still runs. The history is then cleared and the
  result has `history.cleared: true`; the step cannot be undone.

### `setStyle`

- The patch changes only the fields that it has. All other parts of the cell
  style stay (borders, font name and size, vertical alignment, wrap, a custom
  number format when `numFmt` is absent). An unknown field, a wrong type and
  a color that is not `#RRGGBB` are errors.
- `bold`, `italic`, `underline`: `true` sets, `false` removes.
- `color`, `fill`: `"#RRGGBB"` sets; `null` removes.
- `align`: `"general"` removes the horizontal alignment.
- `numFmt` presets:

  | Preset | Format |
  | --- | --- |
  | `general` | General |
  | `number` | `#,##0.00` |
  | `currency` | `"$"#,##0.00` |
  | `percent` | `0%` |
  | `date` | `m/d/yyyy` |

  `getRange` reports the preset of a cell in `style.numFmt`, or `"custom"` for
  all other formats. General is not reported.
- A negative number with a date or time format has the `display` text
  `########`, as in Excel.
- A cell that does not exist is made (empty, with the style). A cell whose
  style becomes the default style and that is empty is removed.
- A range of full rows, full columns or the full sheet is limited to the used
  range of the sheet; `clamped` is then true and `range` is the limited range.
  When the sheet is empty, nothing changes, `changed` is 0 and `range` is the
  top-left cell of the request. Column and row formats are not supported.
- `changed` is the number of cells whose style changed. A call that changes
  no style makes no undo step.
- One call is one undo step with the label `Format`.
- At most 200,000 cells in one call.
- Merged ranges: the style goes to all cells of the range. Spill followers: a
  style on a follower is valid; it does not block the spill.

## Contract additions: milestone 4 (freeze, sort, find)

### Types

```ts
type SortKey = { col: number; desc?: boolean }; // col is a sheet column in the range

type FindOptions = {
  matchCase?: boolean;   // default false
  wholeCell?: boolean;   // default false
  lookIn?: "input" | "display"; // default "input" (the formula text for a formula cell)
  from?: { row: number; col: number }; // the search starts after this cell; default: before A1
  backward?: boolean;    // default false
};
```

### Changes to functions that exist

| Function | Change |
| --- | --- |
| `sheetMeta` | The result gains `freeze: { rows: number; cols: number }` (`0` = not frozen on that axis). |
| `insertRows`, `deleteRows`, `insertCols`, `deleteCols` | The freeze follows: an insert in the frozen rows or columns makes the frozen area larger, and a delete makes it smaller. An undo restores it. |

### New functions

| Function | Result fields |
| --- | --- |
| `setFreeze(sheet, rows, cols)` | `history` |
| `sortRange(sheet, range: Range, keys: SortKey[], header: boolean)` | `history`, `changed: number`, `range: Range` |
| `find(sheet, query: string, opts?: FindOptions)` | `found: boolean`, `row: number`, `col: number`, `count: number`, `wrapped: boolean` |
| `replace(sheet, row, col, query: string, replacement: string, opts?: FindOptions)` | `history`, `replaced: boolean` |
| `replaceAll(sheet, query: string, replacement: string, opts?: FindOptions)` | `history`, `replaced: number`, `skipped: number` |

### `setFreeze`

- `rows` rows at the top and `cols` columns at the left stay in view. `0, 0`
  removes the freeze. `rows` must be 0..1,048,575 and `cols` 0..16,383.
- One call is one undo step (label `Freeze panes`). A call that does not
  change the freeze makes no step. `undo()` returns `select: null`.
- The freeze is saved in the XLSX and read from it. A file that has a split
  that is not frozen is shown as not frozen, and its split is kept in the
  file as long as `setFreeze` is not called for that sheet.
- A freeze does not start a recalculation.

### `sortRange`

- The rows of `range` are sorted by the key columns. `keys` must have one or
  more keys; each `col` must be a column of `range`; the first key is the
  most important one. `header: true` leaves the first row of the range in
  its position.
- Order (Excel): numbers, then text, then `FALSE`, `TRUE`, then errors.
- Text order: letter case is not important; a Latin letter with an accent
  sorts with its base letter (`Émile` is between `b` and `Zeta`); in a text,
  digits are before symbols and symbols are before letters; letters of other
  scripts are after the Latin letters. The rule of Excel that ignores a
  hyphen and an apostrophe is not applied. Errors keep their sequence. Empty cells are
  last for an ascending and for a descending key. The sort is stable: rows
  that are equal on all keys keep their sequence.
- The key is the value that the cell shows (the result of a formula).
- Full rows of the range move: values, formulas and styles. A formula moves
  as a copy does: its relative references shift by the row distance.
- A full-row or full-column range is limited to the used range.
- Fails, with no change, when the range intersects a merged range, has a
  spill anchor, a spill follower or a legacy array formula, or has more than
  200,000 cells.
- `changed` is the number of cells whose content changed. A range that is
  already in order makes no undo step. The label is `Sort`. `range` is the
  range that was sorted (with the header row).

### `find`, `replace`, `replaceAll`

- The search is on one sheet, in row-major order, over the cells that are
  not empty. The match is literal text (no wildcards). An empty `query`
  fails.
- `lookIn: "input"` (default) searches the editor text of each cell: the
  formula with its `=` for a formula cell, else the value. A leading
  apostrophe of the editor text (`'007`) is not part of the searched text.
  `lookIn: "display"` searches the text that the grid shows.
- `matchCase: false` (default) ignores letter case. `wholeCell: true`
  matches only when the full text is equal to the query.
- `find` starts after `from` (before A1 when `from` is absent) and goes
  forward, or backward with `backward: true`. It wraps at the end of the
  sheet; `wrapped` is then true. A sheet with one match gives that cell
  again, with `wrapped: true`, also when `from` is that cell.
- `count` is the number of matching cells on the sheet, at most 10,000. With
  no match: `found: false`, `row: 0`, `col: 0`, `count: 0`.
- A spill follower can be found (its text is its value) but not replaced.
- `replace` changes the cell at `row`, `col` when it matches: each
  occurrence of the query in the editor text becomes `replacement`
  (`wholeCell`: the full text). The new text goes through the input rules,
  as an edit does: text that becomes a number is a number, and a formula
  stays a formula. `replaced` is false, with no change and no undo step,
  when the cell does not match or is a spill follower. Replace always
  searches the editor text; `lookIn` is ignored.
- `replaceAll` does the same for all matching cells as one undo step.
  `skipped` is the number of matching spill followers. `from` and `backward`
  are ignored.
- A text cell that looks like a different type keeps its type when the
  replacement does not change that: `'007` with `7` -> `8` gives the text
  `008`.
- Labels: `Replace`, `Replace all`. `undo()` selects the changed cell, or
  the bounding range of all changed cells.
- A merged range is searched in its top-left cell.

### `currentRegion` (used by sort)

- `range` is the smallest rectangle around the cell that has an empty border
  on all sides (the Ctrl+A region of Excel): it grows while a row or column
  next to it, with the corners, has a non-empty cell. An empty cell with no
  non-empty neighbor gives the range of the cell itself.
- `header` is true when the region has two or more rows, each cell of the
  first row that is not empty is text, the first row has one or more text
  cells, and the second row has a cell that is not empty and not text.
