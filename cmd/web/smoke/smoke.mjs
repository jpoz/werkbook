// Smoke test of the built WASM module against the contract in ../README.md.
// Usage: node cmd/web/smoke/smoke.mjs <static dir>   (or: make web-smoke)
//
// It checks the wrapper layer that the Go unit tests cannot see: argument
// decoding, JSON shapes, Uint8Array results, and error results that do not
// throw. It prints one line for each check and exits with code 1 on the first
// failure.

import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

const dir = path.resolve(process.argv[2] || "cmd/web/static");
const TIMEOUT_MS = 60_000;

const timer = setTimeout(() => {
  console.log("FAIL timeout: the module was not ready or the checks did not end in 60 s");
  process.exit(1);
}, TIMEOUT_MS);

let count = 0;
function check(name, fn) {
  try {
    fn();
  } catch (err) {
    console.log(`FAIL ${name}`);
    console.log(String(err && err.message ? err.message : err));
    process.exit(1);
  }
  count++;
  console.log(`ok   ${name}`);
}

// ok asserts a success result and returns it.
function ok(res) {
  assert.equal(typeof res, "object");
  assert.equal(res.ok, true, `expected ok: true, got ${JSON.stringify(res)}`);
  return res;
}

// fails asserts a failure result with an error text that matches pattern.
function fails(res, pattern) {
  assert.equal(typeof res, "object");
  assert.equal(res.ok, false, `expected ok: false, got ${JSON.stringify(res)}`);
  assert.equal(typeof res.error, "string");
  if (pattern) assert.match(res.error, pattern);
  return res;
}

function cellAt(cells, row, col) {
  return cells.find((c) => c.row === row && c.col === col);
}

function R(row1, col1, row2, col2) {
  return { row1, col1, row2, col2 };
}

function assertHistory(h) {
  assert.equal(typeof h, "object", "the result has no history field");
  assert.equal(typeof h.canUndo, "boolean");
  assert.equal(typeof h.canRedo, "boolean");
  assert.equal(typeof h.undoLabel, "string");
  assert.equal(typeof h.redoLabel, "string");
  assert.equal(typeof h.depth, "number");
  assert.equal(h.canUndo, h.depth > 0);
}

// ---------------------------------------------------------------------------
// Step S0: the contract before stage 2.
// ---------------------------------------------------------------------------

function checksS0(wb) {
  check("S0 each function exists", () => {
    const names = [
      "version", "newWorkbook", "open", "sheets", "getRange", "sheetMeta",
      "setCell", "setCells", "addSheet", "renameSheet", "deleteSheet", "save",
    ];
    for (const n of names) assert.equal(typeof wb[n], "function", `${n} is missing`);
  });

  check("S0 version", () => {
    assert.equal(typeof ok(wb.version()).version, "string");
  });

  check("S0 newWorkbook gives one empty sheet", () => {
    const r = ok(wb.newWorkbook());
    assert.deepEqual(r.sheets, [{ name: "Sheet1", maxRow: 0, maxCol: 0, visible: true }]);
    assert.deepEqual(ok(wb.getRange("Sheet1", 1, 1, 50, 20)).cells, []);
  });

  check("S0 setCell and setCells", () => {
    ok(wb.setCell("Sheet1", 1, 1, "2"));
    ok(wb.setCells("Sheet1", [
      { row: 2, col: 1, input: "3" },
      { row: 3, col: 1, input: "=SUM(A1:A2)" },
      { row: 1, col: 3, input: "=SEQUENCE(2)" },
    ]));
  });

  check("S0 getRange cell shapes", () => {
    const cells = ok(wb.getRange("Sheet1", 1, 1, 50, 20)).cells;
    assert.deepEqual(cellAt(cells, 3, 1), {
      row: 3, col: 1, display: "5", input: "=SUM(A1:A2)", type: "number", formula: true,
    });
    assert.deepEqual(cellAt(cells, 2, 3), {
      row: 2, col: 3, display: "2", input: "2", type: "number", formula: false, spill: true,
    });
  });

  check("S0 a write recalculates dependents", () => {
    ok(wb.setCell("Sheet1", 1, 1, "10"));
    assert.equal(ok(wb.getRange("Sheet1", 3, 1, 3, 1)).cells[0].display, "13");
  });

  check("S0 sheetMeta shape", () => {
    const r = ok(wb.sheetMeta("Sheet1"));
    assert.equal(typeof r.colWidths, "object");
    assert.equal(typeof r.rowHeights, "object");
    assert.ok(Array.isArray(r.merges));
  });

  check("S0 addSheet, renameSheet, deleteSheet", () => {
    assert.equal(ok(wb.addSheet("Two")).sheets.length, 2);
    assert.equal(ok(wb.renameSheet("Two", "Second")).sheets[1].name, "Second");
    assert.equal(ok(wb.deleteSheet("Second")).sheets.length, 1);
    assert.deepEqual(ok(wb.sheets()).sheets.map((s) => s.name), ["Sheet1"]);
  });

  check("S0 usual errors give ok: false and do not throw", () => {
    fails(wb.getRange("Nope", 1, 1, 2, 2), /not found/);
    fails(wb.setCell("Sheet1", 0, 1, "x"), /out of range/);
    fails(wb.deleteSheet("Sheet1"));
    fails(wb.open(new Uint8Array([1, 2, 3])));
  });

  check("S0 a wrong argument type gives ok: false", () => {
    fails(wb.setCell("Sheet1", 1.5, 1, "x"), /integer/);
    fails(wb.setCell("Sheet1", 1, 1), /missing/);
    fails(wb.setCell("Sheet1", "1", 1, "x"));
    fails(wb.setCell(1, 1, 1, "x"));
    fails(wb.setCells("Sheet1", [{ row: 1 }]));
    fails(wb.setCells("Sheet1", "x"));
    fails(wb.getRange("Sheet1", 1, 1, 2));
    fails(wb.open("x"));
    fails(wb.addSheet(5));
  });

  check("S0 save then open gives the same cells", () => {
    const before = ok(wb.getRange("Sheet1", 1, 1, 50, 20)).cells;
    const saved = ok(wb.save());
    assert.ok(saved.data instanceof Uint8Array);
    assert.equal(saved.data[0], 0x50);
    assert.equal(saved.data[1], 0x4b);
    ok(wb.newWorkbook());
    const opened = ok(wb.open(saved.data));
    assert.equal(opened.sheets[0].maxRow, 3);
    assert.deepEqual(ok(wb.getRange("Sheet1", 1, 1, 50, 20)).cells, before);
    // An ArrayBuffer is accepted too.
    const buf = saved.data.buffer.slice(saved.data.byteOffset, saved.data.byteOffset + saved.data.byteLength);
    ok(wb.open(buf));
  });
}

// ---------------------------------------------------------------------------
// Milestone 1: history, clearRange, fill, copy, paste.
// ---------------------------------------------------------------------------

function checksM1(wb) {
  check("M1 each function exists", () => {
    for (const n of ["history", "undo", "redo", "clearRange", "fill", "copy", "paste"]) {
      assert.equal(typeof wb[n], "function", `${n} is missing`);
    }
  });

  check("M1 newWorkbook and open give an empty history", () => {
    const r = ok(wb.newWorkbook());
    assertHistory(r.history);
    assert.deepEqual(r.history, { canUndo: false, canRedo: false, undoLabel: "", redoLabel: "", depth: 0 });
    assert.deepEqual(ok(wb.history()).history, r.history);
    fails(wb.undo(), /nothing to undo/);
    fails(wb.redo(), /nothing to redo/);
  });

  check("M1 history shape after each write", () => {
    let r = ok(wb.setCell("Sheet1", 1, 1, "1"));
    assertHistory(r.history);
    assert.equal(r.history.depth, 1);
    assert.equal(r.history.canRedo, false);
    assert.notEqual(r.history.undoLabel, "");
    r = ok(wb.setCells("Sheet1", [{ row: 2, col: 1, input: "2" }, { row: 1, col: 2, input: "=A1*2" }]));
    assertHistory(r.history);
    assert.equal(r.history.depth, 2);
    r = ok(wb.addSheet("Other"));
    assertHistory(r.history);
    assert.equal(r.history.depth, 3);
    r = ok(wb.renameSheet("Other", "Second"));
    assertHistory(r.history);
    assert.equal(r.history.depth, 4);
  });

  check("M1 undo and redo shapes and effect", () => {
    let r = ok(wb.undo()); // rename
    assertHistory(r.history);
    assert.equal(typeof r.label, "string");
    assert.equal(r.select, null);
    assert.deepEqual(r.sheets.map((s) => s.name), ["Sheet1", "Other"]);
    assert.equal(r.history.canRedo, true);
    assert.equal(r.history.depth, 3);
    ok(wb.undo()); // add sheet
    r = ok(wb.undo()); // setCells
    assert.equal(r.sheet, "Sheet1");
    assert.deepEqual(r.select, R(1, 1, 2, 2));
    assert.deepEqual(ok(wb.getRange("Sheet1", 1, 1, 9, 9)).cells.map((c) => [c.row, c.col, c.input]), [[1, 1, "1"]]);
    r = ok(wb.redo());
    assert.equal(r.sheet, "Sheet1");
    assert.deepEqual(r.select, R(1, 1, 2, 2));
    assert.equal(r.history.depth, 2);
    assert.equal(cellAt(ok(wb.getRange("Sheet1", 1, 1, 9, 9)).cells, 1, 2).input, "=A1*2");
    // A new write clears the redo stack.
    r = ok(wb.setCell("Sheet1", 5, 5, "x"));
    assert.equal(r.history.canRedo, false);
    fails(wb.redo(), /nothing to redo/);
    ok(wb.undo());
  });

  check("M1 errors do not change history.depth", () => {
    const depth = ok(wb.history()).history.depth;
    fails(wb.setCell("Nope", 1, 1, "x"));
    fails(wb.setCells("Sheet1", [{ row: 3, col: 3, input: "ok" }, { row: 0, col: 1, input: "bad" }]));
    fails(wb.clearRange("Sheet1", { row1: 1, col1: 1, row2: 1 }));
    fails(wb.clearRange("Sheet1", { row1: 1, col1: 1, row2: 1, col2: 1, extra: 1 }));
    fails(wb.clearRange("Sheet1", "A1"));
    fails(wb.fill("Sheet1", R(1, 1, 1, 1), R(3, 3, 4, 4), true));
    fails(wb.fill("Sheet1", R(1, 1, 1, 1), R(1, 1, 3, 1), "yes"));
    fails(wb.paste("Sheet1", R(0, 1, 1, 1), "x"));
    fails(wb.paste("Sheet1", R(1, 1, 1, 1), 5));
    fails(wb.copy("Nope", R(1, 1, 1, 1), false));
    assert.equal(ok(wb.history()).history.depth, depth);
    // The rejected setCells did not write its first edit.
    assert.equal(cellAt(ok(wb.getRange("Sheet1", 1, 1, 9, 9)).cells, 3, 3), undefined);
  });

  check("M1 fill with a formula and a series", () => {
    // A1=1, A2=2, B1==A1*2
    let r = ok(wb.fill("Sheet1", R(1, 1, 2, 2), R(1, 1, 5, 2), true));
    assertHistory(r.history);
    assert.deepEqual(r.range, R(1, 1, 5, 2));
    assert.equal(typeof r.changed, "number");
    assert.equal(typeof r.skipped, "number");
    const cells = ok(wb.getRange("Sheet1", 1, 1, 5, 2)).cells;
    assert.equal(cellAt(cells, 5, 1).display, "5");
    assert.equal(cellAt(cells, 3, 2).input, "=A3*2");
    assert.equal(cellAt(cells, 5, 2).display, "10");
    r = ok(wb.undo());
    assert.deepEqual(r.select, R(1, 1, 5, 2));
    assert.equal(ok(wb.getRange("Sheet1", 3, 1, 5, 2)).cells.length, 0);
    ok(wb.redo());
  });

  check("M1 copy then paste shifts a formula", () => {
    let r = ok(wb.copy("Sheet1", R(1, 2, 1, 2), false));
    assert.equal(r.text, "2");
    assert.equal(r.cells, 1);
    assert.equal(r.internal, true);
    assertHistory(r.history);
    const depth = r.history.depth;
    r = ok(wb.paste("Sheet1", R(3, 5, 3, 5), r.text));
    assert.equal(r.mode, "internal");
    assert.deepEqual(r.range, R(3, 5, 3, 5));
    assert.equal(r.clipped, false);
    assert.equal(r.history.depth, depth + 1);
    assert.equal(cellAt(ok(wb.getRange("Sheet1", 3, 5, 3, 5)).cells, 3, 5).input, "=D3*2");
  });

  check("M1 copy id decides the paste mode", () => {
    const cp = ok(wb.copy("Sheet1", R(1, 2, 1, 2), false)); // B1 = =A1*2, shown 2
    assert.equal(typeof cp.id, "string");
    assert.ok(cp.id.length >= 16, `id = ${cp.id}`);
    // Equal id: internal, also when the text is different.
    assert.equal(ok(wb.paste("Sheet1", R(20, 5, 20, 5), "other", cp.id)).mode, "internal");
    // The clipboard has no id: the same text from a different program.
    let r = ok(wb.paste("Sheet1", R(21, 5, 21, 5), cp.text, ""));
    assert.equal(r.mode, "text");
    assert.equal(cellAt(ok(wb.getRange("Sheet1", 21, 5, 21, 5)).cells, 21, 5).input, "2");
    // A different id.
    assert.equal(ok(wb.paste("Sheet1", R(22, 5, 22, 5), cp.text, "0123456789abcdef")).mode, "text");
    // No id argument, undefined and null: the text compare.
    assert.equal(ok(wb.paste("Sheet1", R(23, 5, 23, 5), cp.text)).mode, "internal");
    assert.equal(ok(wb.paste("Sheet1", R(23, 5, 23, 5), cp.text, undefined)).mode, "internal");
    assert.equal(ok(wb.paste("Sheet1", R(23, 5, 23, 5), cp.text, null)).mode, "internal");
    fails(wb.paste("Sheet1", R(23, 5, 23, 5), cp.text, 5));
    // Empty text is nothing to paste.
    r = ok(wb.paste("Sheet1", R(20, 5, 20, 5), "", ""));
    assert.equal(r.changed, 0);
    assert.equal(cellAt(ok(wb.getRange("Sheet1", 20, 5, 20, 5)).cells, 20, 5).input, "=D20*2");
    ok(wb.clearRange("Sheet1", R(20, 5, 23, 5)));
  });

  check("M1 copy text has the input of constants", () => {
    ok(wb.setCell("Sheet1", 30, 1, "'007"));
    const cp = ok(wb.copy("Sheet1", R(30, 1, 30, 1), false));
    assert.equal(cp.text, "'007");
    ok(wb.clearRange("Sheet1", R(30, 1, 30, 1)));
  });

  check("M1 paste of other text is a text paste", () => {
    const r = ok(wb.paste("Sheet1", R(7, 1, 7, 1), "a\t1\r\n\"x\ty\"\t=1+1\r\n"));
    assert.equal(r.mode, "text");
    assert.deepEqual(r.range, R(7, 1, 8, 2));
    const cells = ok(wb.getRange("Sheet1", 7, 1, 8, 2)).cells;
    assert.deepEqual(cells.map((c) => [c.input, c.type]), [["a", "string"], ["1", "number"], ["x\ty", "string"], ["=1+1", "number"]]);
  });

  check("M1 cut is one undo step", () => {
    const before = ok(wb.history()).history.depth;
    let r = ok(wb.copy("Sheet1", R(7, 1, 8, 2), true));
    assert.equal(r.history.depth, before + 1);
    assert.equal(ok(wb.getRange("Sheet1", 7, 1, 8, 2)).cells.length, 0);
    r = ok(wb.undo());
    assert.equal(ok(wb.getRange("Sheet1", 7, 1, 8, 2)).cells.length, 4);
  });

  check("M1 clearRange", () => {
    const r = ok(wb.clearRange("Sheet1", R(8, 2, 7, 1))); // corners in any order
    assertHistory(r.history);
    assert.equal(r.changed, 4);
    assert.equal(r.skipped, 0);
    assert.equal(ok(wb.getRange("Sheet1", 7, 1, 8, 2)).cells.length, 0);
    ok(wb.undo());
    assert.equal(ok(wb.getRange("Sheet1", 7, 1, 8, 2)).cells.length, 4);
  });

  check("M1 deleteSheet clears the history", () => {
    ok(wb.addSheet("Temp"));
    const r = ok(wb.deleteSheet("Temp"));
    assert.deepEqual(r.history, { canUndo: false, canRedo: false, undoLabel: "", redoLabel: "", depth: 0 });
  });
}

// ---------------------------------------------------------------------------
// Milestone 2: formula help, layout, navigation.
// ---------------------------------------------------------------------------

function checksM2Formula(wb) {
	check("M2 functions() has more than 400 entries and SUM has arguments", () => {
		const fns = ok(wb.functions()).functions;
		assert.ok(fns.length > 400, `got ${fns.length} functions`);
		const names = fns.map((f) => f.name);
		assert.deepEqual(names, [...names].sort());
		const sum = fns.find((f) => f.name === "SUM");
		assert.deepEqual(sum.args, [
			{ name: "number1", optional: false, repeat: false },
			{ name: "number2", optional: true, repeat: true },
		]);
		assert.equal(sum.signature, "SUM(number1, [number2], ...)");
		assert.equal(typeof sum.summary, "string");
		assert.equal(typeof sum.minArgs, "number");
		assert.equal(typeof sum.variadic, "boolean");
		// Each record has an args array, also with no authored record.
		assert.ok(fns.every((f) => Array.isArray(f.args) && f.signature.startsWith(f.name + "(")));
		assert.ok(fns.some((f) => f.summary === "" && f.signature === f.name + "(...)"));
	});

	check("M2 formulaContext prefix", () => {
		const r = ok(wb.formulaContext("=SU", 3));
		assert.deepEqual(r.prefix, { start: 1, end: 3, text: "SU" });
		assert.equal(r.call, null);
		assert.deepEqual(r.refs, []);
	});

	check("M2 formulaContext call and refs", () => {
		const r = ok(wb.formulaContext("=SUMIF(A1:A3,", 13));
		assert.equal(r.prefix, null);
		assert.deepEqual(r.call, { name: "SUMIF", argIndex: 1, nameStart: 1 });
		assert.deepEqual(r.refs, [
			{ start: 7, end: 12, text: "A1:A3", sheet: "", kind: "range", range: R(1, 1, 3, 1), color: 0 },
		]);
	});

	check("M2 formulaContext offsets are UTF-16 code units", () => {
		const text = '="\u{1F600}"&A1+SU';
		const r = ok(wb.formulaContext(text, text.length));
		assert.equal(text.slice(r.prefix.start, r.prefix.end), "SU");
		assert.equal(text.slice(r.refs[0].start, r.refs[0].end), "A1");
	});

	check("M2 formulaContext with text that is not a formula", () => {
		assert.deepEqual(ok(wb.formulaContext("SUM(A1", 3)), { ok: true, prefix: null, call: null, refs: [] });
		assert.deepEqual(ok(wb.formulaContext("", 0)), { ok: true, prefix: null, call: null, refs: [] });
		ok(wb.formulaContext("=SUM(", 999));
		fails(wb.formulaContext(5, 0));
		fails(wb.formulaContext("=SU"));
		fails(wb.formulaContext("=SU", 1.5));
	});

	check("M2 cycleRef", () => {
		let r = ok(wb.cycleRef("=A1+B2", 3, 3));
		assert.deepEqual(r, { ok: true, text: "=$A$1+B2", selStart: 5, selEnd: 5, changed: true });
		r = ok(wb.cycleRef(r.text, r.selStart, r.selEnd));
		assert.equal(r.text, "=A$1+B2");
		r = ok(wb.cycleRef("=SUM(B2:C3)", 8, 8));
		assert.equal(r.text, "=SUM($B$2:$C$3)");
		r = ok(wb.cycleRef("=A1+B2", 1, 6));
		assert.deepEqual(r, { ok: true, text: "=$A$1+$B$2", selStart: 1, selEnd: 10, changed: true });
		r = ok(wb.cycleRef("=1+2", 2, 2));
		assert.deepEqual(r, { ok: true, text: "=1+2", selStart: 2, selEnd: 2, changed: false });
		fails(wb.cycleRef("=A1", 1));
		fails(wb.cycleRef(1, 1, 1));
	});
}

function checksM2Layout(wb) {
	check("M2 setColWidth then sheetMeta then save and open keeps the width", () => {
		ok(wb.newWorkbook());
		let r = ok(wb.setColWidth("Sheet1", 2, 4, 123));
		assertHistory(r.history);
		assert.equal(r.history.depth, 1);
		assert.equal(r.history.undoLabel, "Column width");
		assert.deepEqual(ok(wb.sheetMeta("Sheet1")).colWidths, { 2: 123, 3: 123, 4: 123 });
		r = ok(wb.setRowHeight("Sheet1", 3, 3, 41));
		assert.equal(r.history.undoLabel, "Row height");
		assert.deepEqual(ok(wb.sheetMeta("Sheet1")).rowHeights, { 3: 41 });

		const saved = ok(wb.save());
		ok(wb.open(saved.data));
		const meta = ok(wb.sheetMeta("Sheet1"));
		assert.deepEqual(meta.colWidths, { 2: 123, 3: 123, 4: 123 });
		assert.deepEqual(meta.rowHeights, { 3: 41 });
	});

	check("M2 setColWidth and setRowHeight: remove, limits, undo", () => {
		ok(wb.setColWidth("Sheet1", 3, 3, 0));
		assert.deepEqual(ok(wb.sheetMeta("Sheet1")).colWidths, { 2: 123, 4: 123 });
		ok(wb.setColWidth("Sheet1", 2, 2, 1));
		ok(wb.setRowHeight("Sheet1", 3, 3, 99999));
		let meta = ok(wb.sheetMeta("Sheet1"));
		assert.equal(meta.colWidths[2], 12);
		assert.equal(meta.rowHeights[3], 1000);
		let r = ok(wb.undo());
		assert.equal(r.label, "Row height");
		assert.deepEqual(r.select, R(3, 1, 3, 16384));
		r = ok(wb.undo());
		assert.deepEqual(r.select, R(1, 2, 1048576, 2));
		meta = ok(wb.sheetMeta("Sheet1"));
		assert.equal(meta.colWidths[2], 123);
		assert.equal(meta.rowHeights[3], 41);
		const depth = ok(wb.history()).history.depth;
		fails(wb.setColWidth("Sheet1", 0, 1, 50));
		fails(wb.setColWidth("Nope", 1, 1, 50));
		fails(wb.setColWidth("Sheet1", 1, 1, "50"));
		fails(wb.setRowHeight("Sheet1", 1, 1048577, 20));
		fails(wb.setRowHeight("Sheet1", 1, 1));
		assert.equal(ok(wb.history()).history.depth, depth);
	});
}

function checksM2Navigation(wb) {
	check("M2 findEdge", () => {
		ok(wb.newWorkbook());
		ok(wb.setCells("Sheet1", [1, 2, 3, 4, 5, 8].map((row) => ({ row, col: 1, input: String(row) }))));
		assert.deepEqual(ok(wb.findEdge("Sheet1", 1, 1, 1, 0)), { ok: true, row: 5, col: 1 });
		assert.deepEqual(ok(wb.findEdge("Sheet1", 5, 1, 1, 0)), { ok: true, row: 8, col: 1 });
		assert.deepEqual(ok(wb.findEdge("Sheet1", 8, 1, 1, 0)), { ok: true, row: 1048576, col: 1 });
		assert.deepEqual(ok(wb.findEdge("Sheet1", 8, 1, -1, 0)), { ok: true, row: 5, col: 1 });
		assert.deepEqual(ok(wb.findEdge("Sheet1", 1, 1, 0, 1)), { ok: true, row: 1, col: 16384 });
		assert.deepEqual(ok(wb.findEdge("Sheet1", 3, 4, 0, -1)), { ok: true, row: 3, col: 1 });
		fails(wb.findEdge("Sheet1", 1, 1, 1, 1));
		fails(wb.findEdge("Sheet1", 1, 1, 0, 0));
		fails(wb.findEdge("Sheet1", 1, 1, 2, 0));
		fails(wb.findEdge("Sheet1", 0, 1, 1, 0));
		fails(wb.findEdge("Nope", 1, 1, 1, 0));
	});

	check("M2 currentRegion", () => {
		ok(wb.newWorkbook());
		ok(wb.setCells("Sheet1", [
			{ row: 2, col: 2, input: "Name" }, { row: 2, col: 3, input: "Qty" },
			{ row: 3, col: 2, input: "a" }, { row: 3, col: 3, input: "5" },
			{ row: 4, col: 2, input: "b" }, { row: 4, col: 3, input: "7" },
		]));
		assert.deepEqual(ok(wb.currentRegion("Sheet1", 3, 3)), { ok: true, range: R(2, 2, 4, 3), header: true });
		assert.deepEqual(ok(wb.currentRegion("Sheet1", 9, 9)), { ok: true, range: R(9, 9, 9, 9), header: false });
		fails(wb.currentRegion("Sheet1", 0, 1));
		fails(wb.currentRegion("Nope", 1, 1));
	});
}

// ---------------------------------------------------------------------------
// Milestone 3: structure and format.
// ---------------------------------------------------------------------------

function checksM3Structure(wb) {
  const inputs = (sheet) =>
    Object.fromEntries(ok(wb.getRange(sheet, 1, 1, 30, 10)).cells.map((c) => [`${c.row},${c.col}`, c.input]));

  check("M3 insertRows changes a formula; undo restores it", () => {
    ok(wb.newWorkbook());
    ok(wb.setCells("Sheet1", [
      ...[1, 2, 3, 4, 5].map((row) => ({ row, col: 1, input: String(row) })),
      { row: 1, col: 2, input: "=SUM(A1:A5)" },
      { row: 1, col: 3, input: "=A5" },
    ]));
    const before = inputs("Sheet1");
    let r = ok(wb.insertRows("Sheet1", 3, 1));
    assertHistory(r.history);
    assert.equal(r.history.undoLabel, "Insert rows");
    assert.equal(r.sheets[0].maxRow, 6);
    let now = inputs("Sheet1");
    assert.equal(now["1,2"], "=SUM(A1:A6)");
    assert.equal(now["1,3"], "=A6");
    assert.equal(now["3,1"], undefined);
    assert.equal(now["4,1"], "3");
    r = ok(wb.undo());
    assert.equal(r.label, "Insert rows");
    assert.deepEqual(r.select, R(3, 1, 3, 16384));
    assert.deepEqual(inputs("Sheet1"), before);
    ok(wb.redo());
    assert.equal(inputs("Sheet1")["1,3"], "=A6");
    ok(wb.undo());
  });

  check("M3 deleteRows and deleteCols give #REF!; undo restores", () => {
    const before = inputs("Sheet1");
    let r = ok(wb.deleteRows("Sheet1", 5, 1));
    assert.equal(r.history.undoLabel, "Delete rows");
    let now = inputs("Sheet1");
    assert.equal(now["1,2"], "=SUM(A1:A4)");
    assert.equal(now["1,3"], "=#REF!");
    assert.equal(cellAt(ok(wb.getRange("Sheet1", 1, 3, 1, 3)).cells, 1, 3).display, "#REF!");
    ok(wb.undo());
    assert.deepEqual(inputs("Sheet1"), before);

    r = ok(wb.deleteCols("Sheet1", 1, 1));
    assert.equal(r.history.undoLabel, "Delete columns");
    assert.equal(r.sheets[0].maxCol, 2);
    now = inputs("Sheet1");
    assert.equal(now["1,1"], "=SUM(#REF!)");
    assert.equal(now["1,2"], "=#REF!");
    r = ok(wb.undo());
    assert.deepEqual(r.select, R(1, 1, 1048576, 1));
    assert.deepEqual(inputs("Sheet1"), before);
  });

  check("M3 insertCols moves formulas and sizes", () => {
    ok(wb.setColWidth("Sheet1", 2, 2, 111));
    const r = ok(wb.insertCols("Sheet1", 2, 2));
    assert.equal(r.history.undoLabel, "Insert columns");
    const now = inputs("Sheet1");
    assert.equal(now["1,4"], "=SUM(A1:A5)");
    assert.equal(now["1,5"], "=A5");
    assert.deepEqual(ok(wb.sheetMeta("Sheet1")).colWidths, { 4: 111 });
    ok(wb.undo());
    assert.deepEqual(ok(wb.sheetMeta("Sheet1")).colWidths, { 2: 111 });
  });

  check("M3 structure errors do not change the history; sheetMeta has warnings", () => {
    const depth = ok(wb.history()).history.depth;
    fails(wb.insertRows("Sheet1", 0, 1), /out of range/);
    fails(wb.insertRows("Sheet1", 1, 0), /count/);
    fails(wb.deleteRows("Sheet1", 1048576, 2), /count/);
    fails(wb.deleteCols("Sheet1", 16385, 1), /out of range/);
    fails(wb.insertCols("Nope", 1, 1), /not found/);
    fails(wb.insertRows("Sheet1", 1));
    fails(wb.insertRows("Sheet1", "1", 1));
    fails(wb.deleteCols("Sheet1", 1.5, 1));
    assert.equal(ok(wb.history()).history.depth, depth);
    assert.deepEqual(ok(wb.sheetMeta("Sheet1")).warnings, []);
  });
}

function checksM3Style(wb) {
  check("M3 setStyle then getRange shows the style and numFmt; save and open keeps both", () => {
    ok(wb.newWorkbook());
    ok(wb.setCells("Sheet1", [{ row: 1, col: 1, input: "1234.5" }, { row: 1, col: 2, input: "0.25" }]));
    let r = ok(wb.setStyle("Sheet1", R(1, 1, 1, 2), { bold: true, fill: "#ffff00", align: "center" }));
    assertHistory(r.history);
    assert.equal(r.history.undoLabel, "Format");
    assert.equal(r.changed, 2);
    assert.equal(r.clamped, false);
    assert.deepEqual(r.range, R(1, 1, 1, 2));
    ok(wb.setStyle("Sheet1", R(1, 1, 1, 1), { numFmt: "currency" }));
    ok(wb.setStyle("Sheet1", R(1, 2, 1, 2), { numFmt: "percent", color: "#FF0000" }));

    const expectCells = () => {
      const cells = ok(wb.getRange("Sheet1", 1, 1, 1, 2)).cells;
      assert.equal(cells[0].display, "$1,234.50");
      assert.deepEqual(cells[0].style, { bold: true, fill: "#FFFF00", align: "center", numFmt: "currency" });
      assert.equal(cells[1].display, "25%");
      assert.deepEqual(cells[1].style, { bold: true, color: "#FF0000", fill: "#FFFF00", align: "center", numFmt: "percent" });
    };
    expectCells();
    const saved = ok(wb.save());
    ok(wb.open(saved.data));
    expectCells();
  });

  check("M3 setStyle: remove with null, clamp, undo", () => {
    let r = ok(wb.setStyle("Sheet1", R(1, 1, 1, 2), { fill: null, color: null, bold: false, align: "general", numFmt: "general" }));
    assert.equal(r.changed, 2);
    let cells = ok(wb.getRange("Sheet1", 1, 1, 1, 2)).cells;
    assert.equal(cells[0].style, undefined);
    assert.equal(cells[0].display, "1234.5");
    r = ok(wb.undo());
    assert.equal(r.label, "Format");
    cells = ok(wb.getRange("Sheet1", 1, 1, 1, 2)).cells;
    assert.equal(cells[1].style.numFmt, "percent");
    // A full column is limited to the used range.
    r = ok(wb.setStyle("Sheet1", R(1, 1, 1048576, 1), { italic: true }));
    assert.equal(r.clamped, true);
    assert.deepEqual(r.range, R(1, 1, 1, 1));
    assert.equal(r.changed, 1);
  });

  check("M3 setStyle errors", () => {
    const depth = ok(wb.history()).history.depth;
    fails(wb.setStyle("Sheet1", R(1, 1, 1, 1), { bold: "yes" }));
    fails(wb.setStyle("Sheet1", R(1, 1, 1, 1), { color: "red" }), /RRGGBB/);
    fails(wb.setStyle("Sheet1", R(1, 1, 1, 1), { numFmt: "custom" }), /numFmt/);
    fails(wb.setStyle("Sheet1", R(1, 1, 1, 1), { align: "top" }), /align/);
    fails(wb.setStyle("Sheet1", R(1, 1, 1, 1), { border: "thin" }));
    fails(wb.setStyle("Sheet1", R(1, 1, 1, 1), "bold"));
    fails(wb.setStyle("Sheet1", R(1, 1, 1, 1)));
    fails(wb.setStyle("Nope", R(1, 1, 1, 1), { bold: true }), /not found/);
    assert.equal(ok(wb.history()).history.depth, depth);
  });
}

// ---------------------------------------------------------------------------
// Milestone 4: freeze, sort, find.
// ---------------------------------------------------------------------------

function checksM4(wb) {
  const col = (c, r1, r2) => {
    const cells = ok(wb.getRange("Sheet1", r1, c, r2, c)).cells;
    const out = [];
    for (let r = r1; r <= r2; r++) out.push((cellAt(cells, r, c) || { display: "" }).display);
    return out;
  };

  check("M4 setFreeze in sheetMeta and after save and open", () => {
    ok(wb.newWorkbook());
    assert.deepEqual(ok(wb.sheetMeta("Sheet1")).freeze, { rows: 0, cols: 0 });
    let r = ok(wb.setFreeze("Sheet1", 2, 1));
    assertHistory(r.history);
    assert.equal(r.history.undoLabel, "Freeze panes");
    assert.deepEqual(ok(wb.sheetMeta("Sheet1")).freeze, { rows: 2, cols: 1 });
    const saved = ok(wb.save());
    ok(wb.open(saved.data));
    assert.deepEqual(ok(wb.sheetMeta("Sheet1")).freeze, { rows: 2, cols: 1 });
    // A structural edit moves the freeze line; undo moves it back.
    ok(wb.insertRows("Sheet1", 1, 1));
    assert.deepEqual(ok(wb.sheetMeta("Sheet1")).freeze, { rows: 3, cols: 1 });
    r = ok(wb.undo());
    assert.deepEqual(ok(wb.sheetMeta("Sheet1")).freeze, { rows: 2, cols: 1 });
    ok(wb.setFreeze("Sheet1", 0, 0));
    r = ok(wb.undo());
    assert.equal(r.label, "Freeze panes");
    assert.equal(r.select, null);
    assert.deepEqual(ok(wb.sheetMeta("Sheet1")).freeze, { rows: 2, cols: 1 });
    fails(wb.setFreeze("Sheet1", -1, 0), /out of range/);
    fails(wb.setFreeze("Sheet1", 1));
    fails(wb.setFreeze("Nope", 1, 1), /not found/);
  });

  check("M4 sortRange", () => {
    ok(wb.newWorkbook());
    ok(wb.setCells("Sheet1", [
      { row: 1, col: 1, input: "Name" }, { row: 1, col: 2, input: "Qty" }, { row: 1, col: 3, input: "Double" },
      { row: 2, col: 1, input: "pear" }, { row: 2, col: 2, input: "30" }, { row: 2, col: 3, input: "=B2*2" },
      { row: 3, col: 1, input: "apple" }, { row: 3, col: 2, input: "10" }, { row: 3, col: 3, input: "=B3*2" },
      { row: 4, col: 1, input: "fig" }, { row: 4, col: 2, input: "20" }, { row: 4, col: 3, input: "=B4*2" },
    ]));
    assert.deepEqual(ok(wb.currentRegion("Sheet1", 3, 2)), { ok: true, range: R(1, 1, 4, 3), header: true });
    let r = ok(wb.sortRange("Sheet1", R(1, 1, 4, 3), [{ col: 2 }], true));
    assertHistory(r.history);
    assert.equal(r.history.undoLabel, "Sort");
    assert.deepEqual(r.range, R(1, 1, 4, 3));
    assert.ok(r.changed > 0);
    assert.deepEqual(col(1, 1, 4), ["Name", "apple", "fig", "pear"]);
    assert.deepEqual(col(3, 1, 4), ["Double", "20", "40", "60"]);
    assert.equal(cellAt(ok(wb.getRange("Sheet1", 2, 3, 2, 3)).cells, 2, 3).input, "=B2*2");
    r = ok(wb.sortRange("Sheet1", R(1, 1, 4, 3), [{ col: 1, desc: true }], true));
    assert.deepEqual(col(1, 1, 4), ["Name", "pear", "fig", "apple"]);
    r = ok(wb.undo());
    assert.equal(r.label, "Sort");
    assert.deepEqual(col(1, 1, 4), ["Name", "apple", "fig", "pear"]);
    const depth = ok(wb.history()).history.depth;
    fails(wb.sortRange("Sheet1", R(1, 1, 4, 3), [], true), /no sort key/);
    fails(wb.sortRange("Sheet1", R(1, 1, 4, 3), [{ col: 9 }], true), /not in the range/);
    fails(wb.sortRange("Sheet1", R(1, 1, 4, 3), { col: 1 }, true), /array/);
    fails(wb.sortRange("Sheet1", R(1, 1, 4, 3), [{ col: 1, order: "asc" }], true));
    fails(wb.sortRange("Sheet1", R(1, 1, 4, 3), [{ col: 1 }]));
    fails(wb.sortRange("Sheet1", R(1, 1, 4, 3), [{ col: 1 }], "yes"));
    assert.equal(ok(wb.history()).history.depth, depth);
  });

  check("M4 find", () => {
    // A2 apple, A3 fig, A4 pear; C2:C4 formulas.
    let r = ok(wb.find("Sheet1", "P"));
    assert.deepEqual(r, { ok: true, found: true, row: 2, col: 1, count: 2, wrapped: false });
    r = ok(wb.find("Sheet1", "p", { from: { row: 2, col: 1 } }));
    assert.deepEqual(r, { ok: true, found: true, row: 4, col: 1, count: 2, wrapped: false });
    r = ok(wb.find("Sheet1", "p", { from: { row: 4, col: 1 } }));
    assert.equal(r.wrapped, true);
    assert.equal(r.row, 2);
    r = ok(wb.find("Sheet1", "p", { from: { row: 2, col: 1 }, backward: true }));
    assert.deepEqual([r.row, r.wrapped], [4, true]);
    assert.equal(ok(wb.find("Sheet1", "P", { matchCase: true })).found, false);
    assert.equal(ok(wb.find("Sheet1", "fig", { wholeCell: true })).row, 3);
    assert.equal(ok(wb.find("Sheet1", "B3*", {})).col, 3);
    assert.equal(ok(wb.find("Sheet1", "40", { lookIn: "display" })).col, 3);
    assert.equal(ok(wb.find("Sheet1", "B3*", { lookIn: "display" })).found, false);
    assert.deepEqual(ok(wb.find("Sheet1", "zebra", null)), { ok: true, found: false, row: 0, col: 0, count: 0, wrapped: false });
    fails(wb.find("Sheet1", ""), /empty/);
    fails(wb.find("Sheet1", "x", { regex: true }));
    fails(wb.find("Sheet1", "x", { lookIn: "formula" }), /lookIn/);
    fails(wb.find("Sheet1", "x", "opts"));
    fails(wb.find("Sheet1", 5));
    fails(wb.find("Nope", "x"), /not found/);
  });

  check("M4 replace and replaceAll", () => {
    const depth = ok(wb.history()).history.depth;
    let r = ok(wb.replace("Sheet1", 2, 1, "apple", "plum"));
    assertHistory(r.history);
    assert.equal(r.replaced, true);
    assert.equal(r.history.undoLabel, "Replace");
    assert.deepEqual(col(1, 2, 2), ["plum"]);
    r = ok(wb.replace("Sheet1", 2, 1, "zebra", "x"));
    assert.equal(r.replaced, false);
    assert.equal(r.history.depth, depth + 1);
    // A replace in a formula keeps the formula.
    r = ok(wb.replace("Sheet1", 2, 3, "*2", "*3", { matchCase: true }));
    assert.equal(cellAt(ok(wb.getRange("Sheet1", 2, 3, 2, 3)).cells, 2, 3).input, "=B2*3");
    assert.deepEqual(col(3, 2, 2), ["30"]);

    r = ok(wb.replaceAll("Sheet1", "0", "5"));
    assert.equal(r.history.undoLabel, "Replace all");
    assert.equal(r.replaced, 3);
    assert.equal(r.skipped, 0);
    assert.deepEqual(col(2, 2, 4), ["15", "25", "35"]);
    r = ok(wb.undo());
    assert.equal(r.label, "Replace all");
    assert.deepEqual(r.select, R(2, 2, 4, 2));
    assert.deepEqual(col(2, 2, 4), ["10", "20", "30"]);
    r = ok(wb.replaceAll("Sheet1", "zebra", "x", { wholeCell: true }));
    assert.equal(r.replaced, 0);
    fails(wb.replace("Sheet1", 0, 1, "a", "b"), /out of range/);
    fails(wb.replace("Sheet1", 1, 1, "", "b"), /empty/);
    fails(wb.replace("Sheet1", 1, 1, "a"));
    fails(wb.replaceAll("Sheet1", "", "b"), /empty/);
    fails(wb.replaceAll("Sheet1", "a", 5));
    fails(wb.replaceAll("Nope", "a", "b"), /not found/);
  });
}

// ---------------------------------------------------------------------------
// Timing (information only, not a gate).
// ---------------------------------------------------------------------------

function timingFormulaContext(wb) {
	const text = '=IF(SUMIFS(Sales!$C:$C,Sales!$A:$A,$A2,Sales!$B:$B,">="&B$1)>0,VLOOKUP($A2,Data!$A$1:$D$100,3,FALSE),SU';
	const n = 5000;
	const start = performance.now();
	for (let i = 0; i < n; i++) wb.formulaContext(text, text.length);
	const ms = performance.now() - start;
	console.log(`info ${n} formulaContext calls (${text.length} characters): ${ms.toFixed(0)} ms (${(ms / n).toFixed(3)} ms each)`);
}

function timing(wb) {
  ok(wb.newWorkbook());
  const n = 2000;
  const start = performance.now();
  for (let i = 1; i <= n; i++) {
    wb.setCell("Sheet1", i, 1, String(i));
    wb.getRange("Sheet1", 1, 1, 40, 20);
  }
  const ms = performance.now() - start;
  console.log(`info ${n} setCell + getRange pairs: ${ms.toFixed(0)} ms (${(ms / n).toFixed(2)} ms each)`);
}

// ---------------------------------------------------------------------------

function run() {
  const wb = globalThis.werkbook;
  assert.equal(typeof wb, "object", "globalThis.werkbook is not set");
  checksS0(wb);
  // A milestone runs when its first function is in the module, so that the
  // script is valid at each step of the stage. A milestone that is partly
  // built fails at its first missing function.
  if (typeof wb.history === "function") checksM1(wb);
  else console.log("skip M1 (history is not in the module)");
  // Milestone 2 is delivered in three groups.
  if (typeof wb.functions === "function") checksM2Formula(wb);
  else console.log("skip M2 formula help (functions is not in the module)");
  if (typeof wb.setColWidth === "function") checksM2Layout(wb);
  else console.log("skip M2 layout (setColWidth is not in the module)");
  if (typeof wb.findEdge === "function") checksM2Navigation(wb);
  else console.log("skip M2 navigation (findEdge is not in the module)");
  // Milestone 3 is delivered in two groups.
  if (typeof wb.insertRows === "function") checksM3Structure(wb);
  else console.log("skip M3 structure (insertRows is not in the module)");
  if (typeof wb.setStyle === "function") checksM3Style(wb);
  else console.log("skip M3 format (setStyle is not in the module)");
  if (typeof wb.setFreeze === "function") checksM4(wb);
  else console.log("skip M4 (setFreeze is not in the module)");
  if (typeof wb.formulaContext === "function") timingFormulaContext(wb);
  timing(wb);
  console.log(`ALL OK (${count} checks)`);
}

let readyCalls = 0;
globalThis.onWerkbookReady = () => {
  readyCalls++;
  check("onWerkbookReady is called one time", () => assert.equal(readyCalls, 1));
  try {
    run();
  } catch (err) {
    console.log("FAIL", err);
    process.exit(1);
  }
  clearTimeout(timer);
  process.exit(0);
};

await import(pathToFileURL(path.join(dir, "wasm_exec.js")).href);
const go = new globalThis.Go();
const { instance } = await WebAssembly.instantiate(fs.readFileSync(path.join(dir, "werkbook.wasm")), go.importObject);
go.run(instance);
