//go:build js && wasm

// Command wasm is the WebAssembly entry point of the in-browser editor. It
// exposes one bridge.Session to JavaScript as globalThis.werkbook. All
// workbook logic is in cmd/web/internal/bridge; this file only converts
// arguments and results. The contract is in cmd/web/README.md.
package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"math"
	"syscall/js"

	"github.com/jpoz/werkbook/cmd/web/internal/bridge"
)

// handler implements one API function. It returns the result fields (a
// struct or map that encoding/json can encode, or nil for no fields).
type handler func(args []js.Value) (any, error)

// bytesResult marks a result whose single field is binary data. JSON cannot
// carry it, so the wrapper copies it to a Uint8Array.
type bytesResult struct {
	field string
	data  []byte
}

// withHistory marks the result of a mutating function whose bridge method
// does not return the history state. The wrapper adds the `history` field.
type withHistory struct {
	fields any
}

type sheetsResult struct {
	Sheets []bridge.SheetInfo `json:"sheets"`
}

// session is the one workbook session of the module.
var session = bridge.NewSession()

func main() {
	api := map[string]handler{
		"version":     version,
		"newWorkbook": newWorkbook,
		"open":        open,
		"sheets":      sheets,
		"getRange":    getRange,
		"sheetMeta":   sheetMeta,
		"setCell":     setCell,
		"setCells":    setCells,
		"addSheet":    addSheet,
		"renameSheet": renameSheet,
		"deleteSheet": deleteSheet,
		"save":        save,

		// Milestone 1.
		"history":    history,
		"undo":       undo,
		"redo":       redo,
		"clearRange": clearRange,
		"fill":       fill,
		"copy":       copyCells,
		"paste":      paste,

		// Milestone 2.
		"functions":      functions,
		"formulaContext": formulaContext,
		"cycleRef":       cycleRef,
		"setColWidth":    setColWidth,
		"setRowHeight":   setRowHeight,
		"findEdge":       findEdge,
		"currentRegion":  currentRegion,

		// Milestone 3.
		"insertRows": insertRows,
		"deleteRows": deleteRows,
		"insertCols": insertCols,
		"deleteCols": deleteCols,
		"setStyle":   setStyle,

		// Milestone 4.
		"setFreeze":  setFreeze,
		"sortRange":  sortRange,
		"find":       find,
		"replace":    replace,
		"replaceAll": replaceAll,
	}

	obj := js.Global().Get("Object").New()
	for name, h := range api {
		obj.Set(name, wrap(name, h))
	}
	js.Global().Set("werkbook", obj)

	if ready := js.Global().Get("onWerkbookReady"); ready.Type() == js.TypeFunction {
		ready.Invoke()
	}

	// Block forever: the exported functions are valid only while main runs.
	select {}
}

// ---------------------------------------------------------------------------
// Handlers
// ---------------------------------------------------------------------------

func version(args []js.Value) (any, error) {
	return map[string]string{"version": bridge.Version}, nil
}

// sheetsWithHistory is the result of the functions that return the sheet
// list and change the history.
func sheetsWithHistory(infos []bridge.SheetInfo, err error) (any, error) {
	if err != nil {
		return nil, err
	}
	return withHistory{fields: sheetsResult{Sheets: infos}}, nil
}

func newWorkbook(args []js.Value) (any, error) {
	return sheetsWithHistory(session.New(), nil)
}

func open(args []js.Value) (any, error) {
	data, err := argBytes(args, 0, "bytes")
	if err != nil {
		return nil, err
	}
	return sheetsWithHistory(session.Open(data))
}

func sheets(args []js.Value) (any, error) {
	return sheetsResult{Sheets: session.Sheets()}, nil
}

func getRange(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	var n [4]int
	for i, name := range []string{"row1", "col1", "row2", "col2"} {
		if n[i], err = argInt(args, i+1, name); err != nil {
			return nil, err
		}
	}
	cells, err := session.GetRange(sheet, n[0], n[1], n[2], n[3])
	if err != nil {
		return nil, err
	}
	return map[string][]bridge.Cell{"cells": cells}, nil
}

func sheetMeta(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	return session.SheetMeta(sheet)
}

func setCell(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	row, err := argInt(args, 1, "row")
	if err != nil {
		return nil, err
	}
	col, err := argInt(args, 2, "col")
	if err != nil {
		return nil, err
	}
	input, err := argString(args, 3, "input")
	if err != nil {
		return nil, err
	}
	if err := session.SetCell(sheet, row, col, input); err != nil {
		return nil, err
	}
	return withHistory{}, nil
}

func setCells(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	edits, err := argEdits(args, 1, "edits")
	if err != nil {
		return nil, err
	}
	if err := session.SetCells(sheet, edits); err != nil {
		return nil, err
	}
	return withHistory{}, nil
}

func addSheet(args []js.Value) (any, error) {
	name, err := argString(args, 0, "name")
	if err != nil {
		return nil, err
	}
	return sheetsWithHistory(session.AddSheet(name))
}

func renameSheet(args []js.Value) (any, error) {
	oldName, err := argString(args, 0, "oldName")
	if err != nil {
		return nil, err
	}
	newName, err := argString(args, 1, "newName")
	if err != nil {
		return nil, err
	}
	return sheetsWithHistory(session.RenameSheet(oldName, newName))
}

func deleteSheet(args []js.Value) (any, error) {
	name, err := argString(args, 0, "name")
	if err != nil {
		return nil, err
	}
	return sheetsWithHistory(session.DeleteSheet(name))
}

func save(args []js.Value) (any, error) {
	data, err := session.Save()
	if err != nil {
		return nil, err
	}
	return bytesResult{field: "data", data: data}, nil
}

func history(args []js.Value) (any, error) {
	return withHistory{}, nil
}

func undo(args []js.Value) (any, error) {
	return session.Undo()
}

func redo(args []js.Value) (any, error) {
	return session.Redo()
}

func clearRange(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	var rg bridge.Range
	if err := argJSON(args, 1, "range", &rg); err != nil {
		return nil, err
	}
	return session.ClearRange(sheet, rg)
}

func fill(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	var src, dst bridge.Range
	if err := argJSON(args, 1, "src", &src); err != nil {
		return nil, err
	}
	if err := argJSON(args, 2, "dst", &dst); err != nil {
		return nil, err
	}
	series, err := argBool(args, 3, "series")
	if err != nil {
		return nil, err
	}
	return session.Fill(sheet, src, dst, series)
}

func copyCells(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	var rg bridge.Range
	if err := argJSON(args, 1, "range", &rg); err != nil {
		return nil, err
	}
	cut, err := argBool(args, 2, "cut")
	if err != nil {
		return nil, err
	}
	return session.Copy(sheet, rg, cut)
}

func paste(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	var dst bridge.Range
	if err := argJSON(args, 1, "dst", &dst); err != nil {
		return nil, err
	}
	text, err := argString(args, 2, "text")
	if err != nil {
		return nil, err
	}
	// The id is optional. Absent (undefined or null) means that the caller
	// cannot read the id from the system clipboard; "" means that the
	// clipboard has no id.
	var id *string
	if len(args) > 3 && !args[3].IsUndefined() && !args[3].IsNull() {
		v, err := argString(args, 3, "id")
		if err != nil {
			return nil, err
		}
		id = &v
	}
	return session.Paste(sheet, dst, text, id)
}

func functions(args []js.Value) (any, error) {
	return map[string][]bridge.FunctionInfo{"functions": session.Functions()}, nil
}

func formulaContext(args []js.Value) (any, error) {
	text, err := argString(args, 0, "text")
	if err != nil {
		return nil, err
	}
	caret, err := argInt(args, 1, "caret")
	if err != nil {
		return nil, err
	}
	return session.FormulaContext(text, caret), nil
}

func cycleRef(args []js.Value) (any, error) {
	text, err := argString(args, 0, "text")
	if err != nil {
		return nil, err
	}
	selStart, err := argInt(args, 1, "selStart")
	if err != nil {
		return nil, err
	}
	selEnd, err := argInt(args, 2, "selEnd")
	if err != nil {
		return nil, err
	}
	return session.CycleRef(text, selStart, selEnd), nil
}

// argInts reads the integer arguments from index first, one for each name.
func argInts(args []js.Value, first int, names ...string) ([]int, error) {
	out := make([]int, len(names))
	for i, name := range names {
		n, err := argInt(args, first+i, name)
		if err != nil {
			return nil, err
		}
		out[i] = n
	}
	return out, nil
}

func setColWidth(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	n, err := argInts(args, 1, "col1", "col2", "px")
	if err != nil {
		return nil, err
	}
	return session.SetColWidth(sheet, n[0], n[1], n[2])
}

func setRowHeight(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	n, err := argInts(args, 1, "row1", "row2", "px")
	if err != nil {
		return nil, err
	}
	return session.SetRowHeight(sheet, n[0], n[1], n[2])
}

func findEdge(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	n, err := argInts(args, 1, "row", "col", "dRow", "dCol")
	if err != nil {
		return nil, err
	}
	return session.FindEdge(sheet, n[0], n[1], n[2], n[3])
}

func currentRegion(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	n, err := argInts(args, 1, "row", "col")
	if err != nil {
		return nil, err
	}
	return session.CurrentRegion(sheet, n[0], n[1])
}

// structural wraps one of the four insert and delete methods.
func structural(first string, op func(sheet string, at, count int) (bridge.StructureResult, error)) handler {
	return func(args []js.Value) (any, error) {
		sheet, err := argString(args, 0, "sheet")
		if err != nil {
			return nil, err
		}
		n, err := argInts(args, 1, first, "count")
		if err != nil {
			return nil, err
		}
		return op(sheet, n[0], n[1])
	}
}

func insertRows(args []js.Value) (any, error) { return structural("row", session.InsertRows)(args) }
func deleteRows(args []js.Value) (any, error) { return structural("row", session.DeleteRows)(args) }
func insertCols(args []js.Value) (any, error) { return structural("col", session.InsertCols)(args) }
func deleteCols(args []js.Value) (any, error) { return structural("col", session.DeleteCols)(args) }

func setStyle(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	var rg bridge.Range
	if err := argJSON(args, 1, "range", &rg); err != nil {
		return nil, err
	}
	var patch bridge.StylePatch
	if err := argJSON(args, 2, "patch", &patch); err != nil {
		return nil, err
	}
	return session.SetStyle(sheet, rg, patch)
}

func setFreeze(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	n, err := argInts(args, 1, "rows", "cols")
	if err != nil {
		return nil, err
	}
	return session.SetFreeze(sheet, n[0], n[1])
}

func sortRange(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	var rg bridge.Range
	if err := argJSON(args, 1, "range", &rg); err != nil {
		return nil, err
	}
	if _, err := arg(args, 2, "keys"); err != nil {
		return nil, err
	}
	if !js.Global().Get("Array").Call("isArray", args[2]).Bool() {
		return nil, fmt.Errorf("argument %q must be an array", "keys")
	}
	var keys []bridge.SortKey
	if err := argJSON(args, 2, "keys", &keys); err != nil {
		return nil, err
	}
	header, err := argBool(args, 3, "header")
	if err != nil {
		return nil, err
	}
	return session.SortRange(sheet, rg, keys, header)
}

// argFindOptions reads the optional FindOptions argument at index i.
func argFindOptions(args []js.Value, i int) (bridge.FindOptions, error) {
	var opts bridge.FindOptions
	if i >= len(args) || args[i].IsUndefined() || args[i].IsNull() {
		return opts, nil
	}
	if js.Global().Get("Array").Call("isArray", args[i]).Bool() {
		return opts, fmt.Errorf("argument %q must be an object", "opts")
	}
	err := argJSON(args, i, "opts", &opts)
	return opts, err
}

func find(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	query, err := argString(args, 1, "query")
	if err != nil {
		return nil, err
	}
	opts, err := argFindOptions(args, 2)
	if err != nil {
		return nil, err
	}
	return session.Find(sheet, query, opts)
}

func replace(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	n, err := argInts(args, 1, "row", "col")
	if err != nil {
		return nil, err
	}
	query, err := argString(args, 3, "query")
	if err != nil {
		return nil, err
	}
	with, err := argString(args, 4, "replacement")
	if err != nil {
		return nil, err
	}
	opts, err := argFindOptions(args, 5)
	if err != nil {
		return nil, err
	}
	return session.Replace(sheet, n[0], n[1], query, with, opts)
}

func replaceAll(args []js.Value) (any, error) {
	sheet, err := argString(args, 0, "sheet")
	if err != nil {
		return nil, err
	}
	query, err := argString(args, 1, "query")
	if err != nil {
		return nil, err
	}
	with, err := argString(args, 2, "replacement")
	if err != nil {
		return nil, err
	}
	opts, err := argFindOptions(args, 3)
	if err != nil {
		return nil, err
	}
	return session.ReplaceAll(sheet, query, with, opts)
}

// ---------------------------------------------------------------------------
// Results
// ---------------------------------------------------------------------------

// wrap turns a handler into a JavaScript function that always returns
// { ok: true, ...fields } or { ok: false, error }. It never throws; a panic
// in the library becomes an error result.
func wrap(name string, h handler) js.Func {
	return js.FuncOf(func(_ js.Value, args []js.Value) (result any) {
		defer func() {
			if r := recover(); r != nil {
				result = failure(fmt.Errorf("%s: internal error: %v", name, r))
			}
		}()
		fields, err := h(args)
		if err != nil {
			return failure(err)
		}
		return success(fields)
	})
}

func failure(err error) js.Value {
	obj := js.Global().Get("Object").New()
	obj.Set("ok", false)
	obj.Set("error", err.Error())
	return obj
}

// success converts the result fields to a JavaScript object. The fields go
// through JSON: one JSON.parse call is much cheaper than a syscall/js call
// for each property of a large getRange result.
func success(fields any) js.Value {
	addHistory := false
	if wh, ok := fields.(withHistory); ok {
		addHistory = true
		fields = wh.fields
	}

	var obj js.Value
	switch f := fields.(type) {
	case nil:
		obj = js.Global().Get("Object").New()
	case bytesResult:
		obj = js.Global().Get("Object").New()
		arr := js.Global().Get("Uint8Array").New(len(f.data))
		js.CopyBytesToJS(arr, f.data)
		obj.Set(f.field, arr)
	default:
		parsed, err := toJS(fields)
		if err != nil {
			return failure(err)
		}
		obj = parsed
	}
	if addHistory {
		h, err := toJS(session.History())
		if err != nil {
			return failure(err)
		}
		obj.Set("history", h)
	}
	obj.Set("ok", true)
	return obj
}

func toJS(v any) (js.Value, error) {
	encoded, err := json.Marshal(v)
	if err != nil {
		return js.Value{}, fmt.Errorf("encode result: %w", err)
	}
	return js.Global().Get("JSON").Call("parse", string(encoded)), nil
}

// ---------------------------------------------------------------------------
// Arguments
// ---------------------------------------------------------------------------

func arg(args []js.Value, i int, name string) (js.Value, error) {
	if i >= len(args) || args[i].IsUndefined() || args[i].IsNull() {
		return js.Value{}, fmt.Errorf("argument %q is missing", name)
	}
	return args[i], nil
}

func argString(args []js.Value, i int, name string) (string, error) {
	v, err := arg(args, i, name)
	if err != nil {
		return "", err
	}
	if v.Type() != js.TypeString {
		return "", fmt.Errorf("argument %q must be a string", name)
	}
	return v.String(), nil
}

func argBool(args []js.Value, i int, name string) (bool, error) {
	v, err := arg(args, i, name)
	if err != nil {
		return false, err
	}
	if v.Type() != js.TypeBoolean {
		return false, fmt.Errorf("argument %q must be a boolean", name)
	}
	return v.Bool(), nil
}

func argInt(args []js.Value, i int, name string) (int, error) {
	v, err := arg(args, i, name)
	if err != nil {
		return 0, err
	}
	return toInt(v, fmt.Sprintf("argument %q", name))
}

func toInt(v js.Value, what string) (int, error) {
	if v.Type() != js.TypeNumber {
		return 0, fmt.Errorf("%s must be an integer", what)
	}
	f := v.Float()
	if f != math.Trunc(f) || math.IsInf(f, 0) || math.Abs(f) > math.MaxInt32 {
		return 0, fmt.Errorf("%s must be an integer", what)
	}
	return int(f), nil
}

// argJSON decodes an object argument (Range, StylePatch, ...) into dst. The
// value goes through JSON.stringify and a strict decoder: an unknown field
// and a wrong type are errors. A missing field keeps its zero value; the
// bridge rejects a zero row or column.
func argJSON(args []js.Value, i int, name string, dst any) error {
	v, err := arg(args, i, name)
	if err != nil {
		return err
	}
	if v.Type() != js.TypeObject {
		return fmt.Errorf("argument %q must be an object", name)
	}
	text := js.Global().Get("JSON").Call("stringify", v)
	if text.Type() != js.TypeString {
		return fmt.Errorf("argument %q cannot be encoded", name)
	}
	dec := json.NewDecoder(bytes.NewReader([]byte(text.String())))
	dec.DisallowUnknownFields()
	if err := dec.Decode(dst); err != nil {
		return fmt.Errorf("argument %q: %v", name, err)
	}
	return nil
}

func argBytes(args []js.Value, i int, name string) ([]byte, error) {
	v, err := arg(args, i, name)
	if err != nil {
		return nil, err
	}
	// Accept an ArrayBuffer as well: it is what File.arrayBuffer() gives.
	if v.InstanceOf(js.Global().Get("ArrayBuffer")) {
		v = js.Global().Get("Uint8Array").New(v)
	}
	if !v.InstanceOf(js.Global().Get("Uint8Array")) {
		return nil, fmt.Errorf("argument %q must be a Uint8Array", name)
	}
	data := make([]byte, v.Get("byteLength").Int())
	js.CopyBytesToGo(data, v)
	return data, nil
}

func argEdits(args []js.Value, i int, name string) ([]bridge.Edit, error) {
	v, err := arg(args, i, name)
	if err != nil {
		return nil, err
	}
	if !js.Global().Get("Array").Call("isArray", v).Bool() {
		return nil, fmt.Errorf("argument %q must be an array", name)
	}
	edits := make([]bridge.Edit, v.Length())
	for j := range edits {
		item := v.Index(j)
		if item.Type() != js.TypeObject {
			return nil, fmt.Errorf("%s[%d] must be an object {row, col, input}", name, j)
		}
		row, err := toInt(item.Get("row"), fmt.Sprintf("%s[%d].row", name, j))
		if err != nil {
			return nil, err
		}
		col, err := toInt(item.Get("col"), fmt.Sprintf("%s[%d].col", name, j))
		if err != nil {
			return nil, err
		}
		input := item.Get("input")
		if input.Type() != js.TypeString {
			return nil, fmt.Errorf("%s[%d].input must be a string", name, j)
		}
		edits[j] = bridge.Edit{Row: row, Col: col, Input: input.String()}
	}
	return edits, nil
}
