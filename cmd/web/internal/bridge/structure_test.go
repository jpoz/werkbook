package bridge

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"reflect"
	"strings"
	"testing"

	werkbook "github.com/jpoz/werkbook"
)

// newStructureSession builds A1..A5 = 1..5, B1 = SUM(A1:A5), C1 = A5 on
// Sheet1, and a formula on Two that reads Sheet1.
func newStructureSession(t *testing.T) *Session {
	t.Helper()
	s := NewSession()
	if _, err := s.AddSheet("Two"); err != nil {
		t.Fatal(err)
	}
	mustSet(t, s, sheet1, "A1", "1", "A2", "2", "A3", "3", "A4", "4", "A5", "5", "B1", "=SUM(A1:A5)", "C1", "=A5")
	mustSet(t, s, "Two", "A1", "=Sheet1!A5*10")
	return s
}

func TestStructureFunctions(t *testing.T) {
	fullRows := func(r1, r2 int) Range { return Range{r1, 1, r2, werkbook.MaxColumns} }
	fullCols := func(c1, c2 int) Range { return Range{1, c1, werkbook.MaxRows, c2} }

	tests := []struct {
		name       string
		run        func(s *Session) (StructureResult, error)
		label      string
		sel        Range
		wantInputs map[string]string // Sheet1
		wantTwo    string            // input of Two!A1
		wantSize   [2]int            // maxRow, maxCol of Sheet1
	}{
		{
			name:  "insertRows",
			run:   func(s *Session) (StructureResult, error) { return s.InsertRows(sheet1, 3, 2) },
			label: "Insert rows", sel: fullRows(3, 4),
			wantInputs: map[string]string{"A1": "1", "A2": "2", "A5": "3", "A6": "4", "A7": "5", "B1": "=SUM(A1:A7)", "C1": "=A7"},
			wantTwo:    "=Sheet1!A7*10", wantSize: [2]int{7, 3},
		},
		{
			name:  "deleteRows",
			run:   func(s *Session) (StructureResult, error) { return s.DeleteRows(sheet1, 4, 2) },
			label: "Delete rows", sel: fullRows(4, 5),
			wantInputs: map[string]string{"A1": "1", "A2": "2", "A3": "3", "B1": "=SUM(A1:A3)", "C1": "=#REF!"},
			wantTwo:    "=#REF!*10", wantSize: [2]int{3, 3},
		},
		{
			name:  "insertCols",
			run:   func(s *Session) (StructureResult, error) { return s.InsertCols(sheet1, 1, 1) },
			label: "Insert columns", sel: fullCols(1, 1),
			wantInputs: map[string]string{"B1": "1", "B2": "2", "B3": "3", "B4": "4", "B5": "5", "C1": "=SUM(B1:B5)", "D1": "=B5"},
			wantTwo:    "=Sheet1!B5*10", wantSize: [2]int{5, 4},
		},
		{
			name:  "deleteCols",
			run:   func(s *Session) (StructureResult, error) { return s.DeleteCols(sheet1, 1, 2) },
			label: "Delete columns", sel: fullCols(1, 2),
			wantInputs: map[string]string{"A1": "=#REF!"},
			wantTwo:    "=#REF!*10", wantSize: [2]int{1, 1},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := newStructureSession(t)
			before := dump(t, s, sheet1) + dump(t, s, "Two")
			depth := s.History().Depth

			res, err := tt.run(s)
			if err != nil {
				t.Fatal(err)
			}
			if res.History.Depth != depth+1 || res.History.UndoLabel != tt.label || res.History.CanRedo || res.History.Cleared {
				t.Errorf("history = %+v", res.History)
			}
			if len(res.Sheets) != 2 || res.Sheets[0].MaxRow != tt.wantSize[0] || res.Sheets[0].MaxCol != tt.wantSize[1] {
				t.Errorf("sheets = %+v, want Sheet1 with size %v", res.Sheets, tt.wantSize)
			}
			wantInputs(t, s, sheet1, tt.wantInputs)
			if got := inputsOf(t, s, "Two")["A1"]; got != tt.wantTwo {
				t.Errorf("Two!A1 = %q, want %q", got, tt.wantTwo)
			}
			after := dump(t, s, sheet1) + dump(t, s, "Two")

			// Undo restores all, also the formulas that became #REF!.
			undo := mustUndo(t, s)
			if undo.Label != tt.label || undo.Sheet != sheet1 || undo.Select == nil || *undo.Select != tt.sel {
				t.Errorf("undo = label %q sheet %q select %v, want %q %q %v", undo.Label, undo.Sheet, undo.Select, tt.label, sheet1, tt.sel)
			}
			if got := dump(t, s, sheet1) + dump(t, s, "Two"); got != before {
				t.Errorf("state after undo:\n%s\nwant:\n%s", got, before)
			}
			// Redo runs the edit again; a second undo still works.
			redo := mustRedo(t, s)
			if redo.Label != tt.label || *redo.Select != tt.sel {
				t.Errorf("redo = %+v", redo)
			}
			if got := dump(t, s, sheet1) + dump(t, s, "Two"); got != after {
				t.Errorf("state after redo:\n%s\nwant:\n%s", got, after)
			}
			mustUndo(t, s)
			if got := dump(t, s, sheet1) + dump(t, s, "Two"); got != before {
				t.Errorf("state after the second undo is not the start state")
			}
		})
	}
}

// Structural steps and cell steps in one history: each undo is valid only in
// strict sequence, and the history keeps that sequence.
func TestStructureInAHistoryWithCellSteps(t *testing.T) {
	s := newStructureSession(t)
	var states []string
	state := func() string { return dump(t, s, sheet1) + dump(t, s, "Two") }
	step := func(fn func() error) {
		t.Helper()
		states = append(states, state())
		if err := fn(); err != nil {
			t.Fatal(err)
		}
	}
	base := s.History().Depth

	step(func() error { _, err := s.InsertRows(sheet1, 2, 1); return err })
	step(func() error { return s.SetCell(sheet1, 2, 1, "=A1+A3") })
	step(func() error { _, err := s.DeleteCols(sheet1, 2, 1); return err })
	step(func() error { _, err := s.Fill(sheet1, Range{2, 1, 2, 1}, Range{2, 1, 2, 3}, false); return err })
	step(func() error { _, err := s.InsertCols("Two", 1, 2); return err })
	step(func() error { _, err := s.DeleteRows(sheet1, 1, 3); return err })
	step(func() error {
		_, err := s.SetStyle(sheet1, Range{1, 1, 2, 2}, StylePatch{Bold: boolPtr(true)})
		return err
	})
	step(func() error { _, err := s.SetColWidth(sheet1, 1, 2, 90); return err })
	end := state()

	if got := s.History().Depth; got != base+len(states) {
		t.Fatalf("depth = %d, want %d", got, base+len(states))
	}
	for i := len(states) - 1; i >= 0; i-- {
		mustUndo(t, s)
		if got := state(); got != states[i] {
			t.Fatalf("state after undo of step %d:\n%s\nwant:\n%s", i, got, states[i])
		}
	}
	for range states {
		mustRedo(t, s)
	}
	if got := state(); got != end {
		t.Errorf("state after all redos:\n%s\nwant:\n%s", got, end)
	}
	// And back again: the undo data of each redone structural step is new.
	for i := len(states) - 1; i >= 0; i-- {
		mustUndo(t, s)
	}
	if got := state(); got != states[0] {
		t.Errorf("state after the second round of undos is not the start state")
	}
}

func boolPtr(b bool) *bool    { return &b }
func strPtr(s string) *string { return &s }

var _ = strPtr // used by style_test.go

func TestStructureErrorsLeaveWorkbookAndHistoryUnchanged(t *testing.T) {
	s := newStructureSession(t)
	sh := s.file.Sheet(sheet1)
	if err := sh.SetValue("XFD9", "edge"); err != nil {
		t.Fatal(err)
	}
	if err := sh.SetValue("B20", "h"); err != nil {
		t.Fatal(err)
	}
	if err := sh.AddTable(werkbook.Table{Name: "T", Ref: "B20:C22"}); err != nil {
		t.Fatal(err)
	}
	mustUndo(t, s) // one step on the redo stack
	before := dump(t, s, sheet1)
	history := s.History()

	tests := []struct {
		name string
		run  func() (StructureResult, error)
		want string
	}{
		{"unknown sheet", func() (StructureResult, error) { return s.InsertRows("Nope", 1, 1) }, "not found"},
		{"row 0", func() (StructureResult, error) { return s.InsertRows(sheet1, 0, 1) }, "out of range"},
		{"row past the end", func() (StructureResult, error) { return s.DeleteRows(sheet1, werkbook.MaxRows+1, 1) }, "out of range"},
		{"column 0", func() (StructureResult, error) { return s.DeleteCols(sheet1, 0, 1) }, "out of range"},
		{"column past the end", func() (StructureResult, error) { return s.InsertCols(sheet1, werkbook.MaxColumns+1, 1) }, "out of range"},
		{"count 0", func() (StructureResult, error) { return s.InsertRows(sheet1, 1, 0) }, "count 0"},
		{"negative count", func() (StructureResult, error) { return s.DeleteCols(sheet1, 1, -1) }, "count -1"},
		{"count past the end", func() (StructureResult, error) { return s.DeleteRows(sheet1, werkbook.MaxRows, 2) }, "count 2"},
		{"insert pushes a cell off the sheet", func() (StructureResult, error) { return s.InsertCols(sheet1, 1, 1) }, "XFD9 is not empty"},
		{"delete the header row of a table", func() (StructureResult, error) { return s.DeleteRows(sheet1, 20, 1) }, `cannot delete the header row of table "T"`},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, err := tt.run()
			if err == nil || !strings.Contains(err.Error(), tt.want) {
				t.Fatalf("error = %v, want it to contain %q", err, tt.want)
			}
			if got := s.History(); got != history {
				t.Errorf("history = %+v, want %+v", got, history)
			}
			if got := dump(t, s, sheet1); got != before {
				t.Errorf("the workbook changed")
			}
		})
	}
}

func TestStructureAboveTheCostLimitClearsTheHistory(t *testing.T) {
	s := NewSession()
	// 3 batches of 170,000 cells: 510,000 cells in column A..C blocks. The
	// history drops the oldest step, so its total stays in the limit.
	const perBatch = 170_000
	for b := 0; b < 3; b++ {
		edits := make([]Edit, 0, perBatch)
		for i := 0; i < perBatch; i++ {
			edits = append(edits, Edit{Row: i + 1, Col: b + 1, Input: "1"})
		}
		if err := s.SetCells(sheet1, edits); err != nil {
			t.Fatal(err)
		}
	}
	mustSet(t, s, sheet1, "E1", "x")
	if !s.History().CanUndo {
		t.Fatal("test setup: no history")
	}

	// One delete of all rows removes 510,001 cells: above the cost limit.
	res, err := s.DeleteRows(sheet1, 1, perBatch)
	if err != nil {
		t.Fatal(err)
	}
	want := HistoryState{Cleared: true}
	if res.History != want {
		t.Errorf("history = %+v, want %+v", res.History, want)
	}
	if res.Sheets[0].MaxRow != 0 {
		t.Errorf("the delete did not run: MaxRow = %d", res.Sheets[0].MaxRow)
	}
	if _, err := s.Undo(); err == nil {
		t.Errorf("a step above the cost limit can be undone")
	}
	// The cleared mark is for that call only.
	if got := s.History(); got.Cleared {
		t.Errorf("History() reports cleared: %+v", got)
	}
	mustSet(t, s, sheet1, "A1", "next")
	res, err = s.InsertRows(sheet1, 1, 1)
	if err != nil {
		t.Fatal(err)
	}
	if res.History.Cleared || res.History.Depth != 2 {
		t.Errorf("history after a small edit = %+v", res.History)
	}
}

func TestStructureKeepsMergesSizesAndClipboard(t *testing.T) {
	s := newStructureSession(t)
	sh := s.file.Sheet(sheet1)
	if err := sh.MergeCell("E2", "F4"); err != nil {
		t.Fatal(err)
	}
	if _, err := s.SetColWidth(sheet1, 2, 2, 100); err != nil {
		t.Fatal(err)
	}
	if _, err := s.SetRowHeight(sheet1, 5, 5, 40); err != nil {
		t.Fatal(err)
	}
	cp := mustCopy(t, s, sheet1, "B1", false) // =SUM(A1:A5)
	metaBefore := mustMeta(t, s, sheet1)

	if _, err := s.InsertRows(sheet1, 3, 1); err != nil {
		t.Fatal(err)
	}
	if _, err := s.InsertCols(sheet1, 1, 1); err != nil {
		t.Fatal(err)
	}
	meta := mustMeta(t, s, sheet1)
	if !reflect.DeepEqual(meta.Merges, []Merge{{Row1: 2, Col1: 6, Row2: 5, Col2: 7}}) {
		t.Errorf("merges = %+v, want F2:G5", meta.Merges)
	}
	if !reflect.DeepEqual(meta.ColWidths, map[int]int{3: 100}) || !reflect.DeepEqual(meta.RowHeights, map[int]int{6: 40}) {
		t.Errorf("sizes = %v %v, want column C and row 6", meta.ColWidths, meta.RowHeights)
	}
	// The clipboard is a snapshot from before the edits: it pastes the
	// formula text of that time.
	res := mustPasteID(t, s, sheet1, "J1", cp.Text, cp.ID)
	if got := inputsOf(t, s, sheet1)["J1"]; res.Mode != "internal" || got != "=SUM(I1:I5)" {
		t.Errorf("paste after structural edits: mode %q, J1 = %q", res.Mode, got)
	}
	mustUndo(t, s)
	mustUndo(t, s)
	mustUndo(t, s)
	if got := mustMeta(t, s, sheet1); !reflect.DeepEqual(got, metaBefore) {
		t.Errorf("meta after undo = %+v, want %+v", got, metaBefore)
	}
}

// --- warnings -----------------------------------------------------------------

// passthroughWorkbook builds an XLSX with worksheet elements that the
// library keeps as they are: it saves a workbook and adds the elements to
// the sheet XML.
func passthroughWorkbook(t *testing.T, extra string) []byte {
	t.Helper()
	src := werkbook.New()
	if err := src.Sheet(sheet1).SetValue("A1", 1); err != nil {
		t.Fatal(err)
	}
	var buf bytes.Buffer
	if _, err := src.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	patched := false
	for _, file := range zr.File {
		rc, err := file.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		if file.Name == "xl/worksheets/sheet1.xml" {
			text := string(data)
			i := strings.Index(text, "</sheetData>")
			if i < 0 {
				t.Fatalf("sheet XML has no </sheetData>: %s", text)
			}
			i += len("</sheetData>")
			data = []byte(text[:i] + extra + text[i:])
			patched = true
		}
		w, err := zw.Create(file.Name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write(data); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	if !patched {
		t.Fatal("sheet1.xml was not in the file")
	}
	return out.Bytes()
}

func TestSheetMetaWarnings(t *testing.T) {
	s := NewSession()
	meta := mustMeta(t, s, sheet1)
	if meta.Warnings == nil || len(meta.Warnings) != 0 {
		t.Errorf("warnings of a new sheet = %#v, want an empty list that is not nil", meta.Warnings)
	}

	// The elements are in the order that the worksheet schema requires.
	extra := `<conditionalFormatting sqref="A1:A9"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>` +
		`<conditionalFormatting sqref="B1"><cfRule type="cellIs" priority="2" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting>` +
		`<dataValidations count="1"><dataValidation type="whole" sqref="B2"><formula1>1</formula1></dataValidation></dataValidations>` +
		`<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`
	data := passthroughWorkbook(t, extra)
	if _, err := s.Open(data); err != nil {
		t.Fatalf("Open: %v", err)
	}
	meta = mustMeta(t, s, sheet1)
	// Sorted, one text for each kind; page margins hold no cell address.
	want := []string{"conditional formats", "data validation"}
	if !reflect.DeepEqual(meta.Warnings, want) {
		t.Errorf("warnings = %v, want %v", meta.Warnings, want)
	}
	// The JSON has the field.
	encoded, err := json.Marshal(meta)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(encoded), `"warnings":["conditional formats","data validation"]`) {
		t.Errorf("JSON = %s", encoded)
	}

	// A structural edit is permitted; the elements stay in the saved file,
	// with their old addresses (the limit that the warning is for).
	if _, err := s.InsertRows(sheet1, 1, 2); err != nil {
		t.Fatalf("InsertRows: %v", err)
	}
	if got := mustMeta(t, s, sheet1).Warnings; !reflect.DeepEqual(got, want) {
		t.Errorf("warnings after an edit = %v", got)
	}
	saved, err := s.Save()
	if err != nil {
		t.Fatal(err)
	}
	s2 := NewSession()
	if _, err := s2.Open(saved); err != nil {
		t.Fatalf("open of the saved file: %v", err)
	}
	if got := mustMeta(t, s2, sheet1).Warnings; !reflect.DeepEqual(got, want) {
		t.Errorf("warnings after save and open = %v", got)
	}
	if got := inputsOf(t, s2, sheet1)["A3"]; got != "1" {
		t.Errorf("A3 after save and open = %q, want the moved value 1", got)
	}

	// Each mapped element name has a text; the list has no duplicates.
	seen := map[string]bool{}
	for name, text := range passthroughWarnings {
		if text == "" {
			t.Errorf("element %s has an empty warning text", name)
		}
		seen[text] = true
	}
	if len(seen) < 8 {
		t.Errorf("only %d warning texts: %v", len(seen), fmt.Sprint(seen))
	}
}
