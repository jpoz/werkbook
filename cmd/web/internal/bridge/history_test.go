package bridge

import (
	"bytes"
	"encoding/json"
	"fmt"
	"strings"
	"testing"

	werkbook "github.com/jpoz/werkbook"
	"github.com/jpoz/werkbook/ooxml"
)

// dump gives a text form of the stored and shown state of a sheet range, for
// "undo gives the start state" comparisons.
func dump(t *testing.T, s *Session, sheet string) string {
	t.Helper()
	cells, err := s.GetRange(sheet, 1, 1, 60, 30)
	if err != nil {
		t.Fatalf("GetRange: %v", err)
	}
	var b strings.Builder
	for _, c := range cells {
		ref, _ := werkbook.CoordinatesToCellName(c.Col, c.Row)
		fmt.Fprintf(&b, "%s|%s|%s|%s|%v|%v", ref, c.Input, c.Display, c.Type, c.Formula, c.Spill)
		if c.Style != nil {
			fmt.Fprintf(&b, "|%+v", *c.Style)
		}
		b.WriteString("\n")
	}
	for _, info := range s.Sheets() {
		fmt.Fprintf(&b, "%s %dx%d\n", info.Name, info.MaxRow, info.MaxCol)
	}
	return b.String()
}

func mustUndo(t *testing.T, s *Session) HistoryResult {
	t.Helper()
	res, err := s.Undo()
	if err != nil {
		t.Fatalf("Undo: %v", err)
	}
	return res
}

func mustRedo(t *testing.T, s *Session) HistoryResult {
	t.Helper()
	res, err := s.Redo()
	if err != nil {
		t.Fatalf("Redo: %v", err)
	}
	return res
}

// checkUndoRedo checks that one undo gives the state `before` and one redo
// gives the state that the sheet has now.
func checkUndoRedo(t *testing.T, s *Session, sheet, before string) {
	t.Helper()
	after := dump(t, s, sheet)
	depth := s.History().Depth
	mustUndo(t, s)
	if got := dump(t, s, sheet); got != before {
		t.Errorf("state after undo:\n%s\nwant:\n%s", got, before)
	}
	if got := s.History().Depth; got != depth-1 {
		t.Errorf("depth after undo = %d, want %d (the call must be one step)", got, depth-1)
	}
	mustRedo(t, s)
	if got := dump(t, s, sheet); got != after {
		t.Errorf("state after redo:\n%s\nwant:\n%s", got, after)
	}
	if got := s.History().Depth; got != depth {
		t.Errorf("depth after redo = %d, want %d", got, depth)
	}
}

func TestHistoryEmpty(t *testing.T) {
	s := NewSession()
	if got := s.History(); got != (HistoryState{}) {
		t.Fatalf("History of a new session = %+v, want the zero state", got)
	}
	if _, err := s.Undo(); err == nil || err.Error() != "nothing to undo" {
		t.Errorf("Undo error = %v, want nothing to undo", err)
	}
	if _, err := s.Redo(); err == nil || err.Error() != "nothing to redo" {
		t.Errorf("Redo error = %v, want nothing to redo", err)
	}
}

func TestUndoRedoCellWrites(t *testing.T) {
	bold := &werkbook.Style{Font: &werkbook.Font{Bold: true}}
	tests := []struct {
		name  string
		setup func(t *testing.T, s *Session) // not in the history check
		input string                         // written to B2
	}{
		{name: "value over absent cell", input: "5"},
		{name: "formula over absent cell", input: "=A1+1"},
		{name: "value over value", setup: func(t *testing.T, s *Session) { mustSet(t, s, sheet1, "B2", "old") }, input: "new"},
		{name: "value over formula", setup: func(t *testing.T, s *Session) { mustSet(t, s, sheet1, "B2", "=A1*2") }, input: "7"},
		{name: "formula over value", setup: func(t *testing.T, s *Session) { mustSet(t, s, sheet1, "B2", "7") }, input: "=A1*2"},
		{name: "clear of value", setup: func(t *testing.T, s *Session) { mustSet(t, s, sheet1, "B2", "7") }, input: ""},
		{name: "clear of formula", setup: func(t *testing.T, s *Session) { mustSet(t, s, sheet1, "B2", "=A1*2") }, input: ""},
		{
			name: "value over style-only cell",
			setup: func(t *testing.T, s *Session) {
				if err := s.file.Sheet(sheet1).SetStyle("B2", bold); err != nil {
					t.Fatal(err)
				}
			},
			input: "styled",
		},
		{
			name: "clear of styled value keeps the style",
			setup: func(t *testing.T, s *Session) {
				if err := s.file.Sheet(sheet1).SetStyle("B2", bold); err != nil {
					t.Fatal(err)
				}
				mustSet(t, s, sheet1, "B2", "x")
			},
			input: "",
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := NewSession()
			mustSet(t, s, sheet1, "A1", "10", "C2", "=B2&\"!\"")
			if tt.setup != nil {
				tt.setup(t, s)
			}
			before := dump(t, s, sheet1)
			depth := s.History().Depth

			mustSet(t, s, sheet1, "B2", tt.input)
			h := s.History()
			if h.Depth != depth+1 || !h.CanUndo || h.CanRedo || h.UndoLabel != "Edit cell" || h.RedoLabel != "" {
				t.Fatalf("history after the write = %+v", h)
			}
			if dump(t, s, sheet1) == before {
				t.Fatalf("the write did not change the sheet")
			}

			res := mustUndo(t, s)
			if got := dump(t, s, sheet1); got != before {
				t.Errorf("state after undo:\n%s\nwant:\n%s", got, before)
			}
			if res.Label != "Edit cell" || res.Sheet != sheet1 || res.Select == nil || *res.Select != (Range{2, 2, 2, 2}) {
				t.Errorf("undo result = %+v (select %v)", res, res.Select)
			}
			if res.History.Depth != depth || !res.History.CanRedo || res.History.RedoLabel != "Edit cell" {
				t.Errorf("history after undo = %+v", res.History)
			}
			if len(res.Sheets) != 1 {
				t.Errorf("undo result sheets = %+v", res.Sheets)
			}
			mustRedo(t, s)
			mustUndo(t, s)
			if got := dump(t, s, sheet1); got != before {
				t.Errorf("state after undo, redo, undo:\n%s\nwant:\n%s", got, before)
			}
		})
	}
}

func TestUndoRemovesNewCell(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "1", "H40", "far")
	if info := s.Sheets()[0]; info.MaxRow != 40 || info.MaxCol != 8 {
		t.Fatalf("size = %dx%d, want 40x8", info.MaxRow, info.MaxCol)
	}
	res := mustUndo(t, s)
	if info := res.Sheets[0]; info.MaxRow != 1 || info.MaxCol != 1 {
		t.Errorf("size after undo = %dx%d, want 1x1", info.MaxRow, info.MaxCol)
	}

	// A clear of a cell that never existed makes no cell and no step.
	depth := s.History().Depth
	mustSet(t, s, sheet1, "J50", "")
	if info := s.Sheets()[0]; info.MaxRow != 1 || info.MaxCol != 1 {
		t.Errorf("size after a clear of an absent cell = %dx%d, want 1x1", info.MaxRow, info.MaxCol)
	}
	if got := s.History().Depth; got != depth {
		t.Errorf("depth = %d, want %d", got, depth)
	}
	// A clear of a value removes the stored cell, so the used range shrinks.
	mustSet(t, s, sheet1, "H40", "x", "H40", "")
	if info := s.Sheets()[0]; info.MaxRow != 1 || info.MaxCol != 1 {
		t.Errorf("size after a clear = %dx%d, want 1x1", info.MaxRow, info.MaxCol)
	}
}

func TestNoStepWhenNothingChanges(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "5", "A2", "=A1*2", "A3", "text")
	mustUndo(t, s) // A3 is on the redo stack
	base := s.History()

	writes := []struct{ ref, input string }{
		{"A1", "5"}, {"A1", " 5 "}, {"A2", "=A1*2"}, {"B9", ""},
	}
	for _, w := range writes {
		mustSet(t, s, sheet1, w.ref, w.input)
		if got := s.History(); got != base {
			t.Errorf("history after a write of %q to %s = %+v, want unchanged %+v", w.input, w.ref, got, base)
		}
	}
	if err := s.SetCells(sheet1, []Edit{{Row: 1, Col: 1, Input: "5"}, {Row: 9, Col: 9, Input: ""}}); err != nil {
		t.Fatal(err)
	}
	if got := s.History(); got != base {
		t.Errorf("history after a batch with no change = %+v, want %+v", got, base)
	}
	if _, err := s.RenameSheet(sheet1, sheet1); err != nil {
		t.Fatal(err)
	}
	if got := s.History(); got != base {
		t.Errorf("history after a rename to the same name = %+v, want %+v", got, base)
	}

	// A different type with the same text is a change.
	mustSet(t, s, sheet1, "A1", "'5")
	if got := s.History(); got.Depth != base.Depth+1 || got.CanRedo {
		t.Errorf("history after number -> text = %+v", got)
	}
}

func TestSetCellsIsAtomicAndOneStep(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "1", "B1", "=A1+A2")
	before := dump(t, s, sheet1)
	depth := s.History().Depth

	// The huge formula passes the coordinate checks and fails in the write.
	huge := "=\"" + strings.Repeat("a", 70*1024) + "\""
	err := s.SetCells(sheet1, []Edit{
		{Row: 1, Col: 1, Input: "100"},
		{Row: 2, Col: 1, Input: "200"},
		{Row: 3, Col: 3, Input: "new"},
		{Row: 4, Col: 4, Input: huge},
	})
	if err == nil {
		t.Fatalf("SetCells with a formula above the size limit gave no error")
	}
	if !strings.Contains(err.Error(), "edit 3") {
		t.Errorf("error = %q, want it to name edit 3", err)
	}
	if got := dump(t, s, sheet1); got != before {
		t.Errorf("state after a failed batch:\n%s\nwant:\n%s", got, before)
	}
	if got := s.History().Depth; got != depth {
		t.Errorf("depth after a failed batch = %d, want %d", got, depth)
	}

	// A good batch is one step, also when it writes one cell two times.
	err = s.SetCells(sheet1, []Edit{
		{Row: 1, Col: 1, Input: "100"},
		{Row: 2, Col: 1, Input: "200"},
		{Row: 1, Col: 1, Input: "300"},
		{Row: 5, Col: 2, Input: "x"},
	})
	if err != nil {
		t.Fatal(err)
	}
	if got := cellsByRef(t, s, sheet1, 1, 1, 5, 5)["B1"].Display; got != "500" {
		t.Errorf("B1 = %q, want 500", got)
	}
	h := s.History()
	if h.Depth != depth+1 || h.UndoLabel != "Edit cells" {
		t.Errorf("history = %+v", h)
	}
	res := mustUndo(t, s)
	if res.Select == nil || *res.Select != (Range{1, 1, 5, 2}) {
		t.Errorf("select = %v, want the bounding range A1:B5", res.Select)
	}
	if got := dump(t, s, sheet1); got != before {
		t.Errorf("state after undo of the batch:\n%s\nwant:\n%s", got, before)
	}
}

// A panic in a mutating call (for example in the formula compiler) must
// leave the workbook, the history and the batch state of the file as they
// were, and must come back as an error.
func TestMutateReversesChangesOnPanic(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "1", "B1", "=A1*10")
	if _, err := s.AddSheet("Two"); err != nil {
		t.Fatal(err)
	}
	mustUndo(t, s) // "Add sheet" is on the redo stack
	before := dump(t, s, sheet1)
	history := s.History()
	sh := s.file.Sheet(sheet1)

	err := s.mutate("Test", func(rec *recorder) error {
		rec.setSheet(sheet1)
		value := func(text string) werkbook.CellSnapshot {
			return werkbook.CellSnapshot{}.WithValue(werkbook.Value{Type: werkbook.TypeString, String: text})
		}
		if err := rec.write(sh, 1, 1, value("changed")); err != nil {
			return err
		}
		if err := rec.write(sh, 5, 5, value("new")); err != nil {
			return err
		}
		panic("failure in the middle of a call")
	})
	if err == nil || !strings.Contains(err.Error(), "failure in the middle of a call") {
		t.Fatalf("mutate error = %v, want the panic as an error", err)
	}
	if got := dump(t, s, sheet1); got != before {
		t.Errorf("state after the panic:\n%s\nwant:\n%s", got, before)
	}
	if got := s.History(); got != history {
		t.Errorf("history after the panic = %+v, want %+v", got, history)
	}

	// The batch state is closed: a write invalidates its dependents again.
	// (In an open batch this still gives the correct value, so the test
	// reads the counter through a second file-level batch.)
	mustSet(t, s, sheet1, "A1", "2")
	if got := cellsByRef(t, s, sheet1, 1, 2, 1, 2)["B1"].Display; got != "20" {
		t.Errorf("B1 after a write = %q, want 20", got)
	}
	s.file.EndBatchWrite() // a no-op when no batch is open
	if err := sh.SetValue("A1", 3); err != nil {
		t.Fatal(err)
	}
	if deps, _ := s.file.DirectDependents(sheet1, "A1"); len(deps) != 1 {
		t.Errorf("dependents of A1 = %v", deps)
	}
	if v, _ := sh.GetValue("B1"); v.Number != 30 {
		t.Errorf("B1 = %v, want 30", v.Number)
	}

	// A panic with an error value, and a panic before the first write.
	err = s.mutate("Test", func(rec *recorder) error { panic(fmt.Errorf("wrapped %d", 7)) })
	if err == nil || !strings.Contains(err.Error(), "wrapped 7") {
		t.Errorf("mutate error = %v", err)
	}
	// The session still works.
	mustSet(t, s, sheet1, "C1", "ok")
	mustUndo(t, s)
}

func TestRedoIsClearedByNewWrite(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "1", "A2", "2", "A3", "3")
	mustUndo(t, s)
	mustUndo(t, s)
	if h := s.History(); h.Depth != 1 || !h.CanRedo {
		t.Fatalf("history = %+v", h)
	}
	mustSet(t, s, sheet1, "B1", "x")
	h := s.History()
	if h.CanRedo || h.RedoLabel != "" || h.Depth != 2 {
		t.Errorf("history after a new write = %+v, want no redo and depth 2", h)
	}
	if _, err := s.Redo(); err == nil {
		t.Errorf("Redo after a new write gave no error")
	}
	// A failed call keeps the redo stack.
	mustUndo(t, s)
	if err := s.SetCell("Nope", 1, 1, "x"); err == nil {
		t.Fatal("no error")
	}
	if err := s.SetCell(sheet1, 0, 1, "x"); err == nil {
		t.Fatal("no error")
	}
	if h := s.History(); !h.CanRedo || h.Depth != 1 {
		t.Errorf("history after failed calls = %+v, want redo kept", h)
	}
}

func TestHistoryStepLimit(t *testing.T) {
	s := NewSession()
	for i := 1; i <= maxHistorySteps+25; i++ {
		if err := s.SetCell(sheet1, i, 1, fmt.Sprint(i)); err != nil {
			t.Fatal(err)
		}
	}
	if got := s.History().Depth; got != maxHistorySteps {
		t.Fatalf("depth = %d, want %d", got, maxHistorySteps)
	}
	for i := 0; i < maxHistorySteps; i++ {
		mustUndo(t, s)
	}
	if _, err := s.Undo(); err == nil {
		t.Errorf("Undo above the step limit gave no error")
	}
	// The 25 oldest writes stay.
	cells := cellsByRef(t, s, sheet1, 1, 1, 200, 1)
	if len(cells) != 25 || cells["A25"].Display != "25" {
		t.Errorf("got %d cells after all undos, want the 25 oldest", len(cells))
	}
}

func TestHistoryCostLimitDropsOldest(t *testing.T) {
	s := NewSession()
	// Each batch changes 150,000 cells. Four batches are above the limit of
	// 500,000 units, so the oldest step must go.
	const rows, cols = 1500, 100
	batch := func(value string) {
		t.Helper()
		edits := make([]Edit, 0, rows*cols)
		for r := 1; r <= rows; r++ {
			for c := 1; c <= cols; c++ {
				edits = append(edits, Edit{Row: r, Col: c, Input: value})
			}
		}
		if err := s.SetCells(sheet1, edits); err != nil {
			t.Fatal(err)
		}
	}
	for i := 1; i <= 3; i++ {
		batch(fmt.Sprint(i))
	}
	if got := s.History().Depth; got != 3 {
		t.Fatalf("depth after three batches = %d, want 3", got)
	}
	batch("4")
	if got := s.History().Depth; got != 3 {
		t.Fatalf("depth after four batches = %d, want 3 (the oldest step is dropped)", got)
	}
	mustUndo(t, s)
	mustUndo(t, s)
	mustUndo(t, s)
	if _, err := s.Undo(); err == nil {
		t.Errorf("the dropped step can still be undone")
	}
	if got := cellsByRef(t, s, sheet1, 1, 1, 1, 1)["A1"].Display; got != "1" {
		t.Errorf("A1 after all undos = %q, want 1 (the state after the dropped step)", got)
	}

	// One call above the cell limit fails and changes nothing.
	before := s.History()
	edits := make([]Edit, 0, maxOpCells+1)
	for i := 0; i <= maxOpCells; i++ {
		edits = append(edits, Edit{Row: i/100 + 1, Col: i%100 + 1, Input: "z"})
	}
	if err := s.SetCells(sheet1, edits); err == nil || !strings.Contains(err.Error(), "200000") {
		t.Errorf("SetCells above the cell limit: error = %v", err)
	}
	if got := s.History(); got != before {
		t.Errorf("history after the failed call = %+v, want %+v", got, before)
	}
	if got := cellsByRef(t, s, sheet1, 1, 1, 1, 1)["A1"].Display; got != "1" {
		t.Errorf("A1 after the failed call = %q, want 1", got)
	}
}

func TestHistoryClearedByOpenNewDeleteSheet(t *testing.T) {
	fill := func(t *testing.T) *Session {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "1", "A2", "2")
		if _, err := s.AddSheet("Two"); err != nil {
			t.Fatal(err)
		}
		mustUndo(t, s)
		if h := s.History(); h.Depth != 2 || !h.CanRedo {
			t.Fatalf("setup history = %+v", h)
		}
		return s
	}

	t.Run("new", func(t *testing.T) {
		s := fill(t)
		s.New()
		if got := s.History(); got != (HistoryState{}) {
			t.Errorf("history = %+v", got)
		}
	})
	t.Run("open", func(t *testing.T) {
		s := fill(t)
		data, err := s.Save()
		if err != nil {
			t.Fatal(err)
		}
		if _, err := s.Open(data); err != nil {
			t.Fatal(err)
		}
		if got := s.History(); got != (HistoryState{}) {
			t.Errorf("history = %+v", got)
		}
	})
	t.Run("failed open keeps the history", func(t *testing.T) {
		s := fill(t)
		before := s.History()
		if _, err := s.Open([]byte("bad")); err == nil {
			t.Fatal("no error")
		}
		if got := s.History(); got != before {
			t.Errorf("history = %+v, want %+v", got, before)
		}
	})
	t.Run("delete sheet", func(t *testing.T) {
		s := fill(t)
		if _, err := s.AddSheet("Three"); err != nil {
			t.Fatal(err)
		}
		if _, err := s.DeleteSheet("Three"); err != nil {
			t.Fatal(err)
		}
		if got := s.History(); got != (HistoryState{}) {
			t.Errorf("history = %+v", got)
		}
	})
	t.Run("failed delete keeps the history", func(t *testing.T) {
		s := fill(t)
		before := s.History()
		if _, err := s.DeleteSheet(sheet1); err == nil {
			t.Fatal("delete of the last sheet gave no error")
		}
		if _, err := s.DeleteSheet("Nope"); err == nil {
			t.Fatal("no error")
		}
		if got := s.History(); got != before {
			t.Errorf("history = %+v, want %+v", got, before)
		}
	})
}

func TestUndoAddAndRenameSheet(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "5")
	if _, err := s.AddSheet("Data"); err != nil {
		t.Fatal(err)
	}
	if h := s.History(); h.UndoLabel != "Add sheet" {
		t.Errorf("label = %q", h.UndoLabel)
	}
	mustSet(t, s, "Data", "A1", "=Sheet1!A1+1")
	if _, err := s.RenameSheet(sheet1, "Input"); err != nil {
		t.Fatal(err)
	}
	if h := s.History(); h.UndoLabel != "Rename sheet" || h.Depth != 4 {
		t.Errorf("history = %+v", h)
	}
	if got := cellsByRef(t, s, "Data", 1, 1, 1, 1)["A1"].Input; got != "=Input!A1+1" {
		t.Fatalf("formula after rename = %q", got)
	}

	// Undo the rename: the name and the formula text go back.
	res := mustUndo(t, s)
	if res.Label != "Rename sheet" || res.Sheet != sheet1 || res.Select != nil {
		t.Errorf("undo rename result = %+v", res)
	}
	if res.Sheets[0].Name != sheet1 {
		t.Errorf("sheets = %+v", res.Sheets)
	}
	if got := cellsByRef(t, s, "Data", 1, 1, 1, 1)["A1"]; got.Input != "=Sheet1!A1+1" || got.Display != "6" {
		t.Errorf("formula after undo of the rename = %+v", got)
	}

	// Older steps hold the old name; they still apply.
	res = mustUndo(t, s) // the formula on Data
	if res.Sheet != "Data" || *res.Select != (Range{1, 1, 1, 1}) {
		t.Errorf("undo cell result = %+v", res)
	}
	res = mustUndo(t, s) // add sheet
	if res.Label != "Add sheet" || res.Sheet != sheet1 || res.Select != nil || len(res.Sheets) != 1 {
		t.Errorf("undo add result = %+v", res)
	}
	if _, err := s.GetRange("Data", 1, 1, 1, 1); err == nil {
		t.Errorf("sheet Data exists after undo of its add")
	}

	// Redo all: the same end state.
	res = mustRedo(t, s)
	if res.Sheet != "Data" || len(res.Sheets) != 2 {
		t.Errorf("redo add result = %+v", res)
	}
	mustRedo(t, s)
	res = mustRedo(t, s)
	if res.Label != "Rename sheet" || res.Sheet != "Input" {
		t.Errorf("redo rename result = %+v", res)
	}
	if got := cellsByRef(t, s, "Data", 1, 1, 1, 1)["A1"]; got.Input != "=Input!A1+1" || got.Display != "6" {
		t.Errorf("formula after redo = %+v", got)
	}
	if s.History().CanRedo {
		t.Errorf("redo stack is not empty")
	}

	// Errors make no step.
	depth := s.History().Depth
	if _, err := s.AddSheet("Data"); err == nil {
		t.Error("no error for a name in use")
	}
	if _, err := s.RenameSheet("Input", "Data"); err == nil {
		t.Error("no error for a name in use")
	}
	if got := s.History().Depth; got != depth {
		t.Errorf("depth after failed sheet calls = %d, want %d", got, depth)
	}
}

func TestUndoAcrossSheets(t *testing.T) {
	s := NewSession()
	if _, err := s.AddSheet("Two"); err != nil {
		t.Fatal(err)
	}
	mustSet(t, s, sheet1, "A1", "one", "B1", "skip")
	mustSet(t, s, "Two", "C3", "two")
	mustSet(t, s, sheet1, "B2", "three")

	res := mustUndo(t, s)
	if res.Sheet != sheet1 || *res.Select != (Range{2, 2, 2, 2}) {
		t.Errorf("first undo = %+v %v", res, res.Select)
	}
	res = mustUndo(t, s)
	if res.Sheet != "Two" || *res.Select != (Range{3, 3, 3, 3}) {
		t.Errorf("second undo = %+v %v", res, res.Select)
	}
	if cells := cellsByRef(t, s, "Two", 1, 1, 9, 9); len(cells) != 0 {
		t.Errorf("sheet Two after undo = %+v", cells)
	}
	if got := cellsByRef(t, s, sheet1, 1, 1, 9, 9); len(got) != 2 {
		t.Errorf("Sheet1 after undo on Two = %+v, want A1 and B1", got)
	}
	res = mustRedo(t, s)
	if res.Sheet != "Two" || cellsByRef(t, s, "Two", 1, 1, 9, 9)["C3"].Display != "two" {
		t.Errorf("redo on Two = %+v", res)
	}
}

func TestUndoSpillAnchorAndBlocker(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "=SEQUENCE(3)", "C1", "=SUM(A1:A3)")
	spilled := dump(t, s, sheet1)

	// A write into the spill range blocks it; undo makes the anchor spill again.
	mustSet(t, s, sheet1, "A2", "x")
	if got := cellsByRef(t, s, sheet1, 1, 1, 3, 3)["A1"].Display; got != "#SPILL!" {
		t.Fatalf("anchor = %q, want #SPILL!", got)
	}
	mustUndo(t, s)
	if got := dump(t, s, sheet1); got != spilled {
		t.Errorf("state after undo of the blocker:\n%s\nwant:\n%s", got, spilled)
	}
	mustRedo(t, s)
	if got := cellsByRef(t, s, sheet1, 1, 1, 3, 3)["A1"].Display; got != "#SPILL!" {
		t.Errorf("anchor after redo = %q, want #SPILL!", got)
	}
	mustUndo(t, s)

	// Overwrite the anchor, undo: the spill is back.
	mustSet(t, s, sheet1, "A1", "plain")
	if cells := cellsByRef(t, s, sheet1, 1, 1, 3, 1); len(cells) != 1 {
		t.Fatalf("cells after the anchor was replaced = %+v", cells)
	}
	checkUndoRedo(t, s, sheet1, spilled)
	mustUndo(t, s)
	if got := dump(t, s, sheet1); got != spilled {
		t.Errorf("state after undo of the anchor write:\n%s\nwant:\n%s", got, spilled)
	}

	// Undo of the anchor itself removes the followers.
	mustUndo(t, s) // C1
	mustUndo(t, s) // A1
	if cells := cellsByRef(t, s, sheet1, 1, 1, 9, 9); len(cells) != 0 {
		t.Errorf("cells after undo of the anchor = %+v", cells)
	}
}

func TestUndoKeepsLegacyArrayFormula(t *testing.T) {
	// Build a file with a legacy array formula through the OOXML layer: the
	// public API has no setter for it.
	data := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{{
			Name: sheet1,
			Rows: []ooxml.RowData{
				{Num: 1, Cells: []ooxml.CellData{
					{Ref: "A1", Value: "1"},
					{Ref: "B1", Value: "6", Formula: "SUM(A1:A2*2)", FormulaType: "array", FormulaRef: "B1", IsArrayFormula: true},
				}},
				{Num: 2, Cells: []ooxml.CellData{{Ref: "A2", Value: "2"}}},
			},
		}},
	}
	var buf bytes.Buffer
	if err := ooxml.WriteWorkbook(&buf, data); err != nil {
		t.Fatal(err)
	}
	s := NewSession()
	if _, err := s.Open(buf.Bytes()); err != nil {
		t.Fatal(err)
	}
	before := s.file.Sheet(sheet1).SnapshotCell(2, 1)
	if !before.IsArrayFormula() {
		t.Fatalf("test setup: B1 is not a legacy array formula")
	}
	mustSet(t, s, sheet1, "B1", "0")
	mustUndo(t, s)
	after := s.file.Sheet(sheet1).SnapshotCell(2, 1)
	if !after.Equal(before) || after.IsArrayFormula() != before.IsArrayFormula() {
		t.Errorf("formula cell after undo is not equal to the cell before the write")
	}
	if got := cellsByRef(t, s, sheet1, 1, 2, 1, 2)["B1"].Display; got != "6" {
		t.Errorf("B1 = %q, want 6", got)
	}
}

func TestUndoSurvivesSaveAndStyles(t *testing.T) {
	s := NewSession()
	sh := s.file.Sheet(sheet1)
	bold := &werkbook.Style{Font: &werkbook.Font{Bold: true}, NumFmt: "0.00"}
	if err := sh.SetRangeStyle("A1:A2", bold); err != nil {
		t.Fatal(err)
	}
	mustSet(t, s, sheet1, "A1", "1.5")
	before := dump(t, s, sheet1)
	mustSet(t, s, sheet1, "A1", "=1/3")
	if _, err := s.Save(); err != nil { // a save does not change the history
		t.Fatal(err)
	}
	if got := s.History().Depth; got != 2 {
		t.Fatalf("depth after save = %d, want 2", got)
	}
	mustUndo(t, s)
	if got := dump(t, s, sheet1); got != before {
		t.Errorf("state after undo:\n%s\nwant:\n%s", got, before)
	}
	// The shared style pointer is the same object after the undo.
	if st, _ := sh.GetStyle("A1"); st != bold {
		t.Errorf("style pointer after undo = %p, want %p", st, bold)
	}
}

func TestHistoryJSONShape(t *testing.T) {
	sel := Range{Row1: 1, Col1: 2, Row2: 3, Col2: 4}
	tests := []struct {
		name string
		v    any
		want string
	}{
		{
			name: "empty state omits cleared",
			v:    HistoryState{},
			want: `{"canUndo":false,"canRedo":false,"undoLabel":"","redoLabel":"","depth":0}`,
		},
		{
			name: "state",
			v:    HistoryState{CanUndo: true, CanRedo: true, UndoLabel: "Fill", RedoLabel: "Paste", Depth: 3, Cleared: true},
			want: `{"canUndo":true,"canRedo":true,"undoLabel":"Fill","redoLabel":"Paste","depth":3,"cleared":true}`,
		},
		{
			name: "result with a range",
			v:    HistoryResult{Sheet: "S", Select: &sel, Sheets: []SheetInfo{}, Label: "Fill"},
			want: `{"history":{"canUndo":false,"canRedo":false,"undoLabel":"","redoLabel":"","depth":0},"sheet":"S","select":{"row1":1,"col1":2,"row2":3,"col2":4},"sheets":[],"label":"Fill"}`,
		},
		{
			name: "result with no range has select null",
			v:    HistoryResult{Sheet: "S", Sheets: []SheetInfo{}, Label: "Add sheet"},
			want: `{"history":{"canUndo":false,"canRedo":false,"undoLabel":"","redoLabel":"","depth":0},"sheet":"S","select":null,"sheets":[],"label":"Add sheet"}`,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := json.Marshal(tt.v)
			if err != nil {
				t.Fatal(err)
			}
			if string(got) != tt.want {
				t.Errorf("JSON = %s\nwant   %s", got, tt.want)
			}
		})
	}
}
