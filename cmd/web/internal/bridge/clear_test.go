package bridge

import (
	"testing"

	werkbook "github.com/jpoz/werkbook"
)

func rangeOf(t *testing.T, a1 string) Range {
	t.Helper()
	col1, row1, col2, row2, err := werkbook.RangeToCoordinates(a1)
	if err != nil {
		// A single cell.
		col, row, err2 := werkbook.CellNameToCoordinates(a1)
		if err2 != nil {
			t.Fatalf("bad range %q: %v", a1, err)
		}
		return Range{Row1: row, Col1: col, Row2: row, Col2: col}
	}
	return Range{Row1: row1, Col1: col1, Row2: row2, Col2: col2}
}

// inputsOf returns the editor text of each cell in A1:Z60, by reference.
func inputsOf(t *testing.T, s *Session, sheet string) map[string]string {
	t.Helper()
	out := map[string]string{}
	for ref, c := range cellsByRef(t, s, sheet, 1, 1, 60, 26) {
		if c.Type == "empty" {
			continue
		}
		out[ref] = c.Input
	}
	return out
}

func wantInputs(t *testing.T, s *Session, sheet string, want map[string]string) {
	t.Helper()
	got := inputsOf(t, s, sheet)
	for ref, w := range want {
		if got[ref] != w {
			t.Errorf("%s input = %q, want %q", ref, got[ref], w)
		}
	}
	for ref, g := range got {
		if _, ok := want[ref]; !ok {
			t.Errorf("%s = %q, want no cell", ref, g)
		}
	}
}

func TestClearRange(t *testing.T) {
	s := NewSession()
	bold := &werkbook.Style{Font: &werkbook.Font{Bold: true}}
	if err := s.file.Sheet(sheet1).SetStyle("B2", bold); err != nil {
		t.Fatal(err)
	}
	mustSet(t, s, sheet1, "A1", "1", "B1", "text", "A2", "=A1*2", "B2", "styled", "C3", "out", "D1", "=SUM(A1:A2)")
	before := dump(t, s, sheet1)
	depth := s.History().Depth

	// Corners in any order.
	res, err := s.ClearRange(sheet1, Range{Row1: 2, Col1: 2, Row2: 1, Col2: 1})
	if err != nil {
		t.Fatal(err)
	}
	if res.Changed != 4 || res.Skipped != 0 {
		t.Errorf("result = %+v, want changed 4 skipped 0", res)
	}
	if res.History.Depth != depth+1 || res.History.UndoLabel != "Clear" {
		t.Errorf("history = %+v", res.History)
	}
	wantInputs(t, s, sheet1, map[string]string{"C3": "out", "D1": "=SUM(A1:A2)"})
	cells := cellsByRef(t, s, sheet1, 1, 1, 5, 5)
	if got := cells["D1"].Display; got != "0" {
		t.Errorf("D1 = %q, want 0", got)
	}
	// The style stays; the cell is a styled empty cell.
	if got, ok := cells["B2"]; !ok || got.Type != "empty" || got.Style == nil || !got.Style.Bold {
		t.Errorf("B2 after clear = %+v, want a styled empty cell", got)
	}
	// Cells with no style are removed, so the used range is C3.
	if info := s.Sheets()[0]; info.MaxRow != 3 || info.MaxCol != 4 {
		t.Errorf("size = %dx%d, want 3x4", info.MaxRow, info.MaxCol)
	}

	undo := mustUndo(t, s)
	if *undo.Select != (Range{1, 1, 2, 2}) || undo.Label != "Clear" {
		t.Errorf("undo result = %+v %v", undo, undo.Select)
	}
	if got := dump(t, s, sheet1); got != before {
		t.Errorf("state after undo:\n%s\nwant:\n%s", got, before)
	}
	mustRedo(t, s)
	checkUndoRedo(t, s, sheet1, before)
}

func TestClearRangeEdgeCases(t *testing.T) {
	t.Run("empty range makes no step", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "1")
		res, err := s.ClearRange(sheet1, rangeOf(t, "C3:E9"))
		if err != nil {
			t.Fatal(err)
		}
		if res.Changed != 0 || res.History.Depth != 1 {
			t.Errorf("result = %+v", res)
		}
	})

	t.Run("style-only cells are not counted and make no step", func(t *testing.T) {
		s := NewSession()
		if err := s.file.Sheet(sheet1).SetStyle("A1", &werkbook.Style{Fill: &werkbook.Fill{Color: "FF0000"}}); err != nil {
			t.Fatal(err)
		}
		res, err := s.ClearRange(sheet1, rangeOf(t, "A1:B2"))
		if err != nil {
			t.Fatal(err)
		}
		if res.Changed != 0 || res.History.Depth != 0 {
			t.Errorf("result = %+v", res)
		}
		if len(cellsByRef(t, s, sheet1, 1, 1, 2, 2)) != 1 {
			t.Errorf("the styled cell is gone")
		}
	})

	t.Run("full columns and full rows touch only stored cells", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "1", "B5", "2", "C9", "keep")
		res, err := s.ClearRange(sheet1, Range{Row1: 1, Col1: 1, Row2: werkbook.MaxRows, Col2: 2})
		if err != nil {
			t.Fatal(err)
		}
		if res.Changed != 2 {
			t.Errorf("changed = %d, want 2", res.Changed)
		}
		wantInputs(t, s, sheet1, map[string]string{"C9": "keep"})
		res, err = s.ClearRange(sheet1, Range{Row1: 9, Col1: 1, Row2: 9, Col2: werkbook.MaxColumns})
		if err != nil {
			t.Fatal(err)
		}
		if res.Changed != 1 {
			t.Errorf("changed = %d, want 1", res.Changed)
		}
		wantInputs(t, s, sheet1, map[string]string{})
	})

	t.Run("spill followers are skipped; the spill stays", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(4)", "B2", "x")
		before := dump(t, s, sheet1)
		res, err := s.ClearRange(sheet1, rangeOf(t, "A2:B3"))
		if err != nil {
			t.Fatal(err)
		}
		if res.Changed != 1 || res.Skipped != 2 {
			t.Errorf("result = %+v, want changed 1 skipped 2", res)
		}
		cells := cellsByRef(t, s, sheet1, 1, 1, 5, 2)
		if cells["A2"].Display != "2" || !cells["A2"].Spill || cells["A1"].Display != "1" {
			t.Errorf("spill after clear = %+v", cells)
		}
		checkUndoRedo(t, s, sheet1, before)
	})

	t.Run("anchor in the range clears the spill; followers are not skipped", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(4)")
		before := dump(t, s, sheet1)
		res, err := s.ClearRange(sheet1, rangeOf(t, "A1:A2"))
		if err != nil {
			t.Fatal(err)
		}
		if res.Changed != 1 || res.Skipped != 0 {
			t.Errorf("result = %+v, want changed 1 skipped 0", res)
		}
		wantInputs(t, s, sheet1, map[string]string{})
		checkUndoRedo(t, s, sheet1, before)
	})

	t.Run("errors", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "1")
		before := s.History()
		bad := []struct {
			sheet string
			rg    Range
		}{
			{"Nope", Range{1, 1, 1, 1}},
			{sheet1, Range{0, 1, 1, 1}},
			{sheet1, Range{1, 1, 1, 0}},
			{sheet1, Range{1, 1, werkbook.MaxRows + 1, 1}},
			{sheet1, Range{}},
		}
		for _, b := range bad {
			if _, err := s.ClearRange(b.sheet, b.rg); err == nil {
				t.Errorf("ClearRange(%q, %+v) gave no error", b.sheet, b.rg)
			}
		}
		if got := s.History(); got != before {
			t.Errorf("history changed: %+v", got)
		}
		wantInputs(t, s, sheet1, map[string]string{"A1": "1"})
	})
}
