package werkbook

import (
	"bytes"
	"reflect"
	"testing"
)

func TestSetColumnWidthZeroRemovesTheWidth(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	for _, err := range []error{
		s.SetColumnWidth("B", 20),
		s.SetColumnWidth("D", 12.5),
		s.SetColumnWidth("XFD", 9),
	} {
		if err != nil {
			t.Fatal(err)
		}
	}
	want := map[int]float64{2: 20, 4: 12.5, MaxColumns: 9}
	if got := s.ColumnWidths(); !reflect.DeepEqual(got, want) {
		t.Fatalf("ColumnWidths = %v, want %v", got, want)
	}

	// The result is a copy.
	s.ColumnWidths()[2] = 99
	if w, _ := s.GetColumnWidth("B"); w != 20 {
		t.Errorf("a change of the returned map changed the sheet: width = %v", w)
	}

	// Zero and negative widths remove the entry.
	for _, step := range []struct {
		col   string
		width float64
	}{{"B", 0}, {"XFD", -1}, {"Z", 0}} {
		if err := s.SetColumnWidth(step.col, step.width); err != nil {
			t.Fatal(err)
		}
	}
	if got := s.ColumnWidths(); !reflect.DeepEqual(got, map[int]float64{4: 12.5}) {
		t.Errorf("ColumnWidths after removal = %v, want only column D", got)
	}
	if w, _ := s.GetColumnWidth("B"); w != 0 {
		t.Errorf("GetColumnWidth of a removed width = %v, want 0", w)
	}
	if err := s.SetColumnWidth("not a column!", 0); err == nil {
		t.Errorf("SetColumnWidth with a bad column gave no error")
	}

	// Save and open: no <col> for the removed columns.
	var buf bytes.Buffer
	if _, err := f.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	f2, err := OpenReader(&buf)
	if err != nil {
		t.Fatal(err)
	}
	if got := f2.Sheet("Sheet1").ColumnWidths(); !reflect.DeepEqual(got, map[int]float64{4: 12.5}) {
		t.Errorf("ColumnWidths after save and open = %v, want only column D", got)
	}

	// A sheet with no widths gives an empty map that is not nil.
	empty := New().Sheet("Sheet1")
	if got := empty.ColumnWidths(); got == nil || len(got) != 0 {
		t.Errorf("ColumnWidths of a new sheet = %#v", got)
	}
}

func TestSetRowHeightZeroRemovesTheHeight(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	if err := s.SetValue("A1", "x"); err != nil {
		t.Fatal(err)
	}

	// A height on an empty row makes the used range larger; its removal
	// must make it small again.
	if err := s.SetRowHeight(40, 30); err != nil {
		t.Fatal(err)
	}
	if s.MaxRow() != 40 {
		t.Fatalf("MaxRow with a height on row 40 = %d", s.MaxRow())
	}
	if err := s.SetRowHeight(40, 0); err != nil {
		t.Fatal(err)
	}
	if s.MaxRow() != 1 {
		t.Errorf("MaxRow after the height was removed = %d, want 1", s.MaxRow())
	}
	if h, _ := s.GetRowHeight(40); h != 0 {
		t.Errorf("height = %v, want 0", h)
	}

	// A removal on a row that does not exist makes no row.
	if err := s.SetRowHeight(500, 0); err != nil {
		t.Fatal(err)
	}
	if err := s.SetRowHeight(600, -5); err != nil {
		t.Fatal(err)
	}
	if s.MaxRow() != 1 {
		t.Errorf("MaxRow after removals on absent rows = %d, want 1", s.MaxRow())
	}

	// A row with cells keeps its cells.
	if err := s.SetRowHeight(1, 25); err != nil {
		t.Fatal(err)
	}
	if err := s.SetRowHeight(1, 0); err != nil {
		t.Fatal(err)
	}
	if v, _ := s.GetValue("A1"); v.String != "x" {
		t.Errorf("A1 after its row height was removed = %#v", v)
	}
	if h, _ := s.GetRowHeight(1); h != 0 {
		t.Errorf("height of row 1 = %v, want 0", h)
	}

	// A hidden row with no cells stays: the hidden flag is its own state.
	if err := s.SetRowHeight(7, 18); err != nil {
		t.Fatal(err)
	}
	s.rows[7].hidden = true
	if err := s.SetRowHeight(7, 0); err != nil {
		t.Fatal(err)
	}
	if r, ok := s.rows[7]; !ok || !r.hidden || r.height != 0 {
		t.Errorf("hidden row after its height was removed = %+v, %v", r, ok)
	}
	s.rows[7].hidden = false
	delete(s.rows, 7)

	// Out of range is still an error.
	if err := s.SetRowHeight(0, 0); err == nil {
		t.Errorf("SetRowHeight(0, 0) gave no error")
	}
	if err := s.SetRowHeight(MaxRows+1, 0); err == nil {
		t.Errorf("SetRowHeight past the last row gave no error")
	}

	// Save and open: a removed height is not in the file.
	if err := s.SetRowHeight(3, 44); err != nil {
		t.Fatal(err)
	}
	if err := s.SetRowHeight(9, 44); err != nil {
		t.Fatal(err)
	}
	if err := s.SetRowHeight(9, 0); err != nil {
		t.Fatal(err)
	}
	var buf bytes.Buffer
	if _, err := f.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	f2, err := OpenReader(&buf)
	if err != nil {
		t.Fatal(err)
	}
	s2 := f2.Sheet("Sheet1")
	if h, _ := s2.GetRowHeight(3); h != 44 {
		t.Errorf("height of row 3 after save and open = %v, want 44", h)
	}
	if h, _ := s2.GetRowHeight(9); h != 0 {
		t.Errorf("height of row 9 after save and open = %v, want 0", h)
	}
	if s2.MaxRow() != 3 {
		t.Errorf("MaxRow after save and open = %d, want 3", s2.MaxRow())
	}
}
