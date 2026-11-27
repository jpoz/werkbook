package werkbook

import (
	"fmt"
	"path/filepath"
	"testing"
)

func TestSheetIndexAndRename(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	if err := s.SetValue("A1", 5); err != nil {
		t.Fatal(err)
	}
	if err := s.SetFormula("A2", "A1*2"); err != nil {
		t.Fatal(err)
	}

	if got := f.SheetIndex("Sheet1"); got != 0 {
		t.Fatalf("SheetIndex(Sheet1) = %d, want 0", got)
	}
	if got := f.SheetIndex("Missing"); got != -1 {
		t.Fatalf("SheetIndex(Missing) = %d, want -1", got)
	}

	if err := f.SetSheetName("Sheet1", "Data"); err != nil {
		t.Fatal(err)
	}
	if f.Sheet("Sheet1") != nil {
		t.Fatal("old sheet name should not resolve after rename")
	}
	if got := f.SheetIndex("Data"); got != 0 {
		t.Fatalf("SheetIndex(Data) = %d, want 0", got)
	}

	v, err := f.Sheet("Data").GetValue("A2")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeNumber || v.Number != 10 {
		t.Fatalf("A2 = %#v, want 10", v)
	}
}

func TestSetSheetVisibleRoundTrip(t *testing.T) {
	f := New()
	if _, err := f.NewSheet("Hidden"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetSheetVisible("Hidden", false); err != nil {
		t.Fatal(err)
	}

	path := filepath.Join(t.TempDir(), "hidden.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}

	if f2.Sheet("Hidden") == nil {
		t.Fatal("Hidden sheet missing after round-trip")
	}
	if f2.Sheet("Hidden").Visible() {
		t.Fatal("Hidden sheet should remain hidden after round-trip")
	}
}

func TestRemoveRowShiftsRowsAndMetadata(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	if err := s.SetValue("A1", "row1"); err != nil {
		t.Fatal(err)
	}
	if err := s.SetValue("A2", "row2"); err != nil {
		t.Fatal(err)
	}
	if err := s.SetValue("A3", "row3"); err != nil {
		t.Fatal(err)
	}
	if err := s.SetRowHeight(3, 25); err != nil {
		t.Fatal(err)
	}

	if err := s.RemoveRow(2); err != nil {
		t.Fatal(err)
	}

	v, err := s.GetValue("A2")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeString || v.String != "row3" {
		t.Fatalf("A2 = %#v, want row3", v)
	}

	h, err := s.GetRowHeight(2)
	if err != nil {
		t.Fatal(err)
	}
	if h != 25 {
		t.Fatalf("row 2 height = %g, want 25", h)
	}
}

func TestRemoveRowsFromDropsTailAndKeepsHead(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	for i, v := range []string{"hdr", "r2", "r3", "r4"} {
		if err := s.SetValue(fmt.Sprintf("A%d", i+1), v); err != nil {
			t.Fatal(err)
		}
	}
	if err := s.SetFormula("B1", `COUNTA(A2:A100)`); err != nil {
		t.Fatal(err)
	}
	if err := s.MergeCell("A1", "B1"); err != nil {
		t.Fatal(err)
	}
	if err := s.MergeCell("C1", "C3"); err != nil {
		t.Fatal(err)
	}
	if err := s.MergeCell("D3", "D4"); err != nil {
		t.Fatal(err)
	}

	if err := s.RemoveRowsFrom(2); err != nil {
		t.Fatal(err)
	}

	if v, _ := s.GetValue("A1"); v.String != "hdr" {
		t.Fatalf("A1 = %#v, want hdr", v)
	}
	for _, ref := range []string{"A2", "A3", "A4"} {
		if v, _ := s.GetValue(ref); v.Type != TypeEmpty {
			t.Fatalf("%s = %#v, want empty", ref, v)
		}
	}
	if len(s.rows) != 1 {
		t.Fatalf("rows = %d, want 1", len(s.rows))
	}
	f.Recalculate()
	if v, _ := s.GetValue("B1"); v.Number != 0 {
		t.Fatalf("B1 = %#v, want 0 after truncation", v)
	}
	merges := s.MergeCells()
	if len(merges) != 1 || merges[0].Start != "A1" || merges[0].End != "B1" {
		t.Fatalf("merges = %#v, want only A1:B1", merges)
	}

	// Truncating an already-short sheet is a no-op.
	if err := s.RemoveRowsFrom(2); err != nil {
		t.Fatal(err)
	}
	if err := s.RemoveRowsFrom(0); err == nil {
		t.Fatal("expected error for row 0")
	}
}
