package werkbook

import (
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
