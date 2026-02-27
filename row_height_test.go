package werkbook

import (
	"path/filepath"
	"testing"
)

func TestSetGetRowHeight(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	if err := s.SetRowHeight(3, 25.5); err != nil {
		t.Fatal(err)
	}

	h, err := s.GetRowHeight(3)
	if err != nil {
		t.Fatal(err)
	}
	if h != 25.5 {
		t.Errorf("GetRowHeight(3) = %g, want %g", h, 25.5)
	}
}

func TestGetRowHeight_Unset(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	h, err := s.GetRowHeight(1)
	if err != nil {
		t.Fatal(err)
	}
	if h != 0 {
		t.Errorf("GetRowHeight(1) = %g, want 0", h)
	}
}

func TestSetRowHeight_InvalidRow(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	if err := s.SetRowHeight(0, 10); err == nil {
		t.Error("expected error for row 0")
	}
	if err := s.SetRowHeight(MaxRows+1, 10); err == nil {
		t.Error("expected error for row exceeding max")
	}
}

func TestGetRowHeight_InvalidRow(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_, err := s.GetRowHeight(0)
	if err == nil {
		t.Error("expected error for row 0")
	}
}

func TestRowHeight_Roundtrip(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetValue("A1", "Hello")
	_ = s.SetRowHeight(1, 30)
	_ = s.SetRowHeight(3, 50.5)

	path := filepath.Join(t.TempDir(), "rowheight.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	s2 := f2.Sheet("Sheet1")

	h1, _ := s2.GetRowHeight(1)
	if h1 != 30 {
		t.Errorf("row 1 height = %g, want %g", h1, 30.0)
	}
	h3, _ := s2.GetRowHeight(3)
	if h3 != 50.5 {
		t.Errorf("row 3 height = %g, want %g", h3, 50.5)
	}
	// Unset row should be 0.
	h2, _ := s2.GetRowHeight(2)
	if h2 != 0 {
		t.Errorf("row 2 height = %g, want 0", h2)
	}
}

func TestRowHeight_HeightOnlyRow_Roundtrip(t *testing.T) {
	// A row with only a height (no cells) should survive roundtrip.
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetValue("A1", "Data") // need at least one cell for sheet to be non-empty
	_ = s.SetRowHeight(5, 42)

	path := filepath.Join(t.TempDir(), "heightonly.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	s2 := f2.Sheet("Sheet1")
	h, _ := s2.GetRowHeight(5)
	if h != 42 {
		t.Errorf("row 5 height = %g, want %g", h, 42.0)
	}
}

func TestRowHeightAccessor(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetValue("A1", "Hello")
	_ = s.SetRowHeight(1, 25)

	for row := range s.Rows() {
		if row.Num() == 1 && row.Height() != 25 {
			t.Errorf("Row.Height() = %g, want %g", row.Height(), 25.0)
		}
	}
}
