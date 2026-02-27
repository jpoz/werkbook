package werkbook

import (
	"path/filepath"
	"testing"
)

func TestSetGetColumnWidth(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	if err := s.SetColumnWidth("B", 20.5); err != nil {
		t.Fatal(err)
	}

	w, err := s.GetColumnWidth("B")
	if err != nil {
		t.Fatal(err)
	}
	if w != 20.5 {
		t.Errorf("GetColumnWidth(B) = %g, want %g", w, 20.5)
	}
}

func TestGetColumnWidth_Unset(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	w, err := s.GetColumnWidth("A")
	if err != nil {
		t.Fatal(err)
	}
	if w != 0 {
		t.Errorf("GetColumnWidth(A) = %g, want 0", w)
	}
}

func TestSetColumnWidth_InvalidCol(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	err := s.SetColumnWidth("!!!", 10)
	if err == nil {
		t.Error("expected error for invalid column name")
	}
}

func TestGetColumnWidth_InvalidCol(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_, err := s.GetColumnWidth("!!!")
	if err == nil {
		t.Error("expected error for invalid column name")
	}
}

func TestColumnWidth_Roundtrip(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetValue("A1", "Hello")
	_ = s.SetColumnWidth("A", 15.5)
	_ = s.SetColumnWidth("C", 30)

	path := filepath.Join(t.TempDir(), "colwidth.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	s2 := f2.Sheet("Sheet1")

	w1, _ := s2.GetColumnWidth("A")
	if w1 != 15.5 {
		t.Errorf("A width = %g, want %g", w1, 15.5)
	}
	w2, _ := s2.GetColumnWidth("C")
	if w2 != 30 {
		t.Errorf("C width = %g, want %g", w2, 30.0)
	}
	// Unset column should be 0.
	w3, _ := s2.GetColumnWidth("B")
	if w3 != 0 {
		t.Errorf("B width = %g, want 0", w3)
	}
}
