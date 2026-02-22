package werkbook_test

import (
	"path/filepath"
	"testing"
	"time"

	"github.com/jpoz/werkbook"
)

func TestRowIteration(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Sparse data: rows 1, 5, 10
	s.SetValue("A1", "first")
	s.SetValue("B5", 42)
	s.SetValue("C10", true)

	var rowNums []int
	for r := range s.Rows() {
		rowNums = append(rowNums, r.Num())
	}
	if len(rowNums) != 3 {
		t.Fatalf("expected 3 rows, got %d: %v", len(rowNums), rowNums)
	}
	if rowNums[0] != 1 || rowNums[1] != 5 || rowNums[2] != 10 {
		t.Errorf("expected [1, 5, 10], got %v", rowNums)
	}
}

func TestRowCells(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", "a")
	s.SetValue("C1", "c")

	for r := range s.Rows() {
		cells := r.Cells()
		if len(cells) != 2 {
			t.Fatalf("expected 2 cells, got %d", len(cells))
		}
		if cells[0].Col() != 1 || cells[0].Value().Raw() != "a" {
			t.Errorf("cell 0: col=%d val=%v", cells[0].Col(), cells[0].Value().Raw())
		}
		if cells[1].Col() != 3 || cells[1].Value().Raw() != "c" {
			t.Errorf("cell 1: col=%d val=%v", cells[1].Col(), cells[1].Value().Raw())
		}
	}
}

func TestMaxRowMaxCol(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	if s.MaxRow() != 0 || s.MaxCol() != 0 {
		t.Error("empty sheet should have MaxRow=0, MaxCol=0")
	}

	s.SetValue("C5", "data")
	if s.MaxRow() != 5 {
		t.Errorf("MaxRow = %d, want 5", s.MaxRow())
	}
	if s.MaxCol() != 3 {
		t.Errorf("MaxCol = %d, want 3", s.MaxCol())
	}
}

func TestSparseRoundTrip(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set cells at extreme positions.
	s.SetValue("A1", "start")
	s.SetValue("XFD1048576", "end")

	dir := t.TempDir()
	path := filepath.Join(dir, "sparse.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	s2 := f2.Sheet("Sheet1")
	v, _ := s2.GetValue("A1")
	if v.Raw() != "start" {
		t.Errorf("A1 = %v, want start", v.Raw())
	}
	v, _ = s2.GetValue("XFD1048576")
	if v.Raw() != "end" {
		t.Errorf("XFD1048576 = %v, want end", v.Raw())
	}
}

func TestDateValue(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	date := time.Date(2024, 6, 15, 0, 0, 0, 0, time.UTC)
	if err := s.SetValue("A1", date); err != nil {
		t.Fatalf("SetValue: %v", err)
	}

	v, _ := s.GetValue("A1")
	if v.Type != werkbook.TypeNumber {
		t.Fatalf("expected TypeNumber, got %v", v.Type)
	}
	// June 15, 2024 should be serial 45458
	if v.Number < 45000 || v.Number > 46000 {
		t.Errorf("unexpected serial date: %f", v.Number)
	}
}
