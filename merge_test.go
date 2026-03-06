package werkbook

import (
	"path/filepath"
	"testing"
)

func TestMergeCellsRoundTrip(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	if err := s.SetValue("A1", "title"); err != nil {
		t.Fatal(err)
	}
	if err := s.MergeCell("A1", "C2"); err != nil {
		t.Fatal(err)
	}

	path := filepath.Join(t.TempDir(), "merge.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}

	merges := f2.Sheet("Sheet1").MergeCells()
	if len(merges) != 1 {
		t.Fatalf("got %d merges, want 1", len(merges))
	}
	if merges[0].Start != "A1" || merges[0].End != "C2" {
		t.Fatalf("merge = %+v, want A1:C2", merges[0])
	}
}

func TestRemoveRowAdjustsMergedRanges(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	if err := s.MergeCell("A2", "C3"); err != nil {
		t.Fatal(err)
	}

	if err := s.RemoveRow(1); err != nil {
		t.Fatal(err)
	}

	merges := s.MergeCells()
	if len(merges) != 1 {
		t.Fatalf("got %d merges, want 1", len(merges))
	}
	if merges[0].Start != "A1" || merges[0].End != "C2" {
		t.Fatalf("merge = %+v, want A1:C2", merges[0])
	}
}
