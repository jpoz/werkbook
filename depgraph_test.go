package werkbook

import (
	"testing"
)

// TestDepGraphSingleDep verifies a single A1→B1 dependency.
func TestDepGraphSingleDep(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 10)
	s.SetFormula("B1", "A1*2")

	v, err := s.GetValue("B1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Number != 20 {
		t.Fatalf("expected 20, got %v", v.Number)
	}

	// Change A1 — B1 should update.
	s.SetValue("A1", 50)
	v, _ = s.GetValue("B1")
	if v.Number != 100 {
		t.Fatalf("expected 100, got %v", v.Number)
	}
}

// TestDepGraphChain verifies A1→B1→C1 chain propagation.
func TestDepGraphChain(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 5)
	s.SetFormula("B1", "A1+1")
	s.SetFormula("C1", "B1*3")

	v, _ := s.GetValue("C1")
	if v.Number != 18 {
		t.Fatalf("expected 18, got %v", v.Number)
	}

	s.SetValue("A1", 10)
	v, _ = s.GetValue("C1")
	if v.Number != 33 {
		t.Fatalf("expected 33, got %v", v.Number)
	}
}

// TestDepGraphFanOut verifies multiple formulas depending on the same cell.
func TestDepGraphFanOut(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 2)
	s.SetFormula("B1", "A1+1")
	s.SetFormula("C1", "A1*10")
	s.SetFormula("D1", "A1+100")

	s.SetValue("A1", 5)

	v, _ := s.GetValue("B1")
	if v.Number != 6 {
		t.Errorf("B1: expected 6, got %v", v.Number)
	}
	v, _ = s.GetValue("C1")
	if v.Number != 50 {
		t.Errorf("C1: expected 50, got %v", v.Number)
	}
	v, _ = s.GetValue("D1")
	if v.Number != 105 {
		t.Errorf("D1: expected 105, got %v", v.Number)
	}
}

// TestDepGraphRangeDep verifies that SUM(A1:A3) recalculates when a cell in the range changes.
func TestDepGraphRangeDep(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 1)
	s.SetValue("A2", 2)
	s.SetValue("A3", 3)
	s.SetFormula("B1", "SUM(A1:A3)")

	v, _ := s.GetValue("B1")
	if v.Number != 6 {
		t.Fatalf("expected 6, got %v", v.Number)
	}

	s.SetValue("A2", 20)
	v, _ = s.GetValue("B1")
	if v.Number != 24 {
		t.Fatalf("expected 24, got %v", v.Number)
	}
}

// TestDepGraphCrossSheet verifies cross-sheet formula dependency.
func TestDepGraphCrossSheet(t *testing.T) {
	f := New()
	s1 := f.Sheet("Sheet1")
	s2, _ := f.NewSheet("Sheet2")

	s1.SetValue("A1", 7)
	s2.SetFormula("A1", "Sheet1!A1+3")

	v, _ := s2.GetValue("A1")
	if v.Number != 10 {
		t.Fatalf("expected 10, got %v", v.Number)
	}

	s1.SetValue("A1", 20)
	v, _ = s2.GetValue("A1")
	if v.Number != 23 {
		t.Fatalf("expected 23, got %v", v.Number)
	}
}

// TestDepGraphSetFormulaReplacesDeps verifies re-setting a formula updates deps.
func TestDepGraphSetFormulaReplacesDeps(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 10)
	s.SetValue("A2", 20)
	s.SetFormula("B1", "A1")

	v, _ := s.GetValue("B1")
	if v.Number != 10 {
		t.Fatalf("expected 10, got %v", v.Number)
	}

	// Change formula to depend on A2 instead of A1.
	s.SetFormula("B1", "A2")
	v, _ = s.GetValue("B1")
	if v.Number != 20 {
		t.Fatalf("expected 20, got %v", v.Number)
	}

	// Changing A1 should NOT affect B1 anymore.
	s.SetValue("A1", 999)
	v, _ = s.GetValue("B1")
	if v.Number != 20 {
		t.Fatalf("expected 20, got %v", v.Number)
	}

	// Changing A2 should affect B1.
	s.SetValue("A2", 42)
	v, _ = s.GetValue("B1")
	if v.Number != 42 {
		t.Fatalf("expected 42, got %v", v.Number)
	}
}

// TestDepGraphSetValueClearsFormulaDeps verifies SetValue removes formula deps.
func TestDepGraphSetValueClearsFormulaDeps(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 10)
	s.SetFormula("B1", "A1*2")

	v, _ := s.GetValue("B1")
	if v.Number != 20 {
		t.Fatalf("expected 20, got %v", v.Number)
	}

	// Overwrite B1 with a plain value — formula dep should be removed.
	s.SetValue("B1", 99)
	v, _ = s.GetValue("B1")
	if v.Number != 99 {
		t.Fatalf("expected 99, got %v", v.Number)
	}

	// Changing A1 should NOT affect B1.
	s.SetValue("A1", 1000)
	v, _ = s.GetValue("B1")
	if v.Number != 99 {
		t.Fatalf("expected 99, got %v", v.Number)
	}
}

// TestDepGraphRecalculate verifies the Recalculate method evaluates all dirty cells.
func TestDepGraphRecalculate(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 5)
	s.SetFormula("B1", "A1+1")
	s.SetFormula("C1", "B1*2")

	// Force initial evaluation.
	v, _ := s.GetValue("C1")
	if v.Number != 12 {
		t.Fatalf("expected 12, got %v", v.Number)
	}

	// Change A1 but don't read any formula cells.
	s.SetValue("A1", 10)

	// Recalculate should force evaluation.
	f.Recalculate()

	// Now read without triggering lazy eval — values should already be fresh.
	// We verify by checking the cell directly.
	r := s.rows[1]
	b1 := r.cells[2] // B1 = col 2
	if b1.value.Number != 11 {
		t.Fatalf("B1: expected 11, got %v", b1.value.Number)
	}
	c1 := r.cells[3] // C1 = col 3
	if c1.value.Number != 22 {
		t.Fatalf("C1: expected 22, got %v", c1.value.Number)
	}
}

// TestDepGraphFileLoadAndModify verifies dep graph works after loading from file data.
func TestDepGraphFileLoadAndModify(t *testing.T) {
	// Build a file, save its data, reload, then modify.
	f := New()
	s := f.Sheet("Sheet1")
	s.SetValue("A1", 100)
	s.SetFormula("B1", "A1+50")

	// Force evaluation so cached value is set.
	v, _ := s.GetValue("B1")
	if v.Number != 150 {
		t.Fatalf("expected 150, got %v", v.Number)
	}

	// Simulate round-trip through ooxml data.
	data := f.buildWorkbookData()
	f2, err := fileFromData(data)
	if err != nil {
		t.Fatalf("fileFromData error: %v", err)
	}
	s2 := f2.Sheet("Sheet1")

	// Cached value should be trusted.
	v, _ = s2.GetValue("B1")
	if v.Number != 150 {
		t.Fatalf("expected 150 from cache, got %v", v.Number)
	}

	// Now modify A1 and verify B1 recalculates.
	s2.SetValue("A1", 200)
	v, _ = s2.GetValue("B1")
	if v.Number != 250 {
		t.Fatalf("expected 250, got %v", v.Number)
	}
}
