package werkbook

import "testing"

func mustSetFormulaInternal(t *testing.T, s *Sheet, cell, f string) {
	t.Helper()
	if err := s.SetFormula(cell, f); err != nil {
		t.Fatalf("SetFormula(%s): %v", cell, err)
	}
}

func mustSetValueInternal(t *testing.T, s *Sheet, cell string, v any) {
	t.Helper()
	if err := s.SetValue(cell, v); err != nil {
		t.Fatalf("SetValue(%s): %v", cell, err)
	}
}

func mustCellInternal(t *testing.T, s *Sheet, cell string) *Cell {
	t.Helper()
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		t.Fatalf("CellNameToCoordinates(%s): %v", cell, err)
	}
	r, ok := s.rows[row]
	if !ok {
		t.Fatalf("missing row %d for %s", row, cell)
	}
	c, ok := r.cells[col]
	if !ok {
		t.Fatalf("missing cell %s", cell)
	}
	return c
}

func TestSpillBlockerDynamicDepsDirtyAnchor(t *testing.T) {
	f := New(FirstSheet("Spill"))
	spill := f.Sheet("Spill")

	mustSetFormulaInternal(t, spill, "B2", `SEQUENCE(3)`)
	f.Recalculate()

	anchor := mustCellInternal(t, spill, "B2")
	if anchor.dirty {
		t.Fatalf("anchor dirty before blocker")
	}

	mustSetValueInternal(t, spill, "B3", "blocker")
	if !anchor.dirty {
		t.Fatalf("anchor not dirtied after blocker inside attempted spill rect")
	}

	f.Recalculate()
	if anchor.dirty {
		t.Fatalf("anchor dirty after blocked recalc")
	}

	mustSetValueInternal(t, spill, "B3", nil)
	if !anchor.dirty {
		t.Fatalf("anchor not dirtied after blocker removal inside attempted spill rect")
	}
}

func TestMaterializedSpillRangeDynamicDepsDirtyConsumer(t *testing.T) {
	f := New(FirstSheet("Spill"))
	spill := f.Sheet("Spill")
	calc, err := f.NewSheet("Calc")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	mustSetFormulaInternal(t, spill, "B2", `SEQUENCE(3)`)
	mustSetFormulaInternal(t, calc, "A1", `SUM(INDIRECT("Spill!B2:B10"))`)
	f.Recalculate()

	consumer := mustCellInternal(t, calc, "A1")
	if consumer.dirty {
		t.Fatalf("consumer dirty before spill-range mutation")
	}
	got, err := calc.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(Calc!A1): %v", err)
	}
	if got.Type != TypeNumber || got.Number != 6 {
		t.Fatalf("Calc!A1 = %#v, want 6", got)
	}

	mustSetValueInternal(t, spill, "B4", "blocker")
	if !consumer.dirty {
		t.Fatalf("consumer not dirtied after mutation in materialized grown spill range")
	}
}
