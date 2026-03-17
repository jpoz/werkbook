package werkbook

import "testing"

type dynamicArrayAnchorRange struct {
	fromCol int
	fromRow int
	toCol   int
	toRow   int
}

func assertDynamicArrayConsistency(t *testing.T, s *Sheet) {
	t.Helper()

	anchors := make(map[sheetCoord]dynamicArrayAnchorRange)
	for rowNum, r := range s.rows {
		for col, c := range r.cells {
			if c == nil {
				t.Fatalf("nil cell at %d,%d", col, rowNum)
			}
			if !c.isDynamicArray {
				continue
			}
			if c.formula == "" {
				t.Fatalf("dynamic-array anchor %d,%d has no formula", col, rowNum)
			}
			if c.isSpillFollower() {
				t.Fatalf("dynamic-array anchor %d,%d is also marked as a spill follower", col, rowNum)
			}
			if c.formulaRef == "" {
				t.Fatalf("dynamic-array anchor %d,%d has empty formulaRef", col, rowNum)
			}
			fromCol, fromRow, toCol, toRow, err := RangeToCoordinates(c.formulaRef)
			if err != nil {
				t.Fatalf("dynamic-array anchor %d,%d has invalid formulaRef %q: %v", col, rowNum, c.formulaRef, err)
			}
			if fromCol != col || fromRow != rowNum {
				t.Fatalf("dynamic-array anchor %d,%d formulaRef %q does not start at anchor", col, rowNum, c.formulaRef)
			}
			if toCol < col || toRow < rowNum {
				t.Fatalf("dynamic-array anchor %d,%d formulaRef %q has inverted bounds", col, rowNum, c.formulaRef)
			}
			anchors[sheetCoord{col: col, row: rowNum}] = dynamicArrayAnchorRange{
				fromCol: fromCol,
				fromRow: fromRow,
				toCol:   toCol,
				toRow:   toRow,
			}
		}
	}

	for coord, anchor := range anchors {
		for row := anchor.fromRow; row <= anchor.toRow; row++ {
			for col := anchor.fromCol; col <= anchor.toCol; col++ {
				if col == coord.col && row == coord.row {
					continue
				}
				follower := s.cellAt(col, row)
				if follower == nil {
					t.Fatalf("anchor %d,%d formulaRef %q is missing follower at %d,%d", coord.col, coord.row, s.cellAt(coord.col, coord.row).formulaRef, col, row)
				}
				if follower.spillParentCol != coord.col || follower.spillParentRow != coord.row {
					t.Fatalf("follower %d,%d points to %d,%d, want %d,%d", col, row, follower.spillParentCol, follower.spillParentRow, coord.col, coord.row)
				}
				if follower.formula != "" {
					t.Fatalf("spill follower %d,%d unexpectedly has formula %q", col, row, follower.formula)
				}
				if follower.isDynamicArray {
					t.Fatalf("spill follower %d,%d is marked as dynamic-array anchor", col, row)
				}
				if follower.isArrayFormula {
					t.Fatalf("spill follower %d,%d is marked as array formula", col, row)
				}
			}
		}
	}

	for rowNum, r := range s.rows {
		for col, c := range r.cells {
			if c == nil || !c.isSpillFollower() {
				continue
			}
			if c.formula != "" {
				t.Fatalf("spill follower %d,%d unexpectedly has formula %q", col, rowNum, c.formula)
			}
			anchor, ok := anchors[sheetCoord{col: c.spillParentCol, row: c.spillParentRow}]
			if !ok {
				t.Fatalf("spill follower %d,%d points to missing anchor %d,%d", col, rowNum, c.spillParentCol, c.spillParentRow)
			}
			if col == c.spillParentCol && rowNum == c.spillParentRow {
				t.Fatalf("spill follower %d,%d points to itself", col, rowNum)
			}
			if col < anchor.fromCol || col > anchor.toCol || rowNum < anchor.fromRow || rowNum > anchor.toRow {
				t.Fatalf("spill follower %d,%d lies outside anchor %d,%d formulaRef", col, rowNum, c.spillParentCol, c.spillParentRow)
			}
		}
	}
}

func TestDynamicArrayShrinkClearsStaleFollowers(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	if err := s.SetValue("A1", "alpha"); err != nil {
		t.Fatal(err)
	}
	if err := s.SetValue("A2", "beta"); err != nil {
		t.Fatal(err)
	}
	if err := s.SetValue("A3", "gamma"); err != nil {
		t.Fatal(err)
	}
	if err := s.SetFormula("C1", `FILTER(A1:A3,A1:A3<>"")`); err != nil {
		t.Fatal(err)
	}

	for cell, want := range map[string]string{"C1": "alpha", "C2": "beta", "C3": "gamma"} {
		v, err := s.GetValue(cell)
		if err != nil {
			t.Fatalf("GetValue(%s): %v", cell, err)
		}
		if v.Type != TypeString || v.String != want {
			t.Fatalf("%s = %#v, want %q", cell, v, want)
		}
	}
	assertDynamicArrayConsistency(t, s)

	if err := s.SetValue("A3", nil); err != nil {
		t.Fatal(err)
	}

	v, err := s.GetValue("C1")
	if err != nil {
		t.Fatalf("GetValue(C1): %v", err)
	}
	if v.Type != TypeString || v.String != "alpha" {
		t.Fatalf("C1 after shrink = %#v, want alpha", v)
	}

	v, err = s.GetValue("C2")
	if err != nil {
		t.Fatalf("GetValue(C2): %v", err)
	}
	if v.Type != TypeString || v.String != "beta" {
		t.Fatalf("C2 after shrink = %#v, want beta", v)
	}

	v, err = s.GetValue("C3")
	if err != nil {
		t.Fatalf("GetValue(C3): %v", err)
	}
	if v.Type != TypeEmpty {
		t.Fatalf("C3 after shrink = %#v, want empty", v)
	}
	if c := s.cellAt(3, 3); c != nil {
		t.Fatalf("stale spill follower at C3 still materialized: %#v", c)
	}
	assertDynamicArrayConsistency(t, s)
}

func TestDynamicArraySpillCollisionAndRecovery(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	if err := s.SetValue("D2", 99); err != nil {
		t.Fatal(err)
	}
	if err := s.SetFormula("D1", "SEQUENCE(3)"); err != nil {
		t.Fatal(err)
	}

	v, err := s.GetValue("D1")
	if err != nil {
		t.Fatalf("GetValue(D1): %v", err)
	}
	if v.Type != TypeError || v.String != "#SPILL!" {
		t.Fatalf("D1 collision value = %#v, want #SPILL!", v)
	}
	v, err = s.GetValue("D2")
	if err != nil {
		t.Fatalf("GetValue(D2): %v", err)
	}
	if v.Type != TypeNumber || v.Number != 99 {
		t.Fatalf("D2 collision value = %#v, want 99", v)
	}
	assertDynamicArrayConsistency(t, s)

	if err := s.SetValue("D2", nil); err != nil {
		t.Fatal(err)
	}

	for cell, want := range map[string]float64{"D1": 1, "D2": 2, "D3": 3} {
		v, err := s.GetValue(cell)
		if err != nil {
			t.Fatalf("GetValue(%s): %v", cell, err)
		}
		if v.Type != TypeNumber || v.Number != want {
			t.Fatalf("%s after recovery = %#v, want %g", cell, v, want)
		}
	}
	assertDynamicArrayConsistency(t, s)
}
