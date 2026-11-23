package werkbook

import (
	"fmt"
	"testing"

	"github.com/jpoz/werkbook/formula"
)

// A range can be typed with its corners in any order. Excel reads A5:A1 as
// A1:A5; the engine read only the first row before.
func TestReversedRangesReadTheFullRange(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	out, err := f.NewSheet("Out")
	if err != nil {
		t.Fatal(err)
	}
	// A1:C5 holds 1..15, row by row.
	n := 0
	for r := 1; r <= 5; r++ {
		for c := 1; c <= 3; c++ {
			n++
			ref, _ := CoordinatesToCellName(c, r)
			if err := s.SetValue(ref, n); err != nil {
				t.Fatal(err)
			}
		}
	}
	tests := []struct {
		formula string
		want    float64
	}{
		{"SUM(Sheet1!A1:A5)", 35},
		{"SUM(Sheet1!A5:A1)", 35},  // rows reversed
		{"SUM(Sheet1!C1:A1)", 6},   // columns reversed
		{"SUM(Sheet1!C5:A1)", 120}, // both reversed
		{"SUM(Sheet1!C1:A5)", 120}, // columns reversed only
		{"SUM(Sheet1!A5:C1)", 120}, // rows reversed only
		{"SUM(Sheet1!$A$5:$A$1)", 35},
		{"SUM(Sheet1!A$5:$A1)", 35},
		{"SUM(Sheet1!2:1)", 21}, // row-only
		{"SUM(Sheet1!$2:$1)", 21},
		{"SUM(Sheet1!C:A)", 120}, // column-only
		{"SUM(Sheet1!$C:$B)", 85},
		{"ROWS(Sheet1!A5:A1)", 5},
		{"COLUMNS(Sheet1!C1:A1)", 3},
		{"COUNT(Sheet1!B4:A2)", 6},
		{"INDEX(Sheet1!C5:A1,1,1)", 1}, // the first cell is the top-left cell
		{"INDEX(Sheet1!C5:A1,5,3)", 15},
		{"SUM(Sheet1!A5:A1)+SUM(Sheet1!A1:A5)", 70},
		{"SUM(Sheet1!A3:A3)", 7}, // one cell
	}
	for i, tt := range tests {
		ref := fmt.Sprintf("A%d", i+1)
		if err := out.SetFormula(ref, tt.formula); err != nil {
			t.Fatalf("SetFormula(%s): %v", tt.formula, err)
		}
		v, err := out.GetValue(ref)
		if err != nil {
			t.Fatal(err)
		}
		if v.Type != TypeNumber || v.Number != tt.want {
			t.Errorf("%s = %#v, want %v", tt.formula, v, tt.want)
		}
		// The formula text is stored as it was typed.
		if got, _ := out.GetFormula(ref); got != tt.formula {
			t.Errorf("formula text = %q, want %q", got, tt.formula)
		}
	}

	// A write to a cell of a reversed range recalculates the formula: the
	// dependency is on the full range.
	if err := s.SetValue("A3", 107); err != nil {
		t.Fatal(err)
	}
	if v, _ := out.GetValue("A2"); v.Number != 135 {
		t.Errorf("SUM(Sheet1!A5:A1) after a write to A3 = %v, want 135", v.Number)
	}
	f.Recalculate()
	if v, _ := out.GetValue("A4"); v.Number != 220 {
		t.Errorf("SUM(Sheet1!C5:A1) after a write to A3 = %v, want 220", v.Number)
	}
}

func TestParseOrdersRangeCorners(t *testing.T) {
	tests := []struct {
		src      string
		from, to formula.CellRef
	}{
		{"A5:A1", formula.CellRef{Col: 1, Row: 1}, formula.CellRef{Col: 1, Row: 5}},
		{"C1:A3", formula.CellRef{Col: 1, Row: 1}, formula.CellRef{Col: 3, Row: 3}},
		{"C3:A1", formula.CellRef{Col: 1, Row: 1}, formula.CellRef{Col: 3, Row: 3}},
		// Each flag stays with its row or column.
		{"$C3:A$1", formula.CellRef{Col: 1, Row: 1, AbsRow: true}, formula.CellRef{Col: 3, Row: 3, AbsCol: true}},
		{"C$3:$A1", formula.CellRef{Col: 1, Row: 1, AbsCol: true}, formula.CellRef{Col: 3, Row: 3, AbsRow: true}},
		// Ranges in the usual order do not change.
		{"A1:C3", formula.CellRef{Col: 1, Row: 1}, formula.CellRef{Col: 3, Row: 3}},
		{"$A$1:C3", formula.CellRef{Col: 1, Row: 1, AbsCol: true, AbsRow: true}, formula.CellRef{Col: 3, Row: 3}},
		{"B2:B2", formula.CellRef{Col: 2, Row: 2}, formula.CellRef{Col: 2, Row: 2}},
		// Column-only and row-only ranges.
		{"C:A", formula.CellRef{Col: 1, Row: 1}, formula.CellRef{Col: 3, Row: MaxRows}},
		{"$C:A", formula.CellRef{Col: 1, Row: 1}, formula.CellRef{Col: 3, Row: MaxRows, AbsCol: true}},
		{"3:1", formula.CellRef{Col: 1, Row: 1}, formula.CellRef{Col: MaxColumns, Row: 3}},
		{"$3:1", formula.CellRef{Col: 1, Row: 1}, formula.CellRef{Col: MaxColumns, Row: 3, AbsRow: true}},
	}
	for _, tt := range tests {
		node, err := formula.Parse(tt.src)
		if err != nil {
			t.Errorf("Parse(%q): %v", tt.src, err)
			continue
		}
		rng, ok := node.(*formula.RangeRef)
		if !ok {
			t.Errorf("Parse(%q) = %T, want a range", tt.src, node)
			continue
		}
		if *rng.From != tt.from || *rng.To != tt.to {
			t.Errorf("Parse(%q) = %+v .. %+v, want %+v .. %+v", tt.src, *rng.From, *rng.To, tt.from, tt.to)
		}
	}
	// The sheet qualifier stays on the first corner.
	node, err := formula.Parse("Data!B9:A2")
	if err != nil {
		t.Fatal(err)
	}
	if rng := node.(*formula.RangeRef); rng.From.Sheet != "Data" || rng.From.Row != 2 || rng.To.Row != 9 {
		t.Errorf("Parse(Data!B9:A2) = %+v .. %+v", *rng.From, *rng.To)
	}
}
