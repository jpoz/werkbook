package werkbook

import (
	"testing"

	"github.com/jpoz/werkbook/formula"
)

func TestFileResolverFullRowRangeClampsToPopulatedColumns(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	if err := s.SetValue("A1", 10); err != nil {
		t.Fatalf("SetValue(A1): %v", err)
	}
	if err := s.SetValue("C1", 30); err != nil {
		t.Fatalf("SetValue(C1): %v", err)
	}

	fr := &fileResolver{file: f, currentSheet: "Sheet1"}
	rows := fr.GetRangeValues(formula.RangeAddr{
		FromCol: 1,
		FromRow: 1,
		ToCol:   MaxColumns,
		ToRow:   1,
	})

	if len(rows) != 1 {
		t.Fatalf("len(rows) = %d, want 1", len(rows))
	}
	if len(rows[0]) != 3 {
		t.Fatalf("len(rows[0]) = %d, want 3", len(rows[0]))
	}
	if rows[0][0].Type != formula.ValueNumber || rows[0][0].Num != 10 {
		t.Fatalf("rows[0][0] = %#v, want 10", rows[0][0])
	}
	if rows[0][1].Type != formula.ValueEmpty {
		t.Fatalf("rows[0][1] = %#v, want empty", rows[0][1])
	}
	if rows[0][2].Type != formula.ValueNumber || rows[0][2].Num != 30 {
		t.Fatalf("rows[0][2] = %#v, want 30", rows[0][2])
	}
}

func TestFileResolverWholeSheetRangeOverflowReturnsREF(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	if err := s.SetValue("XFD1048576", 1); err != nil {
		t.Fatalf("SetValue(XFD1048576): %v", err)
	}

	fr := &fileResolver{file: f, currentSheet: "Sheet1"}
	rows := fr.GetRangeValues(formula.RangeAddr{
		FromCol: 1,
		FromRow: 1,
		ToCol:   MaxColumns,
		ToRow:   MaxRows,
	})

	if len(rows) != 1 || len(rows[0]) != 1 {
		t.Fatalf("unexpected overflow shape: %dx%d", len(rows), len(rows[0]))
	}
	cell := rows[0][0]
	if cell.Type != formula.ValueError || cell.Err != formula.ErrValREF || !cell.RangeOverflow {
		t.Fatalf("overflow sentinel = %#v, want #REF! overflow", cell)
	}
}
