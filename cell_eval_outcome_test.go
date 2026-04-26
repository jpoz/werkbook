package werkbook

import (
	"testing"

	"github.com/jpoz/werkbook/formula"
)

func TestFormulaDisplayValueAtImplicitIntersectionSingleColumnRef(t *testing.T) {
	origin := formula.RangeAddr{Sheet: "Data", FromCol: 1, FromRow: 2, ToCol: 1, ToRow: 4}
	got := formulaDisplayValueAt(
		formula.Value{
			Type: formula.ValueArray,
			Array: [][]formula.Value{
				{formula.NumberVal(10)},
				{formula.NumberVal(20)},
				{formula.NumberVal(30)},
			},
			RangeOrigin: &origin,
		},
		false,
		true,
		5,
		3,
	)
	if got.Type != formula.ValueNumber || got.Num != 20 {
		t.Fatalf("formulaDisplayValueAt single-column intersection = %#v, want 20", got)
	}
}

func TestFormulaDisplayValueAtImplicitIntersectionSingleRowRef(t *testing.T) {
	origin := formula.RangeAddr{Sheet: "Data", FromCol: 1, FromRow: 5, ToCol: 3, ToRow: 5}
	got := formulaDisplayValueAt(
		formula.Value{
			Type: formula.ValueArray,
			Array: [][]formula.Value{{
				formula.NumberVal(10),
				formula.NumberVal(20),
				formula.NumberVal(30),
			}},
			RangeOrigin: &origin,
		},
		false,
		true,
		2,
		9,
	)
	if got.Type != formula.ValueNumber || got.Num != 20 {
		t.Fatalf("formulaDisplayValueAt single-row intersection = %#v, want 20", got)
	}
}

func TestFormulaDisplayValueAtNoSpillArrayReturnsValueError(t *testing.T) {
	got := formulaDisplayValueAt(
		formula.Value{
			Type:    formula.ValueArray,
			Array:   [][]formula.Value{{formula.NumberVal(1), formula.NumberVal(2)}},
			NoSpill: true,
		},
		false,
		true,
		1,
		1,
	)
	if got.Type != formula.ValueError || got.Err != formula.ErrValVALUE {
		t.Fatalf("formulaDisplayValueAt NoSpill = %#v, want #VALUE!", got)
	}
}
