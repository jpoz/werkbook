package werkbook

import (
	"testing"

	"github.com/jpoz/werkbook/formula"
	"github.com/jpoz/werkbook/ooxml"
)

func TestFormulaValueToValueErrorsAreAlwaysErrors(t *testing.T) {
	// Formula errors must always be represented as TypeError, regardless of
	// whether the cell's cached value was a string. This ensures functions
	// like T() that check for ValueError correctly propagate errors.
	got := formulaValueToValue(formula.ErrorVal(formula.ErrValNAME), true)
	if got.Type != TypeError || got.String != "#NAME?" {
		t.Fatalf("formulaValueToValue(#NAME?, array=true) = %#v, want error #NAME?", got)
	}

	got = formulaValueToValue(formula.ErrorVal(formula.ErrValNAME), false)
	if got.Type != TypeError || got.String != "#NAME?" {
		t.Fatalf("formulaValueToValue(#NAME?, array=false) = %#v, want error #NAME?", got)
	}
}

func TestCellDataToValueKeepsFormulaStrCellsAsStrings(t *testing.T) {
	got := cellDataToValue(ooxml.CellData{
		Type:           "str",
		Value:          "#NAME?",
		Formula:        `LET(day,2,_xlfn.SWITCH(day,1,"Sun",2,Data!F1,"?"))`,
		FormulaType:    "array",
		IsArrayFormula: true,
	}, nil, false)
	if got.Type != TypeString || got.String != "#NAME?" {
		t.Fatalf("cellDataToValue(formula str #NAME?) = %#v, want string #NAME?", got)
	}

	got = cellDataToValue(ooxml.CellData{
		Type:           "str",
		Value:          "#CALC!",
		Formula:        `_xlfn._xlws.FILTER(B1:B10,B1:B10<>"")`,
		FormulaType:    "array",
		IsDynamicArray: true,
	}, nil, false)
	if got.Type != TypeString || got.String != "#CALC!" {
		t.Fatalf("cellDataToValue(dynamic array formula str #CALC!) = %#v, want string #CALC!", got)
	}
}
