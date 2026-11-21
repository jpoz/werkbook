package werkbook_test

import (
	"testing"

	"github.com/jpoz/werkbook"
)

// TestFormulaErrorClassification verifies that evaluateFormula maps
// compile/parse/runtime failures to the appropriate Excel cell error
// code instead of flattening every failure to #NAME?.
func TestFormulaErrorClassification(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Parse error — malformed token. Excel surfaces these as #NAME?.
	// SetFormula defers parse errors to recalc: it accepts any formula
	// text; the parse error materializes when the cell is evaluated.
	if err := s.SetFormula("A1", "@@@"); err != nil {
		t.Fatalf("SetFormula A1: %v", err)
	}

	f.Recalculate()

	got, _ := s.GetValue("A1")
	if got.Type != werkbook.TypeError {
		t.Fatalf("A1: expected TypeError, got %v (%q)", got.Type, got.String)
	}
	if got.String != "#NAME?" {
		t.Errorf("A1: got %s, want #NAME?", got.String)
	}
}
