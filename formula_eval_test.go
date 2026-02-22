package werkbook_test

import (
	"testing"

	"github.com/jpoz/werkbook"
)

// TestFormulaEvaluation verifies that formula cells compute their values.
// Currently fails because the formula engine (formula/ package) is not
// connected to the public API — cells store formula text but never evaluate it.
func TestFormulaEvaluation(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 10)
	s.SetValue("A2", 20)
	s.SetValue("A3", 30)

	s.SetFormula("B1", "SUM(A1:A3)")
	s.SetFormula("B2", "A1*A2")
	s.SetFormula("B3", `IF(A1>5,"yes","no")`)

	tests := []struct {
		cell    string
		wantNum float64
		wantStr string
		wantTyp werkbook.ValueType
	}{
		{"B1", 60, "", werkbook.TypeNumber},
		{"B2", 200, "", werkbook.TypeNumber},
		{"B3", 0, "yes", werkbook.TypeString},
	}

	for _, tt := range tests {
		val, err := s.GetValue(tt.cell)
		if err != nil {
			t.Errorf("GetValue(%s): %v", tt.cell, err)
			continue
		}
		if val.Type != tt.wantTyp {
			t.Errorf("GetValue(%s).Type = %v, want %v", tt.cell, val.Type, tt.wantTyp)
			continue
		}
		switch tt.wantTyp {
		case werkbook.TypeNumber:
			if val.Number != tt.wantNum {
				t.Errorf("GetValue(%s).Number = %g, want %g", tt.cell, val.Number, tt.wantNum)
			}
		case werkbook.TypeString:
			if val.String != tt.wantStr {
				t.Errorf("GetValue(%s).String = %q, want %q", tt.cell, val.String, tt.wantStr)
			}
		}
	}
}
