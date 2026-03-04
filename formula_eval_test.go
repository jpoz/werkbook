package werkbook_test

import (
	"testing"

	"github.com/jpoz/werkbook"
)

// TestFormulaEvaluation verifies that formula cells compute their values.
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

// TestTEXTWithStringFormatCell verifies that TEXT() correctly uses a string
// format from a cell reference, even when the format looks like a number (e.g. "0.00").
func TestTEXTWithStringFormatCell(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("B1", "0.00")   // Format string that looks like a number
	s.SetValue("C1", 12.344)   // Number to format
	s.SetFormula("A1", `TEXT(C1, B1)`)

	val, err := s.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if val.Type != werkbook.TypeString {
		t.Errorf("A1 type = %v, want TypeString", val.Type)
	}
	if val.String != "12.34" {
		t.Errorf("A1 = %q, want %q", val.String, "12.34")
	}
}

// TestArithmeticWithStringCell verifies that arithmetic operations correctly
// coerce string cell values that look like numbers.
func TestArithmeticWithStringCell(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", "42")  // String that looks like a number
	s.SetValue("A2", 8)
	s.SetFormula("B1", "A1+A2")

	val, err := s.GetValue("B1")
	if err != nil {
		t.Fatalf("GetValue(B1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Errorf("B1 type = %v, want TypeNumber", val.Type)
	}
	if val.Number != 50 {
		t.Errorf("B1 = %g, want 50", val.Number)
	}
}

// TestEmptyRefReturnsZero verifies that a formula referencing an empty cell
// returns 0 (TypeNumber), matching Excel behavior where empty formula results
// are coerced to numeric zero.
func TestEmptyRefReturnsZero(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// B1 references A1 which is empty — Excel would show/cache 0.
	s.SetFormula("B1", "A1")

	val, err := s.GetValue("B1")
	if err != nil {
		t.Fatalf("GetValue(B1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Errorf("B1 type = %v, want TypeNumber", val.Type)
	}
	if val.Number != 0 {
		t.Errorf("B1 = %g, want 0", val.Number)
	}
}

// TestCrossSheetEmptyRefReturnsZero verifies that a cross-sheet reference to
// an empty cell returns 0, not empty. This matches Excel behavior where
// formulas like ='Sheet2'!A1 (with A1 empty) cache 0.
func TestCrossSheetEmptyRefReturnsZero(t *testing.T) {
	f := werkbook.New()
	s1 := f.Sheet("Sheet1")
	_, err := f.NewSheet("Sheet2")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	// Reference empty cell on Sheet2
	s1.SetFormula("A1", "'Sheet2'!A1")

	val, err := s1.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Errorf("A1 type = %v, want TypeNumber", val.Type)
	}
	if val.Number != 0 {
		t.Errorf("A1 = %g, want 0", val.Number)
	}
}
