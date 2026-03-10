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

	s.SetValue("B1", "0.00") // Format string that looks like a number
	s.SetValue("C1", 12.344) // Number to format
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

	s.SetValue("A1", "42") // String that looks like a number
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
// returns 0 (TypeNumber), matching expected behavior where empty formula results
// are coerced to numeric zero.
func TestEmptyRefReturnsZero(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// B1 references A1 which is empty — shows/caches 0.
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

// TestWholeColumnRefSUMIF verifies that SUMIF with whole-column references
// like A:A correctly sums matching rows instead of returning 0.
func TestWholeColumnRefSUMIF(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Column A: categories, Column B: values
	s.SetValue("A1", "apple")
	s.SetValue("A2", "banana")
	s.SetValue("A3", "apple")
	s.SetValue("A4", "cherry")
	s.SetValue("A5", "apple")

	s.SetValue("B1", 10)
	s.SetValue("B2", 20)
	s.SetValue("B3", 30)
	s.SetValue("B4", 40)
	s.SetValue("B5", 50)

	// SUMIF with whole-column references: sum column B where column A = "apple"
	s.SetFormula("C1", `SUMIF(A:A,"apple",B:B)`)

	val, err := s.GetValue("C1")
	if err != nil {
		t.Fatalf("GetValue(C1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Errorf("C1 type = %v, want TypeNumber", val.Type)
	}
	// apple rows: B1=10, B3=30, B5=50 → total 90
	if val.Number != 90 {
		t.Errorf("SUMIF(A:A,\"apple\",B:B) = %g, want 90", val.Number)
	}
}

// TestWholeColumnRefMATCH verifies that MATCH with a whole-column reference
// correctly finds the matching row.
func TestWholeColumnRefMATCH(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", "cat")
	s.SetValue("A2", "dog")
	s.SetValue("A3", "bird")

	// MATCH with whole-column reference
	s.SetFormula("B1", `MATCH("dog",A:A,0)`)

	val, err := s.GetValue("B1")
	if err != nil {
		t.Fatalf("GetValue(B1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Errorf("B1 type = %v, want TypeNumber", val.Type)
	}
	if val.Number != 2 {
		t.Errorf("MATCH(\"dog\",A:A,0) = %g, want 2", val.Number)
	}
}

// TestWholeColumnRefCrossSheet verifies that whole-column references work
// correctly with cross-sheet references like Sheet2!R:R.
func TestWholeColumnRefCrossSheet(t *testing.T) {
	f := werkbook.New()
	s1 := f.Sheet("Sheet1")
	s2, err := f.NewSheet("Sheet2")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	// Set up data on Sheet2
	s2.SetValue("A1", "x")
	s2.SetValue("A2", "y")
	s2.SetValue("A3", "x")
	s2.SetValue("B1", 100)
	s2.SetValue("B2", 200)
	s2.SetValue("B3", 300)

	// SUMIF on Sheet1 referencing whole columns on Sheet2
	s1.SetFormula("A1", `SUMIF(Sheet2!A:A,"x",Sheet2!B:B)`)

	val, err := s1.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Errorf("A1 type = %v, want TypeNumber", val.Type)
	}
	// x rows: B1=100, B3=300 → total 400
	if val.Number != 400 {
		t.Errorf("SUMIF(Sheet2!A:A,\"x\",Sheet2!B:B) = %g, want 400", val.Number)
	}
}

func TestOversizedCrossSheetRangeReturnsREF(t *testing.T) {
	f := werkbook.New()
	s1 := f.Sheet("Sheet1")
	s2, err := f.NewSheet("Sheet2")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	s1.SetValue("B524289", 1)
	s2.SetFormula("A1", "SUM(Sheet1!A:B)")

	val, err := s2.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if val.Type != werkbook.TypeError || val.String != "#REF!" {
		t.Errorf("SUM(Sheet1!A:B) = %#v, want #REF!", val)
	}
}

func TestWholeSheetReferenceExtremeCoordinateReturnsREF(t *testing.T) {
	f := werkbook.New()
	data := f.Sheet("Sheet1")
	calc, err := f.NewSheet("Calc")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	data.SetValue("XFD1048576", 1)
	calc.SetFormula("A1", "SUM(Sheet1!A:XFD)")

	val, err := calc.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if val.Type != werkbook.TypeError || val.String != "#REF!" {
		t.Errorf("SUM(Sheet1!A:XFD) = %#v, want #REF!", val)
	}
}

// TestCrossSheetEmptyRefReturnsZero verifies that a cross-sheet reference to
// an empty cell returns 0, not empty. This matches expected behavior where
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

// TestFILTER_ArrayResultReturnsTopLeft verifies that FILTER formulas that
// return an array produce the top-left element in the anchor cell, matching
// dynamic array spill behavior for the formula cell.
func TestFILTER_ArrayResultReturnsTopLeft(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set up data with some non-empty values and gaps.
	s.SetValue("A1", "hello")
	// A2 is empty
	s.SetValue("A3", "world")
	// A4, A5 are empty

	// FILTER(A1:A5, A1:A5<>"") should return {"hello";"world"} as an array.
	// The anchor cell should get the top-left element: "hello".
	s.SetFormula("B1", `FILTER(A1:A5,A1:A5<>"")`)

	val, err := s.GetValue("B1")
	if err != nil {
		t.Fatalf("GetValue(B1): %v", err)
	}
	if val.Type != werkbook.TypeString {
		t.Errorf("B1 type = %v, want TypeString", val.Type)
	}
	if val.String != "hello" {
		t.Errorf("B1 = %q, want %q", val.String, "hello")
	}
}

// TestFILTER_CrossSheetArrayResult verifies that FILTER works with
// cross-sheet references and the array result returns the first element.
func TestFILTER_CrossSheetArrayResult(t *testing.T) {
	f := werkbook.New()
	data, _ := f.NewSheet("Data")
	out := f.Sheet("Sheet1")

	data.SetValue("A1", "alpha")
	data.SetValue("A2", "beta")
	data.SetValue("A3", "gamma")
	// A4..A10 are empty

	out.SetFormula("A1", `FILTER(Data!A1:A10,Data!A1:A10<>"")`)

	val, err := out.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if val.Type != werkbook.TypeString {
		t.Errorf("A1 type = %v, want TypeString", val.Type)
	}
	if val.String != "alpha" {
		t.Errorf("A1 = %q, want %q", val.String, "alpha")
	}
}

// TestSORTUNIQUEFILTER_NestedArrayFunctions verifies the common pattern
// SORT(UNIQUE(FILTER(...))) which chains multiple dynamic array functions.
func TestSORTUNIQUEFILTER_NestedArrayFunctions(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", "cherry")
	s.SetValue("A2", "apple")
	s.SetValue("A3", "banana")
	s.SetValue("A4", "apple") // duplicate
	// A5..A10 are empty

	// SORT(UNIQUE(FILTER(A1:A10, A1:A10<>""))) should produce
	// {"apple";"banana";"cherry"} and the anchor cell gets "apple".
	s.SetFormula("B1", `SORT(UNIQUE(FILTER(A1:A10,A1:A10<>"")))`)

	val, err := s.GetValue("B1")
	if err != nil {
		t.Fatalf("GetValue(B1): %v", err)
	}
	if val.Type != werkbook.TypeString {
		t.Errorf("B1 type = %v, want TypeString", val.Type)
	}
	if val.String != "apple" {
		t.Errorf("B1 = %q, want %q", val.String, "apple")
	}
}

// TestFILTER_NumericArrayResult verifies FILTER with numeric data returns
// the top-left number from the filtered array.
func TestFILTER_NumericArrayResult(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 100)
	s.SetValue("A2", 200)
	// A3..A5 are empty
	s.SetValue("B1", "yes")
	s.SetValue("B2", "no")
	// B3..B5 are empty

	// FILTER(A1:A5, B1:B5<>"") should return {100;200}, anchor gets 100.
	s.SetFormula("C1", `FILTER(A1:A5,B1:B5<>"")`)

	val, err := s.GetValue("C1")
	if err != nil {
		t.Fatalf("GetValue(C1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Errorf("C1 type = %v, want TypeNumber", val.Type)
	}
	if val.Number != 100 {
		t.Errorf("C1 = %g, want 100", val.Number)
	}
}

// TestINDEX_RowZero_ReturnsValueError verifies that INDEX(range,0) returns
// #VALUE! in a non-array cell, matching expected behaviour.
func TestINDEX_RowZero_ReturnsValueError(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 10)
	s.SetValue("A2", 20)
	s.SetValue("A3", 30)

	// INDEX with row_num=0 on a column vector should return #VALUE!
	// in a regular (non-array) cell.
	s.SetFormula("B1", "INDEX(A1:A3,0)")
	val, err := s.GetValue("B1")
	if err != nil {
		t.Fatalf("GetValue B1: %v", err)
	}
	if val.Type != werkbook.TypeError {
		t.Errorf("INDEX(A1:A3,0) type = %v, want TypeError", val.Type)
	}
	if val.String != "#VALUE!" {
		t.Errorf("INDEX(A1:A3,0) = %q, want #VALUE!", val.String)
	}

	// SUM(INDEX(range,0)) should still work — SUM consumes the array.
	s.SetFormula("C1", "SUM(INDEX(A1:A3,0))")
	val, err = s.GetValue("C1")
	if err != nil {
		t.Fatalf("GetValue C1: %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Errorf("SUM(INDEX(A1:A3,0)) type = %v, want TypeNumber", val.Type)
	}
	if val.Number != 60 {
		t.Errorf("SUM(INDEX(A1:A3,0)) = %g, want 60", val.Number)
	}
}
