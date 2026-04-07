package werkbook_test

import (
	"fmt"
	"math"
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

func TestRANDARRAYSpillCellsRemainConsistentWithinCalculation(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetFormula("B3", "RANDARRAY(1,3)")
	s.SetFormula("E3", "SUM(B3:D3)")

	f.Recalculate()

	wantCells := []string{"B3", "C3", "D3"}
	values := make([]float64, 0, len(wantCells))
	for _, cell := range wantCells {
		val, err := s.GetValue(cell)
		if err != nil {
			t.Fatalf("GetValue(%s): %v", cell, err)
		}
		if val.Type != werkbook.TypeNumber {
			t.Fatalf("%s type = %v, want TypeNumber", cell, val.Type)
		}
		if val.Number < 0 || val.Number >= 1 {
			t.Fatalf("%s = %g, want [0,1)", cell, val.Number)
		}
		values = append(values, val.Number)
	}

	total, err := s.GetValue("E3")
	if err != nil {
		t.Fatalf("GetValue(E3): %v", err)
	}
	if total.Type != werkbook.TypeNumber {
		t.Fatalf("E3 type = %v, want TypeNumber", total.Type)
	}

	want := values[0] + values[1] + values[2]
	if math.Abs(total.Number-want) > 1e-12 {
		t.Fatalf("SUM(B3:D3) = %g, want %g", total.Number, want)
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

func newFilteredNumberSpill(values ...float64) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet) {
	f := werkbook.New()
	_ = f.SetSheetName("Sheet1", "Data")
	data := f.Sheet("Data")

	for i, value := range values {
		row := i + 2
		data.SetValue(fmt.Sprintf("A%d", row), true)
		data.SetValue(fmt.Sprintf("B%d", row), value)
	}

	spill, _ := f.NewSheet("Spill")
	spill.SetFormula("B1", fmt.Sprintf(`FILTER(Data!B2:B%d,Data!A2:A%d)`, len(values)+1, len(values)+1))

	calc, _ := f.NewSheet("Calc")
	return f, spill, calc
}

func newFilteredTextSpill(values ...string) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet) {
	f := werkbook.New()
	_ = f.SetSheetName("Sheet1", "Data")
	data := f.Sheet("Data")

	for i, value := range values {
		row := i + 2
		data.SetValue(fmt.Sprintf("A%d", row), true)
		data.SetValue(fmt.Sprintf("B%d", row), value)
	}

	spill, _ := f.NewSheet("Spill")
	spill.SetFormula("B1", fmt.Sprintf(`FILTER(Data!B2:B%d,Data!A2:A%d)`, len(values)+1, len(values)+1))

	calc, _ := f.NewSheet("Calc")
	return f, spill, calc
}

func newFilteredLookupSpill(keys []string, values []float64) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet) {
	f := werkbook.New()
	_ = f.SetSheetName("Sheet1", "Data")
	data := f.Sheet("Data")

	for i, key := range keys {
		row := i + 2
		data.SetValue(fmt.Sprintf("A%d", row), true)
		data.SetValue(fmt.Sprintf("B%d", row), key)
		data.SetValue(fmt.Sprintf("C%d", row), values[i])
	}

	spill, _ := f.NewSheet("Spill")
	spill.SetFormula("A1", fmt.Sprintf(`FILTER(Data!B2:B%d,Data!A2:A%d)`, len(keys)+1, len(keys)+1))
	spill.SetFormula("B1", fmt.Sprintf(`FILTER(Data!C2:C%d,Data!A2:A%d)`, len(values)+1, len(values)+1))

	calc, _ := f.NewSheet("Calc")
	return f, spill, calc
}

// TestSUM_IncludesSpillRows_FullColumn verifies that SUM over a full-column
// reference (e.g. B:B) includes all rows produced by a FILTER dynamic array
// spill, not just the anchor cell.
func TestSUM_IncludesSpillRows_FullColumn(t *testing.T) {
	f := werkbook.New()
	data := f.Sheet("Sheet1")
	f.SetSheetName("Sheet1", "Data")
	data = f.Sheet("Data")

	data.SetValue("A1", "Include")
	data.SetValue("B1", "Amount")
	data.SetValue("A2", true)
	data.SetValue("B2", 10.0)
	data.SetValue("A3", false)
	data.SetValue("B3", 20.0)
	data.SetValue("A4", true)
	data.SetValue("B4", 30.0)
	data.SetValue("A5", false)
	data.SetValue("B5", 40.0)
	data.SetValue("A6", true)
	data.SetValue("B6", 50.0)

	spill, _ := f.NewSheet("Spill")
	spill.SetValue("B1", "Filtered")
	spill.SetFormula("B2", `FILTER(Data!B2:B6,Data!A2:A6)`)

	sum, _ := f.NewSheet("Sum")
	sum.SetFormula("A1", `SUM(Spill!B:B)`)

	f.Recalculate()

	val, err := sum.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Fatalf("A1 type = %v, want TypeNumber", val.Type)
	}
	if val.Number != 90 {
		t.Errorf("SUM(Spill!B:B) = %g, want 90", val.Number)
	}
}

// TestSUM_IncludesSpillRows_BoundedRange verifies that SUM over a bounded
// range that extends beyond the last physical row still picks up spill values.
func TestSUM_IncludesSpillRows_BoundedRange(t *testing.T) {
	f := werkbook.New()
	data := f.Sheet("Sheet1")
	f.SetSheetName("Sheet1", "Data")
	data = f.Sheet("Data")

	data.SetValue("A1", "Include")
	data.SetValue("B1", "Amount")
	data.SetValue("A2", true)
	data.SetValue("B2", 10.0)
	data.SetValue("A3", false)
	data.SetValue("B3", 20.0)
	data.SetValue("A4", true)
	data.SetValue("B4", 30.0)
	data.SetValue("A5", false)
	data.SetValue("B5", 40.0)
	data.SetValue("A6", true)
	data.SetValue("B6", 50.0)

	spill, _ := f.NewSheet("Spill")
	spill.SetValue("B1", "Filtered")
	spill.SetFormula("B2", `FILTER(Data!B2:B6,Data!A2:A6)`)

	sum, _ := f.NewSheet("Sum")
	sum.SetFormula("A1", `SUM(Spill!B1:B10)`)

	f.Recalculate()

	val, err := sum.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Fatalf("A1 type = %v, want TypeNumber", val.Type)
	}
	if val.Number != 90 {
		t.Errorf("SUM(Spill!B1:B10) = %g, want 90", val.Number)
	}
}

// TestAVERAGE_IncludesSpillRows verifies that AVERAGE over a range including
// spill rows computes the correct average.
func TestAVERAGE_IncludesSpillRows(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 100.0)
	s.SetValue("A2", 200.0)
	s.SetValue("A3", 300.0)

	// FILTER returns {100; 300} (rows where value != 200)
	s.SetFormula("B1", `FILTER(A1:A3,A1:A3<>200)`)
	s.SetFormula("C1", `AVERAGE(B:B)`)

	f.Recalculate()

	val, err := s.GetValue("C1")
	if err != nil {
		t.Fatalf("GetValue(C1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Fatalf("C1 type = %v, want TypeNumber", val.Type)
	}
	// AVERAGE(100, 300) = 200
	if val.Number != 200 {
		t.Errorf("AVERAGE(B:B) = %g, want 200", val.Number)
	}
}

// TestCOUNT_IncludesSpillRows verifies that COUNT over a column with spill
// rows counts all spilled numeric values.
func TestCOUNT_IncludesSpillRows(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 10.0)
	s.SetValue("A2", 20.0)
	s.SetValue("A3", 30.0)
	s.SetValue("A4", 40.0)
	s.SetValue("A5", 50.0)

	// FILTER returns {10; 30; 50} (odd-indexed)
	s.SetValue("B1", true)
	s.SetValue("B2", false)
	s.SetValue("B3", true)
	s.SetValue("B4", false)
	s.SetValue("B5", true)
	s.SetFormula("C1", `FILTER(A1:A5,B1:B5)`)
	s.SetFormula("D1", `COUNT(C:C)`)

	f.Recalculate()

	val, err := s.GetValue("D1")
	if err != nil {
		t.Fatalf("GetValue(D1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Fatalf("D1 type = %v, want TypeNumber", val.Type)
	}
	if val.Number != 3 {
		t.Errorf("COUNT(C:C) = %g, want 3", val.Number)
	}
}

// TestCOUNTA_IncludesTextSpillRows verifies that COUNTA over a full-column
// reference includes spilled text values beyond the anchor row.
func TestCOUNTA_IncludesTextSpillRows(t *testing.T) {
	f, _, calc := newFilteredTextSpill("beta", "alpha", "gamma")
	calc.SetFormula("A1", `COUNTA(Spill!B:B)`)

	f.Recalculate()

	val, err := calc.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Fatalf("A1 type = %v, want TypeNumber", val.Type)
	}
	if val.Number != 3 {
		t.Errorf("COUNTA(Spill!B:B) = %g, want 3", val.Number)
	}
}

// TestMINMAX_IncludesSpillRows verifies that full-column MIN/MAX formulas
// see spilled rows after the anchor cell.
func TestMINMAX_IncludesSpillRows(t *testing.T) {
	f, _, calc := newFilteredNumberSpill(30, 10, 50)
	calc.SetFormula("A1", `MIN(Spill!B:B)`)
	calc.SetFormula("A2", `MAX(Spill!B:B)`)

	f.Recalculate()

	minVal, err := calc.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if minVal.Type != werkbook.TypeNumber {
		t.Fatalf("A1 type = %v, want TypeNumber", minVal.Type)
	}
	if minVal.Number != 10 {
		t.Errorf("MIN(Spill!B:B) = %g, want 10", minVal.Number)
	}

	maxVal, err := calc.GetValue("A2")
	if err != nil {
		t.Fatalf("GetValue(A2): %v", err)
	}
	if maxVal.Type != werkbook.TypeNumber {
		t.Fatalf("A2 type = %v, want TypeNumber", maxVal.Type)
	}
	if maxVal.Number != 50 {
		t.Errorf("MAX(Spill!B:B) = %g, want 50", maxVal.Number)
	}
}

// TestSUMIFCOUNTIF_IncludeSpillRows verifies that criteria-based functions
// operating over full-column references include all spilled values.
func TestSUMIFCOUNTIF_IncludeSpillRows(t *testing.T) {
	f, _, calc := newFilteredNumberSpill(30, 10, 50)
	calc.SetFormula("A1", `SUMIF(Spill!B:B,">20")`)
	calc.SetFormula("A2", `COUNTIF(Spill!B:B,">20")`)

	f.Recalculate()

	sumVal, err := calc.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if sumVal.Type != werkbook.TypeNumber {
		t.Fatalf("A1 type = %v, want TypeNumber", sumVal.Type)
	}
	if sumVal.Number != 80 {
		t.Errorf("SUMIF(Spill!B:B,\">20\") = %g, want 80", sumVal.Number)
	}

	countVal, err := calc.GetValue("A2")
	if err != nil {
		t.Fatalf("GetValue(A2): %v", err)
	}
	if countVal.Type != werkbook.TypeNumber {
		t.Fatalf("A2 type = %v, want TypeNumber", countVal.Type)
	}
	if countVal.Number != 2 {
		t.Errorf("COUNTIF(Spill!B:B,\">20\") = %g, want 2", countVal.Number)
	}
}

// TestMATCHXLOOKUP_FindSpillFollowers verifies that lookup functions reading
// full-column ranges can find spilled values beyond the anchor row.
func TestMATCHXLOOKUP_FindSpillFollowers(t *testing.T) {
	f, _, calc := newFilteredLookupSpill(
		[]string{"beta", "alpha", "gamma"},
		[]float64{200, 100, 300},
	)
	calc.SetFormula("A1", `MATCH("gamma",Spill!A:A,0)`)
	calc.SetFormula("A2", `XLOOKUP("gamma",Spill!A:A,Spill!B:B)`)

	f.Recalculate()

	matchVal, err := calc.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if matchVal.Type != werkbook.TypeNumber {
		t.Fatalf("A1 type = %v, want TypeNumber", matchVal.Type)
	}
	if matchVal.Number != 3 {
		t.Errorf("MATCH(\"gamma\",Spill!A:A,0) = %g, want 3", matchVal.Number)
	}

	xlookupVal, err := calc.GetValue("A2")
	if err != nil {
		t.Fatalf("GetValue(A2): %v", err)
	}
	if xlookupVal.Type != werkbook.TypeNumber {
		t.Fatalf("A2 type = %v, want TypeNumber", xlookupVal.Type)
	}
	if xlookupVal.Number != 300 {
		t.Errorf("XLOOKUP(\"gamma\",Spill!A:A,Spill!B:B) = %g, want 300", xlookupVal.Number)
	}
}

// TestSUM_SpillOnSameSheet verifies that SUM picks up spill rows when
// the FILTER formula and the SUM are on the same sheet.
func TestSUM_SpillOnSameSheet(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", true)
	s.SetValue("A2", true)
	s.SetValue("A3", true)
	s.SetValue("B1", 1000.0)
	s.SetValue("B2", 2000.0)
	s.SetValue("B3", 3000.0)

	// FILTER in C1 spills to C1:C3
	s.SetFormula("C1", `FILTER(B1:B3,A1:A3)`)
	// SUM over full column C
	s.SetFormula("D1", `SUM(C:C)`)

	f.Recalculate()

	val, err := s.GetValue("D1")
	if err != nil {
		t.Fatalf("GetValue(D1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Fatalf("D1 type = %v, want TypeNumber", val.Type)
	}
	if val.Number != 6000 {
		t.Errorf("SUM(C:C) = %g, want 6000", val.Number)
	}
}

// TestSUM_FullRowIncludesHorizontalSpill verifies that whole-row references
// include dynamic-array spill values that extend past the last physical column.
func TestSUM_FullRowIncludesHorizontalSpill(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetFormula("B1", `HSTACK(10,20,30)`)
	s.SetFormula("A2", `SUM(1:1)`)

	f.Recalculate()

	val, err := s.GetValue("A2")
	if err != nil {
		t.Fatalf("GetValue(A2): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Fatalf("A2 type = %v, want TypeNumber", val.Type)
	}
	if val.Number != 60 {
		t.Errorf("SUM(1:1) = %g, want 60", val.Number)
	}
}

// TestSUM_FullColumnIgnoresUnrelatedTallSpill verifies that a spill outside
// the referenced columns does not change the materialized height of B:C.
func TestSUM_FullColumnIgnoresUnrelatedTallSpill(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("B1", 1.0)
	s.SetValue("C1", 2.0)
	s.SetFormula("A1", `SUM(B:C)`)
	s.SetFormula("Z1", `SEQUENCE(600000)`)

	f.Recalculate()

	val, err := s.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if val.Type != werkbook.TypeNumber {
		t.Fatalf("A1 type = %v, want TypeNumber", val.Type)
	}
	if val.Number != 3 {
		t.Errorf("SUM(B:C) = %g, want 3", val.Number)
	}
}

// TestSUM_FullColumnSkipsUnrelatedSpillEval verifies that reading B:B does not
// force an unrelated spill anchor outside the requested column into a circular error.
func TestSUM_FullColumnSkipsUnrelatedSpillEval(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("B1", 1.0)
	s.SetFormula("A1", `SUM(B:B)`)
	s.SetFormula("Z1", `SEQUENCE(A1)`)

	f.Recalculate()

	sum, err := s.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if sum.Type != werkbook.TypeNumber || sum.Number != 1 {
		t.Fatalf("A1 = %#v, want 1", sum)
	}

	spill, err := s.GetValue("Z1")
	if err != nil {
		t.Fatalf("GetValue(Z1): %v", err)
	}
	if spill.Type != werkbook.TypeNumber {
		t.Fatalf("Z1 type = %v, want TypeNumber", spill.Type)
	}
	if spill.Number != 1 {
		t.Errorf("Z1 = %g, want 1", spill.Number)
	}
}

// TestFILTER_FullColumnWithDATEVALUE verifies that FILTER with full-column
// references inside IFERROR(DATEVALUE(...)) correctly propagates array context
// through nested non-array functions.
func TestFILTER_FullColumnWithDATEVALUE(t *testing.T) {
	f := werkbook.New()
	data := f.Sheet("Sheet1")
	f.SetSheetName("Sheet1", "Data")
	data = f.Sheet("Data")

	data.SetValue("E1", "Accrued On")
	data.SetValue("I1", "Amount")
	data.SetValue("E2", "2026-07-01")
	data.SetValue("I2", 216994.0)
	data.SetValue("E3", "2026-10-01")
	data.SetValue("I3", 216994.0)
	data.SetValue("E4", "2022-01-01")
	data.SetValue("I4", 50000.0)

	out, _ := f.NewSheet("Out")
	out.SetFormula("A1", `FILTER(Data!I:I/100,IFERROR(DATEVALUE(Data!E:E),0)>=TODAY(),"")`)

	sum, _ := f.NewSheet("Sum")
	sum.SetFormula("A1", `SUM('Out'!A:A)`)

	f.Recalculate()

	// A1 should have the first filtered result (216994/100 = 2169.94).
	v, err := out.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(A1): %v", err)
	}
	if v.Type != werkbook.TypeNumber {
		t.Fatalf("A1 type = %v, want TypeNumber", v.Type)
	}
	if v.Number != 2169.94 {
		t.Errorf("FILTER A1 = %g, want 2169.94", v.Number)
	}

	// SUM should include both filtered rows.
	s, err := sum.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue(Sum!A1): %v", err)
	}
	if s.Type != werkbook.TypeNumber {
		t.Fatalf("Sum!A1 type = %v, want TypeNumber", s.Type)
	}
	if s.Number != 4339.88 {
		t.Errorf("SUM(Out!A:A) = %g, want 4339.88", s.Number)
	}
}
