package formula

import (
	"math"
	"testing"
)

func TestMedian(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 1, Row: 3}: NumberVal(5),
		},
	}

	cf := evalCompile(t, "MEDIAN(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("MEDIAN odd: got %g, want 3", got.Num)
	}

	resolver.cells[CellAddr{Col: 1, Row: 4}] = NumberVal(7)
	cf = evalCompile(t, "MEDIAN(A1:A4)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("MEDIAN even: got %g, want 4", got.Num)
	}
}

func TestLargeSmall(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(30),
			{Col: 1, Row: 3}: NumberVal(20),
		},
	}

	cf := evalCompile(t, "LARGE(A1:A3,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("LARGE k=1: got %g, want 30", got.Num)
	}

	cf = evalCompile(t, "SMALL(A1:A3,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("SMALL k=1: got %g, want 10", got.Num)
	}
}

func TestCountBlank(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			// A2 is empty
			{Col: 1, Row: 3}: NumberVal(3),
		},
	}

	cf := evalCompile(t, "COUNTBLANK(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTBLANK: got %g, want 1", got.Num)
	}
}

func TestSUMIF(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}

	cf := evalCompile(t, `SUMIF(A1:A3,">15")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 50 {
		t.Errorf("SUMIF >15: got %g, want 50", got.Num)
	}
}

func TestCOUNTIF(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple"),
			{Col: 1, Row: 2}: StringVal("banana"),
			{Col: 1, Row: 3}: StringVal("apple"),
		},
	}

	cf := evalCompile(t, `COUNTIF(A1:A3,"apple")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIF: got %g, want 2", got.Num)
	}
}

func TestSUMPRODUCT(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}

	// 1*4 + 2*5 + 3*6 = 4 + 10 + 18 = 32
	cf := evalCompile(t, "SUMPRODUCT(A1:A3,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 32 {
		t.Errorf("SUMPRODUCT: got %g, want 32", got.Num)
	}
}

func TestMatchesCriteria(t *testing.T) {
	tests := []struct {
		v    Value
		crit Value
		want bool
	}{
		{NumberVal(10), StringVal(">5"), true},
		{NumberVal(3), StringVal(">5"), false},
		{NumberVal(5), StringVal(">=5"), true},
		{NumberVal(5), StringVal("<=5"), true},
		{NumberVal(5), StringVal("<>5"), false},
		{NumberVal(5), NumberVal(5), true},
		{StringVal("apple"), StringVal("app*"), true},
		{StringVal("banana"), StringVal("app*"), false},
		{StringVal("cat"), StringVal("c?t"), true},
	}

	for _, tt := range tests {
		got := matchesCriteria(tt.v, tt.crit)
		if got != tt.want {
			t.Errorf("matchesCriteria(%v, %v) = %v, want %v", tt.v, tt.crit, got, tt.want)
		}
	}
}

// ---------------------------------------------------------------------------
// COUNTA — counts all non-empty cells
// ---------------------------------------------------------------------------

func TestCOUNTA(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: StringVal("hello"),
			// A3 is empty
			{Col: 1, Row: 4}: BoolVal(true),
			{Col: 1, Row: 5}: ErrorVal(ErrValNA),
		},
	}

	cf := evalCompile(t, "COUNTA(A1:A5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Number, String, Bool, Error = 4 non-empty cells
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("COUNTA: got %g, want 4", got.Num)
	}

	// All empty
	cf = evalCompile(t, "COUNTA(C1:C3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("COUNTA empty range: got %g, want 0", got.Num)
	}

	// Scalar argument
	cf = evalCompile(t, `COUNTA("hi")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTA scalar: got %g, want 1", got.Num)
	}
}

// ---------------------------------------------------------------------------
// SUMIFS — multiple criteria
// ---------------------------------------------------------------------------

func TestSUMIFS(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Sum range (B)
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: NumberVal(40),
			// Criteria range 1 (A) — category
			{Col: 1, Row: 1}: StringVal("fruit"),
			{Col: 1, Row: 2}: StringVal("veg"),
			{Col: 1, Row: 3}: StringVal("fruit"),
			{Col: 1, Row: 4}: StringVal("veg"),
			// Criteria range 2 (C) — score
			{Col: 3, Row: 1}: NumberVal(5),
			{Col: 3, Row: 2}: NumberVal(15),
			{Col: 3, Row: 3}: NumberVal(25),
			{Col: 3, Row: 4}: NumberVal(35),
		},
	}

	// SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2)
	// Sum B where A="fruit" AND C>10
	cf := evalCompile(t, `SUMIFS(B1:B4,A1:A4,"fruit",C1:C4,">10")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Only row 3 matches (fruit, 25>10) => sum=30
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("SUMIFS: got %g, want 30", got.Num)
	}
}

func TestSUMIFSSingleCriteria(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
		},
	}

	cf := evalCompile(t, `SUMIFS(A1:A3,B1:B3,">=2")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Rows 2,3 match => 20+30=50
	if got.Type != ValueNumber || got.Num != 50 {
		t.Errorf("SUMIFS single: got %g, want 50", got.Num)
	}
}

func TestSUMIFSArgErrors(t *testing.T) {
	resolver := &mockResolver{}

	// Odd number of args (invalid)
	cf := evalCompile(t, "SUMIFS(A1:A3,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("SUMIFS bad args: got %v, want #VALUE!", got)
	}
}

// ---------------------------------------------------------------------------
// COUNTIFS — multiple criteria
// ---------------------------------------------------------------------------

func TestCOUNTIFS(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple"),
			{Col: 1, Row: 2}: StringVal("banana"),
			{Col: 1, Row: 3}: StringVal("apple"),
			{Col: 1, Row: 4}: StringVal("cherry"),
			{Col: 2, Row: 1}: NumberVal(5),
			{Col: 2, Row: 2}: NumberVal(10),
			{Col: 2, Row: 3}: NumberVal(15),
			{Col: 2, Row: 4}: NumberVal(20),
		},
	}

	// Count where A="apple" AND B>10
	cf := evalCompile(t, `COUNTIFS(A1:A4,"apple",B1:B4,">10")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Only row 3 (apple, 15>10) matches
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTIFS: got %g, want 1", got.Num)
	}

	// Single criteria pair
	cf = evalCompile(t, `COUNTIFS(A1:A4,"apple")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIFS single: got %g, want 2", got.Num)
	}
}

// ---------------------------------------------------------------------------
// AVERAGEIF
// ---------------------------------------------------------------------------

func TestAVERAGEIF(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: NumberVal(40),
		},
	}

	cf := evalCompile(t, `AVERAGEIF(A1:A4,">15")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// 20+30+40 = 90, count=3, avg=30
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("AVERAGEIF: got %g, want 30", got.Num)
	}

	// No matches => #DIV/0!
	cf = evalCompile(t, `AVERAGEIF(A1:A4,">100")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("AVERAGEIF no match: got %v, want #DIV/0!", got)
	}
}

func TestAVERAGEIFWithSeparateRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("yes"),
			{Col: 1, Row: 2}: StringVal("no"),
			{Col: 1, Row: 3}: StringVal("yes"),
			{Col: 2, Row: 1}: NumberVal(100),
			{Col: 2, Row: 2}: NumberVal(200),
			{Col: 2, Row: 3}: NumberVal(300),
		},
	}

	cf := evalCompile(t, `AVERAGEIF(A1:A3,"yes",B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// (100+300)/2 = 200
	if got.Type != ValueNumber || got.Num != 200 {
		t.Errorf("AVERAGEIF separate range: got %g, want 200", got.Num)
	}
}

// ---------------------------------------------------------------------------
// AVERAGEIFS
// ---------------------------------------------------------------------------

func TestAVERAGEIFS(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: NumberVal(40),
			{Col: 2, Row: 1}: StringVal("a"),
			{Col: 2, Row: 2}: StringVal("b"),
			{Col: 2, Row: 3}: StringVal("a"),
			{Col: 2, Row: 4}: StringVal("b"),
			{Col: 3, Row: 1}: NumberVal(1),
			{Col: 3, Row: 2}: NumberVal(2),
			{Col: 3, Row: 3}: NumberVal(3),
			{Col: 3, Row: 4}: NumberVal(4),
		},
	}

	// AVERAGEIFS(avg_range, criteria_range1, criteria1, criteria_range2, criteria2)
	// Average of A where B="a" AND C>1 => only row 3 (30), avg=30
	cf := evalCompile(t, `AVERAGEIFS(A1:A4,B1:B4,"a",C1:C4,">1")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("AVERAGEIFS: got %g, want 30", got.Num)
	}

	// No matches => #DIV/0!
	cf = evalCompile(t, `AVERAGEIFS(A1:A4,B1:B4,"z",C1:C4,">0")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("AVERAGEIFS no match: got %v, want #DIV/0!", got)
	}
}

// ---------------------------------------------------------------------------
// SUMIF with separate sum range
// ---------------------------------------------------------------------------

func TestSUMIFWithSeparateRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("yes"),
			{Col: 1, Row: 2}: StringVal("no"),
			{Col: 1, Row: 3}: StringVal("yes"),
			{Col: 2, Row: 1}: NumberVal(100),
			{Col: 2, Row: 2}: NumberVal(200),
			{Col: 2, Row: 3}: NumberVal(300),
		},
	}

	cf := evalCompile(t, `SUMIF(A1:A3,"yes",B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 400 {
		t.Errorf("SUMIF separate range: got %g, want 400", got.Num)
	}
}

// ---------------------------------------------------------------------------
// COUNTIF edge cases
// ---------------------------------------------------------------------------

func TestCOUNTIFWildcard(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple pie"),
			{Col: 1, Row: 2}: StringVal("apple sauce"),
			{Col: 1, Row: 3}: StringVal("banana"),
		},
	}

	cf := evalCompile(t, `COUNTIF(A1:A3,"apple*")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIF wildcard: got %g, want 2", got.Num)
	}
}

func TestCOUNTIFNumericCriteria(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(10),
			{Col: 1, Row: 3}: NumberVal(15),
			{Col: 1, Row: 4}: NumberVal(20),
		},
	}

	// Less-than operator
	cf := evalCompile(t, `COUNTIF(A1:A4,"<15")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIF <15: got %g, want 2", got.Num)
	}

	// Equals with operator
	cf = evalCompile(t, `COUNTIF(A1:A4,"=10")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTIF =10: got %g, want 1", got.Num)
	}

	// Not-equal
	cf = evalCompile(t, `COUNTIF(A1:A4,"<>10")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COUNTIF <>10: got %g, want 3", got.Num)
	}
}

// ---------------------------------------------------------------------------
// AVERAGE / SUM / MIN / MAX with edge cases
// ---------------------------------------------------------------------------

func TestAVERAGEEmpty(t *testing.T) {
	resolver := &mockResolver{}

	// No numeric values => #DIV/0!
	cf := evalCompile(t, "AVERAGE(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("AVERAGE empty: got %v, want #DIV/0!", got)
	}
}

func TestSUMWithMixedTypes(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: StringVal("hello"),
			{Col: 1, Row: 3}: NumberVal(20),
			{Col: 1, Row: 4}: BoolVal(true),
		},
	}

	// In a range, strings and bools are skipped by SUM
	cf := evalCompile(t, "SUM(A1:A4)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("SUM mixed: got %g, want 30", got.Num)
	}
}

func TestMINMAXEmpty(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "MIN(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("MIN empty: got %g, want 0", got.Num)
	}

	cf = evalCompile(t, "MAX(A1:A3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("MAX empty: got %g, want 0", got.Num)
	}
}

func TestMINMAXNegative(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-100),
			{Col: 1, Row: 2}: NumberVal(-50),
			{Col: 1, Row: 3}: NumberVal(-1),
		},
	}

	cf := evalCompile(t, "MIN(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -100 {
		t.Errorf("MIN neg: got %g, want -100", got.Num)
	}

	cf = evalCompile(t, "MAX(A1:A3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -1 {
		t.Errorf("MAX neg: got %g, want -1", got.Num)
	}
}

// ---------------------------------------------------------------------------
// LARGE/SMALL edge cases
// ---------------------------------------------------------------------------

func TestLARGESMALLEdgeCases(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(10),
		},
	}

	// k out of range => #NUM!
	cf := evalCompile(t, "LARGE(A1:A3,0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("LARGE k=0: got %v, want #NUM!", got)
	}

	cf = evalCompile(t, "LARGE(A1:A3,4)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("LARGE k>n: got %v, want #NUM!", got)
	}

	// k=2 with duplicates
	cf = evalCompile(t, "LARGE(A1:A3,2)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("LARGE k=2: got %g, want 5", got.Num)
	}

	cf = evalCompile(t, "SMALL(A1:A3,0)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("SMALL k=0: got %v, want #NUM!", got)
	}
}

// ---------------------------------------------------------------------------
// MEDIAN edge cases
// ---------------------------------------------------------------------------

func TestMEDIANSingleValue(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(42),
		},
	}

	cf := evalCompile(t, "MEDIAN(A1:A1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("MEDIAN single: got %g, want 42", got.Num)
	}
}

// ---------------------------------------------------------------------------
// SUMSQ
// ---------------------------------------------------------------------------

func TestSUMSQ(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    float64
	}{
		{"SUMSQ(3, 4)", 25},
		{"SUMSQ(1, 2, 3)", 14},
		{"SUMSQ(0)", 0},
		{"SUMSQ(-3, 4)", 25},
		{"SUMSQ(5)", 25},
	}
	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != tt.want {
				t.Errorf("%s: got %v, want %g", tt.formula, got, tt.want)
			}
		})
	}
}

func TestSUMSQRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: NumberVal(4),
		},
	}

	cf := evalCompile(t, "SUMSQ(A1:A2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 25 {
		t.Errorf("SUMSQ range: got %g, want 25", got.Num)
	}
}

func TestSUMSQErrorPropagation(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: ErrorVal(ErrValNA),
		},
	}

	cf := evalCompile(t, "SUMSQ(A1:A2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("SUMSQ error propagation: got %v, want #N/A", got)
	}
}

// ---------------------------------------------------------------------------
// PRODUCT
// ---------------------------------------------------------------------------

func TestPRODUCT(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 1, Row: 3}: NumberVal(4),
		},
	}

	cf := evalCompile(t, "PRODUCT(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 24 {
		t.Errorf("PRODUCT: got %g, want 24", got.Num)
	}
}

// ---------------------------------------------------------------------------
// Error propagation in range functions
// ---------------------------------------------------------------------------

func TestSUMErrorInRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: ErrorVal(ErrValNA),
			{Col: 1, Row: 3}: NumberVal(20),
		},
	}

	cf := evalCompile(t, "SUM(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("SUM with error: got %v, want #N/A", got)
	}
}

// ---------------------------------------------------------------------------
// matchesCriteria — extended edge cases
// ---------------------------------------------------------------------------

func TestMatchesCriteriaExtended(t *testing.T) {
	tests := []struct {
		name string
		v    Value
		crit Value
		want bool
	}{
		// Case-insensitive string equality
		{name: "case_insensitive", v: StringVal("Apple"), crit: StringVal("apple"), want: true},
		// Wildcard: ? matches exactly one character
		{name: "question_mark", v: StringVal("bat"), crit: StringVal("b?t"), want: true},
		{name: "question_no_match", v: StringVal("boot"), crit: StringVal("b?t"), want: false},
		// Wildcard: * at end
		{name: "star_end", v: StringVal("hello world"), crit: StringVal("hello*"), want: true},
		// Wildcard: * at start
		{name: "star_start", v: StringVal("hello world"), crit: StringVal("*world"), want: true},
		// Wildcard: * in middle
		{name: "star_middle", v: StringVal("hello world"), crit: StringVal("he*ld"), want: true},
		// Number equality with numeric criteria
		{name: "num_eq_num", v: NumberVal(42), crit: NumberVal(42), want: true},
		{name: "num_ne_num", v: NumberVal(42), crit: NumberVal(43), want: false},
		// String "less than"
		{name: "str_lt", v: NumberVal(3), crit: StringVal("<5"), want: true},
		{name: "str_lt_fail", v: NumberVal(10), crit: StringVal("<5"), want: false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := matchesCriteria(tt.v, tt.crit)
			if got != tt.want {
				t.Errorf("matchesCriteria(%v, %v) = %v, want %v", tt.v, tt.crit, got, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// MAXIFS — maximum value with multiple criteria
// ---------------------------------------------------------------------------

func TestMAXIFS(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Max range (B) — values to take max from
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: NumberVal(40),
			// Criteria range 1 (A) — category
			{Col: 1, Row: 1}: StringVal("fruit"),
			{Col: 1, Row: 2}: StringVal("veg"),
			{Col: 1, Row: 3}: StringVal("fruit"),
			{Col: 1, Row: 4}: StringVal("veg"),
		},
	}

	// MAXIFS with single criteria: max of B where A="fruit"
	cf := evalCompile(t, `MAXIFS(B1:B4,A1:A4,"fruit")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Rows 1 and 3 match (fruit): max(10, 30) = 30
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("MAXIFS single criteria: got %v, want 30", got)
	}
}

func TestMAXIFSNoMatches(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: StringVal("a"),
			{Col: 2, Row: 2}: StringVal("b"),
			{Col: 2, Row: 3}: StringVal("c"),
		},
	}

	// No rows match criteria "z" => return 0
	cf := evalCompile(t, `MAXIFS(A1:A3,B1:B3,"z")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("MAXIFS no matches: got %v, want 0", got)
	}
}

func TestMAXIFSMultipleCriteria(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Max range (A)
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: NumberVal(40),
			// Criteria range 1 (B) — category
			{Col: 2, Row: 1}: StringVal("fruit"),
			{Col: 2, Row: 2}: StringVal("veg"),
			{Col: 2, Row: 3}: StringVal("fruit"),
			{Col: 2, Row: 4}: StringVal("veg"),
			// Criteria range 2 (C) — score
			{Col: 3, Row: 1}: NumberVal(5),
			{Col: 3, Row: 2}: NumberVal(15),
			{Col: 3, Row: 3}: NumberVal(25),
			{Col: 3, Row: 4}: NumberVal(35),
		},
	}

	// MAXIFS(max_range, criteria_range1, criteria1, criteria_range2, criteria2)
	// Max of A where B="fruit" AND C>10
	cf := evalCompile(t, `MAXIFS(A1:A4,B1:B4,"fruit",C1:C4,">10")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Only row 3 matches (fruit, 25>10) => max=30
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("MAXIFS multiple criteria: got %v, want 30", got)
	}
}

func TestMAXIFSArgErrors(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args (only 2 — missing criteria)
	cf := evalCompile(t, "MAXIFS(A1:A3,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("MAXIFS bad args: got %v, want #VALUE!", got)
	}
}

// ---------------------------------------------------------------------------
// MINIFS — minimum value with multiple criteria
// ---------------------------------------------------------------------------

func TestMINIFS(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Min range (B) — values to take min from
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: NumberVal(40),
			// Criteria range 1 (A) — category
			{Col: 1, Row: 1}: StringVal("fruit"),
			{Col: 1, Row: 2}: StringVal("veg"),
			{Col: 1, Row: 3}: StringVal("fruit"),
			{Col: 1, Row: 4}: StringVal("veg"),
		},
	}

	// MINIFS with single criteria: min of B where A="fruit"
	cf := evalCompile(t, `MINIFS(B1:B4,A1:A4,"fruit")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Rows 1 and 3 match (fruit): min(10, 30) = 10
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("MINIFS single criteria: got %v, want 10", got)
	}
}

func TestMINIFSNoMatches(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: StringVal("a"),
			{Col: 2, Row: 2}: StringVal("b"),
			{Col: 2, Row: 3}: StringVal("c"),
		},
	}

	// No rows match criteria "z" => return 0
	cf := evalCompile(t, `MINIFS(A1:A3,B1:B3,"z")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("MINIFS no matches: got %v, want 0", got)
	}
}

// ---------------------------------------------------------------------------
// STDEV — sample standard deviation
// ---------------------------------------------------------------------------

func TestSTDEV(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		tolerance float64
		wantErr   bool
		errVal    ErrorValue
	}{
		{
			name:      "dataset_1",
			formula:   "STDEV(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299)",
			wantNum:   27.46,
			tolerance: 0.01,
		},
		{
			name:      "dataset_2",
			formula:   "STDEV(2, 4, 4, 4, 5, 5, 7, 9)",
			wantNum:   2.138,
			tolerance: 0.001,
		},
		{
			name:    "single_value",
			formula: "STDEV(5)",
			wantErr: true,
			errVal:  ErrValDIV0,
		},
		{
			name:      "all_same",
			formula:   "STDEV(3, 3, 3)",
			wantNum:   0,
			tolerance: 0,
		},
		{
			name:      "two_values",
			formula:   "STDEV(10, 20)",
			wantNum:   7.071,
			tolerance: 0.001,
		},
		{
			name:      "sequential",
			formula:   "STDEV(1, 2, 3, 4, 5)",
			wantNum:   1.5811,
			tolerance: 0.001,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError || got.Err != tt.errVal {
					t.Errorf("%s: got %v, want error %v", tt.formula, got, tt.errVal)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: got type %v, want number", tt.formula, got.Type)
			}
			if tt.tolerance == 0 {
				if got.Num != tt.wantNum {
					t.Errorf("%s: got %g, want %g", tt.formula, got.Num, tt.wantNum)
				}
			} else if math.Abs(got.Num-tt.wantNum) > tt.tolerance {
				t.Errorf("%s: got %g, want %g (tolerance %g)", tt.formula, got.Num, tt.wantNum, tt.tolerance)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// STDEVP — population standard deviation
// ---------------------------------------------------------------------------

func TestSTDEVP(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		tolerance float64
		wantErr   bool
		errVal    ErrorValue
	}{
		{
			name:      "dataset_1",
			formula:   "STDEVP(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299)",
			wantNum:   26.05,
			tolerance: 0.01,
		},
		{
			name:      "dataset_2",
			formula:   "STDEVP(2, 4, 4, 4, 5, 5, 7, 9)",
			wantNum:   2.0,
			tolerance: 0.001,
		},
		{
			name:      "single_value",
			formula:   "STDEVP(5)",
			wantNum:   0,
			tolerance: 0,
		},
		{
			name:      "all_same",
			formula:   "STDEVP(3, 3, 3)",
			wantNum:   0,
			tolerance: 0,
		},
		{
			name:      "two_values",
			formula:   "STDEVP(10, 20)",
			wantNum:   5.0,
			tolerance: 0.001,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError || got.Err != tt.errVal {
					t.Errorf("%s: got %v, want error %v", tt.formula, got, tt.errVal)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: got type %v, want number", tt.formula, got.Type)
			}
			if tt.tolerance == 0 {
				if got.Num != tt.wantNum {
					t.Errorf("%s: got %g, want %g", tt.formula, got.Num, tt.wantNum)
				}
			} else if math.Abs(got.Num-tt.wantNum) > tt.tolerance {
				t.Errorf("%s: got %g, want %g (tolerance %g)", tt.formula, got.Num, tt.wantNum, tt.tolerance)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// VAR — sample variance
// ---------------------------------------------------------------------------

func TestVAR(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		tolerance float64
		wantErr   bool
		errVal    ErrorValue
	}{
		{
			name:      "dataset_1",
			formula:   "VAR(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299)",
			wantNum:   754.27,
			tolerance: 0.01,
		},
		{
			name:      "dataset_2",
			formula:   "VAR(2, 4, 4, 4, 5, 5, 7, 9)",
			wantNum:   4.571,
			tolerance: 0.001,
		},
		{
			name:    "single_value",
			formula: "VAR(5)",
			wantErr: true,
			errVal:  ErrValDIV0,
		},
		{
			name:      "all_same",
			formula:   "VAR(3, 3, 3)",
			wantNum:   0,
			tolerance: 0,
		},
		{
			name:      "two_values",
			formula:   "VAR(10, 20)",
			wantNum:   50.0,
			tolerance: 0.001,
		},
		{
			name:      "sequential",
			formula:   "VAR(1, 2, 3, 4, 5)",
			wantNum:   2.5,
			tolerance: 0.001,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError || got.Err != tt.errVal {
					t.Errorf("%s: got %v, want error %v", tt.formula, got, tt.errVal)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: got type %v, want number", tt.formula, got.Type)
			}
			if tt.tolerance == 0 {
				if got.Num != tt.wantNum {
					t.Errorf("%s: got %g, want %g", tt.formula, got.Num, tt.wantNum)
				}
			} else if math.Abs(got.Num-tt.wantNum) > tt.tolerance {
				t.Errorf("%s: got %g, want %g (tolerance %g)", tt.formula, got.Num, tt.wantNum, tt.tolerance)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// MODE — most frequently occurring value
// ---------------------------------------------------------------------------

func TestMODE(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		wantNA  bool
	}{
		{name: "basic", formula: "MODE(5.6, 4, 4, 3, 2, 4)", wantNum: 4},
		{name: "no_duplicates", formula: "MODE(1, 2, 3, 4, 5)", wantNA: true},
		{name: "tie_first_wins", formula: "MODE(1, 2, 2, 3, 3)", wantNum: 2},
		{name: "all_same", formula: "MODE(7, 7, 7)", wantNum: 7},
		{name: "floats", formula: "MODE(1.5, 1.5, 2.5, 2.5, 2.5)", wantNum: 2.5},
		{name: "clear_winner", formula: "MODE(1, 1, 2, 2, 3, 3, 3)", wantNum: 3},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantNA {
				if got.Type != ValueError || got.Err != ErrValNA {
					t.Errorf("%s: got %v, want #N/A", tt.formula, got)
				}
				return
			}
			if got.Type != ValueNumber || got.Num != tt.wantNum {
				t.Errorf("%s: got %v, want %g", tt.formula, got, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// VARP — population variance
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// PERCENTILE — k-th percentile using linear interpolation
// ---------------------------------------------------------------------------

func TestPERCENTILE(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		tolerance float64
		wantErr   bool
		errVal    ErrorValue
	}{
		{
			name:      "basic_0.3",
			formula:   "PERCENTILE({1, 2, 3, 4}, 0.3)",
			wantNum:   1.9,
			tolerance: 1e-9,
		},
		{
			name:      "minimum",
			formula:   "PERCENTILE({1, 2, 3, 4}, 0)",
			wantNum:   1,
			tolerance: 0,
		},
		{
			name:      "maximum",
			formula:   "PERCENTILE({1, 2, 3, 4}, 1)",
			wantNum:   4,
			tolerance: 0,
		},
		{
			name:      "median",
			formula:   "PERCENTILE({1, 2, 3, 4}, 0.5)",
			wantNum:   2.5,
			tolerance: 1e-9,
		},
		{
			name:    "k_negative",
			formula: "PERCENTILE({1, 2, 3, 4}, -0.1)",
			wantErr: true,
			errVal:  ErrValNUM,
		},
		{
			name:    "k_too_large",
			formula: "PERCENTILE({1, 2, 3, 4}, 1.1)",
			wantErr: true,
			errVal:  ErrValNUM,
		},
		{
			name:      "single_element",
			formula:   "PERCENTILE({10}, 0.5)",
			wantNum:   10,
			tolerance: 0,
		},
		{
			name:      "five_elements_0.25",
			formula:   "PERCENTILE({1, 3, 5, 7, 9}, 0.25)",
			wantNum:   3,
			tolerance: 1e-9,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError || got.Err != tt.errVal {
					t.Errorf("%s: got %v, want error %v", tt.formula, got, tt.errVal)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: got type %v, want number", tt.formula, got.Type)
			}
			if tt.tolerance == 0 {
				if got.Num != tt.wantNum {
					t.Errorf("%s: got %g, want %g", tt.formula, got.Num, tt.wantNum)
				}
			} else if math.Abs(got.Num-tt.wantNum) > tt.tolerance {
				t.Errorf("%s: got %g, want %g (tolerance %g)", tt.formula, got.Num, tt.wantNum, tt.tolerance)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// RANK — rank of a number in a list
// ---------------------------------------------------------------------------

func TestRANK(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		wantErr bool
		errVal  ErrorValue
	}{
		{
			name:    "descending_default",
			formula: "RANK(3.5, {7, 3.5, 3.5, 1, 2})",
			wantNum: 2,
		},
		{
			name:    "descending_largest",
			formula: "RANK(7, {7, 3.5, 3.5, 1, 2})",
			wantNum: 1,
		},
		{
			name:    "descending_smallest",
			formula: "RANK(1, {7, 3.5, 3.5, 1, 2})",
			wantNum: 5,
		},
		{
			name:    "ascending_smallest",
			formula: "RANK(1, {7, 3.5, 3.5, 1, 2}, 1)",
			wantNum: 1,
		},
		{
			name:    "ascending_largest",
			formula: "RANK(7, {7, 3.5, 3.5, 1, 2}, 1)",
			wantNum: 5,
		},
		{
			name:    "not_in_list",
			formula: "RANK(99, {1, 2, 3})",
			wantErr: true,
			errVal:  ErrValNA,
		},
		{
			name:    "explicit_descending",
			formula: "RANK(2, {7, 3.5, 3.5, 1, 2}, 0)",
			wantNum: 4,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError || got.Err != tt.errVal {
					t.Errorf("%s: got %v, want error %v", tt.formula, got, tt.errVal)
				}
				return
			}
			if got.Type != ValueNumber || got.Num != tt.wantNum {
				t.Errorf("%s: got %v, want %g", tt.formula, got, tt.wantNum)
			}
		})
	}
}

func TestVARP(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		tolerance float64
		wantErr   bool
		errVal    ErrorValue
	}{
		{
			name:      "dataset_1",
			formula:   "VARP(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299)",
			wantNum:   678.84,
			tolerance: 0.01,
		},
		{
			name:      "dataset_2",
			formula:   "VARP(2, 4, 4, 4, 5, 5, 7, 9)",
			wantNum:   4.0,
			tolerance: 0.001,
		},
		{
			name:      "single_value",
			formula:   "VARP(5)",
			wantNum:   0,
			tolerance: 0,
		},
		{
			name:      "all_same",
			formula:   "VARP(3, 3, 3)",
			wantNum:   0,
			tolerance: 0,
		},
		{
			name:      "two_values",
			formula:   "VARP(10, 20)",
			wantNum:   25.0,
			tolerance: 0.001,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError || got.Err != tt.errVal {
					t.Errorf("%s: got %v, want error %v", tt.formula, got, tt.errVal)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: got type %v, want number", tt.formula, got.Type)
			}
			if tt.tolerance == 0 {
				if got.Num != tt.wantNum {
					t.Errorf("%s: got %g, want %g", tt.formula, got.Num, tt.wantNum)
				}
			} else if math.Abs(got.Num-tt.wantNum) > tt.tolerance {
				t.Errorf("%s: got %g, want %g (tolerance %g)", tt.formula, got.Num, tt.wantNum, tt.tolerance)
			}
		})
	}
}
