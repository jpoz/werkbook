package formula

import (
	"math"
	"testing"
)

// ---------------------------------------------------------------------------
// LARGE / SMALL
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// COUNTBLANK
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// SUMIF
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// COUNTIF
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// SUMPRODUCT
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// MatchesCriteria — helper used by *IF functions
// ---------------------------------------------------------------------------

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
		got := MatchesCriteria(tt.v, tt.crit)
		if got != tt.want {
			t.Errorf("MatchesCriteria(%v, %v) = %v, want %v", tt.v, tt.crit, got, tt.want)
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
// COUNTIF with mixed positive/negative/zero values
// ---------------------------------------------------------------------------

func TestCOUNTIFMixedSignValues(t *testing.T) {
	// Mirrors the multisheet edge case spec: A1:A5 = [10, -5, 0, 100, 25]
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(-5),
			{Col: 1, Row: 3}: NumberVal(0),
			{Col: 1, Row: 4}: NumberVal(100),
			{Col: 1, Row: 5}: NumberVal(25),
		},
	}

	// >0 should count only strictly positive values (10, 100, 25) => 3
	cf := evalCompile(t, `COUNTIF(A1:A5,">0")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COUNTIF >0 mixed: got %g, want 3", got.Num)
	}

	// <0 should count only negative values (-5) => 1
	cf = evalCompile(t, `COUNTIF(A1:A5,"<0")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTIF <0 mixed: got %g, want 1", got.Num)
	}

	// =0 should count only zero values => 1
	cf = evalCompile(t, `COUNTIF(A1:A5,"=0")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTIF =0 mixed: got %g, want 1", got.Num)
	}

	// >=0 should count zero and positives (10, 0, 100, 25) => 4
	cf = evalCompile(t, `COUNTIF(A1:A5,">=0")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("COUNTIF >=0 mixed: got %g, want 4", got.Num)
	}
}

// ---------------------------------------------------------------------------
// SUMIF with mixed positive/negative/zero values
// ---------------------------------------------------------------------------

func TestSUMIFMixedSignValues(t *testing.T) {
	// Mirrors the multisheet edge case spec: A1:A5 = [10, -5, 0, 100, 25]
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(-5),
			{Col: 1, Row: 3}: NumberVal(0),
			{Col: 1, Row: 4}: NumberVal(100),
			{Col: 1, Row: 5}: NumberVal(25),
		},
	}

	// >0 should sum only strictly positive values (10+100+25) => 135
	cf := evalCompile(t, `SUMIF(A1:A5,">0")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 135 {
		t.Errorf("SUMIF >0 mixed: got %g, want 135", got.Num)
	}

	// <0 should sum only negative values (-5) => -5
	cf = evalCompile(t, `SUMIF(A1:A5,"<0")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -5 {
		t.Errorf("SUMIF <0 mixed: got %g, want -5", got.Num)
	}

	// No matches => 0
	cf = evalCompile(t, `SUMIF(A1:A5,">1000")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("SUMIF no match: got %g, want 0", got.Num)
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
// MatchesCriteria — extended edge cases
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
			got := MatchesCriteria(tt.v, tt.crit)
			if got != tt.want {
				t.Errorf("MatchesCriteria(%v, %v) = %v, want %v", tt.v, tt.crit, got, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// QUARTILE
// ---------------------------------------------------------------------------

func TestQUARTILE(t *testing.T) {
	// Standard dataset from Excel docs: {1,2,4,7,8,9,10,12}
	stdResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(4),
			{Col: 1, Row: 4}: NumberVal(7),
			{Col: 1, Row: 5}: NumberVal(8),
			{Col: 1, Row: 6}: NumberVal(9),
			{Col: 1, Row: 7}: NumberVal(10),
			{Col: 1, Row: 8}: NumberVal(12),
		},
	}

	// Single element in B1
	singleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(42),
		},
	}

	// Two elements in C1:C2
	twoResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(5),
			{Col: 3, Row: 2}: NumberVal(15),
		},
	}

	// Mixed types: numbers, strings, booleans in D1:D5
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(10),
			{Col: 4, Row: 2}: StringVal("hello"),
			{Col: 4, Row: 3}: NumberVal(20),
			{Col: 4, Row: 4}: BoolVal(true),
			{Col: 4, Row: 5}: NumberVal(30),
		},
	}

	// Negative numbers in E1:E4
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(-10),
			{Col: 5, Row: 2}: NumberVal(-5),
			{Col: 5, Row: 3}: NumberVal(0),
			{Col: 5, Row: 4}: NumberVal(5),
		},
	}

	// All same values in F1:F4
	sameResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 6, Row: 1}: NumberVal(7),
			{Col: 6, Row: 2}: NumberVal(7),
			{Col: 6, Row: 3}: NumberVal(7),
			{Col: 6, Row: 4}: NumberVal(7),
		},
	}

	// Empty range
	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	// Large dataset in G1:G20
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 20; i++ {
		largeResolver.cells[CellAddr{Col: 7, Row: i}] = NumberVal(float64(i))
	}

	tests := []struct {
		name     string
		formula  string
		resolver *mockResolver
		wantNum  float64
		wantErr  ErrorValue
	}{
		// quart=0 (minimum)
		{"q0_min", "QUARTILE(A1:A8,0)", stdResolver, 1, 0},
		// quart=1 (25th percentile) - Excel example
		{"q1_25th", "QUARTILE(A1:A8,1)", stdResolver, 3.5, 0},
		// quart=2 (median)
		{"q2_median", "QUARTILE(A1:A8,2)", stdResolver, 7.5, 0},
		// quart=3 (75th percentile)
		{"q3_75th", "QUARTILE(A1:A8,3)", stdResolver, 9.25, 0},
		// quart=4 (maximum)
		{"q4_max", "QUARTILE(A1:A8,4)", stdResolver, 12, 0},
		// quart as float truncated to integer
		{"q_float_1.7_truncates_to_1", "QUARTILE(A1:A8,1.7)", stdResolver, 3.5, 0},
		{"q_float_3.9_truncates_to_3", "QUARTILE(A1:A8,3.9)", stdResolver, 9.25, 0},
		// quart < 0 → #NUM!
		{"q_negative", "QUARTILE(A1:A8,-1)", stdResolver, 0, ErrValNUM},
		// quart > 4 → #NUM!
		{"q_over_4", "QUARTILE(A1:A8,5)", stdResolver, 0, ErrValNUM},
		// Empty array → #NUM!
		{"empty_array", "QUARTILE(Z1:Z3,1)", emptyResolver, 0, ErrValNUM},
		// Single element
		{"single_q0", "QUARTILE(B1:B1,0)", singleResolver, 42, 0},
		{"single_q1", "QUARTILE(B1:B1,1)", singleResolver, 42, 0},
		{"single_q2", "QUARTILE(B1:B1,2)", singleResolver, 42, 0},
		{"single_q3", "QUARTILE(B1:B1,3)", singleResolver, 42, 0},
		{"single_q4", "QUARTILE(B1:B1,4)", singleResolver, 42, 0},
		// Two element array
		{"two_q0", "QUARTILE(C1:C2,0)", twoResolver, 5, 0},
		{"two_q1", "QUARTILE(C1:C2,1)", twoResolver, 7.5, 0},
		{"two_q2", "QUARTILE(C1:C2,2)", twoResolver, 10, 0},
		{"two_q3", "QUARTILE(C1:C2,3)", twoResolver, 12.5, 0},
		{"two_q4", "QUARTILE(C1:C2,4)", twoResolver, 15, 0},
		// Mixed types (non-numeric ignored) → only 10, 20, 30
		{"mixed_q1", "QUARTILE(D1:D5,1)", mixedResolver, 15, 0},
		{"mixed_q2", "QUARTILE(D1:D5,2)", mixedResolver, 20, 0},
		// Negative numbers
		{"neg_q0", "QUARTILE(E1:E4,0)", negResolver, -10, 0},
		{"neg_q1", "QUARTILE(E1:E4,1)", negResolver, -6.25, 0},
		{"neg_q2", "QUARTILE(E1:E4,2)", negResolver, -2.5, 0},
		{"neg_q4", "QUARTILE(E1:E4,4)", negResolver, 5, 0},
		// All same values
		{"same_q0", "QUARTILE(F1:F4,0)", sameResolver, 7, 0},
		{"same_q1", "QUARTILE(F1:F4,1)", sameResolver, 7, 0},
		{"same_q2", "QUARTILE(F1:F4,2)", sameResolver, 7, 0},
		{"same_q4", "QUARTILE(F1:F4,4)", sameResolver, 7, 0},
		// Large dataset
		{"large_q1", "QUARTILE(G1:G20,1)", largeResolver, 5.75, 0},
		{"large_q2", "QUARTILE(G1:G20,2)", largeResolver, 10.5, 0},
		{"large_q3", "QUARTILE(G1:G20,3)", largeResolver, 15.25, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantErr != 0 {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %d (%v), want number", got.Type, got)
			}
			if math.Abs(got.Num-tt.wantNum) > 1e-9 {
				t.Errorf("got %g, want %g", got.Num, tt.wantNum)
			}
		})
	}

	// QUARTILE(x,q) == PERCENTILE(x, q*0.25) equivalence check
	t.Run("equivalence_with_PERCENTILE", func(t *testing.T) {
		for q := 0; q <= 4; q++ {
			qf := evalCompile(t, "QUARTILE(A1:A8,"+string(rune('0'+q))+")")
			qv, err := Eval(qf, stdResolver, nil)
			if err != nil {
				t.Fatalf("Eval QUARTILE q=%d: %v", q, err)
			}

			pctStr := []string{"0", "0.25", "0.5", "0.75", "1"}[q]
			pf := evalCompile(t, "PERCENTILE(A1:A8,"+pctStr+")")
			pv, err := Eval(pf, stdResolver, nil)
			if err != nil {
				t.Fatalf("Eval PERCENTILE q=%d: %v", q, err)
			}

			if math.Abs(qv.Num-pv.Num) > 1e-9 {
				t.Errorf("QUARTILE(q=%d)=%g != PERCENTILE(k=%s)=%g", q, qv.Num, pctStr, pv.Num)
			}
		}
	})

	// Wrong number of arguments
	t.Run("too_few_args", func(t *testing.T) {
		cf := evalCompile(t, "QUARTILE(A1:A8)")
		got, err := Eval(cf, stdResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})
}
