package formula

import (
	"fmt"
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

// ---------------------------------------------------------------------------
// TRIMMEAN
// ---------------------------------------------------------------------------

func TestTRIMMEAN(t *testing.T) {
	// Excel docs example: {4,5,6,7,2,3,4,5,1,2,3} in A1:A11
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}:  NumberVal(4),
			{Col: 1, Row: 2}:  NumberVal(5),
			{Col: 1, Row: 3}:  NumberVal(6),
			{Col: 1, Row: 4}:  NumberVal(7),
			{Col: 1, Row: 5}:  NumberVal(2),
			{Col: 1, Row: 6}:  NumberVal(3),
			{Col: 1, Row: 7}:  NumberVal(4),
			{Col: 1, Row: 8}:  NumberVal(5),
			{Col: 1, Row: 9}:  NumberVal(1),
			{Col: 1, Row: 10}: NumberVal(2),
			{Col: 1, Row: 11}: NumberVal(3),
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

	// All same values in D1:D4
	sameResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(7),
			{Col: 4, Row: 2}: NumberVal(7),
			{Col: 4, Row: 3}: NumberVal(7),
			{Col: 4, Row: 4}: NumberVal(7),
		},
	}

	// Negative numbers in E1:E5
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(-10),
			{Col: 5, Row: 2}: NumberVal(-5),
			{Col: 5, Row: 3}: NumberVal(0),
			{Col: 5, Row: 4}: NumberVal(5),
			{Col: 5, Row: 5}: NumberVal(10),
		},
	}

	// Mixed types in F1:F5 (strings and booleans ignored)
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 6, Row: 1}: NumberVal(10),
			{Col: 6, Row: 2}: StringVal("hello"),
			{Col: 6, Row: 3}: NumberVal(20),
			{Col: 6, Row: 4}: BoolVal(true),
			{Col: 6, Row: 5}: NumberVal(30),
		},
	}

	// Empty range
	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	// Large dataset in G1:G30
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 30; i++ {
		largeResolver.cells[CellAddr{Col: 7, Row: i}] = NumberVal(float64(i))
	}

	// Four elements in H1:H4
	fourResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 8, Row: 1}: NumberVal(1),
			{Col: 8, Row: 2}: NumberVal(2),
			{Col: 8, Row: 3}: NumberVal(3),
			{Col: 8, Row: 4}: NumberVal(4),
		},
	}

	// Six elements in I1:I6
	sixResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 9, Row: 1}: NumberVal(1),
			{Col: 9, Row: 2}: NumberVal(2),
			{Col: 9, Row: 3}: NumberVal(3),
			{Col: 9, Row: 4}: NumberVal(4),
			{Col: 9, Row: 5}: NumberVal(5),
			{Col: 9, Row: 6}: NumberVal(6),
		},
	}

	// Ten elements in J1:J10
	tenResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 10; i++ {
		tenResolver.cells[CellAddr{Col: 10, Row: i}] = NumberVal(float64(i * 10))
	}

	tests := []struct {
		name     string
		formula  string
		resolver *mockResolver
		wantNum  float64
		wantErr  ErrorValue
	}{
		// Excel docs example: {4,5,6,7,2,3,4,5,1,2,3}, 0.2
		// sorted: 1,2,2,3,3,4,4,5,5,6,7; n=11, floor(11*0.2/2)=1, trim 1 each end
		// remaining: 2,2,3,3,4,4,5,5,6 → mean = 34/9
		{"excel_example", "TRIMMEAN(A1:A11,0.2)", excelResolver, 34.0 / 9.0, 0},

		// percent=0 → regular mean
		{"percent_zero", "TRIMMEAN(A1:A11,0)", excelResolver, 42.0 / 11.0, 0},

		// percent=0.5 on 4 elements: floor(4*0.5/2)=1, trim 1 each end → {2,3} mean=2.5
		{"percent_half_4elem", "TRIMMEAN(H1:H4,0.5)", fourResolver, 2.5, 0},

		// percent just under 1 (0.99) on 11 elements: floor(11*0.99/2)=5, trim 5 each end → 1 left
		{"percent_099", "TRIMMEAN(A1:A11,0.99)", excelResolver, 4, 0},

		// percent < 0 → #NUM!
		{"percent_negative", "TRIMMEAN(A1:A11,-0.1)", excelResolver, 0, ErrValNUM},

		// percent >= 1 → #NUM!
		{"percent_one", "TRIMMEAN(A1:A11,1)", excelResolver, 0, ErrValNUM},
		{"percent_over_one", "TRIMMEAN(A1:A11,1.5)", excelResolver, 0, ErrValNUM},

		// Single element, percent=0
		{"single_pct0", "TRIMMEAN(B1:B1,0)", singleResolver, 42, 0},

		// Single element, percent=0.5: floor(1*0.5/2)=0, no trim → mean=42
		{"single_pct05", "TRIMMEAN(B1:B1,0.5)", singleResolver, 42, 0},

		// Two elements, percent=0.5: floor(2*0.5/2)=0, no trim → mean=10
		{"two_pct05", "TRIMMEAN(C1:C2,0.5)", twoResolver, 10, 0},

		// Two elements, percent=0.99: floor(2*0.99/2)=0, no trim → mean=10
		{"two_pct099", "TRIMMEAN(C1:C2,0.99)", twoResolver, 10, 0},

		// All same values
		{"all_same", "TRIMMEAN(D1:D4,0.5)", sameResolver, 7, 0},

		// Negative numbers: {-10,-5,0,5,10}, percent=0.4: floor(5*0.4/2)=1, trim 1 each → {-5,0,5} mean=0
		{"negative_nums", "TRIMMEAN(E1:E5,0.4)", negResolver, 0, 0},

		// Mixed types: only {10,20,30} numeric, percent=0: mean=20
		{"mixed_types_pct0", "TRIMMEAN(F1:F5,0)", mixedResolver, 20, 0},

		// Mixed types: {10,20,30}, percent=0.5: floor(3*0.5/2)=0, no trim → mean=20
		{"mixed_types_pct05", "TRIMMEAN(F1:F5,0.5)", mixedResolver, 20, 0},

		// Empty array → #NUM!
		{"empty_array", "TRIMMEAN(Z1:Z3,0.2)", emptyResolver, 0, ErrValNUM},

		// Large dataset 1..30, percent=0.1: floor(30*0.1/2)=1, trim 1 each → {2..29} mean=15.5
		{"large_pct01", "TRIMMEAN(G1:G30,0.1)", largeResolver, 15.5, 0},

		// Large dataset 1..30, percent=0.2: floor(30*0.2/2)=3, trim 3 each → {4..27} mean=15.5
		{"large_pct02", "TRIMMEAN(G1:G30,0.2)", largeResolver, 15.5, 0},

		// Six elements, percent=0.5: floor(6*0.5/2)=1, trim 1 each → {2,3,4,5} mean=3.5
		{"six_pct05", "TRIMMEAN(I1:I6,0.5)", sixResolver, 3.5, 0},

		// Ten elements {10,20..100}, percent=0.3: floor(10*0.3/2)=1, trim 1 each → {20..90} mean=55
		{"ten_pct03", "TRIMMEAN(J1:J10,0.3)", tenResolver, 55, 0},

		// Four elements, percent=0.99: floor(4*0.99/2)=1, trim 1 each → {2,3} mean=2.5
		{"four_pct099", "TRIMMEAN(H1:H4,0.99)", fourResolver, 2.5, 0},
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

	// Wrong number of arguments
	t.Run("too_few_args", func(t *testing.T) {
		cf := evalCompile(t, "TRIMMEAN(A1:A11)")
		got, err := Eval(cf, excelResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})

	t.Run("too_many_args", func(t *testing.T) {
		cf := evalCompile(t, "TRIMMEAN(A1:A11,0.2,1)")
		got, err := Eval(cf, excelResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})
}

// ---------------------------------------------------------------------------
// HARMEAN
// ---------------------------------------------------------------------------

func TestHARMEAN(t *testing.T) {
	const tol = 1e-6

	// Resolver with {4,5,8,7,11,4,3} in A1:A7
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(4),
			{Col: 1, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(8),
			{Col: 1, Row: 4}: NumberVal(7),
			{Col: 1, Row: 5}: NumberVal(11),
			{Col: 1, Row: 6}: NumberVal(4),
			{Col: 1, Row: 7}: NumberVal(3),
		},
	}

	// Resolver with mixed types in B column
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 2, Row: 2}: StringVal("hello"),
			{Col: 2, Row: 3}: NumberVal(4),
			{Col: 2, Row: 4}: BoolVal(true),
			{Col: 2, Row: 5}: NumberVal(8),
		},
	}

	// Resolver with zero in C column
	zeroResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(5),
			{Col: 3, Row: 2}: NumberVal(0),
			{Col: 3, Row: 3}: NumberVal(10),
		},
	}

	// Resolver with error in D column
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(3),
			{Col: 4, Row: 2}: ErrorVal(ErrValNA),
			{Col: 4, Row: 3}: NumberVal(6),
		},
	}

	// Large dataset in E column (1..20)
	largeCells := map[CellAddr]Value{}
	for i := 1; i <= 20; i++ {
		largeCells[CellAddr{Col: 5, Row: i}] = NumberVal(float64(i))
	}
	largeResolver := &mockResolver{cells: largeCells}

	tests := []struct {
		name     string
		formula  string
		resolver CellResolver
		wantNum  float64
		wantErr  bool // expect ValueError type
	}{
		// Excel documentation example: {4,5,8,7,11,4,3}
		{
			name:     "excel_example_array_ref",
			formula:  "HARMEAN(A1:A7)",
			resolver: excelResolver,
			wantNum:  5.028376,
		},
		// Same values as direct args
		{
			name:     "excel_example_direct",
			formula:  "HARMEAN(4,5,8,7,11,4,3)",
			resolver: nil,
			wantNum:  5.028376,
		},
		// Single value
		{
			name:     "single_value",
			formula:  "HARMEAN(4)",
			resolver: nil,
			wantNum:  4,
		},
		// Two values: 2/(1/1+1/4) = 2/1.25 = 1.6
		{
			name:     "two_values",
			formula:  "HARMEAN(1,4)",
			resolver: nil,
			wantNum:  1.6,
		},
		// All same values
		{
			name:     "all_same",
			formula:  "HARMEAN(5,5,5)",
			resolver: nil,
			wantNum:  5,
		},
		// Zero returns #NUM!
		{
			name:    "zero_direct",
			formula: "HARMEAN(0)",
			wantErr: true,
		},
		// Zero in middle
		{
			name:    "zero_in_middle",
			formula: "HARMEAN(3,0,6)",
			wantErr: true,
		},
		// Negative value returns #NUM!
		{
			name:    "negative_value",
			formula: "HARMEAN(-1,2,3)",
			wantErr: true,
		},
		// Boolean TRUE as direct arg (counted as 1)
		{
			name:    "bool_true_direct",
			formula: "HARMEAN(TRUE,4)",
			wantNum: 1.6, // 2/(1/1+1/4) = 1.6
		},
		// String number as direct arg
		{
			name:    "string_number_direct",
			formula: `HARMEAN("3",6)`,
			wantNum: 4, // 2/(1/3+1/6) = 2/0.5 = 4
		},
		// Array with mixed types: text and bool ignored, only numbers {2,4,8}
		{
			name:     "array_mixed_types",
			formula:  "HARMEAN(B1:B5)",
			resolver: mixedResolver,
			wantNum:  3.428571, // 3/(1/2+1/4+1/8) = 3/0.875
		},
		// Zero in array → #NUM!
		{
			name:     "zero_in_array",
			formula:  "HARMEAN(C1:C3)",
			resolver: zeroResolver,
			wantErr:  true,
		},
		// Error propagation (#N/A in array)
		{
			name:     "error_propagation_na",
			formula:  "HARMEAN(D1:D3)",
			resolver: errResolver,
			wantErr:  true,
		},
		// Error propagation with #VALUE!
		{
			name:    "error_propagation_value",
			formula: `HARMEAN(1,2,1/0)`,
			wantErr: true,
		},
		// Large dataset 1..20
		{
			name:     "large_dataset",
			formula:  "HARMEAN(E1:E20)",
			resolver: largeResolver,
			wantNum:  5.559046, // harmonic mean of 1..20
		},
		// Very small positive numbers
		{
			name:    "very_small_numbers",
			formula: "HARMEAN(0.001,0.002,0.003)",
			wantNum: 0.001636, // 3/(1/0.001+1/0.002+1/0.003)
		},
		// Very large numbers
		{
			name:    "very_large_numbers",
			formula: "HARMEAN(1000000,2000000,3000000)",
			wantNum: 1636363.636364,
		},
		// Single element array
		{
			name:     "single_element_array",
			formula:  "HARMEAN(A1:A1)",
			resolver: excelResolver,
			wantNum:  4,
		},
		// Multiple array arguments
		{
			name:     "multiple_arrays",
			formula:  "HARMEAN(A1:A3,A4:A7)",
			resolver: excelResolver,
			wantNum:  5.028376,
		},
		// Mix of direct and array
		{
			name:     "direct_and_array",
			formula:  "HARMEAN(2,A1:A1)",
			resolver: excelResolver,
			wantNum:  2.666667, // 2/(1/2+1/4) = 2/0.75
		},
		// Equal fractions
		{
			name:    "fractions",
			formula: "HARMEAN(0.5,0.5)",
			wantNum: 0.5,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			resolver := tt.resolver
			if resolver == nil {
				resolver = &mockResolver{}
			}
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr {
				if got.Type != ValueError {
					t.Errorf("got %v (type %d), want error", got, got.Type)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %d (%v), want number", got.Type, got)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// CORREL
// ---------------------------------------------------------------------------

func TestCORREL(t *testing.T) {
	// Basic data: A1:A5 = {3,2,4,5,6}, B1:B5 = {9,7,12,15,17}
	basicResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(4),
			{Col: 1, Row: 4}: NumberVal(5),
			{Col: 1, Row: 5}: NumberVal(6),
			{Col: 2, Row: 1}: NumberVal(9),
			{Col: 2, Row: 2}: NumberVal(7),
			{Col: 2, Row: 3}: NumberVal(12),
			{Col: 2, Row: 4}: NumberVal(15),
			{Col: 2, Row: 5}: NumberVal(17),
		},
	}

	// Perfect positive: A1:A3 = {1,2,3}, B1:B3 = {2,4,6}
	perfectPosResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}

	// Perfect negative: A1:A3 = {1,2,3}, B1:B3 = {6,4,2}
	perfectNegResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(6),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 2, Row: 3}: NumberVal(2),
		},
	}

	// Zero std dev: A1:A3 = {5,5,5}, B1:B3 = {1,2,3}
	zeroStdDevResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(5),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Two pairs: A1:A2 = {1,2}, B1:B2 = {3,5}
	twoPairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 2, Row: 1}: NumberVal(3),
			{Col: 2, Row: 2}: NumberVal(5),
		},
	}

	// Mixed types: A1:A4 = {1,"hello",3,4}, B1:B4 = {10,20,30,"world"}
	// Only pairs (1,10) and (3,30) are numeric in both
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: StringVal("hello"),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(4),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: StringVal("world"),
		},
	}

	// Negative numbers: A1:A4 = {-3,-1,1,3}, B1:B4 = {-9,-3,3,9}
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-3),
			{Col: 1, Row: 2}: NumberVal(-1),
			{Col: 1, Row: 3}: NumberVal(1),
			{Col: 1, Row: 4}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(-9),
			{Col: 2, Row: 2}: NumberVal(-3),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 2, Row: 4}: NumberVal(9),
		},
	}

	// All zeros in one array: A1:A3 = {0,0,0}, B1:B3 = {1,2,3}
	allZerosResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0),
			{Col: 1, Row: 2}: NumberVal(0),
			{Col: 1, Row: 3}: NumberVal(0),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Single pair: A1 = {10}, B1 = {20}
	singlePairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
		},
	}

	// Empty range
	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	// Zeros included: A1:A3 = {0,1,2}, B1:B3 = {0,2,4}
	zerosIncludedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0),
			{Col: 1, Row: 2}: NumberVal(1),
			{Col: 1, Row: 3}: NumberVal(2),
			{Col: 2, Row: 1}: NumberVal(0),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(4),
		},
	}

	// Bool in array: A1:A3 = {1,TRUE,3}, B1:B3 = {10,20,30}
	// TRUE is bool, not numeric in array context; pairs (1,10) and (3,30) only
	boolResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: BoolVal(true),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
		},
	}

	// Large dataset: A1:A20 = {1..20}, B1:B20 = {2..40 step 2}
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 20; i++ {
		largeResolver.cells[CellAddr{Col: 1, Row: i}] = NumberVal(float64(i))
		largeResolver.cells[CellAddr{Col: 2, Row: i}] = NumberVal(float64(i * 2))
	}

	// Weak correlation: A1:A5 = {1,2,3,4,5}, B1:B5 = {2,1,4,3,5}
	weakResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(4),
			{Col: 1, Row: 5}: NumberVal(5),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 2, Row: 2}: NumberVal(1),
			{Col: 2, Row: 3}: NumberVal(4),
			{Col: 2, Row: 4}: NumberVal(3),
			{Col: 2, Row: 5}: NumberVal(5),
		},
	}

	// Error in array: A1:A3 = {1,#VALUE!,3}, B1:B3 = {4,5,6}
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: ErrorVal(ErrValVALUE),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}

	// All text: A1:A3 = {"a","b","c"}, B1:B3 = {"x","y","z"}
	allTextResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: StringVal("b"),
			{Col: 1, Row: 3}: StringVal("c"),
			{Col: 2, Row: 1}: StringVal("x"),
			{Col: 2, Row: 2}: StringVal("y"),
			{Col: 2, Row: 3}: StringVal("z"),
		},
	}

	// Different length arrays: A1:A3, B1:B5
	diffLenResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
			{Col: 2, Row: 4}: NumberVal(7),
			{Col: 2, Row: 5}: NumberVal(8),
		},
	}

	tol := 1e-9

	tests := []struct {
		name      string
		formula   string
		resolver  *mockResolver
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic positive correlation
		{"basic_positive", "CORREL(A1:A5,B1:B5)", basicResolver, 0.997054486, false, 0},
		// Perfect positive correlation
		{"perfect_positive", "CORREL(A1:A3,B1:B3)", perfectPosResolver, 1.0, false, 0},
		// Perfect negative correlation
		{"perfect_negative", "CORREL(A1:A3,B1:B3)", perfectNegResolver, -1.0, false, 0},
		// Zero std dev in first array → #DIV/0!
		{"zero_stddev", "CORREL(A1:A3,B1:B3)", zeroStdDevResolver, 0, true, ErrValDIV0},
		// Two pairs
		{"two_pairs", "CORREL(A1:A2,B1:B2)", twoPairResolver, 1.0, false, 0},
		// Mixed types: pairs (1,10) and (3,30) → perfect positive
		{"mixed_types", "CORREL(A1:A4,B1:B4)", mixedResolver, 1.0, false, 0},
		// Negative numbers: perfect positive
		{"negative_numbers", "CORREL(A1:A4,B1:B4)", negResolver, 1.0, false, 0},
		// All zeros in one array → #DIV/0!
		{"all_zeros_one_array", "CORREL(A1:A3,B1:B3)", allZerosResolver, 0, true, ErrValDIV0},
		// Single pair → #DIV/0! (std dev is 0 with 1 point)
		{"single_pair", "CORREL(A1:A1,B1:B1)", singlePairResolver, 0, true, ErrValDIV0},
		// Empty range → #DIV/0!
		{"empty_range", "CORREL(A1:A3,B1:B3)", emptyResolver, 0, true, ErrValDIV0},
		// Zeros included (0 is numeric): perfect positive
		{"zeros_included", "CORREL(A1:A3,B1:B3)", zerosIncludedResolver, 1.0, false, 0},
		// Bool in array (pair skipped): pairs (1,10) and (3,30) → perfect positive
		{"bool_in_array", "CORREL(A1:A3,B1:B3)", boolResolver, 1.0, false, 0},
		// Large dataset: perfect positive (y = 2x)
		{"large_dataset", "CORREL(A1:A20,B1:B20)", largeResolver, 1.0, false, 0},
		// Weak correlation
		{"weak_correlation", "CORREL(A1:A5,B1:B5)", weakResolver, 0.8, false, 0},
		// Error propagation from first array
		{"error_in_array1", "CORREL(A1:A3,B1:B3)", errResolver, 0, true, ErrValVALUE},
		// All text → no numeric pairs → #DIV/0!
		{"all_text", "CORREL(A1:A3,B1:B3)", allTextResolver, 0, true, ErrValDIV0},
		// Different length arrays → #N/A
		{"different_lengths", "CORREL(A1:A3,B1:B5)", diffLenResolver, 0, true, ErrValNA},
		// Wrong number of arguments (1 arg)
		{"too_few_args", "CORREL(A1:A3)", basicResolver, 0, true, ErrValVALUE},
		// Wrong number of arguments (3 args)
		{"too_many_args", "CORREL(A1:A3,B1:B3,A1:A3)", basicResolver, 0, true, ErrValVALUE},
		// Reversed argument order: same magnitude, same sign
		{"reversed_args", "CORREL(B1:B5,A1:A5)", basicResolver, 0.997054486, false, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// SLOPE
// ---------------------------------------------------------------------------

func TestSLOPE(t *testing.T) {
	// Excel example: y={2,3,9,1,8,7,5}, x={6,5,11,7,5,4,4}
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2), {Col: 2, Row: 1}: NumberVal(6),
			{Col: 1, Row: 2}: NumberVal(3), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(9), {Col: 2, Row: 3}: NumberVal(11),
			{Col: 1, Row: 4}: NumberVal(1), {Col: 2, Row: 4}: NumberVal(7),
			{Col: 1, Row: 5}: NumberVal(8), {Col: 2, Row: 5}: NumberVal(5),
			{Col: 1, Row: 6}: NumberVal(7), {Col: 2, Row: 6}: NumberVal(4),
			{Col: 1, Row: 7}: NumberVal(5), {Col: 2, Row: 7}: NumberVal(4),
		},
	}

	// Perfect slope: y={2,4,6}, x={1,2,3} → slope=2
	perfectResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(6), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Negative slope: y={6,4,2}, x={1,2,3} → slope=-2
	negSlopeResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(6), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(2), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Zero slope: y={5,5,5}, x={1,2,3} → slope=0
	zeroSlopeResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(5), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(5), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Constant x: y={1,2,3}, x={5,5,5} → #DIV/0!
	constXResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(5),
		},
	}

	// Different lengths: A1:A3, B1:B5
	diffLenResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(4),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 2, Row: 4}: NumberVal(7),
			{Col: 2, Row: 5}: NumberVal(8),
		},
	}

	// Empty resolver
	emptyResolver := &mockResolver{cells: map[CellAddr]Value{}}

	// Single pair: y={3}, x={7} → #DIV/0!
	singlePairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3), {Col: 2, Row: 1}: NumberVal(7),
		},
	}

	// Two pairs: y={1,3}, x={2,4} → slope=1
	twoPairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(3), {Col: 2, Row: 2}: NumberVal(4),
		},
	}

	// Mixed types: some non-numeric skipped
	// row1: y=1(num), x="a"(str) → skip
	// row2: y=2(num), x=4(num) → keep
	// row3: y=true(bool), x=6(num) → skip
	// row4: y=8(num), x=8(num) → keep
	// pairs: (2,4),(8,8) → slope = (8-2)/(8-4) = 1.5
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(4),
			{Col: 1, Row: 3}: BoolVal(true), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 1, Row: 4}: NumberVal(8), {Col: 2, Row: 4}: NumberVal(8),
		},
	}

	// Error propagation: y contains #VALUE!
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: ErrorVal(ErrValVALUE), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Large dataset: y = 2x+1 for x=1..20
	largeCells := map[CellAddr]Value{}
	for i := 1; i <= 20; i++ {
		largeCells[CellAddr{Col: 1, Row: i}] = NumberVal(float64(2*i + 1))
		largeCells[CellAddr{Col: 2, Row: i}] = NumberVal(float64(i))
	}
	largeResolver := &mockResolver{cells: largeCells}

	// Fractional slope: y={1.5, 3.0, 4.5}, x={1,2,3} → slope=1.5
	fracResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1.5), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(3.0), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(4.5), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Negative values: y={-6,-4,-2}, x={1,2,3} → slope=2
	negValsResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-6), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(-4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(-2), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	tol := 1e-6

	tests := []struct {
		name      string
		formula   string
		resolver  *mockResolver
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		{"excel_example", "SLOPE(A1:A7,B1:B7)", excelResolver, 0.305555556, false, 0},
		{"perfect_slope_2", "SLOPE(A1:A3,B1:B3)", perfectResolver, 2.0, false, 0},
		{"negative_slope", "SLOPE(A1:A3,B1:B3)", negSlopeResolver, -2.0, false, 0},
		{"zero_slope", "SLOPE(A1:A3,B1:B3)", zeroSlopeResolver, 0.0, false, 0},
		{"constant_x_div0", "SLOPE(A1:A3,B1:B3)", constXResolver, 0, true, ErrValDIV0},
		{"different_lengths_na", "SLOPE(A1:A3,B1:B5)", diffLenResolver, 0, true, ErrValNA},
		{"empty_div0", "SLOPE(A1:A3,B1:B3)", emptyResolver, 0, true, ErrValDIV0},
		{"single_pair_div0", "SLOPE(A1:A1,B1:B1)", singlePairResolver, 0, true, ErrValDIV0},
		{"two_pairs", "SLOPE(A1:A2,B1:B2)", twoPairResolver, 1.0, false, 0},
		{"mixed_types_skip", "SLOPE(A1:A4,B1:B4)", mixedResolver, 1.5, false, 0},
		{"error_propagation", "SLOPE(A1:A3,B1:B3)", errResolver, 0, true, ErrValVALUE},
		{"large_dataset", "SLOPE(A1:A20,B1:B20)", largeResolver, 2.0, false, 0},
		{"fractional_slope", "SLOPE(A1:A3,B1:B3)", fracResolver, 1.5, false, 0},
		{"negative_values", "SLOPE(A1:A3,B1:B3)", negValsResolver, 2.0, false, 0},
		{"too_few_args", "SLOPE(A1:A3)", perfectResolver, 0, true, ErrValVALUE},
		{"too_many_args", "SLOPE(A1:A3,B1:B3,A1:A3)", perfectResolver, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// INTERCEPT
// ---------------------------------------------------------------------------

func TestINTERCEPT(t *testing.T) {
	// Excel example: y={2,3,9,1,8}, x={6,5,11,7,5}
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2), {Col: 2, Row: 1}: NumberVal(6),
			{Col: 1, Row: 2}: NumberVal(3), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(9), {Col: 2, Row: 3}: NumberVal(11),
			{Col: 1, Row: 4}: NumberVal(1), {Col: 2, Row: 4}: NumberVal(7),
			{Col: 1, Row: 5}: NumberVal(8), {Col: 2, Row: 5}: NumberVal(5),
		},
	}

	// y=2x+3: y={5,7,9}, x={1,2,3} → intercept=3
	interceptResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(7), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(9), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Negative intercept: y=2x-3: y={-1,1,3}, x={1,2,3} → intercept=-3
	negInterceptResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-1), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(1), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Zero intercept: y={2,4,6}, x={1,2,3} → intercept=0
	zeroInterceptResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(6), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Constant x → #DIV/0!
	constXResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(5),
		},
	}

	// Different lengths → #N/A
	diffLenResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(4),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 2, Row: 4}: NumberVal(7),
			{Col: 2, Row: 5}: NumberVal(8),
		},
	}

	// Empty → #DIV/0!
	emptyResolver := &mockResolver{cells: map[CellAddr]Value{}}

	// Single pair → #DIV/0!
	singlePairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3), {Col: 2, Row: 1}: NumberVal(7),
		},
	}

	// Two pairs: y={1,5}, x={2,4} → slope=2, intercept=-3
	twoPairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(5), {Col: 2, Row: 2}: NumberVal(4),
		},
	}

	// Mixed types: non-numeric skipped → pairs (2,4),(8,8) → slope=1.5, intercept=-4
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(4),
			{Col: 1, Row: 3}: BoolVal(true), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 1, Row: 4}: NumberVal(8), {Col: 2, Row: 4}: NumberVal(8),
		},
	}

	// Error propagation
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: ErrorVal(ErrValVALUE),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Large dataset: y = 3x+10 for x=1..20 → intercept=10
	largeCells := map[CellAddr]Value{}
	for i := 1; i <= 20; i++ {
		largeCells[CellAddr{Col: 1, Row: i}] = NumberVal(float64(3*i + 10))
		largeCells[CellAddr{Col: 2, Row: i}] = NumberVal(float64(i))
	}
	largeResolver := &mockResolver{cells: largeCells}

	// Fractional intercept: y=0.5x+2.5: y={3.0, 3.5, 4.0}, x={1,2,3} → intercept=2.5
	fracResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3.0), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(3.5), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(4.0), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Negative values: y={-6,-4,-2}, x={1,2,3} → slope=2, intercept=-8
	negValsResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-6), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(-4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(-2), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	tol := 1e-6

	tests := []struct {
		name      string
		formula   string
		resolver  *mockResolver
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		{"excel_example", "INTERCEPT(A1:A5,B1:B5)", excelResolver, 0.0483871, false, 0},
		{"intercept_3", "INTERCEPT(A1:A3,B1:B3)", interceptResolver, 3.0, false, 0},
		{"negative_intercept", "INTERCEPT(A1:A3,B1:B3)", negInterceptResolver, -3.0, false, 0},
		{"zero_intercept", "INTERCEPT(A1:A3,B1:B3)", zeroInterceptResolver, 0.0, false, 0},
		{"constant_x_div0", "INTERCEPT(A1:A3,B1:B3)", constXResolver, 0, true, ErrValDIV0},
		{"different_lengths_na", "INTERCEPT(A1:A3,B1:B5)", diffLenResolver, 0, true, ErrValNA},
		{"empty_div0", "INTERCEPT(A1:A3,B1:B3)", emptyResolver, 0, true, ErrValDIV0},
		{"single_pair_div0", "INTERCEPT(A1:A1,B1:B1)", singlePairResolver, 0, true, ErrValDIV0},
		{"two_pairs", "INTERCEPT(A1:A2,B1:B2)", twoPairResolver, -3.0, false, 0},
		{"mixed_types_skip", "INTERCEPT(A1:A4,B1:B4)", mixedResolver, -4.0, false, 0},
		{"error_propagation", "INTERCEPT(A1:A3,B1:B3)", errResolver, 0, true, ErrValVALUE},
		{"large_dataset", "INTERCEPT(A1:A20,B1:B20)", largeResolver, 10.0, false, 0},
		{"fractional_intercept", "INTERCEPT(A1:A3,B1:B3)", fracResolver, 2.5, false, 0},
		{"negative_values", "INTERCEPT(A1:A3,B1:B3)", negValsResolver, -8.0, false, 0},
		{"too_few_args", "INTERCEPT(A1:A3)", interceptResolver, 0, true, ErrValVALUE},
		{"too_many_args", "INTERCEPT(A1:A3,B1:B3,A1:A3)", interceptResolver, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// FORECAST / FORECAST.LINEAR
// ---------------------------------------------------------------------------

func TestFORECAST(t *testing.T) {
	// Excel example: y={6,7,9,15,21}, x_known={20,28,31,38,40}
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(6), {Col: 2, Row: 1}: NumberVal(20),
			{Col: 1, Row: 2}: NumberVal(7), {Col: 2, Row: 2}: NumberVal(28),
			{Col: 1, Row: 3}: NumberVal(9), {Col: 2, Row: 3}: NumberVal(31),
			{Col: 1, Row: 4}: NumberVal(15), {Col: 2, Row: 4}: NumberVal(38),
			{Col: 1, Row: 5}: NumberVal(21), {Col: 2, Row: 5}: NumberVal(40),
		},
	}

	// y=2x+1: y={3,5,7}, x={1,2,3}
	linearResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(5), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(7), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Two data points: y={1,3}, x={2,4} -> slope=1, intercept=-1
	twoPairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(3), {Col: 2, Row: 2}: NumberVal(4),
		},
	}

	// Constant x: y={1,2,3}, x={5,5,5} -> #DIV/0!
	constXResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(5),
		},
	}

	// Different lengths: A1:A3, B1:B5
	diffLenResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(4),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(5),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 2, Row: 4}: NumberVal(7),
			{Col: 2, Row: 5}: NumberVal(8),
		},
	}

	// Empty resolver
	emptyResolver := &mockResolver{cells: map[CellAddr]Value{}}

	// Single pair: y={3}, x={7} -> #DIV/0!
	singlePairResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3), {Col: 2, Row: 1}: NumberVal(7),
		},
	}

	// Mixed types: skip non-numeric pairs
	// row1: y=1(num), x="a"(str) -> skip
	// row2: y=2(num), x=4(num) -> keep
	// row3: y=true(bool), x=6(num) -> skip
	// row4: y=8(num), x=8(num) -> keep
	// pairs: (2,4),(8,8) -> slope=1.5, intercept=-4
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: NumberVal(4),
			{Col: 1, Row: 3}: BoolVal(true), {Col: 2, Row: 3}: NumberVal(6),
			{Col: 1, Row: 4}: NumberVal(8), {Col: 2, Row: 4}: NumberVal(8),
		},
	}

	// Error propagation: y contains #VALUE!
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: ErrorVal(ErrValVALUE), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Large dataset: y = 2x+1 for x=1..20
	largeCells := map[CellAddr]Value{}
	for i := 1; i <= 20; i++ {
		largeCells[CellAddr{Col: 1, Row: i}] = NumberVal(float64(2*i + 1))
		largeCells[CellAddr{Col: 2, Row: i}] = NumberVal(float64(i))
	}
	largeResolver := &mockResolver{cells: largeCells}

	// Negative values: y={-6,-4,-2}, x={1,2,3} -> slope=2, intercept=-8
	negValsResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(-6), {Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(-4), {Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(-2), {Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// String x that parses as number: "30" -> 30
	stringXResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(6), {Col: 2, Row: 1}: NumberVal(20),
			{Col: 1, Row: 2}: NumberVal(7), {Col: 2, Row: 2}: NumberVal(28),
			{Col: 1, Row: 3}: NumberVal(9), {Col: 2, Row: 3}: NumberVal(31),
			{Col: 1, Row: 4}: NumberVal(15), {Col: 2, Row: 4}: NumberVal(38),
			{Col: 1, Row: 5}: NumberVal(21), {Col: 2, Row: 5}: NumberVal(40),
			{Col: 3, Row: 1}: StringVal("30"),
		},
	}

	tol := 1e-6

	tests := []struct {
		name      string
		formula   string
		resolver  *mockResolver
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Excel example
		{"excel_example", "FORECAST(30,A1:A5,B1:B5)", excelResolver, 10.607253, false, 0},
		// FORECAST.LINEAR identical
		{"forecast_linear_same", "FORECAST.LINEAR(30,A1:A5,B1:B5)", excelResolver, 10.607253, false, 0},
		// Simple y=2x+1, predict x=5 -> 11
		{"linear_2x_plus_1", "FORECAST(5,A1:A3,B1:B3)", linearResolver, 11.0, false, 0},
		// x=0 should return intercept (y=2x+1 -> intercept=1)
		{"x_zero_intercept", "FORECAST(0,A1:A3,B1:B3)", linearResolver, 1.0, false, 0},
		// Negative x value: x=-3 -> 2*(-3)+1 = -5
		{"negative_x", "FORECAST(-3,A1:A3,B1:B3)", linearResolver, -5.0, false, 0},
		// x already in known data: x=2 -> 2*2+1 = 5
		{"x_in_known_data", "FORECAST(2,A1:A3,B1:B3)", linearResolver, 5.0, false, 0},
		// Extrapolation beyond range: x=100 -> 2*100+1 = 201
		{"extrapolation", "FORECAST(100,A1:A3,B1:B3)", linearResolver, 201.0, false, 0},
		// Non-numeric x -> #VALUE!
		{"non_numeric_x", "FORECAST(\"abc\",A1:A3,B1:B3)", linearResolver, 0, true, ErrValVALUE},
		// String x that can be coerced: "30"
		{"string_x_coerced", "FORECAST(C1,A1:A5,B1:B5)", stringXResolver, 10.607253, false, 0},
		// Empty arrays -> #DIV/0!
		{"empty_arrays", "FORECAST(5,A1:A3,B1:B3)", emptyResolver, 0, true, ErrValDIV0},
		// Different length arrays -> #N/A
		{"different_lengths", "FORECAST(5,A1:A3,B1:B5)", diffLenResolver, 0, true, ErrValNA},
		// Constant x values -> #DIV/0!
		{"constant_x_div0", "FORECAST(5,A1:A3,B1:B3)", constXResolver, 0, true, ErrValDIV0},
		// Single data point -> #DIV/0!
		{"single_point_div0", "FORECAST(5,A1:A1,B1:B1)", singlePairResolver, 0, true, ErrValDIV0},
		// Two data points: slope=1, intercept=-1, predict x=10 -> 9
		{"two_points", "FORECAST(10,A1:A2,B1:B2)", twoPairResolver, 9.0, false, 0},
		// Large dataset: y=2x+1, predict x=25 -> 51
		{"large_dataset", "FORECAST(25,A1:A20,B1:B20)", largeResolver, 51.0, false, 0},
		// Mixed types: slope=1.5, intercept=-4, predict x=10 -> 11
		{"mixed_types_skip", "FORECAST(10,A1:A4,B1:B4)", mixedResolver, 11.0, false, 0},
		// Error propagation from arrays
		{"error_propagation", "FORECAST(5,A1:A3,B1:B3)", errResolver, 0, true, ErrValVALUE},
		// Negative values: slope=2, intercept=-8, predict x=5 -> 2
		{"negative_values", "FORECAST(5,A1:A3,B1:B3)", negValsResolver, 2.0, false, 0},
		// Too few args
		{"too_few_args", "FORECAST(5,A1:A3)", linearResolver, 0, true, ErrValVALUE},
		// Too many args
		{"too_many_args", "FORECAST(5,A1:A3,B1:B3,A1:A3)", linearResolver, 0, true, ErrValVALUE},
		// FORECAST.LINEAR too few args
		{"linear_too_few_args", "FORECAST.LINEAR(5)", linearResolver, 0, true, ErrValVALUE},
		// Fractional x value: x=2.5 -> 2*2.5+1 = 6
		{"fractional_x", "FORECAST(2.5,A1:A3,B1:B3)", linearResolver, 6.0, false, 0},
		// FORECAST.LINEAR with two points
		{"linear_two_points", "FORECAST.LINEAR(10,A1:A2,B1:B2)", twoPairResolver, 9.0, false, 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// FISHER
// ---------------------------------------------------------------------------

func TestFISHER(t *testing.T) {
	const tol = 1e-6
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		{"basic_0.75", "FISHER(0.75)", 0.9729551, false, 0},
		{"zero", "FISHER(0)", 0, false, 0},
		{"half", "FISHER(0.5)", 0.5493061, false, 0},
		{"negative_half", "FISHER(-0.5)", -0.5493061, false, 0},
		{"near_one", "FISHER(0.99)", 2.6466524, false, 0},
		{"near_neg_one", "FISHER(-0.99)", -2.6466524, false, 0},
		{"small_value", "FISHER(0.01)", 0.0100003, false, 0},
		{"boundary_one", "FISHER(1)", 0, true, ErrValNUM},
		{"boundary_neg_one", "FISHER(-1)", 0, true, ErrValNUM},
		{"out_of_range_high", "FISHER(1.5)", 0, true, ErrValNUM},
		{"out_of_range_low", "FISHER(-1.5)", 0, true, ErrValNUM},
		{"text_error", `FISHER("text")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

func TestFISHER_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// No args
	cf := evalCompile(t, "FISHER()")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("FISHER() should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "FISHER(0.5,0.3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("FISHER(0.5,0.3) should error, got type=%d", got.Type)
	}
}

// ---------------------------------------------------------------------------
// FISHERINV
// ---------------------------------------------------------------------------

func TestFISHERINV(t *testing.T) {
	const tol = 1e-5
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		{"basic_roundtrip", "FISHERINV(0.972955)", 0.75, false, 0},
		{"zero", "FISHERINV(0)", 0, false, 0},
		{"half", "FISHERINV(0.5493061)", 0.5, false, 0},
		{"negative_half", "FISHERINV(-0.5493061)", -0.5, false, 0},
		{"large_positive", "FISHERINV(10)", 0.99999999, false, 0},
		{"large_negative", "FISHERINV(-10)", -0.99999999, false, 0},
		{"one", "FISHERINV(1)", 0.7615942, false, 0},
		{"negative_one", "FISHERINV(-1)", -0.7615942, false, 0},
		{"small_value", "FISHERINV(0.01)", 0.01, false, 0},
		{"text_error", `FISHERINV("text")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

func TestFISHERINV_argcount(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "FISHERINV()")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("FISHERINV() should error, got type=%d", got.Type)
	}

	cf = evalCompile(t, "FISHERINV(0.5,0.3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("FISHERINV(0.5,0.3) should error, got type=%d", got.Type)
	}
}

func TestFISHERINV_FISHER_roundtrip(t *testing.T) {
	const tol = 1e-10
	resolver := &mockResolver{}

	inputs := []float64{0.3, 0.7, -0.5, 0.99, -0.99, 0.01}
	for _, x := range inputs {
		t.Run(fmt.Sprintf("roundtrip_%g", x), func(t *testing.T) {
			formula := fmt.Sprintf("FISHERINV(FISHER(%g))", x)
			cf := evalCompile(t, formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-x) > tol {
				t.Errorf("FISHERINV(FISHER(%g)) = %g, want %g", x, got.Num, x)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// GAMMALN / GAMMALN.PRECISE
// ---------------------------------------------------------------------------

func TestGAMMALN(t *testing.T) {
	const tol = 1e-4
	resolver := &mockResolver{}

	tests := []struct {
		name      string
		formula   string
		wantNum   float64
		wantError bool
		wantErr   ErrorValue
	}{
		// Basic positive integer values
		{"gammaln_1", "GAMMALN(1)", 0, false, 0},
		{"gammaln_2", "GAMMALN(2)", 0, false, 0},
		{"gammaln_3", "GAMMALN(3)", 0.6931472, false, 0},
		{"gammaln_4", "GAMMALN(4)", 1.7917595, false, 0},
		{"gammaln_5", "GAMMALN(5)", 3.1780538, false, 0},
		{"gammaln_6", "GAMMALN(6)", 4.7874917, false, 0},
		{"gammaln_10", "GAMMALN(10)", 12.80183, false, 0},

		// Fractional values
		{"gammaln_0.5", "GAMMALN(0.5)", 0.5723649, false, 0},
		{"gammaln_1.5", "GAMMALN(1.5)", -0.1207822, false, 0},
		{"gammaln_2.5", "GAMMALN(2.5)", 0.2846829, false, 0},
		{"gammaln_0.001", "GAMMALN(0.001)", 6.9071786, false, 0},
		{"gammaln_0.1", "GAMMALN(0.1)", 2.2527127, false, 0},

		// Large values
		{"gammaln_100", "GAMMALN(100)", 359.1342, false, 0},
		{"gammaln_50", "GAMMALN(50)", 144.5657, false, 0},

		// Boolean coercion (TRUE=1)
		{"gammaln_true", "GAMMALN(TRUE)", 0, false, 0},

		// Error cases: non-positive
		{"gammaln_zero", "GAMMALN(0)", 0, true, ErrValNUM},
		{"gammaln_neg1", "GAMMALN(-1)", 0, true, ErrValNUM},
		{"gammaln_neg0.5", "GAMMALN(-0.5)", 0, true, ErrValNUM},
		{"gammaln_neg100", "GAMMALN(-100)", 0, true, ErrValNUM},

		// Error cases: text
		{"gammaln_text", `GAMMALN("text")`, 0, true, ErrValVALUE},

		// GAMMALN.PRECISE should behave identically
		{"precise_4", "GAMMALN.PRECISE(4)", 1.7917595, false, 0},
		{"precise_1", "GAMMALN.PRECISE(1)", 0, false, 0},
		{"precise_0.5", "GAMMALN.PRECISE(0.5)", 0.5723649, false, 0},
		{"precise_zero", "GAMMALN.PRECISE(0)", 0, true, ErrValNUM},
		{"precise_neg", "GAMMALN.PRECISE(-1)", 0, true, ErrValNUM},
		{"precise_text", `GAMMALN.PRECISE("text")`, 0, true, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("want error %v, got type=%d err=%v num=%g", tt.wantErr, got.Type, got.Err, got.Num)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("want number, got type=%d err=%v", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %f, want %f", got.Num, tt.wantNum)
			}
		})
	}
}

func TestGAMMALN_argcount(t *testing.T) {
	resolver := &mockResolver{}

	// No args
	cf := evalCompile(t, "GAMMALN()")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("GAMMALN() should error, got type=%d", got.Type)
	}

	// Too many args
	cf = evalCompile(t, "GAMMALN(1,2)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval error: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("GAMMALN(1,2) should error, got type=%d", got.Type)
	}
}
