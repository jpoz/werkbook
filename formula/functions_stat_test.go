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
	t.Run("empty cell", func(t *testing.T) {
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
	})

	t.Run("empty string counts as blank", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				// A2 is empty (missing)
				{Col: 1, Row: 3}: StringVal(""),    // empty string = blank
				{Col: 1, Row: 4}: StringVal("hello"),
				{Col: 1, Row: 5}: NumberVal(0),      // zero is not blank
			},
		}

		cf := evalCompile(t, "COUNTBLANK(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("COUNTBLANK: got %g, want 2", got.Num)
		}
	})

	t.Run("single empty string cell", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal(""),
			},
		}

		cf := evalCompile(t, "COUNTBLANK(A1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("COUNTBLANK: got %g, want 1", got.Num)
		}
	})
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
// COUNT – boolean handling
// ---------------------------------------------------------------------------

func TestCOUNTBooleanDirectArgs(t *testing.T) {
	resolver := &mockResolver{cells: map[CellAddr]Value{}}

	// Direct boolean args should be counted (Excel behavior).
	cf := evalCompile(t, "COUNT(TRUE,FALSE,10,20)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("COUNT(TRUE,FALSE,10,20): got %g, want 4", got.Num)
	}

	cf = evalCompile(t, "COUNT(TRUE)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNT(TRUE): got %g, want 1", got.Num)
	}
}

func TestCOUNTBooleanInRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: BoolVal(true),
			{Col: 1, Row: 2}: BoolVal(false),
		},
	}

	// Booleans in a range should NOT be counted.
	cf := evalCompile(t, "COUNT(A1:A2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("COUNT(A1:A2) with booleans: got %g, want 0", got.Num)
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

func TestAVERAGE(t *testing.T) {
	// Helper to build a resolver with numeric values in column A.
	numResolver := func(nums ...float64) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, n := range nums {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = NumberVal(n)
		}
		return m
	}

	// Helper for resolvers with arbitrary values in column A.
	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("single number", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("two numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(10,20)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf("got %v, want 15", got)
		}
	})

	t.Run("multiple numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(2,3,3,5,7,10)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("negative numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(-10,-20,-30)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -20 {
			t.Errorf("got %v, want -20", got)
		}
	})

	t.Run("mixed positive and negative", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(-10,10)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("decimal numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(1.5,2.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("zero values", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(0,0,10)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		want := 10.0 / 3.0
		if got.Type != ValueNumber || math.Abs(got.Num-want) > 1e-10 {
			t.Errorf("got %v, want %g", got, want)
		}
	})

	t.Run("all zeros", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(0,0,0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("large numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(1000000,2000000,3000000)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2000000 {
			t.Errorf("got %v, want 2000000", got)
		}
	})

	t.Run("very small numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(0.001,0.002,0.003)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-0.002) > 1e-15 {
			t.Errorf("got %v, want 0.002", got)
		}
	})

	t.Run("boolean TRUE as direct arg coerces to 1", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(TRUE,3)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE=1, so (1+3)/2 = 2
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("boolean FALSE as direct arg coerces to 0", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(FALSE,4)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// FALSE=0, so (0+4)/2 = 2
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("string number as direct arg coerces", func(t *testing.T) {
		// Direct string arg "5" is coerced by CoerceNum to 5.
		cf := evalCompile(t, `AVERAGE("5",15)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("non-numeric string as direct arg errors", func(t *testing.T) {
		cf := evalCompile(t, `AVERAGE("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			ErrorVal(ErrValNA),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("range with numbers", func(t *testing.T) {
		resolver := numResolver(10, 20, 30)
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("range ignores strings", func(t *testing.T) {
		// In a range, strings are ignored (only ValueNumber counted).
		resolver := valResolver(
			NumberVal(10),
			StringVal("hello"),
			NumberVal(30),
		)
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+30)/2 = 20
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("range ignores booleans", func(t *testing.T) {
		// In a range, booleans are ignored (only ValueNumber counted).
		resolver := valResolver(
			NumberVal(10),
			BoolVal(true),
			NumberVal(30),
		)
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+30)/2 = 20; TRUE is ignored in range
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("range with zero values included", func(t *testing.T) {
		resolver := numResolver(0, 10, 20)
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("range all empty gives DIV0", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("mixed range and literal", func(t *testing.T) {
		resolver := numResolver(10, 20)
		cf := evalCompile(t, "AVERAGE(A1:A2,30)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+20+30)/3 = 20
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("fractional result", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(1,2)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1.5 {
			t.Errorf("got %v, want 1.5", got)
		}
	})

	t.Run("single zero", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGE(0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("error in direct arg propagates", func(t *testing.T) {
		// 1/0 produces #DIV/0!, which should propagate through AVERAGE.
		cf := evalCompile(t, "AVERAGE(1/0,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("range with numeric string ignored", func(t *testing.T) {
		// Per Excel: numeric strings in ranges are ignored by AVERAGE.
		resolver := valResolver(
			NumberVal(10),
			StringVal("5"),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGE(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+20)/2 = 15; "5" in range is ignored
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf("got %v, want 15", got)
		}
	})

	_ = numResolver // suppress unused if needed
	_ = valResolver
}

func TestAVERAGEA(t *testing.T) {
	// Helper to build a resolver with arbitrary values in column A.
	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("single number", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("two numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(10,20)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf("got %v, want 15", got)
		}
	})

	t.Run("TRUE direct arg counts as 1", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(TRUE,3)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (1+3)/2 = 2
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("FALSE direct arg counts as 0", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(FALSE,4)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (0+4)/2 = 2
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("numeric string direct arg", func(t *testing.T) {
		cf := evalCompile(t, `AVERAGEA("5",15)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (5+15)/2 = 10
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("non-numeric string direct arg returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `AVERAGEA("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("range with text counts as 0", func(t *testing.T) {
		// Excel example: {10, 7, 9, 2, "Not available"} => (10+7+9+2+0)/5 = 5.6
		resolver := valResolver(
			NumberVal(10),
			NumberVal(7),
			NumberVal(9),
			NumberVal(2),
			StringVal("Not available"),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5.6 {
			t.Errorf("got %v, want 5.6", got)
		}
	})

	t.Run("range with TRUE counts as 1", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			BoolVal(true),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+1+20)/3 ≈ 10.333...
		want := 31.0 / 3.0
		if got.Type != ValueNumber || math.Abs(got.Num-want) > 1e-10 {
			t.Errorf("got %v, want %g", got, want)
		}
	})

	t.Run("range with FALSE counts as 0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			BoolVal(false),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+0+20)/3 = 10
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("empty cells ignored in range", func(t *testing.T) {
		// A1=10, A2=empty, A3=20
		resolver := valResolver(
			NumberVal(10),
		)
		resolver.cells[CellAddr{Col: 1, Row: 3}] = NumberVal(20)
		// A2 is empty (not set)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+20)/2 = 15; empty cell is ignored
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf("got %v, want 15", got)
		}
	})

	t.Run("all empty range returns DIV0", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("error propagates from range", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			ErrorVal(ErrValNUM),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("got %v, want #NUM!", got)
		}
	})

	t.Run("error propagates from direct arg", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(1/0,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("mixed range and direct args", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A2,30)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+20+30)/3 = 20
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("range all text", func(t *testing.T) {
		resolver := valResolver(
			StringVal("foo"),
			StringVal("bar"),
			StringVal("baz"),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (0+0+0)/3 = 0
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("range numeric string counts as 0 not parsed", func(t *testing.T) {
		// In range, "5" counts as 0, not 5
		resolver := valResolver(
			NumberVal(10),
			StringVal("5"),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+0+20)/3 = 10
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("single TRUE", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(TRUE)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("single FALSE", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(FALSE)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("negative numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(-10,-20,-30)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -20 {
			t.Errorf("got %v, want -20", got)
		}
	})

	t.Run("range with mixed types", func(t *testing.T) {
		// {10, TRUE, "hello", FALSE, 20} => (10+1+0+0+20)/5 = 6.2
		resolver := valResolver(
			NumberVal(10),
			BoolVal(true),
			StringVal("hello"),
			BoolVal(false),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 6.2 {
			t.Errorf("got %v, want 6.2", got)
		}
	})

	t.Run("range with empty string counts as 0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			StringVal(""),
			NumberVal(20),
		)
		cf := evalCompile(t, "AVERAGEA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// (10+0+20)/3 = 10
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("large numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(1000000,2000000,3000000)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2000000 {
			t.Errorf("got %v, want 2000000", got)
		}
	})

	t.Run("decimal numbers", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(1.5,2.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("zero value", func(t *testing.T) {
		cf := evalCompile(t, "AVERAGEA(0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	_ = valResolver
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
// MIN comprehensive tests
// ---------------------------------------------------------------------------

func TestMIN(t *testing.T) {
	// Helper to build a resolver with numeric values in column A.
	numResolver := func(nums ...float64) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, n := range nums {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = NumberVal(n)
		}
		return m
	}

	// Helper for resolvers with arbitrary values in column A.
	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("single number", func(t *testing.T) {
		cf := evalCompile(t, "MIN(42)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 42 {
			t.Errorf("got %v, want 42", got)
		}
	})

	t.Run("two numbers", func(t *testing.T) {
		cf := evalCompile(t, "MIN(10,20)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("multiple numbers picks smallest", func(t *testing.T) {
		cf := evalCompile(t, "MIN(5,3,8,1,9)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("all same values", func(t *testing.T) {
		cf := evalCompile(t, "MIN(7,7,7)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 7 {
			t.Errorf("got %v, want 7", got)
		}
	})

	t.Run("all negative numbers", func(t *testing.T) {
		cf := evalCompile(t, "MIN(-5,-3,-10,-1)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -10 {
			t.Errorf("got %v, want -10", got)
		}
	})

	t.Run("mixed positive and negative", func(t *testing.T) {
		cf := evalCompile(t, "MIN(10,-5,3,-20,0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -20 {
			t.Errorf("got %v, want -20", got)
		}
	})

	t.Run("zero is minimum", func(t *testing.T) {
		cf := evalCompile(t, "MIN(5,10,0,3)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("decimal numbers", func(t *testing.T) {
		cf := evalCompile(t, "MIN(1.5,0.5,2.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0.5 {
			t.Errorf("got %v, want 0.5", got)
		}
	})

	t.Run("large numbers", func(t *testing.T) {
		cf := evalCompile(t, "MIN(1000000,2000000,500000)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 500000 {
			t.Errorf("got %v, want 500000", got)
		}
	})

	t.Run("very small numbers", func(t *testing.T) {
		cf := evalCompile(t, "MIN(0.001,0.002,0.0005)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-0.0005) > 1e-15 {
			t.Errorf("got %v, want 0.0005", got)
		}
	})

	t.Run("boolean TRUE as direct arg coerces to 1", func(t *testing.T) {
		cf := evalCompile(t, "MIN(TRUE,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE=1, so min(1,5) = 1
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("boolean FALSE as direct arg coerces to 0", func(t *testing.T) {
		cf := evalCompile(t, "MIN(FALSE,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// FALSE=0, so min(0,5) = 0
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("boolean TRUE and FALSE as direct args", func(t *testing.T) {
		cf := evalCompile(t, "MIN(TRUE,FALSE)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE=1, FALSE=0, so min = 0
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("string number as direct arg coerces", func(t *testing.T) {
		cf := evalCompile(t, `MIN("3",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// "3" coerced to 3, so min(3,10) = 3
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("non-numeric string as direct arg errors", func(t *testing.T) {
		cf := evalCompile(t, `MIN("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("error propagation from range", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			ErrorVal(ErrValNA),
			NumberVal(5),
		)
		cf := evalCompile(t, "MIN(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("error propagation DIV0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			ErrorVal(ErrValDIV0),
			NumberVal(5),
		)
		cf := evalCompile(t, "MIN(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("no args returns 0", func(t *testing.T) {
		// MIN with empty range => no numbers found => returns 0
		resolver := &mockResolver{}
		cf := evalCompile(t, "MIN(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("range with numbers", func(t *testing.T) {
		resolver := numResolver(10, 20, 5, 15)
		cf := evalCompile(t, "MIN(A1:A4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("range ignores strings", func(t *testing.T) {
		// In a range, strings are ignored (only ValueNumber counted).
		resolver := valResolver(
			NumberVal(10),
			StringVal("hello"),
			NumberVal(5),
		)
		cf := evalCompile(t, "MIN(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("range ignores booleans", func(t *testing.T) {
		// In a range, booleans are ignored (only ValueNumber counted).
		resolver := valResolver(
			NumberVal(10),
			BoolVal(true),
			NumberVal(5),
		)
		cf := evalCompile(t, "MIN(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE is ignored in range, so min(10,5) = 5
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("range ignores empty cells", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			EmptyVal(),
			NumberVal(5),
		)
		cf := evalCompile(t, "MIN(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("range with only strings returns 0", func(t *testing.T) {
		// All non-numeric => no numbers found => returns 0
		resolver := valResolver(
			StringVal("foo"),
			StringVal("bar"),
		)
		cf := evalCompile(t, "MIN(A1:A2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("mixed direct args and range", func(t *testing.T) {
		resolver := numResolver(10, 20)
		cf := evalCompile(t, "MIN(A1:A2,3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// min(10, 20, 3) = 3
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("negative decimal", func(t *testing.T) {
		cf := evalCompile(t, "MIN(-0.5,0.5,-1.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -1.5 {
			t.Errorf("got %v, want -1.5", got)
		}
	})

	t.Run("Excel example from docs", func(t *testing.T) {
		// Data: 10, 7, 9, 27, 2 => MIN = 2
		resolver := numResolver(10, 7, 9, 27, 2)
		cf := evalCompile(t, "MIN(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("Excel example with zero", func(t *testing.T) {
		// MIN(A1:A5, 0) where data is 10, 7, 9, 27, 2 => 0
		resolver := numResolver(10, 7, 9, 27, 2)
		cf := evalCompile(t, "MIN(A1:A5,0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})
}

// ---------------------------------------------------------------------------
// MAX comprehensive tests
// ---------------------------------------------------------------------------

func TestMAX(t *testing.T) {
	// Helper to build a resolver with numeric values in column A.
	numResolver := func(nums ...float64) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, n := range nums {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = NumberVal(n)
		}
		return m
	}

	// Helper for resolvers with arbitrary values in column A.
	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("single number", func(t *testing.T) {
		cf := evalCompile(t, "MAX(42)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 42 {
			t.Errorf("got %v, want 42", got)
		}
	})

	t.Run("two numbers", func(t *testing.T) {
		cf := evalCompile(t, "MAX(10,20)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("multiple numbers picks largest", func(t *testing.T) {
		cf := evalCompile(t, "MAX(5,3,8,1,9)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 9 {
			t.Errorf("got %v, want 9", got)
		}
	})

	t.Run("all same values", func(t *testing.T) {
		cf := evalCompile(t, "MAX(7,7,7)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 7 {
			t.Errorf("got %v, want 7", got)
		}
	})

	t.Run("all negative numbers", func(t *testing.T) {
		cf := evalCompile(t, "MAX(-5,-3,-10,-1)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -1 {
			t.Errorf("got %v, want -1", got)
		}
	})

	t.Run("mixed positive and negative", func(t *testing.T) {
		cf := evalCompile(t, "MAX(10,-5,3,-20,0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("zero is maximum", func(t *testing.T) {
		cf := evalCompile(t, "MAX(-5,-10,0,-3)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("decimal numbers", func(t *testing.T) {
		cf := evalCompile(t, "MAX(1.5,0.5,2.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2.5 {
			t.Errorf("got %v, want 2.5", got)
		}
	})

	t.Run("large numbers", func(t *testing.T) {
		cf := evalCompile(t, "MAX(1000000,2000000,500000)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2000000 {
			t.Errorf("got %v, want 2000000", got)
		}
	})

	t.Run("very small numbers", func(t *testing.T) {
		cf := evalCompile(t, "MAX(0.001,0.002,0.0005)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-0.002) > 1e-15 {
			t.Errorf("got %v, want 0.002", got)
		}
	})

	t.Run("boolean TRUE as direct arg coerces to 1", func(t *testing.T) {
		cf := evalCompile(t, "MAX(TRUE,-5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE=1, so max(1,-5) = 1
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("boolean FALSE as direct arg coerces to 0", func(t *testing.T) {
		cf := evalCompile(t, "MAX(FALSE,-5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// FALSE=0, so max(0,-5) = 0
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("boolean TRUE and FALSE as direct args", func(t *testing.T) {
		cf := evalCompile(t, "MAX(TRUE,FALSE)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE=1, FALSE=0, so max = 1
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("string number as direct arg coerces", func(t *testing.T) {
		cf := evalCompile(t, `MAX("3",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// "3" coerced to 3, so max(3,10) = 10
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("non-numeric string as direct arg errors", func(t *testing.T) {
		cf := evalCompile(t, `MAX("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("error propagation NA from range", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			ErrorVal(ErrValNA),
			NumberVal(5),
		)
		cf := evalCompile(t, "MAX(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("error propagation DIV0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			ErrorVal(ErrValDIV0),
			NumberVal(5),
		)
		cf := evalCompile(t, "MAX(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("no args returns 0", func(t *testing.T) {
		// MAX with empty range => no numbers found => returns 0
		resolver := &mockResolver{}
		cf := evalCompile(t, "MAX(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("range with numbers", func(t *testing.T) {
		resolver := numResolver(10, 20, 5, 15)
		cf := evalCompile(t, "MAX(A1:A4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("range ignores strings", func(t *testing.T) {
		// In a range, strings are ignored (only ValueNumber counted).
		resolver := valResolver(
			NumberVal(10),
			StringVal("hello"),
			NumberVal(5),
		)
		cf := evalCompile(t, "MAX(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("range ignores booleans", func(t *testing.T) {
		// In a range, booleans are ignored (only ValueNumber counted).
		resolver := valResolver(
			NumberVal(10),
			BoolVal(true),
			NumberVal(5),
		)
		cf := evalCompile(t, "MAX(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// TRUE is ignored in range, so max(10,5) = 10
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("range ignores empty cells", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(10),
			EmptyVal(),
			NumberVal(5),
		)
		cf := evalCompile(t, "MAX(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("range with only strings returns 0", func(t *testing.T) {
		// All non-numeric => no numbers found => returns 0
		resolver := valResolver(
			StringVal("foo"),
			StringVal("bar"),
		)
		cf := evalCompile(t, "MAX(A1:A2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("mixed direct args and range", func(t *testing.T) {
		resolver := numResolver(10, 20)
		cf := evalCompile(t, "MAX(A1:A2,30)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// max(10, 20, 30) = 30
		if got.Type != ValueNumber || got.Num != 30 {
			t.Errorf("got %v, want 30", got)
		}
	})

	t.Run("negative decimal", func(t *testing.T) {
		cf := evalCompile(t, "MAX(-0.5,0.5,-1.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0.5 {
			t.Errorf("got %v, want 0.5", got)
		}
	})

	t.Run("Excel example from docs", func(t *testing.T) {
		// Data: 10, 7, 9, 27, 2 => MAX = 27
		resolver := numResolver(10, 7, 9, 27, 2)
		cf := evalCompile(t, "MAX(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 27 {
			t.Errorf("got %v, want 27", got)
		}
	})

	t.Run("Excel example with 30", func(t *testing.T) {
		// MAX(A1:A5, 30) where data is 10, 7, 9, 27, 2 => 30
		resolver := numResolver(10, 7, 9, 27, 2)
		cf := evalCompile(t, "MAX(A1:A5,30)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 30 {
			t.Errorf("got %v, want 30", got)
		}
	})
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

// ---------------------------------------------------------------------------
// PERCENTRANK / PERCENTRANK.INC
// ---------------------------------------------------------------------------

func TestPERCENTRANK(t *testing.T) {
	// Excel example data: {13,12,11,8,4,3,2,1,1,1}
	// Sorted: {1,1,1,2,3,4,8,11,12,13}
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}:  NumberVal(13),
			{Col: 1, Row: 2}:  NumberVal(12),
			{Col: 1, Row: 3}:  NumberVal(11),
			{Col: 1, Row: 4}:  NumberVal(8),
			{Col: 1, Row: 5}:  NumberVal(4),
			{Col: 1, Row: 6}:  NumberVal(3),
			{Col: 1, Row: 7}:  NumberVal(2),
			{Col: 1, Row: 8}:  NumberVal(1),
			{Col: 1, Row: 9}:  NumberVal(1),
			{Col: 1, Row: 10}: NumberVal(1),
		},
	}

	// Simple data for basic tests: {1,2,3,4,5} in B1:B5
	simpleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 2, Row: 4}: NumberVal(4),
			{Col: 2, Row: 5}: NumberVal(5),
		},
	}

	// Single element in C1
	singleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(42),
		},
	}

	// Negative numbers: {-10,-5,0,5,10} in D1:D5
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(-10),
			{Col: 4, Row: 2}: NumberVal(-5),
			{Col: 4, Row: 3}: NumberVal(0),
			{Col: 4, Row: 4}: NumberVal(5),
			{Col: 4, Row: 5}: NumberVal(10),
		},
	}

	// Mixed types: numbers and strings in E1:E4
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(10),
			{Col: 5, Row: 2}: StringVal("hello"),
			{Col: 5, Row: 3}: NumberVal(20),
			{Col: 5, Row: 4}: NumberVal(30),
		},
	}

	tests := []struct {
		name     string
		formula  string
		resolver *mockResolver
		wantNum  float64
		wantErr  ErrorValue // 0 means expect number result
	}{
		// Excel doc examples
		{name: "excel_x=2", formula: "PERCENTRANK(A1:A10,2)", resolver: excelResolver, wantNum: 0.333},
		{name: "excel_x=4", formula: "PERCENTRANK(A1:A10,4)", resolver: excelResolver, wantNum: 0.555},
		{name: "excel_x=8", formula: "PERCENTRANK(A1:A10,8)", resolver: excelResolver, wantNum: 0.666},
		{name: "excel_x=5_interp", formula: "PERCENTRANK(A1:A10,5)", resolver: excelResolver, wantNum: 0.583},

		// Min and max of range
		{name: "x_equals_min", formula: "PERCENTRANK(B1:B5,1)", resolver: simpleResolver, wantNum: 0},
		{name: "x_equals_max", formula: "PERCENTRANK(B1:B5,5)", resolver: simpleResolver, wantNum: 1},

		// x outside range → #N/A
		{name: "x_below_min", formula: "PERCENTRANK(B1:B5,0)", resolver: simpleResolver, wantErr: ErrValNA},
		{name: "x_above_max", formula: "PERCENTRANK(B1:B5,6)", resolver: simpleResolver, wantErr: ErrValNA},

		// Significance parameter
		{name: "sig_1", formula: "PERCENTRANK(A1:A10,5,1)", resolver: excelResolver, wantNum: 0.5},
		{name: "sig_5", formula: "PERCENTRANK(A1:A10,5,5)", resolver: excelResolver, wantNum: 0.58333},
		{name: "default_sig_3", formula: "PERCENTRANK(B1:B5,2)", resolver: simpleResolver, wantNum: 0.25},

		// Single element array
		{name: "single_element_match", formula: "PERCENTRANK(C1:C1,42)", resolver: singleResolver, wantNum: 1},
		{name: "single_element_no_match_below", formula: "PERCENTRANK(C1:C1,10)", resolver: singleResolver, wantErr: ErrValNA},
		{name: "single_element_no_match_above", formula: "PERCENTRANK(C1:C1,50)", resolver: singleResolver, wantErr: ErrValNA},

		// Duplicate values (x=1 appears 3 times at positions 0,1,2 → first occurrence → 0/9=0)
		{name: "duplicate_min", formula: "PERCENTRANK(A1:A10,1)", resolver: excelResolver, wantNum: 0},

		// Negative numbers
		{name: "neg_min", formula: "PERCENTRANK(D1:D5,-10)", resolver: negResolver, wantNum: 0},
		{name: "neg_max", formula: "PERCENTRANK(D1:D5,10)", resolver: negResolver, wantNum: 1},
		{name: "neg_mid", formula: "PERCENTRANK(D1:D5,0)", resolver: negResolver, wantNum: 0.5},
		{name: "neg_interp", formula: "PERCENTRANK(D1:D5,-3)", resolver: negResolver, wantNum: 0.35},

		// Unsorted data produces same result as sorted
		{name: "unsorted_data", formula: "PERCENTRANK(A1:A10,13)", resolver: excelResolver, wantNum: 1},

		// significance < 1 → #NUM!
		{name: "sig_zero", formula: "PERCENTRANK(B1:B5,2,0)", resolver: simpleResolver, wantErr: ErrValNUM},
		{name: "sig_negative", formula: "PERCENTRANK(B1:B5,2,-1)", resolver: simpleResolver, wantErr: ErrValNUM},

		// Mixed types in array (non-numeric ignored)
		{name: "mixed_types", formula: "PERCENTRANK(E1:E4,20)", resolver: mixedResolver, wantNum: 0.5},

		// PERCENTRANK.INC gives same results
		{name: "inc_same_as_base", formula: "PERCENTRANK.INC(A1:A10,2)", resolver: excelResolver, wantNum: 0.333},
		{name: "inc_interp", formula: "PERCENTRANK.INC(A1:A10,5)", resolver: excelResolver, wantNum: 0.583},
		{name: "inc_max", formula: "PERCENTRANK.INC(B1:B5,5)", resolver: simpleResolver, wantNum: 1},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr != 0 {
				if got.Type != ValueError {
					t.Fatalf("expected error %d, got type=%d num=%g", tt.wantErr, got.Type, got.Num)
				}
				if got.Err != tt.wantErr {
					t.Errorf("expected error %d, got %d", tt.wantErr, got.Err)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("expected number, got type=%d err=%d", got.Type, got.Err)
			}
			if math.Abs(got.Num-tt.wantNum) > 1e-9 {
				t.Errorf("got %g, want %g", got.Num, tt.wantNum)
			}
		})
	}

	// Wrong argument count
	t.Run("too_few_args", func(t *testing.T) {
		cf := evalCompile(t, "PERCENTRANK(B1:B5)")
		got, err := Eval(cf, simpleResolver, nil)
		if err != nil {
			t.Fatalf("Eval error: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("expected error, got type=%d", got.Type)
		}
	})

	t.Run("too_many_args", func(t *testing.T) {
		cf := evalCompile(t, "PERCENTRANK(B1:B5,2,3,4)")
		got, err := Eval(cf, simpleResolver, nil)
		if err != nil {
			t.Fatalf("Eval error: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("expected error, got type=%d", got.Type)
		}
	})

	// Empty array → #NUM!
	t.Run("empty_array", func(t *testing.T) {
		emptyResolver := &mockResolver{cells: map[CellAddr]Value{}}
		cf := evalCompile(t, "PERCENTRANK(F1:F3,1)")
		got, err := Eval(cf, emptyResolver, nil)
		if err != nil {
			t.Fatalf("Eval error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("expected #NUM!, got type=%d err=%d", got.Type, got.Err)
		}
	})
}

// ---------------------------------------------------------------------------
// SKEW
// ---------------------------------------------------------------------------

func TestSKEW(t *testing.T) {
	// Resolver with {3,4,5,2,3,4,5,6,4,7} in A1:A10 (Excel docs example)
	excelResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}:  NumberVal(3),
			{Col: 1, Row: 2}:  NumberVal(4),
			{Col: 1, Row: 3}:  NumberVal(5),
			{Col: 1, Row: 4}:  NumberVal(2),
			{Col: 1, Row: 5}:  NumberVal(3),
			{Col: 1, Row: 6}:  NumberVal(4),
			{Col: 1, Row: 7}:  NumberVal(5),
			{Col: 1, Row: 8}:  NumberVal(6),
			{Col: 1, Row: 9}:  NumberVal(4),
			{Col: 1, Row: 10}: NumberVal(7),
		},
	}

	// Symmetric data {1,2,3,4,5} in B1:B5
	symResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 2, Row: 4}: NumberVal(4),
			{Col: 2, Row: 5}: NumberVal(5),
		},
	}

	// Right-skewed {1,1,1,1,1,1,1,1,1,100} in C1:C10
	rightSkewResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}:  NumberVal(1),
			{Col: 3, Row: 2}:  NumberVal(1),
			{Col: 3, Row: 3}:  NumberVal(1),
			{Col: 3, Row: 4}:  NumberVal(1),
			{Col: 3, Row: 5}:  NumberVal(1),
			{Col: 3, Row: 6}:  NumberVal(1),
			{Col: 3, Row: 7}:  NumberVal(1),
			{Col: 3, Row: 8}:  NumberVal(1),
			{Col: 3, Row: 9}:  NumberVal(1),
			{Col: 3, Row: 10}: NumberVal(100),
		},
	}

	// Left-skewed {1,100,100,100,100,100,100,100,100,100} in D1:D10
	leftSkewResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}:  NumberVal(1),
			{Col: 4, Row: 2}:  NumberVal(100),
			{Col: 4, Row: 3}:  NumberVal(100),
			{Col: 4, Row: 4}:  NumberVal(100),
			{Col: 4, Row: 5}:  NumberVal(100),
			{Col: 4, Row: 6}:  NumberVal(100),
			{Col: 4, Row: 7}:  NumberVal(100),
			{Col: 4, Row: 8}:  NumberVal(100),
			{Col: 4, Row: 9}:  NumberVal(100),
			{Col: 4, Row: 10}: NumberVal(100),
		},
	}

	// Exactly 3 data points {1,2,3} in E1:E3
	threeResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(1),
			{Col: 5, Row: 2}: NumberVal(2),
			{Col: 5, Row: 3}: NumberVal(3),
		},
	}

	// Two data points {1,2} in F1:F2
	twoResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 6, Row: 1}: NumberVal(1),
			{Col: 6, Row: 2}: NumberVal(2),
		},
	}

	// Single value {5} in G1
	singleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 7, Row: 1}: NumberVal(5),
		},
	}

	// All same values {4,4,4,4} in H1:H4
	sameResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 8, Row: 1}: NumberVal(4),
			{Col: 8, Row: 2}: NumberVal(4),
			{Col: 8, Row: 3}: NumberVal(4),
			{Col: 8, Row: 4}: NumberVal(4),
		},
	}

	// Negative numbers {-5,-3,-1,0,2} in I1:I5
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 9, Row: 1}: NumberVal(-5),
			{Col: 9, Row: 2}: NumberVal(-3),
			{Col: 9, Row: 3}: NumberVal(-1),
			{Col: 9, Row: 4}: NumberVal(0),
			{Col: 9, Row: 5}: NumberVal(2),
		},
	}

	// Large dataset 1..20 in J1:J20
	largeResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for i := 1; i <= 20; i++ {
		largeResolver.cells[CellAddr{Col: 10, Row: i}] = NumberVal(float64(i))
	}

	// Mixed types: numbers, strings, booleans in K1:K6
	mixedResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 11, Row: 1}: NumberVal(1),
			{Col: 11, Row: 2}: StringVal("hello"),
			{Col: 11, Row: 3}: NumberVal(2),
			{Col: 11, Row: 4}: BoolVal(true),
			{Col: 11, Row: 5}: NumberVal(3),
			{Col: 11, Row: 6}: NumberVal(10),
		},
	}

	// Error in array in L1:L4
	errResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 12, Row: 1}: NumberVal(1),
			{Col: 12, Row: 2}: NumberVal(2),
			{Col: 12, Row: 3}: ErrorVal(ErrValNUM),
			{Col: 12, Row: 4}: NumberVal(4),
		},
	}

	emptyResolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}

	tests := []struct {
		name     string
		formula  string
		resolver CellResolver
		wantNum  float64
		wantErr  ErrorValue
		isErr    bool
		tol      float64
	}{
		// Excel docs example: {3,4,5,2,3,4,5,6,4,7} → 0.359543
		{"excel_example", "SKEW(A1:A10)", excelResolver, 0.359543, 0, false, 1e-4},

		// Symmetric data {1,2,3,4,5} → skew = 0
		{"symmetric", "SKEW(B1:B5)", symResolver, 0, 0, false, 1e-9},

		// Right-skewed data → positive skew
		{"right_skewed", "SKEW(C1:C10)", rightSkewResolver, 3.16227766, 0, false, 1e-4},

		// Left-skewed data → negative skew
		{"left_skewed", "SKEW(D1:D10)", leftSkewResolver, -3.16227766, 0, false, 1e-4},

		// Exactly 3 data points {1,2,3} → 0
		{"three_points", "SKEW(E1:E3)", threeResolver, 0, 0, false, 1e-9},

		// Two data points → #DIV/0!
		{"two_points_div0", "SKEW(F1:F2)", twoResolver, 0, ErrValDIV0, true, 0},

		// Single value → #DIV/0!
		{"single_value_div0", "SKEW(G1)", singleResolver, 0, ErrValDIV0, true, 0},

		// All same values → #DIV/0! (std dev = 0)
		{"all_same_div0", "SKEW(H1:H4)", sameResolver, 0, ErrValDIV0, true, 0},

		// Negative numbers {-5,-3,-1,0,2}
		{"negative_numbers", "SKEW(I1:I5)", negResolver, -0.18252326, 0, false, 1e-4},

		// Large dataset 1..20 → 0 (symmetric)
		{"large_symmetric", "SKEW(J1:J20)", largeResolver, 0, 0, false, 1e-9},

		// Mixed types in array: text and bool are ignored → only {1,2,3,10}
		{"mixed_types_array", "SKEW(K1:K6)", mixedResolver, 1.76363261, 0, false, 1e-4},

		// Direct boolean args are counted: TRUE=1 → SKEW(1,2,3,TRUE) = SKEW(1,2,3,1)
		{"direct_bool_true", "SKEW(1,2,3,TRUE)", emptyResolver, 0.85456304, 0, false, 1e-4},

		// Direct string number args are counted: "5" → 5
		{"direct_string_num", `SKEW(1,2,3,"5")`, emptyResolver, 0.75283720, 0, false, 1e-4},

		// Error propagation from array
		{"error_propagation", "SKEW(L1:L4)", errResolver, 0, ErrValNUM, true, 0},

		// Direct error arg
		{"direct_error", "SKEW(1,2,3,1/0)", emptyResolver, 0, ErrValDIV0, true, 0},

		// Empty range → #DIV/0! (0 values < 3)
		{"empty_range", "SKEW(Z1:Z5)", emptyResolver, 0, ErrValDIV0, true, 0},

		// Direct args: SKEW(1,2,3) with all different → 0
		{"direct_three_sym", "SKEW(1,2,3)", emptyResolver, 0, 0, false, 1e-9},

		// Direct args with more values
		{"direct_many", "SKEW(1,1,1,1,1,100)", emptyResolver, 2.44948975, 0, false, 1e-4},

		// {1,2,3,4,100} right-skewed
		{"moderate_right_skew", "SKEW(1,2,3,4,100)", emptyResolver, 2.23239591, 0, false, 1e-4},

		// Decimals {0.5, 1.5, 2.5, 3.5, 4.5}
		{"decimals", "SKEW(0.5,1.5,2.5,3.5,4.5)", emptyResolver, 0, 0, false, 1e-9},

		// Large positive values
		{"large_values", "SKEW(1000000,2000000,3000000)", emptyResolver, 0, 0, false, 1e-9},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.isErr {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("got type %d (%v), want number", got.Type, got)
			}
			tol := tt.tol
			if tol == 0 {
				tol = 1e-9
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("got %g, want %g (diff %g)", got.Num, tt.wantNum, math.Abs(got.Num-tt.wantNum))
			}
		})
	}

	// 0 args → should error
	t.Run("zero_args", func(t *testing.T) {
		got, err := fnSKEW([]Value{})
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})

	// Verify sign: right-skewed > 0
	t.Run("right_skew_positive", func(t *testing.T) {
		cf := evalCompile(t, "SKEW(C1:C10)")
		got, err := Eval(cf, rightSkewResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num <= 0 {
			t.Errorf("expected positive skew, got %v", got)
		}
	})

	// Verify sign: left-skewed < 0
	t.Run("left_skew_negative", func(t *testing.T) {
		cf := evalCompile(t, "SKEW(D1:D10)")
		got, err := Eval(cf, leftSkewResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num >= 0 {
			t.Errorf("expected negative skew, got %v", got)
		}
	})
}

func TestMAXA(t *testing.T) {
	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("single number", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("two numbers picks larger", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(3,7)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 7 {
			t.Errorf("got %v, want 7", got)
		}
	})

	t.Run("TRUE direct arg counts as 1", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(TRUE,0.5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("FALSE direct arg counts as 0", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(FALSE,-5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("numeric string direct arg", func(t *testing.T) {
		cf := evalCompile(t, `MAXA("10",5)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("non-numeric string direct arg returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `MAXA("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("range with text counts as 0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(-3),
			NumberVal(-5),
			StringVal("hello"),
		)
		cf := evalCompile(t, "MAXA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// text=0 is the largest
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("range with TRUE counts as 1", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(0),
			NumberVal(0.5),
			BoolVal(true),
		)
		cf := evalCompile(t, "MAXA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("range with FALSE counts as 0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(-1),
			NumberVal(-2),
			BoolVal(false),
		)
		cf := evalCompile(t, "MAXA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("empty range returns 0", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(A1:A3)")
		got, err := Eval(cf, &mockResolver{cells: map[CellAddr]Value{}}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("error in range propagates", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(1),
			ErrorVal(ErrValDIV0),
			NumberVal(3),
		)
		cf := evalCompile(t, "MAXA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("error direct arg propagates", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(1/0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("all negative numbers", func(t *testing.T) {
		cf := evalCompile(t, "MAXA(-10,-20,-5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -5 {
			t.Errorf("got %v, want -5", got)
		}
	})

	t.Run("Excel doc example", func(t *testing.T) {
		// {0, 0.2, 0.5, 0.4, TRUE} => max is TRUE=1
		resolver := valResolver(
			NumberVal(0),
			NumberVal(0.2),
			NumberVal(0.5),
			NumberVal(0.4),
			BoolVal(true),
		)
		cf := evalCompile(t, "MAXA(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("mixed range and direct args", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(2),
			NumberVal(3),
		)
		cf := evalCompile(t, "MAXA(A1:A2,10)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})

	t.Run("empty cells in range ignored", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(-5),
			Value{Type: ValueEmpty},
			NumberVal(-3),
		)
		cf := evalCompile(t, "MAXA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -3 {
			t.Errorf("got %v, want -3", got)
		}
	})
}

func TestMINA(t *testing.T) {
	valResolver := func(vals ...Value) *mockResolver {
		m := &mockResolver{cells: map[CellAddr]Value{}}
		for i, v := range vals {
			m.cells[CellAddr{Col: 1, Row: i + 1}] = v
		}
		return m
	}

	t.Run("single number", func(t *testing.T) {
		cf := evalCompile(t, "MINA(5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("two numbers picks smaller", func(t *testing.T) {
		cf := evalCompile(t, "MINA(3,7)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("TRUE direct arg counts as 1", func(t *testing.T) {
		cf := evalCompile(t, "MINA(TRUE,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("FALSE direct arg counts as 0", func(t *testing.T) {
		cf := evalCompile(t, "MINA(FALSE,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("numeric string direct arg", func(t *testing.T) {
		cf := evalCompile(t, `MINA("2",5)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("non-numeric string direct arg returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `MINA("hello",10)`)
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("range with text counts as 0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(3),
			NumberVal(5),
			StringVal("hello"),
		)
		cf := evalCompile(t, "MINA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// text=0 is the smallest
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("range with TRUE counts as 1", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(2),
			NumberVal(3),
			BoolVal(true),
		)
		cf := evalCompile(t, "MINA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("range with FALSE counts as 0", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(1),
			NumberVal(2),
			BoolVal(false),
		)
		cf := evalCompile(t, "MINA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("empty range returns 0", func(t *testing.T) {
		cf := evalCompile(t, "MINA(A1:A3)")
		got, err := Eval(cf, &mockResolver{cells: map[CellAddr]Value{}}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("error in range propagates", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(1),
			ErrorVal(ErrValDIV0),
			NumberVal(3),
		)
		cf := evalCompile(t, "MINA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("error direct arg propagates", func(t *testing.T) {
		cf := evalCompile(t, "MINA(1/0)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})

	t.Run("all positive numbers", func(t *testing.T) {
		cf := evalCompile(t, "MINA(10,20,5)")
		got, err := Eval(cf, &mockResolver{}, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("Excel doc example", func(t *testing.T) {
		// {FALSE, 0.2, 0.5, 0.4, 0.8} => min is FALSE=0
		resolver := valResolver(
			BoolVal(false),
			NumberVal(0.2),
			NumberVal(0.5),
			NumberVal(0.4),
			NumberVal(0.8),
		)
		cf := evalCompile(t, "MINA(A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})

	t.Run("mixed range and direct args", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(2),
			NumberVal(3),
		)
		cf := evalCompile(t, "MINA(A1:A2,-10)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -10 {
			t.Errorf("got %v, want -10", got)
		}
	})

	t.Run("empty cells in range ignored", func(t *testing.T) {
		resolver := valResolver(
			NumberVal(5),
			Value{Type: ValueEmpty},
			NumberVal(3),
		)
		cf := evalCompile(t, "MINA(A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})
}

// ---------------------------------------------------------------------------
// RANK.EQ
// ---------------------------------------------------------------------------

func TestRANKEQ(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(7),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 1, Row: 3}: NumberVal(5),
			{Col: 1, Row: 4}: NumberVal(3),
			{Col: 1, Row: 5}: NumberVal(9),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    float64
		wantErr ErrorValue
	}{
		{name: "desc_top", formula: "RANK.EQ(9,A1:A5)", want: 1},
		{name: "desc_mid", formula: "RANK.EQ(7,A1:A5)", want: 2},
		{name: "desc_tie", formula: "RANK.EQ(3,A1:A5)", want: 4},
		{name: "asc_bottom", formula: "RANK.EQ(3,A1:A5,1)", want: 1},
		{name: "asc_top", formula: "RANK.EQ(9,A1:A5,1)", want: 5},
		{name: "not_found", formula: "RANK.EQ(99,A1:A5)", wantErr: ErrValNA},
		{name: "too_few_args", formula: "RANK.EQ(9)", wantErr: ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr != 0 {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber || got.Num != tt.want {
				t.Errorf("got %v, want %g", got, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// RANK.AVG
// ---------------------------------------------------------------------------

func TestRANKAVG(t *testing.T) {
	// Dataset: {7, 3, 5, 3, 9}
	// Descending sorted: 9(1), 7(2), 5(3), 3(4), 3(5)
	// Ascending sorted:  3(1), 3(2), 5(3), 7(4), 9(5)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(7),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 1, Row: 3}: NumberVal(5),
			{Col: 1, Row: 4}: NumberVal(3),
			{Col: 1, Row: 5}: NumberVal(9),
		},
	}

	// Dataset with triple tie: {4, 4, 4, 1, 6}
	// Descending: 6(1), 4(2), 4(3), 4(4), 1(5)
	tripleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 2, Row: 3}: NumberVal(4),
			{Col: 2, Row: 4}: NumberVal(1),
			{Col: 2, Row: 5}: NumberVal(6),
		},
	}

	// All same values: {5, 5, 5}
	allSameResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 1}: NumberVal(5),
			{Col: 3, Row: 2}: NumberVal(5),
			{Col: 3, Row: 3}: NumberVal(5),
		},
	}

	// Single value: {10}
	singleResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 1}: NumberVal(10),
		},
	}

	// Negative values: {-3, -1, -3, 2, 0}
	negResolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 5, Row: 1}: NumberVal(-3),
			{Col: 5, Row: 2}: NumberVal(-1),
			{Col: 5, Row: 3}: NumberVal(-3),
			{Col: 5, Row: 4}: NumberVal(2),
			{Col: 5, Row: 5}: NumberVal(0),
		},
	}

	tests := []struct {
		name     string
		formula  string
		resolver CellResolver
		want     float64
		wantErr  ErrorValue
	}{
		// Basic descending (no ties)
		{name: "desc_unique_top", formula: "RANK.AVG(9,A1:A5)", resolver: resolver, want: 1},
		{name: "desc_unique_mid", formula: "RANK.AVG(7,A1:A5)", resolver: resolver, want: 2},
		{name: "desc_unique_5", formula: "RANK.AVG(5,A1:A5)", resolver: resolver, want: 3},

		// Descending with ties: two 3s occupy ranks 4 and 5, avg = 4.5
		{name: "desc_tie_avg", formula: "RANK.AVG(3,A1:A5)", resolver: resolver, want: 4.5},

		// Ascending (no ties)
		{name: "asc_unique_top", formula: "RANK.AVG(9,A1:A5,1)", resolver: resolver, want: 5},
		{name: "asc_unique_mid", formula: "RANK.AVG(7,A1:A5,1)", resolver: resolver, want: 4},

		// Ascending with ties: two 3s occupy ranks 1 and 2, avg = 1.5
		{name: "asc_tie_avg", formula: "RANK.AVG(3,A1:A5,1)", resolver: resolver, want: 1.5},

		// Triple tie descending: three 4s occupy ranks 2,3,4 -> avg = 3
		{name: "triple_tie_desc", formula: "RANK.AVG(4,B1:B5)", resolver: tripleResolver, want: 3},

		// Triple tie ascending: three 4s occupy ranks 2,3,4 -> avg = 3
		{name: "triple_tie_asc", formula: "RANK.AVG(4,B1:B5,1)", resolver: tripleResolver, want: 3},

		// No ties in triple dataset
		{name: "triple_unique_top_desc", formula: "RANK.AVG(6,B1:B5)", resolver: tripleResolver, want: 1},
		{name: "triple_unique_bottom_desc", formula: "RANK.AVG(1,B1:B5)", resolver: tripleResolver, want: 5},

		// All same values: three 5s occupy ranks 1,2,3 -> avg = 2
		{name: "all_same_desc", formula: "RANK.AVG(5,C1:C3)", resolver: allSameResolver, want: 2},
		{name: "all_same_asc", formula: "RANK.AVG(5,C1:C3,1)", resolver: allSameResolver, want: 2},

		// Single value
		{name: "single_value", formula: "RANK.AVG(10,D1:D1)", resolver: singleResolver, want: 1},

		// Value not found
		{name: "not_found", formula: "RANK.AVG(99,A1:A5)", resolver: resolver, wantErr: ErrValNA},

		// Negative values descending: sorted desc: 2(1), 0(2), -1(3), -3(4), -3(5)
		{name: "neg_no_tie_desc", formula: "RANK.AVG(2,E1:E5)", resolver: negResolver, want: 1},
		{name: "neg_tie_desc", formula: "RANK.AVG(-3,E1:E5)", resolver: negResolver, want: 4.5},

		// Negative values ascending: sorted asc: -3(1), -3(2), -1(3), 0(4), 2(5)
		{name: "neg_tie_asc", formula: "RANK.AVG(-3,E1:E5,1)", resolver: negResolver, want: 1.5},
		{name: "neg_no_tie_asc", formula: "RANK.AVG(2,E1:E5,1)", resolver: negResolver, want: 5},

		// order=0 means descending (same as omitting)
		{name: "order_zero_desc", formula: "RANK.AVG(3,A1:A5,0)", resolver: resolver, want: 4.5},

		// Wrong argument count
		{name: "too_few_args", formula: "RANK.AVG(9)", resolver: resolver, wantErr: ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, tt.resolver, nil)
			if err != nil {
				t.Fatalf("Eval error: %v", err)
			}
			if tt.wantErr != 0 {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("got %v, want error %v", got, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber || got.Num != tt.want {
				t.Errorf("got %v, want %g", got, tt.want)
			}
		})
	}
}
