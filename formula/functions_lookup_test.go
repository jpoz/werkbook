package formula

import (
	"testing"
)

func TestVLOOKUP(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: StringVal("one"),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 2, Row: 2}: StringVal("two"),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 3}: StringVal("three"),
		},
	}

	cf := evalCompile(t, "VLOOKUP(2,A1:B3,2,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "two" {
		t.Errorf("VLOOKUP exact: got %v, want two", got)
	}

	// Not found
	cf = evalCompile(t, "VLOOKUP(5,A1:B3,2,FALSE)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("VLOOKUP not found: got %v, want #N/A", got)
	}
}

func TestHLOOKUP(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 3, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: StringVal("a"),
			{Col: 2, Row: 2}: StringVal("b"),
			{Col: 3, Row: 2}: StringVal("c"),
		},
	}

	cf := evalCompile(t, "HLOOKUP(2,A1:C2,2,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "b" {
		t.Errorf("HLOOKUP: got %v, want b", got)
	}
}

func TestINDEX(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
			{Col: 1, Row: 2}: NumberVal(30),
			{Col: 2, Row: 2}: NumberVal(40),
		},
	}

	cf := evalCompile(t, "INDEX(A1:B2,2,2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 40 {
		t.Errorf("INDEX: got %g, want 40", got.Num)
	}
}

func TestMATCH(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}

	cf := evalCompile(t, "MATCH(20,A1:A3,0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("MATCH: got %g, want 2", got.Num)
	}
}

// ---------------------------------------------------------------------------
// VLOOKUP edge cases
// ---------------------------------------------------------------------------

func TestVLOOKUPApproximateMatch(t *testing.T) {
	// Sorted data for approximate match (default behavior)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: StringVal("ten"),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 2, Row: 2}: StringVal("twenty"),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 3}: StringVal("thirty"),
		},
	}

	// Approximate match: lookup 25 should find 20 (last value <= 25)
	cf := evalCompile(t, "VLOOKUP(25,A1:B3,2,TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "twenty" {
		t.Errorf("VLOOKUP approx 25: got %v, want twenty", got)
	}

	// Approximate match: exact value
	cf = evalCompile(t, "VLOOKUP(20,A1:B3,2,TRUE)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "twenty" {
		t.Errorf("VLOOKUP approx exact: got %v, want twenty", got)
	}

	// Approximate match: value less than all => #N/A
	cf = evalCompile(t, "VLOOKUP(5,A1:B3,2,TRUE)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("VLOOKUP approx too small: got %v, want #N/A", got)
	}

	// Default (no 4th arg) is approximate match
	cf = evalCompile(t, "VLOOKUP(25,A1:B3,2)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "twenty" {
		t.Errorf("VLOOKUP default mode: got %v, want twenty", got)
	}
}

func TestVLOOKUPColIndexOutOfRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: StringVal("one"),
		},
	}

	// col_index > number of columns in range => #REF!
	cf := evalCompile(t, "VLOOKUP(1,A1:B1,5,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("VLOOKUP col out of range: got %v, want #REF!", got)
	}

	// col_index < 1 => #VALUE!
	cf = evalCompile(t, "VLOOKUP(1,A1:B1,0,FALSE)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("VLOOKUP col 0: got %v, want #VALUE!", got)
	}
}

func TestVLOOKUPStringKeys(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple"),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: StringVal("banana"),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: StringVal("cherry"),
			{Col: 2, Row: 3}: NumberVal(3),
		},
	}

	// Case-insensitive string matching
	cf := evalCompile(t, `VLOOKUP("BANANA",A1:B3,2,FALSE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("VLOOKUP case insensitive: got %v, want 2", got)
	}
}

func TestVLOOKUPStringKeyExactMatch(t *testing.T) {
	// Mirrors the multisheet edge case spec: look up "veggie" in a category/value table
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("fruit"),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: StringVal("veggie"),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: StringVal("grain"),
			{Col: 2, Row: 3}: NumberVal(30),
		},
	}

	// exact match (4th arg = 0) for "veggie" => 20
	cf := evalCompile(t, `VLOOKUP("veggie",A1:B3,2,0)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("VLOOKUP veggie exact: got %v, want 20", got)
	}

	// first entry
	cf = evalCompile(t, `VLOOKUP("fruit",A1:B3,2,0)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("VLOOKUP fruit exact: got %v, want 10", got)
	}

	// last entry
	cf = evalCompile(t, `VLOOKUP("grain",A1:B3,2,0)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("VLOOKUP grain exact: got %v, want 30", got)
	}

	// not found => #N/A
	cf = evalCompile(t, `VLOOKUP("dairy",A1:B3,2,0)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("VLOOKUP dairy not found: got %v, want #N/A", got)
	}
}

func TestVLOOKUPArgErrors(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "VLOOKUP(1,A1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("VLOOKUP too few args: got %v, want #VALUE!", got)
	}
}

// ---------------------------------------------------------------------------
// HLOOKUP edge cases
// ---------------------------------------------------------------------------

func TestHLOOKUPApproximateMatch(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
			{Col: 3, Row: 1}: NumberVal(30),
			{Col: 1, Row: 2}: StringVal("a"),
			{Col: 2, Row: 2}: StringVal("b"),
			{Col: 3, Row: 2}: StringVal("c"),
		},
	}

	cf := evalCompile(t, "HLOOKUP(25,A1:C2,2,TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "b" {
		t.Errorf("HLOOKUP approx: got %v, want b", got)
	}
}

func TestHLOOKUPNotFound(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: StringVal("a"),
			{Col: 2, Row: 2}: StringVal("b"),
		},
	}

	cf := evalCompile(t, "HLOOKUP(5,A1:B2,2,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("HLOOKUP not found: got %v, want #N/A", got)
	}
}

func TestHLOOKUPRowIndexOutOfRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: StringVal("a"),
			{Col: 2, Row: 2}: StringVal("b"),
		},
	}

	cf := evalCompile(t, "HLOOKUP(1,A1:B2,5,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("HLOOKUP row out of range: got %v, want #REF!", got)
	}
}

// ---------------------------------------------------------------------------
// MATCH edge cases — all match types
// ---------------------------------------------------------------------------

func TestMATCHExact(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple"),
			{Col: 1, Row: 2}: StringVal("banana"),
			{Col: 1, Row: 3}: StringVal("cherry"),
		},
	}

	cf := evalCompile(t, `MATCH("banana",A1:A3,0)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("MATCH exact: got %g, want 2", got.Num)
	}

	// Not found
	cf = evalCompile(t, `MATCH("date",A1:A3,0)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("MATCH not found: got %v, want #N/A", got)
	}
}

func TestMATCHAscending(t *testing.T) {
	// match_type=1 (ascending, default)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: NumberVal(40),
		},
	}

	// Exact match in ascending
	cf := evalCompile(t, "MATCH(30,A1:A4,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("MATCH asc exact: got %g, want 3", got.Num)
	}

	// Between values: 25 => position of 20 (last <=)
	cf = evalCompile(t, "MATCH(25,A1:A4,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("MATCH asc between: got %g, want 2", got.Num)
	}

	// Value smaller than all => #N/A
	cf = evalCompile(t, "MATCH(5,A1:A4,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("MATCH asc too small: got %v, want #N/A", got)
	}

	// Default match_type is 1
	cf = evalCompile(t, "MATCH(25,A1:A4)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("MATCH default: got %g, want 2", got.Num)
	}
}

func TestMATCHDescending(t *testing.T) {
	// match_type=-1 (descending)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(40),
			{Col: 1, Row: 2}: NumberVal(30),
			{Col: 1, Row: 3}: NumberVal(20),
			{Col: 1, Row: 4}: NumberVal(10),
		},
	}

	cf := evalCompile(t, "MATCH(25,A1:A4,-1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("MATCH desc between: got %g, want 2", got.Num)
	}

	// Value larger than all => #N/A
	cf = evalCompile(t, "MATCH(50,A1:A4,-1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("MATCH desc too large: got %v, want #N/A", got)
	}
}

// ---------------------------------------------------------------------------
// INDEX edge cases
// ---------------------------------------------------------------------------

func TestINDEXEdgeCases(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
			{Col: 1, Row: 2}: NumberVal(30),
			{Col: 2, Row: 2}: NumberVal(40),
		},
	}

	// First cell
	cf := evalCompile(t, "INDEX(A1:B2,1,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("INDEX(1,1): got %g, want 10", got.Num)
	}

	// Row out of range => #REF!
	cf = evalCompile(t, "INDEX(A1:B2,5,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("INDEX row OOB: got %v, want #REF!", got)
	}

	// Col out of range => #REF!
	cf = evalCompile(t, "INDEX(A1:B2,1,5)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("INDEX col OOB: got %v, want #REF!", got)
	}

	// Two-arg form (row only, col defaults to 0 which is first col)
	cf = evalCompile(t, "INDEX(A1:B2,2)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("INDEX 2-arg: got %g, want 30", got.Num)
	}
}

// ---------------------------------------------------------------------------
// INDEX + MATCH combo
// ---------------------------------------------------------------------------

func TestINDEXMATCHCombo(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple"),
			{Col: 1, Row: 2}: StringVal("banana"),
			{Col: 1, Row: 3}: StringVal("cherry"),
			{Col: 2, Row: 1}: NumberVal(100),
			{Col: 2, Row: 2}: NumberVal(200),
			{Col: 2, Row: 3}: NumberVal(300),
		},
	}

	cf := evalCompile(t, `INDEX(B1:B3,MATCH("banana",A1:A3,0))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 200 {
		t.Errorf("INDEX/MATCH: got %g, want 200", got.Num)
	}
}

// ---------------------------------------------------------------------------
// INDIRECT tests
// ---------------------------------------------------------------------------

func TestINDIRECTSingleCell(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(42),
			{Col: 2, Row: 3}: StringVal("hello"),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	cf := evalCompile(t, `INDIRECT("A1")`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("INDIRECT(A1): got %v, want 42", got)
	}

	cf = evalCompile(t, `INDIRECT("B3")`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "hello" {
		t.Errorf("INDIRECT(B3): got %v, want hello", got)
	}
}

func TestINDIRECTWithDollarSigns(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(99),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	cf := evalCompile(t, `INDIRECT("$A$1")`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 99 {
		t.Errorf("INDIRECT($A$1): got %v, want 99", got)
	}
}

func TestINDIRECTRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	cf := evalCompile(t, `INDIRECT("A1:A3")`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("INDIRECT(A1:A3): expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("INDIRECT(A1:A3): expected 3 rows, got %d", len(got.Array))
	}
	for i, want := range []float64{1, 2, 3} {
		if got.Array[i][0].Num != want {
			t.Errorf("INDIRECT(A1:A3)[%d]: got %g, want %g", i, got.Array[i][0].Num, want)
		}
	}
	if got.RangeOrigin == nil {
		t.Error("INDIRECT(A1:A3): expected RangeOrigin to be set")
	}
}

func TestINDIRECTRowRange(t *testing.T) {
	// INDIRECT("1:3") creates a full-row range from row 1 to 3.
	// ROW(INDIRECT("1:3")) should produce {1,2,3} in array context.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}
	ctx := &EvalContext{Resolver: resolver, IsArrayFormula: true}

	cf := evalCompile(t, `INDIRECT("1:3")`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("INDIRECT(1:3): expected array, got %v", got.Type)
	}
	if got.RangeOrigin == nil {
		t.Fatal("INDIRECT(1:3): expected RangeOrigin to be set")
	}
	if got.RangeOrigin.FromRow != 1 || got.RangeOrigin.ToRow != 3 {
		t.Errorf("INDIRECT(1:3): range rows = %d:%d, want 1:3",
			got.RangeOrigin.FromRow, got.RangeOrigin.ToRow)
	}
	if got.RangeOrigin.FromCol != 1 || got.RangeOrigin.ToCol != maxExcelCols {
		t.Errorf("INDIRECT(1:3): range cols = %d:%d, want 1:%d",
			got.RangeOrigin.FromCol, got.RangeOrigin.ToCol, maxExcelCols)
	}
}

func TestINDIRECTRowRangeWithROW(t *testing.T) {
	// The critical pattern from bond pricing: ROW(INDIRECT("1:5"))
	resolver := &mockResolver{}
	ctx := &EvalContext{Resolver: resolver, IsArrayFormula: true}

	cf := evalCompile(t, `ROW(INDIRECT("1:5"))`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("ROW(INDIRECT(1:5)): expected array, got %v", got.Type)
	}
	if len(got.Array) != 5 {
		t.Fatalf("ROW(INDIRECT(1:5)): expected 5 rows, got %d", len(got.Array))
	}
	for i := 0; i < 5; i++ {
		want := float64(i + 1)
		if got.Array[i][0].Num != want {
			t.Errorf("ROW(INDIRECT(1:5))[%d]: got %g, want %g", i, got.Array[i][0].Num, want)
		}
	}
}

func TestINDIRECTEmptyString(t *testing.T) {
	resolver := &mockResolver{}
	ctx := &EvalContext{Resolver: resolver}

	cf := evalCompile(t, `INDIRECT("")`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("INDIRECT empty: got %v, want #REF!", got)
	}
}

func TestINDIRECTWithSheetName(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Sheet: "Sheet2", Col: 1, Row: 1}: NumberVal(77),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	cf := evalCompile(t, `INDIRECT("Sheet2!A1")`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 77 {
		t.Errorf("INDIRECT(Sheet2!A1): got %v, want 77", got)
	}
}

func TestINDIRECTDynamic(t *testing.T) {
	// Test INDIRECT with a dynamically constructed reference: INDIRECT("A"&"1")
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(55),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	cf := evalCompile(t, `INDIRECT("A"&"1")`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 55 {
		t.Errorf(`INDIRECT("A"&"1"): got %v, want 55`, got)
	}
}

// ---------------------------------------------------------------------------
// TRANSPOSE tests
// ---------------------------------------------------------------------------

func TestTRANSPOSE_2x3(t *testing.T) {
	// 2 rows x 3 cols → 3 rows x 2 cols
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 3, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 3, Row: 2}: NumberVal(6),
		},
	}

	cf := evalCompile(t, "TRANSPOSE(A1:C2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	if len(got.Array[0]) != 2 {
		t.Fatalf("expected 2 cols, got %d", len(got.Array[0]))
	}
	// result[0] = {1, 4}, result[1] = {2, 5}, result[2] = {3, 6}
	want := [][]float64{{1, 4}, {2, 5}, {3, 6}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestTRANSPOSE_3x2(t *testing.T) {
	// 3 rows x 2 cols → 2 rows x 3 cols
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 1, Row: 3}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}

	cf := evalCompile(t, "TRANSPOSE(A1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(got.Array))
	}
	if len(got.Array[0]) != 3 {
		t.Fatalf("expected 3 cols, got %d", len(got.Array[0]))
	}
	want := [][]float64{{1, 3, 5}, {2, 4, 6}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestTRANSPOSE_SingleRow(t *testing.T) {
	// 1 row x 3 cols → 3 rows x 1 col
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
			{Col: 3, Row: 1}: NumberVal(30),
		},
	}

	cf := evalCompile(t, "TRANSPOSE(A1:C1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	for i, w := range []float64{10, 20, 30} {
		if len(got.Array[i]) != 1 {
			t.Fatalf("row %d: expected 1 col, got %d", i, len(got.Array[i]))
		}
		if got.Array[i][0].Num != w {
			t.Errorf("row %d: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestTRANSPOSE_SingleColumn(t *testing.T) {
	// 3 rows x 1 col → 1 row x 3 cols
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}

	cf := evalCompile(t, "TRANSPOSE(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 1 {
		t.Fatalf("expected 1 row, got %d", len(got.Array))
	}
	if len(got.Array[0]) != 3 {
		t.Fatalf("expected 3 cols, got %d", len(got.Array[0]))
	}
	for i, w := range []float64{10, 20, 30} {
		if got.Array[0][i].Num != w {
			t.Errorf("col %d: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestTRANSPOSE_1x1(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(42),
		},
	}

	cf := evalCompile(t, "TRANSPOSE(A1:A1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 1x1, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[0][0].Num != 42 {
		t.Errorf("got %g, want 42", got.Array[0][0].Num)
	}
}

func TestTRANSPOSE_ScalarNumber(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "TRANSPOSE(5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("scalar number: got %v, want 5", got)
	}
}

func TestTRANSPOSE_ScalarString(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TRANSPOSE("hello")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "hello" {
		t.Errorf("scalar string: got %v, want hello", got)
	}
}

func TestTRANSPOSE_ScalarBool(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "TRANSPOSE(TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("scalar bool: got %v, want TRUE", got)
	}
}

func TestTRANSPOSE_MixedTypes(t *testing.T) {
	// 2x2 array with mixed types
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: BoolVal(true),
			{Col: 2, Row: 2}: NumberVal(2),
		},
	}

	cf := evalCompile(t, "TRANSPOSE(A1:B2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	// Transposed: row0={1, TRUE}, row1={"a", 2}
	if got.Array[0][0].Type != ValueNumber || got.Array[0][0].Num != 1 {
		t.Errorf("[0][0]: got %v, want 1", got.Array[0][0])
	}
	if got.Array[0][1].Type != ValueBool || !got.Array[0][1].Bool {
		t.Errorf("[0][1]: got %v, want TRUE", got.Array[0][1])
	}
	if got.Array[1][0].Type != ValueString || got.Array[1][0].Str != "a" {
		t.Errorf("[1][0]: got %v, want a", got.Array[1][0])
	}
	if got.Array[1][1].Type != ValueNumber || got.Array[1][1].Num != 2 {
		t.Errorf("[1][1]: got %v, want 2", got.Array[1][1])
	}
}

func TestTRANSPOSE_ErrorPreserved(t *testing.T) {
	// An error value in the array should be preserved after transpose
	v := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(1), ErrorVal(ErrValDIV0)},
			{NumberVal(3), NumberVal(4)},
		},
	}
	result, err := fnTRANSPOSE([]Value{v})
	if err != nil {
		t.Fatalf("fnTRANSPOSE: %v", err)
	}
	if result.Type != ValueArray {
		t.Fatalf("expected array, got %v", result.Type)
	}
	// Transposed: row0={1, 3}, row1={#DIV/0!, 4}
	if result.Array[1][0].Type != ValueError || result.Array[1][0].Err != ErrValDIV0 {
		t.Errorf("[1][0]: got %v, want #DIV/0!", result.Array[1][0])
	}
}

func TestTRANSPOSE_NoArgs(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "TRANSPOSE()")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("no args: got %v, want #VALUE!", got)
	}
}

func TestTRANSPOSE_TooManyArgs(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "TRANSPOSE(A1:B2,A1:B2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("too many args: got %v, want #VALUE!", got)
	}
}

func TestTRANSPOSE_SingleCellRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 3}: NumberVal(99),
		},
	}

	cf := evalCompile(t, "TRANSPOSE(B3:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if got.Array[0][0].Num != 99 {
		t.Errorf("got %g, want 99", got.Array[0][0].Num)
	}
}

func TestTRANSPOSE_LargeArray4x5(t *testing.T) {
	// 4 rows x 5 cols → 5 rows x 4 cols
	resolver := &mockResolver{
		cells: map[CellAddr]Value{},
	}
	for r := 1; r <= 4; r++ {
		for c := 1; c <= 5; c++ {
			resolver.cells[CellAddr{Col: c, Row: r}] = NumberVal(float64(r*10 + c))
		}
	}

	cf := evalCompile(t, "TRANSPOSE(A1:E4)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 5 {
		t.Fatalf("expected 5 rows, got %d", len(got.Array))
	}
	if len(got.Array[0]) != 4 {
		t.Fatalf("expected 4 cols, got %d", len(got.Array[0]))
	}
	for origR := 1; origR <= 4; origR++ {
		for origC := 1; origC <= 5; origC++ {
			want := float64(origR*10 + origC)
			// In transposed: row=origC-1, col=origR-1
			g := got.Array[origC-1][origR-1].Num
			if g != want {
				t.Errorf("[%d][%d]: got %g, want %g", origC-1, origR-1, g, want)
			}
		}
	}
}

func TestTRANSPOSE_DoubleTranspose(t *testing.T) {
	// TRANSPOSE(TRANSPOSE(x)) should return the original
	inner := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
		},
	}
	first, err := fnTRANSPOSE([]Value{inner})
	if err != nil {
		t.Fatalf("first transpose: %v", err)
	}
	second, err := fnTRANSPOSE([]Value{first})
	if err != nil {
		t.Fatalf("second transpose: %v", err)
	}
	if len(second.Array) != 2 || len(second.Array[0]) != 3 {
		t.Fatalf("double transpose: expected 2x3, got %dx%d", len(second.Array), len(second.Array[0]))
	}
	want := [][]float64{{1, 2, 3}, {4, 5, 6}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if second.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, second.Array[r][c].Num, w)
			}
		}
	}
}

func TestTRANSPOSE_EmptyArray(t *testing.T) {
	v := Value{Type: ValueArray, Array: [][]Value{}}
	result, err := fnTRANSPOSE([]Value{v})
	if err != nil {
		t.Fatalf("fnTRANSPOSE: %v", err)
	}
	if result.Type != ValueArray {
		t.Fatalf("expected array, got %v", result.Type)
	}
	if len(result.Array) != 0 {
		t.Errorf("expected empty array, got %d rows", len(result.Array))
	}
}

func TestTRANSPOSE_AllStrings(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("a"),
			{Col: 2, Row: 1}: StringVal("b"),
			{Col: 1, Row: 2}: StringVal("c"),
			{Col: 2, Row: 2}: StringVal("d"),
		},
	}

	cf := evalCompile(t, "TRANSPOSE(A1:B2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	// Transposed: row0={"a","c"}, row1={"b","d"}
	if got.Array[0][0].Str != "a" || got.Array[0][1].Str != "c" {
		t.Errorf("row 0: got %v %v, want a c", got.Array[0][0], got.Array[0][1])
	}
	if got.Array[1][0].Str != "b" || got.Array[1][1].Str != "d" {
		t.Errorf("row 1: got %v %v, want b d", got.Array[1][0], got.Array[1][1])
	}
}

func TestTRANSPOSE_WithEmptyCells(t *testing.T) {
	// Only some cells filled; empty cells should appear as empty values
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			// B1 is empty
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 2, Row: 2}: NumberVal(4),
		},
	}

	cf := evalCompile(t, "TRANSPOSE(A1:B2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	// Original: {{1, 0}, {3, 4}} → Transposed: {{1, 3}, {0, 4}}
	if got.Array[0][0].Num != 1 {
		t.Errorf("[0][0]: got %g, want 1", got.Array[0][0].Num)
	}
	if got.Array[0][1].Num != 3 {
		t.Errorf("[0][1]: got %g, want 3", got.Array[0][1].Num)
	}
	if got.Array[1][1].Num != 4 {
		t.Errorf("[1][1]: got %g, want 4", got.Array[1][1].Num)
	}
}

func TestTRANSPOSE_MultipleErrors(t *testing.T) {
	// Multiple error values should all be preserved
	v := Value{
		Type: ValueArray,
		Array: [][]Value{
			{ErrorVal(ErrValNA), ErrorVal(ErrValDIV0)},
			{ErrorVal(ErrValVALUE), ErrorVal(ErrValREF)},
		},
	}
	result, err := fnTRANSPOSE([]Value{v})
	if err != nil {
		t.Fatalf("fnTRANSPOSE: %v", err)
	}
	// Transposed: row0={#N/A, #VALUE!}, row1={#DIV/0!, #REF!}
	if result.Array[0][0].Err != ErrValNA {
		t.Errorf("[0][0]: got %v, want #N/A", result.Array[0][0])
	}
	if result.Array[0][1].Err != ErrValVALUE {
		t.Errorf("[0][1]: got %v, want #VALUE!", result.Array[0][1])
	}
	if result.Array[1][0].Err != ErrValDIV0 {
		t.Errorf("[1][0]: got %v, want #DIV/0!", result.Array[1][0])
	}
	if result.Array[1][1].Err != ErrValREF {
		t.Errorf("[1][1]: got %v, want #REF!", result.Array[1][1])
	}
}

func TestTRANSPOSE_SquareMatrix(t *testing.T) {
	// 3x3 square matrix
	v := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
			{NumberVal(7), NumberVal(8), NumberVal(9)},
		},
	}
	result, err := fnTRANSPOSE([]Value{v})
	if err != nil {
		t.Fatalf("fnTRANSPOSE: %v", err)
	}
	want := [][]float64{{1, 4, 7}, {2, 5, 8}, {3, 6, 9}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if result.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, result.Array[r][c].Num, w)
			}
		}
	}
}
