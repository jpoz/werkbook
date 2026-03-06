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

func TestHLOOKUPRowIndex1ReturnsHeaderRow(t *testing.T) {
	// row_index_num=1 should return from the first (lookup) row itself.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("Name"),
			{Col: 2, Row: 1}: StringVal("Age"),
			{Col: 3, Row: 1}: StringVal("City"),
			{Col: 1, Row: 2}: StringVal("Alice"),
			{Col: 2, Row: 2}: NumberVal(30),
			{Col: 3, Row: 2}: StringVal("NYC"),
		},
	}

	cf := evalCompile(t, `HLOOKUP("Age",A1:C2,1,FALSE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "Age" {
		t.Errorf("HLOOKUP row_index=1: got %v, want Age", got)
	}
}

func TestHLOOKUPStringLookup(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("Axles"),
			{Col: 2, Row: 1}: StringVal("Bearings"),
			{Col: 3, Row: 1}: StringVal("Bolts"),
			{Col: 1, Row: 2}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(7),
			{Col: 3, Row: 2}: NumberVal(10),
		},
	}

	cf := evalCompile(t, `HLOOKUP("Bearings",A1:C2,2,FALSE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 7 {
		t.Errorf("HLOOKUP string lookup: got %v, want 7", got)
	}
}

func TestHLOOKUPCaseInsensitive(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("Apple"),
			{Col: 2, Row: 1}: StringVal("Banana"),
			{Col: 3, Row: 1}: StringVal("Cherry"),
			{Col: 1, Row: 2}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 3, Row: 2}: NumberVal(3),
		},
	}

	// Lookup "banana" (lowercase) should match "Banana"
	cf := evalCompile(t, `HLOOKUP("banana",A1:C2,2,FALSE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("HLOOKUP case insensitive: got %v, want 2", got)
	}

	// Also try all-caps
	cf = evalCompile(t, `HLOOKUP("CHERRY",A1:C2,2,FALSE)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("HLOOKUP case insensitive caps: got %v, want 3", got)
	}
}

func TestHLOOKUPMatchInMiddle(t *testing.T) {
	// 5 columns, match in the middle (col 3)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
			{Col: 3, Row: 1}: NumberVal(30),
			{Col: 4, Row: 1}: NumberVal(40),
			{Col: 5, Row: 1}: NumberVal(50),
			{Col: 1, Row: 2}: StringVal("a"),
			{Col: 2, Row: 2}: StringVal("b"),
			{Col: 3, Row: 2}: StringVal("c"),
			{Col: 4, Row: 2}: StringVal("d"),
			{Col: 5, Row: 2}: StringVal("e"),
		},
	}

	cf := evalCompile(t, "HLOOKUP(30,A1:E2,2,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "c" {
		t.Errorf("HLOOKUP match in middle: got %v, want c", got)
	}
}

func TestHLOOKUPFirstMatchWins(t *testing.T) {
	// Duplicate values in header row; first match should win.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 3, Row: 1}: NumberVal(2),
			{Col: 4, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: StringVal("a"),
			{Col: 2, Row: 2}: StringVal("b"),
			{Col: 3, Row: 2}: StringVal("c"),
			{Col: 4, Row: 2}: StringVal("d"),
		},
	}

	cf := evalCompile(t, "HLOOKUP(2,A1:D2,2,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "b" {
		t.Errorf("HLOOKUP first match wins: got %v, want b", got)
	}
}

func TestHLOOKUPApproxExactHit(t *testing.T) {
	// Approximate match where lookup value exactly matches a header value.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
			{Col: 3, Row: 1}: NumberVal(30),
			{Col: 1, Row: 2}: StringVal("ten"),
			{Col: 2, Row: 2}: StringVal("twenty"),
			{Col: 3, Row: 2}: StringVal("thirty"),
		},
	}

	cf := evalCompile(t, "HLOOKUP(20,A1:C2,2,TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "twenty" {
		t.Errorf("HLOOKUP approx exact hit: got %v, want twenty", got)
	}
}

func TestHLOOKUPApproxDefaultOmitted(t *testing.T) {
	// Omitting range_lookup should default to approximate match (TRUE).
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

	cf := evalCompile(t, "HLOOKUP(25,A1:C2,2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "b" {
		t.Errorf("HLOOKUP default approx: got %v, want b", got)
	}
}

func TestHLOOKUPApproxLessThanAll(t *testing.T) {
	// Approximate match with lookup value smaller than all header values → #N/A.
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

	cf := evalCompile(t, "HLOOKUP(5,A1:C2,2,TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("HLOOKUP approx less than all: got %v, want #N/A", got)
	}
}

func TestHLOOKUPUnsortedExact(t *testing.T) {
	// Unsorted header row with exact match should still find the value.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(30),
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 3, Row: 1}: NumberVal(20),
			{Col: 1, Row: 2}: StringVal("thirty"),
			{Col: 2, Row: 2}: StringVal("ten"),
			{Col: 3, Row: 2}: StringVal("twenty"),
		},
	}

	cf := evalCompile(t, "HLOOKUP(10,A1:C2,2,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "ten" {
		t.Errorf("HLOOKUP unsorted exact: got %v, want ten", got)
	}
}

func TestHLOOKUPArgErrors(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args (2 args)
	cf := evalCompile(t, "HLOOKUP(1,A1:C2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("HLOOKUP too few args: got %v, want #VALUE!", got)
	}

	// Too many args (5 args)
	cf = evalCompile(t, "HLOOKUP(1,A1:C2,2,FALSE,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("HLOOKUP too many args: got %v, want #VALUE!", got)
	}
}

func TestHLOOKUPRowIndexZero(t *testing.T) {
	// row_index_num = 0 → #REF! (index out of range)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: StringVal("a"),
			{Col: 2, Row: 2}: StringVal("b"),
		},
	}

	cf := evalCompile(t, "HLOOKUP(1,A1:B2,0,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("HLOOKUP row_index=0: got %v, want #REF!", got)
	}
}

func TestHLOOKUPNegativeRowIndex(t *testing.T) {
	// Negative row_index_num → #REF!
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: StringVal("a"),
			{Col: 2, Row: 2}: StringVal("b"),
		},
	}

	cf := evalCompile(t, "HLOOKUP(1,A1:B2,-1,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("HLOOKUP negative row_index: got %v, want #REF!", got)
	}
}

func TestHLOOKUPBooleanLookup(t *testing.T) {
	// Look up a boolean value in the header row.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: BoolVal(false),
			{Col: 2, Row: 1}: BoolVal(true),
			{Col: 1, Row: 2}: StringVal("no"),
			{Col: 2, Row: 2}: StringVal("yes"),
		},
	}

	cf := evalCompile(t, "HLOOKUP(TRUE,A1:B2,2,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "yes" {
		t.Errorf("HLOOKUP boolean lookup: got %v, want yes", got)
	}
}

func TestHLOOKUPMultipleRows(t *testing.T) {
	// Table with 3 rows; retrieve from the third row.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("X"),
			{Col: 2, Row: 1}: StringVal("Y"),
			{Col: 3, Row: 1}: StringVal("Z"),
			{Col: 1, Row: 2}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 3, Row: 2}: NumberVal(30),
			{Col: 1, Row: 3}: NumberVal(100),
			{Col: 2, Row: 3}: NumberVal(200),
			{Col: 3, Row: 3}: NumberVal(300),
		},
	}

	cf := evalCompile(t, `HLOOKUP("Y",A1:C3,3,FALSE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 200 {
		t.Errorf("HLOOKUP multiple rows: got %v, want 200", got)
	}
}

func TestHLOOKUPApproxLastColumn(t *testing.T) {
	// Approximate match should return the last column when lookup >= all values.
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

	cf := evalCompile(t, "HLOOKUP(99,A1:C2,2,TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "c" {
		t.Errorf("HLOOKUP approx last col: got %v, want c", got)
	}
}

func TestHLOOKUPNumberLookup(t *testing.T) {
	// Exact match for a number in the header row.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(100),
			{Col: 2, Row: 1}: NumberVal(200),
			{Col: 3, Row: 1}: NumberVal(300),
			{Col: 1, Row: 2}: StringVal("low"),
			{Col: 2, Row: 2}: StringVal("mid"),
			{Col: 3, Row: 2}: StringVal("high"),
		},
	}

	cf := evalCompile(t, "HLOOKUP(300,A1:C2,2,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "high" {
		t.Errorf("HLOOKUP number lookup: got %v, want high", got)
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

func TestMATCHDescendingUnsortedReturnsNA(t *testing.T) {
	// When match_type=-1 is used on unsorted data, Excel's binary search
	// typically returns #N/A. Our binary search should replicate that.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(30),
			{Col: 1, Row: 3}: NumberVal(5),
			{Col: 1, Row: 4}: NumberVal(25),
			{Col: 1, Row: 5}: NumberVal(15),
			{Col: 1, Row: 6}: NumberVal(20),
		},
	}

	cf := evalCompile(t, "MATCH(12,A1:A6,-1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("MATCH desc unsorted: got %v, want #N/A", got)
	}
}

func TestMATCHAscendingSkipsEmpty(t *testing.T) {
	// Simulate a whole-column ref where data is sparse: rows 1-3 have
	// sorted ascending values, rows 4-8 are empty. MATCH(matchType=1)
	// must skip the trailing empty cells rather than treating them as 0.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			// rows 4-8 are empty (not in map)
		},
	}

	// MATCH(20,A1:A8,1) should find row 2, not be confused by trailing empties
	cf := evalCompile(t, "MATCH(20,A1:A8,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("MATCH asc skip empty: got %v, want 2", got)
	}

	// MATCH(25,A1:A8,1) should find row 2 (last <= 25)
	cf = evalCompile(t, "MATCH(25,A1:A8,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("MATCH asc between skip empty: got %v, want 2", got)
	}
}

func TestMATCHDescendingSkipsEmpty(t *testing.T) {
	// Descending data with trailing empties.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(30),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(10),
			// rows 4-6 are empty
		},
	}

	cf := evalCompile(t, "MATCH(20,A1:A6,-1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("MATCH desc skip empty: got %v, want 2", got)
	}
}

func TestMATCHAscendingWithLeadingEmpty(t *testing.T) {
	// Leading empty row (e.g. header row is empty in lookup column),
	// followed by sorted data. MATCH should skip the empty and find the value.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// row 1 is empty
			{Col: 1, Row: 2}: NumberVal(10),
			{Col: 1, Row: 3}: NumberVal(20),
			{Col: 1, Row: 4}: NumberVal(30),
		},
	}

	cf := evalCompile(t, "MATCH(20,A1:A6,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("MATCH asc leading empty: got %v, want 3", got)
	}
}

func TestMATCHDefaultUnsortedStringsWithEmpty(t *testing.T) {
	// Real-world scenario from fa.xlsx: MATCH(A10,lfy!Q:Q) where Q:Q has
	// a header row, then unsorted string names, with leading empties.
	// match_type defaults to 1. The implementation must still find an
	// exact match among the non-empty values.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// rows 1-2 empty
			{Col: 1, Row: 3}: StringVal("LPs name"),           // header
			{Col: 1, Row: 4}: StringVal("Brian Schechter"),    // target
			{Col: 1, Row: 5}: StringVal("Foundation Capital"), // after target
			// rows 6-10 empty
		},
	}

	// Default match_type=1, lookup="Brian Schechter"
	cf := evalCompile(t, `MATCH("Brian Schechter",A1:A10)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("MATCH default unsorted strings: got %v, want 4", got)
	}
}

func TestINDEXMATCHWholeColumnPattern(t *testing.T) {
	// Simulates INDEX(D:D,MATCH(val,Q:Q)) with sparse data — the
	// pattern that was failing in the fa.xlsx audit.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Q column (col 17) — lookup array
			{Col: 17, Row: 3}: StringVal("LPs name"),
			{Col: 17, Row: 4}: StringVal("Brian Schechter"),
			{Col: 17, Row: 5}: StringVal("Foundation Capital"),
			// D column (col 4) — result array
			{Col: 4, Row: 3}: StringVal("Header"),
			{Col: 4, Row: 4}: NumberVal(1055),
			{Col: 4, Row: 5}: NumberVal(2000),
		},
	}

	cf := evalCompile(t, `INDEX(D1:D10,MATCH("Brian Schechter",Q1:Q10))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1055 {
		t.Errorf("INDEX/MATCH whole-column pattern: got %v, want 1055", got)
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

	// Two-arg form (row only, col defaults to 1 which is first col)
	cf = evalCompile(t, "INDEX(A1:B2,2)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("INDEX 2-arg: got %g, want 30", got.Num)
	}

	// row_num=0 returns entire column as an array. The caller
	// (formulaValueToValue) converts multi-element arrays to #VALUE!
	// in non-array formula cells.
	cf = evalCompile(t, "INDEX(A1:B2,0,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Errorf("INDEX row=0: got %v, want 2-row array", got)
	}

	// col_num=0 returns entire row as an array.
	cf = evalCompile(t, "INDEX(A1:B2,1,0)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 2 {
		t.Errorf("INDEX col=0: got %v, want 1x2 array", got)
	}

	// Negative row_num => #VALUE!
	cf = evalCompile(t, "INDEX(A1:B2,-1,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("INDEX negative row: got %v, want #VALUE!", got)
	}

	// Negative col_num => #VALUE!
	cf = evalCompile(t, "INDEX(A1:B2,1,-1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("INDEX negative col: got %v, want #VALUE!", got)
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
// INDIRECT R1C1-style tests
// ---------------------------------------------------------------------------

func TestINDIRECT_R1C1_SingleCell(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("Test"),
			{Col: 3, Row: 5}: NumberVal(99),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// R1C1 means row 1, col 1 = A1
	cf := evalCompile(t, `INDIRECT("R1C1",FALSE)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "Test" {
		t.Errorf(`INDIRECT("R1C1",FALSE): got %v, want "Test"`, got)
	}

	// R5C3 means row 5, col 3 = C5
	cf = evalCompile(t, `INDIRECT("R5C3",FALSE)`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 99 {
		t.Errorf(`INDIRECT("R5C3",FALSE): got %v, want 99`, got)
	}
}

func TestINDIRECT_R1C1_Range(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// R1C1:R3C1 = A1:A3
	cf := evalCompile(t, `INDIRECT("R1C1:R3C1",FALSE)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf(`INDIRECT("R1C1:R3C1",FALSE): expected array, got %v`, got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf(`INDIRECT("R1C1:R3C1",FALSE): expected 3 rows, got %d`, len(got.Array))
	}
	for i, want := range []float64{10, 20, 30} {
		if got.Array[i][0].Num != want {
			t.Errorf(`INDIRECT("R1C1:R3C1",FALSE)[%d]: got %g, want %g`, i, got.Array[i][0].Num, want)
		}
	}
}

func TestINDIRECT_R1C1_CaseInsensitive(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(42),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// lowercase r1c1 should also work
	cf := evalCompile(t, `INDIRECT("r1c1",FALSE)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf(`INDIRECT("r1c1",FALSE): got %v, want 42`, got)
	}
}

func TestINDIRECT_R1C1_Invalid(t *testing.T) {
	resolver := &mockResolver{cells: map[CellAddr]Value{}}
	ctx := &EvalContext{Resolver: resolver}

	// Invalid R1C1 reference should return #REF!
	cf := evalCompile(t, `INDIRECT("RXCY",FALSE)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf(`INDIRECT("RXCY",FALSE): got %v, want error`, got)
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

// ---------------------------------------------------------------------------
// UNIQUE tests
// ---------------------------------------------------------------------------

func TestUNIQUE_Basic1D(t *testing.T) {
	// UNIQUE({1;2;1;3;2}) = {1;2;3}
	got, err := fnUNIQUE([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}, {NumberVal(1)}, {NumberVal(3)}, {NumberVal(2)}},
	}})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	want := []float64{1, 2, 3}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestUNIQUE_Strings(t *testing.T) {
	// UNIQUE({"a";"b";"a";"c"}) = {"a";"b";"c"}
	got, err := fnUNIQUE([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{StringVal("a")}, {StringVal("b")}, {StringVal("a")}, {StringVal("c")}},
	}})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	want := []string{"a", "b", "c"}
	for i, w := range want {
		if got.Array[i][0].Str != w {
			t.Errorf("[%d]: got %q, want %q", i, got.Array[i][0].Str, w)
		}
	}
}

func TestUNIQUE_MixedTypes(t *testing.T) {
	// UNIQUE({1;"1";TRUE;1}) → {1;"1";TRUE} — 1 and "1" are different types
	got, err := fnUNIQUE([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(1)}, {StringVal("1")}, {BoolVal(true)}, {NumberVal(1)}},
	}})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	if got.Array[0][0].Type != ValueNumber || got.Array[0][0].Num != 1 {
		t.Errorf("[0]: got %v, want 1", got.Array[0][0])
	}
	if got.Array[1][0].Type != ValueString || got.Array[1][0].Str != "1" {
		t.Errorf("[1]: got %v, want \"1\"", got.Array[1][0])
	}
	if got.Array[2][0].Type != ValueBool || !got.Array[2][0].Bool {
		t.Errorf("[2]: got %v, want TRUE", got.Array[2][0])
	}
}

func TestUNIQUE_ExactlyOnce(t *testing.T) {
	// UNIQUE({1;2;1;3;2},,TRUE) = {3}
	got, err := fnUNIQUE([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}, {NumberVal(1)}, {NumberVal(3)}, {NumberVal(2)}}},
		BoolVal(false),
		BoolVal(true),
	})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	// Single value returned (not wrapped in array).
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("got %v, want 3", got)
	}
}

func TestUNIQUE_AllUnique(t *testing.T) {
	// UNIQUE({1;2;3}) = {1;2;3}
	got, err := fnUNIQUE([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}},
	}})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3-row array, got %v (rows=%d)", got.Type, len(got.Array))
	}
}

func TestUNIQUE_AllSame(t *testing.T) {
	// UNIQUE({5;5;5}) = {5}
	got, err := fnUNIQUE([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(5)}, {NumberVal(5)}, {NumberVal(5)}},
	}})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("got %v, want 5", got)
	}
}

func TestUNIQUE_AllSameExactlyOnce(t *testing.T) {
	// UNIQUE({5;5;5},,TRUE) → #CALC! (no values appear exactly once)
	got, err := fnUNIQUE([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(5)}, {NumberVal(5)}, {NumberVal(5)}}},
		BoolVal(false),
		BoolVal(true),
	})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("got %v, want #CALC!", got)
	}
}

func TestUNIQUE_SingleValue(t *testing.T) {
	// UNIQUE({42}) = 42
	got, err := fnUNIQUE([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(42)}},
	}})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("got %v, want 42", got)
	}
}

func TestUNIQUE_Booleans(t *testing.T) {
	// UNIQUE({TRUE;FALSE;TRUE}) = {TRUE;FALSE}
	got, err := fnUNIQUE([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{BoolVal(true)}, {BoolVal(false)}, {BoolVal(true)}},
	}})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2-row array, got %v", got)
	}
	if !got.Array[0][0].Bool {
		t.Errorf("[0]: got %v, want TRUE", got.Array[0][0])
	}
	if got.Array[1][0].Bool {
		t.Errorf("[1]: got %v, want FALSE", got.Array[1][0])
	}
}

func TestUNIQUE_EmptyHandling(t *testing.T) {
	// UNIQUE with empty values — empties are equal to each other
	got, err := fnUNIQUE([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{EmptyVal()}, {NumberVal(1)}, {EmptyVal()}, {NumberVal(2)}},
	}})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %v (rows=%d)", got.Type, len(got.Array))
	}
	if got.Array[0][0].Type != ValueEmpty {
		t.Errorf("[0]: got type %v, want empty", got.Array[0][0].Type)
	}
	if got.Array[1][0].Num != 1 {
		t.Errorf("[1]: got %v, want 1", got.Array[1][0])
	}
	if got.Array[2][0].Num != 2 {
		t.Errorf("[2]: got %v, want 2", got.Array[2][0])
	}
}

func TestUNIQUE_ErrorsPreserved(t *testing.T) {
	// Errors in the array are treated as values to compare, not propagated
	got, err := fnUNIQUE([]Value{{
		Type: ValueArray,
		Array: [][]Value{
			{ErrorVal(ErrValDIV0)},
			{NumberVal(1)},
			{ErrorVal(ErrValDIV0)},
			{ErrorVal(ErrValNA)},
		},
	}})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %v (rows=%d)", got.Type, len(got.Array))
	}
	if got.Array[0][0].Type != ValueError || got.Array[0][0].Err != ErrValDIV0 {
		t.Errorf("[0]: got %v, want #DIV/0!", got.Array[0][0])
	}
	if got.Array[1][0].Num != 1 {
		t.Errorf("[1]: got %v, want 1", got.Array[1][0])
	}
	if got.Array[2][0].Type != ValueError || got.Array[2][0].Err != ErrValNA {
		t.Errorf("[2]: got %v, want #N/A", got.Array[2][0])
	}
}

func TestUNIQUE_MultiColumnRows(t *testing.T) {
	// Multi-column: rows must match on ALL columns to be duplicates
	// {1,"a"; 2,"b"; 1,"a"; 1,"c"} → {1,"a"; 2,"b"; 1,"c"}
	got, err := fnUNIQUE([]Value{{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(1), StringVal("a")},
			{NumberVal(2), StringVal("b")},
			{NumberVal(1), StringVal("a")},
			{NumberVal(1), StringVal("c")},
		},
	}})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %v (rows=%d)", got.Type, len(got.Array))
	}
	// Row 0: {1, "a"}
	if got.Array[0][0].Num != 1 || got.Array[0][1].Str != "a" {
		t.Errorf("row 0: got %v %v, want 1 a", got.Array[0][0], got.Array[0][1])
	}
	// Row 1: {2, "b"}
	if got.Array[1][0].Num != 2 || got.Array[1][1].Str != "b" {
		t.Errorf("row 1: got %v %v, want 2 b", got.Array[1][0], got.Array[1][1])
	}
	// Row 2: {1, "c"}
	if got.Array[2][0].Num != 1 || got.Array[2][1].Str != "c" {
		t.Errorf("row 2: got %v %v, want 1 c", got.Array[2][0], got.Array[2][1])
	}
}

func TestUNIQUE_ByCol(t *testing.T) {
	// by_col=TRUE: compare columns instead of rows
	// {1,2,1; 3,4,3} with by_col=TRUE → columns {1,3}, {2,4}, {1,3}
	// Unique columns: {1,3}, {2,4} → result: {1,2; 3,4}
	got, err := fnUNIQUE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(1)},
			{NumberVal(3), NumberVal(4), NumberVal(3)},
		}},
		BoolVal(true),  // by_col
		BoolVal(false), // exactly_once
	})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2 rows, got %v (rows=%d)", got.Type, len(got.Array))
	}
	if len(got.Array[0]) != 2 {
		t.Fatalf("expected 2 cols, got %d", len(got.Array[0]))
	}
	// Result: {1,2; 3,4}
	if got.Array[0][0].Num != 1 || got.Array[0][1].Num != 2 {
		t.Errorf("row 0: got %v %v, want 1 2", got.Array[0][0], got.Array[0][1])
	}
	if got.Array[1][0].Num != 3 || got.Array[1][1].Num != 4 {
		t.Errorf("row 1: got %v %v, want 3 4", got.Array[1][0], got.Array[1][1])
	}
}

func TestUNIQUE_ByColExactlyOnce(t *testing.T) {
	// by_col=TRUE, exactly_once=TRUE
	// {1,2,1; 3,4,3} → only column {2,4} appears once
	got, err := fnUNIQUE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(1)},
			{NumberVal(3), NumberVal(4), NumberVal(3)},
		}},
		BoolVal(true), // by_col
		BoolVal(true), // exactly_once
	})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2 rows, got %v (rows=%d)", got.Type, len(got.Array))
	}
	if len(got.Array[0]) != 1 {
		t.Fatalf("expected 1 col, got %d", len(got.Array[0]))
	}
	if got.Array[0][0].Num != 2 || got.Array[1][0].Num != 4 {
		t.Errorf("got %v %v, want 2 4", got.Array[0][0], got.Array[1][0])
	}
}

func TestUNIQUE_NoArgs(t *testing.T) {
	got, err := fnUNIQUE([]Value{})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("no args: got %v, want #VALUE!", got)
	}
}

func TestUNIQUE_TooManyArgs(t *testing.T) {
	got, err := fnUNIQUE([]Value{NumberVal(1), BoolVal(false), BoolVal(false), NumberVal(0)})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("too many args: got %v, want #VALUE!", got)
	}
}

func TestUNIQUE_ScalarValue(t *testing.T) {
	// Single non-array value should return itself
	got, err := fnUNIQUE([]Value{NumberVal(7)})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 7 {
		t.Errorf("got %v, want 7", got)
	}
}

func TestUNIQUE_PreservesOrder(t *testing.T) {
	// UNIQUE({3;1;2;1;3;2}) = {3;1;2} — preserves first occurrence order
	got, err := fnUNIQUE([]Value{{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(3)}, {NumberVal(1)}, {NumberVal(2)},
			{NumberVal(1)}, {NumberVal(3)}, {NumberVal(2)},
		},
	}})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %v", got)
	}
	want := []float64{3, 1, 2}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestUNIQUE_ExactlyOnceMultiple(t *testing.T) {
	// UNIQUE({1;2;3;2;4;3},,TRUE) = {1;4} — only 1 and 4 appear exactly once
	got, err := fnUNIQUE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
			{NumberVal(2)}, {NumberVal(4)}, {NumberVal(3)},
		}},
		BoolVal(false),
		BoolVal(true),
	})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2 rows, got %v (rows=%d)", got.Type, len(got.Array))
	}
	if got.Array[0][0].Num != 1 {
		t.Errorf("[0]: got %v, want 1", got.Array[0][0])
	}
	if got.Array[1][0].Num != 4 {
		t.Errorf("[1]: got %v, want 4", got.Array[1][0])
	}
}

func TestUNIQUE_StringScalar(t *testing.T) {
	got, err := fnUNIQUE([]Value{StringVal("hello")})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueString || got.Str != "hello" {
		t.Errorf("got %v, want hello", got)
	}
}

func TestUNIQUE_ViaEval(t *testing.T) {
	// Test via the formula parser with array literal syntax
	resolver := &mockResolver{}
	cf := evalCompile(t, "UNIQUE({1;2;1;3;2})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	want := []float64{1, 2, 3}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestUNIQUE_ViaEvalExactlyOnce(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "UNIQUE({1;2;1;3;2},,TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("got %v, want 3", got)
	}
}

func TestUNIQUE_BoolNotEqualToNumber(t *testing.T) {
	// TRUE (bool) and 1 (number) should be considered different
	got, err := fnUNIQUE([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{BoolVal(true)}, {NumberVal(1)}, {BoolVal(true)}},
	}})
	if err != nil {
		t.Fatalf("fnUNIQUE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2 rows, got %v (rows=%d)", got.Type, len(got.Array))
	}
	if got.Array[0][0].Type != ValueBool || !got.Array[0][0].Bool {
		t.Errorf("[0]: got %v, want TRUE", got.Array[0][0])
	}
	if got.Array[1][0].Type != ValueNumber || got.Array[1][0].Num != 1 {
		t.Errorf("[1]: got %v, want 1", got.Array[1][0])
	}
}

func TestUNIQUE_FromRange(t *testing.T) {
	// Test with cell range reference
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(10),
			{Col: 1, Row: 4}: NumberVal(30),
		},
	}
	cf := evalCompile(t, "UNIQUE(A1:A4)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got type=%v rows=%d", got.Type, len(got.Array))
	}
	want := []float64{10, 20, 30}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

// ── FILTER tests ──────────────────────────────────────────────────────

func TestFILTER_BasicBoolean(t *testing.T) {
	// FILTER({1;2;3;4;5}, {TRUE;FALSE;TRUE;FALSE;TRUE}) = {1;3;5}
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}, {NumberVal(4)}, {NumberVal(5)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{BoolVal(true)}, {BoolVal(false)}, {BoolVal(true)}, {BoolVal(false)}, {BoolVal(true)},
		}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	want := []float64{1, 3, 5}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestFILTER_NumericBooleans(t *testing.T) {
	// FILTER({1;2;3}, {1;0;1}) = {1;3}
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(0)}, {NumberVal(1)},
		}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][0].Num != 1 || got.Array[1][0].Num != 3 {
		t.Errorf("got %v %v, want 1 3", got.Array[0][0], got.Array[1][0])
	}
}

func TestFILTER_AllMatch(t *testing.T) {
	// FILTER({1;2;3}, {1;1;1}) = {1;2;3}
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(1)}, {NumberVal(1)},
		}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	want := []float64{1, 2, 3}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestFILTER_NoneMatchWithIfEmpty(t *testing.T) {
	// FILTER({1;2;3}, {0;0;0}, "empty") = "empty"
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(0)}, {NumberVal(0)}, {NumberVal(0)},
		}},
		StringVal("empty"),
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueString || got.Str != "empty" {
		t.Errorf("got %v, want string 'empty'", got)
	}
}

func TestFILTER_NoneMatchWithoutIfEmpty(t *testing.T) {
	// FILTER({1;2;3}, {0;0;0}) = #CALC!
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(0)}, {NumberVal(0)}, {NumberVal(0)},
		}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("got %v, want #CALC!", got)
	}
}

func TestFILTER_SingleMatch(t *testing.T) {
	// FILTER({10;20;30}, {0;1;0}) = 20 (scalar)
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(10)}, {NumberVal(20)}, {NumberVal(30)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(0)}, {NumberVal(1)}, {NumberVal(0)},
		}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("got %v, want 20", got)
	}
}

func TestFILTER_MultiColumnRows(t *testing.T) {
	// FILTER({1,"a";2,"b";3,"c"}, {TRUE;FALSE;TRUE}) = {1,"a";3,"c"}
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), StringVal("a")},
			{NumberVal(2), StringVal("b")},
			{NumberVal(3), StringVal("c")},
		}},
		{Type: ValueArray, Array: [][]Value{
			{BoolVal(true)}, {BoolVal(false)}, {BoolVal(true)},
		}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2 rows, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][0].Num != 1 || got.Array[0][1].Str != "a" {
		t.Errorf("row 0: got %v %v, want 1 a", got.Array[0][0], got.Array[0][1])
	}
	if got.Array[1][0].Num != 3 || got.Array[1][1].Str != "c" {
		t.Errorf("row 1: got %v %v, want 3 c", got.Array[1][0], got.Array[1][1])
	}
}

func TestFILTER_StringValues(t *testing.T) {
	// FILTER({"a";"b";"c"}, {1;0;1}) = {"a";"c"}
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{StringVal("a")}, {StringVal("b")}, {StringVal("c")},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(0)}, {NumberVal(1)},
		}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2 rows, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][0].Str != "a" || got.Array[1][0].Str != "c" {
		t.Errorf("got %v %v, want a c", got.Array[0][0], got.Array[1][0])
	}
}

func TestFILTER_ErrorInInclude(t *testing.T) {
	// Error in include array propagates immediately.
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{BoolVal(true)}, {ErrorVal(ErrValDIV0)}, {BoolVal(true)},
		}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("got %v, want #DIV/0!", got)
	}
}

func TestFILTER_MismatchedSizes(t *testing.T) {
	// Include length doesn't match rows or columns.
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(0)},
		}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("got %v, want #VALUE!", got)
	}
}

func TestFILTER_IfEmptyNumber(t *testing.T) {
	// if_empty is a number
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(0)}, {NumberVal(0)},
		}},
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("got %v, want 0", got)
	}
}

func TestFILTER_WrongArgCount(t *testing.T) {
	// Too few args
	got, err := fnFILTER([]Value{NumberVal(1)})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("got %v, want #VALUE!", got)
	}

	// Too many args
	got, err = fnFILTER([]Value{NumberVal(1), NumberVal(1), NumberVal(0), NumberVal(0)})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("got %v, want #VALUE!", got)
	}
}

func TestFILTER_SingleElement(t *testing.T) {
	// FILTER({42}, {TRUE}) = 42
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(42)}}},
		{Type: ValueArray, Array: [][]Value{{BoolVal(true)}}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("got %v, want 42", got)
	}
}

func TestFILTER_SingleElementFalse(t *testing.T) {
	// FILTER({42}, {FALSE}) = #CALC!
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(42)}}},
		{Type: ValueArray, Array: [][]Value{{BoolVal(false)}}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("got %v, want #CALC!", got)
	}
}

func TestFILTER_ScalarInputs(t *testing.T) {
	// Scalar array + scalar include (both treated as 1x1)
	got, err := fnFILTER([]Value{
		NumberVal(5),
		BoolVal(true),
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("got %v, want 5", got)
	}
}

func TestFILTER_NegativeNumberIsTruthy(t *testing.T) {
	// Negative numbers are truthy (non-zero).
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(10)}, {NumberVal(20)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(-1)}, {NumberVal(0)},
		}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("got %v, want 10", got)
	}
}

func TestFILTER_ColumnFiltering(t *testing.T) {
	// FILTER({1,2,3;4,5,6}, {TRUE,FALSE,TRUE}) filters columns → {1,3;4,6}
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{BoolVal(true), BoolVal(false), BoolVal(true)},
		}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2 rows, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if len(got.Array[0]) != 2 {
		t.Fatalf("expected 2 cols, got %d", len(got.Array[0]))
	}
	if got.Array[0][0].Num != 1 || got.Array[0][1].Num != 3 {
		t.Errorf("row 0: got %v %v, want 1 3", got.Array[0][0], got.Array[0][1])
	}
	if got.Array[1][0].Num != 4 || got.Array[1][1].Num != 6 {
		t.Errorf("row 1: got %v %v, want 4 6", got.Array[1][0], got.Array[1][1])
	}
}

func TestFILTER_ColumnFilterNoneMatch(t *testing.T) {
	// Column filter with all FALSE → #CALC!
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{BoolVal(false), BoolVal(false), BoolVal(false)},
		}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("got %v, want #CALC!", got)
	}
}

func TestFILTER_ErrorInArray(t *testing.T) {
	// Error in array argument is propagated.
	got, err := fnFILTER([]Value{
		ErrorVal(ErrValNA),
		{Type: ValueArray, Array: [][]Value{{BoolVal(true)}}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("got %v, want #N/A", got)
	}
}

func TestFILTER_ErrorInIncludeArg(t *testing.T) {
	// Error value as the include argument itself (not an array element).
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}}},
		ErrorVal(ErrValREF),
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("got %v, want #REF!", got)
	}
}

func TestFILTER_ViaEval(t *testing.T) {
	// Test through the formula evaluator.
	resolver := &mockResolver{}
	cf := evalCompile(t, "FILTER({1;2;3;4;5},{TRUE;FALSE;TRUE;FALSE;TRUE})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	want := []float64{1, 3, 5}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestFILTER_ViaEvalIfEmpty(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `FILTER({1;2;3},{0;0;0},"none")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "none" {
		t.Errorf("got %v, want string 'none'", got)
	}
}

func TestFILTER_BoolFalseInInclude(t *testing.T) {
	// All FALSE booleans with no if_empty.
	got, err := fnFILTER([]Value{
		{Type: ValueArray, Array: [][]Value{
			{StringVal("a")}, {StringVal("b")}, {StringVal("c")},
		}},
		{Type: ValueArray, Array: [][]Value{
			{BoolVal(false)}, {BoolVal(false)}, {BoolVal(false)},
		}},
	})
	if err != nil {
		t.Fatalf("fnFILTER: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("got %v, want #CALC!", got)
	}
}

func TestXLOOKUP(t *testing.T) {
	// Vertical data: A1:A5 = lookup, B1:B5 = return
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Vertical lookup/return arrays
			{Col: 1, Row: 1}: StringVal("Apple"),
			{Col: 1, Row: 2}: StringVal("Banana"),
			{Col: 1, Row: 3}: StringVal("Cherry"),
			{Col: 1, Row: 4}: StringVal("Date"),
			{Col: 1, Row: 5}: StringVal("apple"), // duplicate, different case
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: NumberVal(40),
			{Col: 2, Row: 5}: NumberVal(50),

			// Numeric sorted data: C1:C5 = lookup, D1:D5 = return
			{Col: 3, Row: 1}: NumberVal(10),
			{Col: 3, Row: 2}: NumberVal(20),
			{Col: 3, Row: 3}: NumberVal(30),
			{Col: 3, Row: 4}: NumberVal(40),
			{Col: 3, Row: 5}: NumberVal(50),
			{Col: 4, Row: 1}: StringVal("ten"),
			{Col: 4, Row: 2}: StringVal("twenty"),
			{Col: 4, Row: 3}: StringVal("thirty"),
			{Col: 4, Row: 4}: StringVal("forty"),
			{Col: 4, Row: 5}: StringVal("fifty"),

			// Horizontal data: E1:I1 = lookup, E2:I2 = return
			{Col: 5, Row: 1}: StringVal("X"),
			{Col: 6, Row: 1}: StringVal("Y"),
			{Col: 7, Row: 1}: StringVal("Z"),
			{Col: 5, Row: 2}: NumberVal(100),
			{Col: 6, Row: 2}: NumberVal(200),
			{Col: 7, Row: 2}: NumberVal(300),
		},
	}

	t.Run("basic exact match string", func(t *testing.T) {
		cf := evalCompile(t, `XLOOKUP("Cherry",A1:A5,B1:B5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 30 {
			t.Errorf("got %v, want 30", got)
		}
	})

	t.Run("basic exact match number", func(t *testing.T) {
		cf := evalCompile(t, `XLOOKUP(20,C1:C5,D1:D5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "twenty" {
			t.Errorf("got %v, want twenty", got)
		}
	})

	t.Run("not found without if_not_found returns NA", func(t *testing.T) {
		cf := evalCompile(t, `XLOOKUP("Mango",A1:A5,B1:B5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("not found with if_not_found returns custom value", func(t *testing.T) {
		cf := evalCompile(t, `XLOOKUP("Mango",A1:A5,B1:B5,"Not Found")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "Not Found" {
			t.Errorf("got %v, want Not Found", got)
		}
	})

	t.Run("case insensitive matching", func(t *testing.T) {
		cf := evalCompile(t, `XLOOKUP("BANANA",A1:A5,B1:B5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("multiple matches returns first found", func(t *testing.T) {
		// "apple" matches row 1 (Apple) first due to case-insensitive compare
		cf := evalCompile(t, `XLOOKUP("apple",A1:A5,B1:B5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10 (first match)", got)
		}
	})

	t.Run("match_mode -1 exact or next smaller", func(t *testing.T) {
		// Lookup 25 in sorted numeric array; next smaller is 20
		cf := evalCompile(t, `XLOOKUP(25,C1:C5,D1:D5,,-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "twenty" {
			t.Errorf("got %v, want twenty", got)
		}
	})

	t.Run("match_mode 1 exact or next larger", func(t *testing.T) {
		// Lookup 25 in sorted numeric array; next larger is 30
		cf := evalCompile(t, `XLOOKUP(25,C1:C5,D1:D5,,1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "thirty" {
			t.Errorf("got %v, want thirty", got)
		}
	})

	t.Run("match_mode -1 exact match found", func(t *testing.T) {
		// Exact value exists; should return it
		cf := evalCompile(t, `XLOOKUP(30,C1:C5,D1:D5,,-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "thirty" {
			t.Errorf("got %v, want thirty", got)
		}
	})

	t.Run("horizontal lookup array", func(t *testing.T) {
		cf := evalCompile(t, `XLOOKUP("Y",E1:G1,E2:G2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 200 {
			t.Errorf("got %v, want 200", got)
		}
	})

	t.Run("too few args returns error", func(t *testing.T) {
		cf := evalCompile(t, `XLOOKUP("Apple",A1:A5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error for too few args", got)
		}
	})

	t.Run("wildcard star pattern", func(t *testing.T) {
		cf := evalCompile(t, `XLOOKUP("Ch*",A1:A5,B1:B5,,2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 30 {
			t.Errorf("got %v, want 30", got)
		}
	})

	t.Run("wildcard question mark pattern", func(t *testing.T) {
		cf := evalCompile(t, `XLOOKUP("Dat?",A1:A5,B1:B5,,2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 40 {
			t.Errorf("got %v, want 40", got)
		}
	})

	t.Run("if_not_found with numeric zero", func(t *testing.T) {
		cf := evalCompile(t, `XLOOKUP("Missing",A1:A5,B1:B5,0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v, want 0", got)
		}
	})
}

func TestXLOOKUP_WildcardMode(t *testing.T) {
	// Data layout: D2:D4 = lookup values, E2:E4 = return values
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 4, Row: 2}: StringVal("Banana Split"),
			{Col: 4, Row: 3}: StringVal("Apple Pie"),
			{Col: 4, Row: 4}: StringVal("Cherry Tart"),
			{Col: 5, Row: 2}: StringVal("BS"),
			{Col: 5, Row: 3}: StringVal("AP"),
			{Col: 5, Row: 4}: StringVal("CT"),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{
			name:    "wildcard star prefix",
			formula: `XLOOKUP("*Split",D2:D4,E2:E4,"N/A",2)`,
			want:    "BS",
		},
		{
			name:    "wildcard star suffix",
			formula: `XLOOKUP("Cherry*",D2:D4,E2:E4,"N/A",2)`,
			want:    "CT",
		},
		{
			name:    "wildcard question mark",
			formula: `XLOOKUP("Apple Pi?",D2:D4,E2:E4,"N/A",2)`,
			want:    "AP",
		},
		{
			name:    "wildcard no match returns not_found",
			formula: `XLOOKUP("*Mango*",D2:D4,E2:E4,"N/A",2)`,
			want:    "N/A",
		},
		{
			name:    "wildcard case insensitive",
			formula: `XLOOKUP("*split",D2:D4,E2:E4,"N/A",2)`,
			want:    "BS",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("got %v (type %d), want string %q", got, got.Type, tt.want)
			}
		})
	}
}

// ---- TAKE tests ----

func TestTAKE_FirstTwoRows(t *testing.T) {
	// TAKE({1,2,3;4,5,6;7,8,9}, 2) → {1,2,3;4,5,6}
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
			{NumberVal(7), NumberVal(8), NumberVal(9)},
		}},
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	want := [][]float64{{1, 2, 3}, {4, 5, 6}}
	for i, wRow := range want {
		for j, w := range wRow {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestTAKE_LastRow(t *testing.T) {
	// TAKE({1,2,3;4,5,6;7,8,9}, -1) → {7,8,9}
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
			{NumberVal(7), NumberVal(8), NumberVal(9)},
		}},
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 {
		t.Fatalf("expected 1-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	want := []float64{7, 8, 9}
	for j, w := range want {
		if got.Array[0][j].Num != w {
			t.Errorf("[0][%d]: got %g, want %g", j, got.Array[0][j].Num, w)
		}
	}
}

func TestTAKE_RowsAndColumns(t *testing.T) {
	// TAKE({1,2,3;4,5,6}, 1, 2) → {1,2}
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
		}},
		NumberVal(1),
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 1x2 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 1 || got.Array[0][1].Num != 2 {
		t.Errorf("got {%g,%g}, want {1,2}", got.Array[0][0].Num, got.Array[0][1].Num)
	}
}

func TestTAKE_NegRowsNegCols(t *testing.T) {
	// TAKE({1,2,3;4,5,6}, -1, -2) → {5,6}
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
		}},
		NumberVal(-1),
		NumberVal(-2),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 1x2 array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][0].Num != 5 || got.Array[0][1].Num != 6 {
		t.Errorf("got {%g,%g}, want {5,6}", got.Array[0][0].Num, got.Array[0][1].Num)
	}
}

func TestTAKE_ColumnArray(t *testing.T) {
	// TAKE({1;2;3}, 2) → {1;2}
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
			{NumberVal(2)},
			{NumberVal(3)},
		}},
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][0].Num != 1 || got.Array[1][0].Num != 2 {
		t.Errorf("got {%g;%g}, want {1;2}", got.Array[0][0].Num, got.Array[1][0].Num)
	}
}

func TestTAKE_ColumnArrayNegative(t *testing.T) {
	// TAKE({1;2;3}, -2) → {2;3}
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
			{NumberVal(2)},
			{NumberVal(3)},
		}},
		NumberVal(-2),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][0].Num != 2 || got.Array[1][0].Num != 3 {
		t.Errorf("got {%g;%g}, want {2;3}", got.Array[0][0].Num, got.Array[1][0].Num)
	}
}

func TestTAKE_SingleRowArray(t *testing.T) {
	// TAKE({1,2,3}, 1) → {1,2,3} (single row, take 1 row = entire row)
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
		}},
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3 array, got type=%v", got.Type)
	}
}

func TestTAKE_SingleRowTakeCols(t *testing.T) {
	// TAKE({1,2,3}, 1, 2) → {1,2}
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
		}},
		NumberVal(1),
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 1x2 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 1 || got.Array[0][1].Num != 2 {
		t.Errorf("got wrong values")
	}
}

func TestTAKE_RowsZeroError(t *testing.T) {
	// TAKE({1,2,3}, 0) → #VALUE!
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}}},
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTAKE_ColsZeroError(t *testing.T) {
	// TAKE({1,2,3}, 1, 0) → #VALUE!
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}}},
		NumberVal(1),
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTAKE_RowsExceedArray(t *testing.T) {
	// TAKE({1;2;3}, 5) → #VALUE!
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}}},
		NumberVal(5),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTAKE_NegRowsExceedArray(t *testing.T) {
	// TAKE({1;2;3}, -5) → #VALUE!
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}}},
		NumberVal(-5),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTAKE_ColsExceedArray(t *testing.T) {
	// TAKE({1,2,3}, 1, 5) → #VALUE!
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}}},
		NumberVal(1),
		NumberVal(5),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTAKE_Scalar(t *testing.T) {
	// TAKE(42, 1) → 42 (scalar wrapped in {{42}}, take 1 row = scalar)
	got, err := fnTAKE([]Value{
		NumberVal(42),
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("expected 42, got %v", got)
	}
}

func TestTAKE_ScalarNeg(t *testing.T) {
	// TAKE(42, -1) → 42
	got, err := fnTAKE([]Value{
		NumberVal(42),
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("expected 42, got %v", got)
	}
}

func TestTAKE_ScalarExceed(t *testing.T) {
	// TAKE(42, 2) → #VALUE! (scalar = 1 row, can't take 2)
	got, err := fnTAKE([]Value{
		NumberVal(42),
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTAKE_ErrorPropagation(t *testing.T) {
	// TAKE(#REF!, 1) → #REF!
	got, err := fnTAKE([]Value{
		ErrorVal(ErrValREF),
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("expected #REF!, got %v", got)
	}
}

func TestTAKE_TooFewArgs(t *testing.T) {
	got, err := fnTAKE([]Value{NumberVal(1)})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTAKE_TooManyArgs(t *testing.T) {
	got, err := fnTAKE([]Value{NumberVal(1), NumberVal(1), NumberVal(1), NumberVal(1)})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTAKE_AllRows(t *testing.T) {
	// TAKE({1;2;3}, 3) → {1;2;3}
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}}},
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3-row array, got type=%v", got.Type)
	}
}

func TestTAKE_AllRowsNeg(t *testing.T) {
	// TAKE({1;2;3}, -3) → {1;2;3}
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}}},
		NumberVal(-3),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3-row array, got type=%v", got.Type)
	}
}

func TestTAKE_NegCols(t *testing.T) {
	// TAKE({1,2,3;4,5,6}, 2, -1) → {3;6}
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
		}},
		NumberVal(2),
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 2x1 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 3 || got.Array[1][0].Num != 6 {
		t.Errorf("got {%g;%g}, want {3;6}", got.Array[0][0].Num, got.Array[1][0].Num)
	}
}

func TestTAKE_StringValues(t *testing.T) {
	// TAKE({"a","b","c";"d","e","f"}, 1) → {"a","b","c"}
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{StringVal("a"), StringVal("b"), StringVal("c")},
			{StringVal("d"), StringVal("e"), StringVal("f")},
		}},
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Str != "a" || got.Array[0][1].Str != "b" || got.Array[0][2].Str != "c" {
		t.Errorf("wrong string values")
	}
}

func TestTAKE_PosCols(t *testing.T) {
	// TAKE({1,2,3;4,5,6}, 2, 1) → {1;4}
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
		}},
		NumberVal(2),
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 2x1 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 1 || got.Array[1][0].Num != 4 {
		t.Errorf("got {%g;%g}, want {1;4}", got.Array[0][0].Num, got.Array[1][0].Num)
	}
}

func TestTAKE_NegColsExceed(t *testing.T) {
	// TAKE({1,2}, 1, -3) → #VALUE!
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}}},
		NumberVal(1),
		NumberVal(-3),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTAKE_ViaEval(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "TAKE({1,2,3;4,5,6;7,8,9},2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 2x3 array, got type=%v", got.Type)
	}
	if got.Array[1][2].Num != 6 {
		t.Errorf("expected 6, got %g", got.Array[1][2].Num)
	}
}

func TestTAKE_ViaEvalNeg(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "TAKE({1,2,3;4,5,6;7,8,9},-1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 7 {
		t.Errorf("expected 7, got %g", got.Array[0][0].Num)
	}
}

// ---- DROP tests ----

func TestDROP_FirstRow(t *testing.T) {
	// DROP({1,2,3;4,5,6;7,8,9}, 1) → {4,5,6;7,8,9}
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
			{NumberVal(7), NumberVal(8), NumberVal(9)},
		}},
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][0].Num != 4 || got.Array[1][0].Num != 7 {
		t.Errorf("got wrong values")
	}
}

func TestDROP_LastRow(t *testing.T) {
	// DROP({1,2,3;4,5,6;7,8,9}, -1) → {1,2,3;4,5,6}
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
			{NumberVal(7), NumberVal(8), NumberVal(9)},
		}},
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][0].Num != 1 || got.Array[1][0].Num != 4 {
		t.Errorf("got wrong values")
	}
}

func TestDROP_FirstColumn(t *testing.T) {
	// DROP({1,2,3;4,5,6}, 0, 1) → {2,3;5,6}
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
		}},
		NumberVal(0),
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 2 || got.Array[0][1].Num != 3 {
		t.Errorf("row 0: got {%g,%g}, want {2,3}", got.Array[0][0].Num, got.Array[0][1].Num)
	}
	if got.Array[1][0].Num != 5 || got.Array[1][1].Num != 6 {
		t.Errorf("row 1: got {%g,%g}, want {5,6}", got.Array[1][0].Num, got.Array[1][1].Num)
	}
}

func TestDROP_LastColumn(t *testing.T) {
	// DROP({1,2,3;4,5,6}, 0, -1) → {1,2;4,5}
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
		}},
		NumberVal(0),
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 1 || got.Array[0][1].Num != 2 {
		t.Errorf("row 0: got {%g,%g}, want {1,2}", got.Array[0][0].Num, got.Array[0][1].Num)
	}
}

func TestDROP_RowAndColumn(t *testing.T) {
	// DROP({1,2,3;4,5,6;7,8,9}, 1, 1) → {5,6;8,9}
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
			{NumberVal(7), NumberVal(8), NumberVal(9)},
		}},
		NumberVal(1),
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 5 || got.Array[0][1].Num != 6 {
		t.Errorf("row 0: got {%g,%g}, want {5,6}", got.Array[0][0].Num, got.Array[0][1].Num)
	}
	if got.Array[1][0].Num != 8 || got.Array[1][1].Num != 9 {
		t.Errorf("row 1: got {%g,%g}, want {8,9}", got.Array[1][0].Num, got.Array[1][1].Num)
	}
}

func TestDROP_NegRowNegCol(t *testing.T) {
	// DROP({1,2,3;4,5,6;7,8,9}, -1, -1) → {1,2;4,5}
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
			{NumberVal(7), NumberVal(8), NumberVal(9)},
		}},
		NumberVal(-1),
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 1 || got.Array[0][1].Num != 2 {
		t.Errorf("row 0 wrong")
	}
	if got.Array[1][0].Num != 4 || got.Array[1][1].Num != 5 {
		t.Errorf("row 1 wrong")
	}
}

func TestDROP_AllRowsError(t *testing.T) {
	// DROP({1;2;3}, 3) → #VALUE! (drops all rows)
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}}},
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestDROP_MoreThanAllRowsError(t *testing.T) {
	// DROP({1;2}, 5) → #VALUE!
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}}},
		NumberVal(5),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestDROP_AllColsError(t *testing.T) {
	// DROP({1,2,3}, 0, 3) → #VALUE!
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}}},
		NumberVal(0),
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestDROP_AllNegRowsError(t *testing.T) {
	// DROP({1;2;3}, -3) → #VALUE!
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}}},
		NumberVal(-3),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestDROP_ZeroRows(t *testing.T) {
	// DROP({1,2,3;4,5,6}, 0) → {1,2,3;4,5,6} (drop nothing)
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
		}},
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 2x3 array, got type=%v", got.Type)
	}
}

func TestDROP_Scalar(t *testing.T) {
	// DROP(42, 0) → 42 (scalar, drop 0 rows)
	got, err := fnDROP([]Value{
		NumberVal(42),
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("expected 42, got %v", got)
	}
}

func TestDROP_ScalarDropAll(t *testing.T) {
	// DROP(42, 1) → #VALUE! (scalar = 1 row, drop 1 = nothing left)
	got, err := fnDROP([]Value{
		NumberVal(42),
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestDROP_ErrorPropagation(t *testing.T) {
	// DROP(#REF!, 1) → #REF!
	got, err := fnDROP([]Value{
		ErrorVal(ErrValREF),
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("expected #REF!, got %v", got)
	}
}

func TestDROP_TooFewArgs(t *testing.T) {
	got, err := fnDROP([]Value{NumberVal(1)})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestDROP_TooManyArgs(t *testing.T) {
	got, err := fnDROP([]Value{NumberVal(1), NumberVal(1), NumberVal(1), NumberVal(1)})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestDROP_TwoRows(t *testing.T) {
	// DROP({1;2;3;4;5}, 2) → {3;4;5}
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}, {NumberVal(4)}, {NumberVal(5)},
		}},
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][0].Num != 3 || got.Array[1][0].Num != 4 || got.Array[2][0].Num != 5 {
		t.Errorf("wrong values")
	}
}

func TestDROP_LastTwoRows(t *testing.T) {
	// DROP({1;2;3;4;5}, -2) → {1;2;3}
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}, {NumberVal(4)}, {NumberVal(5)},
		}},
		NumberVal(-2),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][0].Num != 1 || got.Array[2][0].Num != 3 {
		t.Errorf("wrong values")
	}
}

func TestDROP_StringValues(t *testing.T) {
	// DROP({"a","b","c";"d","e","f"}, 1) → {"d","e","f"}
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{
			{StringVal("a"), StringVal("b"), StringVal("c")},
			{StringVal("d"), StringVal("e"), StringVal("f")},
		}},
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Str != "d" || got.Array[0][1].Str != "e" || got.Array[0][2].Str != "f" {
		t.Errorf("wrong string values")
	}
}

func TestDROP_TwoCols(t *testing.T) {
	// DROP({1,2,3;4,5,6}, 0, 2) → {3;6}
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
		}},
		NumberVal(0),
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 2x1 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 3 || got.Array[1][0].Num != 6 {
		t.Errorf("got {%g;%g}, want {3;6}", got.Array[0][0].Num, got.Array[1][0].Num)
	}
}

func TestDROP_NegAllCols(t *testing.T) {
	// DROP({1,2}, 0, -2) → #VALUE!
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}}},
		NumberVal(0),
		NumberVal(-2),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestDROP_ViaEval(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "DROP({1,2,3;4,5,6;7,8,9},1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 2x3 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 4 {
		t.Errorf("expected 4, got %g", got.Array[0][0].Num)
	}
}

func TestDROP_ViaEvalNeg(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "DROP({1,2,3;4,5,6;7,8,9},-1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 2x3 array, got type=%v", got.Type)
	}
	if got.Array[1][2].Num != 6 {
		t.Errorf("expected 6, got %g", got.Array[1][2].Num)
	}
}

func TestDROP_SingleResultIsScalar(t *testing.T) {
	// DROP({1,2;3,4}, 1, 1) → 4 (single cell result unwrapped)
	got, err := fnDROP([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
			{NumberVal(3), NumberVal(4)},
		}},
		NumberVal(1),
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnDROP: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("expected scalar 4, got %v", got)
	}
}

func TestTAKE_SingleResultIsScalar(t *testing.T) {
	// TAKE({1,2;3,4}, 1, 1) → 1 (single cell result unwrapped)
	got, err := fnTAKE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
			{NumberVal(3), NumberVal(4)},
		}},
		NumberVal(1),
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnTAKE: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("expected scalar 1, got %v", got)
	}
}

func assertLookupValueEqual(t *testing.T, got, want Value) {
	t.Helper()

	if got.Type != want.Type {
		t.Fatalf("type mismatch: got %v, want %v (got=%v want=%v)", got.Type, want.Type, got, want)
	}

	switch want.Type {
	case ValueEmpty:
		return
	case ValueNumber:
		if got.Num != want.Num {
			t.Fatalf("number mismatch: got %g, want %g", got.Num, want.Num)
		}
	case ValueString:
		if got.Str != want.Str {
			t.Fatalf("string mismatch: got %q, want %q", got.Str, want.Str)
		}
	case ValueBool:
		if got.Bool != want.Bool {
			t.Fatalf("bool mismatch: got %v, want %v", got.Bool, want.Bool)
		}
	case ValueError:
		if got.Err != want.Err {
			t.Fatalf("error mismatch: got %v, want %v", got.Err, want.Err)
		}
	case ValueArray:
		if len(got.Array) != len(want.Array) {
			t.Fatalf("row count mismatch: got %d, want %d", len(got.Array), len(want.Array))
		}
		for r := range want.Array {
			if len(got.Array[r]) != len(want.Array[r]) {
				t.Fatalf("col count mismatch at row %d: got %d, want %d", r, len(got.Array[r]), len(want.Array[r]))
			}
			for c := range want.Array[r] {
				assertLookupValueEqual(t, got.Array[r][c], want.Array[r][c])
			}
		}
	default:
		t.Fatalf("unsupported value type in test helper: %v", want.Type)
	}
}

func TestCHOOSECOLS(t *testing.T) {
	base := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(1), NumberVal(2), NumberVal(3)},
		{NumberVal(4), NumberVal(5), NumberVal(6)},
	}}
	ragged := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(1), NumberVal(2), NumberVal(3)},
		{NumberVal(4)},
	}}
	mixed := Value{Type: ValueArray, Array: [][]Value{
		{StringVal("a"), BoolVal(true), ErrorVal(ErrValNA)},
		{EmptyVal(), NumberVal(2), StringVal("z")},
	}}

	tests := []struct {
		name string
		args []Value
		want Value
	}{
		{
			name: "first_column",
			args: []Value{base, NumberVal(1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1)},
				{NumberVal(4)},
			}},
		},
		{
			name: "last_column_negative",
			args: []Value{base, NumberVal(-1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3)},
				{NumberVal(6)},
			}},
		},
		{
			name: "reorder_columns",
			args: []Value{base, NumberVal(3), NumberVal(1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3), NumberVal(1)},
				{NumberVal(6), NumberVal(4)},
			}},
		},
		{
			name: "duplicate_columns",
			args: []Value{base, NumberVal(2), NumberVal(2), NumberVal(1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2), NumberVal(2), NumberVal(1)},
				{NumberVal(5), NumberVal(5), NumberVal(4)},
			}},
		},
		{
			name: "mixed_positive_and_negative",
			args: []Value{base, NumberVal(-1), NumberVal(2), NumberVal(-3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3), NumberVal(2), NumberVal(1)},
				{NumberVal(6), NumberVal(5), NumberVal(4)},
			}},
		},
		{
			name: "scalar_first_column",
			args: []Value{NumberVal(9), NumberVal(1)},
			want: NumberVal(9),
		},
		{
			name: "scalar_negative_one",
			args: []Value{StringVal("x"), NumberVal(-1)},
			want: StringVal("x"),
		},
		{
			name: "bool_index_true_coerces_to_one",
			args: []Value{base, BoolVal(true)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1)},
				{NumberVal(4)},
			}},
		},
		{
			name: "numeric_string_index",
			args: []Value{base, StringVal("2")},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2)},
				{NumberVal(5)},
			}},
		},
		{
			name: "fractional_index_truncates",
			args: []Value{base, NumberVal(2.9)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2)},
				{NumberVal(5)},
			}},
		},
		{
			name: "fractional_negative_index_truncates",
			args: []Value{base, NumberVal(-1.9)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3)},
				{NumberVal(6)},
			}},
		},
		{
			name: "ragged_rows_fill_missing_with_empty",
			args: []Value{ragged, NumberVal(2), NumberVal(3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2), NumberVal(3)},
				{EmptyVal(), EmptyVal()},
			}},
		},
		{
			name: "preserves_error_values",
			args: []Value{mixed, NumberVal(3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{ErrorVal(ErrValNA)},
				{StringVal("z")},
			}},
		},
		{
			name: "preserves_empty_values",
			args: []Value{mixed, NumberVal(1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{StringVal("a")},
				{EmptyVal()},
			}},
		},
		{
			name: "array_error_passthrough",
			args: []Value{ErrorVal(ErrValREF), NumberVal(1)},
			want: ErrorVal(ErrValREF),
		},
		{
			name: "empty_array_is_value_error",
			args: []Value{{Type: ValueArray, Array: [][]Value{}}, NumberVal(1)},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "zero_index_errors",
			args: []Value{base, NumberVal(0)},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "positive_index_too_large_errors",
			args: []Value{base, NumberVal(4)},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "negative_index_too_large_errors",
			args: []Value{base, NumberVal(-4)},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "non_numeric_string_index_errors",
			args: []Value{base, StringVal("abc")},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "error_index_propagates",
			args: []Value{base, ErrorVal(ErrValDIV0)},
			want: ErrorVal(ErrValDIV0),
		},
		{
			name: "too_few_args_errors",
			args: []Value{base},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "no_args_errors",
			args: nil,
			want: ErrorVal(ErrValVALUE),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := fnCHOOSECOLS(tt.args)
			if err != nil {
				t.Fatalf("fnCHOOSECOLS: %v", err)
			}
			assertLookupValueEqual(t, got, tt.want)
		})
	}
}

func TestCHOOSECOLS_ViaEval(t *testing.T) {
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

	tests := []struct {
		name    string
		formula string
		want    Value
	}{
		{
			name:    "range_reorder",
			formula: "CHOOSECOLS(A1:C2,3,1)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3), NumberVal(1)},
				{NumberVal(6), NumberVal(4)},
			}},
		},
		{
			name:    "range_negative_index",
			formula: "CHOOSECOLS(A1:C2,-2)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2)},
				{NumberVal(5)},
			}},
		},
		{
			name:    "scalar_formula",
			formula: "CHOOSECOLS(42,1)",
			want:    NumberVal(42),
		},
		{
			name:    "string_index_formula",
			formula: `CHOOSECOLS(A1:C2,"2")`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2)},
				{NumberVal(5)},
			}},
		},
		{
			name:    "bool_index_formula",
			formula: "CHOOSECOLS(A1:C2,TRUE,3)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(3)},
				{NumberVal(4), NumberVal(6)},
			}},
		},
		{
			name:    "too_few_args_formula",
			formula: "CHOOSECOLS(A1:C2)",
			want:    ErrorVal(ErrValVALUE),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			assertLookupValueEqual(t, got, tt.want)
		})
	}
}

func TestCHOOSEROWS(t *testing.T) {
	base := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(1), NumberVal(2), NumberVal(3)},
		{NumberVal(4), NumberVal(5), NumberVal(6)},
		{NumberVal(7), NumberVal(8), NumberVal(9)},
	}}
	mixed := Value{Type: ValueArray, Array: [][]Value{
		{StringVal("a"), BoolVal(true), ErrorVal(ErrValNA)},
		{EmptyVal(), NumberVal(2), StringVal("z")},
		{NumberVal(9), StringVal("tail"), BoolVal(false)},
	}}

	tests := []struct {
		name string
		args []Value
		want Value
	}{
		{
			name: "first_row",
			args: []Value{base, NumberVal(1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2), NumberVal(3)},
			}},
		},
		{
			name: "last_row_negative",
			args: []Value{base, NumberVal(-1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(7), NumberVal(8), NumberVal(9)},
			}},
		},
		{
			name: "reorder_rows",
			args: []Value{base, NumberVal(3), NumberVal(1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(7), NumberVal(8), NumberVal(9)},
				{NumberVal(1), NumberVal(2), NumberVal(3)},
			}},
		},
		{
			name: "duplicate_rows",
			args: []Value{base, NumberVal(2), NumberVal(2), NumberVal(1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4), NumberVal(5), NumberVal(6)},
				{NumberVal(4), NumberVal(5), NumberVal(6)},
				{NumberVal(1), NumberVal(2), NumberVal(3)},
			}},
		},
		{
			name: "mixed_positive_and_negative",
			args: []Value{base, NumberVal(-1), NumberVal(2), NumberVal(-3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(7), NumberVal(8), NumberVal(9)},
				{NumberVal(4), NumberVal(5), NumberVal(6)},
				{NumberVal(1), NumberVal(2), NumberVal(3)},
			}},
		},
		{
			name: "scalar_first_row",
			args: []Value{NumberVal(9), NumberVal(1)},
			want: NumberVal(9),
		},
		{
			name: "scalar_negative_one",
			args: []Value{StringVal("x"), NumberVal(-1)},
			want: StringVal("x"),
		},
		{
			name: "bool_index_true_coerces_to_one",
			args: []Value{base, BoolVal(true)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2), NumberVal(3)},
			}},
		},
		{
			name: "numeric_string_index",
			args: []Value{base, StringVal("2")},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4), NumberVal(5), NumberVal(6)},
			}},
		},
		{
			name: "fractional_index_truncates",
			args: []Value{base, NumberVal(2.9)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4), NumberVal(5), NumberVal(6)},
			}},
		},
		{
			name: "fractional_negative_index_truncates",
			args: []Value{base, NumberVal(-1.9)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(7), NumberVal(8), NumberVal(9)},
			}},
		},
		{
			name: "preserves_error_values",
			args: []Value{mixed, NumberVal(1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{StringVal("a"), BoolVal(true), ErrorVal(ErrValNA)},
			}},
		},
		{
			name: "preserves_empty_values",
			args: []Value{mixed, NumberVal(2)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{EmptyVal(), NumberVal(2), StringVal("z")},
			}},
		},
		{
			name: "multiple_rows_with_mixed_types",
			args: []Value{mixed, NumberVal(3), NumberVal(1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(9), StringVal("tail"), BoolVal(false)},
				{StringVal("a"), BoolVal(true), ErrorVal(ErrValNA)},
			}},
		},
		{
			name: "array_error_passthrough",
			args: []Value{ErrorVal(ErrValREF), NumberVal(1)},
			want: ErrorVal(ErrValREF),
		},
		{
			name: "empty_array_is_value_error",
			args: []Value{{Type: ValueArray, Array: [][]Value{}}, NumberVal(1)},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "zero_index_errors",
			args: []Value{base, NumberVal(0)},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "positive_index_too_large_errors",
			args: []Value{base, NumberVal(4)},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "negative_index_too_large_errors",
			args: []Value{base, NumberVal(-4)},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "nonnumeric_string_index_errors",
			args: []Value{base, StringVal("abc")},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "error_index_propagates",
			args: []Value{base, ErrorVal(ErrValDIV0)},
			want: ErrorVal(ErrValDIV0),
		},
		{
			name: "too_few_args_errors",
			args: []Value{base},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "no_args_errors",
			args: nil,
			want: ErrorVal(ErrValVALUE),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := fnCHOOSEROWS(tt.args)
			if err != nil {
				t.Fatalf("fnCHOOSEROWS: %v", err)
			}
			assertLookupValueEqual(t, got, tt.want)
		})
	}
}

func TestCHOOSEROWS_ViaEval(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 3, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 3, Row: 2}: NumberVal(6),
			{Col: 1, Row: 3}: NumberVal(7),
			{Col: 2, Row: 3}: NumberVal(8),
			{Col: 3, Row: 3}: NumberVal(9),
			{Col: 1, Row: 4}: NumberVal(10),
			{Col: 2, Row: 4}: NumberVal(11),
			{Col: 3, Row: 4}: NumberVal(12),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    Value
	}{
		{
			name:    "range_reorder",
			formula: "CHOOSEROWS(A1:C4,4,1)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10), NumberVal(11), NumberVal(12)},
				{NumberVal(1), NumberVal(2), NumberVal(3)},
			}},
		},
		{
			name:    "range_negative_index",
			formula: "CHOOSEROWS(A1:C4,-2)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(7), NumberVal(8), NumberVal(9)},
			}},
		},
		{
			name:    "scalar_formula",
			formula: "CHOOSEROWS(42,1)",
			want:    NumberVal(42),
		},
		{
			name:    "string_index_formula",
			formula: `CHOOSEROWS(A1:C4,"2")`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4), NumberVal(5), NumberVal(6)},
			}},
		},
		{
			name:    "bool_index_formula",
			formula: "CHOOSEROWS(A1:C4,TRUE,4)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2), NumberVal(3)},
				{NumberVal(10), NumberVal(11), NumberVal(12)},
			}},
		},
		{
			name:    "too_few_args_formula",
			formula: "CHOOSEROWS(A1:C4)",
			want:    ErrorVal(ErrValVALUE),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			assertLookupValueEqual(t, got, tt.want)
		})
	}
}

// ---------------------------------------------------------------------------
// TOCOL
// ---------------------------------------------------------------------------

func TestTOCOL_BasicRow(t *testing.T) {
	// TOCOL({1,2,3}) → column {1;2;3}
	got, err := fnTOCOL([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}},
	}})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	for i, w := range []float64{1, 2, 3} {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestTOCOL_2D(t *testing.T) {
	// TOCOL({1,2;3,4}) → {1;2;3;4}
	got, err := fnTOCOL([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(1), NumberVal(2)}, {NumberVal(3), NumberVal(4)}},
	}})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := []float64{1, 2, 3, 4}
	if len(got.Array) != len(want) {
		t.Fatalf("expected %d rows, got %d", len(want), len(got.Array))
	}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestTOCOL_ColumnScan(t *testing.T) {
	// TOCOL({1,2;3,4},,TRUE) → {1;3;2;4}
	got, err := fnTOCOL([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}, {NumberVal(3), NumberVal(4)}}},
		NumberVal(0),
		BoolVal(true),
	})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	want := []float64{1, 3, 2, 4}
	if len(got.Array) != len(want) {
		t.Fatalf("expected %d rows, got %d", len(want), len(got.Array))
	}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestTOCOL_IgnoreBlanks(t *testing.T) {
	// TOCOL({1,"";3,4},1) → {1;3;4}
	got, err := fnTOCOL([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), EmptyVal()}, {NumberVal(3), NumberVal(4)}}},
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	want := []float64{1, 3, 4}
	if len(got.Array) != len(want) {
		t.Fatalf("expected %d rows, got %d", len(want), len(got.Array))
	}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestTOCOL_IgnoreErrors(t *testing.T) {
	// TOCOL({1,#N/A;3,4},2) → {1;3;4}
	got, err := fnTOCOL([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), ErrorVal(ErrValNA)}, {NumberVal(3), NumberVal(4)}}},
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	want := []float64{1, 3, 4}
	if len(got.Array) != len(want) {
		t.Fatalf("expected %d rows, got %d", len(want), len(got.Array))
	}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestTOCOL_IgnoreBlanksAndErrors(t *testing.T) {
	// TOCOL({1,"",#N/A;3,"",4},3) → {1;3;4}
	got, err := fnTOCOL([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), EmptyVal(), ErrorVal(ErrValNA)},
			{NumberVal(3), EmptyVal(), NumberVal(4)},
		}},
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	want := []float64{1, 3, 4}
	if len(got.Array) != len(want) {
		t.Fatalf("expected %d rows, got %d", len(want), len(got.Array))
	}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestTOCOL_Scalar(t *testing.T) {
	// TOCOL(5) → 5
	got, err := fnTOCOL([]Value{NumberVal(5)})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("expected 5, got %v", got)
	}
}

func TestTOCOL_ScalarString(t *testing.T) {
	got, err := fnTOCOL([]Value{StringVal("hello")})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if got.Type != ValueString || got.Str != "hello" {
		t.Errorf("expected hello, got %v", got)
	}
}

func TestTOCOL_Column(t *testing.T) {
	// TOCOL({1;2;3}) → {1;2;3} (already a column)
	got, err := fnTOCOL([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}},
	}})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3-row array, got %v", got)
	}
	for i, w := range []float64{1, 2, 3} {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestTOCOL_3x3ColumnScan(t *testing.T) {
	// TOCOL({1,2,3;4,5,6;7,8,9},,TRUE) → {1;4;7;2;5;8;3;6;9}
	got, err := fnTOCOL([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
			{NumberVal(7), NumberVal(8), NumberVal(9)},
		}},
		NumberVal(0),
		BoolVal(true),
	})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	want := []float64{1, 4, 7, 2, 5, 8, 3, 6, 9}
	if len(got.Array) != len(want) {
		t.Fatalf("expected %d rows, got %d", len(want), len(got.Array))
	}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestTOCOL_ErrorPassthrough(t *testing.T) {
	// TOCOL(#VALUE!) → #VALUE!
	got, err := fnTOCOL([]Value{ErrorVal(ErrValVALUE)})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTOCOL_KeepErrors(t *testing.T) {
	// TOCOL({1,#N/A},0) → {1;#N/A} — errors are kept when ignore=0
	got, err := fnTOCOL([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), ErrorVal(ErrValNA)}}},
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if len(got.Array) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(got.Array))
	}
	if got.Array[0][0].Num != 1 {
		t.Errorf("[0]: got %v, want 1", got.Array[0][0])
	}
	if got.Array[1][0].Type != ValueError || got.Array[1][0].Err != ErrValNA {
		t.Errorf("[1]: got %v, want #N/A", got.Array[1][0])
	}
}

func TestTOCOL_NoArgs(t *testing.T) {
	got, err := fnTOCOL([]Value{})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTOCOL_TooManyArgs(t *testing.T) {
	got, err := fnTOCOL([]Value{NumberVal(1), NumberVal(0), BoolVal(false), NumberVal(0)})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTOCOL_InvalidIgnore(t *testing.T) {
	got, err := fnTOCOL([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}}},
		NumberVal(4),
	})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTOCOL_AllBlanksIgnored(t *testing.T) {
	// TOCOL({"",""},1) → #CALC! (nothing left)
	got, err := fnTOCOL([]Value{
		{Type: ValueArray, Array: [][]Value{{EmptyVal(), EmptyVal()}}},
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("expected #CALC!, got %v", got)
	}
}

func TestTOCOL_MixedTypes(t *testing.T) {
	// TOCOL({1,"a";TRUE,#N/A}) → {1;"a";TRUE;#N/A}
	got, err := fnTOCOL([]Value{{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(1), StringVal("a")},
			{BoolVal(true), ErrorVal(ErrValNA)},
		},
	}})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if len(got.Array) != 4 {
		t.Fatalf("expected 4 rows, got %d", len(got.Array))
	}
	if got.Array[0][0].Type != ValueNumber || got.Array[0][0].Num != 1 {
		t.Errorf("[0]: got %v", got.Array[0][0])
	}
	if got.Array[1][0].Type != ValueString || got.Array[1][0].Str != "a" {
		t.Errorf("[1]: got %v", got.Array[1][0])
	}
	if got.Array[2][0].Type != ValueBool || !got.Array[2][0].Bool {
		t.Errorf("[2]: got %v", got.Array[2][0])
	}
	if got.Array[3][0].Type != ValueError || got.Array[3][0].Err != ErrValNA {
		t.Errorf("[3]: got %v", got.Array[3][0])
	}
}

func TestTOCOL_ColumnScanIgnoreBlanks(t *testing.T) {
	// TOCOL({1,"";3,4},1,TRUE) → {1;3;4} (column scan, ignore blanks)
	got, err := fnTOCOL([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), EmptyVal()}, {NumberVal(3), NumberVal(4)}}},
		NumberVal(1),
		BoolVal(true),
	})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	want := []float64{1, 3, 4}
	if len(got.Array) != len(want) {
		t.Fatalf("expected %d rows, got %d", len(want), len(got.Array))
	}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestTOCOL_SingleElement(t *testing.T) {
	// TOCOL({5}) → 5
	got, err := fnTOCOL([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(5)}},
	}})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("expected 5, got %v", got)
	}
}

func TestTOCOL_BoolInput(t *testing.T) {
	got, err := fnTOCOL([]Value{BoolVal(true)})
	if err != nil {
		t.Fatalf("fnTOCOL: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("expected TRUE, got %v", got)
	}
}

func TestTOCOL_ViaEval(t *testing.T) {
	cf := evalCompile(t, "TOCOL({1,2;3,4})")
	got, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 4 {
		t.Fatalf("expected 4-row array, got %v", got)
	}
	for i, w := range []float64{1, 2, 3, 4} {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

// ---------------------------------------------------------------------------
// TOROW
// ---------------------------------------------------------------------------

func TestTOROW_BasicColumn(t *testing.T) {
	// TOROW({1;2;3}) → row {1,2,3}
	got, err := fnTOROW([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}},
	}})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	for i, w := range []float64{1, 2, 3} {
		if got.Array[0][i].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestTOROW_2D(t *testing.T) {
	// TOROW({1,2;3,4}) → {1,2,3,4}
	got, err := fnTOROW([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(1), NumberVal(2)}, {NumberVal(3), NumberVal(4)}},
	}})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 4 {
		t.Fatalf("expected 1x4 array, got %v", got)
	}
	for i, w := range []float64{1, 2, 3, 4} {
		if got.Array[0][i].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestTOROW_ColumnScan(t *testing.T) {
	// TOROW({1,2;3,4},,TRUE) → {1,3,2,4}
	got, err := fnTOROW([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}, {NumberVal(3), NumberVal(4)}}},
		NumberVal(0),
		BoolVal(true),
	})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	want := []float64{1, 3, 2, 4}
	if len(got.Array[0]) != len(want) {
		t.Fatalf("expected %d cols, got %d", len(want), len(got.Array[0]))
	}
	for i, w := range want {
		if got.Array[0][i].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestTOROW_IgnoreBlanks(t *testing.T) {
	got, err := fnTOROW([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), EmptyVal()}, {NumberVal(3), NumberVal(4)}}},
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	want := []float64{1, 3, 4}
	if len(got.Array[0]) != len(want) {
		t.Fatalf("expected %d cols, got %d", len(want), len(got.Array[0]))
	}
	for i, w := range want {
		if got.Array[0][i].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestTOROW_IgnoreErrors(t *testing.T) {
	got, err := fnTOROW([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), ErrorVal(ErrValNA)}, {NumberVal(3), NumberVal(4)}}},
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	want := []float64{1, 3, 4}
	if len(got.Array[0]) != len(want) {
		t.Fatalf("expected %d cols, got %d", len(want), len(got.Array[0]))
	}
	for i, w := range want {
		if got.Array[0][i].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestTOROW_IgnoreBlanksAndErrors(t *testing.T) {
	got, err := fnTOROW([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), EmptyVal(), ErrorVal(ErrValNA)},
			{NumberVal(3), EmptyVal(), NumberVal(4)},
		}},
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	want := []float64{1, 3, 4}
	if len(got.Array[0]) != len(want) {
		t.Fatalf("expected %d cols, got %d", len(want), len(got.Array[0]))
	}
	for i, w := range want {
		if got.Array[0][i].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestTOROW_Scalar(t *testing.T) {
	got, err := fnTOROW([]Value{NumberVal(42)})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("expected 42, got %v", got)
	}
}

func TestTOROW_ErrorPassthrough(t *testing.T) {
	got, err := fnTOROW([]Value{ErrorVal(ErrValVALUE)})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTOROW_NoArgs(t *testing.T) {
	got, err := fnTOROW([]Value{})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTOROW_TooManyArgs(t *testing.T) {
	got, err := fnTOROW([]Value{NumberVal(1), NumberVal(0), BoolVal(false), NumberVal(0)})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTOROW_InvalidIgnore(t *testing.T) {
	got, err := fnTOROW([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}}},
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestTOROW_AllBlanksIgnored(t *testing.T) {
	got, err := fnTOROW([]Value{
		{Type: ValueArray, Array: [][]Value{{EmptyVal(), EmptyVal()}}},
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("expected #CALC!, got %v", got)
	}
}

func TestTOROW_KeepErrors(t *testing.T) {
	got, err := fnTOROW([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), ErrorVal(ErrValNA)}}},
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	if len(got.Array[0]) != 2 {
		t.Fatalf("expected 2 cols, got %d", len(got.Array[0]))
	}
	if got.Array[0][1].Type != ValueError || got.Array[0][1].Err != ErrValNA {
		t.Errorf("[1]: got %v, want #N/A", got.Array[0][1])
	}
}

func TestTOROW_Row(t *testing.T) {
	// TOROW({1,2,3}) → {1,2,3} (already a row)
	got, err := fnTOROW([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}},
	}})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3 array, got %v", got)
	}
}

func TestTOROW_MixedTypes(t *testing.T) {
	got, err := fnTOROW([]Value{{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(1), StringVal("a")},
			{BoolVal(true), ErrorVal(ErrValNA)},
		},
	}})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 4 {
		t.Fatalf("expected 1x4, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[0][0].Type != ValueNumber {
		t.Errorf("[0]: expected number")
	}
	if got.Array[0][1].Type != ValueString {
		t.Errorf("[1]: expected string")
	}
	if got.Array[0][2].Type != ValueBool {
		t.Errorf("[2]: expected bool")
	}
	if got.Array[0][3].Type != ValueError {
		t.Errorf("[3]: expected error")
	}
}

func TestTOROW_3x3ColumnScan(t *testing.T) {
	got, err := fnTOROW([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
			{NumberVal(7), NumberVal(8), NumberVal(9)},
		}},
		NumberVal(0),
		BoolVal(true),
	})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	want := []float64{1, 4, 7, 2, 5, 8, 3, 6, 9}
	if len(got.Array[0]) != len(want) {
		t.Fatalf("expected %d cols, got %d", len(want), len(got.Array[0]))
	}
	for i, w := range want {
		if got.Array[0][i].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestTOROW_SingleElement(t *testing.T) {
	got, err := fnTOROW([]Value{{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(5)}},
	}})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("expected 5, got %v", got)
	}
}

func TestTOROW_ViaEval(t *testing.T) {
	cf := evalCompile(t, "TOROW({1;2;3})")
	got, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3 array, got %v", got)
	}
	for i, w := range []float64{1, 2, 3} {
		if got.Array[0][i].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestTOROW_ColumnScanIgnoreBlanks(t *testing.T) {
	got, err := fnTOROW([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), EmptyVal()}, {NumberVal(3), NumberVal(4)}}},
		NumberVal(1),
		BoolVal(true),
	})
	if err != nil {
		t.Fatalf("fnTOROW: %v", err)
	}
	want := []float64{1, 3, 4}
	if len(got.Array[0]) != len(want) {
		t.Fatalf("expected %d cols, got %d", len(want), len(got.Array[0]))
	}
	for i, w := range want {
		if got.Array[0][i].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

// ---------------------------------------------------------------------------
// WRAPROWS
// ---------------------------------------------------------------------------

func TestWRAPROWS_Exact(t *testing.T) {
	// WRAPROWS({1,2,3,4,5,6}, 3) → {1,2,3;4,5,6}
	got, err := fnWRAPROWS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5), NumberVal(6)}}},
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2x3, got %v", got)
	}
	want := [][]float64{{1, 2, 3}, {4, 5, 6}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestWRAPROWS_Padding(t *testing.T) {
	// WRAPROWS({1,2,3,4,5}, 3) → {1,2,3;4,5,#N/A}
	got, err := fnWRAPROWS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5)}}},
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[1]) != 3 {
		t.Fatalf("expected 2x3, got %dx%d", len(got.Array), len(got.Array[1]))
	}
	if got.Array[1][2].Type != ValueError || got.Array[1][2].Err != ErrValNA {
		t.Errorf("expected #N/A padding, got %v", got.Array[1][2])
	}
}

func TestWRAPROWS_CustomPad(t *testing.T) {
	// WRAPROWS({1,2,3,4,5}, 3, 0) → {1,2,3;4,5,0}
	got, err := fnWRAPROWS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5)}}},
		NumberVal(3),
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if got.Array[1][2].Type != ValueNumber || got.Array[1][2].Num != 0 {
		t.Errorf("expected 0 padding, got %v", got.Array[1][2])
	}
}

func TestWRAPROWS_WrapOne(t *testing.T) {
	// WRAPROWS({1,2,3}, 1) → {1;2;3}
	got, err := fnWRAPROWS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}}},
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	for i, w := range []float64{1, 2, 3} {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestWRAPROWS_ZeroWrapCount(t *testing.T) {
	got, err := fnWRAPROWS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}}},
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestWRAPROWS_NegativeWrapCount(t *testing.T) {
	got, err := fnWRAPROWS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}}},
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestWRAPROWS_SingleElement(t *testing.T) {
	// WRAPROWS({5}, 1) → 5
	got, err := fnWRAPROWS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(5)}}},
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("expected 5, got %v", got)
	}
}

func TestWRAPROWS_WrapLargerThanVector(t *testing.T) {
	// WRAPROWS({1,2,3}, 5) → {1,2,3,#N/A,#N/A}
	got, err := fnWRAPROWS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}}},
		NumberVal(5),
	})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 5 {
		t.Fatalf("expected 1x5, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[0][3].Type != ValueError || got.Array[0][3].Err != ErrValNA {
		t.Errorf("[3]: expected #N/A, got %v", got.Array[0][3])
	}
	if got.Array[0][4].Type != ValueError || got.Array[0][4].Err != ErrValNA {
		t.Errorf("[4]: expected #N/A, got %v", got.Array[0][4])
	}
}

func TestWRAPROWS_NoArgs(t *testing.T) {
	got, err := fnWRAPROWS([]Value{})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestWRAPROWS_TooManyArgs(t *testing.T) {
	got, err := fnWRAPROWS([]Value{NumberVal(1), NumberVal(1), NumberVal(0), NumberVal(0)})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestWRAPROWS_ErrorPassthrough(t *testing.T) {
	got, err := fnWRAPROWS([]Value{ErrorVal(ErrValVALUE), NumberVal(2)})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestWRAPROWS_ColumnVector(t *testing.T) {
	// WRAPROWS({1;2;3;4;5;6}, 2) → {1,2;3,4;5,6}
	got, err := fnWRAPROWS([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
			{NumberVal(4)}, {NumberVal(5)}, {NumberVal(6)},
		}},
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if len(got.Array) != 3 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 3x2, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	want := [][]float64{{1, 2}, {3, 4}, {5, 6}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestWRAPROWS_StringPad(t *testing.T) {
	// WRAPROWS({1,2,3,4,5}, 3, "x") → {1,2,3;4,5,"x"}
	got, err := fnWRAPROWS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5)}}},
		NumberVal(3),
		StringVal("x"),
	})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if got.Array[1][2].Type != ValueString || got.Array[1][2].Str != "x" {
		t.Errorf("expected 'x' padding, got %v", got.Array[1][2])
	}
}

func TestWRAPROWS_Scalar(t *testing.T) {
	// WRAPROWS(5, 1) → 5
	got, err := fnWRAPROWS([]Value{NumberVal(5), NumberVal(1)})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("expected 5, got %v", got)
	}
}

func TestWRAPROWS_WrapTwo(t *testing.T) {
	// WRAPROWS({1,2,3,4,5}, 2) → {1,2;3,4;5,#N/A}
	got, err := fnWRAPROWS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5)}}},
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	if got.Array[2][1].Type != ValueError || got.Array[2][1].Err != ErrValNA {
		t.Errorf("expected #N/A at [2][1], got %v", got.Array[2][1])
	}
}

func TestWRAPROWS_MixedTypes(t *testing.T) {
	got, err := fnWRAPROWS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), StringVal("a"), BoolVal(true), ErrorVal(ErrValNA)}}},
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[0][0].Type != ValueNumber || got.Array[0][1].Type != ValueString {
		t.Errorf("row 0 types wrong")
	}
	if got.Array[1][0].Type != ValueBool || got.Array[1][1].Type != ValueError {
		t.Errorf("row 1 types wrong")
	}
}

func TestWRAPROWS_WrapEqualToLength(t *testing.T) {
	// WRAPROWS({1,2,3}, 3) → {1,2,3}
	got, err := fnWRAPROWS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}}},
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnWRAPROWS: %v", err)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3, got %dx%d", len(got.Array), len(got.Array[0]))
	}
}

func TestWRAPROWS_ViaEval(t *testing.T) {
	cf := evalCompile(t, "WRAPROWS({1,2,3,4,5,6}, 3)")
	got, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 2x3, got %v", got)
	}
	for i, w := range []float64{1, 2, 3} {
		if got.Array[0][i].Num != w {
			t.Errorf("[0][%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
	for i, w := range []float64{4, 5, 6} {
		if got.Array[1][i].Num != w {
			t.Errorf("[1][%d]: got %g, want %g", i, got.Array[1][i].Num, w)
		}
	}
}

// ---------------------------------------------------------------------------
// WRAPCOLS
// ---------------------------------------------------------------------------

func TestWRAPCOLS_Exact(t *testing.T) {
	// WRAPCOLS({1,2,3,4,5,6}, 3) → {1,4;2,5;3,6}
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5), NumberVal(6)}}},
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3x2, got %v", got)
	}
	want := [][]float64{{1, 4}, {2, 5}, {3, 6}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestWRAPCOLS_Padding(t *testing.T) {
	// WRAPCOLS({1,2,3,4,5}, 3) → {1,4;2,5;3,#N/A}
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5)}}},
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if len(got.Array) != 3 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 3x2, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[2][1].Type != ValueError || got.Array[2][1].Err != ErrValNA {
		t.Errorf("expected #N/A padding, got %v", got.Array[2][1])
	}
}

func TestWRAPCOLS_CustomPad(t *testing.T) {
	// WRAPCOLS({1,2,3,4,5}, 3, 0) → {1,4;2,5;3,0}
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5)}}},
		NumberVal(3),
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if got.Array[2][1].Type != ValueNumber || got.Array[2][1].Num != 0 {
		t.Errorf("expected 0 padding, got %v", got.Array[2][1])
	}
}

func TestWRAPCOLS_WrapOne(t *testing.T) {
	// WRAPCOLS({1,2,3}, 1) → {1,2,3}
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}}},
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	for i, w := range []float64{1, 2, 3} {
		if got.Array[0][i].Num != w {
			t.Errorf("[0][%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestWRAPCOLS_ZeroWrapCount(t *testing.T) {
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}}},
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestWRAPCOLS_NegativeWrapCount(t *testing.T) {
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}}},
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestWRAPCOLS_SingleElement(t *testing.T) {
	// WRAPCOLS({5}, 1) → 5
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(5)}}},
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("expected 5, got %v", got)
	}
}

func TestWRAPCOLS_WrapLargerThanVector(t *testing.T) {
	// WRAPCOLS({1,2,3}, 5) → column of {1;2;3;#N/A;#N/A}
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}}},
		NumberVal(5),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if len(got.Array) != 5 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 5x1, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[3][0].Type != ValueError || got.Array[3][0].Err != ErrValNA {
		t.Errorf("expected #N/A padding at [3][0], got %v", got.Array[3][0])
	}
}

func TestWRAPCOLS_NoArgs(t *testing.T) {
	got, err := fnWRAPCOLS([]Value{})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestWRAPCOLS_TooManyArgs(t *testing.T) {
	got, err := fnWRAPCOLS([]Value{NumberVal(1), NumberVal(1), NumberVal(0), NumberVal(0)})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestWRAPCOLS_ErrorPassthrough(t *testing.T) {
	got, err := fnWRAPCOLS([]Value{ErrorVal(ErrValVALUE), NumberVal(2)})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestWRAPCOLS_ColumnVector(t *testing.T) {
	// WRAPCOLS({1;2;3;4;5;6}, 2) → {1,3,5;2,4,6}
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
			{NumberVal(4)}, {NumberVal(5)}, {NumberVal(6)},
		}},
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 2x3, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	want := [][]float64{{1, 3, 5}, {2, 4, 6}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestWRAPCOLS_StringPad(t *testing.T) {
	// WRAPCOLS({1,2,3,4,5}, 3, "x") → {1,4;2,5;3,"x"}
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5)}}},
		NumberVal(3),
		StringVal("x"),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if got.Array[2][1].Type != ValueString || got.Array[2][1].Str != "x" {
		t.Errorf("expected 'x' padding, got %v", got.Array[2][1])
	}
}

func TestWRAPCOLS_Scalar(t *testing.T) {
	// WRAPCOLS(5, 1) → 5
	got, err := fnWRAPCOLS([]Value{NumberVal(5), NumberVal(1)})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("expected 5, got %v", got)
	}
}

func TestWRAPCOLS_WrapTwo(t *testing.T) {
	// WRAPCOLS({1,2,3,4,5}, 2) → {1,3,5;2,4,#N/A}
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5)}}},
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 2x3, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[0][2].Num != 5 {
		t.Errorf("[0][2]: expected 5, got %g", got.Array[0][2].Num)
	}
	if got.Array[1][2].Type != ValueError || got.Array[1][2].Err != ErrValNA {
		t.Errorf("[1][2]: expected #N/A, got %v", got.Array[1][2])
	}
}

func TestWRAPCOLS_MixedTypes(t *testing.T) {
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), StringVal("a"), BoolVal(true), NumberVal(4)}}},
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[0][0].Num != 1 {
		t.Errorf("[0][0]: expected 1, got %v", got.Array[0][0])
	}
	if got.Array[1][0].Str != "a" {
		t.Errorf("[1][0]: expected 'a', got %v", got.Array[1][0])
	}
	if got.Array[0][1].Type != ValueBool || !got.Array[0][1].Bool {
		t.Errorf("[0][1]: expected TRUE, got %v", got.Array[0][1])
	}
	if got.Array[1][1].Num != 4 {
		t.Errorf("[1][1]: expected 4, got %v", got.Array[1][1])
	}
}

func TestWRAPCOLS_WrapEqualToLength(t *testing.T) {
	// WRAPCOLS({1,2,3}, 3) → {1;2;3}
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}}},
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if len(got.Array) != 3 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 3x1, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	for i, w := range []float64{1, 2, 3} {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d][0]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestWRAPCOLS_ViaEval(t *testing.T) {
	cf := evalCompile(t, "WRAPCOLS({1,2,3,4,5,6}, 3)")
	got, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatalf("eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 3x2 array, got %v", got)
	}
	want := [][]float64{{1, 4}, {2, 5}, {3, 6}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestWRAPCOLS_EightElements(t *testing.T) {
	// WRAPCOLS({1,2,3,4,5,6,7,8}, 4) → {1,5;2,6;3,7;4,8}
	got, err := fnWRAPCOLS([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5), NumberVal(6), NumberVal(7), NumberVal(8)}}},
		NumberVal(4),
	})
	if err != nil {
		t.Fatalf("fnWRAPCOLS: %v", err)
	}
	if len(got.Array) != 4 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 4x2, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	want := [][]float64{{1, 5}, {2, 6}, {3, 7}, {4, 8}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

// ---------------------------------------------------------------------------
// HSTACK
// ---------------------------------------------------------------------------

func TestHSTACK_TwoColumnVectors(t *testing.T) {
	// HSTACK({1;2},{3;4}) → {1,3;2,4}
	got, err := fnHSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3)}, {NumberVal(4)}}},
	})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %v", got)
	}
	want := [][]float64{{1, 3}, {2, 4}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestHSTACK_TwoRowArrays(t *testing.T) {
	// HSTACK({1,2},{3,4}) → {1,2,3,4}
	got, err := fnHSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), NumberVal(4)}}},
	})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 4 {
		t.Fatalf("expected 1x4, got %v", got)
	}
	for i, w := range []float64{1, 2, 3, 4} {
		if got.Array[0][i].Num != w {
			t.Errorf("[0][%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestHSTACK_Scalars(t *testing.T) {
	// HSTACK(1,2,3) → {1,2,3}
	got, err := fnHSTACK([]Value{NumberVal(1), NumberVal(2), NumberVal(3)})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3, got %v", got)
	}
	for i, w := range []float64{1, 2, 3} {
		if got.Array[0][i].Num != w {
			t.Errorf("[0][%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestHSTACK_DifferentRowCounts(t *testing.T) {
	// HSTACK({1;2;3},{4;5}) → {1,4;2,5;3,#N/A}
	got, err := fnHSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(4)}, {NumberVal(5)}}},
	})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if len(got.Array) != 3 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 3x2, got %v", got)
	}
	if got.Array[2][1].Type != ValueError || got.Array[2][1].Err != ErrValNA {
		t.Errorf("expected #N/A at [2][1], got %v", got.Array[2][1])
	}
}

func TestHSTACK_SingleArray(t *testing.T) {
	// HSTACK({1,2;3,4}) → {1,2;3,4}
	got, err := fnHSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}, {NumberVal(3), NumberVal(4)}}},
	})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %v", got)
	}
	want := [][]float64{{1, 2}, {3, 4}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestHSTACK_SingleScalar(t *testing.T) {
	// HSTACK(42) → 42
	got, err := fnHSTACK([]Value{NumberVal(42)})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("expected 42, got %v", got)
	}
}

func TestHSTACK_ThreeArrays(t *testing.T) {
	// HSTACK({1},{2},{3}) with column vectors → {1,2,3}
	got, err := fnHSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3)}}},
	})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3, got %v", got)
	}
	for i, w := range []float64{1, 2, 3} {
		if got.Array[0][i].Num != w {
			t.Errorf("[0][%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestHSTACK_ErrorPassthrough(t *testing.T) {
	got, err := fnHSTACK([]Value{ErrorVal(ErrValVALUE), NumberVal(1)})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestHSTACK_NoArgs(t *testing.T) {
	got, err := fnHSTACK([]Value{})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestHSTACK_MixedScalarAndArray(t *testing.T) {
	// HSTACK(1, {2;3}) → {1,2;#N/A,3}  (scalar treated as 1x1, padded)
	got, err := fnHSTACK([]Value{
		NumberVal(1),
		{Type: ValueArray, Array: [][]Value{{NumberVal(2)}, {NumberVal(3)}}},
	})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %v", got)
	}
	if got.Array[0][0].Num != 1 {
		t.Errorf("[0][0]: expected 1, got %v", got.Array[0][0])
	}
	if got.Array[1][0].Type != ValueError || got.Array[1][0].Err != ErrValNA {
		t.Errorf("[1][0]: expected #N/A, got %v", got.Array[1][0])
	}
}

func TestHSTACK_TwoByTwoArrays(t *testing.T) {
	// HSTACK({1,2;3,4},{5,6;7,8}) → {1,2,5,6;3,4,7,8}
	got, err := fnHSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}, {NumberVal(3), NumberVal(4)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(5), NumberVal(6)}, {NumberVal(7), NumberVal(8)}}},
	})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 4 {
		t.Fatalf("expected 2x4, got %v", got)
	}
	want := [][]float64{{1, 2, 5, 6}, {3, 4, 7, 8}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestHSTACK_StringValues(t *testing.T) {
	got, err := fnHSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{StringVal("a")}}},
		{Type: ValueArray, Array: [][]Value{{StringVal("b")}}},
	})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 1x2, got %v", got)
	}
	if got.Array[0][0].Str != "a" || got.Array[0][1].Str != "b" {
		t.Errorf("expected [a,b], got %v", got.Array[0])
	}
}

func TestHSTACK_MultipleRowPadding(t *testing.T) {
	// HSTACK({1;2;3;4},{5}) → {1,5;2,#N/A;3,#N/A;4,#N/A}
	got, err := fnHSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)}, {NumberVal(4)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(5)}}},
	})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if len(got.Array) != 4 {
		t.Fatalf("expected 4 rows, got %d", len(got.Array))
	}
	if got.Array[0][1].Num != 5 {
		t.Errorf("[0][1]: expected 5, got %v", got.Array[0][1])
	}
	for i := 1; i < 4; i++ {
		if got.Array[i][1].Type != ValueError || got.Array[i][1].Err != ErrValNA {
			t.Errorf("[%d][1]: expected #N/A, got %v", i, got.Array[i][1])
		}
	}
}

func TestHSTACK_ViaEval(t *testing.T) {
	cf := evalCompile(t, "HSTACK({1;2;3},{4;5;6})")
	got, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatalf("eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 3x2 array, got %v", got)
	}
	want := [][]float64{{1, 4}, {2, 5}, {3, 6}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestHSTACK_BoolValues(t *testing.T) {
	got, err := fnHSTACK([]Value{BoolVal(true), BoolVal(false)})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 1x2, got %v", got)
	}
	if got.Array[0][0].Type != ValueBool || !got.Array[0][0].Bool {
		t.Errorf("[0][0]: expected TRUE, got %v", got.Array[0][0])
	}
	if got.Array[0][1].Type != ValueBool || got.Array[0][1].Bool {
		t.Errorf("[0][1]: expected FALSE, got %v", got.Array[0][1])
	}
}

func TestHSTACK_FourScalars(t *testing.T) {
	// HSTACK(1,2,3,4) → {1,2,3,4}
	got, err := fnHSTACK([]Value{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 4 {
		t.Fatalf("expected 1x4, got %v", got)
	}
	for i, w := range []float64{1, 2, 3, 4} {
		if got.Array[0][i].Num != w {
			t.Errorf("[0][%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

func TestHSTACK_DifferentWidths(t *testing.T) {
	// HSTACK({1,2},{3}) → {1,2,3}
	got, err := fnHSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3)}}},
	})
	if err != nil {
		t.Fatalf("fnHSTACK: %v", err)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3, got %v", got)
	}
	for i, w := range []float64{1, 2, 3} {
		if got.Array[0][i].Num != w {
			t.Errorf("[0][%d]: got %g, want %g", i, got.Array[0][i].Num, w)
		}
	}
}

// ---------------------------------------------------------------------------
// VSTACK
// ---------------------------------------------------------------------------

func TestVSTACK_TwoRowArrays(t *testing.T) {
	// VSTACK({1,2},{3,4}) → {1,2;3,4}
	got, err := fnVSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), NumberVal(4)}}},
	})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %v", got)
	}
	want := [][]float64{{1, 2}, {3, 4}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestVSTACK_TwoColumnVectors(t *testing.T) {
	// VSTACK({1;2},{3;4}) → {1;2;3;4}
	got, err := fnVSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3)}, {NumberVal(4)}}},
	})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if len(got.Array) != 4 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 4x1, got %v", got)
	}
	for i, w := range []float64{1, 2, 3, 4} {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d][0]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestVSTACK_Scalars(t *testing.T) {
	// VSTACK(1,2,3) → {1;2;3}
	got, err := fnVSTACK([]Value{NumberVal(1), NumberVal(2), NumberVal(3)})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if len(got.Array) != 3 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 3x1, got %v", got)
	}
	for i, w := range []float64{1, 2, 3} {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d][0]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestVSTACK_DifferentColumnCounts(t *testing.T) {
	// VSTACK({1,2,3},{4,5}) → {1,2,3;4,5,#N/A}
	got, err := fnVSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(4), NumberVal(5)}}},
	})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 2x3, got %v", got)
	}
	if got.Array[1][2].Type != ValueError || got.Array[1][2].Err != ErrValNA {
		t.Errorf("expected #N/A at [1][2], got %v", got.Array[1][2])
	}
}

func TestVSTACK_SingleArray(t *testing.T) {
	// VSTACK({1,2;3,4}) → {1,2;3,4}
	got, err := fnVSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}, {NumberVal(3), NumberVal(4)}}},
	})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %v", got)
	}
	want := [][]float64{{1, 2}, {3, 4}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestVSTACK_SingleScalar(t *testing.T) {
	// VSTACK(42) → 42
	got, err := fnVSTACK([]Value{NumberVal(42)})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("expected 42, got %v", got)
	}
}

func TestVSTACK_ThreeArrays(t *testing.T) {
	// VSTACK({1,2},{3,4},{5,6}) → {1,2;3,4;5,6}
	got, err := fnVSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), NumberVal(4)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(5), NumberVal(6)}}},
	})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if len(got.Array) != 3 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 3x2, got %v", got)
	}
	want := [][]float64{{1, 2}, {3, 4}, {5, 6}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestVSTACK_ErrorPassthrough(t *testing.T) {
	got, err := fnVSTACK([]Value{ErrorVal(ErrValVALUE), NumberVal(1)})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestVSTACK_NoArgs(t *testing.T) {
	got, err := fnVSTACK([]Value{})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestVSTACK_MixedScalarAndArray(t *testing.T) {
	// VSTACK(1, {2,3}) → {1,#N/A;2,3}
	got, err := fnVSTACK([]Value{
		NumberVal(1),
		{Type: ValueArray, Array: [][]Value{{NumberVal(2), NumberVal(3)}}},
	})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %v", got)
	}
	if got.Array[0][0].Num != 1 {
		t.Errorf("[0][0]: expected 1, got %v", got.Array[0][0])
	}
	if got.Array[0][1].Type != ValueError || got.Array[0][1].Err != ErrValNA {
		t.Errorf("[0][1]: expected #N/A, got %v", got.Array[0][1])
	}
}

func TestVSTACK_TwoByTwoArrays(t *testing.T) {
	// VSTACK({1,2;3,4},{5,6;7,8}) → {1,2;3,4;5,6;7,8}
	got, err := fnVSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}, {NumberVal(3), NumberVal(4)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(5), NumberVal(6)}, {NumberVal(7), NumberVal(8)}}},
	})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if len(got.Array) != 4 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 4x2, got %v", got)
	}
	want := [][]float64{{1, 2}, {3, 4}, {5, 6}, {7, 8}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestVSTACK_StringValues(t *testing.T) {
	got, err := fnVSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{StringVal("a")}}},
		{Type: ValueArray, Array: [][]Value{{StringVal("b")}}},
	})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 2x1, got %v", got)
	}
	if got.Array[0][0].Str != "a" || got.Array[1][0].Str != "b" {
		t.Errorf("expected [a;b], got %v", got)
	}
}

func TestVSTACK_MultipleColumnPadding(t *testing.T) {
	// VSTACK({1,2,3,4},{5}) → {1,2,3,4;5,#N/A,#N/A,#N/A}
	got, err := fnVSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(5)}}},
	})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[1]) != 4 {
		t.Fatalf("expected 2x4, got %v", got)
	}
	if got.Array[1][0].Num != 5 {
		t.Errorf("[1][0]: expected 5, got %v", got.Array[1][0])
	}
	for i := 1; i < 4; i++ {
		if got.Array[1][i].Type != ValueError || got.Array[1][i].Err != ErrValNA {
			t.Errorf("[1][%d]: expected #N/A, got %v", i, got.Array[1][i])
		}
	}
}

func TestVSTACK_ViaEval(t *testing.T) {
	cf := evalCompile(t, "VSTACK({1,2,3},{4,5,6})")
	got, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatalf("eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 2x3 array, got %v", got)
	}
	want := [][]float64{{1, 2, 3}, {4, 5, 6}}
	for i, wr := range want {
		for j, w := range wr {
			if got.Array[i][j].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", i, j, got.Array[i][j].Num, w)
			}
		}
	}
}

func TestVSTACK_BoolValues(t *testing.T) {
	got, err := fnVSTACK([]Value{BoolVal(true), BoolVal(false)})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 2x1, got %v", got)
	}
	if got.Array[0][0].Type != ValueBool || !got.Array[0][0].Bool {
		t.Errorf("[0][0]: expected TRUE, got %v", got.Array[0][0])
	}
	if got.Array[1][0].Type != ValueBool || got.Array[1][0].Bool {
		t.Errorf("[1][0]: expected FALSE, got %v", got.Array[1][0])
	}
}

func TestVSTACK_FourScalars(t *testing.T) {
	// VSTACK(1,2,3,4) → {1;2;3;4}
	got, err := fnVSTACK([]Value{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if len(got.Array) != 4 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 4x1, got %v", got)
	}
	for i, w := range []float64{1, 2, 3, 4} {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d][0]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestVSTACK_DifferentWidths(t *testing.T) {
	// VSTACK({1},{2,3}) → {1,#N/A;2,3}
	got, err := fnVSTACK([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(2), NumberVal(3)}}},
	})
	if err != nil {
		t.Fatalf("fnVSTACK: %v", err)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %v", got)
	}
	if got.Array[0][0].Num != 1 {
		t.Errorf("[0][0]: expected 1, got %v", got.Array[0][0])
	}
	if got.Array[0][1].Type != ValueError || got.Array[0][1].Err != ErrValNA {
		t.Errorf("[0][1]: expected #N/A, got %v", got.Array[0][1])
	}
	if got.Array[1][0].Num != 2 {
		t.Errorf("[1][0]: expected 2, got %v", got.Array[1][0])
	}
	if got.Array[1][1].Num != 3 {
		t.Errorf("[1][1]: expected 3, got %v", got.Array[1][1])
	}
}

// mockArrayResolver implements CellResolver and FormulaArrayEvaluator for testing ANCHORARRAY.
type mockArrayResolver struct {
	mockResolver
	arrays map[CellAddr]Value // pre-computed array results keyed by cell address
}

func (m *mockArrayResolver) EvalCellFormula(sheet string, col, row int) Value {
	addr := CellAddr{Sheet: sheet, Col: col, Row: row}
	if v, ok := m.arrays[addr]; ok {
		return v
	}
	return m.GetCellValue(addr)
}

func TestANCHORARRAY(t *testing.T) {
	// Set up a mock where cell A2 (col=1,row=2) has a dynamic array formula
	// that produces a 4-element column array.
	arr := Value{Type: ValueArray, Array: [][]Value{
		{StringVal("Cleotilde")},
		{StringVal("Kenneth")},
		{StringVal("Matilda")},
		{StringVal("Yevette")},
	}}

	resolver := &mockArrayResolver{
		mockResolver: mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 2}: StringVal("Cleotilde"), // scalar value of anchor cell
			},
		},
		arrays: map[CellAddr]Value{
			{Col: 1, Row: 2}: arr, // full array result
		},
	}

	ctx := &EvalContext{
		CurrentCol:   2,
		CurrentRow:   2,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	cf := evalCompile(t, "ANCHORARRAY(A2)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected ValueArray, got %v", got.Type)
	}
	if len(got.Array) != 4 {
		t.Fatalf("expected 4 rows, got %d", len(got.Array))
	}
	if got.Array[0][0].Str != "Cleotilde" {
		t.Errorf("[0][0]: expected Cleotilde, got %v", got.Array[0][0])
	}
	if got.Array[3][0].Str != "Yevette" {
		t.Errorf("[3][0]: expected Yevette, got %v", got.Array[3][0])
	}
}

func TestANCHORARRAY_ScalarFallback(t *testing.T) {
	// When the cell has no array formula, return the scalar value.
	resolver := &mockArrayResolver{
		mockResolver: mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),
			},
		},
		arrays: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(42),
		},
	}

	ctx := &EvalContext{
		CurrentCol:   2,
		CurrentRow:   1,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	cf := evalCompile(t, "ANCHORARRAY(A1)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("expected 42, got %v", got)
	}
}

func TestANCHORARRAY_CrossSheet(t *testing.T) {
	// Test cross-sheet ANCHORARRAY reference.
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(100)},
		{NumberVal(200)},
		{NumberVal(300)},
	}}

	resolver := &mockArrayResolver{
		mockResolver: mockResolver{
			cells: map[CellAddr]Value{},
		},
		arrays: map[CellAddr]Value{
			{Sheet: "Sheet2", Col: 3, Row: 2}: arr,
		},
	}

	ctx := &EvalContext{
		CurrentCol:   1,
		CurrentRow:   1,
		CurrentSheet: "Sheet1",
		Resolver:     resolver,
	}

	cf := evalCompile(t, "SUM(ANCHORARRAY(Sheet2!C2))")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 600 {
		t.Errorf("expected 600, got %v", got)
	}
}

func TestANCHORARRAY_NoResolver(t *testing.T) {
	// Without FormulaArrayEvaluator, fall back to scalar cell value.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(99),
		},
	}

	ctx := &EvalContext{
		CurrentCol:   2,
		CurrentRow:   1,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	cf := evalCompile(t, "ANCHORARRAY(A1)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 99 {
		t.Errorf("expected 99, got %v", got)
	}
}
