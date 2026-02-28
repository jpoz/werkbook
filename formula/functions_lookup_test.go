package formula

import (
	"testing"
)

func TestADDRESS(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{"default abs", "ADDRESS(2,3)", "$C$2"},
		{"abs row relative col", "ADDRESS(2,3,2)", "C$2"},
		{"relative row abs col", "ADDRESS(2,3,3)", "$C2"},
		{"all relative", "ADDRESS(2,3,4)", "C2"},
		{"R1C1 absolute", "ADDRESS(2,3,1,FALSE)", "R2C3"},
		{"R1C1 abs row", "ADDRESS(2,3,2,FALSE)", "R2C[3]"},
		{"R1C1 abs col", "ADDRESS(2,3,3,FALSE)", "R[2]C3"},
		{"R1C1 all relative", "ADDRESS(2,3,4,FALSE)", "R[2]C[3]"},
		{"with sheet", `ADDRESS(1,1,1,TRUE,"Sheet2")`, "Sheet2!$A$1"},
		{"sheet with space", `ADDRESS(1,1,1,TRUE,"My Sheet")`, "'My Sheet'!$A$1"},
		{"multi-letter column", "ADDRESS(1,27)", "$AA$1"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("ADDRESS %s: got %v, want %s", tt.name, got, tt.want)
			}
		})
	}
}

func TestADDRESSErrors(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "ADDRESS(1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ADDRESS too few args: got %v, want #VALUE!", got)
	}

	// Invalid abs_num
	cf = evalCompile(t, "ADDRESS(1,1,5)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ADDRESS invalid abs_num: got %v, want #VALUE!", got)
	}
}

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

func TestXLOOKUP(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: StringVal("b"),
			{Col: 1, Row: 3}: StringVal("c"),
			{Col: 2, Row: 1}: NumberVal(100),
			{Col: 2, Row: 2}: NumberVal(200),
			{Col: 2, Row: 3}: NumberVal(300),
		},
	}

	cf := evalCompile(t, `XLOOKUP("b",A1:A3,B1:B3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 200 {
		t.Errorf("XLOOKUP: got %g, want 200", got.Num)
	}

	// Not found with default
	cf = evalCompile(t, `XLOOKUP("z",A1:A3,B1:B3,"missing")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "missing" {
		t.Errorf("XLOOKUP not found: got %v, want missing", got)
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
// LOOKUP — basic vector form
// ---------------------------------------------------------------------------

func TestLOOKUP(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: StringVal("ten"),
			{Col: 2, Row: 2}: StringVal("twenty"),
			{Col: 2, Row: 3}: StringVal("thirty"),
		},
	}

	// 3-arg form: lookup, lookup_vector, result_vector
	cf := evalCompile(t, "LOOKUP(20,A1:A3,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "twenty" {
		t.Errorf("LOOKUP exact: got %v, want twenty", got)
	}

	// Between values: 25 should find 20's result
	cf = evalCompile(t, "LOOKUP(25,A1:A3,B1:B3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "twenty" {
		t.Errorf("LOOKUP approx: got %v, want twenty", got)
	}

	// Too small => #N/A
	cf = evalCompile(t, "LOOKUP(5,A1:A3,B1:B3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("LOOKUP too small: got %v, want #N/A", got)
	}

	// Arg validation
	cf = evalCompile(t, "LOOKUP(1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("LOOKUP 1 arg: got %v, want #VALUE!", got)
	}
}

// ---------------------------------------------------------------------------
// XLOOKUP — match modes
// ---------------------------------------------------------------------------

func TestXLOOKUPMatchModes(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: StringVal("ten"),
			{Col: 2, Row: 2}: StringVal("twenty"),
			{Col: 2, Row: 3}: StringVal("thirty"),
		},
	}

	// match_mode=-1: exact or next smaller
	cf := evalCompile(t, `XLOOKUP(25,A1:A3,B1:B3,"nope",-1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "twenty" {
		t.Errorf("XLOOKUP next smaller: got %v, want twenty", got)
	}

	// match_mode=1: exact or next larger
	cf = evalCompile(t, `XLOOKUP(25,A1:A3,B1:B3,"nope",1)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "thirty" {
		t.Errorf("XLOOKUP next larger: got %v, want thirty", got)
	}

	// match_mode=0 not found => returns if_not_found
	cf = evalCompile(t, `XLOOKUP(99,A1:A3,B1:B3,"nope",0)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "nope" {
		t.Errorf("XLOOKUP not found: got %v, want nope", got)
	}

	// Default if_not_found is #N/A
	cf = evalCompile(t, `XLOOKUP(99,A1:A3,B1:B3)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("XLOOKUP default not found: got %v, want #N/A", got)
	}
}

func TestXLOOKUPArgErrors(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, `XLOOKUP(1,A1:A3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("XLOOKUP too few: got %v, want #VALUE!", got)
	}
}
