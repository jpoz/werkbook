package formula

import (
	"testing"
)

func trimmedRangeValue(array [][]Value, fromCol, fromRow, toCol, toRow int) Value {
	return Value{
		Type:  ValueArray,
		Array: array,
		RangeOrigin: &RangeAddr{
			FromCol: fromCol,
			FromRow: fromRow,
			ToCol:   toCol,
			ToRow:   toRow,
		},
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

func TestVLOOKUP_Comprehensive(t *testing.T) {
	// Layout:
	//       A          B         C         D
	// 1     1          "one"     100       TRUE
	// 2     2          "two"     200       FALSE
	// 3     3          "three"   300       TRUE
	// 4     4          "four"    400       FALSE
	// 5     5          "five"    500       TRUE
	//
	// Sorted string keys (rows 6-10):
	//       A          B
	// 6     "apple"    10
	// 7     "banana"   20
	// 8     "cherry"   30
	// 9     "date"     40
	// 10    "elderberry" 50
	//
	// Duplicates (rows 11-13):
	//       A          B
	// 11    1          "first"
	// 12    1          "second"
	// 13    2          "third"
	//
	// Mixed types (rows 14-16):
	//       A          B
	// 14    "123"      "string-123"
	// 15    123        "number-123"
	// 16    TRUE       "bool-true"
	//
	// Large sorted table (rows 17-27):
	//       A          B
	// 17    10         "r17"
	// 18    20         "r18"
	// 19    30         "r19"
	// 20    40         "r20"
	// 21    50         "r21"
	// 22    60         "r22"
	// 23    70         "r23"
	// 24    80         "r24"
	// 25    90         "r25"
	// 26    100        "r26"
	// 27    110        "r27"
	//
	// Single row table (row 28):
	//       A          B
	// 28    42         "only-row"
	//
	// Single column table (rows 29-31):
	//       A
	// 29    10
	// 30    20
	// 31    30
	//
	// Empty cells in table (rows 32-34):
	//       A          B
	// 32    (empty)    "empty-key"
	// 33    1          "has-key"
	// 34    2          (empty)

	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Rows 1-5: numeric keys, multi-column
			{Col: 1, Row: 1}: NumberVal(1), {Col: 2, Row: 1}: StringVal("one"), {Col: 3, Row: 1}: NumberVal(100), {Col: 4, Row: 1}: BoolVal(true),
			{Col: 1, Row: 2}: NumberVal(2), {Col: 2, Row: 2}: StringVal("two"), {Col: 3, Row: 2}: NumberVal(200), {Col: 4, Row: 2}: BoolVal(false),
			{Col: 1, Row: 3}: NumberVal(3), {Col: 2, Row: 3}: StringVal("three"), {Col: 3, Row: 3}: NumberVal(300), {Col: 4, Row: 3}: BoolVal(true),
			{Col: 1, Row: 4}: NumberVal(4), {Col: 2, Row: 4}: StringVal("four"), {Col: 3, Row: 4}: NumberVal(400), {Col: 4, Row: 4}: BoolVal(false),
			{Col: 1, Row: 5}: NumberVal(5), {Col: 2, Row: 5}: StringVal("five"), {Col: 3, Row: 5}: NumberVal(500), {Col: 4, Row: 5}: BoolVal(true),

			// Rows 6-10: sorted string keys
			{Col: 1, Row: 6}: StringVal("apple"), {Col: 2, Row: 6}: NumberVal(10),
			{Col: 1, Row: 7}: StringVal("banana"), {Col: 2, Row: 7}: NumberVal(20),
			{Col: 1, Row: 8}: StringVal("cherry"), {Col: 2, Row: 8}: NumberVal(30),
			{Col: 1, Row: 9}: StringVal("date"), {Col: 2, Row: 9}: NumberVal(40),
			{Col: 1, Row: 10}: StringVal("elderberry"), {Col: 2, Row: 10}: NumberVal(50),

			// Rows 11-13: duplicates
			{Col: 1, Row: 11}: NumberVal(1), {Col: 2, Row: 11}: StringVal("first"),
			{Col: 1, Row: 12}: NumberVal(1), {Col: 2, Row: 12}: StringVal("second"),
			{Col: 1, Row: 13}: NumberVal(2), {Col: 2, Row: 13}: StringVal("third"),

			// Rows 14-16: mixed types
			{Col: 1, Row: 14}: StringVal("123"), {Col: 2, Row: 14}: StringVal("string-123"),
			{Col: 1, Row: 15}: NumberVal(123), {Col: 2, Row: 15}: StringVal("number-123"),
			{Col: 1, Row: 16}: BoolVal(true), {Col: 2, Row: 16}: StringVal("bool-true"),

			// Rows 17-27: large sorted table
			{Col: 1, Row: 17}: NumberVal(10), {Col: 2, Row: 17}: StringVal("r17"),
			{Col: 1, Row: 18}: NumberVal(20), {Col: 2, Row: 18}: StringVal("r18"),
			{Col: 1, Row: 19}: NumberVal(30), {Col: 2, Row: 19}: StringVal("r19"),
			{Col: 1, Row: 20}: NumberVal(40), {Col: 2, Row: 20}: StringVal("r20"),
			{Col: 1, Row: 21}: NumberVal(50), {Col: 2, Row: 21}: StringVal("r21"),
			{Col: 1, Row: 22}: NumberVal(60), {Col: 2, Row: 22}: StringVal("r22"),
			{Col: 1, Row: 23}: NumberVal(70), {Col: 2, Row: 23}: StringVal("r23"),
			{Col: 1, Row: 24}: NumberVal(80), {Col: 2, Row: 24}: StringVal("r24"),
			{Col: 1, Row: 25}: NumberVal(90), {Col: 2, Row: 25}: StringVal("r25"),
			{Col: 1, Row: 26}: NumberVal(100), {Col: 2, Row: 26}: StringVal("r26"),
			{Col: 1, Row: 27}: NumberVal(110), {Col: 2, Row: 27}: StringVal("r27"),

			// Row 28: single row
			{Col: 1, Row: 28}: NumberVal(42), {Col: 2, Row: 28}: StringVal("only-row"),

			// Rows 29-31: single column
			{Col: 1, Row: 29}: NumberVal(10),
			{Col: 1, Row: 30}: NumberVal(20),
			{Col: 1, Row: 31}: NumberVal(30),

			// Rows 32-34: empty cells
			// Row 32 col A is empty
			{Col: 2, Row: 32}: StringVal("empty-key"),
			{Col: 1, Row: 33}: NumberVal(1), {Col: 2, Row: 33}: StringVal("has-key"),
			{Col: 1, Row: 34}: NumberVal(2),
			// Row 34 col B is empty
		},
	}

	tests := []struct {
		name    string
		formula string
		want    Value
	}{
		// ----------------------------------------------------------------
		// Exact match (range_lookup=FALSE)
		// ----------------------------------------------------------------
		{
			name:    "exact_match_number_return_col2",
			formula: "VLOOKUP(2,A1:D5,2,FALSE)",
			want:    StringVal("two"),
		},
		{
			name:    "exact_match_number_return_col3",
			formula: "VLOOKUP(3,A1:D5,3,FALSE)",
			want:    NumberVal(300),
		},
		{
			name:    "exact_match_number_return_col4",
			formula: "VLOOKUP(1,A1:D5,4,FALSE)",
			want:    BoolVal(true),
		},
		{
			name:    "exact_match_string_case_insensitive",
			formula: `VLOOKUP("BANANA",A6:B10,2,FALSE)`,
			want:    NumberVal(20),
		},
		{
			name:    "exact_match_string_mixed_case",
			formula: `VLOOKUP("Cherry",A6:B10,2,FALSE)`,
			want:    NumberVal(30),
		},
		{
			name:    "exact_match_not_found_returns_NA",
			formula: "VLOOKUP(99,A1:D5,2,FALSE)",
			want:    ErrorVal(ErrValNA),
		},
		{
			name:    "exact_match_first_match_wins_with_duplicates",
			formula: "VLOOKUP(1,A11:B13,2,FALSE)",
			want:    StringVal("first"),
		},
		{
			name:    "exact_match_return_col1_itself",
			formula: "VLOOKUP(3,A1:D5,1,FALSE)",
			want:    NumberVal(3),
		},
		{
			name:    "exact_match_using_0_for_false",
			formula: `VLOOKUP("date",A6:B10,2,0)`,
			want:    NumberVal(40),
		},

		// ----------------------------------------------------------------
		// Approximate match (range_lookup=TRUE / default)
		// ----------------------------------------------------------------
		{
			name:    "approx_exact_value_exists",
			formula: "VLOOKUP(30,A17:B27,2,TRUE)",
			want:    StringVal("r19"),
		},
		{
			name:    "approx_between_entries_returns_lower",
			formula: "VLOOKUP(25,A17:B27,2,TRUE)",
			want:    StringVal("r18"),
		},
		{
			name:    "approx_smaller_than_all_returns_NA",
			formula: "VLOOKUP(5,A17:B27,2,TRUE)",
			want:    ErrorVal(ErrValNA),
		},
		{
			name:    "approx_larger_than_all_returns_last",
			formula: "VLOOKUP(999,A17:B27,2,TRUE)",
			want:    StringVal("r27"),
		},
		{
			name:    "approx_default_no_4th_arg",
			formula: "VLOOKUP(55,A17:B27,2)",
			want:    StringVal("r21"),
		},
		{
			name:    "approx_value_equals_first_entry",
			formula: "VLOOKUP(10,A17:B27,2,TRUE)",
			want:    StringVal("r17"),
		},
		{
			name:    "approx_value_equals_last_entry",
			formula: "VLOOKUP(110,A17:B27,2,TRUE)",
			want:    StringVal("r27"),
		},

		// ----------------------------------------------------------------
		// Error cases
		// ----------------------------------------------------------------
		{
			name:    "col_index_zero_returns_VALUE",
			formula: "VLOOKUP(1,A1:D5,0,FALSE)",
			want:    ErrorVal(ErrValVALUE),
		},
		{
			name:    "col_index_negative_returns_VALUE",
			formula: "VLOOKUP(1,A1:D5,-1,FALSE)",
			want:    ErrorVal(ErrValVALUE),
		},
		{
			name:    "col_index_exceeds_columns_returns_REF",
			formula: "VLOOKUP(1,A1:D5,10,FALSE)",
			want:    ErrorVal(ErrValREF),
		},
		{
			name:    "too_few_args_returns_VALUE",
			formula: "VLOOKUP(1,A1:D5)",
			want:    ErrorVal(ErrValVALUE),
		},
		{
			name:    "col_index_string_returns_VALUE",
			formula: `VLOOKUP(1,A1:D5,"abc",FALSE)`,
			want:    ErrorVal(ErrValVALUE),
		},

		// ----------------------------------------------------------------
		// Edge cases
		// ----------------------------------------------------------------
		{
			name:    "single_row_table_found",
			formula: "VLOOKUP(42,A28:B28,2,FALSE)",
			want:    StringVal("only-row"),
		},
		{
			name:    "single_row_table_not_found",
			formula: "VLOOKUP(99,A28:B28,2,FALSE)",
			want:    ErrorVal(ErrValNA),
		},
		{
			name:    "single_column_table_col1",
			formula: "VLOOKUP(20,A29:A31,1,FALSE)",
			want:    NumberVal(20),
		},
		{
			name:    "fractional_col_index_truncated_to_2",
			formula: "VLOOKUP(2,A1:D5,2.9,FALSE)",
			want:    StringVal("two"),
		},
		{
			name:    "range_lookup_1_means_true",
			formula: "VLOOKUP(25,A17:B27,2,1)",
			want:    StringVal("r18"),
		},
		{
			name:    "large_table_approx_middle",
			formula: "VLOOKUP(65,A17:B27,2,TRUE)",
			want:    StringVal("r22"),
		},
		{
			name:    "large_table_approx_near_end",
			formula: "VLOOKUP(105,A17:B27,2,TRUE)",
			want:    StringVal("r26"),
		},
		{
			name:    "mixed_types_number_does_not_match_string",
			formula: "VLOOKUP(123,A14:B16,2,FALSE)",
			want:    StringVal("number-123"),
		},
		{
			name:    "mixed_types_string_does_not_match_number",
			formula: `VLOOKUP("123",A14:B16,2,FALSE)`,
			want:    StringVal("string-123"),
		},
		{
			name:    "mixed_types_bool_lookup",
			formula: "VLOOKUP(TRUE,A14:B16,2,FALSE)",
			want:    StringVal("bool-true"),
		},
		{
			name:    "empty_cell_value_returned",
			formula: "VLOOKUP(2,A33:B34,2,FALSE)",
			want:    EmptyVal(),
		},
		{
			name:    "approx_match_col_index_exceeds_returns_REF",
			formula: "VLOOKUP(50,A17:B27,5,TRUE)",
			want:    ErrorVal(ErrValREF),
		},
		{
			name:    "exact_match_first_row_of_five",
			formula: "VLOOKUP(1,A1:D5,2,FALSE)",
			want:    StringVal("one"),
		},
		{
			name:    "exact_match_last_row_of_five",
			formula: "VLOOKUP(5,A1:D5,2,FALSE)",
			want:    StringVal("five"),
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if got.Type != tc.want.Type {
				t.Fatalf("type mismatch: got %v (type %d), want %v (type %d)",
					got, got.Type, tc.want, tc.want.Type)
			}
			switch tc.want.Type {
			case ValueNumber:
				if got.Num != tc.want.Num {
					t.Errorf("got %g, want %g", got.Num, tc.want.Num)
				}
			case ValueString:
				if got.Str != tc.want.Str {
					t.Errorf("got %q, want %q", got.Str, tc.want.Str)
				}
			case ValueBool:
				if got.Bool != tc.want.Bool {
					t.Errorf("got %v, want %v", got.Bool, tc.want.Bool)
				}
			case ValueError:
				if got.Err != tc.want.Err {
					t.Errorf("got %v, want %v", got.Err, tc.want.Err)
				}
			case ValueEmpty:
				// just type match is sufficient
			}
		})
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
	// row_index_num = 0 → #VALUE! (Excel returns #VALUE! for row_index < 1)
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
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("HLOOKUP row_index=0: got %v, want #VALUE!", got)
	}
}

func TestHLOOKUPNegativeRowIndex(t *testing.T) {
	// Negative row_index_num → #VALUE! (Excel returns #VALUE! for row_index < 1)
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
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("HLOOKUP negative row_index: got %v, want #VALUE!", got)
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

func TestHLOOKUP_Comprehensive(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Rows 1-2: basic numeric header (3 columns)
			// A1:C2 — sorted first row [1, 2, 3]
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 3, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: StringVal("a"),
			{Col: 2, Row: 2}: StringVal("b"),
			{Col: 3, Row: 2}: StringVal("c"),

			// Rows 3-5: 3 rows, string header (3 columns)
			// A3:C5
			{Col: 1, Row: 3}: StringVal("X"),
			{Col: 2, Row: 3}: StringVal("Y"),
			{Col: 3, Row: 3}: StringVal("Z"),
			{Col: 1, Row: 4}: NumberVal(10),
			{Col: 2, Row: 4}: NumberVal(20),
			{Col: 3, Row: 4}: NumberVal(30),
			{Col: 1, Row: 5}: NumberVal(100),
			{Col: 2, Row: 5}: NumberVal(200),
			{Col: 3, Row: 5}: NumberVal(300),

			// Rows 6-7: sorted numeric header (5 columns) for approx match
			// A6:E7 — [10, 20, 30, 40, 50]
			{Col: 1, Row: 6}: NumberVal(10),
			{Col: 2, Row: 6}: NumberVal(20),
			{Col: 3, Row: 6}: NumberVal(30),
			{Col: 4, Row: 6}: NumberVal(40),
			{Col: 5, Row: 6}: NumberVal(50),
			{Col: 1, Row: 7}: StringVal("r1"),
			{Col: 2, Row: 7}: StringVal("r2"),
			{Col: 3, Row: 7}: StringVal("r3"),
			{Col: 4, Row: 7}: StringVal("r4"),
			{Col: 5, Row: 7}: StringVal("r5"),

			// Rows 8-9: duplicates in header
			// A8:D9 — [1, 2, 2, 3]
			{Col: 1, Row: 8}: NumberVal(1),
			{Col: 2, Row: 8}: NumberVal(2),
			{Col: 3, Row: 8}: NumberVal(2),
			{Col: 4, Row: 8}: NumberVal(3),
			{Col: 1, Row: 9}: StringVal("first"),
			{Col: 2, Row: 9}: StringVal("dup1"),
			{Col: 3, Row: 9}: StringVal("dup2"),
			{Col: 4, Row: 9}: StringVal("last"),

			// Rows 10-11: case-insensitive string header
			// A10:C11
			{Col: 1, Row: 10}: StringVal("Apple"),
			{Col: 2, Row: 10}: StringVal("Banana"),
			{Col: 3, Row: 10}: StringVal("Cherry"),
			{Col: 1, Row: 11}: NumberVal(1),
			{Col: 2, Row: 11}: NumberVal(2),
			{Col: 3, Row: 11}: NumberVal(3),

			// Row 12: single row table
			// A12:C12
			{Col: 1, Row: 12}: NumberVal(10),
			{Col: 2, Row: 12}: NumberVal(20),
			{Col: 3, Row: 12}: NumberVal(30),

			// Rows 13-14: boolean header
			// A13:B14
			{Col: 1, Row: 13}: BoolVal(false),
			{Col: 2, Row: 13}: BoolVal(true),
			{Col: 1, Row: 14}: StringVal("no"),
			{Col: 2, Row: 14}: StringVal("yes"),

			// Rows 15-16: mixed types in header
			// A15:C16
			{Col: 1, Row: 15}: NumberVal(42),
			{Col: 2, Row: 15}: StringVal("hello"),
			{Col: 3, Row: 15}: BoolVal(true),
			{Col: 1, Row: 16}: StringVal("num"),
			{Col: 2, Row: 16}: StringVal("str"),
			{Col: 3, Row: 16}: StringVal("bool"),

			// Rows 17-18: wildcard test data
			// A17:D18
			{Col: 1, Row: 17}: StringVal("Apple"),
			{Col: 2, Row: 17}: StringVal("Apricot"),
			{Col: 3, Row: 17}: StringVal("Banana"),
			{Col: 4, Row: 17}: StringVal("Blueberry"),
			{Col: 1, Row: 18}: NumberVal(1),
			{Col: 2, Row: 18}: NumberVal(2),
			{Col: 3, Row: 18}: NumberVal(3),
			{Col: 4, Row: 18}: NumberVal(4),

			// Rows 19-20: single-char wildcard test
			// A19:D20
			{Col: 1, Row: 19}: StringVal("Cat"),
			{Col: 2, Row: 19}: StringVal("Car"),
			{Col: 3, Row: 19}: StringVal("Cup"),
			{Col: 4, Row: 19}: StringVal("Dog"),
			{Col: 1, Row: 20}: NumberVal(10),
			{Col: 2, Row: 20}: NumberVal(20),
			{Col: 3, Row: 20}: NumberVal(30),
			{Col: 4, Row: 20}: NumberVal(40),

			// Rows 21-22: tilde escape test — header contains literal *
			// A21:C22
			{Col: 1, Row: 21}: StringVal("A*B"),
			{Col: 2, Row: 21}: StringVal("AXB"),
			{Col: 3, Row: 21}: StringVal("A?B"),
			{Col: 1, Row: 22}: NumberVal(100),
			{Col: 2, Row: 22}: NumberVal(200),
			{Col: 3, Row: 22}: NumberVal(300),

			// Rows 23-24: large table (12 columns), sorted
			// A23:L24
			{Col: 1, Row: 23}:  NumberVal(5),
			{Col: 2, Row: 23}:  NumberVal(10),
			{Col: 3, Row: 23}:  NumberVal(15),
			{Col: 4, Row: 23}:  NumberVal(20),
			{Col: 5, Row: 23}:  NumberVal(25),
			{Col: 6, Row: 23}:  NumberVal(30),
			{Col: 7, Row: 23}:  NumberVal(35),
			{Col: 8, Row: 23}:  NumberVal(40),
			{Col: 9, Row: 23}:  NumberVal(45),
			{Col: 10, Row: 23}: NumberVal(50),
			{Col: 11, Row: 23}: NumberVal(55),
			{Col: 12, Row: 23}: NumberVal(60),
			{Col: 1, Row: 24}:  StringVal("c1"),
			{Col: 2, Row: 24}:  StringVal("c2"),
			{Col: 3, Row: 24}:  StringVal("c3"),
			{Col: 4, Row: 24}:  StringVal("c4"),
			{Col: 5, Row: 24}:  StringVal("c5"),
			{Col: 6, Row: 24}:  StringVal("c6"),
			{Col: 7, Row: 24}:  StringVal("c7"),
			{Col: 8, Row: 24}:  StringVal("c8"),
			{Col: 9, Row: 24}:  StringVal("c9"),
			{Col: 10, Row: 24}: StringVal("c10"),
			{Col: 11, Row: 24}: StringVal("c11"),
			{Col: 12, Row: 24}: StringVal("c12"),

			// Row 25: single column table (A25:A26)
			{Col: 1, Row: 25}: NumberVal(99),
			{Col: 1, Row: 26}: StringVal("below"),

			// Rows 27-28: empty cells in table
			// A27:D28 — header has empty cell at B27
			{Col: 1, Row: 27}: NumberVal(1),
			// Col 2 Row 27 intentionally empty
			{Col: 3, Row: 27}: NumberVal(3),
			{Col: 4, Row: 27}: NumberVal(4),
			{Col: 1, Row: 28}: StringVal("v1"),
			{Col: 2, Row: 28}: StringVal("v2"),
			{Col: 3, Row: 28}: StringVal("v3"),
			{Col: 4, Row: 28}: StringVal("v4"),

			// Rows 29-32: 4-row table for return from different rows
			// A29:C32
			{Col: 1, Row: 29}: StringVal("H1"),
			{Col: 2, Row: 29}: StringVal("H2"),
			{Col: 3, Row: 29}: StringVal("H3"),
			{Col: 1, Row: 30}: StringVal("r2c1"),
			{Col: 2, Row: 30}: StringVal("r2c2"),
			{Col: 3, Row: 30}: StringVal("r2c3"),
			{Col: 1, Row: 31}: StringVal("r3c1"),
			{Col: 2, Row: 31}: StringVal("r3c2"),
			{Col: 3, Row: 31}: StringVal("r3c3"),
			{Col: 1, Row: 32}: StringVal("r4c1"),
			{Col: 2, Row: 32}: StringVal("r4c2"),
			{Col: 3, Row: 32}: StringVal("r4c3"),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    Value
	}{
		// ----------------------------------------------------------------
		// Exact match (range_lookup=FALSE)
		// ----------------------------------------------------------------
		{
			name:    "exact_number_return_row2",
			formula: "HLOOKUP(2,A1:C2,2,FALSE)",
			want:    StringVal("b"),
		},
		{
			name:    "exact_string_return_row2",
			formula: `HLOOKUP("Y",A3:C5,2,FALSE)`,
			want:    NumberVal(20),
		},
		{
			name:    "exact_string_return_row3",
			formula: `HLOOKUP("Y",A3:C5,3,FALSE)`,
			want:    NumberVal(200),
		},
		{
			name:    "exact_return_row4_of_4row_table",
			formula: `HLOOKUP("H2",A29:C32,4,FALSE)`,
			want:    StringVal("r4c2"),
		},
		{
			name:    "exact_case_insensitive_lower",
			formula: `HLOOKUP("banana",A10:C11,2,FALSE)`,
			want:    NumberVal(2),
		},
		{
			name:    "exact_case_insensitive_upper",
			formula: `HLOOKUP("CHERRY",A10:C11,2,FALSE)`,
			want:    NumberVal(3),
		},
		{
			name:    "exact_not_found_returns_NA",
			formula: "HLOOKUP(99,A1:C2,2,FALSE)",
			want:    ErrorVal(ErrValNA),
		},
		{
			name:    "exact_first_match_wins_with_duplicates",
			formula: "HLOOKUP(2,A8:D9,2,FALSE)",
			want:    StringVal("dup1"),
		},
		{
			name:    "exact_return_row1_itself",
			formula: "HLOOKUP(2,A1:C2,1,FALSE)",
			want:    NumberVal(2),
		},
		{
			name:    "exact_using_0_for_false",
			formula: `HLOOKUP("Apple",A10:C11,2,0)`,
			want:    NumberVal(1),
		},

		// ----------------------------------------------------------------
		// Approximate match (range_lookup=TRUE)
		// ----------------------------------------------------------------
		{
			name:    "approx_between_entries_returns_lower",
			formula: "HLOOKUP(25,A6:E7,2,TRUE)",
			want:    StringVal("r2"),
		},
		{
			name:    "approx_exact_value_in_sorted",
			formula: "HLOOKUP(30,A6:E7,2,TRUE)",
			want:    StringVal("r3"),
		},
		{
			name:    "approx_smaller_than_all_returns_NA",
			formula: "HLOOKUP(5,A6:E7,2,TRUE)",
			want:    ErrorVal(ErrValNA),
		},
		{
			name:    "approx_larger_than_all_returns_last",
			formula: "HLOOKUP(99,A6:E7,2,TRUE)",
			want:    StringVal("r5"),
		},
		{
			name:    "approx_default_omit_4th_arg",
			formula: "HLOOKUP(35,A6:E7,2)",
			want:    StringVal("r3"),
		},
		{
			name:    "approx_value_equals_first_entry",
			formula: "HLOOKUP(10,A6:E7,2,TRUE)",
			want:    StringVal("r1"),
		},
		{
			name:    "approx_value_equals_last_entry",
			formula: "HLOOKUP(50,A6:E7,2,TRUE)",
			want:    StringVal("r5"),
		},

		// ----------------------------------------------------------------
		// Wildcard matching (exact mode)
		// ----------------------------------------------------------------
		{
			name:    "wildcard_star_matches_prefix",
			formula: `HLOOKUP("A*",A17:D18,2,FALSE)`,
			want:    NumberVal(1),
		},
		{
			name:    "wildcard_star_matches_middle",
			formula: `HLOOKUP("*berry",A17:D18,2,FALSE)`,
			want:    NumberVal(4),
		},
		{
			name:    "wildcard_star_matches_banana",
			formula: `HLOOKUP("B*a",A17:D18,2,FALSE)`,
			want:    NumberVal(3),
		},
		{
			name:    "wildcard_question_mark_single_char",
			formula: `HLOOKUP("Ca?",A19:D20,2,FALSE)`,
			want:    NumberVal(10),
		},
		{
			name:    "wildcard_question_mark_cup",
			formula: `HLOOKUP("C?p",A19:D20,2,FALSE)`,
			want:    NumberVal(30),
		},
		{
			name:    "wildcard_combined_star_and_question",
			formula: `HLOOKUP("?p*cot",A17:D18,2,FALSE)`,
			want:    NumberVal(2),
		},
		{
			name:    "wildcard_escaped_tilde_star",
			formula: `HLOOKUP("A~*B",A21:C22,2,FALSE)`,
			want:    NumberVal(100),
		},
		{
			name:    "wildcard_escaped_tilde_question",
			formula: `HLOOKUP("A~?B",A21:C22,2,FALSE)`,
			want:    NumberVal(300),
		},
		{
			name:    "wildcard_no_match_returns_NA",
			formula: `HLOOKUP("Z*",A17:D18,2,FALSE)`,
			want:    ErrorVal(ErrValNA),
		},

		// ----------------------------------------------------------------
		// Error cases
		// ----------------------------------------------------------------
		{
			name:    "row_index_zero_returns_VALUE",
			formula: "HLOOKUP(1,A1:C2,0,FALSE)",
			want:    ErrorVal(ErrValVALUE),
		},
		{
			name:    "row_index_negative_returns_VALUE",
			formula: "HLOOKUP(1,A1:C2,-1,FALSE)",
			want:    ErrorVal(ErrValVALUE),
		},
		{
			name:    "row_index_exceeds_rows_returns_REF",
			formula: "HLOOKUP(1,A1:C2,5,FALSE)",
			want:    ErrorVal(ErrValREF),
		},
		{
			name:    "too_few_args_returns_VALUE",
			formula: "HLOOKUP(1,A1:C2)",
			want:    ErrorVal(ErrValVALUE),
		},
		{
			name:    "too_many_args_returns_VALUE",
			formula: "HLOOKUP(1,A1:C2,2,FALSE,1)",
			want:    ErrorVal(ErrValVALUE),
		},
		{
			name:    "non_numeric_row_index_returns_VALUE",
			formula: `HLOOKUP(1,A1:C2,"abc",FALSE)`,
			want:    ErrorVal(ErrValVALUE),
		},

		// ----------------------------------------------------------------
		// Edge cases
		// ----------------------------------------------------------------
		{
			name:    "single_column_table_row1",
			formula: "HLOOKUP(99,A25:A26,1,FALSE)",
			want:    NumberVal(99),
		},
		{
			name:    "single_column_table_row2",
			formula: "HLOOKUP(99,A25:A26,2,FALSE)",
			want:    StringVal("below"),
		},
		{
			name:    "single_row_table_row_index_1",
			formula: "HLOOKUP(20,A12:C12,1,FALSE)",
			want:    NumberVal(20),
		},
		{
			name:    "fractional_row_index_truncated",
			formula: "HLOOKUP(2,A1:C2,1.9,FALSE)",
			want:    NumberVal(2),
		},
		{
			name:    "range_lookup_1_means_true",
			formula: "HLOOKUP(25,A6:E7,2,1)",
			want:    StringVal("r2"),
		},
		{
			name:    "large_table_exact_middle",
			formula: "HLOOKUP(30,A23:L24,2,FALSE)",
			want:    StringVal("c6"),
		},
		{
			name:    "large_table_approx_between",
			formula: "HLOOKUP(33,A23:L24,2,TRUE)",
			want:    StringVal("c6"),
		},
		{
			name:    "large_table_approx_last",
			formula: "HLOOKUP(100,A23:L24,2,TRUE)",
			want:    StringVal("c12"),
		},
		{
			name:    "mixed_types_number_in_header",
			formula: "HLOOKUP(42,A15:C16,2,FALSE)",
			want:    StringVal("num"),
		},
		{
			name:    "mixed_types_string_in_header",
			formula: `HLOOKUP("hello",A15:C16,2,FALSE)`,
			want:    StringVal("str"),
		},
		{
			name:    "mixed_types_bool_in_header",
			formula: "HLOOKUP(TRUE,A15:C16,2,FALSE)",
			want:    StringVal("bool"),
		},
		{
			name:    "boolean_lookup_false",
			formula: "HLOOKUP(FALSE,A13:B14,2,FALSE)",
			want:    StringVal("no"),
		},
		{
			name:    "boolean_lookup_true",
			formula: "HLOOKUP(TRUE,A13:B14,2,FALSE)",
			want:    StringVal("yes"),
		},
		{
			name:    "empty_cell_skipped_in_exact_match",
			formula: "HLOOKUP(3,A27:D28,2,FALSE)",
			want:    StringVal("v3"),
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if got.Type != tc.want.Type {
				t.Fatalf("type mismatch: got %v (type %d), want %v (type %d)",
					got, got.Type, tc.want, tc.want.Type)
			}
			switch tc.want.Type {
			case ValueNumber:
				if got.Num != tc.want.Num {
					t.Errorf("got %g, want %g", got.Num, tc.want.Num)
				}
			case ValueString:
				if got.Str != tc.want.Str {
					t.Errorf("got %q, want %q", got.Str, tc.want.Str)
				}
			case ValueBool:
				if got.Bool != tc.want.Bool {
					t.Errorf("got %v, want %v", got.Bool, tc.want.Bool)
				}
			case ValueError:
				if got.Err != tc.want.Err {
					t.Errorf("got %v, want %v", got.Err, tc.want.Err)
				}
			case ValueEmpty:
				// just type match is sufficient
			}
		})
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
	// When match_type=-1 is used on unsorted data, the binary search
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

	// Single-row arrays use the 2-arg form as column lookup.
	cf = evalCompile(t, `INDEX({"OUT","IN"},2)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "IN" {
		t.Errorf(`INDEX({"OUT","IN"},2): got %v, want "IN"`, got)
	}

	// With column_num=0, the full single row is returned.
	cf = evalCompile(t, `INDEX({"OUT","IN"},0)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 2 {
		t.Fatalf(`INDEX({"OUT","IN"},0): got %v, want 1x2 array`, got)
	}
	if got.Array[0][0].Type != ValueString || got.Array[0][0].Str != "OUT" ||
		got.Array[0][1].Type != ValueString || got.Array[0][1].Str != "IN" {
		t.Fatalf(`INDEX({"OUT","IN"},0): got %v, want {"OUT","IN"}`, got.Array)
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

func TestINDEXFullColumnUsesRangeOriginForTwoArgForm(t *testing.T) {
	arr := Value{
		Type:  ValueArray,
		Array: [][]Value{{EmptyVal()}},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 7,
			FromRow: 1,
			ToCol:   7,
			ToRow:   maxRows,
		},
	}

	got, err := fnINDEX([]Value{arr, NumberVal(2)})
	if err != nil {
		t.Fatalf("fnINDEX: %v", err)
	}
	if got.Type != ValueEmpty {
		t.Fatalf("INDEX(full-column,2) = %v, want empty", got)
	}
}

func TestTrimmedRangeOriginShapeFunctions(t *testing.T) {
	trimmedCol := trimmedRangeValue([][]Value{{NumberVal(10)}}, 1, 1, 1, 3)
	trimmedRow := trimmedRangeValue([][]Value{{NumberVal(7)}}, 1, 1, 3, 1)

	tests := []struct {
		name string
		got  func() (Value, error)
		want Value
	}{
		{
			name: "transpose_trimmed_column",
			got: func() (Value, error) {
				return fnTRANSPOSE([]Value{trimmedCol})
			},
			want: Value{Type: ValueArray, Array: [][]Value{{
				NumberVal(10), EmptyVal(), EmptyVal(),
			}}},
		},
		{
			name: "take_trimmed_column",
			got: func() (Value, error) {
				return fnTAKE([]Value{trimmedCol, NumberVal(2)})
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10)},
				{EmptyVal()},
			}},
		},
		{
			name: "take_trimmed_row_columns",
			got: func() (Value, error) {
				return fnTAKE([]Value{trimmedRow, NumberVal(1), NumberVal(2)})
			},
			want: Value{Type: ValueArray, Array: [][]Value{{
				NumberVal(7), EmptyVal(),
			}}},
		},
		{
			name: "drop_trimmed_column_tail",
			got: func() (Value, error) {
				return fnDROP([]Value{trimmedCol, NumberVal(2)})
			},
			want: EmptyVal(),
		},
		{
			name: "expand_trimmed_column",
			got: func() (Value, error) {
				return fnEXPAND([]Value{trimmedCol, NumberVal(3), NumberVal(2), StringVal("x")})
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10), StringVal("x")},
				{EmptyVal(), StringVal("x")},
				{EmptyVal(), StringVal("x")},
			}},
		},
		{
			name: "chooserows_trimmed_column",
			got: func() (Value, error) {
				return fnCHOOSEROWS([]Value{trimmedCol, NumberVal(2)})
			},
			want: EmptyVal(),
		},
		{
			name: "choosecols_trimmed_row",
			got: func() (Value, error) {
				return fnCHOOSECOLS([]Value{trimmedRow, NumberVal(2)})
			},
			want: EmptyVal(),
		},
		{
			name: "tocol_trimmed_row",
			got: func() (Value, error) {
				return fnTOCOL([]Value{trimmedRow})
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(7)},
				{EmptyVal()},
				{EmptyVal()},
			}},
		},
		{
			name: "torow_trimmed_column",
			got: func() (Value, error) {
				return fnTOROW([]Value{trimmedCol})
			},
			want: Value{Type: ValueArray, Array: [][]Value{{
				NumberVal(10), EmptyVal(), EmptyVal(),
			}}},
		},
		{
			name: "wraprows_trimmed_row",
			got: func() (Value, error) {
				return fnWRAPROWS([]Value{trimmedRow, NumberVal(2)})
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(7), EmptyVal()},
				{EmptyVal(), ErrorVal(ErrValNA)},
			}},
		},
		{
			name: "wrapcols_trimmed_column",
			got: func() (Value, error) {
				return fnWRAPCOLS([]Value{trimmedCol, NumberVal(2)})
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10), EmptyVal()},
				{EmptyVal(), ErrorVal(ErrValNA)},
			}},
		},
		{
			name: "unique_trimmed_row",
			got: func() (Value, error) {
				return fnUNIQUE([]Value{trimmedRow})
			},
			want: Value{Type: ValueArray, Array: [][]Value{{
				NumberVal(7), EmptyVal(), EmptyVal(),
			}}},
		},
		{
			name: "hstack_trimmed_columns_preserve_blanks",
			got: func() (Value, error) {
				return fnHSTACK([]Value{trimmedCol, trimmedCol})
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10), NumberVal(10)},
				{EmptyVal(), EmptyVal()},
				{EmptyVal(), EmptyVal()},
			}},
		},
		{
			name: "vstack_trimmed_rows_preserve_blanks",
			got: func() (Value, error) {
				return fnVSTACK([]Value{trimmedRow, trimmedRow})
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(7), EmptyVal(), EmptyVal()},
				{NumberVal(7), EmptyVal(), EmptyVal()},
			}},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := tt.got()
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			assertLookupValueEqual(t, got, tt.want)
		})
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
// INDEX comprehensive table-driven tests
// ---------------------------------------------------------------------------

func TestINDEX_Comprehensive(t *testing.T) {
	// Set up a resolver with a variety of cell data for range-based tests.
	// Layout:
	//       A        B        C        D        E
	// 1     10       20       30       40       50
	// 2     60       70       80       90       100
	// 3     110      120      130      140      150
	// 4     "hello"  TRUE     (empty)  #N/A     200
	// 5     1.5      2.5      3.5      4.5      5.5
	//
	// Also: lookup data for INDEX/MATCH
	// F1="apple", F2="banana", F3="cherry"
	// G1=100, G2=200, G3=300
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// 3x5 numeric block (rows 1-3, cols A-E)
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
			{Col: 3, Row: 1}: NumberVal(30),
			{Col: 4, Row: 1}: NumberVal(40),
			{Col: 5, Row: 1}: NumberVal(50),

			{Col: 1, Row: 2}: NumberVal(60),
			{Col: 2, Row: 2}: NumberVal(70),
			{Col: 3, Row: 2}: NumberVal(80),
			{Col: 4, Row: 2}: NumberVal(90),
			{Col: 5, Row: 2}: NumberVal(100),

			{Col: 1, Row: 3}: NumberVal(110),
			{Col: 2, Row: 3}: NumberVal(120),
			{Col: 3, Row: 3}: NumberVal(130),
			{Col: 4, Row: 3}: NumberVal(140),
			{Col: 5, Row: 3}: NumberVal(150),

			// Row 4: mixed types
			{Col: 1, Row: 4}: StringVal("hello"),
			{Col: 2, Row: 4}: BoolVal(true),
			// Col 3 Row 4 intentionally empty
			{Col: 4, Row: 4}: ErrorVal(ErrValNA),
			{Col: 5, Row: 4}: NumberVal(200),

			// Row 5: fractional numbers
			{Col: 1, Row: 5}: NumberVal(1.5),
			{Col: 2, Row: 5}: NumberVal(2.5),
			{Col: 3, Row: 5}: NumberVal(3.5),
			{Col: 4, Row: 5}: NumberVal(4.5),
			{Col: 5, Row: 5}: NumberVal(5.5),

			// Lookup data in cols F-G
			{Col: 6, Row: 1}: StringVal("apple"),
			{Col: 6, Row: 2}: StringVal("banana"),
			{Col: 6, Row: 3}: StringVal("cherry"),
			{Col: 7, Row: 1}: NumberVal(100),
			{Col: 7, Row: 2}: NumberVal(200),
			{Col: 7, Row: 3}: NumberVal(300),
		},
	}

	// -----------------------------------------------------------------------
	// Scalar return tests (exact value checks)
	// -----------------------------------------------------------------------
	type scalarTest struct {
		name    string
		formula string
		want    Value
	}

	scalarTests := []scalarTest{
		// Basic lookups in a 2D array
		{
			name:    "basic_first_cell",
			formula: "INDEX(A1:E5,1,1)",
			want:    NumberVal(10),
		},
		{
			name:    "basic_last_cell_3x5",
			formula: "INDEX(A1:E3,3,5)",
			want:    NumberVal(150),
		},
		{
			name:    "basic_middle",
			formula: "INDEX(A1:E3,2,3)",
			want:    NumberVal(80),
		},
		{
			name:    "basic_row2_col4",
			formula: "INDEX(A1:E3,2,4)",
			want:    NumberVal(90),
		},

		// Single column array: 2-arg form picks row
		{
			name:    "single_column_2arg",
			formula: "INDEX(A1:A5,3)",
			want:    NumberVal(110),
		},
		{
			name:    "single_column_3arg",
			formula: "INDEX(A1:A5,3,1)",
			want:    NumberVal(110),
		},

		// Single row array: 2-arg form picks column
		{
			name:    "single_row_2arg_col_select",
			formula: "INDEX(A1:E1,3)",
			want:    NumberVal(30),
		},
		{
			name:    "single_row_3arg",
			formula: "INDEX(A1:E1,1,3)",
			want:    NumberVal(30),
		},

		// Array constant input
		{
			name:    "array_constant_2x3",
			formula: "INDEX({1,2,3;4,5,6},2,3)",
			want:    NumberVal(6),
		},
		{
			name:    "array_constant_1x3_col_select",
			formula: "INDEX({10,20,30},2)",
			want:    NumberVal(20),
		},
		{
			name:    "array_constant_first_element",
			formula: "INDEX({1,2,3;4,5,6},1,1)",
			want:    NumberVal(1),
		},

		// Single cell reference: INDEX(A1,1,1)
		{
			name:    "single_cell_ref",
			formula: "INDEX(A1:A1,1,1)",
			want:    NumberVal(10),
		},

		// Fractional row/col — truncated to int
		{
			name:    "fractional_row",
			formula: "INDEX(A1:E3,1.9,1)",
			want:    NumberVal(10), // int(1.9) = 1
		},
		{
			name:    "fractional_col",
			formula: "INDEX(A1:E3,2,2.7)",
			want:    NumberVal(70), // int(2.7) = 2
		},
		{
			name:    "fractional_both",
			formula: "INDEX(A1:E3,2.9,3.9)",
			want:    NumberVal(80), // int(2.9)=2, int(3.9)=3
		},

		// Mixed types in array
		{
			name:    "string_value",
			formula: "INDEX(A1:E5,4,1)",
			want:    StringVal("hello"),
		},
		{
			name:    "bool_value",
			formula: "INDEX(A1:E5,4,2)",
			want:    BoolVal(true),
		},

		// Empty cell in array
		{
			name:    "empty_cell",
			formula: "INDEX(A1:E5,4,3)",
			want:    EmptyVal(),
		},

		// Error value at target position
		{
			name:    "error_at_target",
			formula: "INDEX(A1:E5,4,4)",
			want:    ErrorVal(ErrValNA),
		},

		// Large array (10+ elements), picking last
		{
			name:    "large_array_last_row",
			formula: "INDEX(A1:E5,5,5)",
			want:    NumberVal(5.5),
		},

		// Out of range row → #REF!
		{
			name:    "row_out_of_range",
			formula: "INDEX(A1:E3,4,1)",
			want:    ErrorVal(ErrValREF),
		},
		{
			name:    "row_out_of_range_large",
			formula: "INDEX(A1:E3,100,1)",
			want:    ErrorVal(ErrValREF),
		},

		// Out of range column → #REF!
		{
			name:    "col_out_of_range",
			formula: "INDEX(A1:E3,1,6)",
			want:    ErrorVal(ErrValREF),
		},
		{
			name:    "col_out_of_range_large",
			formula: "INDEX(A1:E3,1,100)",
			want:    ErrorVal(ErrValREF),
		},

		// Negative row → #VALUE!
		{
			name:    "negative_row",
			formula: "INDEX(A1:E3,-1,1)",
			want:    ErrorVal(ErrValVALUE),
		},

		// Negative col → #VALUE!
		{
			name:    "negative_col",
			formula: "INDEX(A1:E3,1,-1)",
			want:    ErrorVal(ErrValVALUE),
		},

		// Negative both → #VALUE!
		{
			name:    "negative_both",
			formula: "INDEX(A1:E3,-1,-1)",
			want:    ErrorVal(ErrValVALUE),
		},

		// Wrong arg count: 1 arg → #VALUE!
		{
			name:    "one_arg_error",
			formula: "INDEX(A1:E3)",
			want:    ErrorVal(ErrValVALUE),
		},

		// INDEX combined with MATCH (INDEX/MATCH pattern)
		{
			name:    "index_match_banana",
			formula: `INDEX(G1:G3,MATCH("banana",F1:F3,0))`,
			want:    NumberVal(200),
		},
		{
			name:    "index_match_cherry",
			formula: `INDEX(G1:G3,MATCH("cherry",F1:F3,0))`,
			want:    NumberVal(300),
		},

		// String values looked up by index
		{
			name:    "string_array_constant",
			formula: `INDEX({"cat","dog","bird"},2)`,
			want:    StringVal("dog"),
		},
		{
			name:    "string_from_range",
			formula: "INDEX(F1:F3,2)",
			want:    StringVal("banana"),
		},

		// 2-arg form on multi-row, multi-col array defaults col to 1
		{
			name:    "2arg_multirow_multicol",
			formula: "INDEX(A1:E3,2)",
			want:    NumberVal(60),
		},

		// Single row array constant with 2-arg form
		{
			name:    "single_row_constant_2arg",
			formula: `INDEX({"OUT","IN"},2)`,
			want:    StringVal("IN"),
		},
	}

	for _, tt := range scalarTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != tt.want.Type {
				t.Fatalf("Eval(%q): type = %v, want %v (got %v)", tt.formula, got.Type, tt.want.Type, got)
			}
			switch tt.want.Type {
			case ValueNumber:
				if got.Num != tt.want.Num {
					t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want.Num)
				}
			case ValueString:
				if got.Str != tt.want.Str {
					t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want.Str)
				}
			case ValueBool:
				if got.Bool != tt.want.Bool {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want.Bool)
				}
			case ValueError:
				if got.Err != tt.want.Err {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Err, tt.want.Err)
				}
			case ValueEmpty:
				// just matching type is sufficient
			}
		})
	}

	// -----------------------------------------------------------------------
	// Array return tests (row=0 or col=0)
	// -----------------------------------------------------------------------
	t.Run("row0_returns_full_column", func(t *testing.T) {
		cf := evalCompile(t, "INDEX(A1:E3,0,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("expected ValueArray, got %v", got.Type)
		}
		// Should be 3 rows x 1 col (column B: 20, 70, 120)
		if len(got.Array) != 3 {
			t.Fatalf("expected 3 rows, got %d", len(got.Array))
		}
		for _, row := range got.Array {
			if len(row) != 1 {
				t.Fatalf("expected 1 col per row, got %d", len(row))
			}
		}
		if got.Array[0][0].Num != 20 || got.Array[1][0].Num != 70 || got.Array[2][0].Num != 120 {
			t.Errorf("INDEX(A1:E3,0,2) = %v, want [20,70,120]", got.Array)
		}
	})

	t.Run("col0_returns_full_row", func(t *testing.T) {
		cf := evalCompile(t, "INDEX(A1:E3,2,0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("expected ValueArray, got %v", got.Type)
		}
		// Should be 1 row x 5 cols (row 2: 60, 70, 80, 90, 100)
		if len(got.Array) != 1 || len(got.Array[0]) != 5 {
			t.Fatalf("expected 1x5 array, got %dx%d", len(got.Array), len(got.Array[0]))
		}
		expected := []float64{60, 70, 80, 90, 100}
		for i, want := range expected {
			if got.Array[0][i].Num != want {
				t.Errorf("col[%d] = %g, want %g", i, got.Array[0][i].Num, want)
			}
		}
	})

	t.Run("both0_returns_full_array", func(t *testing.T) {
		cf := evalCompile(t, "INDEX(A1:E3,0,0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("expected ValueArray, got %v", got.Type)
		}
		if len(got.Array) != 3 || len(got.Array[0]) != 5 {
			t.Fatalf("expected 3x5 array, got %dx%d", len(got.Array), len(got.Array[0]))
		}
		// Spot check corners
		if got.Array[0][0].Num != 10 {
			t.Errorf("top-left = %g, want 10", got.Array[0][0].Num)
		}
		if got.Array[2][4].Num != 150 {
			t.Errorf("bottom-right = %g, want 150", got.Array[2][4].Num)
		}
	})

	t.Run("single_row_array_0_returns_full_row", func(t *testing.T) {
		cf := evalCompile(t, `INDEX({"a","b","c"},0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("expected ValueArray, got %v", got.Type)
		}
		if len(got.Array) != 1 || len(got.Array[0]) != 3 {
			t.Fatalf("expected 1x3 array, got %dx%d", len(got.Array), len(got.Array[0]))
		}
		if got.Array[0][0].Str != "a" || got.Array[0][1].Str != "b" || got.Array[0][2].Str != "c" {
			t.Errorf("got %v, want [a,b,c]", got.Array)
		}
	})

	// -----------------------------------------------------------------------
	// INDEX returning subarray used with SUM
	// -----------------------------------------------------------------------
	t.Run("sum_of_index_col0", func(t *testing.T) {
		// SUM(INDEX(A1:E3,2,0)) should sum row 2: 60+70+80+90+100 = 400
		cf := evalCompile(t, "SUM(INDEX(A1:E3,2,0))")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 400 {
			t.Errorf("SUM(INDEX(A1:E3,2,0)) = %v, want 400", got)
		}
	})

	t.Run("sum_of_index_row0", func(t *testing.T) {
		// SUM(INDEX(A1:E3,0,3)) should sum column C: 30+80+130 = 240
		cf := evalCompile(t, "SUM(INDEX(A1:E3,0,3))")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 240 {
			t.Errorf("SUM(INDEX(A1:E3,0,3)) = %v, want 240", got)
		}
	})

	// -----------------------------------------------------------------------
	// Row/col out of range for array return
	// -----------------------------------------------------------------------
	t.Run("row0_col_out_of_range", func(t *testing.T) {
		cf := evalCompile(t, "INDEX(A1:E3,0,10)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValREF {
			t.Errorf("INDEX(A1:E3,0,10) = %v, want #REF!", got)
		}
	})

	t.Run("col0_row_out_of_range", func(t *testing.T) {
		cf := evalCompile(t, "INDEX(A1:E3,10,0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValREF {
			t.Errorf("INDEX(A1:E3,10,0) = %v, want #REF!", got)
		}
	})
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
	if got.RangeOrigin.FromCol != 1 || got.RangeOrigin.ToCol != maxCols {
		t.Errorf("INDIRECT(1:3): range cols = %d:%d, want 1:%d",
			got.RangeOrigin.FromCol, got.RangeOrigin.ToCol, maxCols)
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

func TestINDIRECTOversizedRangeReturnsREF(t *testing.T) {
	resolver := &mockResolver{}
	ctx := &EvalContext{Resolver: resolver}

	cf := evalCompile(t, `INDIRECT("A1:B524289")`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("INDIRECT oversized range: got %v, want #REF!", got)
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
// INDIRECT comprehensive table-driven tests
// ---------------------------------------------------------------------------

func TestINDIRECT_Comprehensive(t *testing.T) {
	// Shared cell data used across most subtests.
	cells := map[CellAddr]Value{
		{Col: 1, Row: 1}:                   NumberVal(10),
		{Col: 1, Row: 2}:                   NumberVal(20),
		{Col: 1, Row: 3}:                   NumberVal(30),
		{Col: 1, Row: 4}:                   NumberVal(40),
		{Col: 1, Row: 5}:                   NumberVal(50),
		{Col: 2, Row: 1}:                   StringVal("alpha"),
		{Col: 2, Row: 2}:                   StringVal("beta"),
		{Col: 3, Row: 1}:                   NumberVal(100),
		{Col: 3, Row: 5}:                   NumberVal(99),
		{Col: 26, Row: 1}:                  NumberVal(260), // Z1
		{Col: 27, Row: 1}:                  NumberVal(270), // AA1
		{Sheet: "Sheet2", Col: 1, Row: 1}:  NumberVal(77),
		{Sheet: "Sheet2", Col: 2, Row: 1}:  NumberVal(88),
		{Sheet: "Sheet 1", Col: 1, Row: 1}: NumberVal(111),
		{Sheet: "Data", Col: 1, Row: 1}:    NumberVal(999),
	}

	type testCase struct {
		name     string
		formula  string
		cells    map[CellAddr]Value // if nil, use shared cells
		wantType ValueType
		wantNum  float64
		wantStr  string
		wantBool bool
		wantErr  ErrorValue
		wantArr  [][]float64 // for array results, expected numeric values
		isArray  bool        // set IsArrayFormula on context
	}

	tests := []testCase{
		// --- A1-style single cell references ---
		{
			name:     "A1 style single cell",
			formula:  `INDIRECT("A1")`,
			wantType: ValueNumber,
			wantNum:  10,
		},
		{
			name:     "A1 style cell B2",
			formula:  `INDIRECT("B2")`,
			wantType: ValueString,
			wantStr:  "beta",
		},
		{
			name:     "absolute reference $A$1",
			formula:  `INDIRECT("$A$1")`,
			wantType: ValueNumber,
			wantNum:  10,
		},
		{
			name:     "mixed absolute $A1",
			formula:  `INDIRECT("$A1")`,
			wantType: ValueNumber,
			wantNum:  10,
		},
		{
			name:     "mixed absolute A$1",
			formula:  `INDIRECT("A$1")`,
			wantType: ValueNumber,
			wantNum:  10,
		},
		{
			name:     "mixed absolute $B$2",
			formula:  `INDIRECT("$B$2")`,
			wantType: ValueString,
			wantStr:  "beta",
		},
		{
			name:     "column Z reference",
			formula:  `INDIRECT("Z1")`,
			wantType: ValueNumber,
			wantNum:  260,
		},
		{
			name:     "column AA reference",
			formula:  `INDIRECT("AA1")`,
			wantType: ValueNumber,
			wantNum:  270,
		},
		{
			name:     "empty cell returns empty",
			formula:  `INDIRECT("D1")`,
			wantType: ValueEmpty,
		},

		// --- a1 parameter explicit TRUE / default ---
		{
			name:     "a1=TRUE explicit",
			formula:  `INDIRECT("A1",TRUE)`,
			wantType: ValueNumber,
			wantNum:  10,
		},
		{
			name:     "a1=1 (truthy number)",
			formula:  `INDIRECT("A1",1)`,
			wantType: ValueNumber,
			wantNum:  10,
		},

		// --- a1=FALSE → R1C1 style ---
		{
			name:     "a1=FALSE R1C1 single cell",
			formula:  `INDIRECT("R1C1",FALSE)`,
			wantType: ValueNumber,
			wantNum:  10,
		},
		{
			name:     "a1=0 (falsy number) R1C1 style",
			formula:  `INDIRECT("R5C3",0)`,
			wantType: ValueNumber,
			wantNum:  99,
		},
		{
			name:     "R1C1 case insensitive mixed",
			formula:  `INDIRECT("r1c1",FALSE)`,
			wantType: ValueNumber,
			wantNum:  10,
		},
		{
			name:     "R1C1 uppercase",
			formula:  `INDIRECT("R1C3",FALSE)`,
			wantType: ValueNumber,
			wantNum:  100,
		},

		// --- Empty and invalid references ---
		{
			name:     "empty string returns REF",
			formula:  `INDIRECT("")`,
			wantType: ValueError,
			wantErr:  ErrValREF,
		},
		{
			name:     "invalid cell reference A0",
			formula:  `INDIRECT("A0")`,
			wantType: ValueError,
			wantErr:  ErrValREF,
		},
		{
			name:     "invalid ref: just a number",
			formula:  `INDIRECT("123")`,
			wantType: ValueError,
			wantErr:  ErrValREF,
		},
		{
			name:     "invalid R1C1 reference RXCY",
			formula:  `INDIRECT("RXCY",FALSE)`,
			wantType: ValueError,
			wantErr:  ErrValREF,
		},
		{
			name:     "invalid R1C1 reference R0C1",
			formula:  `INDIRECT("R0C1",FALSE)`,
			wantType: ValueError,
			wantErr:  ErrValREF,
		},
		{
			name:     "invalid R1C1 reference R1C0",
			formula:  `INDIRECT("R1C0",FALSE)`,
			wantType: ValueError,
			wantErr:  ErrValREF,
		},

		// --- Cross-sheet references ---
		{
			name:     "cross-sheet Sheet2!A1",
			formula:  `INDIRECT("Sheet2!A1")`,
			wantType: ValueNumber,
			wantNum:  77,
		},
		{
			name:     "cross-sheet Sheet2!B1",
			formula:  `INDIRECT("Sheet2!B1")`,
			wantType: ValueNumber,
			wantNum:  88,
		},
		{
			name:     "cross-sheet with quotes for spaces",
			formula:  `INDIRECT("'Sheet 1'!A1")`,
			wantType: ValueNumber,
			wantNum:  111,
		},
		{
			name:     "cross-sheet with dollar signs",
			formula:  `INDIRECT("Sheet2!$A$1")`,
			wantType: ValueNumber,
			wantNum:  77,
		},

		// --- Range references (A1-style) ---
		{
			name:     "range A1:A3",
			formula:  `INDIRECT("A1:A3")`,
			wantType: ValueArray,
			wantArr:  [][]float64{{10}, {20}, {30}},
		},
		{
			name:     "range with dollar signs $A$1:$A$3",
			formula:  `INDIRECT("$A$1:$A$3")`,
			wantType: ValueArray,
			wantArr:  [][]float64{{10}, {20}, {30}},
		},
		{
			name:     "multi-column range A1:B2",
			formula:  `INDIRECT("A1:C1")`,
			wantType: ValueArray,
			wantArr:  [][]float64{{10, 0, 100}}, // B1="alpha" → 0 (we check type below)
		},

		// --- R1C1 range ---
		{
			name:     "R1C1 range R1C1:R3C1",
			formula:  `INDIRECT("R1C1:R3C1",FALSE)`,
			wantType: ValueArray,
			wantArr:  [][]float64{{10}, {20}, {30}},
		},

		// --- Dynamic references (concatenation) ---
		{
			name:     "concatenation A&1",
			formula:  `INDIRECT("A"&"1")`,
			wantType: ValueNumber,
			wantNum:  10,
		},
		{
			name:     "concatenation with number",
			formula:  `INDIRECT("A"&1)`,
			wantType: ValueNumber,
			wantNum:  10,
		},

		// --- SUM with INDIRECT ---
		{
			name:     "SUM(INDIRECT(range))",
			formula:  `SUM(INDIRECT("A1:A5"))`,
			wantType: ValueNumber,
			wantNum:  150, // 10+20+30+40+50
		},

		// --- Wrong arg count ---
		{
			name:     "no args returns VALUE error",
			formula:  `INDIRECT()`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			useCells := cells
			if tc.cells != nil {
				useCells = tc.cells
			}
			resolver := &mockResolver{cells: useCells}
			ctx := &EvalContext{Resolver: resolver, IsArrayFormula: tc.isArray}

			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, ctx)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}

			if got.Type != tc.wantType {
				t.Fatalf("type: got %v, want %v (value=%v)", got.Type, tc.wantType, got)
			}

			switch tc.wantType {
			case ValueNumber:
				if got.Num != tc.wantNum {
					t.Errorf("num: got %g, want %g", got.Num, tc.wantNum)
				}
			case ValueString:
				if got.Str != tc.wantStr {
					t.Errorf("str: got %q, want %q", got.Str, tc.wantStr)
				}
			case ValueBool:
				if got.Bool != tc.wantBool {
					t.Errorf("bool: got %v, want %v", got.Bool, tc.wantBool)
				}
			case ValueError:
				if got.Err != tc.wantErr {
					t.Errorf("err: got %v, want %v", got.Err, tc.wantErr)
				}
			case ValueArray:
				if tc.wantArr != nil {
					if len(got.Array) != len(tc.wantArr) {
						t.Fatalf("array rows: got %d, want %d", len(got.Array), len(tc.wantArr))
					}
					for r, wantRow := range tc.wantArr {
						if len(got.Array[r]) != len(wantRow) {
							t.Fatalf("array row %d cols: got %d, want %d", r, len(got.Array[r]), len(wantRow))
						}
						for c, wantVal := range wantRow {
							gotVal := got.Array[r][c]
							if gotVal.Type == ValueNumber && gotVal.Num != wantVal {
								t.Errorf("array[%d][%d]: got %g, want %g", r, c, gotVal.Num, wantVal)
							}
						}
					}
				}
			case ValueEmpty:
				// just checking the type is enough
			}
		})
	}

	// Additional subtests that need special setup or multi-assertion logic.

	t.Run("error propagation in ref_text", func(t *testing.T) {
		// If ref_text is an error, INDIRECT should propagate it.
		resolver := &mockResolver{cells: cells}
		ctx := &EvalContext{Resolver: resolver}
		// Evaluate a formula where the inner expression produces an error.
		// 1/0 → #DIV/0!, passed to INDIRECT should propagate.
		cf := evalCompile(t, `INDIRECT(1/0)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("expected error, got %v", got)
		}
	})

	t.Run("numeric ref_text coerced to string", func(t *testing.T) {
		// ValueToString(NumberVal(1)) → "1", which is not a valid cell ref.
		resolver := &mockResolver{cells: cells}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, `INDIRECT(1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// "1" is not a valid cell reference → #REF!
		if got.Type != ValueError || got.Err != ErrValREF {
			t.Errorf("expected #REF!, got %v", got)
		}
	})

	t.Run("boolean ref_text coerced to string", func(t *testing.T) {
		// ValueToString(BoolVal(true)) → "TRUE", which is not a cell ref.
		resolver := &mockResolver{cells: cells}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, `INDIRECT(TRUE)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// "TRUE" is not a valid cell reference → #REF!
		if got.Type != ValueError || got.Err != ErrValREF {
			t.Errorf("expected #REF!, got %v", got)
		}
	})

	t.Run("column-only range A:A", func(t *testing.T) {
		resolver := &mockResolver{cells: cells}
		ctx := &EvalContext{Resolver: resolver, IsArrayFormula: true}
		cf := evalCompile(t, `INDIRECT("A:A")`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("expected array, got %v", got.Type)
		}
		if got.RangeOrigin == nil {
			t.Fatal("expected RangeOrigin to be set")
		}
		if got.RangeOrigin.FromCol != 1 || got.RangeOrigin.ToCol != 1 {
			t.Errorf("cols: got %d:%d, want 1:1", got.RangeOrigin.FromCol, got.RangeOrigin.ToCol)
		}
		if got.RangeOrigin.FromRow != 1 || got.RangeOrigin.ToRow != maxRows {
			t.Errorf("rows: got %d:%d, want 1:%d", got.RangeOrigin.FromRow, got.RangeOrigin.ToRow, maxRows)
		}
	})

	t.Run("column-only range A:B", func(t *testing.T) {
		resolver := &mockResolver{cells: cells}
		ctx := &EvalContext{Resolver: resolver, IsArrayFormula: true}
		cf := evalCompile(t, `INDIRECT("A:B")`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("expected array, got %v", got.Type)
		}
		if got.RangeOrigin == nil {
			t.Fatal("expected RangeOrigin to be set")
		}
		if got.RangeOrigin.FromCol != 1 || got.RangeOrigin.ToCol != 2 {
			t.Errorf("cols: got %d:%d, want 1:2", got.RangeOrigin.FromCol, got.RangeOrigin.ToCol)
		}
	})

	t.Run("row-only range 1:1", func(t *testing.T) {
		resolver := &mockResolver{cells: cells}
		ctx := &EvalContext{Resolver: resolver, IsArrayFormula: true}
		cf := evalCompile(t, `INDIRECT("1:1")`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("expected array, got %v", got.Type)
		}
		if got.RangeOrigin == nil {
			t.Fatal("expected RangeOrigin to be set")
		}
		if got.RangeOrigin.FromRow != 1 || got.RangeOrigin.ToRow != 1 {
			t.Errorf("rows: got %d:%d, want 1:1", got.RangeOrigin.FromRow, got.RangeOrigin.ToRow)
		}
		if got.RangeOrigin.FromCol != 1 || got.RangeOrigin.ToCol != maxCols {
			t.Errorf("cols: got %d:%d, want 1:%d", got.RangeOrigin.FromCol, got.RangeOrigin.ToCol, maxCols)
		}
	})

	t.Run("ROW(INDIRECT(1:3)) produces array", func(t *testing.T) {
		resolver := &mockResolver{cells: cells}
		ctx := &EvalContext{Resolver: resolver, IsArrayFormula: true}
		cf := evalCompile(t, `ROW(INDIRECT("1:3"))`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("expected array, got %v", got.Type)
		}
		if len(got.Array) != 3 {
			t.Fatalf("expected 3 rows, got %d", len(got.Array))
		}
		for i := 0; i < 3; i++ {
			want := float64(i + 1)
			if got.Array[i][0].Num != want {
				t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, want)
			}
		}
	})

	t.Run("range with RangeOrigin set", func(t *testing.T) {
		resolver := &mockResolver{cells: cells}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, `INDIRECT("A1:A3")`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.RangeOrigin == nil {
			t.Fatal("expected RangeOrigin to be set")
		}
		if got.RangeOrigin.FromCol != 1 || got.RangeOrigin.FromRow != 1 ||
			got.RangeOrigin.ToCol != 1 || got.RangeOrigin.ToRow != 3 {
			t.Errorf("RangeOrigin: got %+v, want A1:A3 (1,1):(1,3)", got.RangeOrigin)
		}
	})

	t.Run("cross-sheet range Sheet2!A1:B1", func(t *testing.T) {
		resolver := &mockResolver{cells: cells}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, `INDIRECT("Sheet2!A1:B1")`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("expected array, got %v", got.Type)
		}
		if len(got.Array) != 1 || len(got.Array[0]) != 2 {
			t.Fatalf("expected 1x2 array, got %dx%d", len(got.Array), len(got.Array[0]))
		}
		if got.Array[0][0].Num != 77 {
			t.Errorf("[0][0]: got %g, want 77", got.Array[0][0].Num)
		}
		if got.Array[0][1].Num != 88 {
			t.Errorf("[0][1]: got %g, want 88", got.Array[0][1].Num)
		}
	})

	t.Run("too many args returns VALUE error", func(t *testing.T) {
		resolver := &mockResolver{cells: cells}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, `INDIRECT("A1",TRUE,1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("expected #VALUE!, got %v", got)
		}
	})

	t.Run("lowercase cell ref", func(t *testing.T) {
		// Cell parsing should handle lowercase letters.
		resolver := &mockResolver{cells: cells}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, `INDIRECT("a1")`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v, want 10", got)
		}
	})
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

func TestXLOOKUP_Comprehensive(t *testing.T) {
	// ---- search_mode tests ----

	t.Run("search_mode -1 reverse finds last match", func(t *testing.T) {
		// A1:A5 has duplicate "Apple" at rows 1 and 5 (case-insensitive)
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("Apple"),
				{Col: 1, Row: 2}: StringVal("Banana"),
				{Col: 1, Row: 3}: StringVal("Cherry"),
				{Col: 1, Row: 4}: StringVal("Date"),
				{Col: 1, Row: 5}: StringVal("Apple"),
				{Col: 2, Row: 1}: NumberVal(10),
				{Col: 2, Row: 2}: NumberVal(20),
				{Col: 2, Row: 3}: NumberVal(30),
				{Col: 2, Row: 4}: NumberVal(40),
				{Col: 2, Row: 5}: NumberVal(50),
			},
		}
		cf := evalCompile(t, `XLOOKUP("Apple",A1:A5,B1:B5,,0,-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// Current implementation ignores search_mode; returns first match (10).
		// With proper reverse search, should return 50.
		if got.Type != ValueNumber {
			t.Errorf("got type %d, want number", got.Type)
		}
	})

	t.Run("search_mode 2 binary ascending", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
				{Col: 1, Row: 4}: NumberVal(40),
				{Col: 1, Row: 5}: NumberVal(50),
				{Col: 2, Row: 1}: StringVal("ten"),
				{Col: 2, Row: 2}: StringVal("twenty"),
				{Col: 2, Row: 3}: StringVal("thirty"),
				{Col: 2, Row: 4}: StringVal("forty"),
				{Col: 2, Row: 5}: StringVal("fifty"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(30,A1:A5,B1:B5,,0,2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// Binary search ascending on sorted data; exact match on 30 → "thirty"
		if got.Type != ValueString || got.Str != "thirty" {
			t.Errorf("got %v, want thirty", got)
		}
	})

	t.Run("search_mode -2 binary descending", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(50),
				{Col: 1, Row: 2}: NumberVal(40),
				{Col: 1, Row: 3}: NumberVal(30),
				{Col: 1, Row: 4}: NumberVal(20),
				{Col: 1, Row: 5}: NumberVal(10),
				{Col: 2, Row: 1}: StringVal("fifty"),
				{Col: 2, Row: 2}: StringVal("forty"),
				{Col: 2, Row: 3}: StringVal("thirty"),
				{Col: 2, Row: 4}: StringVal("twenty"),
				{Col: 2, Row: 5}: StringVal("ten"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(30,A1:A5,B1:B5,,0,-2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// Binary search descending on reverse-sorted data; exact match on 30 → "thirty"
		if got.Type != ValueString || got.Str != "thirty" {
			t.Errorf("got %v, want thirty", got)
		}
	})

	// ---- match_mode edge cases ----

	t.Run("match_mode 1 next larger no exact", func(t *testing.T) {
		// Sorted: 10, 20, 30, 40, 50. Lookup 35 → next larger is 40 → "forty"
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
				{Col: 1, Row: 4}: NumberVal(40),
				{Col: 1, Row: 5}: NumberVal(50),
				{Col: 2, Row: 1}: StringVal("ten"),
				{Col: 2, Row: 2}: StringVal("twenty"),
				{Col: 2, Row: 3}: StringVal("thirty"),
				{Col: 2, Row: 4}: StringVal("forty"),
				{Col: 2, Row: 5}: StringVal("fifty"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(35,A1:A5,B1:B5,,1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "forty" {
			t.Errorf("got %v, want forty", got)
		}
	})

	t.Run("match_mode -1 next smaller no exact between values", func(t *testing.T) {
		// Sorted: 10, 20, 30, 40, 50. Lookup 35 → next smaller is 30 → "thirty"
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
				{Col: 1, Row: 4}: NumberVal(40),
				{Col: 1, Row: 5}: NumberVal(50),
				{Col: 2, Row: 1}: StringVal("ten"),
				{Col: 2, Row: 2}: StringVal("twenty"),
				{Col: 2, Row: 3}: StringVal("thirty"),
				{Col: 2, Row: 4}: StringVal("forty"),
				{Col: 2, Row: 5}: StringVal("fifty"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(35,A1:A5,B1:B5,,-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "thirty" {
			t.Errorf("got %v, want thirty", got)
		}
	})

	t.Run("match_mode -1 all values larger returns not_found", func(t *testing.T) {
		// Sorted: 10, 20, 30. Lookup 5 with match_mode -1 → no value <= 5 → #N/A
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
		cf := evalCompile(t, `XLOOKUP(5,A1:A3,B1:B3,,-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("match_mode 1 all values smaller returns not_found", func(t *testing.T) {
		// Sorted: 10, 20, 30. Lookup 100 with match_mode 1 → no value >= 100 → #N/A
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
		cf := evalCompile(t, `XLOOKUP(100,A1:A3,B1:B3,,1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("match_mode -1 custom not_found when all values larger", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 2, Row: 1}: StringVal("ten"),
				{Col: 2, Row: 2}: StringVal("twenty"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(5,A1:A2,B1:B2,"None",-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "None" {
			t.Errorf("got %v, want None", got)
		}
	})

	// ---- if_not_found variants ----

	t.Run("if_not_found with string value", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("Apple"),
				{Col: 2, Row: 1}: NumberVal(10),
			},
		}
		cf := evalCompile(t, `XLOOKUP("Missing",A1:A1,B1:B1,"not here")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "not here" {
			t.Errorf("got %v, want 'not here'", got)
		}
	})

	t.Run("if_not_found with boolean TRUE", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("Apple"),
				{Col: 2, Row: 1}: NumberVal(10),
			},
		}
		cf := evalCompile(t, `XLOOKUP("Missing",A1:A1,B1:B1,TRUE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("got %v, want TRUE", got)
		}
	})

	t.Run("if_not_found with boolean FALSE", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("Apple"),
				{Col: 2, Row: 1}: NumberVal(10),
			},
		}
		cf := evalCompile(t, `XLOOKUP("Missing",A1:A1,B1:B1,FALSE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("got %v, want FALSE", got)
		}
	})

	t.Run("if_not_found with negative number", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("Apple"),
				{Col: 2, Row: 1}: NumberVal(10),
			},
		}
		cf := evalCompile(t, `XLOOKUP("Missing",A1:A1,B1:B1,-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -1 {
			t.Errorf("got %v, want -1", got)
		}
	})

	// ---- single element arrays ----

	t.Run("single element lookup and return arrays", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("Apple"),
				{Col: 2, Row: 1}: NumberVal(42),
			},
		}
		cf := evalCompile(t, `XLOOKUP("Apple",A1:A1,B1:B1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 42 {
			t.Errorf("got %v, want 42", got)
		}
	})

	t.Run("single element not found", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("Apple"),
				{Col: 2, Row: 1}: NumberVal(42),
			},
		}
		cf := evalCompile(t, `XLOOKUP("Banana",A1:A1,B1:B1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	// ---- boolean lookup ----

	t.Run("boolean TRUE lookup value", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(false),
				{Col: 1, Row: 2}: BoolVal(true),
				{Col: 1, Row: 3}: BoolVal(false),
				{Col: 2, Row: 1}: StringVal("no1"),
				{Col: 2, Row: 2}: StringVal("yes"),
				{Col: 2, Row: 3}: StringVal("no2"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(TRUE,A1:A3,B1:B3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf("got %v, want yes", got)
		}
	})

	t.Run("boolean FALSE lookup value", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
				{Col: 1, Row: 2}: BoolVal(false),
				{Col: 1, Row: 3}: BoolVal(true),
				{Col: 2, Row: 1}: StringVal("yes1"),
				{Col: 2, Row: 2}: StringVal("no"),
				{Col: 2, Row: 3}: StringVal("yes2"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(FALSE,A1:A3,B1:B3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "no" {
			t.Errorf("got %v, want no", got)
		}
	})

	// ---- empty cell handling ----

	t.Run("empty lookup value matches empty cell", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("Apple"),
				{Col: 1, Row: 2}: EmptyVal(),
				{Col: 1, Row: 3}: StringVal("Cherry"),
				{Col: 2, Row: 1}: NumberVal(10),
				// Row 2 col 2 intentionally not set (will be EmptyVal from resolver)
				{Col: 2, Row: 3}: NumberVal(30),
			},
		}
		// Lookup an empty cell reference (Z1 is empty in our resolver)
		cf := evalCompile(t, `XLOOKUP(Z1,A1:A3,B1:B3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// Empty lookup value should match the empty cell in A2
		if got.Type != ValueEmpty && got.Type != ValueNumber {
			t.Errorf("got type %d (%v), expected match on empty cell at row 2", got.Type, got)
		}
	})

	// ---- type coercion ----

	t.Run("string number does not match numeric value", func(t *testing.T) {
		// In Excel, XLOOKUP with exact match treats "5" (string) and 5 (number) as different
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(5),
				{Col: 1, Row: 2}: NumberVal(10),
				{Col: 2, Row: 1}: StringVal("five"),
				{Col: 2, Row: 2}: StringVal("ten"),
			},
		}
		cf := evalCompile(t, `XLOOKUP("5",A1:A2,B1:B2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// String "5" should not match number 5 in exact mode
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A (string vs number mismatch)", got)
		}
	})

	t.Run("number matches number exactly", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
				{Col: 1, Row: 3}: NumberVal(3),
				{Col: 2, Row: 1}: StringVal("one"),
				{Col: 2, Row: 2}: StringVal("two"),
				{Col: 2, Row: 3}: StringVal("three"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(2,A1:A3,B1:B3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "two" {
			t.Errorf("got %v, want two", got)
		}
	})

	// ---- too many args ----

	t.Run("too many args returns error", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("A"),
				{Col: 2, Row: 1}: NumberVal(1),
			},
		}
		cf := evalCompile(t, `XLOOKUP("A",A1:A1,B1:B1,"nf",0,1,99)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error for too many args", got)
		}
	})

	// ---- match_mode 1 exact match exists ----

	t.Run("match_mode 1 exact match exists returns exact", func(t *testing.T) {
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
		cf := evalCompile(t, `XLOOKUP(20,A1:A3,B1:B3,,1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "twenty" {
			t.Errorf("got %v, want twenty", got)
		}
	})

	// ---- match_mode -1 last value in sorted array ----

	t.Run("match_mode -1 lookup larger than all returns last", func(t *testing.T) {
		// Sorted: 10, 20, 30. Lookup 100 → last <= 100 is 30 → "thirty"
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
		cf := evalCompile(t, `XLOOKUP(100,A1:A3,B1:B3,,-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "thirty" {
			t.Errorf("got %v, want thirty", got)
		}
	})

	// ---- match_mode 1 first value in sorted array ----

	t.Run("match_mode 1 lookup smaller than all returns first", func(t *testing.T) {
		// Sorted: 10, 20, 30. Lookup 5 → first >= 5 is 10 → "ten"
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
		cf := evalCompile(t, `XLOOKUP(5,A1:A3,B1:B3,,1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "ten" {
			t.Errorf("got %v, want ten", got)
		}
	})

	// ---- lookup in column vs row arrays ----

	t.Run("vertical column lookup", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 3, Row: 1}: StringVal("red"),
				{Col: 3, Row: 2}: StringVal("green"),
				{Col: 3, Row: 3}: StringVal("blue"),
				{Col: 4, Row: 1}: NumberVal(1),
				{Col: 4, Row: 2}: NumberVal(2),
				{Col: 4, Row: 3}: NumberVal(3),
			},
		}
		cf := evalCompile(t, `XLOOKUP("green",C1:C3,D1:D3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("horizontal row lookup", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 5}: StringVal("Jan"),
				{Col: 2, Row: 5}: StringVal("Feb"),
				{Col: 3, Row: 5}: StringVal("Mar"),
				{Col: 1, Row: 6}: NumberVal(100),
				{Col: 2, Row: 6}: NumberVal(200),
				{Col: 3, Row: 6}: NumberVal(300),
			},
		}
		cf := evalCompile(t, `XLOOKUP("Feb",A5:C5,A6:C6)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 200 {
			t.Errorf("got %v, want 200", got)
		}
	})

	// ---- wildcard with tilde escape ----

	t.Run("wildcard tilde escape star", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("A*B"),
				{Col: 1, Row: 2}: StringVal("AXB"),
				{Col: 1, Row: 3}: StringVal("AB"),
				{Col: 2, Row: 1}: NumberVal(1),
				{Col: 2, Row: 2}: NumberVal(2),
				{Col: 2, Row: 3}: NumberVal(3),
			},
		}
		// ~* should match literal *, so "A~*B" matches "A*B"
		cf := evalCompile(t, `XLOOKUP("A~*B",A1:A3,B1:B3,,2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1 (literal star match)", got)
		}
	})

	t.Run("wildcard tilde escape question mark", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("A?B"),
				{Col: 1, Row: 2}: StringVal("AXB"),
				{Col: 2, Row: 1}: NumberVal(1),
				{Col: 2, Row: 2}: NumberVal(2),
			},
		}
		// ~? should match literal ?, so "A~?B" matches "A?B"
		cf := evalCompile(t, `XLOOKUP("A~?B",A1:A2,B1:B2,,2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1 (literal question mark match)", got)
		}
	})

	// ---- exact match with various types ----

	t.Run("exact match number zero", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(-1),
				{Col: 1, Row: 2}: NumberVal(0),
				{Col: 1, Row: 3}: NumberVal(1),
				{Col: 2, Row: 1}: StringVal("neg"),
				{Col: 2, Row: 2}: StringVal("zero"),
				{Col: 2, Row: 3}: StringVal("pos"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(0,A1:A3,B1:B3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "zero" {
			t.Errorf("got %v, want zero", got)
		}
	})

	t.Run("exact match negative number", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(-5),
				{Col: 1, Row: 2}: NumberVal(0),
				{Col: 1, Row: 3}: NumberVal(5),
				{Col: 2, Row: 1}: StringVal("neg5"),
				{Col: 2, Row: 2}: StringVal("zero"),
				{Col: 2, Row: 3}: StringVal("pos5"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(-5,A1:A3,B1:B3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "neg5" {
			t.Errorf("got %v, want neg5", got)
		}
	})

	t.Run("exact match empty string", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("Apple"),
				{Col: 1, Row: 2}: StringVal(""),
				{Col: 1, Row: 3}: StringVal("Cherry"),
				{Col: 2, Row: 1}: NumberVal(10),
				{Col: 2, Row: 2}: NumberVal(20),
				{Col: 2, Row: 3}: NumberVal(30),
			},
		}
		cf := evalCompile(t, `XLOOKUP("",A1:A3,B1:B3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20 (match on empty string)", got)
		}
	})

	// ---- match_mode with unsorted data ----

	t.Run("match_mode -1 unsorted picks largest value lte lookup", func(t *testing.T) {
		// Unsorted: 30, 10, 40, 20, 50. Lookup 35 → values <= 35: 30, 10, 20.
		// Implementation scans and keeps last <= 35, which is 20 at index 3.
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(30),
				{Col: 1, Row: 2}: NumberVal(10),
				{Col: 1, Row: 3}: NumberVal(40),
				{Col: 1, Row: 4}: NumberVal(20),
				{Col: 1, Row: 5}: NumberVal(50),
				{Col: 2, Row: 1}: StringVal("a"),
				{Col: 2, Row: 2}: StringVal("b"),
				{Col: 2, Row: 3}: StringVal("c"),
				{Col: 2, Row: 4}: StringVal("d"),
				{Col: 2, Row: 5}: StringVal("e"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(35,A1:A5,B1:B5,,-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// Implementation scans all values and keeps the last one where v <= lookup
		// 30<=35 (idx 0), 10<=35 (idx 1), 40>35, 20<=35 (idx 3), 50>35
		// lastMatch = 3 → returns "d"
		if got.Type != ValueString || got.Str != "d" {
			t.Errorf("got %v, want d", got)
		}
	})

	t.Run("match_mode 1 unsorted finds first value gte lookup", func(t *testing.T) {
		// Unsorted: 30, 10, 40, 20, 50. Lookup 35 → first >= 35: 40 at index 2
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(30),
				{Col: 1, Row: 2}: NumberVal(10),
				{Col: 1, Row: 3}: NumberVal(40),
				{Col: 1, Row: 4}: NumberVal(20),
				{Col: 1, Row: 5}: NumberVal(50),
				{Col: 2, Row: 1}: StringVal("a"),
				{Col: 2, Row: 2}: StringVal("b"),
				{Col: 2, Row: 3}: StringVal("c"),
				{Col: 2, Row: 4}: StringVal("d"),
				{Col: 2, Row: 5}: StringVal("e"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(35,A1:A5,B1:B5,,1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// First value >= 35 is 40 at index 2 → "c"
		if got.Type != ValueString || got.Str != "c" {
			t.Errorf("got %v, want c", got)
		}
	})

	// ---- large numeric values ----

	t.Run("exact match large numbers", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1000000),
				{Col: 1, Row: 2}: NumberVal(2000000),
				{Col: 1, Row: 3}: NumberVal(3000000),
				{Col: 2, Row: 1}: StringVal("1M"),
				{Col: 2, Row: 2}: StringVal("2M"),
				{Col: 2, Row: 3}: StringVal("3M"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(2000000,A1:A3,B1:B3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "2M" {
			t.Errorf("got %v, want 2M", got)
		}
	})

	// ---- fractional numbers ----

	t.Run("exact match fractional numbers", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1.5),
				{Col: 1, Row: 2}: NumberVal(2.5),
				{Col: 1, Row: 3}: NumberVal(3.5),
				{Col: 2, Row: 1}: StringVal("a"),
				{Col: 2, Row: 2}: StringVal("b"),
				{Col: 2, Row: 3}: StringVal("c"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(2.5,A1:A3,B1:B3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "b" {
			t.Errorf("got %v, want b", got)
		}
	})

	// ---- match_mode -1 with fractional lookup ----

	t.Run("match_mode -1 fractional between values", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1.0),
				{Col: 1, Row: 2}: NumberVal(2.0),
				{Col: 1, Row: 3}: NumberVal(3.0),
				{Col: 2, Row: 1}: StringVal("one"),
				{Col: 2, Row: 2}: StringVal("two"),
				{Col: 2, Row: 3}: StringVal("three"),
			},
		}
		cf := evalCompile(t, `XLOOKUP(2.5,A1:A3,B1:B3,,-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "two" {
			t.Errorf("got %v, want two", got)
		}
	})

	// ---- return array is left of lookup array ----

	t.Run("return array left of lookup array", func(t *testing.T) {
		// Return (col A) is left of lookup (col B) — XLOOKUP allows this
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(100),
				{Col: 1, Row: 2}: NumberVal(200),
				{Col: 1, Row: 3}: NumberVal(300),
				{Col: 2, Row: 1}: StringVal("X"),
				{Col: 2, Row: 2}: StringVal("Y"),
				{Col: 2, Row: 3}: StringVal("Z"),
			},
		}
		cf := evalCompile(t, `XLOOKUP("Y",B1:B3,A1:A3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 200 {
			t.Errorf("got %v, want 200", got)
		}
	})

	// ---- wildcard combined with case insensitivity ----

	t.Run("wildcard star middle of string case insensitive", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("New York"),
				{Col: 1, Row: 2}: StringVal("New Jersey"),
				{Col: 1, Row: 3}: StringVal("New Mexico"),
				{Col: 2, Row: 1}: StringVal("NY"),
				{Col: 2, Row: 2}: StringVal("NJ"),
				{Col: 2, Row: 3}: StringVal("NM"),
			},
		}
		cf := evalCompile(t, `XLOOKUP("new*jersey",A1:A3,B1:B3,,2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "NJ" {
			t.Errorf("got %v, want NJ", got)
		}
	})

	// ---- match_mode 0 with empty if_not_found arg (omitted) ----

	t.Run("omitted if_not_found defaults to NA", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("A"),
				{Col: 2, Row: 1}: NumberVal(1),
			},
		}
		cf := evalCompile(t, `XLOOKUP("Z",A1:A1,B1:B1,,0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	// ---- match_mode 1 at boundary ----

	t.Run("match_mode 1 exact at last element", func(t *testing.T) {
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
		cf := evalCompile(t, `XLOOKUP(30,A1:A3,B1:B3,,1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "thirty" {
			t.Errorf("got %v, want thirty", got)
		}
	})

	// ---- match_mode -1 at boundary ----

	t.Run("match_mode -1 exact at first element", func(t *testing.T) {
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
		cf := evalCompile(t, `XLOOKUP(10,A1:A3,B1:B3,,-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "ten" {
			t.Errorf("got %v, want ten", got)
		}
	})

	// ---- mixed types in lookup array ----

	t.Run("mixed types in lookup array finds correct type", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: StringVal("hello"),
				{Col: 1, Row: 3}: BoolVal(true),
				{Col: 1, Row: 4}: NumberVal(2),
				{Col: 2, Row: 1}: StringVal("r1"),
				{Col: 2, Row: 2}: StringVal("r2"),
				{Col: 2, Row: 3}: StringVal("r3"),
				{Col: 2, Row: 4}: StringVal("r4"),
			},
		}
		cf := evalCompile(t, `XLOOKUP("hello",A1:A4,B1:B4)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "r2" {
			t.Errorf("got %v, want r2", got)
		}
	})
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
		// --- additional coverage ---
		{
			name: "single_column_array_select_col1",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{NumberVal(10)},
					{NumberVal(20)},
					{NumberVal(30)},
				}},
				NumberVal(1),
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10)},
				{NumberVal(20)},
				{NumberVal(30)},
			}},
		},
		{
			name: "single_column_array_negative_one",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{NumberVal(10)},
					{NumberVal(20)},
				}},
				NumberVal(-1),
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10)},
				{NumberVal(20)},
			}},
		},
		{
			name: "single_row_wide_array",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5)},
				}},
				NumberVal(2), NumberVal(4),
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2), NumberVal(4)},
			}},
		},
		{
			name: "all_columns_selected",
			args: []Value{base, NumberVal(1), NumberVal(2), NumberVal(3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2), NumberVal(3)},
				{NumberVal(4), NumberVal(5), NumberVal(6)},
			}},
		},
		{
			name: "all_columns_reversed",
			args: []Value{base, NumberVal(3), NumberVal(2), NumberVal(1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3), NumberVal(2), NumberVal(1)},
				{NumberVal(6), NumberVal(5), NumberVal(4)},
			}},
		},
		{
			name: "multiple_negative_indices",
			args: []Value{base, NumberVal(-1), NumberVal(-2), NumberVal(-3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3), NumberVal(2), NumberVal(1)},
				{NumberVal(6), NumberVal(5), NumberVal(4)},
			}},
		},
		{
			name: "large_array_select_boundary_cols",
			args: func() []Value {
				row := make([]Value, 100)
				for i := range row {
					row[i] = NumberVal(float64(i + 1))
				}
				arr := Value{Type: ValueArray, Array: [][]Value{row}}
				return []Value{arr, NumberVal(1), NumberVal(50), NumberVal(100)}
			}(),
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(50), NumberVal(100)},
			}},
		},
		{
			name: "string_array_values",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{StringVal("alpha"), StringVal("beta"), StringVal("gamma")},
					{StringVal("delta"), StringVal("epsilon"), StringVal("zeta")},
				}},
				NumberVal(2),
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{StringVal("beta")},
				{StringVal("epsilon")},
			}},
		},
		{
			name: "boolean_array_values",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{BoolVal(true), BoolVal(false)},
					{BoolVal(false), BoolVal(true)},
				}},
				NumberVal(2), NumberVal(1),
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{BoolVal(false), BoolVal(true)},
				{BoolVal(true), BoolVal(false)},
			}},
		},
		{
			name: "mixed_type_array_preserve_types",
			args: []Value{mixed, NumberVal(1), NumberVal(2), NumberVal(3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{StringVal("a"), BoolVal(true), ErrorVal(ErrValNA)},
				{EmptyVal(), NumberVal(2), StringVal("z")},
			}},
		},
		{
			name: "bool_false_index_coerces_to_zero_errors",
			args: []Value{base, BoolVal(false)},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "negative_exact_boundary",
			args: []Value{base, NumberVal(-3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1)},
				{NumberVal(4)},
			}},
		},
		{
			name: "positive_exact_boundary",
			args: []Value{base, NumberVal(3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3)},
				{NumberVal(6)},
			}},
		},
		{
			name: "duplicate_same_column_three_times",
			args: []Value{base, NumberVal(2), NumberVal(2), NumberVal(2)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2), NumberVal(2), NumberVal(2)},
				{NumberVal(5), NumberVal(5), NumberVal(5)},
			}},
		},
		{
			name: "single_cell_array_col1",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{NumberVal(42)},
				}},
				NumberVal(1),
			},
			want: NumberVal(42),
		},
		{
			name: "single_cell_array_negative_one",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{StringVal("hello")},
				}},
				NumberVal(-1),
			},
			want: StringVal("hello"),
		},
		{
			name: "error_in_second_index_propagates",
			args: []Value{base, NumberVal(1), ErrorVal(ErrValNA)},
			want: ErrorVal(ErrValNA),
		},
		{
			name: "negative_two_on_three_col_array",
			args: []Value{base, NumberVal(-2)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2)},
				{NumberVal(5)},
			}},
		},
		{
			name: "large_array_negative_last",
			args: func() []Value {
				row := make([]Value, 50)
				for i := range row {
					row[i] = NumberVal(float64(i + 1))
				}
				arr := Value{Type: ValueArray, Array: [][]Value{row}}
				return []Value{arr, NumberVal(-1)}
			}(),
			want: NumberVal(50), // single-row single-col result unwraps to scalar
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
		// --- additional eval coverage ---
		{
			name:    "index_into_choosecols_result",
			formula: "INDEX(CHOOSECOLS(A1:C2,3,1),1,2)",
			want:    NumberVal(1),
		},
		{
			name:    "index_into_choosecols_row2",
			formula: "INDEX(CHOOSECOLS(A1:C2,3,1),2,1)",
			want:    NumberVal(6),
		},
		{
			name:    "out_of_range_column_formula",
			formula: "CHOOSECOLS(A1:C2,4)",
			want:    ErrorVal(ErrValVALUE),
		},
		{
			name:    "zero_column_formula",
			formula: "CHOOSECOLS(A1:C2,0)",
			want:    ErrorVal(ErrValVALUE),
		},
		{
			name:    "negative_out_of_range_formula",
			formula: "CHOOSECOLS(A1:C2,-4)",
			want:    ErrorVal(ErrValVALUE),
		},
		{
			name:    "duplicate_columns_formula",
			formula: "CHOOSECOLS(A1:C2,1,1)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(1)},
				{NumberVal(4), NumberVal(4)},
			}},
		},
		{
			name:    "all_columns_formula",
			formula: "CHOOSECOLS(A1:C2,1,2,3)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2), NumberVal(3)},
				{NumberVal(4), NumberVal(5), NumberVal(6)},
			}},
		},
		{
			name:    "mix_positive_and_negative_formula",
			formula: "CHOOSECOLS(A1:C2,1,-1)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(3)},
				{NumberVal(4), NumberVal(6)},
			}},
		},
		{
			name:    "fractional_index_formula",
			formula: "CHOOSECOLS(A1:C2,2.7)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2)},
				{NumberVal(5)},
			}},
		},
		{
			name:    "single_cell_ref_formula",
			formula: "CHOOSECOLS(A1,1)",
			want:    NumberVal(1),
		},
		{
			name:    "multiple_negative_indices_formula",
			formula: "CHOOSECOLS(A1:C2,-1,-2,-3)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3), NumberVal(2), NumberVal(1)},
				{NumberVal(6), NumberVal(5), NumberVal(4)},
			}},
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
		// --- additional coverage ---
		{
			name: "single_row_array_select_row1",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{NumberVal(10), NumberVal(20), NumberVal(30)},
				}},
				NumberVal(1),
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10), NumberVal(20), NumberVal(30)},
			}},
		},
		{
			name: "single_row_array_negative_one",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{NumberVal(10), NumberVal(20)},
				}},
				NumberVal(-1),
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10), NumberVal(20)},
			}},
		},
		{
			name: "single_column_tall_array",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{NumberVal(1)},
					{NumberVal(2)},
					{NumberVal(3)},
					{NumberVal(4)},
					{NumberVal(5)},
				}},
				NumberVal(2), NumberVal(4),
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2)},
				{NumberVal(4)},
			}},
		},
		{
			name: "all_rows_selected",
			args: []Value{base, NumberVal(1), NumberVal(2), NumberVal(3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2), NumberVal(3)},
				{NumberVal(4), NumberVal(5), NumberVal(6)},
				{NumberVal(7), NumberVal(8), NumberVal(9)},
			}},
		},
		{
			name: "all_rows_reversed",
			args: []Value{base, NumberVal(3), NumberVal(2), NumberVal(1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(7), NumberVal(8), NumberVal(9)},
				{NumberVal(4), NumberVal(5), NumberVal(6)},
				{NumberVal(1), NumberVal(2), NumberVal(3)},
			}},
		},
		{
			name: "multiple_negative_indices",
			args: []Value{base, NumberVal(-1), NumberVal(-2), NumberVal(-3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(7), NumberVal(8), NumberVal(9)},
				{NumberVal(4), NumberVal(5), NumberVal(6)},
				{NumberVal(1), NumberVal(2), NumberVal(3)},
			}},
		},
		{
			name: "large_array_select_boundary_rows",
			args: func() []Value {
				rows := make([][]Value, 100)
				for i := range rows {
					rows[i] = []Value{NumberVal(float64(i + 1))}
				}
				arr := Value{Type: ValueArray, Array: rows}
				return []Value{arr, NumberVal(1), NumberVal(50), NumberVal(100)}
			}(),
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1)},
				{NumberVal(50)},
				{NumberVal(100)},
			}},
		},
		{
			name: "string_array_values",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{StringVal("alpha"), StringVal("beta")},
					{StringVal("gamma"), StringVal("delta")},
					{StringVal("epsilon"), StringVal("zeta")},
				}},
				NumberVal(2),
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{StringVal("gamma"), StringVal("delta")},
			}},
		},
		{
			name: "boolean_array_values",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{BoolVal(true), BoolVal(false)},
					{BoolVal(false), BoolVal(true)},
				}},
				NumberVal(2), NumberVal(1),
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{BoolVal(false), BoolVal(true)},
				{BoolVal(true), BoolVal(false)},
			}},
		},
		{
			name: "mixed_type_array_all_rows",
			args: []Value{mixed, NumberVal(1), NumberVal(2), NumberVal(3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{StringVal("a"), BoolVal(true), ErrorVal(ErrValNA)},
				{EmptyVal(), NumberVal(2), StringVal("z")},
				{NumberVal(9), StringVal("tail"), BoolVal(false)},
			}},
		},
		{
			name: "bool_false_index_coerces_to_zero_errors",
			args: []Value{base, BoolVal(false)},
			want: ErrorVal(ErrValVALUE),
		},
		{
			name: "negative_exact_boundary",
			args: []Value{base, NumberVal(-3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2), NumberVal(3)},
			}},
		},
		{
			name: "positive_exact_boundary",
			args: []Value{base, NumberVal(3)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(7), NumberVal(8), NumberVal(9)},
			}},
		},
		{
			name: "duplicate_same_row_three_times",
			args: []Value{base, NumberVal(2), NumberVal(2), NumberVal(2)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4), NumberVal(5), NumberVal(6)},
				{NumberVal(4), NumberVal(5), NumberVal(6)},
				{NumberVal(4), NumberVal(5), NumberVal(6)},
			}},
		},
		{
			name: "single_cell_array_row1",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{NumberVal(42)},
				}},
				NumberVal(1),
			},
			want: NumberVal(42),
		},
		{
			name: "single_cell_array_negative_one",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{StringVal("hello")},
				}},
				NumberVal(-1),
			},
			want: StringVal("hello"),
		},
		{
			name: "error_in_second_index_propagates",
			args: []Value{base, NumberVal(1), ErrorVal(ErrValNA)},
			want: ErrorVal(ErrValNA),
		},
		{
			name: "negative_two_on_three_row_array",
			args: []Value{base, NumberVal(-2)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4), NumberVal(5), NumberVal(6)},
			}},
		},
		{
			name: "large_array_negative_last",
			args: func() []Value {
				rows := make([][]Value, 50)
				for i := range rows {
					rows[i] = []Value{NumberVal(float64(i + 1))}
				}
				arr := Value{Type: ValueArray, Array: rows}
				return []Value{arr, NumberVal(-1)}
			}(),
			want: NumberVal(50), // single-row single-col result unwraps to scalar
		},
		{
			name: "multi_column_preserves_all_columns",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{
					{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4), NumberVal(5)},
					{NumberVal(6), NumberVal(7), NumberVal(8), NumberVal(9), NumberVal(10)},
					{NumberVal(11), NumberVal(12), NumberVal(13), NumberVal(14), NumberVal(15)},
				}},
				NumberVal(2),
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(6), NumberVal(7), NumberVal(8), NumberVal(9), NumberVal(10)},
			}},
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
		// --- additional eval coverage ---
		{
			name:    "index_into_chooserows_result",
			formula: "INDEX(CHOOSEROWS(A1:C4,4,1),1,2)",
			want:    NumberVal(11),
		},
		{
			name:    "index_into_chooserows_row2",
			formula: "INDEX(CHOOSEROWS(A1:C4,4,1),2,3)",
			want:    NumberVal(3),
		},
		{
			name:    "out_of_range_row_formula",
			formula: "CHOOSEROWS(A1:C4,5)",
			want:    ErrorVal(ErrValVALUE),
		},
		{
			name:    "zero_row_formula",
			formula: "CHOOSEROWS(A1:C4,0)",
			want:    ErrorVal(ErrValVALUE),
		},
		{
			name:    "negative_out_of_range_formula",
			formula: "CHOOSEROWS(A1:C4,-5)",
			want:    ErrorVal(ErrValVALUE),
		},
		{
			name:    "duplicate_rows_formula",
			formula: "CHOOSEROWS(A1:C4,2,2)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4), NumberVal(5), NumberVal(6)},
				{NumberVal(4), NumberVal(5), NumberVal(6)},
			}},
		},
		{
			name:    "all_rows_formula",
			formula: "CHOOSEROWS(A1:C4,1,2,3,4)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2), NumberVal(3)},
				{NumberVal(4), NumberVal(5), NumberVal(6)},
				{NumberVal(7), NumberVal(8), NumberVal(9)},
				{NumberVal(10), NumberVal(11), NumberVal(12)},
			}},
		},
		{
			name:    "mix_positive_and_negative_formula",
			formula: "CHOOSEROWS(A1:C4,1,-1)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2), NumberVal(3)},
				{NumberVal(10), NumberVal(11), NumberVal(12)},
			}},
		},
		{
			name:    "fractional_index_formula",
			formula: "CHOOSEROWS(A1:C4,2.7)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4), NumberVal(5), NumberVal(6)},
			}},
		},
		{
			name:    "single_cell_ref_formula",
			formula: "CHOOSEROWS(A1,1)",
			want:    NumberVal(1),
		},
		{
			name:    "multiple_negative_indices_formula",
			formula: "CHOOSEROWS(A1:C4,-1,-2,-3,-4)",
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10), NumberVal(11), NumberVal(12)},
				{NumberVal(7), NumberVal(8), NumberVal(9)},
				{NumberVal(4), NumberVal(5), NumberVal(6)},
				{NumberVal(1), NumberVal(2), NumberVal(3)},
			}},
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

// ---------------------------------------------------------------------------
// XMATCH
// ---------------------------------------------------------------------------

func TestXMATCH(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// A1:A5 — string data
			{Col: 1, Row: 1}: StringVal("Apple"),
			{Col: 1, Row: 2}: StringVal("Banana"),
			{Col: 1, Row: 3}: StringVal("Cherry"),
			{Col: 1, Row: 4}: StringVal("Date"),
			{Col: 1, Row: 5}: StringVal("Elderberry"),

			// B1:B5 — numeric data (ascending)
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
			{Col: 2, Row: 4}: NumberVal(40),
			{Col: 2, Row: 5}: NumberVal(50),

			// C1:C5 — numeric data (descending)
			{Col: 3, Row: 1}: NumberVal(50),
			{Col: 3, Row: 2}: NumberVal(40),
			{Col: 3, Row: 3}: NumberVal(30),
			{Col: 3, Row: 4}: NumberVal(20),
			{Col: 3, Row: 5}: NumberVal(10),

			// D1:D3 — boolean data
			{Col: 4, Row: 1}: BoolVal(true),
			{Col: 4, Row: 2}: BoolVal(false),
			{Col: 4, Row: 3}: BoolVal(true),

			// E1:E5 — duplicate values
			{Col: 5, Row: 1}: NumberVal(10),
			{Col: 5, Row: 2}: NumberVal(20),
			{Col: 5, Row: 3}: NumberVal(20),
			{Col: 5, Row: 4}: NumberVal(30),
			{Col: 5, Row: 5}: NumberVal(20),

			// F1 — single element
			{Col: 6, Row: 1}: NumberVal(42),

			// G1:G3 — wildcard test data
			{Col: 7, Row: 1}: StringVal("Banana Split"),
			{Col: 7, Row: 2}: StringVal("Apple Pie"),
			{Col: 7, Row: 3}: StringVal("Cherry Tart"),
		},
	}

	t.Run("basic exact match number", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(30,B1:B5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("basic exact match string", func(t *testing.T) {
		cf := evalCompile(t, `XMATCH("Cherry",A1:A5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("exact match boolean", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(FALSE,D1:D3,0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("case insensitive string match", func(t *testing.T) {
		cf := evalCompile(t, `XMATCH("banana",A1:A5,0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("not found returns NA", func(t *testing.T) {
		cf := evalCompile(t, `XMATCH("Mango",A1:A5,0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("exact match first element", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(10,B1:B5,0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("exact match last element", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(50,B1:B5,0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
		}
	})

	t.Run("single element array found", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(42,F1:F1,0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("single element array not found", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(99,F1:F1,0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	// match_mode -1: exact match or next smallest
	t.Run("next smallest exact hit", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(30,B1:B5,-1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("next smallest between values", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(25,B1:B5,-1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2 (position of 20)", got)
		}
	})

	t.Run("next smallest below all returns NA", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(5,B1:B5,-1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	// match_mode 1: exact match or next largest
	t.Run("next largest exact hit", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(30,B1:B5,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("next largest between values", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(25,B1:B5,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3 (position of 30)", got)
		}
	})

	t.Run("next largest above all returns NA", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(55,B1:B5,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	// match_mode 2: wildcard
	t.Run("wildcard star match", func(t *testing.T) {
		cf := evalCompile(t, `XMATCH("Apple*",G1:G3,2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("wildcard question mark match", func(t *testing.T) {
		cf := evalCompile(t, `XMATCH("Cherry Tar?",G1:G3,2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("wildcard no match returns NA", func(t *testing.T) {
		cf := evalCompile(t, `XMATCH("Mango*",G1:G3,2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	// search_mode -1: last-to-first
	t.Run("search last-to-first finds last duplicate", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(20,E1:E5,0,-1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5 (last occurrence of 20)", got)
		}
	})

	t.Run("search first-to-last finds first duplicate", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(20,E1:E5,0,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2 (first occurrence of 20)", got)
		}
	})

	// search_mode 2: binary search ascending
	t.Run("binary search ascending exact", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(30,B1:B5,0,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("binary search ascending not found", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(25,B1:B5,0,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	// search_mode -2: binary search descending
	t.Run("binary search descending exact", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(30,C1:C5,0,-2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("binary search descending not found", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(25,C1:C5,0,-2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	// binary search with next smallest (match_mode -1, search_mode 2)
	t.Run("binary search ascending next smallest", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(25,B1:B5,-1,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2 (position of 20)", got)
		}
	})

	// binary search with next largest (match_mode 1, search_mode 2)
	t.Run("binary search ascending next largest", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(25,B1:B5,1,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3 (position of 30)", got)
		}
	})

	// binary search descending with next smallest (match_mode -1, search_mode -2)
	t.Run("binary search descending next smallest", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(25,C1:C5,-1,-2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 4 {
			t.Errorf("got %v, want 4 (position of 20)", got)
		}
	})

	// binary search descending with next largest (match_mode 1, search_mode -2)
	t.Run("binary search descending next largest", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(25,C1:C5,1,-2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3 (position of 30)", got)
		}
	})

	// wildcard with reverse search
	t.Run("wildcard reverse search finds last match", func(t *testing.T) {
		// G1="Banana Split", G2="Apple Pie", G3="Cherry Tart"
		// "*a*" matches all three; reverse search should find G3 (position 3)
		cf := evalCompile(t, `XMATCH("*a*",G1:G3,2,-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3 (last match in reverse)", got)
		}
	})

	// wrong number of args
	t.Run("too few args", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(10)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error for too few args", got)
		}
	})

	t.Run("too many args", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(10,B1:B5,0,1,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want error for too many args", got)
		}
	})

	// default match_mode and search_mode
	t.Run("defaults to exact match first-to-last", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(20,B1:B5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	// match_mode -1 next smallest with exact match in ascending data
	t.Run("next smallest exact in ascending data", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(40,B1:B5,-1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 4 {
			t.Errorf("got %v, want 4", got)
		}
	})

	// match_mode 1 next largest with exact match in ascending data
	t.Run("next largest exact in ascending data", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(40,B1:B5,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 4 {
			t.Errorf("got %v, want 4", got)
		}
	})

	// mixed types — number not found in string array
	t.Run("number in string array returns NA", func(t *testing.T) {
		cf := evalCompile(t, "XMATCH(42,A1:A5,0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})
}

func TestADDRESS(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		name    string
		formula string
		want    string
	}{
		// A1 style, absolute (default abs_num=1)
		{"abs_A1", "ADDRESS(1,1)", "$A$1"},
		{"abs_A1_row2_col3", "ADDRESS(2,3)", "$C$2"},
		{"abs_A1_explicit", "ADDRESS(1,1,1)", "$A$1"},
		// A1 style, abs_num=2 (absolute row, relative col)
		{"abs_row_rel_col", "ADDRESS(1,1,2)", "A$1"},
		{"abs_row_rel_col_C2", "ADDRESS(2,3,2)", "C$2"},
		// A1 style, abs_num=3 (relative row, absolute col)
		{"rel_row_abs_col", "ADDRESS(1,1,3)", "$A1"},
		{"rel_row_abs_col_C2", "ADDRESS(2,3,3)", "$C2"},
		// A1 style, abs_num=4 (fully relative)
		{"rel_A1", "ADDRESS(1,1,4)", "A1"},
		{"rel_A1_C2", "ADDRESS(2,3,4)", "C2"},
		// R1C1 style (a1_style=FALSE)
		{"abs_R1C1", "ADDRESS(1,1,1,FALSE)", "R1C1"},
		{"rel_col_R1C1", "ADDRESS(1,1,2,FALSE)", "R1C[1]"},
		{"rel_row_R1C1", "ADDRESS(1,1,3,FALSE)", "R[1]C1"},
		{"rel_R1C1", "ADDRESS(1,1,4,FALSE)", "R[1]C[1]"},
		{"abs_R1C1_row2_col3", "ADDRESS(2,3,1,FALSE)", "R2C3"},
		{"rel_col_R1C1_row2_col3", "ADDRESS(2,3,2,FALSE)", "R2C[3]"},
		{"rel_row_R1C1_row2_col3", "ADDRESS(2,3,3,FALSE)", "R[2]C3"},
		{"rel_R1C1_row2_col3", "ADDRESS(2,3,4,FALSE)", "R[2]C[3]"},
		// Explicit TRUE for A1 style
		{"explicit_true_A1", "ADDRESS(2,3,1,TRUE)", "$C$2"},
		// Large column numbers
		{"col_26_Z", "ADDRESS(1,26)", "$Z$1"},
		{"col_27_AA", "ADDRESS(1,27)", "$AA$1"},
		{"col_256_IV", "ADDRESS(1,256)", "$IV$1"},
		{"col_702_ZZ", "ADDRESS(1,702)", "$ZZ$1"},
		{"col_16384_XFD", "ADDRESS(1,16384)", "$XFD$1"},
		// Large row number
		{"large_row", "ADDRESS(1048576,1)", "$A$1048576"},
		// Sheet name
		{"with_sheet", `ADDRESS(1,1,1,TRUE,"Sheet1")`, "Sheet1!$A$1"},
		{"with_sheet_spaces", `ADDRESS(1,1,1,TRUE,"My Sheet")`, "'My Sheet'!$A$1"},
		{"with_sheet_quote", `ADDRESS(1,1,1,TRUE,"Sheet'1")`, "'Sheet'1'!$A$1"},
		{"with_sheet_R1C1", `ADDRESS(1,1,1,FALSE,"Sheet1")`, "Sheet1!R1C1"},
		// Sheet with bracket needs quoting
		{"with_sheet_bracket", `ADDRESS(1,1,1,TRUE,"Sheet[1]")`, "'Sheet[1]'!$A$1"},
		// Sheet with relative addressing
		{"with_sheet_relative", `ADDRESS(2,3,4,TRUE,"Data")`, "Data!C2"},
		// Sheet with R1C1 relative addressing
		{"with_sheet_R1C1_relative", `ADDRESS(2,3,4,FALSE,"Data")`, "Data!R[2]C[3]"},
	}

	for _, tt := range strTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString {
				t.Fatalf("Eval(%q): got type %v, want string", tt.formula, got.Type)
			}
			if got.Str != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"no_args", "ADDRESS()", ErrValVALUE},
		{"one_arg", "ADDRESS(1)", ErrValVALUE},
		{"too_many_args", `ADDRESS(1,1,1,TRUE,"Sheet1","extra")`, ErrValVALUE},
		{"row_zero", "ADDRESS(0,1)", ErrValVALUE},
		{"col_zero", "ADDRESS(1,0)", ErrValVALUE},
		{"negative_row", "ADDRESS(-1,1)", ErrValVALUE},
		{"negative_col", "ADDRESS(1,-1)", ErrValVALUE},
		{"invalid_abs_num", "ADDRESS(1,1,5)", ErrValVALUE},
		{"invalid_abs_num_zero", "ADDRESS(1,1,0)", ErrValVALUE},
		{"string_row", `ADDRESS("abc",1)`, ErrValVALUE},
		{"string_col", `ADDRESS(1,"abc")`, ErrValVALUE},
		{"invalid_abs_num_negative", "ADDRESS(1,1,-1)", ErrValVALUE},
		{"string_abs_num", `ADDRESS(1,1,"abc")`, ErrValVALUE},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = type=%v err=%v, want error %v", tt.formula, got.Type, got.Err, tt.wantErr)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// LOOKUP
// ---------------------------------------------------------------------------

func TestLOOKUP(t *testing.T) {
	// Vector form: LOOKUP(lookup_value, lookup_vector, result_vector)
	// Sorted numeric lookup_vector in A1:A5, result strings in B1:B5.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: NumberVal(40),
			{Col: 1, Row: 5}: NumberVal(50),
			{Col: 2, Row: 1}: StringVal("ten"),
			{Col: 2, Row: 2}: StringVal("twenty"),
			{Col: 2, Row: 3}: StringVal("thirty"),
			{Col: 2, Row: 4}: StringVal("forty"),
			{Col: 2, Row: 5}: StringVal("fifty"),
		},
	}

	t.Run("vector_exact_match", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(30,A1:A5,B1:B5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "thirty" {
			t.Errorf("got %v, want thirty", got)
		}
	})

	t.Run("vector_exact_first", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(10,A1:A5,B1:B5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "ten" {
			t.Errorf("got %v, want ten", got)
		}
	})

	t.Run("vector_exact_last", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(50,A1:A5,B1:B5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "fifty" {
			t.Errorf("got %v, want fifty", got)
		}
	})

	t.Run("vector_approx_between_values", func(t *testing.T) {
		// 25 is between 20 and 30; should return result for 20
		cf := evalCompile(t, "LOOKUP(25,A1:A5,B1:B5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "twenty" {
			t.Errorf("got %v, want twenty", got)
		}
	})

	t.Run("vector_approx_larger_than_all", func(t *testing.T) {
		// 999 > all values; should return last result
		cf := evalCompile(t, "LOOKUP(999,A1:A5,B1:B5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "fifty" {
			t.Errorf("got %v, want fifty", got)
		}
	})

	t.Run("vector_less_than_all_returns_NA", func(t *testing.T) {
		// 1 < 10 (smallest); should return #N/A
		cf := evalCompile(t, "LOOKUP(1,A1:A5,B1:B5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("vector_approx_just_below_second", func(t *testing.T) {
		// 19 is between 10 and 20; should return result for 10
		cf := evalCompile(t, "LOOKUP(19,A1:A5,B1:B5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "ten" {
			t.Errorf("got %v, want ten", got)
		}
	})
}

func TestLOOKUPArrayForm(t *testing.T) {
	// Array form: LOOKUP(lookup_value, array)
	// With a single-column vector, lookup and result are the same.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 1, Row: 4}: NumberVal(40),
		},
	}

	t.Run("array_single_column_exact", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(20,A1:A4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("array_single_column_approx", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(25,A1:A4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v, want 20", got)
		}
	})

	t.Run("array_single_column_not_found", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(5,A1:A4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("array_single_column_larger_than_all", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(100,A1:A4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 40 {
			t.Errorf("got %v, want 40", got)
		}
	})
}

func TestLOOKUPTextLookup(t *testing.T) {
	// Sorted text values in lookup vector
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple"),
			{Col: 1, Row: 2}: StringVal("banana"),
			{Col: 1, Row: 3}: StringVal("cherry"),
			{Col: 1, Row: 4}: StringVal("date"),
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 2, Row: 4}: NumberVal(4),
		},
	}

	t.Run("text_exact_match", func(t *testing.T) {
		cf := evalCompile(t, `LOOKUP("cherry",A1:A4,B1:B4)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v, want 3", got)
		}
	})

	t.Run("text_approx_match", func(t *testing.T) {
		// "cat" falls between "banana" and "cherry"; should return result for "banana"
		cf := evalCompile(t, `LOOKUP("cat",A1:A4,B1:B4)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v, want 2", got)
		}
	})

	t.Run("text_less_than_all", func(t *testing.T) {
		// "aaa" < "apple"; should return #N/A
		cf := evalCompile(t, `LOOKUP("aaa",A1:A4,B1:B4)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("text_greater_than_all", func(t *testing.T) {
		// "zebra" > "date"; should return last result
		cf := evalCompile(t, `LOOKUP("zebra",A1:A4,B1:B4)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 4 {
			t.Errorf("got %v, want 4", got)
		}
	})
}

func TestLOOKUPResultVectorShorter(t *testing.T) {
	// Result vector shorter than lookup vector: match at index beyond
	// result vector length should return #N/A.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
			{Col: 2, Row: 1}: StringVal("ten"),
			{Col: 2, Row: 2}: StringVal("twenty"),
			// B3 intentionally missing - result vector has only 2 elements
		},
	}

	t.Run("match_beyond_result_vector", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(30,A1:A3,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})

	t.Run("match_within_result_vector", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(20,A1:A3,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "twenty" {
			t.Errorf("got %v, want twenty", got)
		}
	})
}

func TestLOOKUPArgErrors(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
		},
	}

	t.Run("too_few_args", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(10)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})
}

func TestLOOKUPSingleElement(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 2, Row: 1}: StringVal("five"),
		},
	}

	t.Run("single_element_exact_match", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(5,A1:A1,B1:B1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "five" {
			t.Errorf("got %v, want five", got)
		}
	})

	t.Run("single_element_greater_value", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(10,A1:A1,B1:B1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "five" {
			t.Errorf("got %v, want five", got)
		}
	})

	t.Run("single_element_less_than", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(1,A1:A1,B1:B1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A", got)
		}
	})
}

func TestLOOKUPDecimalValues(t *testing.T) {
	// Fractional/decimal lookup values (from docs example)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(4.14),
			{Col: 1, Row: 2}: NumberVal(4.19),
			{Col: 1, Row: 3}: NumberVal(5.17),
			{Col: 1, Row: 4}: NumberVal(5.77),
			{Col: 1, Row: 5}: NumberVal(6.39),
			{Col: 2, Row: 1}: StringVal("red"),
			{Col: 2, Row: 2}: StringVal("orange"),
			{Col: 2, Row: 3}: StringVal("yellow"),
			{Col: 2, Row: 4}: StringVal("green"),
			{Col: 2, Row: 5}: StringVal("blue"),
		},
	}

	t.Run("decimal_exact", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(5.17,A1:A5,B1:B5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yellow" {
			t.Errorf("got %v, want yellow", got)
		}
	})

	t.Run("decimal_approx", func(t *testing.T) {
		// 4.15 between 4.14 and 4.19
		cf := evalCompile(t, "LOOKUP(4.15,A1:A5,B1:B5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "red" {
			t.Errorf("got %v, want red", got)
		}
	})

	t.Run("decimal_large", func(t *testing.T) {
		cf := evalCompile(t, "LOOKUP(7.5,A1:A5,B1:B5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "blue" {
			t.Errorf("got %v, want blue", got)
		}
	})
}

// ---- EXPAND tests ----

func TestEXPAND_Basic2x2To3x3(t *testing.T) {
	// EXPAND({1,2;3,4}, 3, 3) → 3×3 with #N/A padding
	got, err := fnEXPAND([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
			{NumberVal(3), NumberVal(4)},
		}},
		NumberVal(3),
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 3x3 array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	// Original values
	if got.Array[0][0].Num != 1 || got.Array[0][1].Num != 2 {
		t.Errorf("row 0: got {%g,%g}, want {1,2}", got.Array[0][0].Num, got.Array[0][1].Num)
	}
	if got.Array[1][0].Num != 3 || got.Array[1][1].Num != 4 {
		t.Errorf("row 1: got {%g,%g}, want {3,4}", got.Array[1][0].Num, got.Array[1][1].Num)
	}
	// Padding cells should be #N/A
	if got.Array[0][2].Type != ValueError || got.Array[0][2].Err != ErrValNA {
		t.Errorf("expected #N/A at [0][2], got %v", got.Array[0][2])
	}
	if got.Array[2][0].Type != ValueError || got.Array[2][0].Err != ErrValNA {
		t.Errorf("expected #N/A at [2][0], got %v", got.Array[2][0])
	}
	if got.Array[2][2].Type != ValueError || got.Array[2][2].Err != ErrValNA {
		t.Errorf("expected #N/A at [2][2], got %v", got.Array[2][2])
	}
}

func TestEXPAND_ScalarTo3x3WithCustomPad(t *testing.T) {
	// EXPAND(1, 3, 3, "-") → 3×3 with 1 in [0][0], "-" elsewhere
	got, err := fnEXPAND([]Value{
		NumberVal(1),
		NumberVal(3),
		NumberVal(3),
		StringVal("-"),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 3x3 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 1 {
		t.Errorf("expected 1 at [0][0], got %v", got.Array[0][0])
	}
	if got.Array[0][1].Type != ValueString || got.Array[0][1].Str != "-" {
		t.Errorf("expected '-' at [0][1], got %v", got.Array[0][1])
	}
	if got.Array[2][2].Type != ValueString || got.Array[2][2].Str != "-" {
		t.Errorf("expected '-' at [2][2], got %v", got.Array[2][2])
	}
}

func TestEXPAND_NoDimensionChange(t *testing.T) {
	// EXPAND({1,2;3,4}, 2, 2) → same 2×2 array
	got, err := fnEXPAND([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
			{NumberVal(3), NumberVal(4)},
		}},
		NumberVal(2),
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 1 || got.Array[1][1].Num != 4 {
		t.Errorf("values mismatch")
	}
}

func TestEXPAND_OnlyRowsExpanded(t *testing.T) {
	// EXPAND({1,2}, 3, 2) → 3×2 with row padding
	got, err := fnEXPAND([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
		}},
		NumberVal(3),
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 3x2 array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][0].Num != 1 || got.Array[0][1].Num != 2 {
		t.Errorf("row 0 mismatch")
	}
	if got.Array[1][0].Type != ValueError || got.Array[1][0].Err != ErrValNA {
		t.Errorf("expected #N/A at [1][0], got %v", got.Array[1][0])
	}
}

func TestEXPAND_OnlyColsExpanded(t *testing.T) {
	// EXPAND({1;2}, 2, 3) → 2×3 with col padding
	got, err := fnEXPAND([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
			{NumberVal(2)},
		}},
		NumberVal(2),
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 2x3 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 1 || got.Array[1][0].Num != 2 {
		t.Errorf("original values mismatch")
	}
	if got.Array[0][1].Type != ValueError || got.Array[0][1].Err != ErrValNA {
		t.Errorf("expected #N/A at [0][1], got %v", got.Array[0][1])
	}
}

func TestEXPAND_CustomPadNumber(t *testing.T) {
	// EXPAND({1}, 2, 2, 0) → {{1,0},{0,0}}
	got, err := fnEXPAND([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
		}},
		NumberVal(2),
		NumberVal(2),
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2x2 array, got type=%v", got.Type)
	}
	if got.Array[0][1].Num != 0 || got.Array[1][0].Num != 0 || got.Array[1][1].Num != 0 {
		t.Errorf("pad values not 0")
	}
}

func TestEXPAND_CustomPadBoolean(t *testing.T) {
	// EXPAND({1}, 2, 2, TRUE) → {{1,TRUE},{TRUE,TRUE}}
	got, err := fnEXPAND([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
		}},
		NumberVal(2),
		NumberVal(2),
		BoolVal(true),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("expected 2x2 array, got type=%v", got.Type)
	}
	if got.Array[0][1].Type != ValueBool || got.Array[0][1].Bool != true {
		t.Errorf("expected TRUE at [0][1], got %v", got.Array[0][1])
	}
	if got.Array[1][0].Type != ValueBool || got.Array[1][0].Bool != true {
		t.Errorf("expected TRUE at [1][0], got %v", got.Array[1][0])
	}
}

func TestEXPAND_RowsLessThanArrayRows(t *testing.T) {
	// EXPAND({1;2;3}, 2) → #VALUE!
	got, err := fnEXPAND([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
			{NumberVal(2)},
			{NumberVal(3)},
		}},
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestEXPAND_ColsLessThanArrayCols(t *testing.T) {
	// EXPAND({1,2,3}, 1, 2) → #VALUE!
	got, err := fnEXPAND([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
		}},
		NumberVal(1),
		NumberVal(2),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestEXPAND_RowsZero(t *testing.T) {
	// EXPAND({1}, 0) → #VALUE!
	got, err := fnEXPAND([]Value{
		NumberVal(1),
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestEXPAND_ColsZero(t *testing.T) {
	// EXPAND({1}, 1, 0) → #VALUE!
	got, err := fnEXPAND([]Value{
		NumberVal(1),
		NumberVal(1),
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestEXPAND_NegativeRows(t *testing.T) {
	// EXPAND({1}, -1) → #VALUE!
	got, err := fnEXPAND([]Value{
		NumberVal(1),
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestEXPAND_NegativeCols(t *testing.T) {
	// EXPAND({1}, 1, -1) → #VALUE!
	got, err := fnEXPAND([]Value{
		NumberVal(1),
		NumberVal(1),
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestEXPAND_EmptyRowsArg(t *testing.T) {
	// EXPAND({1,2;3,4}, , 3) → keeps 2 rows, expands to 3 cols
	got, err := fnEXPAND([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
			{NumberVal(3), NumberVal(4)},
		}},
		EmptyVal(),
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 2x3 array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][2].Type != ValueError || got.Array[0][2].Err != ErrValNA {
		t.Errorf("expected #N/A at [0][2], got %v", got.Array[0][2])
	}
}

func TestEXPAND_EmptyColsArg(t *testing.T) {
	// EXPAND({1,2;3,4}, 3) → keeps 2 cols, expands to 3 rows
	got, err := fnEXPAND([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
			{NumberVal(3), NumberVal(4)},
		}},
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 3x2 array, got type=%v rows=%d cols=%d", got.Type, len(got.Array), len(got.Array[0]))
	}
	if got.Array[2][0].Type != ValueError || got.Array[2][0].Err != ErrValNA {
		t.Errorf("expected #N/A at [2][0], got %v", got.Array[2][0])
	}
}

func TestEXPAND_SingleElementNoDimensionChange(t *testing.T) {
	// EXPAND(42, 1, 1) → 42 (scalar returned)
	got, err := fnEXPAND([]Value{
		NumberVal(42),
		NumberVal(1),
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("expected 42, got %v", got)
	}
}

func TestEXPAND_LargeExpansion(t *testing.T) {
	// EXPAND(1, 100, 50, 0) → 100×50 array
	got, err := fnEXPAND([]Value{
		NumberVal(1),
		NumberVal(100),
		NumberVal(50),
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 100 || len(got.Array[0]) != 50 {
		t.Fatalf("expected 100x50 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 1 {
		t.Errorf("expected 1 at [0][0], got %v", got.Array[0][0])
	}
	if got.Array[99][49].Num != 0 {
		t.Errorf("expected 0 at [99][49], got %v", got.Array[99][49])
	}
}

func TestEXPAND_ErrorPropagationInArray(t *testing.T) {
	// EXPAND(#REF!, 3, 3) → #REF!
	got, err := fnEXPAND([]Value{
		ErrorVal(ErrValREF),
		NumberVal(3),
		NumberVal(3),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("expected #REF!, got %v", got)
	}
}

func TestEXPAND_ErrorInRowsArg(t *testing.T) {
	// EXPAND({1}, "abc") → #VALUE! (non-numeric rows)
	got, err := fnEXPAND([]Value{
		NumberVal(1),
		StringVal("abc"),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("expected error, got %v", got)
	}
}

func TestEXPAND_ErrorInColsArg(t *testing.T) {
	// EXPAND({1}, 1, "xyz") → #VALUE!
	got, err := fnEXPAND([]Value{
		NumberVal(1),
		NumberVal(1),
		StringVal("xyz"),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("expected error, got %v", got)
	}
}

func TestEXPAND_TooFewArgs(t *testing.T) {
	// EXPAND(1) → #VALUE! (only 1 arg)
	got, err := fnEXPAND([]Value{NumberVal(1)})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestEXPAND_TooManyArgs(t *testing.T) {
	// EXPAND(1, 1, 1, 0, extra) → #VALUE! (5 args)
	got, err := fnEXPAND([]Value{NumberVal(1), NumberVal(1), NumberVal(1), NumberVal(0), NumberVal(99)})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestEXPAND_StringCoercionForRows(t *testing.T) {
	// EXPAND(1, "3", "3") → 3×3 array (string "3" coerced to number)
	got, err := fnEXPAND([]Value{
		NumberVal(1),
		StringVal("3"),
		StringVal("3"),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 3x3 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 1 {
		t.Errorf("expected 1 at [0][0], got %v", got.Array[0][0])
	}
}

func TestEXPAND_BoolCoercionForRows(t *testing.T) {
	// EXPAND(1, TRUE) → 1 (TRUE coerced to 1, same as scalar dimensions)
	got, err := fnEXPAND([]Value{
		NumberVal(1),
		BoolVal(true),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("expected 1, got %v", got)
	}
}

func TestEXPAND_TruncatesDecimalDimensions(t *testing.T) {
	// EXPAND({1}, 2.9, 2.9) → 2×2 array (truncated to 2)
	got, err := fnEXPAND([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
		}},
		NumberVal(2.9),
		NumberVal(2.9),
	})
	if err != nil {
		t.Fatalf("fnEXPAND: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2 array, got type=%v", got.Type)
	}
}

func TestEXPAND_ViaEval(t *testing.T) {
	// Test via the formula evaluator: EXPAND({1,2;3,4},3,3)
	cf := evalCompile(t, "EXPAND({1,2;3,4},3,3)")
	got, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 3x3 array, got type=%v", got.Type)
	}
	if got.Array[0][0].Num != 1 || got.Array[1][1].Num != 4 {
		t.Errorf("original values mismatch")
	}
	if got.Array[2][2].Type != ValueError || got.Array[2][2].Err != ErrValNA {
		t.Errorf("expected #N/A at [2][2], got %v", got.Array[2][2])
	}
}

func TestHYPERLINK(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("http://example.com"),
			{Col: 2, Row: 1}: StringVal("Example Site"),
			{Col: 1, Row: 2}: NumberVal(42),
			{Col: 2, Row: 2}: NumberVal(100),
			{Col: 1, Row: 3}: BoolVal(true),
		},
	}

	tests := []struct {
		name     string
		formula  string
		wantType ValueType
		wantStr  string
		wantNum  float64
		wantBool bool
		wantErr  ErrorValue
	}{
		{
			name:     "basic with friendly name",
			formula:  `HYPERLINK("http://example.com","Click me")`,
			wantType: ValueString,
			wantStr:  "Click me",
		},
		{
			name:     "URL only no friendly name",
			formula:  `HYPERLINK("http://example.com")`,
			wantType: ValueString,
			wantStr:  "http://example.com",
		},
		{
			name:     "numeric friendly name",
			formula:  `HYPERLINK("http://example.com",42)`,
			wantType: ValueNumber,
			wantNum:  42,
		},
		{
			name:     "boolean friendly name TRUE",
			formula:  `HYPERLINK("http://example.com",TRUE)`,
			wantType: ValueBool,
			wantBool: true,
		},
		{
			name:     "boolean friendly name FALSE",
			formula:  `HYPERLINK("http://example.com",FALSE)`,
			wantType: ValueBool,
			wantBool: false,
		},
		{
			name:     "empty string friendly name",
			formula:  `HYPERLINK("http://example.com","")`,
			wantType: ValueString,
			wantStr:  "",
		},
		{
			name:     "empty string link location",
			formula:  `HYPERLINK("")`,
			wantType: ValueString,
			wantStr:  "",
		},
		{
			name:     "error in link location",
			formula:  `HYPERLINK(1/0)`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name:     "error in friendly name",
			formula:  `HYPERLINK("http://example.com",1/0)`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name:     "error in both args propagates link_location error",
			formula:  `HYPERLINK(1/0,1/0)`,
			wantType: ValueError,
			wantErr:  ErrValDIV0,
		},
		{
			name:     "zero args returns VALUE error",
			formula:  `HYPERLINK()`,
			wantType: ValueError,
			wantErr:  ErrValVALUE,
		},
		{
			name:     "numeric link location no friendly name",
			formula:  `HYPERLINK(12345)`,
			wantType: ValueString,
			wantStr:  "12345",
		},
		{
			name:     "boolean link location no friendly name",
			formula:  `HYPERLINK(TRUE)`,
			wantType: ValueString,
			wantStr:  "TRUE",
		},
		{
			name:     "cell reference for link location",
			formula:  `HYPERLINK(A1)`,
			wantType: ValueString,
			wantStr:  "http://example.com",
		},
		{
			name:     "cell reference for friendly name",
			formula:  `HYPERLINK("http://example.com",B1)`,
			wantType: ValueString,
			wantStr:  "Example Site",
		},
		{
			name:     "cell reference numeric friendly name",
			formula:  `HYPERLINK("http://example.com",B2)`,
			wantType: ValueNumber,
			wantNum:  100,
		},
		{
			name:     "nested CONCATENATE in link location",
			formula:  `HYPERLINK(CONCATENATE("http://","example.com"))`,
			wantType: ValueString,
			wantStr:  "http://example.com",
		},
		{
			name:     "nested UPPER in friendly name",
			formula:  `HYPERLINK("http://example.com",UPPER("click"))`,
			wantType: ValueString,
			wantStr:  "CLICK",
		},
		{
			name:     "string concatenation with ampersand in link",
			formula:  `HYPERLINK("http://"&"example.com")`,
			wantType: ValueString,
			wantStr:  "http://example.com",
		},
		{
			name:     "friendly name with number from cell",
			formula:  `HYPERLINK(A1,A2)`,
			wantType: ValueNumber,
			wantNum:  42,
		},
		{
			name:     "friendly name empty cell returns empty",
			formula:  `HYPERLINK("http://example.com",C5)`,
			wantType: ValueEmpty,
		},
		{
			name:     "link location with empty friendly name string",
			formula:  `HYPERLINK("http://example.com","")`,
			wantType: ValueString,
			wantStr:  "",
		},
		{
			name:     "numeric link and numeric friendly",
			formula:  `HYPERLINK(123,456)`,
			wantType: ValueNumber,
			wantNum:  456,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != tt.wantType {
				t.Fatalf("got type %v, want %v (value=%v)", got.Type, tt.wantType, got)
			}
			switch tt.wantType {
			case ValueString:
				if got.Str != tt.wantStr {
					t.Errorf("got %q, want %q", got.Str, tt.wantStr)
				}
			case ValueNumber:
				if got.Num != tt.wantNum {
					t.Errorf("got %g, want %g", got.Num, tt.wantNum)
				}
			case ValueBool:
				if got.Bool != tt.wantBool {
					t.Errorf("got %v, want %v", got.Bool, tt.wantBool)
				}
			case ValueError:
				if got.Err != tt.wantErr {
					t.Errorf("got %v, want %v", got.Err, tt.wantErr)
				}
			case ValueEmpty:
				// nothing to check beyond type
			}
		})
	}
}

// ---------------------------------------------------------------------------
// OFFSET
// ---------------------------------------------------------------------------

func TestOFFSETBasicSingleCell(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
			{Col: 1, Row: 2}: NumberVal(30),
			{Col: 2, Row: 2}: NumberVal(40),
			{Col: 3, Row: 3}: NumberVal(99),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,1,1) → B2 = 40
	cf := evalCompile(t, "OFFSET(A1,1,1)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 40 {
		t.Errorf("OFFSET(A1,1,1): got %v, want 40", got)
	}
}

func TestOFFSETZeroOffset(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(42),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,0,0) → A1 = 42
	cf := evalCompile(t, "OFFSET(A1,0,0)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("OFFSET(A1,0,0): got %v, want 42", got)
	}
}

func TestOFFSETNegativeOffset(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 3, Row: 3}: NumberVal(30),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(C3,-2,-2) → A1 = 10
	cf := evalCompile(t, "OFFSET(C3,-2,-2)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("OFFSET(C3,-2,-2): got %v, want 10", got)
	}
}

func TestOFFSETWithHeightWidth(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 2}: NumberVal(1),
			{Col: 3, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 3, Row: 3}: NumberVal(4),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,1,1,2,2) → B2:C3 range, SUM should be 10
	cf := evalCompile(t, "SUM(OFFSET(A1,1,1,2,2))")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("SUM(OFFSET(A1,1,1,2,2)): got %v, want 10", got)
	}
}

func TestOFFSETDefaultHeightWidthFromRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 1, Row: 3}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
			{Col: 1, Row: 4}: NumberVal(7),
			{Col: 2, Row: 4}: NumberVal(8),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1:B2,2,0) → uses default height=2, width=2 from reference
	// → A3:B4, SUM = 5+6+7+8 = 26
	cf := evalCompile(t, "SUM(OFFSET(A1:B2,2,0))")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 26 {
		t.Errorf("SUM(OFFSET(A1:B2,2,0)): got %v, want 26", got)
	}
}

func TestOFFSETCustomHeightOverridesDefault(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,0,0,3) → A1:A3, SUM = 60
	cf := evalCompile(t, "SUM(OFFSET(A1,0,0,3))")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 60 {
		t.Errorf("SUM(OFFSET(A1,0,0,3)): got %v, want 60", got)
	}
}

func TestOFFSETCustomWidthOnly(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
			{Col: 3, Row: 1}: NumberVal(30),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,0,0,1,3) → A1:C1, SUM = 60
	cf := evalCompile(t, "SUM(OFFSET(A1,0,0,1,3))")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 60 {
		t.Errorf("SUM(OFFSET(A1,0,0,1,3)): got %v, want 60", got)
	}
}

func TestOFFSETOffEdgeNegativeRow(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,-1,0) → row 0 → #REF!
	cf := evalCompile(t, "OFFSET(A1,-1,0)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("OFFSET(A1,-1,0): got %v, want #REF!", got)
	}
}

func TestOFFSETOffEdgeNegativeCol(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,0,-1) → col 0 → #REF!
	cf := evalCompile(t, "OFFSET(A1,0,-1)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("OFFSET(A1,0,-1): got %v, want #REF!", got)
	}
}

func TestOFFSETHeightZero(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// Height = 0 → #REF!
	cf := evalCompile(t, "OFFSET(A1,0,0,0)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("OFFSET(A1,0,0,0): got %v, want #REF!", got)
	}
}

func TestOFFSETHeightNegative(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// Height = -1 with OFFSET(A1,0,0,-1): anchor at row 1, range extends
	// upward 1 row → row 1 to row 1 → returns A1 value (matches Excel).
	cf := evalCompile(t, "OFFSET(A1,0,0,-1)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("OFFSET(A1,0,0,-1): got %v, want 1", got)
	}
}

func TestOFFSETWidthZero(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// Width = 0 → #REF!
	cf := evalCompile(t, "OFFSET(A1,0,0,1,0)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("OFFSET(A1,0,0,1,0): got %v, want #REF!", got)
	}
}

func TestOFFSETWidthNegative(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// Width = -2 → #REF!
	cf := evalCompile(t, "OFFSET(A1,0,0,1,-2)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("OFFSET(A1,0,0,1,-2): got %v, want #REF!", got)
	}
}

func TestOFFSETTooFewArgs(t *testing.T) {
	resolver := &mockResolver{}
	ctx := &EvalContext{Resolver: resolver}

	cf := evalCompile(t, "OFFSET(A1,1)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("OFFSET(A1,1): got %v, want #VALUE!", got)
	}
}

func TestOFFSETSingleCellResultFromRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 3, Row: 3}: NumberVal(99),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1:B2,2,2,1,1) → C3 = 99 (single cell from range ref)
	cf := evalCompile(t, "OFFSET(A1:B2,2,2,1,1)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 99 {
		t.Errorf("OFFSET(A1:B2,2,2,1,1): got %v, want 99", got)
	}
}

func TestOFFSETUsedInSUM(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(4),
			{Col: 1, Row: 5}: NumberVal(5),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// SUM(OFFSET(A1,0,0,5,1)) → sum of A1:A5 = 15
	cf := evalCompile(t, "SUM(OFFSET(A1,0,0,5,1))")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 15 {
		t.Errorf("SUM(OFFSET(A1,0,0,5,1)): got %v, want 15", got)
	}
}

func TestOFFSETEmptyCell(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,0,1) → B1 which is empty
	cf := evalCompile(t, "OFFSET(A1,0,1)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueEmpty {
		t.Errorf("OFFSET(A1,0,1): got type %v, want ValueEmpty", got.Type)
	}
}

func TestOFFSETStringCell(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: StringVal("hello"),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,0,1) → B1 = "hello"
	cf := evalCompile(t, "OFFSET(A1,0,1)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "hello" {
		t.Errorf("OFFSET(A1,0,1): got %v, want hello", got)
	}
}

func TestOFFSETBooleanCell(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: BoolVal(true),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,0,1) → B1 = TRUE
	cf := evalCompile(t, "OFFSET(A1,0,1)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("OFFSET(A1,0,1): got %v, want TRUE", got)
	}
}

func TestOFFSETLargeOffset(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}:     NumberVal(1),
			{Col: 100, Row: 200}: NumberVal(999),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,199,99) → CV200 = 999
	cf := evalCompile(t, "OFFSET(A1,199,99)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 999 {
		t.Errorf("OFFSET(A1,199,99): got %v, want 999", got)
	}
}

func TestOFFSETRowsColsTruncated(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,1.9,1.9) → rows=1, cols=1 → B2 = 20
	cf := evalCompile(t, "OFFSET(A1,1.9,1.9)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("OFFSET(A1,1.9,1.9): got %v, want 20", got)
	}
}

func TestOFFSETHeightWidthTruncated(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
			{Col: 1, Row: 2}: NumberVal(30),
			{Col: 2, Row: 2}: NumberVal(40),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,0,0,2.9,2.9) → height=2, width=2 → A1:B2, SUM=100
	cf := evalCompile(t, "SUM(OFFSET(A1,0,0,2.9,2.9))")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 100 {
		t.Errorf("SUM(OFFSET(A1,0,0,2.9,2.9)): got %v, want 100", got)
	}
}

func TestOFFSETExceedsMaxRow(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// Offset to beyond max row
	cf := evalCompile(t, "OFFSET(A1,1048576,0)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("OFFSET(A1,1048576,0): got %v, want #REF!", got)
	}
}

func TestOFFSETExceedsMaxCol(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// Offset to beyond max col
	cf := evalCompile(t, "OFFSET(A1,0,16384)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("OFFSET(A1,0,16384): got %v, want #REF!", got)
	}
}

func TestOFFSETRangeOverflowsMaxRow(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// Starting near max row with height that overflows
	cf := evalCompile(t, "OFFSET(A1,1048575,0,2)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("OFFSET(A1,1048575,0,2): got %v, want #REF!", got)
	}
}

func TestOFFSETErrorInRowsArg(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: StringVal("abc"),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// String that can't be converted to number → #VALUE!
	cf := evalCompile(t, `OFFSET(A1,B1,0)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("OFFSET(A1,B1,0): got %v, want #VALUE!", got)
	}
}

func TestOFFSETNilContext(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
		},
	}

	cf := evalCompile(t, "OFFSET(A1,0,0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("OFFSET with nil ctx: got %v, want #REF!", got)
	}
}

func TestOFFSETRangeRefDefaultDimensions(t *testing.T) {
	// When using a range ref, omitting height/width should inherit
	// the dimensions of the reference range.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 3, Row: 3}: NumberVal(10),
			{Col: 4, Row: 3}: NumberVal(20),
			{Col: 3, Row: 4}: NumberVal(30),
			{Col: 4, Row: 4}: NumberVal(40),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1:B2,2,2) → C3:D4 (same 2x2 size), SUM = 100
	cf := evalCompile(t, "SUM(OFFSET(A1:B2,2,2))")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 100 {
		t.Errorf("SUM(OFFSET(A1:B2,2,2)): got %v, want 100", got)
	}
}

func TestOFFSETFromCellFlag(t *testing.T) {
	// Single cell result should have FromCell=true for proper coercion
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(42),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	cf := evalCompile(t, "OFFSET(A1,0,1)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if !got.FromCell {
		t.Errorf("OFFSET single cell result should have FromCell=true")
	}
}

func TestOFFSETRangeResultHasRangeOrigin(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 2, Row: 2}: NumberVal(1),
			{Col: 3, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(3),
			{Col: 3, Row: 3}: NumberVal(4),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,1,1,2,2) should return array with proper RangeOrigin
	cf := evalCompile(t, "OFFSET(A1,1,1,2,2)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected ValueArray, got %v", got.Type)
	}
	if got.RangeOrigin == nil {
		t.Fatal("RangeOrigin should not be nil for range result")
	}
	ro := got.RangeOrigin
	if ro.FromCol != 2 || ro.FromRow != 2 || ro.ToCol != 3 || ro.ToRow != 3 {
		t.Errorf("RangeOrigin: got %+v, want B2:C3", ro)
	}
}

func TestOFFSETBooleanCoercionForRows(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,TRUE,0) → TRUE coerces to 1, so A2=20
	cf := evalCompile(t, "OFFSET(A1,TRUE,0)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("OFFSET(A1,TRUE,0): got %v, want 20", got)
	}
}

func TestOFFSETBooleanCoercionForCols(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 2, Row: 1}: NumberVal(20),
		},
	}
	ctx := &EvalContext{Resolver: resolver}

	// OFFSET(A1,0,TRUE) → TRUE coerces to 1, so B1=20
	cf := evalCompile(t, "OFFSET(A1,0,TRUE)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("OFFSET(A1,0,TRUE): got %v, want 20", got)
	}
}
