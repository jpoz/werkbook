package formula

import (
	"math"
	"testing"
)

// mockResolver implements CellResolver for testing.
type mockResolver struct {
	cells map[CellAddr]Value
}

func (m *mockResolver) GetCellValue(addr CellAddr) Value {
	if v, ok := m.cells[addr]; ok {
		return v
	}
	return EmptyVal()
}

func (m *mockResolver) GetRangeValues(addr RangeAddr) [][]Value {
	rows := make([][]Value, addr.ToRow-addr.FromRow+1)
	for r := addr.FromRow; r <= addr.ToRow; r++ {
		row := make([]Value, addr.ToCol-addr.FromCol+1)
		for c := addr.FromCol; c <= addr.ToCol; c++ {
			ca := CellAddr{Sheet: addr.Sheet, Col: c, Row: r}
			if v, ok := m.cells[ca]; ok {
				row[c-addr.FromCol] = v
			}
		}
		rows[r-addr.FromRow] = row
	}
	return rows
}

func evalCompile(t *testing.T, formula string) *CompiledFormula {
	t.Helper()
	node, err := Parse(formula)
	if err != nil {
		t.Fatalf("Parse(%q): %v", formula, err)
	}
	cf, err := Compile(formula, node)
	if err != nil {
		t.Fatalf("Compile(%q): %v", formula, err)
	}
	return cf
}

func TestEvalArithmetic(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    float64
	}{
		{"1+2*3", 7},
		{"(1+2)*3", 9},
		{"10-3", 7},
		{"2^3", 8},
		{"10/4", 2.5},
		{"-5", -5},
		{"50%", 0.5},
		{"2+3*4-1", 13},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueNumber || got.Num != tt.want {
			t.Errorf("Eval(%q) = %v (%g), want %g", tt.formula, got.Type, got.Num, tt.want)
		}
	}
}

func TestEvalCellReference(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
		},
	}

	cf := evalCompile(t, "A1*2")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("got %v (%g), want 20", got.Type, got.Num)
	}
}

func TestEvalSUMRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}

	cf := evalCompile(t, "SUM(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 60 {
		t.Errorf("got %v (%g), want 60", got.Type, got.Num)
	}
}

func TestEvalStringConcat(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
		},
	}

	cf := evalCompile(t, `A1&" items"`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "10 items" {
		t.Errorf("got %v (%q), want %q", got.Type, got.Str, "10 items")
	}
}

func TestEvalIF(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `IF(TRUE,"yes","no")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "yes" {
		t.Errorf("got %v (%q), want %q", got.Type, got.Str, "yes")
	}

	cf = evalCompile(t, `IF(FALSE,"yes","no")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "no" {
		t.Errorf("got %v (%q), want %q", got.Type, got.Str, "no")
	}
}

func TestEvalComparison(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "10>5")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("got %v (%v), want TRUE", got.Type, got.Bool)
	}
}

func TestEvalDivByZero(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "1/0")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("got %v (err=%v), want #DIV/0!", got.Type, got.Err)
	}
}

func TestEvalAVERAGE(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
		},
	}

	cf := evalCompile(t, "AVERAGE(A1:A2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 15 {
		t.Errorf("got %v (%g), want 15", got.Type, got.Num)
	}
}

func TestEvalMINMAX(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(15),
			{Col: 1, Row: 3}: NumberVal(10),
		},
	}

	cf := evalCompile(t, "MIN(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("MIN: got %g, want 5", got.Num)
	}

	cf = evalCompile(t, "MAX(A1:A3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 15 {
		t.Errorf("MAX: got %g, want 15", got.Num)
	}
}

func TestEvalCOUNT(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: StringVal("hello"),
			{Col: 1, Row: 3}: NumberVal(10),
		},
	}

	cf := evalCompile(t, "COUNT(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNT: got %g, want 2", got.Num)
	}
}

func TestEvalROUND(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "ROUND(3.14159,2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || math.Abs(got.Num-3.14) > 1e-10 {
		t.Errorf("ROUND: got %g, want 3.14", got.Num)
	}
}

func TestEvalStringFunctions(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    string
	}{
		{`UPPER("hello")`, "HELLO"},
		{`LOWER("HELLO")`, "hello"},
		{`TRIM("  hello   world  ")`, "hello world"},
		{`LEFT("hello",3)`, "hel"},
		{`RIGHT("hello",3)`, "llo"},
		{`MID("hello",2,3)`, "ell"},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueString || got.Str != tt.want {
			t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
		}
	}
}

func TestEvalLEN(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `LEN("hello")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("LEN: got %g, want 5", got.Num)
	}
}

func TestEvalLogical(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    bool
	}{
		{"AND(TRUE,TRUE)", true},
		{"AND(TRUE,FALSE)", false},
		{"OR(FALSE,TRUE)", true},
		{"OR(FALSE,FALSE)", false},
		{"NOT(TRUE)", false},
		{"NOT(FALSE)", true},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueBool || got.Bool != tt.want {
			t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want)
		}
	}
}

func TestEvalIFERROR(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `IFERROR(1/0,"err")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "err" {
		t.Errorf("got %v (%q), want %q", got.Type, got.Str, "err")
	}

	cf = evalCompile(t, `IFERROR(42,"err")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("got %v (%g), want 42", got.Type, got.Num)
	}
}

func TestEvalImplicitIntersectionFullColumn(t *testing.T) {
	// Simulate: formula in row 2 references F:F (full column).
	// In non-array formula context, F:F should be implicitly intersected
	// to a single cell at the formula's own row.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 6, Row: 1}: StringVal("Header"),
			{Col: 6, Row: 2}: NumberVal(-250264),
			{Col: 6, Row: 3}: NumberVal(250264),
			{Col: 6, Row: 4}: NumberVal(-5750000),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     1,
		CurrentRow:     2,
		CurrentSheet:   "Outputs",
		IsArrayFormula: false,
	}

	// ABS(F:F) in row 2 with implicit intersection → ABS(F2) = 250264
	cf := evalCompile(t, "ABS(F:F)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 250264 {
		t.Errorf("ABS(F:F) with implicit intersection = %v (%g), want 250264", got.Type, got.Num)
	}

}

func TestEvalSUMFullRowRange(t *testing.T) {
	// SUM(5:6) should sum all values in rows 5 and 6 across all columns.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 5}: NumberVal(10),
			{Col: 2, Row: 5}: NumberVal(20),
			{Col: 3, Row: 5}: NumberVal(30),
			{Col: 1, Row: 6}: NumberVal(40),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     5,
		CurrentRow:     1,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: true,
	}

	cf := evalCompile(t, "SUM(5:6)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 100 {
		t.Errorf("SUM(5:6) = %v (%g), want 100", got.Type, got.Num)
	}
}

func TestEvalImplicitIntersectionFullRow(t *testing.T) {
	// In a non-array formula context, 1:1 should intersect at the current column.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(100),
			{Col: 2, Row: 1}: NumberVal(200),
			{Col: 3, Row: 1}: NumberVal(300),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     2,
		CurrentRow:     5,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	// ABS(1:1) in col 2 with implicit intersection → ABS(B1) = 200
	cf := evalCompile(t, "ABS(1:1)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 200 {
		t.Errorf("ABS(1:1) with implicit intersection = %v (%g), want 200", got.Type, got.Num)
	}
}

func TestEvalArrayFormulaFullColumn(t *testing.T) {
	// When IsArrayFormula=true, F:F should load as a full array.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}

	ctx := &EvalContext{
		CurrentCol:     2,
		CurrentRow:     1,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: true,
	}

	cf := evalCompile(t, "SUM(A:A)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 60 {
		t.Errorf("SUM(A:A) array formula = %v (%g), want 60", got.Type, got.Num)
	}
}

// ---------------------------------------------------------------------------
// coerceNum edge cases (exercised through arithmetic operations)
// ---------------------------------------------------------------------------

func TestEvalCoerceNum(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		cells   map[CellAddr]Value
		wantNum float64
		wantErr ErrorValue
		isErr   bool
	}{
		// Empty cell treated as 0
		{name: "empty+1", formula: "A1+1", wantNum: 1},
		// Bool true coerced to 1
		{name: "TRUE+0", formula: "TRUE+0", wantNum: 1},
		// Bool false coerced to 0
		{name: "FALSE+5", formula: "FALSE+5", wantNum: 5},
		// Numeric string coerced
		{name: "numeric_string", formula: `"123"+0`, wantNum: 123},
		{name: "numeric_string_float", formula: `"3.14"+0`, wantNum: 3.14},
		// Empty string coerced to 0
		{name: "empty_string", formula: `""+0`, wantNum: 0},
		// Non-numeric string produces #VALUE!
		{name: "non_numeric_string", formula: `"abc"+0`, isErr: true, wantErr: ErrValVALUE},
		// Error propagates through arithmetic
		{name: "error_propagation_add", formula: `#N/A+1`, isErr: true, wantErr: ErrValNA},
		{name: "error_propagation_mul", formula: `#DIV/0!*2`, isErr: true, wantErr: ErrValDIV0},
		// Large numbers
		{name: "large_number", formula: "1e300+1e300", wantNum: 2e300},
		// Negative numbers
		{name: "negative_arithmetic", formula: "-10+-20", wantNum: -30},
		// Chained operations with coercion
		{name: "bool_chain", formula: "TRUE+TRUE+TRUE", wantNum: 3},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cells := tt.cells
			if cells == nil {
				cells = map[CellAddr]Value{}
			}
			resolver := &mockResolver{cells: cells}
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if tt.isErr {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v (err=%v), want error %v", tt.formula, got.Type, got.Err, tt.wantErr)
				}
			} else {
				if got.Type != ValueNumber || got.Num != tt.wantNum {
					t.Errorf("Eval(%q) = %v (%g), want %g", tt.formula, got.Type, got.Num, tt.wantNum)
				}
			}
		})
	}
}

// ---------------------------------------------------------------------------
// compareValues — exercised through comparison operators
// ---------------------------------------------------------------------------

func TestEvalCompareValues(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: StringVal("hello"),
			{Col: 1, Row: 3}: BoolVal(true),
			// A4 is empty
		},
	}

	tests := []struct {
		name    string
		formula string
		want    bool
	}{
		// Same-type number comparisons
		{name: "num_eq", formula: "10=10", want: true},
		{name: "num_ne", formula: "10<>5", want: true},
		{name: "num_lt", formula: "5<10", want: true},
		{name: "num_le", formula: "10<=10", want: true},
		{name: "num_gt", formula: "10>5", want: true},
		{name: "num_ge", formula: "10>=10", want: true},
		// Same-type string comparisons (case-insensitive)
		{name: "str_eq_case", formula: `"Hello"="hello"`, want: true},
		{name: "str_lt", formula: `"abc"<"def"`, want: true},
		{name: "str_gt", formula: `"xyz">"abc"`, want: true},
		// Same-type bool comparisons
		{name: "bool_eq", formula: "TRUE=TRUE", want: true},
		{name: "bool_ne", formula: "TRUE<>FALSE", want: true},
		{name: "bool_order", formula: "FALSE<TRUE", want: true},
		// Empty cell = 0
		{name: "empty_eq_zero", formula: "A4=0", want: true},
		// Cross-type comparisons (via typeRank: error < number < string < bool)
		{name: "num_lt_str", formula: `10<"hello"`, want: true},
		{name: "str_lt_bool", formula: `"hello"<TRUE`, want: true},
		// Negative numbers
		{name: "negative_lt", formula: "-5<0", want: true},
		{name: "negative_gt", formula: "0>-10", want: true},
		// Decimal comparisons
		{name: "decimal_eq", formula: "0.1+0.2=0.3", want: false}, // floating point!
		{name: "decimal_lt", formula: "0.1<0.2", want: true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueBool || got.Bool != tt.want {
				t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// isTruthy — exercised through IF
// ---------------------------------------------------------------------------

func TestEvalIsTruthy(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(0),
			{Col: 1, Row: 2}: NumberVal(1),
			{Col: 1, Row: 3}: StringVal(""),
			{Col: 1, Row: 4}: StringVal("x"),
			// A5 is empty
		},
	}

	tests := []struct {
		name    string
		formula string
		want    string // "yes" or "no"
	}{
		{name: "bool_true", formula: `IF(TRUE,"yes","no")`, want: "yes"},
		{name: "bool_false", formula: `IF(FALSE,"yes","no")`, want: "no"},
		{name: "num_zero", formula: `IF(A1,"yes","no")`, want: "no"},
		{name: "num_nonzero", formula: `IF(A2,"yes","no")`, want: "yes"},
		{name: "str_empty", formula: `IF(A3,"yes","no")`, want: "no"},
		{name: "str_nonempty", formula: `IF(A4,"yes","no")`, want: "yes"},
		{name: "empty_cell", formula: `IF(A5,"yes","no")`, want: "no"},
		{name: "num_negative", formula: `IF(-1,"yes","no")`, want: "yes"},
		{name: "num_fraction", formula: `IF(0.001,"yes","no")`, want: "yes"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// valueToString — exercised through the & (concat) operator
// ---------------------------------------------------------------------------

func TestEvalValueToString(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(42),
			{Col: 1, Row: 2}: BoolVal(true),
			{Col: 1, Row: 3}: BoolVal(false),
			// A4 empty
		},
	}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "number_to_str", formula: `A1&""`, want: "42"},
		{name: "bool_true_to_str", formula: `A2&""`, want: "TRUE"},
		{name: "bool_false_to_str", formula: `A3&""`, want: "FALSE"},
		{name: "empty_to_str", formula: `A4&"x"`, want: "x"},
		{name: "string_concat", formula: `"hello"&" "&"world"`, want: "hello world"},
		{name: "float_to_str", formula: `3.14&""`, want: "3.14"},
		{name: "error_to_str", formula: `#N/A&""`, want: "#N/A"},
		{name: "div0_to_str", formula: `#DIV/0!&""`, want: "#DIV/0!"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// errorValueToString — exercised through concat with error literals
// ---------------------------------------------------------------------------

func TestEvalErrorValueToString(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "DIV0", formula: `#DIV/0!&""`, want: "#DIV/0!"},
		{name: "NA", formula: `#N/A&""`, want: "#N/A"},
		{name: "NAME", formula: `#NAME?&""`, want: "#NAME?"},
		{name: "NULL", formula: `#NULL!&""`, want: "#NULL!"},
		{name: "NUM", formula: `#NUM!&""`, want: "#NUM!"},
		{name: "REF", formula: `#REF!&""`, want: "#REF!"},
		{name: "VALUE", formula: `#VALUE!&""`, want: "#VALUE!"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Division and power edge cases
// ---------------------------------------------------------------------------

func TestEvalDivisionEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("div_by_zero", func(t *testing.T) {
		cf := evalCompile(t, "1/0")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})
	t.Run("zero_div_zero", func(t *testing.T) {
		cf := evalCompile(t, "0/0")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0!", got)
		}
	})
	t.Run("large_div_overflow", func(t *testing.T) {
		cf := evalCompile(t, "1e300/1e-300")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || !math.IsInf(got.Num, 1) {
			t.Errorf("got %v (%g), want +Inf", got.Type, got.Num)
		}
	})
	t.Run("negative_div", func(t *testing.T) {
		cf := evalCompile(t, "-10/3")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || math.Abs(got.Num-(-10.0/3.0)) > 1e-10 {
			t.Errorf("got %g, want %g", got.Num, -10.0/3.0)
		}
	})
	t.Run("power_zero", func(t *testing.T) {
		cf := evalCompile(t, "0^0")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %g, want 1", got.Num)
		}
	})
	t.Run("power_negative_int", func(t *testing.T) {
		cf := evalCompile(t, "(-2)^3")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -8 {
			t.Errorf("got %g, want -8", got.Num)
		}
	})
	t.Run("power_fractional", func(t *testing.T) {
		cf := evalCompile(t, "4^0.5")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %g, want 2", got.Num)
		}
	})
}

// ---------------------------------------------------------------------------
// Unary negation and percent with various types
// ---------------------------------------------------------------------------

func TestEvalUnaryEdgeCases(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: BoolVal(true),
			{Col: 1, Row: 2}: StringVal("5"),
		},
	}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		isErr   bool
		wantErr ErrorValue
	}{
		{name: "negate_bool", formula: "-A1", wantNum: -1},
		{name: "negate_numeric_string", formula: "-A2", wantNum: -5},
		{name: "percent_100", formula: "100%", wantNum: 1},
		{name: "percent_50", formula: "50%", wantNum: 0.5},
		{name: "negate_zero", formula: "-0", wantNum: 0},
		{name: "double_negate", formula: "--5", wantNum: 5},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if tt.isErr {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			} else {
				if got.Type != ValueNumber || got.Num != tt.wantNum {
					t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
				}
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Error propagation through all binary operators
// ---------------------------------------------------------------------------

func TestEvalErrorPropagation(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: ErrorVal(ErrValNA),
		},
	}

	ops := []struct {
		name string
		expr string
	}{
		{name: "add_left", expr: "A1+1"},
		{name: "add_right", expr: "1+A1"},
		{name: "sub", expr: "A1-1"},
		{name: "mul", expr: "A1*2"},
		{name: "div", expr: "A1/2"},
		{name: "pow", expr: "A1^2"},
		{name: "neg", expr: "-A1"},
	}

	for _, tt := range ops {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.expr)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.expr, err)
			}
			if got.Type != ValueError || got.Err != ErrValNA {
				t.Errorf("Eval(%q) = %v (err=%v), want #N/A", tt.expr, got.Type, got.Err)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// IF with two args (missing else), NOT edge cases
// ---------------------------------------------------------------------------

func TestEvalIFTwoArgs(t *testing.T) {
	resolver := &mockResolver{}

	// IF(FALSE, "yes") with no third arg should return FALSE
	cf := evalCompile(t, `IF(FALSE,"yes")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool != false {
		t.Errorf("IF(FALSE,yes) = %v, want FALSE", got)
	}
}

func TestEvalNOTEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    bool
	}{
		{"NOT(1)", false},
		{"NOT(0)", true},
		{`NOT("")`, true},
		{`NOT("x")`, false},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueBool || got.Bool != tt.want {
			t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want)
		}
	}
}

// ---------------------------------------------------------------------------
// Large number and decimal precision arithmetic
// ---------------------------------------------------------------------------

func TestEvalLargeNumberArithmetic(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		check   func(Value) bool
	}{
		{
			name:    "max_float_mul",
			formula: "1e308*1",
			check:   func(v Value) bool { return v.Type == ValueNumber && v.Num == 1e308 },
		},
		{
			name:    "overflow_to_inf",
			formula: "1e308*10",
			check:   func(v Value) bool { return v.Type == ValueNumber && math.IsInf(v.Num, 1) },
		},
		{
			name:    "very_small",
			formula: "1e-300+0",
			check:   func(v Value) bool { return v.Type == ValueNumber && v.Num == 1e-300 },
		},
		{
			name:    "subtract_near_equal",
			formula: "1.0000000001-1",
			check: func(v Value) bool {
				return v.Type == ValueNumber && math.Abs(v.Num-0.0000000001) < 1e-15
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if !tt.check(got) {
				t.Errorf("Eval(%q) = %v (%g), check failed", tt.formula, got.Type, got.Num)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Empty cell handling in various contexts
// ---------------------------------------------------------------------------

func TestEvalEmptyCellArithmetic(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			// A2, A3 are empty
		},
	}

	tests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		{name: "empty_add", formula: "A1+A2", wantNum: 5},
		{name: "empty_mul", formula: "A1*A2", wantNum: 0},
		{name: "empty_sub", formula: "A2-A1", wantNum: -5},
		{name: "sum_with_empty", formula: "SUM(A1:A3)", wantNum: 5},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueNumber || got.Num != tt.wantNum {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Array literal evaluation
// ---------------------------------------------------------------------------

func TestEvalArrayLiteral(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "SUM({1,2,3})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 6 {
		t.Errorf("SUM({1,2,3}) = %g, want 6", got.Num)
	}
}

// ---------------------------------------------------------------------------
// IFNA
// ---------------------------------------------------------------------------

func TestEvalIFNA(t *testing.T) {
	resolver := &mockResolver{}

	// #N/A should be caught
	cf := evalCompile(t, `IFNA(#N/A,"fallback")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "fallback" {
		t.Errorf("IFNA(#N/A,...) = %v, want fallback", got)
	}

	// Non-#N/A error should pass through
	cf = evalCompile(t, `IFNA(#DIV/0!,"fallback")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("IFNA(#DIV/0!,...) = %v, want #DIV/0!", got)
	}

	// Normal value passes through
	cf = evalCompile(t, `IFNA(42,"fallback")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("IFNA(42,...) = %v, want 42", got)
	}
}

// ---------------------------------------------------------------------------
// 3D sheet references — must return parse error, not panic
// ---------------------------------------------------------------------------

func TestEval3DSheetRefNoPanic(t *testing.T) {
	// SUM(Sheet2:Sheet5!A11) is a 3D sheet reference (multi-sheet range).
	// The formula engine does not support 3D references, but it must return
	// a graceful error instead of panicking.
	formulas := []string{
		"SUM(Sheet2:Sheet5!A11)",
		"SUM('Sheet2:Sheet5'!A11)",
	}

	for _, f := range formulas {
		t.Run(f, func(t *testing.T) {
			node, err := Parse(f)
			if err != nil {
				// Parse error is the expected graceful failure.
				t.Logf("Parse(%q) returned expected error: %v", f, err)
				return
			}
			// If parsing somehow succeeded, compilation should fail or
			// evaluation should not panic.
			cf, err := Compile(f, node)
			if err != nil {
				t.Logf("Compile(%q) returned expected error: %v", f, err)
				return
			}
			resolver := &mockResolver{}
			_, err = Eval(cf, resolver, nil)
			if err != nil {
				t.Logf("Eval(%q) returned expected error: %v", f, err)
				return
			}
			t.Errorf("expected error for 3D sheet reference %q, but got none", f)
		})
	}
}

// ---------------------------------------------------------------------------
// COUNTBLANK range padding — ensures blank rows beyond MaxRow are counted
// ---------------------------------------------------------------------------

func TestEvalCOUNTBLANKPadding(t *testing.T) {
	// Simulate a sheet where only rows 1 and 3 have data in column A.
	// Rows 2, 4, and 5 are blank (not present in the resolver).
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("hello"),
			{Col: 1, Row: 3}: StringVal("world"),
		},
	}

	// COUNTBLANK(A1:A5): range spans 5 rows, rows 2/4/5 are blank → 3
	cf := evalCompile(t, "COUNTBLANK(A1:A5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(COUNTBLANK(A1:A5)): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COUNTBLANK(A1:A5) = %v (%g), want 3", got.Type, got.Num)
	}

	// COUNTBLANK(A1:A3): range spans 3 rows, only row 2 is blank → 1
	cf = evalCompile(t, "COUNTBLANK(A1:A3)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(COUNTBLANK(A1:A3)): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTBLANK(A1:A3) = %v (%g), want 1", got.Type, got.Num)
	}
}
