package formula

import (
	"testing"
)

func TestREDUCE(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Sheet: "Data", Col: 1, Row: 1}: NumberVal(10),
			{Sheet: "Data", Col: 1, Row: 2}: NumberVal(20),
			{Sheet: "Data", Col: 1, Row: 3}: NumberVal(30),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    Value
	}{
		// 1. Basic sum
		{
			name:    "basic sum",
			formula: `REDUCE(0, {1,2,3}, LAMBDA(a,b, a+b))`,
			want:    NumberVal(6),
		},
		// 2. Product
		{
			name:    "product",
			formula: `REDUCE(1, {2,3,4}, LAMBDA(a,b, a*b))`,
			want:    NumberVal(24),
		},
		// 3. String concatenation
		{
			name:    "string concatenation",
			formula: `REDUCE("", {"a","b","c"}, LAMBDA(a,b, a&b))`,
			want:    StringVal("abc"),
		},
		// 4. Omitted initial value (sum)
		{
			name:    "omitted initial value sum",
			formula: `REDUCE(, {1,2,3}, LAMBDA(a,b, a+b))`,
			want:    NumberVal(6),
		},
		// 5. Count even with IF+ISEVEN
		{
			name:    "count even",
			formula: `REDUCE(0, {1,2,3,4,5,6}, LAMBDA(a,n, IF(ISEVEN(n), a+1, a)))`,
			want:    NumberVal(3),
		},
		// 6. Product with IF
		{
			name:    "product with IF",
			formula: `REDUCE(1, {10,60,30,70}, LAMBDA(a,b, IF(b>50, a*b, a)))`,
			want:    NumberVal(4200),
		},
		// 7. Max via reduce
		{
			name:    "max via reduce",
			formula: `REDUCE(0, {5,2,8,1,9}, LAMBDA(a,b, IF(b>a, b, a)))`,
			want:    NumberVal(9),
		},
		// 8. Single element array
		{
			name:    "single element array",
			formula: `REDUCE(0, {42}, LAMBDA(a,b, a+b))`,
			want:    NumberVal(42),
		},
		// 9. Single element no initial
		{
			name:    "single element no initial",
			formula: `REDUCE(, {42}, LAMBDA(a,b, a+b))`,
			want:    NumberVal(42),
		},
		// 10. 2D array flattened
		{
			name:    "2D array flattened",
			formula: `REDUCE(0, {1,2;3,4}, LAMBDA(a,b, a+b))`,
			want:    NumberVal(10),
		},
		// 11. Cell range reference
		{
			name:    "cell range reference",
			formula: `REDUCE(0, Data!A1:A3, LAMBDA(a,b, a+b))`,
			want:    NumberVal(60),
		},
		// 12. Division in lambda
		{
			name:    "division in lambda",
			formula: `REDUCE(120, {2,3,4}, LAMBDA(a,b, a/b))`,
			want:    NumberVal(5),
		},
		// 13. Subtraction
		{
			name:    "subtraction",
			formula: `REDUCE(100, {10,20,30}, LAMBDA(a,b, a-b))`,
			want:    NumberVal(40),
		},
		// 14. Boolean coercion
		{
			name:    "boolean coercion",
			formula: `REDUCE(0, {TRUE,FALSE,TRUE}, LAMBDA(a,b, a+b))`,
			want:    NumberVal(2),
		},
		// 15. Error propagation from array
		{
			name:    "error propagation from array",
			formula: `REDUCE(0, {1,#N/A,3}, LAMBDA(a,b, a+b))`,
			want:    ErrorVal(ErrValNA),
		},
		// 16. Division by zero in lambda
		{
			name:    "division by zero in lambda",
			formula: `REDUCE(1, {1,0,3}, LAMBDA(a,b, a/b))`,
			want:    ErrorVal(ErrValDIV0),
		},
		// 17. Wrong arg count (not 3) - wrapped in IFERROR since it becomes #VALUE! at parse time
		{
			name:    "wrong arg count",
			formula: `IFERROR(REDUCE(0, {1,2}), "err")`,
			want:    StringVal("err"),
		},
		// 18. Lambda with wrong param count
		{
			name:    "lambda wrong param count",
			formula: `IFERROR(REDUCE(0, {1,2,3}, LAMBDA(a, a+1)), "err")`,
			want:    StringVal("err"),
		},
		// 19. Last arg not lambda
		{
			name:    "last arg not lambda",
			formula: `IFERROR(REDUCE(0, {1,2,3}, 5), "err")`,
			want:    StringVal("err"),
		},
		// 20. OOXML prefix
		{
			name:    "OOXML prefix",
			formula: `_XLFN.REDUCE(0, {1,2,3}, LAMBDA(a,b, a+b))`,
			want:    NumberVal(6),
		},
		// 21. Nested function in lambda (ABS)
		{
			name:    "nested function in lambda",
			formula: `REDUCE(0, {-1,-2,3}, LAMBDA(a,b, a+ABS(b)))`,
			want:    NumberVal(6),
		},
		// 22. Power in lambda
		{
			name:    "power in lambda",
			formula: `REDUCE(0, {1,2,3}, LAMBDA(a,b, a+b^2))`,
			want:    NumberVal(14),
		},
		// 23. Conditional sum
		{
			name:    "conditional sum",
			formula: `REDUCE(0, {1,2,3,4,5}, LAMBDA(a,b, IF(b>2, a+b, a)))`,
			want:    NumberVal(12),
		},
		// 24. Large initial value
		{
			name:    "large initial value",
			formula: `REDUCE(1000000, {1,2,3}, LAMBDA(a,b, a+b))`,
			want:    NumberVal(1000006),
		},
		// 25. Empty initial with single element no fold
		{
			name:    "empty initial single element",
			formula: `REDUCE(, {7}, LAMBDA(a,b, a*b))`,
			want:    NumberVal(7),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			assertReduceValueEqual(t, tt.formula, got, tt.want)
		})
	}
}

func TestREDUCEEmptyArrayNoInitial(t *testing.T) {
	// Empty array with no initial value should return #CALC!
	// We can't express an empty literal array in Excel syntax, but we
	// can test a single-element array with omitted initial to verify
	// the accumulator logic works correctly.
	resolver := &mockResolver{}
	cf := evalCompile(t, `REDUCE(, {7}, LAMBDA(a,b, a+b))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// With omitted initial, first element (7) becomes acc, no more elements to iterate
	assertReduceValueEqual(t, "empty initial single", got, NumberVal(7))
}

func TestREDUCEWithXLFNLambda(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `REDUCE(0, {1,2,3}, _XLFN.LAMBDA(a,b, a+b))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	assertReduceValueEqual(t, "XLFN lambda", got, NumberVal(6))
}

func TestREDUCEColumnArray(t *testing.T) {
	resolver := &mockResolver{}
	// Column array using semicolons: {1;2;3} is a 3x1 array
	cf := evalCompile(t, `REDUCE(0, {1;2;3}, LAMBDA(a,b, a+b))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	assertReduceValueEqual(t, "column array", got, NumberVal(6))
}

func TestREDUCE2DArrayOrder(t *testing.T) {
	resolver := &mockResolver{}
	// 2D array {1,2;3,4} should flatten row-by-row: 1, 2, 3, 4
	// Concatenating should produce "1234"
	cf := evalCompile(t, `REDUCE("", {1,2;3,4}, LAMBDA(a,b, a&b))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	assertReduceValueEqual(t, "2D array order", got, StringVal("1234"))
}

// assertReduceValueEqual is a test helper that compares two Values deeply.
func assertReduceValueEqual(t *testing.T, label string, got, want Value) {
	t.Helper()

	if got.Type != want.Type {
		t.Fatalf("%s: type mismatch: got %v, want %v (got=%v want=%v)", label, got.Type, want.Type, got, want)
	}

	switch want.Type {
	case ValueEmpty:
		return
	case ValueNumber:
		if got.Num != want.Num {
			t.Fatalf("%s: number mismatch: got %g, want %g", label, got.Num, want.Num)
		}
	case ValueString:
		if got.Str != want.Str {
			t.Fatalf("%s: string mismatch: got %q, want %q", label, got.Str, want.Str)
		}
	case ValueBool:
		if got.Bool != want.Bool {
			t.Fatalf("%s: bool mismatch: got %v, want %v", label, got.Bool, want.Bool)
		}
	case ValueError:
		if got.Err != want.Err {
			t.Fatalf("%s: error mismatch: got %v, want %v", label, got.Err, want.Err)
		}
	case ValueArray:
		if len(got.Array) != len(want.Array) {
			t.Fatalf("%s: row count mismatch: got %d, want %d", label, len(got.Array), len(want.Array))
		}
		for r := range want.Array {
			if len(got.Array[r]) != len(want.Array[r]) {
				t.Fatalf("%s: col count mismatch at row %d: got %d, want %d", label, r, len(got.Array[r]), len(want.Array[r]))
			}
			for c := range want.Array[r] {
				assertReduceValueEqual(t, label, got.Array[r][c], want.Array[r][c])
			}
		}
	default:
		t.Fatalf("%s: unexpected value type %v", label, want.Type)
	}
}
