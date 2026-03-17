package formula

import (
	"fmt"
	"testing"
)

func TestSCAN(t *testing.T) {
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
		// 1. Running sum
		{
			name:    "running sum",
			formula: `SCAN(0, {1,2,3}, LAMBDA(a,b, a+b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(3), NumberVal(6)}}},
		},
		// 2. Running product
		{
			name:    "running product",
			formula: `SCAN(1, {1,2,3,4}, LAMBDA(a,b, a*b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(6), NumberVal(24)}}},
		},
		// 3. String concat
		{
			name:    "string concatenation",
			formula: `SCAN("", {"a","b","c"}, LAMBDA(a,b, a&b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{StringVal("a"), StringVal("ab"), StringVal("abc")}}},
		},
		// 4. Omitted initial value
		{
			name:    "omitted initial value",
			formula: `SCAN(, {1,2,3}, LAMBDA(a,b, a+b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(3), NumberVal(6)}}},
		},
		// 5. Running max
		{
			name:    "running max",
			formula: `SCAN(0, {3,1,4,1,5}, LAMBDA(a,b, IF(b>a,b,a)))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(3), NumberVal(3), NumberVal(4), NumberVal(4), NumberVal(5)}}},
		},
		// 6. Running min
		{
			name:    "running min",
			formula: `SCAN(999, {5,2,8,1,9}, LAMBDA(a,b, IF(b<a,b,a)))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(5), NumberVal(2), NumberVal(2), NumberVal(1), NumberVal(1)}}},
		},
		// 7. Subtraction
		{
			name:    "subtraction",
			formula: `SCAN(100, {10,20,30}, LAMBDA(a,b, a-b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(90), NumberVal(70), NumberVal(40)}}},
		},
		// 8. Power accumulate
		{
			name:    "power accumulate",
			formula: `SCAN(0, {1,2,3}, LAMBDA(a,b, a+b^2))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(5), NumberVal(14)}}},
		},
		// 9. Single element
		{
			name:    "single element",
			formula: `SCAN(0, {42}, LAMBDA(a,b, a+b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(42)}}},
		},
		// 10. 2D array
		{
			name:    "2D array",
			formula: `SCAN(0, {1,2;3,4}, LAMBDA(a,b, a+b))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(3)},
				{NumberVal(6), NumberVal(10)},
			}},
		},
		// 11. Cell range reference
		{
			name:    "cell range reference",
			formula: `SCAN(0, Data!A1:A3, LAMBDA(a,b, a+b))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10)},
				{NumberVal(30)},
				{NumberVal(60)},
			}},
		},
		// 12. Boolean coercion
		{
			name:    "boolean coercion",
			formula: `SCAN(0, {TRUE,FALSE,TRUE}, LAMBDA(a,b, a+b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(1), NumberVal(2)}}},
		},
		// 13. Conditional accumulation
		{
			name:    "conditional accumulation",
			formula: `SCAN(0, {1,2,3,4,5}, LAMBDA(a,b, IF(b>2,a+b,a)))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(0), NumberVal(0), NumberVal(3), NumberVal(7), NumberVal(12)}}},
		},
		// 14. ABS in lambda
		{
			name:    "ABS in lambda",
			formula: `SCAN(0, {-1,-2,3}, LAMBDA(a,b, a+ABS(b)))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(3), NumberVal(6)}}},
		},
		// 15. Division
		{
			name:    "division",
			formula: `SCAN(120, {2,3,4}, LAMBDA(a,b, a/b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(60), NumberVal(20), NumberVal(5)}}},
		},
		// 16. Division by zero (error propagation)
		{
			name:    "division by zero propagation",
			formula: `SCAN(1, {1,0,3}, LAMBDA(a,b, a/b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), ErrorVal(ErrValDIV0), ErrorVal(ErrValDIV0)}}},
		},
		// 17. Error in array propagation
		{
			name:    "error in array",
			formula: `SCAN(0, {1,#N/A,3}, LAMBDA(a,b, a+b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), ErrorVal(ErrValNA), ErrorVal(ErrValNA)}}},
		},
		// 18. Wrong arg count
		{
			name:    "wrong arg count",
			formula: `IFERROR(SCAN(0, {1,2}), "err")`,
			want:    StringVal("err"),
		},
		// 19. Lambda wrong param count
		{
			name:    "lambda wrong param count",
			formula: `IFERROR(SCAN(0, {1,2,3}, LAMBDA(a, a+1)), "err")`,
			want:    StringVal("err"),
		},
		// 20. Last arg not lambda
		{
			name:    "last arg not lambda",
			formula: `IFERROR(SCAN(0, {1,2,3}, 5), "err")`,
			want:    StringVal("err"),
		},
		// 21. OOXML prefix
		{
			name:    "OOXML prefix",
			formula: `_XLFN.SCAN(0, {1,2,3}, LAMBDA(a,b, a+b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(3), NumberVal(6)}}},
		},
		// 22. Large initial value
		{
			name:    "large initial value",
			formula: `SCAN(1000, {1,2,3}, LAMBDA(a,b, a+b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(1001), NumberVal(1003), NumberVal(1006)}}},
		},
		// 23. Column array
		{
			name:    "column array",
			formula: `SCAN(0, {1;2;3}, LAMBDA(a,b, a+b))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1)},
				{NumberVal(3)},
				{NumberVal(6)},
			}},
		},
		// 24. Nested functions (ROUND)
		{
			name:    "nested ROUND in lambda",
			formula: `SCAN(0, {1.5,2.3,3.7}, LAMBDA(a,b, a+ROUND(b,0)))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(2), NumberVal(4), NumberVal(8)}}},
		},
		// 25. IFERROR in body
		{
			name:    "IFERROR in body",
			formula: `SCAN(0, {1,0,3}, LAMBDA(a,b, IFERROR(a/b,0)+b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(0), NumberVal(3)}}},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			assertScanValueEqual(t, tt.formula, got, tt.want)
		})
	}
}

func TestSCANWithXLFNLambda(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `SCAN(0, {1,2,3}, _XLFN.LAMBDA(a,b, a+b))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	assertScanValueEqual(t, "XLFN lambda", got, Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(3), NumberVal(6)}}})
}

func TestSCANColumnArray(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `SCAN(0, {1;2;3}, LAMBDA(a,b, a+b))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	assertScanValueEqual(t, "column array", got, Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(1)},
		{NumberVal(3)},
		{NumberVal(6)},
	}})
}

func TestSCAN2DArrayOrder(t *testing.T) {
	resolver := &mockResolver{}
	// 2D array {1,2;3,4} should scan row-by-row: 1, 1+2=3, 3+3=6, 6+4=10
	cf := evalCompile(t, `SCAN(0, {1,2;3,4}, LAMBDA(a,b, a+b))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	assertScanValueEqual(t, "2D array order", got, Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(1), NumberVal(3)},
		{NumberVal(6), NumberVal(10)},
	}})
}

func TestSCANSingleElementNoInitial(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `SCAN(, {7}, LAMBDA(a,b, a+b))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	assertScanValueEqual(t, "single element no initial", got, Value{Type: ValueArray, Array: [][]Value{{NumberVal(7)}}})
}

// assertScanValueEqual is a test helper that compares two Values deeply.
func assertScanValueEqual(t *testing.T, label string, got, want Value) {
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
				assertScanValueEqual(t, fmt.Sprintf("%s[%d][%d]", label, r, c), got.Array[r][c], want.Array[r][c])
			}
		}
	default:
		t.Fatalf("%s: unexpected value type %v", label, want.Type)
	}
}
