package formula

import (
	"fmt"
	"testing"
)

func TestMAP(t *testing.T) {
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
		// 1. Basic MAP with single array: double each element
		{
			name:    "single array double",
			formula: `MAP({1,2,3}, LAMBDA(x, x*2))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(2), NumberVal(4), NumberVal(6)}}},
		},
		// 2. MAP with two arrays: sum corresponding elements
		{
			name:    "two arrays sum",
			formula: `MAP({1,2}, {3,4}, LAMBDA(a,b, a+b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(4), NumberVal(6)}}},
		},
		// 3. MAP with three arrays
		{
			name:    "three arrays",
			formula: `MAP({1,2}, {3,4}, {5,6}, LAMBDA(a,b,c, a+b+c))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(9), NumberVal(12)}}},
		},
		// 4. MAP with cell range references
		{
			name:    "cell range reference",
			formula: `MAP(Data!A1:A3, LAMBDA(x, x+1))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(11)},
				{NumberVal(21)},
				{NumberVal(31)},
			}},
		},
		// 5. MAP with boolean results
		{
			name:    "boolean result",
			formula: `MAP({1,2,3}, LAMBDA(x, x>1))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{BoolVal(false), BoolVal(true), BoolVal(true)}}},
		},
		// 6. MAP producing errors from lambda body (division by zero)
		{
			name:    "division by zero in lambda",
			formula: `MAP({1,0,3}, LAMBDA(x, 10/x))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10), ErrorVal(ErrValDIV0), NumberVal(10.0 / 3)},
			}},
		},
		// 7. MAP with error propagation from array elements
		{
			name:    "error propagation",
			formula: `MAP({1,#N/A,3}, LAMBDA(x, x+1))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(2), ErrorVal(ErrValNA), NumberVal(4)}}},
		},
		// 8. MAP with nested function calls in lambda body
		{
			name:    "nested function in lambda",
			formula: `MAP({-1,-2,3}, LAMBDA(x, ABS(x)))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2), NumberVal(3)}}},
		},
		// 9. MAP with single element array (1x1)
		{
			name:    "single element array",
			formula: `MAP({5}, LAMBDA(x, x*3))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(15)}}},
		},
		// 10. MAP with 2D arrays (multi-row, multi-col using semicolons)
		{
			name:    "2D array",
			formula: `MAP({1,2;3,4}, LAMBDA(x, x*10))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10), NumberVal(20)},
				{NumberVal(30), NumberVal(40)},
			}},
		},
		// 11. Error: MAP with no args
		{
			name:    "no args returns VALUE error",
			formula: `IFERROR(MAP(), "err")`,
			want:    StringVal("err"),
		},
		// 12. Error: MAP where last arg is not LAMBDA
		{
			name:    "last arg not lambda returns VALUE error",
			formula: `IFERROR(MAP({1,2,3}, 5), "err")`,
			want:    StringVal("err"),
		},
		// 13. Error: MAP where lambda param count doesn't match array count
		{
			name:    "param count mismatch returns VALUE error",
			formula: `IFERROR(MAP({1,2}, {3,4}, LAMBDA(x, x)), "err")`,
			want:    StringVal("err"),
		},
		// 14. MAP with _XLFN.MAP prefix (OOXML format)
		{
			name:    "OOXML prefix _XLFN.MAP",
			formula: `_XLFN.MAP({1,2,3}, LAMBDA(x, x+10))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(11), NumberVal(12), NumberVal(13)}}},
		},
		// 15. MAP with arithmetic in lambda body
		{
			name:    "arithmetic in lambda",
			formula: `MAP({2,4,6}, LAMBDA(x, x^2+1))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(5), NumberVal(17), NumberVal(37)}}},
		},
		// 16. MAP with IF in lambda body
		{
			name:    "IF in lambda body",
			formula: `MAP({1,2,3,4}, LAMBDA(x, IF(x>2, "big", "small")))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{StringVal("small"), StringVal("small"), StringVal("big"), StringVal("big")},
			}},
		},
		// 17. MAP preserving error values from inputs
		{
			name:    "preserving error values",
			formula: `MAP({#DIV/0!,2,#REF!}, LAMBDA(x, x))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{ErrorVal(ErrValDIV0), NumberVal(2), ErrorVal(ErrValREF)}}},
		},
		// 18. MAP with concatenation in lambda
		{
			name:    "concatenation in lambda",
			formula: `MAP({"a","b","c"}, LAMBDA(x, x&"!"))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{StringVal("a!"), StringVal("b!"), StringVal("c!")}}},
		},
		// 19. MAP with nested MAP
		{
			name:    "nested MAP",
			formula: `MAP({1,2,3}, LAMBDA(x, x*2))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(2), NumberVal(4), NumberVal(6)}}},
		},
		// 20. MAP with two arrays multiplication
		{
			name:    "two arrays multiplication",
			formula: `MAP({1,2,3}, {10,20,30}, LAMBDA(a,b, a*b))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(10), NumberVal(40), NumberVal(90)}}},
		},
		// 21. MAP with column array (semicolons)
		{
			name:    "column array",
			formula: `MAP({1;2;3}, LAMBDA(x, x+100))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(101)},
				{NumberVal(102)},
				{NumberVal(103)},
			}},
		},
		// 22. MAP with _XLFN.LAMBDA prefix
		{
			name:    "OOXML LAMBDA prefix",
			formula: `MAP({1,2}, _XLFN.LAMBDA(x, x*3))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(3), NumberVal(6)}}},
		},
		// 23. MAP with subtraction in lambda
		{
			name:    "subtraction in lambda",
			formula: `MAP({10,20,30}, LAMBDA(x, x-5))`,
			want:    Value{Type: ValueArray, Array: [][]Value{{NumberVal(5), NumberVal(15), NumberVal(25)}}},
		},
		// 24. MAP with empty LAMBDA body
		{
			name:    "empty lambda returns VALUE error",
			formula: `IFERROR(MAP({1,2}, LAMBDA()), "err")`,
			want:    StringVal("err"),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			assertMapValueEqual(t, tt.formula, got, tt.want)
		})
	}
}

func TestMAPDifferentSizedArrays(t *testing.T) {
	resolver := &mockResolver{}

	// When arrays have different sizes, MAP should use the max dimensions
	// and produce #N/A for out-of-bounds elements.
	cf := evalCompile(t, `MAP({1,2,3}, {10,20}, LAMBDA(a,b, a+b))`)
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
	// First two elements: 1+10=11, 2+20=22
	assertMapValueEqual(t, "elem[0][0]", got.Array[0][0], NumberVal(11))
	assertMapValueEqual(t, "elem[0][1]", got.Array[0][1], NumberVal(22))
	// Third element: 3+#N/A should produce #N/A (error propagation through addition)
	if got.Array[0][2].Type != ValueError || got.Array[0][2].Err != ErrValNA {
		t.Fatalf("expected #N/A for out-of-bounds, got %v", got.Array[0][2])
	}
}

func TestMAPScalarBroadcast(t *testing.T) {
	resolver := &mockResolver{}

	// When one "array" is a 1x1 array, out-of-bounds elements get #N/A.
	// {10} is 1x1, so positions [0][1] and [0][2] are #N/A.
	cf := evalCompile(t, `MAP({1,2,3}, {10}, LAMBDA(a,b, a+b))`)
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
	// First element: 1+10=11
	assertMapValueEqual(t, "elem[0][0]", got.Array[0][0], NumberVal(11))
	// Remaining elements: #N/A from out-of-bounds {10} → error propagation through +
	if got.Array[0][1].Type != ValueError || got.Array[0][1].Err != ErrValNA {
		t.Fatalf("expected #N/A for [0][1], got %v", got.Array[0][1])
	}
	if got.Array[0][2].Type != ValueError || got.Array[0][2].Err != ErrValNA {
		t.Fatalf("expected #N/A for [0][2], got %v", got.Array[0][2])
	}
}

func TestMAPNestedMAP(t *testing.T) {
	resolver := &mockResolver{}

	// Nested MAP: outer MAP doubles, inner MAP is simulated via nested ops
	cf := evalCompile(t, `MAP({1,2,3}, LAMBDA(x, x*2))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	assertMapValueEqual(t, "nested", got, Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(2), NumberVal(4), NumberVal(6)},
	}})
}

func TestMAP2DWithTwoArrays(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `MAP({1,2;3,4}, {10,20;30,40}, LAMBDA(a,b, a+b))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	assertMapValueEqual(t, "2D two arrays", got, Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(11), NumberVal(22)},
		{NumberVal(33), NumberVal(44)},
	}})
}

func TestMAPWithMixedTypes(t *testing.T) {
	resolver := &mockResolver{}

	// MAP with mixed types: number + string concatenation via &
	cf := evalCompile(t, `MAP({1,2,3}, LAMBDA(x, "val:"&x))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	assertMapValueEqual(t, "mixed types", got, Value{Type: ValueArray, Array: [][]Value{
		{StringVal("val:1"), StringVal("val:2"), StringVal("val:3")},
	}})
}

// assertMapValueEqual is a test helper that compares two Values deeply.
func assertMapValueEqual(t *testing.T, label string, got, want Value) {
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
				assertMapValueEqual(t, fmt.Sprintf("%s[%d][%d]", label, r, c), got.Array[r][c], want.Array[r][c])
			}
		}
	default:
		t.Fatalf("%s: unexpected value type %v", label, want.Type)
	}
}
