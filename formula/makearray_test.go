package formula

import (
	"fmt"
	"testing"
)

func TestMAKEARRAY(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Sheet: "Data", Col: 1, Row: 1}: NumberVal(10),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    Value
	}{
		// 1. Multiplication table
		{
			name:    "multiplication table",
			formula: `MAKEARRAY(2, 3, LAMBDA(r,c, r*c))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2), NumberVal(3)},
				{NumberVal(2), NumberVal(4), NumberVal(6)},
			}},
		},
		// 2. All ones
		{
			name:    "all ones",
			formula: `MAKEARRAY(2, 2, LAMBDA(r,c, 1))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(1)},
				{NumberVal(1), NumberVal(1)},
			}},
		},
		// 3. Row indices
		{
			name:    "row indices",
			formula: `MAKEARRAY(3, 1, LAMBDA(r,c, r))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1)},
				{NumberVal(2)},
				{NumberVal(3)},
			}},
		},
		// 4. Col indices
		{
			name:    "col indices",
			formula: `MAKEARRAY(1, 3, LAMBDA(r,c, c))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2), NumberVal(3)},
			}},
		},
		// 5. Sum of indices
		{
			name:    "sum of indices",
			formula: `MAKEARRAY(2, 2, LAMBDA(r,c, r+c))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2), NumberVal(3)},
				{NumberVal(3), NumberVal(4)},
			}},
		},
		// 6. Squares
		{
			name:    "squares",
			formula: `MAKEARRAY(3, 1, LAMBDA(r,c, r^2))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1)},
				{NumberVal(4)},
				{NumberVal(9)},
			}},
		},
		// 7. String result
		{
			name:    "string result",
			formula: `MAKEARRAY(2, 2, LAMBDA(r,c, r&","&c))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{StringVal("1,1"), StringVal("1,2")},
				{StringVal("2,1"), StringVal("2,2")},
			}},
		},
		// 8. Boolean (identity matrix booleans)
		{
			name:    "boolean identity",
			formula: `MAKEARRAY(2, 2, LAMBDA(r,c, r=c))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{BoolVal(true), BoolVal(false)},
				{BoolVal(false), BoolVal(true)},
			}},
		},
		// 9. 1x1
		{
			name:    "1x1",
			formula: `MAKEARRAY(1, 1, LAMBDA(r,c, 42))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(42)},
			}},
		},
		// 10. Large 5x5
		{
			name:    "5x5 matrix",
			formula: `MAKEARRAY(5, 5, LAMBDA(r,c, r*10+c))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(11), NumberVal(12), NumberVal(13), NumberVal(14), NumberVal(15)},
				{NumberVal(21), NumberVal(22), NumberVal(23), NumberVal(24), NumberVal(25)},
				{NumberVal(31), NumberVal(32), NumberVal(33), NumberVal(34), NumberVal(35)},
				{NumberVal(41), NumberVal(42), NumberVal(43), NumberVal(44), NumberVal(45)},
				{NumberVal(51), NumberVal(52), NumberVal(53), NumberVal(54), NumberVal(55)},
			}},
		},
		// 11. IF in body
		{
			name:    "IF in body",
			formula: `MAKEARRAY(2, 3, LAMBDA(r,c, IF(r=c,"X","")))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{StringVal("X"), StringVal(""), StringVal("")},
				{StringVal(""), StringVal("X"), StringVal("")},
			}},
		},
		// 12. Power c^r
		{
			name:    "power c^r",
			formula: `MAKEARRAY(2, 3, LAMBDA(r,c, c^r))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2), NumberVal(3)},
				{NumberVal(1), NumberVal(4), NumberVal(9)},
			}},
		},
		// 13. MOD checkerboard
		{
			name:    "MOD checkerboard",
			formula: `MAKEARRAY(3, 3, LAMBDA(r,c, MOD(r+c,2)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(0), NumberVal(1), NumberVal(0)},
				{NumberVal(1), NumberVal(0), NumberVal(1)},
				{NumberVal(0), NumberVal(1), NumberVal(0)},
			}},
		},
		// 14. Negative
		{
			name:    "negative",
			formula: `MAKEARRAY(2, 2, LAMBDA(r,c, -(r*c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(-1), NumberVal(-2)},
				{NumberVal(-2), NumberVal(-4)},
			}},
		},
		// 15. Cell ref in body
		{
			name:    "cell ref in body",
			formula: `MAKEARRAY(2, 1, LAMBDA(r,c, Data!A1*r))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10)},
				{NumberVal(20)},
			}},
		},
		// 16. OOXML prefix
		{
			name:    "OOXML prefix",
			formula: `_XLFN.MAKEARRAY(2, 2, LAMBDA(r,c, r+c))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2), NumberVal(3)},
				{NumberVal(3), NumberVal(4)},
			}},
		},
		// 17. Wrong arg count (2 args)
		{
			name:    "wrong arg count 2",
			formula: `MAKEARRAY(2, LAMBDA(r,c, r+c))`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 18. Wrong arg count (4 args)
		{
			name:    "wrong arg count 4",
			formula: `MAKEARRAY(2, 3, 4, LAMBDA(r,c, r+c))`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 19. Lambda wrong param count (1 param)
		{
			name:    "lambda 1 param",
			formula: `MAKEARRAY(2, 2, LAMBDA(r, r+1))`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 20. Lambda wrong param count (3 params)
		{
			name:    "lambda 3 params",
			formula: `MAKEARRAY(2, 2, LAMBDA(r,c,d, r+c+d))`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 21. Last arg not lambda
		{
			name:    "last arg not lambda",
			formula: `MAKEARRAY(2, 2, 5)`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 22. Rows = 0
		{
			name:    "rows zero",
			formula: `MAKEARRAY(0, 2, LAMBDA(r,c, r+c))`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 23. Cols = 0
		{
			name:    "cols zero",
			formula: `MAKEARRAY(2, 0, LAMBDA(r,c, r+c))`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 24. Rows negative
		{
			name:    "rows negative",
			formula: `MAKEARRAY(-1, 2, LAMBDA(r,c, r+c))`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 25. Rows is string
		{
			name:    "rows string",
			formula: `MAKEARRAY("abc", 2, LAMBDA(r,c, r+c))`,
			want:    ErrorVal(ErrValVALUE),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			assertMakeArrayValueEqual(t, tt.formula, got, tt.want)
		})
	}
}

// assertMakeArrayValueEqual is a test helper that compares two Values deeply.
func assertMakeArrayValueEqual(t *testing.T, label string, got, want Value) {
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
				assertMakeArrayValueEqual(t, fmt.Sprintf("%s[%d][%d]", label, r, c), got.Array[r][c], want.Array[r][c])
			}
		}
	default:
		t.Fatalf("%s: unexpected value type %v", label, want.Type)
	}
}
