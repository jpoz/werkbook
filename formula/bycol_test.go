package formula

import (
	"fmt"
	"testing"
)

func TestBYCOL(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Sheet: "Data", Col: 1, Row: 1}: NumberVal(1),
			{Sheet: "Data", Col: 2, Row: 1}: NumberVal(2),
			{Sheet: "Data", Col: 3, Row: 1}: NumberVal(3),
			{Sheet: "Data", Col: 1, Row: 2}: NumberVal(4),
			{Sheet: "Data", Col: 2, Row: 2}: NumberVal(5),
			{Sheet: "Data", Col: 3, Row: 2}: NumberVal(6),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    Value
	}{
		// 1. SUM each col
		{
			name:    "SUM each col",
			formula: `BYCOL({1,2,3;4,5,6}, LAMBDA(c, SUM(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(5), NumberVal(7), NumberVal(9)},
			}},
		},
		// 2. MAX each col
		{
			name:    "MAX each col",
			formula: `BYCOL({1,5;3,2;7,4}, LAMBDA(c, MAX(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(7), NumberVal(5)},
			}},
		},
		// 3. MIN each col
		{
			name:    "MIN each col",
			formula: `BYCOL({3,1;4,2}, LAMBDA(c, MIN(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3), NumberVal(1)},
			}},
		},
		// 4. AVERAGE each col
		{
			name:    "AVERAGE each col",
			formula: `BYCOL({2,4;6,8}, LAMBDA(c, AVERAGE(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4), NumberVal(6)},
			}},
		},
		// 5. COUNT each col
		{
			name:    "COUNT each col",
			formula: `BYCOL({1,2;3,4;5,6}, LAMBDA(c, COUNT(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3), NumberVal(3)},
			}},
		},
		// 6. Single col array
		{
			name:    "single col array",
			formula: `BYCOL({1;2;3}, LAMBDA(c, SUM(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(6)},
			}},
		},
		// 7. Single row array
		{
			name:    "single row array",
			formula: `BYCOL({1,2,3}, LAMBDA(c, SUM(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2), NumberVal(3)},
			}},
		},
		// 8. PRODUCT per col
		{
			name:    "PRODUCT per col",
			formula: `BYCOL({2,3;4,5}, LAMBDA(c, PRODUCT(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(8), NumberVal(15)},
			}},
		},
		// 9. Cell range
		{
			name:    "cell range",
			formula: `BYCOL(Data!A1:C2, LAMBDA(c, SUM(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(5), NumberVal(7), NumberVal(9)},
			}},
		},
		// 10. IF in body
		{
			name:    "IF in body",
			formula: `BYCOL({1,3;2,4}, LAMBDA(c, IF(SUM(c)>5,"big","small")))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{StringVal("small"), StringVal("big")},
			}},
		},
		// 11. 1x1 input
		{
			name:    "1x1 input",
			formula: `BYCOL({5}, LAMBDA(c, SUM(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(5)},
			}},
		},
		// 12. Arithmetic in body
		{
			name:    "arithmetic in body",
			formula: `BYCOL({1,2;3,4}, LAMBDA(c, SUM(c)*2))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(8), NumberVal(12)},
			}},
		},
		// 13. Boolean result
		{
			name:    "boolean result",
			formula: `BYCOL({1,3;2,4}, LAMBDA(c, SUM(c)>5))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{BoolVal(false), BoolVal(true)},
			}},
		},
		// 14. Error propagation
		{
			name:    "error propagation",
			formula: `BYCOL({1,#N/A;3,4}, LAMBDA(c, SUM(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4), ErrorVal(ErrValNA)},
			}},
		},
		// 15. OOXML prefix
		{
			name:    "OOXML prefix",
			formula: `_XLFN.BYCOL({1,2;3,4}, LAMBDA(c, SUM(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4), NumberVal(6)},
			}},
		},
		// 16. Wrong arg count
		{
			name:    "wrong arg count",
			formula: `BYCOL({1,2})`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 17. Lambda wrong param count
		{
			name:    "lambda wrong param count",
			formula: `BYCOL({1,2}, LAMBDA(a,b, a+b))`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 18. Last arg not lambda
		{
			name:    "last arg not lambda",
			formula: `BYCOL({1,2}, 5)`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 19. COUNTA per col
		{
			name:    "COUNTA per col",
			formula: `BYCOL({"a",1;"",TRUE}, LAMBDA(c, COUNTA(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2), NumberVal(2)},
			}},
		},
		// 20. LARGE per col
		{
			name:    "LARGE per col",
			formula: `BYCOL({3,1;5,4;2,6}, LAMBDA(c, LARGE(c,1)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(5), NumberVal(6)},
			}},
		},
		// 21. LEN of concat
		{
			name:    "LEN of concat",
			formula: `BYCOL({"hi","ab";"wo","cd"}, LAMBDA(c, LEN(CONCAT(c))))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4), NumberVal(4)},
			}},
		},
		// 22. Scalar input
		{
			name:    "scalar input",
			formula: `BYCOL(5, LAMBDA(c, SUM(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(5)},
			}},
		},
		// 23. STDEV per col
		{
			name:    "STDEV per col",
			formula: `BYCOL({1,4;2,5;3,6}, LAMBDA(c, STDEV(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(1)},
			}},
		},
		// 24. Empty body lambda
		{
			name:    "empty body lambda",
			formula: `BYCOL({1,2}, LAMBDA(SUM(c)))`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 25. XLFN LAMBDA prefix
		{
			name:    "XLFN LAMBDA prefix",
			formula: `BYCOL({1,2;3,4}, _XLFN.LAMBDA(c, SUM(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4), NumberVal(6)},
			}},
		},
		// 26. Lambda must return a scalar, not an array
		{
			name:    "array result returns value error",
			formula: `BYCOL({1,2;3,4}, LAMBDA(c, TOCOL(c)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{ErrorVal(ErrValVALUE), ErrorVal(ErrValVALUE)},
			}},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			assertByColValueEqual(t, tt.formula, got, tt.want)
		})
	}
}

// assertByColValueEqual is a test helper that compares two Values deeply.
func assertByColValueEqual(t *testing.T, label string, got, want Value) {
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
				assertByColValueEqual(t, fmt.Sprintf("%s[%d][%d]", label, r, c), got.Array[r][c], want.Array[r][c])
			}
		}
	default:
		t.Fatalf("%s: unexpected value type %v", label, want.Type)
	}
}
