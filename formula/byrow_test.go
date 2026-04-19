package formula

import (
	"fmt"
	"testing"
)

func TestBYROW(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Sheet: "Data", Col: 1, Row: 1}: NumberVal(1),
			{Sheet: "Data", Col: 2, Row: 1}: NumberVal(2),
			{Sheet: "Data", Col: 3, Row: 1}: NumberVal(3),
			{Sheet: "Data", Col: 1, Row: 2}: NumberVal(4),
			{Sheet: "Data", Col: 2, Row: 2}: NumberVal(5),
			{Sheet: "Data", Col: 3, Row: 2}: NumberVal(6),
			{Sheet: "Data", Col: 1, Row: 3}: NumberVal(7),
			{Sheet: "Data", Col: 2, Row: 3}: NumberVal(8),
			{Sheet: "Data", Col: 3, Row: 3}: NumberVal(9),
		},
	}

	tests := []struct {
		name    string
		formula string
		want    Value
	}{
		// 1. SUM each row
		{
			name:    "SUM each row",
			formula: `BYROW({1,2;3,4;5,6}, LAMBDA(r, SUM(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3)},
				{NumberVal(7)},
				{NumberVal(11)},
			}},
		},
		// 2. MAX each row
		{
			name:    "MAX each row",
			formula: `BYROW({1,5;3,2;7,4}, LAMBDA(r, MAX(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(5)},
				{NumberVal(3)},
				{NumberVal(7)},
			}},
		},
		// 3. MIN each row
		{
			name:    "MIN each row",
			formula: `BYROW({3,1;4,2}, LAMBDA(r, MIN(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1)},
				{NumberVal(2)},
			}},
		},
		// 4. AVERAGE each row
		{
			name:    "AVERAGE each row",
			formula: `BYROW({2,4;6,8}, LAMBDA(r, AVERAGE(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3)},
				{NumberVal(7)},
			}},
		},
		// 5. COUNT each row
		{
			name:    "COUNT each row",
			formula: `BYROW({1,2,3;4,5,6}, LAMBDA(r, COUNT(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3)},
				{NumberVal(3)},
			}},
		},
		// 6. Single row array
		{
			name:    "single row array",
			formula: `BYROW({1,2,3}, LAMBDA(r, SUM(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(6)},
			}},
		},
		// 7. Single column array
		{
			name:    "single column array",
			formula: `BYROW({1;2;3}, LAMBDA(r, SUM(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1)},
				{NumberVal(2)},
				{NumberVal(3)},
			}},
		},
		// 8. Product per row
		{
			name:    "product per row",
			formula: `BYROW({2,3;4,5}, LAMBDA(r, PRODUCT(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(6)},
				{NumberVal(20)},
			}},
		},
		// 9. Cell range
		{
			name:    "cell range",
			formula: `BYROW(Data!A1:C3, LAMBDA(r, SUM(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(6)},
				{NumberVal(15)},
				{NumberVal(24)},
			}},
		},
		// 10. Concatenation
		{
			name:    "concatenation",
			formula: `BYROW({"a","b";"c","d"}, LAMBDA(r, CONCAT(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{StringVal("ab")},
				{StringVal("cd")},
			}},
		},
		// 11. IF in body
		{
			name:    "IF in body",
			formula: `BYROW({1,2;3,4}, LAMBDA(r, IF(SUM(r)>5,"big","small")))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{StringVal("small")},
				{StringVal("big")},
			}},
		},
		// 12. LARGE in body
		{
			name:    "LARGE in body",
			formula: `BYROW({3,1,4;1,5,9}, LAMBDA(r, LARGE(r,1)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(4)},
				{NumberVal(9)},
			}},
		},
		// 13. 1x1 input
		{
			name:    "1x1 input",
			formula: `BYROW({5}, LAMBDA(r, SUM(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(5)},
			}},
		},
		// 14. Arithmetic in body
		{
			name:    "arithmetic in body",
			formula: `BYROW({1,2;3,4}, LAMBDA(r, SUM(r)*2))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(6)},
				{NumberVal(14)},
			}},
		},
		// 15. Boolean result
		{
			name:    "boolean result",
			formula: `BYROW({1,2;3,4}, LAMBDA(r, SUM(r)>5))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{BoolVal(false)},
				{BoolVal(true)},
			}},
		},
		// 16. Error propagation
		{
			name:    "error propagation",
			formula: `BYROW({1,#N/A;3,4}, LAMBDA(r, SUM(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{ErrorVal(ErrValNA)},
				{NumberVal(7)},
			}},
		},
		// 17. OOXML prefix
		{
			name:    "OOXML prefix",
			formula: `_XLFN.BYROW({1,2;3,4}, LAMBDA(r, SUM(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3)},
				{NumberVal(7)},
			}},
		},
		// 18. Wrong arg count
		{
			name:    "wrong arg count",
			formula: `BYROW({1,2})`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 19. Lambda wrong param count
		{
			name:    "lambda wrong param count",
			formula: `BYROW({1,2}, LAMBDA(a,b, a+b))`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 20. Last arg not lambda
		{
			name:    "last arg not lambda",
			formula: `BYROW({1,2}, 5)`,
			want:    ErrorVal(ErrValVALUE),
		},
		// 21. COUNTA per row
		{
			name:    "COUNTA per row",
			formula: `BYROW({"a",1;"",TRUE}, LAMBDA(r, COUNTA(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(2)},
				{NumberVal(2)},
			}},
		},
		// 22. SUM with negative values
		{
			name:    "SUM with negatives",
			formula: `BYROW({-1,2;-3,4}, LAMBDA(r, SUM(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1)},
				{NumberVal(1)},
			}},
		},
		// 23. LEN of concat
		{
			name:    "LEN of concat",
			formula: `BYROW({"hello","world";"a","bc"}, LAMBDA(r, LEN(CONCAT(r))))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10)},
				{NumberVal(3)},
			}},
		},
		// 24. XLFN LAMBDA prefix
		{
			name:    "XLFN LAMBDA prefix",
			formula: `BYROW({1,2;3,4}, _XLFN.LAMBDA(r, SUM(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(3)},
				{NumberVal(7)},
			}},
		},
		// 25. Scalar input (not array)
		{
			name:    "scalar input",
			formula: `BYROW(42, LAMBDA(r, SUM(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(42)},
			}},
		},
		// 26. Lambda must return a scalar, not an array
		{
			name:    "array result returns value error",
			formula: `BYROW({1,2;3,4}, LAMBDA(r, TOROW(r)))`,
			want: Value{Type: ValueArray, Array: [][]Value{
				{ErrorVal(ErrValVALUE)},
				{ErrorVal(ErrValVALUE)},
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
			assertByRowValueEqual(t, tt.formula, got, tt.want)
		})
	}
}

// assertByRowValueEqual is a test helper that compares two Values deeply.
func assertByRowValueEqual(t *testing.T, label string, got, want Value) {
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
				assertByRowValueEqual(t, fmt.Sprintf("%s[%d][%d]", label, r, c), got.Array[r][c], want.Array[r][c])
			}
		}
	default:
		t.Fatalf("%s: unexpected value type %v", label, want.Type)
	}
}
