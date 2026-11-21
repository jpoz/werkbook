package formula

import (
	"reflect"
	"testing"
)

func TestSORT_TrimmedRangeOrigin(t *testing.T) {
	got, err := fnSORT([]Value{
		trimmedRangeValue([][]Value{
			{NumberVal(2)},
			{NumberVal(1)},
		}, 1, 1, 1, 3),
		NumberVal(1),
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnSORT: %v", err)
	}
	assertLookupValueEqual(t, got, Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(2)},
		{NumberVal(1)},
		{EmptyVal()},
	}})
}

func TestSORTBY_TrimmedRangeOrigin(t *testing.T) {
	got, err := fnSORTBY([]Value{
		trimmedRangeValue([][]Value{
			{StringVal("b")},
			{StringVal("a")},
		}, 1, 1, 1, 3),
		trimmedRangeValue([][]Value{
			{NumberVal(1)},
			{NumberVal(2)},
		}, 2, 1, 2, 3),
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnSORTBY: %v", err)
	}
	assertLookupValueEqual(t, got, Value{Type: ValueArray, Array: [][]Value{
		{StringVal("a")},
		{StringVal("b")},
		{EmptyVal()},
	}})
}

func TestSORTBY_TrimmedFullColumnRange(t *testing.T) {
	got, err := fnSORTBY([]Value{
		trimmedRangeValue([][]Value{
			{StringVal("b")},
			{StringVal("a")},
		}, 1, 1, 1, maxRows),
		trimmedRangeValue([][]Value{
			{NumberVal(1)},
			{NumberVal(2)},
		}, 2, 1, 2, maxRows),
		NumberVal(-1),
	})
	if err != nil {
		t.Fatalf("fnSORTBY: %v", err)
	}
	assertLookupValueEqual(t, got, Value{Type: ValueArray, Array: [][]Value{
		{StringVal("a")},
		{StringVal("b")},
	}})
}

func TestSORTBY_IndexRowsColumnsCountaViaEval(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 1, Row: 3}: NumberVal(1),
			{Col: 1, Row: 4}: NumberVal(4),
			{Col: 1, Row: 5}: NumberVal(2),
			{Col: 2, Row: 1}: NumberVal(50),
			{Col: 2, Row: 2}: NumberVal(30),
			{Col: 2, Row: 3}: NumberVal(10),
			{Col: 2, Row: 4}: NumberVal(40),
			{Col: 2, Row: 5}: NumberVal(20),
		},
	}

	tests := []struct {
		formula string
		want    Value
	}{
		{formula: "INDEX(SORTBY(A1:A5,B1:B5),4,1)", want: NumberVal(4)},
		{formula: "ROWS(SORTBY(A1:A5,B1:B5))", want: NumberVal(5)},
		{formula: "COLUMNS(SORTBY(A1:A5,B1:B5))", want: NumberVal(1)},
		{formula: "COUNTA(SORTBY(A1:A5,B1:B5))", want: NumberVal(5)},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			got, err := Eval(evalCompile(t, tt.formula), resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			assertLookupValueEqual(t, got, tt.want)
		})
	}
}

func TestTRUE(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("value tests", func(t *testing.T) {
		tests := []struct {
			formula string
			want    Value
		}{
			// Basic: TRUE() returns boolean TRUE
			{"TRUE()", BoolVal(true)},
			// Numeric coercion: TRUE()+0 = 1
			{"TRUE()+0", NumberVal(1)},
			// Arithmetic with TRUE: TRUE()*5 = 5
			{"TRUE()*5", NumberVal(5)},
			// TRUE used in IF condition
			{"IF(TRUE(),\"yes\",\"no\")", StringVal("yes")},
			// TRUE combined with AND
			{"AND(TRUE(),TRUE())", BoolVal(true)},
			// NOT(TRUE()) = FALSE
			{"NOT(TRUE())", BoolVal(false)},
			// TRUE in OR
			{"OR(TRUE(),FALSE())", BoolVal(true)},
			// XOR with TRUE
			{"XOR(TRUE(),FALSE())", BoolVal(true)},
			// TRUE + TRUE = 2 (numeric coercion in arithmetic)
			{"TRUE()+TRUE()", NumberVal(2)},
		}

		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if !reflect.DeepEqual(got, tt.want) {
					t.Fatalf("Eval(%q) = %#v, want %#v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("error: argument provided", func(t *testing.T) {
		cf := evalCompile(t, "TRUE(1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(TRUE(1)): %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Fatalf("Eval(TRUE(1)) = %#v, want #VALUE!", got)
		}
	})
}

func TestFALSE(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("value tests", func(t *testing.T) {
		tests := []struct {
			formula string
			want    Value
		}{
			// Basic: FALSE() returns boolean FALSE
			{"FALSE()", BoolVal(false)},
			// Numeric coercion: FALSE()+0 = 0
			{"FALSE()+0", NumberVal(0)},
			// Arithmetic with FALSE: FALSE()*5 = 0
			{"FALSE()*5", NumberVal(0)},
			// FALSE used in IF condition
			{"IF(FALSE(),\"yes\",\"no\")", StringVal("no")},
			// AND with FALSE
			{"AND(TRUE(),FALSE())", BoolVal(false)},
			// OR with all FALSE
			{"OR(FALSE(),FALSE())", BoolVal(false)},
			// NOT(FALSE()) = TRUE
			{"NOT(FALSE())", BoolVal(true)},
			// XOR with two FALSE values
			{"XOR(FALSE(),FALSE())", BoolVal(false)},
			// FALSE + FALSE = 0 (numeric coercion in arithmetic)
			{"FALSE()+FALSE()", NumberVal(0)},
			// TRUE + FALSE = 1
			{"TRUE()+FALSE()", NumberVal(1)},
		}

		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if !reflect.DeepEqual(got, tt.want) {
					t.Fatalf("Eval(%q) = %#v, want %#v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("error: argument provided", func(t *testing.T) {
		cf := evalCompile(t, "FALSE(1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(FALSE(1)): %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Fatalf("Eval(FALSE(1)) = %#v, want #VALUE!", got)
		}
	})
}

func TestAND(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("boolean results", func(t *testing.T) {
		tests := []struct {
			formula string
			want    bool
		}{
			// Original tests
			{"AND(TRUE,TRUE)", true},
			{"AND(TRUE,FALSE)", false},
			{"AND(FALSE,FALSE)", false},
			{"AND(TRUE,TRUE,TRUE)", true},
			{"AND(TRUE,TRUE,FALSE)", false},

			// Single argument
			{"AND(TRUE)", true},
			{"AND(FALSE)", false},

			// One FALSE among many TRUE
			{"AND(TRUE,TRUE,TRUE,TRUE,FALSE)", false},

			// All TRUE with many args
			{"AND(TRUE,TRUE,TRUE,TRUE,TRUE)", true},

			// Numbers: non-zero is TRUE, zero is FALSE
			{"AND(1)", true},
			{"AND(0)", false},
			{"AND(1,2,3)", true},
			{"AND(1,0,3)", false},
			{"AND(-1)", true},
			{"AND(0.5)", true},

			// Mixed booleans and numbers
			{"AND(TRUE,1)", true},
			{"AND(TRUE,0)", false},
			{"AND(FALSE,1)", false},
			{"AND(1,TRUE,2,TRUE)", true},

			// Doc example: AND(A2>1,A2<100) style comparisons
			{"AND(50>1,50<100)", true},
			{"AND(150>1,150<100)", false},
		}

		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want bool %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("no args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, "AND()")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("AND() = %v, want #VALUE! error", got)
		}
	})

	t.Run("direct string arg returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `AND("hello")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`AND("hello") = %v, want #VALUE! error`, got)
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, "AND(TRUE,1/0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("AND(TRUE,1/0) = %v, want error", got)
		}
	})

	t.Run("range with all TRUE values", func(t *testing.T) {
		// Build an array of all TRUE values
		arr := Value{
			Type: ValueArray,
			Array: [][]Value{
				{BoolVal(true), BoolVal(true)},
				{BoolVal(true), BoolVal(true)},
			},
		}
		got, err := fnAND([]Value{arr})
		if err != nil {
			t.Fatalf("fnAND: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("AND(all-true-range) = %v, want true", got)
		}
	})

	t.Run("range with one FALSE", func(t *testing.T) {
		arr := Value{
			Type: ValueArray,
			Array: [][]Value{
				{BoolVal(true), BoolVal(true)},
				{BoolVal(true), BoolVal(false)},
			},
		}
		got, err := fnAND([]Value{arr})
		if err != nil {
			t.Fatalf("fnAND: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("AND(range-with-false) = %v, want false", got)
		}
	})

	t.Run("strings in range are ignored", func(t *testing.T) {
		arr := Value{
			Type: ValueArray,
			Array: [][]Value{
				{BoolVal(true), StringVal("hello")},
				{BoolVal(true), BoolVal(true)},
			},
		}
		got, err := fnAND([]Value{arr})
		if err != nil {
			t.Fatalf("fnAND: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("AND(range-with-string) = %v, want true (strings ignored)", got)
		}
	})

	t.Run("error in range propagates", func(t *testing.T) {
		arr := Value{
			Type: ValueArray,
			Array: [][]Value{
				{BoolVal(true), ErrorVal(ErrValDIV0)},
			},
		}
		got, err := fnAND([]Value{arr})
		if err != nil {
			t.Fatalf("fnAND: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("AND(range-with-error) = %v, want #DIV/0! error", got)
		}
	})

	t.Run("range with numbers", func(t *testing.T) {
		arr := Value{
			Type: ValueArray,
			Array: [][]Value{
				{NumberVal(1), NumberVal(5)},
				{NumberVal(3), NumberVal(0)},
			},
		}
		got, err := fnAND([]Value{arr})
		if err != nil {
			t.Fatalf("fnAND: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("AND(range-with-zero) = %v, want false", got)
		}
	})
}

func TestOR(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("basic boolean args", func(t *testing.T) {
		tests := []struct {
			formula string
			want    bool
		}{
			{"OR(TRUE)", true},
			{"OR(FALSE)", false},
			{"OR(TRUE,FALSE)", true},
			{"OR(FALSE,FALSE)", false},
			{"OR(TRUE,TRUE)", true},
			{"OR(FALSE,TRUE)", true},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("multiple args", func(t *testing.T) {
		tests := []struct {
			formula string
			want    bool
		}{
			{"OR(FALSE,FALSE,FALSE)", false},
			{"OR(FALSE,FALSE,TRUE)", true},
			{"OR(FALSE,TRUE,FALSE)", true},
			{"OR(TRUE,FALSE,FALSE)", true},
			{"OR(TRUE,TRUE,TRUE)", true},
			{"OR(FALSE,FALSE,FALSE,FALSE,TRUE)", true},
			{"OR(FALSE,FALSE,FALSE,FALSE,FALSE)", false},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("numeric args", func(t *testing.T) {
		tests := []struct {
			formula string
			want    bool
		}{
			{"OR(1)", true},
			{"OR(0)", false},
			{"OR(0,0)", false},
			{"OR(0,1)", true},
			{"OR(1,0)", true},
			{"OR(1,1)", true},
			{"OR(-1)", true},
			{"OR(0.5)", true},
			{"OR(0,0,0,0.001)", true},
			{"OR(99)", true},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("mixed boolean and numeric", func(t *testing.T) {
		tests := []struct {
			formula string
			want    bool
		}{
			{"OR(FALSE,0)", false},
			{"OR(TRUE,0)", true},
			{"OR(FALSE,1)", true},
			{"OR(0,FALSE)", false},
			{"OR(1,FALSE)", true},
			{"OR(0,TRUE)", true},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("direct string args return VALUE error", func(t *testing.T) {
		// Direct string arguments to OR cause #VALUE!
		tests := []string{
			`OR("hello")`,
			`OR("TRUE")`,
			`OR("FALSE")`,
			`OR("")`,
			`OR("text",FALSE)`,
			`OR("1")`,
			`OR("0")`,
		}
		for _, formula := range tests {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", formula, err)
				}
				if got.Type != ValueError || got.Err != ErrValVALUE {
					t.Errorf("Eval(%q) = %v, want #VALUE!", formula, got)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		tests := []struct {
			formula string
			wantErr ErrorValue
		}{
			{"OR(1/0)", ErrValDIV0},
			{"OR(1/0,TRUE)", ErrValDIV0},
			{"OR(FALSE,1/0)", ErrValDIV0},
			{"OR(1/0,FALSE)", ErrValDIV0},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected Go error: %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %d", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("short circuits past error when true found first", func(t *testing.T) {
		cf := evalCompile(t, "OR(TRUE,1/0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("OR(TRUE,1/0) = %v, want true", got)
		}
	})

	t.Run("with expressions", func(t *testing.T) {
		tests := []struct {
			formula string
			want    bool
		}{
			{"OR(1>2)", false},
			{"OR(1<2)", true},
			{"OR(1>2,3>4)", false},
			{"OR(1>2,3<4)", true},
			{"OR(1=1,2=3)", true},
			{"OR(1=2,2=3)", false},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("range with cell references", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(false),
				{Col: 1, Row: 2}: BoolVal(false),
				{Col: 1, Row: 3}: BoolVal(true),
			},
		}
		cf := evalCompile(t, "OR(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("OR(A1:A3) = %v, want true", got)
		}
	})

	t.Run("range all false", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(false),
				{Col: 1, Row: 2}: BoolVal(false),
				{Col: 1, Row: 3}: BoolVal(false),
			},
		}
		cf := evalCompile(t, "OR(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("OR(A1:A3) = %v, want false", got)
		}
	})

	t.Run("range with error propagates", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(false),
				{Col: 1, Row: 2}: ErrorVal(ErrValNA),
				{Col: 1, Row: 3}: BoolVal(true),
			},
		}
		cf := evalCompile(t, "OR(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("OR(A1:A3) = %v, want #N/A error", got)
		}
	})

	t.Run("short circuits on first true in range", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
				{Col: 1, Row: 2}: BoolVal(false),
			},
		}
		cf := evalCompile(t, "OR(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("OR(A1:A2) = %v, want true", got)
		}
	})

	t.Run("range with numbers", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(0),
				{Col: 1, Row: 2}: NumberVal(0),
				{Col: 1, Row: 3}: NumberVal(5),
			},
		}
		cf := evalCompile(t, "OR(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("OR(A1:A3) = %v, want true", got)
		}
	})
}

func TestANDStringArgs(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("direct string args return VALUE error", func(t *testing.T) {
		tests := []string{
			`AND("text")`,
			`AND("1")`,
			`AND("TRUE")`,
		}
		for _, formula := range tests {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", formula, err)
				}
				if got.Type != ValueError || got.Err != ErrValVALUE {
					t.Errorf("Eval(%q) = %v, want #VALUE!", formula, got)
				}
			})
		}
	})

	t.Run("string in range is skipped", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("text"),
				{Col: 1, Row: 2}: BoolVal(true),
			},
		}
		cf := evalCompile(t, "AND(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("AND(A1:A2) = %v, want true", got)
		}
	})
}

func TestXOR(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("basic boolean results", func(t *testing.T) {
		tests := []struct {
			formula string
			want    bool
		}{
			// Single args
			{"XOR(TRUE)", true},
			{"XOR(FALSE)", false},
			// Two booleans
			{"XOR(TRUE,TRUE)", false},   // even TRUE count
			{"XOR(TRUE,FALSE)", true},   // odd TRUE count
			{"XOR(FALSE,FALSE)", false}, // zero TRUE count
			// Three booleans
			{"XOR(TRUE,TRUE,TRUE)", true},     // 3 TRUE = odd
			{"XOR(TRUE,TRUE,FALSE)", false},   // 2 TRUE = even
			{"XOR(TRUE,FALSE,FALSE)", true},   // 1 TRUE = odd
			{"XOR(FALSE,FALSE,FALSE)", false}, // 0 TRUE
			// Four TRUE (even count) → FALSE
			{"XOR(TRUE,TRUE,TRUE,TRUE)", false},
			// Numbers: non-zero = TRUE, zero = FALSE
			{"XOR(1)", true},
			{"XOR(0)", false},
			{"XOR(1,0)", true},    // 1 TRUE
			{"XOR(1,1)", false},   // 2 TRUE = even
			{"XOR(5,0,0)", true},  // 1 TRUE
			{"XOR(5,3,0)", false}, // 2 TRUE = even
			// Mixed booleans and numbers
			{"XOR(TRUE,1)", false},  // 2 TRUE = even
			{"XOR(TRUE,0)", true},   // 1 TRUE
			{"XOR(FALSE,1)", true},  // 1 TRUE
			{"XOR(FALSE,0)", false}, // 0 TRUE
			// Doc examples: =XOR(3>0,2<9) → FALSE (both TRUE, even)
			{"XOR(3>0,2<9)", false},
			// =XOR(3>12,4>6) → FALSE (both FALSE)
			{"XOR(3>12,4>6)", false},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("no args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, "XOR()")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("XOR() = %v, want #VALUE!", got)
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, "XOR(1/0,TRUE)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("XOR(1/0,TRUE) = %v, want error", got)
		}
	})

	t.Run("range with mixed values", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
				{Col: 1, Row: 2}: BoolVal(false),
				{Col: 1, Row: 3}: NumberVal(1),
			},
		}
		// TRUE + FALSE + 1(TRUE) = 2 TRUE = even → FALSE
		cf := evalCompile(t, "XOR(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("XOR(A1:A3) = %v, want false", got)
		}
	})

	t.Run("error in range propagates", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
				{Col: 1, Row: 2}: ErrorVal(ErrValDIV0),
			},
		}
		cf := evalCompile(t, "XOR(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("XOR(A1:A2) = %v, want #DIV/0!", got)
		}
	})
}

func TestXORStringArgs(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("direct string args are skipped", func(t *testing.T) {
		// XOR skips strings (both direct and from ranges), unlike AND/OR which error.
		tests := []struct {
			formula string
			want    bool
		}{
			{`XOR("text",1)`, true},    // skip "text", XOR(1) = true
			{`XOR("text",TRUE)`, true}, // skip "text", XOR(TRUE) = true
			{`XOR("1",0)`, false},      // skip "1", XOR(0) = false
			{`XOR("0",0)`, false},      // skip "0", XOR(0) = false
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("string in cell reference is skipped", func(t *testing.T) {
		// XOR(Data!A1, 1) where A1 contains "text" → string is skipped, XOR(1) = TRUE
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("text"),
				{Col: 1, Row: 2}: NumberVal(1),
			},
		}
		cf := evalCompile(t, "XOR(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("XOR(A1:A2) = %v, want true (string skipped, only 1 counts)", got)
		}
	})
}

func TestORStringInRange(t *testing.T) {
	t.Run("string in range is skipped", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("text"),
				{Col: 1, Row: 2}: BoolVal(true),
			},
		}
		cf := evalCompile(t, "OR(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("OR(A1:A2) = %v, want true", got)
		}
	})
}

func TestNOT(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    bool
	}{
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

func TestNOTComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("value tests", func(t *testing.T) {
		tests := []struct {
			formula string
			want    Value
		}{
			// Basic boolean negation
			{"NOT(TRUE)", BoolVal(false)},
			{"NOT(FALSE)", BoolVal(true)},

			// Numeric arguments: nonzero = TRUE, zero = FALSE
			{"NOT(1)", BoolVal(false)},
			{"NOT(0)", BoolVal(true)},
			{"NOT(-1)", BoolVal(false)},
			{"NOT(0.001)", BoolVal(false)},
			{"NOT(100)", BoolVal(false)},
			{"NOT(-0.5)", BoolVal(false)},

			// Double negation
			{"NOT(NOT(TRUE))", BoolVal(true)},
			{"NOT(NOT(FALSE))", BoolVal(false)},

			// NOT with comparison operators
			{"NOT(1>2)", BoolVal(true)},
			{"NOT(2>1)", BoolVal(false)},
			{"NOT(1=1)", BoolVal(false)},
			{"NOT(1=2)", BoolVal(true)},
			{"NOT(3>=3)", BoolVal(false)},
			{"NOT(2<=1)", BoolVal(true)},

			// NOT with nested logical functions
			{"NOT(AND(TRUE,FALSE))", BoolVal(true)},
			{"NOT(AND(TRUE,TRUE))", BoolVal(false)},
			{"NOT(OR(FALSE,FALSE))", BoolVal(true)},
			{"NOT(OR(TRUE,FALSE))", BoolVal(false)},

			// NOT in arithmetic context: TRUE=1, FALSE=0
			{"NOT(TRUE)+NOT(FALSE)", NumberVal(1)},
			{"NOT(FALSE)*2", NumberVal(2)},
			{"NOT(TRUE)+0", NumberVal(0)},

			// String "1" is truthy (non-empty), so NOT("1") = FALSE
			{`NOT("1")`, BoolVal(false)},
			// Empty string is falsy, so NOT("") = TRUE
			{`NOT("")`, BoolVal(true)},

			// NOT with TRUE()/FALSE() function calls
			{"NOT(TRUE())", BoolVal(false)},
			{"NOT(FALSE())", BoolVal(true)},

			// Deeply nested NOT
			{"NOT(NOT(NOT(TRUE)))", BoolVal(false)},
			{"NOT(NOT(NOT(FALSE)))", BoolVal(true)},

			// NOT used inside IF
			{`IF(NOT(FALSE),"yes","no")`, StringVal("yes")},
			{`IF(NOT(TRUE),"yes","no")`, StringVal("no")},
		}

		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if !reflect.DeepEqual(got, tt.want) {
					t.Fatalf("Eval(%q) = %#v, want %#v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		tests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// #N/A propagation
			{"NOT(NA())", ErrValNA},
			// #DIV/0! propagation
			{"NOT(1/0)", ErrValDIV0},
			// Wrong arg count: 0 args
			{"NOT()", ErrValVALUE},
		}

		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected Go error: %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Fatalf("Eval(%q) = %#v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count two args", func(t *testing.T) {
		// NOT with 2 arguments should return #VALUE!
		// Direct function call since parser may not allow 2 args easily
		got, err := fnNOT([]Value{BoolVal(true), BoolVal(false)})
		if err != nil {
			t.Fatalf("fnNOT: unexpected Go error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Fatalf("fnNOT(TRUE,FALSE) = %#v, want #VALUE!", got)
		}
	})

	t.Run("direct function call with error values", func(t *testing.T) {
		// Passing an error value directly should propagate it
		got, err := fnNOT([]Value{ErrorVal(ErrValNA)})
		if err != nil {
			t.Fatalf("fnNOT: unexpected Go error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Fatalf("fnNOT(#N/A) = %#v, want #N/A", got)
		}

		got, err = fnNOT([]Value{ErrorVal(ErrValDIV0)})
		if err != nil {
			t.Fatalf("fnNOT: unexpected Go error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Fatalf("fnNOT(#DIV/0!) = %#v, want #DIV/0!", got)
		}

		got, err = fnNOT([]Value{ErrorVal(ErrValREF)})
		if err != nil {
			t.Fatalf("fnNOT: unexpected Go error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValREF {
			t.Fatalf("fnNOT(#REF!) = %#v, want #REF!", got)
		}
	})

	t.Run("direct function call with empty value", func(t *testing.T) {
		// Empty value is falsy, so NOT(empty) = TRUE
		got, err := fnNOT([]Value{EmptyVal()})
		if err != nil {
			t.Fatalf("fnNOT: unexpected Go error: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Fatalf("fnNOT(empty) = %#v, want TRUE", got)
		}
	})

	t.Run("direct function call with zero args", func(t *testing.T) {
		got, err := fnNOT([]Value{})
		if err != nil {
			t.Fatalf("fnNOT: unexpected Go error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Fatalf("fnNOT() = %#v, want #VALUE!", got)
		}
	})
}

func TestIF(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("true branch", func(t *testing.T) {
		cf := evalCompile(t, `IF(TRUE, "yes", "no")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IF(TRUE, "yes", "no") = %v, want "yes"`, got)
		}
	})

	t.Run("false branch", func(t *testing.T) {
		cf := evalCompile(t, `IF(FALSE, "yes", "no")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "no" {
			t.Errorf(`IF(FALSE, "yes", "no") = %v, want "no"`, got)
		}
	})

	t.Run("false without else returns FALSE", func(t *testing.T) {
		cf := evalCompile(t, `IF(FALSE, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf(`IF(FALSE, "yes") = %v, want FALSE`, got)
		}
	})
}

func TestIFTrimmedRangeOriginCondition(t *testing.T) {
	got, err := fnIF([]Value{
		trimmedRangeValue([][]Value{
			{BoolVal(true)},
		}, 1, 1, 1, 3),
		NumberVal(1),
		NumberVal(0),
	})
	if err != nil {
		t.Fatalf("fnIF: %v", err)
	}

	assertLookupValueEqual(t, got, Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(1)},
		{NumberVal(0)},
		{NumberVal(0)},
	}})
	if got.RangeOrigin == nil || got.RangeOrigin.FromRow != 1 || got.RangeOrigin.ToRow != 3 {
		t.Fatalf("fnIF RangeOrigin = %+v, want rows 1:3", got.RangeOrigin)
	}
}

func TestIFERROR(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("non-error returns value", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(1, "fallback")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf(`IFERROR(1, "fallback") = %v, want 1`, got)
		}
	})

	t.Run("error returns fallback", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(1/0, "fallback")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "fallback" {
			t.Errorf(`IFERROR(1/0, "fallback") = %v, want "fallback"`, got)
		}
	})

	t.Run("DIV/0 error returns value_if_error", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(1/0, 42)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 42 {
			t.Errorf(`IFERROR(1/0, 42) = %v, want 42`, got)
		}
	})

	t.Run("VALUE error returns value_if_error", func(t *testing.T) {
		// AND with no args produces #VALUE!
		cf := evalCompile(t, `IFERROR(AND(), "caught")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "caught" {
			t.Errorf(`IFERROR(AND(), "caught") = %v, want "caught"`, got)
		}
	})

	t.Run("NA error returns value_if_error", func(t *testing.T) {
		// IFS with no true condition produces #N/A
		cf := evalCompile(t, `IFERROR(IFS(FALSE,"x"), "na caught")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "na caught" {
			t.Errorf(`IFERROR(IFS(FALSE,"x"), "na caught") = %v, want "na caught"`, got)
		}
	})

	t.Run("NUM error returns value_if_error", func(t *testing.T) {
		// SQRT(-1) produces #NUM!
		cf := evalCompile(t, `IFERROR(SQRT(-1), "num error")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "num error" {
			t.Errorf(`IFERROR(SQRT(-1), "num error") = %v, want "num error"`, got)
		}
	})

	t.Run("string value no error returns string", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR("hello", "fallback")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "hello" {
			t.Errorf(`IFERROR("hello", "fallback") = %v, want "hello"`, got)
		}
	})

	t.Run("number value no error returns number", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(99.5, 0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 99.5 {
			t.Errorf(`IFERROR(99.5, 0) = %v, want 99.5`, got)
		}
	})

	t.Run("boolean value no error returns boolean", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(TRUE, FALSE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf(`IFERROR(TRUE, FALSE) = %v, want TRUE`, got)
		}
	})

	t.Run("nested IFERROR", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(IFERROR(1/0, 1/0), "double fallback")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "double fallback" {
			t.Errorf(`IFERROR(IFERROR(1/0, 1/0), "double fallback") = %v, want "double fallback"`, got)
		}
	})

	t.Run("nested IFERROR inner catches", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(IFERROR(1/0, 42), "outer")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 42 {
			t.Errorf(`IFERROR(IFERROR(1/0, 42), "outer") = %v, want 42`, got)
		}
	})

	t.Run("value_if_error is 0", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(1/0, 0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf(`IFERROR(1/0, 0) = %v, want 0`, got)
		}
	})

	t.Run("value_if_error is empty string", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(1/0, "")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "" {
			t.Errorf(`IFERROR(1/0, "") = %v, want ""`, got)
		}
	})

	t.Run("too few args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`IFERROR(1) = %v, want #VALUE!`, got)
		}
	})

	t.Run("too many args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(1, 2, 3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`IFERROR(1, 2, 3) = %v, want #VALUE!`, got)
		}
	})

	t.Run("doc example 210/35 no error", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(210/35, "Error in calculation")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 6 {
			t.Errorf(`IFERROR(210/35, "Error in calculation") = %v, want 6`, got)
		}
	})

	t.Run("doc example 55/0 error", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(55/0, "Error in calculation")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "Error in calculation" {
			t.Errorf(`IFERROR(55/0, "Error in calculation") = %v, want "Error in calculation"`, got)
		}
	})

	t.Run("expression result no error", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(2+3, "err")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf(`IFERROR(2+3, "err") = %v, want 5`, got)
		}
	})
}

func TestIFS(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("single condition true returns value", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IFS(TRUE, "yes") = %v, want "yes"`, got)
		}
	})

	t.Run("single condition false returns NA", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`IFS(FALSE, "yes") = %v, want #N/A`, got)
		}
	})

	t.Run("first true wins", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, "first", TRUE, "second")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "first" {
			t.Errorf(`IFS(TRUE,"first",TRUE,"second") = %v, want "first"`, got)
		}
	})

	t.Run("second condition true", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE, "first", TRUE, "second")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "second" {
			t.Errorf(`IFS(FALSE,"first",TRUE,"second") = %v, want "second"`, got)
		}
	})

	t.Run("last condition true wins", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE, "a", FALSE, "b", TRUE, "c")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "c" {
			t.Errorf(`IFS(FALSE,"a",FALSE,"b",TRUE,"c") = %v, want "c"`, got)
		}
	})

	t.Run("all conditions false returns NA", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE, 1, FALSE, 2, FALSE, 3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`IFS(FALSE,1,FALSE,2,FALSE,3) = %v, want #N/A`, got)
		}
	})

	t.Run("returns number value", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, 42)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 42 {
			t.Errorf(`IFS(TRUE, 42) = %v, want 42`, got)
		}
	})

	t.Run("returns boolean value", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, FALSE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf(`IFS(TRUE, FALSE) = %v, want FALSE`, got)
		}
	})

	t.Run("returns string value", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, "hello")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "hello" {
			t.Errorf(`IFS(TRUE, "hello") = %v, want "hello"`, got)
		}
	})

	t.Run("numeric condition non-zero is true", func(t *testing.T) {
		cf := evalCompile(t, `IFS(1, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IFS(1, "yes") = %v, want "yes"`, got)
		}
	})

	t.Run("numeric condition zero is false", func(t *testing.T) {
		cf := evalCompile(t, `IFS(0, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`IFS(0, "yes") = %v, want #N/A`, got)
		}
	})

	t.Run("negative number is truthy", func(t *testing.T) {
		cf := evalCompile(t, `IFS(-1, "neg")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "neg" {
			t.Errorf(`IFS(-1, "neg") = %v, want "neg"`, got)
		}
	})

	t.Run("fractional number is truthy", func(t *testing.T) {
		cf := evalCompile(t, `IFS(0.5, "frac")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "frac" {
			t.Errorf(`IFS(0.5, "frac") = %v, want "frac"`, got)
		}
	})

	t.Run("non-empty string condition is truthy", func(t *testing.T) {
		cf := evalCompile(t, `IFS("text", "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IFS("text", "yes") = %v, want "yes"`, got)
		}
	})

	t.Run("empty string condition is falsy", func(t *testing.T) {
		cf := evalCompile(t, `IFS("", "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`IFS("", "yes") = %v, want #N/A`, got)
		}
	})

	t.Run("odd number of args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`IFS(TRUE) = %v, want #VALUE!`, got)
		}
	})

	t.Run("three args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, 1, FALSE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`IFS(TRUE, 1, FALSE) = %v, want #VALUE!`, got)
		}
	})

	t.Run("error in condition propagates", func(t *testing.T) {
		cf := evalCompile(t, `IFS(1/0, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf(`IFS(1/0, "yes") = %v, want #DIV/0!`, got)
		}
	})

	t.Run("error in second condition propagates when first is false", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE, "a", 1/0, "b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf(`IFS(FALSE,"a",1/0,"b") = %v, want #DIV/0!`, got)
		}
	})

	t.Run("error in value not taken is not propagated", func(t *testing.T) {
		// When the first condition is TRUE, the second pair's value (1/0) is never reached
		cf := evalCompile(t, `IFS(TRUE, "ok", FALSE, 1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "ok" {
			t.Errorf(`IFS(TRUE,"ok",FALSE,1/0) = %v, want "ok"`, got)
		}
	})

	t.Run("error in taken value propagates", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, 1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf(`IFS(TRUE, 1/0) = %v, want #DIV/0!`, got)
		}
	})

	t.Run("expression conditions", func(t *testing.T) {
		cf := evalCompile(t, `IFS(1>2, "a", 2>1, "b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "b" {
			t.Errorf(`IFS(1>2,"a",2>1,"b") = %v, want "b"`, got)
		}
	})

	t.Run("nested IFS", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, IFS(FALSE, 1, TRUE, 2))`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf(`IFS(TRUE, IFS(FALSE,1,TRUE,2)) = %v, want 2`, got)
		}
	})

	t.Run("cell reference as condition true", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
			},
		}
		cf := evalCompile(t, `IFS(A1, "yes")`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IFS(A1, "yes") with A1=TRUE = %v, want "yes"`, got)
		}
	})

	t.Run("cell reference as condition false", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(false),
			},
		}
		cf := evalCompile(t, `IFS(A1, "yes")`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`IFS(A1, "yes") with A1=FALSE = %v, want #N/A`, got)
		}
	})

	t.Run("cell reference numeric condition", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(0),
				{Col: 2, Row: 1}: NumberVal(5),
			},
		}
		cf := evalCompile(t, `IFS(A1, "zero", B1, "five")`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "five" {
			t.Errorf(`IFS(A1,"zero",B1,"five") = %v, want "five"`, got)
		}
	})

	t.Run("cell reference with error", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNA),
			},
		}
		cf := evalCompile(t, `IFS(A1, "yes")`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`IFS(A1, "yes") with A1=#N/A = %v, want #N/A`, got)
		}
	})

	t.Run("many conditions middle match", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE,1,FALSE,2,TRUE,3,FALSE,4,FALSE,5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf(`IFS with 5 pairs, 3rd true = %v, want 3`, got)
		}
	})

	t.Run("value is zero", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, 0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf(`IFS(TRUE, 0) = %v, want 0`, got)
		}
	})

	t.Run("value is empty string", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, "")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "" {
			t.Errorf(`IFS(TRUE, "") = %v, want ""`, got)
		}
	})

	t.Run("computed value expression", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, 10+5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf(`IFS(TRUE, 10+5) = %v, want 15`, got)
		}
	})
}

func TestSORT_BasicAscending(t *testing.T) {
	// SORT({3;1;2}) → {1;2;3}
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(3)},
		{NumberVal(1)},
		{NumberVal(2)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 ||
		got.Array[0][0].Num != 1 || got.Array[1][0].Num != 2 || got.Array[2][0].Num != 3 {
		t.Errorf("SORT ascending = %v, want {1;2;3}", got)
	}
}

func TestSORT_Descending(t *testing.T) {
	// SORT({3;1;2}, 1, -1) → {3;2;1}
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(3)},
		{NumberVal(1)},
		{NumberVal(2)},
	}}
	got, err := fnSORT([]Value{arr, NumberVal(1), NumberVal(-1)})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 ||
		got.Array[0][0].Num != 3 || got.Array[1][0].Num != 2 || got.Array[2][0].Num != 1 {
		t.Errorf("SORT descending = %v, want {3;2;1}", got)
	}
}

func TestSORT_BySecondColumnAscending(t *testing.T) {
	// SORT({{"a",3};{"b",1};{"c",2}}, 2) → {{"b",1};{"c",2};{"a",3}}
	arr := Value{Type: ValueArray, Array: [][]Value{
		{StringVal("a"), NumberVal(3)},
		{StringVal("b"), NumberVal(1)},
		{StringVal("c"), NumberVal(2)},
	}}
	got, err := fnSORT([]Value{arr, NumberVal(2)})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 ||
		got.Array[0][0].Str != "b" || got.Array[0][1].Num != 1 ||
		got.Array[1][0].Str != "c" || got.Array[1][1].Num != 2 ||
		got.Array[2][0].Str != "a" || got.Array[2][1].Num != 3 {
		t.Errorf("SORT by col 2 = %v, want b1,c2,a3", got)
	}
}

func TestSORT_BySecondColumnDescending(t *testing.T) {
	arr := Value{Type: ValueArray, Array: [][]Value{
		{StringVal("a"), NumberVal(3)},
		{StringVal("b"), NumberVal(1)},
		{StringVal("c"), NumberVal(2)},
	}}
	got, err := fnSORT([]Value{arr, NumberVal(2), NumberVal(-1)})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 ||
		got.Array[0][0].Str != "a" || got.Array[0][1].Num != 3 ||
		got.Array[1][0].Str != "c" || got.Array[1][1].Num != 2 ||
		got.Array[2][0].Str != "b" || got.Array[2][1].Num != 1 {
		t.Errorf("SORT by col 2 desc = %v, want a3,c2,b1", got)
	}
}

func TestSORT_SingleElement(t *testing.T) {
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(42)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || got.Array[0][0].Num != 42 {
		t.Errorf("SORT single = %v, want {42}", got)
	}
}

func TestSORT_AlreadySorted(t *testing.T) {
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(1)},
		{NumberVal(2)},
		{NumberVal(3)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 ||
		got.Array[0][0].Num != 1 || got.Array[1][0].Num != 2 || got.Array[2][0].Num != 3 {
		t.Errorf("SORT already sorted = %v, want {1;2;3}", got)
	}
}

func TestSORT_Numbers(t *testing.T) {
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(100)},
		{NumberVal(5)},
		{NumberVal(50)},
		{NumberVal(10)},
		{NumberVal(1)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 5 ||
		got.Array[0][0].Num != 1 || got.Array[1][0].Num != 5 ||
		got.Array[2][0].Num != 10 || got.Array[3][0].Num != 50 ||
		got.Array[4][0].Num != 100 {
		t.Errorf("SORT numbers = %v, want {1;5;10;50;100}", got)
	}
}

func TestSORT_Strings(t *testing.T) {
	arr := Value{Type: ValueArray, Array: [][]Value{
		{StringVal("cherry")},
		{StringVal("apple")},
		{StringVal("banana")},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 ||
		got.Array[0][0].Str != "apple" ||
		got.Array[1][0].Str != "banana" ||
		got.Array[2][0].Str != "cherry" {
		t.Errorf("SORT strings = %v, want apple,banana,cherry", got)
	}
}

func TestSORT_MixedTypes(t *testing.T) {
	// Numbers sort before strings
	arr := Value{Type: ValueArray, Array: [][]Value{
		{StringVal("b")},
		{NumberVal(2)},
		{StringVal("a")},
		{NumberVal(1)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 4 {
		t.Fatalf("SORT mixed = %v, want 4 rows", got)
	}
	// Numbers should come first, then strings
	if got.Array[0][0].Num != 1 || got.Array[1][0].Num != 2 ||
		got.Array[2][0].Str != "a" || got.Array[3][0].Str != "b" {
		t.Errorf("SORT mixed = %v, want {1;2;a;b}", got)
	}
}

func TestSORT_MultipleColumns(t *testing.T) {
	// Sort 3x3 array by first column
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(3), StringVal("c"), NumberVal(30)},
		{NumberVal(1), StringVal("a"), NumberVal(10)},
		{NumberVal(2), StringVal("b"), NumberVal(20)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 ||
		got.Array[0][0].Num != 1 || got.Array[0][1].Str != "a" || got.Array[0][2].Num != 10 ||
		got.Array[1][0].Num != 2 || got.Array[1][1].Str != "b" || got.Array[1][2].Num != 20 ||
		got.Array[2][0].Num != 3 || got.Array[2][1].Str != "c" || got.Array[2][2].Num != 30 {
		t.Errorf("SORT multi-col = %v, want sorted by col 1", got)
	}
}

func TestSORT_NegativeValues(t *testing.T) {
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(0)},
		{NumberVal(-5)},
		{NumberVal(3)},
		{NumberVal(-10)},
		{NumberVal(1)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 5 ||
		got.Array[0][0].Num != -10 || got.Array[1][0].Num != -5 ||
		got.Array[2][0].Num != 0 || got.Array[3][0].Num != 1 ||
		got.Array[4][0].Num != 3 {
		t.Errorf("SORT negative = %v, want {-10;-5;0;1;3}", got)
	}
}

func TestSORT_Duplicates(t *testing.T) {
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(3)},
		{NumberVal(1)},
		{NumberVal(3)},
		{NumberVal(1)},
		{NumberVal(2)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 5 ||
		got.Array[0][0].Num != 1 || got.Array[1][0].Num != 1 ||
		got.Array[2][0].Num != 2 || got.Array[3][0].Num != 3 ||
		got.Array[4][0].Num != 3 {
		t.Errorf("SORT duplicates = %v, want {1;1;2;3;3}", got)
	}
}

func TestSORT_TooFewArgs(t *testing.T) {
	got, err := fnSORT([]Value{})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("SORT() = %v, want #VALUE!", got)
	}
}

func TestSORT_TooManyArgs(t *testing.T) {
	arr := Value{Type: ValueArray, Array: [][]Value{{NumberVal(1)}}}
	got, err := fnSORT([]Value{arr, NumberVal(1), NumberVal(1), BoolVal(false), NumberVal(99)})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("SORT with 5 args = %v, want #VALUE!", got)
	}
}

func TestSORT_NonArrayInput(t *testing.T) {
	got, err := fnSORT([]Value{NumberVal(42)})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("SORT(42) = %v, want 42", got)
	}
}

func TestSORT_AscendingExplicit(t *testing.T) {
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(3)},
		{NumberVal(1)},
		{NumberVal(2)},
	}}
	got, err := fnSORT([]Value{arr, NumberVal(1), NumberVal(1)})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 ||
		got.Array[0][0].Num != 1 || got.Array[1][0].Num != 2 || got.Array[2][0].Num != 3 {
		t.Errorf("SORT asc explicit = %v, want {1;2;3}", got)
	}
}

func TestSORT_DoesNotMutateOriginal(t *testing.T) {
	original := [][]Value{
		{NumberVal(3)},
		{NumberVal(1)},
		{NumberVal(2)},
	}
	arr := Value{Type: ValueArray, Array: original}
	_, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	// Original should be unchanged
	if original[0][0].Num != 3 || original[1][0].Num != 1 || original[2][0].Num != 2 {
		t.Errorf("SORT mutated original array: %v", original)
	}
}

func TestSORT_StableSort(t *testing.T) {
	// Sort by col 1 (all same), original row order should be preserved
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(1), StringVal("first")},
		{NumberVal(1), StringVal("second")},
		{NumberVal(1), StringVal("third")},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 ||
		got.Array[0][1].Str != "first" ||
		got.Array[1][1].Str != "second" ||
		got.Array[2][1].Str != "third" {
		t.Errorf("SORT stable = %v, want original order preserved", got)
	}
}

func TestSORTBY(t *testing.T) {
	t.Run("basic ascending by single key", func(t *testing.T) {
		// SORTBY({5;3;1;4;2}, {50;30;10;40;20}) → {1;2;3;4;5}
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(5)}, {NumberVal(3)}, {NumberVal(1)}, {NumberVal(4)}, {NumberVal(2)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(50)}, {NumberVal(30)}, {NumberVal(10)}, {NumberVal(40)}, {NumberVal(20)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 5 ||
			got.Array[0][0].Num != 1 || got.Array[1][0].Num != 2 ||
			got.Array[2][0].Num != 3 || got.Array[3][0].Num != 4 ||
			got.Array[4][0].Num != 5 {
			t.Errorf("SORTBY ascending = %v, want {1;2;3;4;5}", got)
		}
	})

	t.Run("descending sort", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(10)}, {NumberVal(20)}, {NumberVal(30)},
		}}
		got, err := fnSORTBY([]Value{arr, by, NumberVal(-1)})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Num != 3 || got.Array[1][0].Num != 2 || got.Array[2][0].Num != 1 {
			t.Errorf("SORTBY descending = %v, want {3;2;1}", got)
		}
	})

	t.Run("multiple sort keys primary and secondary", func(t *testing.T) {
		// Sort by first key ascending, then second key descending to break ties.
		// Rows: (A,2), (B,1), (A,1), (B,2)
		// By1:   1, 2, 1, 2  (group: A=1, B=2)
		// By2:   2, 1, 1, 2  (secondary sort descending)
		// Expected: (A,2), (A,1), (B,2), (B,1) → by1 asc, by2 desc
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("A"), NumberVal(2)},
			{StringVal("B"), NumberVal(1)},
			{StringVal("A"), NumberVal(1)},
			{StringVal("B"), NumberVal(2)},
		}}
		by1 := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(1)}, {NumberVal(2)},
		}}
		by2 := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(2)}, {NumberVal(1)}, {NumberVal(1)}, {NumberVal(2)},
		}}
		got, err := fnSORTBY([]Value{arr, by1, NumberVal(1), by2, NumberVal(-1)})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 4 ||
			got.Array[0][0].Str != "A" || got.Array[0][1].Num != 2 ||
			got.Array[1][0].Str != "A" || got.Array[1][1].Num != 1 ||
			got.Array[2][0].Str != "B" || got.Array[2][1].Num != 2 ||
			got.Array[3][0].Str != "B" || got.Array[3][1].Num != 1 {
			t.Errorf("SORTBY multi-key = %v, want A2,A1,B2,B1", got)
		}
	})

	t.Run("sort multi-column array by external key", func(t *testing.T) {
		// Array has multiple columns, sort by an external key.
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("c"), NumberVal(30)},
			{StringVal("a"), NumberVal(10)},
			{StringVal("b"), NumberVal(20)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(3)}, {NumberVal(1)}, {NumberVal(2)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Str != "a" || got.Array[0][1].Num != 10 ||
			got.Array[1][0].Str != "b" || got.Array[1][1].Num != 20 ||
			got.Array[2][0].Str != "c" || got.Array[2][1].Num != 30 {
			t.Errorf("SORTBY multi-col = %v, want a10,b20,c30", got)
		}
	})

	t.Run("single element array", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(42)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 1 || got.Array[0][0].Num != 42 {
			t.Errorf("SORTBY single = %v, want {42}", got)
		}
	})

	t.Run("already sorted data", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(10)}, {NumberVal(20)}, {NumberVal(30)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Num != 1 || got.Array[1][0].Num != 2 || got.Array[2][0].Num != 3 {
			t.Errorf("SORTBY already sorted = %v, want {1;2;3}", got)
		}
	})

	t.Run("reverse sorted data", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(3)}, {NumberVal(2)}, {NumberVal(1)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(30)}, {NumberVal(20)}, {NumberVal(10)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Num != 1 || got.Array[1][0].Num != 2 || got.Array[2][0].Num != 3 {
			t.Errorf("SORTBY reverse = %v, want {1;2;3}", got)
		}
	})

	t.Run("duplicate sort keys preserves order (stability)", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("first")},
			{StringVal("second")},
			{StringVal("third")},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(1)}, {NumberVal(1)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Str != "first" ||
			got.Array[1][0].Str != "second" ||
			got.Array[2][0].Str != "third" {
			t.Errorf("SORTBY stable = %v, want original order preserved", got)
		}
	})

	t.Run("string sort keys", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("cherry")}, {StringVal("apple")}, {StringVal("banana")},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Num != 2 || got.Array[1][0].Num != 3 || got.Array[2][0].Num != 1 {
			t.Errorf("SORTBY string keys = %v, want {2;3;1} (sorted by apple,banana,cherry)", got)
		}
	})

	t.Run("mixed type sort keys", func(t *testing.T) {
		// Numbers sort before strings.
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("A")}, {StringVal("B")}, {StringVal("C")}, {StringVal("D")},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("z")}, {NumberVal(2)}, {StringVal("a")}, {NumberVal(1)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 4 ||
			got.Array[0][0].Str != "D" || got.Array[1][0].Str != "B" ||
			got.Array[2][0].Str != "C" || got.Array[3][0].Str != "A" {
			t.Errorf("SORTBY mixed keys = %v, want {D;B;C;A} (nums first, then strings)", got)
		}
	})

	t.Run("numeric sort keys with negatives and decimals", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("A")}, {StringVal("B")}, {StringVal("C")}, {StringVal("D")}, {StringVal("E")},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(0)}, {NumberVal(-5.5)}, {NumberVal(3.14)}, {NumberVal(-10)}, {NumberVal(1.5)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 5 ||
			got.Array[0][0].Str != "D" || // -10
			got.Array[1][0].Str != "B" || // -5.5
			got.Array[2][0].Str != "A" || // 0
			got.Array[3][0].Str != "E" || // 1.5
			got.Array[4][0].Str != "C" { // 3.14
			t.Errorf("SORTBY negative/decimal keys = %v, want {D;B;A;E;C}", got)
		}
	})

	t.Run("error propagation from array", func(t *testing.T) {
		arr := ErrorVal(ErrValREF)
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueError || got.Err != ErrValREF {
			t.Errorf("SORTBY error array = %v, want #REF!", got)
		}
	})

	t.Run("error propagation from by_array", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)},
		}}
		by := ErrorVal(ErrValNA)
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("SORTBY error by_array = %v, want #N/A", got)
		}
	})

	t.Run("error values in by_array sort last", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("A")}, {StringVal("B")}, {StringVal("C")},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(10)}, {ErrorVal(ErrValDIV0)}, {NumberVal(30)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Str != "A" || got.Array[1][0].Str != "C" || got.Array[2][0].Str != "B" {
			t.Errorf("SORTBY error in by_array = %v, want {A;C;B}", got)
		}
	})

	t.Run("invalid sort order not 1 or -1 returns VALUE error", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(10)}, {NumberVal(20)},
		}}
		got, err := fnSORTBY([]Value{arr, by, NumberVal(2)})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("SORTBY invalid order = %v, want #VALUE!", got)
		}
	})

	t.Run("invalid sort order zero returns VALUE error", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(10)}, {NumberVal(20)},
		}}
		got, err := fnSORTBY([]Value{arr, by, NumberVal(0)})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("SORTBY order=0 = %v, want #VALUE!", got)
		}
	})

	t.Run("by_array size mismatch returns VALUE error", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(10)}, {NumberVal(20)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("SORTBY size mismatch = %v, want #VALUE!", got)
		}
	})

	t.Run("too few args returns VALUE error", func(t *testing.T) {
		got, err := fnSORTBY([]Value{})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("SORTBY() = %v, want #VALUE!", got)
		}
	})

	t.Run("one arg returns VALUE error", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
		}}
		got, err := fnSORTBY([]Value{arr})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("SORTBY(arr) = %v, want #VALUE!", got)
		}
	})

	t.Run("scalar inputs treated as 1x1 array", func(t *testing.T) {
		got, err := fnSORTBY([]Value{NumberVal(42), NumberVal(1)})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 1 || got.Array[0][0].Num != 42 {
			t.Errorf("SORTBY(42,1) = %v, want {{42}}", got)
		}
	})

	t.Run("does not mutate original array", func(t *testing.T) {
		original := [][]Value{
			{NumberVal(3)}, {NumberVal(1)}, {NumberVal(2)},
		}
		arr := Value{Type: ValueArray, Array: original}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(30)}, {NumberVal(10)}, {NumberVal(20)},
		}}
		_, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if original[0][0].Num != 3 || original[1][0].Num != 1 || original[2][0].Num != 2 {
			t.Errorf("SORTBY mutated original array: %v", original)
		}
	})

	t.Run("row vector by_array mismatched against column source", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("c")}, {StringVal("a")}, {StringVal("b")},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(30), NumberVal(10), NumberVal(20)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("SORTBY row vector mismatch = %v, want #VALUE!", got)
		}
	})

	t.Run("row vector by_array sorts columns of row source", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("c"), StringVal("a"), StringVal("b")},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(30), NumberVal(10), NumberVal(20)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 3 ||
			got.Array[0][0].Str != "a" || got.Array[0][1].Str != "b" || got.Array[0][2].Str != "c" {
			t.Errorf("SORTBY row source columns = %v, want {a,b,c}", got)
		}
	})

	t.Run("by_array is 2D array returns VALUE error", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(10), NumberVal(20)},
			{NumberVal(30), NumberVal(40)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("SORTBY 2D by_array = %v, want #VALUE!", got)
		}
	})

	t.Run("multiple keys with default sort order on last", func(t *testing.T) {
		// SORTBY(arr, by1, 1, by2) — 4 args: second key defaults to ascending.
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("A")}, {StringVal("B")}, {StringVal("C")}, {StringVal("D")},
		}}
		by1 := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(2)}, {NumberVal(1)}, {NumberVal(2)}, {NumberVal(1)},
		}}
		by2 := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(2)}, {NumberVal(2)}, {NumberVal(1)}, {NumberVal(1)},
		}}
		got, err := fnSORTBY([]Value{arr, by1, NumberVal(1), by2})
		if err != nil {
			t.Fatal(err)
		}
		// Sort by1 asc: B(1),D(1) then A(2),C(2). By2 asc within: D(1),B(2), C(1),A(2).
		if got.Type != ValueArray || len(got.Array) != 4 ||
			got.Array[0][0].Str != "D" || got.Array[1][0].Str != "B" ||
			got.Array[2][0].Str != "C" || got.Array[3][0].Str != "A" {
			t.Errorf("SORTBY multi-key default order = %v, want {D;B;C;A}", got)
		}
	})

	t.Run("cell reference based test", func(t *testing.T) {
		// Set up A1:A5 = {5,3,1,4,2}, B1:B5 = {50,30,10,40,20}
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(5),
				{Col: 1, Row: 2}: NumberVal(3),
				{Col: 1, Row: 3}: NumberVal(1),
				{Col: 1, Row: 4}: NumberVal(4),
				{Col: 1, Row: 5}: NumberVal(2),
				{Col: 2, Row: 1}: NumberVal(50),
				{Col: 2, Row: 2}: NumberVal(30),
				{Col: 2, Row: 3}: NumberVal(10),
				{Col: 2, Row: 4}: NumberVal(40),
				{Col: 2, Row: 5}: NumberVal(20),
			},
		}
		cf := evalCompile(t, "SORTBY(A1:A5,B1:B5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 5 ||
			got.Array[0][0].Num != 1 || got.Array[1][0].Num != 2 ||
			got.Array[2][0].Num != 3 || got.Array[3][0].Num != 4 ||
			got.Array[4][0].Num != 5 {
			t.Errorf("SORTBY(A1:A5,B1:B5) = %v, want {1;2;3;4;5}", got)
		}
	})

	t.Run("cell reference descending", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(5),
				{Col: 1, Row: 2}: NumberVal(3),
				{Col: 1, Row: 3}: NumberVal(1),
				{Col: 2, Row: 1}: NumberVal(50),
				{Col: 2, Row: 2}: NumberVal(30),
				{Col: 2, Row: 3}: NumberVal(10),
			},
		}
		cf := evalCompile(t, "SORTBY(A1:A3,B1:B3,-1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Num != 5 || got.Array[1][0].Num != 3 || got.Array[2][0].Num != 1 {
			t.Errorf("SORTBY descending cell ref = %v, want {5;3;1}", got)
		}
	})

	t.Run("boolean sort keys", func(t *testing.T) {
		// Booleans: FALSE sorts before TRUE.
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("A")}, {StringVal("B")}, {StringVal("C")},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{BoolVal(true)}, {BoolVal(false)}, {BoolVal(true)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Str != "B" ||
			got.Array[1][0].Str != "A" ||
			got.Array[2][0].Str != "C" {
			t.Errorf("SORTBY boolean keys = %v, want {B;A;C}", got)
		}
	})

	t.Run("sort order as string coerced to number", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(10)}, {NumberVal(20)}, {NumberVal(30)},
		}}
		got, err := fnSORTBY([]Value{arr, by, StringVal("-1")})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Num != 3 || got.Array[1][0].Num != 2 || got.Array[2][0].Num != 1 {
			t.Errorf("SORTBY string order = %v, want {3;2;1}", got)
		}
	})

	t.Run("empty array returns VALUE error", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{}}
		by := Value{Type: ValueArray, Array: [][]Value{}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("SORTBY empty = %v, want #VALUE!", got)
		}
	})

	t.Run("empty cells in by_array sort as zero among numbers", func(t *testing.T) {
		// Empty adapts to number context → 0, so should sort between -1 and 1.
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("A")}, {StringVal("B")}, {StringVal("C")},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {EmptyVal()}, {NumberVal(-1)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		// Expected order: C(-1), B(0/empty), A(1)
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Str != "C" ||
			got.Array[1][0].Str != "B" ||
			got.Array[2][0].Str != "A" {
			t.Errorf("SORTBY empty cells = %v, want {C;B;A}", got)
		}
	})

	t.Run("INDEX into SORTBY result", func(t *testing.T) {
		// INDEX(SORTBY({5;3;1;4;2},{50;30;10;40;20}),3,1) → 3 (third element after sort)
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(5),
				{Col: 1, Row: 2}: NumberVal(3),
				{Col: 1, Row: 3}: NumberVal(1),
				{Col: 1, Row: 4}: NumberVal(4),
				{Col: 1, Row: 5}: NumberVal(2),
				{Col: 2, Row: 1}: NumberVal(50),
				{Col: 2, Row: 2}: NumberVal(30),
				{Col: 2, Row: 3}: NumberVal(10),
				{Col: 2, Row: 4}: NumberVal(40),
				{Col: 2, Row: 5}: NumberVal(20),
			},
		}
		cf := evalCompile(t, "INDEX(SORTBY(A1:A5,B1:B5),3,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("INDEX(SORTBY(...),3,1) = %v, want 3", got)
		}
	})

	t.Run("INDEX into SORTBY result first element", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(5),
				{Col: 1, Row: 2}: NumberVal(3),
				{Col: 1, Row: 3}: NumberVal(1),
				{Col: 2, Row: 1}: NumberVal(50),
				{Col: 2, Row: 2}: NumberVal(30),
				{Col: 2, Row: 3}: NumberVal(10),
			},
		}
		cf := evalCompile(t, "INDEX(SORTBY(A1:A3,B1:B3),1,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("INDEX(SORTBY(...),1,1) = %v, want 1", got)
		}
	})

	t.Run("large array sorting 20 elements", func(t *testing.T) {
		// Sort 20 elements in descending order by key.
		n := 20
		arrRows := make([][]Value, n)
		byRows := make([][]Value, n)
		for i := 0; i < n; i++ {
			arrRows[i] = []Value{NumberVal(float64(i + 1))}
			byRows[i] = []Value{NumberVal(float64(i + 1))}
		}
		arr := Value{Type: ValueArray, Array: arrRows}
		by := Value{Type: ValueArray, Array: byRows}
		got, err := fnSORTBY([]Value{arr, by, NumberVal(-1)})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != n {
			t.Fatalf("SORTBY large array len = %d, want %d", len(got.Array), n)
		}
		// Should be 20, 19, 18, ..., 1
		for i := 0; i < n; i++ {
			expected := float64(n - i)
			if got.Array[i][0].Num != expected {
				t.Errorf("SORTBY large array[%d] = %v, want %v", i, got.Array[i][0].Num, expected)
			}
		}
	})

	t.Run("sort by one column return different column via cell refs", func(t *testing.T) {
		// A1:B3 = {{Alice,90},{Bob,70},{Carol,80}}, sort by B (scores) ascending
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("Alice"),
				{Col: 1, Row: 2}: StringVal("Bob"),
				{Col: 1, Row: 3}: StringVal("Carol"),
				{Col: 2, Row: 1}: NumberVal(90),
				{Col: 2, Row: 2}: NumberVal(70),
				{Col: 2, Row: 3}: NumberVal(80),
			},
		}
		cf := evalCompile(t, "SORTBY(A1:B3,B1:B3,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Str != "Bob" || got.Array[0][1].Num != 70 ||
			got.Array[1][0].Str != "Carol" || got.Array[1][1].Num != 80 ||
			got.Array[2][0].Str != "Alice" || got.Array[2][1].Num != 90 {
			t.Errorf("SORTBY by score = %v, want Bob70,Carol80,Alice90", got)
		}
	})

	t.Run("three sort keys", func(t *testing.T) {
		// 6 rows, 3 levels of grouping:
		// by1: {1,1,2,2,1,2}, by2: {1,2,1,2,1,2}, by3: {2,1,2,1,1,2}
		// Sort all ascending. Expected order by (by1,by2,by3):
		// (1,1,1)=E, (1,1,2)=A, (1,2,1)=B, (2,1,2)=C, (2,2,1)=D, (2,2,2)=F
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("A")}, {StringVal("B")}, {StringVal("C")},
			{StringVal("D")}, {StringVal("E")}, {StringVal("F")},
		}}
		by1 := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(1)}, {NumberVal(2)},
			{NumberVal(2)}, {NumberVal(1)}, {NumberVal(2)},
		}}
		by2 := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(1)},
			{NumberVal(2)}, {NumberVal(1)}, {NumberVal(2)},
		}}
		by3 := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(2)}, {NumberVal(1)}, {NumberVal(2)},
			{NumberVal(1)}, {NumberVal(1)}, {NumberVal(2)},
		}}
		got, err := fnSORTBY([]Value{arr, by1, NumberVal(1), by2, NumberVal(1), by3, NumberVal(1)})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 6 ||
			got.Array[0][0].Str != "E" || // (1,1,1)
			got.Array[1][0].Str != "A" || // (1,1,2)
			got.Array[2][0].Str != "B" || // (1,2,1)
			got.Array[3][0].Str != "C" || // (2,1,2)
			got.Array[4][0].Str != "D" || // (2,2,1)
			got.Array[5][0].Str != "F" { // (2,2,2)
			t.Errorf("SORTBY 3 keys = %v, want {E;A;B;C;D;F}", got)
		}
	})

	t.Run("descending string sort keys", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("cherry")}, {StringVal("apple")}, {StringVal("banana")},
		}}
		got, err := fnSORTBY([]Value{arr, by, NumberVal(-1)})
		if err != nil {
			t.Fatal(err)
		}
		// Descending alphabetical: cherry, banana, apple → items 1, 3, 2
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Num != 1 || got.Array[1][0].Num != 3 || got.Array[2][0].Num != 2 {
			t.Errorf("SORTBY descending strings = %v, want {1;3;2}", got)
		}
	})

	t.Run("all same values in array", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(7)}, {NumberVal(7)}, {NumberVal(7)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(3)}, {NumberVal(1)}, {NumberVal(2)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Num != 7 || got.Array[1][0].Num != 7 || got.Array[2][0].Num != 7 {
			t.Errorf("SORTBY all same = %v, want {7;7;7}", got)
		}
	})

	t.Run("sort order as boolean TRUE treated as 1", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(3)}, {NumberVal(1)}, {NumberVal(2)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(30)}, {NumberVal(10)}, {NumberVal(20)},
		}}
		got, err := fnSORTBY([]Value{arr, by, BoolVal(true)})
		if err != nil {
			t.Fatal(err)
		}
		// TRUE coerces to 1 → ascending
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Num != 1 || got.Array[1][0].Num != 2 || got.Array[2][0].Num != 3 {
			t.Errorf("SORTBY order=TRUE = %v, want {1;2;3}", got)
		}
	})

	t.Run("second by_array size mismatch returns VALUE error", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}}
		by1 := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(10)}, {NumberVal(20)}, {NumberVal(30)},
		}}
		by2 := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, // only 2 elements, need 3
		}}
		got, err := fnSORTBY([]Value{arr, by1, NumberVal(1), by2, NumberVal(1)})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("SORTBY second by_array mismatch = %v, want #VALUE!", got)
		}
	})

	t.Run("error in second by_array sorts last within ties", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("A")}, {StringVal("B")}, {StringVal("C")},
		}}
		by1 := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(1)}, {NumberVal(1)},
		}}
		by2 := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(2)}, {ErrorVal(ErrValNUM)}, {NumberVal(1)},
		}}
		got, err := fnSORTBY([]Value{arr, by1, NumberVal(1), by2, NumberVal(1)})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Str != "C" || got.Array[1][0].Str != "A" || got.Array[2][0].Str != "B" {
			t.Errorf("SORTBY error in second by_array = %v, want {C;A;B}", got)
		}
	})

	t.Run("case insensitive string sort keys", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(3)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("Banana")}, {StringVal("apple")}, {StringVal("CHERRY")},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		// Case insensitive: apple, Banana, CHERRY → 2, 1, 3
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Num != 2 || got.Array[1][0].Num != 1 || got.Array[2][0].Num != 3 {
			t.Errorf("SORTBY case insensitive = %v, want {2;1;3}", got)
		}
	})

	t.Run("multi-column result preserves all columns", func(t *testing.T) {
		// 3 columns, sort by external key
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("c"), NumberVal(30), BoolVal(false)},
			{StringVal("a"), NumberVal(10), BoolVal(true)},
			{StringVal("b"), NumberVal(20), BoolVal(false)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(3)}, {NumberVal(1)}, {NumberVal(2)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 3 || len(got.Array[0]) != 3 ||
			got.Array[0][0].Str != "a" || got.Array[0][1].Num != 10 || got.Array[0][2].Bool != true ||
			got.Array[1][0].Str != "b" || got.Array[1][1].Num != 20 || got.Array[1][2].Bool != false ||
			got.Array[2][0].Str != "c" || got.Array[2][1].Num != 30 || got.Array[2][2].Bool != false {
			t.Errorf("SORTBY 3-col = %v, want a10true,b20false,c30false", got)
		}
	})

	t.Run("non-numeric sort order string returns VALUE error", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(10)}, {NumberVal(20)},
		}}
		got, err := fnSORTBY([]Value{arr, by, StringVal("abc")})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("SORTBY order='abc' = %v, want #VALUE!", got)
		}
	})

	t.Run("two element swap", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("second")}, {StringVal("first")},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(2)}, {NumberVal(1)},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		if got.Type != ValueArray || len(got.Array) != 2 ||
			got.Array[0][0].Str != "first" || got.Array[1][0].Str != "second" {
			t.Errorf("SORTBY two element = %v, want {first;second}", got)
		}
	})

	t.Run("duplicate sort keys with descending preserves stability", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("X")}, {StringVal("Y")}, {StringVal("Z")},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(5)}, {NumberVal(5)}, {NumberVal(5)},
		}}
		got, err := fnSORTBY([]Value{arr, by, NumberVal(-1)})
		if err != nil {
			t.Fatal(err)
		}
		// All keys equal + stable sort → original order preserved
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Str != "X" ||
			got.Array[1][0].Str != "Y" ||
			got.Array[2][0].Str != "Z" {
			t.Errorf("SORTBY stable descending = %v, want {X;Y;Z}", got)
		}
	})

	t.Run("multi-key with both descending", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("A")}, {StringVal("B")}, {StringVal("C")}, {StringVal("D")},
		}}
		by1 := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(2)}, {NumberVal(1)}, {NumberVal(2)},
		}}
		by2 := Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)}, {NumberVal(1)}, {NumberVal(2)}, {NumberVal(2)},
		}}
		got, err := fnSORTBY([]Value{arr, by1, NumberVal(-1), by2, NumberVal(-1)})
		if err != nil {
			t.Fatal(err)
		}
		// by1 desc: B(2),D(2),A(1),C(1). by2 desc within: D(2),B(1),C(2),A(1)
		if got.Type != ValueArray || len(got.Array) != 4 ||
			got.Array[0][0].Str != "D" || got.Array[1][0].Str != "B" ||
			got.Array[2][0].Str != "C" || got.Array[3][0].Str != "A" {
			t.Errorf("SORTBY both desc = %v, want {D;B;C;A}", got)
		}
	})

	t.Run("INDEX into SORTBY multi-column result", func(t *testing.T) {
		// Sort names by scores, then extract name from 2nd row
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("Alice"),
				{Col: 1, Row: 2}: StringVal("Bob"),
				{Col: 1, Row: 3}: StringVal("Carol"),
				{Col: 2, Row: 1}: NumberVal(90),
				{Col: 2, Row: 2}: NumberVal(70),
				{Col: 2, Row: 3}: NumberVal(80),
			},
		}
		cf := evalCompile(t, "INDEX(SORTBY(A1:B3,B1:B3),2,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatal(err)
		}
		// Sorted by score asc: Bob(70), Carol(80), Alice(90). Row 2 col 1 = Carol.
		if got.Type != ValueString || got.Str != "Carol" {
			t.Errorf("INDEX(SORTBY(...),2,1) = %v, want Carol", got)
		}
	})

	t.Run("workbook parity for dimension mismatch and error-key sorting", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Sheet: "data", Col: 1, Row: 2}: StringVal("a"),
				{Sheet: "data", Col: 1, Row: 3}: StringVal("b"),
				{Sheet: "data", Col: 1, Row: 4}: StringVal("c"),
				{Sheet: "data", Col: 1, Row: 5}: StringVal("d"),
				{Sheet: "data", Col: 1, Row: 6}: StringVal("e"),

				{Sheet: "data", Col: 2, Row: 2}: NumberVal(97),
				{Sheet: "data", Col: 2, Row: 3}: NumberVal(98),
				{Sheet: "data", Col: 2, Row: 4}: NumberVal(99),
				{Sheet: "data", Col: 2, Row: 5}: NumberVal(100),
				{Sheet: "data", Col: 2, Row: 6}: NumberVal(101),

				{Sheet: "data", Col: 4, Row: 2}: ErrorVal(ErrValNA),
				{Sheet: "data", Col: 4, Row: 3}: NumberVal(10),
				{Sheet: "data", Col: 4, Row: 4}: NumberVal(20),
				{Sheet: "data", Col: 4, Row: 5}: NumberVal(30),
				{Sheet: "data", Col: 4, Row: 6}: NumberVal(40),

				{Sheet: "data", Col: 3, Row: 7}: NumberVal(100),
				{Sheet: "data", Col: 4, Row: 7}: NumberVal(200),
				{Sheet: "data", Col: 5, Row: 7}: NumberVal(300),
				{Sheet: "data", Col: 6, Row: 7}: NumberVal(400),
				{Sheet: "data", Col: 7, Row: 7}: NumberVal(500),
			},
		}

		tests := []struct {
			formula string
			want    Value
		}{
			{
				formula: `IFERROR(INDEX(SORTBY(data!A2:A6, data!C7:G7), 1), "err")`,
				want:    StringVal("err"),
			},
			{
				formula: `IFERROR(INDEX(SORTBY(data!A2:A6, data!B2:B6, 1, data!C7:F7, 1), 1), "err")`,
				want:    StringVal("err"),
			},
			{
				formula: `IFERROR(INDEX(SORTBY(data!A2:A6, data!D2:D6, 1), 5), "err")`,
				want:    StringVal("a"),
			},
			{
				formula: `IFERROR(INDEX(SORTBY(data!A2:A6, data!Z1:Z5), 1), "err")`,
				want:    StringVal("a"),
			},
		}

		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				got, err := Eval(evalCompile(t, tt.formula), resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertLookupValueEqual(t, got, tt.want)
			})
		}
	})

	t.Run("all empty cells in by_array", func(t *testing.T) {
		arr := Value{Type: ValueArray, Array: [][]Value{
			{StringVal("A")}, {StringVal("B")}, {StringVal("C")},
		}}
		by := Value{Type: ValueArray, Array: [][]Value{
			{EmptyVal()}, {EmptyVal()}, {EmptyVal()},
		}}
		got, err := fnSORTBY([]Value{arr, by})
		if err != nil {
			t.Fatal(err)
		}
		// All empty → all equal → stable sort preserves order
		if got.Type != ValueArray || len(got.Array) != 3 ||
			got.Array[0][0].Str != "A" ||
			got.Array[1][0].Str != "B" ||
			got.Array[2][0].Str != "C" {
			t.Errorf("SORTBY all empty keys = %v, want {A;B;C}", got)
		}
	})
}

func TestSWITCH(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("match first value", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1, 1, "one", 2, "two", 3, "three")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "one" {
			t.Errorf(`SWITCH(1, 1,"one", 2,"two", 3,"three") = %v, want "one"`, got)
		}
	})

	t.Run("match second value", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(2, 1, "one", 2, "two", 3, "three")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "two" {
			t.Errorf(`SWITCH(2, 1,"one", 2,"two", 3,"three") = %v, want "two"`, got)
		}
	})

	t.Run("match last value", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(3, 1, "one", 2, "two", 3, "three")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "three" {
			t.Errorf(`SWITCH(3, 1,"one", 2,"two", 3,"three") = %v, want "three"`, got)
		}
	})

	t.Run("no match without default returns NA", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(99, 1, "one", 2, "two")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`SWITCH(99, 1,"one", 2,"two") = %v, want #N/A`, got)
		}
	})

	t.Run("no match with default returns default", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(99, 1, "one", 2, "two", "none")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "none" {
			t.Errorf(`SWITCH(99, 1,"one", 2,"two", "none") = %v, want "none"`, got)
		}
	})

	t.Run("default value is numeric", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(99, 1, "one", 2, "two", -1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -1 {
			t.Errorf(`SWITCH(99, 1,"one", 2,"two", -1) = %v, want -1`, got)
		}
	})

	t.Run("string matching", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH("b", "a", 1, "b", 2, "c", 3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf(`SWITCH("b", "a",1, "b",2, "c",3) = %v, want 2`, got)
		}
	})

	t.Run("number matching", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(3.14, 1, "int", 3.14, "pi", 2.72, "e")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "pi" {
			t.Errorf(`SWITCH(3.14, ...) = %v, want "pi"`, got)
		}
	})

	t.Run("boolean matching TRUE", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(TRUE, FALSE, "no", TRUE, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`SWITCH(TRUE, FALSE,"no", TRUE,"yes") = %v, want "yes"`, got)
		}
	})

	t.Run("boolean matching FALSE", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(FALSE, TRUE, "yes", FALSE, "no")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "no" {
			t.Errorf(`SWITCH(FALSE, TRUE,"yes", FALSE,"no") = %v, want "no"`, got)
		}
	})

	t.Run("string matching is case insensitive", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH("HELLO", "hello", 1, "world", 2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf(`SWITCH("HELLO", "hello",1, "world",2) = %v, want 1`, got)
		}
	})

	t.Run("expression as first arg", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1+1, 1, "one", 2, "two", 3, "three")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "two" {
			t.Errorf(`SWITCH(1+1, 1,"one", 2,"two", 3,"three") = %v, want "two"`, got)
		}
	})

	t.Run("multiple pairs match returns first match", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1, 1, "first", 1, "second", 1, "third")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "first" {
			t.Errorf(`SWITCH(1, 1,"first", 1,"second", 1,"third") = %v, want "first"`, got)
		}
	})

	t.Run("error expression propagates without default", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1/0, 1, "one", 2, "two")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf(`SWITCH(1/0, 1,"one", 2,"two") = %v, want #DIV/0!`, got)
		}
	})

	t.Run("error expression propagates with default", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1/0, 1, "one", 2, "two", "fallback")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf(`SWITCH(1/0, ..., "fallback") = %v, want #DIV/0!`, got)
		}
	})

	t.Run("udf error expression stays catchable", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(SWITCH(_xludf.IFNA(1,"x"), "x", "ok", "fallback"), "caught")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "caught" {
			t.Errorf(`IFERROR(SWITCH(_xludf.IFNA(...)), "caught") = %v, want "caught"`, got)
		}
	})

	t.Run("error in matched result propagates", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1, 1, 1/0, 2, "two")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf(`SWITCH(1, 1, 1/0, ...) = %v, want #DIV/0!`, got)
		}
	})

	t.Run("error in unmatched result does not propagate", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(2, 1, 1/0, 2, "two")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "two" {
			t.Errorf(`SWITCH(2, 1, 1/0, 2,"two") = %v, want "two"`, got)
		}
	})

	t.Run("mixed types no match", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH("hello", 1, "num", TRUE, "bool", "nomatch")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "nomatch" {
			t.Errorf(`SWITCH("hello", 1,"num", TRUE,"bool", "nomatch") = %v, want "nomatch"`, got)
		}
	})

	t.Run("too few args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1, 2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`SWITCH(1, 2) = %v, want #VALUE!`, got)
		}
	})

	t.Run("single pair match", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(5, 5, "five")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "five" {
			t.Errorf(`SWITCH(5, 5,"five") = %v, want "five"`, got)
		}
	})

	t.Run("single pair no match returns NA", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(5, 6, "six")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`SWITCH(5, 6,"six") = %v, want #N/A`, got)
		}
	})

	t.Run("zero matches zero", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(0, 1, "one", 0, "zero")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "zero" {
			t.Errorf(`SWITCH(0, 1,"one", 0,"zero") = %v, want "zero"`, got)
		}
	})

	t.Run("empty string matches empty string", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH("", "", "empty", "a", "letter")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "empty" {
			t.Errorf(`SWITCH("", "","empty", "a","letter") = %v, want "empty"`, got)
		}
	})

	t.Run("result can be numeric", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH("x", "a", 10, "x", 20, "z", 30)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf(`SWITCH("x", "a",10, "x",20, "z",30) = %v, want 20`, got)
		}
	})

	t.Run("result can be boolean", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1, 1, TRUE, 2, FALSE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf(`SWITCH(1, 1,TRUE, 2,FALSE) = %v, want TRUE`, got)
		}
	})

	t.Run("weekday example from docs", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(2, 1, "Sunday", 2, "Monday", 3, "Tuesday", "No match")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "Monday" {
			t.Errorf(`SWITCH(2, 1,"Sunday", 2,"Monday", ...) = %v, want "Monday"`, got)
		}
	})

	t.Run("weekday no match with default", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(99, 1, "Sunday", 2, "Monday", 3, "Tuesday", "No match")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "No match" {
			t.Errorf(`SWITCH(99, ..., "No match") = %v, want "No match"`, got)
		}
	})
}

// ── Additional SORT tests ────────────────────────────────────────────

func TestSORT_ByThirdColumn(t *testing.T) {
	// SORT({{"a",1,30};{"b",2,10};{"c",3,20}}, 3) → sort by col 3
	arr := Value{Type: ValueArray, Array: [][]Value{
		{StringVal("a"), NumberVal(1), NumberVal(30)},
		{StringVal("b"), NumberVal(2), NumberVal(10)},
		{StringVal("c"), NumberVal(3), NumberVal(20)},
	}}
	got, err := fnSORT([]Value{arr, NumberVal(3)})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 ||
		got.Array[0][0].Str != "b" || got.Array[0][2].Num != 10 ||
		got.Array[1][0].Str != "c" || got.Array[1][2].Num != 20 ||
		got.Array[2][0].Str != "a" || got.Array[2][2].Num != 30 {
		t.Errorf("SORT by col 3 = %v, want b10,c20,a30", got)
	}
}

func TestSORT_SingleColumn(t *testing.T) {
	// Single column array sorted ascending
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(5)},
		{NumberVal(2)},
		{NumberVal(8)},
		{NumberVal(1)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 4 ||
		got.Array[0][0].Num != 1 || got.Array[1][0].Num != 2 ||
		got.Array[2][0].Num != 5 || got.Array[3][0].Num != 8 {
		t.Errorf("SORT single col = %v, want {1;2;5;8}", got)
	}
}

func TestSORT_SingleRow(t *testing.T) {
	// Single row array (1x4) — sorts as a single-element array (no reordering needed)
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(3), NumberVal(1), NumberVal(2)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 {
		t.Fatalf("SORT single row = %v, want 1 row", got)
	}
	// Single row stays as-is since there is only one row to sort
	if got.Array[0][0].Num != 3 || got.Array[0][1].Num != 1 || got.Array[0][2].Num != 2 {
		t.Errorf("SORT single row = %v, want {3,1,2}", got.Array[0])
	}
}

func TestSORT_1x1Array(t *testing.T) {
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(99)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || got.Array[0][0].Num != 99 {
		t.Errorf("SORT 1x1 = %v, want {99}", got)
	}
}

func TestSORT_ReverseSortedAscending(t *testing.T) {
	// Reverse sorted input + ascending order
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(5)},
		{NumberVal(4)},
		{NumberVal(3)},
		{NumberVal(2)},
		{NumberVal(1)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 5 {
		t.Fatalf("expected 5 rows, got %d", len(got.Array))
	}
	for i := 0; i < 5; i++ {
		want := float64(i + 1)
		if got.Array[i][0].Num != want {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, want)
		}
	}
}

func TestSORT_SortIndexOutOfRange(t *testing.T) {
	// sort_index beyond the number of columns — should use empty values
	// (implementation uses zero value when index is out of range)
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(3)},
		{NumberVal(1)},
		{NumberVal(2)},
	}}
	got, err := fnSORT([]Value{arr, NumberVal(5)})
	if err != nil {
		t.Fatal(err)
	}
	// When sort_index is out of range, all comparison values are empty
	// so original order is preserved (stable sort)
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][0].Num != 3 || got.Array[1][0].Num != 1 || got.Array[2][0].Num != 2 {
		t.Errorf("SORT out-of-range index = %v, want original order {3;1;2}", got)
	}
}

func TestSORT_LargeArray(t *testing.T) {
	// Sort 100 elements in reverse order
	rows := make([][]Value, 100)
	for i := 0; i < 100; i++ {
		rows[i] = []Value{NumberVal(float64(100 - i))}
	}
	arr := Value{Type: ValueArray, Array: rows}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 100 {
		t.Fatalf("expected 100 rows, got %d", len(got.Array))
	}
	for i := 0; i < 100; i++ {
		want := float64(i + 1)
		if got.Array[i][0].Num != want {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, want)
			break
		}
	}
}

func TestSORT_WithEmptyValues(t *testing.T) {
	// Empty values should sort (empty compares as less than numbers in Excel)
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(3)},
		{EmptyVal()},
		{NumberVal(1)},
		{EmptyVal()},
		{NumberVal(2)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 5 {
		t.Fatalf("expected 5 rows, got %d", len(got.Array))
	}
	// Empty values should sort before numbers
	if got.Array[0][0].Type != ValueEmpty || got.Array[1][0].Type != ValueEmpty {
		t.Errorf("first two should be empty, got %v %v", got.Array[0][0], got.Array[1][0])
	}
	if got.Array[2][0].Num != 1 || got.Array[3][0].Num != 2 || got.Array[4][0].Num != 3 {
		t.Errorf("numbers should be sorted: got %v %v %v", got.Array[2][0], got.Array[3][0], got.Array[4][0])
	}
}

func TestSORT_DefaultArguments(t *testing.T) {
	// SORT with only the array argument — defaults: sort_index=1, sort_order=1 (asc)
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(30), StringVal("c")},
		{NumberVal(10), StringVal("a")},
		{NumberVal(20), StringVal("b")},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 ||
		got.Array[0][0].Num != 10 || got.Array[0][1].Str != "a" ||
		got.Array[1][0].Num != 20 || got.Array[1][1].Str != "b" ||
		got.Array[2][0].Num != 30 || got.Array[2][1].Str != "c" {
		t.Errorf("SORT defaults = %v, want sorted by col 1 asc", got)
	}
}

func TestSORT_ViaEvalBasic(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SORT({5;3;1;4;2})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 5 {
		t.Fatalf("expected 5-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	want := []float64{1, 2, 3, 4, 5}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestSORT_ViaEvalDescending(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SORT({5;3;1;4;2}, 1, -1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 5 {
		t.Fatalf("expected 5-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	want := []float64{5, 4, 3, 2, 1}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestSORT_ViaEvalByColumn(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(30), {Col: 2, Row: 1}: StringVal("c"),
			{Col: 1, Row: 2}: NumberVal(10), {Col: 2, Row: 2}: StringVal("a"),
			{Col: 1, Row: 3}: NumberVal(20), {Col: 2, Row: 3}: StringVal("b"),
		},
	}
	cf := evalCompile(t, "SORT(A1:B3, 2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3-row array, got type=%v rows=%d", got.Type, len(got.Array))
	}
	if got.Array[0][1].Str != "a" || got.Array[1][1].Str != "b" || got.Array[2][1].Str != "c" {
		t.Errorf("SORT by col 2 = %v, want sorted by string col", got)
	}
}

func TestSORT_StableSortMultiColumn(t *testing.T) {
	// Stable sort: ties in sort column preserve original order
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(2), StringVal("x")},
		{NumberVal(1), StringVal("a")},
		{NumberVal(2), StringVal("y")},
		{NumberVal(1), StringVal("b")},
		{NumberVal(2), StringVal("z")},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 5 {
		t.Fatalf("expected 5 rows, got %d", len(got.Array))
	}
	// 1s come first (in original relative order), then 2s
	if got.Array[0][1].Str != "a" || got.Array[1][1].Str != "b" {
		t.Errorf("1-group: got %v %v, want a b", got.Array[0][1], got.Array[1][1])
	}
	if got.Array[2][1].Str != "x" || got.Array[3][1].Str != "y" || got.Array[4][1].Str != "z" {
		t.Errorf("2-group: got %v %v %v, want x y z", got.Array[2][1], got.Array[3][1], got.Array[4][1])
	}
}

func TestSORT_StringsSorted(t *testing.T) {
	// Sort strings alphabetically
	arr := Value{Type: ValueArray, Array: [][]Value{
		{StringVal("delta")},
		{StringVal("alpha")},
		{StringVal("charlie")},
		{StringVal("bravo")},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 4 ||
		got.Array[0][0].Str != "alpha" ||
		got.Array[1][0].Str != "bravo" ||
		got.Array[2][0].Str != "charlie" ||
		got.Array[3][0].Str != "delta" {
		t.Errorf("SORT strings = %v, want alpha,bravo,charlie,delta", got)
	}
}

func TestSORT_MixedTypesDescending(t *testing.T) {
	// Mixed types descending: strings first (higher), then numbers
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(2)},
		{StringVal("b")},
		{NumberVal(1)},
		{StringVal("a")},
	}}
	got, err := fnSORT([]Value{arr, NumberVal(1), NumberVal(-1)})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 4 {
		t.Fatalf("expected 4 rows, got %d", len(got.Array))
	}
	// Descending: strings first (b, a), then numbers (2, 1)
	if got.Array[0][0].Str != "b" || got.Array[1][0].Str != "a" ||
		got.Array[2][0].Num != 2 || got.Array[3][0].Num != 1 {
		t.Errorf("SORT mixed desc = %v, want {b;a;2;1}", got)
	}
}

func TestSORT_ErrorInArray(t *testing.T) {
	// Error value as input argument propagates
	got, err := fnSORT([]Value{ErrorVal(ErrValREF)})
	if err != nil {
		t.Fatal(err)
	}
	// Non-array input is returned as-is
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("SORT(#REF!) = %v, want #REF!", got)
	}
}

func TestSORT_BoolInput(t *testing.T) {
	// Boolean input (non-array) is returned as-is
	got, err := fnSORT([]Value{BoolVal(true)})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("SORT(TRUE) = %v, want TRUE", got)
	}
}

func TestSORT_StringInput(t *testing.T) {
	// String input (non-array) is returned as-is
	got, err := fnSORT([]Value{StringVal("hello")})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueString || got.Str != "hello" {
		t.Errorf("SORT(\"hello\") = %v, want hello", got)
	}
}

func TestSORT_SortIndexZero(t *testing.T) {
	// sort_index = 0 → all comparisons use empty values, order preserved
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(3)},
		{NumberVal(1)},
		{NumberVal(2)},
	}}
	got, err := fnSORT([]Value{arr, NumberVal(0)})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	// With index 0, si = -1, so all values are empty and original order preserved
	if got.Array[0][0].Num != 3 || got.Array[1][0].Num != 1 || got.Array[2][0].Num != 2 {
		t.Errorf("SORT index=0 = %v, want original order {3;1;2}", got)
	}
}

func TestSORT_FloatingPointValues(t *testing.T) {
	arr := Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(3.14)},
		{NumberVal(2.71)},
		{NumberVal(1.41)},
		{NumberVal(1.73)},
	}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if got.Type != ValueArray || len(got.Array) != 4 ||
		got.Array[0][0].Num != 1.41 || got.Array[1][0].Num != 1.73 ||
		got.Array[2][0].Num != 2.71 || got.Array[3][0].Num != 3.14 {
		t.Errorf("SORT floats = %v, want {1.41;1.73;2.71;3.14}", got)
	}
}

func TestSORT_EmptyArrayInput(t *testing.T) {
	// Empty array input
	arr := Value{Type: ValueArray, Array: [][]Value{}}
	got, err := fnSORT([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	// Empty array: len(arr.Array) == 0, so returns arr as-is
	if got.Type != ValueArray || len(got.Array) != 0 {
		t.Errorf("SORT empty = %v, want empty array", got)
	}
}
