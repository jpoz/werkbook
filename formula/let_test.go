package formula

import (
	"reflect"
	"testing"
)

func TestLET(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10), // A1
			{Col: 2, Row: 1}: NumberVal(20), // B1
			{Col: 1, Row: 2}: NumberVal(30), // A2
			{Col: 1, Row: 3}: NumberVal(40), // A3
		},
	}

	tests := []struct {
		formula string
		want    Value
	}{
		// Basic LET: single binding
		{"LET(x, 5, x+1)", NumberVal(6)},
		// Name used multiple times in calculation
		{"LET(x, 3, x+x)", NumberVal(6)},
		// Multiple bindings
		{"LET(x, 5, y, x+1, y*2)", NumberVal(12)},
		// Later binding references earlier
		{"LET(x, 2, y, x*3, y+1)", NumberVal(7)},
		// Three bindings
		{"LET(a, 1, b, 2, c, 3, a+b+c)", NumberVal(6)},
		// Three bindings with forward reference chain
		{"LET(a, 10, b, a+5, c, b*2, c)", NumberVal(30)},
		// String values
		{`LET(x, "hello", UPPER(x))`, StringVal("HELLO")},
		// Boolean values
		{"LET(x, TRUE, IF(x, 1, 0))", NumberVal(1)},
		// LET with function calls in values
		{"LET(x, SUM(1,2,3), x*2)", NumberVal(12)},
		// Nested LET
		{"LET(x, LET(y, 5, y+1), x*2)", NumberVal(12)},
		// LET used inside other functions
		{"SUM(LET(x, 5, x), 1)", NumberVal(6)},
		// LET with cell references
		{"LET(x, A1, x+1)", NumberVal(11)},
		// LET with cell references in expressions
		{"LET(x, A1+B1, x*2)", NumberVal(60)},
		// Name that shadows a column letter (A is treated as the name, not column A)
		{"LET(A, 5, A+1)", NumberVal(6)},
		// Complex expressions in values and calculation
		{"LET(x, (2+3)*4, y, x/2, y+1)", NumberVal(11)},
		// LET with string concatenation
		{`LET(x, "hello", y, " world", x&y)`, StringVal("hello world")},
		// LET with unary minus
		{"LET(x, -5, x*-1)", NumberVal(5)},
		// LET with percentage
		{"LET(x, 50, x%)", NumberVal(0.5)},
		// Single binding with arithmetic
		{"LET(x, 10, x^2)", NumberVal(100)},
		// LET with comparison in calculation
		{"LET(x, 5, IF(x>3, \"big\", \"small\"))", StringVal("big")},
		// LET with nested function calls
		{"LET(x, 3, y, 4, MAX(x, y))", NumberVal(4)},
		// LET with _xlfn prefix
		{"_xlfn.LET(x, 5, x+1)", NumberVal(6)},
		// Four bindings
		{"LET(a, 1, b, a+1, c, b+1, d, c+1, d)", NumberVal(4)},
		// Binding value is zero
		{"LET(x, 0, x+1)", NumberVal(1)},
		// Binding value is negative
		{"LET(x, -10, ABS(x))", NumberVal(10)},
		// Long parameter names (previously broke because col > XFD was rejected at parse time)
		{"LET(result, 5, result+1)", NumberVal(6)},
		{"LET(total, 10, item, total+5, item)", NumberVal(15)},
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
}

func TestLETBoundLambdaInvocation(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: NumberVal(4),
		},
	}

	tests := []struct {
		formula string
		want    Value
	}{
		// LET-bound LAMBDA invoked by name
		{"LET(sq, LAMBDA(n, n*n), sq(A1) + sq(A2))", NumberVal(25)},
		// Currying via nested LAMBDA
		{"LET(mul, LAMBDA(a, LAMBDA(b, a*b)), double, mul(2), double(A2))", NumberVal(8)},
		// Long parameter names
		{"LET(cube, LAMBDA(n, n*n*n), cube(A1))", NumberVal(27)},
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
}

func TestLETErrors(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
	}{
		// Even number of args (2 args: name + value, no calculation)
		{"two_args", "LET(x, 5)"},
		// Less than 3 args (1 arg)
		{"one_arg", "LET(x)"},
		// Even number of args (4 args)
		{"four_args", "LET(x, 5, y, 10)"},
		// Invalid name: number literal as name
		{"number_as_name", "LET(5, 5, 10)"},
		// Invalid name: string literal as name
		{"string_as_name", `LET("x", 5, 10)`},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != ErrValVALUE {
				t.Fatalf("Eval(%q) = %#v, want #VALUE!", tt.formula, got)
			}
		})
	}
}

func TestLETParseAST(t *testing.T) {
	// LET is desugared at parse time, so the resulting AST should not contain a FuncCall for LET.
	tests := []struct {
		formula string
		want    string
	}{
		// Basic: LET(x, 5, x+1) => (+ 5 1)
		{"LET(x, 5, x+1)", "(+ 5 1)"},
		// Multiple bindings: LET(x, 5, y, x+1, y*2) => (* (+ 5 1) 2)
		{"LET(x, 5, y, x+1, y*2)", "(* (+ 5 1) 2)"},
		// Simple passthrough: LET(x, 42, x) => 42
		{"LET(x, 42, x)", "42"},
		// Long parameter names
		{"LET(result, 5, result+1)", "(+ 5 1)"},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			node, err := Parse(tt.formula)
			if err != nil {
				t.Fatalf("Parse(%q): %v", tt.formula, err)
			}
			got := node.String()
			if got != tt.want {
				t.Errorf("Parse(%q)\n  got:  %s\n  want: %s", tt.formula, got, tt.want)
			}
		})
	}
}

func TestLETErrorPropagation(t *testing.T) {
	resolver := &mockResolver{}

	// LET with a #DIV/0! in value - the error should propagate
	cf := evalCompile(t, "LET(x, 1/0, x+1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Fatalf("LET(x, 1/0, x+1) = %#v, want #DIV/0!", got)
	}
}

func TestLETWithIFERROR(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `IFERROR(LET(x, 1/0, x+1), "safe")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "safe" {
		t.Fatalf(`IFERROR(LET(x, 1/0, x+1), "safe") = %#v, want "safe"`, got)
	}
}

func TestLETWithRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
		},
	}

	cf := evalCompile(t, "LET(x, SUM(A1:A3), x*2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 12 {
		t.Fatalf("LET(x, SUM(A1:A3), x*2) = %#v, want 12", got)
	}
}
