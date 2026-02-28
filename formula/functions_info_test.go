package formula

import (
	"testing"
)

func TestCOLUMN(t *testing.T) {
	resolver := &mockResolver{}
	ctx := &EvalContext{CurrentCol: 3, CurrentRow: 5}

	cf := evalCompile(t, "COLUMN()")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COLUMN() = %g, want 3", got.Num)
	}
}

func TestROW(t *testing.T) {
	resolver := &mockResolver{}
	ctx := &EvalContext{CurrentCol: 3, CurrentRow: 5}

	cf := evalCompile(t, "ROW()")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("ROW() = %g, want 5", got.Num)
	}
}

func TestCOLUMNS(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 3, Row: 1}: NumberVal(3),
		},
	}

	cf := evalCompile(t, "COLUMNS(A1:C1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("COLUMNS = %g, want 3", got.Num)
	}
}

func TestROWS(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
		},
	}

	cf := evalCompile(t, "ROWS(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("ROWS = %g, want 3", got.Num)
	}
}

func TestISEVEN(t *testing.T) {
	resolver := &mockResolver{}
	tests := []struct {
		expr string
		want bool
	}{
		{"ISEVEN(0)", true},
		{"ISEVEN(1)", false},
		{"ISEVEN(2)", true},
		{"ISEVEN(2.5)", true},
		{"ISEVEN(-1)", false},
		{"ISEVEN(-4)", true},
	}
	for _, tt := range tests {
		cf := evalCompile(t, tt.expr)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", tt.expr, err)
		}
		if got.Type != ValueBool || got.Bool != tt.want {
			t.Errorf("%s = %v, want %v", tt.expr, got, tt.want)
		}
	}

	// Non-numeric should return #VALUE!
	cf := evalCompile(t, `ISEVEN("abc")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf(`ISEVEN("abc") = %v, want #VALUE!`, got)
	}
}

func TestISODD(t *testing.T) {
	resolver := &mockResolver{}
	tests := []struct {
		expr string
		want bool
	}{
		{"ISODD(1)", true},
		{"ISODD(2)", false},
		{"ISODD(0)", false},
		{"ISODD(2.5)", false},
		{"ISODD(-3)", true},
	}
	for _, tt := range tests {
		cf := evalCompile(t, tt.expr)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", tt.expr, err)
		}
		if got.Type != ValueBool || got.Bool != tt.want {
			t.Errorf("%s = %v, want %v", tt.expr, got, tt.want)
		}
	}

	// Non-numeric should return #VALUE!
	cf := evalCompile(t, `ISODD("abc")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf(`ISODD("abc") = %v, want #VALUE!`, got)
	}
}

func TestISNA(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "ISNA(#N/A)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("ISNA(#N/A) = %v, want TRUE", got)
	}

	cf = evalCompile(t, "ISNA(42)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool {
		t.Errorf("ISNA(42) = %v, want FALSE", got)
	}
}

func TestIFNA(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `IFNA(#N/A,"default")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "default" {
		t.Errorf("IFNA(#N/A) = %v, want default", got)
	}

	cf = evalCompile(t, `IFNA(42,"default")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("IFNA(42) = %v, want 42", got)
	}
}

func TestISLOGICAL(t *testing.T) {
	resolver := &mockResolver{}
	tests := []struct {
		expr string
		want bool
	}{
		{"ISLOGICAL(TRUE)", true},
		{"ISLOGICAL(FALSE)", true},
		{`ISLOGICAL("TRUE")`, false},
		{"ISLOGICAL(1)", false},
		{"ISLOGICAL(0)", false},
		{`ISLOGICAL("")`, false},
		{"ISLOGICAL(1/0)", false},  // error value is not boolean
		{"ISLOGICAL(1>0)", true},   // comparison result is boolean
	}
	for _, tt := range tests {
		t.Run(tt.expr, func(t *testing.T) {
			cf := evalCompile(t, tt.expr)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tt.expr, err)
			}
			if got.Type != ValueBool || got.Bool != tt.want {
				t.Errorf("%s = %v, want %v", tt.expr, got, tt.want)
			}
		})
	}
}

func TestISNONTEXT(t *testing.T) {
	resolver := &mockResolver{}
	tests := []struct {
		expr string
		want bool
	}{
		{"ISNONTEXT(123)", true},
		{`ISNONTEXT("hello")`, false},
		{"ISNONTEXT(TRUE)", true},
		{"ISNONTEXT(1/0)", true},  // error is non-text
		{`ISNONTEXT("")`, false},  // empty string is still text type
	}
	for _, tt := range tests {
		t.Run(tt.expr, func(t *testing.T) {
			cf := evalCompile(t, tt.expr)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tt.expr, err)
			}
			if got.Type != ValueBool || got.Bool != tt.want {
				t.Errorf("%s = %v, want %v", tt.expr, got, tt.want)
			}
		})
	}
}

func TestERRORTYPE(t *testing.T) {
	resolver := &mockResolver{}

	// Test cases that return a number
	numTests := []struct {
		expr    string
		wantNum float64
	}{
		{"ERROR.TYPE(1/0)", 2},     // #DIV/0!
		{"ERROR.TYPE(#N/A)", 7},    // #N/A
		{"ERROR.TYPE(#VALUE!)", 3}, // #VALUE!
		{"ERROR.TYPE(#REF!)", 4},   // #REF!
		{"ERROR.TYPE(#NAME?)", 5},  // #NAME?
		{"ERROR.TYPE(#NUM!)", 6},   // #NUM!
		{"ERROR.TYPE(#NULL!)", 1},  // #NULL!
	}
	for _, tt := range numTests {
		t.Run(tt.expr, func(t *testing.T) {
			cf := evalCompile(t, tt.expr)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tt.expr, err)
			}
			if got.Type != ValueNumber || got.Num != tt.wantNum {
				t.Errorf("%s = %v, want %g", tt.expr, got, tt.wantNum)
			}
		})
	}

	// Test cases that return #N/A (non-error input)
	naTests := []struct {
		expr string
	}{
		{"ERROR.TYPE(123)"},
		{`ERROR.TYPE("hello")`},
	}
	for _, tt := range naTests {
		t.Run(tt.expr, func(t *testing.T) {
			cf := evalCompile(t, tt.expr)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tt.expr, err)
			}
			if got.Type != ValueError || got.Err != ErrValNA {
				t.Errorf("%s = %v, want #N/A", tt.expr, got)
			}
		})
	}
}

func TestNA(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "NA()")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("NA() = %v, want #N/A", got)
	}
}

func TestTYPE(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		expr    string
		wantNum float64
	}{
		{"TYPE(1)", 1},
		{"TYPE(1.5)", 1},
		{`TYPE("text")`, 2},
		{`TYPE("")`, 2},
		{"TYPE(TRUE)", 4},
		{"TYPE(FALSE)", 4},
		{"TYPE(1/0)", 16},
		{"TYPE(#N/A)", 16},
		{"TYPE({1,2,3})", 64},
	}
	for _, tt := range tests {
		cf := evalCompile(t, tt.expr)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", tt.expr, err)
		}
		if got.Type != ValueNumber || got.Num != tt.wantNum {
			t.Errorf("%s = %v, want %g", tt.expr, got, tt.wantNum)
		}
	}
}

func TestN(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		expr    string
		wantNum float64
	}{
		{"N(7)", 7},
		{`N("text")`, 0},
		{"N(TRUE)", 1},
		{"N(FALSE)", 0},
		{"N(0)", 0},
	}
	for _, tt := range tests {
		cf := evalCompile(t, tt.expr)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", tt.expr, err)
		}
		if got.Type != ValueNumber || got.Num != tt.wantNum {
			t.Errorf("%s = %v, want %g", tt.expr, got, tt.wantNum)
		}
	}

	// Error value should be returned as-is
	cf := evalCompile(t, "N(#N/A)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(N(#N/A)): %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("N(#N/A) = %v, want #N/A", got)
	}
}
