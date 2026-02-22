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

func TestEvalISFunctions(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(42),
			{Col: 1, Row: 2}: StringVal("hi"),
			// A3 is empty
		},
	}

	tests := []struct {
		formula string
		want    bool
	}{
		{"ISNUMBER(A1)", true},
		{"ISNUMBER(A2)", false},
		{"ISTEXT(A2)", true},
		{"ISTEXT(A1)", false},
		{"ISBLANK(A3)", true},
		{"ISBLANK(A1)", false},
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
