package formula

import (
	"testing"
)

func TestISEVEN(t *testing.T) {
	resolver := &mockResolver{}

	// ISEVEN(4) = TRUE
	cf := evalCompile(t, `ISEVEN(4)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("ISEVEN(4) = %v, want true", got)
	}

	// ISEVEN(3) = FALSE
	cf = evalCompile(t, `ISEVEN(3)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool {
		t.Errorf("ISEVEN(3) = %v, want false", got)
	}

	// ISEVEN(TRUE) = #VALUE!
	cf = evalCompile(t, `ISEVEN(TRUE)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISEVEN(TRUE) = %v, want #VALUE!", got)
	}

	// ISEVEN(FALSE) = #VALUE!
	cf = evalCompile(t, `ISEVEN(FALSE)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISEVEN(FALSE) = %v, want #VALUE!", got)
	}
}

func TestISODD(t *testing.T) {
	resolver := &mockResolver{}

	// ISODD(3) = TRUE
	cf := evalCompile(t, `ISODD(3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("ISODD(3) = %v, want true", got)
	}

	// ISODD(4) = FALSE
	cf = evalCompile(t, `ISODD(4)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool {
		t.Errorf("ISODD(4) = %v, want false", got)
	}

	// ISODD(TRUE) = #VALUE!
	cf = evalCompile(t, `ISODD(TRUE)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISODD(TRUE) = %v, want #VALUE!", got)
	}

	// ISODD(FALSE) = #VALUE!
	cf = evalCompile(t, `ISODD(FALSE)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISODD(FALSE) = %v, want #VALUE!", got)
	}
}

// mockFormulaResolver extends mockResolver with FormulaIntrospector support.
type mockFormulaResolver struct {
	mockResolver
	formulas map[CellAddr]string
}

func (m *mockFormulaResolver) HasFormula(sheet string, col, row int) bool {
	_, ok := m.formulas[CellAddr{Sheet: sheet, Col: col, Row: row}]
	return ok
}

func (m *mockFormulaResolver) GetFormulaText(sheet string, col, row int) string {
	return m.formulas[CellAddr{Sheet: sheet, Col: col, Row: row}]
}

func TestISFORMULA(t *testing.T) {
	resolver := &mockFormulaResolver{
		mockResolver: mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),  // A1: constant
				{Col: 2, Row: 1}: NumberVal(100), // B1: has formula
			},
		},
		formulas: map[CellAddr]string{
			{Col: 2, Row: 1}: "A1+58", // B1 has a formula
		},
	}
	ctx := &EvalContext{
		CurrentCol:   3,
		CurrentRow:   1,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	// ISFORMULA(B1) = TRUE (cell with formula)
	cf := evalCompile(t, `ISFORMULA(B1)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("ISFORMULA(B1) = %v, want TRUE", got)
	}

	// ISFORMULA(A1) = FALSE (constant value)
	cf = evalCompile(t, `ISFORMULA(A1)`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool {
		t.Errorf("ISFORMULA(A1) = %v, want FALSE", got)
	}

	// ISFORMULA(C1) = FALSE (empty cell)
	cf = evalCompile(t, `ISFORMULA(C1)`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool {
		t.Errorf("ISFORMULA(C1) = %v, want FALSE", got)
	}

	// ISFORMULA(123) = #VALUE! (non-reference argument)
	cf = evalCompile(t, `ISFORMULA(123)`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISFORMULA(123) = %v, want #VALUE!", got)
	}

	// ISFORMULA() with no args = #VALUE!
	cf = evalCompile(t, `ISFORMULA()`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISFORMULA() = %v, want #VALUE!", got)
	}
}

func TestFORMULATEXT(t *testing.T) {
	resolver := &mockFormulaResolver{
		mockResolver: mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),  // A1: constant
				{Col: 2, Row: 1}: NumberVal(100), // B1: has formula
			},
		},
		formulas: map[CellAddr]string{
			{Col: 2, Row: 1}: "A1+58", // B1 has a formula
		},
	}
	ctx := &EvalContext{
		CurrentCol:   3,
		CurrentRow:   1,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	// FORMULATEXT(B1) = "=A1+58" (cell with formula)
	cf := evalCompile(t, `FORMULATEXT(B1)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "=A1+58" {
		t.Errorf("FORMULATEXT(B1) = %v, want =A1+58", got)
	}

	// FORMULATEXT(A1) = #N/A (constant value, no formula)
	cf = evalCompile(t, `FORMULATEXT(A1)`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("FORMULATEXT(A1) = %v, want #N/A", got)
	}

	// FORMULATEXT(C1) = #N/A (empty cell)
	cf = evalCompile(t, `FORMULATEXT(C1)`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("FORMULATEXT(C1) = %v, want #N/A", got)
	}

	// FORMULATEXT(123) = #VALUE! (non-reference argument)
	cf = evalCompile(t, `FORMULATEXT(123)`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("FORMULATEXT(123) = %v, want #VALUE!", got)
	}

	// FORMULATEXT() with no args = #VALUE!
	cf = evalCompile(t, `FORMULATEXT()`)
	got, err = Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("FORMULATEXT() = %v, want #VALUE!", got)
	}
}

func TestISFORMULA_NoIntrospector(t *testing.T) {
	// When the resolver doesn't implement FormulaIntrospector, ISFORMULA returns FALSE.
	resolver := &mockResolver{}
	ctx := &EvalContext{
		CurrentCol:   1,
		CurrentRow:   1,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	cf := evalCompile(t, `ISFORMULA(A1)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool {
		t.Errorf("ISFORMULA(A1) with basic resolver = %v, want FALSE", got)
	}
}

func TestFORMULATEXT_NoIntrospector(t *testing.T) {
	// When the resolver doesn't implement FormulaIntrospector, FORMULATEXT returns #N/A.
	resolver := &mockResolver{}
	ctx := &EvalContext{
		CurrentCol:   1,
		CurrentRow:   1,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	cf := evalCompile(t, `FORMULATEXT(A1)`)
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("FORMULATEXT(A1) with basic resolver = %v, want #N/A", got)
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
