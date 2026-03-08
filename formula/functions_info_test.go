package formula

import (
	"testing"
)


func TestISBLANK(t *testing.T) {
	t.Run("empty_cell_reference", func(t *testing.T) {
		// ISBLANK should return TRUE for a reference to an empty cell.
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}

		// A1 is not set in the resolver, so GetCellValue returns EmptyVal()
		cf := evalCompile(t, `ISBLANK(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISBLANK(A1) empty cell = %v, want TRUE", got)
		}
	})

	t.Run("non_empty_cell_with_number", func(t *testing.T) {
		// ISBLANK should return FALSE for a cell containing a number.
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}

		cf := evalCompile(t, `ISBLANK(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISBLANK(A1) with number = %v, want FALSE", got)
		}
	})

	t.Run("non_empty_cell_with_text", func(t *testing.T) {
		// ISBLANK should return FALSE for a cell containing text.
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("hello"),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}

		cf := evalCompile(t, `ISBLANK(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISBLANK(A1) with text = %v, want FALSE", got)
		}
	})

	t.Run("non_empty_cell_with_boolean", func(t *testing.T) {
		// ISBLANK should return FALSE for a cell containing a boolean.
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}

		cf := evalCompile(t, `ISBLANK(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISBLANK(A1) with TRUE = %v, want FALSE", got)
		}
	})

	t.Run("empty_string_literal_not_blank", func(t *testing.T) {
		// In Excel, ISBLANK("") returns FALSE. An empty string is NOT the same as blank.
		resolver := &mockResolver{}

		cf := evalCompile(t, `ISBLANK("")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf(`ISBLANK("") = %v, want FALSE`, got)
		}
	})

	t.Run("number_zero_not_blank", func(t *testing.T) {
		// ISBLANK(0) should return FALSE.
		resolver := &mockResolver{}

		cf := evalCompile(t, `ISBLANK(0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISBLANK(0) = %v, want FALSE", got)
		}
	})

	t.Run("positive_number_not_blank", func(t *testing.T) {
		// ISBLANK(1) should return FALSE.
		resolver := &mockResolver{}

		cf := evalCompile(t, `ISBLANK(1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISBLANK(1) = %v, want FALSE", got)
		}
	})

	t.Run("negative_number_not_blank", func(t *testing.T) {
		// ISBLANK(-5) should return FALSE.
		resolver := &mockResolver{}

		cf := evalCompile(t, `ISBLANK(-5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISBLANK(-5) = %v, want FALSE", got)
		}
	})

	t.Run("text_string_not_blank", func(t *testing.T) {
		// ISBLANK("text") should return FALSE.
		resolver := &mockResolver{}

		cf := evalCompile(t, `ISBLANK("text")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf(`ISBLANK("text") = %v, want FALSE`, got)
		}
	})

	t.Run("boolean_true_not_blank", func(t *testing.T) {
		// ISBLANK(TRUE) should return FALSE.
		resolver := &mockResolver{}

		cf := evalCompile(t, `ISBLANK(TRUE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISBLANK(TRUE) = %v, want FALSE", got)
		}
	})

	t.Run("boolean_false_not_blank", func(t *testing.T) {
		// ISBLANK(FALSE) should return FALSE.
		resolver := &mockResolver{}

		cf := evalCompile(t, `ISBLANK(FALSE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISBLANK(FALSE) = %v, want FALSE", got)
		}
	})

	t.Run("decimal_number_not_blank", func(t *testing.T) {
		// ISBLANK(3.14) should return FALSE.
		resolver := &mockResolver{}

		cf := evalCompile(t, `ISBLANK(3.14)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISBLANK(3.14) = %v, want FALSE", got)
		}
	})

	t.Run("string_with_space_not_blank", func(t *testing.T) {
		// ISBLANK(" ") should return FALSE - a space is text, not blank.
		resolver := &mockResolver{}

		cf := evalCompile(t, `ISBLANK(" ")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf(`ISBLANK(" ") = %v, want FALSE`, got)
		}
	})

	t.Run("cell_with_zero_not_blank", func(t *testing.T) {
		// A cell containing 0 is not blank.
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(0),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}

		cf := evalCompile(t, `ISBLANK(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISBLANK(A1) with 0 = %v, want FALSE", got)
		}
	})

	t.Run("cell_with_empty_string_not_blank", func(t *testing.T) {
		// A cell containing an empty string is not blank.
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal(""),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}

		cf := evalCompile(t, `ISBLANK(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf(`ISBLANK(A1) with "" = %v, want FALSE`, got)
		}
	})

	t.Run("cell_with_error_not_blank", func(t *testing.T) {
		// A cell containing an error value is not blank.
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNA),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}

		cf := evalCompile(t, `ISBLANK(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISBLANK(A1) with #N/A = %v, want FALSE", got)
		}
	})

	t.Run("multiple_empty_cells", func(t *testing.T) {
		// Verify ISBLANK returns TRUE for different empty cell references.
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10), // A1 has data
			},
		}
		ctx := &EvalContext{
			CurrentCol:   5,
			CurrentRow:   5,
			CurrentSheet: "",
			Resolver:     resolver,
		}

		// B1 is empty
		cf := evalCompile(t, `ISBLANK(B1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISBLANK(B1) empty = %v, want TRUE", got)
		}

		// C5 is empty
		cf = evalCompile(t, `ISBLANK(C5)`)
		got, err = Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISBLANK(C5) empty = %v, want TRUE", got)
		}

		// Z100 is empty
		cf = evalCompile(t, `ISBLANK(Z100)`)
		got, err = Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISBLANK(Z100) empty = %v, want TRUE", got)
		}
	})

	t.Run("expression_result_not_blank", func(t *testing.T) {
		// ISBLANK(1+1) should return FALSE since the result is a number.
		resolver := &mockResolver{}

		cf := evalCompile(t, `ISBLANK(1+1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISBLANK(1+1) = %v, want FALSE", got)
		}
	})

	t.Run("no_args_returns_error", func(t *testing.T) {
		// ISBLANK() with no arguments should return #VALUE!
		resolver := &mockResolver{}

		cf := evalCompile(t, `ISBLANK()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISBLANK() = %v, want #VALUE!", got)
		}
	})

	t.Run("too_many_args_returns_error", func(t *testing.T) {
		// ISBLANK(A1, B1) with two arguments should return #VALUE!
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   3,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}

		cf := evalCompile(t, `ISBLANK(A1,B1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISBLANK(A1,B1) = %v, want #VALUE!", got)
		}
	})

	t.Run("large_number_not_blank", func(t *testing.T) {
		// ISBLANK(999999) should return FALSE.
		resolver := &mockResolver{}

		cf := evalCompile(t, `ISBLANK(999999)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISBLANK(999999) = %v, want FALSE", got)
		}
	})
}

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

func TestCOLUMN(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("no_args_returns_current_col", func(t *testing.T) {
		tests := []struct {
			name string
			col  int
			want float64
		}{
			{"col_1", 1, 1},
			{"col_3", 3, 3},
			{"col_10", 10, 10},
			{"col_26_Z", 26, 26},
			{"col_256", 256, 256},
			{"col_16384_XFD", 16384, 16384},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				ctx := &EvalContext{
					CurrentCol:   tt.col,
					CurrentRow:   1,
					CurrentSheet: "",
					Resolver:     resolver,
				}
				cf := evalCompile(t, `COLUMN()`)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("COLUMN() with CurrentCol=%d = %v, want %v", tt.col, got, tt.want)
				}
			})
		}
	})

	t.Run("no_args_nil_context", func(t *testing.T) {
		// COLUMN() with no EvalContext should return #VALUE!
		cf := evalCompile(t, `COLUMN()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("COLUMN() with nil ctx = %v, want #VALUE!", got)
		}
	})

	t.Run("single_cell_ref", func(t *testing.T) {
		tests := []struct {
			name    string
			formula string
			want    float64
		}{
			{"A1_col_1", `COLUMN(A1)`, 1},
			{"B1_col_2", `COLUMN(B1)`, 2},
			{"C5_col_3", `COLUMN(C5)`, 3},
			{"Z1_col_26", `COLUMN(Z1)`, 26},
			{"AA1_col_27", `COLUMN(AA1)`, 27},
			{"AZ1_col_52", `COLUMN(AZ1)`, 52},
			{"BA1_col_53", `COLUMN(BA1)`, 53},
			{"IV1_col_256", `COLUMN(IV1)`, 256},
			{"XFD1_col_16384", `COLUMN(XFD1)`, 16384},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				ctx := &EvalContext{
					CurrentCol:   1,
					CurrentRow:   1,
					CurrentSheet: "",
					Resolver:     resolver,
				}
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("%s = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("ref_different_rows_same_col", func(t *testing.T) {
		// COLUMN should return the column regardless of which row the ref is in
		tests := []struct {
			name    string
			formula string
			want    float64
		}{
			{"D1", `COLUMN(D1)`, 4},
			{"D10", `COLUMN(D10)`, 4},
			{"D100", `COLUMN(D100)`, 4},
			{"D1000", `COLUMN(D1000)`, 4},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				ctx := &EvalContext{
					CurrentCol:   1,
					CurrentRow:   1,
					CurrentSheet: "",
					Resolver:     resolver,
				}
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("%s = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("non_reference_arg_returns_VALUE_error", func(t *testing.T) {
		// Non-reference arguments (numbers, strings, booleans) should return #VALUE!
		tests := []struct {
			name    string
			formula string
		}{
			{"number", `COLUMN(42)`},
			{"string", `COLUMN("hello")`},
			{"boolean_true", `COLUMN(TRUE)`},
			{"boolean_false", `COLUMN(FALSE)`},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				ctx := &EvalContext{
					CurrentCol:   1,
					CurrentRow:   1,
					CurrentSheet: "",
					Resolver:     resolver,
				}
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError || got.Err != ErrValVALUE {
					t.Errorf("%s = %v, want #VALUE!", tt.formula, got)
				}
			})
		}
	})

	t.Run("range_ref_returns_leftmost_column", func(t *testing.T) {
		// When a range is passed, COLUMN returns the leftmost column.
		// In the current implementation, a range resolves to a ValueArray
		// which is not ValueRef, so it returns #VALUE!.
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `COLUMN(A1:C3)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("COLUMN(A1:C3) = %v, want #VALUE!", got)
		}
	})

	t.Run("absolute_ref", func(t *testing.T) {
		// Absolute references ($C$1) should work identically
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `COLUMN($C$1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("COLUMN($C$1) = %v, want 3", got)
		}
	})

	t.Run("absolute_col_only", func(t *testing.T) {
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `COLUMN($E1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("COLUMN($E1) = %v, want 5", got)
		}
	})

	t.Run("absolute_row_only", func(t *testing.T) {
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `COLUMN(F$1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 6 {
			t.Errorf("COLUMN(F$1) = %v, want 6", got)
		}
	})
}

func TestISTEXT(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    bool
	}{
		// String values → TRUE
		{"simple_string", `ISTEXT("hello")`, true},
		{"empty_string", `ISTEXT("")`, true},
		{"string_with_spaces", `ISTEXT("  ")`, true},
		{"numeric_string", `ISTEXT("19")`, true},
		{"string_true", `ISTEXT("TRUE")`, true},
		{"string_false", `ISTEXT("FALSE")`, true},
		{"string_with_number", `ISTEXT("123abc")`, true},
		{"single_char", `ISTEXT("A")`, true},
		{"string_with_special_chars", `ISTEXT("hello world!")`, true},

		// Numeric values → FALSE
		{"integer", `ISTEXT(1)`, false},
		{"zero", `ISTEXT(0)`, false},
		{"negative", `ISTEXT(-1)`, false},
		{"decimal", `ISTEXT(3.14)`, false},
		{"large_number", `ISTEXT(1000000)`, false},

		// Boolean values → FALSE
		{"true", `ISTEXT(TRUE)`, false},
		{"false", `ISTEXT(FALSE)`, false},

		// Expressions resulting in non-text → FALSE
		{"addition", `ISTEXT(1+1)`, false},
		{"comparison", `ISTEXT(1>0)`, false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueBool {
				t.Fatalf("%s = %v (type %d), want bool", tt.formula, got, got.Type)
			}
			if got.Bool != tt.want {
				t.Errorf("%s = %v, want %v", tt.formula, got.Bool, tt.want)
			}
		})
	}

	// Error values → FALSE (errors are not text)
	t.Run("error_div0", func(t *testing.T) {
		cf := evalCompile(t, `ISTEXT(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("ISTEXT(1/0) = %v, want FALSE", got)
		}
	})

	t.Run("error_NA", func(t *testing.T) {
		cf := evalCompile(t, `ISTEXT(#N/A)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("ISTEXT(#N/A) = %v, want FALSE", got)
		}
	})

	t.Run("error_VALUE", func(t *testing.T) {
		cf := evalCompile(t, `ISTEXT(#VALUE!)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("ISTEXT(#VALUE!) = %v, want FALSE", got)
		}
	})

	// Wrong argument count → #VALUE!
	t.Run("no_args", func(t *testing.T) {
		cf := evalCompile(t, `ISTEXT()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISTEXT() = %v, want #VALUE!", got)
		}
	})

	t.Run("too_many_args", func(t *testing.T) {
		cf := evalCompile(t, `ISTEXT("a","b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISTEXT(\"a\",\"b\") = %v, want #VALUE!", got)
		}
	})

	// Concatenation result is text → TRUE
	t.Run("concatenation", func(t *testing.T) {
		cf := evalCompile(t, `ISTEXT("hello"&" world")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf(`ISTEXT("hello"&" world") = %v, want TRUE`, got)
		}
	})
}

func TestISNONTEXT(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    bool
	}{
		// Numeric values → TRUE (numbers are non-text)
		{"integer", `ISNONTEXT(1)`, true},
		{"zero", `ISNONTEXT(0)`, true},
		{"negative", `ISNONTEXT(-1)`, true},
		{"decimal", `ISNONTEXT(3.14)`, true},
		{"large_number", `ISNONTEXT(1000000)`, true},

		// Boolean values → TRUE (booleans are non-text)
		{"true", `ISNONTEXT(TRUE)`, true},
		{"false", `ISNONTEXT(FALSE)`, true},

		// String values → FALSE (text is NOT non-text)
		{"simple_string", `ISNONTEXT("hello")`, false},
		{"empty_string", `ISNONTEXT("")`, false},
		{"string_with_spaces", `ISNONTEXT("  ")`, false},
		{"numeric_string", `ISNONTEXT("19")`, false},
		{"string_true", `ISNONTEXT("TRUE")`, false},
		{"string_false", `ISNONTEXT("FALSE")`, false},
		{"string_with_number", `ISNONTEXT("123abc")`, false},
		{"single_char", `ISNONTEXT("A")`, false},

		// Expressions resulting in non-text → TRUE
		{"addition", `ISNONTEXT(1+1)`, true},
		{"comparison", `ISNONTEXT(1>0)`, true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueBool {
				t.Fatalf("%s = %v (type %d), want bool", tt.formula, got, got.Type)
			}
			if got.Bool != tt.want {
				t.Errorf("%s = %v, want %v", tt.formula, got.Bool, tt.want)
			}
		})
	}

	// Error values → TRUE (errors are non-text)
	t.Run("error_div0", func(t *testing.T) {
		cf := evalCompile(t, `ISNONTEXT(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("ISNONTEXT(1/0) = %v, want TRUE", got)
		}
	})

	t.Run("error_NA", func(t *testing.T) {
		cf := evalCompile(t, `ISNONTEXT(#N/A)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("ISNONTEXT(#N/A) = %v, want TRUE", got)
		}
	})

	t.Run("error_VALUE", func(t *testing.T) {
		cf := evalCompile(t, `ISNONTEXT(#VALUE!)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("ISNONTEXT(#VALUE!) = %v, want TRUE", got)
		}
	})

	// Blank cell → TRUE (Excel docs: "this function returns TRUE if the value refers to a blank cell")
	t.Run("blank_cell", func(t *testing.T) {
		blankResolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     blankResolver,
		}

		cf := evalCompile(t, `ISNONTEXT(A1)`)
		got, err := Eval(cf, blankResolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("ISNONTEXT(A1) blank cell = %v, want TRUE", got)
		}
	})

	// Cell containing text → FALSE
	t.Run("cell_with_text", func(t *testing.T) {
		textResolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("hello"),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     textResolver,
		}

		cf := evalCompile(t, `ISNONTEXT(A1)`)
		got, err := Eval(cf, textResolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("ISNONTEXT(A1) with text = %v, want FALSE", got)
		}
	})

	// Cell containing number → TRUE
	t.Run("cell_with_number", func(t *testing.T) {
		numResolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     numResolver,
		}

		cf := evalCompile(t, `ISNONTEXT(A1)`)
		got, err := Eval(cf, numResolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("ISNONTEXT(A1) with number = %v, want TRUE", got)
		}
	})

	// Concatenation result is text → FALSE
	t.Run("concatenation", func(t *testing.T) {
		cf := evalCompile(t, `ISNONTEXT("hello"&" world")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf(`ISNONTEXT("hello"&" world") = %v, want FALSE`, got)
		}
	})

	// Wrong argument count → #VALUE!
	t.Run("no_args", func(t *testing.T) {
		cf := evalCompile(t, `ISNONTEXT()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISNONTEXT() = %v, want #VALUE!", got)
		}
	})

	t.Run("too_many_args", func(t *testing.T) {
		cf := evalCompile(t, `ISNONTEXT("a","b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISNONTEXT(\"a\",\"b\") = %v, want #VALUE!", got)
		}
	})
}

func TestCOLUMNS(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),  // A1
			{Col: 2, Row: 1}: NumberVal(20),  // B1
			{Col: 3, Row: 1}: NumberVal(30),  // C1
			{Col: 4, Row: 1}: NumberVal(40),  // D1
			{Col: 5, Row: 1}: NumberVal(50),  // E1
			{Col: 1, Row: 2}: NumberVal(100), // A2
			{Col: 2, Row: 2}: NumberVal(200), // B2
			{Col: 3, Row: 2}: NumberVal(300), // C2
			{Col: 4, Row: 2}: NumberVal(400), // D2
			{Col: 5, Row: 2}: NumberVal(500), // E2
			{Col: 1, Row: 3}: NumberVal(1),   // A3
			{Col: 2, Row: 3}: NumberVal(2),   // B3
			{Col: 3, Row: 3}: NumberVal(3),   // C3
			{Col: 4, Row: 3}: NumberVal(4),   // D3
			{Col: 5, Row: 3}: NumberVal(5),   // E3
			{Col: 1, Row: 4}: NumberVal(6),   // A4
			{Col: 1, Row: 5}: NumberVal(7),   // A5
		},
	}
	ctx := &EvalContext{
		CurrentCol:   1,
		CurrentRow:   1,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	t.Run("range_references", func(t *testing.T) {
		tests := []struct {
			name    string
			formula string
			want    float64
		}{
			// Basic range references
			{"three_cols_A1_C1", `COLUMNS(A1:C1)`, 3},
			{"single_col_A1_A1", `COLUMNS(A1:A1)`, 1},
			{"five_cols_A1_E1", `COLUMNS(A1:E1)`, 5},
			{"five_cols_A1_E5", `COLUMNS(A1:E5)`, 5},
			{"two_cols_B1_C1", `COLUMNS(B1:C1)`, 2},
			{"four_cols_B2_E2", `COLUMNS(B2:E2)`, 4},

			// Multi-row ranges (column count should be based on columns, not rows)
			{"three_cols_A1_C3", `COLUMNS(A1:C3)`, 3},
			{"two_cols_A1_B5", `COLUMNS(A1:B5)`, 2},

			// Single column, multiple rows
			{"single_col_A1_A5", `COLUMNS(A1:A5)`, 1},

			// Large range
			{"five_cols_A1_E3", `COLUMNS(A1:E3)`, 5},

			// Range starting from non-A column
			{"three_cols_C1_E1", `COLUMNS(C1:E1)`, 3},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("%s = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("single_cell_reference", func(t *testing.T) {
		// A single cell reference resolves to a scalar, so COLUMNS returns 1.
		tests := []struct {
			name    string
			formula string
		}{
			{"A1", `COLUMNS(A1)`},
			{"B1", `COLUMNS(B1)`},
			{"E5", `COLUMNS(E5)`},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != 1 {
					t.Errorf("%s = %v, want 1", tt.formula, got)
				}
			})
		}
	})

	t.Run("non_array_arguments", func(t *testing.T) {
		// Non-array values (numbers, strings, booleans) should return 1.
		tests := []struct {
			name    string
			formula string
		}{
			{"number", `COLUMNS(42)`},
			{"string", `COLUMNS("hello")`},
			{"boolean_true", `COLUMNS(TRUE)`},
			{"boolean_false", `COLUMNS(FALSE)`},
			{"zero", `COLUMNS(0)`},
			{"negative", `COLUMNS(-1)`},
			{"decimal", `COLUMNS(3.14)`},
			{"empty_string", `COLUMNS("")`},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != 1 {
					t.Errorf("%s = %v, want 1", tt.formula, got)
				}
			})
		}
	})

	t.Run("wrong_arg_count", func(t *testing.T) {
		// COLUMNS requires exactly 1 argument.
		tests := []struct {
			name    string
			formula string
		}{
			{"no_args", `COLUMNS()`},
			{"two_args", `COLUMNS(A1,B1)`},
			{"three_args", `COLUMNS(A1,B1,C1)`},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError || got.Err != ErrValVALUE {
					t.Errorf("%s = %v, want #VALUE!", tt.formula, got)
				}
			})
		}
	})

	t.Run("nil_context", func(t *testing.T) {
		// COLUMNS with a scalar should work with nil context
		// because fnCOLUMNS is a NoCtx function.
		cf := evalCompile(t, `COLUMNS(42)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("COLUMNS(42) with nil ctx = %v, want 1", got)
		}
	})

	t.Run("absolute_refs", func(t *testing.T) {
		tests := []struct {
			name    string
			formula string
			want    float64
		}{
			{"absolute_range", `COLUMNS($A$1:$C$1)`, 3},
			{"mixed_absolute", `COLUMNS($A1:C$3)`, 3},
			{"absolute_col", `COLUMNS($B1:$D1)`, 3},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("%s = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("expression_result", func(t *testing.T) {
		// When COLUMNS receives a result from another function that
		// produces a scalar, it should return 1.
		cf := evalCompile(t, `COLUMNS(SUM(A1:C1))`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("COLUMNS(SUM(A1:C1)) = %v, want 1", got)
		}
	})
}

func TestROWS(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),  // A1
			{Col: 2, Row: 1}: NumberVal(20),  // B1
			{Col: 3, Row: 1}: NumberVal(30),  // C1
			{Col: 1, Row: 2}: NumberVal(100), // A2
			{Col: 2, Row: 2}: NumberVal(200), // B2
			{Col: 3, Row: 2}: NumberVal(300), // C2
			{Col: 1, Row: 3}: NumberVal(1),   // A3
			{Col: 2, Row: 3}: NumberVal(2),   // B3
			{Col: 3, Row: 3}: NumberVal(3),   // C3
			{Col: 1, Row: 4}: NumberVal(4),   // A4
			{Col: 1, Row: 5}: NumberVal(5),   // A5
		},
	}
	ctx := &EvalContext{
		CurrentCol:   1,
		CurrentRow:   1,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	t.Run("range_references", func(t *testing.T) {
		tests := []struct {
			name    string
			formula string
			want    float64
		}{
			// Basic range references
			{"five_rows_A1_A5", `ROWS(A1:A5)`, 5},
			{"three_rows_A1_C3", `ROWS(A1:C3)`, 3},
			{"single_row_A1_A1", `ROWS(A1:A1)`, 1},
			{"single_row_A1_C1", `ROWS(A1:C1)`, 1},
			{"two_rows_A1_C2", `ROWS(A1:C2)`, 2},
			{"four_rows_A1_A4", `ROWS(A1:A4)`, 4},

			// Multi-column ranges (row count based on rows, not columns)
			{"three_rows_A1_B3", `ROWS(A1:B3)`, 3},
			{"two_rows_B1_C2", `ROWS(B1:C2)`, 2},

			// Range starting from non-row-1
			{"two_rows_A2_A3", `ROWS(A2:A3)`, 2},
			{"three_rows_B2_C4", `ROWS(B2:C4)`, 3},

			// Excel documentation example: ROWS(C1:E4) = 4
			{"excel_doc_example", `ROWS(C1:E4)`, 4},

			// Large range
			{"hundred_rows_A1_A100", `ROWS(A1:A100)`, 100},
			{"thousand_rows_A1_B1000", `ROWS(A1:B1000)`, 1000},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("%s = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("single_cell_reference", func(t *testing.T) {
		// A single cell reference resolves to a scalar, so ROWS returns 1.
		tests := []struct {
			name    string
			formula string
		}{
			{"A1", `ROWS(A1)`},
			{"B2", `ROWS(B2)`},
			{"C3", `ROWS(C3)`},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != 1 {
					t.Errorf("%s = %v, want 1", tt.formula, got)
				}
			})
		}
	})

	t.Run("non_array_arguments", func(t *testing.T) {
		// Non-array values (numbers, strings, booleans) should return 1.
		tests := []struct {
			name    string
			formula string
		}{
			{"number", `ROWS(42)`},
			{"string", `ROWS("hello")`},
			{"boolean_true", `ROWS(TRUE)`},
			{"boolean_false", `ROWS(FALSE)`},
			{"zero", `ROWS(0)`},
			{"negative", `ROWS(-1)`},
			{"decimal", `ROWS(3.14)`},
			{"empty_string", `ROWS("")`},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != 1 {
					t.Errorf("%s = %v, want 1", tt.formula, got)
				}
			})
		}
	})

	t.Run("wrong_arg_count", func(t *testing.T) {
		// ROWS requires exactly 1 argument.
		tests := []struct {
			name    string
			formula string
		}{
			{"no_args", `ROWS()`},
			{"two_args", `ROWS(A1,B1)`},
			{"three_args", `ROWS(A1,B1,C1)`},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError || got.Err != ErrValVALUE {
					t.Errorf("%s = %v, want #VALUE!", tt.formula, got)
				}
			})
		}
	})

	t.Run("nil_context", func(t *testing.T) {
		// ROWS with a scalar should work with nil context
		// because fnROWS is a NoCtx function.
		cf := evalCompile(t, `ROWS(42)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("ROWS(42) with nil ctx = %v, want 1", got)
		}
	})

	t.Run("absolute_refs", func(t *testing.T) {
		tests := []struct {
			name    string
			formula string
			want    float64
		}{
			{"absolute_range", `ROWS($A$1:$A$3)`, 3},
			{"mixed_absolute", `ROWS($A1:C$3)`, 3},
			{"absolute_row", `ROWS(A$1:A$5)`, 5},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("%s = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("expression_result", func(t *testing.T) {
		// When ROWS receives a result from another function that
		// produces a scalar, it should return 1.
		cf := evalCompile(t, `ROWS(SUM(A1:C1))`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("ROWS(SUM(A1:C1)) = %v, want 1", got)
		}
	})
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

func TestISERR(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name     string
		formula  string
		wantTyp  ValueType
		wantBool bool
		wantErr  ErrorValue
	}{
		// Error types (except #N/A) should return TRUE
		{"DIV0_error", `ISERR(1/0)`, ValueBool, true, 0},
		{"VALUE_error", `ISERR(#VALUE!)`, ValueBool, true, 0},
		{"REF_error", `ISERR(#REF!)`, ValueBool, true, 0},
		{"NUM_error", `ISERR(#NUM!)`, ValueBool, true, 0},
		{"NULL_error", `ISERR(#NULL!)`, ValueBool, true, 0},
		{"NAME_error", `ISERR(#NAME?)`, ValueBool, true, 0},

		// #N/A must return FALSE — this is the key difference from ISERROR
		{"NA_function", `ISERR(NA())`, ValueBool, false, 0},
		{"NA_literal", `ISERR(#N/A)`, ValueBool, false, 0},

		// Non-error values should return FALSE
		{"number_positive", `ISERR(1)`, ValueBool, false, 0},
		{"number_zero", `ISERR(0)`, ValueBool, false, 0},
		{"number_negative", `ISERR(-5)`, ValueBool, false, 0},
		{"number_decimal", `ISERR(3.14)`, ValueBool, false, 0},
		{"text", `ISERR("text")`, ValueBool, false, 0},
		{"empty_string", `ISERR("")`, ValueBool, false, 0},
		{"bool_true", `ISERR(TRUE)`, ValueBool, false, 0},
		{"bool_false", `ISERR(FALSE)`, ValueBool, false, 0},

		// Expressions that produce errors
		{"div_by_zero_expr", `ISERR(0/0)`, ValueBool, true, 0},

		// Expressions that do not produce errors
		{"valid_arithmetic", `ISERR(2+3)`, ValueBool, false, 0},
		{"valid_division", `ISERR(10/2)`, ValueBool, false, 0},

		// Wrong number of arguments returns #VALUE!
		{"no_args", `ISERR()`, ValueError, false, ErrValVALUE},
		{"two_args", `ISERR(1,2)`, ValueError, false, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantTyp == ValueBool {
				if got.Type != ValueBool || got.Bool != tt.wantBool {
					t.Errorf("%s = %v, want bool %v", tt.formula, got, tt.wantBool)
				}
			} else {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("%s = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			}
		})
	}
}

func TestISERROR(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name     string
		formula  string
		wantTyp  ValueType
		wantBool bool
		wantErr  ErrorValue
	}{
		// All error types should return TRUE
		{"DIV0_error", `ISERROR(1/0)`, ValueBool, true, 0},
		{"NA_function", `ISERROR(NA())`, ValueBool, true, 0},
		{"NA_literal", `ISERROR(#N/A)`, ValueBool, true, 0},
		{"VALUE_error", `ISERROR(#VALUE!)`, ValueBool, true, 0},
		{"REF_error", `ISERROR(#REF!)`, ValueBool, true, 0},
		{"NUM_error", `ISERROR(#NUM!)`, ValueBool, true, 0},
		{"NULL_error", `ISERROR(#NULL!)`, ValueBool, true, 0},
		{"NAME_error", `ISERROR(#NAME?)`, ValueBool, true, 0},

		// Non-error values should return FALSE
		{"number_positive", `ISERROR(1)`, ValueBool, false, 0},
		{"number_zero", `ISERROR(0)`, ValueBool, false, 0},
		{"number_negative", `ISERROR(-5)`, ValueBool, false, 0},
		{"number_decimal", `ISERROR(3.14)`, ValueBool, false, 0},
		{"text", `ISERROR("text")`, ValueBool, false, 0},
		{"empty_string", `ISERROR("")`, ValueBool, false, 0},
		{"bool_true", `ISERROR(TRUE)`, ValueBool, false, 0},
		{"bool_false", `ISERROR(FALSE)`, ValueBool, false, 0},

		// Expressions that produce errors
		{"div_by_zero_expr", `ISERROR(0/0)`, ValueBool, true, 0},

		// Expressions that do not produce errors
		{"valid_arithmetic", `ISERROR(2+3)`, ValueBool, false, 0},
		{"valid_division", `ISERROR(10/2)`, ValueBool, false, 0},

		// Wrong number of arguments returns #VALUE!
		{"no_args", `ISERROR()`, ValueError, false, ErrValVALUE},
		{"two_args", `ISERROR(1,2)`, ValueError, false, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantTyp == ValueBool {
				if got.Type != ValueBool || got.Bool != tt.wantBool {
					t.Errorf("%s = %v, want bool %v", tt.formula, got, tt.wantBool)
				}
			} else {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("%s = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			}
		})
	}
}

func TestISNA(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name     string
		formula  string
		wantTyp  ValueType
		wantBool bool
		wantErr  ErrorValue
	}{
		// #N/A error should return TRUE
		{"NA_function", `ISNA(NA())`, ValueBool, true, 0},
		{"NA_literal", `ISNA(#N/A)`, ValueBool, true, 0},

		// Other error types should return FALSE (only #N/A is TRUE)
		{"DIV0_error", `ISNA(1/0)`, ValueBool, false, 0},
		{"DIV0_zero_div_zero", `ISNA(0/0)`, ValueBool, false, 0},
		{"VALUE_error", `ISNA(#VALUE!)`, ValueBool, false, 0},
		{"REF_error", `ISNA(#REF!)`, ValueBool, false, 0},
		{"NUM_error", `ISNA(#NUM!)`, ValueBool, false, 0},
		{"NULL_error", `ISNA(#NULL!)`, ValueBool, false, 0},
		{"NAME_error", `ISNA(#NAME?)`, ValueBool, false, 0},

		// Numbers should return FALSE
		{"number_positive", `ISNA(1)`, ValueBool, false, 0},
		{"number_zero", `ISNA(0)`, ValueBool, false, 0},
		{"number_negative", `ISNA(-5)`, ValueBool, false, 0},
		{"number_decimal", `ISNA(3.14)`, ValueBool, false, 0},

		// Text should return FALSE
		{"text_string", `ISNA("text")`, ValueBool, false, 0},
		{"empty_string", `ISNA("")`, ValueBool, false, 0},

		// Boolean should return FALSE
		{"bool_true", `ISNA(TRUE)`, ValueBool, false, 0},
		{"bool_false", `ISNA(FALSE)`, ValueBool, false, 0},

		// Expressions that do not produce #N/A
		{"valid_arithmetic", `ISNA(2+3)`, ValueBool, false, 0},
		{"valid_division", `ISNA(10/2)`, ValueBool, false, 0},

		// Wrong number of arguments returns #VALUE!
		{"no_args", `ISNA()`, ValueError, false, ErrValVALUE},
		{"two_args", `ISNA(1,2)`, ValueError, false, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantTyp == ValueBool {
				if got.Type != ValueBool || got.Bool != tt.wantBool {
					t.Errorf("%s = %v, want bool %v", tt.formula, got, tt.wantBool)
				}
			} else {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("%s = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			}
		})
	}
}

func TestISNUMBER(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("positive_integer", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER(1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(1) = %v, want TRUE", got)
		}
	})

	t.Run("zero", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER(0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(0) = %v, want TRUE", got)
		}
	})

	t.Run("negative_decimal", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER(-5.5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(-5.5) = %v, want TRUE", got)
		}
	})

	t.Run("large_number", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER(999999999)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(999999999) = %v, want TRUE", got)
		}
	})

	t.Run("very_small_decimal", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER(0.0001)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(0.0001) = %v, want TRUE", got)
		}
	})

	t.Run("expression_result", func(t *testing.T) {
		// Result of arithmetic is a number
		cf := evalCompile(t, `ISNUMBER(2+3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(2+3) = %v, want TRUE", got)
		}
	})

	t.Run("negative_integer", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER(-100)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(-100) = %v, want TRUE", got)
		}
	})

	t.Run("text_string", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER("text")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISNUMBER(\"text\") = %v, want FALSE", got)
		}
	})

	t.Run("empty_string", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER("")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISNUMBER(\"\") = %v, want FALSE", got)
		}
	})

	t.Run("string_looks_like_number", func(t *testing.T) {
		// Per Excel docs: numeric values in double quotes are treated as text
		cf := evalCompile(t, `ISNUMBER("123")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISNUMBER(\"123\") = %v, want FALSE", got)
		}
	})

	t.Run("string_looks_like_negative_number", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER("-42.5")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISNUMBER(\"-42.5\") = %v, want FALSE", got)
		}
	})

	t.Run("boolean_TRUE_is_not_number", func(t *testing.T) {
		// Booleans are NOT numbers in Excel
		cf := evalCompile(t, `ISNUMBER(TRUE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISNUMBER(TRUE) = %v, want FALSE", got)
		}
	})

	t.Run("boolean_FALSE_is_not_number", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER(FALSE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISNUMBER(FALSE) = %v, want FALSE", got)
		}
	})

	t.Run("error_DIV0_is_not_number", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISNUMBER(1/0) = %v, want FALSE", got)
		}
	})

	t.Run("error_NA_is_not_number", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER(#N/A)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISNUMBER(#N/A) = %v, want FALSE", got)
		}
	})

	t.Run("error_VALUE_is_not_number", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER(#VALUE!)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISNUMBER(#VALUE!) = %v, want FALSE", got)
		}
	})

	t.Run("nested_SUM_returns_number", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER(SUM(1,2,3))`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(SUM(1,2,3)) = %v, want TRUE", got)
		}
	})

	t.Run("PI_is_number", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER(PI())`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(PI()) = %v, want TRUE", got)
		}
	})

	t.Run("cell_ref_number", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(330.92),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     r,
		}
		cf := evalCompile(t, `ISNUMBER(A1)`)
		got, err := Eval(cf, r, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(A1) with numeric cell = %v, want TRUE", got)
		}
	})

	t.Run("cell_ref_text", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("hello"),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     r,
		}
		cf := evalCompile(t, `ISNUMBER(A1)`)
		got, err := Eval(cf, r, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISNUMBER(A1) with text cell = %v, want FALSE", got)
		}
	})

	t.Run("cell_ref_empty", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     r,
		}
		cf := evalCompile(t, `ISNUMBER(A1)`)
		got, err := Eval(cf, r, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISNUMBER(A1) with empty cell = %v, want FALSE", got)
		}
	})

	t.Run("no_args_error", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISNUMBER() = %v, want #VALUE!", got)
		}
	})

	t.Run("too_many_args_error", func(t *testing.T) {
		cf := evalCompile(t, `ISNUMBER(1,2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISNUMBER(1,2) = %v, want #VALUE!", got)
		}
	})
}

func TestISLOGICAL(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    bool
	}{
		// Boolean literals → TRUE
		{"true_literal", `ISLOGICAL(TRUE)`, true},
		{"false_literal", `ISLOGICAL(FALSE)`, true},

		// Comparison expressions produce boolean results → TRUE
		{"greater_than", `ISLOGICAL(1>0)`, true},
		{"less_than", `ISLOGICAL(0<1)`, true},
		{"equal_comparison", `ISLOGICAL(1=1)`, true},
		{"not_equal", `ISLOGICAL(1<>2)`, true},
		{"greater_equal", `ISLOGICAL(5>=5)`, true},
		{"less_equal", `ISLOGICAL(3<=4)`, true},

		// Numbers are NOT logical → FALSE
		{"integer_one", `ISLOGICAL(1)`, false},
		{"integer_zero", `ISLOGICAL(0)`, false},
		{"negative_number", `ISLOGICAL(-1)`, false},
		{"decimal", `ISLOGICAL(3.14)`, false},
		{"large_number", `ISLOGICAL(1000000)`, false},

		// Strings are NOT logical → FALSE
		{"string_true", `ISLOGICAL("TRUE")`, false},
		{"string_false", `ISLOGICAL("FALSE")`, false},
		{"string_text", `ISLOGICAL("hello")`, false},
		{"empty_string", `ISLOGICAL("")`, false},
		{"numeric_string", `ISLOGICAL("1")`, false},

		// Arithmetic expressions produce numbers → FALSE
		{"addition", `ISLOGICAL(1+1)`, false},
		{"multiplication", `ISLOGICAL(2*3)`, false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueBool {
				t.Fatalf("%s = %v (type %d), want bool", tt.formula, got, got.Type)
			}
			if got.Bool != tt.want {
				t.Errorf("%s = %v, want %v", tt.formula, got.Bool, tt.want)
			}
		})
	}

	// Error values → FALSE (errors are not logical)
	t.Run("error_div0", func(t *testing.T) {
		cf := evalCompile(t, `ISLOGICAL(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("ISLOGICAL(1/0) = %v, want FALSE", got)
		}
	})

	t.Run("error_NA", func(t *testing.T) {
		cf := evalCompile(t, `ISLOGICAL(#N/A)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("ISLOGICAL(#N/A) = %v, want FALSE", got)
		}
	})

	t.Run("error_VALUE", func(t *testing.T) {
		cf := evalCompile(t, `ISLOGICAL(#VALUE!)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("ISLOGICAL(#VALUE!) = %v, want FALSE", got)
		}
	})

	// Cell reference containing a boolean → TRUE
	t.Run("cell_with_true", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     r,
		}
		cf := evalCompile(t, `ISLOGICAL(A1)`)
		got, err := Eval(cf, r, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISLOGICAL(A1) with TRUE = %v, want TRUE", got)
		}
	})

	// Cell reference containing a number → FALSE
	t.Run("cell_with_number", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     r,
		}
		cf := evalCompile(t, `ISLOGICAL(A1)`)
		got, err := Eval(cf, r, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISLOGICAL(A1) with number = %v, want FALSE", got)
		}
	})

	// Wrong argument count → #VALUE!
	t.Run("no_args_error", func(t *testing.T) {
		cf := evalCompile(t, `ISLOGICAL()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISLOGICAL() = %v, want #VALUE!", got)
		}
	})

	t.Run("too_many_args_error", func(t *testing.T) {
		cf := evalCompile(t, `ISLOGICAL(TRUE,FALSE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISLOGICAL(TRUE,FALSE) = %v, want #VALUE!", got)
		}
	})
}

func TestN(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantTyp ValueType
		wantNum float64
		wantErr ErrorValue
	}{
		// Numbers: N returns the number itself
		{"positive_integer", `N(42)`, ValueNumber, 42, 0},
		{"zero", `N(0)`, ValueNumber, 0, 0},
		{"negative_integer", `N(-7)`, ValueNumber, -7, 0},
		{"decimal", `N(3.14)`, ValueNumber, 3.14, 0},
		{"large_number", `N(1000000)`, ValueNumber, 1000000, 0},

		// Booleans: TRUE=1, FALSE=0
		{"true", `N(TRUE)`, ValueNumber, 1, 0},
		{"false", `N(FALSE)`, ValueNumber, 0, 0},

		// Strings: any string returns 0
		{"text_string", `N("hello")`, ValueNumber, 0, 0},
		{"empty_string", `N("")`, ValueNumber, 0, 0},
		{"numeric_string", `N("123")`, ValueNumber, 0, 0},

		// Errors: N returns the error value unchanged
		{"error_na", `N(#N/A)`, ValueError, 0, ErrValNA},
		{"error_value", `N(#VALUE!)`, ValueError, 0, ErrValVALUE},
		{"error_div0", `N(1/0)`, ValueError, 0, ErrValDIV0},
		{"error_ref", `N(#REF!)`, ValueError, 0, ErrValREF},

		// Expressions that produce numbers
		{"arithmetic_expr", `N(2+3)`, ValueNumber, 5, 0},

		// Wrong number of arguments
		{"no_args", `N()`, ValueError, 0, ErrValVALUE},
		{"two_args", `N(1,2)`, ValueError, 0, ErrValVALUE},
		{"three_args", `N(1,2,3)`, ValueError, 0, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantTyp == ValueError {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("%s = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			} else {
				if got.Type != ValueNumber || got.Num != tt.wantNum {
					t.Errorf("%s = %v, want number %v", tt.formula, got, tt.wantNum)
				}
			}
		})
	}
}

func TestNA(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns_na_error", func(t *testing.T) {
		cf := evalCompile(t, `NA()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("NA() = %v, want #N/A", got)
		}
	})

	t.Run("with_one_arg_returns_value_error", func(t *testing.T) {
		cf := evalCompile(t, `NA(1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("NA(1) = %v, want #VALUE!", got)
		}
	})

	t.Run("with_two_args_returns_value_error", func(t *testing.T) {
		cf := evalCompile(t, `NA(1,2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("NA(1,2) = %v, want #VALUE!", got)
		}
	})

	t.Run("with_string_arg_returns_value_error", func(t *testing.T) {
		cf := evalCompile(t, `NA("test")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`NA("test") = %v, want #VALUE!`, got)
		}
	})

	t.Run("nested_in_isna", func(t *testing.T) {
		cf := evalCompile(t, `ISNA(NA())`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNA(NA()) = %v, want TRUE", got)
		}
	})

	t.Run("nested_in_iserror", func(t *testing.T) {
		cf := evalCompile(t, `ISERROR(NA())`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISERROR(NA()) = %v, want TRUE", got)
		}
	})
}

func TestERROR_TYPE(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantTyp ValueType
		wantNum float64
		wantErr ErrorValue
	}{
		// Each error type maps to a specific number
		{"null_error", `ERROR.TYPE(#NULL!)`, ValueNumber, 1, 0},
		{"div0_error", `ERROR.TYPE(1/0)`, ValueNumber, 2, 0},
		{"value_error", `ERROR.TYPE(#VALUE!)`, ValueNumber, 3, 0},
		{"ref_error", `ERROR.TYPE(#REF!)`, ValueNumber, 4, 0},
		{"name_error", `ERROR.TYPE(#NAME?)`, ValueNumber, 5, 0},
		{"num_error", `ERROR.TYPE(#NUM!)`, ValueNumber, 6, 0},
		{"na_error", `ERROR.TYPE(#N/A)`, ValueNumber, 7, 0},
		{"na_function", `ERROR.TYPE(NA())`, ValueNumber, 7, 0},

		// Division by zero via expression
		{"div0_expr", `ERROR.TYPE(0/0)`, ValueNumber, 2, 0},

		// Non-error values return #N/A
		{"number", `ERROR.TYPE(42)`, ValueError, 0, ErrValNA},
		{"text", `ERROR.TYPE("hello")`, ValueError, 0, ErrValNA},
		{"bool_true", `ERROR.TYPE(TRUE)`, ValueError, 0, ErrValNA},
		{"bool_false", `ERROR.TYPE(FALSE)`, ValueError, 0, ErrValNA},
		{"zero", `ERROR.TYPE(0)`, ValueError, 0, ErrValNA},
		{"empty_string", `ERROR.TYPE("")`, ValueError, 0, ErrValNA},

		// Wrong argument count returns #VALUE!
		{"no_args", `ERROR.TYPE()`, ValueError, 0, ErrValVALUE},
		{"two_args", `ERROR.TYPE(1,2)`, ValueError, 0, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantTyp == ValueNumber {
				if got.Type != ValueNumber || got.Num != tt.wantNum {
					t.Errorf("%s = %v, want %v", tt.formula, got, tt.wantNum)
				}
			} else {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("%s = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			}
		})
	}
}

func TestTYPE(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantTyp ValueType
		wantNum float64
		wantErr ErrorValue
	}{
		// Number -> 1
		{"positive_integer", `TYPE(42)`, ValueNumber, 1, 0},
		{"zero", `TYPE(0)`, ValueNumber, 1, 0},
		{"negative", `TYPE(-5)`, ValueNumber, 1, 0},
		{"decimal", `TYPE(3.14)`, ValueNumber, 1, 0},
		{"arithmetic_expr", `TYPE(1+2)`, ValueNumber, 1, 0},

		// String -> 2
		{"text", `TYPE("hello")`, ValueNumber, 2, 0},
		{"empty_string", `TYPE("")`, ValueNumber, 2, 0},
		{"numeric_string", `TYPE("123")`, ValueNumber, 2, 0},

		// Boolean -> 4
		{"true", `TYPE(TRUE)`, ValueNumber, 4, 0},
		{"false", `TYPE(FALSE)`, ValueNumber, 4, 0},
		{"comparison", `TYPE(1>0)`, ValueNumber, 4, 0},

		// Error -> 16
		{"error_div0", `TYPE(1/0)`, ValueNumber, 16, 0},
		{"error_value", `TYPE(#VALUE!)`, ValueNumber, 16, 0},
		{"error_na", `TYPE(#N/A)`, ValueNumber, 16, 0},
		{"error_ref", `TYPE(#REF!)`, ValueNumber, 16, 0},
		{"error_name", `TYPE(#NAME?)`, ValueNumber, 16, 0},
		{"error_num", `TYPE(#NUM!)`, ValueNumber, 16, 0},
		{"error_null", `TYPE(#NULL!)`, ValueNumber, 16, 0},

		// Wrong argument count returns #VALUE!
		{"no_args", `TYPE()`, ValueError, 0, ErrValVALUE},
		{"two_args", `TYPE(1,2)`, ValueError, 0, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if tt.wantTyp == ValueNumber {
				if got.Type != ValueNumber || got.Num != tt.wantNum {
					t.Errorf("%s = %v, want %v", tt.formula, got, tt.wantNum)
				}
			} else {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("%s = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			}
		})
	}

	// TYPE(array) -> 64 -- test with a cell range that produces an array
	t.Run("array_type", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   3,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     r,
		}
		cf := evalCompile(t, `TYPE(A1:A2)`)
		got, err := Eval(cf, r, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 64 {
			t.Errorf("TYPE(A1:A2) = %v, want 64", got)
		}
	})

	// TYPE on an empty cell reference -> 1 (empty treated as number)
	t.Run("empty_cell", func(t *testing.T) {
		r := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     r,
		}
		cf := evalCompile(t, `TYPE(A1)`)
		got, err := Eval(cf, r, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("TYPE(A1) empty = %v, want 1", got)
		}
	})
}
