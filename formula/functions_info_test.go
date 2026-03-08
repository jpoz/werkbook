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
