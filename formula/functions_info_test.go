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
		// ISBLANK("") returns FALSE. An empty string is NOT the same as blank.
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
	// --- Literal / expression tests (no cell references) ---

	t.Run("even_positive_4", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(4)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(4) = %v, want TRUE", got)
		}
	})

	t.Run("odd_positive_3", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISEVEN(3) = %v, want FALSE", got)
		}
	})

	t.Run("zero_is_even", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(0) = %v, want TRUE", got)
		}
	})

	t.Run("even_positive_2", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(2) = %v, want TRUE", got)
		}
	})

	t.Run("even_positive_100", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(100)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(100) = %v, want TRUE", got)
		}
	})

	t.Run("odd_positive_1", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISEVEN(1) = %v, want FALSE", got)
		}
	})

	t.Run("odd_positive_5", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISEVEN(5) = %v, want FALSE", got)
		}
	})

	t.Run("negative_even", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(-2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(-2) = %v, want TRUE", got)
		}
	})

	t.Run("negative_even_minus4", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(-4)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(-4) = %v, want TRUE", got)
		}
	})

	t.Run("negative_odd_minus1", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISEVEN(-1) = %v, want FALSE", got)
		}
	})

	t.Run("negative_odd_minus3", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(-3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISEVEN(-3) = %v, want FALSE", got)
		}
	})

	t.Run("decimal_truncated_to_even", func(t *testing.T) {
		// ISEVEN(2.5) truncates to 2, which is even → TRUE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(2.5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(2.5) = %v, want TRUE", got)
		}
	})

	t.Run("decimal_truncated_to_odd", func(t *testing.T) {
		// ISEVEN(3.9) truncates to 3, which is odd → FALSE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(3.9)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISEVEN(3.9) = %v, want FALSE", got)
		}
	})

	t.Run("negative_decimal_truncated_to_even", func(t *testing.T) {
		// ISEVEN(-2.7) truncates to -2, which is even → TRUE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(-2.7)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(-2.7) = %v, want TRUE", got)
		}
	})

	t.Run("negative_decimal_truncated_to_odd", func(t *testing.T) {
		// ISEVEN(-3.1) truncates to -3, which is odd → FALSE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(-3.1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISEVEN(-3.1) = %v, want FALSE", got)
		}
	})

	t.Run("large_even_number", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(1000000)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(1000000) = %v, want TRUE", got)
		}
	})

	t.Run("large_odd_number", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(999999)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISEVEN(999999) = %v, want FALSE", got)
		}
	})

	t.Run("string_coercion_even", func(t *testing.T) {
		// ISEVEN("4") coerces "4" to 4 → TRUE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN("4")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf(`ISEVEN("4") = %v, want TRUE`, got)
		}
	})

	t.Run("string_coercion_odd", func(t *testing.T) {
		// ISEVEN("3") coerces "3" to 3 → FALSE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN("3")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf(`ISEVEN("3") = %v, want FALSE`, got)
		}
	})

	t.Run("non_numeric_string_returns_value_error", func(t *testing.T) {
		// ISEVEN("hello") → #VALUE!
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN("hello")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`ISEVEN("hello") = %v, want #VALUE!`, got)
		}
	})

	t.Run("empty_string_returns_value_error", func(t *testing.T) {
		// ISEVEN("") → #VALUE! (empty string is not numeric)
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN("")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`ISEVEN("") = %v, want #VALUE!`, got)
		}
	})

	t.Run("boolean_true_returns_value_error", func(t *testing.T) {
		// Excel: ISEVEN(TRUE) → #VALUE!
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(TRUE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISEVEN(TRUE) = %v, want #VALUE!", got)
		}
	})

	t.Run("boolean_false_returns_value_error", func(t *testing.T) {
		// Excel: ISEVEN(FALSE) → #VALUE!
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(FALSE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISEVEN(FALSE) = %v, want #VALUE!", got)
		}
	})

	t.Run("error_propagation_div0", func(t *testing.T) {
		// ISEVEN(1/0) → #DIV/0!
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("ISEVEN(1/0) = %v, want #DIV/0!", got)
		}
	})

	t.Run("expression_result_even", func(t *testing.T) {
		// ISEVEN(2+2) → TRUE (4 is even)
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(2+2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(2+2) = %v, want TRUE", got)
		}
	})

	t.Run("expression_result_odd", func(t *testing.T) {
		// ISEVEN(2+1) → FALSE (3 is odd)
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(2+1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISEVEN(2+1) = %v, want FALSE", got)
		}
	})

	t.Run("string_coercion_decimal", func(t *testing.T) {
		// ISEVEN("2.9") coerces to 2.9, truncated to 2 → TRUE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN("2.9")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf(`ISEVEN("2.9") = %v, want TRUE`, got)
		}
	})

	// --- Cell reference tests ---

	t.Run("cell_reference_even_number", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(6),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISEVEN(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(A1) with A1=6 = %v, want TRUE", got)
		}
	})

	t.Run("cell_reference_odd_number", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(7),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISEVEN(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISEVEN(A1) with A1=7 = %v, want FALSE", got)
		}
	})

	t.Run("cell_reference_empty_cell_coerces_to_zero", func(t *testing.T) {
		// Empty cell → EmptyVal → CoerceNum returns 0 → even → TRUE
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISEVEN(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(A1) with empty A1 = %v, want TRUE", got)
		}
	})

	t.Run("cell_reference_with_error_propagates", func(t *testing.T) {
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
		cf := evalCompile(t, `ISEVEN(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("ISEVEN(A1) with A1=#N/A = %v, want #N/A", got)
		}
	})

	t.Run("very_small_decimal", func(t *testing.T) {
		// ISEVEN(0.1) truncates to 0, which is even → TRUE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(0.1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(0.1) = %v, want TRUE", got)
		}
	})

	t.Run("negative_zero", func(t *testing.T) {
		// ISEVEN(-0) is even → TRUE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISEVEN(-0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISEVEN(-0) = %v, want TRUE", got)
		}
	})
}

func TestISODD(t *testing.T) {
	// --- Basic odd numbers (TRUE) ---

	t.Run("odd_positive_1", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISODD(1) = %v, want TRUE", got)
		}
	})

	t.Run("odd_positive_3", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISODD(3) = %v, want TRUE", got)
		}
	})

	t.Run("odd_positive_5", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISODD(5) = %v, want TRUE", got)
		}
	})

	t.Run("odd_positive_99", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(99)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISODD(99) = %v, want TRUE", got)
		}
	})

	// --- Basic even numbers (FALSE) ---

	t.Run("even_positive_2", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISODD(2) = %v, want FALSE", got)
		}
	})

	t.Run("even_positive_4", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(4)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISODD(4) = %v, want FALSE", got)
		}
	})

	// --- Zero ---

	t.Run("zero_is_not_odd", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISODD(0) = %v, want FALSE", got)
		}
	})

	// --- Negative odd numbers ---

	t.Run("negative_odd_minus1", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(-1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISODD(-1) = %v, want TRUE", got)
		}
	})

	t.Run("negative_odd_minus3", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(-3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISODD(-3) = %v, want TRUE", got)
		}
	})

	// --- Negative even numbers ---

	t.Run("negative_even_minus2", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(-2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISODD(-2) = %v, want FALSE", got)
		}
	})

	t.Run("negative_even_minus4", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(-4)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISODD(-4) = %v, want FALSE", got)
		}
	})

	// --- Decimal truncation ---

	t.Run("decimal_truncated_to_odd", func(t *testing.T) {
		// ISODD(3.5) truncates to 3, which is odd -> TRUE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(3.5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISODD(3.5) = %v, want TRUE", got)
		}
	})

	t.Run("decimal_truncated_to_even", func(t *testing.T) {
		// ISODD(2.9) truncates to 2, which is even -> FALSE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(2.9)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISODD(2.9) = %v, want FALSE", got)
		}
	})

	t.Run("negative_decimal_truncated_to_odd", func(t *testing.T) {
		// ISODD(-3.1) truncates to -3, which is odd -> TRUE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(-3.1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISODD(-3.1) = %v, want TRUE", got)
		}
	})

	t.Run("negative_decimal_truncated_to_even", func(t *testing.T) {
		// ISODD(-2.7) truncates to -2, which is even -> FALSE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(-2.7)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISODD(-2.7) = %v, want FALSE", got)
		}
	})

	t.Run("very_small_decimal", func(t *testing.T) {
		// ISODD(0.1) truncates to 0, which is even -> FALSE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(0.1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISODD(0.1) = %v, want FALSE", got)
		}
	})

	// --- Large numbers ---

	t.Run("large_odd_number", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(999999)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISODD(999999) = %v, want TRUE", got)
		}
	})

	t.Run("large_even_number", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(1000000)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISODD(1000000) = %v, want FALSE", got)
		}
	})

	// --- String coercion ---

	t.Run("string_coercion_odd", func(t *testing.T) {
		// ISODD("3") coerces "3" to 3 -> TRUE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD("3")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf(`ISODD("3") = %v, want TRUE`, got)
		}
	})

	t.Run("string_coercion_even", func(t *testing.T) {
		// ISODD("4") coerces "4" to 4 -> FALSE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD("4")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf(`ISODD("4") = %v, want FALSE`, got)
		}
	})

	t.Run("string_coercion_decimal", func(t *testing.T) {
		// ISODD("3.9") coerces to 3.9, truncated to 3 -> TRUE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD("3.9")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf(`ISODD("3.9") = %v, want TRUE`, got)
		}
	})

	// --- Non-numeric string ---

	t.Run("non_numeric_string_returns_value_error", func(t *testing.T) {
		// ISODD("hello") -> #VALUE!
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD("hello")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`ISODD("hello") = %v, want #VALUE!`, got)
		}
	})

	t.Run("empty_string_returns_value_error", func(t *testing.T) {
		// ISODD("") -> #VALUE! (empty string is not numeric)
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD("")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`ISODD("") = %v, want #VALUE!`, got)
		}
	})

	// --- Booleans ---

	t.Run("boolean_true_returns_value_error", func(t *testing.T) {
		// Excel: ISODD(TRUE) -> #VALUE!
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(TRUE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISODD(TRUE) = %v, want #VALUE!", got)
		}
	})

	t.Run("boolean_false_returns_value_error", func(t *testing.T) {
		// Excel: ISODD(FALSE) -> #VALUE!
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(FALSE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ISODD(FALSE) = %v, want #VALUE!", got)
		}
	})

	// --- Error propagation ---

	t.Run("error_propagation_div0", func(t *testing.T) {
		// ISODD(1/0) -> #DIV/0!
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("ISODD(1/0) = %v, want #DIV/0!", got)
		}
	})

	// --- Expressions ---

	t.Run("expression_result_odd", func(t *testing.T) {
		// ISODD(2+1) -> TRUE (3 is odd)
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(2+1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISODD(2+1) = %v, want TRUE", got)
		}
	})

	t.Run("expression_result_even", func(t *testing.T) {
		// ISODD(2+2) -> FALSE (4 is even)
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(2+2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISODD(2+2) = %v, want FALSE", got)
		}
	})

	// --- Edge: negative zero ---

	t.Run("negative_zero", func(t *testing.T) {
		// ISODD(-0) is even -> FALSE
		resolver := &mockResolver{}
		cf := evalCompile(t, `ISODD(-0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISODD(-0) = %v, want FALSE", got)
		}
	})

	// --- Cell reference tests ---

	t.Run("cell_reference_odd_number", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(7),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISODD(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISODD(A1) with A1=7 = %v, want TRUE", got)
		}
	})

	t.Run("cell_reference_even_number", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(6),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISODD(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISODD(A1) with A1=6 = %v, want FALSE", got)
		}
	})

	t.Run("cell_reference_empty_cell_coerces_to_zero", func(t *testing.T) {
		// Empty cell -> EmptyVal -> CoerceNum returns 0 -> even -> FALSE
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISODD(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISODD(A1) with empty A1 = %v, want FALSE", got)
		}
	})

	t.Run("cell_reference_with_error_propagates", func(t *testing.T) {
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
		cf := evalCompile(t, `ISODD(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("ISODD(A1) with A1=#N/A = %v, want #N/A", got)
		}
	})
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
	// Set up a rich mock with many formula and non-formula cells.
	//
	// Layout:
	//   A1: 42         (constant number)
	//   B1: =A1+58     (simple arithmetic formula, value 100)
	//   C1: "hello"    (constant text)
	//   D1: TRUE       (constant boolean)
	//   A2: =SUM(A1:A1) (SUM formula, value 42)
	//   B2: =IF(A1>10,7,0) (IF formula, value 7)
	//   C2: =CONCATENATE("hello","world") (text-returning formula)
	//   D2: =SUM(IF(A1>0,A1,0)) (nested functions formula, value 42)
	//   A3: =$A$1      (absolute reference formula, value 42)
	//   B3: =A$1+$B1   (mixed reference formula, value 142)
	//   C3: #DIV/0!    (constant error, no formula)
	//   D3: =1/0       (formula that returns error, #DIV/0!)
	//   A4: =FORMULATEXT(B1) (meta formula)
	//   B4: =CONCATENATE(...) (complex text formula)
	//   C4: =LEN(FORMULATEXT(B1)) (deeply nested formula)
	//   D4: =COLUMN(A1) (function with reference arg)
	//   A5: =AVERAGE(A1,B1) (statistical formula)
	//   B5: =MAX(1,2,3) (MAX formula)
	//   C5: =MIN(A1,0) (MIN formula)
	//   D5: =UPPER("hello") (text-returning formula)
	//   E1: (empty cell, no value at all)
	//   Sheet2!A1: =Sheet1!A1*2+15 (cross-sheet formula)
	resolver := &mockFormulaResolver{
		mockResolver: mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),           // A1: constant number
				{Col: 2, Row: 1}: NumberVal(100),          // B1: has formula =A1+58
				{Col: 3, Row: 1}: StringVal("hello"),      // C1: constant text
				{Col: 4, Row: 1}: BoolVal(true),           // D1: constant boolean
				{Col: 1, Row: 2}: NumberVal(42),           // A2: has formula =SUM(A1:A1)
				{Col: 2, Row: 2}: NumberVal(7),            // B2: has formula =IF(A1>10,7,0)
				{Col: 3, Row: 2}: StringVal("helloworld"), // C2: has formula =CONCATENATE(...)
				{Col: 4, Row: 2}: NumberVal(42),           // D2: has formula =SUM(IF(A1>0,A1,0))
				{Col: 1, Row: 3}: NumberVal(42),           // A3: has formula =$A$1
				{Col: 2, Row: 3}: NumberVal(142),          // B3: has formula =A$1+$B1
				{Col: 3, Row: 3}: ErrorVal(ErrValDIV0),    // C3: constant error (no formula)
				{Col: 4, Row: 3}: ErrorVal(ErrValDIV0),    // D3: has formula =1/0
				{Col: 1, Row: 4}: StringVal("=A1+58"),     // A4: has formula =FORMULATEXT(B1)
				{Col: 2, Row: 4}: StringVal("complex"),    // B4: has formula
				{Col: 3, Row: 4}: NumberVal(6),            // C4: has formula =LEN(FORMULATEXT(B1))
				{Col: 4, Row: 4}: NumberVal(1),            // D4: has formula =COLUMN(A1)
				{Col: 1, Row: 5}: NumberVal(71),           // A5: has formula =AVERAGE(A1,B1)
				{Col: 2, Row: 5}: NumberVal(3),            // B5: has formula =MAX(1,2,3)
				{Col: 3, Row: 5}: NumberVal(0),            // C5: has formula =MIN(A1,0)
				{Col: 4, Row: 5}: StringVal("HELLO"),      // D5: has formula =UPPER("hello")
				// E1 (Col:5, Row:1) intentionally missing — empty cell
				// Cross-sheet cell
				{Sheet: "Sheet2", Col: 1, Row: 1}: NumberVal(99), // Sheet2!A1: has formula
			},
		},
		formulas: map[CellAddr]string{
			{Col: 2, Row: 1}:                    "A1+58",
			{Col: 1, Row: 2}:                    "SUM(A1:A1)",
			{Col: 2, Row: 2}:                    "IF(A1>10,7,0)",
			{Col: 3, Row: 2}:                    `CONCATENATE("hello","world")`,
			{Col: 4, Row: 2}:                    "SUM(IF(A1>0,A1,0))",
			{Col: 1, Row: 3}:                    "$A$1",
			{Col: 2, Row: 3}:                    "A$1+$B1",
			{Col: 4, Row: 3}:                    "1/0",
			{Col: 1, Row: 4}:                    "FORMULATEXT(B1)",
			{Col: 2, Row: 4}:                    `CONCATENATE("she said ",CHAR(34),"hi",CHAR(34))`,
			{Col: 3, Row: 4}:                    "LEN(FORMULATEXT(B1))",
			{Col: 4, Row: 4}:                    "COLUMN(A1)",
			{Col: 1, Row: 5}:                    "AVERAGE(A1,B1)",
			{Col: 2, Row: 5}:                    "MAX(1,2,3)",
			{Col: 3, Row: 5}:                    "MIN(A1,0)",
			{Col: 4, Row: 5}:                    `UPPER("hello")`,
			{Sheet: "Sheet2", Col: 1, Row: 1}:   "Sheet1!A1*2+15",
		},
	}
	ctx := &EvalContext{
		CurrentCol:   10,
		CurrentRow:   10,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	// ---- Table-driven: cells that should return TRUE (formula cells) ----
	trueCases := []struct {
		name    string
		formula string
	}{
		{"simple_arithmetic_formula", `ISFORMULA(B1)`},
		{"SUM_formula", `ISFORMULA(A2)`},
		{"IF_formula", `ISFORMULA(B2)`},
		{"CONCATENATE_formula_returns_text", `ISFORMULA(C2)`},
		{"nested_functions_SUM_IF", `ISFORMULA(D2)`},
		{"absolute_reference_formula", `ISFORMULA(A3)`},
		{"mixed_reference_formula", `ISFORMULA(B3)`},
		{"formula_returning_error_DIV0", `ISFORMULA(D3)`},
		{"FORMULATEXT_meta_formula", `ISFORMULA(A4)`},
		{"complex_text_formula", `ISFORMULA(B4)`},
		{"deeply_nested_LEN_FORMULATEXT", `ISFORMULA(C4)`},
		{"COLUMN_function_formula", `ISFORMULA(D4)`},
		{"AVERAGE_formula", `ISFORMULA(A5)`},
		{"MAX_formula", `ISFORMULA(B5)`},
		{"MIN_formula", `ISFORMULA(C5)`},
		{"UPPER_text_formula", `ISFORMULA(D5)`},
	}
	for _, tc := range trueCases {
		t.Run("true_"+tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, ctx)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if got.Type != ValueBool || !got.Bool {
				t.Errorf("%s = %v, want TRUE", tc.formula, got)
			}
		})
	}

	// ---- Table-driven: cells that should return FALSE (non-formula cells) ----
	falseCases := []struct {
		name    string
		formula string
	}{
		{"constant_number", `ISFORMULA(A1)`},
		{"constant_text", `ISFORMULA(C1)`},
		{"constant_boolean", `ISFORMULA(D1)`},
		{"empty_cell", `ISFORMULA(E1)`},
		{"constant_error_no_formula", `ISFORMULA(C3)`},
		{"completely_unused_cell", `ISFORMULA(Z99)`},
	}
	for _, tc := range falseCases {
		t.Run("false_"+tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, ctx)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if got.Type != ValueBool || got.Bool {
				t.Errorf("%s = %v, want FALSE", tc.formula, got)
			}
		})
	}

	// ---- Error cases: wrong arg count or non-reference arguments ----
	errorCases := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"no_args", `ISFORMULA()`, ErrValVALUE},
		{"number_literal_arg", `ISFORMULA(123)`, ErrValVALUE},
		{"string_literal_arg", `ISFORMULA("hello")`, ErrValVALUE},
		{"boolean_literal_arg", `ISFORMULA(TRUE)`, ErrValVALUE},
	}
	for _, tc := range errorCases {
		t.Run("error_"+tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, ctx)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if got.Type != ValueError || got.Err != tc.wantErr {
				t.Errorf("%s = %v, want %v", tc.formula, got, tc.wantErr)
			}
		})
	}

	// ---- Composition: ISFORMULA used in larger expressions ----

	t.Run("NOT_ISFORMULA_on_constant_cell", func(t *testing.T) {
		// NOT(ISFORMULA(A1)) should be TRUE since A1 is a constant.
		cf := evalCompile(t, `NOT(ISFORMULA(A1))`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("NOT(ISFORMULA(A1)) = %v, want TRUE", got)
		}
	})

	t.Run("NOT_ISFORMULA_on_formula_cell", func(t *testing.T) {
		// NOT(ISFORMULA(B1)) should be FALSE since B1 has a formula.
		cf := evalCompile(t, `NOT(ISFORMULA(B1))`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("NOT(ISFORMULA(B1)) = %v, want FALSE", got)
		}
	})

	t.Run("IF_ISFORMULA_pattern", func(t *testing.T) {
		// IF(ISFORMULA(B1), "formula", "value") should return "formula".
		cf := evalCompile(t, `IF(ISFORMULA(B1), "formula", "value")`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "formula" {
			t.Errorf(`IF(ISFORMULA(B1), "formula", "value") = %v, want "formula"`, got)
		}
	})

	t.Run("IF_ISFORMULA_constant_pattern", func(t *testing.T) {
		// IF(ISFORMULA(A1), "formula", "value") should return "value".
		cf := evalCompile(t, `IF(ISFORMULA(A1), "formula", "value")`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "value" {
			t.Errorf(`IF(ISFORMULA(A1), "formula", "value") = %v, want "value"`, got)
		}
	})

	// ---- Consistency: ISFORMULA and FORMULATEXT agree ----

	t.Run("consistency_formula_cell", func(t *testing.T) {
		// B1 has a formula: ISFORMULA returns TRUE, FORMULATEXT returns text.
		cfIs := evalCompile(t, `ISFORMULA(B1)`)
		gotIs, err := Eval(cfIs, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval ISFORMULA: %v", err)
		}
		if gotIs.Type != ValueBool || !gotIs.Bool {
			t.Errorf("ISFORMULA(B1) = %v, want TRUE", gotIs)
		}
		cfFt := evalCompile(t, `FORMULATEXT(B1)`)
		gotFt, err := Eval(cfFt, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval FORMULATEXT: %v", err)
		}
		if gotFt.Type != ValueString {
			t.Errorf("FORMULATEXT(B1) = %v, want a string value", gotFt)
		}
	})

	t.Run("consistency_constant_cell", func(t *testing.T) {
		// A1 is a constant: ISFORMULA returns FALSE, FORMULATEXT returns #N/A.
		cfIs := evalCompile(t, `ISFORMULA(A1)`)
		gotIs, err := Eval(cfIs, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval ISFORMULA: %v", err)
		}
		if gotIs.Type != ValueBool || gotIs.Bool {
			t.Errorf("ISFORMULA(A1) = %v, want FALSE", gotIs)
		}
		cfFt := evalCompile(t, `FORMULATEXT(A1)`)
		gotFt, err := Eval(cfFt, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval FORMULATEXT: %v", err)
		}
		if gotFt.Type != ValueError || gotFt.Err != ErrValNA {
			t.Errorf("FORMULATEXT(A1) = %v, want #N/A", gotFt)
		}
	})

	t.Run("consistency_empty_cell", func(t *testing.T) {
		// E1 is empty: ISFORMULA returns FALSE, FORMULATEXT returns #N/A.
		cfIs := evalCompile(t, `ISFORMULA(E1)`)
		gotIs, err := Eval(cfIs, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval ISFORMULA: %v", err)
		}
		if gotIs.Type != ValueBool || gotIs.Bool {
			t.Errorf("ISFORMULA(E1) = %v, want FALSE", gotIs)
		}
		cfFt := evalCompile(t, `FORMULATEXT(E1)`)
		gotFt, err := Eval(cfFt, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval FORMULATEXT: %v", err)
		}
		if gotFt.Type != ValueError || gotFt.Err != ErrValNA {
			t.Errorf("FORMULATEXT(E1) = %v, want #N/A", gotFt)
		}
	})

	t.Run("consistency_error_formula_cell", func(t *testing.T) {
		// D3 has formula =1/0: ISFORMULA returns TRUE, FORMULATEXT returns text.
		cfIs := evalCompile(t, `ISFORMULA(D3)`)
		gotIs, err := Eval(cfIs, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval ISFORMULA: %v", err)
		}
		if gotIs.Type != ValueBool || !gotIs.Bool {
			t.Errorf("ISFORMULA(D3) = %v, want TRUE", gotIs)
		}
		cfFt := evalCompile(t, `FORMULATEXT(D3)`)
		gotFt, err := Eval(cfFt, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval FORMULATEXT: %v", err)
		}
		if gotFt.Type != ValueString || gotFt.Str != "=1/0" {
			t.Errorf("FORMULATEXT(D3) = %v, want =1/0", gotFt)
		}
	})
}

func TestFORMULATEXT(t *testing.T) {
	// Set up a rich mock with many formula cells for comprehensive testing.
	resolver := &mockFormulaResolver{
		mockResolver: mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),           // A1: constant number
				{Col: 2, Row: 1}: NumberVal(100),          // B1: has formula =A1+58
				{Col: 3, Row: 1}: StringVal("hello"),      // C1: constant text
				{Col: 4, Row: 1}: BoolVal(true),           // D1: constant boolean
				{Col: 1, Row: 2}: NumberVal(600),          // A2: has formula =SUM(A1:A1)
				{Col: 2, Row: 2}: NumberVal(7),            // B2: has formula =IF(A1>10,7,0)
				{Col: 3, Row: 2}: StringVal("helloworld"), // C2: has formula =CONCATENATE("hello","world")
				{Col: 4, Row: 2}: NumberVal(99),           // D2: has formula =SUM(IF(A1>0,A1,0))
				{Col: 1, Row: 3}: NumberVal(42),           // A3: has formula =$A$1
				{Col: 2, Row: 3}: NumberVal(42),           // B3: has formula =A$1+$B1
				{Col: 3, Row: 3}: ErrorVal(ErrValDIV0),    // C3: constant error (no formula)
				{Col: 4, Row: 3}: ErrorVal(ErrValDIV0),    // D3: has formula =1/0
				{Col: 1, Row: 4}: StringVal("=A1+58"),     // A4: has formula =FORMULATEXT(B1)
				{Col: 2, Row: 4}: StringVal(`she said "hi"`), // B4: has formula =CONCATENATE("she said ",CHAR(34),"hi",CHAR(34))
				{Col: 3, Row: 4}: NumberVal(6),            // C4: has formula =LEN(FORMULATEXT(B1))
				{Col: 4, Row: 4}: NumberVal(1),            // D4: has formula =COLUMN(A1)
				{Col: 1, Row: 5}: NumberVal(10),           // A5: has formula =AVERAGE(A1,B1)
				{Col: 2, Row: 5}: NumberVal(3),            // B5: has formula =MAX(1,2,3)
				{Col: 3, Row: 5}: NumberVal(0),            // C5: has formula =MIN(A1,0)
				{Col: 4, Row: 5}: StringVal("HELLO"),      // D5: has formula =UPPER("hello")
				// E1 (Col:5, Row:1) intentionally missing — empty cell
				// Cross-sheet cell
				{Sheet: "Sheet2", Col: 1, Row: 1}: NumberVal(99), // Sheet2!A1: has formula =Sheet1!A1*2+15
			},
		},
		formulas: map[CellAddr]string{
			{Col: 2, Row: 1}: "A1+58",
			{Col: 1, Row: 2}: "SUM(A1:A1)",
			{Col: 2, Row: 2}: "IF(A1>10,7,0)",
			{Col: 3, Row: 2}: `CONCATENATE("hello","world")`,
			{Col: 4, Row: 2}: "SUM(IF(A1>0,A1,0))",
			{Col: 1, Row: 3}: "$A$1",
			{Col: 2, Row: 3}: "A$1+$B1",
			{Col: 4, Row: 3}: "1/0",
			{Col: 1, Row: 4}: "FORMULATEXT(B1)",
			{Col: 2, Row: 4}: `CONCATENATE("she said ",CHAR(34),"hi",CHAR(34))`,
			{Col: 3, Row: 4}: "LEN(FORMULATEXT(B1))",
			{Col: 4, Row: 4}: "COLUMN(A1)",
			{Col: 1, Row: 5}: "AVERAGE(A1,B1)",
			{Col: 2, Row: 5}: "MAX(1,2,3)",
			{Col: 3, Row: 5}: "MIN(A1,0)",
			{Col: 4, Row: 5}: `UPPER("hello")`,
			{Sheet: "Sheet2", Col: 1, Row: 1}: "Sheet1!A1*2+15",
		},
	}
	ctx := &EvalContext{
		CurrentCol:   10,
		CurrentRow:   10,
		CurrentSheet: "",
		Resolver:     resolver,
	}

	// --- Sub-tests returning formula text (string) ---

	t.Run("simple_addition_formula", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(B1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=A1+58" {
			t.Errorf("FORMULATEXT(B1) = %v, want =A1+58", got)
		}
	})

	t.Run("SUM_formula", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(A2)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=SUM(A1:A1)" {
			t.Errorf("FORMULATEXT(A2) = %v, want =SUM(A1:A1)", got)
		}
	})

	t.Run("IF_formula", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(B2)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=IF(A1>10,7,0)" {
			t.Errorf("FORMULATEXT(B2) = %v, want =IF(A1>10,7,0)", got)
		}
	})

	t.Run("CONCATENATE_string_formula", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(C2)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != `=CONCATENATE("hello","world")` {
			t.Errorf("FORMULATEXT(C2) = %v, want =CONCATENATE(\"hello\",\"world\")", got)
		}
	})

	t.Run("nested_functions_SUM_IF", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(D2)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=SUM(IF(A1>0,A1,0))" {
			t.Errorf("FORMULATEXT(D2) = %v, want =SUM(IF(A1>0,A1,0))", got)
		}
	})

	t.Run("absolute_reference_formula", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(A3)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=$A$1" {
			t.Errorf("FORMULATEXT(A3) = %v, want =$A$1", got)
		}
	})

	t.Run("mixed_absolute_relative_reference", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(B3)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=A$1+$B1" {
			t.Errorf("FORMULATEXT(B3) = %v, want =A$1+$B1", got)
		}
	})

	t.Run("formula_returning_error_still_shows_text", func(t *testing.T) {
		// D3 has formula =1/0 which evaluates to #DIV/0!, but
		// FORMULATEXT should return the formula text, not the error.
		cf := evalCompile(t, `FORMULATEXT(D3)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=1/0" {
			t.Errorf("FORMULATEXT(D3) = %v, want =1/0", got)
		}
	})

	t.Run("FORMULATEXT_of_FORMULATEXT", func(t *testing.T) {
		// A4 has formula =FORMULATEXT(B1)
		cf := evalCompile(t, `FORMULATEXT(A4)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=FORMULATEXT(B1)" {
			t.Errorf("FORMULATEXT(A4) = %v, want =FORMULATEXT(B1)", got)
		}
	})

	t.Run("formula_with_string_containing_quotes", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(B4)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		want := `=CONCATENATE("she said ",CHAR(34),"hi",CHAR(34))`
		if got.Type != ValueString || got.Str != want {
			t.Errorf("FORMULATEXT(B4) = %v, want %s", got, want)
		}
	})

	t.Run("LEN_of_FORMULATEXT", func(t *testing.T) {
		// C4 has formula =LEN(FORMULATEXT(B1))
		cf := evalCompile(t, `FORMULATEXT(C4)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=LEN(FORMULATEXT(B1))" {
			t.Errorf("FORMULATEXT(C4) = %v, want =LEN(FORMULATEXT(B1))", got)
		}
	})

	t.Run("COLUMN_formula", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(D4)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=COLUMN(A1)" {
			t.Errorf("FORMULATEXT(D4) = %v, want =COLUMN(A1)", got)
		}
	})

	t.Run("AVERAGE_formula", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(A5)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=AVERAGE(A1,B1)" {
			t.Errorf("FORMULATEXT(A5) = %v, want =AVERAGE(A1,B1)", got)
		}
	})

	t.Run("MAX_formula", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(B5)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=MAX(1,2,3)" {
			t.Errorf("FORMULATEXT(B5) = %v, want =MAX(1,2,3)", got)
		}
	})

	t.Run("MIN_formula", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(C5)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=MIN(A1,0)" {
			t.Errorf("FORMULATEXT(C5) = %v, want =MIN(A1,0)", got)
		}
	})

	t.Run("UPPER_text_function_formula", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(D5)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != `=UPPER("hello")` {
			t.Errorf(`FORMULATEXT(D5) = %v, want =UPPER("hello")`, got)
		}
	})

	t.Run("cross_sheet_reference", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(Sheet2!A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "=Sheet1!A1*2+15" {
			t.Errorf("FORMULATEXT(Sheet2!A1) = %v, want =Sheet1!A1*2+15", got)
		}
	})

	t.Run("formula_text_always_has_leading_equals", func(t *testing.T) {
		// Verify every formula cell returns text starting with '='.
		cf := evalCompile(t, `FORMULATEXT(B1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || len(got.Str) == 0 || got.Str[0] != '=' {
			t.Errorf("FORMULATEXT(B1) = %v, expected string starting with '='", got)
		}
	})

	// --- Sub-tests returning #N/A (cell without formula) ---

	t.Run("constant_number_returns_NA", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("FORMULATEXT(A1) = %v, want #N/A", got)
		}
	})

	t.Run("constant_text_returns_NA", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(C1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("FORMULATEXT(C1) = %v, want #N/A", got)
		}
	})

	t.Run("constant_boolean_returns_NA", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(D1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("FORMULATEXT(D1) = %v, want #N/A", got)
		}
	})

	t.Run("empty_cell_returns_NA", func(t *testing.T) {
		// E1 is not in the cells map at all — empty cell.
		cf := evalCompile(t, `FORMULATEXT(E1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("FORMULATEXT(E1) = %v, want #N/A", got)
		}
	})

	t.Run("constant_error_returns_NA", func(t *testing.T) {
		// C3 has a #DIV/0! error but no formula — should return #N/A.
		cf := evalCompile(t, `FORMULATEXT(C3)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("FORMULATEXT(C3) = %v, want #N/A", got)
		}
	})

	// --- Error handling sub-tests ---

	t.Run("no_args_returns_VALUE_error", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT()`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("FORMULATEXT() = %v, want #VALUE!", got)
		}
	})

	t.Run("numeric_literal_arg_returns_VALUE_error", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(123)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("FORMULATEXT(123) = %v, want #VALUE!", got)
		}
	})

	t.Run("string_literal_arg_returns_VALUE_error", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT("hello")`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`FORMULATEXT("hello") = %v, want #VALUE!`, got)
		}
	})

	t.Run("boolean_literal_arg_returns_VALUE_error", func(t *testing.T) {
		cf := evalCompile(t, `FORMULATEXT(TRUE)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("FORMULATEXT(TRUE) = %v, want #VALUE!", got)
		}
	})

	// --- Consistency: ISFORMULA and FORMULATEXT agree ---

	t.Run("ISFORMULA_FORMULATEXT_consistency_formula_cell", func(t *testing.T) {
		// B1 has a formula — ISFORMULA returns TRUE, FORMULATEXT returns text.
		cfIs := evalCompile(t, `ISFORMULA(B1)`)
		gotIs, err := Eval(cfIs, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval ISFORMULA: %v", err)
		}
		if gotIs.Type != ValueBool || !gotIs.Bool {
			t.Errorf("ISFORMULA(B1) = %v, want TRUE", gotIs)
		}
		cfFt := evalCompile(t, `FORMULATEXT(B1)`)
		gotFt, err := Eval(cfFt, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval FORMULATEXT: %v", err)
		}
		if gotFt.Type != ValueString {
			t.Errorf("FORMULATEXT(B1) = %v, want a string value", gotFt)
		}
	})

	t.Run("ISFORMULA_FORMULATEXT_consistency_constant_cell", func(t *testing.T) {
		// A1 is a constant — ISFORMULA returns FALSE, FORMULATEXT returns #N/A.
		cfIs := evalCompile(t, `ISFORMULA(A1)`)
		gotIs, err := Eval(cfIs, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval ISFORMULA: %v", err)
		}
		if gotIs.Type != ValueBool || gotIs.Bool {
			t.Errorf("ISFORMULA(A1) = %v, want FALSE", gotIs)
		}
		cfFt := evalCompile(t, `FORMULATEXT(A1)`)
		gotFt, err := Eval(cfFt, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval FORMULATEXT: %v", err)
		}
		if gotFt.Type != ValueError || gotFt.Err != ErrValNA {
			t.Errorf("FORMULATEXT(A1) = %v, want #N/A", gotFt)
		}
	})
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

	// Blank cell → TRUE ("this function returns TRUE if the value refers to a blank cell")
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

			// Documentation example: ROWS(C1:E4) = 4
			{"doc_example", `ROWS(C1:E4)`, 4},

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
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// A1:A3 — lookup range for MATCH tests
			{Col: 1, Row: 1}: StringVal("apple"),
			{Col: 1, Row: 2}: StringVal("banana"),
			{Col: 1, Row: 3}: StringVal("cherry"),
			// B1:B3 — result range for VLOOKUP tests
			{Col: 2, Row: 1}: NumberVal(10),
			{Col: 2, Row: 2}: NumberVal(20),
			{Col: 2, Row: 3}: NumberVal(30),
		},
	}

	type want struct {
		typ  ValueType
		num  float64
		str  string
		bool bool
		err  ErrorValue
	}

	tests := []struct {
		name    string
		formula string
		want    want
	}{
		// --- Existing basic tests (preserved) ---
		{"na_returns_fallback", `IFNA(#N/A,"default")`, want{typ: ValueString, str: "default"}},
		{"number_returns_value", `IFNA(42,"default")`, want{typ: ValueNumber, num: 42}},

		// --- #N/A caught ---
		{"na_literal_string_fallback", `IFNA(#N/A,"not available")`, want{typ: ValueString, str: "not available"}},
		{"na_function_caught", `IFNA(NA(),"caught")`, want{typ: ValueString, str: "caught"}},
		{"na_fallback_number", `IFNA(#N/A,99)`, want{typ: ValueNumber, num: 99}},
		{"na_fallback_bool_true", `IFNA(#N/A,TRUE)`, want{typ: ValueBool, bool: true}},
		{"na_fallback_bool_false", `IFNA(#N/A,FALSE)`, want{typ: ValueBool, bool: false}},
		{"na_fallback_zero", `IFNA(#N/A,0)`, want{typ: ValueNumber, num: 0}},
		{"na_fallback_empty_string", `IFNA(#N/A,"")`, want{typ: ValueString, str: ""}},
		{"na_fallback_is_na", `IFNA(#N/A,#N/A)`, want{typ: ValueError, err: ErrValNA}},
		{"na_fallback_is_error", `IFNA(#N/A,#VALUE!)`, want{typ: ValueError, err: ErrValVALUE}},
		{"both_args_na", `IFNA(#N/A,NA())`, want{typ: ValueError, err: ErrValNA}},

		// --- Other errors pass through (NOT caught by IFNA) ---
		{"value_error_passthrough", `IFNA(#VALUE!,"fallback")`, want{typ: ValueError, err: ErrValVALUE}},
		{"div0_error_passthrough", `IFNA(1/0,"fallback")`, want{typ: ValueError, err: ErrValDIV0}},
		{"ref_error_passthrough", `IFNA(#REF!,"fallback")`, want{typ: ValueError, err: ErrValREF}},
		{"num_error_passthrough", `IFNA(#NUM!,"fallback")`, want{typ: ValueError, err: ErrValNUM}},
		{"name_error_passthrough", `IFNA(#NAME?,"fallback")`, want{typ: ValueError, err: ErrValNAME}},
		{"null_error_passthrough", `IFNA(#NULL!,"fallback")`, want{typ: ValueError, err: ErrValNULL}},

		// --- Value is not #N/A — returns value ---
		{"positive_number", `IFNA(100,"x")`, want{typ: ValueNumber, num: 100}},
		{"negative_number", `IFNA(-7,"x")`, want{typ: ValueNumber, num: -7}},
		{"decimal_number", `IFNA(3.14,"x")`, want{typ: ValueNumber, num: 3.14}},
		{"zero_value", `IFNA(0,"x")`, want{typ: ValueNumber, num: 0}},
		{"string_value", `IFNA("hello","x")`, want{typ: ValueString, str: "hello"}},
		{"empty_string_value", `IFNA("","x")`, want{typ: ValueString, str: ""}},
		{"bool_true_value", `IFNA(TRUE,"x")`, want{typ: ValueBool, bool: true}},
		{"bool_false_value", `IFNA(FALSE,"x")`, want{typ: ValueBool, bool: false}},

		// --- Nested IFNA ---
		{"nested_inner_catches", `IFNA(IFNA(#N/A,#N/A),"caught")`, want{typ: ValueString, str: "caught"}},
		{"nested_inner_ok", `IFNA(IFNA(5,#N/A),"caught")`, want{typ: ValueNumber, num: 5}},
		{"nested_outer_catches", `IFNA(IFNA(#N/A,"ok"),99)`, want{typ: ValueString, str: "ok"}},

		// --- With MATCH that returns #N/A (not found) ---
		{"match_not_found", `IFNA(MATCH("grape",A1:A3,0),"not found")`, want{typ: ValueString, str: "not found"}},
		{"match_found", `IFNA(MATCH("banana",A1:A3,0),"not found")`, want{typ: ValueNumber, num: 2}},

		// --- Arithmetic expressions as value ---
		{"arithmetic_ok", `IFNA(2+3,"x")`, want{typ: ValueNumber, num: 5}},
		{"arithmetic_div0", `IFNA(0/0,"x")`, want{typ: ValueError, err: ErrValDIV0}},

		// --- Wrong argument count ---
		{"no_args", `IFNA()`, want{typ: ValueError, err: ErrValVALUE}},
		{"one_arg", `IFNA(1)`, want{typ: ValueError, err: ErrValVALUE}},
		{"three_args", `IFNA(1,2,3)`, want{typ: ValueError, err: ErrValVALUE}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tt.formula, err)
			}
			switch tt.want.typ {
			case ValueNumber:
				if got.Type != ValueNumber || got.Num != tt.want.num {
					t.Errorf("%s = %v (num %g), want number %g", tt.formula, got.Type, got.Num, tt.want.num)
				}
			case ValueString:
				if got.Type != ValueString || got.Str != tt.want.str {
					t.Errorf("%s = %v (str %q), want string %q", tt.formula, got.Type, got.Str, tt.want.str)
				}
			case ValueBool:
				if got.Type != ValueBool || got.Bool != tt.want.bool {
					t.Errorf("%s = %v (bool %v), want bool %v", tt.formula, got.Type, got.Bool, tt.want.bool)
				}
			case ValueError:
				if got.Type != ValueError || got.Err != tt.want.err {
					t.Errorf("%s = %v (err %v), want error %v", tt.formula, got.Type, got.Err, tt.want.err)
				}
			case ValueEmpty:
				if got.Type != ValueEmpty {
					t.Errorf("%s = %v, want empty", tt.formula, got.Type)
				}
			default:
				t.Fatalf("unexpected want type %v", tt.want.typ)
			}
		})
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
		// Per docs: numeric values in double quotes are treated as text
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
		// Booleans are NOT numbers
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

	type want struct {
		typ  ValueType
		num  float64
		str  string
		bool bool
		err  ErrorValue
	}

	tests := []struct {
		name    string
		formula string
		want    want
	}{
		// --- Basic: NA() returns #N/A error ---
		{"returns_na_error", `NA()`, want{typ: ValueError, err: ErrValNA}},
		{"result_type_is_error", `NA()`, want{typ: ValueError, err: ErrValNA}},

		// --- NA() with wrong argument counts returns #VALUE! ---
		{"with_one_arg_returns_value_error", `NA(1)`, want{typ: ValueError, err: ErrValVALUE}},
		{"with_two_args_returns_value_error", `NA(1,2)`, want{typ: ValueError, err: ErrValVALUE}},
		{"with_string_arg_returns_value_error", `NA("test")`, want{typ: ValueError, err: ErrValVALUE}},
		{"with_bool_arg_returns_value_error", `NA(TRUE)`, want{typ: ValueError, err: ErrValVALUE}},

		// --- NA() is specifically #N/A, not other error types ---
		{"error_type_returns_7", `ERROR.TYPE(NA())`, want{typ: ValueNumber, num: 7}},

		// --- NA() inside ISNA → TRUE ---
		{"isna_of_na_is_true", `ISNA(NA())`, want{typ: ValueBool, bool: true}},

		// --- ISNA with non-NA values → FALSE ---
		{"isna_of_number_is_false", `ISNA(1)`, want{typ: ValueBool, bool: false}},
		{"isna_of_text_is_false", `ISNA("text")`, want{typ: ValueBool, bool: false}},
		{"isna_of_bool_is_false", `ISNA(TRUE)`, want{typ: ValueBool, bool: false}},
		{"isna_of_value_error_is_false", `ISNA(#VALUE!)`, want{typ: ValueBool, bool: false}},
		{"isna_of_div0_error_is_false", `ISNA(1/0)`, want{typ: ValueBool, bool: false}},

		// --- NA() inside ISERROR → TRUE ---
		{"iserror_of_na_is_true", `ISERROR(NA())`, want{typ: ValueBool, bool: true}},

		// --- NA() inside ISERR → FALSE (ISERR catches all errors EXCEPT #N/A) ---
		{"iserr_of_na_is_false", `ISERR(NA())`, want{typ: ValueBool, bool: false}},

		// --- NA() inside IFERROR → returns alternative value ---
		{"iferror_catches_na", `IFERROR(NA(),"fallback")`, want{typ: ValueString, str: "fallback"}},
		{"iferror_catches_na_number", `IFERROR(NA(),99)`, want{typ: ValueNumber, num: 99}},

		// --- NA() inside IFNA → returns alternative value ---
		{"ifna_catches_na", `IFNA(NA(),"caught")`, want{typ: ValueString, str: "caught"}},
		{"ifna_catches_na_number", `IFNA(NA(),42)`, want{typ: ValueNumber, num: 42}},
		{"ifna_catches_na_bool", `IFNA(NA(),TRUE)`, want{typ: ValueBool, bool: true}},

		// --- NA() in arithmetic → #N/A propagation ---
		{"na_plus_one", `NA()+1`, want{typ: ValueError, err: ErrValNA}},
		{"one_plus_na", `1+NA()`, want{typ: ValueError, err: ErrValNA}},
		{"na_minus_one", `NA()-1`, want{typ: ValueError, err: ErrValNA}},
		{"na_times_two", `NA()*2`, want{typ: ValueError, err: ErrValNA}},
		{"na_div_two", `NA()/2`, want{typ: ValueError, err: ErrValNA}},

		// --- NA() in comparison → #N/A propagation ---
		{"na_gt_one", `NA()>1`, want{typ: ValueError, err: ErrValNA}},
		{"na_lt_one", `NA()<1`, want{typ: ValueError, err: ErrValNA}},
		{"na_eq_one", `NA()=1`, want{typ: ValueError, err: ErrValNA}},

		// --- NA() in string concatenation → #N/A propagation ---
		{"na_concat_string", `NA()&"text"`, want{typ: ValueError, err: ErrValNA}},
		{"string_concat_na", `"text"&NA()`, want{typ: ValueError, err: ErrValNA}},

		// --- NA() as argument to aggregate functions → #N/A propagation ---
		{"sum_with_na", `SUM(NA(),1,2)`, want{typ: ValueError, err: ErrValNA}},
		{"average_with_na", `AVERAGE(NA(),1,2)`, want{typ: ValueError, err: ErrValNA}},

		// --- NA() as IF condition → #N/A propagation ---
		{"if_na_condition", `IF(NA(),"yes","no")`, want{typ: ValueError, err: ErrValNA}},

		// --- NA() as argument to AND/OR → #N/A propagation ---
		{"and_with_na", `AND(NA(),TRUE)`, want{typ: ValueError, err: ErrValNA}},
		{"or_with_na", `OR(NA(),FALSE)`, want{typ: ValueError, err: ErrValNA}},

		// --- Multiple NA() calls return same error type ---
		{"two_na_same_type", `ISNA(NA())`, want{typ: ValueBool, bool: true}},
		{"na_eq_na_propagates", `NA()=NA()`, want{typ: ValueError, err: ErrValNA}},

		// --- NA() in nested function chains ---
		{"nested_if_iferror_na", `IFERROR(IF(TRUE,NA(),"ok"),"recovered")`, want{typ: ValueString, str: "recovered"}},
		{"nested_ifna_sum", `IFNA(SUM(NA(),1),"sum failed")`, want{typ: ValueString, str: "sum failed"}},
		{"double_iferror_na", `IFERROR(IFERROR(NA(),"inner"),"outer")`, want{typ: ValueString, str: "inner"}},
		{"isna_iferror_na", `ISNA(IFERROR(NA(),"ok"))`, want{typ: ValueBool, bool: false}},
		{"error_type_of_iferror_na", `IFNA(NA(),0)+1`, want{typ: ValueNumber, num: 1}},

		// --- NA() compared to #N/A literal ---
		{"na_literal_also_na", `ISNA(#N/A)`, want{typ: ValueBool, bool: true}},
		{"error_type_na_literal", `ERROR.TYPE(#N/A)`, want{typ: ValueNumber, num: 7}},

		// --- NA() result used in TYPE function → 16 (error) ---
		{"type_of_na_is_16", `TYPE(NA())`, want{typ: ValueNumber, num: 16}},

		// --- NA() inside N() propagates error ---
		{"n_of_na_propagates", `N(NA())`, want{typ: ValueError, err: ErrValNA}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tt.formula, err)
			}
			switch tt.want.typ {
			case ValueNumber:
				if got.Type != ValueNumber || got.Num != tt.want.num {
					t.Errorf("%s = %v (num %g), want number %g", tt.formula, got.Type, got.Num, tt.want.num)
				}
			case ValueString:
				if got.Type != ValueString || got.Str != tt.want.str {
					t.Errorf("%s = %v (str %q), want string %q", tt.formula, got.Type, got.Str, tt.want.str)
				}
			case ValueBool:
				if got.Type != ValueBool || got.Bool != tt.want.bool {
					t.Errorf("%s = %v (bool %v), want bool %v", tt.formula, got.Type, got.Bool, tt.want.bool)
				}
			case ValueError:
				if got.Type != ValueError || got.Err != tt.want.err {
					t.Errorf("%s = %v (err %v), want error %v", tt.formula, got.Type, got.Err, tt.want.err)
				}
			}
		})
	}
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

func TestISREF(t *testing.T) {
	t.Run("cell_reference_A1", func(t *testing.T) {
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
		cf := evalCompile(t, `ISREF(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISREF(A1) = %v, want TRUE", got)
		}
	})

	t.Run("cell_reference_empty_cell", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   2,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISREF(A1) empty cell = %v, want TRUE (still a reference)", got)
		}
	})

	t.Run("range_reference", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(A1:B5)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISREF(A1:B5) = %v, want TRUE", got)
		}
	})

	t.Run("range_reference_single_column", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(A1:A5)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISREF(A1:A5) = %v, want TRUE", got)
		}
	})

	t.Run("number_literal", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISREF(1) = %v, want FALSE", got)
		}
	})

	t.Run("string_literal", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF("text")`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf(`ISREF("text") = %v, want FALSE`, got)
		}
	})

	t.Run("boolean_true", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(TRUE)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISREF(TRUE) = %v, want FALSE", got)
		}
	})

	t.Run("boolean_false", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(FALSE)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISREF(FALSE) = %v, want FALSE", got)
		}
	})

	t.Run("error_div_by_zero", func(t *testing.T) {
		// ISREF does NOT propagate errors — ISREF(1/0) returns FALSE
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(1/0)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISREF(1/0) = %v, want FALSE (not error propagation)", got)
		}
	})

	t.Run("function_result", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(SUM(1,2))`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISREF(SUM(1,2)) = %v, want FALSE", got)
		}
	})

	t.Run("no_args_error", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF()`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("ISREF() = %v, want #VALUE!", got)
		}
	})

	t.Run("too_many_args_error", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(1,2)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("ISREF(1,2) = %v, want #VALUE!", got)
		}
	})

	t.Run("cross_sheet_reference", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Sheet: "Sheet1", Col: 1, Row: 1}: NumberVal(99),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(Sheet1!A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISREF(Sheet1!A1) = %v, want TRUE", got)
		}
	})

	t.Run("number_zero", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(0)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISREF(0) = %v, want FALSE", got)
		}
	})

	t.Run("negative_number", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(-5)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISREF(-5) = %v, want FALSE", got)
		}
	})

	t.Run("empty_string", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF("")`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf(`ISREF("") = %v, want FALSE`, got)
		}
	})

	t.Run("expression_result", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(2+3)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISREF(2+3) = %v, want FALSE", got)
		}
	})

	t.Run("na_error", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(#N/A)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISREF(#N/A) = %v, want FALSE", got)
		}
	})

	t.Run("value_error", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(#VALUE!)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISREF(#VALUE!) = %v, want FALSE", got)
		}
	})

	t.Run("cell_ref_with_number_value", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 2, Row: 3}: NumberVal(100),
			},
		}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(B3)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISREF(B3) = %v, want TRUE", got)
		}
	})

	t.Run("cell_ref_with_string_value", func(t *testing.T) {
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
		cf := evalCompile(t, `ISREF(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISREF(A1) with string cell = %v, want TRUE", got)
		}
	})

	t.Run("cell_ref_with_bool_value", func(t *testing.T) {
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
		cf := evalCompile(t, `ISREF(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISREF(A1) with bool cell = %v, want TRUE", got)
		}
	})

	t.Run("concatenation_result", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF("a"&"b")`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf(`ISREF("a"&"b") = %v, want FALSE`, got)
		}
	})

	t.Run("pi_function_result", func(t *testing.T) {
		resolver := &mockResolver{}
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ISREF(PI())`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISREF(PI()) = %v, want FALSE", got)
		}
	})
}

func TestROW(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("no_args_returns_current_row", func(t *testing.T) {
		tests := []struct {
			name string
			row  int
			want float64
		}{
			{"row_1", 1, 1},
			{"row_3", 3, 3},
			{"row_5", 5, 5},
			{"row_10", 10, 10},
			{"row_100", 100, 100},
			{"row_256", 256, 256},
			{"row_1000", 1000, 1000},
			{"row_1048576", 1048576, 1048576},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				ctx := &EvalContext{
					CurrentCol:   1,
					CurrentRow:   tt.row,
					CurrentSheet: "",
					Resolver:     resolver,
				}
				cf := evalCompile(t, `ROW()`)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("ROW() with CurrentRow=%d = %v, want %v", tt.row, got, tt.want)
				}
			})
		}
	})

	t.Run("no_args_nil_context", func(t *testing.T) {
		// ROW() with no EvalContext should return #VALUE!
		cf := evalCompile(t, `ROW()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ROW() with nil ctx = %v, want #VALUE!", got)
		}
	})

	t.Run("single_cell_ref", func(t *testing.T) {
		tests := []struct {
			name    string
			formula string
			want    float64
		}{
			{"A1_row_1", `ROW(A1)`, 1},
			{"A2_row_2", `ROW(A2)`, 2},
			{"A5_row_5", `ROW(A5)`, 5},
			{"B10_row_10", `ROW(B10)`, 10},
			{"C50_row_50", `ROW(C50)`, 50},
			{"A100_row_100", `ROW(A100)`, 100},
			{"Z999_row_999", `ROW(Z999)`, 999},
			{"A1000_row_1000", `ROW(A1000)`, 1000},
			{"AA10000_row_10000", `ROW(AA10000)`, 10000},
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

	t.Run("ref_different_cols_same_row", func(t *testing.T) {
		// ROW should return the row regardless of which column the ref is in
		tests := []struct {
			name    string
			formula string
			want    float64
		}{
			{"A7", `ROW(A7)`, 7},
			{"B7", `ROW(B7)`, 7},
			{"C7", `ROW(C7)`, 7},
			{"Z7", `ROW(Z7)`, 7},
			{"AA7", `ROW(AA7)`, 7},
			{"XFD7", `ROW(XFD7)`, 7},
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
			{"number", `ROW(42)`},
			{"string", `ROW("hello")`},
			{"boolean_true", `ROW(TRUE)`},
			{"boolean_false", `ROW(FALSE)`},
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

	t.Run("range_ref_returns_array", func(t *testing.T) {
		// ROW(A1:A5) should return an array {1;2;3;4;5}
		ctx := &EvalContext{
			CurrentCol:     1,
			CurrentRow:     1,
			CurrentSheet:   "",
			Resolver:       resolver,
			IsArrayFormula: true,
		}
		cf := evalCompile(t, `ROW(A1:A5)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("ROW(A1:A5): expected array, got %v", got.Type)
		}
		if len(got.Array) != 5 {
			t.Fatalf("ROW(A1:A5): expected 5 rows, got %d", len(got.Array))
		}
		for i := 0; i < 5; i++ {
			want := float64(i + 1)
			if got.Array[i][0].Num != want {
				t.Errorf("ROW(A1:A5)[%d]: got %g, want %g", i, got.Array[i][0].Num, want)
			}
		}
	})

	t.Run("range_ref_multi_col", func(t *testing.T) {
		// ROW(B3:D7) should return array {3;4;5;6;7} — the rows of the range
		ctx := &EvalContext{
			CurrentCol:     1,
			CurrentRow:     1,
			CurrentSheet:   "",
			Resolver:       resolver,
			IsArrayFormula: true,
		}
		cf := evalCompile(t, `ROW(B3:D7)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("ROW(B3:D7): expected array, got %v", got.Type)
		}
		if len(got.Array) != 5 {
			t.Fatalf("ROW(B3:D7): expected 5 rows, got %d", len(got.Array))
		}
		for i := 0; i < 5; i++ {
			want := float64(i + 3)
			if got.Array[i][0].Num != want {
				t.Errorf("ROW(B3:D7)[%d]: got %g, want %g", i, got.Array[i][0].Num, want)
			}
		}
	})

	t.Run("range_ref_single_row", func(t *testing.T) {
		// ROW(A5:C5) — single-row range should return array with one element {5}
		ctx := &EvalContext{
			CurrentCol:     1,
			CurrentRow:     1,
			CurrentSheet:   "",
			Resolver:       resolver,
			IsArrayFormula: true,
		}
		cf := evalCompile(t, `ROW(A5:C5)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("ROW(A5:C5): expected array, got %v", got.Type)
		}
		if len(got.Array) != 1 {
			t.Fatalf("ROW(A5:C5): expected 1 row, got %d", len(got.Array))
		}
		if got.Array[0][0].Num != 5 {
			t.Errorf("ROW(A5:C5)[0]: got %g, want 5", got.Array[0][0].Num)
		}
	})

	t.Run("range_ref_starting_not_at_1", func(t *testing.T) {
		// ROW(A10:A12) should return {10;11;12}
		ctx := &EvalContext{
			CurrentCol:     1,
			CurrentRow:     1,
			CurrentSheet:   "",
			Resolver:       resolver,
			IsArrayFormula: true,
		}
		cf := evalCompile(t, `ROW(A10:A12)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("ROW(A10:A12): expected array, got %v", got.Type)
		}
		if len(got.Array) != 3 {
			t.Fatalf("ROW(A10:A12): expected 3 rows, got %d", len(got.Array))
		}
		expected := []float64{10, 11, 12}
		for i, want := range expected {
			if got.Array[i][0].Num != want {
				t.Errorf("ROW(A10:A12)[%d]: got %g, want %g", i, got.Array[i][0].Num, want)
			}
		}
	})

	t.Run("absolute_ref", func(t *testing.T) {
		// Absolute references ($A$5) should work identically
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ROW($A$5)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("ROW($A$5) = %v, want 5", got)
		}
	})

	t.Run("absolute_row_only", func(t *testing.T) {
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ROW(A$10)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("ROW(A$10) = %v, want 10", got)
		}
	})

	t.Run("absolute_col_only", func(t *testing.T) {
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ROW($B3)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("ROW($B3) = %v, want 3", got)
		}
	})

	t.Run("absolute_range", func(t *testing.T) {
		ctx := &EvalContext{
			CurrentCol:     1,
			CurrentRow:     1,
			CurrentSheet:   "",
			Resolver:       resolver,
			IsArrayFormula: true,
		}
		cf := evalCompile(t, `ROW($A$2:$A$4)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("ROW($A$2:$A$4): expected array, got %v", got.Type)
		}
		if len(got.Array) != 3 {
			t.Fatalf("ROW($A$2:$A$4): expected 3 rows, got %d", len(got.Array))
		}
		for i := 0; i < 3; i++ {
			want := float64(i + 2)
			if got.Array[i][0].Num != want {
				t.Errorf("ROW($A$2:$A$4)[%d]: got %g, want %g", i, got.Array[i][0].Num, want)
			}
		}
	})

	t.Run("error_propagation", func(t *testing.T) {
		// ROW(1/0) — the argument evaluates to #DIV/0! which is not a ref, so #VALUE!
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ROW(1/0)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("ROW(1/0) = %v, want error", got)
		}
	})

	t.Run("row_in_arithmetic", func(t *testing.T) {
		// ROW(A3)+10 should return 13
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ROW(A3)+10`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 13 {
			t.Errorf("ROW(A3)+10 = %v, want 13", got)
		}
	})

	t.Run("row_multiply", func(t *testing.T) {
		// ROW(A5)*2 should return 10
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ROW(A5)*2`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("ROW(A5)*2 = %v, want 10", got)
		}
	})

	t.Run("row_no_args_in_arithmetic", func(t *testing.T) {
		// ROW()+5 when current row is 3 should return 8
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   3,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ROW()+5`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 8 {
			t.Errorf("ROW()+5 with CurrentRow=3 = %v, want 8", got)
		}
	})

	t.Run("row_first_row", func(t *testing.T) {
		// ROW(A1) — the very first row should be 1
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ROW(A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("ROW(A1) = %v, want 1", got)
		}
	})

	t.Run("row_with_indirect_range", func(t *testing.T) {
		// ROW(INDIRECT("1:5")) should return array {1;2;3;4;5}
		ctx := &EvalContext{
			CurrentCol:     1,
			CurrentRow:     1,
			CurrentSheet:   "",
			Resolver:       resolver,
			IsArrayFormula: true,
		}
		cf := evalCompile(t, `ROW(INDIRECT("1:5"))`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("ROW(INDIRECT(\"1:5\")): expected array, got %v", got.Type)
		}
		if len(got.Array) != 5 {
			t.Fatalf("ROW(INDIRECT(\"1:5\")): expected 5 rows, got %d", len(got.Array))
		}
		for i := 0; i < 5; i++ {
			want := float64(i + 1)
			if got.Array[i][0].Num != want {
				t.Errorf("ROW(INDIRECT(\"1:5\"))[%d]: got %g, want %g", i, got.Array[i][0].Num, want)
			}
		}
	})

	t.Run("row_with_indirect_cell", func(t *testing.T) {
		// INDIRECT("B7") for a single cell resolves the cell value (not a ref),
		// so ROW(INDIRECT("B7")) receives a non-ref and returns #VALUE!.
		ctx := &EvalContext{
			CurrentCol:   1,
			CurrentRow:   1,
			CurrentSheet: "",
			Resolver:     resolver,
		}
		cf := evalCompile(t, `ROW(INDIRECT("B7"))`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("ROW(INDIRECT(\"B7\")) = %v, want #VALUE!", got)
		}
	})

	t.Run("row_range_single_cell_range", func(t *testing.T) {
		// ROW(A1:A1) — single cell range should return array with one element {1}
		ctx := &EvalContext{
			CurrentCol:     1,
			CurrentRow:     1,
			CurrentSheet:   "",
			Resolver:       resolver,
			IsArrayFormula: true,
		}
		cf := evalCompile(t, `ROW(A1:A1)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("ROW(A1:A1): expected array, got %v", got.Type)
		}
		if len(got.Array) != 1 {
			t.Fatalf("ROW(A1:A1): expected 1 row, got %d", len(got.Array))
		}
		if got.Array[0][0].Num != 1 {
			t.Errorf("ROW(A1:A1)[0]: got %g, want 1", got.Array[0][0].Num)
		}
	})

	t.Run("row_no_args_different_positions", func(t *testing.T) {
		// Verify ROW() returns the row of the cell regardless of the column
		tests := []struct {
			name string
			col  int
			row  int
		}{
			{"A1", 1, 1},
			{"B5", 2, 5},
			{"Z100", 26, 100},
			{"AA50", 27, 50},
		}
		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				ctx := &EvalContext{
					CurrentCol:   tt.col,
					CurrentRow:   tt.row,
					CurrentSheet: "",
					Resolver:     resolver,
				}
				cf := evalCompile(t, `ROW()`)
				got, err := Eval(cf, resolver, ctx)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueNumber || got.Num != float64(tt.row) {
					t.Errorf("ROW() at col=%d, row=%d = %v, want %v", tt.col, tt.row, got, float64(tt.row))
				}
			})
		}
	})

	t.Run("range_ref_large", func(t *testing.T) {
		// ROW(A1:A10) should return array {1;2;3;4;5;6;7;8;9;10}
		ctx := &EvalContext{
			CurrentCol:     1,
			CurrentRow:     1,
			CurrentSheet:   "",
			Resolver:       resolver,
			IsArrayFormula: true,
		}
		cf := evalCompile(t, `ROW(A1:A10)`)
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("ROW(A1:A10): expected array, got %v", got.Type)
		}
		if len(got.Array) != 10 {
			t.Fatalf("ROW(A1:A10): expected 10 rows, got %d", len(got.Array))
		}
		for i := 0; i < 10; i++ {
			want := float64(i + 1)
			if got.Array[i][0].Num != want {
				t.Errorf("ROW(A1:A10)[%d]: got %g, want %g", i, got.Array[i][0].Num, want)
			}
		}
	})
}
