package formula

import (
	"math"
	"testing"
)

func TestMathFunctions(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		formula string
		want    float64
		tol     float64
	}{
		{"ABS(-5)", 5, 0},
		{"ABS(3)", 3, 0},
		{"POWER(2,3)", 8, 0},
		{"POWER(4,0.5)", 2, 1e-10},
		{"SQRT(16)", 4, 0},
		{"SQRT(2)", math.Sqrt2, 1e-10},
		{"INT(5.9)", 5, 0},
		{"INT(-5.1)", -6, 0},
		{"MOD(10,3)", 1, 0},
		{"MOD(-10,3)", 2, 0},
		{"ROUND(3.14159,2)", 3.14, 1e-10},
		{"ROUNDDOWN(3.7,0)", 3, 0},
		{"ROUNDDOWN(-3.7,0)", -3, 0},
		{"ROUNDUP(3.2,0)", 4, 0},
		{"ROUNDUP(-3.2,0)", -4, 0},
		{"CEILING(4.1,1)", 5, 0},
		{"CEILING(4.5,2)", 6, 0},
		{"FLOOR(4.9,1)", 4, 0},
		{"FLOOR(4.5,2)", 4, 0},
	}

	for _, tt := range numTests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueNumber {
			t.Errorf("Eval(%q): got type %v, want number", tt.formula, got.Type)
			continue
		}
		if tt.tol == 0 {
			if got.Num != tt.want {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		} else {
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g (tol %g)", tt.formula, got.Num, tt.want, tt.tol)
			}
		}
	}
}

func TestMathErrors(t *testing.T) {
	resolver := &mockResolver{}

	errTests := []struct {
		formula string
		errVal  ErrorValue
	}{
		{"SQRT(-1)", ErrValNUM},
		{"MOD(10,0)", ErrValDIV0},
		{"MOD(100,-1e-15)", ErrValNUM},
		// MOD precision overflow: |n/d| exceeds 2^53
		{"MOD(1e18,3)", ErrValNUM},
		{"FLOOR(2.5,0)", ErrValDIV0},
	}

	for _, tt := range errTests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueError || got.Err != tt.errVal {
			t.Errorf("Eval(%q) = type=%v err=%v, want %v", tt.formula, got.Type, got.Err, tt.errVal)
		}
	}
}

func TestCEILINGMATH(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Basic positive number rounding (default significance=1)
		{"pos_int_ceil", "CEILING.MATH(6.3)", 7},
		{"pos_int_exact", "CEILING.MATH(7)", 7},
		{"pos_small_frac", "CEILING.MATH(0.1)", 1},
		{"pos_half", "CEILING.MATH(2.5)", 3},

		// Positive with significance
		{"pos_sig_2", "CEILING.MATH(6.3,2)", 8},
		{"pos_sig_5", "CEILING.MATH(24.3,5)", 25},
		{"pos_sig_3", "CEILING.MATH(7,3)", 9},
		{"pos_sig_exact", "CEILING.MATH(6,3)", 6},
		{"pos_sig_0.1", "CEILING.MATH(6.31,0.1)", 6.4},
		{"pos_sig_0.5", "CEILING.MATH(6.3,0.5)", 6.5},

		// Negative number, default mode=0 (toward +infinity / toward zero)
		{"neg_default", "CEILING.MATH(-6.3)", -6},
		{"neg_sig_2_default", "CEILING.MATH(-6.3,2)", -6},
		{"neg_sig_2_mode0", "CEILING.MATH(-6.3,2,0)", -6},
		{"neg_exact", "CEILING.MATH(-6)", -6},
		{"neg_sig_5_default", "CEILING.MATH(-8.1,2)", -8},

		// Negative number, mode≠0 (away from zero)
		{"neg_mode1", "CEILING.MATH(-6.3,2,1)", -8},
		{"neg_mode_neg1", "CEILING.MATH(-5.5,2,-1)", -6},
		{"neg_mode1_sig1", "CEILING.MATH(-6.3,1,1)", -7},
		{"neg_mode1_sig5", "CEILING.MATH(-3,5,1)", -5},
		{"neg_mode1_exact", "CEILING.MATH(-6,3,1)", -6},

		// Significance of 0 returns 0
		{"sig_zero_pos", "CEILING.MATH(6.3,0)", 0},
		{"sig_zero_neg", "CEILING.MATH(-6.3,0)", 0},

		// Zero as number
		{"zero_number", "CEILING.MATH(0)", 0},
		{"zero_with_sig", "CEILING.MATH(0,5)", 0},

		// Negative significance (sign of significance is ignored in CEILING.MATH)
		{"neg_sig_pos_num", "CEILING.MATH(6.3,-2)", 8},
		{"neg_sig_neg_num", "CEILING.MATH(-6.3,-2)", -6},

		// Large numbers
		{"large_pos", "CEILING.MATH(1234567,1000)", 1235000},
		{"large_neg", "CEILING.MATH(-1234567,1000)", -1234000},
		{"large_neg_mode1", "CEILING.MATH(-1234567,1000,1)", -1235000},

		// Boolean coercion
		{"bool_true", "CEILING.MATH(TRUE)", 1},
		{"bool_false", "CEILING.MATH(FALSE)", 0},

		// String coercion of numeric strings
		{"string_num", "CEILING.MATH(\"6.3\")", 7},
	}

	for _, tt := range numTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("Eval(%q) = type %v, want ValueNumber", tt.formula, got.Type)
			}
			if math.Abs(got.Num-tt.wantNum) > 1e-10 {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"no_args", "CEILING.MATH()", ErrValVALUE},
		{"too_many_args", "CEILING.MATH(1,2,3,4)", ErrValVALUE},
		{"non_numeric", "CEILING.MATH(\"abc\")", ErrValVALUE},
		{"non_numeric_sig", "CEILING.MATH(1,\"abc\")", ErrValVALUE},
		{"non_numeric_mode", "CEILING.MATH(1,1,\"abc\")", ErrValVALUE},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = type=%v err=%v, want error %v", tt.formula, got.Type, got.Err, tt.wantErr)
			}
		})
	}
}

func TestFLOORMATH(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Basic positive number rounding (default significance=1)
		{"pos_int_floor", "FLOOR.MATH(6.7)", 6},
		{"pos_int_exact", "FLOOR.MATH(7)", 7},
		{"pos_small_frac", "FLOOR.MATH(0.9)", 0},
		{"pos_half", "FLOOR.MATH(2.5)", 2},

		// Positive with significance
		{"pos_sig_2", "FLOOR.MATH(6.7,2)", 6},
		{"pos_sig_5", "FLOOR.MATH(24.3,5)", 20},
		{"pos_sig_3", "FLOOR.MATH(7,3)", 6},
		{"pos_sig_exact", "FLOOR.MATH(6,3)", 6},
		{"pos_sig_0.1", "FLOOR.MATH(6.39,0.1)", 6.3},
		{"pos_sig_0.5", "FLOOR.MATH(6.7,0.5)", 6.5},

		// Negative number, default mode=0 (toward -infinity / away from zero)
		{"neg_default", "FLOOR.MATH(-6.7)", -7},
		{"neg_sig_2_default", "FLOOR.MATH(-6.3,2)", -8},
		{"neg_sig_2_mode0", "FLOOR.MATH(-8.1,2,0)", -10},
		{"neg_exact", "FLOOR.MATH(-6)", -6},
		{"neg_sig_5_default", "FLOOR.MATH(-3,5)", -5},

		// Negative number, mode≠0 (toward zero)
		{"neg_mode1", "FLOOR.MATH(-6.3,2,1)", -6},
		{"neg_mode_neg1", "FLOOR.MATH(-5.5,2,-1)", -4},
		{"neg_mode1_sig1", "FLOOR.MATH(-6.7,1,1)", -6},
		{"neg_mode1_sig5", "FLOOR.MATH(-3,5,1)", 0},
		{"neg_mode1_exact", "FLOOR.MATH(-6,3,1)", -6},

		// Significance of 0 returns 0
		{"sig_zero_pos", "FLOOR.MATH(6.3,0)", 0},
		{"sig_zero_neg", "FLOOR.MATH(-6.3,0)", 0},

		// Zero as number
		{"zero_number", "FLOOR.MATH(0)", 0},
		{"zero_with_sig", "FLOOR.MATH(0,5)", 0},

		// Negative significance (sign of significance is ignored in FLOOR.MATH)
		{"neg_sig_pos_num", "FLOOR.MATH(6.7,-2)", 6},
		{"neg_sig_neg_num", "FLOOR.MATH(-6.3,-2)", -8},

		// Large numbers
		{"large_pos", "FLOOR.MATH(1234567,1000)", 1234000},
		{"large_neg", "FLOOR.MATH(-1234567,1000)", -1235000},
		{"large_neg_mode1", "FLOOR.MATH(-1234567,1000,1)", -1234000},

		// Boolean coercion
		{"bool_true", "FLOOR.MATH(TRUE)", 1},
		{"bool_false", "FLOOR.MATH(FALSE)", 0},

		// String coercion of numeric strings
		{"string_num", "FLOOR.MATH(\"6.7\")", 6},
	}

	for _, tt := range numTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("Eval(%q) = type %v, want ValueNumber", tt.formula, got.Type)
			}
			if math.Abs(got.Num-tt.wantNum) > 1e-10 {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"no_args", "FLOOR.MATH()", ErrValVALUE},
		{"too_many_args", "FLOOR.MATH(1,2,3,4)", ErrValVALUE},
		{"non_numeric", "FLOOR.MATH(\"abc\")", ErrValVALUE},
		{"non_numeric_sig", "FLOOR.MATH(1,\"abc\")", ErrValVALUE},
		{"non_numeric_mode", "FLOOR.MATH(1,1,\"abc\")", ErrValVALUE},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = type=%v err=%v, want error %v", tt.formula, got.Type, got.Err, tt.wantErr)
			}
		})
	}
}

func TestCEILINGPRECISE(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Basic positive number rounding (default significance=1)
		{"pos_default", "CEILING.PRECISE(4.3)", 5},
		{"pos_exact", "CEILING.PRECISE(7)", 7},
		{"pos_small_frac", "CEILING.PRECISE(0.1)", 1},
		{"pos_half", "CEILING.PRECISE(2.5)", 3},

		// Positive with positive significance
		{"pos_sig_2", "CEILING.PRECISE(4.3,2)", 6},
		{"pos_sig_5", "CEILING.PRECISE(24.3,5)", 25},
		{"pos_sig_3", "CEILING.PRECISE(7,3)", 9},
		{"pos_sig_exact", "CEILING.PRECISE(6,3)", 6},
		{"pos_sig_0.1", "CEILING.PRECISE(6.31,0.1)", 6.4},
		{"pos_sig_0.5", "CEILING.PRECISE(6.3,0.5)", 6.5},

		// Negative numbers — always rounds toward +infinity (toward zero for negatives)
		{"neg_default", "CEILING.PRECISE(-4.3)", -4},
		{"neg_sig_2", "CEILING.PRECISE(-4.3,2)", -4},
		{"neg_sig_1", "CEILING.PRECISE(-4.1,1)", -4},
		{"neg_exact", "CEILING.PRECISE(-6)", -6},
		{"neg_sig_5", "CEILING.PRECISE(-8.1,5)", -5},
		{"neg_small_frac", "CEILING.PRECISE(-0.1)", 0},

		// Sign of significance is always ignored
		{"neg_sig_neg_sig", "CEILING.PRECISE(-4.3,-2)", -4},
		{"pos_neg_sig", "CEILING.PRECISE(4.3,-2)", 6},

		// Significance of 0 returns 0
		{"sig_zero_pos", "CEILING.PRECISE(6.3,0)", 0},
		{"sig_zero_neg", "CEILING.PRECISE(-6.3,0)", 0},

		// Zero as number
		{"zero_number", "CEILING.PRECISE(0)", 0},
		{"zero_with_sig", "CEILING.PRECISE(0,5)", 0},

		// Large numbers
		{"large_pos", "CEILING.PRECISE(1234567,1000)", 1235000},
		{"large_neg", "CEILING.PRECISE(-1234567,1000)", -1234000},

		// Boolean coercion
		{"bool_true", "CEILING.PRECISE(TRUE)", 1},
		{"bool_false", "CEILING.PRECISE(FALSE)", 0},

		// String coercion of numeric strings
		{"string_num", "CEILING.PRECISE(\"4.3\")", 5},

		// Excel doc examples
		{"doc_example_1", "CEILING.PRECISE(4.3)", 5},
		{"doc_example_2", "CEILING.PRECISE(-4.3)", -4},
		{"doc_example_3", "CEILING.PRECISE(4.3,2)", 6},
		{"doc_example_4", "CEILING.PRECISE(4.3,-2)", 6},
		{"doc_example_5", "CEILING.PRECISE(-4.3,2)", -4},
		{"doc_example_6", "CEILING.PRECISE(-4.3,-2)", -4},
	}

	for _, tt := range numTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("Eval(%q) = type %v, want ValueNumber", tt.formula, got.Type)
			}
			if math.Abs(got.Num-tt.wantNum) > 1e-10 {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"no_args", "CEILING.PRECISE()", ErrValVALUE},
		{"too_many_args", "CEILING.PRECISE(1,2,3)", ErrValVALUE},
		{"non_numeric", "CEILING.PRECISE(\"abc\")", ErrValVALUE},
		{"non_numeric_sig", "CEILING.PRECISE(1,\"abc\")", ErrValVALUE},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = type=%v err=%v, want error %v", tt.formula, got.Type, got.Err, tt.wantErr)
			}
		})
	}
}

func TestFLOORPRECISE(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Basic positive number rounding (default significance=1)
		{"pos_default", "FLOOR.PRECISE(3.2)", 3},
		{"pos_exact", "FLOOR.PRECISE(7)", 7},
		{"pos_small_frac", "FLOOR.PRECISE(0.9)", 0},
		{"pos_half", "FLOOR.PRECISE(2.5)", 2},

		// Positive with positive significance
		{"pos_sig_2", "FLOOR.PRECISE(4.3,2)", 4},
		{"pos_sig_5", "FLOOR.PRECISE(24.3,5)", 20},
		{"pos_sig_3", "FLOOR.PRECISE(7,3)", 6},
		{"pos_sig_exact", "FLOOR.PRECISE(6,3)", 6},
		{"pos_sig_0.1", "FLOOR.PRECISE(6.39,0.1)", 6.3},
		{"pos_sig_0.5", "FLOOR.PRECISE(6.7,0.5)", 6.5},

		// Negative numbers — always rounds toward -infinity (away from zero for negatives)
		{"neg_default", "FLOOR.PRECISE(-4.1)", -5},
		{"neg_sig_2", "FLOOR.PRECISE(-4.1,2)", -6},
		{"neg_sig_1", "FLOOR.PRECISE(-4.1,1)", -5},
		{"neg_exact", "FLOOR.PRECISE(-6)", -6},
		{"neg_sig_5", "FLOOR.PRECISE(-3,5)", -5},
		{"neg_small_frac", "FLOOR.PRECISE(-0.1)", -1},

		// Sign of significance is always ignored
		{"neg_sig_neg_sig", "FLOOR.PRECISE(-3.2,-1)", -4},
		{"pos_neg_sig", "FLOOR.PRECISE(3.2,-1)", 3},

		// Significance of 0 returns 0
		{"sig_zero_pos", "FLOOR.PRECISE(6.3,0)", 0},
		{"sig_zero_neg", "FLOOR.PRECISE(-6.3,0)", 0},

		// Zero as number
		{"zero_number", "FLOOR.PRECISE(0)", 0},
		{"zero_with_sig", "FLOOR.PRECISE(0,5)", 0},

		// Large numbers
		{"large_pos", "FLOOR.PRECISE(1234567,1000)", 1234000},
		{"large_neg", "FLOOR.PRECISE(-1234567,1000)", -1235000},

		// Boolean coercion
		{"bool_true", "FLOOR.PRECISE(TRUE)", 1},
		{"bool_false", "FLOOR.PRECISE(FALSE)", 0},

		// String coercion of numeric strings
		{"string_num", "FLOOR.PRECISE(\"3.2\")", 3},

		// Excel doc examples
		{"doc_example_1", "FLOOR.PRECISE(-3.2,-1)", -4},
		{"doc_example_2", "FLOOR.PRECISE(3.2,1)", 3},
		{"doc_example_3", "FLOOR.PRECISE(-3.2,1)", -4},
		{"doc_example_4", "FLOOR.PRECISE(3.2,-1)", 3},
		{"doc_example_5", "FLOOR.PRECISE(3.2)", 3},
	}

	for _, tt := range numTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("Eval(%q) = type %v, want ValueNumber", tt.formula, got.Type)
			}
			if math.Abs(got.Num-tt.wantNum) > 1e-10 {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"no_args", "FLOOR.PRECISE()", ErrValVALUE},
		{"too_many_args", "FLOOR.PRECISE(1,2,3)", ErrValVALUE},
		{"non_numeric", "FLOOR.PRECISE(\"abc\")", ErrValVALUE},
		{"non_numeric_sig", "FLOOR.PRECISE(1,\"abc\")", ErrValVALUE},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = type=%v err=%v, want error %v", tt.formula, got.Type, got.Err, tt.wantErr)
			}
		})
	}
}

func TestMOD(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		{"positive_mod", "MOD(10,3)", 1},
		{"exact_divisor", "MOD(10,5)", 0},
		{"negative_dividend", "MOD(-10,3)", 2},
		{"negative_divisor", "MOD(10,-3)", -2},
		{"both_negative", "MOD(-10,-3)", -1},
		{"fractional", "MOD(7.5,2)", 1.5},
		{"small_divisor_ok", "MOD(10,0.1)", 0},
	}

	const epsilon = 1e-10
	for _, tt := range numTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueNumber {
				t.Errorf("Eval(%q) = type %v, want ValueNumber", tt.formula, got.Type)
			} else if math.Abs(got.Num-tt.wantNum) > epsilon {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// Division by zero -> #DIV/0!
		{"div_by_zero", "MOD(10,0)", ErrValDIV0},
		// Precision overflow: |n/d| >= 1e13 -> #NUM!
		{"precision_overflow", "MOD(100,-1e-15)", ErrValNUM},
		{"precision_overflow_large", "MOD(1e18,1)", ErrValNUM},
		{"precision_overflow_excel", "MOD(10^15,7)", ErrValNUM},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = type=%v err=%v, want error %v", tt.formula, got.Type, got.Err, tt.wantErr)
			}
		})
	}
}

func TestROUND(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
		tol     float64
	}{
		// Basic positive rounding
		{"round_2dp", "ROUND(3.14159,2)", 3.14, 1e-10},
		{"round_1dp", "ROUND(2.15,1)", 2.2, 1e-10},
		{"round_1dp_down", "ROUND(2.149,1)", 2.1, 1e-10},
		{"round_3dp", "ROUND(1.23456,3)", 1.235, 1e-10},
		{"round_0dp", "ROUND(3.7,0)", 4, 0},
		{"round_0dp_down", "ROUND(3.2,0)", 3, 0},

		// Negative numbers
		{"neg_2dp", "ROUND(-1.475,2)", -1.48, 1e-10},
		{"neg_0dp", "ROUND(-3.7,0)", -4, 0},
		{"neg_0dp_down", "ROUND(-3.2,0)", -3, 0},
		{"neg_1dp", "ROUND(-2.15,1)", -2.2, 1e-10},

		// Negative num_digits (round to tens, hundreds, thousands)
		{"neg_digits_tens", "ROUND(21.5,-1)", 20, 0},
		{"neg_digits_thousands", "ROUND(626.3,-3)", 1000, 0},
		{"neg_digits_tens_small", "ROUND(1.98,-1)", 0, 0},
		{"neg_digits_hundreds_neg", "ROUND(-50.55,-2)", -100, 0},
		{"neg_digits_tens_pos", "ROUND(55,-1)", 60, 0},
		{"neg_digits_hundreds", "ROUND(149,-2)", 100, 0},
		{"neg_digits_hundreds_up", "ROUND(150,-2)", 200, 0},

		// Rounding 0.5 (away from zero in Excel)
		{"half_up", "ROUND(2.5,0)", 3, 0},
		{"half_up_neg", "ROUND(-2.5,0)", -3, 0},
		{"half_up_1dp", "ROUND(0.15,1)", 0.2, 1e-10},
		{"half_up_0.5", "ROUND(0.5,0)", 1, 0},

		// Zero
		{"zero", "ROUND(0,2)", 0, 0},
		{"zero_neg_digits", "ROUND(0,-1)", 0, 0},

		// Large numbers
		{"large_number", "ROUND(123456789.123,2)", 123456789.12, 1e-2},
		{"large_neg_digits", "ROUND(12345,-3)", 12000, 0},

		// Very small numbers
		{"small_positive", "ROUND(0.001,2)", 0, 0},
		{"small_positive_3dp", "ROUND(0.001,3)", 0.001, 1e-10},

		// Boolean coercion (TRUE=1, FALSE=0)
		{"bool_true_number", "ROUND(TRUE,0)", 1, 0},
		{"bool_false_number", "ROUND(FALSE,0)", 0, 0},
		{"bool_true_digits", "ROUND(3.14,TRUE)", 3.1, 1e-10},
		{"bool_false_digits", "ROUND(3.14,FALSE)", 3, 0},

		// Integer input (no fractional part)
		{"integer", "ROUND(5,2)", 5, 0},
		{"integer_neg", "ROUND(-5,2)", -5, 0},
	}

	const epsilon = 1e-10
	for _, tt := range numTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("Eval(%q) = type %v, want ValueNumber", tt.formula, got.Type)
			}
			tol := tt.tol
			if tol == 0 {
				tol = epsilon
			}
			if math.Abs(got.Num-tt.wantNum) > tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"too_few_args", "ROUND(1)", ErrValVALUE},
		{"too_many_args", "ROUND(1,2,3)", ErrValVALUE},
		{"non_numeric_number", "ROUND(\"abc\",2)", ErrValVALUE},
		{"non_numeric_digits", "ROUND(1,\"abc\")", ErrValVALUE},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = type=%v err=%v, want error %v", tt.formula, got.Type, got.Err, tt.wantErr)
			}
		})
	}
}

func TestRandFunctions(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "RAND()")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(RAND()): %v", err)
	}
	if got.Type != ValueNumber || got.Num < 0 || got.Num >= 1 {
		t.Errorf("RAND() = %g, want [0,1)", got.Num)
	}

	cf = evalCompile(t, "RANDBETWEEN(1,10)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(RANDBETWEEN): %v", err)
	}
	if got.Type != ValueNumber || got.Num < 1 || got.Num > 10 {
		t.Errorf("RANDBETWEEN(1,10) = %g, want [1,10]", got.Num)
	}
}
