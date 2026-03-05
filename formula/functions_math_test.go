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
		{"POWER(-8,1/3)", -2, 1e-10},
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
		{"CEILING(-2.5,1)", -2, 0},
		{"CEILING(-2.5,-1)", -3, 0},
		{"FLOOR(4.9,1)", 4, 0},
		{"FLOOR(4.5,2)", 4, 0},
		{"FLOOR(0,0)", 0, 0},
		{"FLOOR(-2.5,1)", -3, 0},
		{"FLOOR(-2.5,-1)", -2, 0},
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
		{"POWER(0,0)", ErrValNUM},
		{"POWER(0,-1)", ErrValDIV0},
		{"LOG(1,1)", ErrValDIV0},
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

// ---------------------------------------------------------------------------
// SEQUENCE tests
// ---------------------------------------------------------------------------

func TestSEQUENCE_SingleColumn4Rows(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(4)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 4 {
		t.Fatalf("expected 4 rows, got %d", len(got.Array))
	}
	for i, want := range []float64{1, 2, 3, 4} {
		if len(got.Array[i]) != 1 {
			t.Fatalf("row %d: expected 1 col, got %d", i, len(got.Array[i]))
		}
		if got.Array[i][0].Num != want {
			t.Errorf("row %d: got %g, want %g", i, got.Array[i][0].Num, want)
		}
	}
}

func TestSEQUENCE_SingleRow5Cols(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(1,5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 5 {
		t.Fatalf("expected 1x5, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	for i, want := range []float64{1, 2, 3, 4, 5} {
		if got.Array[0][i].Num != want {
			t.Errorf("col %d: got %g, want %g", i, got.Array[0][i].Num, want)
		}
	}
}

func TestSEQUENCE_2x3(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(2,3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{1, 2, 3}, {4, 5, 6}}
	if len(got.Array) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(got.Array))
	}
	for r, wantRow := range want {
		if len(got.Array[r]) != 3 {
			t.Fatalf("row %d: expected 3 cols, got %d", r, len(got.Array[r]))
		}
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestSEQUENCE_CustomStartAndStep(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(3,1,10,10)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	for i, want := range []float64{10, 20, 30} {
		if got.Array[i][0].Num != want {
			t.Errorf("row %d: got %g, want %g", i, got.Array[i][0].Num, want)
		}
	}
}

func TestSEQUENCE_SingleCell(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(1,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("SEQUENCE(1,1) = %v, want 1", got)
	}
}

func TestSEQUENCE_SingleCellCustomStart(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(1,1,42)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("SEQUENCE(1,1,42) = %v, want 42", got)
	}
}

func TestSEQUENCE_FractionalStep(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(2,3,0,0.5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{0, 0.5, 1}, {1.5, 2, 2.5}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestSEQUENCE_NegativeStep(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(1,4,100,-10)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	for i, want := range []float64{100, 90, 80, 70} {
		if got.Array[0][i].Num != want {
			t.Errorf("col %d: got %g, want %g", i, got.Array[0][i].Num, want)
		}
	}
}

func TestSEQUENCE_NegativeStepVertical(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(3,1,10,-5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	for i, want := range []float64{10, 5, 0} {
		if got.Array[i][0].Num != want {
			t.Errorf("row %d: got %g, want %g", i, got.Array[i][0].Num, want)
		}
	}
}

func TestSEQUENCE_ZeroStep(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(3,1,5,0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	for i, want := range []float64{5, 5, 5} {
		if got.Array[i][0].Num != want {
			t.Errorf("row %d: got %g, want %g", i, got.Array[i][0].Num, want)
		}
	}
}

func TestSEQUENCE_StartAtZero(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(3,1,0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	for i, want := range []float64{0, 1, 2} {
		if got.Array[i][0].Num != want {
			t.Errorf("row %d: got %g, want %g", i, got.Array[i][0].Num, want)
		}
	}
}

func TestSEQUENCE_DefaultsOnly3Rows(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	for i, want := range []float64{1, 2, 3} {
		if got.Array[i][0].Num != want {
			t.Errorf("row %d: got %g, want %g", i, got.Array[i][0].Num, want)
		}
	}
}

func TestSEQUENCE_NegativeStart(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(4,1,-2,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	for i, want := range []float64{-2, -1, 0, 1} {
		if got.Array[i][0].Num != want {
			t.Errorf("row %d: got %g, want %g", i, got.Array[i][0].Num, want)
		}
	}
}

func TestSEQUENCE_LargeArray(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(5,5,1,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if len(got.Array) != 5 || len(got.Array[0]) != 5 {
		t.Fatalf("expected 5x5, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	// First row: 1-5, last row: 21-25
	if got.Array[0][0].Num != 1 {
		t.Errorf("[0][0]: got %g, want 1", got.Array[0][0].Num)
	}
	if got.Array[4][4].Num != 25 {
		t.Errorf("[4][4]: got %g, want 25", got.Array[4][4].Num)
	}
}

func TestSEQUENCE_ErrorZeroRows(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("SEQUENCE(0): got %v, want #CALC!", got)
	}
}

func TestSEQUENCE_ErrorNegativeRows(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(-1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("SEQUENCE(-1): got %v, want #CALC!", got)
	}
}

func TestSEQUENCE_ErrorZeroCols(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(2,0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("SEQUENCE(2,0): got %v, want #CALC!", got)
	}
}

func TestSEQUENCE_ErrorTooManyArgs(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(1,1,1,1,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("SEQUENCE(5 args): got %v, want #VALUE!", got)
	}
}

func TestSEQUENCE_ErrorNoArgs(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE()")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("SEQUENCE(): got %v, want #VALUE!", got)
	}
}

func TestSEQUENCE_StringCoercion(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `SEQUENCE("3","2","10","5")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{10, 15}, {20, 25}, {30, 35}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestSEQUENCE_RowsTruncated(t *testing.T) {
	// 2.9 should be truncated to 2
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(2.9,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 2 {
		t.Errorf("expected 2 rows, got %d", len(got.Array))
	}
}

func TestSEQUENCE_ErrorNonNumericRows(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `SEQUENCE("abc")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("SEQUENCE(\"abc\"): got type %v, want error", got.Type)
	}
}

func TestSEQUENCE_3x3StartNegStep(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(3,3,9,-1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	want := [][]float64{{9, 8, 7}, {6, 5, 4}, {3, 2, 1}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

// ---------------------------------------------------------------------------
// BITAND tests
// ---------------------------------------------------------------------------

func TestBITAND(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"basic_1_5", "BITAND(1,5)", 1},
		{"basic_13_25", "BITAND(13,25)", 9},
		{"both_zero", "BITAND(0,0)", 0},
		{"one_zero", "BITAND(0,255)", 0},
		{"zero_other", "BITAND(255,0)", 0},
		{"same_number", "BITAND(42,42)", 42},
		{"all_bits", "BITAND(255,255)", 255},
		{"no_overlap", "BITAND(170,85)", 0},
		{"full_overlap", "BITAND(15,15)", 15},
		{"large_near_max", "BITAND(281474976710655,281474976710655)", 281474976710655},
		{"large_partial", "BITAND(281474976710655,1)", 1},
		{"bool_true", "BITAND(TRUE,3)", 1},
		{"string_num", `BITAND("7",3)`, 3},
		{"power_of_two", "BITAND(8,12)", 8},
		{"mixed_bits", "BITAND(6,3)", 2},
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
			if got.Num != tt.want {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"negative_first", "BITAND(-1,5)", ErrValNUM},
		{"negative_second", "BITAND(5,-1)", ErrValNUM},
		{"non_integer_first", "BITAND(1.5,5)", ErrValNUM},
		{"non_integer_second", "BITAND(5,1.5)", ErrValNUM},
		{"over_max_first", "BITAND(281474976710656,1)", ErrValNUM},
		{"over_max_second", "BITAND(1,281474976710656)", ErrValNUM},
		{"non_numeric_first", `BITAND("abc",1)`, ErrValVALUE},
		{"non_numeric_second", `BITAND(1,"abc")`, ErrValVALUE},
		{"too_few_args", "BITAND(1)", ErrValVALUE},
		{"too_many_args", "BITAND(1,2,3)", ErrValVALUE},
		{"no_args", "BITAND()", ErrValVALUE},
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

// ---------------------------------------------------------------------------
// BITOR tests
// ---------------------------------------------------------------------------

func TestBITOR(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"basic_1_5", "BITOR(1,5)", 5},
		{"basic_13_25", "BITOR(13,25)", 29},
		{"both_zero", "BITOR(0,0)", 0},
		{"one_zero", "BITOR(0,255)", 255},
		{"zero_other", "BITOR(255,0)", 255},
		{"same_number", "BITOR(42,42)", 42},
		{"all_bits", "BITOR(255,255)", 255},
		{"no_overlap", "BITOR(170,85)", 255},
		{"full_overlap", "BITOR(15,15)", 15},
		{"large_near_max", "BITOR(281474976710655,0)", 281474976710655},
		{"bool_true", "BITOR(TRUE,2)", 3},
		{"string_num", `BITOR("4",2)`, 6},
		{"power_of_two", "BITOR(8,4)", 12},
		{"mixed_bits", "BITOR(6,3)", 7},
		{"one_and_two", "BITOR(1,2)", 3},
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
			if got.Num != tt.want {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"negative_first", "BITOR(-1,5)", ErrValNUM},
		{"negative_second", "BITOR(5,-1)", ErrValNUM},
		{"non_integer_first", "BITOR(1.5,5)", ErrValNUM},
		{"non_integer_second", "BITOR(5,1.5)", ErrValNUM},
		{"over_max_first", "BITOR(281474976710656,1)", ErrValNUM},
		{"over_max_second", "BITOR(1,281474976710656)", ErrValNUM},
		{"non_numeric_first", `BITOR("abc",1)`, ErrValVALUE},
		{"non_numeric_second", `BITOR(1,"abc")`, ErrValVALUE},
		{"too_few_args", "BITOR(1)", ErrValVALUE},
		{"too_many_args", "BITOR(1,2,3)", ErrValVALUE},
		{"no_args", "BITOR()", ErrValVALUE},
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

// ---------------------------------------------------------------------------
// BITXOR tests
// ---------------------------------------------------------------------------

func TestBITXOR(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"basic_1_5", "BITXOR(1,5)", 4},
		{"basic_13_25", "BITXOR(13,25)", 20},
		{"both_zero", "BITXOR(0,0)", 0},
		{"one_zero", "BITXOR(0,255)", 255},
		{"zero_other", "BITXOR(255,0)", 255},
		{"same_number", "BITXOR(42,42)", 0},
		{"all_bits", "BITXOR(255,255)", 0},
		{"no_overlap", "BITXOR(170,85)", 255},
		{"full_overlap", "BITXOR(15,15)", 0},
		{"large_near_max", "BITXOR(281474976710655,0)", 281474976710655},
		{"large_xor_self", "BITXOR(281474976710655,281474976710655)", 0},
		{"bool_true", "BITXOR(TRUE,3)", 2},
		{"string_num", `BITXOR("7",3)`, 4},
		{"power_of_two", "BITXOR(8,12)", 4},
		{"mixed_bits", "BITXOR(6,3)", 5},
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
			if got.Num != tt.want {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"negative_first", "BITXOR(-1,5)", ErrValNUM},
		{"negative_second", "BITXOR(5,-1)", ErrValNUM},
		{"non_integer_first", "BITXOR(1.5,5)", ErrValNUM},
		{"non_integer_second", "BITXOR(5,1.5)", ErrValNUM},
		{"over_max_first", "BITXOR(281474976710656,1)", ErrValNUM},
		{"over_max_second", "BITXOR(1,281474976710656)", ErrValNUM},
		{"non_numeric_first", `BITXOR("abc",1)`, ErrValVALUE},
		{"non_numeric_second", `BITXOR(1,"abc")`, ErrValVALUE},
		{"too_few_args", "BITXOR(1)", ErrValVALUE},
		{"too_many_args", "BITXOR(1,2,3)", ErrValVALUE},
		{"no_args", "BITXOR()", ErrValVALUE},
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

// ---------------------------------------------------------------------------
// BITLSHIFT tests
// ---------------------------------------------------------------------------

func TestBITLSHIFT(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"basic_1_3", "BITLSHIFT(1,3)", 8},
		{"basic_4_2", "BITLSHIFT(4,2)", 16},
		{"shift_zero", "BITLSHIFT(255,0)", 255},
		{"number_zero", "BITLSHIFT(0,10)", 0},
		{"both_zero", "BITLSHIFT(0,0)", 0},
		{"shift_one", "BITLSHIFT(1,1)", 2},
		{"large_number", "BITLSHIFT(1,47)", 140737488355328},
		{"negative_shift_right", "BITLSHIFT(16,-2)", 4},
		{"negative_shift_zero", "BITLSHIFT(1,-1)", 0},
		{"bool_true", "BITLSHIFT(TRUE,2)", 4},
		{"string_num", `BITLSHIFT("8",1)`, 16},
		{"shift_by_one", "BITLSHIFT(5,1)", 10},
		{"max_val_shift_zero", "BITLSHIFT(281474976710655,0)", 281474976710655},
		{"power_of_two", "BITLSHIFT(1,10)", 1024},
		{"large_shift_neg", "BITLSHIFT(1024,-10)", 1},
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
			if got.Num != tt.want {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"negative_number", "BITLSHIFT(-1,3)", ErrValNUM},
		{"non_integer_number", "BITLSHIFT(1.5,3)", ErrValNUM},
		{"non_integer_shift", "BITLSHIFT(1,1.5)", ErrValNUM},
		{"over_max_number", "BITLSHIFT(281474976710656,1)", ErrValNUM},
		{"result_over_max", "BITLSHIFT(1,48)", ErrValNUM},
		{"result_over_max_large", "BITLSHIFT(281474976710655,1)", ErrValNUM},
		{"non_numeric_number", `BITLSHIFT("abc",1)`, ErrValVALUE},
		{"non_numeric_shift", `BITLSHIFT(1,"abc")`, ErrValVALUE},
		{"too_few_args", "BITLSHIFT(1)", ErrValVALUE},
		{"too_many_args", "BITLSHIFT(1,2,3)", ErrValVALUE},
		{"no_args", "BITLSHIFT()", ErrValVALUE},
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

// ---------------------------------------------------------------------------
// BITRSHIFT tests
// ---------------------------------------------------------------------------

func TestBITRSHIFT(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"basic_8_3", "BITRSHIFT(8,3)", 1},
		{"basic_16_2", "BITRSHIFT(16,2)", 4},
		{"shift_zero", "BITRSHIFT(255,0)", 255},
		{"number_zero", "BITRSHIFT(0,10)", 0},
		{"both_zero", "BITRSHIFT(0,0)", 0},
		{"shift_one", "BITRSHIFT(2,1)", 1},
		{"large_number", "BITRSHIFT(140737488355328,47)", 1},
		{"negative_shift_left", "BITRSHIFT(4,-2)", 16},
		{"shift_to_zero", "BITRSHIFT(1,1)", 0},
		{"bool_true", "BITRSHIFT(TRUE,0)", 1},
		{"string_num", `BITRSHIFT("16",1)`, 8},
		{"shift_by_one", "BITRSHIFT(10,1)", 5},
		{"max_val_shift_zero", "BITRSHIFT(281474976710655,0)", 281474976710655},
		{"power_of_two", "BITRSHIFT(1024,10)", 1},
		{"large_shift_neg", "BITRSHIFT(1,-10)", 1024},
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
			if got.Num != tt.want {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"negative_number", "BITRSHIFT(-1,3)", ErrValNUM},
		{"non_integer_number", "BITRSHIFT(1.5,3)", ErrValNUM},
		{"non_integer_shift", "BITRSHIFT(1,1.5)", ErrValNUM},
		{"over_max_number", "BITRSHIFT(281474976710656,1)", ErrValNUM},
		{"result_over_max_neg_shift", "BITRSHIFT(1,-48)", ErrValNUM},
		{"result_over_max_large", "BITRSHIFT(281474976710655,-1)", ErrValNUM},
		{"non_numeric_number", `BITRSHIFT("abc",1)`, ErrValVALUE},
		{"non_numeric_shift", `BITRSHIFT(1,"abc")`, ErrValVALUE},
		{"too_few_args", "BITRSHIFT(1)", ErrValVALUE},
		{"too_many_args", "BITRSHIFT(1,2,3)", ErrValVALUE},
		{"no_args", "BITRSHIFT()", ErrValVALUE},
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

func TestSERIESSUM(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Excel documentation example: cos(PI/4) via Taylor series
		{"excel_example", "SERIESSUM(PI()/4,0,2,{1,-0.5,0.041666667,-0.001388889})", 0.707103, 1e-4},
		// Single coefficient
		{"single_coeff", "SERIESSUM(2,3,1,{5})", 40, 0},
		// Single coefficient as scalar
		{"scalar_coeff", "SERIESSUM(2,3,1,5)", 40, 0},
		// x=0, n>0: 0^positive = 0, so result = 0
		{"x_zero_n_pos", "SERIESSUM(0,1,1,{3,4,5})", 0, 0},
		// x=0, n=0: Excel returns #NUM! (0^0 indeterminate) — tested in error cases below
		// x=1: all powers are 1, so result = sum of coefficients
		{"x_one", "SERIESSUM(1,2,3,{2,3,4})", 9, 0},
		// Negative x
		{"neg_x", "SERIESSUM(-2,0,1,{1,1})", -1, 0},
		// Negative n: 2^(-1)=0.5, 2^0=1 => 3*0.5 + 4*1 = 5.5
		{"neg_n", "SERIESSUM(2,-1,1,{3,4})", 5.5, 1e-10},
		// m=0: all same power, result = sum(coeffs) * x^n
		{"m_zero", "SERIESSUM(3,2,0,{1,2,3})", 54, 0},
		// m=1: standard power series 1 + 2 + 4 + 8 = 15
		{"m_one", "SERIESSUM(2,0,1,{1,1,1,1})", 15, 0},
		// Large number of coefficients (10 terms, all 1, x=1)
		{"many_coeffs", "SERIESSUM(1,0,1,{1,1,1,1,1,1,1,1,1,1})", 10, 0},
		// Zero coefficients
		{"zero_coeffs", "SERIESSUM(5,1,1,{0,0,0})", 0, 0},
		// Fractional n and m: sqrt(4)=2, 4^1=4 => 2+4=6... wait, 4^0.5=2, 4^1=4 => 1*2+1*4=6
		{"frac_n_m", "SERIESSUM(4,0.5,0.5,{1,1})", 6, 0},
		// Boolean coercion for x: TRUE=1, 1^0=1, 1^1=1 => 2+3=5
		{"bool_x", "SERIESSUM(TRUE,0,1,{2,3})", 5, 0},
		// Negative m: 2^4=16, 2^2=4 => 16+4=20
		{"neg_m", "SERIESSUM(2,4,-2,{1,1})", 20, 0},
		// Single term with n=0, m=0: 10^0=1 => 3*1=3
		{"single_n0_m0", "SERIESSUM(10,0,0,{3})", 3, 0},
		// Two-row array of coefficients: {1,2;3,4} flattened = [1,2,3,4], x=1 => sum=10
		{"multi_row_array", "SERIESSUM(1,0,1,{1,2;3,4})", 10, 0},
		// x=2, n=0, m=2: 1*1 + 2*4 + 3*16 = 57
		{"power_step_2", "SERIESSUM(2,0,2,{1,2,3})", 57, 0},
		// Very small x: 0.001^0=1 => 1*1 + 1000*0.001 = 2
		{"small_x", "SERIESSUM(0.001,0,1,{1,1000})", 2, 1e-10},
		// x=0, n=0, m=0: Excel returns #NUM! (0^0 indeterminate) — tested in error cases below
	}

	for _, tt := range numTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("Eval(%q): got type %v, want number", tt.formula, got.Type)
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
		})
	}

	// Error tests
	seriesSumErrTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"too_few_args", "SERIESSUM(1,0,1)", ErrValVALUE},
		{"too_many_args", "SERIESSUM(1,0,1,{1},2)", ErrValVALUE},
		{"no_args", "SERIESSUM()", ErrValVALUE},
		{"non_numeric_x", `SERIESSUM("abc",0,1,{1})`, ErrValVALUE},
		{"non_numeric_n", `SERIESSUM(1,"abc",1,{1})`, ErrValVALUE},
		{"non_numeric_m", `SERIESSUM(1,0,"abc",{1})`, ErrValVALUE},
		{"non_numeric_coeff", `SERIESSUM(1,0,1,"abc")`, ErrValVALUE},
		{"x_zero_n_zero", "SERIESSUM(0,0,1,{7,3,5})", ErrValNUM},
		{"x0_n0_m0", "SERIESSUM(0,0,0,{2,3,5})", ErrValNUM},
	}

	for _, tt := range seriesSumErrTests {
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

// ---------------------------------------------------------------------------
// MMULT tests
// ---------------------------------------------------------------------------

func TestMMULT_2x2_times_2x2(t *testing.T) {
	// {1,2;3,4} * {5,6;7,8} = {19,22;43,50}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1,2;3,4},{5,6;7,8})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{19, 22}, {43, 50}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestMMULT_2x3_times_3x2(t *testing.T) {
	// {1,2,3;4,5,6} * {7,8;9,10;11,12} = {58,64;139,154}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1,2,3;4,5,6},{7,8;9,10;11,12})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{58, 64}, {139, 154}}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestMMULT_DotProduct_1x3_times_3x1(t *testing.T) {
	// {1,2,3} * {4;5;6} = {32} (dot product: 1*4+2*5+3*6=32)
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1,2,3},{4;5;6})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 1x1, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[0][0].Num != 32 {
		t.Errorf("got %g, want 32", got.Array[0][0].Num)
	}
}

func TestMMULT_OuterProduct_3x1_times_1x3(t *testing.T) {
	// {1;2;3} * {4,5,6} = {4,5,6;8,10,12;12,15,18}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1;2;3},{4,5,6})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{4, 5, 6}, {8, 10, 12}, {12, 15, 18}}
	if len(got.Array) != 3 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 3x3, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestMMULT_IdentityMatrix(t *testing.T) {
	// {1,2;3,4} * {1,0;0,1} = {1,2;3,4}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1,2;3,4},{1,0;0,1})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{1, 2}, {3, 4}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestMMULT_DimensionMismatch(t *testing.T) {
	// {1,2;3,4} (2x2) * {1,2,3} (1x3) => columns(2) != rows(1) => #VALUE!
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1,2;3,4},{1,2,3})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("dimension mismatch: got %v, want #VALUE!", got)
	}
}

func TestMMULT_TextInArray1(t *testing.T) {
	// Text in first array => #VALUE!
	result, err := fnMMULT([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), StringVal("x")},
			{NumberVal(3), NumberVal(4)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(5), NumberVal(6)},
			{NumberVal(7), NumberVal(8)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("text in array1: got %v, want #VALUE!", result)
	}
}

func TestMMULT_TextInArray2(t *testing.T) {
	// Text in second array => #VALUE!
	result, err := fnMMULT([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(3)},
			{StringVal("hello")},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("text in array2: got %v, want #VALUE!", result)
	}
}

func TestMMULT_EmptyCellInArray(t *testing.T) {
	// Empty cell in array => #VALUE!
	result, err := fnMMULT([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), EmptyVal()},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(3)},
			{NumberVal(4)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("empty cell in array: got %v, want #VALUE!", result)
	}
}

func TestMMULT_SingleElements(t *testing.T) {
	// {5} * {3} = {15}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({5},{3})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if got.Array[0][0].Num != 15 {
		t.Errorf("got %g, want 15", got.Array[0][0].Num)
	}
}

func TestMMULT_ZeroMatrix(t *testing.T) {
	// {0,0;0,0} * {1,2;3,4} = {0,0;0,0}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({0,0;0,0},{1,2;3,4})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			if got.Array[r][c].Num != 0 {
				t.Errorf("[%d][%d]: got %g, want 0", r, c, got.Array[r][c].Num)
			}
		}
	}
}

func TestMMULT_NegativeNumbers(t *testing.T) {
	// {-1,-2;-3,-4} * {1,0;0,1} = {-1,-2;-3,-4}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({-1,-2;-3,-4},{1,0;0,1})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{-1, -2}, {-3, -4}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestMMULT_WrongArgCount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, "MMULT({1,2})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("too few args: got %v, want #VALUE!", got)
	}

	// Too many args
	cf = evalCompile(t, "MMULT({1},{2},{3})")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("too many args: got %v, want #VALUE!", got)
	}

	// No args
	cf = evalCompile(t, "MMULT()")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("no args: got %v, want #VALUE!", got)
	}
}

func TestMMULT_4x4Matrix(t *testing.T) {
	// 4x4 identity * 4x4 matrix = same matrix
	result, err := fnMMULT([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(0), NumberVal(0), NumberVal(0)},
			{NumberVal(0), NumberVal(1), NumberVal(0), NumberVal(0)},
			{NumberVal(0), NumberVal(0), NumberVal(1), NumberVal(0)},
			{NumberVal(0), NumberVal(0), NumberVal(0), NumberVal(1)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3), NumberVal(4)},
			{NumberVal(5), NumberVal(6), NumberVal(7), NumberVal(8)},
			{NumberVal(9), NumberVal(10), NumberVal(11), NumberVal(12)},
			{NumberVal(13), NumberVal(14), NumberVal(15), NumberVal(16)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueArray {
		t.Fatalf("expected array, got %v", result.Type)
	}
	want := [][]float64{
		{1, 2, 3, 4},
		{5, 6, 7, 8},
		{9, 10, 11, 12},
		{13, 14, 15, 16},
	}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if result.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, result.Array[r][c].Num, w)
			}
		}
	}
}

func TestMMULT_ScalarTimesArray(t *testing.T) {
	// scalar 3 coerced to {3} (1x1), times {2;4} (2x1) => dimension mismatch
	// columns(1) == rows(2)? No => #VALUE!
	result, err := fnMMULT([]Value{
		NumberVal(3),
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(2)},
			{NumberVal(4)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("scalar times 2x1: got %v, want #VALUE!", result)
	}
}

func TestMMULT_ScalarTimesScalar(t *testing.T) {
	// scalar 3 * scalar 5 => 1x1 * 1x1 = {15}
	result, err := fnMMULT([]Value{
		NumberVal(3),
		NumberVal(5),
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueArray {
		t.Fatalf("expected array, got %v", result.Type)
	}
	if result.Array[0][0].Num != 15 {
		t.Errorf("got %g, want 15", result.Array[0][0].Num)
	}
}

func TestMMULT_BooleanValues(t *testing.T) {
	// Booleans are numeric: TRUE=1, FALSE=0
	result, err := fnMMULT([]Value{
		{Type: ValueArray, Array: [][]Value{
			{BoolVal(true), BoolVal(false)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(5)},
			{NumberVal(10)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueArray {
		t.Fatalf("expected array, got %v", result.Type)
	}
	// TRUE*5 + FALSE*10 = 5
	if result.Array[0][0].Num != 5 {
		t.Errorf("got %g, want 5", result.Array[0][0].Num)
	}
}

func TestMMULT_FractionalNumbers(t *testing.T) {
	// {0.5,1.5;2.5,3.5} * {1,0;0,1} = {0.5,1.5;2.5,3.5}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({0.5,1.5;2.5,3.5},{1,0;0,1})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{0.5, 1.5}, {2.5, 3.5}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestMMULT_ErrorInArray(t *testing.T) {
	// An error value in the array should propagate
	result, err := fnMMULT([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), ErrorVal(ErrValDIV0)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
			{NumberVal(2)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValDIV0 {
		t.Errorf("error in array: got %v, want #DIV/0!", result)
	}
}

func TestMMULT_3x2_times_2x3(t *testing.T) {
	// {1,2;3,4;5,6} * {1,2,3;4,5,6} = {9,12,15;19,26,33;29,40,51}
	result, err := fnMMULT([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
			{NumberVal(3), NumberVal(4)},
			{NumberVal(5), NumberVal(6)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueArray {
		t.Fatalf("expected array, got %v", result.Type)
	}
	want := [][]float64{{9, 12, 15}, {19, 26, 33}, {29, 40, 51}}
	if len(result.Array) != 3 || len(result.Array[0]) != 3 {
		t.Fatalf("expected 3x3, got %dx%d", len(result.Array), len(result.Array[0]))
	}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if result.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, result.Array[r][c].Num, w)
			}
		}
	}
}

func TestMMULT_WithCellRange(t *testing.T) {
	// Use cell references: A1:B2 * C1:D2
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 2, Row: 2}: NumberVal(4),
			{Col: 3, Row: 1}: NumberVal(5),
			{Col: 4, Row: 1}: NumberVal(6),
			{Col: 3, Row: 2}: NumberVal(7),
			{Col: 4, Row: 2}: NumberVal(8),
		},
	}
	cf := evalCompile(t, "MMULT(A1:B2,C1:D2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{19, 22}, {43, 50}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestMMULT_LargeValues(t *testing.T) {
	// Large numbers: {1000,2000;3000,4000} * {5000,6000;7000,8000}
	// = {1000*5000+2000*7000, 1000*6000+2000*8000; 3000*5000+4000*7000, 3000*6000+4000*8000}
	// = {19000000, 22000000; 43000000, 50000000}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1000,2000;3000,4000},{5000,6000;7000,8000})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{19000000, 22000000}, {43000000, 50000000}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

// ---------------------------------------------------------------------------
// MDETERM tests
// ---------------------------------------------------------------------------

func TestMDETERM(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		formula string
		want    float64
		tol     float64
	}{
		// 1x1 matrix
		{"MDETERM({5})", 5, 0},
		{"MDETERM({0})", 0, 0},
		{"MDETERM({-7})", -7, 0},

		// 2x2 matrices
		{"MDETERM({3,6;1,1})", -3, 0},
		{"MDETERM({1,0;0,1})", 1, 0},    // identity
		{"MDETERM({0,0;0,0})", 0, 0},    // zero matrix
		{"MDETERM({2,0;0,3})", 6, 0},    // diagonal
		{"MDETERM({1,2;2,4})", 0, 0},    // singular
		{"MDETERM({-1,-2;3,4})", 2, 0},  // negative numbers

		// 3x3 matrices
		{"MDETERM({3,6,1;1,1,0;3,10,2})", 1, 1e-10},
		{"MDETERM({1,0,0;0,1,0;0,0,1})", 1, 0},        // 3x3 identity
		{"MDETERM({2,0,0;0,3,0;0,0,4})", 24, 1e-10},    // diagonal
		{"MDETERM({1,2,3;4,5,6;7,8,9})", 0, 1e-10},     // singular

		// 4x4 matrix
		{"MDETERM({1,3,8,5;1,3,6,1;1,1,1,0;7,3,10,2})", 88, 1e-10},

		// 5x5 identity
		{"MDETERM({1,0,0,0,0;0,1,0,0,0;0,0,1,0,0;0,0,0,1,0;0,0,0,0,1})", 1, 0},

		// Fractional values
		{"MDETERM({0.5,1.5;2.5,3.5})", -2, 1e-10},
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

func TestMDETERM_Errors(t *testing.T) {
	resolver := &mockResolver{}

	errTests := []struct {
		name    string
		formula string
		errVal  ErrorValue
	}{
		{"non-square 2x3", "MDETERM({1,2,3;4,5,6})", ErrValVALUE},
		{"non-square 3x2", "MDETERM({1,2;3,4;5,6})", ErrValVALUE},
		{"single row 1x3", "MDETERM({1,2,3})", ErrValVALUE},
		{"wrong args: none", "MDETERM()", ErrValVALUE},
		{"wrong args: two", "MDETERM({1},{2})", ErrValVALUE},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.errVal {
				t.Errorf("Eval(%q) = type=%v err=%v, want %v", tt.formula, got.Type, got.Err, tt.errVal)
			}
		})
	}
}

func TestMDETERM_TextInArray(t *testing.T) {
	result, err := fnMDETERM([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), StringVal("x")},
			{NumberVal(3), NumberVal(4)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMDETERM: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("text in array: got %v, want #VALUE!", result)
	}
}

func TestMDETERM_EmptyInArray(t *testing.T) {
	result, err := fnMDETERM([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), EmptyVal()},
			{NumberVal(3), NumberVal(4)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMDETERM: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("empty in array: got %v, want #VALUE!", result)
	}
}

func TestMDETERM_ErrorPropagation(t *testing.T) {
	result, err := fnMDETERM([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), ErrorVal(ErrValDIV0)},
			{NumberVal(3), NumberVal(4)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMDETERM: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValDIV0 {
		t.Errorf("error propagation: got %v, want #DIV/0!", result)
	}
}

func TestMDETERM_ErrorArgPropagation(t *testing.T) {
	result, err := fnMDETERM([]Value{
		ErrorVal(ErrValNA),
	})
	if err != nil {
		t.Fatalf("fnMDETERM: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValNA {
		t.Errorf("error arg propagation: got %v, want #N/A", result)
	}
}

func TestMDETERM_ScalarInput(t *testing.T) {
	// Scalar input is treated as 1x1 matrix.
	result, err := fnMDETERM([]Value{NumberVal(42)})
	if err != nil {
		t.Fatalf("fnMDETERM: %v", err)
	}
	if result.Type != ValueNumber || result.Num != 42 {
		t.Errorf("scalar input: got %v, want 42", result)
	}
}

func TestMDETERM_BooleanValues(t *testing.T) {
	// TRUE=1, FALSE=0 => det({TRUE,FALSE;FALSE,TRUE}) = 1*1 - 0*0 = 1
	result, err := fnMDETERM([]Value{
		{Type: ValueArray, Array: [][]Value{
			{BoolVal(true), BoolVal(false)},
			{BoolVal(false), BoolVal(true)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMDETERM: %v", err)
	}
	if result.Type != ValueNumber || result.Num != 1 {
		t.Errorf("boolean identity: got %v, want 1", result)
	}
}

func TestMDETERM_LargeValues(t *testing.T) {
	// {1000,2000;3000,4000} => det = 1000*4000 - 2000*3000 = -2000000
	resolver := &mockResolver{}
	cf := evalCompile(t, "MDETERM({1000,2000;3000,4000})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber {
		t.Fatalf("expected number, got %v", got.Type)
	}
	if got.Num != -2000000 {
		t.Errorf("got %g, want -2000000", got.Num)
	}
}

func TestMDETERM_PermutationMatrix(t *testing.T) {
	// Permutation matrix: det = -1 (single swap from identity)
	result, err := fnMDETERM([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(0), NumberVal(1), NumberVal(0)},
			{NumberVal(1), NumberVal(0), NumberVal(0)},
			{NumberVal(0), NumberVal(0), NumberVal(1)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMDETERM: %v", err)
	}
	if result.Type != ValueNumber {
		t.Fatalf("expected number, got %v", result.Type)
	}
	if result.Num != -1 {
		t.Errorf("permutation matrix: got %g, want -1", result.Num)
	}
}

// ---------------------------------------------------------------------------
// MINVERSE tests
// ---------------------------------------------------------------------------

func TestMINVERSE_1x1(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({2})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 1x1 array, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if math.Abs(got.Array[0][0].Num-0.5) > 1e-10 {
		t.Errorf("got %g, want 0.5", got.Array[0][0].Num)
	}
}

func TestMINVERSE_2x2(t *testing.T) {
	// {1,2;3,4} => inverse is {-2,1;1.5,-0.5}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({1,2;3,4})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{-2, 1}, {1.5, -0.5}}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			if math.Abs(got.Array[r][c].Num-want[r][c]) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestMINVERSE_2x2_Known(t *testing.T) {
	// {4,7;2,6} => inverse is {0.6,-0.7;-0.2,0.4}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({4,7;2,6})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{0.6, -0.7}, {-0.2, 0.4}}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			if math.Abs(got.Array[r][c].Num-want[r][c]) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestMINVERSE_3x3_Identity(t *testing.T) {
	// Inverse of identity is identity.
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({1,0,0;0,1,0;0,0,1})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	for r := 0; r < 3; r++ {
		for c := 0; c < 3; c++ {
			want := 0.0
			if r == c {
				want = 1.0
			}
			if math.Abs(got.Array[r][c].Num-want) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want)
			}
		}
	}
}

func TestMINVERSE_3x3_General(t *testing.T) {
	// {1,2,3;0,1,4;5,6,0} => inverse is {-24,18,5;20,-15,-4;-5,4,1}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({1,2,3;0,1,4;5,6,0})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{-24, 18, 5}, {20, -15, -4}, {-5, 4, 1}}
	for r := 0; r < 3; r++ {
		for c := 0; c < 3; c++ {
			if math.Abs(got.Array[r][c].Num-want[r][c]) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestMINVERSE_Singular(t *testing.T) {
	// {1,2;2,4} is singular (det=0) => #NUM!
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({1,2;2,4})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("expected #NUM!, got %v", got)
	}
}

func TestMINVERSE_NonSquare(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({1,2,3;4,5,6})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", got)
	}
}

func TestMINVERSE_TextInArray(t *testing.T) {
	result, err := fnMINVERSE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), StringVal("x")},
			{NumberVal(3), NumberVal(4)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", result)
	}
}

func TestMINVERSE_EmptyInArray(t *testing.T) {
	result, err := fnMINVERSE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), {Type: ValueEmpty}},
			{NumberVal(3), NumberVal(4)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", result)
	}
}

func TestMINVERSE_WrongArgCount(t *testing.T) {
	// No args.
	result, err := fnMINVERSE([]Value{})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", result)
	}

	// Two args.
	result, err = fnMINVERSE([]Value{NumberVal(1), NumberVal(2)})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", result)
	}
}

func TestMINVERSE_ScalarInput(t *testing.T) {
	// Scalar 5 => 1/5 = 0.2
	result, err := fnMINVERSE([]Value{NumberVal(5)})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueNumber {
		t.Fatalf("expected number, got %v", result.Type)
	}
	if math.Abs(result.Num-0.2) > 1e-10 {
		t.Errorf("got %g, want 0.2", result.Num)
	}
}

func TestMINVERSE_ScalarZero(t *testing.T) {
	// Scalar 0 => singular => #NUM!
	result, err := fnMINVERSE([]Value{NumberVal(0)})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValNUM {
		t.Errorf("expected #NUM!, got %v", result)
	}
}

func TestMINVERSE_DiagonalMatrix(t *testing.T) {
	// Diagonal {2,0;0,4} => inverse {0.5,0;0,0.25}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({2,0;0,4})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{0.5, 0}, {0, 0.25}}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			if math.Abs(got.Array[r][c].Num-want[r][c]) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestMINVERSE_3x3_Diagonal(t *testing.T) {
	// {2,0,0;0,5,0;0,0,10} => {0.5,0,0;0,0.2,0;0,0,0.1}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({2,0,0;0,5,0;0,0,10})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{0.5, 0, 0}, {0, 0.2, 0}, {0, 0, 0.1}}
	for r := 0; r < 3; r++ {
		for c := 0; c < 3; c++ {
			if math.Abs(got.Array[r][c].Num-want[r][c]) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestMINVERSE_NegativeNumbers(t *testing.T) {
	// {-1,0;0,-2} => {-1,0;0,-0.5}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({-1,0;0,-2})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{-1, 0}, {0, -0.5}}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			if math.Abs(got.Array[r][c].Num-want[r][c]) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestMINVERSE_VerifyProduct(t *testing.T) {
	// A * A^-1 should be identity. Using {1,2;3,4}.
	// A^-1 = {-2,1;1.5,-0.5}
	// Product: {1*-2+2*1.5, 1*1+2*-0.5; 3*-2+4*1.5, 3*1+4*-0.5} = {1,0;0,1}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1,2;3,4},MINVERSE({1,2;3,4}))")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			want := 0.0
			if r == c {
				want = 1.0
			}
			if math.Abs(got.Array[r][c].Num-want) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want)
			}
		}
	}
}

func TestMINVERSE_BooleanValues(t *testing.T) {
	// {TRUE,FALSE;FALSE,TRUE} = identity => inverse is identity
	result, err := fnMINVERSE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{BoolVal(true), BoolVal(false)},
			{BoolVal(false), BoolVal(true)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueArray {
		t.Fatalf("expected array, got %v", result.Type)
	}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			want := 0.0
			if r == c {
				want = 1.0
			}
			if math.Abs(result.Array[r][c].Num-want) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, result.Array[r][c].Num, want)
			}
		}
	}
}

func TestMINVERSE_ErrorPropagation(t *testing.T) {
	result, err := fnMINVERSE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), ErrorVal(ErrValDIV0)},
			{NumberVal(3), NumberVal(4)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValDIV0 {
		t.Errorf("expected #DIV/0!, got %v", result)
	}
}

func TestMINVERSE_ErrorArgPropagation(t *testing.T) {
	result, err := fnMINVERSE([]Value{
		ErrorVal(ErrValREF),
	})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValREF {
		t.Errorf("expected #REF!, got %v", result)
	}
}

func TestMINVERSE_3x3_Singular(t *testing.T) {
	// {1,2,3;4,5,6;7,8,9} is singular => #NUM!
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({1,2,3;4,5,6;7,8,9})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("expected #NUM!, got %v", got)
	}
}

func TestMINVERSE_Fractional(t *testing.T) {
	// {0.5,0;0,2} => {2,0;0,0.5}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({0.5,0;0,2})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{2, 0}, {0, 0.5}}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			if math.Abs(got.Array[r][c].Num-want[r][c]) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestERF(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		{"zero", "ERF(0)", 0},
		{"one", "ERF(1)", 0.8427007929497149},
		{"two", "ERF(2)", 0.9953222650189527},
		{"negative_one", "ERF(-1)", -0.8427007929497149},
		{"half", "ERF(0.5)", 0.5204998778130465},
		{"0.745", "ERF(0.745)", 0.7079289200957377},
		{"large", "ERF(6)", 1.0},
		{"large_neg", "ERF(-6)", -1.0},
		{"small", "ERF(0.01)", 0.011283415555849618},
		{"bool_true", "ERF(TRUE)", 0.8427007929497149},
		{"bool_false", "ERF(FALSE)", 0},
		{"string_num", "ERF(\"1\")", 0.8427007929497149},
		{"two_args_0_1", "ERF(0,1)", 0.8427007929497149},
		{"two_args_1_2", "ERF(1,2)", 0.15262147207903781},
		{"two_args_same", "ERF(1,1)", 0},
		{"two_args_neg", "ERF(-1,1)", 1.6854015858994299},
		{"two_args_reverse", "ERF(2,1)", -0.15262147207903781},
		{"string_num_two", "ERF(\"0\",\"1\")", 0.8427007929497149},
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
		{"no_args", "ERF()", ErrValVALUE},
		{"too_many_args", "ERF(1,2,3)", ErrValVALUE},
		{"non_numeric", "ERF(\"abc\")", ErrValVALUE},
		{"non_numeric_second", "ERF(1,\"abc\")", ErrValVALUE},
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

func TestERFPRECISE(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		{"zero", "ERF.PRECISE(0)", 0},
		{"one", "ERF.PRECISE(1)", 0.8427007929497149},
		{"two", "ERF.PRECISE(2)", 0.9953222650189527},
		{"negative_one", "ERF.PRECISE(-1)", -0.8427007929497149},
		{"half", "ERF.PRECISE(0.5)", 0.5204998778130465},
		{"0.745", "ERF.PRECISE(0.745)", 0.7079289200957377},
		{"large", "ERF.PRECISE(6)", 1.0},
		{"large_neg", "ERF.PRECISE(-6)", -1.0},
		{"small", "ERF.PRECISE(0.01)", 0.011283415555849618},
		{"bool_true", "ERF.PRECISE(TRUE)", 0.8427007929497149},
		{"bool_false", "ERF.PRECISE(FALSE)", 0},
		{"string_num", "ERF.PRECISE(\"1\")", 0.8427007929497149},
		{"three", "ERF.PRECISE(3)", 0.9999779095030014},
		{"neg_half", "ERF.PRECISE(-0.5)", -0.5204998778130465},
		{"quarter", "ERF.PRECISE(0.25)", 0.27632639016823696},
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
		{"no_args", "ERF.PRECISE()", ErrValVALUE},
		{"too_many_args", "ERF.PRECISE(1,2)", ErrValVALUE},
		{"non_numeric", "ERF.PRECISE(\"abc\")", ErrValVALUE},
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

func TestERFC(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		{"zero", "ERFC(0)", 1},
		{"one", "ERFC(1)", 0.15729920705028513},
		{"two", "ERFC(2)", 0.004677734981047266},
		{"negative_one", "ERFC(-1)", 1.8427007929497148},
		{"half", "ERFC(0.5)", 0.4795001221869535},
		{"0.745", "ERFC(0.745)", 0.2920710799042623},
		{"large", "ERFC(6)", 0},
		{"large_neg", "ERFC(-6)", 2},
		{"small", "ERFC(0.01)", 0.9887165844441506},
		{"bool_true", "ERFC(TRUE)", 0.15729920705028513},
		{"bool_false", "ERFC(FALSE)", 1},
		{"string_num", "ERFC(\"1\")", 0.15729920705028513},
		{"three", "ERFC(3)", 0.000022090496998585438},
		{"neg_half", "ERFC(-0.5)", 1.5204998778130465},
		{"quarter", "ERFC(0.25)", 0.723673609831763},
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
		{"no_args", "ERFC()", ErrValVALUE},
		{"too_many_args", "ERFC(1,2)", ErrValVALUE},
		{"non_numeric", "ERFC(\"abc\")", ErrValVALUE},
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

func TestERFCPRECISE(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		{"zero", "ERFC.PRECISE(0)", 1},
		{"one", "ERFC.PRECISE(1)", 0.15729920705028513},
		{"two", "ERFC.PRECISE(2)", 0.004677734981047266},
		{"negative_one", "ERFC.PRECISE(-1)", 1.8427007929497148},
		{"half", "ERFC.PRECISE(0.5)", 0.4795001221869535},
		{"0.745", "ERFC.PRECISE(0.745)", 0.2920710799042623},
		{"large", "ERFC.PRECISE(6)", 0},
		{"large_neg", "ERFC.PRECISE(-6)", 2},
		{"small", "ERFC.PRECISE(0.01)", 0.9887165844441506},
		{"bool_true", "ERFC.PRECISE(TRUE)", 0.15729920705028513},
		{"bool_false", "ERFC.PRECISE(FALSE)", 1},
		{"string_num", "ERFC.PRECISE(\"1\")", 0.15729920705028513},
		{"three", "ERFC.PRECISE(3)", 0.000022090496998585438},
		{"neg_half", "ERFC.PRECISE(-0.5)", 1.5204998778130465},
		{"quarter", "ERFC.PRECISE(0.25)", 0.723673609831763},
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
		{"no_args", "ERFC.PRECISE()", ErrValVALUE},
		{"too_many_args", "ERFC.PRECISE(1,2)", ErrValVALUE},
		{"non_numeric", "ERFC.PRECISE(\"abc\")", ErrValVALUE},
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
