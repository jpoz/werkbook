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

func TestPOWER(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
		tol     float64
	}{
		// Basic exponentiation
		{"2_cubed", "POWER(2,3)", 8, 0},
		{"5_squared", "POWER(5,2)", 25, 0},

		// Power of 0 → 1
		{"power_of_zero", "POWER(7,0)", 1, 0},
		{"large_base_power_zero", "POWER(999,0)", 1, 0},

		// Power of 1 → same number
		{"power_of_one", "POWER(7,1)", 7, 0},
		{"neg_base_power_one", "POWER(-3,1)", -3, 0},

		// Number 0, positive power → 0
		{"zero_base_pos_power", "POWER(0,5)", 0, 0},
		{"zero_base_large_power", "POWER(0,100)", 0, 0},

		// Negative base, integer power
		{"neg_base_even_power", "POWER(-3,2)", 9, 0},
		{"neg_base_odd_power", "POWER(-3,3)", -27, 0},
		{"neg_base_power_4", "POWER(-2,4)", 16, 0},

		// Negative base, fractional power (odd root) → real result
		{"neg_base_cube_root", "POWER(-8,1/3)", -2, 1e-10},

		// Fractional power (square root)
		{"square_root", "POWER(4,0.5)", 2, 1e-10},
		{"cube_root_pos", "POWER(27,1/3)", 3, 1e-10},
		{"fourth_root", "POWER(16,0.25)", 2, 1e-10},

		// Negative power (reciprocal)
		{"neg_power", "POWER(2,-1)", 0.5, 1e-10},
		{"neg_power_squared", "POWER(2,-2)", 0.25, 1e-10},
		{"neg_power_3", "POWER(10,-3)", 0.001, 1e-10},

		// Large values
		{"large_result", "POWER(10,10)", 1e10, 0},
		{"large_base", "POWER(100,3)", 1e6, 0},

		// String coercion
		{"string_base", "POWER(\"2\",3)", 8, 0},
		{"string_power", "POWER(2,\"3\")", 8, 0},
		{"string_both", "POWER(\"5\",\"2\")", 25, 0},

		// Boolean coercion
		{"bool_true_base", "POWER(TRUE,5)", 1, 0},
		{"bool_false_base_pos", "POWER(FALSE,5)", 0, 0},
		{"bool_true_power", "POWER(3,TRUE)", 3, 0},
		{"bool_false_power", "POWER(3,FALSE)", 1, 0},

		// Excel doc examples
		{"doc_ex1", "POWER(5,2)", 25, 0},
		{"doc_ex2", "POWER(98.6,3.2)", 2401077.2220695773, 1e-3},
		{"doc_ex3", "POWER(4,5/4)", 5.656854249, 1e-6},
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
				if got.Num != tt.wantNum {
					t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
				}
			} else {
				if math.Abs(got.Num-tt.wantNum) > tt.tol {
					t.Errorf("Eval(%q) = %g, want %g (tol %g)", tt.formula, got.Num, tt.wantNum, tt.tol)
				}
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// 0^0 → #NUM!
		{"zero_power_zero", "POWER(0,0)", ErrValNUM},
		// 0^negative → #DIV/0!
		{"zero_neg_power", "POWER(0,-1)", ErrValDIV0},
		{"zero_neg_power_2", "POWER(0,-5)", ErrValDIV0},
		// Negative base, fractional (non-odd-root) power → #NUM!
		{"neg_base_frac_power", "POWER(-4,0.5)", ErrValNUM},
		{"neg_base_even_frac", "POWER(-8,1/4)", ErrValNUM},
		{"neg_base_arb_frac", "POWER(-2,1.5)", ErrValNUM},
		// Too few args
		{"no_args", "POWER()", ErrValVALUE},
		{"one_arg", "POWER(2)", ErrValVALUE},
		// Too many args
		{"three_args", "POWER(2,3,4)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric_base", "POWER(\"abc\",2)", ErrValVALUE},
		{"non_numeric_power", "POWER(2,\"abc\")", ErrValVALUE},
		// Error propagation
		{"err_div0_base", "POWER(1/0,2)", ErrValDIV0},
		{"err_div0_power", "POWER(2,1/0)", ErrValDIV0},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = type=%v err=%v, want %v", tt.formula, got.Type, got.Err, tt.wantErr)
			}
		})
	}
}

func TestABS(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Positive numbers → unchanged
		{"pos_int", "ABS(5)", 5},
		{"pos_decimal", "ABS(3.14)", 3.14},

		// Negative numbers → positive
		{"neg_int", "ABS(-5)", 5},
		{"neg_decimal", "ABS(-3.14)", 3.14},

		// Zero
		{"zero", "ABS(0)", 0},

		// Large values
		{"large_pos", "ABS(999999999)", 999999999},
		{"large_neg", "ABS(-999999999)", 999999999},

		// Small decimals
		{"small_pos", "ABS(0.000001)", 0.000001},
		{"small_neg", "ABS(-0.000001)", 0.000001},

		// String coercion
		{"string_pos", "ABS(\"5\")", 5},
		{"string_neg", "ABS(\"-3.5\")", 3.5},

		// Boolean coercion
		{"bool_true", "ABS(TRUE)", 1},
		{"bool_false", "ABS(FALSE)", 0},

		// Expression argument
		{"expr_neg_result", "ABS(2-5)", 3},
		{"expr_pos_result", "ABS(5-2)", 3},

		// Excel doc examples
		{"doc_ex1", "ABS(2)", 2},
		{"doc_ex2", "ABS(-2)", 2},
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
			if got.Num != tt.wantNum {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No args
		{"no_args", "ABS()", ErrValVALUE},
		// Too many args
		{"too_many_args", "ABS(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", "ABS(\"abc\")", ErrValVALUE},
		// Error propagation
		{"err_div0", "ABS(1/0)", ErrValDIV0},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = type=%v err=%v, want %v", tt.formula, got.Type, got.Err, tt.wantErr)
			}
		})
	}
}

func TestFLOOR(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Basic: round down to nearest integer (significance=1)
		{"pos_int_floor", "FLOOR(4.9,1)", 4},
		{"pos_half_floor", "FLOOR(2.5,1)", 2},
		{"pos_exact", "FLOOR(7,1)", 7},
		{"pos_small_frac", "FLOOR(0.1,1)", 0},

		// Various significance (0.5, 0.1, 10, 100)
		{"sig_0.5", "FLOOR(4.7,0.5)", 4.5},
		{"sig_0.1", "FLOOR(4.29,0.1)", 4.2},
		{"sig_10", "FLOOR(42,10)", 40},
		{"sig_100", "FLOOR(450,100)", 400},

		// Negative number with negative significance (rounds toward zero)
		{"neg_neg_sig", "FLOOR(-2.5,-1)", -2},
		{"neg_neg_sig_2", "FLOOR(-2.5,-2)", -2},
		{"neg_neg_sig_5", "FLOOR(-7,-5)", -5},

		// Zero number → 0 (regardless of significance, even sig=0)
		{"zero_pos_sig", "FLOOR(0,1)", 0},
		{"zero_neg_sig", "FLOOR(0,5)", 0},
		{"zero_zero", "FLOOR(0,0)", 0},

		// At boundary → unchanged
		{"exact_boundary", "FLOOR(6,3)", 6},
		{"exact_boundary_neg", "FLOOR(-6,-3)", -6},
		{"exact_boundary_2", "FLOOR(10,5)", 10},

		// Negative number with positive significance (rounds away from zero)
		{"neg_pos_sig", "FLOOR(-2.5,1)", -3},
		{"neg_pos_sig_2", "FLOOR(-4.1,2)", -6},

		// Large numbers
		{"large_num", "FLOOR(1234567,1000)", 1234000},
		{"large_exact", "FLOOR(1000000,1000)", 1000000},

		// Decimal significance
		{"dec_sig_0.01", "FLOOR(0.234,0.01)", 0.23},
		{"dec_sig_0.05", "FLOOR(4.42,0.05)", 4.4},

		// Excel doc examples
		{"doc_ex1", "FLOOR(3.7,2)", 2},
		{"doc_ex2", "FLOOR(-2.5,-2)", -2},
		{"doc_ex3", "FLOOR(1.58,0.1)", 1.5},
		{"doc_ex4", "FLOOR(0.234,0.01)", 0.23},

		// String coercion of numeric strings
		{"string_num", "FLOOR(\"4.9\",1)", 4},
		{"string_sig", "FLOOR(4.9,\"1\")", 4},
		{"string_both", "FLOOR(\"4.9\",\"1\")", 4},

		// Boolean coercion
		{"bool_true_num", "FLOOR(TRUE,1)", 1},
		{"bool_false_num", "FLOOR(FALSE,1)", 0},
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
		// Too few args
		{"no_args", "FLOOR()", ErrValVALUE},
		{"one_arg", "FLOOR(1)", ErrValVALUE},
		// Too many args
		{"too_many_args", "FLOOR(1,2,3)", ErrValVALUE},
		// Non-numeric
		{"non_numeric_num", "FLOOR(\"abc\",1)", ErrValVALUE},
		{"non_numeric_sig", "FLOOR(1,\"abc\")", ErrValVALUE},
		// Positive number with negative significance → #NUM!
		{"pos_neg_sig", "FLOOR(2.5,-1)", ErrValNUM},
		{"pos_neg_sig_2", "FLOOR(2.5,-2)", ErrValNUM},
		{"pos_neg_sig_3", "FLOOR(10,-5)", ErrValNUM},
		// Significance = 0 with non-zero number → #DIV/0!
		{"sig_zero_pos", "FLOOR(2.5,0)", ErrValDIV0},
		{"sig_zero_neg", "FLOOR(-2.5,0)", ErrValDIV0},
		// Error propagation
		{"err_prop_num", "FLOOR(1/0,1)", ErrValDIV0},
		{"err_prop_sig", "FLOOR(1,1/0)", ErrValDIV0},
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

func TestCEILING(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Basic: round up to nearest integer (significance=1)
		{"pos_int_ceil", "CEILING(4.1,1)", 5},
		{"pos_half_ceil", "CEILING(2.5,1)", 3},
		{"pos_exact", "CEILING(7,1)", 7},
		{"pos_small_frac", "CEILING(0.1,1)", 1},

		// Significance of 0.5, 0.1, 10, 100
		{"sig_0.5", "CEILING(4.2,0.5)", 4.5},
		{"sig_0.1", "CEILING(4.21,0.1)", 4.3},
		{"sig_10", "CEILING(42,10)", 50},
		{"sig_100", "CEILING(450,100)", 500},

		// Negative number with negative significance (rounds away from zero)
		{"neg_neg_sig", "CEILING(-2.5,-1)", -3},
		{"neg_neg_sig_2", "CEILING(-2.5,-2)", -4},
		{"neg_neg_sig_5", "CEILING(-7,-5)", -10},

		// Negative number with positive significance (rounds toward zero)
		{"neg_pos_sig", "CEILING(-2.5,1)", -2},
		{"neg_pos_sig_2", "CEILING(-4.1,2)", -4},
		{"neg_pos_sig_5", "CEILING(-8.1,5)", -5},

		// Zero → 0
		{"zero_zero", "CEILING(0,1)", 0},
		{"zero_neg_sig", "CEILING(0,5)", 0},

		// Number already at significance boundary → unchanged
		{"exact_boundary", "CEILING(6,3)", 6},
		{"exact_boundary_neg", "CEILING(-6,-3)", -6},
		{"exact_boundary_neg_pos_sig", "CEILING(-6,3)", -6},

		// Significance = 0 → 0
		{"sig_zero_pos", "CEILING(6.3,0)", 0},
		{"sig_zero_neg", "CEILING(-6.3,0)", 0},
		{"sig_zero_zero", "CEILING(0,0)", 0},

		// Very small significance
		{"small_sig", "CEILING(0.234,0.01)", 0.24},
		{"small_sig_2", "CEILING(1.001,0.001)", 1.001},

		// Large numbers
		{"large_num", "CEILING(1234567,1000)", 1235000},
		{"large_exact", "CEILING(1000000,1000)", 1000000},

		// Decimal significance
		{"dec_sig_0.05", "CEILING(4.42,0.05)", 4.45},
		{"dec_sig_0.25", "CEILING(1.1,0.25)", 1.25},

		// Excel doc examples
		{"doc_ex1", "CEILING(2.5,1)", 3},
		{"doc_ex2", "CEILING(-2.5,-2)", -4},
		{"doc_ex3", "CEILING(-2.5,2)", -2},
		{"doc_ex4", "CEILING(1.5,0.1)", 1.5},
		{"doc_ex5", "CEILING(0.234,0.01)", 0.24},

		// String coercion of numeric strings
		{"string_num", "CEILING(\"4.1\",1)", 5},
		{"string_sig", "CEILING(4.1,\"1\")", 5},
		{"string_both", "CEILING(\"4.1\",\"1\")", 5},

		// Boolean coercion
		{"bool_true_num", "CEILING(TRUE,1)", 1},
		{"bool_false_num", "CEILING(FALSE,1)", 0},
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
		// Too few args
		{"no_args", "CEILING()", ErrValVALUE},
		{"one_arg", "CEILING(1)", ErrValVALUE},
		// Too many args
		{"too_many_args", "CEILING(1,2,3)", ErrValVALUE},
		// Non-numeric
		{"non_numeric_num", "CEILING(\"abc\",1)", ErrValVALUE},
		{"non_numeric_sig", "CEILING(1,\"abc\")", ErrValVALUE},
		// Positive number with negative significance → #NUM!
		{"pos_neg_sig", "CEILING(2.5,-1)", ErrValNUM},
		{"pos_neg_sig_2", "CEILING(10,-5)", ErrValNUM},
		// Error propagation
		{"err_prop_num", "CEILING(1/0,1)", ErrValDIV0},
		{"err_prop_sig", "CEILING(1,1/0)", ErrValDIV0},
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

func TestROUNDDOWN(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
		tol     float64
	}{
		// Basic positive number, various decimal places
		{"basic_2dp", "ROUNDDOWN(3.14159,2)", 3.14, 1e-10},
		{"basic_1dp", "ROUNDDOWN(2.19,1)", 2.1, 1e-10},
		{"basic_3dp", "ROUNDDOWN(1.23456,3)", 1.234, 1e-10},

		// Zero digits
		{"zero_digits_pos", "ROUNDDOWN(3.7,0)", 3, 0},
		{"zero_digits_pos2", "ROUNDDOWN(3.2,0)", 3, 0},

		// Negative numbers (toward zero means truncation)
		{"neg_0dp", "ROUNDDOWN(-3.7,0)", -3, 0},
		{"neg_1dp", "ROUNDDOWN(-3.14159,1)", -3.1, 1e-10},
		{"neg_2dp", "ROUNDDOWN(-2.789,2)", -2.78, 1e-10},

		// Negative num_digits (round left of decimal)
		{"neg_digits_tens", "ROUNDDOWN(123.456,-1)", 120, 0},
		{"neg_digits_hundreds", "ROUNDDOWN(31415.92654,-2)", 31400, 0},
		{"neg_digits_thousands", "ROUNDDOWN(9876,-3)", 9000, 0},

		// Zero input
		{"zero_input", "ROUNDDOWN(0,2)", 0, 0},
		{"zero_input_neg_digits", "ROUNDDOWN(0,-1)", 0, 0},

		// Already rounded (no change expected)
		{"already_rounded", "ROUNDDOWN(3.14,2)", 3.14, 1e-10},
		{"already_integer", "ROUNDDOWN(5,0)", 5, 0},
		{"already_integer_dp", "ROUNDDOWN(5,3)", 5, 0},

		// Large positive num_digits
		{"large_digits", "ROUNDDOWN(3.14159265,8)", 3.14159265, 1e-10},

		// Negative num_digits with large numbers
		{"neg_digits_large", "ROUNDDOWN(123456789,-4)", 123450000, 0},

		// Boolean coercion (TRUE=1, FALSE=0)
		{"bool_true_number", "ROUNDDOWN(TRUE,0)", 1, 0},
		{"bool_false_number", "ROUNDDOWN(FALSE,0)", 0, 0},
		{"bool_true_digits", "ROUNDDOWN(3.14,TRUE)", 3.1, 1e-10},
		{"bool_false_digits", "ROUNDDOWN(3.14,FALSE)", 3, 0},

		// String coercion
		{"string_number", `ROUNDDOWN("3.7",0)`, 3, 0},
		{"string_digits", `ROUNDDOWN(3.14159,"2")`, 3.14, 1e-10},

		// Excel doc examples
		{"doc_ex1", "ROUNDDOWN(3.2,0)", 3, 0},
		{"doc_ex2", "ROUNDDOWN(76.9,0)", 76, 0},
		{"doc_ex3", "ROUNDDOWN(3.14159,3)", 3.141, 1e-10},
		{"doc_ex4", "ROUNDDOWN(-3.14159,1)", -3.1, 1e-10},
		{"doc_ex5", "ROUNDDOWN(31415.92654,-2)", 31400, 0},
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
		{"no_args", "ROUNDDOWN()", ErrValVALUE},
		{"one_arg", "ROUNDDOWN(1)", ErrValVALUE},
		{"too_many_args", "ROUNDDOWN(1,2,3)", ErrValVALUE},
		{"non_numeric_number", `ROUNDDOWN("abc",2)`, ErrValVALUE},
		{"non_numeric_digits", `ROUNDDOWN(1,"abc")`, ErrValVALUE},
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

func TestROUNDUP(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
		tol     float64
	}{
		// Basic positive number, various decimal places
		{"basic_2dp", "ROUNDUP(3.14159,2)", 3.15, 1e-10},
		{"basic_1dp", "ROUNDUP(2.11,1)", 2.2, 1e-10},
		{"basic_3dp", "ROUNDUP(1.23412,3)", 1.235, 1e-10},

		// Zero digits
		{"zero_digits_pos", "ROUNDUP(3.2,0)", 4, 0},
		{"zero_digits_pos2", "ROUNDUP(76.9,0)", 77, 0},

		// Negative numbers (away from zero)
		{"neg_0dp", "ROUNDUP(-3.2,0)", -4, 0},
		{"neg_1dp", "ROUNDUP(-3.14159,1)", -3.2, 1e-10},
		{"neg_2dp", "ROUNDUP(-2.781,2)", -2.79, 1e-10},

		// Negative num_digits (round left of decimal)
		{"neg_digits_tens", "ROUNDUP(123.456,-1)", 130, 0},
		{"neg_digits_hundreds", "ROUNDUP(31415.92654,-2)", 31500, 0},
		{"neg_digits_thousands", "ROUNDUP(9111,-3)", 10000, 0},

		// Zero input
		{"zero_input", "ROUNDUP(0,2)", 0, 0},
		{"zero_input_neg_digits", "ROUNDUP(0,-1)", 0, 0},

		// Already rounded (no change expected)
		{"already_rounded", "ROUNDUP(3.14,2)", 3.14, 1e-10},
		{"already_integer", "ROUNDUP(5,0)", 5, 0},
		{"already_integer_dp", "ROUNDUP(5,3)", 5, 0},

		// Large positive num_digits
		{"large_digits", "ROUNDUP(3.14159265,8)", 3.14159265, 1e-10},

		// Negative num_digits with large numbers
		{"neg_digits_large", "ROUNDUP(123456789,-4)", 123460000, 0},

		// Boolean coercion (TRUE=1, FALSE=0)
		{"bool_true_number", "ROUNDUP(TRUE,0)", 1, 0},
		{"bool_false_number", "ROUNDUP(FALSE,0)", 0, 0},
		{"bool_true_digits", "ROUNDUP(3.14,TRUE)", 3.2, 1e-10},
		{"bool_false_digits", "ROUNDUP(3.14,FALSE)", 4, 0},

		// String coercion
		{"string_number", `ROUNDUP("3.2",0)`, 4, 0},
		{"string_digits", `ROUNDUP(3.14159,"2")`, 3.15, 1e-10},

		// Excel doc examples
		{"doc_ex1", "ROUNDUP(3.2,0)", 4, 0},
		{"doc_ex2", "ROUNDUP(76.9,0)", 77, 0},
		{"doc_ex3", "ROUNDUP(3.14159,3)", 3.142, 1e-10},
		{"doc_ex4", "ROUNDUP(-3.14159,1)", -3.2, 1e-10},
		{"doc_ex5", "ROUNDUP(31415.92654,-2)", 31500, 0},
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
		{"no_args", "ROUNDUP()", ErrValVALUE},
		{"one_arg", "ROUNDUP(1)", ErrValVALUE},
		{"too_many_args", "ROUNDUP(1,2,3)", ErrValVALUE},
		{"non_numeric_number", `ROUNDUP("abc",2)`, ErrValVALUE},
		{"non_numeric_digits", `ROUNDUP(1,"abc")`, ErrValVALUE},
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

func TestACOT(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		{"zero", "ACOT(0)", math.Pi / 2},
		{"one", "ACOT(1)", math.Pi / 4},
		{"neg_one", "ACOT(-1)", 3 * math.Pi / 4},
		{"two", "ACOT(2)", math.Pi/2 - math.Atan(2)},
		{"neg_two", "ACOT(-2)", math.Pi/2 - math.Atan(-2)},
		{"half", "ACOT(0.5)", math.Pi/2 - math.Atan(0.5)},
		{"neg_half", "ACOT(-0.5)", math.Pi/2 - math.Atan(-0.5)},
		{"large", "ACOT(1000000)", math.Pi/2 - math.Atan(1000000)},
		{"large_neg", "ACOT(-1000000)", math.Pi/2 - math.Atan(-1000000)},
		{"ten", "ACOT(10)", math.Pi/2 - math.Atan(10)},
		{"neg_ten", "ACOT(-10)", math.Pi/2 - math.Atan(-10)},
		{"small_pos", "ACOT(0.001)", math.Pi/2 - math.Atan(0.001)},
		{"small_neg", "ACOT(-0.001)", math.Pi/2 - math.Atan(-0.001)},
		{"sqrt2", "ACOT(SQRT(2))", math.Pi/2 - math.Atan(math.Sqrt2)},
		{"pi", "ACOT(PI())", math.Pi/2 - math.Atan(math.Pi)},
		{"hundred", "ACOT(100)", math.Pi/2 - math.Atan(100)},
		{"third", "ACOT(1/3)", math.Pi/2 - math.Atan(1.0/3.0)},
		{"string_coerce", "ACOT(\"1\")", math.Pi / 4},
		{"bool_true", "ACOT(TRUE)", math.Pi / 4},
		{"bool_false", "ACOT(FALSE)", math.Pi / 2},
		{"negative_frac", "ACOT(-0.1)", math.Pi/2 - math.Atan(-0.1)},
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
		{"no_args", "ACOT()", ErrValVALUE},
		{"too_many_args", "ACOT(1,2)", ErrValVALUE},
		{"non_numeric", "ACOT(\"abc\")", ErrValVALUE},
		{"error_prop", "ACOT(1/0)", ErrValDIV0},
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

func TestACOTH(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		{"two", "ACOTH(2)", 0.5 * math.Log(3.0/1.0)},
		{"neg_two", "ACOTH(-2)", 0.5 * math.Log((-2.0+1.0)/(-2.0-1.0))},
		{"ten", "ACOTH(10)", 0.5 * math.Log(11.0/9.0)},
		{"neg_ten", "ACOTH(-10)", 0.5 * math.Log((-10.0+1.0)/(-10.0-1.0))},
		{"large", "ACOTH(1000)", 0.5 * math.Log(1001.0/999.0)},
		{"neg_large", "ACOTH(-1000)", 0.5 * math.Log((-1000.0+1.0)/(-1000.0-1.0))},
		{"one_point_one", "ACOTH(1.1)", 0.5 * math.Log(2.1/0.1)},
		{"neg_one_point_one", "ACOTH(-1.1)", 0.5 * math.Log((-1.1+1.0)/(-1.1-1.0))},
		{"five", "ACOTH(5)", 0.5 * math.Log(6.0/4.0)},
		{"neg_five", "ACOTH(-5)", 0.5 * math.Log((-5.0+1.0)/(-5.0-1.0))},
		{"three", "ACOTH(3)", 0.5 * math.Log(4.0/2.0)},
		{"hundred", "ACOTH(100)", 0.5 * math.Log(101.0/99.0)},
		{"one_point_five", "ACOTH(1.5)", 0.5 * math.Log(2.5/0.5)},
		{"two_point_five", "ACOTH(2.5)", 0.5 * math.Log(3.5/1.5)},
		{"string_coerce", "ACOTH(\"2\")", 0.5 * math.Log(3.0/1.0)},
		{"bool_symmetry", "ACOTH(2)+ACOTH(-2)", 0},
		{"large_val", "ACOTH(1000000)", 0.5 * math.Log(1000001.0/999999.0)},
		{"neg_three", "ACOTH(-3)", 0.5 * math.Log((-3.0+1.0)/(-3.0-1.0))},
		{"twenty", "ACOTH(20)", 0.5 * math.Log(21.0/19.0)},
		{"neg_twenty", "ACOTH(-20)", 0.5 * math.Log((-20.0+1.0)/(-20.0-1.0))},
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
		{"no_args", "ACOTH()", ErrValVALUE},
		{"too_many_args", "ACOTH(1,2)", ErrValVALUE},
		{"non_numeric", "ACOTH(\"abc\")", ErrValVALUE},
		{"one", "ACOTH(1)", ErrValNUM},
		{"neg_one", "ACOTH(-1)", ErrValNUM},
		{"zero", "ACOTH(0)", ErrValNUM},
		{"half", "ACOTH(0.5)", ErrValNUM},
		{"neg_half", "ACOTH(-0.5)", ErrValNUM},
		{"point_nine", "ACOTH(0.9)", ErrValNUM},
		{"neg_point_nine", "ACOTH(-0.9)", ErrValNUM},
		{"error_prop", "ACOTH(1/0)", ErrValDIV0},
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
// MUNIT tests
// ---------------------------------------------------------------------------

func TestMUNIT_1x1(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(1)")
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
	if got.Array[0][0].Num != 1 {
		t.Errorf("got %g, want 1", got.Array[0][0].Num)
	}
}

func TestMUNIT_2x2(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{1, 0}, {0, 1}}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			if got.Array[r][c].Num != want[r][c] {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestMUNIT_3x3(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 3x3, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	want := [][]float64{{1, 0, 0}, {0, 1, 0}, {0, 0, 1}}
	for r := 0; r < 3; r++ {
		for c := 0; c < 3; c++ {
			if got.Array[r][c].Num != want[r][c] {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestMUNIT_4x4(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(4)")
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
	for r := 0; r < 4; r++ {
		if len(got.Array[r]) != 4 {
			t.Fatalf("row %d: expected 4 cols, got %d", r, len(got.Array[r]))
		}
		for c := 0; c < 4; c++ {
			var expected float64
			if r == c {
				expected = 1
			}
			if got.Array[r][c].Num != expected {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, expected)
			}
		}
	}
}

func TestMUNIT_5x5_Diagonal(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	// Check all diagonal elements are 1.
	for i := 0; i < 5; i++ {
		if got.Array[i][i].Num != 1 {
			t.Errorf("diagonal [%d][%d]: got %g, want 1", i, i, got.Array[i][i].Num)
		}
	}
	// Check some off-diagonal elements are 0.
	if got.Array[0][1].Num != 0 {
		t.Errorf("[0][1]: got %g, want 0", got.Array[0][1].Num)
	}
	if got.Array[3][2].Num != 0 {
		t.Errorf("[3][2]: got %g, want 0", got.Array[3][2].Num)
	}
}

func TestMUNIT_Truncation_2_9(t *testing.T) {
	// MUNIT(2.9) should truncate to MUNIT(2).
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(2.9)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[0][0].Num != 1 || got.Array[1][1].Num != 1 {
		t.Error("diagonal values not 1")
	}
	if got.Array[0][1].Num != 0 || got.Array[1][0].Num != 0 {
		t.Error("off-diagonal values not 0")
	}
}

func TestMUNIT_Truncation_3_1(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(3.1)")
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
}

func TestMUNIT_Truncation_1_5(t *testing.T) {
	// MUNIT(1.5) should truncate to MUNIT(1).
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(1.5)")
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
	if got.Array[0][0].Num != 1 {
		t.Errorf("got %g, want 1", got.Array[0][0].Num)
	}
}

func TestMUNIT_StringCoercion(t *testing.T) {
	// MUNIT("3") should work the same as MUNIT(3).
	resolver := &mockResolver{}
	cf := evalCompile(t, `MUNIT("3")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 3x3, got %dx%d", len(got.Array), len(got.Array[0]))
	}
}

func TestMUNIT_Error_Zero(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestMUNIT_Error_Negative(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(-1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestMUNIT_Error_NegativeLarge(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(-5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestMUNIT_Error_String(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `MUNIT("abc")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("expected error, got type=%v", got.Type)
	}
}

func TestMUNIT_Error_TooManyArgs(t *testing.T) {
	got, err := fnMUNIT([]Value{NumberVal(3), NumberVal(2)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestMUNIT_Error_NoArgs(t *testing.T) {
	got, err := fnMUNIT([]Value{})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestMUNIT_Error_FractionTruncatesToZero(t *testing.T) {
	// MUNIT(0.9) truncates to 0 => #VALUE!
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(0.9)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestMUNIT_OffDiagonalAllZero(t *testing.T) {
	// Verify every off-diagonal element is 0 for a 4x4 matrix.
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(4)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	for r := 0; r < 4; r++ {
		for c := 0; c < 4; c++ {
			if r != c && got.Array[r][c].Num != 0 {
				t.Errorf("off-diagonal [%d][%d]: got %g, want 0", r, c, got.Array[r][c].Num)
			}
		}
	}
}

func TestMUNIT_AllDiagonalOne(t *testing.T) {
	// Verify every diagonal element is 1 for a 4x4 matrix.
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(4)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	for i := 0; i < 4; i++ {
		if got.Array[i][i].Num != 1 {
			t.Errorf("diagonal [%d][%d]: got %g, want 1", i, i, got.Array[i][i].Num)
		}
	}
}

func TestMUNIT_ValueTypes(t *testing.T) {
	// All values in the result should be numbers.
	resolver := &mockResolver{}
	cf := evalCompile(t, "MUNIT(3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	for r := 0; r < 3; r++ {
		for c := 0; c < 3; c++ {
			if got.Array[r][c].Type != ValueNumber {
				t.Errorf("[%d][%d]: expected ValueNumber, got %v", r, c, got.Array[r][c].Type)
			}
		}
	}
}

func TestMUNIT_BoolTrue(t *testing.T) {
	// TRUE coerces to 1, so MUNIT(TRUE) = MUNIT(1) = {{1}}.
	got, err := fnMUNIT([]Value{BoolVal(true)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 1x1, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[0][0].Num != 1 {
		t.Errorf("got %g, want 1", got.Array[0][0].Num)
	}
}

func TestMUNIT_BoolFalse(t *testing.T) {
	// FALSE coerces to 0 => #VALUE!
	got, err := fnMUNIT([]Value{BoolVal(false)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

// --------------- SUMX2MY2 tests ---------------

func TestSUMX2MY2_Basic(t *testing.T) {
	resolver := &mockResolver{}
	// {1,2,3} vs {4,5,6}: (1-16)+(4-25)+(9-36) = -15-21-27 = -63
	cf := evalCompile(t, "SUMX2MY2({1,2,3},{4,5,6})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -63 {
		t.Errorf("got %v (%g), want -63", got.Type, got.Num)
	}
}

func TestSUMX2MY2_SinglePair(t *testing.T) {
	resolver := &mockResolver{}
	// {3} vs {4}: 9-16 = -7
	cf := evalCompile(t, "SUMX2MY2({3},{4})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -7 {
		t.Errorf("got %v (%g), want -7", got.Type, got.Num)
	}
}

func TestSUMX2MY2_Zeros(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2MY2({0,0},{0,0})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("got %v (%g), want 0", got.Type, got.Num)
	}
}

func TestSUMX2MY2_SameValues(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2MY2({5,5},{5,5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("got %v (%g), want 0", got.Type, got.Num)
	}
}

func TestSUMX2MY2_Negative(t *testing.T) {
	resolver := &mockResolver{}
	// {-1,-2} vs {3,4}: (1-9)+(4-16) = -8-12 = -20
	cf := evalCompile(t, "SUMX2MY2({-1,-2},{3,4})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -20 {
		t.Errorf("got %v (%g), want -20", got.Type, got.Num)
	}
}

func TestSUMX2MY2_LargerArray(t *testing.T) {
	resolver := &mockResolver{}
	// {2,3,9,1,8,7,5} vs {6,5,11,7,5,4,4}
	// (4-36)+(9-25)+(81-121)+(1-49)+(64-25)+(49-16)+(25-16)
	// = -32 + -16 + -40 + -48 + 39 + 33 + 9 = -55
	cf := evalCompile(t, "SUMX2MY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -55 {
		t.Errorf("got %v (%g), want -55", got.Type, got.Num)
	}
}

func TestSUMX2MY2_DifferentLengths(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2MY2({1,2,3},{4,5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("expected #N/A, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2MY2_WrongArgCount(t *testing.T) {
	got, err := fnSumx2my2([]Value{NumberVal(1)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2MY2_ThreeArgs(t *testing.T) {
	got, err := fnSumx2my2([]Value{NumberVal(1), NumberVal(2), NumberVal(3)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2MY2_TextInArray(t *testing.T) {
	// Text that cannot be parsed as a number => #VALUE!
	got, err := fnSumx2my2([]Value{
		{Type: ValueArray, Array: [][]Value{{StringVal("abc"), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), NumberVal(4)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2MY2_ErrorPropagation(t *testing.T) {
	got, err := fnSumx2my2([]Value{
		{Type: ValueArray, Array: [][]Value{{ErrorVal(ErrValDIV0), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), NumberVal(4)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("expected #DIV/0!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2MY2_ErrorInSecondArray(t *testing.T) {
	got, err := fnSumx2my2([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), ErrorVal(ErrValNAME)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNAME {
		t.Errorf("expected #NAME?, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2MY2_EmptyCellsAsZero(t *testing.T) {
	// Empty cells should be treated as 0
	got, err := fnSumx2my2([]Value{
		{Type: ValueArray, Array: [][]Value{{Value{Type: ValueEmpty}, NumberVal(3)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(4), NumberVal(5)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (0-16)+(9-25) = -16-16 = -32
	if got.Type != ValueNumber || got.Num != -32 {
		t.Errorf("got %v (%g), want -32", got.Type, got.Num)
	}
}

func TestSUMX2MY2_BoolValues(t *testing.T) {
	// TRUE=1, FALSE=0
	got, err := fnSumx2my2([]Value{
		{Type: ValueArray, Array: [][]Value{{BoolVal(true), BoolVal(false)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(2), NumberVal(3)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (1-4)+(0-9) = -3-9 = -12
	if got.Type != ValueNumber || got.Num != -12 {
		t.Errorf("got %v (%g), want -12", got.Type, got.Num)
	}
}

func TestSUMX2MY2_NumericString(t *testing.T) {
	// Numeric string "5" should coerce to 5
	got, err := fnSumx2my2([]Value{
		{Type: ValueArray, Array: [][]Value{{StringVal("5")}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// 25-9 = 16
	if got.Type != ValueNumber || got.Num != 16 {
		t.Errorf("got %v (%g), want 16", got.Type, got.Num)
	}
}

func TestSUMX2MY2_LargeNumbers(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2MY2({100,200},{300,400})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (10000-90000)+(40000-160000) = -80000-120000 = -200000
	if got.Type != ValueNumber || got.Num != -200000 {
		t.Errorf("got %v (%g), want -200000", got.Type, got.Num)
	}
}

func TestSUMX2MY2_Decimals(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2MY2({1.5,2.5},{0.5,1.5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (2.25-0.25)+(6.25-2.25) = 2+4 = 6
	if got.Type != ValueNumber || got.Num != 6 {
		t.Errorf("got %v (%g), want 6", got.Type, got.Num)
	}
}

func TestSUMX2MY2_ScalarArgs(t *testing.T) {
	// Scalar values should work as single-element arrays
	got, err := fnSumx2my2([]Value{NumberVal(5), NumberVal(3)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// 25-9 = 16
	if got.Type != ValueNumber || got.Num != 16 {
		t.Errorf("got %v (%g), want 16", got.Type, got.Num)
	}
}

func TestSUMX2MY2_AllNegative(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2MY2({-3,-4},{-1,-2})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (9-1)+(16-4) = 8+12 = 20
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("got %v (%g), want 20", got.Type, got.Num)
	}
}

func TestSUMX2MY2_CellRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}
	cf := evalCompile(t, "SUMX2MY2(A1:A3,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (1-16)+(4-25)+(9-36) = -63
	if got.Type != ValueNumber || got.Num != -63 {
		t.Errorf("got %v (%g), want -63", got.Type, got.Num)
	}
}

// --------------- SUMX2PY2 tests ---------------

func TestSUMX2PY2_Basic(t *testing.T) {
	resolver := &mockResolver{}
	// {1,2,3} vs {4,5,6}: (1+16)+(4+25)+(9+36) = 17+29+45 = 91
	cf := evalCompile(t, "SUMX2PY2({1,2,3},{4,5,6})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 91 {
		t.Errorf("got %v (%g), want 91", got.Type, got.Num)
	}
}

func TestSUMX2PY2_SinglePair(t *testing.T) {
	resolver := &mockResolver{}
	// {3} vs {4}: 9+16 = 25
	cf := evalCompile(t, "SUMX2PY2({3},{4})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 25 {
		t.Errorf("got %v (%g), want 25", got.Type, got.Num)
	}
}

func TestSUMX2PY2_Zeros(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2PY2({0,0},{0,0})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("got %v (%g), want 0", got.Type, got.Num)
	}
}

func TestSUMX2PY2_SameValues(t *testing.T) {
	resolver := &mockResolver{}
	// {5,5} vs {5,5}: (25+25)+(25+25) = 100
	cf := evalCompile(t, "SUMX2PY2({5,5},{5,5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 100 {
		t.Errorf("got %v (%g), want 100", got.Type, got.Num)
	}
}

func TestSUMX2PY2_Negative(t *testing.T) {
	resolver := &mockResolver{}
	// {-1,-2} vs {3,4}: (1+9)+(4+16) = 10+20 = 30
	cf := evalCompile(t, "SUMX2PY2({-1,-2},{3,4})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("got %v (%g), want 30", got.Type, got.Num)
	}
}

func TestSUMX2PY2_LargerArray(t *testing.T) {
	resolver := &mockResolver{}
	// {2,3,9,1,8,7,5} vs {6,5,11,7,5,4,4}
	// (4+36)+(9+25)+(81+121)+(1+49)+(64+25)+(49+16)+(25+16)
	// = 40+34+202+50+89+65+41 = 521
	cf := evalCompile(t, "SUMX2PY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 521 {
		t.Errorf("got %v (%g), want 521", got.Type, got.Num)
	}
}

func TestSUMX2PY2_DifferentLengths(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2PY2({1,2,3},{4,5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("expected #N/A, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2PY2_WrongArgCount(t *testing.T) {
	got, err := fnSumx2py2([]Value{NumberVal(1)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2PY2_TextInArray(t *testing.T) {
	got, err := fnSumx2py2([]Value{
		{Type: ValueArray, Array: [][]Value{{StringVal("abc"), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), NumberVal(4)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2PY2_ErrorPropagation(t *testing.T) {
	got, err := fnSumx2py2([]Value{
		{Type: ValueArray, Array: [][]Value{{ErrorVal(ErrValDIV0), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), NumberVal(4)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("expected #DIV/0!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2PY2_EmptyCellsAsZero(t *testing.T) {
	got, err := fnSumx2py2([]Value{
		{Type: ValueArray, Array: [][]Value{{Value{Type: ValueEmpty}, NumberVal(3)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(4), NumberVal(5)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (0+16)+(9+25) = 16+34 = 50
	if got.Type != ValueNumber || got.Num != 50 {
		t.Errorf("got %v (%g), want 50", got.Type, got.Num)
	}
}

func TestSUMX2PY2_BoolValues(t *testing.T) {
	got, err := fnSumx2py2([]Value{
		{Type: ValueArray, Array: [][]Value{{BoolVal(true), BoolVal(false)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(2), NumberVal(3)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (1+4)+(0+9) = 5+9 = 14
	if got.Type != ValueNumber || got.Num != 14 {
		t.Errorf("got %v (%g), want 14", got.Type, got.Num)
	}
}

func TestSUMX2PY2_Decimals(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2PY2({1.5,2.5},{0.5,1.5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (2.25+0.25)+(6.25+2.25) = 2.5+8.5 = 11
	if got.Type != ValueNumber || got.Num != 11 {
		t.Errorf("got %v (%g), want 11", got.Type, got.Num)
	}
}

func TestSUMX2PY2_ScalarArgs(t *testing.T) {
	got, err := fnSumx2py2([]Value{NumberVal(3), NumberVal(4)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// 9+16 = 25
	if got.Type != ValueNumber || got.Num != 25 {
		t.Errorf("got %v (%g), want 25", got.Type, got.Num)
	}
}

func TestSUMX2PY2_AllNegative(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2PY2({-3,-4},{-1,-2})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (9+1)+(16+4) = 10+20 = 30
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("got %v (%g), want 30", got.Type, got.Num)
	}
}

func TestSUMX2PY2_CellRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}
	cf := evalCompile(t, "SUMX2PY2(A1:A3,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (1+16)+(4+25)+(9+36) = 91
	if got.Type != ValueNumber || got.Num != 91 {
		t.Errorf("got %v (%g), want 91", got.Type, got.Num)
	}
}

func TestSUMX2PY2_LargeNumbers(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2PY2({100,200},{300,400})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (10000+90000)+(40000+160000) = 100000+200000 = 300000
	if got.Type != ValueNumber || got.Num != 300000 {
		t.Errorf("got %v (%g), want 300000", got.Type, got.Num)
	}
}

func TestSUMX2PY2_NumericString(t *testing.T) {
	got, err := fnSumx2py2([]Value{
		{Type: ValueArray, Array: [][]Value{{StringVal("5")}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// 25+9 = 34
	if got.Type != ValueNumber || got.Num != 34 {
		t.Errorf("got %v (%g), want 34", got.Type, got.Num)
	}
}

func TestSUMX2PY2_ThreeArgs(t *testing.T) {
	got, err := fnSumx2py2([]Value{NumberVal(1), NumberVal(2), NumberVal(3)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2PY2_ErrorInSecondArray(t *testing.T) {
	got, err := fnSumx2py2([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1)}}},
		{Type: ValueArray, Array: [][]Value{{ErrorVal(ErrValREF)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("expected #REF!, got type=%v err=%v", got.Type, got.Err)
	}
}

// --------------- SUMXMY2 tests ---------------

func TestSUMXMY2_Basic(t *testing.T) {
	resolver := &mockResolver{}
	// {1,2,3} vs {4,5,6}: (1-4)²+(2-5)²+(3-6)² = 9+9+9 = 27
	cf := evalCompile(t, "SUMXMY2({1,2,3},{4,5,6})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 27 {
		t.Errorf("got %v (%g), want 27", got.Type, got.Num)
	}
}

func TestSUMXMY2_SinglePair(t *testing.T) {
	resolver := &mockResolver{}
	// {3} vs {4}: (3-4)² = 1
	cf := evalCompile(t, "SUMXMY2({3},{4})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("got %v (%g), want 1", got.Type, got.Num)
	}
}

func TestSUMXMY2_Zeros(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMXMY2({0,0},{0,0})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("got %v (%g), want 0", got.Type, got.Num)
	}
}

func TestSUMXMY2_SameValues(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMXMY2({5,5,5},{5,5,5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("got %v (%g), want 0", got.Type, got.Num)
	}
}

func TestSUMXMY2_Negative(t *testing.T) {
	resolver := &mockResolver{}
	// {-1,-2} vs {3,4}: (-1-3)²+(-2-4)² = 16+36 = 52
	cf := evalCompile(t, "SUMXMY2({-1,-2},{3,4})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 52 {
		t.Errorf("got %v (%g), want 52", got.Type, got.Num)
	}
}

func TestSUMXMY2_LargerArray(t *testing.T) {
	resolver := &mockResolver{}
	// {2,3,9,1,8,7,5} vs {6,5,11,7,5,4,4}
	// (2-6)²+(3-5)²+(9-11)²+(1-7)²+(8-5)²+(7-4)²+(5-4)²
	// = 16+4+4+36+9+9+1 = 79
	cf := evalCompile(t, "SUMXMY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 79 {
		t.Errorf("got %v (%g), want 79", got.Type, got.Num)
	}
}

func TestSUMXMY2_DifferentLengths(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMXMY2({1,2,3},{4,5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("expected #N/A, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMXMY2_WrongArgCount(t *testing.T) {
	got, err := fnSumxmy2([]Value{NumberVal(1)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMXMY2_ThreeArgs(t *testing.T) {
	got, err := fnSumxmy2([]Value{NumberVal(1), NumberVal(2), NumberVal(3)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMXMY2_TextInArray(t *testing.T) {
	got, err := fnSumxmy2([]Value{
		{Type: ValueArray, Array: [][]Value{{StringVal("abc"), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), NumberVal(4)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMXMY2_ErrorPropagation(t *testing.T) {
	got, err := fnSumxmy2([]Value{
		{Type: ValueArray, Array: [][]Value{{ErrorVal(ErrValDIV0), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), NumberVal(4)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("expected #DIV/0!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMXMY2_ErrorInSecondArray(t *testing.T) {
	got, err := fnSumxmy2([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), ErrorVal(ErrValNUM)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("expected #NUM!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMXMY2_EmptyCellsAsZero(t *testing.T) {
	got, err := fnSumxmy2([]Value{
		{Type: ValueArray, Array: [][]Value{{Value{Type: ValueEmpty}, NumberVal(3)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(4), NumberVal(5)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (0-4)²+(3-5)² = 16+4 = 20
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("got %v (%g), want 20", got.Type, got.Num)
	}
}

func TestSUMXMY2_BoolValues(t *testing.T) {
	got, err := fnSumxmy2([]Value{
		{Type: ValueArray, Array: [][]Value{{BoolVal(true), BoolVal(false)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(2), NumberVal(3)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (1-2)²+(0-3)² = 1+9 = 10
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("got %v (%g), want 10", got.Type, got.Num)
	}
}

func TestSUMXMY2_Decimals(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMXMY2({1.5,2.5},{0.5,1.5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (1)²+(1)² = 1+1 = 2
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("got %v (%g), want 2", got.Type, got.Num)
	}
}

func TestSUMXMY2_ScalarArgs(t *testing.T) {
	got, err := fnSumxmy2([]Value{NumberVal(7), NumberVal(3)})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (7-3)² = 16
	if got.Type != ValueNumber || got.Num != 16 {
		t.Errorf("got %v (%g), want 16", got.Type, got.Num)
	}
}

func TestSUMXMY2_LargeNumbers(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMXMY2({100,200},{300,400})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (100-300)²+(200-400)² = 40000+40000 = 80000
	if got.Type != ValueNumber || got.Num != 80000 {
		t.Errorf("got %v (%g), want 80000", got.Type, got.Num)
	}
}

func TestSUMXMY2_NumericString(t *testing.T) {
	got, err := fnSumxmy2([]Value{
		{Type: ValueArray, Array: [][]Value{{StringVal("5")}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (5-3)² = 4
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("got %v (%g), want 4", got.Type, got.Num)
	}
}

func TestSUMXMY2_CellRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}
	cf := evalCompile(t, "SUMXMY2(A1:A3,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (1-4)²+(2-5)²+(3-6)² = 9+9+9 = 27
	if got.Type != ValueNumber || got.Num != 27 {
		t.Errorf("got %v (%g), want 27", got.Type, got.Num)
	}
}

func TestSUMXMY2_AllNegative(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMXMY2({-3,-4},{-1,-2})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (-3-(-1))²+(-4-(-2))² = 4+4 = 8
	if got.Type != ValueNumber || got.Num != 8 {
		t.Errorf("got %v (%g), want 8", got.Type, got.Num)
	}
}

func TestINT(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Basic positive decimal — rounds down
		{"pos_decimal", "INT(3.7)", 3},
		// Positive integer — unchanged
		{"pos_integer", "INT(5)", 5},
		// Zero
		{"zero", "INT(0)", 0},
		// Negative decimal — rounds away from zero (toward -infinity)
		{"neg_decimal", "INT(-3.2)", -4},
		// Negative integer — unchanged
		{"neg_integer", "INT(-7)", -7},
		// Very small positive
		{"small_pos", "INT(0.1)", 0},
		// Very small negative — rounds to -1
		{"small_neg", "INT(-0.1)", -1},
		// Negative fraction -0.5
		{"neg_half", "INT(-0.5)", -1},
		// Large number
		{"large_number", "INT(1000000.9)", 1000000},
		// String number coercion
		{"string_num", `INT("5.5")`, 5},
		// Boolean TRUE → 1
		{"bool_true", "INT(TRUE)", 1},
		// Boolean FALSE → 0
		{"bool_false", "INT(FALSE)", 0},
		// Already integer positive
		{"already_int_pos", "INT(10)", 10},
		// Already integer negative
		{"already_int_neg", "INT(-10)", -10},
		// Excel doc example: INT(8.9) = 8
		{"excel_doc_pos", "INT(8.9)", 8},
		// Excel doc example: INT(-8.9) = -9
		{"excel_doc_neg", "INT(-8.9)", -9},
		// Positive number just below next integer
		{"just_below", "INT(2.999999)", 2},
		// Negative number just above integer
		{"neg_just_above", "INT(-2.000001)", -3},
	}

	for _, tt := range numTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): unexpected error: %v", tt.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("Eval(%q): got type %v, want number", tt.formula, got.Type)
			}
			if got.Num != tt.want {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		errVal  ErrorValue
	}{
		// Non-numeric string
		{"non_numeric_string", `INT("hello")`, ErrValVALUE},
		// No arguments
		{"no_args", "INT()", ErrValVALUE},
		// Too many arguments
		{"too_many_args", "INT(1,2)", ErrValVALUE},
		// Error propagation
		{"error_propagation", "INT(1/0)", ErrValDIV0},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): unexpected error: %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.errVal {
				t.Errorf("Eval(%q) = type=%v err=%v, want %v", tt.formula, got.Type, got.Err, tt.errVal)
			}
		})
	}
}

func TestSQRT(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Perfect squares
		{"sqrt_0", "SQRT(0)", 0},
		{"sqrt_1", "SQRT(1)", 1},
		{"sqrt_4", "SQRT(4)", 2},
		{"sqrt_9", "SQRT(9)", 3},
		{"sqrt_16", "SQRT(16)", 4},
		{"sqrt_25", "SQRT(25)", 5},
		{"sqrt_100", "SQRT(100)", 10},
		// Non-perfect squares
		{"sqrt_2", "SQRT(2)", math.Sqrt2},
		{"sqrt_3", "SQRT(3)", math.Sqrt(3)},
		{"sqrt_10", "SQRT(10)", math.Sqrt(10)},
		// Large number
		{"sqrt_large", "SQRT(1000000)", 1000},
		{"sqrt_large_non_perfect", "SQRT(999999)", math.Sqrt(999999)},
		// Small decimals
		{"sqrt_0.25", "SQRT(0.25)", 0.5},
		{"sqrt_0.01", "SQRT(0.01)", 0.1},
		{"sqrt_0.0001", "SQRT(0.0001)", 0.01},
		// String coercion
		{"string_coerce", "SQRT(\"9\")", 3},
		{"string_coerce_decimal", "SQRT(\"0.25\")", 0.5},
		// Boolean coercion
		{"bool_true", "SQRT(TRUE)", 1},
		{"bool_false", "SQRT(FALSE)", 0},
		// Excel doc example: SQRT(ABS(-16)) = 4
		{"excel_doc_abs_neg16", "SQRT(ABS(-16))", 4},
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
		{"negative", "SQRT(-1)", ErrValNUM},
		{"negative_large", "SQRT(-100)", ErrValNUM},
		{"no_args", "SQRT()", ErrValVALUE},
		{"too_many_args", "SQRT(4,2)", ErrValVALUE},
		{"non_numeric", "SQRT(\"abc\")", ErrValVALUE},
		{"error_propagation", "SQRT(1/0)", ErrValDIV0},
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

func TestLOG(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Basic default base 10
		{"log10_of_10", "LOG(10)", 1, 0},
		{"log10_of_100", "LOG(100)", 2, 0},
		{"log10_of_1000", "LOG(1000)", 3, 1e-10},

		// Explicit base 10
		{"log_10_explicit", "LOG(10,10)", 1, 0},
		{"log_100_explicit", "LOG(100,10)", 2, 0},

		// Base 2
		{"log2_of_8", "LOG(8,2)", 3, 0},
		{"log2_of_16", "LOG(16,2)", 4, 0},
		{"log2_of_1", "LOG(1,2)", 0, 0},

		// Base e (natural log)
		{"log_e_of_e", "LOG(EXP(1),EXP(1))", 1, 1e-10},

		// LOG(1, any base) = 0
		{"log_1_base10", "LOG(1)", 0, 0},
		{"log_1_base7", "LOG(1,7)", 0, 0},
		{"log_1_base100", "LOG(1,100)", 0, 0},

		// LOG(base, base) = 1
		{"log_base_base_5", "LOG(5,5)", 1, 1e-10},
		{"log_base_base_3", "LOG(3,3)", 1, 1e-10},

		// Fractional base
		{"log_frac_base", "LOG(0.125,0.5)", 3, 1e-10},

		// Large numbers
		{"log_large", "LOG(1000000)", 6, 1e-10},
		{"log_large_base2", "LOG(1048576,2)", 20, 1e-10},

		// Excel doc examples
		{"excel_example_1", "LOG(10)", 1, 0},
		{"excel_example_2", "LOG(8,2)", 3, 0},
		{"excel_example_3", "LOG(86,2.7182818)", 4.4543473, 1e-4},

		// Boolean coercion: TRUE=1
		{"log_true", "LOG(TRUE)", 0, 0},

		// String coercion
		{"log_string_10", "LOG(\"10\")", 1, 0},
		{"log_string_base", "LOG(100,\"10\")", 2, 0},
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

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// Negative number
		{"negative_num", "LOG(-1)", ErrValNUM},
		{"negative_num2", "LOG(-10)", ErrValNUM},

		// Zero
		{"zero_num", "LOG(0)", ErrValNUM},

		// Base = 1 → #DIV/0!
		{"base_1", "LOG(10,1)", ErrValDIV0},
		{"base_1_any", "LOG(1,1)", ErrValDIV0},

		// Base = 0 → #NUM!
		{"base_0", "LOG(10,0)", ErrValNUM},

		// Negative base → #NUM!
		{"negative_base", "LOG(10,-2)", ErrValNUM},

		// No args
		{"no_args", "LOG()", ErrValVALUE},

		// Too many args
		{"too_many_args", "LOG(10,2,3)", ErrValVALUE},

		// Error propagation
		{"error_prop_num", "LOG(1/0)", ErrValDIV0},
		{"error_prop_base", "LOG(10,1/0)", ErrValDIV0},
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

func TestGAMMA(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Positive integers: GAMMA(n) = (n-1)!
		{"int_1", "GAMMA(1)", 1, 0},
		{"int_2", "GAMMA(2)", 1, 0},
		{"int_3", "GAMMA(3)", 2, 0},
		{"int_4", "GAMMA(4)", 6, 0},
		{"int_5", "GAMMA(5)", 24, 0},
		{"int_10", "GAMMA(10)", 362880, 0},

		// Large integer
		{"int_20", "GAMMA(20)", 121645100408832000, 1e3},

		// Half-integer values
		{"half_0.5", "GAMMA(0.5)", 1.7724538509055159, 1e-9},
		{"half_1.5", "GAMMA(1.5)", 0.886226925452758, 1e-9},
		{"half_2.5", "GAMMA(2.5)", 1.3293403881791370, 1e-9},

		// Fractional values
		{"frac_0.1", "GAMMA(0.1)", 9.513507698668732, 1e-6},
		{"frac_0.001", "GAMMA(0.001)", 999.4237724845955, 1e-3},

		// Negative non-integer values
		{"neg_0.5", "GAMMA(-0.5)", -3.544907701811032, 1e-9},
		{"neg_1.5", "GAMMA(-1.5)", 2.363271801207354, 1e-9},
		{"neg_3.75", "GAMMA(-3.75)", 0.267866128861417, 1e-9},

		// String coercion
		{"string_num", "GAMMA(\"2.5\")", 1.3293403881791370, 1e-9},

		// Boolean coercion
		{"bool_true", "GAMMA(TRUE)", 1, 0},
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

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// Zero → #NUM!
		{"zero", "GAMMA(0)", ErrValNUM},

		// Negative integers → #NUM!
		{"neg_int_1", "GAMMA(-1)", ErrValNUM},
		{"neg_int_2", "GAMMA(-2)", ErrValNUM},
		{"neg_int_100", "GAMMA(-100)", ErrValNUM},

		// Overflow → #NUM!
		{"overflow", "GAMMA(200)", ErrValNUM},

		// Non-numeric string → #VALUE!
		{"non_numeric", "GAMMA(\"abc\")", ErrValVALUE},

		// Wrong arg count → #VALUE!
		{"no_args", "GAMMA()", ErrValVALUE},
		{"too_many_args", "GAMMA(1,2)", ErrValVALUE},

		// Error propagation
		{"err_div0", "GAMMA(1/0)", ErrValDIV0},
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
