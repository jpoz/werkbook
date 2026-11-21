package formula

import (
	"fmt"
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

		// Doc examples
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

	t.Run("trimmed_range_origin", func(t *testing.T) {
		got, err := fnABS([]Value{
			trimmedRangeValue([][]Value{
				{NumberVal(-5)},
			}, 1, 1, 1, 3),
		})
		if err != nil {
			t.Fatalf("fnABS: %v", err)
		}
		assertLookupValueEqual(t, got, Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(5)},
			{NumberVal(0)},
			{NumberVal(0)},
		}})
		if got.RangeOrigin == nil || got.RangeOrigin.FromRow != 1 || got.RangeOrigin.ToRow != 3 {
			t.Fatalf("fnABS RangeOrigin = %+v, want rows 1:3", got.RangeOrigin)
		}
	})

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

		// Doc examples
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

		// Doc examples
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

		// Doc examples
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

		// Doc examples
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

		// Doc examples
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
		// Original tests
		{"positive_mod", "MOD(10,3)", 1},
		{"exact_divisor", "MOD(10,5)", 0},
		{"negative_dividend", "MOD(-10,3)", 2},
		{"negative_divisor", "MOD(10,-3)", -2},
		{"both_negative", "MOD(-10,-3)", -1},
		{"fractional", "MOD(7.5,2)", 1.5},
		{"small_divisor_ok", "MOD(10,0.1)", 0},

		// Basic cases
		{"basic_10_3", "MOD(10,3)", 1},
		{"basic_3_2", "MOD(3,2)", 1},
		{"basic_17_5", "MOD(17,5)", 2},
		{"basic_100_7", "MOD(100,7)", 2},

		// Zero dividend
		{"zero_dividend", "MOD(0,5)", 0},
		{"zero_dividend_neg_divisor", "MOD(0,-3)", 0},

		// Number mod itself = 0
		{"self_mod", "MOD(5,5)", 0},
		{"self_mod_large", "MOD(123,123)", 0},
		{"self_mod_neg", "MOD(-7,-7)", 0},

		// Mod by 1 = 0 for integers
		{"mod_by_1", "MOD(5,1)", 0},
		{"mod_by_1_large", "MOD(999,1)", 0},
		{"mod_by_1_neg_num", "MOD(-5,1)", 0},

		// Sign follows divisor (Excel semantics)
		{"sign_follows_divisor_pos", "MOD(3,2)", 1},
		{"sign_follows_divisor_neg_d", "MOD(3,-2)", -1},
		{"sign_follows_divisor_neg_n", "MOD(-3,2)", 1},
		{"sign_follows_divisor_both_neg", "MOD(-3,-2)", -1},

		// Decimal arguments
		{"decimal_5_5_div_2", "MOD(5.5,2)", 1.5},
		{"decimal_2_5_div_1_5", "MOD(2.5,1.5)", 1},
		{"decimal_small", "MOD(0.7,0.3)", 0.1},
		{"decimal_neg_dividend", "MOD(-5.5,2)", 0.5},
		{"decimal_neg_divisor", "MOD(5.5,-2)", -0.5},

		// Small divisor (near zero but not zero)
		{"small_divisor_0_01", "MOD(1,0.01)", 0},
		{"small_divisor_0_001", "MOD(1,0.001)", 0},

		// Larger numbers within precision threshold
		{"large_nums_ok", "MOD(10000,3)", 1},
		{"large_nums_ok2", "MOD(999999,7)", 999999 - 142857*7},

		// String coercion
		{"string_num", `MOD("10","3")`, 1},
		{"string_dividend", `MOD("10",3)`, 1},
		{"string_divisor", `MOD(10,"3")`, 1},

		// Boolean coercion
		{"bool_true_dividend", "MOD(TRUE,2)", 1},
		{"bool_false_dividend", "MOD(FALSE,5)", 0},
		{"bool_true_divisor", "MOD(10,TRUE)", 0},
		{"bool_true_both", "MOD(TRUE,TRUE)", 0},
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
		{"div_by_zero_neg_num", "MOD(-5,0)", ErrValDIV0},
		{"div_zero_by_zero", "MOD(0,0)", ErrValDIV0},

		// Precision overflow: |n/d| >= 1e13 -> #NUM!
		{"precision_overflow_tiny_neg_d", "MOD(100,-1e-15)", ErrValNUM},
		{"precision_overflow_large_n", "MOD(1e18,1)", ErrValNUM},
		{"precision_overflow_power", "MOD(10^15,7)", ErrValNUM},

		// Wrong argument count
		{"no_args", "MOD()", ErrValVALUE},
		{"one_arg", "MOD(5)", ErrValVALUE},
		{"three_args", "MOD(5,3,1)", ErrValVALUE},

		// Non-numeric strings -> #VALUE!
		{"non_numeric_num", `MOD("abc",2)`, ErrValVALUE},
		{"non_numeric_den", `MOD(2,"abc")`, ErrValVALUE},
		{"non_numeric_both", `MOD("abc","def")`, ErrValVALUE},

		// Error propagation
		{"err_div0_num", "MOD(1/0,2)", ErrValDIV0},
		{"err_div0_den", "MOD(2,1/0)", ErrValDIV0},
		{"err_na_num", "MOD(NA(),2)", ErrValNA},

		// Boolean FALSE as divisor (coerces to 0 -> #DIV/0!)
		{"bool_false_divisor", "MOD(5,FALSE)", ErrValDIV0},
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

func TestMROUND(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
		tol     float64
	}{
		// Documentation examples
		{"doc_ex1", "MROUND(10,3)", 9, 0},
		{"doc_ex2", "MROUND(-10,-3)", -9, 0},
		{"doc_ex3", "MROUND(1.3,0.2)", 1.4, 1e-10},

		// Multiple of zero returns 0
		{"multiple_zero", "MROUND(5,0)", 0, 0},
		{"multiple_zero_neg", "MROUND(-5,0)", 0, 0},
		{"zero_multiple_zero", "MROUND(0,0)", 0, 0},

		// Number is zero
		{"zero_number", "MROUND(0,3)", 0, 0},
		{"zero_number_neg", "MROUND(0,-3)", 0, 0},

		// Positive rounding - rounds to nearest multiple
		{"pos_round_down", "MROUND(7,5)", 5, 0},
		{"pos_round_up", "MROUND(8,5)", 10, 0},
		{"pos_exact", "MROUND(10,5)", 10, 0},
		{"pos_round_nearest", "MROUND(13,5)", 15, 0},
		{"pos_round_mid_away", "MROUND(7.5,5)", 10, 0}, // midpoint rounds away from zero

		// Negative rounding - both signs negative
		{"neg_round_down", "MROUND(-7,-5)", -5, 0},
		{"neg_round_up", "MROUND(-8,-5)", -10, 0},
		{"neg_exact", "MROUND(-10,-5)", -10, 0},
		{"neg_round_nearest", "MROUND(-13,-5)", -15, 0},
		{"neg_round_mid_away", "MROUND(-7.5,-5)", -10, 0}, // midpoint rounds away from zero

		// Decimal multiples
		{"decimal_mult_1", "MROUND(1.05,0.1)", 1.1, 1e-10}, // midpoint with decimal multiple: direction undefined per docs
		{"decimal_mult_2", "MROUND(1.15,0.1)", 1.1, 1e-10}, // midpoint with decimal multiple: direction undefined per docs
		{"decimal_mult_3", "MROUND(0.5,0.25)", 0.5, 1e-10},
		{"decimal_mult_4", "MROUND(0.6,0.25)", 0.5, 1e-10},
		{"decimal_mult_5", "MROUND(0.63,0.25)", 0.75, 1e-10},

		// String coercion
		{"string_number", "MROUND(\"10\",3)", 9, 0},
		{"string_multiple", "MROUND(10,\"3\")", 9, 0},
		{"string_both", "MROUND(\"10\",\"3\")", 9, 0},

		// Boolean coercion
		{"bool_true_number", "MROUND(TRUE,1)", 1, 0},
		{"bool_false_number", "MROUND(FALSE,1)", 0, 0},
		{"bool_true_multiple", "MROUND(5,TRUE)", 5, 0},
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
		// Signs must match: positive number with negative multiple
		{"pos_neg_mismatch", "MROUND(5,-2)", ErrValNUM},
		// Signs must match: negative number with positive multiple
		{"neg_pos_mismatch", "MROUND(-5,2)", ErrValNUM},

		// Wrong argument count
		{"no_args", "MROUND()", ErrValVALUE},
		{"one_arg", "MROUND(10)", ErrValVALUE},
		{"three_args", "MROUND(10,3,1)", ErrValVALUE},

		// Non-numeric string
		{"non_numeric_number", "MROUND(\"abc\",3)", ErrValVALUE},
		{"non_numeric_multiple", "MROUND(10,\"abc\")", ErrValVALUE},

		// Error propagation
		{"err_div0_number", "MROUND(1/0,3)", ErrValDIV0},
		{"err_div0_multiple", "MROUND(10,1/0)", ErrValDIV0},
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

		// Rounding 0.5 (away from zero)
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

		// Doc examples
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

		// Doc examples
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
// RAND comprehensive tests
// ---------------------------------------------------------------------------

func TestRAND(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns_number_type", func(t *testing.T) {
		cf := evalCompile(t, "RAND()")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Errorf("RAND() type = %v, want ValueNumber", got.Type)
		}
	})

	t.Run("no_args_works", func(t *testing.T) {
		cf := evalCompile(t, "RAND()")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Errorf("RAND() should succeed with no args, got type %v", got.Type)
		}
	})

	t.Run("wrong_arg_count_one_number", func(t *testing.T) {
		cf := evalCompile(t, "RAND(1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("RAND(1) = %v, want #VALUE! error", got)
		}
	})

	t.Run("wrong_arg_count_string", func(t *testing.T) {
		cf := evalCompile(t, `RAND("x")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`RAND("x") = %v, want #VALUE! error`, got)
		}
	})

	t.Run("wrong_arg_count_two_args", func(t *testing.T) {
		cf := evalCompile(t, "RAND(1,2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("RAND(1,2) = %v, want #VALUE! error", got)
		}
	})

	t.Run("lower_bound_inclusive", func(t *testing.T) {
		// RAND() >= 0 must always be true
		for i := 0; i < 100; i++ {
			cf := evalCompile(t, "RAND()>=0")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueBool || !got.Bool {
				t.Fatalf("RAND()>=0 = %v, want TRUE (iteration %d)", got, i)
			}
		}
	})

	t.Run("upper_bound_exclusive", func(t *testing.T) {
		// RAND() < 1 must always be true
		for i := 0; i < 100; i++ {
			cf := evalCompile(t, "RAND()<1")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueBool || !got.Bool {
				t.Fatalf("RAND()<1 = %v, want TRUE (iteration %d)", got, i)
			}
		}
	})

	t.Run("not_negative", func(t *testing.T) {
		for i := 0; i < 100; i++ {
			cf := evalCompile(t, "RAND()")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Num < 0 {
				t.Fatalf("RAND() = %g, want non-negative (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("TYPE_equals_1", func(t *testing.T) {
		cf := evalCompile(t, "TYPE(RAND())")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("TYPE(RAND()) = %v, want 1 (number)", got)
		}
	})

	t.Run("ISNUMBER_true", func(t *testing.T) {
		cf := evalCompile(t, "ISNUMBER(RAND())")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(RAND()) = %v, want TRUE", got)
		}
	})

	t.Run("ISTEXT_false", func(t *testing.T) {
		cf := evalCompile(t, "ISTEXT(RAND())")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISTEXT(RAND()) = %v, want FALSE", got)
		}
	})

	t.Run("arithmetic_plus_one", func(t *testing.T) {
		// RAND()+1 should be between 1 (inclusive) and 2 (exclusive)
		for i := 0; i < 50; i++ {
			cf := evalCompile(t, "RAND()+1")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < 1 || got.Num >= 2 {
				t.Fatalf("RAND()+1 = %g, want [1,2) (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("scaling_times_10", func(t *testing.T) {
		// RAND()*10 should be in [0, 10)
		for i := 0; i < 50; i++ {
			cf := evalCompile(t, "RAND()*10")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < 0 || got.Num >= 10 {
				t.Fatalf("RAND()*10 = %g, want [0,10) (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("INT_of_RAND_is_zero", func(t *testing.T) {
		// Since 0 <= RAND() < 1, INT(RAND()) = 0 always
		for i := 0; i < 50; i++ {
			cf := evalCompile(t, "INT(RAND())")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != 0 {
				t.Fatalf("INT(RAND()) = %g, want 0 (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("in_IF_condition", func(t *testing.T) {
		// RAND()+1 > 0 is always true, so IF should return "yes"
		cf := evalCompile(t, `IF(RAND()+1>0,"yes","no")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IF(RAND()+1>0,"yes","no") = %v, want "yes"`, got)
		}
	})

	t.Run("comparison_chain_AND", func(t *testing.T) {
		// AND(RAND()>=0, RAND()<1) should always be TRUE
		for i := 0; i < 50; i++ {
			cf := evalCompile(t, "AND(RAND()>=0, RAND()<1)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueBool || !got.Bool {
				t.Fatalf("AND(RAND()>=0, RAND()<1) = %v, want TRUE (iteration %d)", got, i)
			}
		}
	})

	t.Run("dice_roll_simulation", func(t *testing.T) {
		// INT(RAND()*6)+1 should be in [1, 6]
		for i := 0; i < 100; i++ {
			cf := evalCompile(t, "INT(RAND()*6)+1")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < 1 || got.Num > 6 {
				t.Fatalf("INT(RAND()*6)+1 = %g, want [1,6] (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("ROUND_to_zero_decimals", func(t *testing.T) {
		// ROUND(RAND(),0) should be either 0 or 1
		for i := 0; i < 50; i++ {
			cf := evalCompile(t, "ROUND(RAND(),0)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || (got.Num != 0 && got.Num != 1) {
				t.Fatalf("ROUND(RAND(),0) = %g, want 0 or 1 (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("two_calls_independent", func(t *testing.T) {
		// Two separate RAND() evaluations should both be valid numbers in [0,1)
		cf1 := evalCompile(t, "RAND()")
		got1, err := Eval(cf1, resolver, nil)
		if err != nil {
			t.Fatalf("Eval RAND() #1: %v", err)
		}
		cf2 := evalCompile(t, "RAND()")
		got2, err := Eval(cf2, resolver, nil)
		if err != nil {
			t.Fatalf("Eval RAND() #2: %v", err)
		}
		if got1.Type != ValueNumber || got1.Num < 0 || got1.Num >= 1 {
			t.Errorf("RAND() #1 = %g, want [0,1)", got1.Num)
		}
		if got2.Type != ValueNumber || got2.Num < 0 || got2.Num >= 1 {
			t.Errorf("RAND() #2 = %g, want [0,1)", got2.Num)
		}
	})

	t.Run("SUM_of_two_RANDs", func(t *testing.T) {
		// SUM(RAND(), RAND()) should be in [0, 2)
		for i := 0; i < 50; i++ {
			cf := evalCompile(t, "SUM(RAND(), RAND())")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < 0 || got.Num >= 2 {
				t.Fatalf("SUM(RAND(), RAND()) = %g, want [0,2) (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("AVERAGE_of_RAND", func(t *testing.T) {
		// AVERAGE(RAND()) should be in [0, 1)
		for i := 0; i < 50; i++ {
			cf := evalCompile(t, "AVERAGE(RAND())")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < 0 || got.Num >= 1 {
				t.Fatalf("AVERAGE(RAND()) = %g, want [0,1) (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("RAND_times_100_int", func(t *testing.T) {
		// INT(RAND()*100) should be in [0, 99] (Excel doc example)
		for i := 0; i < 50; i++ {
			cf := evalCompile(t, "INT(RAND()*100)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < 0 || got.Num > 99 {
				t.Fatalf("INT(RAND()*100) = %g, want [0,99] (iteration %d)", got.Num, i)
			}
			// Should be an integer
			if got.Num != math.Floor(got.Num) {
				t.Fatalf("INT(RAND()*100) = %g, want integer (iteration %d)", got.Num, i)
			}
		}
	})
}

// ---------------------------------------------------------------------------
// RANDBETWEEN comprehensive tests
// ---------------------------------------------------------------------------

func TestRANDBETWEEN(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("basic_returns_number", func(t *testing.T) {
		cf := evalCompile(t, "RANDBETWEEN(1,10)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Errorf("RANDBETWEEN(1,10) type = %v, want ValueNumber", got.Type)
		}
	})

	t.Run("lower_bound_inclusive", func(t *testing.T) {
		for i := 0; i < 100; i++ {
			cf := evalCompile(t, "RANDBETWEEN(1,10)>=1")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueBool || !got.Bool {
				t.Fatalf("RANDBETWEEN(1,10)>=1 = %v, want TRUE (iteration %d)", got, i)
			}
		}
	})

	t.Run("upper_bound_inclusive", func(t *testing.T) {
		for i := 0; i < 100; i++ {
			cf := evalCompile(t, "RANDBETWEEN(1,10)<=10")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueBool || !got.Bool {
				t.Fatalf("RANDBETWEEN(1,10)<=10 = %v, want TRUE (iteration %d)", got, i)
			}
		}
	})

	t.Run("same_bounds_returns_that_value", func(t *testing.T) {
		for i := 0; i < 20; i++ {
			cf := evalCompile(t, "RANDBETWEEN(1,1)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != 1 {
				t.Fatalf("RANDBETWEEN(1,1) = %g, want 1 (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("same_bounds_zero", func(t *testing.T) {
		for i := 0; i < 20; i++ {
			cf := evalCompile(t, "RANDBETWEEN(0,0)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != 0 {
				t.Fatalf("RANDBETWEEN(0,0) = %g, want 0 (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("negative_range", func(t *testing.T) {
		for i := 0; i < 100; i++ {
			cf := evalCompile(t, "RANDBETWEEN(-10,-1)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < -10 || got.Num > -1 {
				t.Fatalf("RANDBETWEEN(-10,-1) = %g, want [-10,-1] (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("negative_to_positive_range", func(t *testing.T) {
		for i := 0; i < 100; i++ {
			cf := evalCompile(t, "RANDBETWEEN(-5,5)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < -5 || got.Num > 5 {
				t.Fatalf("RANDBETWEEN(-5,5) = %g, want [-5,5] (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("binary_range_0_or_1", func(t *testing.T) {
		for i := 0; i < 100; i++ {
			cf := evalCompile(t, "RANDBETWEEN(0,1)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || (got.Num != 0 && got.Num != 1) {
				t.Fatalf("RANDBETWEEN(0,1) = %g, want 0 or 1 (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("returns_integer_no_fractional_part", func(t *testing.T) {
		for i := 0; i < 100; i++ {
			cf := evalCompile(t, "RANDBETWEEN(1,100)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != math.Floor(got.Num) {
				t.Fatalf("RANDBETWEEN(1,100) = %g, want integer (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("TYPE_equals_1", func(t *testing.T) {
		cf := evalCompile(t, "TYPE(RANDBETWEEN(1,10))")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("TYPE(RANDBETWEEN(1,10)) = %v, want 1 (number)", got)
		}
	})

	t.Run("ISNUMBER_true", func(t *testing.T) {
		cf := evalCompile(t, "ISNUMBER(RANDBETWEEN(1,10))")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(RANDBETWEEN(1,10)) = %v, want TRUE", got)
		}
	})

	t.Run("error_bottom_greater_than_top", func(t *testing.T) {
		cf := evalCompile(t, "RANDBETWEEN(10,1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("RANDBETWEEN(10,1) = %v, want #NUM! error", got)
		}
	})

	t.Run("error_zero_args", func(t *testing.T) {
		cf := evalCompile(t, "RANDBETWEEN()")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("RANDBETWEEN() = %v, want error", got)
		}
	})

	t.Run("error_one_arg", func(t *testing.T) {
		cf := evalCompile(t, "RANDBETWEEN(1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("RANDBETWEEN(1) = %v, want #VALUE! error", got)
		}
	})

	t.Run("error_three_args", func(t *testing.T) {
		cf := evalCompile(t, "RANDBETWEEN(1,5,10)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("RANDBETWEEN(1,5,10) = %v, want #VALUE! error", got)
		}
	})

	t.Run("error_propagation_first_arg_VALUE", func(t *testing.T) {
		cf := evalCompile(t, `RANDBETWEEN(VALUE("x"),10)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf(`RANDBETWEEN(VALUE("x"),10) = %v, want error`, got)
		}
	})

	t.Run("error_propagation_second_arg_NA", func(t *testing.T) {
		cf := evalCompile(t, "RANDBETWEEN(1,NA())")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("RANDBETWEEN(1,NA()) = %v, want #N/A error", got)
		}
	})

	t.Run("decimal_bounds_ceil_floor", func(t *testing.T) {
		// RANDBETWEEN(1.5, 5.5) → Ceil(1.5)=2, Floor(5.5)=5 → result in [2,5]
		for i := 0; i < 100; i++ {
			cf := evalCompile(t, "RANDBETWEEN(1.5,5.5)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < 2 || got.Num > 5 {
				t.Fatalf("RANDBETWEEN(1.5,5.5) = %g, want [2,5] (iteration %d)", got.Num, i)
			}
			if got.Num != math.Floor(got.Num) {
				t.Fatalf("RANDBETWEEN(1.5,5.5) = %g, want integer (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("large_range", func(t *testing.T) {
		for i := 0; i < 50; i++ {
			cf := evalCompile(t, "RANDBETWEEN(1,1000000)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < 1 || got.Num > 1000000 {
				t.Fatalf("RANDBETWEEN(1,1000000) = %g, want [1,1000000] (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("negative_to_zero", func(t *testing.T) {
		for i := 0; i < 100; i++ {
			cf := evalCompile(t, "RANDBETWEEN(-100,0)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < -100 || got.Num > 0 {
				t.Fatalf("RANDBETWEEN(-100,0) = %g, want [-100,0] (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("boolean_coercion_TRUE_as_1", func(t *testing.T) {
		// TRUE coerces to 1, so RANDBETWEEN(TRUE,5) → RANDBETWEEN(1,5)
		for i := 0; i < 50; i++ {
			cf := evalCompile(t, "RANDBETWEEN(TRUE,5)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < 1 || got.Num > 5 {
				t.Fatalf("RANDBETWEEN(TRUE,5) = %g, want [1,5] (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("string_coercion_numeric_strings", func(t *testing.T) {
		for i := 0; i < 50; i++ {
			cf := evalCompile(t, `RANDBETWEEN("1","10")`)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < 1 || got.Num > 10 {
				t.Fatalf(`RANDBETWEEN("1","10") = %g, want [1,10] (iteration %d)`, got.Num, i)
			}
		}
	})

	t.Run("error_non_numeric_string", func(t *testing.T) {
		cf := evalCompile(t, `RANDBETWEEN("abc",10)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`RANDBETWEEN("abc",10) = %v, want #VALUE! error`, got)
		}
	})

	t.Run("error_non_numeric_string_second_arg", func(t *testing.T) {
		cf := evalCompile(t, `RANDBETWEEN(1,"xyz")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`RANDBETWEEN(1,"xyz") = %v, want #VALUE! error`, got)
		}
	})

	t.Run("decimal_bounds_that_cross_after_ceil_floor", func(t *testing.T) {
		// RANDBETWEEN(1.9, 1.1) → Ceil(1.9)=2, Floor(1.1)=1 → lo>hi → #NUM!
		cf := evalCompile(t, "RANDBETWEEN(1.9,1.1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("RANDBETWEEN(1.9,1.1) = %v, want #NUM! error", got)
		}
	})

	t.Run("boolean_FALSE_as_0", func(t *testing.T) {
		// FALSE coerces to 0, so RANDBETWEEN(FALSE,3) → RANDBETWEEN(0,3)
		for i := 0; i < 50; i++ {
			cf := evalCompile(t, "RANDBETWEEN(FALSE,3)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num < 0 || got.Num > 3 {
				t.Fatalf("RANDBETWEEN(FALSE,3) = %g, want [0,3] (iteration %d)", got.Num, i)
			}
		}
	})

	t.Run("negative_same_bounds", func(t *testing.T) {
		for i := 0; i < 20; i++ {
			cf := evalCompile(t, "RANDBETWEEN(-7,-7)")
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval: %v", err)
			}
			if got.Type != ValueNumber || got.Num != -7 {
				t.Fatalf("RANDBETWEEN(-7,-7) = %g, want -7 (iteration %d)", got.Num, i)
			}
		}
	})
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
		// Documentation example: cos(PI/4) via Taylor series
		{"doc_example", "SERIESSUM(PI()/4,0,2,{1,-0.5,0.041666667,-0.001388889})", 0.707103, 1e-4},
		// Single coefficient
		{"single_coeff", "SERIESSUM(2,3,1,{5})", 40, 0},
		// Single coefficient as scalar
		{"scalar_coeff", "SERIESSUM(2,3,1,5)", 40, 0},
		// x=0, n>0: 0^positive = 0, so result = 0
		{"x_zero_n_pos", "SERIESSUM(0,1,1,{3,4,5})", 0, 0},
		// x=0, n=0: returns #NUM! (0^0 indeterminate) — tested in error cases below
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
		// x=0, n=0, m=0: returns #NUM! (0^0 indeterminate) — tested in error cases below
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
// Additional MMULT tests
// ---------------------------------------------------------------------------

func TestMMULT_3x3_times_3x3(t *testing.T) {
	// {1,2,3;4,5,6;7,8,9} * {9,8,7;6,5,4;3,2,1}
	// Row0: 1*9+2*6+3*3=30, 1*8+2*5+3*2=24, 1*7+2*4+3*1=18
	// Row1: 4*9+5*6+6*3=84, 4*8+5*5+6*2=69, 4*7+5*4+6*1=54
	// Row2: 7*9+8*6+9*3=138, 7*8+8*5+9*2=114, 7*7+8*4+9*1=90
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1,2,3;4,5,6;7,8,9},{9,8,7;6,5,4;3,2,1})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{30, 24, 18}, {84, 69, 54}, {138, 114, 90}}
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

func TestMMULT_IdentityLeft(t *testing.T) {
	// I * M = M
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1,0;0,1},{7,8;9,10})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{7, 8}, {9, 10}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestMMULT_IdentityRight3x3(t *testing.T) {
	// M * I = M for 3x3
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1,2,3;4,5,6;7,8,9},{1,0,0;0,1,0;0,0,1})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{1, 2, 3}, {4, 5, 6}, {7, 8, 9}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if got.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, w)
			}
		}
	}
}

func TestMMULT_IncompatibleDimensions_3x2_times_3x2(t *testing.T) {
	// 3x2 * 3x2 => cols(2) != rows(3) => #VALUE!
	result, err := fnMMULT([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
			{NumberVal(3), NumberVal(4)},
			{NumberVal(5), NumberVal(6)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
			{NumberVal(3), NumberVal(4)},
			{NumberVal(5), NumberVal(6)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", result)
	}
}

func TestMMULT_ZeroMatrixRight(t *testing.T) {
	// M * 0 = 0
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1,2;3,4},{0,0;0,0})")
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

func TestMMULT_NegativeTimesNegative(t *testing.T) {
	// {-1,-2;-3,-4} * {-5,-6;-7,-8}
	// [0][0]: (-1)(-5)+(-2)(-7)=5+14=19
	// [0][1]: (-1)(-6)+(-2)(-8)=6+16=22
	// [1][0]: (-3)(-5)+(-4)(-7)=15+28=43
	// [1][1]: (-3)(-6)+(-4)(-8)=18+32=50
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({-1,-2;-3,-4},{-5,-6;-7,-8})")
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

func TestMMULT_1x4_times_4x1(t *testing.T) {
	// Dot product: {1,2,3,4} * {5;6;7;8} = {1*5+2*6+3*7+4*8} = {70}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1,2,3,4},{5;6;7;8})")
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
	if got.Array[0][0].Num != 70 {
		t.Errorf("got %g, want 70", got.Array[0][0].Num)
	}
}

func TestMMULT_4x1_times_1x4(t *testing.T) {
	// Outer product: {1;2;3;4} * {5,6,7,8}
	result, err := fnMMULT([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
			{NumberVal(2)},
			{NumberVal(3)},
			{NumberVal(4)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(5), NumberVal(6), NumberVal(7), NumberVal(8)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueArray {
		t.Fatalf("expected array, got %v", result.Type)
	}
	want := [][]float64{
		{5, 6, 7, 8},
		{10, 12, 14, 16},
		{15, 18, 21, 24},
		{20, 24, 28, 32},
	}
	if len(result.Array) != 4 || len(result.Array[0]) != 4 {
		t.Fatalf("expected 4x4, got %dx%d", len(result.Array), len(result.Array[0]))
	}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if result.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, result.Array[r][c].Num, w)
			}
		}
	}
}

func TestMMULT_ErrorInArray1_Propagate(t *testing.T) {
	// Error in arg1 (as a scalar error) propagates directly
	result, err := fnMMULT([]Value{
		ErrorVal(ErrValNA),
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValNA {
		t.Errorf("expected #N/A, got %v", result)
	}
}

func TestMMULT_ErrorInArray2_Propagate(t *testing.T) {
	// Error in arg2 (as a scalar error) propagates directly
	result, err := fnMMULT([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1)},
		}},
		ErrorVal(ErrValREF),
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValREF {
		t.Errorf("expected #REF!, got %v", result)
	}
}

func TestMMULT_ErrorInArray2_Cell(t *testing.T) {
	// Error cell in array2 should propagate
	result, err := fnMMULT([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(3)},
			{ErrorVal(ErrValNUM)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValNUM {
		t.Errorf("expected #NUM!, got %v", result)
	}
}

func TestMMULT_VeryLargeValues(t *testing.T) {
	// 1e10 * 1e10 matrix product
	result, err := fnMMULT([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1e10), NumberVal(2e10)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(3e10)},
			{NumberVal(4e10)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueArray {
		t.Fatalf("expected array, got %v", result.Type)
	}
	// 1e10*3e10 + 2e10*4e10 = 3e20 + 8e20 = 1.1e21
	if math.Abs(result.Array[0][0].Num-1.1e21) > 1e10 {
		t.Errorf("got %g, want 1.1e21", result.Array[0][0].Num)
	}
}

func TestMMULT_NonSquareResult_2x3_times_3x4(t *testing.T) {
	// Result should be 2x4
	result, err := fnMMULT([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
		}},
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(0), NumberVal(0), NumberVal(1)},
			{NumberVal(0), NumberVal(1), NumberVal(0), NumberVal(1)},
			{NumberVal(0), NumberVal(0), NumberVal(1), NumberVal(1)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueArray {
		t.Fatalf("expected array, got %v", result.Type)
	}
	if len(result.Array) != 2 || len(result.Array[0]) != 4 {
		t.Fatalf("expected 2x4, got %dx%d", len(result.Array), len(result.Array[0]))
	}
	// Row0: 1*1+2*0+3*0=1, 1*0+2*1+3*0=2, 1*0+2*0+3*1=3, 1*1+2*1+3*1=6
	// Row1: 4*1+5*0+6*0=4, 4*0+5*1+6*0=5, 4*0+5*0+6*1=6, 4*1+5*1+6*1=15
	want := [][]float64{{1, 2, 3, 6}, {4, 5, 6, 15}}
	for r, wantRow := range want {
		for c, w := range wantRow {
			if result.Array[r][c].Num != w {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, result.Array[r][c].Num, w)
			}
		}
	}
}

func TestMMULT_CrossCheck_MINVERSE(t *testing.T) {
	// MMULT(M, MINVERSE(M)) should be identity
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({2,1;5,3},MINVERSE({2,1;5,3}))")
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

func TestMMULT_CrossCheck_MINVERSE_3x3(t *testing.T) {
	// MMULT(M, MINVERSE(M)) should be identity for 3x3
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1,2,3;0,1,4;5,6,0},MINVERSE({1,2,3;0,1,4;5,6,0}))")
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

func TestMMULT_ScalarTimesArray_1x1(t *testing.T) {
	// scalar 3 coerced to {3} (1x1), times {2} (1x1) => {6}
	result, err := fnMMULT([]Value{
		NumberVal(3),
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(2)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMMULT: %v", err)
	}
	if result.Type != ValueArray {
		t.Fatalf("expected array, got %v", result.Type)
	}
	if result.Array[0][0].Num != 6 {
		t.Errorf("got %g, want 6", result.Array[0][0].Num)
	}
}

func TestMMULT_EvalInlineArray(t *testing.T) {
	// Verify inline arrays work through full eval pipeline
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({2,0;0,3},{4;5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	// {2*4+0*5; 0*4+3*5} = {8; 15}
	if len(got.Array) != 2 || len(got.Array[0]) != 1 {
		t.Fatalf("expected 2x1, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[0][0].Num != 8 {
		t.Errorf("[0][0]: got %g, want 8", got.Array[0][0].Num)
	}
	if got.Array[1][0].Num != 15 {
		t.Errorf("[1][0]: got %g, want 15", got.Array[1][0].Num)
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
		{"MDETERM({1,0;0,1})", 1, 0},   // identity
		{"MDETERM({0,0;0,0})", 0, 0},   // zero matrix
		{"MDETERM({2,0;0,3})", 6, 0},   // diagonal
		{"MDETERM({1,2;2,4})", 0, 0},   // singular
		{"MDETERM({-1,-2;3,4})", 2, 0}, // negative numbers

		// 3x3 matrices
		{"MDETERM({3,6,1;1,1,0;3,10,2})", 1, 1e-10},
		{"MDETERM({1,0,0;0,1,0;0,0,1})", 1, 0},      // 3x3 identity
		{"MDETERM({2,0,0;0,3,0;0,0,4})", 24, 1e-10}, // diagonal
		{"MDETERM({1,2,3;4,5,6;7,8,9})", 0, 1e-10},  // singular

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

func TestMDETERM_WithCellRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 2, Row: 2}: NumberVal(4),
		},
	}
	cf := evalCompile(t, "MDETERM(A1:B2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -2 {
		t.Fatalf("MDETERM(A1:B2) = %v, want -2", got)
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

// ---------------------------------------------------------------------------
// Additional MINVERSE tests
// ---------------------------------------------------------------------------

func TestMINVERSE_2x2_Identity(t *testing.T) {
	// Inverse of 2x2 identity is identity.
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({1,0;0,1})")
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

func TestMINVERSE_1x1_Array(t *testing.T) {
	// {5} as array => {{0.2}}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({5})")
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
	if math.Abs(got.Array[0][0].Num-0.2) > 1e-10 {
		t.Errorf("got %g, want 0.2", got.Array[0][0].Num)
	}
}

func TestMINVERSE_4x4_Identity(t *testing.T) {
	// Inverse of 4x4 identity is identity.
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({1,0,0,0;0,1,0,0;0,0,1,0;0,0,0,1})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	for r := 0; r < 4; r++ {
		for c := 0; c < 4; c++ {
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

func TestMINVERSE_3x3_NonSquare_2x3(t *testing.T) {
	// 2x3 is not square => #VALUE!
	result, err := fnMINVERSE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2), NumberVal(3)},
			{NumberVal(4), NumberVal(5), NumberVal(6)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", result)
	}
}

func TestMINVERSE_3x3_NonSquare_3x2(t *testing.T) {
	// 3x2 is not square => #VALUE!
	result, err := fnMINVERSE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
			{NumberVal(3), NumberVal(4)},
			{NumberVal(5), NumberVal(6)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", result)
	}
}

func TestMINVERSE_LargeValues(t *testing.T) {
	// {1000,0;0,500} => {0.001,0;0,0.002}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({1000,0;0,500})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{0.001, 0}, {0, 0.002}}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			if math.Abs(got.Array[r][c].Num-want[r][c]) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestMINVERSE_WithCellRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 2, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 2, Row: 2}: NumberVal(4),
		},
	}
	cf := evalCompile(t, "MINVERSE(A1:B2)")
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

func TestMINVERSE_NearSingular(t *testing.T) {
	// A matrix with a very small determinant should still invert if above threshold.
	// {1, 1; 1, 1.0000000001} has det ~ 1e-10 which is above 1e-15 threshold.
	result, err := fnMINVERSE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(1)},
			{NumberVal(1), NumberVal(1.0000000001)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	// Should succeed (not singular) because det is ~1e-10, above the 1e-15 threshold
	if result.Type == ValueError {
		t.Errorf("expected successful inverse for near-singular matrix, got error %v", result.Err)
	}
}

func TestMINVERSE_NegativeElements_2x2(t *testing.T) {
	// {-3,-4;-1,-2} => det = (-3)(-2)-(-4)(-1) = 6-4 = 2
	// inverse = 1/2 * {-2,4;1,-3} = {-1,2;0.5,-1.5}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({-3,-4;-1,-2})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{-1, 2}, {0.5, -1.5}}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			if math.Abs(got.Array[r][c].Num-want[r][c]) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestMINVERSE_MTimesInverseM_3x3(t *testing.T) {
	// Cross-check: MMULT(M, MINVERSE(M)) ≈ identity for 3x3
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({2,1,1;1,3,2;1,0,0},MINVERSE({2,1,1;1,3,2;1,0,0}))")
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
			if math.Abs(got.Array[r][c].Num-want) > 1e-9 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want)
			}
		}
	}
}

func TestMINVERSE_InverseTimesM_3x3(t *testing.T) {
	// Cross-check from the other side: MMULT(MINVERSE(M), M) ≈ identity
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT(MINVERSE({2,1,1;1,3,2;1,0,0}),{2,1,1;1,3,2;1,0,0})")
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
			if math.Abs(got.Array[r][c].Num-want) > 1e-9 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want)
			}
		}
	}
}

func TestMINVERSE_ErrorInArray_NA(t *testing.T) {
	// #N/A error inside array should propagate
	result, err := fnMINVERSE([]Value{
		{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), ErrorVal(ErrValNA)},
			{NumberVal(3), NumberVal(4)},
		}},
	})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValNA {
		t.Errorf("expected #N/A, got %v", result)
	}
}

func TestMINVERSE_ScalarNegative(t *testing.T) {
	// Scalar -4 => 1/(-4) = -0.25
	result, err := fnMINVERSE([]Value{NumberVal(-4)})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueNumber {
		t.Fatalf("expected number, got %v", result.Type)
	}
	if math.Abs(result.Num-(-0.25)) > 1e-10 {
		t.Errorf("got %g, want -0.25", result.Num)
	}
}

func TestMINVERSE_4x4_General(t *testing.T) {
	// Verify 4x4 matrix inverse via product check
	resolver := &mockResolver{}
	cf := evalCompile(t, "MMULT({1,0,2,0;1,1,0,0;1,2,0,1;1,1,1,1},MINVERSE({1,0,2,0;1,1,0,0;1,2,0,1;1,1,1,1}))")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	for r := 0; r < 4; r++ {
		for c := 0; c < 4; c++ {
			want := 0.0
			if r == c {
				want = 1.0
			}
			if math.Abs(got.Array[r][c].Num-want) > 1e-9 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want)
			}
		}
	}
}

func TestMINVERSE_2x2_AllNegative(t *testing.T) {
	// {-2,-5;-1,-3} => det = (-2)(-3)-(-5)(-1) = 6-5 = 1
	// inverse = {-3,5;1,-2}
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({-2,-5;-1,-3})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := [][]float64{{-3, 5}, {1, -2}}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			if math.Abs(got.Array[r][c].Num-want[r][c]) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestMINVERSE_EvalInlineArray(t *testing.T) {
	// Verify the full eval pipeline with inline arrays
	resolver := &mockResolver{}
	cf := evalCompile(t, "MINVERSE({3,1;5,2})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	// det = 3*2 - 1*5 = 1, inverse = {2,-1;-5,3}
	want := [][]float64{{2, -1}, {-5, 3}}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			if math.Abs(got.Array[r][c].Num-want[r][c]) > 1e-10 {
				t.Errorf("[%d][%d]: got %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestMINVERSE_StringArg(t *testing.T) {
	// String scalar arg => #VALUE!
	result, err := fnMINVERSE([]Value{StringVal("hello")})
	if err != nil {
		t.Fatalf("fnMINVERSE: %v", err)
	}
	if result.Type != ValueError || result.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got %v", result)
	}
}

func TestMatrixFunctions_InvalidMatrixInputs(t *testing.T) {
	tests := []struct {
		name string
		fn   func([]Value) (Value, error)
		arg  Value
	}{
		{
			name: "mdeterm_jagged_array",
			fn:   fnMDETERM,
			arg: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2)},
				{NumberVal(3)},
			}},
		},
		{
			name: "mdeterm_trimmed_range_missing_row",
			fn:   fnMDETERM,
			arg: trimmedRangeValue([][]Value{
				{NumberVal(1), NumberVal(2)},
			}, 1, 1, 2, 2),
		},
		{
			name: "minverse_jagged_array",
			fn:   fnMINVERSE,
			arg: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(1), NumberVal(2)},
				{NumberVal(3)},
			}},
		},
		{
			name: "minverse_trimmed_range_missing_row",
			fn:   fnMINVERSE,
			arg: trimmedRangeValue([][]Value{
				{NumberVal(1), NumberVal(2)},
			}, 1, 1, 2, 2),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := tt.fn([]Value{tt.arg})
			if err != nil {
				t.Fatalf("fn: %v", err)
			}
			if got.Type != ValueError || got.Err != ErrValVALUE {
				t.Fatalf("got %v, want #VALUE!", got)
			}
		})
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

// --------------- SUMX2MY2 additional tests ---------------

func TestSUMX2MY2_ZeroArgs(t *testing.T) {
	got, err := fnSumx2my2(nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2MY2_OneArrayAllZeros(t *testing.T) {
	resolver := &mockResolver{}
	// When y=0: SUMX2MY2 = SUMSQ(x) = 1+4+9 = 14
	cf := evalCompile(t, "SUMX2MY2({1,2,3},{0,0,0})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 14 {
		t.Errorf("got %v (%g), want 14", got.Type, got.Num)
	}
}

func TestSUMX2MY2_IdentitySumSqDiff(t *testing.T) {
	// Mathematical identity: SUMX2MY2(x,y) = SUMSQ(x) - SUMSQ(y)
	// x={3,4,5}, y={1,2,3}: SUMSQ(x)=9+16+25=50, SUMSQ(y)=1+4+9=14, diff=36
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2MY2({3,4,5},{1,2,3})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 36 {
		t.Errorf("got %v (%g), want 36", got.Type, got.Num)
	}
}

func TestSUMX2MY2_MixedPositiveNegativeLarger(t *testing.T) {
	resolver := &mockResolver{}
	// {-5,10,-15,20} vs {3,-7,11,-13}
	// (25-9)+(100-49)+(225-121)+(400-169) = 16+51+104+231 = 402
	cf := evalCompile(t, "SUMX2MY2({-5,10,-15,20},{3,-7,11,-13})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 402 {
		t.Errorf("got %v (%g), want 402", got.Type, got.Num)
	}
}

func TestSUMX2MY2_VeryLargeValues(t *testing.T) {
	resolver := &mockResolver{}
	// {10000,20000} vs {30000,40000}
	// (1e8-9e8)+(4e8-16e8) = -8e8 + -12e8 = -2e9
	cf := evalCompile(t, "SUMX2MY2({10000,20000},{30000,40000})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -2e9 {
		t.Errorf("got %v (%g), want -2e9", got.Type, got.Num)
	}
}

func TestSUMX2MY2_FractionalDecimals(t *testing.T) {
	resolver := &mockResolver{}
	// {0.1,0.2,0.3} vs {0.4,0.5,0.6}
	// (0.01-0.16)+(0.04-0.25)+(0.09-0.36) = -0.15-0.21-0.27 = -0.63
	cf := evalCompile(t, "SUMX2MY2({0.1,0.2,0.3},{0.4,0.5,0.6})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if math.Abs(got.Num-(-0.63)) > 1e-10 {
		t.Errorf("got %v (%g), want -0.63", got.Type, got.Num)
	}
}

func TestSUMX2MY2_TextInSecondArray(t *testing.T) {
	got, err := fnSumx2my2([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), StringVal("hello")}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2MY2_MultiRow2DArray(t *testing.T) {
	// 2D array: {{1,2},{3,4}} vs {{5,6},{7,8}} flattened: {1,2,3,4} vs {5,6,7,8}
	// (1-25)+(4-36)+(9-49)+(16-64) = -24-32-40-48 = -144
	got, err := fnSumx2my2([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}, {NumberVal(3), NumberVal(4)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(5), NumberVal(6)}, {NumberVal(7), NumberVal(8)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -144 {
		t.Errorf("got %v (%g), want -144", got.Type, got.Num)
	}
}

func TestSUMX2MY2_BothBoolArrays(t *testing.T) {
	// {TRUE,FALSE} vs {FALSE,TRUE} => {1,0} vs {0,1}
	// (1-0)+(0-1) = 1-1 = 0
	got, err := fnSumx2my2([]Value{
		{Type: ValueArray, Array: [][]Value{{BoolVal(true), BoolVal(false)}}},
		{Type: ValueArray, Array: [][]Value{{BoolVal(false), BoolVal(true)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("got %v (%g), want 0", got.Type, got.Num)
	}
}

func TestSUMX2MY2_NegativeDecimals(t *testing.T) {
	resolver := &mockResolver{}
	// {-1.5,-2.5} vs {-0.5,-1.5}
	// (2.25-0.25)+(6.25-2.25) = 2+4 = 6
	cf := evalCompile(t, "SUMX2MY2({-1.5,-2.5},{-0.5,-1.5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 6 {
		t.Errorf("got %v (%g), want 6", got.Type, got.Num)
	}
}

func TestSUMX2MY2_AntiSymmetry(t *testing.T) {
	// SUMX2MY2(x,y) = -SUMX2MY2(y,x)
	resolver := &mockResolver{}
	cf1 := evalCompile(t, "SUMX2MY2({2,5,8},{3,4,6})")
	got1, err := Eval(cf1, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	cf2 := evalCompile(t, "SUMX2MY2({3,4,6},{2,5,8})")
	got2, err := Eval(cf2, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got1.Num != -got2.Num {
		t.Errorf("anti-symmetry failed: SUMX2MY2(x,y)=%g, SUMX2MY2(y,x)=%g", got1.Num, got2.Num)
	}
}

func TestSUMX2MY2_CellRangeWithEmptyCell(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: NumberVal(4),
			// A3 is empty (treated as 0)
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(5),
		},
	}
	cf := evalCompile(t, "SUMX2MY2(A1:A3,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (9-1)+(16-4)+(0-25) = 8+12-25 = -5
	if got.Type != ValueNumber || got.Num != -5 {
		t.Errorf("got %v (%g), want -5", got.Type, got.Num)
	}
}

// --------------- SUMX2PY2 additional tests ---------------

func TestSUMX2PY2_ZeroArgs(t *testing.T) {
	got, err := fnSumx2py2(nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2PY2_OneArrayAllZeros(t *testing.T) {
	resolver := &mockResolver{}
	// When y=0: SUMX2PY2 = SUMSQ(x) = 1+4+9 = 14
	cf := evalCompile(t, "SUMX2PY2({1,2,3},{0,0,0})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 14 {
		t.Errorf("got %v (%g), want 14", got.Type, got.Num)
	}
}

func TestSUMX2PY2_IdentityWithSUMX2MY2(t *testing.T) {
	// SUMX2PY2(x,y) = SUMX2MY2(x,y) + 2*SUMSQ(y)
	// x={2,3}, y={4,5}: SUMX2PY2=(4+16)+(9+25)=54
	// SUMX2MY2=(4-16)+(9-25)=-28, SUMSQ(y)=16+25=41, -28+82=54 ✓
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2PY2({2,3},{4,5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 54 {
		t.Errorf("got %v (%g), want 54", got.Type, got.Num)
	}
}

func TestSUMX2PY2_MixedPositiveNegativeLarger(t *testing.T) {
	resolver := &mockResolver{}
	// {-5,10,-15,20} vs {3,-7,11,-13}
	// (25+9)+(100+49)+(225+121)+(400+169) = 34+149+346+569 = 1098
	cf := evalCompile(t, "SUMX2PY2({-5,10,-15,20},{3,-7,11,-13})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1098 {
		t.Errorf("got %v (%g), want 1098", got.Type, got.Num)
	}
}

func TestSUMX2PY2_VeryLargeValues(t *testing.T) {
	resolver := &mockResolver{}
	// {10000,20000} vs {30000,40000}
	// (1e8+9e8)+(4e8+16e8) = 10e8+20e8 = 3e9
	cf := evalCompile(t, "SUMX2PY2({10000,20000},{30000,40000})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3e9 {
		t.Errorf("got %v (%g), want 3e9", got.Type, got.Num)
	}
}

func TestSUMX2PY2_FractionalDecimals(t *testing.T) {
	resolver := &mockResolver{}
	// {0.1,0.2,0.3} vs {0.4,0.5,0.6}
	// (0.01+0.16)+(0.04+0.25)+(0.09+0.36) = 0.17+0.29+0.45 = 0.91
	cf := evalCompile(t, "SUMX2PY2({0.1,0.2,0.3},{0.4,0.5,0.6})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if math.Abs(got.Num-0.91) > 1e-10 {
		t.Errorf("got %v (%g), want 0.91", got.Type, got.Num)
	}
}

func TestSUMX2PY2_TextInSecondArray(t *testing.T) {
	got, err := fnSumx2py2([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), StringVal("xyz")}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMX2PY2_MultiRow2DArray(t *testing.T) {
	// 2D array: {{1,2},{3,4}} vs {{5,6},{7,8}} flattened: {1,2,3,4} vs {5,6,7,8}
	// (1+25)+(4+36)+(9+49)+(16+64) = 26+40+58+80 = 204
	got, err := fnSumx2py2([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}, {NumberVal(3), NumberVal(4)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(5), NumberVal(6)}, {NumberVal(7), NumberVal(8)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 204 {
		t.Errorf("got %v (%g), want 204", got.Type, got.Num)
	}
}

func TestSUMX2PY2_BothBoolArrays(t *testing.T) {
	// {TRUE,FALSE} vs {FALSE,TRUE} => {1,0} vs {0,1}
	// (1+0)+(0+1) = 1+1 = 2
	got, err := fnSumx2py2([]Value{
		{Type: ValueArray, Array: [][]Value{{BoolVal(true), BoolVal(false)}}},
		{Type: ValueArray, Array: [][]Value{{BoolVal(false), BoolVal(true)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("got %v (%g), want 2", got.Type, got.Num)
	}
}

func TestSUMX2PY2_NegativeDecimals(t *testing.T) {
	resolver := &mockResolver{}
	// {-1.5,-2.5} vs {-0.5,-1.5}
	// (2.25+0.25)+(6.25+2.25) = 2.5+8.5 = 11
	cf := evalCompile(t, "SUMX2PY2({-1.5,-2.5},{-0.5,-1.5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 11 {
		t.Errorf("got %v (%g), want 11", got.Type, got.Num)
	}
}

func TestSUMX2PY2_Symmetric(t *testing.T) {
	// SUMX2PY2(x,y) = SUMX2PY2(y,x) — it's commutative
	resolver := &mockResolver{}
	cf1 := evalCompile(t, "SUMX2PY2({2,5,8},{3,4,6})")
	got1, err := Eval(cf1, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	cf2 := evalCompile(t, "SUMX2PY2({3,4,6},{2,5,8})")
	got2, err := Eval(cf2, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got1.Num != got2.Num {
		t.Errorf("symmetry failed: SUMX2PY2(x,y)=%g, SUMX2PY2(y,x)=%g", got1.Num, got2.Num)
	}
}

func TestSUMX2PY2_CellRangeWithEmptyCell(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: NumberVal(4),
			// A3 is empty (treated as 0)
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(5),
		},
	}
	cf := evalCompile(t, "SUMX2PY2(A1:A3,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (9+1)+(16+4)+(0+25) = 10+20+25 = 55
	if got.Type != ValueNumber || got.Num != 55 {
		t.Errorf("got %v (%g), want 55", got.Type, got.Num)
	}
}

func TestSUMX2PY2_AlwaysNonNegative(t *testing.T) {
	// SUMX2PY2 result is always >= 0 for any real inputs
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMX2PY2({-99,-50,0,50,99},{-88,-44,0,44,88})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num < 0 {
		t.Errorf("expected non-negative, got %v (%g)", got.Type, got.Num)
	}
	// Exact: (9801+7744)+(2500+1936)+(0+0)+(2500+1936)+(9801+7744) = 43962
	if got.Num != 43962 {
		t.Errorf("got %g, want 43962", got.Num)
	}
}

// --------------- SUMXMY2 additional tests ---------------

func TestSUMXMY2_ZeroArgs(t *testing.T) {
	got, err := fnSumxmy2(nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMXMY2_OneArrayAllZeros(t *testing.T) {
	resolver := &mockResolver{}
	// When y=0: SUMXMY2 = SUMSQ(x) = 1+4+9 = 14
	cf := evalCompile(t, "SUMXMY2({1,2,3},{0,0,0})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 14 {
		t.Errorf("got %v (%g), want 14", got.Type, got.Num)
	}
}

func TestSUMXMY2_IdentityCrossCheck(t *testing.T) {
	// SUMXMY2(x,y) = SUMX2PY2(x,y) - 2*SUMPRODUCT(x,y)
	// x={2,3,4}, y={1,5,2}
	// SUMX2PY2 = (4+1)+(9+25)+(16+4) = 5+34+20 = 59
	// SUMPRODUCT = 2*1 + 3*5 + 4*2 = 2+15+8 = 25
	// SUMXMY2 = 59 - 50 = 9
	// Direct: (2-1)²+(3-5)²+(4-2)² = 1+4+4 = 9 ✓
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMXMY2({2,3,4},{1,5,2})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 9 {
		t.Errorf("got %v (%g), want 9", got.Type, got.Num)
	}
}

func TestSUMXMY2_MixedPositiveNegativeLarger(t *testing.T) {
	resolver := &mockResolver{}
	// {-5,10,-15,20} vs {3,-7,11,-13}
	// (-5-3)²+(10-(-7))²+(-15-11)²+(20-(-13))²
	// = (-8)²+(17)²+(-26)²+(33)² = 64+289+676+1089 = 2118
	cf := evalCompile(t, "SUMXMY2({-5,10,-15,20},{3,-7,11,-13})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2118 {
		t.Errorf("got %v (%g), want 2118", got.Type, got.Num)
	}
}

func TestSUMXMY2_VeryLargeValues(t *testing.T) {
	resolver := &mockResolver{}
	// {10000,20000} vs {30000,40000}
	// (-20000)²+(-20000)² = 4e8+4e8 = 8e8
	cf := evalCompile(t, "SUMXMY2({10000,20000},{30000,40000})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 8e8 {
		t.Errorf("got %v (%g), want 8e8", got.Type, got.Num)
	}
}

func TestSUMXMY2_FractionalDecimals(t *testing.T) {
	resolver := &mockResolver{}
	// {0.1,0.2,0.3} vs {0.4,0.5,0.6}
	// (-0.3)²+(-0.3)²+(-0.3)² = 0.09+0.09+0.09 = 0.27
	cf := evalCompile(t, "SUMXMY2({0.1,0.2,0.3},{0.4,0.5,0.6})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if math.Abs(got.Num-0.27) > 1e-10 {
		t.Errorf("got %v (%g), want 0.27", got.Type, got.Num)
	}
}

func TestSUMXMY2_TextInSecondArray(t *testing.T) {
	got, err := fnSumxmy2([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(3), StringVal("nope")}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE!, got type=%v err=%v", got.Type, got.Err)
	}
}

func TestSUMXMY2_MultiRow2DArray(t *testing.T) {
	// 2D array: {{1,2},{3,4}} vs {{5,6},{7,8}} flattened: {1,2,3,4} vs {5,6,7,8}
	// (-4)²+(-4)²+(-4)²+(-4)² = 16+16+16+16 = 64
	got, err := fnSumxmy2([]Value{
		{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}, {NumberVal(3), NumberVal(4)}}},
		{Type: ValueArray, Array: [][]Value{{NumberVal(5), NumberVal(6)}, {NumberVal(7), NumberVal(8)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 64 {
		t.Errorf("got %v (%g), want 64", got.Type, got.Num)
	}
}

func TestSUMXMY2_BothBoolArrays(t *testing.T) {
	// {TRUE,FALSE} vs {FALSE,TRUE} => {1,0} vs {0,1}
	// (1-0)²+(0-1)² = 1+1 = 2
	got, err := fnSumxmy2([]Value{
		{Type: ValueArray, Array: [][]Value{{BoolVal(true), BoolVal(false)}}},
		{Type: ValueArray, Array: [][]Value{{BoolVal(false), BoolVal(true)}}},
	})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("got %v (%g), want 2", got.Type, got.Num)
	}
}

func TestSUMXMY2_NegativeDecimals(t *testing.T) {
	resolver := &mockResolver{}
	// {-1.5,-2.5} vs {-0.5,-1.5}
	// (-1.5-(-0.5))²+(-2.5-(-1.5))² = (-1)²+(-1)² = 1+1 = 2
	cf := evalCompile(t, "SUMXMY2({-1.5,-2.5},{-0.5,-1.5})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("got %v (%g), want 2", got.Type, got.Num)
	}
}

func TestSUMXMY2_Symmetric(t *testing.T) {
	// SUMXMY2(x,y) = SUMXMY2(y,x) because (x-y)² = (y-x)²
	resolver := &mockResolver{}
	cf1 := evalCompile(t, "SUMXMY2({2,5,8},{3,4,6})")
	got1, err := Eval(cf1, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	cf2 := evalCompile(t, "SUMXMY2({3,4,6},{2,5,8})")
	got2, err := Eval(cf2, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got1.Num != got2.Num {
		t.Errorf("symmetry failed: SUMXMY2(x,y)=%g, SUMXMY2(y,x)=%g", got1.Num, got2.Num)
	}
}

func TestSUMXMY2_CellRangeWithEmptyCell(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(3),
			{Col: 1, Row: 2}: NumberVal(4),
			// A3 is empty (treated as 0)
			{Col: 2, Row: 1}: NumberVal(1),
			{Col: 2, Row: 2}: NumberVal(2),
			{Col: 2, Row: 3}: NumberVal(5),
		},
	}
	cf := evalCompile(t, "SUMXMY2(A1:A3,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// (3-1)²+(4-2)²+(0-5)² = 4+4+25 = 33
	if got.Type != ValueNumber || got.Num != 33 {
		t.Errorf("got %v (%g), want 33", got.Type, got.Num)
	}
}

func TestSUMXMY2_AlwaysNonNegative(t *testing.T) {
	// SUMXMY2 result is always >= 0
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUMXMY2({-99,-50,0,50,99},{-88,-44,0,44,88})")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num < 0 {
		t.Errorf("expected non-negative, got %v (%g)", got.Type, got.Num)
	}
	// Exact: (-11)²+(-6)²+0+(6)²+(11)² = 121+36+0+36+121 = 314
	if got.Num != 314 {
		t.Errorf("got %g, want 314", got.Num)
	}
}

// --------------- Cross-function identity tests ---------------

func TestSUMX_TripleIdentity(t *testing.T) {
	// For x={1,2,3,4}, y={5,6,7,8}:
	// SUMX2PY2 = SUMX2MY2 + 2*SUMSQ(y)
	// SUMXMY2 = SUMX2PY2 - 2*SUMPRODUCT(x,y)
	resolver := &mockResolver{}

	cf1 := evalCompile(t, "SUMX2MY2({1,2,3,4},{5,6,7,8})")
	x2my2, err := Eval(cf1, resolver, nil)
	if err != nil {
		t.Fatalf("SUMX2MY2 error: %v", err)
	}

	cf2 := evalCompile(t, "SUMX2PY2({1,2,3,4},{5,6,7,8})")
	x2py2, err := Eval(cf2, resolver, nil)
	if err != nil {
		t.Fatalf("SUMX2PY2 error: %v", err)
	}

	cf3 := evalCompile(t, "SUMXMY2({1,2,3,4},{5,6,7,8})")
	xmy2, err := Eval(cf3, resolver, nil)
	if err != nil {
		t.Fatalf("SUMXMY2 error: %v", err)
	}

	// Manual: SUMSQ(y) = 25+36+49+64 = 174
	sumsqY := 174.0
	// Manual: SUMPRODUCT(x,y) = 5+12+21+32 = 70
	sumprod := 70.0

	// Identity 1: SUMX2PY2 = SUMX2MY2 + 2*SUMSQ(y)
	lhs1 := x2py2.Num
	rhs1 := x2my2.Num + 2*sumsqY
	if math.Abs(lhs1-rhs1) > 1e-10 {
		t.Errorf("identity SUMX2PY2 = SUMX2MY2 + 2*SUMSQ(y) failed: %g != %g", lhs1, rhs1)
	}

	// Identity 2: SUMXMY2 = SUMX2PY2 - 2*SUMPRODUCT(x,y)
	lhs2 := xmy2.Num
	rhs2 := x2py2.Num - 2*sumprod
	if math.Abs(lhs2-rhs2) > 1e-10 {
		t.Errorf("identity SUMXMY2 = SUMX2PY2 - 2*SUMPRODUCT failed: %g != %g", lhs2, rhs2)
	}

	// Identity 3: SUMX2MY2 = SUMSQ(x) - SUMSQ(y)
	// SUMSQ(x) = 1+4+9+16 = 30
	sumsqX := 30.0
	lhs3 := x2my2.Num
	rhs3 := sumsqX - sumsqY
	if math.Abs(lhs3-rhs3) > 1e-10 {
		t.Errorf("identity SUMX2MY2 = SUMSQ(x) - SUMSQ(y) failed: %g != %g", lhs3, rhs3)
	}
}

func TestSUMX_IdentityWithNegatives(t *testing.T) {
	// Same identities with negative and mixed values
	// x={-3,7,-2,10}, y={4,-5,8,-1}
	resolver := &mockResolver{}

	cf1 := evalCompile(t, "SUMX2MY2({-3,7,-2,10},{4,-5,8,-1})")
	x2my2, err := Eval(cf1, resolver, nil)
	if err != nil {
		t.Fatalf("SUMX2MY2 error: %v", err)
	}

	cf2 := evalCompile(t, "SUMX2PY2({-3,7,-2,10},{4,-5,8,-1})")
	x2py2, err := Eval(cf2, resolver, nil)
	if err != nil {
		t.Fatalf("SUMX2PY2 error: %v", err)
	}

	cf3 := evalCompile(t, "SUMXMY2({-3,7,-2,10},{4,-5,8,-1})")
	xmy2, err := Eval(cf3, resolver, nil)
	if err != nil {
		t.Fatalf("SUMXMY2 error: %v", err)
	}

	// SUMSQ(x) = 9+49+4+100 = 162
	sumsqX := 162.0
	// SUMSQ(y) = 16+25+64+1 = 106
	sumsqY := 106.0
	// SUMPRODUCT = -12 + -35 + -16 + -10 = -73
	sumprod := -73.0

	// Identity 1: SUMX2MY2 = SUMSQ(x) - SUMSQ(y)
	if math.Abs(x2my2.Num-(sumsqX-sumsqY)) > 1e-10 {
		t.Errorf("identity SUMX2MY2 = SUMSQ(x) - SUMSQ(y) failed: %g != %g", x2my2.Num, sumsqX-sumsqY)
	}

	// Identity 2: SUMX2PY2 = SUMX2MY2 + 2*SUMSQ(y)
	if math.Abs(x2py2.Num-(x2my2.Num+2*sumsqY)) > 1e-10 {
		t.Errorf("identity SUMX2PY2 = SUMX2MY2 + 2*SUMSQ(y) failed: %g != %g", x2py2.Num, x2my2.Num+2*sumsqY)
	}

	// Identity 3: SUMXMY2 = SUMX2PY2 - 2*SUMPRODUCT
	if math.Abs(xmy2.Num-(x2py2.Num-2*sumprod)) > 1e-10 {
		t.Errorf("identity SUMXMY2 = SUMX2PY2 - 2*SUMPRODUCT failed: %g != %g", xmy2.Num, x2py2.Num-2*sumprod)
	}
}

func TestSUMX_EqualArraysProperties(t *testing.T) {
	// When x = y:
	// SUMX2MY2 = 0 (x²-y² = 0 for each pair)
	// SUMXMY2 = 0 ((x-y)² = 0 for each pair)
	// SUMX2PY2 = 2*SUMSQ(x)
	resolver := &mockResolver{}

	cf1 := evalCompile(t, "SUMX2MY2({3,7,11},{3,7,11})")
	x2my2, err := Eval(cf1, resolver, nil)
	if err != nil {
		t.Fatalf("SUMX2MY2 error: %v", err)
	}
	if x2my2.Num != 0 {
		t.Errorf("SUMX2MY2 with equal arrays: got %g, want 0", x2my2.Num)
	}

	cf2 := evalCompile(t, "SUMXMY2({3,7,11},{3,7,11})")
	xmy2, err := Eval(cf2, resolver, nil)
	if err != nil {
		t.Fatalf("SUMXMY2 error: %v", err)
	}
	if xmy2.Num != 0 {
		t.Errorf("SUMXMY2 with equal arrays: got %g, want 0", xmy2.Num)
	}

	cf3 := evalCompile(t, "SUMX2PY2({3,7,11},{3,7,11})")
	x2py2, err := Eval(cf3, resolver, nil)
	if err != nil {
		t.Fatalf("SUMX2PY2 error: %v", err)
	}
	// 2*SUMSQ(3,7,11) = 2*(9+49+121) = 2*179 = 358
	if x2py2.Num != 358 {
		t.Errorf("SUMX2PY2 with equal arrays: got %g, want 358", x2py2.Num)
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
		// Doc example: INT(8.9) = 8
		{"doc_pos", "INT(8.9)", 8},
		// Doc example: INT(-8.9) = -9
		{"doc_neg", "INT(-8.9)", -9},
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
		// Doc example: SQRT(ABS(-16)) = 4
		{"doc_abs_neg16", "SQRT(ABS(-16))", 4},
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

func TestSQRTPI(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Doc examples
		{"sqrtpi_1", "SQRTPI(1)", math.Sqrt(math.Pi)},
		{"sqrtpi_2", "SQRTPI(2)", math.Sqrt(2 * math.Pi)},
		// Zero
		{"sqrtpi_0", "SQRTPI(0)", 0},
		// Small values
		{"sqrtpi_0.5", "SQRTPI(0.5)", math.Sqrt(0.5 * math.Pi)},
		{"sqrtpi_0.1", "SQRTPI(0.1)", math.Sqrt(0.1 * math.Pi)},
		{"sqrtpi_0.01", "SQRTPI(0.01)", math.Sqrt(0.01 * math.Pi)},
		// Integer values
		{"sqrtpi_3", "SQRTPI(3)", math.Sqrt(3 * math.Pi)},
		{"sqrtpi_4", "SQRTPI(4)", math.Sqrt(4 * math.Pi)},
		{"sqrtpi_10", "SQRTPI(10)", math.Sqrt(10 * math.Pi)},
		// Large values
		{"sqrtpi_100", "SQRTPI(100)", math.Sqrt(100 * math.Pi)},
		{"sqrtpi_1000000", "SQRTPI(1000000)", math.Sqrt(1000000 * math.Pi)},
		// String coercion
		{"string_coerce_1", "SQRTPI(\"1\")", math.Sqrt(math.Pi)},
		{"string_coerce_2", "SQRTPI(\"2\")", math.Sqrt(2 * math.Pi)},
		{"string_coerce_decimal", "SQRTPI(\"0.5\")", math.Sqrt(0.5 * math.Pi)},
		// Boolean coercion
		{"bool_true", "SQRTPI(TRUE)", math.Sqrt(math.Pi)},
		{"bool_false", "SQRTPI(FALSE)", 0},
		// Known approximate values
		{"approx_1", "SQRTPI(1)", 1.7724538509055159},
		{"approx_2", "SQRTPI(2)", 2.5066282746310002},
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
		{"negative", "SQRTPI(-1)", ErrValNUM},
		{"negative_large", "SQRTPI(-100)", ErrValNUM},
		{"negative_small", "SQRTPI(-0.001)", ErrValNUM},
		{"no_args", "SQRTPI()", ErrValVALUE},
		{"too_many_args", "SQRTPI(1,2)", ErrValVALUE},
		{"non_numeric", "SQRTPI(\"abc\")", ErrValVALUE},
		{"error_propagation", "SQRTPI(1/0)", ErrValDIV0},
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

		// Doc examples
		{"doc_example_1", "LOG(10)", 1, 0},
		{"doc_example_2", "LOG(8,2)", 3, 0},
		{"doc_example_3", "LOG(86,2.7182818)", 4.4543473, 1e-4},

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

func TestLOG10(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Powers of 10 — exact integer results
		{"log10_1", "LOG10(1)", 0, 0},
		{"log10_10", "LOG10(10)", 1, 0},
		{"log10_100", "LOG10(100)", 2, 0},
		{"log10_1000", "LOG10(1000)", 3, 1e-10},
		{"log10_10000", "LOG10(10000)", 4, 1e-10},

		// Negative powers of 10
		{"log10_0_1", "LOG10(0.1)", -1, 1e-10},
		{"log10_0_01", "LOG10(0.01)", -2, 1e-10},
		{"log10_0_001", "LOG10(0.001)", -3, 1e-10},

		// Non-power-of-10 values
		{"log10_2", "LOG10(2)", 0.30102999566398120, 1e-10},
		{"log10_5", "LOG10(5)", 0.69897000433601880, 1e-10},
		{"log10_50", "LOG10(50)", 1.69897000433601880, 1e-10},

		// Large values
		{"log10_1e6", "LOG10(1000000)", 6, 1e-10},
		{"log10_1e10", "LOG10(10000000000)", 10, 1e-10},
		{"log10_1e15", "LOG10(1E15)", 15, 1e-6},

		// Small positive values
		{"log10_1e-5", "LOG10(0.00001)", -5, 1e-10},
		{"log10_1e-10", "LOG10(1E-10)", -10, 1e-6},

		// Documentation example
		{"doc_example_86", "LOG10(86)", 1.93449845124357, 1e-10},
		{"doc_example_10", "LOG10(10)", 1, 0},
		{"doc_example_1e5", "LOG10(1E5)", 5, 1e-10},

		// Boolean coercion: TRUE=1 -> LOG10(1)=0
		{"log10_true", "LOG10(TRUE)", 0, 0},

		// String coercion
		{"log10_string_10", `LOG10("10")`, 1, 0},
		{"log10_string_100", `LOG10("100")`, 2, 0},
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
		// LOG10(0) -> #NUM!
		{"log10_zero", "LOG10(0)", ErrValNUM},

		// Negative values -> #NUM!
		{"negative_1", "LOG10(-1)", ErrValNUM},
		{"negative_10", "LOG10(-10)", ErrValNUM},
		{"negative_small", "LOG10(-0.001)", ErrValNUM},

		// Boolean coercion: FALSE=0 -> #NUM!
		{"bool_false", "LOG10(FALSE)", ErrValNUM},

		// No args -> #VALUE!
		{"no_args", "LOG10()", ErrValVALUE},

		// Too many args -> #VALUE!
		{"too_many_args", "LOG10(10,2)", ErrValVALUE},

		// Non-numeric string -> #VALUE!
		{"non_numeric_string", `LOG10("abc")`, ErrValVALUE},

		// Error propagation
		{"error_prop_div0", "LOG10(1/0)", ErrValDIV0},
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

func TestACOS(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Key identities
		{"acos_1", "ACOS(1)", 0, 0},
		{"acos_0", "ACOS(0)", math.Pi / 2, 1e-10},
		{"acos_neg1", "ACOS(-1)", math.Pi, 1e-10},

		// Well-known angles
		{"acos_0.5_pi_over_3", "ACOS(0.5)", math.Pi / 3, 1e-10},
		{"acos_neg0.5_2pi_over_3", "ACOS(-0.5)", 2 * math.Pi / 3, 1e-10},
		{"acos_sqrt2_over_2_pi_over_4", "ACOS(SQRT(2)/2)", math.Pi / 4, 1e-10},
		{"acos_sqrt3_over_2_pi_over_6", "ACOS(SQRT(3)/2)", math.Pi / 6, 1e-10},
		{"acos_neg_sqrt2_over_2_3pi_over_4", "ACOS(-SQRT(2)/2)", 3 * math.Pi / 4, 1e-10},
		{"acos_neg_sqrt3_over_2_5pi_over_6", "ACOS(-SQRT(3)/2)", 5 * math.Pi / 6, 1e-10},

		// Values between -1 and 1
		{"acos_0.25", "ACOS(0.25)", math.Acos(0.25), 1e-10},
		{"acos_0.75", "ACOS(0.75)", math.Acos(0.75), 1e-10},
		{"acos_neg0.25", "ACOS(-0.25)", math.Acos(-0.25), 1e-10},
		{"acos_neg0.75", "ACOS(-0.75)", math.Acos(-0.75), 1e-10},
		{"acos_near_1", "ACOS(0.99)", math.Acos(0.99), 1e-10},
		{"acos_near_neg1", "ACOS(-0.99)", math.Acos(-0.99), 1e-10},

		// Boolean coercion: TRUE->1 (ACOS=0), FALSE->0 (ACOS=PI/2)
		{"acos_bool_true", "ACOS(TRUE)", 0, 0},
		{"acos_bool_false", "ACOS(FALSE)", math.Pi / 2, 1e-10},

		// String coercion
		{"acos_string_num", `ACOS("0.5")`, math.Pi / 3, 1e-10},
		{"acos_string_neg", `ACOS("-0.5")`, 2 * math.Pi / 3, 1e-10},
		{"acos_string_zero", `ACOS("0")`, math.Pi / 2, 1e-10},
		{"acos_string_one", `ACOS("1")`, 0, 0},

		// Doc example: ACOS(-0.5) = 2.094395102
		{"doc_example_neg0.5", "ACOS(-0.5)", 2.094395102, 1e-6},
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
		// Out of range
		{"out_of_range_high", "ACOS(1.1)", ErrValNUM},
		{"out_of_range_low", "ACOS(-1.1)", ErrValNUM},
		{"out_of_range_2", "ACOS(2)", ErrValNUM},
		{"out_of_range_neg2", "ACOS(-2)", ErrValNUM},
		{"out_of_range_large", "ACOS(100)", ErrValNUM},
		{"out_of_range_large_neg", "ACOS(-100)", ErrValNUM},

		// Wrong number of arguments
		{"no_args", "ACOS()", ErrValVALUE},
		{"too_many_args", "ACOS(1,2)", ErrValVALUE},

		// Non-numeric string
		{"string_non_num", `ACOS("abc")`, ErrValVALUE},

		// Error propagation
		{"error_propagation_na", "ACOS(#N/A)", ErrValNA},
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

func TestACOSH(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Doc examples
		{"doc_example_1", "ACOSH(1)", 0, 0},
		{"doc_example_10", "ACOSH(10)", 2.9932228461263808, 1e-10},

		// Domain boundary: ACOSH(1) = 0 exactly
		{"boundary_1", "ACOSH(1)", 0, 0},

		// Common values
		{"acosh_2", "ACOSH(2)", 1.3169578969248166, 1e-10},
		{"acosh_1_5", "ACOSH(1.5)", math.Acosh(1.5), 1e-10},
		{"acosh_3", "ACOSH(3)", math.Acosh(3), 1e-10},
		{"acosh_5", "ACOSH(5)", math.Acosh(5), 1e-10},

		// Large values
		{"acosh_100", "ACOSH(100)", math.Acosh(100), 1e-10},
		{"acosh_1000", "ACOSH(1000)", math.Acosh(1000), 1e-8},
		{"acosh_1e6", "ACOSH(1000000)", math.Acosh(1e6), 1e-6},

		// Value just above domain boundary
		{"just_above_1", "ACOSH(1.0001)", math.Acosh(1.0001), 1e-10},
		{"just_above_1_tiny", "ACOSH(1.00000001)", math.Acosh(1.00000001), 1e-10},

		// Expression input
		{"expr_add", "ACOSH(1+1)", math.Acosh(2), 1e-10},
		{"expr_mul", "ACOSH(2*3)", math.Acosh(6), 1e-10},

		// Identity: ACOSH(COSH(x)) = x for x >= 0
		{"identity_cosh_1", "ACOSH(COSH(1))", 1, 1e-10},
		{"identity_cosh_2", "ACOSH(COSH(2))", 2, 1e-10},

		// Boolean coercion: TRUE = 1 -> ACOSH(1) = 0
		{"bool_true", "ACOSH(TRUE)", 0, 0},

		// String coercion with numeric strings
		{"str_2", `ACOSH("2")`, math.Acosh(2), 1e-10},
		{"str_10", `ACOSH("10")`, math.Acosh(10), 1e-10},
		{"str_1", `ACOSH("1")`, 0, 0},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// Domain errors: ACOSH requires number >= 1
		{"below_domain_0", "ACOSH(0)", ErrValNUM},
		{"below_domain_neg1", "ACOSH(-1)", ErrValNUM},
		{"below_domain_neg100", "ACOSH(-100)", ErrValNUM},
		{"below_domain_0_5", "ACOSH(0.5)", ErrValNUM},
		{"below_domain_0_999", "ACOSH(0.999)", ErrValNUM},
		// Boolean coercion: FALSE = 0 -> below domain
		{"bool_false", "ACOSH(FALSE)", ErrValNUM},
		// String below domain
		{"str_below_domain", `ACOSH("0.5")`, ErrValNUM},
		// No args
		{"no_args", "ACOSH()", ErrValVALUE},
		// Too many args
		{"too_many_args", "ACOSH(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", `ACOSH("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "ACOSH(1/0)", ErrValDIV0},
		{"err_na", "ACOSH(NA())", ErrValNA},
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

func TestARABIC(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Single roman numerals
		{"I", `ARABIC("I")`, 1},
		{"V", `ARABIC("V")`, 5},
		{"X", `ARABIC("X")`, 10},
		{"L", `ARABIC("L")`, 50},
		{"C", `ARABIC("C")`, 100},
		{"D", `ARABIC("D")`, 500},
		{"M", `ARABIC("M")`, 1000},

		// Subtractive combinations
		{"IV", `ARABIC("IV")`, 4},
		{"IX", `ARABIC("IX")`, 9},
		{"XL", `ARABIC("XL")`, 40},
		{"XC", `ARABIC("XC")`, 90},
		{"CD", `ARABIC("CD")`, 400},
		{"CM", `ARABIC("CM")`, 900},

		// Multi-character compound numerals
		{"MCMXCIX", `ARABIC("MCMXCIX")`, 1999},
		{"MMXXVI", `ARABIC("MMXXVI")`, 2026},
		{"MMMDCCCLXXXVIII", `ARABIC("MMMDCCCLXXXVIII")`, 3888},
		{"XLII", `ARABIC("XLII")`, 42},
		{"CDXLIV", `ARABIC("CDXLIV")`, 444},
		{"DCCCXC", `ARABIC("DCCCXC")`, 890},
		{"XIV", `ARABIC("XIV")`, 14},

		// Empty string and empty cell
		{"empty_string", `ARABIC("")`, 0},
		{"empty_cell", "ARABIC(B1)", 0},

		// Negative roman numerals
		{"negative_IV", `ARABIC("-IV")`, -4},
		{"negative_X", `ARABIC("-X")`, -10},
		{"negative_MCMXCIX", `ARABIC("-MCMXCIX")`, -1999},

		// Case insensitive: lowercase
		{"lowercase_iv", `ARABIC("iv")`, 4},
		{"lowercase_mcmxcix", `ARABIC("mcmxcix")`, 1999},

		// Case insensitive: mixed case
		{"mixed_case", `ARABIC("McmXcIx")`, 1999},

		// Whitespace trimming
		{"whitespace_both", `ARABIC("  X  ")`, 10},
		{"leading_space", `ARABIC(" IV")`, 4},
		{"trailing_space", `ARABIC("IV ")`, 4},
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
		// Wrong argument count
		{"no_args", "ARABIC()", ErrValVALUE},
		{"too_many_args", `ARABIC("I","V")`, ErrValVALUE},
		// Invalid roman numeral characters
		{"invalid_char_A", `ARABIC("ABC")`, ErrValVALUE},
		{"invalid_char_Z", `ARABIC("Z")`, ErrValVALUE},
		{"invalid_mixed", `ARABIC("XIZ")`, ErrValVALUE},
		// Non-string argument types
		{"number_input", "ARABIC(123)", ErrValVALUE},
		{"bool_input", "ARABIC(TRUE)", ErrValVALUE},
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

func TestASIN(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Exact / well-known values
		{"asin_0", "ASIN(0)", 0, 0},
		{"asin_1", "ASIN(1)", math.Pi / 2, 1e-10},
		{"asin_neg1", "ASIN(-1)", -math.Pi / 2, 1e-10},
		{"asin_0.5", "ASIN(0.5)", math.Pi / 6, 1e-10},
		{"asin_neg0.5", "ASIN(-0.5)", -math.Pi / 6, 1e-10},
		// Additional values between -1 and 1
		{"asin_0.25", "ASIN(0.25)", math.Asin(0.25), 1e-10},
		{"asin_neg0.25", "ASIN(-0.25)", math.Asin(-0.25), 1e-10},
		{"asin_0.75", "ASIN(0.75)", math.Asin(0.75), 1e-10},
		{"asin_neg0.75", "ASIN(-0.75)", math.Asin(-0.75), 1e-10},
		{"asin_sqrt2_over2", "ASIN(SQRT(2)/2)", math.Pi / 4, 1e-10},
		{"asin_small", "ASIN(0.01)", math.Asin(0.01), 1e-10},
		{"asin_near1", "ASIN(0.999)", math.Asin(0.999), 1e-10},
		{"asin_near_neg1", "ASIN(-0.999)", math.Asin(-0.999), 1e-10},
		// Boolean coercion
		{"asin_bool_true", "ASIN(TRUE)", math.Pi / 2, 1e-10},
		{"asin_bool_false", "ASIN(FALSE)", 0, 0},
		// String coercion (numeric strings)
		{"asin_string_0.5", `ASIN("0.5")`, math.Pi / 6, 1e-10},
		{"asin_string_1", `ASIN("1")`, math.Pi / 2, 1e-10},
		{"asin_string_neg1", `ASIN("-1")`, -math.Pi / 2, 1e-10},
		{"asin_string_0", `ASIN("0")`, 0, 0},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// Out of range
		{"out_of_range_2", "ASIN(2)", ErrValNUM},
		{"out_of_range_neg2", "ASIN(-2)", ErrValNUM},
		{"out_of_range_high", "ASIN(1.1)", ErrValNUM},
		{"out_of_range_low", "ASIN(-1.1)", ErrValNUM},
		{"out_of_range_large", "ASIN(100)", ErrValNUM},
		{"out_of_range_neg_large", "ASIN(-100)", ErrValNUM},
		// Wrong arity
		{"no_args", "ASIN()", ErrValVALUE},
		{"too_many_args", "ASIN(1,2)", ErrValVALUE},
		// Non-numeric string
		{"string_non_num", `ASIN("abc")`, ErrValVALUE},
		{"string_empty", `ASIN("")`, ErrValVALUE},
		// Error propagation
		{"error_propagation_div0", "ASIN(1/0)", ErrValDIV0},
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

func TestASINH(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Identity: ASINH(0) = 0
		{"asinh_0", "ASINH(0)", 0, 0},

		// Well-known values
		{"asinh_1", "ASINH(1)", math.Asinh(1), 1e-10},
		{"asinh_neg1", "ASINH(-1)", math.Asinh(-1), 1e-10},

		// Doc examples
		{"doc_neg2.5", "ASINH(-2.5)", -1.6472311463710958, 1e-10},
		{"doc_10", "ASINH(10)", 2.99822295029797, 1e-10},

		// Odd function property: ASINH(-x) = -ASINH(x)
		{"odd_2", "ASINH(2)+ASINH(-2)", 0, 1e-10},
		{"odd_5", "ASINH(5)+ASINH(-5)", 0, 1e-10},

		// Fractional inputs
		{"asinh_0.5", "ASINH(0.5)", math.Asinh(0.5), 1e-10},
		{"asinh_neg0.5", "ASINH(-0.5)", math.Asinh(-0.5), 1e-10},
		{"asinh_0.25", "ASINH(0.25)", math.Asinh(0.25), 1e-10},

		// Large values
		{"asinh_100", "ASINH(100)", math.Asinh(100), 1e-10},
		{"asinh_neg100", "ASINH(-100)", math.Asinh(-100), 1e-10},
		{"asinh_1000", "ASINH(1000)", math.Asinh(1000), 1e-10},

		// Small values near zero (ASINH(x) ~ x for small x)
		{"asinh_small", "ASINH(0.001)", math.Asinh(0.001), 1e-15},
		{"asinh_tiny", "ASINH(0.0000001)", math.Asinh(0.0000001), 1e-18},
		{"asinh_neg_small", "ASINH(-0.001)", math.Asinh(-0.001), 1e-15},

		// Boolean coercion: TRUE=1, FALSE=0
		{"bool_true", "ASINH(TRUE)", math.Asinh(1), 1e-10},
		{"bool_false", "ASINH(FALSE)", 0, 0},

		// String coercion (numeric strings)
		{"string_2", `ASINH("2")`, math.Asinh(2), 1e-10},
		{"string_neg3", `ASINH("-3")`, math.Asinh(-3), 1e-10},
		{"string_0", `ASINH("0")`, 0, 0},
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
					t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
				}
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// Wrong arity
		{"no_args", "ASINH()", ErrValVALUE},
		{"too_many_args", "ASINH(1,2)", ErrValVALUE},
		// Non-numeric string
		{"string_non_num", `ASINH("abc")`, ErrValVALUE},
		{"string_empty", `ASINH("")`, ErrValVALUE},
		// Error propagation
		{"error_propagation_div0", "ASINH(1/0)", ErrValDIV0},
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

func TestATAN(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Identity: ATAN(0) = 0
		{"atan_0", "ATAN(0)", 0, 0},

		// Fundamental values: ATAN(1) = PI/4, ATAN(-1) = -PI/4
		{"atan_1", "ATAN(1)", math.Pi / 4, 1e-10},
		{"atan_neg1", "ATAN(-1)", -math.Pi / 4, 1e-10},

		// Doc example: ATAN(1) = 0.785398163...
		{"doc_ex1", "ATAN(1)", 0.785398163, 1e-9},

		// Fractional inputs
		{"atan_0.5", "ATAN(0.5)", math.Atan(0.5), 1e-10},
		{"atan_neg0.5", "ATAN(-0.5)", math.Atan(-0.5), 1e-10},
		{"atan_2", "ATAN(2)", math.Atan(2), 1e-10},
		{"atan_neg2", "ATAN(-2)", math.Atan(-2), 1e-10},

		// Large values approaching PI/2
		{"atan_1000", "ATAN(1000)", math.Atan(1000), 1e-10},
		{"atan_neg1000", "ATAN(-1000)", math.Atan(-1000), 1e-10},
		{"atan_1e10", "ATAN(10000000000)", math.Atan(1e10), 1e-10},
		{"atan_neg1e10", "ATAN(-10000000000)", math.Atan(-1e10), 1e-10},

		// Small values near zero (ATAN(x) ~ x for small x)
		{"atan_small", "ATAN(0.001)", math.Atan(0.001), 1e-10},
		{"atan_tiny", "ATAN(0.0000001)", math.Atan(0.0000001), 1e-15},
		{"atan_neg_small", "ATAN(-0.001)", math.Atan(-0.001), 1e-10},

		// Boolean coercion: TRUE=1, FALSE=0
		{"bool_true", "ATAN(TRUE)", math.Pi / 4, 1e-10},
		{"bool_false", "ATAN(FALSE)", 0, 0},

		// String coercion
		{"string_1", `ATAN("1")`, math.Pi / 4, 1e-10},
		{"string_0", `ATAN("0")`, 0, 0},
		{"string_neg1", `ATAN("-1")`, -math.Pi / 4, 1e-10},

		// Expression argument
		{"expr_add", "ATAN(0.5+0.5)", math.Pi / 4, 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No arguments
		{"no_args", "ATAN()", ErrValVALUE},
		// Too many arguments
		{"too_many_args", "ATAN(1,2)", ErrValVALUE},
		// Non-numeric string
		{"string_non_num", `ATAN("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "ATAN(1/0)", ErrValDIV0},
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

func TestATAN2(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Quadrant I: x>0, y>0 => result in (0, PI/2)
		{"q1_1_1", "ATAN2(1,1)", math.Pi / 4, 1e-10},
		{"q1_3_4", "ATAN2(3,4)", math.Atan2(4, 3), 1e-10},
		{"q1_large", "ATAN2(1000,2000)", math.Atan2(2000, 1000), 1e-10},
		// Quadrant II: x<0, y>0 => result in (PI/2, PI)
		{"q2_neg1_1", "ATAN2(-1,1)", math.Atan2(1, -1), 1e-10},
		{"q2_neg3_4", "ATAN2(-3,4)", math.Atan2(4, -3), 1e-10},
		// Quadrant III: x<0, y<0 => result in (-PI, -PI/2)
		{"q3_neg1_neg1", "ATAN2(-1,-1)", math.Atan2(-1, -1), 1e-10},
		{"q3_neg3_neg4", "ATAN2(-3,-4)", math.Atan2(-4, -3), 1e-10},
		// Quadrant IV: x>0, y<0 => result in (-PI/2, 0)
		{"q4_1_neg1", "ATAN2(1,-1)", math.Atan2(-1, 1), 1e-10},
		{"q4_3_neg4", "ATAN2(3,-4)", math.Atan2(-4, 3), 1e-10},
		// Axis-aligned: positive x-axis (y=0, x>0) => 0
		{"pos_x_axis", "ATAN2(1,0)", 0, 0},
		// Axis-aligned: negative x-axis (y=0, x<0) => PI
		{"neg_x_axis", "ATAN2(-1,0)", math.Pi, 1e-10},
		// Axis-aligned: positive y-axis (x=0, y>0) => PI/2
		{"pos_y_axis", "ATAN2(0,1)", math.Pi / 2, 1e-10},
		// Axis-aligned: negative y-axis (x=0, y<0) => -PI/2
		{"neg_y_axis", "ATAN2(0,-1)", -math.Pi / 2, 1e-10},
		// Very small values
		{"small_values", "ATAN2(0.0001,0.0001)", math.Pi / 4, 1e-10},
		// Boolean coercion
		{"bool_true_true", "ATAN2(TRUE,TRUE)", math.Pi / 4, 1e-10},
		{"bool_false_pos", "ATAN2(FALSE,1)", math.Pi / 2, 1e-10},
		{"bool_pos_false", "ATAN2(1,FALSE)", 0, 0},
		// String numeric coercion
		{"string_num_both", `ATAN2("1","1")`, math.Pi / 4, 1e-10},
		{"string_num_x", `ATAN2("3",4)`, math.Atan2(4, 3), 1e-10},
		{"string_num_y", `ATAN2(3,"4")`, math.Atan2(4, 3), 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// Both zero => #DIV/0!
		{"both_zero", "ATAN2(0,0)", ErrValDIV0},
		{"both_zero_float", "ATAN2(0.0,0.0)", ErrValDIV0},
		{"both_false", "ATAN2(FALSE,FALSE)", ErrValDIV0},
		// Wrong argument count
		{"no_args", "ATAN2()", ErrValVALUE},
		{"one_arg", "ATAN2(1)", ErrValVALUE},
		{"too_many_args", "ATAN2(1,2,3)", ErrValVALUE},
		// Non-numeric strings
		{"string_non_num_x", `ATAN2("abc",1)`, ErrValVALUE},
		{"string_non_num_y", `ATAN2(1,"abc")`, ErrValVALUE},
		// Error propagation
		{"err_propagate_x", "ATAN2(1/0,1)", ErrValDIV0},
		{"err_propagate_y", "ATAN2(1,1/0)", ErrValDIV0},
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

func TestATANH(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Identity: ATANH(0) = 0
		{"atanh_0", "ATANH(0)", 0, 0},

		// Fundamental values
		{"atanh_0.5", "ATANH(0.5)", 0.5493061443340549, 1e-10},
		{"atanh_neg0.5", "ATANH(-0.5)", -0.5493061443340549, 1e-10},

		// Doc examples
		{"doc_ex1", "ATANH(0.76159416)", 1.00000001, 1e-8},
		{"doc_ex2", "ATANH(-0.1)", -0.100335348, 1e-9},

		// Values near domain boundaries
		{"atanh_0.99", "ATANH(0.99)", math.Atanh(0.99), 1e-10},
		{"atanh_neg0.99", "ATANH(-0.99)", math.Atanh(-0.99), 1e-10},
		{"atanh_0.999", "ATANH(0.999)", math.Atanh(0.999), 1e-10},
		{"atanh_neg0.999", "ATANH(-0.999)", math.Atanh(-0.999), 1e-10},
		{"atanh_0.9999999", "ATANH(0.9999999)", math.Atanh(0.9999999), 1e-10},

		// Small values near zero (ATANH(x) ~ x for small x)
		{"atanh_small", "ATANH(0.001)", math.Atanh(0.001), 1e-10},
		{"atanh_tiny", "ATANH(0.0000001)", math.Atanh(0.0000001), 1e-15},
		{"atanh_neg_small", "ATANH(-0.001)", math.Atanh(-0.001), 1e-10},

		// Boolean coercion: FALSE=0 => ATANH(0)=0
		{"bool_false", "ATANH(FALSE)", 0, 0},

		// String coercion with valid numeric strings
		{"string_0.5", `ATANH("0.5")`, math.Atanh(0.5), 1e-10},
		{"string_0", `ATANH("0")`, 0, 0},
		{"string_neg0.5", `ATANH("-0.5")`, math.Atanh(-0.5), 1e-10},

		// Expression argument
		{"expr_add", "ATANH(0.25+0.25)", math.Atanh(0.5), 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// Domain boundary: ATANH(1) and ATANH(-1) are undefined => #NUM!
		{"at_1", "ATANH(1)", ErrValNUM},
		{"at_neg1", "ATANH(-1)", ErrValNUM},

		// Outside domain: |x| > 1 => #NUM!
		{"above_1", "ATANH(1.5)", ErrValNUM},
		{"below_neg1", "ATANH(-1.5)", ErrValNUM},
		{"far_above", "ATANH(100)", ErrValNUM},
		{"far_below", "ATANH(-100)", ErrValNUM},

		// Boolean coercion: TRUE=1 => #NUM! (at domain boundary)
		{"bool_true", "ATANH(TRUE)", ErrValNUM},

		// String that coerces to out-of-domain value
		{"string_1", `ATANH("1")`, ErrValNUM},
		{"string_neg1", `ATANH("-1")`, ErrValNUM},

		// No arguments
		{"no_args", "ATANH()", ErrValVALUE},

		// Too many arguments
		{"too_many_args", "ATANH(0.5,0.5)", ErrValVALUE},

		// Non-numeric string
		{"string_non_num", `ATANH("abc")`, ErrValVALUE},

		// Error propagation
		{"err_div0", "ATANH(1/0)", ErrValDIV0},
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

func TestFACT(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Basic factorials
		{"zero", "FACT(0)", 1},
		{"one", "FACT(1)", 1},
		{"two", "FACT(2)", 2},
		{"three", "FACT(3)", 6},
		{"four", "FACT(4)", 24},
		{"five", "FACT(5)", 120},
		{"six", "FACT(6)", 720},
		{"seven", "FACT(7)", 5040},
		{"ten", "FACT(10)", 3628800},
		{"twelve", "FACT(12)", 479001600},
		// Decimal inputs — truncated to integer (doc: FACT(1.9) = 1)
		{"decimal_1.9", "FACT(1.9)", 1},
		{"decimal_5.9", "FACT(5.9)", 120},
		{"decimal_0.5", "FACT(0.5)", 1},
		{"decimal_3.1", "FACT(3.1)", 6},
		{"decimal_0.999", "FACT(0.999)", 1},
		// Large factorials
		{"twenty", "FACT(20)", 2432902008176640000},
		{"fact_170", "FACT(170)", 7.257415615307994e+306},
		// Boolean coercion
		{"bool_true", "FACT(TRUE)", 1},
		{"bool_false", "FACT(FALSE)", 1},
		// Negative decimal that truncates to zero — FACT(0) = 1
		{"negative_small_decimal", "FACT(-0.1)", 1},
		// String number coercion
		{"string_5", `FACT("5")`, 120},
		{"string_0", `FACT("0")`, 1},
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
		// Negative numbers → #NUM! (doc: FACT(-1) = #NUM!)
		{"negative_one", "FACT(-1)", ErrValNUM},
		{"negative_five", "FACT(-5)", ErrValNUM},
		// FACT(-0.1): math.Trunc(-0.1) = -0, which is not < 0, so returns FACT(0) = 1
		// Tested in numTests as negative_small_decimal
		// Non-numeric string → #VALUE!
		{"non_numeric_string", `FACT("hello")`, ErrValVALUE},
		// No arguments → #VALUE!
		{"no_args", "FACT()", ErrValVALUE},
		// Too many arguments → #VALUE!
		{"too_many_args", "FACT(5,2)", ErrValVALUE},
		// Error propagation
		{"error_propagation_div0", "FACT(1/0)", ErrValDIV0},
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

func TestFACTDOUBLE(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Base cases
		{"zero", "FACTDOUBLE(0)", 1},
		{"one", "FACTDOUBLE(1)", 1},
		// Even numbers: n!! = n*(n-2)*(n-4)...(4)(2)
		{"two", "FACTDOUBLE(2)", 2},
		{"four", "FACTDOUBLE(4)", 8},    // 4*2
		{"six", "FACTDOUBLE(6)", 48},    // 6*4*2
		{"eight", "FACTDOUBLE(8)", 384}, // 8*6*4*2
		{"ten", "FACTDOUBLE(10)", 3840}, // 10*8*6*4*2
		// Odd numbers: n!! = n*(n-2)*(n-4)...(3)(1)
		{"three", "FACTDOUBLE(3)", 3},   // 3*1
		{"five", "FACTDOUBLE(5)", 15},   // 5*3*1
		{"seven", "FACTDOUBLE(7)", 105}, // 7*5*3*1
		{"nine", "FACTDOUBLE(9)", 945},  // 9*7*5*3*1
		// Larger values
		{"fifteen", "FACTDOUBLE(15)", 2027025},   // 15*13*11*9*7*5*3*1
		{"twenty", "FACTDOUBLE(20)", 3715891200}, // 20*18*16*14*12*10*8*6*4*2
		// FACTDOUBLE(-1) = 1 by convention
		{"negative_one", "FACTDOUBLE(-1)", 1},
		// Decimal inputs — truncated to integer (doc: non-integer is truncated)
		{"decimal_5.9", "FACTDOUBLE(5.9)", 15},
		{"decimal_6.1", "FACTDOUBLE(6.1)", 48},
		{"decimal_0.7", "FACTDOUBLE(0.7)", 1},
		{"decimal_1.999", "FACTDOUBLE(1.999)", 1},
		// Negative decimal that truncates to 0 → FACTDOUBLE(0) = 1
		{"negative_small_decimal", "FACTDOUBLE(-0.5)", 1},
		// Boolean coercion
		{"bool_true", "FACTDOUBLE(TRUE)", 1},
		{"bool_false", "FACTDOUBLE(FALSE)", 1},
		// String number coercion
		{"string_7", `FACTDOUBLE("7")`, 105},
		{"string_0", `FACTDOUBLE("0")`, 1},
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
		// Negative numbers < -1 → #NUM!
		{"negative_two", "FACTDOUBLE(-2)", ErrValNUM},
		{"negative_five", "FACTDOUBLE(-5)", ErrValNUM},
		// Negative decimal that truncates to < -1 → #NUM!
		{"negative_decimal", "FACTDOUBLE(-2.5)", ErrValNUM},
		// Non-numeric string → #VALUE!
		{"non_numeric_string", `FACTDOUBLE("hello")`, ErrValVALUE},
		// No arguments → #VALUE!
		{"no_args", "FACTDOUBLE()", ErrValVALUE},
		// Too many arguments → #VALUE!
		{"too_many_args", "FACTDOUBLE(5,2)", ErrValVALUE},
		// Error propagation
		{"error_propagation_div0", "FACTDOUBLE(1/0)", ErrValDIV0},
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

func TestEVEN(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Basic positive numbers
		{"pos_1", "EVEN(1)", 2},
		{"pos_2", "EVEN(2)", 2},
		{"pos_3", "EVEN(3)", 4},
		{"pos_5", "EVEN(5)", 6},

		// Already even positive numbers
		{"already_even_4", "EVEN(4)", 4},
		{"already_even_10", "EVEN(10)", 10},

		// Zero
		{"zero", "EVEN(0)", 0},

		// Negative numbers (rounds away from zero)
		{"neg_1", "EVEN(-1)", -2},
		{"neg_3", "EVEN(-3)", -4},
		{"neg_5", "EVEN(-5)", -6},

		// Already even negative numbers
		{"neg_even_2", "EVEN(-2)", -2},
		{"neg_even_4", "EVEN(-4)", -4},

		// Positive decimals (rounds up away from zero to nearest even)
		{"pos_decimal_1_5", "EVEN(1.5)", 2},
		{"pos_decimal_0_1", "EVEN(0.1)", 2},
		{"pos_decimal_2_1", "EVEN(2.1)", 4},
		{"pos_decimal_3_9", "EVEN(3.9)", 4},

		// Negative decimals (rounds away from zero to nearest even)
		{"neg_decimal_1_5", "EVEN(-1.5)", -2},
		{"neg_decimal_0_1", "EVEN(-0.1)", -2},
		{"neg_decimal_2_1", "EVEN(-2.1)", -4},

		// Large numbers
		{"large_odd", "EVEN(999999)", 1000000},
		{"large_even", "EVEN(1000000)", 1000000},
		{"large_neg", "EVEN(-999999)", -1000000},

		// String coercion
		{"string_pos", `EVEN("3")`, 4},
		{"string_neg", `EVEN("-1")`, -2},

		// Boolean coercion
		{"bool_true", "EVEN(TRUE)", 2},
		{"bool_false", "EVEN(FALSE)", 0},

		// Doc examples
		{"doc_ex1", "EVEN(1.5)", 2},
		{"doc_ex2", "EVEN(3)", 4},
		{"doc_ex3", "EVEN(2)", 2},
		{"doc_ex4", "EVEN(-1)", -2},
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
		{"no_args", "EVEN()", ErrValVALUE},
		// Too many args
		{"too_many_args", "EVEN(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", `EVEN("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "EVEN(1/0)", ErrValDIV0},
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

func TestODD(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Basic positive numbers
		{"pos_1", "ODD(1)", 1},
		{"pos_2", "ODD(2)", 3},
		{"pos_3", "ODD(3)", 3},
		{"pos_4", "ODD(4)", 5},
		{"pos_5", "ODD(5)", 5},
		{"pos_6", "ODD(6)", 7},

		// Already-odd positive numbers
		{"already_odd_7", "ODD(7)", 7},
		{"already_odd_9", "ODD(9)", 9},
		{"already_odd_11", "ODD(11)", 11},

		// Zero (returns 1)
		{"zero", "ODD(0)", 1},

		// Negative numbers (rounds away from zero)
		{"neg_1", "ODD(-1)", -1},
		{"neg_2", "ODD(-2)", -3},
		{"neg_3", "ODD(-3)", -3},
		{"neg_4", "ODD(-4)", -5},

		// Already-odd negative numbers
		{"neg_odd_5", "ODD(-5)", -5},
		{"neg_odd_7", "ODD(-7)", -7},

		// Positive decimals (rounds up away from zero to nearest odd)
		{"pos_decimal_1_5", "ODD(1.5)", 3},
		{"pos_decimal_0_1", "ODD(0.1)", 1},
		{"pos_decimal_2_1", "ODD(2.1)", 3},
		{"pos_decimal_3_9", "ODD(3.9)", 5},
		{"pos_decimal_4_5", "ODD(4.5)", 5},

		// Negative decimals (rounds away from zero to nearest odd)
		{"neg_decimal_1_5", "ODD(-1.5)", -3},
		{"neg_decimal_0_1", "ODD(-0.1)", -1},
		{"neg_decimal_2_1", "ODD(-2.1)", -3},

		// Large numbers
		{"large_even", "ODD(1000000)", 1000001},
		{"large_odd", "ODD(999999)", 999999},
		{"large_neg", "ODD(-1000000)", -1000001},

		// String coercion
		{"string_pos", `ODD("3")`, 3},
		{"string_neg", `ODD("-2")`, -3},

		// Boolean coercion
		{"bool_true", "ODD(TRUE)", 1},
		{"bool_false", "ODD(FALSE)", 1},

		// Doc examples
		{"doc_ex1", "ODD(1.5)", 3},
		{"doc_ex2", "ODD(3)", 3},
		{"doc_ex3", "ODD(2)", 3},
		{"doc_ex4", "ODD(-1)", -1},
		{"doc_ex5", "ODD(-2)", -3},
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
		{"no_args", "ODD()", ErrValVALUE},
		// Too many args
		{"too_many_args", "ODD(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", `ODD("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "ODD(1/0)", ErrValDIV0},
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

func TestEXP(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Basic cases
		{"exp_0", "EXP(0)", 1, 0},
		{"exp_1", "EXP(1)", math.E, 1e-10},
		{"exp_2", "EXP(2)", math.E * math.E, 1e-10},

		// Negative exponents
		{"exp_neg1", "EXP(-1)", 1 / math.E, 1e-10},
		{"exp_neg2", "EXP(-2)", 1 / (math.E * math.E), 1e-10},
		{"exp_neg0.5", "EXP(-0.5)", math.Exp(-0.5), 1e-10},

		// Decimal exponents
		{"exp_0.5", "EXP(0.5)", math.Sqrt(math.E), 1e-10},
		{"exp_0.1", "EXP(0.1)", math.Exp(0.1), 1e-10},
		{"exp_1.5", "EXP(1.5)", math.Exp(1.5), 1e-10},
		{"exp_3.14", "EXP(3.14)", math.Exp(3.14), 1e-10},

		// Large positive values
		{"exp_10", "EXP(10)", math.Exp(10), 1e-3},
		{"exp_20", "EXP(20)", math.Exp(20), 1e3},
		{"exp_50", "EXP(50)", math.Exp(50), 1e6},

		// Small negative values (approaching zero)
		{"exp_neg10", "EXP(-10)", math.Exp(-10), 1e-14},
		{"exp_neg20", "EXP(-20)", math.Exp(-20), 1e-18},

		// EXP(LN(x)) = x identity
		{"exp_ln_1", "EXP(LN(1))", 1, 1e-10},
		{"exp_ln_5", "EXP(LN(5))", 5, 1e-10},
		{"exp_ln_100", "EXP(LN(100))", 100, 1e-10},

		// Boolean coercion
		{"bool_true", "EXP(TRUE)", math.E, 1e-10},
		{"bool_false", "EXP(FALSE)", 1, 0},

		// String coercion
		{"string_0", `EXP("0")`, 1, 0},
		{"string_1", `EXP("1")`, math.E, 1e-10},
		{"string_neg1", `EXP("-1")`, 1 / math.E, 1e-10},

		// Expression argument
		{"expr_add", "EXP(1+1)", math.Exp(2), 1e-10},
		{"expr_sub", "EXP(3-1)", math.Exp(2), 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No args
		{"no_args", "EXP()", ErrValVALUE},
		// Too many args
		{"too_many_args", "EXP(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", `EXP("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "EXP(1/0)", ErrValDIV0},
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

func TestSIGN(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Basic: positive, negative, zero
		{"pos_int", "SIGN(5)", 1},
		{"neg_int", "SIGN(-5)", -1},
		{"zero", "SIGN(0)", 0},

		// Various positive numbers
		{"pos_one", "SIGN(1)", 1},
		{"pos_large_int", "SIGN(42)", 1},
		{"pos_hundred", "SIGN(100)", 1},

		// Various negative numbers
		{"neg_one", "SIGN(-1)", -1},
		{"neg_large_int", "SIGN(-42)", -1},
		{"neg_hundred", "SIGN(-100)", -1},

		// Decimal numbers
		{"pos_decimal", "SIGN(0.5)", 1},
		{"neg_decimal", "SIGN(-0.5)", -1},
		{"small_pos_decimal", "SIGN(0.001)", 1},
		{"small_neg_decimal", "SIGN(-0.001)", -1},

		// Large positive and negative numbers
		{"large_pos", "SIGN(999999999)", 1},
		{"large_neg", "SIGN(-999999999)", -1},

		// Boolean coercion
		{"bool_true", "SIGN(TRUE)", 1},
		{"bool_false", "SIGN(FALSE)", 0},

		// String coercion
		{"string_pos", `SIGN("5")`, 1},
		{"string_neg", `SIGN("-3")`, -1},
		{"string_zero", `SIGN("0")`, 0},

		// Expression argument
		{"expr_pos_result", "SIGN(5-2)", 1},
		{"expr_neg_result", "SIGN(2-5)", -1},
		{"expr_zero_result", "SIGN(4-4)", 0},

		// Doc examples
		{"doc_ex1", "SIGN(10)", 1},
		{"doc_ex2", "SIGN(4-4)", 0},
		{"doc_ex3", "SIGN(-0.00001)", -1},
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
		{"no_args", "SIGN()", ErrValVALUE},
		// Too many args
		{"too_many_args", "SIGN(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", `SIGN("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "SIGN(1/0)", ErrValDIV0},
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

func TestSIN(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Basic: SIN(0) = 0
		{"sin_0", "SIN(0)", 0, 0},

		// Standard angles (radians)
		{"sin_pi_6", "SIN(PI()/6)", 0.5, 1e-10},
		{"sin_pi_4", "SIN(PI()/4)", math.Sqrt2 / 2, 1e-10},
		{"sin_pi_3", "SIN(PI()/3)", math.Sqrt(3) / 2, 1e-10},
		{"sin_pi_2", "SIN(PI()/2)", 1, 1e-10},
		{"sin_pi", "SIN(PI())", 0, 1e-10},

		// Negative angles
		{"sin_neg_pi_6", "SIN(-PI()/6)", -0.5, 1e-10},
		{"sin_neg_pi_4", "SIN(-PI()/4)", -math.Sqrt2 / 2, 1e-10},
		{"sin_neg_pi_2", "SIN(-PI()/2)", -1, 1e-10},
		{"sin_neg_pi", "SIN(-PI())", 0, 1e-10},

		// Multiples of PI
		{"sin_2pi", "SIN(2*PI())", 0, 1e-10},
		{"sin_3pi_2", "SIN(3*PI()/2)", -1, 1e-10},
		{"sin_4pi", "SIN(4*PI())", 0, 1e-9},

		// Large angles
		{"sin_10pi", "SIN(10*PI())", 0, 1e-8},
		{"sin_100", "SIN(100)", math.Sin(100), 1e-10},
		{"sin_1000", "SIN(1000)", math.Sin(1000), 1e-10},

		// Small angle
		{"sin_small", "SIN(0.001)", math.Sin(0.001), 1e-10},

		// Degrees via conversion (doc examples)
		{"doc_ex1_pi", "SIN(PI())", 0, 1e-10},
		{"doc_ex2_pi_2", "SIN(PI()/2)", 1, 1e-10},
		{"doc_ex3_30deg", "SIN(30*PI()/180)", 0.5, 1e-10},

		// Boolean coercion
		{"bool_true", "SIN(TRUE)", math.Sin(1), 1e-10},
		{"bool_false", "SIN(FALSE)", 0, 0},

		// String coercion
		{"string_0", `SIN("0")`, 0, 0},
		{"string_1", `SIN("1")`, math.Sin(1), 1e-10},
		{"string_neg1", `SIN("-1")`, math.Sin(-1), 1e-10},

		// Expression argument
		{"expr_add", "SIN(PI()/4+PI()/4)", 1, 1e-10},
		{"expr_mul", "SIN(2*PI()/6)", math.Sqrt(3) / 2, 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No args
		{"no_args", "SIN()", ErrValVALUE},
		// Too many args
		{"too_many_args", "SIN(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", `SIN("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "SIN(1/0)", ErrValDIV0},
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

func TestCOMBINA(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Doc examples
		{"doc_ex1", "COMBINA(4,3)", 20},
		{"doc_ex2", "COMBINA(10,3)", 220},

		// k=0 always returns 1
		{"n0_k0", "COMBINA(0,0)", 1},
		{"n5_k0", "COMBINA(5,0)", 1},
		{"n1_k0", "COMBINA(1,0)", 1},
		{"n100_k0", "COMBINA(100,0)", 1},

		// k=1 returns n
		{"n1_k1", "COMBINA(1,1)", 1},
		{"n5_k1", "COMBINA(5,1)", 5},
		{"n10_k1", "COMBINA(10,1)", 10},

		// k=2: C(n+1, 2) = n*(n+1)/2
		{"n3_k2", "COMBINA(3,2)", 6},
		{"n6_k2", "COMBINA(6,2)", 21},

		// Larger values
		{"n10_k5", "COMBINA(10,5)", 2002},
		{"n20_k3", "COMBINA(20,3)", 1540},
		{"n15_k4", "COMBINA(15,4)", 3060},

		// Decimal truncation: non-integer values are truncated
		{"dec_n", "COMBINA(4.9,3)", 20},
		{"dec_k", "COMBINA(4,3.7)", 20},
		{"dec_both", "COMBINA(10.8,3.2)", 220},

		// Boolean coercion
		{"bool_true_n", "COMBINA(TRUE,0)", 1},
		{"bool_true_k", "COMBINA(4,TRUE)", 4},
		{"bool_false_k", "COMBINA(5,FALSE)", 1},

		// String coercion
		{"str_n", `COMBINA("4",3)`, 20},
		{"str_k", `COMBINA(10,"3")`, 220},
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
		// Negative n
		{"neg_n", "COMBINA(-1,3)", ErrValNUM},
		// Negative k
		{"neg_k", "COMBINA(4,-1)", ErrValNUM},
		// Both negative
		{"neg_both", "COMBINA(-2,-3)", ErrValNUM},
		// No args
		{"no_args", "COMBINA()", ErrValVALUE},
		// Too few args
		{"one_arg", "COMBINA(4)", ErrValVALUE},
		// Too many args
		{"too_many_args", "COMBINA(4,3,1)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric_n", `COMBINA("abc",3)`, ErrValVALUE},
		{"non_numeric_k", `COMBINA(4,"xyz")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "COMBINA(1/0,3)", ErrValDIV0},
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

func TestCOMBIN(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Doc example
		{"doc_example", "COMBIN(8,2)", 28},
		// Basic combinations
		{"5_choose_2", "COMBIN(5,2)", 10},
		{"10_choose_3", "COMBIN(10,3)", 120},
		{"6_choose_3", "COMBIN(6,3)", 20},
		// COMBIN(n,0) = 1 for any n
		{"4_choose_0", "COMBIN(4,0)", 1},
		{"0_choose_0", "COMBIN(0,0)", 1},
		{"10_choose_0", "COMBIN(10,0)", 1},
		// COMBIN(n,n) = 1 for any n
		{"4_choose_4", "COMBIN(4,4)", 1},
		{"1_choose_1", "COMBIN(1,1)", 1},
		{"7_choose_7", "COMBIN(7,7)", 1},
		// COMBIN(n,1) = n
		{"5_choose_1", "COMBIN(5,1)", 5},
		{"12_choose_1", "COMBIN(12,1)", 12},
		// Larger values
		{"20_choose_10", "COMBIN(20,10)", 184756},
		{"15_choose_5", "COMBIN(15,5)", 3003},
		{"100_choose_2", "COMBIN(100,2)", 4950},
		{"52_choose_5", "COMBIN(52,5)", 2598960},
		// Decimal truncation — arguments are truncated to integers
		{"decimal_n", "COMBIN(5.9,2)", 10},
		{"decimal_k", "COMBIN(5,2.7)", 10},
		{"decimal_both", "COMBIN(8.9,2.1)", 28},
		// String coercion
		{"string_n", `COMBIN("10",3)`, 120},
		{"string_k", `COMBIN(10,"3")`, 120},
		{"string_both", `COMBIN("5","2")`, 10},
		// Boolean coercion
		{"bool_true_n", "COMBIN(TRUE,0)", 1},
		{"bool_true_k", "COMBIN(5,TRUE)", 5},
		{"bool_false_k", "COMBIN(5,FALSE)", 1},
		{"bool_true_both", "COMBIN(TRUE,TRUE)", 1},
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
		// n < k → #NUM!
		{"n_less_than_k", "COMBIN(3,5)", ErrValNUM},
		{"n_less_than_k_2", "COMBIN(2,10)", ErrValNUM},
		// Negative n → #NUM!
		{"negative_n", "COMBIN(-1,2)", ErrValNUM},
		{"negative_n_large", "COMBIN(-5,3)", ErrValNUM},
		// Negative k → #NUM!
		{"negative_k", "COMBIN(5,-1)", ErrValNUM},
		{"negative_k_large", "COMBIN(10,-3)", ErrValNUM},
		// Both negative → #NUM!
		{"both_negative", "COMBIN(-3,-1)", ErrValNUM},
		// Wrong number of arguments → #VALUE!
		{"no_args", "COMBIN()", ErrValVALUE},
		{"one_arg", "COMBIN(5)", ErrValVALUE},
		{"three_args", "COMBIN(5,2,1)", ErrValVALUE},
		// Non-numeric string → #VALUE!
		{"non_numeric_n", `COMBIN("abc",2)`, ErrValVALUE},
		{"non_numeric_k", `COMBIN(5,"xyz")`, ErrValVALUE},
		// Error propagation
		{"error_div0_n", "COMBIN(1/0,2)", ErrValDIV0},
		{"error_div0_k", "COMBIN(5,1/0)", ErrValDIV0},
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

func TestCOS(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Doc example
		{"doc_example", "COS(1.047)", 0.5001710745970701, 1e-10},

		// Key angles
		{"cos_0", "COS(0)", 1, 0},
		{"cos_pi_6", "COS(PI()/6)", math.Sqrt(3) / 2, 1e-10},
		{"cos_pi_4", "COS(PI()/4)", math.Sqrt(2) / 2, 1e-10},
		{"cos_pi_3", "COS(PI()/3)", 0.5, 1e-10},
		{"cos_pi_2", "COS(PI()/2)", 0, 1e-10},
		{"cos_pi", "COS(PI())", -1, 1e-10},

		// Negative angles: COS(-x) = COS(x) (even function)
		{"cos_neg_pi_6", "COS(-PI()/6)", math.Sqrt(3) / 2, 1e-10},
		{"cos_neg_pi_4", "COS(-PI()/4)", math.Sqrt(2) / 2, 1e-10},
		{"cos_neg_pi_3", "COS(-PI()/3)", 0.5, 1e-10},
		{"cos_neg_pi", "COS(-PI())", -1, 1e-10},
		{"cos_neg_1", "COS(-1)", math.Cos(-1), 1e-10},

		// Multiples of PI
		{"cos_2pi", "COS(2*PI())", 1, 1e-10},
		{"cos_3pi_2", "COS(3*PI()/2)", 0, 1e-10},
		{"cos_4pi", "COS(4*PI())", 1, 1e-10},

		// Large angles
		{"cos_100", "COS(100)", math.Cos(100), 1e-10},
		{"cos_1000", "COS(1000)", math.Cos(1000), 1e-10},

		// Boolean coercion: TRUE=1, FALSE=0
		{"bool_true", "COS(TRUE)", math.Cos(1), 1e-10},
		{"bool_false", "COS(FALSE)", 1, 0},

		// String coercion
		{"str_zero", `COS("0")`, 1, 0},
		{"str_1_047", `COS("1.047")`, 0.5001710745970701, 1e-10},

		// Degree conversion via expression (60 degrees)
		{"degrees_60", "COS(60*PI()/180)", 0.5, 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No args
		{"no_args", "COS()", ErrValVALUE},
		// Too many args
		{"too_many_args", "COS(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", `COS("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "COS(1/0)", ErrValDIV0},
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

func TestCOSH(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Doc examples
		{"doc_example_4", "COSH(4)", 27.308232836016487, 1e-10},
		{"doc_example_e", "COSH(EXP(1))", 7.6101251386622884, 1e-10},

		// COSH(0) = 1
		{"cosh_0", "COSH(0)", 1, 0},

		// COSH(1) ≈ 1.5430806348152437
		{"cosh_1", "COSH(1)", math.Cosh(1), 1e-10},

		// COSH(-1) = COSH(1) — even function
		{"cosh_neg1", "COSH(-1)", math.Cosh(1), 1e-10},

		// Even function: COSH(-x) = COSH(x)
		{"even_2", "COSH(-2)", math.Cosh(2), 1e-10},
		{"even_3", "COSH(-3)", math.Cosh(3), 1e-10},
		{"even_5", "COSH(-5)", math.Cosh(5), 1e-10},

		// Larger values
		{"cosh_5", "COSH(5)", 74.20994852478785, 1e-8},
		{"cosh_10", "COSH(10)", 11013.232920103324, 1e-6},

		// Small values
		{"cosh_0_1", "COSH(0.1)", math.Cosh(0.1), 1e-10},
		{"cosh_0_01", "COSH(0.01)", math.Cosh(0.01), 1e-10},
		{"cosh_0_001", "COSH(0.001)", math.Cosh(0.001), 1e-14},
		{"cosh_neg_0_5", "COSH(-0.5)", math.Cosh(0.5), 1e-10},

		// Expression input
		{"cosh_expr", "COSH(2+1)", math.Cosh(3), 1e-10},
		{"cosh_ln_2", "COSH(LN(2))", 1.25, 1e-10},

		// Boolean coercion: TRUE=1, FALSE=0
		{"bool_true", "COSH(TRUE)", math.Cosh(1), 1e-10},
		{"bool_false", "COSH(FALSE)", 1, 0},

		// String coercion
		{"str_zero", `COSH("0")`, 1, 0},
		{"str_1", `COSH("1")`, math.Cosh(1), 1e-10},
		{"str_neg_1", `COSH("-1")`, math.Cosh(1), 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No args
		{"no_args", "COSH()", ErrValVALUE},
		// Too many args
		{"too_many_args", "COSH(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", `COSH("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "COSH(1/0)", ErrValDIV0},
		// Overflow to infinity
		{"overflow", "COSH(710)", ErrValNUM},
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

func TestCOT(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Key angles
		{"cot_pi_4", "COT(PI()/4)", 1, 1e-10},
		{"cot_pi_6", "COT(PI()/6)", math.Sqrt(3), 1e-10},
		{"cot_pi_3", "COT(PI()/3)", 1 / math.Sqrt(3), 1e-10},
		{"cot_pi_2", "COT(PI()/2)", 0, 1e-10},

		// COT(x) = 1/tan(x)
		{"cot_1", "COT(1)", 1 / math.Tan(1), 1e-10},
		{"cot_2", "COT(2)", 1 / math.Tan(2), 1e-10},
		{"cot_0_5", "COT(0.5)", 1 / math.Tan(0.5), 1e-10},

		// Negative angles: COT(-x) = -COT(x) (odd function)
		{"cot_neg_pi_4", "COT(-PI()/4)", -1, 1e-10},
		{"cot_neg_pi_6", "COT(-PI()/6)", -math.Sqrt(3), 1e-10},
		{"cot_neg_1", "COT(-1)", 1 / math.Tan(-1), 1e-10},
		{"cot_neg_2", "COT(-2)", 1 / math.Tan(-2), 1e-10},

		// Large angles
		{"cot_100", "COT(100)", 1 / math.Tan(100), 1e-10},
		{"cot_1000", "COT(1000)", 1 / math.Tan(1000), 1e-10},

		// Small angle close to zero
		{"cot_0_01", "COT(0.01)", 1 / math.Tan(0.01), 1e-6},

		// Boolean coercion: TRUE=1
		{"bool_true", "COT(TRUE)", 1 / math.Tan(1), 1e-10},

		// String coercion
		{"str_1", `COT("1")`, 1 / math.Tan(1), 1e-10},
		{"str_0_5", `COT("0.5")`, 1 / math.Tan(0.5), 1e-10},

		// Degree conversion via expression (45 degrees = PI/4)
		{"degrees_45", "COT(45*PI()/180)", 1, 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No args
		{"no_args", "COT()", ErrValVALUE},
		// Too many args
		{"too_many_args", "COT(1,2)", ErrValVALUE},
		// COT(0) = 1/tan(0) => division by zero
		{"zero_div0", "COT(0)", ErrValDIV0},
		// FALSE coerces to 0 => DIV/0!
		{"bool_false_div0", "COT(FALSE)", ErrValDIV0},
		// String "0" coerces to 0 => DIV/0!
		{"str_zero_div0", `COT("0")`, ErrValDIV0},
		// Non-numeric string
		{"non_numeric", `COT("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "COT(1/0)", ErrValDIV0},
		{"err_na", "COT(NA())", ErrValNA},
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

func TestTAN(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Basic values
		{"tan_0", "TAN(0)", 0, 0},
		{"tan_pi_over_4", "TAN(PI()/4)", 1, 1e-10},
		{"tan_pi_over_6", "TAN(PI()/6)", 1.0 / math.Sqrt(3), 1e-10},
		{"tan_pi_over_3", "TAN(PI()/3)", math.Sqrt(3), 1e-10},
		{"tan_pi", "TAN(PI())", 0, 1e-10},

		// Negative angles
		{"tan_neg_pi_over_4", "TAN(-PI()/4)", -1, 1e-10},
		{"tan_neg_pi_over_6", "TAN(-PI()/6)", -1.0 / math.Sqrt(3), 1e-10},
		{"tan_neg_pi_over_3", "TAN(-PI()/3)", -math.Sqrt(3), 1e-10},
		{"tan_neg1", "TAN(-1)", math.Tan(-1), 1e-10},

		// Large angles
		{"tan_large", "TAN(100)", math.Tan(100), 1e-10},
		{"tan_neg_large", "TAN(-100)", math.Tan(-100), 1e-10},

		// Fractional values
		{"tan_0.785", "TAN(0.785)", math.Tan(0.785), 1e-10},
		{"tan_0.5", "TAN(0.5)", math.Tan(0.5), 1e-10},
		{"tan_1", "TAN(1)", math.Tan(1), 1e-10},

		// Boolean coercion
		{"tan_bool_true", "TAN(TRUE)", math.Tan(1), 1e-10},
		{"tan_bool_false", "TAN(FALSE)", 0, 0},

		// String coercion
		{"tan_string_num", `TAN("0.785")`, math.Tan(0.785), 1e-10},
		{"tan_string_zero", `TAN("0")`, 0, 0},

		// Two pi periodicity
		{"tan_two_pi", "TAN(2*PI())", 0, 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"no_args", "TAN()", ErrValVALUE},
		{"too_many_args", "TAN(1,2)", ErrValVALUE},
		{"string_non_num", `TAN("abc")`, ErrValVALUE},
		{"error_prop_div0", "TAN(1/0)", ErrValDIV0},
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

func TestTANH(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Basic values (from docs)
		{"tanh_0", "TANH(0)", 0, 0},
		{"tanh_0.5", "TANH(0.5)", 0.462117157, 1e-8},
		{"tanh_1", "TANH(1)", 0.761594156, 1e-8},
		{"tanh_neg2", "TANH(-2)", -0.964027580, 1e-8},

		// Odd function symmetry: TANH(-x) = -TANH(x)
		{"tanh_neg1", "TANH(-1)", -0.761594156, 1e-8},
		{"tanh_neg0.5", "TANH(-0.5)", -0.462117157, 1e-8},

		// Small values (close to linear: tanh(x) ≈ x for small x)
		{"tanh_small_0.001", "TANH(0.001)", math.Tanh(0.001), 1e-12},
		{"tanh_small_0.0001", "TANH(0.0001)", math.Tanh(0.0001), 1e-14},
		{"tanh_small_neg", "TANH(-0.001)", math.Tanh(-0.001), 1e-12},

		// Large values approach ±1 asymptotically
		{"tanh_10", "TANH(10)", 1.0, 1e-8},
		{"tanh_neg10", "TANH(-10)", -1.0, 1e-8},
		{"tanh_20", "TANH(20)", 1.0, 1e-15},
		{"tanh_neg20", "TANH(-20)", -1.0, 1e-15},
		{"tanh_100", "TANH(100)", 1.0, 0},
		{"tanh_neg100", "TANH(-100)", -1.0, 0},

		// Moderate values
		{"tanh_2", "TANH(2)", math.Tanh(2), 1e-10},
		{"tanh_3", "TANH(3)", math.Tanh(3), 1e-10},
		{"tanh_5", "TANH(5)", math.Tanh(5), 1e-10},

		// Boolean coercion
		{"tanh_bool_true", "TANH(TRUE)", math.Tanh(1), 1e-10},
		{"tanh_bool_false", "TANH(FALSE)", 0, 0},

		// String coercion
		{"tanh_string_num", `TANH("0.5")`, math.Tanh(0.5), 1e-10},
		{"tanh_string_zero", `TANH("0")`, 0, 0},
		{"tanh_string_neg", `TANH("-2")`, math.Tanh(-2), 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		{"no_args", "TANH()", ErrValVALUE},
		{"too_many_args", "TANH(1,2)", ErrValVALUE},
		{"string_non_num", `TANH("abc")`, ErrValVALUE},
		{"error_prop_div0", "TANH(1/0)", ErrValDIV0},
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

func TestSINH(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Basic: SINH(0) = 0
		{"sinh_0", "SINH(0)", 0, 0},

		// Standard values
		{"sinh_1", "SINH(1)", math.Sinh(1), 1e-10},       // ~1.1752
		{"sinh_neg1", "SINH(-1)", math.Sinh(-1), 1e-10},  // ~-1.1752
		{"sinh_0_5", "SINH(0.5)", math.Sinh(0.5), 1e-10}, // ~0.5211
		{"sinh_neg0_5", "SINH(-0.5)", math.Sinh(-0.5), 1e-10},

		// Odd function: SINH(-x) = -SINH(x)
		{"odd_2", "SINH(-2)", -math.Sinh(2), 1e-10},
		{"odd_3", "SINH(-3)", -math.Sinh(3), 1e-10},

		// Larger values
		{"sinh_5", "SINH(5)", math.Sinh(5), 1e-10},   // ~74.2032
		{"sinh_10", "SINH(10)", math.Sinh(10), 1e-4}, // ~11013.2329

		// Small values (near-linear region: SINH(x) ~ x for small x)
		{"sinh_small", "SINH(0.001)", math.Sinh(0.001), 1e-14},
		{"sinh_tiny", "SINH(0.0001)", math.Sinh(0.0001), 1e-15},

		// Doc example: 2.868*SINH(0.0342*1.03)
		{"doc_example", "2.868*SINH(0.0342*1.03)", 2.868 * math.Sinh(0.0342*1.03), 1e-7},

		// Boolean coercion
		{"bool_true", "SINH(TRUE)", math.Sinh(1), 1e-10},
		{"bool_false", "SINH(FALSE)", 0, 0},

		// String coercion
		{"string_0", `SINH("0")`, 0, 0},
		{"string_1", `SINH("1")`, math.Sinh(1), 1e-10},
		{"string_neg1", `SINH("-1")`, math.Sinh(-1), 1e-10},

		// Expression argument
		{"expr_add", "SINH(1+1)", math.Sinh(2), 1e-10},
		{"expr_mul", "SINH(2*3)", math.Sinh(6), 1e-8},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No args
		{"no_args", "SINH()", ErrValVALUE},
		// Too many args
		{"too_many_args", "SINH(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", `SINH("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "SINH(1/0)", ErrValDIV0},
		// Overflow to infinity returns #NUM!
		{"overflow", "SINH(1000)", ErrValNUM},
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

func TestCOTH(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Standard values
		{"coth_1", "COTH(1)", 1.3130352854993315, 1e-10},
		{"coth_2", "COTH(2)", 1.0373147207275481, 1e-10},
		{"coth_3", "COTH(3)", 1.0049698233136892, 1e-10},
		{"coth_0_5", "COTH(0.5)", 2.163953413738653, 1e-10},

		// Odd function: COTH(-x) = -COTH(x)
		{"odd_neg1", "COTH(-1)", -1.3130352854993315, 1e-10},
		{"odd_neg2", "COTH(-2)", -1.0373147207275481, 1e-10},
		{"odd_neg0_5", "COTH(-0.5)", -2.163953413738653, 1e-10},
		{"odd_neg3", "COTH(-3)", -1.0049698233136892, 1e-10},

		// Large values approaching ±1
		{"large_10", "COTH(10)", 1.0000000041223073, 1e-12},
		{"large_20", "COTH(20)", 1.0, 1e-8},
		{"large_neg10", "COTH(-10)", -1.0000000041223073, 1e-12},
		{"large_neg20", "COTH(-20)", -1.0, 1e-8},

		// Small values (large magnitude result)
		{"small_0_1", "COTH(0.1)", 1 / math.Tanh(0.1), 1e-10},
		{"small_0_01", "COTH(0.01)", 1 / math.Tanh(0.01), 1e-8},

		// Boolean coercion: TRUE=1, FALSE handled in error tests (=0 -> DIV/0)
		{"bool_true", "COTH(TRUE)", 1.3130352854993315, 1e-10},

		// String coercion
		{"str_1", `COTH("1")`, 1.3130352854993315, 1e-10},
		{"str_neg1", `COTH("-1")`, -1.3130352854993315, 1e-10},
		{"str_2", `COTH("2")`, 1.0373147207275481, 1e-10},

		// Expression input
		{"expr_add", "COTH(1+1)", 1 / math.Tanh(2), 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No args
		{"no_args", "COTH()", ErrValVALUE},
		// Too many args
		{"too_many_args", "COTH(1,2)", ErrValVALUE},
		// COTH(0) is undefined (division by zero)
		{"zero", "COTH(0)", ErrValDIV0},
		// Boolean FALSE coerces to 0 -> DIV/0
		{"bool_false", "COTH(FALSE)", ErrValDIV0},
		// String "0" coerces to 0 -> DIV/0
		{"str_zero", `COTH("0")`, ErrValDIV0},
		// Non-numeric string
		{"non_numeric", `COTH("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "COTH(1/0)", ErrValDIV0},
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

func TestLN(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// LN(1) = 0
		{"ln_1", "LN(1)", 0, 0},

		// LN(e) = 1
		{"ln_e", "LN(2.7182818284590452)", 1, 1e-10},
		{"ln_e_via_exp", "LN(EXP(1))", 1, 1e-10},

		// LN(e^2) = 2
		{"ln_e_squared", "LN(EXP(2))", 2, 1e-10},

		// LN(EXP(x)) = x identity
		{"ln_exp_0", "LN(EXP(0))", 0, 1e-10},
		{"ln_exp_3", "LN(EXP(3))", 3, 1e-10},
		{"ln_exp_5", "LN(EXP(5))", 5, 1e-10},
		{"ln_exp_0_5", "LN(EXP(0.5))", 0.5, 1e-10},

		// Documentation examples
		{"doc_example_86", "LN(86)", 4.4543473, 1e-4},
		{"doc_example_e", "LN(2.7182818)", 1, 1e-4},
		{"doc_example_exp3", "LN(EXP(3))", 3, 0},

		// Common values
		{"ln_2", "LN(2)", 0.69314718055994530, 1e-10},
		{"ln_10", "LN(10)", 2.30258509299404568, 1e-10},

		// Large values
		{"ln_1000", "LN(1000)", 6.90775527898213705, 1e-10},
		{"ln_1e6", "LN(1000000)", 13.8155105579642741, 1e-10},
		{"ln_1e15", "LN(1E15)", 34.5387763949107352, 1e-6},

		// Small positive values (near zero)
		{"ln_0_5", "LN(0.5)", -0.69314718055994530, 1e-10},
		{"ln_0_1", "LN(0.1)", -2.30258509299404568, 1e-10},
		{"ln_0_001", "LN(0.001)", -6.90775527898213705, 1e-10},
		{"ln_1e-10", "LN(1E-10)", -23.0258509299404568, 1e-6},

		// Boolean coercion: TRUE=1 -> LN(1)=0
		{"ln_true", "LN(TRUE)", 0, 0},

		// String coercion
		{"ln_string_1", `LN("1")`, 0, 0},
		{"ln_string_10", `LN("10")`, 2.30258509299404568, 1e-10},
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
		// LN(0) -> #NUM!
		{"ln_zero", "LN(0)", ErrValNUM},

		// Negative values -> #NUM!
		{"ln_negative_1", "LN(-1)", ErrValNUM},
		{"ln_negative_10", "LN(-10)", ErrValNUM},
		{"ln_negative_small", "LN(-0.001)", ErrValNUM},

		// Boolean coercion: FALSE=0 -> #NUM!
		{"ln_false", "LN(FALSE)", ErrValNUM},

		// No args -> #VALUE!
		{"no_args", "LN()", ErrValVALUE},

		// Too many args -> #VALUE!
		{"too_many_args", "LN(1,2)", ErrValVALUE},

		// Non-numeric string -> #VALUE!
		{"non_numeric_string", `LN("abc")`, ErrValVALUE},

		// Error propagation
		{"error_prop_div0", "LN(1/0)", ErrValDIV0},
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

func TestCSC(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Key angles: CSC(x) = 1/SIN(x)
		{"csc_pi_2", "CSC(PI()/2)", 1, 1e-10},
		{"csc_pi_6", "CSC(PI()/6)", 2, 1e-10},
		{"csc_pi_4", "CSC(PI()/4)", math.Sqrt(2), 1e-10},
		{"csc_pi_3", "CSC(PI()/3)", 2 / math.Sqrt(3), 1e-10},

		// Generic values
		{"csc_1", "CSC(1)", 1 / math.Sin(1), 1e-10},
		{"csc_2", "CSC(2)", 1 / math.Sin(2), 1e-10},
		{"csc_0_5", "CSC(0.5)", 1 / math.Sin(0.5), 1e-10},

		// Negative angles: CSC(-x) = -CSC(x) (odd function)
		{"csc_neg_pi_2", "CSC(-PI()/2)", -1, 1e-10},
		{"csc_neg_pi_6", "CSC(-PI()/6)", -2, 1e-10},
		{"csc_neg_1", "CSC(-1)", 1 / math.Sin(-1), 1e-10},
		{"csc_neg_2", "CSC(-2)", 1 / math.Sin(-2), 1e-10},

		// Large angles
		{"csc_100", "CSC(100)", 1 / math.Sin(100), 1e-10},
		{"csc_1000", "CSC(1000)", 1 / math.Sin(1000), 1e-10},

		// Small angle close to zero (large magnitude result)
		{"csc_0_01", "CSC(0.01)", 1 / math.Sin(0.01), 1e-6},

		// Boolean coercion: TRUE=1
		{"bool_true", "CSC(TRUE)", 1 / math.Sin(1), 1e-10},

		// String coercion
		{"str_1", `CSC("1")`, 1 / math.Sin(1), 1e-10},
		{"str_0_5", `CSC("0.5")`, 1 / math.Sin(0.5), 1e-10},

		// Degree conversion via expression (30 degrees = PI/6)
		{"degrees_30", "CSC(30*PI()/180)", 2, 1e-10},

		// Example: CSC(15) ≈ 1.5377...
		{"doc_example", "CSC(15)", 1 / math.Sin(15), 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No args
		{"no_args", "CSC()", ErrValVALUE},
		// Too many args
		{"too_many_args", "CSC(1,2)", ErrValVALUE},
		// CSC(0) = 1/sin(0) => division by zero
		{"zero_div0", "CSC(0)", ErrValDIV0},
		// FALSE coerces to 0 => DIV/0!
		{"bool_false_div0", "CSC(FALSE)", ErrValDIV0},
		// String "0" coerces to 0 => DIV/0!
		{"str_zero_div0", `CSC("0")`, ErrValDIV0},
		// Non-numeric string
		{"non_numeric", `CSC("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "CSC(1/0)", ErrValDIV0},
		{"err_na", "CSC(NA())", ErrValNA},
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

func TestSEC(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Doc examples (angles in radians)
		{"doc_sec_45", "SEC(45)", 1.0 / math.Cos(45), 1e-5},
		{"doc_sec_30", "SEC(30)", 1.0 / math.Cos(30), 1e-5},

		// Key angles: SEC(x) = 1/COS(x)
		{"sec_0", "SEC(0)", 1, 0},
		{"sec_pi_6", "SEC(PI()/6)", 2 / math.Sqrt(3), 1e-10},
		{"sec_pi_4", "SEC(PI()/4)", math.Sqrt(2), 1e-10},
		{"sec_pi_3", "SEC(PI()/3)", 2, 1e-10},
		{"sec_pi", "SEC(PI())", -1, 1e-10},

		// Negative angles: SEC(-x) = SEC(x) (even function)
		{"sec_neg_pi_6", "SEC(-PI()/6)", 2 / math.Sqrt(3), 1e-10},
		{"sec_neg_pi_4", "SEC(-PI()/4)", math.Sqrt(2), 1e-10},
		{"sec_neg_pi_3", "SEC(-PI()/3)", 2, 1e-10},
		{"sec_neg_pi", "SEC(-PI())", -1, 1e-10},
		{"sec_neg_1", "SEC(-1)", 1.0 / math.Cos(-1), 1e-10},

		// Multiples of PI
		{"sec_2pi", "SEC(2*PI())", 1, 1e-10},
		{"sec_4pi", "SEC(4*PI())", 1, 1e-10},

		// Large angles
		{"sec_100", "SEC(100)", 1.0 / math.Cos(100), 1e-10},
		{"sec_1000", "SEC(1000)", 1.0 / math.Cos(1000), 1e-10},

		// Boolean coercion: TRUE=1, FALSE=0
		{"bool_true", "SEC(TRUE)", 1.0 / math.Cos(1), 1e-10},
		{"bool_false", "SEC(FALSE)", 1, 0},

		// String coercion
		{"str_zero", `SEC("0")`, 1, 0},
		{"str_1", `SEC("1")`, 1.0 / math.Cos(1), 1e-10},

		// Degree conversion via expression (60 degrees -> PI/3 radians)
		{"degrees_60", "SEC(60*PI()/180)", 2, 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No args
		{"no_args", "SEC()", ErrValVALUE},
		// Too many args
		{"too_many_args", "SEC(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", `SEC("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "SEC(1/0)", ErrValDIV0},
		{"err_na", "SEC(NA())", ErrValNA},
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

func TestSECH(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Doc examples
		{"doc_example_45", "SECH(45)", 5.725037161098787e-20, 1e-30},
		{"doc_example_30", "SECH(30)", 1.871524593768035e-13, 1e-23},

		// SECH(0) = 1
		{"sech_0", "SECH(0)", 1, 0},

		// SECH(1) ≈ 0.6481
		{"sech_1", "SECH(1)", 0.6480542736638855, 1e-10},

		// SECH(-1) = SECH(1) — even function
		{"sech_neg1", "SECH(-1)", 0.6480542736638855, 1e-10},

		// Even function: SECH(-x) = SECH(x)
		{"even_2", "SECH(-2)", 1 / math.Cosh(2), 1e-10},
		{"even_3", "SECH(-3)", 1 / math.Cosh(3), 1e-10},
		{"even_5", "SECH(-5)", 1 / math.Cosh(5), 1e-10},

		// Large values approach 0
		{"sech_5", "SECH(5)", 0.013475282221304556, 1e-10},
		{"sech_10", "SECH(10)", 9.079985933781724e-05, 1e-10},
		{"sech_100", "SECH(100)", 7.440151952041671e-44, 1e-54},
		{"sech_710", "SECH(710)", 0, 1e-300},

		// Small values near 1
		{"sech_0_1", "SECH(0.1)", 1 / math.Cosh(0.1), 1e-10},
		{"sech_0_01", "SECH(0.01)", 1 / math.Cosh(0.01), 1e-14},
		{"sech_0_001", "SECH(0.001)", 1 / math.Cosh(0.001), 1e-14},
		{"sech_neg_0_5", "SECH(-0.5)", 1 / math.Cosh(0.5), 1e-10},

		// Expression input
		{"sech_expr", "SECH(2+1)", 1 / math.Cosh(3), 1e-10},
		{"sech_ln_2", "SECH(LN(2))", 1 / math.Cosh(math.Log(2)), 1e-10},

		// Boolean coercion: TRUE=1, FALSE=0
		{"bool_true", "SECH(TRUE)", 1 / math.Cosh(1), 1e-10},
		{"bool_false", "SECH(FALSE)", 1, 0},

		// String coercion
		{"str_zero", `SECH("0")`, 1, 0},
		{"str_1", `SECH("1")`, 1 / math.Cosh(1), 1e-10},
		{"str_neg_1", `SECH("-1")`, 1 / math.Cosh(1), 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No args
		{"no_args", "SECH()", ErrValVALUE},
		// Too many args
		{"too_many_args", "SECH(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", `SECH("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "SECH(1/0)", ErrValDIV0},
		{"err_na", "SECH(NA())", ErrValNA},
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

func TestCSCH(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Standard values: CSCH(x) = 1/SINH(x)
		{"csch_1", "CSCH(1)", 0.8509181282393216, 1e-10},
		{"csch_2", "CSCH(2)", 0.2757205647717832, 1e-10},
		{"csch_3", "CSCH(3)", 0.09982156966382391, 1e-10},
		{"csch_0_5", "CSCH(0.5)", 1.9190347513349437, 1e-10},
		{"csch_1_5", "CSCH(1.5)", 0.46964244059522464, 1e-10},

		// Odd function: CSCH(-x) = -CSCH(x)
		{"odd_neg1", "CSCH(-1)", -0.8509181282393216, 1e-10},
		{"odd_neg2", "CSCH(-2)", -0.2757205647717832, 1e-10},
		{"odd_neg0_5", "CSCH(-0.5)", -1.9190347513349437, 1e-10},
		{"odd_neg3", "CSCH(-3)", -0.09982156966382391, 1e-10},

		// Large values approach 0
		{"large_10", "CSCH(10)", 1 / math.Sinh(10), 1e-12},
		{"large_20", "CSCH(20)", 1 / math.Sinh(20), 1e-12},
		{"large_neg10", "CSCH(-10)", 1 / math.Sinh(-10), 1e-12},
		{"large_neg20", "CSCH(-20)", 1 / math.Sinh(-20), 1e-12},

		// Small values (large magnitude result)
		{"small_0_1", "CSCH(0.1)", 1 / math.Sinh(0.1), 1e-10},
		{"small_0_01", "CSCH(0.01)", 1 / math.Sinh(0.01), 1e-8},

		// Boolean coercion: TRUE=1, FALSE handled in error tests (=0 -> DIV/0)
		{"bool_true", "CSCH(TRUE)", 0.8509181282393216, 1e-10},

		// String coercion
		{"str_1", `CSCH("1")`, 0.8509181282393216, 1e-10},
		{"str_neg1", `CSCH("-1")`, -0.8509181282393216, 1e-10},
		{"str_2", `CSCH("2")`, 0.2757205647717832, 1e-10},

		// Expression input
		{"expr_add", "CSCH(1+1)", 1 / math.Sinh(2), 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No args
		{"no_args", "CSCH()", ErrValVALUE},
		// Too many args
		{"too_many_args", "CSCH(1,2)", ErrValVALUE},
		// CSCH(0) is undefined (division by zero)
		{"zero", "CSCH(0)", ErrValDIV0},
		// Boolean FALSE coerces to 0 -> DIV/0
		{"bool_false", "CSCH(FALSE)", ErrValDIV0},
		// String "0" coerces to 0 -> DIV/0
		{"str_zero", `CSCH("0")`, ErrValDIV0},
		// Non-numeric string
		{"non_numeric", `CSCH("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "CSCH(1/0)", ErrValDIV0},
		{"err_na", "CSCH(NA())", ErrValNA},
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

func TestPI(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
		tol     float64
	}{
		// Basic value
		{"pi_value", "PI()", math.Pi, 0},
		{"pi_approx", "PI()", 3.14159265358979, 1e-14},

		// Division by constants
		{"pi_over_2", "PI()/2", math.Pi / 2, 0},
		{"pi_over_4", "PI()/4", math.Pi / 4, 0},
		{"pi_over_180", "PI()/180", math.Pi / 180, 1e-18},

		// Multiplication by constants
		{"two_pi", "2*PI()", 2 * math.Pi, 0},
		{"pi_times_2", "PI()*2", math.Pi * 2, 0},
		{"three_pi", "3*PI()", 3 * math.Pi, 0},

		// Arithmetic expressions
		{"pi_plus_1", "PI()+1", math.Pi + 1, 0},
		{"pi_minus_3", "PI()-3", math.Pi - 3, 1e-15},
		{"pi_squared", "PI()*PI()", math.Pi * math.Pi, 1e-14},
		{"pi_power_2", "POWER(PI(),2)", math.Pi * math.Pi, 1e-10},

		// Trigonometric identities using PI
		{"sin_pi", "SIN(PI())", 0, 1e-10},
		{"cos_pi", "COS(PI())", -1, 1e-10},
		{"sin_pi_over_2", "SIN(PI()/2)", 1, 1e-10},
		{"cos_pi_over_2", "COS(PI()/2)", 0, 1e-10},

		// Degrees conversion: 180 degrees = PI radians
		{"radians_180", "RADIANS(180)", math.Pi, 1e-10},
		{"degrees_pi", "DEGREES(PI())", 180, 1e-10},
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
		// PI takes no arguments
		{"one_arg_number", "PI(1)", ErrValVALUE},
		{"one_arg_string", "PI(\"a\")", ErrValVALUE},
		{"one_arg_bool", "PI(TRUE)", ErrValVALUE},
		{"two_args", "PI(1,2)", ErrValVALUE},
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

func TestRADIANS(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// Zero
		{"zero", "RADIANS(0)", 0, 0},

		// Standard degree-to-radian conversions
		{"30_deg", "RADIANS(30)", math.Pi / 6, 1e-10},
		{"45_deg", "RADIANS(45)", math.Pi / 4, 1e-10},
		{"60_deg", "RADIANS(60)", math.Pi / 3, 1e-10},
		{"90_deg", "RADIANS(90)", math.Pi / 2, 1e-10},
		{"120_deg", "RADIANS(120)", 2 * math.Pi / 3, 1e-10},
		{"180_deg", "RADIANS(180)", math.Pi, 1e-10},
		{"270_deg", "RADIANS(270)", 3 * math.Pi / 2, 1e-10},
		{"360_deg", "RADIANS(360)", 2 * math.Pi, 1e-10},

		// Negative angles
		{"neg_90", "RADIANS(-90)", -math.Pi / 2, 1e-10},
		{"neg_180", "RADIANS(-180)", -math.Pi, 1e-10},
		{"neg_360", "RADIANS(-360)", -2 * math.Pi, 1e-10},

		// Large values
		{"720_deg", "RADIANS(720)", 4 * math.Pi, 1e-10},
		{"large", "RADIANS(36000)", 200 * math.Pi, 1e-8},
		{"very_large", "RADIANS(1000000)", 1000000 * math.Pi / 180, 1e-6},

		// Decimal degrees
		{"decimal_45_5", "RADIANS(45.5)", 45.5 * math.Pi / 180, 1e-10},
		{"decimal_0_1", "RADIANS(0.1)", 0.1 * math.Pi / 180, 1e-15},

		// Boolean coercion
		{"bool_true", "RADIANS(TRUE)", math.Pi / 180, 1e-15},
		{"bool_false", "RADIANS(FALSE)", 0, 0},

		// String coercion
		{"string_90", `RADIANS("90")`, math.Pi / 2, 1e-10},
		{"string_neg_180", `RADIANS("-180")`, -math.Pi, 1e-10},

		// Expression argument
		{"expr_mul", "RADIANS(2*90)", math.Pi, 1e-10},
		{"expr_add", "RADIANS(45+45)", math.Pi / 2, 1e-10},

		// Doc example
		{"doc_ex1", "RADIANS(270)", 3 * math.Pi / 2, 1e-10},
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
			if math.Abs(got.Num-tt.want) > tt.tol {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// No args
		{"no_args", "RADIANS()", ErrValVALUE},
		// Too many args
		{"too_many_args", "RADIANS(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", `RADIANS("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "RADIANS(1/0)", ErrValDIV0},
		{"err_na", "RADIANS(NA())", ErrValNA},
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

func TestDEGREES(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
		tol     float64
	}{
		// Core conversions using PI()
		{"pi", "DEGREES(PI())", 180, 0},
		{"half_pi", "DEGREES(PI()/2)", 90, 1e-10},
		{"quarter_pi", "DEGREES(PI()/4)", 45, 1e-10},
		{"two_pi", "DEGREES(2*PI())", 360, 1e-10},
		{"three_halves_pi", "DEGREES(3*PI()/2)", 270, 1e-10},

		// Zero
		{"zero", "DEGREES(0)", 0, 0},

		// One radian ~ 57.29577951308232
		{"one_radian", "DEGREES(1)", 180 / math.Pi, 1e-10},

		// Negative values
		{"neg_pi", "DEGREES(-PI())", -180, 1e-10},
		{"neg_half_pi", "DEGREES(-PI()/2)", -90, 1e-10},
		{"neg_one", "DEGREES(-1)", -180 / math.Pi, 1e-10},

		// Small values
		{"small_pos", "DEGREES(0.001)", 0.001 * 180 / math.Pi, 1e-10},
		{"small_neg", "DEGREES(-0.001)", -0.001 * 180 / math.Pi, 1e-10},

		// Large values
		{"large", "DEGREES(100)", 100 * 180 / math.Pi, 1e-6},

		// Boolean coercion: TRUE=1, FALSE=0
		{"bool_true", "DEGREES(TRUE)", 180 / math.Pi, 1e-10},
		{"bool_false", "DEGREES(FALSE)", 0, 0},

		// String coercion
		{"string_num", `DEGREES("1")`, 180 / math.Pi, 1e-10},
		{"string_zero", `DEGREES("0")`, 0, 0},
		{"string_neg", `DEGREES("-1")`, -180 / math.Pi, 1e-10},

		// Doc example
		{"doc_example", "DEGREES(PI())", 180, 0},
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
		// No arguments
		{"no_args", "DEGREES()", ErrValVALUE},
		// Too many arguments
		{"too_many_args", "DEGREES(1,2)", ErrValVALUE},
		// Non-numeric string
		{"non_numeric", `DEGREES("abc")`, ErrValVALUE},
		// Error propagation
		{"err_div0", "DEGREES(1/0)", ErrValDIV0},
		{"err_na", "DEGREES(NA())", ErrValNA},
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

func TestQUOTIENT(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Documentation examples
		{"doc_ex1", "QUOTIENT(5,2)", 2},
		{"doc_ex2", "QUOTIENT(4.5,3.1)", 1},
		{"doc_ex3", "QUOTIENT(-10,3)", -3},

		// Basic positive division
		{"10_div_3", "QUOTIENT(10,3)", 3},
		{"7_div_2", "QUOTIENT(7,2)", 3},
		{"100_div_10", "QUOTIENT(100,10)", 10},
		{"1_div_1", "QUOTIENT(1,1)", 1},
		{"17_div_5", "QUOTIENT(17,5)", 3},
		{"99_div_10", "QUOTIENT(99,10)", 9},

		// QUOTIENT(x,1) equals INT(x) for positive x
		{"x_div_1_integer", "QUOTIENT(7,1)", 7},
		{"x_div_1_decimal", "QUOTIENT(5.9,1)", 5},
		{"x_div_1_large", "QUOTIENT(123,1)", 123},

		// Negative numerator (truncates toward zero)
		{"neg_num_pos_den", "QUOTIENT(-7,2)", -3},
		{"neg_num_pos_den2", "QUOTIENT(-1,2)", 0},
		{"neg_num_pos_den3", "QUOTIENT(-13,4)", -3},
		{"neg_num_pos_den4", "QUOTIENT(-17,5)", -3},
		{"neg_num_pos_den5", "QUOTIENT(-99,10)", -9},

		// Negative denominator
		{"pos_num_neg_den", "QUOTIENT(7,-2)", -3},
		{"pos_num_neg_den2", "QUOTIENT(10,-3)", -3},
		{"pos_num_neg_den3", "QUOTIENT(17,-5)", -3},
		{"pos_num_neg_den4", "QUOTIENT(1,-3)", 0},

		// Both negative (result positive, truncated toward zero)
		{"both_neg", "QUOTIENT(-10,-3)", 3},
		{"both_neg2", "QUOTIENT(-7,-2)", 3},
		{"both_neg3", "QUOTIENT(-17,-5)", 3},
		{"both_neg4", "QUOTIENT(-1,-3)", 0},

		// Decimal truncation behavior (truncates toward zero)
		{"trunc_pos", "QUOTIENT(7,3)", 2},
		{"trunc_neg_num", "QUOTIENT(-7,3)", -2},
		{"trunc_decimal_args", "QUOTIENT(9.9,3.1)", 3},
		{"trunc_small_result", "QUOTIENT(1,3)", 0},
		{"trunc_decimal_num", "QUOTIENT(5.5,2)", 2},
		{"trunc_decimal_den", "QUOTIENT(10,2.5)", 4},
		{"trunc_both_decimal", "QUOTIENT(7.7,2.2)", 3},
		{"trunc_neg_decimal", "QUOTIENT(-5.5,2)", -2},

		// Exact division (no truncation needed)
		{"exact_div", "QUOTIENT(10,2)", 5},
		{"exact_div2", "QUOTIENT(100,25)", 4},
		{"exact_div_neg", "QUOTIENT(-10,2)", -5},
		{"exact_div_self", "QUOTIENT(7,7)", 1},

		// Zero numerator
		{"zero_num", "QUOTIENT(0,5)", 0},
		{"zero_num_neg_den", "QUOTIENT(0,-3)", 0},
		{"zero_num_large_den", "QUOTIENT(0,1000000)", 0},

		// Large numbers
		{"large_nums", "QUOTIENT(1000000,3)", 333333},
		{"large_nums2", "QUOTIENT(999999,1000)", 999},

		// Numerator smaller than denominator
		{"num_lt_den", "QUOTIENT(2,5)", 0},
		{"num_lt_den_neg", "QUOTIENT(-2,5)", 0},
		{"num_lt_den2", "QUOTIENT(1,100)", 0},

		// String coercion
		{"string_num", `QUOTIENT("10",3)`, 3},
		{"string_den", `QUOTIENT(10,"3")`, 3},
		{"string_both", `QUOTIENT("10","3")`, 3},
		{"string_decimal", `QUOTIENT("7.5","2")`, 3},

		// Boolean coercion
		{"bool_true_num", "QUOTIENT(TRUE,1)", 1},
		{"bool_false_num", "QUOTIENT(FALSE,1)", 0},
		{"bool_true_den", "QUOTIENT(5,TRUE)", 5},
		{"bool_true_both", "QUOTIENT(TRUE,TRUE)", 1},
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
		// Division by zero
		{"div_by_zero", "QUOTIENT(5,0)", ErrValDIV0},
		{"div_by_zero_neg", "QUOTIENT(-5,0)", ErrValDIV0},
		{"div_zero_by_zero", "QUOTIENT(0,0)", ErrValDIV0},

		// Wrong argument count
		{"no_args", "QUOTIENT()", ErrValVALUE},
		{"one_arg", "QUOTIENT(5)", ErrValVALUE},
		{"three_args", "QUOTIENT(5,2,1)", ErrValVALUE},

		// Non-numeric string
		{"non_numeric_num", `QUOTIENT("abc",2)`, ErrValVALUE},
		{"non_numeric_den", `QUOTIENT(2,"abc")`, ErrValVALUE},
		{"non_numeric_both", `QUOTIENT("abc","def")`, ErrValVALUE},

		// Error propagation
		{"err_div0_num", "QUOTIENT(1/0,2)", ErrValDIV0},
		{"err_div0_den", "QUOTIENT(2,1/0)", ErrValDIV0},
		{"err_na_num", "QUOTIENT(NA(),2)", ErrValNA},
		{"err_na_den", "QUOTIENT(2,NA())", ErrValNA},

		// Boolean FALSE as denominator (coerces to 0 -> #DIV/0!)
		{"bool_false_den", "QUOTIENT(5,FALSE)", ErrValDIV0},

		// String "0" as denominator
		{"string_zero_den", `QUOTIENT(5,"0")`, ErrValDIV0},
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

// TestMODQUOTIENTCrossCheck verifies properties of MOD and QUOTIENT.
//
// Note: In Excel, MOD and QUOTIENT use *different* division semantics:
//   - MOD(a,b)     = a - b*INT(a/b)   — remainder with INT (floor) division
//   - QUOTIENT(a,b) = TRUNC(a/b)      — truncation toward zero
//
// Therefore a = QUOTIENT(a,b)*b + MOD(a,b) only holds when a and b have the
// same sign (where INT and TRUNC agree). The always-true identity for MOD is:
//
//	a = INT(a/b)*b + MOD(a,b)
//
// We test both properties here.
func TestMODQUOTIENTCrossCheck(t *testing.T) {
	resolver := &mockResolver{}
	const epsilon = 1e-10

	cases := []struct {
		name string
		a, b float64
	}{
		{"10_3", 10, 3},
		{"neg10_3", -10, 3},
		{"10_neg3", 10, -3},
		{"neg10_neg3", -10, -3},
		{"7_2", 7, 2},
		{"neg7_2", -7, 2},
		{"7_neg2", 7, -2},
		{"neg7_neg2", -7, -2},
		{"0_5", 0, 5},
		{"5_5", 5, 5},
		{"5_1", 5, 1},
		{"17_5", 17, 5},
		{"100_7", 100, 7},
		{"5_5_2", 5.5, 2},
		{"9_9_3_1", 9.9, 3.1},
		{"1_3", 1, 3},
		{"neg1_3", -1, 3},
		{"3_2", 3, 2},
		{"neg3_2", -3, 2},
		{"3_neg2", 3, -2},
		{"neg3_neg2", -3, -2},
	}

	for _, tt := range cases {
		t.Run(tt.name, func(t *testing.T) {
			modFormula := fmt.Sprintf("MOD(%g,%g)", tt.a, tt.b)
			quotFormula := fmt.Sprintf("QUOTIENT(%g,%g)", tt.a, tt.b)
			intFormula := fmt.Sprintf("INT(%g/%g)", tt.a, tt.b)

			cfMod := evalCompile(t, modFormula)
			modVal, err := Eval(cfMod, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", modFormula, err)
			}
			if modVal.Type != ValueNumber {
				t.Skipf("MOD returned non-number type %v (formula might hit precision limit)", modVal.Type)
			}

			cfQuot := evalCompile(t, quotFormula)
			quotVal, err := Eval(cfQuot, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", quotFormula, err)
			}
			if quotVal.Type != ValueNumber {
				t.Skipf("QUOTIENT returned non-number type %v", quotVal.Type)
			}

			cfInt := evalCompile(t, intFormula)
			intVal, err := Eval(cfInt, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", intFormula, err)
			}
			if intVal.Type != ValueNumber {
				t.Skipf("INT returned non-number type %v", intVal.Type)
			}

			// Always-true identity: a = INT(a/b)*b + MOD(a,b)
			reconstructedINT := intVal.Num*tt.b + modVal.Num
			if math.Abs(reconstructedINT-tt.a) > epsilon {
				t.Errorf("INT(%g/%g)*%g + MOD(%g,%g) = %g*%g + %g = %g, want %g",
					tt.a, tt.b, tt.b, tt.a, tt.b, intVal.Num, tt.b, modVal.Num, reconstructedINT, tt.a)
			}

			// When signs agree, QUOTIENT and INT should match, so the QUOTIENT identity also holds
			sameSign := (tt.a >= 0 && tt.b > 0) || (tt.a <= 0 && tt.b < 0)
			if sameSign {
				reconstructedQUOT := quotVal.Num*tt.b + modVal.Num
				if math.Abs(reconstructedQUOT-tt.a) > epsilon {
					t.Errorf("same-sign: QUOTIENT(%g,%g)*%g + MOD(%g,%g) = %g*%g + %g = %g, want %g",
						tt.a, tt.b, tt.b, tt.a, tt.b, quotVal.Num, tt.b, modVal.Num, reconstructedQUOT, tt.a)
				}
			}

			// MOD result sign always matches divisor sign (or is zero)
			if modVal.Num != 0 {
				if (modVal.Num > 0) != (tt.b > 0) {
					t.Errorf("MOD(%g,%g) = %g, sign should follow divisor %g", tt.a, tt.b, modVal.Num, tt.b)
				}
			}

			// QUOTIENT truncates toward zero: |QUOTIENT(a,b)| <= |a/b|
			exactQuot := tt.a / tt.b
			if math.Abs(quotVal.Num) > math.Abs(exactQuot)+epsilon {
				t.Errorf("QUOTIENT(%g,%g) = %g, should truncate toward zero from %g",
					tt.a, tt.b, quotVal.Num, exactQuot)
			}
		})
	}
}

func TestDECIMAL(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Documentation examples
		{"hex_FF", `DECIMAL("FF",16)`, 255},
		{"binary_111", `DECIMAL("111",2)`, 7},
		{"base36_zap", `DECIMAL("zap",36)`, 45745},

		// Hexadecimal (base 16)
		{"hex_1F", `DECIMAL("1F",16)`, 31},
		{"hex_A", `DECIMAL("A",16)`, 10},
		{"hex_lowercase_ff", `DECIMAL("ff",16)`, 255},
		{"hex_0", `DECIMAL("0",16)`, 0},

		// Binary (base 2)
		{"binary_0", `DECIMAL("0",2)`, 0},
		{"binary_1", `DECIMAL("1",2)`, 1},
		{"binary_10", `DECIMAL("10",2)`, 2},
		{"binary_1010", `DECIMAL("1010",2)`, 10},
		{"binary_11111111", `DECIMAL("11111111",2)`, 255},

		// Octal (base 8)
		{"octal_77", `DECIMAL("77",8)`, 63},
		{"octal_10", `DECIMAL("10",8)`, 8},

		// Decimal (base 10)
		{"decimal_0", `DECIMAL("0",10)`, 0},
		{"decimal_100", `DECIMAL("100",10)`, 100},
		{"decimal_42", `DECIMAL("42",10)`, 42},

		// Base 36 edge cases
		{"base36_z", `DECIMAL("z",36)`, 35},
		{"base36_10", `DECIMAL("10",36)`, 36},

		// Numeric first arg (number coerced to string)
		{"number_input_111_base2", `DECIMAL(111,2)`, 7},
		{"number_input_100_base10", `DECIMAL(100,10)`, 100},

		// Radix as float (truncated to integer)
		{"radix_float_16_9", `DECIMAL("FF",16.9)`, 255},
		{"radix_float_2_7", `DECIMAL("111",2.7)`, 7},

		// Whitespace trimming
		{"leading_space", `DECIMAL(" FF",16)`, 255},
		{"trailing_space", `DECIMAL("FF ",16)`, 255},
		{"both_spaces", `DECIMAL(" FF ",16)`, 255},
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
		// Invalid digits for base
		{"invalid_digit_base2", `DECIMAL("2",2)`, ErrValNUM},
		{"invalid_digit_base8", `DECIMAL("8",8)`, ErrValNUM},
		{"invalid_hex_char", `DECIMAL("GG",16)`, ErrValNUM},

		// Base out of range
		{"base_too_low_0", `DECIMAL("1",0)`, ErrValNUM},
		{"base_too_low_1", `DECIMAL("1",1)`, ErrValNUM},
		{"base_too_high_37", `DECIMAL("1",37)`, ErrValNUM},

		// Empty text
		{"empty_string", `DECIMAL("",16)`, ErrValNUM},
		{"whitespace_only", `DECIMAL("  ",16)`, ErrValNUM},

		// Wrong argument count
		{"no_args", `DECIMAL()`, ErrValVALUE},
		{"one_arg", `DECIMAL("FF")`, ErrValVALUE},
		{"three_args", `DECIMAL("FF",16,1)`, ErrValVALUE},

		// Boolean first arg
		{"bool_input", `DECIMAL(TRUE,10)`, ErrValVALUE},
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

func TestTRUNC(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
		tol     float64
	}{
		// Basic truncation toward zero (default num_digits=0)
		{"trunc_positive", "TRUNC(8.9)", 8, 0},
		{"trunc_negative", "TRUNC(-8.9)", -8, 0},
		{"trunc_between_0_and_1", "TRUNC(0.45)", 0, 0},

		// Truncation with num_digits > 0
		{"1dp_positive", "TRUNC(8.59,1)", 8.5, 1e-10},
		{"2dp_positive", "TRUNC(0.545,2)", 0.54, 1e-10},
		{"3dp_positive", "TRUNC(1.23789,3)", 1.237, 1e-10},
		{"1dp_negative", "TRUNC(-8.59,1)", -8.5, 1e-10},
		{"2dp_negative", "TRUNC(-0.545,2)", -0.54, 1e-10},

		// Truncation with negative num_digits (truncate to tens, hundreds)
		{"neg_digits_tens", "TRUNC(123,-1)", 120, 0},
		{"neg_digits_hundreds", "TRUNC(123,-2)", 100, 0},
		{"neg_digits_thousands", "TRUNC(5678,-3)", 5000, 0},
		{"neg_digits_tens_neg_num", "TRUNC(-123,-1)", -120, 0},
		{"neg_digits_hundreds_neg_num", "TRUNC(-567,-2)", -500, 0},

		// Explicit num_digits=0
		{"explicit_zero_digits", "TRUNC(3.7,0)", 3, 0},
		{"explicit_zero_digits_neg", "TRUNC(-3.7,0)", -3, 0},

		// Zero as input
		{"zero_value", "TRUNC(0)", 0, 0},
		{"zero_with_digits", "TRUNC(0,5)", 0, 0},
		{"zero_with_neg_digits", "TRUNC(0,-2)", 0, 0},

		// Very small decimals
		{"small_positive", "TRUNC(0.0001,3)", 0, 0},
		{"small_positive_4dp", "TRUNC(0.0001,4)", 0.0001, 1e-12},

		// Integer input (no fractional part)
		{"integer", "TRUNC(5)", 5, 0},
		{"integer_with_digits", "TRUNC(5,2)", 5, 0},

		// Boolean coercion (TRUE=1, FALSE=0)
		{"bool_true_number", "TRUNC(TRUE)", 1, 0},
		{"bool_false_number", "TRUNC(FALSE)", 0, 0},
		{"bool_true_digits", "TRUNC(3.14,TRUE)", 3.1, 1e-10},
		{"bool_false_digits", "TRUNC(3.14,FALSE)", 3, 0},

		// String coercion (numeric strings)
		{"string_number", `TRUNC("8.9")`, 8, 0},
		{"string_digits", `TRUNC(8.59,"1")`, 8.5, 1e-10},
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
		{"no_args", "TRUNC()", ErrValVALUE},
		{"too_many_args", "TRUNC(1,2,3)", ErrValVALUE},
		{"non_numeric_number", `TRUNC("abc")`, ErrValVALUE},
		{"non_numeric_digits", `TRUNC(1,"abc")`, ErrValVALUE},
		{"error_propagation_number", "TRUNC(1/0)", ErrValDIV0},
		{"error_propagation_digits", "TRUNC(1,1/0)", ErrValDIV0},
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

func TestLCM(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Documentation examples
		{"doc_ex1", "LCM(5,2)", 10},
		{"doc_ex2", "LCM(24,36)", 72},
		// Basic two-argument cases
		{"4_and_6", "LCM(4,6)", 12},
		{"3_and_7", "LCM(3,7)", 21},
		{"6_and_8", "LCM(6,8)", 24},
		// Zero cases — any zero argument yields 0
		{"5_and_0", "LCM(5,0)", 0},
		{"0_and_5", "LCM(0,5)", 0},
		{"0_and_0", "LCM(0,0)", 0},
		{"0_0_0", "LCM(0,0,0)", 0},
		// Multiple arguments
		{"three_args", "LCM(2,3,5)", 30},
		{"three_args2", "LCM(4,6,10)", 60},
		{"four_args", "LCM(2,3,4,5)", 60},
		// LCM(n,1) = n
		{"n_and_1_small", "LCM(7,1)", 7},
		{"n_and_1_large", "LCM(100,1)", 100},
		{"1_and_n", "LCM(1,12)", 12},
		// LCM(n,n) = n
		{"same_5", "LCM(5,5)", 5},
		{"same_12", "LCM(12,12)", 12},
		// Single argument
		{"single_arg", "LCM(7)", 7},
		{"single_arg_1", "LCM(1)", 1},
		{"single_arg_0", "LCM(0)", 0},
		// Decimal truncation — decimals truncated to integer
		{"decimal_5_9", "LCM(5.9,2)", 10},
		{"decimal_both", "LCM(4.7,6.3)", 12},
		{"decimal_0_9", "LCM(0.9,5)", 0},
		// Boolean coercion
		{"bool_true", "LCM(TRUE,5)", 5},
		{"bool_false", "LCM(FALSE,5)", 0},
		{"bool_true_true", "LCM(TRUE,TRUE)", 1},
		// String coercion — numeric strings
		{"string_5", `LCM("5",2)`, 10},
		{"string_both", `LCM("4","6")`, 12},
		// Negative decimal that truncates to -0 (not < 0), so treated as zero
		{"neg_decimal_trunc_zero", "LCM(-0.5,5)", 0},
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
		// Negative values → #NUM!
		{"negative_first", "LCM(-1,5)", ErrValNUM},
		{"negative_second", "LCM(5,-3)", ErrValNUM},
		{"negative_both", "LCM(-2,-3)", ErrValNUM},
		{"negative_decimal_truncates_neg", "LCM(-1.5,5)", ErrValNUM},
		// No arguments → #VALUE!
		{"no_args", "LCM()", ErrValVALUE},
		// Non-numeric string → #VALUE!
		{"non_numeric_string", `LCM("hello",5)`, ErrValVALUE},
		{"non_numeric_second", `LCM(5,"abc")`, ErrValVALUE},
		// Error propagation
		{"error_propagation_div0", "LCM(1/0,5)", ErrValDIV0},
		{"error_propagation_na", "LCM(NA(),5)", ErrValNA},
		{"error_in_second_arg", "LCM(5,1/0)", ErrValDIV0},
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

func TestGCD(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Basic two-argument cases (from docs)
		{"basic_12_8", "GCD(12,8)", 4},
		{"basic_24_36", "GCD(24,36)", 12},
		{"basic_5_2", "GCD(5,2)", 1},
		{"basic_7_1", "GCD(7,1)", 1},
		// GCD with zero: GCD(n,0) = n
		{"zero_second", "GCD(5,0)", 5},
		{"zero_first", "GCD(0,5)", 5},
		{"both_zero", "GCD(0,0)", 0},
		// Single argument
		{"single_arg", "GCD(42)", 42},
		{"single_zero", "GCD(0)", 0},
		// Multiple arguments
		{"three_args", "GCD(12,8,6)", 2},
		{"four_args", "GCD(60,48,36,24)", 12},
		{"three_args_common", "GCD(100,75,50)", 25},
		// GCD with 1 always returns 1
		{"one_first", "GCD(1,100)", 1},
		{"one_middle", "GCD(12,1,8)", 1},
		// Coprime numbers (GCD = 1)
		{"coprime", "GCD(7,13)", 1},
		{"coprime_large", "GCD(17,31)", 1},
		// Decimal truncation: non-integer values are truncated
		{"decimal_trunc_first", "GCD(12.5,8)", 4},
		{"decimal_trunc_second", "GCD(12,8.9)", 4},
		{"decimal_trunc_both", "GCD(12.7,8.3)", 4},
		{"decimal_trunc_large_frac", "GCD(24.999,36.001)", 12},
		// Same number: GCD(n,n) = n
		{"same_number", "GCD(15,15)", 15},
		// Large values
		{"large_values", "GCD(1000000,500000)", 500000},
		// Boolean coercion (TRUE=1, FALSE=0)
		{"bool_true", "GCD(TRUE,1)", 1},
		{"bool_false", "GCD(FALSE,5)", 5},
		{"bool_true_true", "GCD(TRUE,TRUE)", 1},
		// String number coercion
		{"string_number", `GCD("12","8")`, 4},
		{"string_single", `GCD("42")`, 42},
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
		// Negative values -> #NUM!
		{"negative_first", "GCD(-5,10)", ErrValNUM},
		{"negative_second", "GCD(10,-5)", ErrValNUM},
		{"negative_both", "GCD(-3,-7)", ErrValNUM},
		{"negative_decimal", "GCD(-1.5,8)", ErrValNUM},
		// Non-numeric string -> #VALUE!
		{"non_numeric_string", `GCD("hello",5)`, ErrValVALUE},
		{"non_numeric_second", `GCD(5,"abc")`, ErrValVALUE},
		{"empty_string", `GCD("",5)`, ErrValVALUE},
		// Error propagation
		{"error_propagation_div0", "GCD(1/0,5)", ErrValDIV0},
		{"error_propagation_na", "GCD(NA(),5)", ErrValNA},
		{"error_propagation_second", "GCD(5,1/0)", ErrValDIV0},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): unexpected error: %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.errVal {
				t.Errorf("Eval(%q) = type=%v err=%v, want error %v", tt.formula, got.Type, got.Err, tt.errVal)
			}
		})
	}
}

// --------------- PRODUCT tests ---------------

func TestPRODUCT(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// Basic multiplication
		{"two_args", "PRODUCT(2,3)", 6},
		{"five_args", "PRODUCT(1,2,3,4,5)", 120},

		// Doc examples
		{"doc_ex1", "PRODUCT(5,15,30)", 2250},
		{"doc_ex1_with_multiplier", "PRODUCT(5,15,30,2)", 4500},

		// Single value
		{"single_value", "PRODUCT(5)", 5},
		{"single_one", "PRODUCT(1)", 1},

		// Product with zero
		{"zero_factor", "PRODUCT(5,0)", 0},
		{"zero_first", "PRODUCT(0,3)", 0},
		{"zero_middle", "PRODUCT(2,0,5)", 0},

		// Negative numbers
		{"neg_times_pos", "PRODUCT(-3,4)", -12},
		{"neg_times_neg", "PRODUCT(-3,-4)", 12},
		{"two_neg_one_pos", "PRODUCT(-2,-3,5)", 30},
		{"odd_negatives", "PRODUCT(-1,-2,-3)", -6},

		// Decimal numbers
		{"decimals", "PRODUCT(2.5,4)", 10},
		{"two_decimals", "PRODUCT(0.5,0.5)", 0.25},
		{"neg_decimal", "PRODUCT(-1.5,2)", -3},

		// Boolean coercion (direct args are coerced)
		{"bool_true", "PRODUCT(TRUE,5)", 5},
		{"bool_false", "PRODUCT(FALSE,5)", 0},
		{"bool_true_true", "PRODUCT(TRUE,TRUE)", 1},

		// String coercion (direct numeric string args are coerced)
		{"string_num", `PRODUCT("3",4)`, 12},
		{"string_neg", `PRODUCT("-2",5)`, -10},

		// Large numbers
		{"large", "PRODUCT(1000,1000,1000)", 1e9},
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
		// Non-numeric string (direct arg)
		{"non_numeric_string", `PRODUCT("abc",5)`, ErrValVALUE},
		// Error propagation
		{"err_div0", "PRODUCT(1/0,5)", ErrValDIV0},
		{"err_na", "PRODUCT(NA(),5)", ErrValNA},
		{"err_in_second_arg", "PRODUCT(5,1/0)", ErrValDIV0},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): unexpected error: %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = type=%v err=%v, want error %v", tt.formula, got.Type, got.Err, tt.wantErr)
			}
		})
	}
}

func TestPRODUCT_RangeIgnoresTextAndBool(t *testing.T) {
	// In ranges, text and logical values are ignored (only numbers are used)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: StringVal("hello"),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: BoolVal(true),
			{Col: 1, Row: 5}: NumberVal(5),
		},
	}
	cf := evalCompile(t, "PRODUCT(A1:A5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// Only 2*3*5 = 30; text "hello" and TRUE are ignored in ranges
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("got %v (%g), want 30", got.Type, got.Num)
	}
}

func TestPRODUCT_RangeWithEmptyCells(t *testing.T) {
	// Empty cells in ranges are ignored (not treated as zero)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(4),
			// Row 2 is empty
			{Col: 1, Row: 3}: NumberVal(5),
		},
	}
	cf := evalCompile(t, "PRODUCT(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// 4*5 = 20; empty cell is ignored
	if got.Type != ValueNumber || got.Num != 20 {
		t.Errorf("got %v (%g), want 20", got.Type, got.Num)
	}
}

func TestPRODUCT_RangeErrorPropagation(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2),
			{Col: 1, Row: 2}: ErrorVal(ErrValDIV0),
			{Col: 1, Row: 3}: NumberVal(3),
		},
	}
	cf := evalCompile(t, "PRODUCT(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("got type=%v err=%v, want #DIV/0!", got.Type, got.Err)
	}
}

func TestPRODUCT_RangeAndDirectArg(t *testing.T) {
	// Mix of range and direct argument: PRODUCT(A1:A3, 2) = 5*15*30*2 = 4500
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(5),
			{Col: 1, Row: 2}: NumberVal(15),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}
	cf := evalCompile(t, "PRODUCT(A1:A3,2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 4500 {
		t.Errorf("got %v (%g), want 4500", got.Type, got.Num)
	}
}

func TestPRODUCT_AllEmptyRange(t *testing.T) {
	// Range with no numeric values returns 0 (count == 0)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("a"),
			{Col: 1, Row: 2}: StringVal("b"),
		},
	}
	cf := evalCompile(t, "PRODUCT(A1:A2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("got %v (%g), want 0", got.Type, got.Num)
	}
}

func TestPERMUT(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Doc examples
		{"doc_example_100_3", "PERMUT(100,3)", 970200},
		{"doc_example_3_2", "PERMUT(3,2)", 6},
		// Basic permutations
		{"5_perm_2", "PERMUT(5,2)", 20},
		{"6_perm_3", "PERMUT(6,3)", 120},
		{"10_perm_4", "PERMUT(10,4)", 5040},
		// PERMUT(n,0) = 1 for any n > 0
		{"4_perm_0", "PERMUT(4,0)", 1},
		{"10_perm_0", "PERMUT(10,0)", 1},
		{"1_perm_0", "PERMUT(1,0)", 1},
		// PERMUT(n,n) = n!
		{"4_perm_4", "PERMUT(4,4)", 24},
		{"5_perm_5", "PERMUT(5,5)", 120},
		{"1_perm_1", "PERMUT(1,1)", 1},
		// PERMUT(n,1) = n
		{"5_perm_1", "PERMUT(5,1)", 5},
		{"12_perm_1", "PERMUT(12,1)", 12},
		// Larger values
		{"20_perm_5", "PERMUT(20,5)", 1860480},
		{"8_perm_3", "PERMUT(8,3)", 336},
		// Decimal truncation — arguments are truncated to integers
		{"decimal_n", "PERMUT(5.9,2)", 20},
		{"decimal_k", "PERMUT(5,2.7)", 20},
		{"decimal_both", "PERMUT(6.9,3.1)", 120},
		// String coercion
		{"string_n", `PERMUT("5",2)`, 20},
		{"string_k", `PERMUT(5,"2")`, 20},
		{"string_both", `PERMUT("10","3")`, 720},
		// Boolean coercion
		{"bool_true_n", "PERMUT(TRUE,0)", 1},
		{"bool_true_k", "PERMUT(5,TRUE)", 5},
		{"bool_false_k", "PERMUT(5,FALSE)", 1},
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
		// n < k → #NUM!
		{"n_less_than_k", "PERMUT(3,5)", ErrValNUM},
		{"n_less_than_k_2", "PERMUT(2,10)", ErrValNUM},
		// n <= 0 → #NUM!
		{"zero_n", "PERMUT(0,0)", ErrValNUM},
		{"negative_n", "PERMUT(-1,2)", ErrValNUM},
		{"negative_n_large", "PERMUT(-5,3)", ErrValNUM},
		// Negative k → #NUM!
		{"negative_k", "PERMUT(5,-1)", ErrValNUM},
		{"negative_k_large", "PERMUT(10,-3)", ErrValNUM},
		// Both negative → #NUM!
		{"both_negative", "PERMUT(-3,-1)", ErrValNUM},
		// Wrong number of arguments → #VALUE!
		{"no_args", "PERMUT()", ErrValVALUE},
		{"one_arg", "PERMUT(5)", ErrValVALUE},
		{"three_args", "PERMUT(5,2,1)", ErrValVALUE},
		// Non-numeric string → #VALUE!
		{"non_numeric_n", `PERMUT("abc",2)`, ErrValVALUE},
		{"non_numeric_k", `PERMUT(5,"xyz")`, ErrValVALUE},
		// Error propagation
		{"error_div0_n", "PERMUT(1/0,2)", ErrValDIV0},
		{"error_div0_k", "PERMUT(5,1/0)", ErrValDIV0},
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

func TestBASE(t *testing.T) {
	resolver := &mockResolver{}

	// String-result tests: BASE returns a string representation.
	strTests := []struct {
		name    string
		formula string
		want    string
	}{
		// Documentation examples
		{"doc_binary_7", "BASE(7,2)", "111"},
		{"doc_hex_100", "BASE(100,16)", "64"},
		{"doc_binary_15_padded", "BASE(15,2,10)", "0000001111"},

		// Additional basic conversions
		{"hex_15", "BASE(15,16)", "F"},
		{"hex_255", "BASE(255,16)", "FF"},
		{"octal_8", "BASE(8,8)", "10"},
		{"octal_63", "BASE(63,8)", "77"},
		{"base36", "BASE(35,36)", "Z"},
		{"base36_large", "BASE(1295,36)", "ZZ"},

		// Zero
		{"zero_binary", "BASE(0,2)", "0"},
		{"zero_hex", "BASE(0,16)", "0"},
		{"zero_base10", "BASE(0,10)", "0"},

		// Min-length padding
		{"pad_binary_7", "BASE(7,2,8)", "00000111"},
		{"pad_hex_1", "BASE(1,16,4)", "0001"},
		{"pad_zero", "BASE(0,2,5)", "00000"},
		// min_length shorter than result: no truncation
		{"pad_shorter", "BASE(255,2,1)", "11111111"},
		// min_length equals result length: no change
		{"pad_exact", "BASE(7,2,3)", "111"},
		// min_length zero: same as omitting
		{"pad_zero_len", "BASE(7,2,0)", "111"},

		// Decimal truncation: non-integer number is truncated
		{"trunc_number", "BASE(7.9,2)", "111"},
		{"trunc_radix", "BASE(7,2.9)", "111"},
		{"trunc_minlen", "BASE(7,2,8.7)", "00000111"},

		// String and boolean coercion
		{"string_number", `BASE("7",2)`, "111"},
		{"string_radix", `BASE(7,"2")`, "111"},
		{"bool_true_number", "BASE(TRUE,2)", "1"},
	}

	for _, tt := range strTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): unexpected error: %v", tt.formula, err)
			}
			if got.Type != ValueString {
				t.Fatalf("Eval(%q): got type %v, want string", tt.formula, got.Type)
			}
			if got.Str != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
			}
		})
	}

	// Error tests
	errTests := []struct {
		name    string
		formula string
		errVal  ErrorValue
	}{
		// Radix out of range
		{"radix_too_low", "BASE(7,1)", ErrValNUM},
		{"radix_too_high", "BASE(7,37)", ErrValNUM},

		// Negative number
		{"negative_number", "BASE(-1,2)", ErrValNUM},

		// Negative min_length
		{"negative_minlen", "BASE(7,2,-1)", ErrValNUM},

		// Wrong argument count
		{"too_few_args", "BASE(7)", ErrValVALUE},
		{"too_many_args", "BASE(7,2,8,1)", ErrValVALUE},

		// Non-numeric string
		{"non_numeric_string", `BASE("abc",2)`, ErrValVALUE},
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

func TestMULTINOMIAL(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Doc example
		{"doc_example_2_3_4", "MULTINOMIAL(2,3,4)", 1260},
		// Basic multinomial identities
		{"two_ones", "MULTINOMIAL(1,1)", 2},
		{"three_ones", "MULTINOMIAL(1,1,1)", 6},
		{"four_ones", "MULTINOMIAL(1,1,1,1)", 24},
		// Single argument: MULTINOMIAL(n) = n!/n! = 1
		{"single_5", "MULTINOMIAL(5)", 1},
		{"single_1", "MULTINOMIAL(1)", 1},
		{"single_0", "MULTINOMIAL(0)", 1},
		{"single_10", "MULTINOMIAL(10)", 1},
		// Zeros: MULTINOMIAL(0,0) = 0!/(0!*0!) = 1
		{"zero_zero", "MULTINOMIAL(0,0)", 1},
		{"zero_zero_zero", "MULTINOMIAL(0,0,0)", 1},
		{"zero_and_val", "MULTINOMIAL(0,5)", 1},
		{"val_and_zero", "MULTINOMIAL(3,0)", 1},
		// Known combinatorial values
		{"2_2", "MULTINOMIAL(2,2)", 6},
		{"3_3", "MULTINOMIAL(3,3)", 20},
		{"2_2_2", "MULTINOMIAL(2,2,2)", 90},
		{"1_2_3", "MULTINOMIAL(1,2,3)", 60},
		{"5_3", "MULTINOMIAL(5,3)", 56},
		// Decimal truncation — arguments are truncated to integers
		{"decimal_2_9", "MULTINOMIAL(2.9,3,4)", 1260},
		{"decimal_3_7", "MULTINOMIAL(2,3.7,4)", 1260},
		{"decimal_all", "MULTINOMIAL(2.1,3.9,4.5)", 1260},
		// String coercion
		{"string_first", `MULTINOMIAL("2",3,4)`, 1260},
		{"string_second", `MULTINOMIAL(2,"3",4)`, 1260},
		{"string_all", `MULTINOMIAL("2","3","4")`, 1260},
		// Boolean coercion: TRUE=1, FALSE=0
		{"bool_true", "MULTINOMIAL(TRUE,1)", 2},
		{"bool_false", "MULTINOMIAL(FALSE,5)", 1},
		{"bool_true_true", "MULTINOMIAL(TRUE,TRUE,TRUE)", 6},
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
		// No arguments → #VALUE!
		{"no_args", "MULTINOMIAL()", ErrValVALUE},
		// Negative values → #NUM!
		{"negative_single", "MULTINOMIAL(-1)", ErrValNUM},
		{"negative_first", "MULTINOMIAL(-1,2,3)", ErrValNUM},
		{"negative_second", "MULTINOMIAL(2,-3,4)", ErrValNUM},
		{"negative_last", "MULTINOMIAL(2,3,-4)", ErrValNUM},
		{"negative_decimal_trunc", "MULTINOMIAL(-1.5,3)", ErrValNUM},
		// Non-numeric string → #VALUE!
		{"non_numeric", `MULTINOMIAL("abc")`, ErrValVALUE},
		{"non_numeric_second", `MULTINOMIAL(2,"xyz",4)`, ErrValVALUE},
		// Error propagation
		{"error_div0", "MULTINOMIAL(1/0,2)", ErrValDIV0},
		{"error_div0_second", "MULTINOMIAL(2,1/0)", ErrValDIV0},
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

func TestSUBTOTAL(t *testing.T) {
	// Set up a resolver with values in A1:A5 = {1, 2, 3, 4, 5}.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 1, Row: 4}: NumberVal(4),
			{Col: 1, Row: 5}: NumberVal(5),
		},
	}

	// SUBTOTAL delegates to aggregate functions. For A1:A5 = {1,2,3,4,5}:
	//   AVERAGE = 3, COUNT = 5, COUNTA = 5, MAX = 5, MIN = 1,
	//   PRODUCT = 120, SUM = 15
	//   STDEV  = sqrt(2.5) ~ 1.5811388
	//   STDEVP = sqrt(2)   ~ 1.4142136
	//   VAR    = 2.5
	//   VARP   = 2
	numTests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
	}{
		// function_num 1-11
		{"avg_1", "SUBTOTAL(1,A1:A5)", 3, 0},
		{"count_2", "SUBTOTAL(2,A1:A5)", 5, 0},
		{"counta_3", "SUBTOTAL(3,A1:A5)", 5, 0},
		{"max_4", "SUBTOTAL(4,A1:A5)", 5, 0},
		{"min_5", "SUBTOTAL(5,A1:A5)", 1, 0},
		{"product_6", "SUBTOTAL(6,A1:A5)", 120, 0},
		{"stdev_7", "SUBTOTAL(7,A1:A5)", math.Sqrt(2.5), 1e-10},
		{"stdevp_8", "SUBTOTAL(8,A1:A5)", math.Sqrt(2), 1e-10},
		{"sum_9", "SUBTOTAL(9,A1:A5)", 15, 0},
		{"var_10", "SUBTOTAL(10,A1:A5)", 2.5, 1e-10},
		{"varp_11", "SUBTOTAL(11,A1:A5)", 2, 1e-10},

		// function_num 101-111 (ignore hidden rows variants, same results
		// without a HiddenRowChecker implementation)
		{"avg_101", "SUBTOTAL(101,A1:A5)", 3, 0},
		{"count_102", "SUBTOTAL(102,A1:A5)", 5, 0},
		{"counta_103", "SUBTOTAL(103,A1:A5)", 5, 0},
		{"max_104", "SUBTOTAL(104,A1:A5)", 5, 0},
		{"min_105", "SUBTOTAL(105,A1:A5)", 1, 0},
		{"product_106", "SUBTOTAL(106,A1:A5)", 120, 0},
		{"stdev_107", "SUBTOTAL(107,A1:A5)", math.Sqrt(2.5), 1e-10},
		{"stdevp_108", "SUBTOTAL(108,A1:A5)", math.Sqrt(2), 1e-10},
		{"sum_109", "SUBTOTAL(109,A1:A5)", 15, 0},
		{"var_110", "SUBTOTAL(110,A1:A5)", 2.5, 1e-10},
		{"varp_111", "SUBTOTAL(111,A1:A5)", 2, 1e-10},

		// Single cell range
		{"sum_single", "SUBTOTAL(9,A1:A1)", 1, 0},
		{"avg_single", "SUBTOTAL(1,A1:A1)", 1, 0},
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
	errTests := []struct {
		name    string
		formula string
		errVal  ErrorValue
	}{
		// Invalid function_num values
		{"invalid_0", "SUBTOTAL(0,A1:A5)", ErrValVALUE},
		{"invalid_12", "SUBTOTAL(12,A1:A5)", ErrValVALUE},
		{"invalid_100", "SUBTOTAL(100,A1:A5)", ErrValVALUE},
		{"invalid_112", "SUBTOTAL(112,A1:A5)", ErrValVALUE},
		{"invalid_neg", "SUBTOTAL(-1,A1:A5)", ErrValVALUE},
		{"invalid_99", "SUBTOTAL(99,A1:A5)", ErrValVALUE},

		// Wrong argument count (too few)
		{"too_few_args", "SUBTOTAL(9)", ErrValVALUE},

		// Error propagation: error in function_num
		{"error_in_funcnum", "SUBTOTAL(1/0,A1:A5)", ErrValDIV0},
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

// subtotalMockResolver implements CellResolver, SubtotalChecker, and
// HiddenRowChecker for comprehensive SUBTOTAL testing.
type subtotalMockResolver struct {
	cells          map[CellAddr]Value
	subtotalCells  map[CellAddr]bool       // cells that contain SUBTOTAL formulas
	hiddenRows     map[string]map[int]bool // sheet -> row -> hidden
	autoFilterRows map[string]map[int]bool // sheet -> row -> filtered by autoFilter
}

func (m *subtotalMockResolver) GetCellValue(addr CellAddr) Value {
	if v, ok := m.cells[addr]; ok {
		return v
	}
	return EmptyVal()
}

func (m *subtotalMockResolver) GetRangeValues(addr RangeAddr) [][]Value {
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

func (m *subtotalMockResolver) IsSubtotalCell(sheet string, col, row int) bool {
	if m.subtotalCells == nil {
		return false
	}
	return m.subtotalCells[CellAddr{Sheet: sheet, Col: col, Row: row}]
}

func (m *subtotalMockResolver) IsRowHidden(sheet string, row int) bool {
	if m.hiddenRows == nil {
		return false
	}
	if sheetMap, ok := m.hiddenRows[sheet]; ok {
		return sheetMap[row]
	}
	return false
}

func (m *subtotalMockResolver) IsRowFilteredByAutoFilter(sheet string, row int) bool {
	if m.autoFilterRows == nil {
		return false
	}
	if sheetMap, ok := m.autoFilterRows[sheet]; ok {
		return sheetMap[row]
	}
	return false
}

type subtotalSparseResolver struct {
	subtotalMockResolver
}

func (m *subtotalSparseResolver) GetRangeValues(addr RangeAddr) [][]Value {
	return (&sparseResolver{cells: m.cells}).GetRangeValues(addr)
}

func TestSUBTOTALComprehensive(t *testing.T) {
	// ---------------------------------------------------------------
	// 1. Multiple ref arguments: SUBTOTAL(9, A1:A3, B1:B3)
	// ---------------------------------------------------------------
	t.Run("multiple_refs_sum", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
				{Col: 1, Row: 3}: NumberVal(3),
				{Col: 2, Row: 1}: NumberVal(10),
				{Col: 2, Row: 2}: NumberVal(20),
				{Col: 2, Row: 3}: NumberVal(30),
			},
		}
		cf := evalCompile(t, "SUBTOTAL(9,A1:A3,B1:B3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 66 {
			t.Errorf("got %v (%g), want 66", got.Type, got.Num)
		}
	})

	t.Run("multiple_refs_count", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
				{Col: 1, Row: 2}: NumberVal(2),
				{Col: 2, Row: 1}: NumberVal(10),
				{Col: 2, Row: 2}: NumberVal(20),
			},
		}
		cf := evalCompile(t, "SUBTOTAL(2,A1:A2,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 4 {
			t.Errorf("got %v (%g), want 4", got.Type, got.Num)
		}
	})

	t.Run("multiple_refs_average", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 2, Row: 1}: NumberVal(30),
				{Col: 2, Row: 2}: NumberVal(40),
			},
		}
		cf := evalCompile(t, "SUBTOTAL(1,A1:A2,B1:B2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 25 {
			t.Errorf("got %v (%g), want 25", got.Type, got.Num)
		}
	})

	// ---------------------------------------------------------------
	// 2. Empty range — COUNT returns 0, SUM returns 0, AVERAGE returns #DIV/0!
	// ---------------------------------------------------------------
	t.Run("empty_range_sum", func(t *testing.T) {
		resolver := &mockResolver{cells: map[CellAddr]Value{}}
		cf := evalCompile(t, "SUBTOTAL(9,A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got type=%v num=%g, want 0", got.Type, got.Num)
		}
	})

	t.Run("empty_range_count", func(t *testing.T) {
		resolver := &mockResolver{cells: map[CellAddr]Value{}}
		cf := evalCompile(t, "SUBTOTAL(2,A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got type=%v num=%g, want 0", got.Type, got.Num)
		}
	})

	t.Run("empty_range_average_div0", func(t *testing.T) {
		resolver := &mockResolver{cells: map[CellAddr]Value{}}
		cf := evalCompile(t, "SUBTOTAL(1,A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got type=%v err=%v, want #DIV/0!", got.Type, got.Err)
		}
	})

	t.Run("empty_range_max", func(t *testing.T) {
		resolver := &mockResolver{cells: map[CellAddr]Value{}}
		cf := evalCompile(t, "SUBTOTAL(4,A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		// MAX of empty range = 0
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got type=%v num=%g, want 0", got.Type, got.Num)
		}
	})

	t.Run("empty_range_min", func(t *testing.T) {
		resolver := &mockResolver{cells: map[CellAddr]Value{}}
		cf := evalCompile(t, "SUBTOTAL(5,A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		// MIN of empty range = 0
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got type=%v num=%g, want 0", got.Type, got.Num)
		}
	})

	// ---------------------------------------------------------------
	// 3. Range with mixed types (text ignored for numeric modes)
	// ---------------------------------------------------------------
	t.Run("mixed_types_sum_ignores_text", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: StringVal("hello"),
				{Col: 1, Row: 3}: NumberVal(20),
				{Col: 1, Row: 4}: BoolVal(true),
				{Col: 1, Row: 5}: NumberVal(30),
			},
		}
		// SUM ignores text and booleans in ranges
		cf := evalCompile(t, "SUBTOTAL(9,A1:A5)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 60 {
			t.Errorf("got %v (%g), want 60", got.Type, got.Num)
		}
	})

	t.Run("mixed_types_count_numbers_only", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: StringVal("hello"),
				{Col: 1, Row: 3}: NumberVal(20),
				{Col: 1, Row: 4}: BoolVal(true),
			},
		}
		// COUNT counts only numeric values (not strings, not bools in ranges)
		cf := evalCompile(t, "SUBTOTAL(2,A1:A4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v (%g), want 2", got.Type, got.Num)
		}
	})

	t.Run("mixed_types_counta", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: StringVal("hello"),
				{Col: 1, Row: 3}: NumberVal(20),
				{Col: 1, Row: 4}: BoolVal(true),
			},
		}
		// COUNTA counts all non-empty values
		cf := evalCompile(t, "SUBTOTAL(3,A1:A4)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 4 {
			t.Errorf("got %v (%g), want 4", got.Type, got.Num)
		}
	})

	t.Run("mixed_types_max_ignores_text", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: StringVal("zzz"),
				{Col: 1, Row: 3}: NumberVal(5),
			},
		}
		cf := evalCompile(t, "SUBTOTAL(4,A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v (%g), want 10", got.Type, got.Num)
		}
	})

	t.Run("mixed_types_min_ignores_text", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: StringVal("aaa"),
				{Col: 1, Row: 3}: NumberVal(5),
			},
		}
		cf := evalCompile(t, "SUBTOTAL(5,A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v (%g), want 5", got.Type, got.Num)
		}
	})

	// ---------------------------------------------------------------
	// 4. Range with errors → propagate
	// ---------------------------------------------------------------
	t.Run("error_in_range_propagates_sum", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: ErrorVal(ErrValNA),
				{Col: 1, Row: 3}: NumberVal(20),
			},
		}
		cf := evalCompile(t, "SUBTOTAL(9,A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got type=%v err=%v, want #N/A", got.Type, got.Err)
		}
	})

	t.Run("error_in_range_propagates_average", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: ErrorVal(ErrValDIV0),
				{Col: 1, Row: 3}: NumberVal(20),
			},
		}
		cf := evalCompile(t, "SUBTOTAL(1,A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got type=%v err=%v, want #DIV/0!", got.Type, got.Err)
		}
	})

	t.Run("count_skips_errors_in_range", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: ErrorVal(ErrValREF),
				{Col: 1, Row: 3}: NumberVal(20),
			},
		}
		cf := evalCompile(t, "SUBTOTAL(2,A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		// COUNT only counts numeric values; errors are not numeric so they are skipped
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got type=%v num=%g, want 2", got.Type, got.Num)
		}
	})

	// ---------------------------------------------------------------
	// 5. SUBTOTAL nested in another SUBTOTAL (inner should be ignored)
	// ---------------------------------------------------------------
	t.Run("nested_subtotal_ignored", func(t *testing.T) {
		// A1=10, A2=20, A3=SUBTOTAL(9,...) which evaluates to 100.
		// SUBTOTAL(9, A1:A3) should skip A3 and return 10+20=30.
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(100), // contains a SUBTOTAL formula
			},
			subtotalCells: map[CellAddr]bool{
				{Col: 1, Row: 3}: true,
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(9,A1:A3)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 30 {
			t.Errorf("got %v (%g), want 30", got.Type, got.Num)
		}
	})

	t.Run("nested_subtotal_ignored_average", func(t *testing.T) {
		// A1=10, A2=20, A3=SUBTOTAL (value 100, should be ignored)
		// AVERAGE of {10, 20} = 15
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(100),
			},
			subtotalCells: map[CellAddr]bool{
				{Col: 1, Row: 3}: true,
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(1,A1:A3)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf("got %v (%g), want 15", got.Type, got.Num)
		}
	})

	t.Run("nested_subtotal_ignored_count", func(t *testing.T) {
		// A1=10, A2=20, A3=SUBTOTAL (value 100, should be ignored)
		// COUNT should be 2, not 3
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(100),
			},
			subtotalCells: map[CellAddr]bool{
				{Col: 1, Row: 3}: true,
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(2,A1:A3)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v (%g), want 2", got.Type, got.Num)
		}
	})

	// ---------------------------------------------------------------
	// 6. Hidden rows with 101-111 modes
	// ---------------------------------------------------------------
	t.Run("hidden_rows_sum_109", func(t *testing.T) {
		// A1=10, A2=20 (hidden), A3=30
		// SUBTOTAL(109, A1:A3) should skip row 2 → 10+30=40
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
			},
			hiddenRows: map[string]map[int]bool{
				"": {2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(109,A1:A3)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 40 {
			t.Errorf("got %v (%g), want 40", got.Type, got.Num)
		}
	})

	t.Run("hidden_rows_not_skipped_mode_9", func(t *testing.T) {
		// A1=10, A2=20 (hidden but NOT auto-filtered), A3=30
		// SUBTOTAL(9, A1:A3) should include hidden rows → 10+20+30=60
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
			},
			hiddenRows: map[string]map[int]bool{
				"": {2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(9,A1:A3)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 60 {
			t.Errorf("got %v (%g), want 60", got.Type, got.Num)
		}
	})

	t.Run("autofilter_hidden_skipped_mode_9", func(t *testing.T) {
		// A1=10, A2=20 (auto-filter hidden), A3=30
		// SUBTOTAL(9, A1:A3) should skip auto-filtered rows → 10+30=40
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
			},
			autoFilterRows: map[string]map[int]bool{
				"": {2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(9,A1:A3)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 40 {
			t.Errorf("got %v (%g), want 40", got.Type, got.Num)
		}
	})

	t.Run("hidden_rows_average_101", func(t *testing.T) {
		// A1=10, A2=20 (hidden), A3=30
		// SUBTOTAL(101, A1:A3) should skip row 2 → AVERAGE(10,30)=20
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
			},
			hiddenRows: map[string]map[int]bool{
				"": {2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(101,A1:A3)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf("got %v (%g), want 20", got.Type, got.Num)
		}
	})

	t.Run("hidden_rows_max_104", func(t *testing.T) {
		// A1=10, A2=50 (hidden), A3=30
		// SUBTOTAL(104, A1:A3) should skip row 2 → MAX(10,30)=30
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(50),
				{Col: 1, Row: 3}: NumberVal(30),
			},
			hiddenRows: map[string]map[int]bool{
				"": {2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(104,A1:A3)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 30 {
			t.Errorf("got %v (%g), want 30", got.Type, got.Num)
		}
	})

	t.Run("hidden_rows_min_105", func(t *testing.T) {
		// A1=10, A2=1 (hidden), A3=30
		// SUBTOTAL(105, A1:A3) should skip row 2 → MIN(10,30)=10
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(1),
				{Col: 1, Row: 3}: NumberVal(30),
			},
			hiddenRows: map[string]map[int]bool{
				"": {2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(105,A1:A3)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 10 {
			t.Errorf("got %v (%g), want 10", got.Type, got.Num)
		}
	})

	t.Run("hidden_rows_product_106", func(t *testing.T) {
		// A1=2, A2=100 (hidden), A3=3
		// SUBTOTAL(106, A1:A3) should skip row 2 → PRODUCT(2,3)=6
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(2),
				{Col: 1, Row: 2}: NumberVal(100),
				{Col: 1, Row: 3}: NumberVal(3),
			},
			hiddenRows: map[string]map[int]bool{
				"": {2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(106,A1:A3)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 6 {
			t.Errorf("got %v (%g), want 6", got.Type, got.Num)
		}
	})

	t.Run("hidden_rows_count_102", func(t *testing.T) {
		// A1=10, A2=20 (hidden), A3=30
		// SUBTOTAL(102, A1:A3) should skip row 2 → COUNT=2
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
			},
			hiddenRows: map[string]map[int]bool{
				"": {2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(102,A1:A3)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v (%g), want 2", got.Type, got.Num)
		}
	})

	t.Run("hidden_rows_counta_103", func(t *testing.T) {
		// A1=10, A2="text" (hidden), A3=30
		// SUBTOTAL(103, A1:A3) should skip row 2 → COUNTA=2
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: StringVal("text"),
				{Col: 1, Row: 3}: NumberVal(30),
			},
			hiddenRows: map[string]map[int]bool{
				"": {2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(103,A1:A3)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf("got %v (%g), want 2", got.Type, got.Num)
		}
	})

	// ---------------------------------------------------------------
	// 7. Boolean values in range
	// ---------------------------------------------------------------
	t.Run("bool_values_counta", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
				{Col: 1, Row: 2}: BoolVal(false),
				{Col: 1, Row: 3}: NumberVal(5),
			},
		}
		// COUNTA counts all non-empty values including booleans
		cf := evalCompile(t, "SUBTOTAL(3,A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf("got %v (%g), want 3", got.Type, got.Num)
		}
	})

	t.Run("bool_values_count_skips", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
				{Col: 1, Row: 2}: BoolVal(false),
				{Col: 1, Row: 3}: NumberVal(5),
			},
		}
		// COUNT only counts numeric values (booleans in range are not counted)
		cf := evalCompile(t, "SUBTOTAL(2,A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v (%g), want 1", got.Type, got.Num)
		}
	})

	// ---------------------------------------------------------------
	// 8. String function_num ("9" stored in cell)
	// ---------------------------------------------------------------
	t.Run("string_funcnum_coerced", func(t *testing.T) {
		// Use "9" as a string that gets coerced to number by SUBTOTAL
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 2, Row: 1}: StringVal("9"),
			},
		}
		// SUBTOTAL(B1, A1:A2) where B1="9"
		cf := evalCompile(t, "SUBTOTAL(B1,A1:A2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 30 {
			t.Errorf("got %v (%g), want 30", got.Type, got.Num)
		}
	})

	t.Run("string_funcnum_non_numeric_error", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 2, Row: 1}: StringVal("abc"),
			},
		}
		// SUBTOTAL("abc", A1:A1) → #VALUE!
		cf := evalCompile(t, "SUBTOTAL(B1,A1:A1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got type=%v err=%v, want #VALUE!", got.Type, got.Err)
		}
	})

	// ---------------------------------------------------------------
	// 9. Cross-check: SUBTOTAL(9, range) = SUM(range) for simple ranges
	// ---------------------------------------------------------------
	t.Run("crosscheck_sum", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(7),
				{Col: 1, Row: 2}: NumberVal(13),
				{Col: 1, Row: 3}: NumberVal(22),
			},
		}
		cfSub := evalCompile(t, "SUBTOTAL(9,A1:A3)")
		cfSum := evalCompile(t, "SUM(A1:A3)")
		gotSub, err := Eval(cfSub, resolver, nil)
		if err != nil {
			t.Fatalf("SUBTOTAL error: %v", err)
		}
		gotSum, err := Eval(cfSum, resolver, nil)
		if err != nil {
			t.Fatalf("SUM error: %v", err)
		}
		if gotSub.Num != gotSum.Num {
			t.Errorf("SUBTOTAL(9)=%g != SUM=%g", gotSub.Num, gotSum.Num)
		}
	})

	// ---------------------------------------------------------------
	// 10. Cross-check: SUBTOTAL(1, range) = AVERAGE(range) for simple ranges
	// ---------------------------------------------------------------
	t.Run("crosscheck_average", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(7),
				{Col: 1, Row: 2}: NumberVal(13),
				{Col: 1, Row: 3}: NumberVal(22),
			},
		}
		cfSub := evalCompile(t, "SUBTOTAL(1,A1:A3)")
		cfAvg := evalCompile(t, "AVERAGE(A1:A3)")
		gotSub, err := Eval(cfSub, resolver, nil)
		if err != nil {
			t.Fatalf("SUBTOTAL error: %v", err)
		}
		gotAvg, err := Eval(cfAvg, resolver, nil)
		if err != nil {
			t.Fatalf("AVERAGE error: %v", err)
		}
		if gotSub.Num != gotAvg.Num {
			t.Errorf("SUBTOTAL(1)=%g != AVERAGE=%g", gotSub.Num, gotAvg.Num)
		}
	})

	// ---------------------------------------------------------------
	// 11. Cross-check all 11 modes against their direct functions
	// ---------------------------------------------------------------
	t.Run("crosscheck_all_modes", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(4),
				{Col: 1, Row: 2}: NumberVal(8),
				{Col: 1, Row: 3}: NumberVal(15),
				{Col: 1, Row: 4}: NumberVal(16),
				{Col: 1, Row: 5}: NumberVal(23),
			},
		}
		crossChecks := []struct {
			mode    int
			directF string
		}{
			{1, "AVERAGE(A1:A5)"},
			{2, "COUNT(A1:A5)"},
			{3, "COUNTA(A1:A5)"},
			{4, "MAX(A1:A5)"},
			{5, "MIN(A1:A5)"},
			{6, "PRODUCT(A1:A5)"},
			{7, "STDEV(A1:A5)"},
			{8, "STDEVP(A1:A5)"},
			{9, "SUM(A1:A5)"},
			{10, "VAR(A1:A5)"},
			{11, "VARP(A1:A5)"},
		}
		for _, cc := range crossChecks {
			subtotalFormula := fmt.Sprintf("SUBTOTAL(%d,A1:A5)", cc.mode)
			cfSub := evalCompile(t, subtotalFormula)
			cfDirect := evalCompile(t, cc.directF)
			gotSub, err := Eval(cfSub, resolver, nil)
			if err != nil {
				t.Fatalf("SUBTOTAL(%d) error: %v", cc.mode, err)
			}
			gotDirect, err := Eval(cfDirect, resolver, nil)
			if err != nil {
				t.Fatalf("%s error: %v", cc.directF, err)
			}
			if gotSub.Type != gotDirect.Type {
				t.Errorf("SUBTOTAL(%d) type=%v != %s type=%v",
					cc.mode, gotSub.Type, cc.directF, gotDirect.Type)
				continue
			}
			if gotSub.Type == ValueNumber && math.Abs(gotSub.Num-gotDirect.Num) > 1e-10 {
				t.Errorf("SUBTOTAL(%d)=%g != %s=%g",
					cc.mode, gotSub.Num, cc.directF, gotDirect.Num)
			}
		}
	})

	// ---------------------------------------------------------------
	// 12. Additional invalid function_num boundary tests
	// ---------------------------------------------------------------
	t.Run("invalid_funcnum_50", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
			},
		}
		cf := evalCompile(t, "SUBTOTAL(50,A1:A1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got type=%v err=%v, want #VALUE!", got.Type, got.Err)
		}
	})

	t.Run("invalid_funcnum_float_truncated", func(t *testing.T) {
		// 9.9 should be treated as 9 (truncated to int) → SUM
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
			},
		}
		cf := evalCompile(t, "SUBTOTAL(9.9,A1:A2)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 30 {
			t.Errorf("got %v (%g), want 30", got.Type, got.Num)
		}
	})

	// ---------------------------------------------------------------
	// 13. Combined: hidden rows + subtotal cells + 101-111 mode
	// ---------------------------------------------------------------
	t.Run("hidden_and_subtotal_combined", func(t *testing.T) {
		// A1=10, A2=20 (hidden), A3=SUBTOTAL(value 100), A4=40
		// SUBTOTAL(109, A1:A4) should skip row 2 (hidden) and A3 (subtotal) → 10+40=50
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(100),
				{Col: 1, Row: 4}: NumberVal(40),
			},
			subtotalCells: map[CellAddr]bool{
				{Col: 1, Row: 3}: true,
			},
			hiddenRows: map[string]map[int]bool{
				"": {2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(109,A1:A4)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 50 {
			t.Errorf("got %v (%g), want 50", got.Type, got.Num)
		}
	})

	// ---------------------------------------------------------------
	// 14. Product of range with single value
	// ---------------------------------------------------------------
	t.Run("product_single_value", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),
			},
		}
		cf := evalCompile(t, "SUBTOTAL(6,A1:A1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 42 {
			t.Errorf("got %v (%g), want 42", got.Type, got.Num)
		}
	})

	// ---------------------------------------------------------------
	// 15. Large function_num (negative)
	// ---------------------------------------------------------------
	t.Run("large_negative_funcnum", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(1),
			},
		}
		cf := evalCompile(t, "SUBTOTAL(-100,A1:A1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got type=%v err=%v, want #VALUE!", got.Type, got.Err)
		}
	})

	// ---------------------------------------------------------------
	// 16. All rows hidden → empty result behaviors
	// ---------------------------------------------------------------
	t.Run("all_rows_hidden_sum_109", func(t *testing.T) {
		// All rows hidden, SUBTOTAL(109,...) → SUM of nothing = 0
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
			},
			hiddenRows: map[string]map[int]bool{
				"": {1: true, 2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(109,A1:A2)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf("got %v (%g), want 0", got.Type, got.Num)
		}
	})

	t.Run("all_rows_hidden_average_101_div0", func(t *testing.T) {
		// All rows hidden, SUBTOTAL(101,...) → AVERAGE of nothing = #DIV/0!
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
			},
			hiddenRows: map[string]map[int]bool{
				"": {1: true, 2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(101,A1:A2)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got type=%v err=%v, want #DIV/0!", got.Type, got.Err)
		}
	})

	// ---------------------------------------------------------------
	// 17. Stdev/Var with hidden rows
	// ---------------------------------------------------------------
	t.Run("hidden_rows_stdev_107", func(t *testing.T) {
		// A1=2, A2=4 (hidden), A3=6, A4=8
		// SUBTOTAL(107, A1:A4) should skip row 2 → STDEV.S(2,6,8) = sqrt(((2-16/3)^2+(6-16/3)^2+(8-16/3)^2)/2)
		// Mean=16/3, deviations: -10/3, 2/3, 8/3; sum_sq = 100/9+4/9+64/9 = 168/9; var = 168/18 = 28/3
		// stdev = sqrt(28/3) ≈ 3.05505
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(2),
				{Col: 1, Row: 2}: NumberVal(4),
				{Col: 1, Row: 3}: NumberVal(6),
				{Col: 1, Row: 4}: NumberVal(8),
			},
			hiddenRows: map[string]map[int]bool{
				"": {2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(107,A1:A4)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		expected := math.Sqrt(28.0 / 3.0)
		if got.Type != ValueNumber || math.Abs(got.Num-expected) > 1e-10 {
			t.Errorf("got %v (%g), want %g", got.Type, got.Num, expected)
		}
	})

	t.Run("hidden_rows_var_110", func(t *testing.T) {
		// A1=2, A2=4 (hidden), A3=6, A4=8
		// SUBTOTAL(110, A1:A4) should skip row 2 → VAR.S(2,6,8) = 28/3 ≈ 9.33333
		resolver := &subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(2),
				{Col: 1, Row: 2}: NumberVal(4),
				{Col: 1, Row: 3}: NumberVal(6),
				{Col: 1, Row: 4}: NumberVal(8),
			},
			hiddenRows: map[string]map[int]bool{
				"": {2: true},
			},
		}
		ctx := &EvalContext{Resolver: resolver}
		cf := evalCompile(t, "SUBTOTAL(110,A1:A4)")
		got, err := Eval(cf, resolver, ctx)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		expected := 28.0 / 3.0
		if got.Type != ValueNumber || math.Abs(got.Num-expected) > 1e-10 {
			t.Errorf("got %v (%g), want %g", got.Type, got.Num, expected)
		}
	})
}

func TestSubtotalFilterArgsPreservesTrimmedRangeOrigin(t *testing.T) {
	arg := trimmedRangeValue([][]Value{
		{NumberVal(10)},
	}, 1, 1, 1, 3)
	ctx := &EvalContext{
		Resolver: &subtotalMockResolver{
			autoFilterRows: map[string]map[int]bool{
				"": {1: true},
			},
		},
	}

	gotArgs := subtotalFilterArgs([]Value{arg}, ctx, false)
	if len(gotArgs) != 1 {
		t.Fatalf("len(filtered) = %d, want 1", len(gotArgs))
	}
	got := gotArgs[0]
	if got.RangeOrigin == nil || got.RangeOrigin.FromRow != 1 || got.RangeOrigin.ToRow != 3 {
		t.Fatalf("RangeOrigin = %+v, want rows 1:3", got.RangeOrigin)
	}
	rows, cols := arrayOpBounds(got)
	if rows != 3 || cols != 1 {
		t.Fatalf("arrayOpBounds(filtered) = %dx%d, want 3x1", rows, cols)
	}
	for row := 0; row < rows; row++ {
		if cell := arrayElementDirect(got, rows, cols, row, 0); cell.Type != ValueEmpty {
			t.Fatalf("filtered[%d,0] = %#v, want empty", row, cell)
		}
	}
}

func TestSUBTOTALFullColumnSparseRefUsesLiveGrid(t *testing.T) {
	resolver := &subtotalSparseResolver{
		subtotalMockResolver: subtotalMockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
				{Col: 1, Row: 4}: NumberVal(40),
			},
			subtotalCells: map[CellAddr]bool{
				{Col: 1, Row: 4}: true,
			},
			hiddenRows: map[string]map[int]bool{
				"": {2: true},
			},
		},
	}
	ctx := &EvalContext{Resolver: resolver}
	cf := evalCompile(t, "SUBTOTAL(109,A:A)")
	got, err := Eval(cf, resolver, ctx)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 40 {
		t.Fatalf("SUBTOTAL(109,A:A) = %v (%g), want 40", got.Type, got.Num)
	}
}

func TestISOCEILING(t *testing.T) {
	resolver := &mockResolver{}

	numTests := []struct {
		name    string
		formula string
		wantNum float64
	}{
		// documented examples
		{"doc_iso_1", "ISO.CEILING(4.3)", 5},
		{"doc_iso_2", "ISO.CEILING(-4.3)", -4},
		{"doc_iso_3", "ISO.CEILING(4.3,2)", 6},
		{"doc_iso_4", "ISO.CEILING(4.3,-2)", 6},
		{"doc_iso_5", "ISO.CEILING(-4.3,2)", -4},
		{"doc_iso_6", "ISO.CEILING(-4.3,-2)", -4},

		// Default significance (omitted = 1)
		{"pos_default", "ISO.CEILING(4.3)", 5},
		{"neg_default", "ISO.CEILING(-4.3)", -4},
		{"pos_exact_default", "ISO.CEILING(7)", 7},

		// Zero number returns 0
		{"zero_number", "ISO.CEILING(0)", 0},
		{"zero_with_sig", "ISO.CEILING(0,5)", 0},
		{"zero_with_neg_sig", "ISO.CEILING(0,-3)", 0},

		// Zero significance returns 0
		{"sig_zero_pos", "ISO.CEILING(6.3,0)", 0},
		{"sig_zero_neg", "ISO.CEILING(-6.3,0)", 0},

		// Positive number, positive significance
		{"pos_sig_2", "ISO.CEILING(4.3,2)", 6},
		{"pos_sig_5", "ISO.CEILING(24.3,5)", 25},
		{"pos_sig_3", "ISO.CEILING(7,3)", 9},
		{"pos_sig_exact", "ISO.CEILING(6,3)", 6},

		// Negative number, positive significance (rounds toward +inf)
		{"neg_sig_pos_2", "ISO.CEILING(-4.3,2)", -4},
		{"neg_sig_pos_1", "ISO.CEILING(-4.1,1)", -4},
		{"neg_sig_pos_5", "ISO.CEILING(-8.1,5)", -5},

		// Positive number, negative significance (uses abs, same as positive sig)
		{"pos_neg_sig_2", "ISO.CEILING(4.3,-2)", 6},
		{"pos_neg_sig_5", "ISO.CEILING(24.3,-5)", 25},

		// Negative number, negative significance (uses abs)
		{"neg_neg_sig_2", "ISO.CEILING(-4.3,-2)", -4},
		{"neg_neg_sig_5", "ISO.CEILING(-8.1,-5)", -5},

		// Fractional significance
		{"frac_sig_0.1", "ISO.CEILING(6.31,0.1)", 6.4},
		{"frac_sig_0.5", "ISO.CEILING(6.3,0.5)", 6.5},
		{"frac_sig_0.1_neg", "ISO.CEILING(-6.31,0.1)", -6.3},
		{"frac_sig_0.5_neg", "ISO.CEILING(-6.3,0.5)", -6},

		// Large numbers
		{"large_pos", "ISO.CEILING(1234567,1000)", 1235000},
		{"large_neg", "ISO.CEILING(-1234567,1000)", -1234000},

		// Very small numbers
		{"small_pos", "ISO.CEILING(0.001,0.01)", 0.01},
		{"small_neg", "ISO.CEILING(-0.001,0.01)", 0},
		{"small_frac", "ISO.CEILING(0.1)", 1},
		{"neg_small_frac", "ISO.CEILING(-0.1)", 0},

		// Already-rounded numbers (exact multiples)
		{"exact_multiple_pos", "ISO.CEILING(6,3)", 6},
		{"exact_multiple_neg", "ISO.CEILING(-6,3)", -6},
		{"exact_multiple_neg_sig", "ISO.CEILING(6,-3)", 6},

		// String coercion of numeric strings
		{"string_num", "ISO.CEILING(\"4.3\")", 5},
		{"string_num_neg", "ISO.CEILING(\"-4.3\")", -4},
		{"string_sig", "ISO.CEILING(4.3,\"2\")", 6},

		// Boolean coercion (TRUE = 1, FALSE = 0)
		{"bool_true", "ISO.CEILING(TRUE)", 1},
		{"bool_false", "ISO.CEILING(FALSE)", 0},
		{"bool_true_sig", "ISO.CEILING(2.3,TRUE)", 3},
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
		{"no_args", "ISO.CEILING()", ErrValVALUE},
		{"too_many_args", "ISO.CEILING(1,2,3)", ErrValVALUE},
		{"non_numeric", "ISO.CEILING(\"abc\")", ErrValVALUE},
		{"non_numeric_sig", "ISO.CEILING(1,\"abc\")", ErrValVALUE},
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
// RANDARRAY tests
// ---------------------------------------------------------------------------

func TestRANDARRAY_NoArgs(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY()")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber {
		t.Fatalf("RANDARRAY() type = %v, want number", got.Type)
	}
	if got.Num < 0 || got.Num >= 1 {
		t.Errorf("RANDARRAY() = %g, want [0,1)", got.Num)
	}
}

func TestRANDARRAY_3x2(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(3,2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("RANDARRAY(3,2) type = %v, want array", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("RANDARRAY(3,2) rows = %d, want 3", len(got.Array))
	}
	for r, row := range got.Array {
		if len(row) != 2 {
			t.Fatalf("RANDARRAY(3,2) row %d cols = %d, want 2", r, len(row))
		}
		for c, v := range row {
			if v.Type != ValueNumber || v.Num < 0 || v.Num >= 1 {
				t.Errorf("RANDARRAY(3,2)[%d][%d] = %g, want [0,1)", r, c, v.Num)
			}
		}
	}
}

func TestRANDARRAY_1x1Scalar(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(1,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber {
		t.Fatalf("RANDARRAY(1,1) type = %v, want number (scalar)", got.Type)
	}
	if got.Num < 0 || got.Num >= 1 {
		t.Errorf("RANDARRAY(1,1) = %g, want [0,1)", got.Num)
	}
}

func TestRANDARRAY_CustomMinMax(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(2,3,10,20)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("type = %v, want array", got.Type)
	}
	if len(got.Array) != 2 {
		t.Fatalf("rows = %d, want 2", len(got.Array))
	}
	for r, row := range got.Array {
		if len(row) != 3 {
			t.Fatalf("row %d cols = %d, want 3", r, len(row))
		}
		for c, v := range row {
			if v.Type != ValueNumber || v.Num < 10 || v.Num > 20 {
				t.Errorf("[%d][%d] = %g, want [10,20]", r, c, v.Num)
			}
		}
	}
}

func TestRANDARRAY_WholeNumberTrue(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(3,3,1,10,TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("type = %v, want array", got.Type)
	}
	for r, row := range got.Array {
		for c, v := range row {
			if v.Type != ValueNumber {
				t.Errorf("[%d][%d] type = %v, want number", r, c, v.Type)
				continue
			}
			if v.Num != math.Floor(v.Num) {
				t.Errorf("[%d][%d] = %g, want integer", r, c, v.Num)
			}
			if v.Num < 1 || v.Num > 10 {
				t.Errorf("[%d][%d] = %g, want [1,10]", r, c, v.Num)
			}
		}
	}
}

func TestRANDARRAY_WholeNumberFalse(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(2,2,0,100,FALSE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("type = %v, want array", got.Type)
	}
	for r, row := range got.Array {
		for c, v := range row {
			if v.Type != ValueNumber || v.Num < 0 || v.Num > 100 {
				t.Errorf("[%d][%d] = %g, want [0,100]", r, c, v.Num)
			}
		}
	}
}

func TestRANDARRAY_MinGreaterThanMax(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(1,1,10,5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("RANDARRAY(1,1,10,5) = type=%v err=%v, want #VALUE!", got.Type, got.Err)
	}
}

func TestRANDARRAY_ZeroRows(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(0,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("RANDARRAY(0,1) = type=%v err=%v, want #CALC!", got.Type, got.Err)
	}
}

func TestRANDARRAY_ZeroCols(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(1,0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("RANDARRAY(1,0) = type=%v err=%v, want #CALC!", got.Type, got.Err)
	}
}

func TestRANDARRAY_NegativeRows(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(-3,2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("RANDARRAY(-3,2) = type=%v err=%v, want #CALC!", got.Type, got.Err)
	}
}

func TestRANDARRAY_NegativeCols(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(2,-1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("RANDARRAY(2,-1) = type=%v err=%v, want #CALC!", got.Type, got.Err)
	}
}

func TestRANDARRAY_WholeNoValidIntegers(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(1,1,1.5,1.7,TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("RANDARRAY(1,1,1.5,1.7,TRUE) = type=%v err=%v, want #VALUE!", got.Type, got.Err)
	}
}

func TestRANDARRAY_DefaultMinMaxIntegerMode(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(2,2,,,TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("type = %v, want array", got.Type)
	}
	for r, row := range got.Array {
		for c, v := range row {
			if v.Type != ValueNumber {
				t.Errorf("[%d][%d] type = %v, want number", r, c, v.Type)
				continue
			}
			// With default min=0, max=1 and whole=TRUE, valid integers are 0 and 1
			if v.Num != 0 && v.Num != 1 {
				t.Errorf("[%d][%d] = %g, want 0 or 1", r, c, v.Num)
			}
		}
	}
}

func TestRANDARRAY_LargeArray10x10(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(10,10,0,100)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("type = %v, want array", got.Type)
	}
	if len(got.Array) != 10 {
		t.Fatalf("rows = %d, want 10", len(got.Array))
	}
	for r, row := range got.Array {
		if len(row) != 10 {
			t.Fatalf("row %d cols = %d, want 10", r, len(row))
		}
		for c, v := range row {
			if v.Type != ValueNumber || v.Num < 0 || v.Num > 100 {
				t.Errorf("[%d][%d] = %g, want [0,100]", r, c, v.Num)
			}
		}
	}
}

func TestRANDARRAY_SingleRowMultipleCols(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(1,5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("type = %v, want array", got.Type)
	}
	if len(got.Array) != 1 {
		t.Fatalf("rows = %d, want 1", len(got.Array))
	}
	if len(got.Array[0]) != 5 {
		t.Fatalf("cols = %d, want 5", len(got.Array[0]))
	}
}

func TestRANDARRAY_MultipleRowsSingleCol(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(4,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("type = %v, want array", got.Type)
	}
	if len(got.Array) != 4 {
		t.Fatalf("rows = %d, want 4", len(got.Array))
	}
	for r, row := range got.Array {
		if len(row) != 1 {
			t.Fatalf("row %d cols = %d, want 1", r, len(row))
		}
	}
}

func TestRANDARRAY_StringCoercion(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `RANDARRAY("2","3","5","15")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("type = %v, want array", got.Type)
	}
	if len(got.Array) != 2 {
		t.Fatalf("rows = %d, want 2", len(got.Array))
	}
	for r, row := range got.Array {
		if len(row) != 3 {
			t.Fatalf("row %d cols = %d, want 3", r, len(row))
		}
		for c, v := range row {
			if v.Type != ValueNumber || v.Num < 5 || v.Num > 15 {
				t.Errorf("[%d][%d] = %g, want [5,15]", r, c, v.Num)
			}
		}
	}
}

func TestRANDARRAY_BooleanCoercion(t *testing.T) {
	// TRUE coerces to 1, so RANDARRAY(TRUE,TRUE) = RANDARRAY(1,1) = scalar
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(TRUE,TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber {
		t.Fatalf("type = %v, want number (scalar)", got.Type)
	}
	if got.Num < 0 || got.Num >= 1 {
		t.Errorf("RANDARRAY(TRUE,TRUE) = %g, want [0,1)", got.Num)
	}
}

func TestRANDARRAY_ErrorTooManyArgs(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(1,1,0,1,FALSE,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("RANDARRAY(6 args) = type=%v err=%v, want #VALUE!", got.Type, got.Err)
	}
}

func TestRANDARRAY_ErrorNonNumericRows(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `RANDARRAY("abc")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("RANDARRAY(\"abc\") type = %v, want error", got.Type)
	}
}

func TestRANDARRAY_WholeScalar(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(1,1,5,10,TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber {
		t.Fatalf("type = %v, want number (scalar)", got.Type)
	}
	if got.Num != math.Floor(got.Num) {
		t.Errorf("got %g, want integer", got.Num)
	}
	if got.Num < 5 || got.Num > 10 {
		t.Errorf("got %g, want [5,10]", got.Num)
	}
}

func TestRANDARRAY_RowsTruncated(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(2.9,1.8)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("type = %v, want array", got.Type)
	}
	// 2.9 truncated to 2 rows, 1.8 truncated to 1 col
	if len(got.Array) != 2 {
		t.Fatalf("rows = %d, want 2", len(got.Array))
	}
	if len(got.Array[0]) != 1 {
		t.Fatalf("cols = %d, want 1", len(got.Array[0]))
	}
}

func TestRANDARRAY_MinEqualsMax(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(2,2,7,7)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("type = %v, want array", got.Type)
	}
	for r, row := range got.Array {
		for c, v := range row {
			if v.Type != ValueNumber || v.Num != 7 {
				t.Errorf("[%d][%d] = %g, want 7", r, c, v.Num)
			}
		}
	}
}

func TestRANDARRAY_NegativeRange(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(2,2,-10,-5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("type = %v, want array", got.Type)
	}
	for r, row := range got.Array {
		for c, v := range row {
			if v.Type != ValueNumber || v.Num < -10 || v.Num > -5 {
				t.Errorf("[%d][%d] = %g, want [-10,-5]", r, c, v.Num)
			}
		}
	}
}

func TestRANDARRAY_OnlyRowsArg(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, "RANDARRAY(3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("type = %v, want array", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("rows = %d, want 3", len(got.Array))
	}
	for r, row := range got.Array {
		if len(row) != 1 {
			t.Fatalf("row %d cols = %d, want 1 (default)", r, len(row))
		}
		if row[0].Type != ValueNumber || row[0].Num < 0 || row[0].Num >= 1 {
			t.Errorf("[%d][0] = %g, want [0,1)", r, row[0].Num)
		}
	}
}

// ---------------------------------------------------------------------------
// SEQUENCE – additional comprehensive tests
// ---------------------------------------------------------------------------

func TestSEQUENCE_Basic5Rows(t *testing.T) {
	// SEQUENCE(5) = {1;2;3;4;5}
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 5 {
		t.Fatalf("type=%v rows=%d, want 5-row array", got.Type, len(got.Array))
	}
	for i := 0; i < 5; i++ {
		if len(got.Array[i]) != 1 {
			t.Fatalf("row %d cols = %d, want 1", i, len(got.Array[i]))
		}
		want := float64(i + 1)
		if got.Array[i][0].Num != want {
			t.Errorf("[%d][0] = %g, want %g", i, got.Array[i][0].Num, want)
		}
	}
}

func TestSEQUENCE_3x3Matrix(t *testing.T) {
	// SEQUENCE(3,3) = 3x3 matrix: {1,2,3;4,5,6;7,8,9}
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(3,3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("type=%v rows=%d, want 3-row array", got.Type, len(got.Array))
	}
	n := 1.0
	for r := 0; r < 3; r++ {
		if len(got.Array[r]) != 3 {
			t.Fatalf("row %d cols = %d, want 3", r, len(got.Array[r]))
		}
		for c := 0; c < 3; c++ {
			if got.Array[r][c].Num != n {
				t.Errorf("[%d][%d] = %g, want %g", r, c, got.Array[r][c].Num, n)
			}
			n++
		}
	}
}

func TestSEQUENCE_CustomStartOnly(t *testing.T) {
	// SEQUENCE(5,1,10) = {10;11;12;13;14}
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(5,1,10)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 5 {
		t.Fatalf("got type=%v rows=%d, want 5 rows", got.Type, len(got.Array))
	}
	want := []float64{10, 11, 12, 13, 14}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestSEQUENCE_CustomStepEven(t *testing.T) {
	// SEQUENCE(5,1,0,2) = {0;2;4;6;8}
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(5,1,0,2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 5 {
		t.Fatalf("got type=%v rows=%d, want 5 rows", got.Type, len(got.Array))
	}
	want := []float64{0, 2, 4, 6, 8}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestSEQUENCE_NegativeStepDown(t *testing.T) {
	// SEQUENCE(5,1,10,-2) = {10;8;6;4;2}
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(5,1,10,-2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 5 {
		t.Fatalf("got type=%v rows=%d, want 5 rows", got.Type, len(got.Array))
	}
	want := []float64{10, 8, 6, 4, 2}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestSEQUENCE_WideRow10Cols(t *testing.T) {
	// SEQUENCE(1,10) = row of 1-10
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(1,10)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 10 {
		t.Fatalf("got type=%v, want 1x10 array", got.Type)
	}
	for c := 0; c < 10; c++ {
		want := float64(c + 1)
		if got.Array[0][c].Num != want {
			t.Errorf("[0][%d] = %g, want %g", c, got.Array[0][c].Num, want)
		}
	}
}

func TestSEQUENCE_FractionalStartAndStep(t *testing.T) {
	// SEQUENCE(4,1,0.5,0.25) = {0.5;0.75;1.0;1.25}
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(4,1,0.5,0.25)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 4 {
		t.Fatalf("got type=%v rows=%d, want 4 rows", got.Type, len(got.Array))
	}
	want := []float64{0.5, 0.75, 1.0, 1.25}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestSEQUENCE_Large100Rows(t *testing.T) {
	// SEQUENCE(100,1) = 100 rows, 1 to 100
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(100,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 100 {
		t.Fatalf("got type=%v rows=%d, want 100 rows", got.Type, len(got.Array))
	}
	for i := 0; i < 100; i++ {
		want := float64(i + 1)
		if got.Array[i][0].Num != want {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, want)
		}
	}
}

func TestSEQUENCE_NegativeColsError(t *testing.T) {
	// SEQUENCE(3,-1) => #CALC!
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(3,-1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValCALC {
		t.Errorf("got type=%v err=%v, want #CALC!", got.Type, got.Err)
	}
}

func TestSEQUENCE_ColsTruncated(t *testing.T) {
	// SEQUENCE(1,3.7) should truncate cols to 3
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(1,3.7)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("got type=%v, want 1x3 array", got.Type)
	}
	for c := 0; c < 3; c++ {
		want := float64(c + 1)
		if got.Array[0][c].Num != want {
			t.Errorf("[0][%d] = %g, want %g", c, got.Array[0][c].Num, want)
		}
	}
}

func TestSEQUENCE_NegativeStartLargeStep(t *testing.T) {
	// SEQUENCE(4,1,-5,3) = {-5;-2;1;4}
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(4,1,-5,3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 4 {
		t.Fatalf("got type=%v rows=%d, want 4 rows", got.Type, len(got.Array))
	}
	want := []float64{-5, -2, 1, 4}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestSEQUENCE_2x3Custom(t *testing.T) {
	// SEQUENCE(2,3,10,10) = {10,20,30;40,50,60}
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(2,3,10,10)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("got type=%v rows=%d, want 2 rows", got.Type, len(got.Array))
	}
	want := [][]float64{{10, 20, 30}, {40, 50, 60}}
	for r := 0; r < 2; r++ {
		if len(got.Array[r]) != 3 {
			t.Fatalf("row %d cols = %d, want 3", r, len(got.Array[r]))
		}
		for c := 0; c < 3; c++ {
			if got.Array[r][c].Num != want[r][c] {
				t.Errorf("[%d][%d] = %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestSEQUENCE_LargeStep1000(t *testing.T) {
	// SEQUENCE(3,1,100,1000) = {100;1100;2100}
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(3,1,100,1000)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("got type=%v rows=%d, want 3 rows", got.Type, len(got.Array))
	}
	want := []float64{100, 1100, 2100}
	for i, w := range want {
		if got.Array[i][0].Num != w {
			t.Errorf("[%d]: got %g, want %g", i, got.Array[i][0].Num, w)
		}
	}
}

func TestSEQUENCE_BoolCoercion(t *testing.T) {
	// SEQUENCE(TRUE) should coerce TRUE to 1 => scalar 1
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(TRUE)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("got %v, want scalar 1", got)
	}
}

func TestSEQUENCE_2x2Matrix(t *testing.T) {
	// SEQUENCE(2,2) = {1,2;3,4}
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(2,2)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("got type=%v rows=%d, want 2 rows", got.Type, len(got.Array))
	}
	want := [][]float64{{1, 2}, {3, 4}}
	for r := 0; r < 2; r++ {
		if len(got.Array[r]) != 2 {
			t.Fatalf("row %d cols = %d, want 2", r, len(got.Array[r]))
		}
		for c := 0; c < 2; c++ {
			if got.Array[r][c].Num != want[r][c] {
				t.Errorf("[%d][%d] = %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestSEQUENCE_ErrorPropagationFromCell(t *testing.T) {
	// SEQUENCE with a cell that resolves to an error
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: ErrorVal(ErrValNA),
		},
	}
	cf := evalCompile(t, "SEQUENCE(A1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("got type=%v err=%v, want #N/A", got.Type, got.Err)
	}
}

func TestSEQUENCE_DirectCallFn(t *testing.T) {
	// Test fnSEQUENCE directly with all four arguments
	got, err := fnSEQUENCE([]Value{NumberVal(3), NumberVal(2), NumberVal(5), NumberVal(3)})
	if err != nil {
		t.Fatalf("fnSEQUENCE: %v", err)
	}
	// 3x2 starting at 5, step 3: {5,8;11,14;17,20}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("got type=%v rows=%d, want 3 rows", got.Type, len(got.Array))
	}
	want := [][]float64{{5, 8}, {11, 14}, {17, 20}}
	for r := 0; r < 3; r++ {
		if len(got.Array[r]) != 2 {
			t.Fatalf("row %d cols = %d, want 2", r, len(got.Array[r]))
		}
		for c := 0; c < 2; c++ {
			if got.Array[r][c].Num != want[r][c] {
				t.Errorf("[%d][%d] = %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestSEQUENCE_2DZeroStepFill(t *testing.T) {
	// SEQUENCE(2,3,7,0) => all cells = 7
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(2,3,7,0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("got type=%v rows=%d, want 2 rows", got.Type, len(got.Array))
	}
	for r := 0; r < 2; r++ {
		if len(got.Array[r]) != 3 {
			t.Fatalf("row %d cols = %d, want 3", r, len(got.Array[r]))
		}
		for c := 0; c < 3; c++ {
			if got.Array[r][c].Num != 7 {
				t.Errorf("[%d][%d] = %g, want 7", r, c, got.Array[r][c].Num)
			}
		}
	}
}

func TestSEQUENCE_2DNegativeStepCountdown(t *testing.T) {
	// SEQUENCE(2,3,100,-10) = {100,90,80;70,60,50}
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(2,3,100,-10)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 2 {
		t.Fatalf("got type=%v rows=%d, want 2 rows", got.Type, len(got.Array))
	}
	want := [][]float64{{100, 90, 80}, {70, 60, 50}}
	for r := 0; r < 2; r++ {
		for c := 0; c < 3; c++ {
			if got.Array[r][c].Num != want[r][c] {
				t.Errorf("[%d][%d] = %g, want %g", r, c, got.Array[r][c].Num, want[r][c])
			}
		}
	}
}

func TestSEQUENCE_SumComposition(t *testing.T) {
	// SUM(SEQUENCE(5)) = 1+2+3+4+5 = 15
	resolver := &mockResolver{}
	cf := evalCompile(t, "SUM(SEQUENCE(5))")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 15 {
		t.Errorf("SUM(SEQUENCE(5)) = %v, want 15", got)
	}
}

func TestSEQUENCE_ErrorInStepArg(t *testing.T) {
	// Error in step argument should propagate
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: ErrorVal(ErrValDIV0),
		},
	}
	cf := evalCompile(t, "SEQUENCE(3,1,1,A1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("got type=%v err=%v, want #DIV/0!", got.Type, got.Err)
	}
}

func TestSEQUENCE_ErrorInColsArg(t *testing.T) {
	// Error in cols argument should propagate
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: ErrorVal(ErrValREF),
		},
	}
	cf := evalCompile(t, "SEQUENCE(3,A1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValREF {
		t.Errorf("got type=%v err=%v, want #REF!", got.Type, got.Err)
	}
}

func TestSEQUENCE_ErrorInStartArg(t *testing.T) {
	// Error in start argument should propagate
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: ErrorVal(ErrValNUM),
		},
	}
	cf := evalCompile(t, "SEQUENCE(3,1,A1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("got type=%v err=%v, want #NUM!", got.Type, got.Err)
	}
}

func TestSEQUENCE_4x4Identity(t *testing.T) {
	// SEQUENCE(4,4) = 4x4 matrix 1-16
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(4,4)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 4 {
		t.Fatalf("got type=%v rows=%d, want 4 rows", got.Type, len(got.Array))
	}
	n := 1.0
	for r := 0; r < 4; r++ {
		if len(got.Array[r]) != 4 {
			t.Fatalf("row %d cols = %d, want 4", r, len(got.Array[r]))
		}
		for c := 0; c < 4; c++ {
			if got.Array[r][c].Num != n {
				t.Errorf("[%d][%d] = %g, want %g", r, c, got.Array[r][c].Num, n)
			}
			n++
		}
	}
}

func TestSEQUENCE_SingleRowOnlyRowsArg(t *testing.T) {
	// SEQUENCE(1) = scalar 1 (rows=1, cols=1 default => single cell)
	resolver := &mockResolver{}
	cf := evalCompile(t, "SEQUENCE(1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("got %v, want scalar 1", got)
	}
}

func TestSEQUENCE_StringCoercionNonNumericCols(t *testing.T) {
	// SEQUENCE(3,"abc") => #VALUE! (can't coerce cols)
	resolver := &mockResolver{}
	cf := evalCompile(t, `SEQUENCE(3,"abc")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("got type=%v err=%v, want #VALUE!", got.Type, got.Err)
	}
}

func TestSEQUENCE_MaxComposition(t *testing.T) {
	// MAX(SEQUENCE(5)) = 5
	resolver := &mockResolver{}
	cf := evalCompile(t, "MAX(SEQUENCE(5))")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 5 {
		t.Errorf("MAX(SEQUENCE(5)) = %v, want 5", got)
	}
}

func TestSEQUENCE_MinComposition(t *testing.T) {
	// MIN(SEQUENCE(5,1,10,-3)) = MIN(10,7,4,1,-2) = -2
	resolver := &mockResolver{}
	cf := evalCompile(t, "MIN(SEQUENCE(5,1,10,-3))")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != -2 {
		t.Errorf("MIN(SEQUENCE(5,1,10,-3)) = %v, want -2", got)
	}
}
