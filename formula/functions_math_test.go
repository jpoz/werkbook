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
