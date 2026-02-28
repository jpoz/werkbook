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
		{"ACOS(1)", 0, 1e-10},
		{"ACOS(0)", math.Pi / 2, 1e-10},
		{"ASIN(0)", 0, 1e-10},
		{"ASIN(1)", math.Pi / 2, 1e-10},
		{"ATAN(0)", 0, 1e-10},
		{"ATAN(1)", math.Pi / 4, 1e-10},
		{"ATAN2(1,1)", math.Pi / 4, 1e-10},
		{"COS(0)", 1, 1e-10},
		{"SIN(0)", 0, 1e-10},
		{"TAN(0)", 0, 1e-10},
		{"EXP(0)", 1, 1e-10},
		{"EXP(1)", math.E, 1e-10},
		{"LN(1)", 0, 1e-10},
		{`LN(2.718281828459045)`, 1, 1e-10},
		{"LOG(100)", 2, 1e-10},
		{"LOG(8,2)", 3, 1e-10},
		{"LOG10(1000)", 3, 1e-10},
		{"PI()", math.Pi, 1e-10},
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
		{"PRODUCT(2,3,4)", 24, 0},
		{"COMBIN(8,2)", 28, 0},
		{"COMBIN(5,0)", 1, 0},
		{"COMBIN(5,5)", 1, 0},
		{"COMBIN(10,3)", 120, 0},
		{"DEGREES(PI())", 180, 1e-10},
		{"EVEN(1.5)", 2, 0},
		{"EVEN(3)", 4, 0},
		{"EVEN(2)", 2, 0},
		{"EVEN(-1)", -2, 0},
		{"EVEN(-2)", -2, 0},
		{"EVEN(0)", 0, 0},
		{"DEGREES(0)", 0, 0},
		{"DEGREES(1)", 57.29577951308232, 1e-10},
		{"FACT(5)", 120, 0},
		{"FACT(0)", 1, 0},
		{"FACT(1)", 1, 0},
		{"FACT(1.9)", 1, 0},
		{"FACT(10)", 3628800, 0},
		{"GCD(5,2)", 1, 0},
		{"GCD(24,36)", 12, 0},
		{"GCD(7,1)", 1, 0},
		{"GCD(5,0)", 5, 0},
		{"GCD(0,0)", 0, 0},
		{"GCD(12,8,4)", 4, 0},
		{"LCM(5,2)", 10, 0},
		{"LCM(24,36)", 72, 0},
		{"LCM(3,4,5)", 60, 0},
		{"LCM(5,0)", 0, 0},
		{"LCM(7)", 7, 0},
		{"MROUND(10,3)", 9, 0},
		{"MROUND(-10,-3)", -9, 0},
		{"MROUND(1.3,0.2)", 1.4, 1e-10},
		{"MROUND(5,0)", 0, 0},
		{"MROUND(7.5,5)", 10, 0},
		{"ODD(1.5)", 3, 0},
		{"ODD(3)", 3, 0},
		{"ODD(2)", 3, 0},
		{"ODD(-1)", -1, 0},
		{"ODD(-2)", -3, 0},
		{"ODD(0)", 1, 0},
		{"QUOTIENT(5,2)", 2, 0},
		{"QUOTIENT(4.5,3.1)", 1, 0},
		{"QUOTIENT(-10,3)", -3, 0},
		{"QUOTIENT(7,7)", 1, 0},
		{"RADIANS(180)", math.Pi, 1e-10},
		{"RADIANS(0)", 0, 0},
		{"RADIANS(360)", 2 * math.Pi, 1e-10},
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
		{"ACOS(2)", ErrValNUM},
		{"ASIN(2)", ErrValNUM},
		{"LN(-1)", ErrValNUM},
		{"LOG(-1)", ErrValNUM},
		{"LOG(10,1)", ErrValNUM},
		{"SQRT(-1)", ErrValNUM},
		{"MOD(10,0)", ErrValDIV0},
		{"ATAN2(0,0)", ErrValDIV0},
		{"COMBIN(3,5)", ErrValNUM},
		{"COMBIN(-1,2)", ErrValNUM},
		{"FACT(-1)", ErrValNUM},
		{"GCD(-5,2)", ErrValNUM},
		{"LCM(-5,2)", ErrValNUM},
		{"MROUND(5,-2)", ErrValNUM},
		{"QUOTIENT(10,0)", ErrValDIV0},
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
