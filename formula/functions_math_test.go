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
		{"ACOSH(1)", 0, 1e-10},
		{"ACOSH(2)", math.Acosh(2), 1e-10},
		{"ACOSH(10)", math.Acosh(10), 1e-10},
		{"ASIN(0)", 0, 1e-10},
		{"ASIN(1)", math.Pi / 2, 1e-10},
		{"ASINH(0)", 0, 1e-10},
		{"ASINH(1)", math.Asinh(1), 1e-10},
		{"ASINH(-1)", math.Asinh(-1), 1e-10},
		{"ASINH(10)", math.Asinh(10), 1e-10},
		{"ATAN(0)", 0, 1e-10},
		{"ATAN(1)", math.Pi / 4, 1e-10},
		{"ATAN2(1,1)", math.Pi / 4, 1e-10},
		{"ATANH(0)", 0, 1e-10},
		{"ATANH(0.5)", math.Atanh(0.5), 1e-10},
		{"ATANH(-0.5)", math.Atanh(-0.5), 1e-10},
		{"ATANH(0.9)", math.Atanh(0.9), 1e-10},
		{"COS(0)", 1, 1e-10},
		{"COSH(0)", 1, 1e-10},
		{"COSH(1)", math.Cosh(1), 1e-10},
		{"COSH(-1)", math.Cosh(-1), 1e-10},
		{"COSH(10)", math.Cosh(10), 1e-10},
		{"COT(1)", 1 / math.Tan(1), 1e-10},
		{"COT(PI()/4)", 1, 1e-10},
		{"COTH(1)", 1 / math.Tanh(1), 1e-10},
		{"CSC(1)", 1 / math.Sin(1), 1e-10},
		{"CSC(PI()/2)", 1, 1e-10},
		{"CSCH(1)", 1 / math.Sinh(1), 1e-10},
		{"SEC(0)", 1, 1e-10},
		{"SEC(PI()/3)", 2, 1e-10},
		{"SECH(0)", 1, 1e-10},
		{"SECH(1)", 1 / math.Cosh(1), 1e-10},
		{"SIN(0)", 0, 1e-10},
		{"SINH(0)", 0, 1e-10},
		{"SINH(1)", math.Sinh(1), 1e-10},
		{"SINH(-1)", math.Sinh(-1), 1e-10},
		{"SINH(5)", math.Sinh(5), 1e-10},
		{"TAN(0)", 0, 1e-10},
		{"TANH(0)", 0, 1e-10},
		{"TANH(1)", math.Tanh(1), 1e-10},
		{"TANH(-1)", math.Tanh(-1), 1e-10},
		{"TANH(10)", math.Tanh(10), 1e-10},
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
		{"COMBINA(4,3)", 20, 0},
		{"COMBINA(10,3)", 220, 0},
		{"COMBINA(0,0)", 1, 0},
		{"COMBINA(5,0)", 1, 0},
		{"COMBINA(1,1)", 1, 0},
		{"COMBINA(5,2)", 15, 0},
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
		{"SQRTPI(1)", math.Sqrt(math.Pi), 1e-10},
		{"SQRTPI(2)", math.Sqrt(2 * math.Pi), 1e-10},
		{"SQRTPI(0)", 0, 0},
		{"SQRTPI(4)", math.Sqrt(4 * math.Pi), 1e-10},
		{"FACTDOUBLE(6)", 48, 0},
		{"FACTDOUBLE(7)", 105, 0},
		{"FACTDOUBLE(0)", 1, 0},
		{"FACTDOUBLE(1)", 1, 0},
		{"FACTDOUBLE(10)", 3840, 0},
		{"FACTDOUBLE(-1)", 1, 0},
		{`DECIMAL("FF",16)`, 255, 0},
		{`DECIMAL("111",2)`, 7, 0},
		{`DECIMAL("77",8)`, 63, 0},
		{`DECIMAL("ZZ",36)`, 1295, 0},
		{`DECIMAL("0",10)`, 0, 0},
		{`DECIMAL("10",10)`, 10, 0},
		{"MULTINOMIAL(2,3,4)", 1260, 0},
		{"MULTINOMIAL(1,2,3)", 60, 0},
		{"MULTINOMIAL(5)", 1, 0},
		{"MULTINOMIAL(0,0)", 1, 0},
		{"MULTINOMIAL(3,3)", 20, 0},
		{"MULTINOMIAL(2,2,2)", 90, 0},
		{`ARABIC("I")`, 1, 0},
		{`ARABIC("IV")`, 4, 0},
		{`ARABIC("IX")`, 9, 0},
		{`ARABIC("MCMXCIX")`, 1999, 0},
		{`ARABIC("MMXXVI")`, 2026, 0},
		{`ARABIC("XLII")`, 42, 0},
		{`ARABIC("CD")`, 400, 0},
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

func TestBASE(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		formula string
		want    string
	}{
		{`BASE(255,16)`, "FF"},
		{`BASE(7,2)`, "111"},
		{`BASE(255,16,4)`, "00FF"},
		{`BASE(10,10)`, "10"},
		{`BASE(0,2)`, "0"},
		{`BASE(100,8)`, "144"},
		{`BASE(15,16,1)`, "F"},
	}

	for _, tt := range strTests {
		t.Run(tt.formula, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString {
				t.Errorf("Eval(%q): got type %v, want string", tt.formula, got.Type)
			} else if got.Str != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
			}
		})
	}

	errTests := []struct {
		formula string
		errVal  ErrorValue
	}{
		{"BASE(-1,16)", ErrValNUM},
		{"BASE(255,1)", ErrValNUM},
		{"BASE(255,37)", ErrValNUM},
	}

	for _, tt := range errTests {
		t.Run(tt.formula, func(t *testing.T) {
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

func TestMathErrors(t *testing.T) {
	resolver := &mockResolver{}

	errTests := []struct {
		formula string
		errVal  ErrorValue
	}{
		{"ACOS(2)", ErrValNUM},
		{"ACOSH(0)", ErrValNUM},
		{"ACOSH(0.5)", ErrValNUM},
		{"ACOSH(-1)", ErrValNUM},
		{"ASIN(2)", ErrValNUM},
		{"COT(0)", ErrValDIV0},
		{"COTH(0)", ErrValDIV0},
		{"CSC(0)", ErrValDIV0},
		{"CSCH(0)", ErrValDIV0},
		{"LN(-1)", ErrValNUM},
		{"LOG(-1)", ErrValNUM},
		{"LOG(10,1)", ErrValNUM},
		{"SQRT(-1)", ErrValNUM},
		{"SQRTPI(-1)", ErrValNUM},
		{"MOD(10,0)", ErrValDIV0},
		{"MOD(100,-1e-15)", ErrValNUM},
		{"ATANH(1)", ErrValNUM},
		{"ATANH(-1)", ErrValNUM},
		{"ATANH(2)", ErrValNUM},
		{"ATAN2(0,0)", ErrValDIV0},
		{"COMBIN(3,5)", ErrValNUM},
		{"COMBIN(-1,2)", ErrValNUM},
		{"COMBINA(-1,2)", ErrValNUM},
		{"COMBINA(5,-1)", ErrValNUM},
		{"FACT(-1)", ErrValNUM},
		{"FACTDOUBLE(-2)", ErrValNUM},
		{"GCD(-5,2)", ErrValNUM},
		{"LCM(-5,2)", ErrValNUM},
		{"MROUND(5,-2)", ErrValNUM},
		{"QUOTIENT(10,0)", ErrValDIV0},
		// MOD precision overflow: |n/d| exceeds 2^53
		{"MOD(1e18,3)", ErrValNUM},
		{`DECIMAL("G",16)`, ErrValNUM},
		{`DECIMAL("FF",1)`, ErrValNUM},
		{"MULTINOMIAL(-1,2)", ErrValNUM},
		{`ARABIC("ABC")`, ErrValVALUE},
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

func TestPERMUT(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		wantNum float64
		wantErr bool
	}{
		{"PERMUT(100,3)", 970200, false},
		{"PERMUT(3,2)", 6, false},
		{"PERMUT(5,0)", 1, false},
		{"PERMUT(5,5)", 120, false},
		{"PERMUT(10,1)", 10, false},
		{"PERMUT(7,3)", 210, false},
		{"PERMUT(3,5)", 0, true},
		{"PERMUT(-1,2)", 0, true},
		{"PERMUT(0,0)", 0, true},
		// Pre-truncation negativity checks
		{"PERMUT(-1e-15,1)", 0, true},  // tiny negative n, int truncation would give 0
		{"PERMUT(-0.5,1)", 0, true},    // negative n before truncation
		{"PERMUT(5,-1e-15)", 0, true},  // tiny negative k
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if tt.wantErr {
				if got.Type != ValueError {
					t.Errorf("Eval(%q) = type %v, want ValueError", tt.formula, got.Type)
				}
			} else {
				if got.Type != ValueNumber {
					t.Errorf("Eval(%q) = type %v, want ValueNumber", tt.formula, got.Type)
				} else if got.Num != tt.wantNum {
					t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
				}
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
		// Division by zero → #DIV/0!
		{"div_by_zero", "MOD(10,0)", ErrValDIV0},
		// Precision overflow: |n/d| > 2^53 → #NUM!
		{"precision_overflow", "MOD(100,-1e-15)", ErrValNUM},
		{"precision_overflow_large", "MOD(1e18,1)", ErrValNUM},
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

func TestSUBTOTAL(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
		wantErr ErrorValue
	}{
		{"SUM", "SUBTOTAL(9,{120,10,150,23})", 303, 0, 0},
		{"SUM_100series", "SUBTOTAL(109,{120,10,150,23})", 303, 0, 0},
		{"AVERAGE", "SUBTOTAL(1,{120,10,150,23})", 75.75, 1e-10, 0},
		{"MAX", "SUBTOTAL(4,{120,10,150,23})", 150, 0, 0},
		{"MIN", "SUBTOTAL(5,{120,10,150,23})", 10, 0, 0},
		{"COUNT", "SUBTOTAL(2,{120,10,150,23})", 4, 0, 0},
		{"PRODUCT", "SUBTOTAL(6,{2,3,4})", 24, 0, 0},
		{"invalid_0", "SUBTOTAL(0,{1,2})", 0, 0, ErrValVALUE},
		{"invalid_12", "SUBTOTAL(12,{1,2})", 0, 0, ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if tt.wantErr != 0 {
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = type=%v err=%v, want error %v", tt.formula, got.Type, got.Err, tt.wantErr)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Errorf("Eval(%q) = type %v, want ValueNumber", tt.formula, got.Type)
				return
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
