package formula

import (
	"math"
	"testing"
)

// assertComplexStringClose compares two complex number strings with tolerance.
// It parses both as complex numbers and checks that the real and imaginary parts
// are within 1e-12 relative tolerance. This avoids cross-platform floating point
// formatting differences in the last ULP.
func assertComplexStringClose(t *testing.T, formula string, got Value, want string) {
	t.Helper()
	if got.Type != ValueString {
		t.Errorf("Eval(%q) = %v (type %d), want string %q", formula, got, got.Type, want)
		return
	}
	if got.Str == want {
		return // exact match
	}
	// Parse both as complex numbers
	gr, gi, gfail := parseComplex(got.Str)
	wr, wi, wfail := parseComplex(want)
	if gfail || wfail {
		t.Errorf("Eval(%q) = %q, want %q (could not parse as complex)", formula, got.Str, want)
		return
	}
	const tol = 1e-12
	rdiff := math.Abs(gr - wr)
	idiff := math.Abs(gi - wi)
	rmax := math.Max(1, math.Abs(wr))
	imax := math.Max(1, math.Abs(wi))
	if rdiff/rmax > tol || idiff/imax > tol {
		t.Errorf("Eval(%q) = %q, want %q (real diff=%e, imag diff=%e)", formula, got.Str, want, rdiff, idiff)
	}
}

func TestBIN2DEC(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
		}{
			// Basic string inputs
			{`BIN2DEC("0")`, 0},
			{`BIN2DEC("1")`, 1},
			{`BIN2DEC("10")`, 2},
			{`BIN2DEC("11")`, 3},
			{`BIN2DEC("100")`, 4},
			{`BIN2DEC("1001")`, 9},
			{`BIN2DEC("1010")`, 10},
			{`BIN2DEC("1100100")`, 100},
			{`BIN2DEC("11111111")`, 255},

			// Numeric inputs (coerced to string)
			{"BIN2DEC(0)", 0},
			{"BIN2DEC(1)", 1},
			{"BIN2DEC(10)", 2},
			{"BIN2DEC(11)", 3},
			{"BIN2DEC(1001)", 9},
			{"BIN2DEC(11111111)", 255},

			// Max positive (9 digits)
			{`BIN2DEC("111111111")`, 511},

			// Negative two's complement (10 digits starting with 1)
			{`BIN2DEC("1111111111")`, -1},
			{`BIN2DEC("1111111110")`, -2},
			{`BIN2DEC("1111111100")`, -4},
			{`BIN2DEC("1111110000")`, -16},
			{`BIN2DEC("1000000000")`, -512},
			{`BIN2DEC("1000000001")`, -511},

			// Padded with leading zeros
			{`BIN2DEC("0000000001")`, 1},
			{`BIN2DEC("0000000000")`, 0},
			{`BIN2DEC("0111111111")`, 511},

			// Single digit
			{`BIN2DEC("0")`, 0},
			{`BIN2DEC("1")`, 1},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Non-binary characters
			{`BIN2DEC("2")`, ErrValNUM},
			{`BIN2DEC("12345")`, ErrValNUM},
			{`BIN2DEC("abc")`, ErrValNUM},
			{`BIN2DEC("10102")`, ErrValNUM},
			{`BIN2DEC("1010a")`, ErrValNUM},

			// Too many digits (11 digits)
			{`BIN2DEC("10000000000")`, ErrValNUM},
			{`BIN2DEC("11111111111")`, ErrValNUM},

			// Boolean rejection
			{"BIN2DEC(TRUE)", ErrValVALUE},
			{"BIN2DEC(FALSE)", ErrValVALUE},

			// Non-numeric, non-string input
			{`BIN2DEC("xyz")`, ErrValNUM},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"BIN2DEC()", "BIN2DEC(1,2)"} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestBIN2HEX(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic conversions
			{`BIN2HEX("0")`, "0"},
			{`BIN2HEX("1")`, "1"},
			{`BIN2HEX("10")`, "2"},
			{`BIN2HEX("11")`, "3"},
			{`BIN2HEX("100")`, "4"},
			{`BIN2HEX("1001")`, "9"},
			{`BIN2HEX("1010")`, "A"},
			{`BIN2HEX("1011")`, "B"},
			{`BIN2HEX("1100")`, "C"},
			{`BIN2HEX("1101")`, "D"},
			{`BIN2HEX("1110")`, "E"},
			{`BIN2HEX("1111")`, "F"},
			{`BIN2HEX("10000")`, "10"},
			{`BIN2HEX("11111111")`, "FF"},
			{`BIN2HEX("100000000")`, "100"},
			{`BIN2HEX("111111111")`, "1FF"},

			// Numeric input coercion
			{"BIN2HEX(0)", "0"},
			{"BIN2HEX(1)", "1"},
			{"BIN2HEX(1010)", "A"},
			{"BIN2HEX(11111111)", "FF"},

			// With places
			{`BIN2HEX("1111",4)`, "000F"},
			{`BIN2HEX("1",8)`, "00000001"},
			{`BIN2HEX("0",4)`, "0000"},
			{`BIN2HEX("1010",1)`, "A"},
			{`BIN2HEX("11111111",10)`, "00000000FF"},

			// Negative two's complement (10 digits starting with 1)
			{`BIN2HEX("1111111111")`, "FFFFFFFFFF"},
			{`BIN2HEX("1111111110")`, "FFFFFFFFFE"},
			{`BIN2HEX("1111111100")`, "FFFFFFFFFC"},
			{`BIN2HEX("1111110000")`, "FFFFFFFFF0"},
			{`BIN2HEX("1000000000")`, "FFFFFFFE00"},
			{`BIN2HEX("1000000001")`, "FFFFFFFE01"},

			// Padded with leading zeros
			{`BIN2HEX("0000000001")`, "1"},
			{`BIN2HEX("0000000000")`, "0"},
			{`BIN2HEX("0111111111")`, "1FF"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Non-binary characters
			{`BIN2HEX("2")`, ErrValNUM},
			{`BIN2HEX("12345")`, ErrValNUM},
			{`BIN2HEX("abc")`, ErrValNUM},
			{`BIN2HEX("10102")`, ErrValNUM},

			// Too many digits (11 digits)
			{`BIN2HEX("10000000000")`, ErrValNUM},
			{`BIN2HEX("11111111111")`, ErrValNUM},

			// Empty string
			{`BIN2HEX("")`, ErrValNUM},

			// Boolean rejection
			{"BIN2HEX(TRUE)", ErrValVALUE},
			{"BIN2HEX(FALSE)", ErrValVALUE},

			// Places too small
			{`BIN2HEX("11111111",1)`, ErrValNUM},

			// Places = 0
			{`BIN2HEX("1",0)`, ErrValNUM},

			// Negative places
			{`BIN2HEX("1",-1)`, ErrValNUM},

			// Places > 10
			{`BIN2HEX("1",11)`, ErrValNUM},

			// Non-numeric places
			{`BIN2HEX("1","abc")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"BIN2HEX()", `BIN2HEX("1","2","3")`} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestBIN2OCT(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic conversions
			{`BIN2OCT("0")`, "0"},
			{`BIN2OCT("1")`, "1"},
			{`BIN2OCT("10")`, "2"},
			{`BIN2OCT("11")`, "3"},
			{`BIN2OCT("100")`, "4"},
			{`BIN2OCT("111")`, "7"},
			{`BIN2OCT("1000")`, "10"},
			{`BIN2OCT("1001")`, "11"},
			{`BIN2OCT("1010")`, "12"},
			{`BIN2OCT("11111111")`, "377"},
			{`BIN2OCT("100000000")`, "400"},
			{`BIN2OCT("111111111")`, "777"},

			// Numeric input coercion
			{"BIN2OCT(0)", "0"},
			{"BIN2OCT(1)", "1"},
			{"BIN2OCT(1001)", "11"},
			{"BIN2OCT(11111111)", "377"},

			// With places
			{`BIN2OCT("1001",4)`, "0011"},
			{`BIN2OCT("1",8)`, "00000001"},
			{`BIN2OCT("0",4)`, "0000"},
			{`BIN2OCT("111",1)`, "7"},
			{`BIN2OCT("111111111",10)`, "0000000777"},

			// Negative two's complement (10 digits starting with 1)
			{`BIN2OCT("1111111111")`, "7777777777"},
			{`BIN2OCT("1111111110")`, "7777777776"},
			{`BIN2OCT("1111111100")`, "7777777774"},
			{`BIN2OCT("1111110000")`, "7777777760"},
			{`BIN2OCT("1000000000")`, "7777777000"},
			{`BIN2OCT("1000000001")`, "7777777001"},

			// Padded with leading zeros
			{`BIN2OCT("0000000001")`, "1"},
			{`BIN2OCT("0000000000")`, "0"},
			{`BIN2OCT("0111111111")`, "777"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Non-binary characters
			{`BIN2OCT("2")`, ErrValNUM},
			{`BIN2OCT("12345")`, ErrValNUM},
			{`BIN2OCT("abc")`, ErrValNUM},
			{`BIN2OCT("10102")`, ErrValNUM},

			// Too many digits (11 digits)
			{`BIN2OCT("10000000000")`, ErrValNUM},
			{`BIN2OCT("11111111111")`, ErrValNUM},

			// Empty string
			{`BIN2OCT("")`, ErrValNUM},

			// Boolean rejection
			{"BIN2OCT(TRUE)", ErrValVALUE},
			{"BIN2OCT(FALSE)", ErrValVALUE},

			// Places too small
			{`BIN2OCT("11111111",2)`, ErrValNUM},

			// Places = 0
			{`BIN2OCT("1",0)`, ErrValNUM},

			// Negative places
			{`BIN2OCT("1",-1)`, ErrValNUM},

			// Places > 10
			{`BIN2OCT("1",11)`, ErrValNUM},

			// Non-numeric places
			{`BIN2OCT("1","abc")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"BIN2OCT()", `BIN2OCT("1","2","3")`} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestCOMPLEX(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic: positive real and imaginary
			{`COMPLEX(3,4)`, "3+4i"},
			{`COMPLEX(3,4,"j")`, "3+4j"},

			// Zero imaginary: no imaginary part
			{`COMPLEX(5,0)`, "5"},
			{`COMPLEX(-3,0)`, "-3"},
			{`COMPLEX(1,0)`, "1"},
			{`COMPLEX(0.5,0)`, "0.5"},

			// Zero real: no real part
			{`COMPLEX(0,4)`, "4i"},
			{`COMPLEX(0,-4)`, "-4i"},
			{`COMPLEX(0,2)`, "2i"},
			{`COMPLEX(0,-2)`, "-2i"},

			// Both zero
			{`COMPLEX(0,0)`, "0"},

			// Unit imaginary (coefficient 1 or -1)
			{`COMPLEX(0,1)`, "i"},
			{`COMPLEX(0,-1)`, "-i"},
			{`COMPLEX(3,1)`, "3+i"},
			{`COMPLEX(3,-1)`, "3-i"},
			{`COMPLEX(-3,1)`, "-3+i"},
			{`COMPLEX(-3,-1)`, "-3-i"},

			// Negative real
			{`COMPLEX(-3,4)`, "-3+4i"},
			{`COMPLEX(-3,-4)`, "-3-4i"},
			{`COMPLEX(-1,0)`, "-1"},
			{`COMPLEX(-5,2)`, "-5+2i"},

			// Decimal values
			{`COMPLEX(1.5,2.5)`, "1.5+2.5i"},
			{`COMPLEX(0.1,0.2)`, "0.1+0.2i"},
			{`COMPLEX(-1.5,2.5)`, "-1.5+2.5i"},
			{`COMPLEX(1.5,-2.5)`, "1.5-2.5i"},

			// Large numbers
			{`COMPLEX(1000000,2000000)`, "1000000+2000000i"},

			// Small decimals
			{`COMPLEX(0.001,0.002)`, "0.001+0.002i"},

			// j suffix variations
			{`COMPLEX(3,4,"j")`, "3+4j"},
			{`COMPLEX(0,1,"j")`, "j"},
			{`COMPLEX(0,-1,"j")`, "-j"},
			{`COMPLEX(3,1,"j")`, "3+j"},
			{`COMPLEX(3,-1,"j")`, "3-j"},
			{`COMPLEX(0,0,"j")`, "0"},
			{`COMPLEX(5,0,"j")`, "5"},

			// String coercion for numeric args
			{`COMPLEX("3","4")`, "3+4i"},
			{`COMPLEX("1.5","2.5")`, "1.5+2.5i"},

			// Boolean coercion (TRUE=1, FALSE=0)
			{"COMPLEX(TRUE,FALSE)", "1"},
			{"COMPLEX(FALSE,TRUE)", "i"},
			{"COMPLEX(TRUE,TRUE)", "1+i"},

			// i suffix explicit
			{`COMPLEX(3,4,"i")`, "3+4i"},

			// Integer formatting (no trailing .0)
			{`COMPLEX(3.0,4.0)`, "3+4i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid suffix: uppercase
			{`COMPLEX(3,4,"I")`, ErrValVALUE},
			{`COMPLEX(3,4,"J")`, ErrValVALUE},

			// Invalid suffix: other strings
			{`COMPLEX(3,4,"k")`, ErrValVALUE},
			{`COMPLEX(3,4,"x")`, ErrValVALUE},
			{`COMPLEX(3,4,"ij")`, ErrValVALUE},
			{`COMPLEX(3,4,"")`, ErrValVALUE},

			// Non-numeric real_num
			{`COMPLEX("abc",4)`, ErrValVALUE},

			// Non-numeric i_num
			{`COMPLEX(3,"abc")`, ErrValVALUE},

			// Both non-numeric
			{`COMPLEX("abc","def")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"COMPLEX(1)", "COMPLEX(1,2,3,4)"} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestDELTA(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
		}{
			// Doc examples
			{"DELTA(5,4)", 0},
			{"DELTA(5,5)", 1},
			{"DELTA(0.5,0)", 0},

			// Default second argument (0)
			{"DELTA(0)", 1},
			{"DELTA(1)", 0},
			{"DELTA(-1)", 0},

			// Both zero
			{"DELTA(0,0)", 1},

			// Negative numbers
			{"DELTA(-3,-3)", 1},
			{"DELTA(-3,3)", 0},
			{"DELTA(-1,-2)", 0},

			// Floating point
			{"DELTA(1.5,1.5)", 1},
			{"DELTA(1.5,1.6)", 0},
			{"DELTA(0.1,0.1)", 1},

			// Large numbers
			{"DELTA(1000000,1000000)", 1},
			{"DELTA(1000000,1000001)", 0},

			// Boolean coercion (TRUE=1, FALSE=0)
			{"DELTA(TRUE,1)", 1},
			{"DELTA(FALSE,0)", 1},
			{"DELTA(TRUE,0)", 0},
			{"DELTA(FALSE)", 1},
			{"DELTA(TRUE)", 0},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			{`DELTA("abc",1)`, ErrValVALUE},
			{`DELTA(1,"abc")`, ErrValVALUE},
			{`DELTA("x","y")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		cf := evalCompile(t, "DELTA()")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("DELTA() = %v, want error", got)
		}
	})
}

func TestDEC2BIN(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic positive numbers
			{"DEC2BIN(1)", "1"},
			{"DEC2BIN(2)", "10"},
			{"DEC2BIN(9)", "1001"},
			{"DEC2BIN(10)", "1010"},
			{"DEC2BIN(100)", "1100100"},
			{"DEC2BIN(255)", "11111111"},

			// Zero
			{"DEC2BIN(0)", "0"},

			// With places
			{"DEC2BIN(9,4)", "1001"},
			{"DEC2BIN(1,8)", "00000001"},
			{"DEC2BIN(0,4)", "0000"},
			{"DEC2BIN(2,10)", "0000000010"},
			{"DEC2BIN(9,10)", "0000001001"},

			// Negative numbers (two's complement)
			{"DEC2BIN(-1)", "1111111111"},
			{"DEC2BIN(-2)", "1111111110"},
			{"DEC2BIN(-100)", "1110011100"},
			{"DEC2BIN(-512)", "1000000000"},

			// Boundaries
			{"DEC2BIN(511)", "111111111"},
			{"DEC2BIN(-512)", "1000000000"},

			// Non-integer truncation
			{"DEC2BIN(9.9)", "1001"},
			{"DEC2BIN(9.1)", "1001"},
			{"DEC2BIN(-1.9)", "1111111111"},

			// String coercion
			{`DEC2BIN("9")`, "1001"},
			{`DEC2BIN("0")`, "0"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Out of range
			{"DEC2BIN(512)", ErrValNUM},
			{"DEC2BIN(-513)", ErrValNUM},
			{"DEC2BIN(1000)", ErrValNUM},
			{"DEC2BIN(-1000)", ErrValNUM},

			// Places too small
			{"DEC2BIN(9,2)", ErrValNUM},
			{"DEC2BIN(255,4)", ErrValNUM},

			// Places = 0
			{"DEC2BIN(1,0)", ErrValNUM},

			// Negative places
			{"DEC2BIN(1,-1)", ErrValNUM},

			// Places > 10
			{"DEC2BIN(1,11)", ErrValNUM},

			// Boolean input (engineering functions reject booleans)
			{"DEC2BIN(TRUE)", ErrValVALUE},
			{"DEC2BIN(FALSE)", ErrValVALUE},

			// Non-numeric input
			{`DEC2BIN("abc")`, ErrValVALUE},

			// Non-numeric places
			{`DEC2BIN(1,"abc")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"DEC2BIN()", "DEC2BIN(1,2,3)"} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestDEC2HEX(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic positive numbers
			{"DEC2HEX(0)", "0"},
			{"DEC2HEX(1)", "1"},
			{"DEC2HEX(9)", "9"},
			{"DEC2HEX(10)", "A"},
			{"DEC2HEX(15)", "F"},
			{"DEC2HEX(16)", "10"},
			{"DEC2HEX(100)", "64"},
			{"DEC2HEX(255)", "FF"},
			{"DEC2HEX(256)", "100"},
			{"DEC2HEX(1000)", "3E8"},
			{"DEC2HEX(65535)", "FFFF"},

			// With places
			{"DEC2HEX(255,4)", "00FF"},
			{"DEC2HEX(1,8)", "00000001"},
			{"DEC2HEX(0,4)", "0000"},
			{"DEC2HEX(10,1)", "A"},
			{"DEC2HEX(255,10)", "00000000FF"},

			// Negative numbers (two's complement)
			{"DEC2HEX(-1)", "FFFFFFFFFF"},
			{"DEC2HEX(-2)", "FFFFFFFFFE"},
			{"DEC2HEX(-100)", "FFFFFFFF9C"},
			{"DEC2HEX(-256)", "FFFFFFFF00"},
			{"DEC2HEX(-16)", "FFFFFFFFF0"},

			// Boundaries
			{"DEC2HEX(549755813887)", "7FFFFFFFFF"},
			{"DEC2HEX(-549755813888)", "8000000000"},

			// Non-integer truncation
			{"DEC2HEX(15.9)", "F"},
			{"DEC2HEX(15.1)", "F"},
			{"DEC2HEX(-1.9)", "FFFFFFFFFF"},
			{"DEC2HEX(100.7)", "64"},

			// String coercion
			{`DEC2HEX("255")`, "FF"},
			{`DEC2HEX("0")`, "0"},
			{`DEC2HEX("100")`, "64"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Out of range
			{"DEC2HEX(549755813888)", ErrValNUM},
			{"DEC2HEX(-549755813889)", ErrValNUM},
			{"DEC2HEX(999999999999)", ErrValNUM},
			{"DEC2HEX(-999999999999)", ErrValNUM},

			// Places too small
			{"DEC2HEX(255,1)", ErrValNUM},
			{"DEC2HEX(256,2)", ErrValNUM},
			{"DEC2HEX(65535,3)", ErrValNUM},

			// Places = 0
			{"DEC2HEX(1,0)", ErrValNUM},

			// Negative places
			{"DEC2HEX(1,-1)", ErrValNUM},

			// Places > 10
			{"DEC2HEX(1,11)", ErrValNUM},

			// Boolean input (engineering functions reject booleans)
			{"DEC2HEX(TRUE)", ErrValVALUE},
			{"DEC2HEX(FALSE)", ErrValVALUE},

			// Non-numeric input
			{`DEC2HEX("abc")`, ErrValVALUE},

			// Non-numeric places
			{`DEC2HEX(1,"abc")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"DEC2HEX()", "DEC2HEX(1,2,3)"} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestDEC2OCT(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic positive numbers
			{"DEC2OCT(0)", "0"},
			{"DEC2OCT(1)", "1"},
			{"DEC2OCT(7)", "7"},
			{"DEC2OCT(8)", "10"},
			{"DEC2OCT(9)", "11"},
			{"DEC2OCT(100)", "144"},
			{"DEC2OCT(255)", "377"},
			{"DEC2OCT(511)", "777"},
			{"DEC2OCT(512)", "1000"},
			{"DEC2OCT(1000)", "1750"},

			// With places
			{"DEC2OCT(100,4)", "0144"},
			{"DEC2OCT(1,8)", "00000001"},
			{"DEC2OCT(0,4)", "0000"},
			{"DEC2OCT(7,1)", "7"},
			{"DEC2OCT(8,3)", "010"},
			{"DEC2OCT(100,10)", "0000000144"},

			// Negative numbers (two's complement)
			{"DEC2OCT(-1)", "7777777777"},
			{"DEC2OCT(-2)", "7777777776"},
			{"DEC2OCT(-100)", "7777777634"},
			{"DEC2OCT(-8)", "7777777770"},
			{"DEC2OCT(-256)", "7777777400"},

			// Boundaries
			{"DEC2OCT(536870911)", "3777777777"},
			{"DEC2OCT(-536870912)", "4000000000"},

			// Non-integer truncation
			{"DEC2OCT(7.9)", "7"},
			{"DEC2OCT(7.1)", "7"},
			{"DEC2OCT(-1.9)", "7777777777"},
			{"DEC2OCT(100.7)", "144"},

			// String coercion
			{`DEC2OCT("100")`, "144"},
			{`DEC2OCT("0")`, "0"},
			{`DEC2OCT("255")`, "377"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Out of range
			{"DEC2OCT(536870912)", ErrValNUM},
			{"DEC2OCT(-536870913)", ErrValNUM},
			{"DEC2OCT(1000000000)", ErrValNUM},
			{"DEC2OCT(-1000000000)", ErrValNUM},

			// Places too small
			{"DEC2OCT(100,2)", ErrValNUM},
			{"DEC2OCT(255,2)", ErrValNUM},
			{"DEC2OCT(8,1)", ErrValNUM},

			// Places = 0
			{"DEC2OCT(1,0)", ErrValNUM},

			// Negative places
			{"DEC2OCT(1,-1)", ErrValNUM},

			// Places > 10
			{"DEC2OCT(1,11)", ErrValNUM},

			// Boolean input (engineering functions reject booleans)
			{"DEC2OCT(TRUE)", ErrValVALUE},
			{"DEC2OCT(FALSE)", ErrValVALUE},

			// Non-numeric input
			{`DEC2OCT("abc")`, ErrValVALUE},

			// Non-numeric places
			{`DEC2OCT(1,"abc")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"DEC2OCT()", "DEC2OCT(1,2,3)"} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestHEX2DEC(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
		}{
			// Basic string inputs
			{`HEX2DEC("0")`, 0},
			{`HEX2DEC("1")`, 1},
			{`HEX2DEC("A")`, 10},
			{`HEX2DEC("F")`, 15},
			{`HEX2DEC("10")`, 16},
			{`HEX2DEC("FF")`, 255},
			{`HEX2DEC("64")`, 100},
			{`HEX2DEC("100")`, 256},
			{`HEX2DEC("1F4")`, 500},
			{`HEX2DEC("3E8")`, 1000},
			{`HEX2DEC("FFFF")`, 65535},

			// Case insensitive
			{`HEX2DEC("ff")`, 255},
			{`HEX2DEC("aB")`, 171},
			{`HEX2DEC("a")`, 10},
			{`HEX2DEC("f")`, 15},

			// Numeric inputs (coerced to string)
			{"HEX2DEC(0)", 0},
			{"HEX2DEC(1)", 1},
			{"HEX2DEC(100)", 256},
			{"HEX2DEC(10)", 16},

			// Negative two's complement (10 digits, first digit >= 8)
			{`HEX2DEC("FFFFFFFFFF")`, -1},
			{`HEX2DEC("FFFFFFFFFE")`, -2},
			{`HEX2DEC("8000000000")`, -549755813888},
			{`HEX2DEC("FFFFFFFF9C")`, -100},

			// Max positive (10 digits, first digit < 8)
			{`HEX2DEC("7FFFFFFFFF")`, 549755813887},

			// Padded with leading zeros
			{`HEX2DEC("00000000FF")`, 255},
			{`HEX2DEC("0000000001")`, 1},
			{`HEX2DEC("0000000000")`, 0},

			// Single digit
			{`HEX2DEC("0")`, 0},
			{`HEX2DEC("9")`, 9},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Non-hex characters
			{`HEX2DEC("G")`, ErrValNUM},
			{`HEX2DEC("XYZZY")`, ErrValNUM},
			{`HEX2DEC("1G")`, ErrValNUM},
			{`HEX2DEC("ZZ")`, ErrValNUM},

			// Too many digits (11 hex digits)
			{`HEX2DEC("10000000000")`, ErrValNUM},
			{`HEX2DEC("FFFFFFFFFFF")`, ErrValNUM},

			// Empty string
			{`HEX2DEC("")`, ErrValNUM},

			// Boolean rejection
			{"HEX2DEC(TRUE)", ErrValVALUE},
			{"HEX2DEC(FALSE)", ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"HEX2DEC()", `HEX2DEC("1","2")`} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestOCT2DEC(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
		}{
			// Basic string inputs
			{`OCT2DEC("0")`, 0},
			{`OCT2DEC("1")`, 1},
			{`OCT2DEC("2")`, 2},
			{`OCT2DEC("7")`, 7},
			{`OCT2DEC("10")`, 8},
			{`OCT2DEC("11")`, 9},
			{`OCT2DEC("17")`, 15},
			{`OCT2DEC("20")`, 16},
			{`OCT2DEC("144")`, 100},
			{`OCT2DEC("377")`, 255},
			{`OCT2DEC("777")`, 511},
			{`OCT2DEC("1000")`, 512},
			{`OCT2DEC("1750")`, 1000},

			// Numeric inputs (coerced to string)
			{"OCT2DEC(0)", 0},
			{"OCT2DEC(1)", 1},
			{"OCT2DEC(10)", 8},
			{"OCT2DEC(11)", 9},
			{"OCT2DEC(144)", 100},
			{"OCT2DEC(377)", 255},

			// Max positive (10 digits, first digit < 4)
			{`OCT2DEC("3777777777")`, 536870911},

			// Negative two's complement (10 digits, first digit >= 4)
			{`OCT2DEC("7777777777")`, -1},
			{`OCT2DEC("7777777776")`, -2},
			{`OCT2DEC("7777777770")`, -8},
			{`OCT2DEC("7777777634")`, -100},
			{`OCT2DEC("4000000000")`, -536870912},
			{`OCT2DEC("4000000001")`, -536870911},

			// Padded with leading zeros
			{`OCT2DEC("0000000001")`, 1},
			{`OCT2DEC("0000000000")`, 0},
			{`OCT2DEC("0000000010")`, 8},

			// Single digit
			{`OCT2DEC("3")`, 3},
			{`OCT2DEC("5")`, 5},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Non-octal characters
			{`OCT2DEC("8")`, ErrValNUM},
			{`OCT2DEC("9")`, ErrValNUM},
			{`OCT2DEC("abc")`, ErrValNUM},
			{`OCT2DEC("1238")`, ErrValNUM},
			{`OCT2DEC("12a")`, ErrValNUM},
			{`OCT2DEC("G")`, ErrValNUM},

			// Too many digits (11 digits)
			{`OCT2DEC("40000000000")`, ErrValNUM},
			{`OCT2DEC("77777777777")`, ErrValNUM},

			// Empty string
			{`OCT2DEC("")`, ErrValNUM},

			// Boolean rejection
			{"OCT2DEC(TRUE)", ErrValVALUE},
			{"OCT2DEC(FALSE)", ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"OCT2DEC()", `OCT2DEC("1","2")`} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestHEX2BIN(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic positive numbers
			{`HEX2BIN("0")`, "0"},
			{`HEX2BIN("1")`, "1"},
			{`HEX2BIN("2")`, "10"},
			{`HEX2BIN("A")`, "1010"},
			{`HEX2BIN("F")`, "1111"},
			{`HEX2BIN("10")`, "10000"},
			{`HEX2BIN("1F")`, "11111"},
			{`HEX2BIN("FF")`, "11111111"},
			{`HEX2BIN("1FF")`, "111111111"},

			// Case insensitive
			{`HEX2BIN("a")`, "1010"},
			{`HEX2BIN("f")`, "1111"},
			{`HEX2BIN("ff")`, "11111111"},

			// With places
			{`HEX2BIN("A",8)`, "00001010"},
			{`HEX2BIN("1",4)`, "0001"},
			{`HEX2BIN("0",4)`, "0000"},
			{`HEX2BIN("F",8)`, "00001111"},
			{`HEX2BIN("1FF",10)`, "0111111111"},

			// Numeric input (coerced to string)
			{"HEX2BIN(0)", "0"},
			{"HEX2BIN(1)", "1"},
			{"HEX2BIN(10)", "10000"},

			// Negative numbers (two's complement hex → binary)
			{`HEX2BIN("FFFFFFFFFF")`, "1111111111"}, // -1
			{`HEX2BIN("FFFFFFFFFE")`, "1111111110"}, // -2
			{`HEX2BIN("FFFFFFFE00")`, "1000000000"}, // -512

			// Max positive
			{`HEX2BIN("1FF")`, "111111111"}, // 511
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Out of range (512 > 511)
			{`HEX2BIN("200")`, ErrValNUM},
			// Out of range (large positive)
			{`HEX2BIN("1000")`, ErrValNUM},
			// Non-hex characters
			{`HEX2BIN("G")`, ErrValNUM},
			{`HEX2BIN("XYZ")`, ErrValNUM},
			// Too many digits
			{`HEX2BIN("10000000000")`, ErrValNUM},
			// Empty string
			{`HEX2BIN("")`, ErrValNUM},
			// Boolean rejection
			{"HEX2BIN(TRUE)", ErrValVALUE},
			{"HEX2BIN(FALSE)", ErrValVALUE},
			// Places too small
			{`HEX2BIN("FF",4)`, ErrValNUM},
			// Places = 0
			{`HEX2BIN("1",0)`, ErrValNUM},
			// Places > 10
			{`HEX2BIN("1",11)`, ErrValNUM},
			// Negative places
			{`HEX2BIN("1",-1)`, ErrValNUM},
			// Non-numeric places
			{`HEX2BIN("1","abc")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"HEX2BIN()", `HEX2BIN("1","2","3")`} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestHEX2OCT(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic positive numbers
			{`HEX2OCT("0")`, "0"},
			{`HEX2OCT("1")`, "1"},
			{`HEX2OCT("8")`, "10"},
			{`HEX2OCT("A")`, "12"},
			{`HEX2OCT("F")`, "17"},
			{`HEX2OCT("10")`, "20"},
			{`HEX2OCT("FF")`, "377"},
			{`HEX2OCT("1FF")`, "777"},
			{`HEX2OCT("64")`, "144"},
			{`HEX2OCT("3E8")`, "1750"},

			// Case insensitive
			{`HEX2OCT("a")`, "12"},
			{`HEX2OCT("ff")`, "377"},

			// With places
			{`HEX2OCT("A",4)`, "0012"},
			{`HEX2OCT("1",4)`, "0001"},
			{`HEX2OCT("0",4)`, "0000"},
			{`HEX2OCT("FF",4)`, "0377"},

			// Numeric input
			{"HEX2OCT(0)", "0"},
			{"HEX2OCT(1)", "1"},

			// Negative numbers (two's complement)
			{`HEX2OCT("FFFFFFFFFF")`, "7777777777"}, // -1
			{`HEX2OCT("FFFFFFFFFE")`, "7777777776"}, // -2
			{`HEX2OCT("FFE0000000")`, "4000000000"}, // -536870912

			// Max positive
			{`HEX2OCT("1FFFFFFF")`, "3777777777"}, // 536870911
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Out of range (too large positive for octal)
			{`HEX2OCT("20000000")`, ErrValNUM},
			// Out of range (too negative for octal)
			{`HEX2OCT("8000000000")`, ErrValNUM},
			// Non-hex characters
			{`HEX2OCT("G")`, ErrValNUM},
			{`HEX2OCT("XYZ")`, ErrValNUM},
			// Too many digits
			{`HEX2OCT("10000000000")`, ErrValNUM},
			// Empty string
			{`HEX2OCT("")`, ErrValNUM},
			// Boolean rejection
			{"HEX2OCT(TRUE)", ErrValVALUE},
			{"HEX2OCT(FALSE)", ErrValVALUE},
			// Places too small
			{`HEX2OCT("FF",2)`, ErrValNUM},
			// Places = 0
			{`HEX2OCT("1",0)`, ErrValNUM},
			// Places > 10
			{`HEX2OCT("1",11)`, ErrValNUM},
			// Non-numeric places
			{`HEX2OCT("1","abc")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"HEX2OCT()", `HEX2OCT("1","2","3")`} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestOCT2BIN(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic positive numbers
			{`OCT2BIN("0")`, "0"},
			{`OCT2BIN("1")`, "1"},
			{`OCT2BIN("7")`, "111"},
			{`OCT2BIN("10")`, "1000"},
			{`OCT2BIN("11")`, "1001"},
			{`OCT2BIN("17")`, "1111"},
			{`OCT2BIN("77")`, "111111"},
			{`OCT2BIN("100")`, "1000000"},
			{`OCT2BIN("777")`, "111111111"},

			// With places
			{`OCT2BIN("7",6)`, "000111"},
			{`OCT2BIN("1",4)`, "0001"},
			{`OCT2BIN("0",4)`, "0000"},
			{`OCT2BIN("10",8)`, "00001000"},
			{`OCT2BIN("777",10)`, "0111111111"},

			// Numeric input
			{"OCT2BIN(0)", "0"},
			{"OCT2BIN(1)", "1"},
			{"OCT2BIN(7)", "111"},
			{"OCT2BIN(10)", "1000"},
			{"OCT2BIN(77)", "111111"},

			// Negative numbers (two's complement)
			{`OCT2BIN("7777777777")`, "1111111111"}, // -1
			{`OCT2BIN("7777777776")`, "1111111110"}, // -2
			{`OCT2BIN("7777777000")`, "1000000000"}, // -512

			// Max positive
			{`OCT2BIN("777")`, "111111111"}, // 511
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Out of range (512 > 511)
			{`OCT2BIN("1000")`, ErrValNUM},
			// Out of range (large)
			{`OCT2BIN("7777")`, ErrValNUM},
			// Non-octal characters
			{`OCT2BIN("8")`, ErrValNUM},
			{`OCT2BIN("9")`, ErrValNUM},
			{`OCT2BIN("abc")`, ErrValNUM},
			// Too many digits
			{`OCT2BIN("40000000000")`, ErrValNUM},
			// Empty string
			{`OCT2BIN("")`, ErrValNUM},
			// Boolean rejection
			{"OCT2BIN(TRUE)", ErrValVALUE},
			{"OCT2BIN(FALSE)", ErrValVALUE},
			// Places too small
			{`OCT2BIN("77",4)`, ErrValNUM},
			// Places = 0
			{`OCT2BIN("1",0)`, ErrValNUM},
			// Places > 10
			{`OCT2BIN("1",11)`, ErrValNUM},
			// Negative places
			{`OCT2BIN("1",-1)`, ErrValNUM},
			// Non-numeric places
			{`OCT2BIN("1","abc")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"OCT2BIN()", `OCT2BIN("1","2","3")`} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestOCT2HEX(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic positive numbers
			{`OCT2HEX("0")`, "0"},
			{`OCT2HEX("1")`, "1"},
			{`OCT2HEX("7")`, "7"},
			{`OCT2HEX("10")`, "8"},
			{`OCT2HEX("12")`, "A"},
			{`OCT2HEX("17")`, "F"},
			{`OCT2HEX("20")`, "10"},
			{`OCT2HEX("144")`, "64"},
			{`OCT2HEX("377")`, "FF"},
			{`OCT2HEX("777")`, "1FF"},
			{`OCT2HEX("1750")`, "3E8"},

			// With places
			{`OCT2HEX("377",4)`, "00FF"},
			{`OCT2HEX("1",4)`, "0001"},
			{`OCT2HEX("0",4)`, "0000"},
			{`OCT2HEX("12",4)`, "000A"},

			// Numeric input
			{"OCT2HEX(0)", "0"},
			{"OCT2HEX(1)", "1"},
			{"OCT2HEX(10)", "8"},
			{"OCT2HEX(377)", "FF"},

			// Negative numbers (two's complement)
			{`OCT2HEX("7777777777")`, "FFFFFFFFFF"}, // -1
			{`OCT2HEX("7777777776")`, "FFFFFFFFFE"}, // -2
			{`OCT2HEX("4000000000")`, "FFE0000000"}, // -536870912

			// Max positive
			{`OCT2HEX("3777777777")`, "1FFFFFFF"}, // 536870911
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Non-octal characters
			{`OCT2HEX("8")`, ErrValNUM},
			{`OCT2HEX("9")`, ErrValNUM},
			{`OCT2HEX("abc")`, ErrValNUM},
			// Too many digits
			{`OCT2HEX("40000000000")`, ErrValNUM},
			// Empty string
			{`OCT2HEX("")`, ErrValNUM},
			// Boolean rejection
			{"OCT2HEX(TRUE)", ErrValVALUE},
			{"OCT2HEX(FALSE)", ErrValVALUE},
			// Places too small
			{`OCT2HEX("377",1)`, ErrValNUM},
			// Places = 0
			{`OCT2HEX("1",0)`, ErrValNUM},
			// Places > 10
			{`OCT2HEX("1",11)`, ErrValNUM},
			// Negative places
			{`OCT2HEX("1",-1)`, ErrValNUM},
			// Non-numeric places
			{`OCT2HEX("1","abc")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"OCT2HEX()", `OCT2HEX("1","2","3")`} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestCONVERT(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
			tol     float64 // 0 means exact match
		}{
			// Temperature conversions
			{`CONVERT(1,"C","F")`, 33.8, 1e-9},
			{`CONVERT(32,"F","C")`, 0, 1e-9},
			{`CONVERT(0,"C","K")`, 273.15, 1e-9},
			{`CONVERT(100,"C","F")`, 212, 1e-9},
			{`CONVERT(212,"F","C")`, 100, 1e-9},
			{`CONVERT(0,"K","C")`, -273.15, 1e-9},
			{`CONVERT(0,"C","Rank")`, 491.67, 1e-9},
			{`CONVERT(0,"C","Reau")`, 0, 1e-9},
			{`CONVERT(100,"C","Reau")`, 80, 1e-9},
			{`CONVERT(80,"Reau","C")`, 100, 1e-9},

			// Same unit returns same number
			{`CONVERT(42,"C","C")`, 42, 0},
			{`CONVERT(42,"kg","kg")`, 42, 0},

			// Mass conversions
			{`CONVERT(1,"kg","lbm")`, 2.20462262184878, 1e-6},
			{`CONVERT(1,"lbm","kg")`, 0.45359237, 1e-6},
			{`CONVERT(1,"kg","g")`, 1000, 1e-9},
			{`CONVERT(1,"lbm","ozm")`, 16, 1e-6},
			{`CONVERT(1,"stone","lbm")`, 14, 1e-4},

			// Distance conversions
			{`CONVERT(1,"mi","km")`, 1.609344, 1e-6},
			{`CONVERT(1,"ft","in")`, 12, 1e-9},
			{`CONVERT(1,"yd","ft")`, 3, 1e-9},
			{`CONVERT(1,"m","ft")`, 3.28083989501312, 1e-6},
			{`CONVERT(1,"Nmi","m")`, 1852, 1e-9},
			{`CONVERT(1,"in","m")`, 0.0254, 1e-9},

			// Time conversions
			{`CONVERT(1,"hr","mn")`, 60, 1e-9},
			{`CONVERT(1,"day","hr")`, 24, 1e-9},
			{`CONVERT(1,"day","sec")`, 86400, 1e-9},
			{`CONVERT(1,"yr","day")`, 365.25, 1e-9},

			// Volume conversions
			{`CONVERT(1,"gal","l")`, 3.785411784, 1e-6},
			{`CONVERT(1,"l","gal")`, 0.264172052358148, 1e-6},
			{`CONVERT(1,"cup","tbs")`, 16, 1e-4},
			{`CONVERT(1,"gal","qt")`, 4, 1e-9},
			{`CONVERT(1,"gal","pt")`, 8, 1e-9},

			// Information conversions
			{`CONVERT(1,"byte","bit")`, 8, 0},
			{`CONVERT(8,"bit","byte")`, 1, 0},

			// SI prefix conversions
			{`CONVERT(1,"km","m")`, 1000, 1e-9},
			{`CONVERT(1,"m","cm")`, 100, 1e-9},
			{`CONVERT(1,"mg","g")`, 0.001, 1e-12},
			{`CONVERT(1,"kg","mg")`, 1e6, 1e-3},
			{`CONVERT(1,"kW","W")`, 1000, 1e-9},
			{`CONVERT(1,"MW","kW")`, 1000, 1e-9},

			// Pressure
			{`CONVERT(1,"atm","Pa")`, 101325, 1e-3},

			// Force
			{`CONVERT(1,"lbf","N")`, 4.4482216152605, 1e-6},

			// Energy
			{`CONVERT(1,"BTU","J")`, 1055.05585262, 1e-3},

			// Power
			{`CONVERT(1,"HP","W")`, 745.69987158227, 1e-3},

			// Magnetism
			{`CONVERT(1,"T","ga")`, 10000, 1e-9},

			// Speed
			{`CONVERT(1,"mph","m/s")`, 0.44704, 1e-6},
			{`CONVERT(1,"kn","mph")`, 1.15077944802354, 1e-4},

			// Area
			{`CONVERT(1,"ha","m2")`, 10000, 1e-9},
			{`CONVERT(1,"us_acre","m2")`, 4046.8564224, 1e-6},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber {
					t.Fatalf("Eval(%q) = %v, want number", tt.formula, got)
				}
				if tt.tol == 0 {
					if got.Num != tt.want {
						t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want)
					}
				} else {
					diff := got.Num - tt.want
					if diff < 0 {
						diff = -diff
					}
					if diff > tt.tol {
						t.Errorf("Eval(%q) = %v, want %v (tol %v, diff %v)", tt.formula, got.Num, tt.want, tt.tol, diff)
					}
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Cross-category: mass vs distance
			{`CONVERT(1,"kg","m")`, ErrValNA},
			// Unknown units
			{`CONVERT(1,"foo","bar")`, ErrValNA},
			// Unknown from unit
			{`CONVERT(1,"foo","kg")`, ErrValNA},
			// Unknown to unit
			{`CONVERT(1,"kg","bar")`, ErrValNA},
			// Temperature vs mass
			{`CONVERT(1,"C","kg")`, ErrValNA},
			// Wrong arg count
			{`CONVERT(1,"kg")`, ErrValVALUE},
			{`CONVERT(1,"kg","g","extra")`, ErrValVALUE},
			// Non-numeric first arg
			{`CONVERT("abc","kg","g")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})
}

func TestGESTEP(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
		}{
			// Doc examples
			{"GESTEP(5,4)", 1},
			{"GESTEP(5,5)", 1},
			{"GESTEP(-4,-5)", 1},
			{"GESTEP(-1)", 0},

			// Default step (0)
			{"GESTEP(0)", 1},
			{"GESTEP(1)", 1},
			{"GESTEP(5)", 1},
			{"GESTEP(-0.001)", 0},

			// Equal values
			{"GESTEP(0,0)", 1},
			{"GESTEP(3,3)", 1},
			{"GESTEP(-3,-3)", 1},

			// Greater than step
			{"GESTEP(10,5)", 1},
			{"GESTEP(0.6,0.5)", 1},
			{"GESTEP(-1,-2)", 1},

			// Less than step
			{"GESTEP(4,5)", 0},
			{"GESTEP(-5,-4)", 0},
			{"GESTEP(0,1)", 0},
			{"GESTEP(0.4,0.5)", 0},

			// Boolean coercion
			{"GESTEP(TRUE,1)", 1},
			{"GESTEP(FALSE,0)", 1},
			{"GESTEP(FALSE,1)", 0},
			{"GESTEP(TRUE)", 1},
			{"GESTEP(FALSE)", 1},

			// Large numbers
			{"GESTEP(1000000,999999)", 1},
			{"GESTEP(999999,1000000)", 0},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			{`GESTEP("abc",1)`, ErrValVALUE},
			{`GESTEP(1,"abc")`, ErrValVALUE},
			{`GESTEP("x","y")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		cf := evalCompile(t, "GESTEP()")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("GESTEP() = %v, want error", got)
		}
	})
}

func TestIMAGINARY(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
		}{
			// Basic complex numbers
			{`IMAGINARY("3+4i")`, 4},
			{`IMAGINARY("3-4i")`, -4},
			{`IMAGINARY("-3+4i")`, 4},
			{`IMAGINARY("-3-4i")`, -4},

			// Pure imaginary
			{`IMAGINARY("i")`, 1},
			{`IMAGINARY("-i")`, -1},
			{`IMAGINARY("2i")`, 2},
			{`IMAGINARY("-3.5i")`, -3.5},
			{`IMAGINARY("0.5i")`, 0.5},

			// j suffix
			{`IMAGINARY("3+4j")`, 4},
			{`IMAGINARY("3-4j")`, -4},
			{`IMAGINARY("j")`, 1},
			{`IMAGINARY("-j")`, -1},
			{`IMAGINARY("0-j")`, -1},

			// Pure real (imaginary part is 0)
			{`IMAGINARY("3")`, 0},
			{`IMAGINARY("-5")`, 0},
			{`IMAGINARY("0")`, 0},
			{`IMAGINARY("3.14")`, 0},

			// Plain number (not a string)
			{"IMAGINARY(4)", 0},
			{"IMAGINARY(0)", 0},
			{"IMAGINARY(-7)", 0},
			{"IMAGINARY(3.14)", 0},

			// Boolean treated as real number
			{"IMAGINARY(TRUE)", 0},
			{"IMAGINARY(FALSE)", 0},

			// Unit imaginary with real part
			{`IMAGINARY("3+i")`, 1},
			{`IMAGINARY("3-i")`, -1},

			// Decimal coefficients
			{`IMAGINARY("-5.5+2.3i")`, 2.3},
			{`IMAGINARY("1.5-2.5i")`, -2.5},

			// Zero real, explicit
			{`IMAGINARY("0+4i")`, 4},
			{`IMAGINARY("0-4i")`, -4},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMAGINARY("invalid")`, ErrValNUM},
			{`IMAGINARY("")`, ErrValNUM},
			{`IMAGINARY("abc")`, ErrValNUM},
			{`IMAGINARY("3+4")`, ErrValNUM},
			{`IMAGINARY("3+4k")`, ErrValNUM},
			{`IMAGINARY("+")`, ErrValNUM},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMAGINARY(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMAGINARY(1/0) = %v, want error", got)
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"IMAGINARY()", `IMAGINARY("3+4i","extra")`} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestIMABS(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
		}{
			// Pythagorean triples
			{`IMABS("3+4i")`, 5},
			{`IMABS("5+12i")`, 13},
			{`IMABS("8+15i")`, 17},

			// Pure real (positive and negative)
			{`IMABS("3")`, 3},
			{`IMABS("-3")`, 3},
			{`IMABS("-5")`, 5},
			{`IMABS("0")`, 0},

			// Plain number (not a string)
			{"IMABS(4)", 4},
			{"IMABS(-7)", 7},
			{"IMABS(0)", 0},

			// Pure imaginary
			{`IMABS("4i")`, 4},
			{`IMABS("-3i")`, 3},
			{`IMABS("i")`, 1},
			{`IMABS("-i")`, 1},

			// Both parts negative
			{`IMABS("-3-4i")`, 5},

			// j suffix
			{`IMABS("3+4j")`, 5},
			{`IMABS("j")`, 1},
			{`IMABS("-j")`, 1},

			// Decimal coefficients: sqrt(1.5² + 2²) = sqrt(2.25+4) = sqrt(6.25) = 2.5
			{`IMABS("1.5+2i")`, 2.5},

			// Large numbers: sqrt(300² + 400²) = sqrt(250000) = 500
			{`IMABS("300+400i")`, 500},

			// Unit imaginary with real part: sqrt(1² + 1²) = sqrt(2)
			{`IMABS("1+i")`, math.Sqrt(2)},
			{`IMABS("1-i")`, math.Sqrt(2)},

			// Boolean: TRUE=1, FALSE=0
			{"IMABS(TRUE)", 1},
			{"IMABS(FALSE)", 0},

			// Combined with COMPLEX
			{`IMABS(COMPLEX(3,4))`, 5},
			{`IMABS(COMPLEX(0,0))`, 0},
			{`IMABS(COMPLEX(5,12))`, 13},
			{`IMABS(COMPLEX(0,1))`, 1},

			// Explicit zero parts
			{`IMABS("0+4i")`, 4},
			{`IMABS("3+0i")`, 3},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMABS("invalid")`, ErrValNUM},
			{`IMABS("")`, ErrValNUM},
			{`IMABS("abc")`, ErrValNUM},
			{`IMABS("3+4")`, ErrValNUM},
			{`IMABS("3+4k")`, ErrValNUM},
			{`IMABS("+")`, ErrValNUM},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMABS(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMABS(1/0) = %v, want error", got)
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"IMABS()", `IMABS("3+4i","extra")`} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestIMREAL(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
		}{
			// Basic complex numbers
			{`IMREAL("6-9i")`, 6},
			{`IMREAL("3+4i")`, 3},
			{`IMREAL("-3+4i")`, -3},
			{`IMREAL("-3-4i")`, -3},

			// Pure imaginary (real part is 0)
			{`IMREAL("i")`, 0},
			{`IMREAL("-i")`, 0},
			{`IMREAL("4i")`, 0},
			{`IMREAL("-4i")`, 0},
			{`IMREAL("2.5i")`, 0},

			// Pure real
			{`IMREAL("3")`, 3},
			{`IMREAL("-5")`, -5},
			{`IMREAL("0")`, 0},
			{`IMREAL("3.14")`, 3.14},

			// Plain number (not a string)
			{"IMREAL(4)", 4},
			{"IMREAL(5)", 5},
			{"IMREAL(0)", 0},
			{"IMREAL(-7)", -7},
			{"IMREAL(3.14)", 3.14},

			// Boolean: TRUE=1, FALSE=0
			{"IMREAL(TRUE)", 1},
			{"IMREAL(FALSE)", 0},

			// j suffix
			{`IMREAL("3-4j")`, 3},
			{`IMREAL("3+4j")`, 3},
			{`IMREAL("j")`, 0},
			{`IMREAL("-j")`, 0},

			// Decimal coefficients
			{`IMREAL("-5.5+2.3i")`, -5.5},
			{`IMREAL("1.5-2.5i")`, 1.5},

			// Unit imaginary with real part
			{`IMREAL("3+i")`, 3},
			{`IMREAL("3-i")`, 3},

			// Zero real, explicit
			{`IMREAL("0+4i")`, 0},
			{`IMREAL("0-4i")`, 0},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber || got.Num != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMREAL("invalid")`, ErrValNUM},
			{`IMREAL("")`, ErrValNUM},
			{`IMREAL("abc")`, ErrValNUM},
			{`IMREAL("3+4")`, ErrValNUM},
			{`IMREAL("3+4k")`, ErrValNUM},
			{`IMREAL("+")`, ErrValNUM},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMREAL(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMREAL(1/0) = %v, want error", got)
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"IMREAL()", `IMREAL("3+4i","extra")`} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestIMSUM(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic: two complex numbers
			{`IMSUM("3+6i","5-2i")`, "8+4i"},
			// Three arguments
			{`IMSUM("1+i","2+2i","3+3i")`, "6+6i"},
			// Four arguments
			{`IMSUM("1+i","2+2i","3+3i","4+4i")`, "10+10i"},
			// Pure real addition
			{`IMSUM("3","5")`, "8"},
			// Pure imaginary addition
			{`IMSUM("3i","4i")`, "7i"},
			// Mixed: complex + plain number
			{`IMSUM("3+4i",5)`, "8+4i"},
			// Mixed: plain number + complex
			{`IMSUM(5,"3+4i")`, "8+4i"},
			// Cancellation to zero
			{`IMSUM("3+4i","-3-4i")`, "0"},
			// Single argument
			{`IMSUM("3+4i")`, "3+4i"},
			// Single pure real
			{`IMSUM("7")`, "7"},
			// Single pure imaginary
			{`IMSUM("5i")`, "5i"},
			// j suffix
			{`IMSUM("1+2j","3+4j")`, "4+6j"},
			// Unit imaginary result
			{`IMSUM("3+i","2")`, "5+i"},
			// Negative result
			{`IMSUM("1+i","-3-2i")`, "-2-i"},
			// Result with negative imaginary
			{`IMSUM("3+2i","1-5i")`, "4-3i"},
			// Only real part cancels
			{`IMSUM("3+4i","-3+2i")`, "6i"},
			// Only imaginary part cancels
			{`IMSUM("3+4i","2-4i")`, "5"},
			// Decimal coefficients
			{`IMSUM("1.5+2.5i","0.5+0.5i")`, "2+3i"},
			// Large number of args
			{`IMSUM("1+i","1+i","1+i","1+i","1+i")`, "5+5i"},
			// Two plain numbers (no suffix in input)
			{`IMSUM(3,5)`, "8"},
			// Zero + complex
			{`IMSUM("0","3+4i")`, "3+4i"},
			// Negative real result, positive imaginary
			{`IMSUM("-5+2i","1+3i")`, "-4+5i"},
			// Unit imaginary inputs
			{`IMSUM("i","i")`, "2i"},
			// Negative unit imaginary
			{`IMSUM("-i","-i")`, "-2i"},
			// Mixed: j suffix with pure real
			{`IMSUM("3","2j")`, "3+2j"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number string
			{`IMSUM("invalid")`, ErrValNUM},
			{`IMSUM("")`, ErrValNUM},
			{`IMSUM("abc")`, ErrValNUM},
			{`IMSUM("3+4")`, ErrValNUM},
			{`IMSUM("3+4k")`, ErrValNUM},
			// Mixed suffix: i and j → #NUM!
			{`IMSUM("1+2i","3+4j")`, ErrValNUM},
			// Second arg invalid
			{`IMSUM("3+4i","invalid")`, ErrValNUM},
			// Boolean → #VALUE!
			{`IMSUM(TRUE)`, ErrValVALUE},
			{`IMSUM(FALSE)`, ErrValVALUE},
			{`IMSUM("3+4i",TRUE)`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMSUM(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMSUM(1/0) = %v, want error", got)
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		cf := evalCompile(t, "IMSUM()")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMSUM() = %v, want error", got)
		}
	})
}

func TestIMSUB(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic subtraction
			{`IMSUB("13+4i","5+3i")`, "8+i"},
			// Result is zero
			{`IMSUB("3+4i","3+4i")`, "0"},
			// Pure real subtraction
			{`IMSUB("10","3")`, "7"},
			// Pure imaginary subtraction
			{`IMSUB("5i","2i")`, "3i"},
			// Negative result
			{`IMSUB("1+i","3+4i")`, "-2-3i"},
			// j suffix
			{`IMSUB("5+3j","2+j")`, "3+2j"},
			// Mixed: complex minus number
			{`IMSUB("3+4i",5)`, "-2+4i"},
			// Mixed: number minus complex
			{`IMSUB(5,"3+4i")`, "2-4i"},
			// Both plain numbers
			{`IMSUB(10,3)`, "7"},
			// Subtract to get negative imaginary
			{`IMSUB("3+2i","1+5i")`, "2-3i"},
			// Unit imaginary result (positive)
			{`IMSUB("3+2i","3+i")`, "i"},
			// Unit imaginary result (negative)
			{`IMSUB("3+i","3+2i")`, "-i"},
			// Subtract zero
			{`IMSUB("3+4i","0")`, "3+4i"},
			// Subtract from zero
			{`IMSUB("0","3+4i")`, "-3-4i"},
			// Decimal coefficients
			{`IMSUB("5.5+3.5i","2.5+1.5i")`, "3+2i"},
			// Result with only real part
			{`IMSUB("3+4i","1+4i")`, "2"},
			// Result with only imaginary part
			{`IMSUB("3+4i","3+2i")`, "2i"},
			// Large numbers
			{`IMSUB("100+200i","50+100i")`, "50+100i"},
			// Negative inputs
			{`IMSUB("-3-4i","-1-2i")`, "-2-2i"},
			// Subtracting negative (should add)
			{`IMSUB("3+4i","-1-2i")`, "4+6i"},
			// j suffix with pure real
			{`IMSUB("3+2j","3")`, "2j"},
			// Pure number minus pure number (no suffix)
			{`IMSUB("8","3")`, "5"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number string
			{`IMSUB("invalid","3+4i")`, ErrValNUM},
			{`IMSUB("3+4i","invalid")`, ErrValNUM},
			{`IMSUB("","3+4i")`, ErrValNUM},
			{`IMSUB("3+4i","")`, ErrValNUM},
			{`IMSUB("abc","3+4i")`, ErrValNUM},
			{`IMSUB("3+4k","1+2i")`, ErrValNUM},
			// Mixed suffix: i and j → #NUM!
			{`IMSUB("1+2i","3+4j")`, ErrValNUM},
			{`IMSUB("1+2j","3+4i")`, ErrValNUM},
			// Boolean → #VALUE!
			{`IMSUB(TRUE,"3+4i")`, ErrValVALUE},
			{`IMSUB("3+4i",FALSE)`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMSUB(1/0,"3+4i")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMSUB(1/0,...) = %v, want error", got)
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"IMSUB()", `IMSUB("3+4i")`, `IMSUB("1","2","3")`} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestIMPRODUCT(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic: two complex numbers (3+4i)(5-3i) = 15-9i+20i-12i² = 27+11i
			{`IMPRODUCT("3+4i","5-3i")`, "27+11i"},
			// Scalar multiplication: (1+2i)*30 = 30+60i
			{`IMPRODUCT("1+2i",30)`, "30+60i"},
			// Complex * real string: (3+4i)*2 = 6+8i
			{`IMPRODUCT("3+4i","2")`, "6+8i"},
			// Three args: (1+i)(1+i)(1+i) = (2i)(1+i) = 2i+2i² = -2+2i
			{`IMPRODUCT("1+i","1+i","1+i")`, "-2+2i"},
			// Pure imaginary: i*i = i² = -1
			{`IMPRODUCT("i","i")`, "-1"},
			// Single arg: identity
			{`IMPRODUCT("3+4i")`, "3+4i"},
			// Multiply by zero
			{`IMPRODUCT("3+4i","0")`, "0"},
			// j suffix
			{`IMPRODUCT("1+2j","3+4j")`, "-5+10j"},
			// Negative: (-1-i)(1+i) = -1-i-i-i² = -1-2i+1 = -2i
			{`IMPRODUCT("-1-i","1+i")`, "-2i"},
			// Unit: i*(-i) = -i² = 1
			{`IMPRODUCT("i","-i")`, "1"},
			// Real * real
			{`IMPRODUCT("3","5")`, "15"},
			// Pure imaginary * pure imaginary: (2i)(3i) = 6i² = -6
			{`IMPRODUCT("2i","3i")`, "-6"},
			// Complex conjugate product: (3+4i)(3-4i) = 9+16 = 25
			{`IMPRODUCT("3+4i","3-4i")`, "25"},
			// Multiply by 1 (identity)
			{`IMPRODUCT("3+4i","1")`, "3+4i"},
			// Multiply by -1
			{`IMPRODUCT("3+4i","-1")`, "-3-4i"},
			// Four args: (1+i)(1+i)(1+i)(1+i) = (-2+2i)(1+i) = -2-2i+2i+2i² = -4
			{`IMPRODUCT("1+i","1+i","1+i","1+i")`, "-4"},
			// Decimal coefficients: (1.5+2.5i)(2+0i) = 3+5i
			{`IMPRODUCT("1.5+2.5i","2")`, "3+5i"},
			// Mixed: plain number and complex
			{`IMPRODUCT(5,"3+4i")`, "15+20i"},
			// Two plain numbers
			{`IMPRODUCT(3,5)`, "15"},
			// Pure imaginary * real: (3i)(4) = 12i
			{`IMPRODUCT("3i","4")`, "12i"},
			// Negative real * complex: (-2)(3+4i) = -6-8i
			{`IMPRODUCT("-2","3+4i")`, "-6-8i"},
			// i * (1+i) = i + i² = -1+i
			{`IMPRODUCT("i","1+i")`, "-1+i"},
			// Single pure real
			{`IMPRODUCT("7")`, "7"},
			// Single pure imaginary
			{`IMPRODUCT("5i")`, "5i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number string
			{`IMPRODUCT("invalid")`, ErrValNUM},
			{`IMPRODUCT("")`, ErrValNUM},
			{`IMPRODUCT("abc")`, ErrValNUM},
			{`IMPRODUCT("3+4")`, ErrValNUM},
			{`IMPRODUCT("3+4k")`, ErrValNUM},
			// Mixed suffix: i and j → #NUM!
			{`IMPRODUCT("1+2i","3+4j")`, ErrValNUM},
			// Second arg invalid
			{`IMPRODUCT("3+4i","invalid")`, ErrValNUM},
			// Boolean → #VALUE!
			{`IMPRODUCT(TRUE)`, ErrValVALUE},
			{`IMPRODUCT(FALSE)`, ErrValVALUE},
			{`IMPRODUCT("3+4i",TRUE)`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMPRODUCT(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMPRODUCT(1/0) = %v, want error", got)
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		cf := evalCompile(t, "IMPRODUCT()")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMPRODUCT() = %v, want error", got)
		}
	})
}

func TestIMDIV(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic: (-238+240i)/(10+24i) = 5+12i
			{`IMDIV("-238+240i","10+24i")`, "5+12i"},
			// Divide by real: (6+8i)/2 = 3+4i
			{`IMDIV("6+8i","2")`, "3+4i"},
			// Divide by imaginary: 4/(2i) = 4*(0-2i)/(0+4) = -2i
			{`IMDIV("4","2i")`, "-2i"},
			// Self-division: (3+4i)/(3+4i) = 1
			{`IMDIV("3+4i","3+4i")`, "1"},
			// j suffix
			{`IMDIV("4+6j","2+j")`, "2.8+1.6j"},
			// Divide complex by plain number
			{`IMDIV("10+20i",5)`, "2+4i"},
			// Divide plain number by complex: 5/(1+2i) = 5(1-2i)/(1+4) = (5-10i)/5 = 1-2i
			{`IMDIV(5,"1+2i")`, "1-2i"},
			// Pure imaginary / pure imaginary: (6i)/(3i) = 2
			{`IMDIV("6i","3i")`, "2"},
			// Real / real
			{`IMDIV("10","5")`, "2"},
			// Divide by 1 (identity)
			{`IMDIV("3+4i","1")`, "3+4i"},
			// Divide by -1
			{`IMDIV("3+4i","-1")`, "-3-4i"},
			// Zero numerator
			{`IMDIV("0","3+4i")`, "0"},
			// Complex result: (1+i)/(1-i) = (1+i)²/2 = 2i/2 = i
			{`IMDIV("1+i","1-i")`, "i"},
			// Conjugate division: (3-4i)/(3+4i) = (9-16-24i)/(9+16) = (-7-24i)/25
			{`IMDIV("3-4i","3+4i")`, "-0.28-0.96i"},
			// Negative numerator and denominator: (-6-8i)/(-2) = 3+4i
			{`IMDIV("-6-8i","-2")`, "3+4i"},
			// Two plain numbers
			{`IMDIV(10,2)`, "5"},
			// Unit imaginary numerator: i/(1+i) = i(1-i)/2 = (i+1)/2 = 0.5+0.5i
			{`IMDIV("i","1+i")`, "0.5+0.5i"},
			// Large numbers: (100+200i)/(10+20i) = (1000+4000+2000i-2000i)/500 = 10
			{`IMDIV("100+200i","10+20i")`, "10"},
			// Divide pure imaginary by real: (8i)/4 = 2i
			{`IMDIV("8i","4")`, "2i"},
			// Pure imaginary divided by pure imaginary
			{`IMDIV("4i","2i")`, "2"},
			// Negative imaginary result: (2)/(i) = -2i
			{`IMDIV("2","i")`, "-2i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Division by zero: both real and imag of divisor are 0
			{`IMDIV("3+4i","0")`, ErrValNUM},
			// Invalid complex number string
			{`IMDIV("invalid","3+4i")`, ErrValNUM},
			{`IMDIV("3+4i","invalid")`, ErrValNUM},
			{`IMDIV("","3+4i")`, ErrValNUM},
			{`IMDIV("3+4i","")`, ErrValNUM},
			{`IMDIV("abc","3+4i")`, ErrValNUM},
			{`IMDIV("3+4k","1+2i")`, ErrValNUM},
			// Mixed suffix: i and j → #NUM!
			{`IMDIV("1+2i","3+4j")`, ErrValNUM},
			{`IMDIV("1+2j","3+4i")`, ErrValNUM},
			// Boolean → #VALUE!
			{`IMDIV(TRUE,"3+4i")`, ErrValVALUE},
			{`IMDIV("3+4i",FALSE)`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMDIV(1/0,"3+4i")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMDIV(1/0,...) = %v, want error", got)
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"IMDIV()", `IMDIV("3+4i")`, `IMDIV("1","2","3")`} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval: %v", err)
				}
				if got.Type != ValueError {
					t.Errorf("%s = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestIMARGUMENT(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
		}{
			// Standard complex numbers
			{`IMARGUMENT("3+4i")`, math.Atan2(4, 3)},
			{`IMARGUMENT("1+i")`, math.Pi / 4},
			{`IMARGUMENT("1-i")`, -math.Pi / 4},
			{`IMARGUMENT("-1+i")`, 3 * math.Pi / 4},
			{`IMARGUMENT("-1-i")`, -3 * math.Pi / 4},

			// Pure imaginary
			{`IMARGUMENT("i")`, math.Pi / 2},
			{`IMARGUMENT("-i")`, -math.Pi / 2},
			{`IMARGUMENT("2i")`, math.Pi / 2},
			{`IMARGUMENT("-3i")`, -math.Pi / 2},

			// Pure real positive → 0
			{`IMARGUMENT("1")`, 0},
			{`IMARGUMENT("5")`, 0},
			{`IMARGUMENT("3.14")`, 0},

			// Pure real negative → π
			{`IMARGUMENT("-1")`, math.Pi},
			{`IMARGUMENT("-5")`, math.Pi},

			// Numeric input (not string)
			{"IMARGUMENT(1)", 0},
			{"IMARGUMENT(-1)", math.Pi},
			{"IMARGUMENT(5)", 0},
			{"IMARGUMENT(-7)", math.Pi},

			// Boolean: TRUE=1 (positive real)
			{"IMARGUMENT(TRUE)", 0},

			// j suffix
			{`IMARGUMENT("1+j")`, math.Pi / 4},
			{`IMARGUMENT("3+4j")`, math.Atan2(4, 3)},
			{`IMARGUMENT("j")`, math.Pi / 2},
			{`IMARGUMENT("-j")`, -math.Pi / 2},

			// Decimal coefficients
			{`IMARGUMENT("1.5+2.5i")`, math.Atan2(2.5, 1.5)},
			{`IMARGUMENT("-1.5+2.5i")`, math.Atan2(2.5, -1.5)},

			// Combined with COMPLEX
			{`IMARGUMENT(COMPLEX(3,4))`, math.Atan2(4, 3)},
			{`IMARGUMENT(COMPLEX(1,1))`, math.Pi / 4},
			{`IMARGUMENT(COMPLEX(0,1))`, math.Pi / 2},
			{`IMARGUMENT(COMPLEX(-1,0))`, math.Pi},

			// Explicit zero parts
			{`IMARGUMENT("0+4i")`, math.Pi / 2},
			{`IMARGUMENT("0-4i")`, -math.Pi / 2},
			{`IMARGUMENT("3+0i")`, 0},
			{`IMARGUMENT("-3+0i")`, math.Pi},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber {
					t.Fatalf("Eval(%q) = %v, want number %v", tt.formula, got, tt.want)
				}
				if math.Abs(got.Num-tt.want) > 1e-12 {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Zero: argument is undefined
			{`IMARGUMENT("0")`, ErrValDIV0},
			{"IMARGUMENT(0)", ErrValDIV0},
			{`IMARGUMENT("0+0i")`, ErrValDIV0},
			{`IMARGUMENT(COMPLEX(0,0))`, ErrValDIV0},
			{"IMARGUMENT(FALSE)", ErrValDIV0},

			// Invalid complex number strings
			{`IMARGUMENT("invalid")`, ErrValNUM},
			{`IMARGUMENT("")`, ErrValNUM},
			{`IMARGUMENT("abc")`, ErrValNUM},
			{`IMARGUMENT("3+4")`, ErrValNUM},
			{`IMARGUMENT("3+4k")`, ErrValNUM},
			{`IMARGUMENT("+")`, ErrValNUM},

			// Wrong number of arguments
			{`IMARGUMENT()`, ErrValVALUE},
			{`IMARGUMENT("1","2")`, ErrValVALUE},

			// Error propagation
			{`IMARGUMENT(1/0)`, ErrValDIV0},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})
}

func TestIMCONJUGATE(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Basic conjugation
			{`IMCONJUGATE("3+4i")`, "3-4i"},
			{`IMCONJUGATE("3-4i")`, "3+4i"},
			{`IMCONJUGATE("-3+4i")`, "-3-4i"},
			{`IMCONJUGATE("-3-4i")`, "-3+4i"},

			// Pure imaginary
			{`IMCONJUGATE("i")`, "-i"},
			{`IMCONJUGATE("-i")`, "i"},
			{`IMCONJUGATE("4i")`, "-4i"},
			{`IMCONJUGATE("-4i")`, "4i"},
			{`IMCONJUGATE("2i")`, "-2i"},
			{`IMCONJUGATE("-2i")`, "2i"},

			// Pure real (unchanged)
			{`IMCONJUGATE("5")`, "5"},
			{`IMCONJUGATE("-5")`, "-5"},
			{`IMCONJUGATE("3.14")`, "3.14"},

			// Zero
			{`IMCONJUGATE("0")`, "0"},
			{`IMCONJUGATE("0+0i")`, "0"},

			// j suffix
			{`IMCONJUGATE("3+4j")`, "3-4j"},
			{`IMCONJUGATE("3-4j")`, "3+4j"},
			{`IMCONJUGATE("j")`, "-j"},
			{`IMCONJUGATE("-j")`, "j"},

			// Decimal coefficients
			{`IMCONJUGATE("1.5+2.5i")`, "1.5-2.5i"},
			{`IMCONJUGATE("1.5-2.5i")`, "1.5+2.5i"},

			// Unit imaginary with real
			{`IMCONJUGATE("3+i")`, "3-i"},
			{`IMCONJUGATE("3-i")`, "3+i"},

			// Combined with COMPLEX
			{`IMCONJUGATE(COMPLEX(3,4))`, "3-4i"},
			{`IMCONJUGATE(COMPLEX(3,-4))`, "3+4i"},
			{`IMCONJUGATE(COMPLEX(0,1))`, "-i"},
			{`IMCONJUGATE(COMPLEX(5,0))`, "5"},
			{`IMCONJUGATE(COMPLEX(0,0))`, "0"},

			// Explicit zero parts
			{`IMCONJUGATE("0+4i")`, "-4i"},
			{`IMCONJUGATE("0-4i")`, "4i"},
			{`IMCONJUGATE("3+0i")`, "3"},

			// Numeric input (treat as real)
			{"IMCONJUGATE(5)", "5"},
			{"IMCONJUGATE(0)", "0"},
			{"IMCONJUGATE(-7)", "-7"},

			// Boolean: TRUE=1, FALSE=0
			{"IMCONJUGATE(TRUE)", "1"},
			{"IMCONJUGATE(FALSE)", "0"},

			// Large numbers
			{`IMCONJUGATE("100+200i")`, "100-200i"},
			{`IMCONJUGATE("-100-200i")`, "-100+200i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMCONJUGATE("invalid")`, ErrValNUM},
			{`IMCONJUGATE("")`, ErrValNUM},
			{`IMCONJUGATE("abc")`, ErrValNUM},
			{`IMCONJUGATE("3+4")`, ErrValNUM},
			{`IMCONJUGATE("3+4k")`, ErrValNUM},
			{`IMCONJUGATE("+")`, ErrValNUM},

			// Wrong number of arguments
			{`IMCONJUGATE()`, ErrValVALUE},
			{`IMCONJUGATE("1","2")`, ErrValVALUE},

			// Error propagation
			{`IMCONJUGATE(1/0)`, ErrValDIV0},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})
}

func TestIMSQRT(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Pure real: positive
			{`IMSQRT("4")`, "2"},
			{`IMSQRT("9")`, "3"},
			{`IMSQRT("1")`, "1"},
			{`IMSQRT("0.25")`, "0.5"},

			// Pure real: zero
			{`IMSQRT("0")`, "0"},

			// Pure real: negative → purely imaginary
			{`IMSQRT("-1")`, "i"},
			{`IMSQRT("-4")`, "2i"},
			{`IMSQRT("-9")`, "3i"},

			// Complex: sqrt(3+4i) = 2+i (since (2+i)² = 3+4i)
			{`IMSQRT("3+4i")`, "2+i"},

			// Complex: sqrt(5+12i) = 3+2i (since (3+2i)² = 5+12i)
			{`IMSQRT("5+12i")`, "3+2i"},

			// Complex: sqrt(8+6i) = 3+i (since (3+i)² = 8+6i)
			{`IMSQRT("8+6i")`, "3+i"},

			// Complex: sqrt(-3+4i) = 1+2i (since (1+2i)² = -3+4i)
			{`IMSQRT("-3+4i")`, "1+2i"},

			// Complex: sqrt(-3-4i) = 1-2i (since (1-2i)² = -3-4i)
			{`IMSQRT("-3-4i")`, "1-2i"},

			// j suffix
			{`IMSQRT("3+4j")`, "2+j"},

			// Decimal result: sqrt(2i) = 1+i (since (1+i)² = 2i)
			{`IMSQRT("2i")`, "1+i"},

			// Numeric input (not string)
			{`IMSQRT(4)`, "2"},
			{`IMSQRT(0)`, "0"},

			// COMPLEX composition
			{`IMSQRT(COMPLEX(3,4))`, "2+i"},
			{`IMSQRT(COMPLEX(0,0))`, "0"},
			{`IMSQRT(COMPLEX(5,12))`, "3+2i"},
			{`IMSQRT(COMPLEX(-1,0))`, "i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMSQRT("invalid")`, ErrValNUM},
			{`IMSQRT("")`, ErrValNUM},
			{`IMSQRT("abc")`, ErrValNUM},
			{`IMSQRT("3+4")`, ErrValNUM},
			{`IMSQRT("3+4k")`, ErrValNUM},
			{`IMSQRT("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMSQRT(TRUE)`, ErrValVALUE},
			{`IMSQRT(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMSQRT()`, ErrValVALUE},
			{`IMSQRT("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMSQRT(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMSQRT(1/0) = %v, want error", got)
		}
	})
}

func TestIMPOWER(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// i^2 = -1
			{`IMPOWER("i",2)`, "-1"},

			// i^4 = 1
			{`IMPOWER("i",4)`, "1"},

			// Pure real: 2^3 = 8
			{`IMPOWER("2",3)`, "8"},

			// Pure real: 3^2 = 9
			{`IMPOWER("3",2)`, "9"},

			// (1+i)^2 = 2i
			{`IMPOWER("1+i",2)`, "2i"},

			// Power 0: anything^0 = 1
			{`IMPOWER("3+4i",0)`, "1"},
			{`IMPOWER("i",0)`, "1"},
			{`IMPOWER("5",0)`, "1"},
			{`IMPOWER("0",0)`, "1"},

			// Power 1: returns same
			{`IMPOWER("3+4i",1)`, "3+4i"},
			{`IMPOWER("i",1)`, "i"},
			{`IMPOWER("5",1)`, "5"},

			// 0^positive = 0
			{`IMPOWER("0",2)`, "0"},
			{`IMPOWER("0",5)`, "0"},

			// Negative power: 2^-1 = 0.5
			{`IMPOWER("2",-1)`, "0.5"},

			// Decimal power: 4^0.5 = 2
			{`IMPOWER("4",0.5)`, "2"},

			// j suffix
			{`IMPOWER("j",2)`, "-1"},
			{`IMPOWER("1+j",2)`, "2j"},

			// Numeric first arg
			{`IMPOWER(2,3)`, "8"},
			{`IMPOWER(4,0.5)`, "2"},

			// (1+i)^4 = -4
			{`IMPOWER("1+i",4)`, "-4"},

			// Complex power: (2+3i)^2 = 4+12i-9 = -5+12i
			{`IMPOWER("2+3i",2)`, "-5+12i"},

			// COMPLEX composition
			{`IMPOWER(COMPLEX(1,1),2)`, "2i"},
			{`IMPOWER(COMPLEX(0,1),2)`, "-1"},
			{`IMPOWER(COMPLEX(2,0),3)`, "8"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number string
			{`IMPOWER("invalid",2)`, ErrValNUM},
			{`IMPOWER("",2)`, ErrValNUM},
			{`IMPOWER("abc",2)`, ErrValNUM},
			{`IMPOWER("3+4",2)`, ErrValNUM},
			{`IMPOWER("3+4k",2)`, ErrValNUM},

			// 0^negative → #NUM! (division by zero)
			{`IMPOWER("0",-1)`, ErrValNUM},
			{`IMPOWER("0",-2)`, ErrValNUM},

			// Boolean first arg → #VALUE!
			{`IMPOWER(TRUE,2)`, ErrValVALUE},
			{`IMPOWER(FALSE,2)`, ErrValVALUE},

			// Wrong arg count
			{`IMPOWER("1")`, ErrValVALUE},
			{`IMPOWER()`, ErrValVALUE},
			{`IMPOWER("1",2,3)`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMPOWER(1/0,2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMPOWER(1/0,2) = %v, want error", got)
		}
	})

	t.Run("error propagation second arg", func(t *testing.T) {
		cf := evalCompile(t, `IMPOWER("1+i",1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMPOWER(\"1+i\",1/0) = %v, want error", got)
		}
	})
}

func TestIMEXP(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Zero → e^0 = 1
			{`IMEXP("0")`, "1"},
			{`IMEXP(0)`, "1"},

			// Pure real positive
			{`IMEXP("1")`, "2.718281828459045"},
			{`IMEXP(1)`, "2.718281828459045"},
			{`IMEXP("2")`, "7.38905609893065"},
			{`IMEXP("0.5")`, "1.6487212707001282"},

			// Pure real negative
			{`IMEXP("-1")`, "0.36787944117144233"},
			{`IMEXP("-2")`, "0.1353352832366127"},

			// Pure imaginary: e^(i) = cos(1)+sin(1)*i
			{`IMEXP("i")`, "0.5403023058681398+0.8414709848078965i"},
			{`IMEXP("-i")`, "0.5403023058681398-0.8414709848078965i"},
			{`IMEXP("2i")`, "-0.4161468365471424+0.9092974268256816i"},

			// Complex: e^(1+i) = e*(cos(1)+sin(1)*i)
			{`IMEXP("1+i")`, "1.4686939399158851+2.2873552871788423i"},
			{`IMEXP("1-i")`, "1.4686939399158851-2.2873552871788423i"},

			// Euler's formula: e^(πi) ≈ -1
			{`IMEXP("3.14159265358979i")`, "-1"},

			// j suffix
			{`IMEXP("j")`, "0.5403023058681398+0.8414709848078965j"},
			{`IMEXP("1+j")`, "1.4686939399158851+2.2873552871788423j"},

			// COMPLEX composition
			{`IMEXP(COMPLEX(0,0))`, "1"},
			{`IMEXP(COMPLEX(1,0))`, "2.718281828459045"},
			{`IMEXP(COMPLEX(0,1))`, "0.5403023058681398+0.8414709848078965i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMEXP("invalid")`, ErrValNUM},
			{`IMEXP("")`, ErrValNUM},
			{`IMEXP("abc")`, ErrValNUM},
			{`IMEXP("3+4")`, ErrValNUM},
			{`IMEXP("3+4k")`, ErrValNUM},
			{`IMEXP("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMEXP(TRUE)`, ErrValVALUE},
			{`IMEXP(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMEXP()`, ErrValVALUE},
			{`IMEXP("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMEXP(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMEXP(1/0) = %v, want error", got)
		}
	})
}

func TestIMLN(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// ln(1) = 0
			{`IMLN("1")`, "0"},
			{`IMLN(1)`, "0"},

			// ln(e) = 1
			{`IMLN("2.718281828459045")`, "1"},

			// ln(-1) = πi
			{`IMLN("-1")`, "3.141592653589793i"},

			// ln(i) = (π/2)i
			{`IMLN("i")`, "1.5707963267948966i"},

			// ln(-i) = -(π/2)i
			{`IMLN("-i")`, "-1.5707963267948966i"},

			// ln(3+4i) = ln(5) + atan2(4,3)*i
			{`IMLN("3+4i")`, "1.6094379124341003+0.9272952180016122i"},

			// Pure real
			{`IMLN("2")`, "0.6931471805599453"},
			{`IMLN("5")`, "1.6094379124341003"},
			{`IMLN("10")`, "2.302585092994046"},

			// Negative real
			{`IMLN("-2")`, "0.6931471805599453+3.141592653589793i"},

			// Complex
			{`IMLN("1+i")`, "0.3465735902799727+0.7853981633974483i"},
			{`IMLN("-1-i")`, "0.3465735902799727-2.356194490192345i"},

			// j suffix
			{`IMLN("j")`, "1.5707963267948966j"},
			{`IMLN("3+4j")`, "1.6094379124341003+0.9272952180016122j"},

			// COMPLEX composition
			{`IMLN(COMPLEX(1,0))`, "0"},
			{`IMLN(COMPLEX(3,4))`, "1.6094379124341003+0.9272952180016122i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("log of zero", func(t *testing.T) {
		// IMLN(0) → #NUM!
		cf := evalCompile(t, `IMLN("0")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("IMLN(\"0\") = %v, want #NUM!", got)
		}

		// Also test with numeric 0
		cf = evalCompile(t, `IMLN(0)`)
		got, err = Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("IMLN(0) = %v, want #NUM!", got)
		}
	})

	t.Run("IFERROR wraps log of zero", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(IMLN("0"),"err")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "err" {
			t.Errorf(`IFERROR(IMLN("0"),"err") = %v, want "err"`, got)
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMLN("invalid")`, ErrValNUM},
			{`IMLN("")`, ErrValNUM},
			{`IMLN("abc")`, ErrValNUM},
			{`IMLN("3+4")`, ErrValNUM},
			{`IMLN("3+4k")`, ErrValNUM},
			{`IMLN("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMLN(TRUE)`, ErrValVALUE},
			{`IMLN(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMLN()`, ErrValVALUE},
			{`IMLN("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMLN(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMLN(1/0) = %v, want error", got)
		}
	})
}

func TestIMLOG2(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Powers of 2
			{`IMLOG2("1")`, "0"},
			{`IMLOG2("2")`, "1"},
			{`IMLOG2("4")`, "2"},
			{`IMLOG2("8")`, "3"},
			{`IMLOG2("16")`, "4"},
			{`IMLOG2("32")`, "5"},
			{`IMLOG2("0.5")`, "-1"},
			{`IMLOG2(1)`, "0"},
			{`IMLOG2(4)`, "2"},
			{`IMLOG2(8)`, "3"},

			// Complex
			{`IMLOG2("3+4i")`, "2.321928094887362+1.3378042124509761i"},
			{`IMLOG2("i")`, "2.266180070913597i"},
			{`IMLOG2("-1")`, "4.532360141827194i"},
			{`IMLOG2("1+i")`, "0.5000000000000001+1.1330900354567985i"},

			// j suffix
			{`IMLOG2("j")`, "2.266180070913597j"},
			{`IMLOG2("3+4j")`, "2.321928094887362+1.3378042124509761j"},

			// COMPLEX composition
			{`IMLOG2(COMPLEX(4,0))`, "2"},
			{`IMLOG2(COMPLEX(3,4))`, "2.321928094887362+1.3378042124509761i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("log of zero", func(t *testing.T) {
		cf := evalCompile(t, `IMLOG2("0")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("IMLOG2(\"0\") = %v, want #NUM!", got)
		}

		cf = evalCompile(t, `IMLOG2(0)`)
		got, err = Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("IMLOG2(0) = %v, want #NUM!", got)
		}
	})

	t.Run("IFERROR wraps log of zero", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(IMLOG2("0"),"err")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "err" {
			t.Errorf(`IFERROR(IMLOG2("0"),"err") = %v, want "err"`, got)
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMLOG2("invalid")`, ErrValNUM},
			{`IMLOG2("")`, ErrValNUM},
			{`IMLOG2("abc")`, ErrValNUM},
			{`IMLOG2("3+4")`, ErrValNUM},
			{`IMLOG2("3+4k")`, ErrValNUM},
			{`IMLOG2("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMLOG2(TRUE)`, ErrValVALUE},
			{`IMLOG2(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMLOG2()`, ErrValVALUE},
			{`IMLOG2("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMLOG2(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMLOG2(1/0) = %v, want error", got)
		}
	})
}

func TestIMLOG10(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Powers of 10
			{`IMLOG10("1")`, "0"},
			{`IMLOG10("10")`, "1"},
			{`IMLOG10("100")`, "2"},
			{`IMLOG10("1000")`, "3"},
			{`IMLOG10("0.1")`, "-1"},
			{`IMLOG10("0.01")`, "-2"},
			{`IMLOG10(1)`, "0"},
			{`IMLOG10(10)`, "1"},
			{`IMLOG10(100)`, "2"},

			// Complex
			{`IMLOG10("3+4i")`, "0.6989700043360187+0.4027191962733731i"},
			{`IMLOG10("i")`, "0.6821881769209206i"},
			{`IMLOG10("-1")`, "1.3643763538418412i"},
			{`IMLOG10("1+i")`, "0.1505149978319906+0.3410940884604603i"},

			// j suffix
			{`IMLOG10("j")`, "0.6821881769209206j"},
			{`IMLOG10("3+4j")`, "0.6989700043360187+0.4027191962733731j"},

			// COMPLEX composition
			{`IMLOG10(COMPLEX(10,0))`, "1"},
			{`IMLOG10(COMPLEX(3,4))`, "0.6989700043360187+0.4027191962733731i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("log of zero", func(t *testing.T) {
		cf := evalCompile(t, `IMLOG10("0")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("IMLOG10(\"0\") = %v, want #NUM!", got)
		}

		cf = evalCompile(t, `IMLOG10(0)`)
		got, err = Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("IMLOG10(0) = %v, want #NUM!", got)
		}
	})

	t.Run("IFERROR wraps log of zero", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(IMLOG10("0"),"err")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "err" {
			t.Errorf(`IFERROR(IMLOG10("0"),"err") = %v, want "err"`, got)
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMLOG10("invalid")`, ErrValNUM},
			{`IMLOG10("")`, ErrValNUM},
			{`IMLOG10("abc")`, ErrValNUM},
			{`IMLOG10("3+4")`, ErrValNUM},
			{`IMLOG10("3+4k")`, ErrValNUM},
			{`IMLOG10("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMLOG10(TRUE)`, ErrValVALUE},
			{`IMLOG10(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMLOG10()`, ErrValVALUE},
			{`IMLOG10("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMLOG10(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMLOG10(1/0) = %v, want error", got)
		}
	})
}

func TestIMSIN(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Pure real: zero
			{`IMSIN("0")`, "0"},

			// Pure real: positive
			{`IMSIN("1")`, "0.8414709848078965"},

			// Pure real: negative
			{`IMSIN("-1")`, "-0.8414709848078965"},

			// Pure real: pi ≈ 0
			{`IMSIN("3.14159265358979")`, "0"},

			// Pure imaginary: positive
			{`IMSIN("i")`, "1.1752011936438014i"},

			// Pure imaginary: negative
			{`IMSIN("-i")`, "-1.1752011936438014i"},

			// Pure imaginary: 2i
			{`IMSIN("2i")`, "3.626860407847019i"},

			// Complex: 1+i
			{`IMSIN("1+i")`, "1.2984575814159773+0.6349639147847361i"},

			// Complex: 1-i
			{`IMSIN("1-i")`, "1.2984575814159773-0.6349639147847361i"},

			// Complex: -1-i
			{`IMSIN("-1-i")`, "-1.2984575814159773-0.6349639147847361i"},

			// Complex: 3+4i
			{`IMSIN("3+4i")`, "3.853738037919377-27.016813258003932i"},

			// Complex: 2+3i
			{`IMSIN("2+3i")`, "9.154499146911428-4.168906959966565i"},

			// j suffix
			{`IMSIN("3+4j")`, "3.853738037919377-27.016813258003932j"},

			// Numeric input (not string)
			{`IMSIN(1)`, "0.8414709848078965"},
			{`IMSIN(0)`, "0"},

			// COMPLEX composition
			{`IMSIN(COMPLEX(1,1))`, "1.2984575814159773+0.6349639147847361i"},
			{`IMSIN(COMPLEX(0,0))`, "0"},
			{`IMSIN(COMPLEX(3,4))`, "3.853738037919377-27.016813258003932i"},
			{`IMSIN(COMPLEX(0,1))`, "1.1752011936438014i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMSIN("invalid")`, ErrValNUM},
			{`IMSIN("")`, ErrValNUM},
			{`IMSIN("abc")`, ErrValNUM},
			{`IMSIN("3+4")`, ErrValNUM},
			{`IMSIN("3+4k")`, ErrValNUM},
			{`IMSIN("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMSIN(TRUE)`, ErrValVALUE},
			{`IMSIN(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMSIN()`, ErrValVALUE},
			{`IMSIN("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMSIN(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMSIN(1/0) = %v, want error", got)
		}
	})
}

func TestIMCOS(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Pure real: zero → cos(0) = 1
			{`IMCOS("0")`, "1"},

			// Pure real: positive
			{`IMCOS("1")`, "0.5403023058681398"},

			// Pure real: negative (cos is even)
			{`IMCOS("-1")`, "0.5403023058681398"},

			// Pure real: pi → -1
			{`IMCOS("3.14159265358979")`, "-1"},

			// Pure imaginary: positive → cosh(1)
			{`IMCOS("i")`, "1.5430806348152437"},

			// Pure imaginary: negative → cosh(1) (cos is even)
			{`IMCOS("-i")`, "1.5430806348152437"},

			// Pure imaginary: 2i → cosh(2)
			{`IMCOS("2i")`, "3.7621956910836314"},

			// Complex: 1+i
			{`IMCOS("1+i")`, "0.8337300251311491-0.9888977057628651i"},

			// Complex: 1-i
			{`IMCOS("1-i")`, "0.8337300251311491+0.9888977057628651i"},

			// Complex: -1-i (cos(-z)=cos(z))
			{`IMCOS("-1-i")`, "0.8337300251311491-0.9888977057628651i"},

			// Complex: 3+4i
			{`IMCOS("3+4i")`, "-27.034945603074224-3.851153334811777i"},

			// Complex: 2+3i
			{`IMCOS("2+3i")`, "-4.189625690968807-9.109227893755337i"},

			// j suffix
			{`IMCOS("3+4j")`, "-27.034945603074224-3.851153334811777j"},

			// Numeric input (not string)
			{`IMCOS(1)`, "0.5403023058681398"},
			{`IMCOS(0)`, "1"},

			// COMPLEX composition
			{`IMCOS(COMPLEX(1,1))`, "0.8337300251311491-0.9888977057628651i"},
			{`IMCOS(COMPLEX(0,0))`, "1"},
			{`IMCOS(COMPLEX(3,4))`, "-27.034945603074224-3.851153334811777i"},
			{`IMCOS(COMPLEX(0,1))`, "1.5430806348152437"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMCOS("invalid")`, ErrValNUM},
			{`IMCOS("")`, ErrValNUM},
			{`IMCOS("abc")`, ErrValNUM},
			{`IMCOS("3+4")`, ErrValNUM},
			{`IMCOS("3+4k")`, ErrValNUM},
			{`IMCOS("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMCOS(TRUE)`, ErrValVALUE},
			{`IMCOS(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMCOS()`, ErrValVALUE},
			{`IMCOS("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMCOS(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMCOS(1/0) = %v, want error", got)
		}
	})
}

func TestIMTAN(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Pure real: zero
			{`IMTAN("0")`, "0"},

			// Pure real: positive
			{`IMTAN("1")`, "1.557407724654902"},

			// Pure real: negative
			{`IMTAN("-1")`, "-1.557407724654902"},

			// Pure imaginary: positive → tanh(1)*i
			{`IMTAN("i")`, "0.761594155955765i"},

			// Pure imaginary: negative → -tanh(1)*i
			{`IMTAN("-i")`, "-0.761594155955765i"},

			// Pure imaginary: 2i → tanh(2)*i
			{`IMTAN("2i")`, "0.9640275800758168i"},

			// Complex: 1+i
			{`IMTAN("1+i")`, "0.2717525853195117+1.0839233273386948i"},

			// Complex: 1-i
			{`IMTAN("1-i")`, "0.2717525853195117-1.0839233273386948i"},

			// Complex: 3+4i
			{`IMTAN("3+4i")`, "-0.00018734620462947842+0.9993559873814732i"},

			// Complex: 2+3i
			{`IMTAN("2+3i")`, "-0.0037640256415042484+1.0032386273536098i"},

			// j suffix
			{`IMTAN("3+4j")`, "-0.00018734620462947842+0.9993559873814732j"},

			// Numeric input (not string)
			{`IMTAN(1)`, "1.557407724654902"},
			{`IMTAN(0)`, "0"},

			// COMPLEX composition
			{`IMTAN(COMPLEX(1,1))`, "0.2717525853195117+1.0839233273386948i"},
			{`IMTAN(COMPLEX(0,0))`, "0"},
			{`IMTAN(COMPLEX(3,4))`, "-0.00018734620462947842+0.9993559873814732i"},
			{`IMTAN(COMPLEX(0,1))`, "0.761594155955765i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMTAN("invalid")`, ErrValNUM},
			{`IMTAN("")`, ErrValNUM},
			{`IMTAN("abc")`, ErrValNUM},
			{`IMTAN("3+4")`, ErrValNUM},
			{`IMTAN("3+4k")`, ErrValNUM},
			{`IMTAN("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMTAN(TRUE)`, ErrValVALUE},
			{`IMTAN(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMTAN()`, ErrValVALUE},
			{`IMTAN("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMTAN(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMTAN(1/0) = %v, want error", got)
		}
	})
}

func TestIMSINH(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Pure real: zero
			{`IMSINH("0")`, "0"},

			// Pure real: positive
			{`IMSINH("1")`, "1.1752011936438014"},

			// Pure real: negative
			{`IMSINH("-1")`, "-1.1752011936438014"},

			// Pure imaginary: positive
			{`IMSINH("i")`, "0.8414709848078965i"},

			// Pure imaginary: negative
			{`IMSINH("-i")`, "-0.8414709848078965i"},

			// Complex: 1+i
			{`IMSINH("1+i")`, "0.6349639147847361+1.2984575814159773i"},

			// Complex: 1-i
			{`IMSINH("1-i")`, "0.6349639147847361-1.2984575814159773i"},

			// Complex: 3+4i
			{`IMSINH("3+4i")`, "-6.5481200409110025-7.61923172032141i"},

			// Complex: 2+3i
			{`IMSINH("2+3i")`, "-3.59056458998578+0.5309210862485197i"},

			// j suffix
			{`IMSINH("3+4j")`, "-6.5481200409110025-7.61923172032141j"},

			// Numeric input (not string)
			{`IMSINH(1)`, "1.1752011936438014"},
			{`IMSINH(0)`, "0"},

			// COMPLEX composition
			{`IMSINH(COMPLEX(1,1))`, "0.6349639147847361+1.2984575814159773i"},
			{`IMSINH(COMPLEX(0,0))`, "0"},
			{`IMSINH(COMPLEX(3,4))`, "-6.5481200409110025-7.61923172032141i"},
			{`IMSINH(COMPLEX(0,1))`, "0.8414709848078965i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMSINH("invalid")`, ErrValNUM},
			{`IMSINH("")`, ErrValNUM},
			{`IMSINH("abc")`, ErrValNUM},
			{`IMSINH("3+4")`, ErrValNUM},
			{`IMSINH("3+4k")`, ErrValNUM},
			{`IMSINH("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMSINH(TRUE)`, ErrValVALUE},
			{`IMSINH(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMSINH()`, ErrValVALUE},
			{`IMSINH("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMSINH(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMSINH(1/0) = %v, want error", got)
		}
	})
}

func TestIMCOSH(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Pure real: zero
			{`IMCOSH("0")`, "1"},

			// Pure real: positive
			{`IMCOSH("1")`, "1.5430806348152437"},

			// Pure real: negative (cosh is even)
			{`IMCOSH("-1")`, "1.5430806348152437"},

			// Pure imaginary: positive
			{`IMCOSH("i")`, "0.5403023058681398"},

			// Pure imaginary: negative (cosh is even)
			{`IMCOSH("-i")`, "0.5403023058681398"},

			// Complex: 1+i
			{`IMCOSH("1+i")`, "0.8337300251311491+0.9888977057628651i"},

			// Complex: 1-i
			{`IMCOSH("1-i")`, "0.8337300251311491-0.9888977057628651i"},

			// Complex: 3+4i
			{`IMCOSH("3+4i")`, "-6.580663040551157-7.581552742746545i"},

			// Complex: 2+3i
			{`IMCOSH("2+3i")`, "-3.7245455049153224+0.5118225699873846i"},

			// j suffix
			{`IMCOSH("3+4j")`, "-6.580663040551157-7.581552742746545j"},

			// Numeric input (not string)
			{`IMCOSH(1)`, "1.5430806348152437"},
			{`IMCOSH(0)`, "1"},

			// COMPLEX composition
			{`IMCOSH(COMPLEX(1,1))`, "0.8337300251311491+0.9888977057628651i"},
			{`IMCOSH(COMPLEX(0,0))`, "1"},
			{`IMCOSH(COMPLEX(3,4))`, "-6.580663040551157-7.581552742746545i"},
			{`IMCOSH(COMPLEX(0,1))`, "0.5403023058681398"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMCOSH("invalid")`, ErrValNUM},
			{`IMCOSH("")`, ErrValNUM},
			{`IMCOSH("abc")`, ErrValNUM},
			{`IMCOSH("3+4")`, ErrValNUM},
			{`IMCOSH("3+4k")`, ErrValNUM},
			{`IMCOSH("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMCOSH(TRUE)`, ErrValVALUE},
			{`IMCOSH(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMCOSH()`, ErrValVALUE},
			{`IMCOSH("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMCOSH(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMCOSH(1/0) = %v, want error", got)
		}
	})
}

func TestIMSEC(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// known result: IMSEC("4+3i")
			{`IMSEC("4+3i")`, "-0.0652940278579471-0.0752249603027732i"},

			// Pure real: zero → sec(0) = 1/cos(0) = 1
			{`IMSEC("0")`, "1"},

			// Pure real: positive
			{`IMSEC("1")`, "1.8508157176809252"},

			// Pure real: negative (sec is even)
			{`IMSEC("-1")`, "1.8508157176809252"},

			// Pure imaginary: positive
			{`IMSEC("i")`, "0.6480542736638853"},

			// Pure imaginary: negative (sech is even)
			{`IMSEC("-i")`, "0.6480542736638853"},

			// Complex: 1+i
			{`IMSEC("1+i")`, "0.4983370305551868+0.591083841721045i"},

			// Complex: 1-i
			{`IMSEC("1-i")`, "0.4983370305551868-0.591083841721045i"},

			// Complex: 2+3i
			{`IMSEC("2+3i")`, "-0.041674964411144266+0.0906111371962376i"},

			// j suffix
			{`IMSEC("3+4j")`, "-0.03625349691586887+0.00516434460775318j"},

			// Numeric input
			{`IMSEC(1)`, "1.8508157176809252"},
			{`IMSEC(0)`, "1"},

			// Negative numeric input
			{`IMSEC(-1)`, "1.8508157176809252"},

			// COMPLEX composition
			{`IMSEC(COMPLEX(1,1))`, "0.4983370305551868+0.591083841721045i"},
			{`IMSEC(COMPLEX(0,0))`, "1"},

			// Large imaginary part
			{`IMSEC("0+10i")`, "0.00009079985933781"},

			// Pure imaginary with j suffix
			{`IMSEC("2j")`, "0.2658022288340797"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMSEC("invalid")`, ErrValNUM},
			{`IMSEC("")`, ErrValNUM},
			{`IMSEC("abc")`, ErrValNUM},
			{`IMSEC("3+4")`, ErrValNUM},
			{`IMSEC("3+4k")`, ErrValNUM},
			{`IMSEC("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMSEC(TRUE)`, ErrValVALUE},
			{`IMSEC(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMSEC()`, ErrValVALUE},
			{`IMSEC("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMSEC(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMSEC(1/0) = %v, want error", got)
		}
	})

	t.Run("array input", func(t *testing.T) {
		cf := evalCompile(t, `IMSEC({"1","i"})`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("expected array, got %v", got.Type)
		}
		if len(got.Array) != 1 || len(got.Array[0]) != 2 {
			t.Fatalf("expected 1x2 array, got %dx%d", len(got.Array), len(got.Array[0]))
		}
	})
}

func TestIMSECH(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Pure real: zero → sech(0) = 1/cosh(0) = 1
			{`IMSECH("0")`, "1"},

			// Pure real: positive
			{`IMSECH("1")`, "0.6480542736638853"},

			// Pure real: negative (sech is even)
			{`IMSECH("-1")`, "0.6480542736638853"},

			// Pure imaginary: positive
			{`IMSECH("i")`, "1.8508157176809252"},

			// Pure imaginary: negative
			{`IMSECH("-i")`, "1.8508157176809252"},

			// Complex: 1+i
			{`IMSECH("1+i")`, "0.4983370305551868-0.591083841721045i"},

			// Complex: 1-i
			{`IMSECH("1-i")`, "0.4983370305551868+0.591083841721045i"},

			// Complex: 3+4i
			{`IMSECH("3+4i")`, "-0.06529402785794705+0.07522496030277323i"},

			// Complex: 2+3i
			{`IMSECH("2+3i")`, "-0.26351297515838934-0.03621163655876852i"},

			// j suffix
			{`IMSECH("3+4j")`, "-0.06529402785794705+0.07522496030277323j"},

			// Numeric input
			{`IMSECH(1)`, "0.6480542736638853"},
			{`IMSECH(0)`, "1"},

			// COMPLEX composition
			{`IMSECH(COMPLEX(1,1))`, "0.4983370305551868-0.591083841721045i"},
			{`IMSECH(COMPLEX(0,0))`, "1"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// Invalid complex number strings
			{`IMSECH("invalid")`, ErrValNUM},
			{`IMSECH("")`, ErrValNUM},
			{`IMSECH("abc")`, ErrValNUM},
			{`IMSECH("3+4")`, ErrValNUM},
			{`IMSECH("3+4k")`, ErrValNUM},
			{`IMSECH("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMSECH(TRUE)`, ErrValVALUE},
			{`IMSECH(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMSECH()`, ErrValVALUE},
			{`IMSECH("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMSECH(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMSECH(1/0) = %v, want error", got)
		}
	})
}

func TestIMCSC(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Pure real: positive
			{`IMCSC("1")`, "1.1883951057781212"},

			// Pure real: negative
			{`IMCSC("-1")`, "-1.1883951057781212"},

			// Pure imaginary: positive
			{`IMCSC("i")`, "-0.8509181282393217i"},

			// Pure imaginary: negative
			{`IMCSC("-i")`, "0.8509181282393217i"},

			// Complex: 1+i
			{`IMCSC("1+i")`, "0.6215180171704283-0.3039310016284264i"},

			// Complex: 1-i
			{`IMCSC("1-i")`, "0.6215180171704283+0.3039310016284264i"},

			// Complex: 3+4i
			{`IMCSC("3+4i")`, "0.005174473184019398+0.03627588962862602i"},

			// Complex: 2+3i
			{`IMCSC("2+3i")`, "0.09047320975320745+0.04120098628857414i"},

			// j suffix
			{`IMCSC("3+4j")`, "0.005174473184019398+0.03627588962862602j"},

			// Numeric input
			{`IMCSC(1)`, "1.1883951057781212"},

			// COMPLEX composition
			{`IMCSC(COMPLEX(1,1))`, "0.6215180171704283-0.3039310016284264i"},
			{`IMCSC(COMPLEX(0,1))`, "-0.8509181282393217i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// csc(0) = 1/sin(0) → #NUM!
			{`IMCSC("0")`, ErrValNUM},
			{`IMCSC(0)`, ErrValNUM},

			// Invalid complex number strings
			{`IMCSC("invalid")`, ErrValNUM},
			{`IMCSC("")`, ErrValNUM},
			{`IMCSC("abc")`, ErrValNUM},
			{`IMCSC("3+4")`, ErrValNUM},
			{`IMCSC("3+4k")`, ErrValNUM},
			{`IMCSC("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMCSC(TRUE)`, ErrValVALUE},
			{`IMCSC(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMCSC()`, ErrValVALUE},
			{`IMCSC("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMCSC(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMCSC(1/0) = %v, want error", got)
		}
	})
}

func TestIMCOT(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Pure real: positive
			{`IMCOT("1")`, "0.6420926159343308"},

			// Pure real: negative
			{`IMCOT("-1")`, "-0.6420926159343308"},

			// Pure imaginary: positive
			{`IMCOT("i")`, "-1.3130352854993315i"},

			// Pure imaginary: negative
			{`IMCOT("-i")`, "1.3130352854993315i"},

			// Complex: 1+i
			{`IMCOT("1+i")`, "0.2176215618544027-0.8680141428959249i"},

			// Complex: 1-i
			{`IMCOT("1-i")`, "0.2176215618544027+0.8680141428959249i"},

			// Complex: 3+4i
			{`IMCOT("3+4i")`, "-0.00018758773798367492-1.0006443924715591i"},

			// Complex: 2+3i
			{`IMCOT("2+3i")`, "-0.0037397103763368027-0.9967577965693585i"},

			// j suffix
			{`IMCOT("3+4j")`, "-0.00018758773798367492-1.0006443924715591j"},

			// Numeric input
			{`IMCOT(1)`, "0.6420926159343308"},

			// COMPLEX composition
			{`IMCOT(COMPLEX(1,1))`, "0.2176215618544027-0.8680141428959249i"},
			{`IMCOT(COMPLEX(0,1))`, "-1.3130352854993315i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// cot(0) = cos(0)/sin(0) → #NUM!
			{`IMCOT("0")`, ErrValNUM},
			{`IMCOT(0)`, ErrValNUM},

			// Invalid complex number strings
			{`IMCOT("invalid")`, ErrValNUM},
			{`IMCOT("")`, ErrValNUM},
			{`IMCOT("abc")`, ErrValNUM},
			{`IMCOT("3+4")`, ErrValNUM},
			{`IMCOT("3+4k")`, ErrValNUM},
			{`IMCOT("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMCOT(TRUE)`, ErrValVALUE},
			{`IMCOT(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMCOT()`, ErrValVALUE},
			{`IMCOT("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMCOT(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMCOT(1/0) = %v, want error", got)
		}
	})
}

func TestIMCSCH(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns string", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Pure real: positive
			{`IMCSCH("1")`, "0.8509181282393217"},

			// Pure real: negative
			{`IMCSCH("-1")`, "-0.8509181282393217"},

			// Pure imaginary: positive
			{`IMCSCH("i")`, "-1.1883951057781212i"},

			// Pure imaginary: negative
			{`IMCSCH("-i")`, "1.1883951057781212i"},

			// Complex: 1+i
			{`IMCSCH("1+i")`, "0.3039310016284264-0.6215180171704283i"},

			// Complex: 1-i
			{`IMCSCH("1-i")`, "0.3039310016284264+0.6215180171704283i"},

			// Complex: 3+4i
			{`IMCSCH("3+4i")`, "-0.0648774713706355+0.0754898329158637i"},

			// Complex: 2+3i
			{`IMCSCH("2+3i")`, "-0.2725486614629402-0.04030057885689152i"},

			// j suffix
			{`IMCSCH("3+4j")`, "-0.0648774713706355+0.0754898329158637j"},

			// Numeric input
			{`IMCSCH(1)`, "0.8509181282393217"},

			// COMPLEX composition
			{`IMCSCH(COMPLEX(1,1))`, "0.3039310016284264-0.6215180171704283i"},
			{`IMCSCH(COMPLEX(0,1))`, "-1.1883951057781212i"},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				assertComplexStringClose(t, tt.formula, got, tt.want)
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// csch(0) = 1/sinh(0) → #NUM!
			{`IMCSCH("0")`, ErrValNUM},
			{`IMCSCH(0)`, ErrValNUM},

			// Invalid complex number strings
			{`IMCSCH("invalid")`, ErrValNUM},
			{`IMCSCH("")`, ErrValNUM},
			{`IMCSCH("abc")`, ErrValNUM},
			{`IMCSCH("3+4")`, ErrValNUM},
			{`IMCSCH("3+4k")`, ErrValNUM},
			{`IMCSCH("+")`, ErrValNUM},

			// Boolean → #VALUE!
			{`IMCSCH(TRUE)`, ErrValVALUE},
			{`IMCSCH(FALSE)`, ErrValVALUE},

			// Wrong arg count
			{`IMCSCH()`, ErrValVALUE},
			{`IMCSCH("1","2")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, `IMCSCH(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("IMCSCH(1/0) = %v, want error", got)
		}
	})
}

func TestBESSELI(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
			tol     float64
		}{
			// Excel documentation example
			{"BESSELI(1.5,1)", 0.981666428475166, 1e-12},

			// x = 0 cases
			{"BESSELI(0,0)", 1, 0},
			{"BESSELI(0,1)", 0, 0},
			{"BESSELI(0,5)", 0, 0},

			// Known values for n=0 (Excel-verified)
			{"BESSELI(1,0)", 1.2660658480342601, 1e-12},
			{"BESSELI(2,0)", 2.279585307296026, 1e-12},

			// Known values for n=1 (Excel-verified)
			{"BESSELI(1,1)", 0.5651590975819435, 1e-12},
			{"BESSELI(2,1)", 1.5906368572633083, 1e-12},

			// Negative x, odd n → negative result
			{"BESSELI(-1.5,1)", -0.981666428475166, 1e-12},
			{"BESSELI(-1,1)", -0.5651590975819435, 1e-12},

			// Negative x, even n → positive result
			{"BESSELI(-1.5,0)", 1.6467232021476756, 1e-12},
			{"BESSELI(-1.5,2)", 0.33783462087443816, 1e-12},
			{"BESSELI(-2,0)", 2.279585307296026, 1e-12},

			// n truncated: BESSELI(1.5, 1.9) same as BESSELI(1.5, 1)
			{"BESSELI(1.5,1.9)", 0.981666428475166, 1e-12},
			{"BESSELI(1.5,1.1)", 0.981666428475166, 1e-12},

			// Larger n
			{"BESSELI(1,5)", 0.0002714631559, 1e-9},
			{"BESSELI(1,10)", 2.752948e-10, 1e-15},

			// Larger x (Excel-verified)
			{"BESSELI(5,0)", 27.239871894394888, 1e-9},
			{"BESSELI(5,2)", 17.50561501211748, 1e-9},
			{"BESSELI(20,0)", 4.355828255628764e7, 1e2},

			// Boolean coercion (TRUE=1, FALSE=0)
			{"BESSELI(TRUE,0)", 1.2660658480342601, 1e-12},
			{"BESSELI(1,FALSE)", 1.2660658480342601, 1e-12},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber {
					t.Fatalf("Eval(%q) = %v (type %d), want number", tt.formula, got, got.Type)
				}
				diff := math.Abs(got.Num - tt.want)
				if tt.tol == 0 {
					if got.Num != tt.want {
						t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want)
					}
				} else if diff > tt.tol {
					t.Errorf("Eval(%q) = %v, want %v (diff=%e, tol=%e)", tt.formula, got.Num, tt.want, diff, tt.tol)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// n < 0 → #NUM!
			{"BESSELI(1,-1)", ErrValNUM},
			{"BESSELI(1,-5)", ErrValNUM},

			// Non-numeric x → #VALUE!
			{`BESSELI("abc",1)`, ErrValVALUE},

			// Non-numeric n → #VALUE!
			{`BESSELI(1,"abc")`, ErrValVALUE},

			// Both non-numeric → #VALUE!
			{`BESSELI("a","b")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"BESSELI()", "BESSELI(1)", "BESSELI(1,2,3)"} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", formula, err)
				}
				if got.Type != ValueError {
					t.Errorf("Eval(%q) = %v, want error", formula, got)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		for _, formula := range []string{"BESSELI(1/0,1)", "BESSELI(1,1/0)"} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", formula, err)
				}
				if got.Type != ValueError {
					t.Errorf("Eval(%q) = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestBESSELJ(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
			tol     float64
		}{
			// Excel documentation example
			{"BESSELJ(1.9,2)", 0.329925727692387, 1e-12},

			// x = 0 cases
			{"BESSELJ(0,0)", 1, 0},
			{"BESSELJ(0,1)", 0, 0},
			{"BESSELJ(0,5)", 0, 0},

			// Known values for n=0
			{"BESSELJ(1,0)", 0.765197686557967, 1e-12},
			{"BESSELJ(2,0)", 0.223890779141236, 1e-12},

			// Known values for n=1
			{"BESSELJ(1,1)", 0.440050585744934, 1e-12},
			{"BESSELJ(2,1)", 0.576724807756874, 1e-12},

			// Negative x, even n → positive (same as positive x)
			{"BESSELJ(-1.9,2)", 0.329925727692387, 1e-12},
			{"BESSELJ(-2,0)", 0.223890779141236, 1e-12},

			// Negative x, odd n → negative
			{"BESSELJ(-1.9,1)", -0.581157072713434, 1e-12},
			{"BESSELJ(-1,1)", -0.440050585744934, 1e-12},

			// n truncated: BESSELJ(1.9, 2.7) same as BESSELJ(1.9, 2)
			{"BESSELJ(1.9,2.7)", 0.329925727692387, 1e-12},
			{"BESSELJ(1.9,2.1)", 0.329925727692387, 1e-12},

			// Oscillating region: larger x
			{"BESSELJ(5,0)", -0.177596771314338, 1e-12},
			{"BESSELJ(5,2)", 0.046565116277752, 1e-12},
			{"BESSELJ(10,0)", -0.245935764451321, 1e-12},
			{"BESSELJ(10,2)", 0.254630313685208, 1e-12},
			{"BESSELJ(20,0)", 0.167024664372723, 1e-9},

			// Higher order n
			{"BESSELJ(1,5)", 0.000249757730211234, 1e-14},
			{"BESSELJ(5,5)", 0.261140546120170, 1e-12},

			// Boolean coercion (TRUE=1, FALSE=0)
			{"BESSELJ(TRUE,0)", 0.765197686557967, 1e-12},
			{"BESSELJ(1,FALSE)", 0.765197686557967, 1e-12},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber {
					t.Fatalf("Eval(%q) = %v (type %d), want number", tt.formula, got, got.Type)
				}
				diff := math.Abs(got.Num - tt.want)
				if tt.tol == 0 {
					if got.Num != tt.want {
						t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want)
					}
				} else if diff > tt.tol {
					t.Errorf("Eval(%q) = %v, want %v (diff=%e, tol=%e)", tt.formula, got.Num, tt.want, diff, tt.tol)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// n < 0 → #NUM!
			{"BESSELJ(1,-1)", ErrValNUM},
			{"BESSELJ(1,-5)", ErrValNUM},

			// Non-numeric x → #VALUE!
			{`BESSELJ("abc",1)`, ErrValVALUE},

			// Non-numeric n → #VALUE!
			{`BESSELJ(1,"abc")`, ErrValVALUE},

			// Both non-numeric → #VALUE!
			{`BESSELJ("a","b")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"BESSELJ()", "BESSELJ(1)", "BESSELJ(1,2,3)"} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", formula, err)
				}
				if got.Type != ValueError {
					t.Errorf("Eval(%q) = %v, want error", formula, got)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		for _, formula := range []string{"BESSELJ(1/0,1)", "BESSELJ(1,1/0)"} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", formula, err)
				}
				if got.Type != ValueError {
					t.Errorf("Eval(%q) = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestBESSELK(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
			tol     float64
		}{
			// Excel documentation example
			{"BESSELK(1.5,1)", 0.277387804, 1e-8},

			// Known values for n=0
			{"BESSELK(1,0)", 0.4210244382, 1e-7},
			{"BESSELK(2,0)", 0.1138938727, 1e-7},
			{"BESSELK(0.5,0)", 0.9244190716, 1e-7},

			// Known values for n=1
			{"BESSELK(1,1)", 0.6019072169, 1e-7},
			{"BESSELK(2,1)", 0.1398658818, 1e-7},

			// Larger x for n=0
			{"BESSELK(5,0)", 0.0036910982, 1e-8},
			{"BESSELK(10,0)", 0.00001778006232, 1e-11},

			// Larger x for n=1
			{"BESSELK(5,1)", 0.004044613162, 1e-8},

			// Higher order n
			{"BESSELK(5,3)", 0.008291768520, 1e-8},
			{"BESSELK(10,2)", 0.00002150982, 1e-9},
			{"BESSELK(2,3)", 0.6473853909, 1e-8},

			// Very small x (near divergence)
			{"BESSELK(0.01,0)", 4.721244471, 1e-6},
			{"BESSELK(0.01,1)", 99.97389659, 1e-4},

			// n truncated: BESSELK(1.5, 1.9) same as BESSELK(1.5, 1)
			{"BESSELK(1.5,1.9)", 0.277387804, 1e-8},
			{"BESSELK(1.5,1.1)", 0.277387804, 1e-8},

			// Large order n
			{"BESSELK(1,5)", 360.9605896, 1e-3},
			{"BESSELK(5,5)", 0.03270627362, 1e-7},

			// Boolean coercion (TRUE=1, FALSE=0)
			{"BESSELK(TRUE,0)", 0.4210244382, 1e-7},
			{"BESSELK(TRUE,FALSE)", 0.4210244382, 1e-7},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber {
					t.Fatalf("Eval(%q) = %v (type %d), want number", tt.formula, got, got.Type)
				}
				diff := math.Abs(got.Num - tt.want)
				if tt.tol == 0 {
					if got.Num != tt.want {
						t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want)
					}
				} else if diff > tt.tol {
					t.Errorf("Eval(%q) = %v, want %v (diff=%e, tol=%e)", tt.formula, got.Num, tt.want, diff, tt.tol)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// x <= 0 → #NUM!
			{"BESSELK(0,0)", ErrValNUM},
			{"BESSELK(0,1)", ErrValNUM},
			{"BESSELK(-1,0)", ErrValNUM},
			{"BESSELK(-0.5,1)", ErrValNUM},
			{"BESSELK(-5,2)", ErrValNUM},

			// n < 0 → #NUM!
			{"BESSELK(1,-1)", ErrValNUM},
			{"BESSELK(1,-5)", ErrValNUM},

			// Non-numeric x → #VALUE!
			{`BESSELK("abc",1)`, ErrValVALUE},

			// Non-numeric n → #VALUE!
			{`BESSELK(1,"abc")`, ErrValVALUE},

			// Both non-numeric → #VALUE!
			{`BESSELK("a","b")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"BESSELK()", "BESSELK(1)", "BESSELK(1,2,3)"} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", formula, err)
				}
				if got.Type != ValueError {
					t.Errorf("Eval(%q) = %v, want error", formula, got)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		for _, formula := range []string{"BESSELK(1/0,1)", "BESSELK(1,1/0)"} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", formula, err)
				}
				if got.Type != ValueError {
					t.Errorf("Eval(%q) = %v, want error", formula, got)
				}
			})
		}
	})
}

func TestBESSELY(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
			tol     float64
		}{
			// Excel documentation example
			{"BESSELY(2.5,1)", 0.145918138, 1e-8},

			// Known values for n=0
			{"BESSELY(1,0)", 0.0882569642, 1e-9},
			{"BESSELY(2,0)", 0.5103756726, 1e-9},
			{"BESSELY(0.5,0)", -0.4445187335, 1e-9},

			// Known values for n=1
			{"BESSELY(1,1)", -0.7812128213, 1e-9},
			{"BESSELY(2,1)", -0.1070324315, 1e-9},

			// Larger x for n=0
			{"BESSELY(5,0)", -0.3085176252, 1e-9},
			{"BESSELY(10,0)", 0.0556711673, 1e-9},

			// Larger x for n=2
			{"BESSELY(5,2)", 0.3676628826, 1e-9},

			// Higher order n
			{"BESSELY(10,3)", -0.2513626571838373, 1e-12},
			{"BESSELY(20,5)", -0.10003576788953243, 1e-12},

			// n truncated: BESSELY(2.5, 1.9) same as BESSELY(2.5, 1)
			{"BESSELY(2.5,1.9)", 0.145918138, 1e-8},
			{"BESSELY(2.5,1.1)", 0.145918138, 1e-8},

			// Very small x (near divergence)
			{"BESSELY(0.01,0)", -3.005455698, 1e-6},

			// Large order n
			{"BESSELY(1,5)", -260.4058666, 1e-3},
			{"BESSELY(5,5)", -0.4536948225, 1e-8},

			// Boolean coercion (TRUE=1, FALSE=0)
			{"BESSELY(TRUE,0)", 0.0882569642, 1e-9},
			{"BESSELY(TRUE,FALSE)", 0.0882569642, 1e-9},

			// Additional values
			{"BESSELY(3,0)", 0.3768500100, 1e-9},
			{"BESSELY(3,1)", 0.3246744248, 1e-9},
			{"BESSELY(0.1,0)", -1.5342386513, 1e-8},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueNumber {
					t.Fatalf("Eval(%q) = %v (type %d), want number", tt.formula, got, got.Type)
				}
				diff := math.Abs(got.Num - tt.want)
				if tt.tol == 0 {
					if got.Num != tt.want {
						t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want)
					}
				} else if diff > tt.tol {
					t.Errorf("Eval(%q) = %v, want %v (diff=%e, tol=%e)", tt.formula, got.Num, tt.want, diff, tt.tol)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		errTests := []struct {
			formula string
			wantErr ErrorValue
		}{
			// x <= 0 → #NUM!
			{"BESSELY(0,0)", ErrValNUM},
			{"BESSELY(0,1)", ErrValNUM},
			{"BESSELY(-1,0)", ErrValNUM},
			{"BESSELY(-0.5,1)", ErrValNUM},
			{"BESSELY(-5,2)", ErrValNUM},

			// n < 0 → #NUM!
			{"BESSELY(1,-1)", ErrValNUM},
			{"BESSELY(1,-5)", ErrValNUM},

			// Non-numeric x → #VALUE!
			{`BESSELY("abc",1)`, ErrValVALUE},

			// Non-numeric n → #VALUE!
			{`BESSELY(1,"abc")`, ErrValVALUE},

			// Both non-numeric → #VALUE!
			{`BESSELY("a","b")`, ErrValVALUE},
		}
		for _, tt := range errTests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected error %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("wrong arg count", func(t *testing.T) {
		for _, formula := range []string{"BESSELY()", "BESSELY(1)", "BESSELY(1,2,3)"} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", formula, err)
				}
				if got.Type != ValueError {
					t.Errorf("Eval(%q) = %v, want error", formula, got)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		for _, formula := range []string{"BESSELY(1/0,1)", "BESSELY(1,1/0)"} {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", formula, err)
				}
				if got.Type != ValueError {
					t.Errorf("Eval(%q) = %v, want error", formula, got)
				}
			})
		}
	})
}
