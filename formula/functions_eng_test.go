package formula

import (
	"testing"
)

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

func TestDELTA(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
		}{
			// Excel doc examples
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
				if got.Type != ValueString || got.Str != tt.want {
					t.Errorf("Eval(%q) = %v, want %q", tt.formula, got, tt.want)
				}
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
				if got.Type != ValueString || got.Str != tt.want {
					t.Errorf("Eval(%q) = %v, want %q", tt.formula, got, tt.want)
				}
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
				if got.Type != ValueString || got.Str != tt.want {
					t.Errorf("Eval(%q) = %v, want %q", tt.formula, got, tt.want)
				}
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

func TestGESTEP(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("returns number", func(t *testing.T) {
		tests := []struct {
			formula string
			want    float64
		}{
			// Excel doc examples
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
