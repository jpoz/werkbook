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
			{`HEX2BIN("FFFFFFFFFF")`, "1111111111"},   // -1
			{`HEX2BIN("FFFFFFFFFE")`, "1111111110"},   // -2
			{`HEX2BIN("FFFFFFFE00")`, "1000000000"},   // -512

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
