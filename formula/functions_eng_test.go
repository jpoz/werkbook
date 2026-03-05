package formula

import (
	"testing"
)

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
