package formula

import (
	"testing"
)

func TestAND(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    bool
	}{
		{"AND(TRUE,TRUE)", true},
		{"AND(TRUE,FALSE)", false},
		{"AND(FALSE,FALSE)", false},
		{"AND(TRUE,TRUE,TRUE)", true},
		{"AND(TRUE,TRUE,FALSE)", false},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueBool || got.Bool != tt.want {
			t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want)
		}
	}
}

func TestOR(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("basic boolean args", func(t *testing.T) {
		tests := []struct {
			formula string
			want    bool
		}{
			{"OR(TRUE)", true},
			{"OR(FALSE)", false},
			{"OR(TRUE,FALSE)", true},
			{"OR(FALSE,FALSE)", false},
			{"OR(TRUE,TRUE)", true},
			{"OR(FALSE,TRUE)", true},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("multiple args", func(t *testing.T) {
		tests := []struct {
			formula string
			want    bool
		}{
			{"OR(FALSE,FALSE,FALSE)", false},
			{"OR(FALSE,FALSE,TRUE)", true},
			{"OR(FALSE,TRUE,FALSE)", true},
			{"OR(TRUE,FALSE,FALSE)", true},
			{"OR(TRUE,TRUE,TRUE)", true},
			{"OR(FALSE,FALSE,FALSE,FALSE,TRUE)", true},
			{"OR(FALSE,FALSE,FALSE,FALSE,FALSE)", false},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("numeric args", func(t *testing.T) {
		tests := []struct {
			formula string
			want    bool
		}{
			{"OR(1)", true},
			{"OR(0)", false},
			{"OR(0,0)", false},
			{"OR(0,1)", true},
			{"OR(1,0)", true},
			{"OR(1,1)", true},
			{"OR(-1)", true},
			{"OR(0.5)", true},
			{"OR(0,0,0,0.001)", true},
			{"OR(99)", true},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("mixed boolean and numeric", func(t *testing.T) {
		tests := []struct {
			formula string
			want    bool
		}{
			{"OR(FALSE,0)", false},
			{"OR(TRUE,0)", true},
			{"OR(FALSE,1)", true},
			{"OR(0,FALSE)", false},
			{"OR(1,FALSE)", true},
			{"OR(0,TRUE)", true},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("direct string args return VALUE error", func(t *testing.T) {
		// In Excel, direct string arguments to OR cause #VALUE!
		tests := []string{
			`OR("hello")`,
			`OR("TRUE")`,
			`OR("FALSE")`,
			`OR("")`,
			`OR("text",FALSE)`,
			`OR("1")`,
			`OR("0")`,
		}
		for _, formula := range tests {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", formula, err)
				}
				if got.Type != ValueError || got.Err != ErrValVALUE {
					t.Errorf("Eval(%q) = %v, want #VALUE!", formula, got)
				}
			})
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		tests := []struct {
			formula string
			wantErr ErrorValue
		}{
			{"OR(1/0)", ErrValDIV0},
			{"OR(1/0,TRUE)", ErrValDIV0},
			{"OR(FALSE,1/0)", ErrValDIV0},
			{"OR(1/0,FALSE)", ErrValDIV0},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): unexpected Go error: %v", tt.formula, err)
				}
				if got.Type != ValueError || got.Err != tt.wantErr {
					t.Errorf("Eval(%q) = %v, want error %d", tt.formula, got, tt.wantErr)
				}
			})
		}
	})

	t.Run("short circuits past error when true found first", func(t *testing.T) {
		cf := evalCompile(t, "OR(TRUE,1/0)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("OR(TRUE,1/0) = %v, want true", got)
		}
	})

	t.Run("with expressions", func(t *testing.T) {
		tests := []struct {
			formula string
			want    bool
		}{
			{"OR(1>2)", false},
			{"OR(1<2)", true},
			{"OR(1>2,3>4)", false},
			{"OR(1>2,3<4)", true},
			{"OR(1=1,2=3)", true},
			{"OR(1=2,2=3)", false},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("range with cell references", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(false),
				{Col: 1, Row: 2}: BoolVal(false),
				{Col: 1, Row: 3}: BoolVal(true),
			},
		}
		cf := evalCompile(t, "OR(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("OR(A1:A3) = %v, want true", got)
		}
	})

	t.Run("range all false", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(false),
				{Col: 1, Row: 2}: BoolVal(false),
				{Col: 1, Row: 3}: BoolVal(false),
			},
		}
		cf := evalCompile(t, "OR(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("OR(A1:A3) = %v, want false", got)
		}
	})

	t.Run("range with error propagates", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(false),
				{Col: 1, Row: 2}: ErrorVal(ErrValNA),
				{Col: 1, Row: 3}: BoolVal(true),
			},
		}
		cf := evalCompile(t, "OR(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("OR(A1:A3) = %v, want #N/A error", got)
		}
	})

	t.Run("short circuits on first true in range", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
				{Col: 1, Row: 2}: BoolVal(false),
			},
		}
		cf := evalCompile(t, "OR(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("OR(A1:A2) = %v, want true", got)
		}
	})

	t.Run("range with numbers", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(0),
				{Col: 1, Row: 2}: NumberVal(0),
				{Col: 1, Row: 3}: NumberVal(5),
			},
		}
		cf := evalCompile(t, "OR(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("OR(A1:A3) = %v, want true", got)
		}
	})
}

func TestANDStringArgs(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("direct string args return VALUE error", func(t *testing.T) {
		tests := []string{
			`AND("text")`,
			`AND("1")`,
			`AND("TRUE")`,
		}
		for _, formula := range tests {
			t.Run(formula, func(t *testing.T) {
				cf := evalCompile(t, formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", formula, err)
				}
				if got.Type != ValueError || got.Err != ErrValVALUE {
					t.Errorf("Eval(%q) = %v, want #VALUE!", formula, got)
				}
			})
		}
	})

	t.Run("string in range is skipped", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("text"),
				{Col: 1, Row: 2}: BoolVal(true),
			},
		}
		cf := evalCompile(t, "AND(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("AND(A1:A2) = %v, want true", got)
		}
	})
}

func TestXOR(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("basic boolean results", func(t *testing.T) {
		tests := []struct {
			formula string
			want    bool
		}{
			// Single args
			{"XOR(TRUE)", true},
			{"XOR(FALSE)", false},
			// Two booleans
			{"XOR(TRUE,TRUE)", false},   // even TRUE count
			{"XOR(TRUE,FALSE)", true},   // odd TRUE count
			{"XOR(FALSE,FALSE)", false}, // zero TRUE count
			// Three booleans
			{"XOR(TRUE,TRUE,TRUE)", true},    // 3 TRUE = odd
			{"XOR(TRUE,TRUE,FALSE)", false},  // 2 TRUE = even
			{"XOR(TRUE,FALSE,FALSE)", true},  // 1 TRUE = odd
			{"XOR(FALSE,FALSE,FALSE)", false}, // 0 TRUE
			// Four TRUE (even count) → FALSE
			{"XOR(TRUE,TRUE,TRUE,TRUE)", false},
			// Numbers: non-zero = TRUE, zero = FALSE
			{"XOR(1)", true},
			{"XOR(0)", false},
			{"XOR(1,0)", true},    // 1 TRUE
			{"XOR(1,1)", false},   // 2 TRUE = even
			{"XOR(5,0,0)", true},  // 1 TRUE
			{"XOR(5,3,0)", false}, // 2 TRUE = even
			// Mixed booleans and numbers
			{"XOR(TRUE,1)", false},  // 2 TRUE = even
			{"XOR(TRUE,0)", true},   // 1 TRUE
			{"XOR(FALSE,1)", true},  // 1 TRUE
			{"XOR(FALSE,0)", false}, // 0 TRUE
			// Excel doc examples: =XOR(3>0,2<9) → FALSE (both TRUE, even)
			{"XOR(3>0,2<9)", false},
			// =XOR(3>12,4>6) → FALSE (both FALSE)
			{"XOR(3>12,4>6)", false},
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("no args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, "XOR()")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("XOR() = %v, want #VALUE!", got)
		}
	})

	t.Run("error propagation", func(t *testing.T) {
		cf := evalCompile(t, "XOR(1/0,TRUE)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("XOR(1/0,TRUE) = %v, want error", got)
		}
	})

	t.Run("range with mixed values", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
				{Col: 1, Row: 2}: BoolVal(false),
				{Col: 1, Row: 3}: NumberVal(1),
			},
		}
		// TRUE + FALSE + 1(TRUE) = 2 TRUE = even → FALSE
		cf := evalCompile(t, "XOR(A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf("XOR(A1:A3) = %v, want false", got)
		}
	})

	t.Run("error in range propagates", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
				{Col: 1, Row: 2}: ErrorVal(ErrValDIV0),
			},
		}
		cf := evalCompile(t, "XOR(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("XOR(A1:A2) = %v, want #DIV/0!", got)
		}
	})
}

func TestXORStringArgs(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("direct string args are skipped", func(t *testing.T) {
		// XOR skips strings (both direct and from ranges), unlike AND/OR which error.
		tests := []struct {
			formula string
			want    bool
		}{
			{`XOR("text",1)`, true},    // skip "text", XOR(1) = true
			{`XOR("text",TRUE)`, true}, // skip "text", XOR(TRUE) = true
			{`XOR("1",0)`, false},      // skip "1", XOR(0) = false
			{`XOR("0",0)`, false},      // skip "0", XOR(0) = false
		}
		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueBool || got.Bool != tt.want {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got, tt.want)
				}
			})
		}
	})

	t.Run("string in cell reference is skipped", func(t *testing.T) {
		// XOR(Data!A1, 1) where A1 contains "text" → string is skipped, XOR(1) = TRUE
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("text"),
				{Col: 1, Row: 2}: NumberVal(1),
			},
		}
		cf := evalCompile(t, "XOR(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("XOR(A1:A2) = %v, want true (string skipped, only 1 counts)", got)
		}
	})
}

func TestORStringInRange(t *testing.T) {
	t.Run("string in range is skipped", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("text"),
				{Col: 1, Row: 2}: BoolVal(true),
			},
		}
		cf := evalCompile(t, "OR(A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf("OR(A1:A2) = %v, want true", got)
		}
	})
}

func TestNOT(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    bool
	}{
		{"NOT(TRUE)", false},
		{"NOT(FALSE)", true},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueBool || got.Bool != tt.want {
			t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want)
		}
	}
}

func TestIF(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("true branch", func(t *testing.T) {
		cf := evalCompile(t, `IF(TRUE, "yes", "no")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IF(TRUE, "yes", "no") = %v, want "yes"`, got)
		}
	})

	t.Run("false branch", func(t *testing.T) {
		cf := evalCompile(t, `IF(FALSE, "yes", "no")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "no" {
			t.Errorf(`IF(FALSE, "yes", "no") = %v, want "no"`, got)
		}
	})

	t.Run("false without else returns FALSE", func(t *testing.T) {
		cf := evalCompile(t, `IF(FALSE, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf(`IF(FALSE, "yes") = %v, want FALSE`, got)
		}
	})
}

func TestIFERROR(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("non-error returns value", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(1, "fallback")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf(`IFERROR(1, "fallback") = %v, want 1`, got)
		}
	})

	t.Run("error returns fallback", func(t *testing.T) {
		cf := evalCompile(t, `IFERROR(1/0, "fallback")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "fallback" {
			t.Errorf(`IFERROR(1/0, "fallback") = %v, want "fallback"`, got)
		}
	})
}

func TestIFS(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("single condition true returns value", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IFS(TRUE, "yes") = %v, want "yes"`, got)
		}
	})

	t.Run("single condition false returns NA", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`IFS(FALSE, "yes") = %v, want #N/A`, got)
		}
	})

	t.Run("first true wins", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, "first", TRUE, "second")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "first" {
			t.Errorf(`IFS(TRUE,"first",TRUE,"second") = %v, want "first"`, got)
		}
	})

	t.Run("second condition true", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE, "first", TRUE, "second")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "second" {
			t.Errorf(`IFS(FALSE,"first",TRUE,"second") = %v, want "second"`, got)
		}
	})

	t.Run("last condition true wins", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE, "a", FALSE, "b", TRUE, "c")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "c" {
			t.Errorf(`IFS(FALSE,"a",FALSE,"b",TRUE,"c") = %v, want "c"`, got)
		}
	})

	t.Run("all conditions false returns NA", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE, 1, FALSE, 2, FALSE, 3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`IFS(FALSE,1,FALSE,2,FALSE,3) = %v, want #N/A`, got)
		}
	})

	t.Run("returns number value", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, 42)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 42 {
			t.Errorf(`IFS(TRUE, 42) = %v, want 42`, got)
		}
	})

	t.Run("returns boolean value", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, FALSE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != false {
			t.Errorf(`IFS(TRUE, FALSE) = %v, want FALSE`, got)
		}
	})

	t.Run("returns string value", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, "hello")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "hello" {
			t.Errorf(`IFS(TRUE, "hello") = %v, want "hello"`, got)
		}
	})

	t.Run("numeric condition non-zero is true", func(t *testing.T) {
		cf := evalCompile(t, `IFS(1, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IFS(1, "yes") = %v, want "yes"`, got)
		}
	})

	t.Run("numeric condition zero is false", func(t *testing.T) {
		cf := evalCompile(t, `IFS(0, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`IFS(0, "yes") = %v, want #N/A`, got)
		}
	})

	t.Run("negative number is truthy", func(t *testing.T) {
		cf := evalCompile(t, `IFS(-1, "neg")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "neg" {
			t.Errorf(`IFS(-1, "neg") = %v, want "neg"`, got)
		}
	})

	t.Run("fractional number is truthy", func(t *testing.T) {
		cf := evalCompile(t, `IFS(0.5, "frac")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "frac" {
			t.Errorf(`IFS(0.5, "frac") = %v, want "frac"`, got)
		}
	})

	t.Run("non-empty string condition is truthy", func(t *testing.T) {
		cf := evalCompile(t, `IFS("text", "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IFS("text", "yes") = %v, want "yes"`, got)
		}
	})

	t.Run("empty string condition is falsy", func(t *testing.T) {
		cf := evalCompile(t, `IFS("", "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`IFS("", "yes") = %v, want #N/A`, got)
		}
	})

	t.Run("odd number of args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`IFS(TRUE) = %v, want #VALUE!`, got)
		}
	})

	t.Run("three args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, 1, FALSE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`IFS(TRUE, 1, FALSE) = %v, want #VALUE!`, got)
		}
	})

	t.Run("error in condition propagates", func(t *testing.T) {
		cf := evalCompile(t, `IFS(1/0, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf(`IFS(1/0, "yes") = %v, want #DIV/0!`, got)
		}
	})

	t.Run("error in second condition propagates when first is false", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE, "a", 1/0, "b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf(`IFS(FALSE,"a",1/0,"b") = %v, want #DIV/0!`, got)
		}
	})

	t.Run("error in value not taken is not propagated", func(t *testing.T) {
		// When the first condition is TRUE, the second pair's value (1/0) is never reached
		cf := evalCompile(t, `IFS(TRUE, "ok", FALSE, 1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "ok" {
			t.Errorf(`IFS(TRUE,"ok",FALSE,1/0) = %v, want "ok"`, got)
		}
	})

	t.Run("error in taken value propagates", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, 1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf(`IFS(TRUE, 1/0) = %v, want #DIV/0!`, got)
		}
	})

	t.Run("expression conditions", func(t *testing.T) {
		cf := evalCompile(t, `IFS(1>2, "a", 2>1, "b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "b" {
			t.Errorf(`IFS(1>2,"a",2>1,"b") = %v, want "b"`, got)
		}
	})

	t.Run("nested IFS", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, IFS(FALSE, 1, TRUE, 2))`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf(`IFS(TRUE, IFS(FALSE,1,TRUE,2)) = %v, want 2`, got)
		}
	})

	t.Run("cell reference as condition true", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
			},
		}
		cf := evalCompile(t, `IFS(A1, "yes")`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IFS(A1, "yes") with A1=TRUE = %v, want "yes"`, got)
		}
	})

	t.Run("cell reference as condition false", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(false),
			},
		}
		cf := evalCompile(t, `IFS(A1, "yes")`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`IFS(A1, "yes") with A1=FALSE = %v, want #N/A`, got)
		}
	})

	t.Run("cell reference numeric condition", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(0),
				{Col: 2, Row: 1}: NumberVal(5),
			},
		}
		cf := evalCompile(t, `IFS(A1, "zero", B1, "five")`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "five" {
			t.Errorf(`IFS(A1,"zero",B1,"five") = %v, want "five"`, got)
		}
	})

	t.Run("cell reference with error", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNA),
			},
		}
		cf := evalCompile(t, `IFS(A1, "yes")`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`IFS(A1, "yes") with A1=#N/A = %v, want #N/A`, got)
		}
	})

	t.Run("many conditions middle match", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE,1,FALSE,2,TRUE,3,FALSE,4,FALSE,5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 3 {
			t.Errorf(`IFS with 5 pairs, 3rd true = %v, want 3`, got)
		}
	})

	t.Run("value is zero", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, 0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 0 {
			t.Errorf(`IFS(TRUE, 0) = %v, want 0`, got)
		}
	})

	t.Run("value is empty string", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, "")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "" {
			t.Errorf(`IFS(TRUE, "") = %v, want ""`, got)
		}
	})

	t.Run("computed value expression", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, 10+5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 15 {
			t.Errorf(`IFS(TRUE, 10+5) = %v, want 15`, got)
		}
	})
}

func TestSWITCH(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("match first value", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1, 1, "one", 2, "two", 3, "three")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "one" {
			t.Errorf(`SWITCH(1, 1,"one", 2,"two", 3,"three") = %v, want "one"`, got)
		}
	})

	t.Run("match second value", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(2, 1, "one", 2, "two", 3, "three")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "two" {
			t.Errorf(`SWITCH(2, 1,"one", 2,"two", 3,"three") = %v, want "two"`, got)
		}
	})

	t.Run("match last value", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(3, 1, "one", 2, "two", 3, "three")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "three" {
			t.Errorf(`SWITCH(3, 1,"one", 2,"two", 3,"three") = %v, want "three"`, got)
		}
	})

	t.Run("no match without default returns NA", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(99, 1, "one", 2, "two")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`SWITCH(99, 1,"one", 2,"two") = %v, want #N/A`, got)
		}
	})

	t.Run("no match with default returns default", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(99, 1, "one", 2, "two", "none")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "none" {
			t.Errorf(`SWITCH(99, 1,"one", 2,"two", "none") = %v, want "none"`, got)
		}
	})

	t.Run("default value is numeric", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(99, 1, "one", 2, "two", -1)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != -1 {
			t.Errorf(`SWITCH(99, 1,"one", 2,"two", -1) = %v, want -1`, got)
		}
	})

	t.Run("string matching", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH("b", "a", 1, "b", 2, "c", 3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 2 {
			t.Errorf(`SWITCH("b", "a",1, "b",2, "c",3) = %v, want 2`, got)
		}
	})

	t.Run("number matching", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(3.14, 1, "int", 3.14, "pi", 2.72, "e")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "pi" {
			t.Errorf(`SWITCH(3.14, ...) = %v, want "pi"`, got)
		}
	})

	t.Run("boolean matching TRUE", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(TRUE, FALSE, "no", TRUE, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`SWITCH(TRUE, FALSE,"no", TRUE,"yes") = %v, want "yes"`, got)
		}
	})

	t.Run("boolean matching FALSE", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(FALSE, TRUE, "yes", FALSE, "no")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "no" {
			t.Errorf(`SWITCH(FALSE, TRUE,"yes", FALSE,"no") = %v, want "no"`, got)
		}
	})

	t.Run("string matching is case insensitive", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH("HELLO", "hello", 1, "world", 2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf(`SWITCH("HELLO", "hello",1, "world",2) = %v, want 1`, got)
		}
	})

	t.Run("expression as first arg", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1+1, 1, "one", 2, "two", 3, "three")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "two" {
			t.Errorf(`SWITCH(1+1, 1,"one", 2,"two", 3,"three") = %v, want "two"`, got)
		}
	})

	t.Run("multiple pairs match returns first match", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1, 1, "first", 1, "second", 1, "third")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "first" {
			t.Errorf(`SWITCH(1, 1,"first", 1,"second", 1,"third") = %v, want "first"`, got)
		}
	})

	t.Run("error expression no match returns NA", func(t *testing.T) {
		// 1/0 evaluates to #DIV/0! which doesn't match 1 or 2, so returns #N/A
		cf := evalCompile(t, `SWITCH(1/0, 1, "one", 2, "two")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`SWITCH(1/0, 1,"one", 2,"two") = %v, want #N/A`, got)
		}
	})

	t.Run("error expression with default returns default", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1/0, 1, "one", 2, "two", "fallback")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "fallback" {
			t.Errorf(`SWITCH(1/0, ..., "fallback") = %v, want "fallback"`, got)
		}
	})

	t.Run("error in matched result propagates", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1, 1, 1/0, 2, "two")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf(`SWITCH(1, 1, 1/0, ...) = %v, want #DIV/0!`, got)
		}
	})

	t.Run("error in unmatched result does not propagate", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(2, 1, 1/0, 2, "two")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "two" {
			t.Errorf(`SWITCH(2, 1, 1/0, 2,"two") = %v, want "two"`, got)
		}
	})

	t.Run("mixed types no match", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH("hello", 1, "num", TRUE, "bool", "nomatch")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "nomatch" {
			t.Errorf(`SWITCH("hello", 1,"num", TRUE,"bool", "nomatch") = %v, want "nomatch"`, got)
		}
	})

	t.Run("too few args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1, 2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`SWITCH(1, 2) = %v, want #VALUE!`, got)
		}
	})

	t.Run("single pair match", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(5, 5, "five")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "five" {
			t.Errorf(`SWITCH(5, 5,"five") = %v, want "five"`, got)
		}
	})

	t.Run("single pair no match returns NA", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(5, 6, "six")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`SWITCH(5, 6,"six") = %v, want #N/A`, got)
		}
	})

	t.Run("zero matches zero", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(0, 1, "one", 0, "zero")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "zero" {
			t.Errorf(`SWITCH(0, 1,"one", 0,"zero") = %v, want "zero"`, got)
		}
	})

	t.Run("empty string matches empty string", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH("", "", "empty", "a", "letter")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "empty" {
			t.Errorf(`SWITCH("", "","empty", "a","letter") = %v, want "empty"`, got)
		}
	})

	t.Run("result can be numeric", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH("x", "a", 10, "x", 20, "z", 30)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 20 {
			t.Errorf(`SWITCH("x", "a",10, "x",20, "z",30) = %v, want 20`, got)
		}
	})

	t.Run("result can be boolean", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1, 1, TRUE, 2, FALSE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueBool || got.Bool != true {
			t.Errorf(`SWITCH(1, 1,TRUE, 2,FALSE) = %v, want TRUE`, got)
		}
	})

	t.Run("weekday example from docs", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(2, 1, "Sunday", 2, "Monday", 3, "Tuesday", "No match")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "Monday" {
			t.Errorf(`SWITCH(2, 1,"Sunday", 2,"Monday", ...) = %v, want "Monday"`, got)
		}
	})

	t.Run("weekday no match with default", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(99, 1, "Sunday", 2, "Monday", 3, "Tuesday", "No match")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "No match" {
			t.Errorf(`SWITCH(99, ..., "No match") = %v, want "No match"`, got)
		}
	})
}
