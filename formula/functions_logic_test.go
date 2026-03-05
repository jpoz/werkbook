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
