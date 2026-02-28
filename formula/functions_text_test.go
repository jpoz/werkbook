package formula

import (
	"testing"
)

func TestNewTextFunctions(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		formula string
		want    string
	}{
		{`CHAR(65)`, "A"},
		{`CHAR(97)`, "a"},
		{`CLEAN("hello` + "\x01\x02" + `world")`, "helloworld"},
		{`EXACT("hello","hello")`, ""},
		{`PROPER("hello world")`, "Hello World"},
		{`PROPER("it's a test")`, "It'S A Test"},
		{`REPLACE("abcdef",3,2,"XY")`, "abXYef"},
		{`REPT("ab",3)`, "ababab"},
		{`SUBSTITUTE("hello world","world","earth")`, "hello earth"},
		{`CONCATENATE("hello"," ","world")`, "hello world"},
	}

	for _, tt := range strTests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if tt.formula == `EXACT("hello","hello")` {
			if got.Type != ValueBool || !got.Bool {
				t.Errorf("Eval(%q) = %v, want TRUE", tt.formula, got)
			}
			continue
		}
		if got.Type != ValueString || got.Str != tt.want {
			t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
		}
	}
}

func TestCODE(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `CODE("A")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 65 {
		t.Errorf("CODE(A) = %g, want 65", got.Num)
	}
}

func TestFIND(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `FIND("lo","hello")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("FIND: got %g, want 4", got.Num)
	}
}

func TestSEARCH(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `SEARCH("LO","Hello")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("SEARCH: got %g, want 4", got.Num)
	}
}

func TestCHOOSE(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `CHOOSE(2,"a","b","c")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "b" {
		t.Errorf("CHOOSE: got %v, want b", got)
	}
}

func TestVALUE(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `VALUE("123.45")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 123.45 {
		t.Errorf("VALUE: got %g, want 123.45", got.Num)
	}
}

func TestTEXTFormat(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    string
	}{
		{`TEXT(1234.5,"0.00")`, "1234.50"},
		{`TEXT(0.75,"0%")`, "75%"},
		{`TEXT(1234,"#,##0")`, "1,234"},
		{`TEXT(42,"0")`, "42"},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueString || got.Str != tt.want {
			t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
		}
	}
}

// ---------------------------------------------------------------------------
// SUBSTITUTE edge cases
// ---------------------------------------------------------------------------

func TestSUBSTITUTE(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		// Replace all occurrences (no instance_num)
		{name: "replace_all", formula: `SUBSTITUTE("aabbcc","b","X")`, want: "aaXXcc"},
		// Replace specific instance
		{name: "replace_2nd", formula: `SUBSTITUTE("abab","a","X",2)`, want: "abXb"},
		{name: "replace_1st", formula: `SUBSTITUTE("abab","a","X",1)`, want: "Xbab"},
		// No match — return original
		{name: "no_match", formula: `SUBSTITUTE("hello","z","X")`, want: "hello"},
		// Empty old_text — Go ReplaceAll inserts between every rune
		{name: "empty_old", formula: `SUBSTITUTE("hello","","X")`, want: "XhXeXlXlXoX"},
		// Empty new_text — delete occurrences
		{name: "delete_all", formula: `SUBSTITUTE("hello","l","")`, want: "heo"},
		// Empty source text
		{name: "empty_source", formula: `SUBSTITUTE("","a","X")`, want: ""},
		// Replace with longer string
		{name: "longer_replace", formula: `SUBSTITUTE("abc","b","XYZ")`, want: "aXYZc"},
		// Instance_num beyond count — no replacement
		{name: "instance_beyond", formula: `SUBSTITUTE("aaa","a","X",5)`, want: "aaa"},
		// Case-sensitive
		{name: "case_sensitive", formula: `SUBSTITUTE("Hello","h","X")`, want: "Hello"},
		{name: "case_match", formula: `SUBSTITUTE("Hello","H","X")`, want: "Xello"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
			}
		})
	}
}

func TestSUBSTITUTEInvalidArgs(t *testing.T) {
	resolver := &mockResolver{}

	// instance_num < 1 => #VALUE!
	cf := evalCompile(t, `SUBSTITUTE("abc","a","X",0)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("SUBSTITUTE instance 0: got %v, want #VALUE!", got)
	}
}

// ---------------------------------------------------------------------------
// formatNumber / TEXT — extended format codes
// ---------------------------------------------------------------------------

func TestTEXTFormatExtended(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		// Decimal formats
		{name: "one_decimal", formula: `TEXT(3.14159,"0.0")`, want: "3.1"},
		{name: "three_decimals", formula: `TEXT(3.14159,"0.000")`, want: "3.142"},
		{name: "zero_decimals", formula: `TEXT(3.7,"0")`, want: "4"},
		// Percent with decimals
		{name: "percent_2dec", formula: `TEXT(0.1234,"0.00%")`, want: "12.34%"},
		{name: "percent_nodec", formula: `TEXT(0.5,"0%")`, want: "50%"},
		// Comma formatting
		{name: "comma_thousands", formula: `TEXT(1234567,"#,##0")`, want: "1,234,567"},
		{name: "comma_with_dec", formula: `TEXT(1234.56,"#,##0.00")`, want: "1,234.56"},
		{name: "comma_negative", formula: `TEXT(-1234,"#,##0")`, want: "-1,234"},
		{name: "comma_zero", formula: `TEXT(0,"#,##0")`, want: "0"},
		// Integer format
		{name: "integer", formula: `TEXT(99.9,"0")`, want: "100"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// NUMBERVALUE
// ---------------------------------------------------------------------------

func TestNUMBERVALUE(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		isErr   bool
	}{
		{name: "european_format", formula: `NUMBERVALUE("2.500,27", ",", ".")`, wantNum: 2500.27},
		{name: "percent", formula: `NUMBERVALUE("3.5%")`, wantNum: 0.035},
		{name: "empty_string", formula: `NUMBERVALUE("")`, wantNum: 0},
		{name: "spaces", formula: `NUMBERVALUE("  3 000  ")`, wantNum: 3000},
		{name: "us_format", formula: `NUMBERVALUE("1,234.56")`, wantNum: 1234.56},
		{name: "european_full", formula: `NUMBERVALUE("1.234,56", ",", ".")`, wantNum: 1234.56},
		{name: "multiple_decimals", formula: `NUMBERVALUE("1.2.3")`, isErr: true},
		{name: "double_percent", formula: `NUMBERVALUE("50%%")`, wantNum: 0.005},
		{name: "integer", formula: `NUMBERVALUE("100")`, wantNum: 100},
	}

	const epsilon = 0.0001

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if tt.isErr {
				if got.Type != ValueError || got.Err != ErrValVALUE {
					t.Errorf("Eval(%q) = %v, want #VALUE!", tt.formula, got)
				}
			} else {
				if got.Type != ValueNumber {
					t.Errorf("Eval(%q) = %v, want number %g", tt.formula, got, tt.wantNum)
				} else {
					diff := got.Num - tt.wantNum
					if diff < -epsilon || diff > epsilon {
						t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
					}
				}
			}
		})
	}
}

// ---------------------------------------------------------------------------
// FIND edge cases
// ---------------------------------------------------------------------------

func TestFINDEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		isErr   bool
	}{
		{name: "basic", formula: `FIND("lo","hello")`, wantNum: 4},
		{name: "start_pos", formula: `FIND("l","hello world",5)`, wantNum: 10},
		{name: "not_found", formula: `FIND("z","hello")`, isErr: true},
		{name: "case_sensitive", formula: `FIND("H","hello")`, isErr: true},
		{name: "empty_find", formula: `FIND("","hello")`, wantNum: 1},
		{name: "start_too_large", formula: `FIND("h","hello",99)`, isErr: true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if tt.isErr {
				if got.Type != ValueError {
					t.Errorf("Eval(%q) = %v, want error", tt.formula, got)
				}
			} else {
				if got.Type != ValueNumber || got.Num != tt.wantNum {
					t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
				}
			}
		})
	}
}

// ---------------------------------------------------------------------------
// SEARCH edge cases (case-insensitive FIND)
// ---------------------------------------------------------------------------

func TestSEARCHEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		isErr   bool
	}{
		{name: "case_insensitive", formula: `SEARCH("LO","Hello")`, wantNum: 4},
		{name: "with_start", formula: `SEARCH("l","hello world",5)`, wantNum: 10},
		{name: "not_found", formula: `SEARCH("z","hello")`, isErr: true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if tt.isErr {
				if got.Type != ValueError {
					t.Errorf("Eval(%q) = %v, want error", tt.formula, got)
				}
			} else {
				if got.Type != ValueNumber || got.Num != tt.wantNum {
					t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
				}
			}
		})
	}
}

// ---------------------------------------------------------------------------
// LEFT/RIGHT/MID edge cases
// ---------------------------------------------------------------------------

func TestLEFTRIGHTMIDEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
		isErr   bool
	}{
		// LEFT
		{name: "left_default", formula: `LEFT("hello")`, want: "h"},
		{name: "left_zero", formula: `LEFT("hello",0)`, want: ""},
		{name: "left_exceeds", formula: `LEFT("hi",10)`, want: "hi"},
		{name: "left_negative", formula: `LEFT("hello",-1)`, isErr: true},
		// RIGHT
		{name: "right_default", formula: `RIGHT("hello")`, want: "o"},
		{name: "right_zero", formula: `RIGHT("hello",0)`, want: ""},
		{name: "right_exceeds", formula: `RIGHT("hi",10)`, want: "hi"},
		{name: "right_negative", formula: `RIGHT("hello",-1)`, isErr: true},
		// MID
		{name: "mid_basic", formula: `MID("hello",2,3)`, want: "ell"},
		{name: "mid_start_beyond", formula: `MID("hi",10,1)`, want: ""},
		{name: "mid_len_beyond", formula: `MID("hello",3,100)`, want: "llo"},
		{name: "mid_zero_len", formula: `MID("hello",1,0)`, want: ""},
		{name: "mid_neg_start", formula: `MID("hello",0,3)`, isErr: true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if tt.isErr {
				if got.Type != ValueError {
					t.Errorf("Eval(%q) = %v, want error", tt.formula, got)
				}
			} else {
				if got.Type != ValueString || got.Str != tt.want {
					t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
				}
			}
		})
	}
}

// ---------------------------------------------------------------------------
// REPLACE edge cases
// ---------------------------------------------------------------------------

func TestREPLACEEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "basic", formula: `REPLACE("abcdef",3,2,"XY")`, want: "abXYef"},
		{name: "insert", formula: `REPLACE("abcdef",3,0,"XY")`, want: "abXYcdef"},
		{name: "delete", formula: `REPLACE("abcdef",3,2,"")`, want: "abef"},
		{name: "replace_at_end", formula: `REPLACE("abc",3,1,"XYZ")`, want: "abXYZ"},
		{name: "start_beyond", formula: `REPLACE("abc",10,1,"X")`, want: "abcX"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// REPT edge cases
// ---------------------------------------------------------------------------

func TestREPTEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `REPT("ab",0)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "" {
		t.Errorf("REPT 0: got %q, want empty", got.Str)
	}

	cf = evalCompile(t, `REPT("x",-1)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("REPT negative: got %v, want #VALUE!", got)
	}
}

// ---------------------------------------------------------------------------
// PROPER edge cases
// ---------------------------------------------------------------------------

func TestPROPEREdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    string
	}{
		{`PROPER("hello world")`, "Hello World"},
		{`PROPER("HELLO WORLD")`, "Hello World"},
		{`PROPER("123abc")`, "123Abc"},
		{`PROPER("it's a test")`, "It'S A Test"},
		{`PROPER("")`, ""},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueString || got.Str != tt.want {
			t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
		}
	}
}

// ---------------------------------------------------------------------------
// CHAR/CODE edge cases
// ---------------------------------------------------------------------------

func TestCHAREdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	// Valid
	cf := evalCompile(t, `CHAR(65)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "A" {
		t.Errorf("CHAR(65) = %q, want A", got.Str)
	}

	// Out of range
	cf = evalCompile(t, `CHAR(0)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("CHAR(0) = %v, want #VALUE!", got)
	}

	cf = evalCompile(t, `CHAR(256)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("CHAR(256) = %v, want #VALUE!", got)
	}
}

func TestCODEEmpty(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `CODE("")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("CODE empty: got %v, want #VALUE!", got)
	}
}

// ---------------------------------------------------------------------------
// VALUE edge cases
// ---------------------------------------------------------------------------

func TestVALUEEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		isErr   bool
	}{
		{name: "basic", formula: `VALUE("123.45")`, wantNum: 123.45},
		{name: "with_commas", formula: `VALUE("1,234")`, wantNum: 1234},
		{name: "with_dollar", formula: `VALUE("$100")`, wantNum: 100},
		{name: "percent", formula: `VALUE("50%")`, wantNum: 0.5},
		{name: "whitespace", formula: `VALUE("  42  ")`, wantNum: 42},
		{name: "non_numeric", formula: `VALUE("abc")`, isErr: true},
		{name: "number_passthrough", formula: `VALUE(42)`, wantNum: 42},
		{name: "negative", formula: `VALUE("-99")`, wantNum: -99},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if tt.isErr {
				if got.Type != ValueError || got.Err != ErrValVALUE {
					t.Errorf("Eval(%q) = %v, want #VALUE!", tt.formula, got)
				}
			} else {
				if got.Type != ValueNumber || got.Num != tt.wantNum {
					t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.wantNum)
				}
			}
		})
	}
}

// ---------------------------------------------------------------------------
// T function
// ---------------------------------------------------------------------------

func TestT(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "string_value", formula: `T("hello")`, want: "hello"},
		{name: "number_value", formula: `T(123)`, want: ""},
		{name: "bool_value", formula: `T(TRUE)`, want: ""},
		{name: "empty_string", formula: `T("")`, want: ""},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %v (str=%q), want %q", tt.formula, got, got.Str, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// TEXTJOIN
// ---------------------------------------------------------------------------

func TestTEXTJOIN(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "basic", formula: `TEXTJOIN(", ", TRUE, "a", "b", "c")`, want: "a, b, c"},
		{name: "ignore_empty", formula: `TEXTJOIN(", ", TRUE, "a", "", "c")`, want: "a, c"},
		{name: "include_empty", formula: `TEXTJOIN(", ", FALSE, "a", "", "c")`, want: "a, , c"},
		{name: "empty_delimiter", formula: `TEXTJOIN("", TRUE, "a", "b")`, want: "ab"},
		{name: "single_value", formula: `TEXTJOIN("-", TRUE, "only")`, want: "only"},
		{name: "all_empty_ignored", formula: `TEXTJOIN(", ", TRUE, "", "", "")`, want: ""},
		{name: "numbers", formula: `TEXTJOIN("-", TRUE, 1, 2, 3)`, want: "1-2-3"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
			}
		})
	}
}

func TestTEXTJOINTooFewArgs(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `TEXTJOIN(", ", TRUE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("TEXTJOIN too few args: got %v, want #VALUE!", got)
	}
}

// ---------------------------------------------------------------------------
// CLEAN edge cases
// ---------------------------------------------------------------------------

func TestCLEANExtended(t *testing.T) {
	resolver := &mockResolver{}

	// Normal text stays unchanged
	cf := evalCompile(t, `CLEAN("hello")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "hello" {
		t.Errorf("CLEAN normal: got %q, want hello", got.Str)
	}
}

// ---------------------------------------------------------------------------
// TRIM edge cases
// ---------------------------------------------------------------------------

func TestTRIMEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    string
	}{
		{`TRIM("  hello  ")`, "hello"},
		{`TRIM("  hello   world  ")`, "hello world"},
		{`TRIM("")`, ""},
		{`TRIM("   ")`, ""},
		{`TRIM("hello")`, "hello"},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueString || got.Str != tt.want {
			t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
		}
	}
}

// ---------------------------------------------------------------------------
// EXACT edge cases
// ---------------------------------------------------------------------------

func TestEXACTEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    bool
	}{
		{`EXACT("hello","hello")`, true},
		{`EXACT("Hello","hello")`, false}, // case-sensitive!
		{`EXACT("","")`, true},
		{`EXACT("a","b")`, false},
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

// ---------------------------------------------------------------------------
// CHOOSE edge cases
// ---------------------------------------------------------------------------

func TestCHOOSEEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	// Out of range
	cf := evalCompile(t, `CHOOSE(5,"a","b","c")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("CHOOSE OOB: got %v, want #VALUE!", got)
	}

	// Index 0
	cf = evalCompile(t, `CHOOSE(0,"a","b")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("CHOOSE 0: got %v, want #VALUE!", got)
	}
}

// ---------------------------------------------------------------------------
// CONCATENATE / CONCAT with multiple types
// ---------------------------------------------------------------------------

func TestCONCATENATETypes(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `CONCATENATE("Value: ",42,", OK: ",TRUE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "Value: 42, OK: TRUE" {
		t.Errorf("CONCATENATE types: got %q, want 'Value: 42, OK: TRUE'", got.Str)
	}
}

// ---------------------------------------------------------------------------
// LEN with Unicode
// ---------------------------------------------------------------------------

func TestLENUnicode(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `LEN("caf`+"\u00e9"+`")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// "caf\u00e9" is 4 runes
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("LEN unicode: got %g, want 4", got.Num)
	}
}

// ---------------------------------------------------------------------------
// FIXED
// ---------------------------------------------------------------------------

func TestFIXED(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "with_1_decimal", formula: `FIXED(1234.567, 1)`, want: "1,234.6"},
		{name: "negative_decimals", formula: `FIXED(1234.567, -1)`, want: "1,230"},
		{name: "negative_num_neg_dec_no_commas", formula: `FIXED(-1234.567, -1, TRUE)`, want: "-1230"},
		{name: "default_decimals", formula: `FIXED(44.332)`, want: "44.33"},
		{name: "large_number_with_commas", formula: `FIXED(1234567.89, 2)`, want: "1,234,567.89"},
		{name: "large_number_no_commas", formula: `FIXED(1234567.89, 2, TRUE)`, want: "1234567.89"},
		{name: "zero", formula: `FIXED(0, 2)`, want: "0.00"},
		{name: "round_half_up", formula: `FIXED(1.5, 0)`, want: "2"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
			}
		})
	}
}
