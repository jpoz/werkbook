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
		if got.Type != ValueString || got.Str != tt.want {
			t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
		}
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
		// Cases from NumberFormatTests.xlsx audit
		{`TEXT(12.344,"0.00")`, "12.34"},
		{`TEXT(12.344,"0.0")`, "12.3"},
		{`TEXT(12.3,"###.##")`, "12.3"},
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
		// Empty old_text — Excel returns original text unchanged
		{name: "empty_old", formula: `SUBSTITUTE("hello","","X")`, want: "hello"},
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
		{name: "percent_fp_rounding", formula: `TEXT(0.00035,"#,##0.00%")`, want: "0.04%"},
		// Question mark (space-padded digit) formatting
		{name: "qmark_pad_comma", formula: `TEXT(1234567,"?,?????????")`, want: "    1,234,567"},
		{name: "qmark_pad_simple", formula: `TEXT(5,"???")`, want: "  5"},
		{name: "qmark_pad_exact", formula: `TEXT(123,"???")`, want: "123"},
		// Question mark in decimal positions
		{name: "qmark_dec_11", formula: `TEXT(1.1,"?.?")`, want: "1.1"},
		{name: "qmark_dec_10", formula: `TEXT(1,"?.?")`, want: "1. "},
		{name: "qmark_dec_pipe", formula: `TEXT(1.1,"|?.?|")`, want: "|1.1|"},
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
// TEXT — audit failure categories
// ---------------------------------------------------------------------------

func TestTEXTDateTimeFormats(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		// Date formatting from Excel serial numbers
		{name: "mm-dd-yy", formula: `TEXT(17816.607951388887, "mm-dd-yy")`, want: "10-10-48"},
		{name: "yyyy-mm-dd", formula: `TEXT(44197, "yyyy-mm-dd")`, want: "2021-01-01"},
		{name: "mm/dd/yyyy", formula: `TEXT(44197, "mm/dd/yyyy")`, want: "01/01/2021"},
		{name: "d-mmm-yy", formula: `TEXT(44197, "d-mmm-yy")`, want: "1-Jan-21"},
		{name: "dd-mmm-yyyy", formula: `TEXT(44197, "dd-mmm-yyyy")`, want: "01-Jan-2021"},
		{name: "mmmm_d_yyyy", formula: `TEXT(44197, "mmmm d, yyyy")`, want: "January 1, 2021"},

		// Time formatting
		{name: "h:mm:ss", formula: `TEXT(0.5, "h:mm:ss")`, want: "12:00:00"},
		{name: "hh:mm:ss", formula: `TEXT(0.25, "hh:mm:ss")`, want: "06:00:00"},
		{name: "h:mm_AM/PM", formula: `TEXT(0.75, "h:mm AM/PM")`, want: "6:00 PM"},
		{name: "h:mm_AM/PM_morning", formula: `TEXT(0.25, "h:mm AM/PM")`, want: "6:00 AM"},

		// Combined date/time
		{name: "mm/dd/yyyy_hh:mm", formula: `TEXT(44197.5, "mm/dd/yyyy hh:mm")`, want: "01/01/2021 12:00"},

		// Single m/d (no leading zero)
		{name: "m/d/yy", formula: `TEXT(44197, "m/d/yy")`, want: "1/1/21"},
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

func TestTEXTDateTime1904(t *testing.T) {
	resolver := &mockResolver{}

	// In the 1904 date system, serial 0 = 1904-01-01.
	// Serial 17816.607951388887 in the 1904 system = Oct 10, 1952 (vs Oct 10, 1948 in 1900).
	ctx1904 := &EvalContext{Date1904: true}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "mm-dd-yy 1904", formula: `TEXT(17816.607951388887, "mm-dd-yy")`, want: "10-11-52"},
		{name: "yyyy-mm-dd 1904", formula: `TEXT(1, "yyyy-mm-dd")`, want: "1904-01-02"},
		{name: "yyyy-mm-dd serial 0", formula: `TEXT(0, "yyyy-mm-dd hh:mm:ss.000")`, want: "1904-01-01 00:00:00.000"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, ctx1904)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
			}
		})
	}
}

func TestTEXTElapsedTime(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "elapsed_hours", formula: `TEXT(3.14159, "[h]:mm:ss")`, want: "75:23:53"},
		{name: "elapsed_with_frac", formula: `TEXT(3.14159, "[h]:mm:ss.000")`, want: "75:23:53.376"},
		{name: "simple_elapsed", formula: `TEXT(1.5, "[h]:mm:ss")`, want: "36:00:00"},
		{name: "zero_elapsed", formula: `TEXT(0, "[h]:mm:ss")`, want: "0:00:00"},
		// [ss].000 — fractional seconds after elapsed bracket code
		{name: "elapsed_sec_frac", formula: `TEXT(3.14159, "[ss].000")`, want: "271433.376"},
		// bare ss with [s] present — bare ss should show total seconds
		{name: "elapsed_sec_bare_ss", formula: `TEXT(3.14159, "[s]"" [yes, ""ss""] seconds""")`, want: `271433 [yes, 271433] seconds`},
		// [s] with small value
		{name: "elapsed_sec_small", formula: `TEXT(0.08546296296296296, "[s]"" [yes, ""ss""] seconds""")`, want: `7384 [yes, 7384] seconds`},
		// s:m with [hh] — bare s/m show seconds/minutes within the hour, [hh] shows total hours
		{name: "elapsed_hh_with_bare_sm", formula: `TEXT(3.14159, "s:m"" @ hour ""[hh]")`, want: `53:23 @ hour 75`},
		// [h] with literal brackets in quoted text
		{name: "elapsed_h_with_quotes", formula: `TEXT(3.14159, """It was ""[h]"" [yes, ""h""] hours and ""mm:ss")`, want: `It was 75 [yes, 75] hours and 23:53`},
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

func TestTEXTCommaGrouping(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "millions", formula: `TEXT(1234567, "#,##0")`, want: "1,234,567"},
		{name: "with_decimals", formula: `TEXT(1234.56, "#,##0.00")`, want: "1,234.56"},
		{name: "negative", formula: `TEXT(-1234567, "#,##0")`, want: "-1,234,567"},
		{name: "small_number", formula: `TEXT(42, "#,##0")`, want: "42"},
		{name: "zero", formula: `TEXT(0, "#,##0.00")`, want: "0.00"},
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

func TestTEXTCurrency(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "dollar_prefix", formula: `TEXT(1234.5, "$#,##0.00")`, want: "$1,234.50"},
		{name: "dollar_simple", formula: `TEXT(12.3, "$#0.00")`, want: "$12.30"},
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

func TestTEXTPercentWithComma(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "percent_comma", formula: `TEXT(-1.2345, "#,##0.00%")`, want: "-123.45%"},
		{name: "large_percent", formula: `TEXT(123.45, "#,##0%")`, want: "12,345%"},
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

func TestTEXTScientific(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "sci_basic", formula: `TEXT(123456.789, "0.00E+00")`, want: "1.23E+05"},
		{name: "sci_small", formula: `TEXT(0.00123, "0.00E+00")`, want: "1.23E-03"},
		{name: "sci_negative", formula: `TEXT(-5678, "0.0E+0")`, want: "-5.7E+3"},
		// E- format: no sign for positive exponents.
		{name: "sci_eminus_pos", formula: `TEXT(123456.789, "0.0E-0")`, want: "1.2E5"},
		{name: "sci_eminus_neg", formula: `TEXT(0.00123, "0.0E-0")`, want: "1.2E-3"},
		// Pipe literals around coefficient and exponent.
		{name: "sci_pipe_eminus", formula: `TEXT(123456.789, "|#.#|E-|#|")`, want: "|1.2|E|5|"},
		{name: "sci_pipe_eplus", formula: `TEXT(123456.789, "|#.#|E+|#|")`, want: "|1.2|E|+5|"},
		{name: "sci_pipe_eplus_neg", formula: `TEXT(0.0000123456789, "|#.#|E+|#|")`, want: "|1.2|E|-5|"},
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

func TestTEXTLowercaseEReturnsVALUE(t *testing.T) {
	resolver := &mockResolver{}

	// Excel only recognises uppercase E for scientific notation.
	// Lowercase e+/e- in format strings → #VALUE!.
	formats := []string{
		`"|#,e-#|"`,
		`"|#e-#,|"`,
		`"|#%e-#|"`,
		`"|#,e+#|"`,
		`"|#.####|e+|#|"`,
		`"|#|e+|#|"`,
		`"|#.#|e+|#|"`,
	}
	for _, fmt := range formats {
		t.Run(fmt, func(t *testing.T) {
			formula := `TEXT(123456.789, ` + fmt + `)`
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
}

func TestTEXTFraction(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "simple_fraction", formula: `TEXT(0.5, "# #/#")`, want: "1/2"},
		{name: "mixed_fraction", formula: `TEXT(3.25, "# #/#")`, want: "3 1/4"},
		// Fraction with literal characters
		{name: "frac_literal_pipe", formula: `TEXT(0.75, "|#\:#/#|")`, want: "|3/4|"},
		{name: "frac_literal_whole", formula: `TEXT(-23.75, "|#\:#/#|")`, want: "-|23:3/4|"},
		{name: "frac_literal_zero", formula: `TEXT(0, "|#\:#/#|")`, want: "|0|"},
		{name: "frac_literal_eq", formula: `TEXT(0.75, "|#\:#=/=#|")`, want: "|3=/=4|"},
		{name: "frac_literal_eq_whole", formula: `TEXT(1, "|#\:#=/=#|")`, want: "|1|"},
		// ? padding in fractions (space-padded when fraction is zero)
		{name: "frac_qmark_zero", formula: `TEXT(1, "|#\:?=/=?|")`, want: "|1      |"},
		{name: "frac_qmark_zero_val0", formula: `TEXT(0, "|#\:? ?=/=?|")`, want: "|:0      |"},
		{name: "frac_underscore", formula: `TEXT(0.75, "|#_#/#|")`, want: "|3 /4|"},
		{name: "frac_underscore_neg", formula: `TEXT(-3.75, "|#_#/#|")`, want: "-|15 /4|"},
		{name: "frac_underscore_zero", formula: `TEXT(0, "|#_#/#|")`, want: "|0 /1|"},
		// Multi-digit whole part with interleaved literals
		{name: "frac_multi_whole", formula: `TEXT(23.75, "|#-#-#\:#/#|")`, want: "|-2-3:3/4|"},
		{name: "frac_multi_whole_neg", formula: `TEXT(-23.75, "|#-#-#\:#/#|")`, want: "-|-2-3:3/4|"},
		{name: "frac_multi_whole_zero", formula: `TEXT(0.75, "|#-#-#\:#/#|")`, want: "|--3/4|"},
		{name: "frac_multi_whole_neg_zero", formula: `TEXT(-0.75, "|#-#-#\:#/#|")`, want: "-|--3/4|"},
		// Zero-padded numerator and denominator
		{name: "frac_zero_pad", formula: `TEXT(23.75, "|#\:? ?0#/000")`, want: "|2:3  03/004"},
		{name: "frac_zero_pad_val1", formula: `TEXT(1, "|#\:? ?0#/000")`, want: "|:1  00/001"},
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

func TestTEXTSections(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "pos_neg_sections", formula: `TEXT(-42, "0;(0)")`, want: "(42)"},
		{name: "pos_section", formula: `TEXT(42, "0;(0)")`, want: "42"},
		{name: "three_sections_zero", formula: `TEXT(0, "0;(0);""zero""")`, want: "zero"},
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

func TestTEXTLiterals(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "quoted_literal", formula: `TEXT(12.3, """Value: ""0.00")`, want: "Value: 12.30"},
		{name: "general_format", formula: `TEXT(1, "General")`, want: "1"},
		{name: "general_float", formula: `TEXT(3.14, "General")`, want: "3.14"},
		{name: "general_true", formula: `TEXT(TRUE, "General")`, want: "TRUE"},
		{name: "general_false", formula: `TEXT(FALSE, "General")`, want: "FALSE"},
		{name: "general_large_1.1e11", formula: `TEXT(110000000000, "General")`, want: "1.1E+11"},
		{name: "general_large_1.12345e11", formula: `TEXT(112345000000, "General")`, want: "1.12345E+11"},
		{name: "general_large_1e11", formula: `TEXT(100000000000, "General")`, want: "1E+11"},
		{name: "general_small_1.123e-10", formula: `TEXT(1.123E-10, "General")`, want: "1.123E-10"},
		{name: "general_small_1e-5", formula: `TEXT(0.00001, "General")`, want: "1E-5"},
		{name: "general_not_sci_1e10", formula: `TEXT(10000000000, "General")`, want: "10000000000"},
		{name: "general_negative_large", formula: `TEXT(-110000000000, "General")`, want: "-1.1E+11"},
		{name: "bool_true_numeric_fmt", formula: `TEXT(TRUE, "0")`, want: "1"},
		{name: "bool_false_numeric_fmt", formula: `TEXT(FALSE, "0")`, want: "0"},
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

func TestTEXTZeroPad(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "zero_pad_6", formula: `TEXT(42, "000000")`, want: "000042"},
		{name: "zero_pad_3", formula: `TEXT(7, "000")`, want: "007"},
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
// CONCAT with range arguments
// ---------------------------------------------------------------------------

func TestCONCATRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("hello"),
			{Col: 1, Row: 2}: StringVal("world"),
		},
	}

	cf := evalCompile(t, `CONCAT(A1:A2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "helloworld" {
		t.Errorf("CONCAT(A1:A2): got %q, want 'helloworld'", got.Str)
	}

	// CONCAT with a range and a scalar
	cf2 := evalCompile(t, `CONCAT(A1:A2,"!")`)
	got2, err := Eval(cf2, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got2.Type != ValueString || got2.Str != "helloworld!" {
		t.Errorf("CONCAT(A1:A2,\"!\"): got %q, want 'helloworld!'", got2.Str)
	}

	// CONCAT with numbers in range
	resolver2 := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
		},
	}
	cf3 := evalCompile(t, `CONCAT(A1:A3)`)
	got3, err := Eval(cf3, resolver2, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got3.Type != ValueString || got3.Str != "123" {
		t.Errorf("CONCAT(A1:A3) numbers: got %q, want '123'", got3.Str)
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

// ---------------------------------------------------------------------------
// SEARCH with wildcards
// ---------------------------------------------------------------------------

func TestSEARCHWildcards(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		isErr   bool
	}{
		// Basic (no wildcard)
		{name: "basic", formula: `SEARCH("lo","hello")`, wantNum: 4},
		{name: "case_insensitive", formula: `SEARCH("LO","hello")`, wantNum: 4},
		// * wildcard — matches any sequence of characters
		{name: "star_any", formula: `SEARCH("*le","Apple")`, wantNum: 1},
		{name: "star_middle", formula: `SEARCH("A*e","Apple")`, wantNum: 1},
		{name: "star_empty_match", formula: `SEARCH("A*p","Apple")`, wantNum: 1},
		// ? wildcard — matches exactly one character
		{name: "question_mark", formula: `SEARCH("A?p","Apple")`, wantNum: 1},
		{name: "question_mid", formula: `SEARCH("?pp","Apple")`, wantNum: 1},
		{name: "question_no_match", formula: `SEARCH("A?e","Apple")`, isErr: true},
		// Combined wildcards
		{name: "star_and_question", formula: `SEARCH("A?p*","Apple pie")`, wantNum: 1},
		// Tilde escape: ~* matches literal *, ~? matches literal ?
		{name: "tilde_star", formula: `SEARCH("~*","a*b")`, wantNum: 2},
		{name: "tilde_question", formula: `SEARCH("~?","a?b")`, wantNum: 2},
		{name: "tilde_tilde", formula: `SEARCH("~~","a~b")`, wantNum: 2},
		// Start position
		{name: "start_pos", formula: `SEARCH("l","hello world",5)`, wantNum: 10},
		// Not found
		{name: "not_found", formula: `SEARCH("z","hello")`, isErr: true},
		// Empty search text matches position 1
		{name: "empty_find", formula: `SEARCH("","hello")`, wantNum: 1},
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
