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
		{name: "mmmmm_first_letter", formula: `TEXT(45720, "MMMMM")`, want: "M"},
		{name: "mmmmm_jan", formula: `TEXT(44197, "MMMMM")`, want: "J"},

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
		// Denominator left-justified (right-padded) with ?? placeholders
		{name: "frac_denom_left_align", formula: `TEXT(1/3, "# ??/??")`, want: "  1/3 "},
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
		{name: "bool_true_numeric_fmt", formula: `TEXT(TRUE, "0")`, want: "TRUE"},
		{name: "bool_false_numeric_fmt", formula: `TEXT(FALSE, "0")`, want: "FALSE"},
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
// LEFT comprehensive tests
// ---------------------------------------------------------------------------

func TestLEFT(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
		isErr   bool
	}{
		// Basic usage
		{name: "basic_2_chars", formula: `LEFT("hello",2)`, want: "he"},
		{name: "basic_3_chars", formula: `LEFT("hello",3)`, want: "hel"},
		{name: "basic_5_chars", formula: `LEFT("hello",5)`, want: "hello"},
		{name: "basic_1_char", formula: `LEFT("hello",1)`, want: "h"},

		// Default num_chars (1 when omitted)
		{name: "default_num_chars", formula: `LEFT("hello")`, want: "h"},
		{name: "default_single_char", formula: `LEFT("A")`, want: "A"},
		{name: "default_space", formula: `LEFT(" hello")`, want: " "},

		// num_chars = 0 returns empty string
		{name: "zero_chars", formula: `LEFT("hello",0)`, want: ""},
		{name: "zero_chars_empty", formula: `LEFT("",0)`, want: ""},

		// num_chars greater than string length returns full string
		{name: "exceeds_length", formula: `LEFT("hi",10)`, want: "hi"},
		{name: "exceeds_length_single", formula: `LEFT("A",100)`, want: "A"},

		// Empty string input
		{name: "empty_string", formula: `LEFT("")`, want: ""},
		{name: "empty_string_with_n", formula: `LEFT("",5)`, want: ""},

		// Numeric first argument coerced to string
		{name: "numeric_input", formula: `LEFT(12345,3)`, want: "123"},
		{name: "numeric_input_default", formula: `LEFT(12345)`, want: "1"},
		{name: "numeric_float", formula: `LEFT(3.14,3)`, want: "3.1"},
		{name: "numeric_zero", formula: `LEFT(0)`, want: "0"},

		// Boolean first argument coerced to string
		{name: "bool_true", formula: `LEFT(TRUE,2)`, want: "TR"},
		{name: "bool_false", formula: `LEFT(FALSE,3)`, want: "FAL"},
		{name: "bool_true_default", formula: `LEFT(TRUE)`, want: "T"},
		{name: "bool_false_full", formula: `LEFT(FALSE,5)`, want: "FALSE"},

		// Negative num_chars should error
		{name: "negative_num_chars", formula: `LEFT("hello",-1)`, isErr: true},
		{name: "negative_num_chars_large", formula: `LEFT("hello",-100)`, isErr: true},

		// Non-numeric num_chars should error
		{name: "non_numeric_num_chars", formula: `LEFT("hello","abc")`, isErr: true},

		// num_chars as float (truncated to int)
		{name: "float_num_chars", formula: `LEFT("hello",2.9)`, want: "he"},
		{name: "float_num_chars_1_5", formula: `LEFT("hello",1.5)`, want: "h"},

		// Special characters and spaces
		{name: "spaces", formula: `LEFT("  hello  ",4)`, want: "  he"},
		{name: "leading_space", formula: `LEFT(" a",2)`, want: " a"},
		{name: "backslash_t", formula: "LEFT(\"\\thello\",2)", want: "\\t"}, // formula parser treats \t as literal characters
		{name: "punctuation", formula: `LEFT("!@#$%",3)`, want: "!@#"},
		{name: "digits_in_string", formula: `LEFT("123abc",4)`, want: "123a"},
		{name: "mixed_case", formula: `LEFT("AbCdEf",3)`, want: "AbC"},

		// Unicode / multibyte characters
		{name: "unicode_chars", formula: "LEFT(\"日本語\",2)", want: "日本"},
		{name: "unicode_single", formula: "LEFT(\"日本語\",1)", want: "日"},
		{name: "unicode_exceeds", formula: "LEFT(\"日本語\",10)", want: "日本語"},

		// Wrong argument count
		{name: "no_args", formula: `LEFT()`, isErr: true},
		{name: "three_args", formula: `LEFT("hello",2,3)`, isErr: true},
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
					t.Errorf("Eval(%q) = %q (type=%d), want %q", tt.formula, got.Str, got.Type, tt.want)
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
// MID comprehensive tests
// ---------------------------------------------------------------------------

func TestMIDComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
		isErr   bool
	}{
		// Basic usage
		{name: "basic", formula: `MID("hello",2,3)`, want: "ell"},
		{name: "basic_single_char", formula: `MID("hello",3,1)`, want: "l"},
		{name: "basic_two_chars", formula: `MID("hello",4,2)`, want: "lo"},
		{name: "basic_full_string", formula: `MID("hello",1,5)`, want: "hello"},

		// start_num = 1 (beginning of string)
		{name: "start_at_one", formula: `MID("abcdef",1,3)`, want: "abc"},
		{name: "start_at_one_single", formula: `MID("abcdef",1,1)`, want: "a"},
		{name: "start_at_one_all", formula: `MID("abcdef",1,6)`, want: "abcdef"},

		// num_chars = 0 returns empty string
		{name: "zero_num_chars", formula: `MID("hello",1,0)`, want: ""},
		{name: "zero_num_chars_mid", formula: `MID("hello",3,0)`, want: ""},
		{name: "zero_num_chars_end", formula: `MID("hello",5,0)`, want: ""},

		// num_chars exceeds remaining length returns rest of string
		{name: "exceeds_remaining", formula: `MID("hello",3,100)`, want: "llo"},
		{name: "exceeds_from_start", formula: `MID("hi",1,10)`, want: "hi"},
		{name: "exceeds_from_last", formula: `MID("hello",5,50)`, want: "o"},
		{name: "exceeds_by_one", formula: `MID("abc",2,3)`, want: "bc"},

		// start_num exceeds string length returns empty string
		{name: "start_beyond_length", formula: `MID("hi",10,1)`, want: ""},
		{name: "start_just_beyond", formula: `MID("abc",4,1)`, want: ""},
		{name: "start_way_beyond", formula: `MID("x",100,5)`, want: ""},

		// Empty string input
		{name: "empty_string", formula: `MID("",1,0)`, want: ""},
		{name: "empty_string_with_chars", formula: `MID("",1,5)`, want: ""},

		// Numeric input coerced to string
		{name: "numeric_input", formula: `MID(12345,2,3)`, want: "234"},
		{name: "numeric_single_digit", formula: `MID(12345,1,1)`, want: "1"},
		{name: "numeric_zero", formula: `MID(0,1,1)`, want: "0"},
		{name: "numeric_negative", formula: `MID(-123,1,2)`, want: "-1"},
		{name: "numeric_decimal", formula: `MID(3.14,2,2)`, want: ".1"},

		// Boolean input coerced to string
		{name: "bool_true", formula: `MID(TRUE,1,2)`, want: "TR"},
		{name: "bool_false", formula: `MID(FALSE,2,3)`, want: "ALS"},
		{name: "bool_true_end", formula: `MID(TRUE,3,2)`, want: "UE"},
		{name: "bool_false_full", formula: `MID(FALSE,1,5)`, want: "FALSE"},

		// start_num <= 0 (should error)
		{name: "start_zero", formula: `MID("hello",0,3)`, isErr: true},
		{name: "start_negative", formula: `MID("hello",-1,3)`, isErr: true},
		{name: "start_negative_large", formula: `MID("hello",-100,3)`, isErr: true},

		// Negative num_chars (should error)
		{name: "negative_num_chars", formula: `MID("hello",1,-1)`, isErr: true},
		{name: "negative_num_chars_large", formula: `MID("hello",1,-100)`, isErr: true},

		// Non-numeric args (should error)
		{name: "non_numeric_start", formula: `MID("hello","abc",3)`, isErr: true},
		{name: "non_numeric_num_chars", formula: `MID("hello",1,"abc")`, isErr: true},

		// Special characters and spaces
		{name: "with_spaces", formula: `MID("hello world",6,1)`, want: " "},
		{name: "extract_word", formula: `MID("hello world",7,5)`, want: "world"},
		{name: "special_chars", formula: `MID("abc!@#def",4,3)`, want: "!@#"},
		{name: "newline_char", formula: "MID(\"abc\ndef\",4,1)", want: "\n"},

		// Float args truncated to int
		{name: "float_start", formula: `MID("hello",2.9,3)`, want: "ell"},
		{name: "float_num_chars", formula: `MID("hello",1,2.9)`, want: "he"},
		{name: "float_both", formula: `MID("hello",1.7,3.8)`, want: "hel"},
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

func TestMIDWrongArgCount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, `MID("hello",2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("MID with 2 args: got %v, want error", got)
	}

	// Too many args
	cf = evalCompile(t, `MID("hello",2,3,4)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("MID with 4 args: got %v, want error", got)
	}
}

// ---------------------------------------------------------------------------
// RIGHT comprehensive tests
// ---------------------------------------------------------------------------

func TestRIGHTComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
		isErr   bool
	}{
		// Basic usage
		{name: "basic_two_chars", formula: `RIGHT("hello",2)`, want: "lo"},
		{name: "basic_three_chars", formula: `RIGHT("hello",3)`, want: "llo"},
		{name: "basic_one_char", formula: `RIGHT("hello",1)`, want: "o"},
		{name: "basic_full_length", formula: `RIGHT("hello",5)`, want: "hello"},

		// Default num_chars (1 when omitted)
		{name: "default_num_chars", formula: `RIGHT("hello")`, want: "o"},
		{name: "default_single_char_string", formula: `RIGHT("x")`, want: "x"},
		{name: "default_longer_string", formula: `RIGHT("abcdef")`, want: "f"},

		// num_chars = 0 returns empty string
		{name: "zero_num_chars", formula: `RIGHT("hello",0)`, want: ""},
		{name: "zero_num_chars_empty", formula: `RIGHT("",0)`, want: ""},

		// num_chars greater than string length returns full string
		{name: "exceeds_length", formula: `RIGHT("hi",10)`, want: "hi"},
		{name: "exceeds_by_one", formula: `RIGHT("abc",4)`, want: "abc"},
		{name: "exceeds_single_char", formula: `RIGHT("x",100)`, want: "x"},

		// Empty string input
		{name: "empty_string", formula: `RIGHT("")`, want: ""},
		{name: "empty_string_with_n", formula: `RIGHT("",5)`, want: ""},
		{name: "empty_string_zero", formula: `RIGHT("",0)`, want: ""},

		// Numeric input coerced to string
		{name: "numeric_input", formula: `RIGHT(12345,3)`, want: "345"},
		{name: "numeric_input_default", formula: `RIGHT(12345)`, want: "5"},
		{name: "numeric_zero", formula: `RIGHT(0)`, want: "0"},
		{name: "numeric_negative", formula: `RIGHT(-123,2)`, want: "23"},
		{name: "numeric_decimal", formula: `RIGHT(3.14,2)`, want: "14"},

		// Boolean input coerced to string
		{name: "bool_true", formula: `RIGHT(TRUE,2)`, want: "UE"},
		{name: "bool_false", formula: `RIGHT(FALSE,3)`, want: "LSE"},
		{name: "bool_true_default", formula: `RIGHT(TRUE)`, want: "E"},
		{name: "bool_false_default", formula: `RIGHT(FALSE)`, want: "E"},

		// Negative num_chars (should error)
		{name: "negative_num_chars", formula: `RIGHT("hello",-1)`, isErr: true},
		{name: "negative_num_chars_large", formula: `RIGHT("hello",-100)`, isErr: true},

		// Non-numeric num_chars (should error)
		{name: "non_numeric_num_chars", formula: `RIGHT("hello","abc")`, isErr: true},

		// Special characters and spaces
		{name: "with_spaces", formula: `RIGHT("hello world",5)`, want: "world"},
		{name: "trailing_space", formula: `RIGHT("hello ",1)`, want: " "},
		{name: "special_chars", formula: `RIGHT("abc!@#",3)`, want: "!@#"},
		{name: "newline_char", formula: "RIGHT(\"abc\ndef\",3)", want: "def"},

		// Float num_chars truncated to int
		{name: "float_num_chars", formula: `RIGHT("hello",2.9)`, want: "lo"},
		{name: "float_num_chars_one", formula: `RIGHT("hello",1.5)`, want: "o"},
		{name: "float_num_chars_zero", formula: `RIGHT("hello",0.9)`, want: ""},
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

func TestRIGHTWrongArgCount(t *testing.T) {
	resolver := &mockResolver{}

	// 0 args
	cf := evalCompile(t, `RIGHT()`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(RIGHT()): %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("RIGHT() = %v, want error", got)
	}

	// 3 args
	cf = evalCompile(t, `RIGHT("a","b","c")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(RIGHT(a,b,c)): %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("RIGHT(a,b,c) = %v, want error", got)
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
// DOLLAR comprehensive tests
// ---------------------------------------------------------------------------

func TestDOLLAR(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		name    string
		formula string
		want    string
	}{
		// Basic positive numbers
		{name: "basic", formula: `DOLLAR(1234.567, 2)`, want: "$1,234.57"},
		{name: "default_decimals", formula: `DOLLAR(99)`, want: "$99.00"},
		{name: "zero_decimals", formula: `DOLLAR(1234.567, 0)`, want: "$1,235"},
		{name: "negative_decimals", formula: `DOLLAR(1234.567, -2)`, want: "$1,200"},

		// Negative numbers
		{name: "negative", formula: `DOLLAR(-1234.567, 2)`, want: "($1,234.57)"},
		{name: "negative_zero_dec", formula: `DOLLAR(-1234.567, 0)`, want: "($1,235)"},
		{name: "negative_neg_dec", formula: `DOLLAR(-1234.567, -2)`, want: "($1,200)"},

		// Zero
		{name: "zero", formula: `DOLLAR(0, 2)`, want: "$0.00"},
		{name: "zero_default", formula: `DOLLAR(0)`, want: "$0.00"},

		// Small values
		{name: "small_positive", formula: `DOLLAR(0.5, 2)`, want: "$0.50"},
		{name: "small_negative", formula: `DOLLAR(-0.5, 2)`, want: "($0.50)"},

		// String coercion
		{name: "string_number", formula: `DOLLAR("1234", 2)`, want: "$1,234.00"},

		// Boolean coercion
		{name: "bool_true", formula: `DOLLAR(TRUE, 2)`, want: "$1.00"},
		{name: "bool_false", formula: `DOLLAR(FALSE, 2)`, want: "$0.00"},

		// Large numbers
		{name: "large_number", formula: `DOLLAR(1234567.89, 2)`, want: "$1,234,567.89"},
		{name: "millions", formula: `DOLLAR(1000000, 0)`, want: "$1,000,000"},

		// Many decimal places
		{name: "many_decimals", formula: `DOLLAR(1.5, 5)`, want: "$1.50000"},

		// Negative zero edge case: -0.001 with 2 decimals rounds to 0.00
		{name: "neg_zero_round", formula: `DOLLAR(-0.001, 2)`, want: "$0.00"},

		// No decimal part
		{name: "integer_input", formula: `DOLLAR(42, 0)`, want: "$42"},

		// Small number no comma
		{name: "small_no_comma", formula: `DOLLAR(5, 2)`, want: "$5.00"},

		// Negative decimals rounding
		{name: "neg_dec_round_up", formula: `DOLLAR(1250, -2)`, want: "$1,300"},
		{name: "neg_dec_thousands", formula: `DOLLAR(12345, -3)`, want: "$12,000"},
	}

	for _, tt := range strTests {
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

	// Error cases
	errTests := []struct {
		name    string
		formula string
	}{
		{name: "no_args", formula: `DOLLAR()`},
		{name: "too_many_args", formula: `DOLLAR(1,2,3)`},
		{name: "non_numeric_string", formula: `DOLLAR("abc")`},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError {
				t.Errorf("Eval(%q) = %v, want error", tt.formula, got)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// LEN comprehensive tests
// ---------------------------------------------------------------------------

func TestLEN(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		isErr   bool
	}{
		// Basic usage
		{name: "basic_hello", formula: `LEN("hello")`, wantNum: 5},
		{name: "basic_word", formula: `LEN("world")`, wantNum: 5},
		{name: "basic_sentence", formula: `LEN("hello world")`, wantNum: 11},

		// Empty string
		{name: "empty_string", formula: `LEN("")`, wantNum: 0},

		// Single character
		{name: "single_char", formula: `LEN("A")`, wantNum: 1},
		{name: "single_space", formula: `LEN(" ")`, wantNum: 1},

		// Strings with spaces
		{name: "leading_trailing_spaces", formula: `LEN("  hello  ")`, wantNum: 9},
		{name: "only_spaces", formula: `LEN("   ")`, wantNum: 3},
		{name: "internal_spaces", formula: `LEN("a b c")`, wantNum: 5},

		// Numeric input coerced to string
		{name: "integer", formula: `LEN(123)`, wantNum: 3},
		{name: "negative_integer", formula: `LEN(-42)`, wantNum: 3},
		{name: "decimal", formula: `LEN(1.5)`, wantNum: 3},
		{name: "zero", formula: `LEN(0)`, wantNum: 1},
		{name: "large_number", formula: `LEN(123456789)`, wantNum: 9},

		// Boolean input
		{name: "bool_true", formula: `LEN(TRUE)`, wantNum: 4},
		{name: "bool_false", formula: `LEN(FALSE)`, wantNum: 5},

		// Special characters
		{name: "tab_char", formula: "LEN(\"\t\")", wantNum: 1},
		{name: "newline_in_string", formula: "LEN(\"a\nb\")", wantNum: 3},
		{name: "punctuation", formula: `LEN("!@#$%")`, wantNum: 5},

		// Unicode characters
		{name: "unicode_accented", formula: `LEN("caf` + "\u00e9" + `")`, wantNum: 4},
		{name: "unicode_emoji", formula: `LEN("` + "\U0001F600" + `")`, wantNum: 1},
		{name: "unicode_chinese", formula: `LEN("` + "\u4F60\u597D" + `")`, wantNum: 2},
		{name: "unicode_mixed", formula: `LEN("a` + "\u00e9\u4F60" + `b")`, wantNum: 4},

		// Long string
		{name: "long_string", formula: `LEN("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")`, wantNum: 52},

		// Wrong argument count
		{name: "no_args", formula: `LEN()`, isErr: true},
		{name: "two_args", formula: `LEN("a","b")`, isErr: true},

		// Nested formula
		{name: "nested_concat", formula: `LEN(CONCATENATE("ab","cd"))`, wantNum: 4},
		{name: "nested_upper", formula: `LEN(UPPER("hello"))`, wantNum: 5},
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

// ---------------------------------------------------------------------------
// TEXT — boolean, @, color code, and digit placeholder fixes
// ---------------------------------------------------------------------------

func TestTEXTBooleanAlwaysText(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "true_zero_fmt", formula: `TEXT(TRUE,"0")`, want: "TRUE"},
		{name: "false_zero_fmt", formula: `TEXT(FALSE,"0")`, want: "FALSE"},
		{name: "true_general", formula: `TEXT(TRUE,"General")`, want: "TRUE"},
		{name: "false_general", formula: `TEXT(FALSE,"General")`, want: "FALSE"},
		{name: "true_decimal", formula: `TEXT(TRUE,"0.00")`, want: "TRUE"},
		{name: "false_decimal", formula: `TEXT(FALSE,"0.00")`, want: "FALSE"},
		{name: "true_percent", formula: `TEXT(TRUE,"0%")`, want: "TRUE"},
		{name: "false_percent", formula: `TEXT(FALSE,"0%")`, want: "FALSE"},
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

func TestTEXTAtFormat(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "at_format", formula: `TEXT("hello","@")`, want: "hello"},
		{name: "at_with_prefix", formula: `TEXT("world","@ rocks")`, want: "world rocks"},
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

func TestTEXTColorCodes(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "red_positive", formula: `TEXT(5,"[Red]0.00")`, want: "5.00"},
		{name: "red_negative", formula: `TEXT(-5,"[Red]0.00")`, want: "-5.00"},
		{name: "blue_integer", formula: `TEXT(42,"[Blue]0")`, want: "42"},
		{name: "green_percent", formula: `TEXT(0.5,"[Green]0%")`, want: "50%"},
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

func TestTEXTNoFormatCodesReturnsVALUE(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
	}{
		{name: "pos_neg_zero_pos", formula: `TEXT(42,"pos;neg;zero")`},
		{name: "pos_neg_zero_neg", formula: `TEXT(-42,"pos;neg;zero")`},
		{name: "pos_neg_zero_zero", formula: `TEXT(0,"pos;neg;zero")`},
		// Unquoted alphabetic text mixed with number format codes is invalid.
		{name: "unquoted_text_with_digit", formula: `TEXT(42,"Value: 0")`},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != ErrValVALUE {
				t.Errorf("Eval(%q) = %v, want #VALUE!", tt.formula, got)
			}
		})
	}
}

func TestTEXTDigitPlaceholders(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{name: "phone_format", formula: `TEXT(5551234567,"(###) ###-####")`, want: "(555) 123-4567"},
		{name: "ssn_format", formula: `TEXT(123456789,"000-00-0000")`, want: "123-45-6789"},
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
// T comprehensive tests
// ---------------------------------------------------------------------------

func TestT(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("string_returns_string", func(t *testing.T) {
		cf := evalCompile(t, `T("hello")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "hello" {
			t.Errorf("T(\"hello\") = %v, want string(hello)", got)
		}
	})

	t.Run("number_returns_empty", func(t *testing.T) {
		cf := evalCompile(t, `T(123)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "" {
			t.Errorf("T(123) = %v, want empty string", got)
		}
	})

	t.Run("bool_returns_empty", func(t *testing.T) {
		cf := evalCompile(t, `T(TRUE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "" {
			t.Errorf("T(TRUE) = %v, want empty string", got)
		}
	})

	t.Run("error_propagates_div0", func(t *testing.T) {
		cf := evalCompile(t, `T(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("T(1/0) = %v, want error", got)
		}
	})

	t.Run("error_propagates_na", func(t *testing.T) {
		cf := evalCompile(t, `T(NA())`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("T(NA()) = %v, want error", got)
		}
	})
}

func TestCODE_Windows1252(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		// ASCII characters — same in Unicode and Windows-1252
		{"ASCII A", `CODE("A")`, 65},
		{"ASCII underscore", `CODE("_")`, 95},
		{"ASCII space", `CODE(" ")`, 32},

		// Windows-1252 range 0x80–0x9F — these Unicode code points
		// must map back to their Windows-1252 byte values.
		{"Euro sign U+20AC -> 0x80", `CODE("€")`, 0x80},
		{"Left single quote U+2018 -> 0x91", "CODE(\"\u2018\")", 0x91},
		{"Em dash U+2014 -> 0x97", `CODE("—")`, 0x97},
		{"Trademark U+2122 -> 0x99", `CODE("™")`, 0x99},

		// Latin-1 supplement (0xA0–0xFF) — same in Unicode and Windows-1252
		{"Latin A-grave U+00C0 -> 0xC0", `CODE("À")`, 0xC0},
		{"Section sign U+00A7 -> 0xA7", `CODE("§")`, 0xA7},

		// Characters outside Windows-1252 → replacement '?' = 63
		{"CJK char -> replacement", `CODE("日")`, 63},
		{"Greek alpha -> replacement", `CODE("α")`, 63},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueNumber || got.Num != tt.want {
				t.Errorf("Eval(%q) = %v, want %g", tt.formula, got, tt.want)
			}
		})
	}
}

func TestCHAR_Windows1252(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		// ASCII range — same as before
		{"CHAR(65)", `CHAR(65)`, "A"},
		{"CHAR(95)", `CHAR(95)`, "_"},

		// Windows-1252 range 0x80–0x9F — should produce the correct
		// Unicode character, not the raw byte value.
		{"CHAR(128) = Euro", `CHAR(128)`, "€"},
		{"CHAR(151) = Em dash", `CHAR(151)`, "—"},
		{"CHAR(153) = Trademark", `CHAR(153)`, "™"},

		// Latin-1 supplement — same in both encodings
		{"CHAR(192) = A-grave", `CHAR(192)`, "À"},
		{"CHAR(167) = Section", `CHAR(167)`, "§"},
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

func TestROMAN(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		formula string
		want    string
	}{
		// Basic values
		{`ROMAN(1)`, "I"},
		{`ROMAN(2)`, "II"},
		{`ROMAN(3)`, "III"},
		{`ROMAN(4)`, "IV"},
		{`ROMAN(5)`, "V"},
		{`ROMAN(9)`, "IX"},
		{`ROMAN(10)`, "X"},
		{`ROMAN(14)`, "XIV"},
		{`ROMAN(40)`, "XL"},
		{`ROMAN(44)`, "XLIV"},
		{`ROMAN(49)`, "XLIX"},
		{`ROMAN(50)`, "L"},
		{`ROMAN(90)`, "XC"},
		{`ROMAN(99)`, "XCIX"},
		{`ROMAN(100)`, "C"},
		{`ROMAN(400)`, "CD"},
		{`ROMAN(499)`, "CDXCIX"},
		{`ROMAN(500)`, "D"},
		{`ROMAN(900)`, "CM"},
		{`ROMAN(1000)`, "M"},
		{`ROMAN(1999)`, "MCMXCIX"},
		{`ROMAN(2000)`, "MM"},
		{`ROMAN(3999)`, "MMMCMXCIX"},
		// Zero returns empty string
		{`ROMAN(0)`, ""},
		// String coercion
		{`ROMAN("14")`, "XIV"},
		// Default form (classic)
		{`ROMAN(499, 0)`, "CDXCIX"},
		// Boolean form: TRUE = 0 (Classic)
		{`ROMAN(499, TRUE)`, "CDXCIX"},
		// Form 4 (simplified)
		{`ROMAN(499, 4)`, "ID"},
		// Boolean form: FALSE = 4 (Simplified)
		{`ROMAN(499, FALSE)`, "ID"},
		// Form 1
		{`ROMAN(999, 1)`, "LMVLIV"},
		// Form 2
		{`ROMAN(999, 2)`, "XMIX"},
		// Form 3
		{`ROMAN(999, 3)`, "VMIV"},
		// Form 4
		{`ROMAN(999, 4)`, "IM"},
	}

	for _, tt := range strTests {
		t.Run(tt.formula, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %v (%q), want %q", tt.formula, got, got.Str, tt.want)
			}
		})
	}

	// Error cases
	errTests := []struct {
		formula string
	}{
		{`ROMAN(-1)`},
		{`ROMAN(4000)`},
		{`ROMAN()`},
		{`ROMAN(1,2,3)`},
		{`ROMAN("abc")`},
	}

	for _, tt := range errTests {
		t.Run(tt.formula, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError {
				t.Errorf("Eval(%q) = %v, want error", tt.formula, got)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// VALUETOTEXT tests
// ---------------------------------------------------------------------------

func TestVALUETOTEXT(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		name    string
		formula string
		want    string
	}{
		// Numbers
		{"integer", `VALUETOTEXT(123)`, "123"},
		{"float", `VALUETOTEXT(1.5)`, "1.5"},
		{"negative", `VALUETOTEXT(-42)`, "-42"},
		{"zero", `VALUETOTEXT(0)`, "0"},
		{"large_number", `VALUETOTEXT(1000000)`, "1000000"},
		{"small_float", `VALUETOTEXT(0.001)`, "0.001"},

		// Strings concise (format 0)
		{"string_concise", `VALUETOTEXT("hello")`, "hello"},
		{"string_concise_explicit", `VALUETOTEXT("hello",0)`, "hello"},
		{"empty_string", `VALUETOTEXT("")`, ""},
		{"string_with_spaces", `VALUETOTEXT("hello world")`, "hello world"},

		// Strings strict (format 1)
		{"string_strict", `VALUETOTEXT("hello",1)`, `"hello"`},
		{"empty_string_strict", `VALUETOTEXT("",1)`, `""`},
		{"string_strict_spaces", `VALUETOTEXT("hello world",1)`, `"hello world"`},

		// Booleans
		{"bool_true", `VALUETOTEXT(TRUE)`, "TRUE"},
		{"bool_false", `VALUETOTEXT(FALSE)`, "FALSE"},
		{"bool_true_strict", `VALUETOTEXT(TRUE,1)`, "TRUE"},
		{"bool_false_strict", `VALUETOTEXT(FALSE,1)`, "FALSE"},

		// Numbers with format
		{"number_concise", `VALUETOTEXT(123,0)`, "123"},
		{"number_strict", `VALUETOTEXT(123,1)`, "123"},
		{"float_strict", `VALUETOTEXT(1.5,1)`, "1.5"},
	}

	for _, tt := range strTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %v (%q), want %q", tt.formula, got, got.Str, tt.want)
			}
		})
	}

	// Error cases
	errTests := []struct {
		name    string
		formula string
	}{
		{"no_args", `VALUETOTEXT()`},
		{"too_many_args", `VALUETOTEXT(1,2,3)`},
		{"invalid_format_2", `VALUETOTEXT(1,2)`},
		{"invalid_format_neg", `VALUETOTEXT(1,-1)`},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError {
				t.Errorf("Eval(%q) = %v, want error", tt.formula, got)
			}
		})
	}

	// Error propagation
	t.Run("error_propagation_div0", func(t *testing.T) {
		cf := evalCompile(t, `VALUETOTEXT(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("VALUETOTEXT(1/0) = %v, want error", got)
		}
	})

	t.Run("error_propagation_na", func(t *testing.T) {
		cf := evalCompile(t, `VALUETOTEXT(NA())`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("VALUETOTEXT(NA()) = %v, want error", got)
		}
	})
}

// ---------------------------------------------------------------------------
// ARRAYTOTEXT tests
// ---------------------------------------------------------------------------

func TestARRAYTOTEXT(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		name    string
		formula string
		want    string
	}{
		// Concise format (default)
		{"numbers_concise", `ARRAYTOTEXT({1,2,3})`, "1, 2, 3"},
		{"numbers_concise_explicit", `ARRAYTOTEXT({1,2,3},0)`, "1, 2, 3"},
		{"mixed_concise", `ARRAYTOTEXT({1,"hello",TRUE})`, "1, hello, TRUE"},
		{"single_value", `ARRAYTOTEXT({42})`, "42"},
		{"single_string", `ARRAYTOTEXT({"test"})`, "test"},
		{"floats_concise", `ARRAYTOTEXT({1.5,2.5,3.5})`, "1.5, 2.5, 3.5"},
		{"bools_concise", `ARRAYTOTEXT({TRUE,FALSE})`, "TRUE, FALSE"},

		// Strict format
		{"numbers_strict", `ARRAYTOTEXT({1,2,3},1)`, "{1,2,3}"},
		{"mixed_strict", `ARRAYTOTEXT({1,"hello",TRUE},1)`, `{1,"hello",TRUE}`},
		{"single_value_strict", `ARRAYTOTEXT({42},1)`, "{42}"},
		{"single_string_strict", `ARRAYTOTEXT({"test"},1)`, `{"test"}`},
		{"floats_strict", `ARRAYTOTEXT({1.5,2.5},1)`, "{1.5,2.5}"},
		{"bools_strict", `ARRAYTOTEXT({TRUE,FALSE},1)`, "{TRUE,FALSE}"},

		// Multi-row arrays (semicolon separator)
		{"multirow_concise", `ARRAYTOTEXT({1,2;3,4})`, "1, 2, 3, 4"},
		{"multirow_strict", `ARRAYTOTEXT({1,2;3,4},1)`, "{1,2;3,4}"},

		// Scalar (non-array) input
		{"scalar_number", `ARRAYTOTEXT(42)`, "42"},
		{"scalar_string", `ARRAYTOTEXT("hello")`, "hello"},
		{"scalar_string_strict", `ARRAYTOTEXT("hello",1)`, `{"hello"}`},
		{"scalar_bool", `ARRAYTOTEXT(TRUE)`, "TRUE"},
		{"scalar_number_strict", `ARRAYTOTEXT(42,1)`, "{42}"},
	}

	for _, tt := range strTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %v (%q), want %q", tt.formula, got, got.Str, tt.want)
			}
		})
	}

	// Error cases
	errTests := []struct {
		name    string
		formula string
	}{
		{"no_args", `ARRAYTOTEXT()`},
		{"too_many_args", `ARRAYTOTEXT({1},2,3)`},
		{"invalid_format_2", `ARRAYTOTEXT({1},2)`},
		{"invalid_format_neg", `ARRAYTOTEXT({1},-1)`},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError {
				t.Errorf("Eval(%q) = %v, want error", tt.formula, got)
			}
		})
	}
}
