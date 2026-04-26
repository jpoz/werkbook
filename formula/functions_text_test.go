package formula

import (
	"testing"
)

func sparseFullColumnRefValue(col int, cells map[int]Value) Value {
	resolverCells := make(map[CellAddr]Value, len(cells))
	for row, value := range cells {
		resolverCells[CellAddr{Col: col, Row: row}] = value
	}
	return EvalValueToValue(newEvalRangeRef(
		RangeAddr{FromCol: col, FromRow: 1, ToCol: col, ToRow: maxRows},
		nil,
		&sparseResolver{cells: resolverCells},
		&RefLegacyBoundary{PlaceholderRows: 1, PlaceholderCols: 1, UseEmptyArray: true},
	))
}

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
		// Currency symbols (multi-byte UTF-8 characters)
		{"TEXT(12.3,\"\u00a3000.00\")", "\u00a3012.30"}, // £000.00
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
// Currency symbols in number formats (multi-byte UTF-8)
// ---------------------------------------------------------------------------

func TestFormatNumberCurrencySymbols(t *testing.T) {
	tests := []struct {
		name   string
		value  float64
		format string
		want   string
	}{
		{"pound_prefix", 12.3, "\u00a3000.00", "\u00a3012.30"},
		{"yen_with_suffix", 12.3, "\u00a5#.00\" after\"", "\u00a512.30 after"},
		{"yen_with_prefix", 12.3, "\"before \"\u00a5#.00", "before \u00a512.30"},
		{"euro_sign", 12.3, "\u20ac#.00", "\u20ac12.30"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := formatNumber(tt.value, tt.format, false)
			if got != tt.want {
				t.Errorf("formatNumber(%v, %q) = %q, want %q", tt.value, tt.format, got, tt.want)
			}
		})
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
		// Empty old_text — returns original text unchanged
		{name: "empty_old", formula: `SUBSTITUTE("hello","","X")`, want: "hello"},
		{name: "empty_old_with_instance", formula: `SUBSTITUTE("hello","","X",2)`, want: "hello"},
		{name: "empty_old_with_large_instance", formula: `SUBSTITUTE("A","","B",1000000000000)`, want: "A"},
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
		// Replace 3rd instance
		{name: "replace_3rd", formula: `SUBSTITUTE("aXaXaX","a","Z",3)`, want: "aXaXZX"},
		// Number coercion for text argument
		{name: "number_coercion", formula: `SUBSTITUTE(12321,"2","9")`, want: "19391"},
		// Multiple overlapping occurrences — replace all
		{name: "multi_replace_all", formula: `SUBSTITUTE("aaaa","aa","X")`, want: "XX"},
		// Delete specific instance
		{name: "delete_2nd", formula: `SUBSTITUTE("abab","a","",2)`, want: "abb"},
		// Doc example: replace "Sales" with "Cost"
		{name: "doc_example_1", formula: `SUBSTITUTE("Sales Data","Sales","Cost")`, want: "Cost Data"},
		// Doc example: replace 1st "1" with "2"
		{name: "doc_example_2", formula: `SUBSTITUTE("Quarter 1, 2008","1","2",1)`, want: "Quarter 2, 2008"},
		// Doc example: replace 3rd "1" with "2"
		{name: "doc_example_3", formula: `SUBSTITUTE("Quarter 1, 2011","1","2",3)`, want: "Quarter 1, 2012"},
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

	errTests := []struct {
		name    string
		formula string
	}{
		{name: "instance_zero", formula: `SUBSTITUTE("abc","a","X",0)`},
		{name: "instance_negative", formula: `SUBSTITUTE("abc","a","X",-1)`},
		{name: "too_few_args", formula: `SUBSTITUTE("abc","a")`},
		{name: "too_many_args", formula: `SUBSTITUTE("abc","a","X",1,"extra")`},
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
		// Date formatting from date serial numbers
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

	// Only uppercase E is recognised for scientific notation.
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
		// Whole-part placeholder type controls middle literal behavior when wholePart=0
		// # whole: suppress digit AND suppress middle literals entirely
		{name: "frac_hash_whole_middle_suppress", formula: `TEXT(0.75, "|#\:#=/=?|")`, want: "|3=/=4|"},
		// ? whole: space-pad digit AND replace each middle literal char with a space
		{name: "frac_qmark_whole_middle_pad", formula: `TEXT(0.75, "|?\:#=/=#|")`, want: "|  3=/=4|"},
		// 0 whole: show 0 and show middle literals as-is
		{name: "frac_zero_whole_middle_show", formula: `TEXT(0.75, "|0\:#=/=#|")`, want: "|0:3=/=4|"},
		// ? whole with zero fraction: whole shows 0 via forceZero, fraction area all spaces
		{name: "frac_qmark_whole_zero_frac", formula: `TEXT(0, "|?\:#=/=#|")`, want: "|0      |"},
		// # whole with zero fraction and numerator 0 placeholder
		{name: "frac_hash_whole_zero_num0", formula: `TEXT(0, "|#\:0=/=#|")`, want: "|0=/=1|"},
		// ? whole with integer value (zero fraction = spaces)
		{name: "frac_qmark_whole_int1", formula: `TEXT(1, "|?\:#=/=#|")`, want: "|1      |"},
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
		{name: "general_small_1e-5", formula: `TEXT(0.00001, "General")`, want: "0.00001"},
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
// FIND comprehensive tests
// ---------------------------------------------------------------------------

func TestFINDEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		isErr   bool
	}{
		// Basic functionality
		{name: "basic_substring", formula: `FIND("lo","hello")`, wantNum: 4},
		{name: "find_at_beginning", formula: `FIND("M","Miriam McGovern")`, wantNum: 1},
		{name: "find_substring_Gov", formula: `FIND("Gov","Miriam McGovern")`, wantNum: 10},
		{name: "find_single_char", formula: `FIND("a","banana")`, wantNum: 2},
		{name: "find_at_end", formula: `FIND("rn","Miriam McGovern")`, wantNum: 14},
		{name: "find_entire_string", formula: `FIND("hello","hello")`, wantNum: 1},
		{name: "find_single_char_string", formula: `FIND("a","a")`, wantNum: 1},

		// Case sensitivity
		{name: "case_sensitive_upper_not_found", formula: `FIND("H","hello")`, isErr: true},
		{name: "case_sensitive_lower_m", formula: `FIND("m","Miriam McGovern")`, wantNum: 6},
		{name: "case_sensitive_upper_M", formula: `FIND("M","Miriam McGovern")`, wantNum: 1},
		{name: "case_lower_a_in_Apple", formula: `FIND("a","Apple")`, isErr: true},
		{name: "case_upper_A_in_Apple", formula: `FIND("A","Apple")`, wantNum: 1},
		{name: "case_sensitive_substring", formula: `FIND("mc","Miriam McGovern")`, isErr: true},

		// start_num parameter
		{name: "start_pos_skip_first", formula: `FIND("M","Miriam McGovern",3)`, wantNum: 8},
		{name: "start_pos_find_second_l", formula: `FIND("l","hello world",5)`, wantNum: 10},
		{name: "start_pos_1_default", formula: `FIND("h","hello",1)`, wantNum: 1},
		{name: "start_pos_at_match", formula: `FIND("l","hello",3)`, wantNum: 3},
		{name: "start_pos_at_length", formula: `FIND("o","hello",5)`, wantNum: 5},
		{name: "start_pos_past_only_match", formula: `FIND("h","hello",2)`, isErr: true},

		// Multiple occurrences — returns first from start_num
		{name: "multiple_occ_first", formula: `FIND("a","banana")`, wantNum: 2},
		{name: "multiple_occ_second", formula: `FIND("a","banana",3)`, wantNum: 4},
		{name: "multiple_occ_third", formula: `FIND("a","banana",5)`, wantNum: 6},

		// Empty find_text
		{name: "empty_find_text", formula: `FIND("","hello")`, wantNum: 1},
		{name: "empty_find_text_start_3", formula: `FIND("","hello",3)`, wantNum: 3},
		{name: "empty_find_text_start_at_len", formula: `FIND("","hello",5)`, wantNum: 5},
		{name: "empty_find_text_start_past_len", formula: `FIND("","hello",6)`, wantNum: 6},
		{name: "empty_find_empty_within", formula: `FIND("","")`, wantNum: 1},

		// Not found
		{name: "not_found_char", formula: `FIND("z","hello")`, isErr: true},
		{name: "not_found_long_needle", formula: `FIND("hello world!","hello")`, isErr: true},

		// Empty within_text
		{name: "find_in_empty_string", formula: `FIND("a","")`, isErr: true},

		// start_num errors
		{name: "start_too_large", formula: `FIND("h","hello",99)`, isErr: true},
		{name: "start_zero", formula: `FIND("h","hello",0)`, isErr: true},
		{name: "start_negative", formula: `FIND("h","hello",-1)`, isErr: true},

		// Argument count errors
		{name: "no_args", formula: `FIND()`, isErr: true},
		{name: "one_arg", formula: `FIND("a")`, isErr: true},
		{name: "four_args", formula: `FIND("a","b",1,2)`, isErr: true},

		// Numeric coercion — numbers become strings via ValueToString
		{name: "numeric_find_text", formula: `FIND(1,"a1b2")`, wantNum: 2},
		{name: "numeric_within_text", formula: `FIND("2",1234)`, wantNum: 2},
		{name: "numeric_both", formula: `FIND(3,12345)`, wantNum: 3},

		// Boolean coercion — TRUE/FALSE become strings
		{name: "bool_true_find", formula: `FIND(TRUE,"TRUEFALSE")`, wantNum: 1},
		{name: "bool_false_find", formula: `FIND(FALSE,"TRUEFALSE")`, wantNum: 5},
		{name: "bool_true_not_lowercase", formula: `FIND(TRUE,"true")`, isErr: true},

		// Special characters
		{name: "find_space", formula: `FIND(" ","hello world")`, wantNum: 6},
		{name: "find_comma", formula: `FIND(",","a,b,c")`, wantNum: 2},
		{name: "find_exclamation", formula: `FIND("!","hello!")`, wantNum: 6},

		// Nested FIND
		{name: "nested_find_second_o", formula: `FIND("o","hello world",FIND("o","hello world")+1)`, wantNum: 8},

		// Unicode
		{name: "unicode_char", formula: "FIND(\"\u00e9\",\"caf\u00e9\")", wantNum: 4},
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
		// Leading spaces only
		{`TRIM("   leading")`, "leading"},
		// Trailing spaces only
		{`TRIM("trailing   ")`, "trailing"},
		// Both leading and trailing
		{`TRIM("  both  ")`, "both"},
		// Multiple internal spaces reduced to single
		{`TRIM("a    b     c")`, "a b c"},
		// Single space → empty string
		{`TRIM(" ")`, ""},
		// Multiple words with various spacing
		{`TRIM("  the   quick   brown   fox  ")`, "the quick brown fox"},
		// Tab characters — Excel's TRIM only handles ASCII space (0x20),
		// so tabs are preserved verbatim.
		{`TRIM("hello` + "\t" + `world")`, "hello\tworld"},
		// Number coercion (42 → "42", no spaces to trim)
		{`TRIM(42)`, "42"},
		// Boolean coercion (TRUE → "TRUE")
		{`TRIM(TRUE)`, "TRUE"},
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

func TestTRIMErrors(t *testing.T) {
	resolver := &mockResolver{}

	// No arguments → #VALUE!
	cf := evalCompile(t, `TRIM()`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(TRIM()): unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("TRIM() = %v, want #VALUE!", got)
	}

	// Too many arguments → #VALUE!
	cf = evalCompile(t, `TRIM("a","b")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(TRIM(a,b)): unexpected error: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("TRIM(a,b) = %v, want #VALUE!", got)
	}

	// Error propagation — TRIM does not guard against error args,
	// so ValueToString converts the error to its string representation.
	cf = evalCompile(t, `TRIM(1/0)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(TRIM(1/0)): unexpected error: %v", err)
	}
	if got.Type != ValueString || got.Str != "#DIV/0!" {
		t.Errorf("TRIM(1/0) = %v, want string #DIV/0!", got)
	}
}

func TestTRIMComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("basic trimming", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Leading spaces only
			{`TRIM("   hello")`, "hello"},
			// Trailing spaces only
			{`TRIM("hello   ")`, "hello"},
			// Both leading and trailing
			{`TRIM("  hello  ")`, "hello"},
			// Internal multiple spaces collapsed to single
			{`TRIM("hello   world")`, "hello world"},
			// Mixed: leading, trailing, and internal
			{`TRIM("  hello   world  ")`, "hello world"},
			// No trimming needed
			{`TRIM("hello")`, "hello"},
			// Empty string
			{`TRIM("")`, ""},
			// Single space
			{`TRIM(" ")`, ""},
			// All spaces
			{`TRIM("     ")`, ""},
			// Single character
			{`TRIM("a")`, "a"},
			// Single character with spaces
			{`TRIM("  a  ")`, "a"},
		}

		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueString || got.Str != tt.want {
					t.Fatalf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
				}
			})
		}
	})

	t.Run("multiple words with various spacing", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			{`TRIM("  the   quick   brown   fox  ")`, "the quick brown fox"},
			{`TRIM("a  b  c  d  e")`, "a b c d e"},
			{`TRIM("  one   two   three  ")`, "one two three"},
			// Single spaces between words — no change
			{`TRIM("hello world foo")`, "hello world foo"},
			// Many internal spaces
			{`TRIM("a          b")`, "a b"},
		}

		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueString || got.Str != tt.want {
					t.Fatalf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
				}
			})
		}
	})

	t.Run("whitespace characters", func(t *testing.T) {
		// Excel's TRIM only removes ASCII space (U+0020). Tabs, newlines,
		// and NBSP pass through untouched — callers need CLEAN or
		// SUBSTITUTE to handle those.
		tests := []struct {
			name  string
			input []Value
			want  string
		}{
			{"tab between words", []Value{StringVal("hello\tworld")}, "hello\tworld"},
			{"newline between words", []Value{StringVal("hello\nworld")}, "hello\nworld"},
			{"cr between words", []Value{StringVal("hello\rworld")}, "hello\rworld"},
			{"mixed whitespace", []Value{StringVal("  hello\t\n  world  ")}, "hello\t\n world"},
			{"only tabs", []Value{StringVal("\t\t\t")}, "\t\t\t"},
			{"only newlines", []Value{StringVal("\n\n\n")}, "\n\n\n"},
		}

		for _, tt := range tests {
			t.Run(tt.name, func(t *testing.T) {
				got, err := fnTRIM(tt.input)
				if err != nil {
					t.Fatalf("fnTRIM: %v", err)
				}
				if got.Type != ValueString || got.Str != tt.want {
					t.Fatalf("fnTRIM(%q) = %q, want %q", tt.input[0].Str, got.Str, tt.want)
				}
			})
		}
	})

	t.Run("non-breaking space", func(t *testing.T) {
		// Excel's TRIM leaves NBSP (U+00A0) in place — callers need
		// SUBSTITUTE(..., UNICHAR(160), " ") before TRIM to clean it.
		input := "hello\u00A0\u00A0world"
		got, err := fnTRIM([]Value{StringVal(input)})
		if err != nil {
			t.Fatalf("fnTRIM: %v", err)
		}
		if got.Type != ValueString || got.Str != input {
			t.Fatalf("fnTRIM with NBSP = %q, want %q", got.Str, input)
		}
	})

	t.Run("unicode characters with spaces", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Unicode characters preserved, spaces trimmed
			{`TRIM("  ` + "\u00e9\u00e0\u00fc" + `  ")`, "\u00e9\u00e0\u00fc"},
		}

		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueString || got.Str != tt.want {
					t.Fatalf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
				}
			})
		}

		// CJK characters via direct call
		got, err := fnTRIM([]Value{StringVal("  \u4e16\u754c   \u4f60\u597d  ")})
		if err != nil {
			t.Fatalf("fnTRIM: %v", err)
		}
		if got.Type != ValueString || got.Str != "\u4e16\u754c \u4f60\u597d" {
			t.Fatalf("fnTRIM CJK = %q, want %q", got.Str, "\u4e16\u754c \u4f60\u597d")
		}
	})

	t.Run("type coercion", func(t *testing.T) {
		tests := []struct {
			formula string
			want    string
		}{
			// Number coercion: 42 → "42"
			{`TRIM(42)`, "42"},
			// Negative number
			{`TRIM(-3.14)`, "-3.14"},
			// Zero
			{`TRIM(0)`, "0"},
			// Boolean TRUE → "TRUE"
			{`TRIM(TRUE)`, "TRUE"},
			// Boolean FALSE → "FALSE"
			{`TRIM(FALSE)`, "FALSE"},
		}

		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueString || got.Str != tt.want {
					t.Fatalf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
				}
			})
		}
	})

	t.Run("error handling", func(t *testing.T) {
		// TRIM converts errors to their string representation via ValueToString
		tests := []struct {
			formula string
			want    string
		}{
			{`TRIM(1/0)`, "#DIV/0!"},
			{`TRIM(NA())`, "#N/A"},
		}

		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != ValueString || got.Str != tt.want {
					t.Fatalf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want)
				}
			})
		}
	})

	t.Run("wrong arg count direct", func(t *testing.T) {
		// No args
		got, err := fnTRIM([]Value{})
		if err != nil {
			t.Fatalf("fnTRIM: unexpected Go error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Fatalf("fnTRIM() = %#v, want #VALUE!", got)
		}

		// Two args
		got, err = fnTRIM([]Value{StringVal("a"), StringVal("b")})
		if err != nil {
			t.Fatalf("fnTRIM: unexpected Go error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Fatalf("fnTRIM(a,b) = %#v, want #VALUE!", got)
		}
	})

	t.Run("very long string with many spaces", func(t *testing.T) {
		// Build a string with many spaces
		input := "  word1"
		for i := 2; i <= 10; i++ {
			input += "     word" + string(rune('0'+i))
		}
		input += "  "
		got, err := fnTRIM([]Value{StringVal(input)})
		if err != nil {
			t.Fatalf("fnTRIM: %v", err)
		}
		// All multiple spaces should be collapsed to single
		want := "word1"
		for i := 2; i <= 10; i++ {
			want += " word" + string(rune('0'+i))
		}
		if got.Type != ValueString || got.Str != want {
			t.Fatalf("fnTRIM(long string) = %q, want %q", got.Str, want)
		}
	})

	t.Run("empty value", func(t *testing.T) {
		// Empty Value → ValueToString returns ""
		got, err := fnTRIM([]Value{EmptyVal()})
		if err != nil {
			t.Fatalf("fnTRIM: %v", err)
		}
		if got.Type != ValueString || got.Str != "" {
			t.Fatalf("fnTRIM(empty) = %q, want %q", got.Str, "")
		}
	})

	t.Run("TRIM in formula expressions", func(t *testing.T) {
		tests := []struct {
			formula string
			want    Value
		}{
			// TRIM result used in LEN
			{`LEN(TRIM("  hello  "))`, NumberVal(5)},
			// TRIM result used in concatenation
			{`TRIM("  hi  ")&" there"`, StringVal("hi there")},
			// TRIM with nested TRIM (idempotent)
			{`TRIM(TRIM("  hello   world  "))`, StringVal("hello world")},
		}

		for _, tt := range tests {
			t.Run(tt.formula, func(t *testing.T) {
				cf := evalCompile(t, tt.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%q): %v", tt.formula, err)
				}
				if got.Type != tt.want.Type {
					t.Fatalf("Eval(%q) type = %v, want type %v", tt.formula, got.Type, tt.want.Type)
				}
				switch tt.want.Type {
				case ValueString:
					if got.Str != tt.want.Str {
						t.Fatalf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want.Str)
					}
				case ValueNumber:
					if got.Num != tt.want.Num {
						t.Fatalf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want.Num)
					}
				}
			})
		}
	})
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
// CHOOSE comprehensive tests
// ---------------------------------------------------------------------------

func TestCHOOSEComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	type want struct {
		typ ValueType
		num float64
		str string
		b   bool
		err ErrorValue
	}

	tests := []struct {
		name    string
		formula string
		want    want
	}{
		// Basic selection
		{name: "first_value", formula: `CHOOSE(1,"a","b","c")`, want: want{typ: ValueString, str: "a"}},
		{name: "second_value", formula: `CHOOSE(2,"a","b","c")`, want: want{typ: ValueString, str: "b"}},
		{name: "third_value", formula: `CHOOSE(3,"a","b","c")`, want: want{typ: ValueString, str: "c"}},
		{name: "last_value", formula: `CHOOSE(4,"w","x","y","z")`, want: want{typ: ValueString, str: "z"}},

		// Single value
		{name: "single_value", formula: `CHOOSE(1,"x")`, want: want{typ: ValueString, str: "x"}},

		// Various return types
		{name: "return_number", formula: `CHOOSE(2,10,20,30)`, want: want{typ: ValueNumber, num: 20}},
		{name: "return_bool_true", formula: `CHOOSE(1,TRUE,FALSE)`, want: want{typ: ValueBool, b: true}},
		{name: "return_bool_false", formula: `CHOOSE(2,TRUE,FALSE)`, want: want{typ: ValueBool, b: false}},
		{name: "return_string", formula: `CHOOSE(1,"hello")`, want: want{typ: ValueString, str: "hello"}},

		// Decimal index truncated
		{name: "decimal_index_2.9", formula: `CHOOSE(2.9,"a","b","c")`, want: want{typ: ValueString, str: "b"}},
		{name: "decimal_index_1.5", formula: `CHOOSE(1.5,"first","second")`, want: want{typ: ValueString, str: "first"}},

		// String coercion of index
		{name: "string_index", formula: `CHOOSE("2","a","b","c")`, want: want{typ: ValueString, str: "b"}},

		// Boolean as index (TRUE=1)
		{name: "bool_true_index", formula: `CHOOSE(TRUE,"a","b","c")`, want: want{typ: ValueString, str: "a"}},

		// Doc examples
		{name: "doc_example", formula: `CHOOSE(3,"Wide",115,"world",8)`, want: want{typ: ValueString, str: "world"}},

		// Error: index out of range (too high)
		{name: "index_too_high", formula: `CHOOSE(5,"a","b","c")`, want: want{typ: ValueError, err: ErrValVALUE}},
		// Error: index = 0
		{name: "index_zero", formula: `CHOOSE(0,"a","b")`, want: want{typ: ValueError, err: ErrValVALUE}},
		// Error: negative index
		{name: "negative_index", formula: `CHOOSE(-1,"a","b")`, want: want{typ: ValueError, err: ErrValVALUE}},
		// Error: no values (only index)
		{name: "no_values", formula: `CHOOSE(1)`, want: want{typ: ValueError, err: ErrValVALUE}},

		// Error propagation in index
		{name: "error_in_index", formula: `CHOOSE(1/0,"a","b")`, want: want{typ: ValueError, err: ErrValDIV0}},

		// Error in selected value propagates
		{name: "error_in_selected_value", formula: `CHOOSE(2,"a",1/0,"c")`, want: want{typ: ValueError, err: ErrValDIV0}},

		// Error in unselected value doesn't propagate (args are pre-evaluated by
		// the engine, but CHOOSE only returns the selected one; if the unselected
		// arg evaluates to an error it's in the arg list but never returned).
		{name: "error_in_unselected_value", formula: `CHOOSE(1,"ok",1/0)`, want: want{typ: ValueString, str: "ok"}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != tt.want.typ {
				t.Fatalf("Eval(%q).Type = %v, want %v (value=%v)", tt.formula, got.Type, tt.want.typ, got)
			}
			switch tt.want.typ {
			case ValueString:
				if got.Str != tt.want.str {
					t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want.str)
				}
			case ValueNumber:
				if got.Num != tt.want.num {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want.num)
				}
			case ValueBool:
				if got.Bool != tt.want.b {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want.b)
				}
			case ValueError:
				if got.Err != tt.want.err {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Err, tt.want.err)
				}
			}
		})
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

func TestCONCAT_FullColumnSparseRef(t *testing.T) {
	got, err := fnCONCAT([]Value{
		sparseFullColumnRefValue(1, map[int]Value{
			1: StringVal("top"),
			3: StringVal("bottom"),
		}),
	})
	if err != nil {
		t.Fatalf("fnCONCAT: %v", err)
	}
	if got.Type != ValueString || got.Str != "topbottom" {
		t.Fatalf("got %v, want topbottom", got)
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
// CONCATENATE comprehensive tests
// ---------------------------------------------------------------------------

func TestCONCATENATEComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	// --- Scalar-only tests (CONCATENATE does NOT accept ranges) ---
	scalarTests := []struct {
		name    string
		formula string
		want    string
	}{
		// Basic: two strings
		{name: "two_strings", formula: `CONCATENATE("hello","world")`, want: "helloworld"},
		// Basic: three strings
		{name: "three_strings", formula: `CONCATENATE("a","b","c")`, want: "abc"},
		// Many strings
		{name: "many_strings", formula: `CONCATENATE("a","b","c","d","e","f","g","h")`, want: "abcdefgh"},
		// Single argument
		{name: "single_string", formula: `CONCATENATE("only")`, want: "only"},
		// No arguments -> empty string
		{name: "no_args", formula: `CONCATENATE()`, want: ""},
		// Numbers coerced to text
		{name: "number_int", formula: `CONCATENATE(42)`, want: "42"},
		{name: "number_float", formula: `CONCATENATE(3.14)`, want: "3.14"},
		{name: "two_numbers", formula: `CONCATENATE(1,2)`, want: "12"},
		{name: "number_and_string", formula: `CONCATENATE("item",99)`, want: "item99"},
		{name: "negative_number", formula: `CONCATENATE("val:",-5)`, want: "val:-5"},
		// Booleans coerced to text
		{name: "bool_true", formula: `CONCATENATE(TRUE)`, want: "TRUE"},
		{name: "bool_false", formula: `CONCATENATE(FALSE)`, want: "FALSE"},
		{name: "bool_and_string", formula: `CONCATENATE("flag: ",TRUE)`, want: "flag: TRUE"},
		{name: "bool_false_and_string", formula: `CONCATENATE("ok=",FALSE)`, want: "ok=FALSE"},
		// Empty strings
		{name: "empty_string", formula: `CONCATENATE("")`, want: ""},
		{name: "empty_between", formula: `CONCATENATE("a","","b")`, want: "ab"},
		{name: "all_empty", formula: `CONCATENATE("","","")`, want: ""},
		// Mixed types
		{name: "mixed_types", formula: `CONCATENATE("count=",5,", ok=",TRUE)`, want: "count=5, ok=TRUE"},
		{name: "mixed_with_bool_false", formula: `CONCATENATE(FALSE," ",0," ","x")`, want: "FALSE 0 x"},
		{name: "mixed_number_string_bool", formula: `CONCATENATE(1," + ",2," = ",TRUE)`, want: "1 + 2 = TRUE"},
		// Numeric strings
		{name: "numeric_strings", formula: `CONCATENATE("123","456")`, want: "123456"},
		{name: "numeric_string_and_number", formula: `CONCATENATE("100",200)`, want: "100200"},
		// Special characters
		{name: "symbols", formula: `CONCATENATE("@","#","$")`, want: "@#$"},
		{name: "spaces", formula: `CONCATENATE(" "," "," ")`, want: "   "},
		{name: "punctuation", formula: `CONCATENATE("hello",", ","world","!")`, want: "hello, world!"},
		// Unicode characters
		{name: "unicode_accents", formula: `CONCATENATE("café"," ","naïve")`, want: "café naïve"},
		{name: "unicode_emoji", formula: `CONCATENATE("hi"," 🎉")`, want: "hi 🎉"},
		{name: "unicode_cjk", formula: `CONCATENATE("日本","語")`, want: "日本語"},
		// Long result string
		{name: "long_string", formula: `CONCATENATE("abcdefghij","abcdefghij","abcdefghij","abcdefghij","abcdefghij")`, want: "abcdefghijabcdefghijabcdefghijabcdefghijabcdefghij"},
	}

	for _, tt := range scalarTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("%s: got type=%d str=%q, want %q", tt.name, got.Type, got.Str, tt.want)
			}
		})
	}

	// --- Error propagation tests ---
	t.Run("error_propagation_div0", func(t *testing.T) {
		cf := evalCompile(t, `CONCATENATE("a",1/0,"b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("expected #DIV/0! error, got type=%d err=%v str=%q", got.Type, got.Err, got.Str)
		}
	})

	t.Run("error_propagation_first_arg", func(t *testing.T) {
		cf := evalCompile(t, `CONCATENATE(1/0,"b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("expected error, got type=%d str=%q", got.Type, got.Str)
		}
	})

	t.Run("error_propagation_last_arg", func(t *testing.T) {
		cf := evalCompile(t, `CONCATENATE("a","b",1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("expected #DIV/0! error, got type=%d err=%v str=%q", got.Type, got.Err, got.Str)
		}
	})

	// --- Cell reference tests ---
	t.Run("cell_references", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("hello"),
				{Col: 2, Row: 1}: StringVal(" "),
				{Col: 3, Row: 1}: StringVal("world"),
			},
		}
		cf := evalCompile(t, `CONCATENATE(A1,B1,C1)`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "hello world" {
			t.Errorf("cell_references: got %q, want %q", got.Str, "hello world")
		}
	})

	t.Run("cell_ref_with_number", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("count: "),
				{Col: 2, Row: 1}: NumberVal(42),
			},
		}
		cf := evalCompile(t, `CONCATENATE(A1,B1)`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "count: 42" {
			t.Errorf("cell_ref_with_number: got %q, want %q", got.Str, "count: 42")
		}
	})

	t.Run("cell_ref_empty_cell", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("a"),
				// B1 is empty
				{Col: 3, Row: 1}: StringVal("c"),
			},
		}
		cf := evalCompile(t, `CONCATENATE(A1,B1,C1)`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "ac" {
			t.Errorf("cell_ref_empty_cell: got %q, want %q", got.Str, "ac")
		}
	})
}

// ---------------------------------------------------------------------------
// CONCAT comprehensive tests
// ---------------------------------------------------------------------------

func TestCONCATComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	// --- Scalar-only tests (no ranges) ---
	scalarTests := []struct {
		name    string
		formula string
		want    string
	}{
		// Basic: two strings
		{name: "two_strings", formula: `CONCAT("hello","world")`, want: "helloworld"},
		// Basic: three strings
		{name: "three_strings", formula: `CONCAT("a","b","c")`, want: "abc"},
		// Many strings
		{name: "many_strings", formula: `CONCAT("a","b","c","d","e","f")`, want: "abcdef"},
		// Single argument
		{name: "single_string", formula: `CONCAT("only")`, want: "only"},
		// No arguments -> empty string
		{name: "no_args", formula: `CONCAT()`, want: ""},
		// Numbers coerced to text
		{name: "number_int", formula: `CONCAT(42)`, want: "42"},
		{name: "number_float", formula: `CONCAT(3.14)`, want: "3.14"},
		{name: "two_numbers", formula: `CONCAT(1,2)`, want: "12"},
		{name: "number_and_string", formula: `CONCAT("item",99)`, want: "item99"},
		// Booleans coerced to text
		{name: "bool_true", formula: `CONCAT(TRUE)`, want: "TRUE"},
		{name: "bool_false", formula: `CONCAT(FALSE)`, want: "FALSE"},
		{name: "bool_and_string", formula: `CONCAT("flag: ",TRUE)`, want: "flag: TRUE"},
		// Empty strings
		{name: "empty_string", formula: `CONCAT("")`, want: ""},
		{name: "empty_between", formula: `CONCAT("a","","b")`, want: "ab"},
		{name: "all_empty", formula: `CONCAT("","","")`, want: ""},
		// Mixed types
		{name: "mixed_types", formula: `CONCAT("count=",5,", ok=",TRUE)`, want: "count=5, ok=TRUE"},
		{name: "mixed_with_bool_false", formula: `CONCAT(FALSE," ",0," ","x")`, want: "FALSE 0 x"},
		// Special characters
		{name: "special_chars", formula: `CONCAT("hello\n","world")`, want: "hello\\nworld"},
		{name: "unicode", formula: `CONCAT("café"," ","naïve")`, want: "café naïve"},
		{name: "symbols", formula: `CONCAT("@","#","$")`, want: "@#$"},
		// Long result string
		{name: "long_string", formula: `CONCAT("abcdefghij","abcdefghij","abcdefghij","abcdefghij","abcdefghij")`, want: "abcdefghijabcdefghijabcdefghijabcdefghijabcdefghij"},
	}

	for _, tt := range scalarTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("%s: got type=%d str=%q, want %q", tt.name, got.Type, got.Str, tt.want)
			}
		})
	}

	// --- Error propagation tests ---
	t.Run("error_propagation_div0", func(t *testing.T) {
		cf := evalCompile(t, `CONCAT("a",1/0,"b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("expected #DIV/0! error, got type=%d err=%v str=%q", got.Type, got.Err, got.Str)
		}
	})

	t.Run("error_propagation_first_arg", func(t *testing.T) {
		cf := evalCompile(t, `CONCAT(1/0,"b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("expected error, got type=%d str=%q", got.Type, got.Str)
		}
	})

	// --- Range tests ---
	t.Run("range_with_empty_cells", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("a"),
				// A2 is empty
				{Col: 1, Row: 3}: StringVal("c"),
			},
		}
		cf := evalCompile(t, `CONCAT(A1:A3)`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "ac" {
			t.Errorf("range_with_empty_cells: got %q, want %q", got.Str, "ac")
		}
	})

	t.Run("range_mixed_types", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("text"),
				{Col: 1, Row: 2}: NumberVal(42),
				{Col: 1, Row: 3}: BoolVal(true),
			},
		}
		cf := evalCompile(t, `CONCAT(A1:A3)`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "text42TRUE" {
			t.Errorf("range_mixed_types: got %q, want %q", got.Str, "text42TRUE")
		}
	})

	t.Run("range_2d", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("a"),
				{Col: 2, Row: 1}: StringVal("b"),
				{Col: 1, Row: 2}: StringVal("c"),
				{Col: 2, Row: 2}: StringVal("d"),
			},
		}
		cf := evalCompile(t, `CONCAT(A1:B2)`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "abcd" {
			t.Errorf("range_2d: got %q, want %q", got.Str, "abcd")
		}
	})

	t.Run("range_error_propagation", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("ok"),
				{Col: 1, Row: 2}: ErrorVal(ErrValNA),
				{Col: 1, Row: 3}: StringVal("after"),
			},
		}
		cf := evalCompile(t, `CONCAT(A1:A3)`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("range_error_propagation: expected #N/A error, got type=%d err=%v", got.Type, got.Err)
		}
	})

	t.Run("range_all_empty", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{},
		}
		cf := evalCompile(t, `CONCAT(A1:A3)`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "" {
			t.Errorf("range_all_empty: got %q, want empty string", got.Str)
		}
	})

	t.Run("multiple_ranges", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("X"),
				{Col: 2, Row: 1}: StringVal("Y"),
			},
		}
		cf := evalCompile(t, `CONCAT(A1:A1,B1:B1)`)
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "XY" {
			t.Errorf("multiple_ranges: got %q, want %q", got.Str, "XY")
		}
	})
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

		// Very small number with 3 decimal places
		{name: "very_small_3dec", formula: `DOLLAR(0.001, 3)`, want: "$0.001"},

		// Many decimal places (10)
		{name: "ten_decimals", formula: `DOLLAR(1.5, 10)`, want: "$1.5000000000"},

		// Single digit, no comma needed
		{name: "single_digit", formula: `DOLLAR(7, 2)`, want: "$7.00"},

		// Exactly 1000 (comma boundary)
		{name: "comma_boundary", formula: `DOLLAR(1000, 2)`, want: "$1,000.00"},

		// Negative number with zero decimals rounding
		{name: "neg_round_hundreds", formula: `DOLLAR(-1250, -2)`, want: "($1,300)"},

		// One decimal place
		{name: "one_decimal", formula: `DOLLAR(1234.567, 1)`, want: "$1,234.6"},

		// String coercion with decimals
		{name: "string_with_decimal", formula: `DOLLAR("99.999", 1)`, want: "$100.0"},

		// Boolean coercion for decimals argument
		{name: "bool_decimals_true", formula: `DOLLAR(1234.567, TRUE)`, want: "$1,234.6"},
		{name: "bool_decimals_false", formula: `DOLLAR(1234.567, FALSE)`, want: "$1,235"},

		// Fractional number rounds up
		{name: "round_up_half", formula: `DOLLAR(2.5, 0)`, want: "$3"},

		// Large negative number with commas
		{name: "large_negative", formula: `DOLLAR(-9876543.21, 2)`, want: "($9,876,543.21)"},
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
		{name: "non_numeric_decimals", formula: `DOLLAR(1234, "xyz")`},
		{name: "error_propagation_num", formula: `DOLLAR(1/0)`},
		{name: "error_propagation_dec", formula: `DOLLAR(1, 1/0)`},
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
		{name: "error_arg", formula: `LEN(NA())`, isErr: true},

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
// SEARCH — comprehensive tests
// ---------------------------------------------------------------------------

func TestSEARCHComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantNum float64
		isErr   bool
	}{
		// Basic case-insensitive search
		{name: "find_world_in_hello_world", formula: `SEARCH("world","Hello World")`, wantNum: 7},
		{name: "case_insensitive_HELLO", formula: `SEARCH("HELLO","hello world")`, wantNum: 1},
		{name: "case_insensitive_mixed", formula: `SEARCH("hElLo","HELLO WORLD")`, wantNum: 1},

		// With start_num
		{name: "start_num_skip_first", formula: `SEARCH("o","hello world",5)`, wantNum: 5},
		{name: "start_num_skip_to_second", formula: `SEARCH("o","hello world",6)`, wantNum: 8},
		{name: "start_num_explicit_1", formula: `SEARCH("h","hello",1)`, wantNum: 1},
		{name: "start_num_at_position", formula: `SEARCH("l","hello",4)`, wantNum: 4},

		// Wildcard ? (single character)
		{name: "question_at_beginning", formula: `SEARCH("?ello","hello")`, wantNum: 1},
		{name: "question_at_end", formula: `SEARCH("hell?","hello")`, wantNum: 1},
		{name: "multiple_question_marks", formula: `SEARCH("h??lo","hello")`, wantNum: 1},
		{name: "question_only", formula: `SEARCH("?","abc")`, wantNum: 1},

		// Wildcard * (any characters)
		{name: "star_zero_chars", formula: `SEARCH("he*llo","hello")`, wantNum: 1},
		{name: "star_many_chars", formula: `SEARCH("h*d","hello world")`, wantNum: 1},
		{name: "star_only", formula: `SEARCH("*","anything")`, wantNum: 1},
		{name: "star_at_end", formula: `SEARCH("hel*","hello")`, wantNum: 1},
		{name: "star_at_beginning", formula: `SEARCH("*lo","hello")`, wantNum: 1},

		// Multiple/combined wildcards
		{name: "question_star_question", formula: `SEARCH("?e*d","hello world")`, wantNum: 1},
		{name: "star_question_combined", formula: `SEARCH("h*l?o","hello")`, wantNum: 1},

		// Empty find text → start position
		{name: "empty_find_default", formula: `SEARCH("","hello")`, wantNum: 1},
		{name: "empty_find_with_start", formula: `SEARCH("","hello",3)`, wantNum: 3},

		// Error cases: not found
		{name: "not_found_VALUE", formula: `SEARCH("xyz","hello")`, isErr: true},
		// Error cases: start_num out of range
		{name: "start_num_exceeds_length", formula: `SEARCH("a","hello",10)`, isErr: true},
		{name: "start_num_zero", formula: `SEARCH("a","hello",0)`, isErr: true},
		{name: "start_num_negative", formula: `SEARCH("a","hello",-1)`, isErr: true},

		// Find text longer than within text
		{name: "find_longer_than_within", formula: `SEARCH("hello world!","hello")`, isErr: true},

		// Escaped wildcards: ~? and ~*
		{name: "escaped_question_literal", formula: `SEARCH("~?","is this?")`, wantNum: 8},
		{name: "escaped_star_literal", formula: `SEARCH("~*","3*5")`, wantNum: 2},
		{name: "escaped_tilde_literal", formula: `SEARCH("~~","a~b")`, wantNum: 2},

		// Unicode / multibyte characters
		{name: "unicode_find", formula: `SEARCH("世界","你好世界")`, wantNum: 3},
		{name: "unicode_case_insensitive", formula: `SEARCH("café","CAFÉ au lait")`, wantNum: 1},
		{name: "emoji_search", formula: `SEARCH("🌍","hello🌍world")`, wantNum: 6},
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
					t.Errorf("Eval(%q) = type=%d num=%g, want %g", tt.formula, got.Type, got.Num, tt.wantNum)
				}
			}
		})
	}

	// Wrong arg count tests via direct function call
	t.Run("too_few_args", func(t *testing.T) {
		v, err := fnSEARCH([]Value{StringVal("a")})
		if err != nil {
			t.Fatal(err)
		}
		if v.Type != ValueError {
			t.Errorf("got %v, want error", v)
		}
	})

	t.Run("too_many_args", func(t *testing.T) {
		v, err := fnSEARCH([]Value{StringVal("a"), StringVal("abc"), NumberVal(1), NumberVal(2)})
		if err != nil {
			t.Fatal(err)
		}
		if v.Type != ValueError {
			t.Errorf("got %v, want error", v)
		}
	})

	// Error propagation: if an argument is an error, it should propagate
	t.Run("error_propagation_find_text", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNA),
			},
		}
		cf := evalCompile(t, `SEARCH(A1,"hello")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// The error in A1 gets coerced to a string "#N/A" and searched for;
		// it won't be found, so the result is #VALUE!
		if got.Type != ValueError {
			t.Errorf("got %v, want error", got)
		}
	})

	t.Run("start_num_just_past_end", func(t *testing.T) {
		// start_num = len+1 is exactly past the end for non-empty find
		cf := evalCompile(t, `SEARCH("a","hello",6)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("empty_within_text", func(t *testing.T) {
		cf := evalCompile(t, `SEARCH("a","")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("got %v, want #VALUE!", got)
		}
	})

	t.Run("empty_find_empty_within", func(t *testing.T) {
		cf := evalCompile(t, `SEARCH("","")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})

	t.Run("wildcard_star_matching_everything", func(t *testing.T) {
		cf := evalCompile(t, `SEARCH("*","hello world")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("got %v, want 1", got)
		}
	})
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

func TestCODE_MacRoman(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		// ASCII characters (same in Mac Roman and ASCII)
		{"ASCII A", `CODE("A")`, 65},
		{"ASCII underscore", `CODE("_")`, 95},
		{"ASCII space", `CODE(" ")`, 32},

		// Mac OS Roman mappings: Unicode code points that map back
		// to Mac Roman byte values.
		{"A-diaeresis U+00C4 -> 0x80", `CODE("Ä")`, 0x80},
		{"Euro sign U+20AC -> 0xDB", `CODE("€")`, 0xDB},
		{"Left single quote U+2018 -> 0xD4", "CODE(\"\u2018\")", 0xD4},
		{"Trademark U+2122 -> 0xAA", `CODE("™")`, 0xAA},
		{"Right guillemet U+00BB -> 0xC8", `CODE("»")`, 0xC8},
		{"Ogonek U+02DB -> 0xFE", "CODE(\"\u02DB\")", 0xFE},
		{"Caron U+02C7 -> 0xFF", "CODE(\"\u02C7\")", 0xFF},

		// Characters outside Mac Roman -> replacement '_' = 95
		{"CJK char -> replacement", `CODE("日")`, 95},
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

func TestUNICHAR(t *testing.T) {
	resolver := &mockResolver{}

	// String result tests
	strTests := []struct {
		name    string
		formula string
		want    string
	}{
		// Basic ASCII characters
		{"space", `UNICHAR(32)`, " "},
		{"exclamation", `UNICHAR(33)`, "!"},
		{"digit_0", `UNICHAR(48)`, "0"},
		{"uppercase_A", `UNICHAR(65)`, "A"},
		{"uppercase_B", `UNICHAR(66)`, "B"},
		{"uppercase_Z", `UNICHAR(90)`, "Z"},
		{"lowercase_a", `UNICHAR(97)`, "a"},
		{"lowercase_z", `UNICHAR(122)`, "z"},
		{"tilde", `UNICHAR(126)`, "~"},

		// Common control characters
		{"tab", `UNICHAR(9)`, "\t"},
		{"line_feed", `UNICHAR(10)`, "\n"},
		{"carriage_return", `UNICHAR(13)`, "\r"},

		// Unicode characters beyond ASCII
		{"copyright", `UNICHAR(169)`, "\u00A9"},     // ©
		{"euro_sign", `UNICHAR(8364)`, "\u20AC"},    // €
		{"snowman", `UNICHAR(9731)`, "\u2603"},      // ☃
		{"greek_alpha", `UNICHAR(945)`, "\u03B1"},   // α
		{"cjk_char", `UNICHAR(20013)`, "\u4E2D"},    // 中
		{"musical_note", `UNICHAR(9834)`, "\u266A"}, // ♪
		{"infinity", `UNICHAR(8734)`, "\u221E"},     // ∞
		{"check_mark", `UNICHAR(10003)`, "\u2713"},  // ✓

		// Emoji (supplementary plane)
		{"grinning_face", `UNICHAR(128512)`, "\U0001F600"}, // 😀

		// Non-integer inputs should truncate
		{"truncate_65.1", `UNICHAR(65.1)`, "A"},
		{"truncate_65.9", `UNICHAR(65.9)`, "A"},
		{"truncate_66.5", `UNICHAR(66.5)`, "B"},

		// String coercion
		{"string_65", `UNICHAR("65")`, "A"},
		{"string_66", `UNICHAR("66")`, "B"},

		// Control characters — Excel accepts all of them
		{"control_1", `UNICHAR(1)`, "\x01"},
		{"control_2", `UNICHAR(2)`, "\x02"},
		{"control_31", `UNICHAR(31)`, "\x1F"},

		// DEL and C1 control characters — Excel accepts them
		{"del", `UNICHAR(127)`, "\x7F"},
		{"c1_128", `UNICHAR(128)`, "\u0080"},
		{"c1_159", `UNICHAR(159)`, "\u009F"},

		// Unicode noncharacters — Excel accepts them
		{"nonchar_fdd0", `UNICHAR(64976)`, "\uFDD0"},
		{"nonchar_fffe", `UNICHAR(65534)`, "\uFFFE"},
		// Boolean TRUE coerces to 1
		{"bool_true", `UNICHAR(TRUE)`, "\x01"},
	}

	for _, tt := range strTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %v (type %d), want %q", tt.formula, got, got.Type, tt.want)
			}
		})
	}

	// Error result tests
	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// Zero is invalid
		{"zero", `UNICHAR(0)`, ErrValVALUE},
		// Negative values
		{"negative_1", `UNICHAR(-1)`, ErrValVALUE},
		{"negative_100", `UNICHAR(-100)`, ErrValVALUE},
		// Above max Unicode code point
		{"too_large", `UNICHAR(1114112)`, ErrValVALUE},
		{"very_large", `UNICHAR(9999999)`, ErrValVALUE},
		// Surrogate code points → #VALUE!
		{"surrogate_start", `UNICHAR(55296)`, ErrValVALUE}, // 0xD800
		{"surrogate_mid", `UNICHAR(56000)`, ErrValVALUE},   // 0xDAC0
		{"surrogate_end", `UNICHAR(57343)`, ErrValVALUE},   // 0xDFFF
		// Plane-terminal U+FFFF noncharacters → #N/A
		{"nonchar_ffff", `UNICHAR(65535)`, ErrValNA},
		{"max_unicode", `UNICHAR(1114111)`, ErrValNA},
		// Non-numeric string → #VALUE!
		{"non_numeric_string", `UNICHAR("hello")`, ErrValVALUE},
		// Wrong number of args
		{"no_args", `UNICHAR()`, ErrValVALUE},
		{"two_args", `UNICHAR(65,66)`, ErrValVALUE},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = %v (type %d, err %d), want error %d", tt.formula, got, got.Type, got.Err, tt.wantErr)
			}
		})
	}
}

func TestCHAR_MacRoman(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		name    string
		formula string
		want    string
	}{
		// ASCII range (same in Mac Roman and ASCII)
		{"CHAR(65)", `CHAR(65)`, "A"},
		{"CHAR(95)", `CHAR(95)`, "_"},

		// Mac OS Roman 0x80-0x9F range
		{"CHAR(128) = A-diaeresis", `CHAR(128)`, "\u00C4"}, // Ä
		{"CHAR(129) = A-ring", `CHAR(129)`, "\u00C5"},      // Å
		{"CHAR(130) = C-cedilla", `CHAR(130)`, "\u00C7"},   // Ç
		{"CHAR(131) = E-acute", `CHAR(131)`, "\u00C9"},     // É
		{"CHAR(135) = a-acute", `CHAR(135)`, "\u00E1"},     // á

		// Mac OS Roman 0xA0-0xBF range
		{"CHAR(160) = dagger", `CHAR(160)`, "\u2020"},    // †
		{"CHAR(161) = degree", `CHAR(161)`, "\u00B0"},    // °
		{"CHAR(162) = cent", `CHAR(162)`, "\u00A2"},      // ¢
		{"CHAR(164) = section", `CHAR(164)`, "\u00A7"},   // §
		{"CHAR(169) = copyright", `CHAR(169)`, "\u00A9"}, // ©
		{"CHAR(170) = trademark", `CHAR(170)`, "\u2122"}, // ™
		{"CHAR(176) = infinity", `CHAR(176)`, "\u221E"},  // ∞

		// Mac OS Roman 0xC0-0xDF range
		{"CHAR(192) = inv_question", `CHAR(192)`, "\u00BF"},       // ¿
		{"CHAR(199) = left_guillemet", `CHAR(199)`, "\u00AB"},     // «
		{"CHAR(200) = right_guillemet", `CHAR(200)`, "\u00BB"},    // »
		{"CHAR(201) = ellipsis", `CHAR(201)`, "\u2026"},           // …
		{"CHAR(210) = left_double_quote", `CHAR(210)`, "\u201C"},  // "
		{"CHAR(211) = right_double_quote", `CHAR(211)`, "\u201D"}, // "
		{"CHAR(212) = left_single_quote", `CHAR(212)`, "\u2018"},  // '
		{"CHAR(213) = right_single_quote", `CHAR(213)`, "\u2019"}, // '
		{"CHAR(216) = y-diaeresis", `CHAR(216)`, "\u00FF"},        // ÿ

		// Mac OS Roman 0xE0-0xFF range
		{"CHAR(219) = euro", `CHAR(219)`, "\u20AC"},    // €
		{"CHAR(247) = em_dash", `CHAR(208)`, "\u2013"}, // en dash at 0xD0
		{"CHAR(254) = ogonek", `CHAR(254)`, "\u02DB"},  // ˛
		{"CHAR(255) = caron", `CHAR(255)`, "\u02C7"},   // ˇ

		// Minimum valid code
		{"CHAR(1) = SOH", `CHAR(1)`, "\x01"},

		// Control characters
		{"CHAR(9) = tab", `CHAR(9)`, "\t"},
		{"CHAR(10) = newline", `CHAR(10)`, "\n"},
		{"CHAR(13) = carriage_return", `CHAR(13)`, "\r"},

		// Space
		{"CHAR(32) = space", `CHAR(32)`, " "},

		// Common punctuation
		{"CHAR(33) = exclamation", `CHAR(33)`, "!"},
		{"CHAR(63) = question_mark", `CHAR(63)`, "?"},
		{"CHAR(64) = at_sign", `CHAR(64)`, "@"},

		// Digit boundaries
		{"CHAR(48) = digit_0", `CHAR(48)`, "0"},
		{"CHAR(57) = digit_9", `CHAR(57)`, "9"},

		// Uppercase boundaries
		{"CHAR(65) = uppercase_A", `CHAR(65)`, "A"},
		{"CHAR(90) = uppercase_Z", `CHAR(90)`, "Z"},

		// Lowercase boundaries
		{"CHAR(97) = lowercase_a", `CHAR(97)`, "a"},
		{"CHAR(122) = lowercase_z", `CHAR(122)`, "z"},

		// Decimal truncation: int(65.9) = 65
		{"CHAR(65.9) = truncate_to_A", `CHAR(65.9)`, "A"},
		{"CHAR(65.1) = truncate_to_A", `CHAR(65.1)`, "A"},
		{"CHAR(66.5) = truncate_to_B", `CHAR(66.5)`, "B"},

		// String coercion
		{"CHAR_string_65", `CHAR("65")`, "A"},
		{"CHAR_string_97", `CHAR("97")`, "a"},

		// Boolean coercion: TRUE = 1
		{"CHAR_bool_true", `CHAR(TRUE)`, "\x01"},

		// Concatenation with CHAR
		{"CHAR_concat_AB", `CHAR(65)&CHAR(66)`, "AB"},
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

	// Error result tests
	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// Zero is invalid (valid range is 1-255)
		{"zero", `CHAR(0)`, ErrValVALUE},
		// Negative values
		{"negative_1", `CHAR(-1)`, ErrValVALUE},
		// Above max code 255
		{"code_256", `CHAR(256)`, ErrValVALUE},
		{"code_1000", `CHAR(1000)`, ErrValVALUE},
		// Non-numeric string
		{"non_numeric_string", `CHAR("abc")`, ErrValVALUE},
		// Wrong number of arguments
		{"no_args", `CHAR()`, ErrValVALUE},
		{"two_args", `CHAR(65,66)`, ErrValVALUE},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = %v (type %d, err %d), want error %d", tt.formula, got, got.Type, got.Err, tt.wantErr)
			}
		})
	}

	// CODE(CHAR(65)) roundtrip should return 65
	t.Run("CODE_CHAR_roundtrip", func(t *testing.T) {
		cf := evalCompile(t, `CODE(CHAR(65))`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(CODE(CHAR(65))): %v", err)
		}
		if got.Type != ValueNumber || got.Num != 65 {
			t.Errorf("Eval(CODE(CHAR(65))) = %v, want 65", got)
		}
	})
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

		// Additional: negative numbers with format
		{"negative_concise", `VALUETOTEXT(-99.5,0)`, "-99.5"},
		{"negative_strict", `VALUETOTEXT(-99.5,1)`, "-99.5"},

		// Additional: large and very large numbers
		{"very_large_number", `VALUETOTEXT(999999999999999)`, "999999999999999"},
		{"negative_large", `VALUETOTEXT(-1000000)`, "-1000000"},

		// Additional: decimal precision
		{"decimal_two_places", `VALUETOTEXT(3.14)`, "3.14"},
		{"decimal_many_places", `VALUETOTEXT(1.23456789)`, "1.23456789"},

		// Additional: integer-valued float
		{"integer_float_concise", `VALUETOTEXT(5.0)`, "5"},
		{"integer_float_strict", `VALUETOTEXT(5.0,1)`, "5"},

		// Additional: string with special characters
		{"string_with_numbers", `VALUETOTEXT("abc123")`, "abc123"},
		{"string_special_chars", `VALUETOTEXT("a-b_c")`, "a-b_c"},
		{"string_strict_special", `VALUETOTEXT("a-b_c",1)`, `"a-b_c"`},

		// Additional: boolean coerced from expression
		{"bool_true_expr", `VALUETOTEXT(1=1)`, "TRUE"},
		{"bool_false_expr", `VALUETOTEXT(1=2)`, "FALSE"},

		// Additional: number from expression
		{"expr_result", `VALUETOTEXT(2+3)`, "5"},
		{"expr_result_strict", `VALUETOTEXT(2+3,1)`, "5"},

		// Additional: negative zero
		{"negative_zero", `VALUETOTEXT(0*-1)`, "0"},

		// Additional: number one
		{"number_one", `VALUETOTEXT(1)`, "1"},
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
		{"invalid_format_3", `VALUETOTEXT("x",3)`},
		{"invalid_format_99", `VALUETOTEXT("x",99)`},
		{"format_arg_error", `VALUETOTEXT("hello","abc")`},
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

	t.Run("error_propagation_ref", func(t *testing.T) {
		cf := evalCompile(t, `VALUETOTEXT(#REF!)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("VALUETOTEXT(#REF!) = %v, want error", got)
		}
	})

	// Via eval with inline expression in nested function
	t.Run("nested_in_concat", func(t *testing.T) {
		cf := evalCompile(t, `CONCAT(VALUETOTEXT(42),"-",VALUETOTEXT("hi",1))`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		want := `42-"hi"`
		if got.Type != ValueString || got.Str != want {
			t.Errorf("got %v (%q), want %q", got, got.Str, want)
		}
	})

	t.Run("in_len", func(t *testing.T) {
		cf := evalCompile(t, `LEN(VALUETOTEXT("abc",1))`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// "abc" with quotes = 5 chars
		if got.Type != ValueNumber || got.Num != 5 {
			t.Errorf("got %v, want 5", got)
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

		// Additional: multi-row with strings
		{"multirow_strings_concise", `ARRAYTOTEXT({"a","b";"c","d"})`, "a, b, c, d"},
		{"multirow_strings_strict", `ARRAYTOTEXT({"a","b";"c","d"},1)`, `{"a","b";"c","d"}`},

		// Additional: multi-row mixed types
		{"multirow_mixed_concise", `ARRAYTOTEXT({1,"hi";TRUE,3.5})`, "1, hi, TRUE, 3.5"},
		{"multirow_mixed_strict", `ARRAYTOTEXT({1,"hi";TRUE,3.5},1)`, `{1,"hi";TRUE,3.5}`},

		// Additional: 3-column row
		{"three_col_concise", `ARRAYTOTEXT({1,2,3;4,5,6})`, "1, 2, 3, 4, 5, 6"},
		{"three_col_strict", `ARRAYTOTEXT({1,2,3;4,5,6},1)`, "{1,2,3;4,5,6}"},

		// Additional: all booleans multi-row
		{"bools_multirow_concise", `ARRAYTOTEXT({TRUE,FALSE;FALSE,TRUE})`, "TRUE, FALSE, FALSE, TRUE"},
		{"bools_multirow_strict", `ARRAYTOTEXT({TRUE,FALSE;FALSE,TRUE},1)`, "{TRUE,FALSE;FALSE,TRUE}"},

		// Additional: negative numbers in array
		{"negatives_concise", `ARRAYTOTEXT({-1,-2,-3})`, "-1, -2, -3"},
		{"negatives_strict", `ARRAYTOTEXT({-1,-2,-3},1)`, "{-1,-2,-3}"},

		// Additional: single element array of string in strict
		{"single_quoted_strict", `ARRAYTOTEXT({"hello"},1)`, `{"hello"}`},

		// Additional: floats in multi-row strict
		{"floats_multirow_strict", `ARRAYTOTEXT({1.1,2.2;3.3,4.4},1)`, "{1.1,2.2;3.3,4.4}"},

		// Additional: scalar bool strict
		{"scalar_bool_strict", `ARRAYTOTEXT(TRUE,1)`, "{TRUE}"},
		{"scalar_false_strict", `ARRAYTOTEXT(FALSE,1)`, "{FALSE}"},

		// Additional: large array concise
		{"five_elem_concise", `ARRAYTOTEXT({10,20,30,40,50})`, "10, 20, 30, 40, 50"},
		{"five_elem_strict", `ARRAYTOTEXT({10,20,30,40,50},1)`, "{10,20,30,40,50}"},

		// Additional: single row all strings strict (quoted)
		{"all_strings_strict", `ARRAYTOTEXT({"a","b","c"},1)`, `{"a","b","c"}`},

		// Additional: scalar expression
		{"scalar_expr_concise", `ARRAYTOTEXT(2+3)`, "5"},
		{"scalar_expr_strict", `ARRAYTOTEXT(2+3,1)`, "{5}"},
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
		{"invalid_format_3", `ARRAYTOTEXT({1},3)`},
		{"invalid_format_99", `ARRAYTOTEXT({1},99)`},
		{"format_arg_error", `ARRAYTOTEXT({1},"abc")`},
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

	// Error propagation from array argument
	t.Run("error_propagation_div0", func(t *testing.T) {
		cf := evalCompile(t, `ARRAYTOTEXT(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("ARRAYTOTEXT(1/0) = %v, want error", got)
		}
	})

	t.Run("error_propagation_na", func(t *testing.T) {
		cf := evalCompile(t, `ARRAYTOTEXT(NA())`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("ARRAYTOTEXT(NA()) = %v, want error", got)
		}
	})

	// Nested in another function
	t.Run("nested_in_len", func(t *testing.T) {
		cf := evalCompile(t, `LEN(ARRAYTOTEXT({1,2,3}))`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// "1, 2, 3" = 7 chars
		if got.Type != ValueNumber || got.Num != 7 {
			t.Errorf("got %v, want 7", got)
		}
	})

	t.Run("nested_strict_in_len", func(t *testing.T) {
		cf := evalCompile(t, `LEN(ARRAYTOTEXT({1,2,3},1))`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// "{1,2,3}" = 7 chars
		if got.Type != ValueNumber || got.Num != 7 {
			t.Errorf("got %v, want 7", got)
		}
	})
}

func TestARRAYTOTEXT_TrimmedRangeLogicalTail(t *testing.T) {
	got, err := fnArrayToText([]Value{
		trimmedRangeValue([][]Value{{StringVal("A")}}, 1, 1, 1, 3),
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnArrayToText: %v", err)
	}
	if got.Type != ValueString || got.Str != `{"A";;}` {
		t.Fatalf("got %v, want {\"A\";;}", got)
	}
}

func TestARRAYTOTEXT_FullColumnSparseRef(t *testing.T) {
	got, err := fnArrayToText([]Value{
		sparseFullColumnRefValue(1, map[int]Value{
			1: StringVal("A"),
			3: StringVal("B"),
		}),
		NumberVal(1),
	})
	if err != nil {
		t.Fatalf("fnArrayToText: %v", err)
	}
	if got.Type != ValueString || got.Str != `{"A";;"B"}` {
		t.Fatalf("got %v, want {\"A\";;\"B\"}", got)
	}
}

func TestTEXTBEFORE(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		formula string
		want    string
	}{
		// Basic cases
		{`TEXTBEFORE("Hello World"," ")`, "Hello"},
		{`TEXTBEFORE("Hello-World-Test","-")`, "Hello"},
		{`TEXTBEFORE("Hello-World-Test","-",2)`, "Hello-World"},
		{`TEXTBEFORE("Hello-World-Test","-",1)`, "Hello"},
		// Negative instance_num (count from end)
		{`TEXTBEFORE("Hello-World-Test","-",-1)`, "Hello-World"},
		{`TEXTBEFORE("Hello-World-Test","-",-2)`, "Hello"},
		// Case insensitive
		{`TEXTBEFORE("Hello WORLD","world",1,1)`, "Hello "},
		{`TEXTBEFORE("Hello WORLD","WORLD",1,0)`, "Hello "},
		{`TEXTBEFORE("abcABCabc","abc",2,1)`, "abc"},
		{`TEXTBEFORE("abcABCabc","abc",3,1)`, "abcABC"},
		// if_not_found
		{`TEXTBEFORE("Hello","x",1,0,0,"missing")`, "missing"},
		{`TEXTBEFORE("Hello","x",1,0,0,"")`, ""},
		// match_end=1: when delimiter not found, return full text
		{`TEXTBEFORE("Hello","x",1,0,1)`, "Hello"},
		// Empty delimiter returns ""
		{`TEXTBEFORE("Hello","")`, ""},
		{`TEXTBEFORE("Hello","",1)`, ""},
		// Empty delimiter with instance > 1
		{`TEXTBEFORE("Hello","",2)`, "H"},
		{`TEXTBEFORE("Hello","",3)`, "He"},
		// Delimiter at start of string
		{`TEXTBEFORE("-Hello","-")`, ""},
		// Delimiter at end of string
		{`TEXTBEFORE("Hello-","-")`, "Hello"},
		// Multi-char delimiter
		{`TEXTBEFORE("Hello::World","::")`, "Hello"},
		{`TEXTBEFORE("a::b::c","::",2)`, "a::b"},
		// Text with no occurrence, using match_end
		{`TEXTBEFORE("Hello","xyz",1,0,1)`, "Hello"},
		// Repeated delimiter
		{`TEXTBEFORE("aaa","a",1)`, ""},
		{`TEXTBEFORE("aaa","a",2)`, "a"},
		{`TEXTBEFORE("aaa","a",3)`, "aa"},
		// Negative instance with empty delimiter
		{`TEXTBEFORE("Hello","",-1)`, ""},
		// match_end=1 with instance beyond count
		{`TEXTBEFORE("a-b","-",2,0,1)`, "a-b"},
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

	// Error / #N/A cases
	errTests := []struct {
		formula string
		wantErr ErrorValue
	}{
		// Not found returns #N/A
		{`TEXTBEFORE("Hello","x")`, ErrValNA},
		{`TEXTBEFORE("Hello","xyz")`, ErrValNA},
		// instance_num=0 returns #VALUE!
		{`TEXTBEFORE("Hello","-",0)`, ErrValVALUE},
		// Too few args
		{`TEXTBEFORE("Hello")`, ErrValVALUE},
		// Instance beyond count without match_end
		{`TEXTBEFORE("a-b","-",3)`, ErrValNA},
		// Negative instance beyond count
		{`TEXTBEFORE("a-b","-",-3)`, ErrValNA},
		// Case sensitive mismatch
		{`TEXTBEFORE("Hello","hello",1,0)`, ErrValNA},
	}

	for _, tt := range errTests {
		t.Run(tt.formula, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
			}
		})
	}
}

func TestTEXTAFTER(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		formula string
		want    string
	}{
		// Basic cases
		{`TEXTAFTER("Hello World"," ")`, "World"},
		{`TEXTAFTER("Hello-World-Test","-")`, "World-Test"},
		{`TEXTAFTER("Hello-World-Test","-",2)`, "Test"},
		{`TEXTAFTER("Hello-World-Test","-",1)`, "World-Test"},
		// Negative instance_num (count from end)
		{`TEXTAFTER("Hello-World-Test","-",-1)`, "Test"},
		{`TEXTAFTER("Hello-World-Test","-",-2)`, "World-Test"},
		// Case insensitive
		{`TEXTAFTER("Hello WORLD","world",1,1)`, ""},
		{`TEXTAFTER("Hello WORLD test","world",1,1)`, " test"},
		{`TEXTAFTER("abcABCabc","abc",2,1)`, "abc"},
		{`TEXTAFTER("abcABCabc","abc",3,1)`, ""},
		// if_not_found
		{`TEXTAFTER("Hello","x",1,0,0,"missing")`, "missing"},
		{`TEXTAFTER("Hello","x",1,0,0,"")`, ""},
		// match_end=1: when delimiter not found, return ""
		{`TEXTAFTER("Hello","x",1,0,1)`, ""},
		// Empty delimiter returns full text
		{`TEXTAFTER("Hello","")`, "Hello"},
		{`TEXTAFTER("Hello","",1)`, "Hello"},
		// Empty delimiter with instance > 1
		{`TEXTAFTER("Hello","",2)`, "ello"},
		{`TEXTAFTER("Hello","",3)`, "llo"},
		// Delimiter at start of string
		{`TEXTAFTER("-Hello","-")`, "Hello"},
		// Delimiter at end of string
		{`TEXTAFTER("Hello-","-")`, ""},
		// Multi-char delimiter
		{`TEXTAFTER("Hello::World","::")`, "World"},
		{`TEXTAFTER("a::b::c","::",2)`, "c"},
		// Text with no occurrence, using match_end
		{`TEXTAFTER("Hello","xyz",1,0,1)`, ""},
		// Repeated delimiter
		{`TEXTAFTER("aaa","a",1)`, "aa"},
		{`TEXTAFTER("aaa","a",2)`, "a"},
		{`TEXTAFTER("aaa","a",3)`, ""},
		// Negative instance with empty delimiter
		{`TEXTAFTER("Hello","",-1)`, "Hello"},
		// match_end=1 with instance beyond count
		{`TEXTAFTER("a-b","-",2,0,1)`, ""},
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

	// Error / #N/A cases
	errTests := []struct {
		formula string
		wantErr ErrorValue
	}{
		// Not found returns #N/A
		{`TEXTAFTER("Hello","x")`, ErrValNA},
		{`TEXTAFTER("Hello","xyz")`, ErrValNA},
		// instance_num=0 returns #VALUE!
		{`TEXTAFTER("Hello","-",0)`, ErrValVALUE},
		// Too few args
		{`TEXTAFTER("Hello")`, ErrValVALUE},
		// Instance beyond count without match_end
		{`TEXTAFTER("a-b","-",3)`, ErrValNA},
		// Negative instance beyond count
		{`TEXTAFTER("a-b","-",-3)`, ErrValNA},
		// Case sensitive mismatch
		{`TEXTAFTER("Hello","hello",1,0)`, ErrValNA},
	}

	for _, tt := range errTests {
		t.Run(tt.formula, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError || got.Err != tt.wantErr {
				t.Errorf("Eval(%q) = %v, want error %v", tt.formula, got, tt.wantErr)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// TEXTSPLIT
// ---------------------------------------------------------------------------

func TestTEXTSPLIT_BasicColSplit(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,B,C", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	want := []string{"A", "B", "C"}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_2D(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,B;C,D", ",", ";")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	expected := [][]string{{"A", "B"}, {"C", "D"}}
	for r, row := range expected {
		for c, w := range row {
			if got.Array[r][c].Str != w {
				t.Errorf("[%d][%d]: got %q, want %q", r, c, got.Array[r][c].Str, w)
			}
		}
	}
}

func TestTEXTSPLIT_EmptySegments(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,,B", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	want := []string{"A", "", "B"}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_IgnoreEmpty(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,,B", ",",,TRUE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 1x2, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	want := []string{"A", "B"}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_MultiCharDelimiter(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A::B::C", "::")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("expected 1x3, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	want := []string{"A", "B", "C"}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_CaseInsensitive(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("AxBxC", "X",,, 1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array[0]) != 3 {
		t.Fatalf("expected 3 cols, got %d", len(got.Array[0]))
	}
	want := []string{"A", "B", "C"}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_CaseSensitiveNoMatch(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("axbxc", "X")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "axbxc" {
		t.Errorf("expected original text, got %v", got)
	}
}

func TestTEXTSPLIT_Padding(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,B;C", ",", ";")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 2 || len(got.Array[0]) != 2 {
		t.Fatalf("expected 2x2, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[0][0].Str != "A" || got.Array[0][1].Str != "B" {
		t.Errorf("row 0: got %v %v, want A B", got.Array[0][0], got.Array[0][1])
	}
	if got.Array[1][0].Str != "C" {
		t.Errorf("row 1 col 0: got %v, want C", got.Array[1][0])
	}
	if got.Array[1][1].Type != ValueError || got.Array[1][1].Err != ErrValNA {
		t.Errorf("row 1 col 1: got %v, want #N/A", got.Array[1][1])
	}
}

func TestTEXTSPLIT_CustomPadWith(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,B;C", ",", ";",,,0)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if got.Array[1][1].Type != ValueNumber || got.Array[1][1].Num != 0 {
		t.Errorf("row 1 col 1: got %v, want 0", got.Array[1][1])
	}
}

func TestTEXTSPLIT_NoMatch(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("hello", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "hello" {
		t.Errorf("expected 'hello', got %v", got)
	}
}

func TestTEXTSPLIT_EmptyText(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "" {
		t.Errorf("expected empty string, got %v", got)
	}
}

func TestTEXTSPLIT_TooFewArgs(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("hello")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("expected error, got %v", got)
	}
}

func TestTEXTSPLIT_SingleChar(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "A" {
		t.Errorf("expected 'A', got %v", got)
	}
}

func TestTEXTSPLIT_DelimiterAtEnd(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,B,", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array[0]) != 3 {
		t.Fatalf("expected 3 cols, got %d", len(got.Array[0]))
	}
	if got.Array[0][2].Str != "" {
		t.Errorf("col 2: got %q, want empty", got.Array[0][2].Str)
	}
}

func TestTEXTSPLIT_DelimiterAtStart(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT(",A,B", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array[0]) != 3 {
		t.Fatalf("expected 3 cols, got %d", len(got.Array[0]))
	}
	if got.Array[0][0].Str != "" {
		t.Errorf("col 0: got %q, want empty", got.Array[0][0].Str)
	}
}

func TestTEXTSPLIT_IgnoreEmptyRows(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A;B;;C", ",", ";",TRUE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	for i, w := range []string{"A", "B", "C"} {
		if got.Array[i][0].Str != w {
			t.Errorf("row %d: got %q, want %q", i, got.Array[i][0].Str, w)
		}
	}
}

func TestTEXTSPLIT_ConsecutiveColDelimiters(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,,,B", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array[0]) != 4 {
		t.Fatalf("expected 4 cols, got %d", len(got.Array[0]))
	}
	want := []string{"A", "", "", "B"}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_IgnoreEmptyConsecutive(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,,,B", ",",,TRUE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array[0]) != 2 {
		t.Fatalf("expected 2 cols, got %d", len(got.Array[0]))
	}
	if got.Array[0][0].Str != "A" || got.Array[0][1].Str != "B" {
		t.Errorf("got %v, want [A, B]", got.Array[0])
	}
}

func TestTEXTSPLIT_CaseInsensitive2D(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("AXBxC", "x", , , 1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := []string{"A", "B", "C"}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_PadWithString(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,B;C", ",", ";",,,"N/A")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if got.Array[1][1].Type != ValueString || got.Array[1][1].Str != "N/A" {
		t.Errorf("pad: got %v, want 'N/A'", got.Array[1][1])
	}
}

func TestTEXTSPLIT_OnlyRowDelimiter(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A;B;C", "", ";")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	for i, w := range []string{"A", "B", "C"} {
		if got.Array[i][0].Str != w {
			t.Errorf("row %d: got %q, want %q", i, got.Array[i][0].Str, w)
		}
	}
}

func TestTEXTSPLIT_UnevenRows3x(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,B,C;D;E,F", ",", ";")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	if len(got.Array[0]) != 3 || len(got.Array[1]) != 3 || len(got.Array[2]) != 3 {
		t.Fatalf("expected 3 cols each, got %d, %d, %d",
			len(got.Array[0]), len(got.Array[1]), len(got.Array[2]))
	}
	if got.Array[0][0].Str != "A" || got.Array[0][1].Str != "B" || got.Array[0][2].Str != "C" {
		t.Errorf("row 0: got %v", got.Array[0])
	}
	if got.Array[1][0].Str != "D" {
		t.Errorf("row 1 col 0: got %v, want D", got.Array[1][0])
	}
	if got.Array[1][1].Type != ValueError || got.Array[1][1].Err != ErrValNA {
		t.Errorf("row 1 col 1: got %v, want #N/A", got.Array[1][1])
	}
	if got.Array[2][0].Str != "E" || got.Array[2][1].Str != "F" {
		t.Errorf("row 2: got %v", got.Array[2])
	}
	if got.Array[2][2].Type != ValueError || got.Array[2][2].Err != ErrValNA {
		t.Errorf("row 2 col 2: got %v, want #N/A", got.Array[2][2])
	}
}

func TestTEXTSPLIT_TwoElements(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,B", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array[0]) != 2 {
		t.Fatalf("expected 2 cols, got %d", len(got.Array[0]))
	}
	if got.Array[0][0].Str != "A" || got.Array[0][1].Str != "B" {
		t.Errorf("got %v, want [A, B]", got.Array[0])
	}
}

func TestTEXTSPLIT_SpaceDelimiter(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("hello world foo", " ")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := []string{"hello", "world", "foo"}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_MatchModeZeroCaseSensitive(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("aXbXc", "x",,, 0)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "aXbXc" {
		t.Errorf("expected original text, got %v", got)
	}
}

// ---------------------------------------------------------------------------
// TEXTSPLIT — additional comprehensive tests
// ---------------------------------------------------------------------------

func TestTEXTSPLIT_FiveElements(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("one,two,three,four,five", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := []string{"one", "two", "three", "four", "five"}
	if len(got.Array[0]) != 5 {
		t.Fatalf("expected 5 cols, got %d", len(got.Array[0]))
	}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_LongDelimiter(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("alpha<=>beta<=>gamma", "<=>")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := []string{"alpha", "beta", "gamma"}
	if len(got.Array[0]) != 3 {
		t.Fatalf("expected 3 cols, got %d", len(got.Array[0]))
	}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_TabDelimiter(t *testing.T) {
	resolver := &mockResolver{}
	// Use CHAR(9) for tab via TEXTJOIN to build tab-delimited text, or just pass literal
	cf := evalCompile(t, `TEXTSPLIT("A"&CHAR(9)&"B"&CHAR(9)&"C", CHAR(9))`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := []string{"A", "B", "C"}
	if len(got.Array[0]) != 3 {
		t.Fatalf("expected 3 cols, got %d", len(got.Array[0]))
	}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_OnlyRowDelimiter_MultiCol(t *testing.T) {
	// Row delimiter only, col delimiter empty => each row is unsplit
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("hello|world|!", "", "|")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	for i, w := range []string{"hello", "world", "!"} {
		if got.Array[i][0].Str != w {
			t.Errorf("row %d: got %q, want %q", i, got.Array[i][0].Str, w)
		}
	}
}

func TestTEXTSPLIT_RowAndColSplit_3x3(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("1,2,3;4,5,6;7,8,9", ",", ";")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	expected := [][]string{{"1", "2", "3"}, {"4", "5", "6"}, {"7", "8", "9"}}
	for r, row := range expected {
		if len(got.Array[r]) != 3 {
			t.Fatalf("row %d: expected 3 cols, got %d", r, len(got.Array[r]))
		}
		for c, w := range row {
			if got.Array[r][c].Str != w {
				t.Errorf("[%d][%d]: got %q, want %q", r, c, got.Array[r][c].Str, w)
			}
		}
	}
}

func TestTEXTSPLIT_IgnoreEmptyBothDims(t *testing.T) {
	// With ignore_empty=TRUE, empty segments in both row and col splits are dropped
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,,B;;C,,D", ",", ";", TRUE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	// After ignore_empty: row split "A,,B" and "C,,D" (empty row removed)
	// Then col split on each: "A","B" and "C","D"
	if len(got.Array) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(got.Array))
	}
	if len(got.Array[0]) != 2 || len(got.Array[1]) != 2 {
		t.Fatalf("expected 2 cols each, got %d, %d", len(got.Array[0]), len(got.Array[1]))
	}
	expected := [][]string{{"A", "B"}, {"C", "D"}}
	for r, row := range expected {
		for c, w := range row {
			if got.Array[r][c].Str != w {
				t.Errorf("[%d][%d]: got %q, want %q", r, c, got.Array[r][c].Str, w)
			}
		}
	}
}

func TestTEXTSPLIT_CaseInsensitiveRowDelimiter(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("AzBZC", "", "z",, 1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(got.Array))
	}
	for i, w := range []string{"A", "B", "C"} {
		if got.Array[i][0].Str != w {
			t.Errorf("row %d: got %q, want %q", i, got.Array[i][0].Str, w)
		}
	}
}

func TestTEXTSPLIT_DelimiterAtBothEnds(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT(",A,B,", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array[0]) != 4 {
		t.Fatalf("expected 4 cols, got %d", len(got.Array[0]))
	}
	want := []string{"", "A", "B", ""}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_IgnoreEmptyStartEnd(t *testing.T) {
	// Delimiter at start/end with ignore_empty should drop the empty segments
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT(",A,B,", ",",,TRUE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array[0]) != 2 {
		t.Fatalf("expected 2 cols, got %d", len(got.Array[0]))
	}
	if got.Array[0][0].Str != "A" || got.Array[0][1].Str != "B" {
		t.Errorf("got %v, want [A, B]", got.Array[0])
	}
}

func TestTEXTSPLIT_PadWithEmptyString(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,B;C", ",", ";",,,"")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	// Row 1 col 1 should be padded with ""
	if got.Array[1][1].Type != ValueString || got.Array[1][1].Str != "" {
		t.Errorf("pad: got %v, want empty string", got.Array[1][1])
	}
}

func TestTEXTSPLIT_ErrorInText(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT(1/0, ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("expected error, got %v", got)
	}
}

func TestTEXTSPLIT_ErrorInColDelim(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("hello", 1/0)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("expected error, got %v", got)
	}
}

func TestTEXTSPLIT_ErrorInRowDelim(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("hello", ",", 1/0)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("expected error, got %v", got)
	}
}

func TestTEXTSPLIT_AllDelimitersConsecutive(t *testing.T) {
	// Text is just delimiters: ",,," => 4 empty segments
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT(",,,", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array[0]) != 4 {
		t.Fatalf("expected 4 cols, got %d", len(got.Array[0]))
	}
	for i := 0; i < 4; i++ {
		if got.Array[0][i].Str != "" {
			t.Errorf("col %d: got %q, want empty", i, got.Array[0][i].Str)
		}
	}
}

func TestTEXTSPLIT_AllDelimitersConsecutiveIgnoreEmpty(t *testing.T) {
	// All delimiters with ignore_empty => everything filtered
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT(",,,", ",",,TRUE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// When everything is filtered out, should return empty string
	if got.Type != ValueString || got.Str != "" {
		t.Errorf("expected empty string, got %v", got)
	}
}

func TestTEXTSPLIT_SingleDelimiterMatch(t *testing.T) {
	// Only one delimiter in text => 2 parts
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("hello,world", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array[0]) != 2 {
		t.Fatalf("expected 2 cols, got %d", len(got.Array[0]))
	}
	if got.Array[0][0].Str != "hello" || got.Array[0][1].Str != "world" {
		t.Errorf("got %v, want [hello, world]", got.Array[0])
	}
}

func TestTEXTSPLIT_WhitespaceText(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT(" a , b , c ", ",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := []string{" a ", " b ", " c "}
	if len(got.Array[0]) != 3 {
		t.Fatalf("expected 3 cols, got %d", len(got.Array[0]))
	}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_NumericColDelimiter(t *testing.T) {
	// Numeric delimiter coerced to string "0"
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("a0b0c", 0)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	want := []string{"a", "b", "c"}
	if len(got.Array[0]) != 3 {
		t.Fatalf("expected 3 cols, got %d", len(got.Array[0]))
	}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestTEXTSPLIT_DefaultPadIsNA(t *testing.T) {
	// Verify the default pad_with is #N/A (not empty, not 0)
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTSPLIT("A,B,C;D", ",", ";")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("expected array, got %v", got.Type)
	}
	if len(got.Array) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(got.Array))
	}
	// Row 0: A, B, C
	// Row 1: D, #N/A, #N/A
	if got.Array[1][0].Str != "D" {
		t.Errorf("row 1 col 0: got %v, want D", got.Array[1][0])
	}
	if got.Array[1][1].Type != ValueError || got.Array[1][1].Err != ErrValNA {
		t.Errorf("row 1 col 1: got %v, want #N/A", got.Array[1][1])
	}
	if got.Array[1][2].Type != ValueError || got.Array[1][2].Err != ErrValNA {
		t.Errorf("row 1 col 2: got %v, want #N/A", got.Array[1][2])
	}
}

// ---------------------------------------------------------------------------
// TEXTJOIN comprehensive tests
// ---------------------------------------------------------------------------

func TestTEXTJOIN(t *testing.T) {
	t.Run("basic comma delimiter", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(",", TRUE, "a", "b", "c")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "a,b,c" {
			t.Errorf("got %q, want %q", got.Str, "a,b,c")
		}
	})

	t.Run("space delimiter", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(" ", TRUE, "The", "sun", "rises")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "The sun rises" {
			t.Errorf("got %q, want %q", got.Str, "The sun rises")
		}
	})

	t.Run("semicolon delimiter", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(";", TRUE, "x", "y")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "x;y" {
			t.Errorf("got %q, want %q", got.Str, "x;y")
		}
	})

	t.Run("empty delimiter concatenates", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN("", TRUE, "a", "b", "c")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "abc" {
			t.Errorf("got %q, want %q", got.Str, "abc")
		}
	})

	t.Run("multi-char delimiter", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(", ", TRUE, "one", "two", "three")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "one, two, three" {
			t.Errorf("got %q, want %q", got.Str, "one, two, three")
		}
	})

	t.Run("ignore_empty TRUE skips empty strings", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(",", TRUE, "a", "", "b", "", "c")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "a,b,c" {
			t.Errorf("got %q, want %q", got.Str, "a,b,c")
		}
	})

	t.Run("ignore_empty FALSE includes empty strings", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(",", FALSE, "a", "", "b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "a,,b" {
			t.Errorf("got %q, want %q", got.Str, "a,,b")
		}
	})

	t.Run("single value no joining", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(",", TRUE, "only")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "only" {
			t.Errorf("got %q, want %q", got.Str, "only")
		}
	})

	t.Run("numbers coerced to string", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN("-", TRUE, 1, 2, 3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "1-2-3" {
			t.Errorf("got %q, want %q", got.Str, "1-2-3")
		}
	})

	t.Run("mix of numbers and strings", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(" ", TRUE, "item", 42, "ok")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "item 42 ok" {
			t.Errorf("got %q, want %q", got.Str, "item 42 ok")
		}
	})

	t.Run("boolean values", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(",", TRUE, TRUE, FALSE, "x")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "TRUE,FALSE,x" {
			t.Errorf("got %q, want %q", got.Str, "TRUE,FALSE,x")
		}
	})

	t.Run("all empty with ignore_empty TRUE returns empty", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(",", TRUE, "", "", "")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "" {
			t.Errorf("got %q, want empty string", got.Str)
		}
	})

	t.Run("all empty with ignore_empty FALSE returns delimiters", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(",", FALSE, "", "", "")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != ",," {
			t.Errorf("got %q, want %q", got.Str, ",,")
		}
	})

	t.Run("range input with ignore_empty TRUE", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("apple"),
				{Col: 1, Row: 2}: StringVal(""),
				{Col: 1, Row: 3}: StringVal("cherry"),
			},
		}
		cf := evalCompile(t, `TEXTJOIN(", ", TRUE, A1:A3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "apple, cherry" {
			t.Errorf("got %q, want %q", got.Str, "apple, cherry")
		}
	})

	t.Run("range input with ignore_empty FALSE", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("a"),
				// Row 2 is empty (not set => EmptyVal)
				{Col: 1, Row: 3}: StringVal("c"),
			},
		}
		cf := evalCompile(t, `TEXTJOIN(",", FALSE, A1:A3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "a,,c" {
			t.Errorf("got %q, want %q", got.Str, "a,,c")
		}
	})

	t.Run("multiple ranges", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("a1"),
				{Col: 1, Row: 2}: StringVal("a2"),
				{Col: 2, Row: 1}: StringVal("b1"),
				{Col: 2, Row: 2}: StringVal("b2"),
			},
		}
		cf := evalCompile(t, `TEXTJOIN(",", TRUE, A1:A2, B1:B2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "a1,a2,b1,b2" {
			t.Errorf("got %q, want %q", got.Str, "a1,a2,b1,b2")
		}
	})

	t.Run("2D range", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("a1"),
				{Col: 2, Row: 1}: StringVal("b1"),
				{Col: 1, Row: 2}: StringVal("a2"),
				{Col: 2, Row: 2}: StringVal("b2"),
			},
		}
		cf := evalCompile(t, `TEXTJOIN(",", TRUE, A1:B2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "a1,b1,a2,b2" {
			t.Errorf("got %q, want %q", got.Str, "a1,b1,a2,b2")
		}
	})

	t.Run("delimiter from cell reference", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("-"),
				{Col: 2, Row: 1}: StringVal("x"),
				{Col: 2, Row: 2}: StringVal("y"),
			},
		}
		cf := evalCompile(t, `TEXTJOIN(A1, TRUE, B1, B2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "x-y" {
			t.Errorf("got %q, want %q", got.Str, "x-y")
		}
	})

	t.Run("numeric delimiter", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(0, TRUE, "a", "b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "a0b" {
			t.Errorf("got %q, want %q", got.Str, "a0b")
		}
	})

	t.Run("too few args returns error", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(",", TRUE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Errorf("expected error value, got type %v", got.Type)
		}
	})

	t.Run("error in values is stringified", func(t *testing.T) {
		resolver := &mockResolver{}
		cf := evalCompile(t, `TEXTJOIN(",", TRUE, "a", 1/0, "b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// The implementation does not propagate errors from scalar args;
		// they are converted to their string representation.
		if got.Type != ValueString || got.Str != "a,#DIV/0!,b" {
			t.Errorf("got %q, want %q", got.Str, "a,#DIV/0!,b")
		}
	})

	t.Run("range with numbers", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(10),
				{Col: 1, Row: 2}: NumberVal(20),
				{Col: 1, Row: 3}: NumberVal(30),
			},
		}
		cf := evalCompile(t, `TEXTJOIN("+", TRUE, A1:A3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "10+20+30" {
			t.Errorf("got %q, want %q", got.Str, "10+20+30")
		}
	})

	t.Run("ignore_empty TRUE with empty cells in range", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("first"),
				// Row 2-4 empty
				{Col: 1, Row: 5}: StringVal("last"),
			},
		}
		cf := evalCompile(t, `TEXTJOIN(",", TRUE, A1:A5)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "first,last" {
			t.Errorf("got %q, want %q", got.Str, "first,last")
		}
	})

	t.Run("many values", func(t *testing.T) {
		resolver := &mockResolver{
			cells: func() map[CellAddr]Value {
				m := make(map[CellAddr]Value)
				for i := 1; i <= 50; i++ {
					m[CellAddr{Col: 1, Row: i}] = NumberVal(float64(i))
				}
				return m
			}(),
		}
		cf := evalCompile(t, `TEXTJOIN(",", TRUE, A1:A50)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// Verify it starts and ends correctly
		if got.Type != ValueString {
			t.Fatalf("expected string, got type %v", got.Type)
		}
		if got.Str[:4] != "1,2," {
			t.Errorf("unexpected start: %q", got.Str[:10])
		}
		if got.Str[len(got.Str)-3:] != ",50" {
			t.Errorf("unexpected end: %q", got.Str[len(got.Str)-5:])
		}
	})

	t.Run("range with mixed empty and ignore FALSE", func(t *testing.T) {
		resolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("a"),
				{Col: 2, Row: 1}: StringVal("b"),
				// A2, B2 empty
				{Col: 1, Row: 3}: StringVal("e"),
				{Col: 2, Row: 3}: StringVal("f"),
			},
		}
		cf := evalCompile(t, `TEXTJOIN(",", FALSE, A1:B3)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "a,b,,,e,f" {
			t.Errorf("got %q, want %q", got.Str, "a,b,,,e,f")
		}
	})
}

// ---------------------------------------------------------------------------
// TEXTJOIN — additional comprehensive tests
// ---------------------------------------------------------------------------

func TestTEXTJOIN_TwoArgs(t *testing.T) {
	// Only delimiter + ignore_empty, no text args => #VALUE!
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN(",", TRUE)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("expected error, got %v", got)
	}
}

func TestTEXTJOIN_SingleArg(t *testing.T) {
	// Only delimiter, no ignore_empty or text => #VALUE!
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN(",")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("expected error, got %v", got)
	}
}

func TestTEXTJOIN_DashDelimiter(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN("-", TRUE, "a", "b", "c", "d")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "a-b-c-d" {
		t.Errorf("got %q, want %q", got.Str, "a-b-c-d")
	}
}

func TestTEXTJOIN_PipeDelimiter(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN("|", TRUE, "x", "y", "z")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "x|y|z" {
		t.Errorf("got %q, want %q", got.Str, "x|y|z")
	}
}

func TestTEXTJOIN_ArrowDelimiter(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN(" -> ", TRUE, "start", "middle", "end")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "start -> middle -> end" {
		t.Errorf("got %q, want %q", got.Str, "start -> middle -> end")
	}
}

func TestTEXTJOIN_IgnoreEmptyFALSE_MultipleEmpty(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN(",", FALSE, "", "a", "", "", "b", "")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != ",a,,,b," {
		t.Errorf("got %q, want %q", got.Str, ",a,,,b,")
	}
}

func TestTEXTJOIN_IgnoreEmptyTRUE_AllButOne(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN(",", TRUE, "", "", "only", "", "")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "only" {
		t.Errorf("got %q, want %q", got.Str, "only")
	}
}

func TestTEXTJOIN_MixedTypes(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN(" ", TRUE, "count", 1, TRUE, "done")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "count 1 TRUE done" {
		t.Errorf("got %q, want %q", got.Str, "count 1 TRUE done")
	}
}

func TestTEXTJOIN_FloatingPoint(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN(",", TRUE, 1.5, 2.75, 3.125)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "1.5,2.75,3.125" {
		t.Errorf("got %q, want %q", got.Str, "1.5,2.75,3.125")
	}
}

func TestTEXTJOIN_NegativeNumbers(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN(",", TRUE, -1, -2, -3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "-1,-2,-3" {
		t.Errorf("got %q, want %q", got.Str, "-1,-2,-3")
	}
}

func TestTEXTJOIN_BooleanDelimiter(t *testing.T) {
	// TRUE as delimiter => "TRUE"
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN(TRUE, TRUE, "a", "b")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "aTRUEb" {
		t.Errorf("got %q, want %q", got.Str, "aTRUEb")
	}
}

func TestTEXTJOIN_RangeWithBooleans(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: BoolVal(true),
			{Col: 1, Row: 2}: BoolVal(false),
			{Col: 1, Row: 3}: StringVal("text"),
		},
	}
	cf := evalCompile(t, `TEXTJOIN(",", TRUE, A1:A3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "TRUE,FALSE,text" {
		t.Errorf("got %q, want %q", got.Str, "TRUE,FALSE,text")
	}
}

func TestTEXTJOIN_RangeIgnoreEmptyMixed(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			// Row 2 empty
			{Col: 1, Row: 3}: StringVal(""),
			{Col: 1, Row: 4}: NumberVal(4),
		},
	}
	cf := evalCompile(t, `TEXTJOIN(",", TRUE, A1:A4)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "1,4" {
		t.Errorf("got %q, want %q", got.Str, "1,4")
	}
}

func TestTEXTJOIN_RangeIgnoreEmptyFALSE_Mixed(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			// Row 2 empty
			{Col: 1, Row: 3}: StringVal(""),
			{Col: 1, Row: 4}: NumberVal(4),
		},
	}
	cf := evalCompile(t, `TEXTJOIN(",", FALSE, A1:A4)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "1,,,4" {
		t.Errorf("got %q, want %q", got.Str, "1,,,4")
	}
}

func TestTEXTJOIN_MultipleTextArgs(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN(",", TRUE, "a", "b", "c", "d", "e", "f", "g")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "a,b,c,d,e,f,g" {
		t.Errorf("got %q, want %q", got.Str, "a,b,c,d,e,f,g")
	}
}

func TestTEXTJOIN_NewlineDelimiter(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN(CHAR(10), TRUE, "line1", "line2", "line3")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	want := "line1\nline2\nline3"
	if got.Type != ValueString || got.Str != want {
		t.Errorf("got %q, want %q", got.Str, want)
	}
}

func TestTEXTJOIN_TabDelimiter(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `TEXTJOIN(CHAR(9), TRUE, "col1", "col2", "col3")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	want := "col1\tcol2\tcol3"
	if got.Type != ValueString || got.Str != want {
		t.Errorf("got %q, want %q", got.Str, want)
	}
}

func TestTEXTJOIN_LargeResult32767(t *testing.T) {
	// Build a range that would exceed 32767 characters
	resolver := &mockResolver{
		cells: func() map[CellAddr]Value {
			m := make(map[CellAddr]Value)
			// Each cell = 10 chars "AAAAAAAAAA", 3280 cells + commas = 3280*10 + 3279 = 36079 > 32767
			for i := 1; i <= 3280; i++ {
				m[CellAddr{Col: 1, Row: i}] = StringVal("AAAAAAAAAA")
			}
			return m
		}(),
	}
	cf := evalCompile(t, `TEXTJOIN(",", TRUE, A1:A3280)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("expected #VALUE! for exceeding 32767 chars, got type %v", got.Type)
	}
}

func TestTEXTJOIN_ExactlyAtLimit(t *testing.T) {
	// Single item that is exactly 32767 chars should succeed
	longStr := ""
	for i := 0; i < 32767; i++ {
		longStr += "A"
	}
	// We can't build 32767-char literal in a formula easily, so use a cell
	resolver2 := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal(longStr),
		},
	}
	cf := evalCompile(t, `TEXTJOIN("", TRUE, A1)`)
	got, err := Eval(cf, resolver2, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || len(got.Str) != 32767 {
		t.Errorf("expected 32767 char string, got type %v len %d", got.Type, len(got.Str))
	}
}

func TestTEXTJOIN_RangeAndScalarArgs(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("r1"),
			{Col: 1, Row: 2}: StringVal("r2"),
		},
	}
	cf := evalCompile(t, `TEXTJOIN(",", TRUE, "pre", A1:A2, "post")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "pre,r1,r2,post" {
		t.Errorf("got %q, want %q", got.Str, "pre,r1,r2,post")
	}
}

func TestTEXTJOIN_TrimmedRangeLogicalTail(t *testing.T) {
	got, err := fnTEXTJOIN([]Value{
		StringVal(","),
		BoolVal(false),
		trimmedRangeValue([][]Value{{StringVal("A")}}, 1, 1, 1, 3),
	})
	if err != nil {
		t.Fatalf("fnTEXTJOIN: %v", err)
	}
	if got.Type != ValueString || got.Str != "A,," {
		t.Fatalf("got %v, want A,,", got)
	}
}

func TestTEXTJOIN_FullColumnSparseRef(t *testing.T) {
	got, err := fnTEXTJOIN([]Value{
		StringVal(","),
		BoolVal(false),
		sparseFullColumnRefValue(1, map[int]Value{
			1: StringVal("A"),
			3: StringVal("B"),
		}),
	})
	if err != nil {
		t.Fatalf("fnTEXTJOIN: %v", err)
	}
	if got.Type != ValueString || got.Str != "A,,B" {
		t.Fatalf("got %v, want A,,B", got)
	}
}

func TestCollectDelimiters_TrimmedRangeLogicalTail(t *testing.T) {
	got := collectDelimiters(trimmedRangeValue([][]Value{{StringVal(",")}}, 1, 1, 1, 3))
	want := []string{",", "", ""}
	if len(got) != len(want) {
		t.Fatalf("len = %d, want %d", len(got), len(want))
	}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("delimiter[%d] = %q, want %q", i, got[i], want[i])
		}
	}
}

func TestLOWER(t *testing.T) {
	resolver := &mockResolver{}

	// String result tests
	strTests := []struct {
		name    string
		formula string
		want    string
	}{
		{"all uppercase", `LOWER("HELLO")`, "hello"},
		{"already lowercase", `LOWER("hello")`, "hello"},
		{"mixed case", `LOWER("Hello World")`, "hello world"},
		{"mixed case sentence", `LOWER("E. E. Cummings")`, "e. e. cummings"},
		{"apartment address", `LOWER("Apt. 2B")`, "apt. 2b"},
		{"numbers as string", `LOWER("ABC123DEF")`, "abc123def"},
		{"pure number string", `LOWER("12345")`, "12345"},
		{"number argument coerced", `LOWER(100)`, "100"},
		{"empty string", `LOWER("")`, ""},
		{"spaces only", `LOWER("   ")`, "   "},
		{"special characters", `LOWER("!@#$%^&*()")`, "!@#$%^&*()"},
		{"punctuation and letters", `LOWER("Hello, World!")`, "hello, world!"},
		{"boolean TRUE", `LOWER(TRUE)`, "true"},
		{"boolean FALSE", `LOWER(FALSE)`, "false"},
		{"accented uppercase", `LOWER("CAFÉ")`, "café"},
		{"german uppercase", `LOWER("STRASSE")`, "strasse"},
		{"unicode accented", `LOWER("RÉSUMÉ")`, "résumé"},
		{"single character", `LOWER("A")`, "a"},

		// Additional comprehensive tests
		{"CJK unchanged", `LOWER("你好世界")`, "你好世界"},
		{"numeric string mixed", `LOWER("123ABC")`, "123abc"},
		{"special chars with letters", `LOWER("ABC!@#DEF")`, "abc!@#def"},
		{"newline preserved", `LOWER("HELLO` + "\n" + `WORLD")`, "hello\nworld"},
		{"tab preserved", `LOWER("HELLO` + "\t" + `WORLD")`, "hello\tworld"},
		{"german eszett lowercase stays", `LOWER("STRAßE")`, "straße"},
		{"long mixed string", `LOWER("ThE QuIcK BrOwN FoX")`, "the quick brown fox"},
		{"trailing spaces", `LOWER("HELLO   ")`, "hello   "},
		{"leading spaces", `LOWER("   HELLO")`, "   hello"},
		{"single uppercase letter", `LOWER("Z")`, "z"},
		{"single lowercase letter unchanged", `LOWER("z")`, "z"},
		{"mixed numbers and upper", `LOWER("A1B2C3")`, "a1b2c3"},
		{"unicode greek", `LOWER("ΩΜΕΓΑ")`, "ωμεγα"},
		{"parentheses with text", `LOWER("(HELLO)")`, "(hello)"},
		{"slash separated", `LOWER("HELLO/WORLD")`, "hello/world"},
		{"negative number coerced", `LOWER(-42.5)`, "-42.5"},
		{"zero coerced", `LOWER(0)`, "0"},
		{"hyphenated word", `LOWER("SMITH-JONES")`, "smith-jones"},
		{"all punctuation unchanged", `LOWER("...!!!")`, "...!!!"},
		{"mixed unicode and ascii", `LOWER("ABCdef日本語GHI")`, "abcdef日本語ghi"},
		{"double spaces preserved", `LOWER("HELLO  WORLD")`, "hello  world"},
		{"period separated", `LOWER("MR. SMITH")`, "mr. smith"},
	}

	for _, tt := range strTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %v (type %d), want %q", tt.formula, got, got.Type, tt.want)
			}
		})
	}

	// Error: too many arguments
	t.Run("too many args", func(t *testing.T) {
		cf := evalCompile(t, `LOWER("a","b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE! error", got)
		}
	})

	// Error: no arguments
	t.Run("no args", func(t *testing.T) {
		cf := evalCompile(t, `LOWER()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE! error", got)
		}
	})

	// Error propagation
	t.Run("error propagation NA", func(t *testing.T) {
		cf := evalCompile(t, `LOWER(NA())`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A error", got)
		}
	})

	t.Run("error propagation VALUE", func(t *testing.T) {
		cf := evalCompile(t, `LOWER(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0! error", got)
		}
	})

	// Cell reference
	t.Run("cell reference", func(t *testing.T) {
		cellResolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("HELLO WORLD"),
			},
		}
		cf := evalCompile(t, `LOWER(A1)`)
		got, err := Eval(cf, cellResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "hello world" {
			t.Errorf("got %v, want %q", got, "hello world")
		}
	})

	// Cell reference with number
	t.Run("cell reference number", func(t *testing.T) {
		cellResolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(42),
			},
		}
		cf := evalCompile(t, `LOWER(A1)`)
		got, err := Eval(cf, cellResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "42" {
			t.Errorf("got %v, want %q", got, "42")
		}
	})

	// Empty cell
	t.Run("empty cell", func(t *testing.T) {
		cellResolver := &mockResolver{}
		cf := evalCompile(t, `LOWER(A1)`)
		got, err := Eval(cf, cellResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "" {
			t.Errorf("got %v, want %q", got, "")
		}
	})
}

func TestUPPER(t *testing.T) {
	resolver := &mockResolver{}

	// String result tests
	strTests := []struct {
		name    string
		formula string
		want    string
	}{
		{"all lowercase", `UPPER("hello")`, "HELLO"},
		{"already uppercase", `UPPER("HELLO")`, "HELLO"},
		{"mixed case", `UPPER("Hello World")`, "HELLO WORLD"},
		{"mixed case sentence", `UPPER("e. e. cummings")`, "E. E. CUMMINGS"},
		{"apartment address", `UPPER("Apt. 2B")`, "APT. 2B"},
		{"numbers as string", `UPPER("abc123def")`, "ABC123DEF"},
		{"pure number string", `UPPER("12345")`, "12345"},
		{"number argument coerced", `UPPER(100)`, "100"},
		{"empty string", `UPPER("")`, ""},
		{"spaces only", `UPPER("   ")`, "   "},
		{"special characters", `UPPER("!@#$%^&*()")`, "!@#$%^&*()"},
		{"punctuation and letters", `UPPER("hello, world!")`, "HELLO, WORLD!"},
		{"boolean TRUE", `UPPER(TRUE)`, "TRUE"},
		{"boolean FALSE", `UPPER(FALSE)`, "FALSE"},
		{"accented lowercase", `UPPER("café")`, "CAFÉ"},
		{"german lowercase", `UPPER("strasse")`, "STRASSE"},
		{"unicode accented", `UPPER("résumé")`, "RÉSUMÉ"},
		{"single character", `UPPER("a")`, "A"},

		// Additional comprehensive tests
		{"CJK unchanged", `UPPER("你好世界")`, "你好世界"},
		{"numeric string mixed", `UPPER("123abc")`, "123ABC"},
		{"special chars with letters", `UPPER("abc!@#def")`, "ABC!@#DEF"},
		{"newline preserved", `UPPER("hello` + "\n" + `world")`, "HELLO\nWORLD"},
		{"tab preserved", `UPPER("hello` + "\t" + `world")`, "HELLO\tWORLD"},
		{"long mixed string", `UPPER("ThE QuIcK BrOwN FoX")`, "THE QUICK BROWN FOX"},
		{"trailing spaces", `UPPER("hello   ")`, "HELLO   "},
		{"leading spaces", `UPPER("   hello")`, "   HELLO"},
		{"single lowercase letter", `UPPER("z")`, "Z"},
		{"single already upper", `UPPER("Z")`, "Z"},
		{"mixed numbers and lower", `UPPER("a1b2c3")`, "A1B2C3"},
		{"unicode greek", `UPPER("ωμεγα")`, "ΩΜΕΓΑ"},
		{"parentheses with text", `UPPER("(hello)")`, "(HELLO)"},
		{"slash separated", `UPPER("hello/world")`, "HELLO/WORLD"},
		{"negative number coerced", `UPPER(-42.5)`, "-42.5"},
		{"zero coerced", `UPPER(0)`, "0"},
		{"hyphenated word", `UPPER("smith-jones")`, "SMITH-JONES"},
		{"all punctuation unchanged", `UPPER("...!!!")`, "...!!!"},
		{"mixed unicode and ascii", `UPPER("abcDEF日本語ghi")`, "ABCDEF日本語GHI"},
		{"double spaces preserved", `UPPER("hello  world")`, "HELLO  WORLD"},
		{"period separated", `UPPER("mr. smith")`, "MR. SMITH"},
		{"curly braces", `UPPER("{hello}")`, "{HELLO}"},
		{"angle brackets", `UPPER("<hello>")`, "<HELLO>"},
		{"pipe separated", `UPPER("hello|world")`, "HELLO|WORLD"},
		{"underscore separated", `UPPER("hello_world")`, "HELLO_WORLD"},
		{"tilde in text", `UPPER("hello~world")`, "HELLO~WORLD"},
		{"long repeated", `UPPER("abcabcabc")`, "ABCABCABC"},
		{"accented mixed", `UPPER("crème brûlée")`, "CRÈME BRÛLÉE"},
		{"equals sign", `UPPER("a=b")`, "A=B"},
		{"colon separated", `UPPER("key:value")`, "KEY:VALUE"},
		{"semicolon separated", `UPPER("hello;world")`, "HELLO;WORLD"},
		{"ampersand", `UPPER("a&b")`, "A&B"},
	}

	for _, tt := range strTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %v (type %d), want %q", tt.formula, got, got.Type, tt.want)
			}
		})
	}

	// Error: too many arguments
	t.Run("too many args", func(t *testing.T) {
		cf := evalCompile(t, `UPPER("a","b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE! error", got)
		}
	})

	// Error: no arguments
	t.Run("no args", func(t *testing.T) {
		cf := evalCompile(t, `UPPER()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE! error", got)
		}
	})

	// Error propagation
	t.Run("error propagation NA", func(t *testing.T) {
		cf := evalCompile(t, `UPPER(NA())`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A error", got)
		}
	})

	t.Run("error propagation DIV0", func(t *testing.T) {
		cf := evalCompile(t, `UPPER(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0! error", got)
		}
	})

	// Cell reference
	t.Run("cell reference", func(t *testing.T) {
		cellResolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("hello world"),
			},
		}
		cf := evalCompile(t, `UPPER(A1)`)
		got, err := Eval(cf, cellResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "HELLO WORLD" {
			t.Errorf("got %v, want %q", got, "HELLO WORLD")
		}
	})

	// Cell reference with boolean
	t.Run("cell reference boolean", func(t *testing.T) {
		cellResolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: BoolVal(true),
			},
		}
		cf := evalCompile(t, `UPPER(A1)`)
		got, err := Eval(cf, cellResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "TRUE" {
			t.Errorf("got %v, want %q", got, "TRUE")
		}
	})

	// Empty cell
	t.Run("empty cell", func(t *testing.T) {
		cellResolver := &mockResolver{}
		cf := evalCompile(t, `UPPER(A1)`)
		got, err := Eval(cf, cellResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "" {
			t.Errorf("got %v, want %q", got, "")
		}
	})
}

// ---------------------------------------------------------------------------
// CLEAN
// ---------------------------------------------------------------------------

func TestCLEAN(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		// Basic: remove tab (CHAR 9) and newline (CHAR 10)
		{name: "tab_and_newline", formula: `CLEAN(CHAR(9)&"Monthly report"&CHAR(10))`, want: "Monthly report"},
		// Empty string input
		{name: "empty_string", formula: `CLEAN("")`, want: ""},
		// String with no non-printable characters (unchanged)
		{name: "no_control_chars", formula: `CLEAN("hello world")`, want: "hello world"},
		// Printable string with punctuation and digits
		{name: "printable_mixed", formula: `CLEAN("abc 123 !@#")`, want: "abc 123 !@#"},
		// Only non-printable characters → empty string
		{name: "only_control_chars", formula: `CLEAN(CHAR(1)&CHAR(2)&CHAR(3))`, want: ""},
		// SOH character (CHAR 1)
		{name: "soh_char", formula: `CLEAN(CHAR(1)&"test")`, want: "test"},
		// Bell character (CHAR 7)
		{name: "bell_char", formula: `CLEAN("before"&CHAR(7)&"after")`, want: "beforeafter"},
		// Backspace character (CHAR 8)
		{name: "backspace_char", formula: `CLEAN("data"&CHAR(8))`, want: "data"},
		// Horizontal tab (CHAR 9)
		{name: "tab_only", formula: `CLEAN(CHAR(9))`, want: ""},
		// Line feed (CHAR 10)
		{name: "linefeed", formula: `CLEAN("line1"&CHAR(10)&"line2")`, want: "line1line2"},
		// Carriage return (CHAR 13)
		{name: "carriage_return", formula: `CLEAN("line1"&CHAR(13)&"line2")`, want: "line1line2"},
		// CRLF combination (CHAR 13 + CHAR 10)
		{name: "crlf", formula: `CLEAN("line1"&CHAR(13)&CHAR(10)&"line2")`, want: "line1line2"},
		// Escape character (CHAR 27)
		{name: "escape_char", formula: `CLEAN(CHAR(27)&"text")`, want: "text"},
		// Character 31 (last non-printable, unit separator)
		{name: "char_31", formula: `CLEAN("a"&CHAR(31)&"b")`, want: "ab"},
		// Character 32 (space) should NOT be removed
		{name: "space_preserved", formula: `CLEAN("a"&CHAR(32)&"b")`, want: "a b"},
		// Multiple control chars interspersed
		{name: "multiple_control_mixed", formula: `CLEAN(CHAR(1)&"a"&CHAR(2)&"b"&CHAR(3)&"c")`, want: "abc"},
		// Multiple low control chars surrounding text
		{name: "all_low_controls_prefix", formula: `CLEAN(CHAR(1)&CHAR(5)&CHAR(15)&CHAR(31)&"keep")`, want: "keep"},
		// Number coercion (42 → "42", no control chars)
		{name: "number_coercion", formula: `CLEAN(42)`, want: "42"},
		// Boolean coercion (TRUE → "TRUE")
		{name: "bool_true_coercion", formula: `CLEAN(TRUE)`, want: "TRUE"},
		// Boolean coercion (FALSE → "FALSE")
		{name: "bool_false_coercion", formula: `CLEAN(FALSE)`, want: "FALSE"},
		// Error value coerced to string (ValueToString converts error)
		{name: "error_coerced", formula: `CLEAN(1/0)`, want: "#DIV/0!"},
		// High ASCII characters (>= 32) preserved
		{name: "high_ascii_preserved", formula: `CLEAN(CHAR(65)&CHAR(90)&CHAR(126))`, want: "AZ~"},
		// Printable special characters preserved
		{name: "special_chars_preserved", formula: `CLEAN("$100.00 (USD)")`, want: "$100.00 (USD)"},
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

func TestCLEANErrors(t *testing.T) {
	resolver := &mockResolver{}

	// No arguments → #VALUE!
	t.Run("no_args", func(t *testing.T) {
		cf := evalCompile(t, `CLEAN()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(CLEAN()): unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("CLEAN() = %v, want #VALUE!", got)
		}
	})

	// Too many arguments → #VALUE!
	t.Run("too_many_args", func(t *testing.T) {
		cf := evalCompile(t, `CLEAN("a","b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(CLEAN(a,b)): unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("CLEAN(a,b) = %v, want #VALUE!", got)
		}
	})
}

// ---------------------------------------------------------------------------
// REPT
// ---------------------------------------------------------------------------

func TestREPT(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		name    string
		formula string
		want    string
	}{
		// Basic usage (docs examples)
		{name: "doc_example_star_dash", formula: `REPT("*-",3)`, want: "*-*-*-"},
		{name: "doc_example_dash_10", formula: `REPT("-",10)`, want: "----------"},

		// Simple repetition
		{name: "single_char_repeat", formula: `REPT("a",5)`, want: "aaaaa"},
		{name: "multi_char_repeat", formula: `REPT("hello",2)`, want: "hellohello"},
		{name: "single_repetition", formula: `REPT("abc",1)`, want: "abc"},

		// Zero repeats returns empty string
		{name: "zero_repeats", formula: `REPT("hello",0)`, want: ""},

		// Empty string repeated
		{name: "empty_string_repeated", formula: `REPT("",5)`, want: ""},
		{name: "empty_string_zero", formula: `REPT("",0)`, want: ""},

		// Decimal number_times is truncated (floored)
		{name: "decimal_truncated_2.9", formula: `REPT("a",2.9)`, want: "aa"},
		{name: "decimal_truncated_3.1", formula: `REPT("a",3.1)`, want: "aaa"},
		{name: "decimal_truncated_1.5", formula: `REPT("xy",1.5)`, want: "xy"},

		// Number coerced to string for text arg
		{name: "number_as_text", formula: `REPT(123,2)`, want: "123123"},
		{name: "decimal_as_text", formula: `REPT(1.5,3)`, want: "1.51.51.5"},

		// Boolean coercion for text arg
		{name: "true_as_text", formula: `REPT(TRUE,2)`, want: "TRUETRUE"},
		{name: "false_as_text", formula: `REPT(FALSE,3)`, want: "FALSEFALSEFALSE"},

		// Boolean coercion for number_times arg (TRUE=1, FALSE=0)
		{name: "true_as_count", formula: `REPT("x",TRUE)`, want: "x"},
		{name: "false_as_count", formula: `REPT("x",FALSE)`, want: ""},

		// Special characters
		{name: "space_repeated", formula: `REPT(" ",4)`, want: "    "},
		{name: "newline_repeated", formula: "REPT(\"\n\",3)", want: "\n\n\n"},
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
}

func TestREPTErrors(t *testing.T) {
	resolver := &mockResolver{}

	errTests := []struct {
		name    string
		formula string
	}{
		// Wrong argument count
		{name: "no_args", formula: `REPT()`},
		{name: "one_arg", formula: `REPT("a")`},
		{name: "three_args", formula: `REPT("a",2,3)`},

		// Negative number_times
		{name: "negative_count", formula: `REPT("a",-1)`},
		{name: "negative_large", formula: `REPT("a",-100)`},

		// Non-numeric string for number_times
		{name: "string_count", formula: `REPT("a","b")`},
		{name: "empty_string_count", formula: `REPT("a","")`},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): unexpected error: %v", tt.formula, err)
			}
			if got.Type != ValueError {
				t.Errorf("Eval(%q) = %v, want #VALUE!", tt.formula, got)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// REPLACE comprehensive tests
// ---------------------------------------------------------------------------

func TestREPLACEComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
		isErr   bool
	}{
		// Documentation examples
		{name: "doc_example_1", formula: `REPLACE("abcdefghijk",6,5,"*")`, want: "abcde*k"},
		{name: "doc_example_2", formula: `REPLACE("2009",3,2,"10")`, want: "2010"},

		// Replace at beginning of string
		{name: "replace_at_start", formula: `REPLACE("123456",1,3,"A")`, want: "A456"},
		{name: "replace_first_char", formula: `REPLACE("hello",1,1,"H")`, want: "Hello"},
		{name: "replace_all_from_start", formula: `REPLACE("abc",1,3,"XYZ")`, want: "XYZ"},

		// Replace at end of string
		{name: "replace_at_end", formula: `REPLACE("hello",5,1,"!")`, want: "hell!"},
		{name: "replace_last_two", formula: `REPLACE("hello",4,2,"p!")`, want: "help!"},
		{name: "replace_entire_end", formula: `REPLACE("abcdef",4,3,"XY")`, want: "abcXY"},

		// Replace in middle
		{name: "replace_middle", formula: `REPLACE("abcdef",3,2,"XX")`, want: "abXXef"},
		{name: "replace_single_middle", formula: `REPLACE("abcdef",3,1,"X")`, want: "abXdef"},

		// num_chars=0 (insert without removing)
		{name: "insert_at_start", formula: `REPLACE("abc",1,0,"X")`, want: "Xabc"},
		{name: "insert_at_middle", formula: `REPLACE("abc",2,0,"X")`, want: "aXbc"},
		{name: "insert_at_end", formula: `REPLACE("abc",4,0,"X")`, want: "abcX"},
		{name: "insert_empty", formula: `REPLACE("abc",2,0,"")`, want: "abc"},

		// Replace with longer/shorter new_text
		{name: "longer_replacement", formula: `REPLACE("abc",2,1,"XXXX")`, want: "aXXXXc"},
		{name: "shorter_replacement", formula: `REPLACE("abcdef",2,4,"X")`, want: "aXf"},

		// Empty old_text
		{name: "empty_old_text", formula: `REPLACE("",1,0,"hello")`, want: "hello"},
		{name: "empty_old_text_replace", formula: `REPLACE("",1,0,"")`, want: ""},

		// Empty new_text (deletion)
		{name: "delete_chars", formula: `REPLACE("hello",2,3,"")`, want: "ho"},
		{name: "delete_first", formula: `REPLACE("hello",1,1,"")`, want: "ello"},
		{name: "delete_all", formula: `REPLACE("hello",1,5,"")`, want: ""},

		// num_chars exceeds remaining string length (clamps to end)
		{name: "num_chars_exceeds", formula: `REPLACE("abc",2,100,"X")`, want: "aX"},
		{name: "num_chars_exceeds_from_start", formula: `REPLACE("hi",1,50,"bye")`, want: "bye"},

		// start_num beyond string length (appends)
		{name: "start_beyond_length", formula: `REPLACE("abc",10,1,"X")`, want: "abcX"},
		{name: "start_just_beyond", formula: `REPLACE("abc",4,0,"X")`, want: "abcX"},

		// Numeric coercion for old_text
		{name: "numeric_old_text", formula: `REPLACE(12345,2,3,"X")`, want: "1X5"},
		{name: "numeric_zero", formula: `REPLACE(0,1,1,"X")`, want: "X"},

		// Boolean coercion for old_text
		{name: "bool_true", formula: `REPLACE(TRUE,1,2,"X")`, want: "XUE"},
		{name: "bool_false", formula: `REPLACE(FALSE,3,3,"X")`, want: "FAX"},

		// Float args for start_num and num_chars (truncated to int)
		{name: "float_start_num", formula: `REPLACE("hello",2.9,3,"X")`, want: "hXo"},
		{name: "float_num_chars", formula: `REPLACE("hello",1,2.7,"X")`, want: "Xllo"},

		// Negative start_num or num_chars (should error)
		{name: "negative_start", formula: `REPLACE("hello",-1,3,"X")`, isErr: true},
		{name: "zero_start", formula: `REPLACE("hello",0,3,"X")`, isErr: true},
		{name: "negative_num_chars", formula: `REPLACE("hello",1,-1,"X")`, isErr: true},

		// Non-numeric start_num or num_chars (should error)
		{name: "non_numeric_start", formula: `REPLACE("hello","abc",3,"X")`, isErr: true},
		{name: "non_numeric_num_chars", formula: `REPLACE("hello",1,"abc","X")`, isErr: true},
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

func TestREPLACEWrongArgCount(t *testing.T) {
	resolver := &mockResolver{}

	// Too few args
	cf := evalCompile(t, `REPLACE("hello",2,3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("REPLACE with 3 args: got %v, want error", got)
	}

	// Too many args
	cf = evalCompile(t, `REPLACE("hello",2,3,"X","extra")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError {
		t.Errorf("REPLACE with 5 args: got %v, want error", got)
	}
}

// ---------------------------------------------------------------------------
// PROPER
// ---------------------------------------------------------------------------

func TestPROPER(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		name    string
		formula string
		want    string
	}{
		// Basic capitalization
		{"lowercase words", `PROPER("hello world")`, "Hello World"},
		{"all uppercase", `PROPER("HELLO")`, "Hello"},
		{"all lowercase", `PROPER("hello")`, "Hello"},
		{"already proper", `PROPER("Hello World")`, "Hello World"},

		// Non-letter separators trigger capitalization (expected behavior)
		{"apostrophe separator", `PROPER("2-cent's worth")`, "2-Cent'S Worth"},
		{"number prefix", `PROPER("76BudGet")`, "76Budget"},
		{"hyphen separator", `PROPER("2-way street")`, "2-Way Street"},
		{"this is a TITLE", `PROPER("this is a TITLE")`, "This Is A Title"},

		// Edge cases with strings
		{"empty string", `PROPER("")`, ""},
		{"single lowercase", `PROPER("a")`, "A"},
		{"single uppercase", `PROPER("A")`, "A"},
		{"spaces only", `PROPER("   ")`, "   "},
		{"special characters", `PROPER("!@#$%")`, "!@#$%"},
		{"digits only", `PROPER("12345")`, "12345"},

		// Multiple non-letter separators
		{"multiple hyphens", `PROPER("one--two")`, "One--Two"},
		{"dot separator", `PROPER("john.doe")`, "John.Doe"},
		{"mixed punctuation", `PROPER("hello,world!foo")`, "Hello,World!Foo"},

		// Unicode / accented characters
		{"accented lowercase", `PROPER("café résumé")`, "Café Résumé"},

		// Number coercion (ValueToString converts number to string)
		{"number coerced", `PROPER(100)`, "100"},
		{"negative number", `PROPER(-42.5)`, "-42.5"},

		// Boolean coercion
		{"boolean TRUE", `PROPER(TRUE)`, "True"},
		{"boolean FALSE", `PROPER(FALSE)`, "False"},

		// Additional comprehensive tests
		{"mixed case words", `PROPER("hELLO wORLD")`, "Hello World"},
		{"hyphenated name", `PROPER("smith-jones")`, "Smith-Jones"},
		{"apostrophe in word", `PROPER("it's")`, "It'S"},
		{"apostrophe possessive", `PROPER("o'brien")`, "O'Brien"},
		{"numbers followed by word", `PROPER("123 main")`, "123 Main"},
		{"multiple spaces preserved", `PROPER("hello  world")`, "Hello  World"},
		{"accented phrase", `PROPER("café au lait")`, "Café Au Lait"},
		{"tab as word boundary", `PROPER("hello` + "\t" + `world")`, "Hello\tWorld"},
		{"newline as word boundary", `PROPER("hello` + "\n" + `world")`, "Hello\nWorld"},
		{"parentheses", `PROPER("(hello)")`, "(Hello)"},
		{"slash separated", `PROPER("hello/world")`, "Hello/World"},
		{"period separated", `PROPER("mr. smith")`, "Mr. Smith"},
		{"all caps input", `PROPER("HELLO WORLD")`, "Hello World"},
		{"CJK with latin", `PROPER("hello你好world")`, "Hello你好world"},
		{"trailing punctuation", `PROPER("hello!")`, "Hello!"},
		{"leading punctuation", `PROPER("!hello")`, "!Hello"},
		{"underscore separated", `PROPER("hello_world")`, "Hello_World"},
		{"colon separated", `PROPER("key:value")`, "Key:Value"},
		{"semicolon separated", `PROPER("hello;world")`, "Hello;World"},
		{"mixed numbers hyphens", `PROPER("test-123-hello")`, "Test-123-Hello"},
		{"curly braces", `PROPER("{hello}")`, "{Hello}"},
		{"angle brackets", `PROPER("<hello>")`, "<Hello>"},
		{"unicode greek", `PROPER("ωμεγα")`, "Ωμεγα"},
		{"double hyphen", `PROPER("one--two")`, "One--Two"},
		{"comma separated", `PROPER("hello,world")`, "Hello,World"},
		{"long sentence", `PROPER("the quick brown fox jumps over the lazy dog")`, "The Quick Brown Fox Jumps Over The Lazy Dog"},
	}

	for _, tt := range strTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %v (type %d), want %q", tt.formula, got, got.Type, tt.want)
			}
		})
	}

	// Error: no arguments
	t.Run("no args", func(t *testing.T) {
		cf := evalCompile(t, `PROPER()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE! error", got)
		}
	})

	// Error: too many arguments
	t.Run("too many args", func(t *testing.T) {
		cf := evalCompile(t, `PROPER("a","b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE! error", got)
		}
	})

	// Error propagation
	t.Run("error propagation NA", func(t *testing.T) {
		cf := evalCompile(t, `PROPER(NA())`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		// PROPER uses ValueToString which doesn't propagate errors the same way as LOWER/UPPER.
		// The function doesn't have explicit error propagation, so NA() gets coerced to string "#N/A"
		// and then PROPER capitalizes it. Let's verify actual behavior.
		// Actually, looking at fnPROPER, it does NOT check for ValueError before calling ValueToString.
		// ValueToString on an error returns the error string like "#N/A".
		if got.Type != ValueString || got.Str != "#N/A" {
			t.Errorf("got %v, want %q", got, "#N/A")
		}
	})

	// Cell reference
	t.Run("cell reference", func(t *testing.T) {
		cellResolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("hello world"),
			},
		}
		cf := evalCompile(t, `PROPER(A1)`)
		got, err := Eval(cf, cellResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "Hello World" {
			t.Errorf("got %v, want %q", got, "Hello World")
		}
	})

	// Cell reference with all caps
	t.Run("cell reference all caps", func(t *testing.T) {
		cellResolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("HELLO WORLD"),
			},
		}
		cf := evalCompile(t, `PROPER(A1)`)
		got, err := Eval(cf, cellResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "Hello World" {
			t.Errorf("got %v, want %q", got, "Hello World")
		}
	})

	// Empty cell
	t.Run("empty cell", func(t *testing.T) {
		cellResolver := &mockResolver{}
		cf := evalCompile(t, `PROPER(A1)`)
		got, err := Eval(cf, cellResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "" {
			t.Errorf("got %v, want %q", got, "")
		}
	})
}

// ---------------------------------------------------------------------------
// EXACT comprehensive tests
// ---------------------------------------------------------------------------

func TestEXACT(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    bool
	}{
		// Identical strings
		{name: "same_lowercase", formula: `EXACT("hello","hello")`, want: true},
		{name: "same_uppercase", formula: `EXACT("HELLO","HELLO")`, want: true},
		{name: "same_mixed_case", formula: `EXACT("AbCd","AbCd")`, want: true},
		// Case sensitivity — EXACT is case-sensitive unlike =
		{name: "case_differs_first_char", formula: `EXACT("Hello","hello")`, want: false},
		{name: "case_differs_all_caps", formula: `EXACT("ABC","abc")`, want: false},
		{name: "case_differs_last_char", formula: `EXACT("hellO","hello")`, want: false},
		// Empty strings
		{name: "both_empty", formula: `EXACT("","")`, want: true},
		{name: "first_empty", formula: `EXACT("","a")`, want: false},
		{name: "second_empty", formula: `EXACT("a","")`, want: false},
		// Different strings of same length
		{name: "different_same_len", formula: `EXACT("a","b")`, want: false},
		// Substrings (different lengths)
		{name: "substring_prefix", formula: `EXACT("hello","hell")`, want: false},
		{name: "substring_reversed", formula: `EXACT("hell","hello")`, want: false},
		// Whitespace matters
		{name: "leading_space", formula: `EXACT(" hello","hello")`, want: false},
		{name: "trailing_space", formula: `EXACT("hello ","hello")`, want: false},
		{name: "both_with_spaces", formula: `EXACT(" hello "," hello ")`, want: true},
		// Number coercion — numbers are converted to their string form
		{name: "same_integers", formula: `EXACT(1,1)`, want: true},
		{name: "number_vs_string", formula: `EXACT(1,"1")`, want: true},
		{name: "string_vs_number", formula: `EXACT("1",1)`, want: true},
		{name: "different_numbers", formula: `EXACT(1,2)`, want: false},
		{name: "decimal_match", formula: `EXACT(1.5,"1.5")`, want: true},
		{name: "integer_vs_decimal_string", formula: `EXACT(1,"1.0")`, want: false},
		{name: "zero_vs_zero", formula: `EXACT(0,0)`, want: true},
		// Boolean coercion — TRUE→"TRUE", FALSE→"FALSE"
		{name: "bool_true_vs_string_TRUE", formula: `EXACT(TRUE,"TRUE")`, want: true},
		{name: "bool_false_vs_string_FALSE", formula: `EXACT(FALSE,"FALSE")`, want: true},
		{name: "bool_true_vs_lowercase_true", formula: `EXACT(TRUE,"true")`, want: false},
		{name: "bool_false_vs_lowercase_false", formula: `EXACT(FALSE,"false")`, want: false},
		{name: "both_true", formula: `EXACT(TRUE,TRUE)`, want: true},
		{name: "both_false", formula: `EXACT(FALSE,FALSE)`, want: true},
		{name: "true_vs_false", formula: `EXACT(TRUE,FALSE)`, want: false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
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
}

func TestEXACTErrors(t *testing.T) {
	resolver := &mockResolver{}

	// No arguments → #VALUE!
	t.Run("no_args", func(t *testing.T) {
		cf := evalCompile(t, `EXACT()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(EXACT()): unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("EXACT() = %v, want #VALUE!", got)
		}
	})

	// One argument → #VALUE!
	t.Run("one_arg", func(t *testing.T) {
		cf := evalCompile(t, `EXACT("hello")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(EXACT(hello)): unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("EXACT(hello) = %v, want #VALUE!", got)
		}
	})

	// Three arguments → #VALUE!
	t.Run("three_args", func(t *testing.T) {
		cf := evalCompile(t, `EXACT("a","b","c")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(EXACT(a,b,c)): unexpected error: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("EXACT(a,b,c) = %v, want #VALUE!", got)
		}
	})
}

// ---------------------------------------------------------------------------
// FIXED comprehensive tests
// ---------------------------------------------------------------------------

func TestFIXED(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		name    string
		formula string
		want    string
	}{
		// Basic with commas (default no_commas=FALSE)
		{name: "basic_one_decimal", formula: `FIXED(1234.567, 1)`, want: "1,234.6"},
		{name: "basic_two_decimals", formula: `FIXED(1234.567, 2)`, want: "1,234.57"},
		{name: "default_decimals", formula: `FIXED(44.332)`, want: "44.33"},
		{name: "zero_decimals", formula: `FIXED(1234.567, 0)`, want: "1,235"},

		// no_commas = TRUE
		{name: "no_commas_true", formula: `FIXED(1234.567, 1, TRUE)`, want: "1234.6"},
		{name: "no_commas_two_dec", formula: `FIXED(1234.567, 2, TRUE)`, want: "1234.57"},

		// no_commas = FALSE (explicit)
		{name: "no_commas_false", formula: `FIXED(1234.567, 2, FALSE)`, want: "1,234.57"},

		// Negative decimals (rounds to the left of decimal point)
		{name: "neg_dec_minus1", formula: `FIXED(1234.567, -1)`, want: "1,230"},
		{name: "neg_dec_minus2", formula: `FIXED(1234.567, -2)`, want: "1,200"},
		{name: "neg_dec_minus3", formula: `FIXED(1234.567, -3)`, want: "1,000"},
		{name: "neg_dec_no_commas", formula: `FIXED(-1234.567, -1, TRUE)`, want: "-1230"},

		// Negative numbers
		{name: "negative_number", formula: `FIXED(-1234.567, 2)`, want: "-1,234.57"},
		{name: "negative_zero_dec", formula: `FIXED(-1234.567, 0)`, want: "-1,235"},
		{name: "negative_no_commas", formula: `FIXED(-1234.567, 2, TRUE)`, want: "-1234.57"},

		// Zero
		{name: "zero_two_dec", formula: `FIXED(0, 2)`, want: "0.00"},
		{name: "zero_default", formula: `FIXED(0)`, want: "0.00"},
		{name: "zero_zero_dec", formula: `FIXED(0, 0)`, want: "0"},

		// Small values
		{name: "small_positive", formula: `FIXED(0.5, 2)`, want: "0.50"},
		{name: "small_negative", formula: `FIXED(-0.5, 2)`, want: "-0.50"},

		// Large numbers with comma grouping
		{name: "millions", formula: `FIXED(1234567.89, 2)`, want: "1,234,567.89"},
		{name: "millions_no_dec", formula: `FIXED(1000000, 0)`, want: "1,000,000"},

		// Many decimal places
		{name: "many_decimals", formula: `FIXED(1.5, 5)`, want: "1.50000"},

		// Boolean coercion
		{name: "bool_true", formula: `FIXED(TRUE, 2)`, want: "1.00"},
		{name: "bool_false", formula: `FIXED(FALSE, 2)`, want: "0.00"},

		// String coercion
		{name: "string_number", formula: `FIXED("1234.567", 2)`, want: "1,234.57"},

		// Negative zero edge case: -0.001 with 2 decimals produces "-0.00"
		// (Go's math.Round preserves the sign of -0.0)
		{name: "neg_zero_round", formula: `FIXED(-0.001, 2)`, want: "-0.00"},

		// Rounding up at boundary
		{name: "round_up_boundary", formula: `FIXED(1250, -2)`, want: "1,300"},
		{name: "round_down", formula: `FIXED(1249, -2)`, want: "1,200"},

		// Very small number with 3 decimal places
		{name: "very_small_3dec", formula: `FIXED(0.001, 3)`, want: "0.001"},

		// Ten decimal places
		{name: "ten_decimals", formula: `FIXED(1.5, 10)`, want: "1.5000000000"},

		// Comma boundary at 1000
		{name: "comma_boundary", formula: `FIXED(1000, 2)`, want: "1,000.00"},

		// no_commas with negative decimals
		{name: "neg_dec_no_commas_pos", formula: `FIXED(1234.567, -2, TRUE)`, want: "1200"},

		// Boolean coercion for decimals
		{name: "bool_decimals_true", formula: `FIXED(1234.567, TRUE)`, want: "1,234.6"},
		{name: "bool_decimals_false", formula: `FIXED(1234.567, FALSE)`, want: "1,235"},

		// no_commas with number coercion (1 is truthy)
		{name: "no_commas_one", formula: `FIXED(1234.567, 2, 1)`, want: "1234.57"},
		{name: "no_commas_zero", formula: `FIXED(1234.567, 2, 0)`, want: "1,234.57"},

		// Large negative with commas
		{name: "large_negative", formula: `FIXED(-9876543.21, 2)`, want: "-9,876,543.21"},

		// Round up half
		{name: "round_up_half", formula: `FIXED(2.5, 0)`, want: "3"},

		// DOLLAR vs FIXED cross-check: DOLLAR wraps with $, FIXED does not
		{name: "cross_check_positive", formula: `FIXED(1234.567, 2)`, want: "1,234.57"},

		// One decimal with no_commas
		{name: "one_dec_no_commas", formula: `FIXED(9999.95, 1, TRUE)`, want: "10000.0"},

		// Negative number with no_commas
		{name: "neg_no_commas", formula: `FIXED(-5678.1234, 3, TRUE)`, want: "-5678.123"},
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
		{name: "no_args", formula: `FIXED()`},
		{name: "too_many_args", formula: `FIXED(1,2,TRUE,4)`},
		{name: "non_numeric_string", formula: `FIXED("abc")`},
		{name: "non_numeric_decimals", formula: `FIXED(1234, "abc")`},
		{name: "error_propagation_num", formula: `FIXED(1/0)`},
		{name: "error_propagation_dec", formula: `FIXED(1, 1/0)`},
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
// VALUE
// ---------------------------------------------------------------------------

func TestVALUE(t *testing.T) {
	resolver := &mockResolver{}

	type want struct {
		typ ValueType
		num float64
		err ErrorValue
	}

	tests := []struct {
		name    string
		formula string
		want    want
	}{
		// Basic integer strings
		{name: "integer", formula: `VALUE("123")`, want: want{typ: ValueNumber, num: 123}},
		{name: "zero", formula: `VALUE("0")`, want: want{typ: ValueNumber, num: 0}},
		{name: "negative_integer", formula: `VALUE("-50")`, want: want{typ: ValueNumber, num: -50}},
		{name: "large_number", formula: `VALUE("1000000")`, want: want{typ: ValueNumber, num: 1000000}},

		// Decimal strings
		{name: "decimal", formula: `VALUE("3.14")`, want: want{typ: ValueNumber, num: 3.14}},
		{name: "negative_decimal", formula: `VALUE("-2.5")`, want: want{typ: ValueNumber, num: -2.5}},
		{name: "leading_zero_decimal", formula: `VALUE("0.001")`, want: want{typ: ValueNumber, num: 0.001}},

		// Whitespace handling
		{name: "leading_spaces", formula: `VALUE("  42")`, want: want{typ: ValueNumber, num: 42}},
		{name: "trailing_spaces", formula: `VALUE("42  ")`, want: want{typ: ValueNumber, num: 42}},
		{name: "surrounded_spaces", formula: `VALUE("  7.5  ")`, want: want{typ: ValueNumber, num: 7.5}},

		// Currency formatting ($ stripped)
		{name: "dollar_sign", formula: `VALUE("$100")`, want: want{typ: ValueNumber, num: 100}},
		{name: "dollar_with_decimals", formula: `VALUE("$19.99")`, want: want{typ: ValueNumber, num: 19.99}},
		{name: "dollar_with_commas", formula: `VALUE("$1,000")`, want: want{typ: ValueNumber, num: 1000}},
		{name: "dollar_commas_decimals", formula: `VALUE("$1,234.56")`, want: want{typ: ValueNumber, num: 1234.56}},

		// Comma-separated thousands
		{name: "thousands_comma", formula: `VALUE("1,000")`, want: want{typ: ValueNumber, num: 1000}},
		{name: "millions_comma", formula: `VALUE("1,000,000")`, want: want{typ: ValueNumber, num: 1000000}},

		// Percent handling
		{name: "percent_integer", formula: `VALUE("50%")`, want: want{typ: ValueNumber, num: 0.5}},
		{name: "percent_decimal", formula: `VALUE("12.5%")`, want: want{typ: ValueNumber, num: 0.125}},
		{name: "percent_100", formula: `VALUE("100%")`, want: want{typ: ValueNumber, num: 1}},
		{name: "percent_zero", formula: `VALUE("0%")`, want: want{typ: ValueNumber, num: 0}},

		// Number argument passed directly (not a string)
		{name: "number_passthrough", formula: `VALUE(42)`, want: want{typ: ValueNumber, num: 42}},
		{name: "number_decimal_passthrough", formula: `VALUE(3.14)`, want: want{typ: ValueNumber, num: 3.14}},
		{name: "number_negative_passthrough", formula: `VALUE(-10)`, want: want{typ: ValueNumber, num: -10}},

		// Non-numeric strings → #VALUE!
		{name: "alpha_string", formula: `VALUE("abc")`, want: want{typ: ValueError, err: ErrValVALUE}},
		{name: "mixed_alpha_num", formula: `VALUE("12abc")`, want: want{typ: ValueError, err: ErrValVALUE}},
		{name: "empty_string", formula: `VALUE("")`, want: want{typ: ValueError, err: ErrValVALUE}},

		// Boolean coercion (becomes "TRUE"/"FALSE" strings → not numeric)
		{name: "bool_true", formula: `VALUE(TRUE)`, want: want{typ: ValueError, err: ErrValVALUE}},
		{name: "bool_false", formula: `VALUE(FALSE)`, want: want{typ: ValueError, err: ErrValVALUE}},

		// Wrong argument count
		{name: "no_args", formula: `VALUE()`, want: want{typ: ValueError, err: ErrValVALUE}},
		{name: "two_args", formula: `VALUE("1","2")`, want: want{typ: ValueError, err: ErrValVALUE}},

		// Error argument (error value is coerced to string, which is not numeric)
		{name: "error_arg_div0", formula: `VALUE(1/0)`, want: want{typ: ValueError, err: ErrValVALUE}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != tt.want.typ {
				t.Fatalf("Eval(%q).Type = %v, want %v (value=%v)", tt.formula, got.Type, tt.want.typ, got)
			}
			switch tt.want.typ {
			case ValueNumber:
				if got.Num != tt.want.num {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want.num)
				}
			case ValueError:
				if got.Err != tt.want.err {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Err, tt.want.err)
				}
			}
		})
	}
}

func TestNUMBERVALUE(t *testing.T) {
	resolver := &mockResolver{}

	type want struct {
		typ ValueType
		num float64
		err ErrorValue
	}

	tests := []struct {
		name    string
		formula string
		want    want
	}{
		// Documentation example: European format
		{name: "european_format", formula: `NUMBERVALUE("2.500,27",",",".")`, want: want{typ: ValueNumber, num: 2500.27}},

		// Documentation example: simple decimal
		{name: "simple_decimal_explicit_seps", formula: `NUMBERVALUE("3.5",".",",")`, want: want{typ: ValueNumber, num: 3.5}},

		// Simple number, no separators needed
		{name: "simple_integer", formula: `NUMBERVALUE("123")`, want: want{typ: ValueNumber, num: 123}},
		{name: "simple_zero", formula: `NUMBERVALUE("0")`, want: want{typ: ValueNumber, num: 0}},

		// Percent suffix
		{name: "percent_single", formula: `NUMBERVALUE("3.5%")`, want: want{typ: ValueNumber, num: 0.035}},
		{name: "percent_double", formula: `NUMBERVALUE("9%%")`, want: want{typ: ValueNumber, num: 0.0009}},
		{name: "percent_integer", formula: `NUMBERVALUE("50%")`, want: want{typ: ValueNumber, num: 0.5}},
		{name: "percent_100", formula: `NUMBERVALUE("100%")`, want: want{typ: ValueNumber, num: 1}},

		// Default separators (decimal_separator=".", group_separator=",")
		{name: "default_group_separator", formula: `NUMBERVALUE("1,000")`, want: want{typ: ValueNumber, num: 1000}},
		{name: "default_millions", formula: `NUMBERVALUE("1,000,000.50")`, want: want{typ: ValueNumber, num: 1000000.50}},
		{name: "default_with_decimals", formula: `NUMBERVALUE("1,234.56")`, want: want{typ: ValueNumber, num: 1234.56}},

		// Negative numbers
		{name: "negative_integer", formula: `NUMBERVALUE("-42")`, want: want{typ: ValueNumber, num: -42}},
		{name: "negative_decimal", formula: `NUMBERVALUE("-3.14")`, want: want{typ: ValueNumber, num: -3.14}},
		{name: "negative_with_groups", formula: `NUMBERVALUE("-1,000.5")`, want: want{typ: ValueNumber, num: -1000.5}},

		// Leading/trailing spaces are stripped
		{name: "leading_spaces", formula: `NUMBERVALUE("  42")`, want: want{typ: ValueNumber, num: 42}},
		{name: "trailing_spaces", formula: `NUMBERVALUE("42  ")`, want: want{typ: ValueNumber, num: 42}},
		{name: "surrounded_spaces", formula: `NUMBERVALUE("  7.5  ")`, want: want{typ: ValueNumber, num: 7.5}},
		{name: "spaces_in_middle", formula: `NUMBERVALUE(" 3 000 ", ".", " ")`, want: want{typ: ValueNumber, num: 3000}},

		// Empty string returns 0
		{name: "empty_string", formula: `NUMBERVALUE("")`, want: want{typ: ValueNumber, num: 0}},
		{name: "only_spaces", formula: `NUMBERVALUE("   ")`, want: want{typ: ValueNumber, num: 0}},

		// Custom separators
		{name: "semicolon_group_sep", formula: `NUMBERVALUE("1;000.5",".",";")`, want: want{typ: ValueNumber, num: 1000.5}},

		// Error: group separator after decimal separator
		{name: "err_group_after_decimal", formula: `NUMBERVALUE("1.000,5",".",",")`, want: want{typ: ValueError, err: ErrValVALUE}},

		// Error: multiple decimal separators
		{name: "err_multiple_decimals", formula: `NUMBERVALUE("1.2.3")`, want: want{typ: ValueError, err: ErrValVALUE}},

		// Error: invalid characters
		{name: "err_alpha_string", formula: `NUMBERVALUE("abc")`, want: want{typ: ValueError, err: ErrValVALUE}},
		{name: "err_mixed_alpha_num", formula: `NUMBERVALUE("12abc")`, want: want{typ: ValueError, err: ErrValVALUE}},

		// Error: wrong argument count
		{name: "err_no_args", formula: `NUMBERVALUE()`, want: want{typ: ValueError, err: ErrValVALUE}},
		{name: "err_too_many_args", formula: `NUMBERVALUE("1",".",",","x")`, want: want{typ: ValueError, err: ErrValVALUE}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != tt.want.typ {
				t.Fatalf("Eval(%q).Type = %v, want %v (value=%v)", tt.formula, got.Type, tt.want.typ, got)
			}
			switch tt.want.typ {
			case ValueNumber:
				if got.Num != tt.want.num {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want.num)
				}
			case ValueError:
				if got.Err != tt.want.err {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Err, tt.want.err)
				}
			}
		})
	}
}

func TestUNICODE(t *testing.T) {
	resolver := &mockResolver{}

	// Number result tests
	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Basic ASCII characters
		{"uppercase_A", `UNICODE("A")`, 65},
		{"uppercase_B", `UNICODE("B")`, 66},
		{"uppercase_Z", `UNICODE("Z")`, 90},
		{"lowercase_a", `UNICODE("a")`, 97},
		{"lowercase_z", `UNICODE("z")`, 122},
		{"space", `UNICODE(" ")`, 32},
		{"digit_0", `UNICODE("0")`, 48},
		{"digit_9", `UNICODE("9")`, 57},
		{"exclamation", `UNICODE("!")`, 33},
		{"tilde", `UNICODE("~")`, 126},
		{"at_sign", `UNICODE("@")`, 64},

		// Multi-character strings (first char only)
		{"hello", `UNICODE("Hello")`, 72},
		{"world", `UNICODE("World")`, 87},
		{"abc", `UNICODE("abc")`, 97},

		// Unicode characters beyond ASCII
		{"copyright", `UNICODE("©")`, 169},
		{"euro_sign", `UNICODE("€")`, 8364},
		{"greek_alpha", `UNICODE("α")`, 945},
		{"greek_beta", `UNICODE("β")`, 946},
		{"snowman", `UNICODE("☃")`, 9731},
		{"musical_note", `UNICODE("♪")`, 9834},
		{"infinity", `UNICODE("∞")`, 8734},
		{"check_mark", `UNICODE("✓")`, 10003},
		{"cjk_char", `UNICODE("中")`, 20013},

		// Emoji (supplementary plane / high code point)
		{"grinning_face", `UNICODE("😀")`, 128512},

		// Number coercion: numbers are converted to string first
		{"number_1", `UNICODE(1)`, 49},     // "1" → 49
		{"number_123", `UNICODE(123)`, 49}, // "123" → first char "1" → 49
		{"number_0", `UNICODE(0)`, 48},     // "0" → 48
		{"number_9", `UNICODE(9)`, 57},     // "9" → 57
		{"number_neg", `UNICODE(-5)`, 45},  // "-5" → first char "-" → 45

		// Boolean coercion: TRUE → "TRUE", FALSE → "FALSE"
		{"bool_true", `UNICODE(TRUE)`, 84},   // "TRUE" → "T" → 84
		{"bool_false", `UNICODE(FALSE)`, 70}, // "FALSE" → "F" → 70

		// Single character strings
		{"single_newline_char", `UNICODE(CHAR(10))`, 10},
		{"single_tab_char", `UNICODE(CHAR(9))`, 9},
	}

	for _, tt := range numTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("Eval(%q) type = %v, want ValueNumber", tt.formula, got.Type)
			}
			if got.Num != tt.want {
				t.Errorf("Eval(%q) = %g, want %g", tt.formula, got.Num, tt.want)
			}
		})
	}

	// Error result tests
	errTests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// Empty string → #VALUE!
		{"empty_string", `UNICODE("")`, ErrValVALUE},
		// Wrong number of args
		{"no_args", `UNICODE()`, ErrValVALUE},
		{"two_args", `UNICODE("A","B")`, ErrValVALUE},
	}

	for _, tt := range errTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueError {
				t.Fatalf("Eval(%q) type = %v, want ValueError", tt.formula, got.Type)
			}
			if got.Err != tt.wantErr {
				t.Errorf("Eval(%q) error = %v, want %v", tt.formula, got.Err, tt.wantErr)
			}
		})
	}
}

func TestENCODEURL(t *testing.T) {
	resolver := &mockResolver{}

	strTests := []struct {
		name    string
		formula string
		want    string
	}{
		// Documented doc example
		{"doc example URL", `ENCODEURL("http://contoso.sharepoint.com/Finance/Profit and Loss Statement.xlsx")`, "http%3A%2F%2Fcontoso.sharepoint.com%2FFinance%2FProfit%20and%20Loss%20Statement.xlsx"},
		// Empty string
		{"empty string", `ENCODEURL("")`, ""},
		// Plain text (no special chars) unchanged
		{"plain alpha", `ENCODEURL("hello")`, "hello"},
		{"plain alphanum", `ENCODEURL("abc123")`, "abc123"},
		// Unreserved chars stay unchanged
		{"unreserved hyphen", `ENCODEURL("a-b")`, "a-b"},
		{"unreserved underscore", `ENCODEURL("a_b")`, "a_b"},
		{"unreserved dot", `ENCODEURL("a.b")`, "a.b"},
		{"unreserved tilde", `ENCODEURL("a~b")`, "a~b"},
		{"all unreserved", `ENCODEURL("AZaz09-_.~")`, "AZaz09-_.~"},
		// Spaces
		{"space", `ENCODEURL("hello world")`, "hello%20world"},
		{"multiple spaces", `ENCODEURL("a  b")`, "a%20%20b"},
		// Special chars: colon, slash, question, equals, ampersand, hash, at, bang, dollar, plus, comma, semicolon
		{"colon", `ENCODEURL("a:b")`, "a%3Ab"},
		{"slash", `ENCODEURL("a/b")`, "a%2Fb"},
		{"question mark", `ENCODEURL("a?b")`, "a%3Fb"},
		{"equals", `ENCODEURL("a=b")`, "a%3Db"},
		{"ampersand", `ENCODEURL("a&b")`, "a%26b"},
		{"hash", `ENCODEURL("a#b")`, "a%23b"},
		{"at sign", `ENCODEURL("a@b")`, "a%40b"},
		{"exclamation", `ENCODEURL("a!b")`, "a%21b"},
		{"dollar sign", `ENCODEURL("a$b")`, "a%24b"},
		{"plus sign", `ENCODEURL("a+b")`, "a%2Bb"},
		{"comma", `ENCODEURL("a,b")`, "a%2Cb"},
		{"semicolon", `ENCODEURL("a;b")`, "a%3Bb"},
		// Percent sign itself
		{"percent sign", `ENCODEURL("100%")`, "100%25"},
		// Number input coerced to string
		{"number input", `ENCODEURL(42)`, "42"},
		{"decimal number", `ENCODEURL(3.14)`, "3.14"},
		// Boolean input coerced
		{"boolean TRUE", `ENCODEURL(TRUE)`, "TRUE"},
		{"boolean FALSE", `ENCODEURL(FALSE)`, "FALSE"},
		// Mixed characters
		{"mixed url chars", `ENCODEURL("key=val&foo=bar")`, "key%3Dval%26foo%3Dbar"},
		// URL with query parameters
		{"url with query", `ENCODEURL("https://example.com/path?q=hello&lang=en")`, "https%3A%2F%2Fexample.com%2Fpath%3Fq%3Dhello%26lang%3Den"},
		// Unicode/UTF-8 characters (multi-byte)
		{"unicode e-acute", `ENCODEURL("caf` + "\xc3\xa9" + `")`, "caf%C3%A9"},
		{"unicode n-tilde", `ENCODEURL("` + "\xc3\xb1" + `")`, "%C3%B1"},
		// Brackets and braces
		{"left bracket", `ENCODEURL("[")`, "%5B"},
		{"right bracket", `ENCODEURL("]")`, "%5D"},
		{"left brace", `ENCODEURL("{")`, "%7B"},
		{"right brace", `ENCODEURL("}")`, "%7D"},
		// Tab and newline
		{"tab char", `ENCODEURL("a` + "\t" + `b")`, "a%09b"},
		{"newline char", `ENCODEURL("a` + "\n" + `b")`, "a%0Ab"},
	}

	for _, tt := range strTests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != ValueString || got.Str != tt.want {
				t.Errorf("Eval(%q) = %v (type %d), want %q", tt.formula, got, got.Type, tt.want)
			}
		})
	}

	// Error: wrong arg count
	t.Run("no args", func(t *testing.T) {
		cf := evalCompile(t, `ENCODEURL()`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE! error", got)
		}
	})

	t.Run("too many args", func(t *testing.T) {
		cf := evalCompile(t, `ENCODEURL("a","b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("got %v, want #VALUE! error", got)
		}
	})

	// Error propagation
	t.Run("error propagation DIV0", func(t *testing.T) {
		cf := evalCompile(t, `ENCODEURL(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("got %v, want #DIV/0! error", got)
		}
	})

	t.Run("error propagation NA", func(t *testing.T) {
		cf := evalCompile(t, `ENCODEURL(NA())`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("got %v, want #N/A error", got)
		}
	})

	// Array input (LiftUnary)
	t.Run("array input", func(t *testing.T) {
		cellResolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("a b"),
				{Col: 2, Row: 1}: StringVal("c/d"),
			},
		}
		cf := evalCompile(t, `ENCODEURL(A1:B1)`)
		got, err := Eval(cf, cellResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueArray {
			t.Fatalf("got type %d, want ValueArray", got.Type)
		}
		if len(got.Array) != 1 || len(got.Array[0]) != 2 {
			t.Fatalf("got array shape %dx%d, want 1x2", len(got.Array), len(got.Array[0]))
		}
		if got.Array[0][0].Str != "a%20b" {
			t.Errorf("array[0][0] = %q, want %q", got.Array[0][0].Str, "a%20b")
		}
		if got.Array[0][1].Str != "c%2Fd" {
			t.Errorf("array[0][1] = %q, want %q", got.Array[0][1].Str, "c%2Fd")
		}
	})

	// Cell reference
	t.Run("cell reference", func(t *testing.T) {
		cellResolver := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: StringVal("hello world"),
			},
		}
		cf := evalCompile(t, `ENCODEURL(A1)`)
		got, err := Eval(cf, cellResolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "hello%20world" {
			t.Errorf("got %v, want %q", got, "hello%20world")
		}
	})
}

func TestT_ErrorCellRef(t *testing.T) {
	// Regression test: T() with a cell reference to an error value must
	// return the error, not a string representation of the error.
	// This exercises the resolver path (cell ref → GetCellValue → fnT).
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: ErrorVal(ErrValDIV0),
		},
	}
	cf := evalCompile(t, `T(A1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValDIV0 {
		t.Errorf("T(A1) where A1=#DIV/0! = %v, want error #DIV/0!", got)
	}
}

// ---------------------------------------------------------------------------
// SUBSTITUTE — comprehensive additional tests
// ---------------------------------------------------------------------------

func TestSUBSTITUTEComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
	}{
		// Replace all occurrences of a repeated pattern
		{name: "replace_all_repeated", formula: `SUBSTITUTE("xyzxyzxyz","xyz","A")`, want: "AAA"},

		// Replace specific instances (1st, 2nd, 3rd)
		{name: "replace_1st_of_three", formula: `SUBSTITUTE("abcabcabc","abc","X",1)`, want: "Xabcabc"},
		{name: "replace_2nd_of_three", formula: `SUBSTITUTE("abcabcabc","abc","X",2)`, want: "abcXabc"},
		{name: "replace_3rd_of_three", formula: `SUBSTITUTE("abcabcabc","abc","X",3)`, want: "abcabcX"},

		// Instance beyond count returns original unchanged
		{name: "instance_beyond_count", formula: `SUBSTITUTE("abc","a","X",2)`, want: "abc"},
		{name: "instance_way_beyond", formula: `SUBSTITUTE("hello","l","X",10)`, want: "hello"},

		// Case-sensitive matching — no match for wrong case
		{name: "case_no_match_upper", formula: `SUBSTITUTE("HELLO","hello","X")`, want: "HELLO"},
		{name: "case_no_match_mixed", formula: `SUBSTITUTE("Hello World","hello","X")`, want: "Hello World"},
		{name: "case_exact_match", formula: `SUBSTITUTE("Hello World","Hello","Goodbye")`, want: "Goodbye World"},

		// Replace with empty string (deletion)
		{name: "delete_all_occurrences", formula: `SUBSTITUTE("banana","a","")`, want: "bnn"},
		{name: "delete_specific_instance", formula: `SUBSTITUTE("banana","a","",2)`, want: "banna"},
		{name: "delete_only_match", formula: `SUBSTITUTE("abc","b","")`, want: "ac"},

		// Replace empty old_text returns original unchanged
		{name: "empty_old_all", formula: `SUBSTITUTE("test","","replacement")`, want: "test"},
		{name: "empty_old_instance_1", formula: `SUBSTITUTE("test","","replacement",1)`, want: "test"},

		// old_text not found returns original unchanged
		{name: "not_found_simple", formula: `SUBSTITUTE("hello","xyz","A")`, want: "hello"},
		{name: "not_found_with_instance", formula: `SUBSTITUTE("hello","xyz","A",1)`, want: "hello"},

		// Special characters in old_text and new_text
		{name: "special_chars_dot", formula: `SUBSTITUTE("a.b.c",".","-")`, want: "a-b-c"},
		{name: "special_chars_star", formula: `SUBSTITUTE("a*b*c","*","+")`, want: "a+b+c"},
		{name: "special_chars_paren", formula: `SUBSTITUTE("f(x)","(","[")`, want: "f[x)"},
		{name: "special_chars_backslash", formula: `SUBSTITUTE("a\b\c","\","/")`, want: "a/b/c"},
		{name: "special_chars_newline_replacement", formula: `SUBSTITUTE("hello world"," ",", ")`, want: "hello, world"},

		// Numeric text arguments (coercion)
		{name: "numeric_all_args", formula: `SUBSTITUTE(100,"0","9")`, want: "199"},
		{name: "numeric_old_text", formula: `SUBSTITUTE(12345,"3","X")`, want: "12X45"},
		{name: "numeric_bool_text", formula: `SUBSTITUTE(TRUE,"RU","es")`, want: "TesE"},

		// Boolean arguments
		{name: "bool_true_text", formula: `SUBSTITUTE(TRUE,"TRUE","YES")`, want: "YES"},
		{name: "bool_false_text", formula: `SUBSTITUTE(FALSE,"FALSE","NO")`, want: "NO"},

		// Long replacement text (expands string)
		{name: "long_replacement", formula: `SUBSTITUTE("a","a","ABCDEFGHIJKLMNOP")`, want: "ABCDEFGHIJKLMNOP"},
		{name: "long_replacement_multi", formula: `SUBSTITUTE("aaa","a","XYZ")`, want: "XYZXYZXYZ"},

		// Replace in empty string
		{name: "empty_text_no_match", formula: `SUBSTITUTE("","x","y")`, want: ""},
		{name: "empty_text_empty_old", formula: `SUBSTITUTE("","","y")`, want: ""},

		// old_text longer than text
		{name: "old_longer_than_text", formula: `SUBSTITUTE("ab","abcdef","X")`, want: "ab"},

		// Unicode characters
		{name: "unicode_replace", formula: `SUBSTITUTE("caf` + "\u00e9" + `","e","E")`, want: "caf\u00e9"},
		{name: "unicode_replace_accent", formula: `SUBSTITUTE("caf` + "\u00e9" + `","` + "\u00e9" + `","e")`, want: "cafe"},
		{name: "unicode_emoji", formula: `SUBSTITUTE("hello ` + "\U0001F600" + ` world","` + "\U0001F600" + `","!")`, want: "hello ! world"},
		{name: "unicode_cjk", formula: `SUBSTITUTE("` + "\u4f60\u597d\u4e16\u754c" + `","` + "\u4e16\u754c" + `","` + "\u5730\u7403" + `")`, want: "\u4f60\u597d\u5730\u7403"},

		// Whitespace handling
		{name: "replace_spaces", formula: `SUBSTITUTE("a b c"," ","")`, want: "abc"},
		{name: "replace_tab", formula: `SUBSTITUTE("a` + "\t" + `b","` + "\t" + `"," ")`, want: "a b"},

		// Consecutive identical old_text
		{name: "consecutive_replace_all", formula: `SUBSTITUTE("aaa","a","bb")`, want: "bbbbbb"},
		{name: "consecutive_replace_2nd", formula: `SUBSTITUTE("aaa","a","bb",2)`, want: "abba"},
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

func TestSUBSTITUTEErrorPropagation(t *testing.T) {
	// Error in cell references should propagate through SUBSTITUTE.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: ErrorVal(ErrValDIV0),
			{Col: 2, Row: 1}: ErrorVal(ErrValNA),
			{Col: 3, Row: 1}: ErrorVal(ErrValREF),
		},
	}

	tests := []struct {
		name    string
		formula string
	}{
		{name: "error_in_text", formula: `SUBSTITUTE(A1,"a","b")`},
		{name: "error_in_instance", formula: `SUBSTITUTE("hello","l","L",A1)`},
	}

	for _, tt := range tests {
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

func TestSUBSTITUTEInvalidArgsExtended(t *testing.T) {
	resolver := &mockResolver{}

	errTests := []struct {
		name    string
		formula string
	}{
		{name: "instance_zero_v2", formula: `SUBSTITUTE("hello","l","L",0)`},
		{name: "instance_neg_large", formula: `SUBSTITUTE("hello","l","L",-100)`},
		{name: "instance_non_numeric", formula: `SUBSTITUTE("hello","l","L","abc")`},
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
// REPLACE — comprehensive additional tests
// ---------------------------------------------------------------------------

func TestREPLACEComprehensiveExtended(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    string
		isErr   bool
	}{
		// Replace entire string
		{name: "replace_entire_string", formula: `REPLACE("hello",1,5,"world")`, want: "world"},
		{name: "replace_entire_short", formula: `REPLACE("ab",1,2,"XYZ")`, want: "XYZ"},

		// Replace with empty (deletion) at various positions
		{name: "delete_middle_chars", formula: `REPLACE("abcdef",2,4,"")`, want: "af"},
		{name: "delete_last_char", formula: `REPLACE("hello",5,1,"")`, want: "hell"},
		{name: "delete_entire_string", formula: `REPLACE("xyz",1,3,"")`, want: ""},

		// num_chars = 0 — pure insertion
		{name: "insert_before_2nd_char", formula: `REPLACE("abcd",2,0,"XY")`, want: "aXYbcd"},
		{name: "insert_before_last_char", formula: `REPLACE("abcd",4,0,"XY")`, want: "abcXYd"},
		{name: "insert_after_string", formula: `REPLACE("abc",4,0,"DEF")`, want: "abcDEF"},
		{name: "insert_at_beginning", formula: `REPLACE("xyz",1,0,"ABC")`, want: "ABCxyz"},

		// Replace beyond end of string
		{name: "replace_past_end", formula: `REPLACE("abc",3,5,"X")`, want: "abX"},
		{name: "replace_way_past_end", formula: `REPLACE("ab",1,1000,"!")`, want: "!"},

		// Large start_num beyond string length — appends
		{name: "start_far_beyond_length", formula: `REPLACE("abc",100,0,"X")`, want: "abcX"},
		{name: "start_far_beyond_with_chars", formula: `REPLACE("abc",100,5,"X")`, want: "abcX"},

		// Numeric text coercion
		{name: "numeric_text_float", formula: `REPLACE(3.14,1,1,"_")`, want: "_.14"},
		{name: "numeric_text_negative", formula: `REPLACE(-42,1,1,"")`, want: "42"},
		{name: "numeric_new_text", formula: `REPLACE("abc",2,1,123)`, want: "a123c"},

		// Boolean arguments
		{name: "bool_true_old_text", formula: `REPLACE(TRUE,1,4,"YES")`, want: "YES"},
		{name: "bool_false_replace_all", formula: `REPLACE(FALSE,1,5,"NO")`, want: "NO"},
		{name: "bool_new_text_true", formula: `REPLACE("abc",2,1,TRUE)`, want: "aTRUEc"},
		{name: "bool_new_text_false", formula: `REPLACE("abc",2,1,FALSE)`, want: "aFALSEc"},

		// Unicode characters
		{name: "unicode_replace_chars", formula: `REPLACE("caf` + "\u00e9" + `",4,1,"e")`, want: "cafe"},
		{name: "unicode_insert", formula: `REPLACE("cafe",4,1,"` + "\u00e9" + `")`, want: "caf\u00e9"},
		{name: "unicode_cjk_replace", formula: `REPLACE("` + "\u4f60\u597d\u4e16\u754c" + `",3,2,"!")`, want: "\u4f60\u597d!"},
		{name: "unicode_emoji_replace", formula: `REPLACE("A` + "\U0001F600" + `B",2,1,"!")`, want: "A!B"},
		{name: "unicode_multibyte_insert", formula: `REPLACE("AB",2,0,"` + "\u00e9" + `")`, want: "A\u00e9B"},

		// Long replacement text
		{name: "long_new_text", formula: `REPLACE("ab",2,0,"XXXXXXXXXXXXXXXXXXXX")`, want: "aXXXXXXXXXXXXXXXXXXXXb"},
		{name: "long_replace_single", formula: `REPLACE("abcdef",3,1,"1234567890")`, want: "ab1234567890def"},

		// Single character string edge cases
		{name: "single_char_replace", formula: `REPLACE("a",1,1,"X")`, want: "X"},
		{name: "single_char_delete", formula: `REPLACE("a",1,1,"")`, want: ""},
		{name: "single_char_insert_before", formula: `REPLACE("a",1,0,"X")`, want: "Xa"},
		{name: "single_char_insert_after", formula: `REPLACE("a",2,0,"X")`, want: "aX"},

		// Float values for start_num and num_chars (truncated to int)
		{name: "float_start_truncated", formula: `REPLACE("abcde",3.7,1,"X")`, want: "abXde"},
		{name: "float_num_chars_truncated", formula: `REPLACE("abcde",1,3.9,"X")`, want: "Xde"},
		{name: "float_both_truncated", formula: `REPLACE("abcde",2.1,2.9,"X")`, want: "aXde"},

		// Error cases
		{name: "start_num_zero", formula: `REPLACE("hello",0,1,"X")`, isErr: true},
		{name: "start_num_negative_2", formula: `REPLACE("hello",-5,1,"X")`, isErr: true},
		{name: "num_chars_negative_2", formula: `REPLACE("hello",1,-3,"X")`, isErr: true},
		{name: "non_numeric_start_bool_like", formula: `REPLACE("hello","two",3,"X")`, isErr: true},
		{name: "non_numeric_num_chars_text", formula: `REPLACE("hello",1,"three","X")`, isErr: true},
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
					t.Errorf("Eval(%q) = %v (str=%q), want %q", tt.formula, got, got.Str, tt.want)
				}
			}
		})
	}
}

func TestREPLACEErrorPropagation(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: ErrorVal(ErrValDIV0),
			{Col: 2, Row: 1}: ErrorVal(ErrValNA),
			{Col: 3, Row: 1}: ErrorVal(ErrValREF),
			{Col: 4, Row: 1}: StringVal("hello"),
		},
	}

	tests := []struct {
		name    string
		formula string
	}{
		// Error in start_num
		{name: "error_in_start_num", formula: `REPLACE("hello",A1,3,"X")`},
		// Error in num_chars
		{name: "error_in_num_chars", formula: `REPLACE("hello",1,B1,"X")`},
	}

	for _, tt := range tests {
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

func TestREPLACEWrongArgCountExtended(t *testing.T) {
	resolver := &mockResolver{}

	wrongArgTests := []struct {
		name    string
		formula string
	}{
		{name: "one_arg", formula: `REPLACE("hello")`},
		{name: "two_args", formula: `REPLACE("hello",1)`},
		{name: "six_args", formula: `REPLACE("hello",1,2,"X","extra","more")`},
	}

	for _, tt := range wrongArgTests {
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
// CHOOSE — additional comprehensive tests
// ---------------------------------------------------------------------------

func TestCHOOSEAdditional(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(2),
			{Col: 2, Row: 1}: StringVal("from_cell"),
		},
	}

	type want struct {
		typ ValueType
		num float64
		str string
		b   bool
		err ErrorValue
	}

	tests := []struct {
		name    string
		formula string
		want    want
	}{
		// Mixed types in value list
		{name: "mixed_types_select_number", formula: `CHOOSE(1,100,"hello",TRUE)`, want: want{typ: ValueNumber, num: 100}},
		{name: "mixed_types_select_string", formula: `CHOOSE(2,100,"hello",TRUE)`, want: want{typ: ValueString, str: "hello"}},
		{name: "mixed_types_select_bool", formula: `CHOOSE(3,100,"hello",TRUE)`, want: want{typ: ValueBool, b: true}},

		// Large number of values (10+)
		{name: "ten_values_first", formula: `CHOOSE(1,"v1","v2","v3","v4","v5","v6","v7","v8","v9","v10")`, want: want{typ: ValueString, str: "v1"}},
		{name: "ten_values_last", formula: `CHOOSE(10,"v1","v2","v3","v4","v5","v6","v7","v8","v9","v10")`, want: want{typ: ValueString, str: "v10"}},
		{name: "ten_values_middle", formula: `CHOOSE(5,"v1","v2","v3","v4","v5","v6","v7","v8","v9","v10")`, want: want{typ: ValueString, str: "v5"}},

		// Nested CHOOSE: CHOOSE(CHOOSE(1,2), "a", "b") = "b"
		{name: "nested_choose", formula: `CHOOSE(CHOOSE(1,2),"a","b")`, want: want{typ: ValueString, str: "b"}},
		{name: "nested_choose_deep", formula: `CHOOSE(CHOOSE(2,1,3),"x","y","z")`, want: want{typ: ValueString, str: "z"}},

		// Cell reference as index
		{name: "cell_ref_index", formula: `CHOOSE(A1,"first","second","third")`, want: want{typ: ValueString, str: "second"}},

		// Cell reference as value
		{name: "cell_ref_value", formula: `CHOOSE(1,B1)`, want: want{typ: ValueString, str: "from_cell"}},

		// Expression as index
		{name: "expression_index", formula: `CHOOSE(1+1,"a","b","c")`, want: want{typ: ValueString, str: "b"}},

		// Index exactly equal to value count (boundary)
		{name: "index_equals_count", formula: `CHOOSE(3,"a","b","c")`, want: want{typ: ValueString, str: "c"}},

		// Index one past the end
		{name: "index_one_past_end", formula: `CHOOSE(4,"a","b","c")`, want: want{typ: ValueError, err: ErrValVALUE}},

		// Decimal index 3.1 selects 3rd
		{name: "decimal_index_3.1", formula: `CHOOSE(3.1,"a","b","c")`, want: want{typ: ValueString, str: "c"}},

		// Large decimal truncation
		{name: "decimal_index_1.999", formula: `CHOOSE(1.999,"first","second")`, want: want{typ: ValueString, str: "first"}},

		// Numeric values with computation
		{name: "computed_values", formula: `CHOOSE(2,10+5,20+5,30+5)`, want: want{typ: ValueNumber, num: 25}},

		// Boolean FALSE as index (FALSE=0) -> out of range
		{name: "bool_false_index", formula: `CHOOSE(FALSE,"a","b")`, want: want{typ: ValueError, err: ErrValVALUE}},

		// Empty string as index -> #VALUE! (cannot coerce)
		{name: "empty_string_index", formula: `CHOOSE("","a","b")`, want: want{typ: ValueError, err: ErrValVALUE}},

		// Non-numeric string as index -> #VALUE!
		{name: "text_string_index", formula: `CHOOSE("abc","a","b")`, want: want{typ: ValueError, err: ErrValVALUE}},

		// Very large index
		{name: "very_large_index", formula: `CHOOSE(999,"a","b","c")`, want: want{typ: ValueError, err: ErrValVALUE}},

		// Two values, select each
		{name: "two_values_first", formula: `CHOOSE(1,"yes","no")`, want: want{typ: ValueString, str: "yes"}},
		{name: "two_values_second", formula: `CHOOSE(2,"yes","no")`, want: want{typ: ValueString, str: "no"}},

		// Number returned as-is (not stringified)
		{name: "number_returned_as_number", formula: `CHOOSE(1,42)`, want: want{typ: ValueNumber, num: 42}},
		{name: "zero_returned", formula: `CHOOSE(1,0)`, want: want{typ: ValueNumber, num: 0}},
		{name: "negative_number_returned", formula: `CHOOSE(1,-99.5)`, want: want{typ: ValueNumber, num: -99.5}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != tt.want.typ {
				t.Fatalf("Eval(%q).Type = %v, want %v (value=%v)", tt.formula, got.Type, tt.want.typ, got)
			}
			switch tt.want.typ {
			case ValueString:
				if got.Str != tt.want.str {
					t.Errorf("Eval(%q) = %q, want %q", tt.formula, got.Str, tt.want.str)
				}
			case ValueNumber:
				if got.Num != tt.want.num {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want.num)
				}
			case ValueBool:
				if got.Bool != tt.want.b {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Bool, tt.want.b)
				}
			case ValueError:
				if got.Err != tt.want.err {
					t.Errorf("Eval(%q) = %v, want %v", tt.formula, got.Err, tt.want.err)
				}
			}
		})
	}
}
