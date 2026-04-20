package formula

import (
	"testing"
)

func TestREGEXTEST(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    bool
	}{
		{`REGEXTEST("Order 12345", "[0-9]+")`, true},
		{`REGEXTEST("Order 12345", "^Order")`, true},
		{`REGEXTEST("Order 12345", "^shipped")`, false},
		{`REGEXTEST("", "a")`, false},
		{`REGEXTEST("", "a?")`, true},
		{`REGEXTEST("abc", "")`, true},
		{`REGEXTEST("line1" & CHAR(10) & "line3", "line3$")`, true},
		{`REGEXTEST("Apple BANANA cherry", "banana")`, false},
		{`REGEXTEST("Apple BANANA cherry", "banana", 1)`, true},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueBool || got.Bool != tt.want {
			t.Errorf("Eval(%q) = %+v, want %v", tt.formula, got, tt.want)
		}
	}
}

func TestREGEXEXTRACT_Mode0(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    string
	}{
		{`REGEXEXTRACT("Order 12345 shipped", "[0-9]+")`, "12345"},
		{`REGEXEXTRACT("abc123def456", "[0-9]+", 0)`, "123"},
		{`REGEXEXTRACT("3.14 and 31415", "\d+\.\d+")`, "3.14"},
		{`REGEXEXTRACT("aaabbbccc", "a+b+")`, "aaabbb"},
		{`REGEXEXTRACT("aaabbbccc", "a{2}")`, "aa"},
		// Case-insensitive flag.
		{`REGEXEXTRACT("Apple BANANA cherry", "[a-z]+")`, "pple"},
		{`REGEXEXTRACT("Apple BANANA cherry", "[a-z]+", 0, 1)`, "Apple"},
		// Zero-width match against empty input: empty string, not #N/A.
		{`REGEXEXTRACT("", "a?")`, ""},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueString || got.Str != tt.want {
			t.Errorf("Eval(%q) = %+v, want %q", tt.formula, got, tt.want)
		}
	}
}

func TestREGEXEXTRACT_Mode0_NoMatch(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `REGEXEXTRACT("xyz", "[0-9]+")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNA {
		t.Errorf("got %+v, want #N/A", got)
	}
}

func TestREGEXEXTRACT_Mode1_AllMatches(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `REGEXEXTRACT("abc123def456ghi789", "[0-9]+", 1)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("want array, got %+v", got)
	}
	if len(got.Array) != 3 || len(got.Array[0]) != 1 {
		t.Fatalf("want 3x1, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	want := []string{"123", "456", "789"}
	for i, w := range want {
		if got.Array[i][0].Str != w {
			t.Errorf("row %d: got %q, want %q", i, got.Array[i][0].Str, w)
		}
	}
}

func TestREGEXEXTRACT_Mode2_CaptureGroups(t *testing.T) {
	resolver := &mockResolver{}
	cf := evalCompile(t, `REGEXEXTRACT("hello-world-2024", "(\w+)-(\w+)-(\d+)", 2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("want array, got %+v", got)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 3 {
		t.Fatalf("want 1x3, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	want := []string{"hello", "world", "2024"}
	for i, w := range want {
		if got.Array[0][i].Str != w {
			t.Errorf("col %d: got %q, want %q", i, got.Array[0][i].Str, w)
		}
	}
}

func TestREGEXEXTRACT_Mode2_NoCaptureGroups(t *testing.T) {
	// Mode 2 without capture groups returns the whole match as a 1x1 array
	// so scalar and array-input paths share the same shape semantics.
	resolver := &mockResolver{}
	cf := evalCompile(t, `REGEXEXTRACT("abc123", "[0-9]+", 2)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray {
		t.Fatalf("want array, got %+v", got)
	}
	if len(got.Array) != 1 || len(got.Array[0]) != 1 {
		t.Fatalf("want 1x1, got %dx%d", len(got.Array), len(got.Array[0]))
	}
	if got.Array[0][0].Str != "123" {
		t.Errorf("got %q, want 123", got.Array[0][0].Str)
	}
}

func TestREGEXREPLACE_ZeroWidthMultibyte(t *testing.T) {
	// Zero-width match with multi-byte UTF-8 must advance by a full rune,
	// not a single byte, or the result is corrupt.
	resolver := &mockResolver{}
	cf := evalCompile(t, `REGEXREPLACE("日本語", "^", "X")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "X日本語" {
		t.Errorf("got %+v, want %q", got, "X日本語")
	}
	// \b is a zero-width boundary: every word boundary gets a '|'. For an
	// all-word input the boundaries are at the start and end, so the output
	// must be well-formed UTF-8 with both delimiters.
	cf2 := evalCompile(t, `REGEXREPLACE("日本", "\b", "|")`)
	got2, err := Eval(cf2, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	// Go's \b is ASCII-word-aware; non-ASCII letters are not word chars, so
	// \b won't match at all here. We only care that the string is not
	// corrupted — result equals the input.
	if got2.Type != ValueString || got2.Str != "日本" {
		t.Errorf("got %+v, want %q", got2, "日本")
	}
}

func TestREGEXREPLACE_BasicAndBackrefs(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    string
	}{
		{`REGEXREPLACE("Order 12345", "[0-9]", "*")`, "Order *****"},
		{`REGEXREPLACE("John Smith", "(\w+) (\w+)", "$2, $1")`, "Smith, John"},
		{`REGEXREPLACE("2024-07-15", "(\d{4})-(\d{2})-(\d{2})", "$2/$3/$1")`, "07/15/2024"},
		{`REGEXREPLACE("foo", "(foo)", "($1)")`, "(foo)"},
		{`REGEXREPLACE("abc", "b", "[$0]")`, "a[b]c"},
		{`REGEXREPLACE("catscan is not cat", "\bcat\b", "dog")`, "catscan is not dog"},
		{`REGEXREPLACE("abc", "^", "X")`, "Xabc"},
		{`REGEXREPLACE("abc", "\b", "|")`, "|abc|"},
		{`REGEXREPLACE("abc", "z", "Q")`, "abc"},
		// Case-insensitive flag.
		{`REGEXREPLACE("Apple BANANA cherry", "apple", "PEAR", , 1)`, "PEAR BANANA cherry"},
		{`REGEXREPLACE("Apple BANANA cherry", "apple", "PEAR")`, "Apple BANANA cherry"},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueString || got.Str != tt.want {
			t.Errorf("Eval(%q) = %+v, want %q", tt.formula, got, tt.want)
		}
	}
}

func TestREGEXREPLACE_InstanceNum(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    string
	}{
		// Replace only the first match.
		{`REGEXREPLACE("a1b2c3", "\d", "X", 1)`, "aXb2c3"},
		// Replace only the second match.
		{`REGEXREPLACE("a1b2c3", "\d", "X", 2)`, "a1bXc3"},
		// Negative: last match.
		{`REGEXREPLACE("a1b2c3", "\d", "X", -1)`, "a1b2cX"},
		// Out of range: unchanged.
		{`REGEXREPLACE("a1b2c3", "\d", "X", 99)`, "a1b2c3"},
	}

	for _, tt := range tests {
		cf := evalCompile(t, tt.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", tt.formula, err)
			continue
		}
		if got.Type != ValueString || got.Str != tt.want {
			t.Errorf("Eval(%q) = %+v, want %q", tt.formula, got, tt.want)
		}
	}
}

func TestREGEX_InvalidPattern(t *testing.T) {
	resolver := &mockResolver{}

	// Go's RE2 rejects backreferences inside the pattern. This is a known
	// limitation vs Excel's ICU; we surface it as #VALUE!.
	tests := []string{
		`REGEXTEST("abc", "(?<")`,
		`REGEXEXTRACT("abc", "[")`,
		`REGEXREPLACE("abc", "(", "x")`,
	}

	for _, f := range tests {
		cf := evalCompile(t, f)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Errorf("Eval(%q): %v", f, err)
			continue
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("Eval(%q) = %+v, want #VALUE!", f, got)
		}
	}
}

func TestREGEX_ArrayLift(t *testing.T) {
	// A range argument spills: A1:A3 → a 3-row column result.
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("abc123"),
			{Col: 1, Row: 2}: StringVal("no numbers here"),
			{Col: 1, Row: 3}: StringVal("42"),
		},
	}

	cf := evalCompile(t, `REGEXTEST(A1:A3, "[0-9]")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("want 3-row array, got %+v", got)
	}
	wantBools := []bool{true, false, true}
	for i, w := range wantBools {
		if got.Array[i][0].Type != ValueBool || got.Array[i][0].Bool != w {
			t.Errorf("row %d: got %+v, want %v", i, got.Array[i][0], w)
		}
	}
}

func TestREGEXEXTRACT_Mode0_ArrayLift(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("abc123"),
			{Col: 1, Row: 2}: StringVal("no numbers here"),
			{Col: 1, Row: 3}: StringVal("42"),
		},
	}

	cf := evalCompile(t, `REGEXEXTRACT(A1:A3, "[0-9]+")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("want 3-row array, got %+v", got)
	}
	if got.Array[0][0].Str != "123" {
		t.Errorf("row 0: got %+v, want 123", got.Array[0][0])
	}
	if got.Array[1][0].Type != ValueError || got.Array[1][0].Err != ErrValNA {
		t.Errorf("row 1: got %+v, want #N/A", got.Array[1][0])
	}
	if got.Array[2][0].Str != "42" {
		t.Errorf("row 2: got %+v, want 42", got.Array[2][0])
	}
}

func TestREGEX_PatternCache(t *testing.T) {
	// Repeated calls with the same pattern should not recompile. Not a
	// behavioural test so much as a sanity check that the cache is keyed
	// correctly: case-sensitive and case-insensitive must not collide.
	resolver := &mockResolver{}
	a := evalCompile(t, `REGEXTEST("APPLE", "apple")`)
	b := evalCompile(t, `REGEXTEST("APPLE", "apple", 1)`)
	ra, _ := Eval(a, resolver, nil)
	rb, _ := Eval(b, resolver, nil)
	if ra.Type != ValueBool || ra.Bool != false {
		t.Errorf("case-sensitive: got %+v, want false", ra)
	}
	if rb.Type != ValueBool || rb.Bool != true {
		t.Errorf("case-insensitive: got %+v, want true", rb)
	}
}
