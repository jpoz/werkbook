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
