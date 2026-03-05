package formula

import (
	"testing"
)

func TestISEVEN(t *testing.T) {
	resolver := &mockResolver{}

	// ISEVEN(4) = TRUE
	cf := evalCompile(t, `ISEVEN(4)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("ISEVEN(4) = %v, want true", got)
	}

	// ISEVEN(3) = FALSE
	cf = evalCompile(t, `ISEVEN(3)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool {
		t.Errorf("ISEVEN(3) = %v, want false", got)
	}

	// ISEVEN(TRUE) = #VALUE!
	cf = evalCompile(t, `ISEVEN(TRUE)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISEVEN(TRUE) = %v, want #VALUE!", got)
	}

	// ISEVEN(FALSE) = #VALUE!
	cf = evalCompile(t, `ISEVEN(FALSE)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISEVEN(FALSE) = %v, want #VALUE!", got)
	}
}

func TestISODD(t *testing.T) {
	resolver := &mockResolver{}

	// ISODD(3) = TRUE
	cf := evalCompile(t, `ISODD(3)`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || !got.Bool {
		t.Errorf("ISODD(3) = %v, want true", got)
	}

	// ISODD(4) = FALSE
	cf = evalCompile(t, `ISODD(4)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueBool || got.Bool {
		t.Errorf("ISODD(4) = %v, want false", got)
	}

	// ISODD(TRUE) = #VALUE!
	cf = evalCompile(t, `ISODD(TRUE)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISODD(TRUE) = %v, want #VALUE!", got)
	}

	// ISODD(FALSE) = #VALUE!
	cf = evalCompile(t, `ISODD(FALSE)`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValVALUE {
		t.Errorf("ISODD(FALSE) = %v, want #VALUE!", got)
	}
}

func TestIFNA(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, `IFNA(#N/A,"default")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueString || got.Str != "default" {
		t.Errorf("IFNA(#N/A) = %v, want default", got)
	}

	cf = evalCompile(t, `IFNA(42,"default")`)
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 42 {
		t.Errorf("IFNA(42) = %v, want 42", got)
	}
}
