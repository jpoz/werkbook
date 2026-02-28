package formula

import (
	"testing"
)

func TestXOR(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    bool
	}{
		{"XOR(TRUE,FALSE)", true},
		{"XOR(TRUE,TRUE)", false},
		{"XOR(FALSE,FALSE)", false},
		{"XOR(TRUE,TRUE,TRUE)", true},
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

func TestIFS(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("first true", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE, "yes")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf("IFS(TRUE, \"yes\") = %v, want \"yes\"", got)
		}
	})

	t.Run("second pair true", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE, "a", TRUE, "b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "b" {
			t.Errorf(`IFS(FALSE, "a", TRUE, "b") = %v, want "b"`, got)
		}
	})

	t.Run("no true returns NA", func(t *testing.T) {
		cf := evalCompile(t, `IFS(FALSE, "a", FALSE, "b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`IFS(FALSE, "a", FALSE, "b") = %v, want #N/A`, got)
		}
	})

	t.Run("odd arg count returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `IFS(TRUE)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("IFS(TRUE) = %v, want #VALUE!", got)
		}
	})
}

func TestSWITCH(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("match second value", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(2, 1, "a", 2, "b", 3, "c")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "b" {
			t.Errorf(`SWITCH(2,1,"a",2,"b",3,"c") = %v, want "b"`, got)
		}
	})

	t.Run("no match returns NA", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(99, 1, "a", 2, "b")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf(`SWITCH(99,1,"a",2,"b") = %v, want #N/A`, got)
		}
	})

	t.Run("no match with default", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(99, 1, "a", "default")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "default" {
			t.Errorf(`SWITCH(99,1,"a","default") = %v, want "default"`, got)
		}
	})

	t.Run("string match", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH("x", "x", 1, "y", 2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf(`SWITCH("x","x",1,"y",2) = %v, want 1`, got)
		}
	})

	t.Run("single pair match", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1, 1, "match")`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueString || got.Str != "match" {
			t.Errorf(`SWITCH(1,1,"match") = %v, want "match"`, got)
		}
	})

	t.Run("too few args returns VALUE error", func(t *testing.T) {
		cf := evalCompile(t, `SWITCH(1, 2)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("SWITCH(1,2) = %v, want #VALUE!", got)
		}
	})
}

func TestSORT(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(30),
			{Col: 1, Row: 2}: NumberVal(10),
			{Col: 1, Row: 3}: NumberVal(20),
		},
	}

	cf := evalCompile(t, "SORT(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueArray || len(got.Array) != 3 {
		t.Fatalf("SORT: got type=%v len=%d", got.Type, len(got.Array))
	}
	if got.Array[0][0].Num != 10 || got.Array[1][0].Num != 20 || got.Array[2][0].Num != 30 {
		t.Errorf("SORT: got [%g,%g,%g], want [10,20,30]",
			got.Array[0][0].Num, got.Array[1][0].Num, got.Array[2][0].Num)
	}
}
