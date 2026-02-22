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
