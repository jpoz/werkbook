package formula

import (
	"testing"
)

func TestMedian(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(3),
			{Col: 1, Row: 3}: NumberVal(5),
		},
	}

	cf := evalCompile(t, "MEDIAN(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 3 {
		t.Errorf("MEDIAN odd: got %g, want 3", got.Num)
	}

	resolver.cells[CellAddr{Col: 1, Row: 4}] = NumberVal(7)
	cf = evalCompile(t, "MEDIAN(A1:A4)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 4 {
		t.Errorf("MEDIAN even: got %g, want 4", got.Num)
	}
}

func TestLargeSmall(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(30),
			{Col: 1, Row: 3}: NumberVal(20),
		},
	}

	cf := evalCompile(t, "LARGE(A1:A3,1)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("LARGE k=1: got %g, want 30", got.Num)
	}

	cf = evalCompile(t, "SMALL(A1:A3,1)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 10 {
		t.Errorf("SMALL k=1: got %g, want 10", got.Num)
	}
}

func TestCountBlank(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			// A2 is empty
			{Col: 1, Row: 3}: NumberVal(3),
		},
	}

	cf := evalCompile(t, "COUNTBLANK(A1:A3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("COUNTBLANK: got %g, want 1", got.Num)
	}
}

func TestSUMIF(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(10),
			{Col: 1, Row: 2}: NumberVal(20),
			{Col: 1, Row: 3}: NumberVal(30),
		},
	}

	cf := evalCompile(t, `SUMIF(A1:A3,">15")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 50 {
		t.Errorf("SUMIF >15: got %g, want 50", got.Num)
	}
}

func TestCOUNTIF(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: StringVal("apple"),
			{Col: 1, Row: 2}: StringVal("banana"),
			{Col: 1, Row: 3}: StringVal("apple"),
		},
	}

	cf := evalCompile(t, `COUNTIF(A1:A3,"apple")`)
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2 {
		t.Errorf("COUNTIF: got %g, want 2", got.Num)
	}
}

func TestSUMPRODUCT(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(1),
			{Col: 1, Row: 2}: NumberVal(2),
			{Col: 1, Row: 3}: NumberVal(3),
			{Col: 2, Row: 1}: NumberVal(4),
			{Col: 2, Row: 2}: NumberVal(5),
			{Col: 2, Row: 3}: NumberVal(6),
		},
	}

	// 1*4 + 2*5 + 3*6 = 4 + 10 + 18 = 32
	cf := evalCompile(t, "SUMPRODUCT(A1:A3,B1:B3)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 32 {
		t.Errorf("SUMPRODUCT: got %g, want 32", got.Num)
	}
}

func TestMatchesCriteria(t *testing.T) {
	tests := []struct {
		v    Value
		crit Value
		want bool
	}{
		{NumberVal(10), StringVal(">5"), true},
		{NumberVal(3), StringVal(">5"), false},
		{NumberVal(5), StringVal(">=5"), true},
		{NumberVal(5), StringVal("<=5"), true},
		{NumberVal(5), StringVal("<>5"), false},
		{NumberVal(5), NumberVal(5), true},
		{StringVal("apple"), StringVal("app*"), true},
		{StringVal("banana"), StringVal("app*"), false},
		{StringVal("cat"), StringVal("c?t"), true},
	}

	for _, tt := range tests {
		got := matchesCriteria(tt.v, tt.crit)
		if got != tt.want {
			t.Errorf("matchesCriteria(%v, %v) = %v, want %v", tt.v, tt.crit, got, tt.want)
		}
	}
}
