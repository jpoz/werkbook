package formula

import (
	"math"
	"testing"
)

func TestDATE(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "DATE(2024,1,15)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber {
		t.Fatalf("DATE: got type %v, want number", got.Type)
	}
	// Jan 15, 2024 should be serial 45306
	if got.Num != 45306 {
		t.Errorf("DATE(2024,1,15) = %g, want 45306", got.Num)
	}
}

func TestYEAR(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "YEAR(45306)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 2024 {
		t.Errorf("YEAR: got %g, want 2024", got.Num)
	}
}

func TestMONTH(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "MONTH(45306)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("MONTH: got %g, want 1", got.Num)
	}
}

func TestDAY(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "DAY(45306)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 15 {
		t.Errorf("DAY: got %g, want 15", got.Num)
	}
}

func TestTIME(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "TIME(12,30,0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber {
		t.Fatalf("TIME: got type %v", got.Type)
	}
	want := (12*3600.0 + 30*60.0) / 86400.0
	if math.Abs(got.Num-want) > 1e-10 {
		t.Errorf("TIME(12,30,0) = %g, want %g", got.Num, want)
	}
}

func TestHOURMINUTESECOND(t *testing.T) {
	resolver := &mockResolver{}

	// TIME(14,30,45) = fractional day
	serial := (14*3600.0 + 30*60.0 + 45) / 86400.0

	resolver2 := &mockResolver{
		cells: map[CellAddr]Value{
			{Col: 1, Row: 1}: NumberVal(serial),
		},
	}

	cf := evalCompile(t, "HOUR(A1)")
	got, err := Eval(cf, resolver2, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 14 {
		t.Errorf("HOUR: got %g, want 14", got.Num)
	}

	cf = evalCompile(t, "MINUTE(A1)")
	got, err = Eval(cf, resolver2, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 30 {
		t.Errorf("MINUTE: got %g, want 30", got.Num)
	}

	cf = evalCompile(t, "SECOND(A1)")
	got, err = Eval(cf, resolver2, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueNumber || got.Num != 45 {
		t.Errorf("SECOND: got %g, want 45", got.Num)
	}

	_ = resolver
}

func TestNOWTODAY(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "NOW()")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(NOW): %v", err)
	}
	if got.Type != ValueNumber || got.Num < 40000 {
		t.Errorf("NOW() = %g, expected a large serial date", got.Num)
	}

	cf = evalCompile(t, "TODAY()")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(TODAY): %v", err)
	}
	if got.Type != ValueNumber || got.Num < 40000 {
		t.Errorf("TODAY() = %g, expected a large serial date", got.Num)
	}
	if got.Num != math.Floor(got.Num) {
		t.Errorf("TODAY() = %g, expected integer (no fractional time)", got.Num)
	}
}
