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

func TestDAYS(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    float64
	}{
		{"DAYS(DATE(2021,3,15), DATE(2021,2,1))", 42},
		{"DAYS(DATE(2021,12,31), DATE(2021,1,1))", 364},
		{"DAYS(DATE(2020,1,1), DATE(2021,1,1))", -366},
		{"DAYS(100, 50)", 50},
	}

	for _, tc := range tests {
		cf := evalCompile(t, tc.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", tc.formula, err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("DAYS: got type %v, want number for %s", got.Type, tc.formula)
		}
		if got.Num != tc.want {
			t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
		}
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

func TestWEEKDAY(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    float64
	}{
		// Feb 14, 2008 is a Thursday
		// Type 1 (default): Sunday=1, so Thursday=5
		{"WEEKDAY(DATE(2008,2,14))", 5},
		// Type 2: Monday=1, so Thursday=4
		{"WEEKDAY(DATE(2008,2,14), 2)", 4},
		// Type 3: Monday=0, so Thursday=3
		{"WEEKDAY(DATE(2008,2,14), 3)", 3},

		// Jan 1, 2024 is a Monday
		// Type 1 (default): Sunday=1, so Monday=2
		{"WEEKDAY(DATE(2024,1,1))", 2},
		// Type 2: Monday=1, so Monday=1
		{"WEEKDAY(DATE(2024,1,1), 2)", 1},

		// Type 11 same as type 2: Monday=1
		{"WEEKDAY(DATE(2024,1,1), 11)", 1},
		// Type 17 same as type 1: Sunday=1, Monday=2
		{"WEEKDAY(DATE(2024,1,1), 17)", 2},

		// Jan 7, 2024 is a Sunday
		// Type 1: Sunday=1
		{"WEEKDAY(DATE(2024,1,7))", 1},
		// Type 2: Sunday=7
		{"WEEKDAY(DATE(2024,1,7), 2)", 7},
		// Type 3: Sunday=6
		{"WEEKDAY(DATE(2024,1,7), 3)", 6},

		// Type 12: Tuesday=1; for Thursday => (4-2+7)%7+1 = 3
		{"WEEKDAY(DATE(2008,2,14), 12)", 3},
		// Type 13: Wednesday=1; for Thursday => (4-3+7)%7+1 = 2
		{"WEEKDAY(DATE(2008,2,14), 13)", 2},
		// Type 14: Thursday=1; for Thursday => (4-4+7)%7+1 = 1
		{"WEEKDAY(DATE(2008,2,14), 14)", 1},
		// Type 15: Friday=1; for Thursday => (4-5+7)%7+1 = 7
		{"WEEKDAY(DATE(2008,2,14), 15)", 7},
		// Type 16: Saturday=1; for Thursday => (4-6+7)%7+1 = 6
		{"WEEKDAY(DATE(2008,2,14), 16)", 6},
	}

	for _, tc := range tests {
		cf := evalCompile(t, tc.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", tc.formula, err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("WEEKDAY: got type %v, want number for %s", got.Type, tc.formula)
		}
		if got.Num != tc.want {
			t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
		}
	}
}

func TestWEEKDAY_InvalidReturnType(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "WEEKDAY(DATE(2024,1,1), 5)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("WEEKDAY with invalid return_type: got %v, want #NUM!", got)
	}
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
