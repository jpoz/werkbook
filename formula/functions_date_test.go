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

func TestFnDATEEdgeCases(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		// Pre-truncation negative year: -0.5 < 0 even though int(-0.5) == 0
		{"negative_fraction_year", "DATE(-0.5, 1, 1)", 0, true, ErrValNUM},
		// Tiny negative year: -1e-15 < 0
		{"tiny_negative_year", "DATE(-0.000000000000001, 1, 1)", 0, true, ErrValNUM},
		// Year >= 10000
		{"year_10000", "DATE(10000, 1, 1)", 0, true, ErrValNUM},
		// Year correction 0-1899: DATE(108,1,2) → year 2008, serial 39449
		{"year_correction_108", "DATE(108, 1, 2)", 39449, false, 0},
		// Year 0 correction: DATE(0,1,1) → year 1900, serial 1
		{"year_zero_correction", "DATE(0, 1, 1)", 1, false, 0},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if tc.isErr {
				if got.Type != ValueError || got.Err != tc.errVal {
					t.Errorf("%s: got %v, want error %v", tc.formula, got, tc.errVal)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: got type %v, want number", tc.formula, got.Type)
			}
			if got.Num != tc.want {
				t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
			}
		})
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

func TestDATEVALUE(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		{`DATEVALUE("8/22/2011")`, 40777, false, 0},
		{`DATEVALUE("2011/02/23")`, 40597, false, 0},
		{`DATEVALUE("2011-02-23")`, 40597, false, 0},
		{`DATEVALUE("not a date")`, 0, true, ErrValVALUE},
	}

	for _, tc := range tests {
		cf := evalCompile(t, tc.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", tc.formula, err)
		}
		if tc.isErr {
			if got.Type != ValueError || got.Err != tc.errVal {
				t.Errorf("%s: got %v, want error %v", tc.formula, got, tc.errVal)
			}
			continue
		}
		if got.Type != ValueNumber {
			t.Fatalf("%s: got type %v, want number", tc.formula, got.Type)
		}
		if got.Num != tc.want {
			t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
		}
	}
}

func TestEDATE(t *testing.T) {
	resolver := &mockResolver{}

	// Helper: evaluate a formula and return its numeric value.
	evalNum := func(formula string) float64 {
		t.Helper()
		cf := evalCompile(t, formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", formula, err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("%s: got type %v, want number", formula, got.Type)
		}
		return got.Num
	}

	tests := []struct {
		formula  string
		expected string
	}{
		{"EDATE(DATE(2011,1,15),1)", "DATE(2011,2,15)"},
		{"EDATE(DATE(2011,1,15),-1)", "DATE(2010,12,15)"},
		{"EDATE(DATE(2011,1,15),0)", "DATE(2011,1,15)"},
		{"EDATE(DATE(2011,1,31),1)", "DATE(2011,2,28)"}, // clamped
	}

	for _, tc := range tests {
		got := evalNum(tc.formula)
		want := evalNum(tc.expected)
		if got != want {
			t.Errorf("%s = %g, want %g (from %s)", tc.formula, got, want, tc.expected)
		}
	}
}

func TestEOMONTH(t *testing.T) {
	resolver := &mockResolver{}

	evalNum := func(formula string) float64 {
		t.Helper()
		cf := evalCompile(t, formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", formula, err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("%s: got type %v, want number", formula, got.Type)
		}
		return got.Num
	}

	tests := []struct {
		formula  string
		expected string
	}{
		{"EOMONTH(DATE(2011,1,1),1)", "DATE(2011,2,28)"},
		{"EOMONTH(DATE(2011,1,1),0)", "DATE(2011,1,31)"},
		{"EOMONTH(DATE(2020,1,1),1)", "DATE(2020,2,29)"}, // leap year
		{"EOMONTH(DATE(2011,1,1),-3)", "DATE(2010,10,31)"},
	}

	for _, tc := range tests {
		got := evalNum(tc.formula)
		want := evalNum(tc.expected)
		if got != want {
			t.Errorf("%s = %g, want %g (from %s)", tc.formula, got, want, tc.expected)
		}
	}
}

func TestNETWORKDAYS(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    float64
	}{
		// Oct 1 2012 = Monday, Oct 5 = Friday → 5 weekdays
		{"NETWORKDAYS(DATE(2012,10,1),DATE(2012,10,5))", 5},
		// Single weekday
		{"NETWORKDAYS(DATE(2012,10,1),DATE(2012,10,1))", 1},
		// Saturday only → 0
		{"NETWORKDAYS(DATE(2012,10,6),DATE(2012,10,6))", 0},
		// Mon-Sun: 5 weekdays (Mon-Fri), Sat+Sun excluded
		{"NETWORKDAYS(DATE(2012,10,1),DATE(2012,10,7))", 5},
	}

	for _, tc := range tests {
		cf := evalCompile(t, tc.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", tc.formula, err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("NETWORKDAYS: got type %v, want number for %s", got.Type, tc.formula)
		}
		if got.Num != tc.want {
			t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
		}
	}
}

func TestWEEKNUM(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    float64
	}{
		// March 9, 2012 is a Friday
		{"WEEKNUM(DATE(2012,3,9))", 10},      // Sunday start (default)
		{"WEEKNUM(DATE(2012,3,9),2)", 11},     // Monday start
		{"WEEKNUM(DATE(2012,1,1))", 1},        // Jan 1 always week 1
		{"WEEKNUM(DATE(2012,3,9),21)", 10},    // ISO week
	}

	for _, tc := range tests {
		cf := evalCompile(t, tc.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", tc.formula, err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("WEEKNUM: got type %v, want number for %s", got.Type, tc.formula)
		}
		if got.Num != tc.want {
			t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
		}
	}
}

func TestWEEKNUM_Invalid(t *testing.T) {
	resolver := &mockResolver{}

	cf := evalCompile(t, "WEEKNUM(DATE(2012,3,9),99)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval: %v", err)
	}
	if got.Type != ValueError || got.Err != ErrValNUM {
		t.Errorf("WEEKNUM with invalid return_type: got %v, want #NUM!", got)
	}
}

func TestDATEDIF(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		{`DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"Y")`, 2, false, 0},
		{`DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"D")`, 440, false, 0},
		{`DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YD")`, 75, false, 0},
		{`DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"M")`, 14, false, 0},
		{`DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YM")`, 2, false, 0},
		{`DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"MD")`, 14, false, 0},
		{`DATEDIF(DATE(2003,1,1),DATE(2001,1,1),"Y")`, 0, true, ErrValNUM},
		{`DATEDIF(DATE(2020,1,15),DATE(2020,3,10),"MD")`, 24, false, 0},
	}

	for _, tc := range tests {
		cf := evalCompile(t, tc.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", tc.formula, err)
		}
		if tc.isErr {
			if got.Type != ValueError || got.Err != tc.errVal {
				t.Errorf("%s: got %v, want error %v", tc.formula, got, tc.errVal)
			}
			continue
		}
		if got.Type != ValueNumber {
			t.Fatalf("%s: got type %v, want number", tc.formula, got.Type)
		}
		if got.Num != tc.want {
			t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
		}
	}
}

func TestISOWEEKNUM(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		formula string
		want    float64
	}{
		{"ISOWEEKNUM(DATE(2012,3,9))", 10},
		{"ISOWEEKNUM(DATE(2023,1,1))", 52},  // Jan 1 2023 is a Sunday, belongs to week 52 of 2022
		{"ISOWEEKNUM(DATE(2023,1,2))", 1},   // first Monday of 2023
		{"ISOWEEKNUM(DATE(2020,12,31))", 53}, // 2020 has 53 ISO weeks
		{"ISOWEEKNUM(DATE(2021,1,4))", 1},   // first Monday of 2021
		{"ISOWEEKNUM(DATE(2015,12,31))", 53},
	}

	for _, tc := range tests {
		cf := evalCompile(t, tc.formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", tc.formula, err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("ISOWEEKNUM: got type %v, want number for %s", got.Type, tc.formula)
		}
		if got.Num != tc.want {
			t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
		}
	}
}

func TestWORKDAY(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		expect  string // formula that produces expected result
	}{
		{"same_day", "WORKDAY(DATE(2008,10,1), 0)", "DATE(2008,10,1)"},
		{"friday_plus_one", "WORKDAY(DATE(2008,10,3), 1)", "DATE(2008,10,6)"},
		{"monday_minus_one", "WORKDAY(DATE(2008,10,6), -1)", "DATE(2008,10,3)"},
		{"five_workdays", "WORKDAY(DATE(2008,10,1), 5)", "DATE(2008,10,8)"},
		{"ten_workdays", "WORKDAY(DATE(2008,10,1), 10)", "DATE(2008,10,15)"},
		{"skip_christmas_weekend", "WORKDAY(DATE(2024,12,23), 5)", "DATE(2024,12,30)"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, nil, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tt.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("WORKDAY: got type %v, want number for %s", got.Type, tt.formula)
			}

			cfExpect := evalCompile(t, tt.expect)
			want, err := Eval(cfExpect, nil, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tt.expect, err)
			}

			if got.Num != want.Num {
				t.Errorf("%s = %g, want %g (from %s)", tt.formula, got.Num, want.Num, tt.expect)
			}
		})
	}
}

func TestYEARFRAC(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		tol     float64
		isErr   bool
		errVal  ErrorValue
	}{
		{"basis0", "YEARFRAC(DATE(2012,1,1), DATE(2012,7,30), 0)", 0.58056, 0.001, false, 0},
		{"basis1", "YEARFRAC(DATE(2012,1,1), DATE(2012,7,30), 1)", 0.57650, 0.001, false, 0},
		{"basis2", "YEARFRAC(DATE(2012,1,1), DATE(2012,7,30), 2)", 0.58611, 0.001, false, 0},
		{"basis3", "YEARFRAC(DATE(2012,1,1), DATE(2012,7,30), 3)", 0.57808, 0.001, false, 0},
		{"basis4", "YEARFRAC(DATE(2012,1,1), DATE(2012,7,30), 4)", 0.58056, 0.001, false, 0},
		{"same_date", "YEARFRAC(DATE(2006,1,1), DATE(2006,1,1))", 0, 0, false, 0},
		{"invalid_basis", "YEARFRAC(DATE(2006,1,1), DATE(2006,1,1), 5)", 0, 0, true, ErrValNUM},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if tc.isErr {
				if got.Type != ValueError || got.Err != tc.errVal {
					t.Errorf("%s: got %v, want error %v", tc.formula, got, tc.errVal)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: got type %v, want number", tc.formula, got.Type)
			}
			if math.Abs(got.Num-tc.want) > tc.tol {
				t.Errorf("%s = %g, want %g (tolerance %g)", tc.formula, got.Num, tc.want, tc.tol)
			}
		})
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
