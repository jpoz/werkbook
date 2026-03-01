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

func TestDAYSErrorPropagation(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		wantErr ErrorValue
	}{
		// DATE(-0.5,1,1) should return #NUM!, which DAYS should propagate
		{"negative_year_start", "DAYS(100, DATE(-0.5,1,1))", ErrValNUM},
		{"negative_year_end", "DAYS(DATE(-0.5,1,1), 100)", ErrValNUM},
		// DATE(10000,1,1) → #NUM!
		{"year_too_large", "DAYS(DATE(10000,1,1), 100)", ErrValNUM},
		// Tiny negative year that truncates to 0 but is still < 0
		{"tiny_negative_year", "DAYS(DATE(-1e-15,1,1), 100)", ErrValNUM},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if got.Type != ValueError || got.Err != tc.wantErr {
				t.Errorf("%s: got type=%v err=%v, want error %v", tc.formula, got.Type, got.Err, tc.wantErr)
			}
		})
	}
}

func TestDAYS360(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"full_year_us", "DAYS360(DATE(2024,1,1), DATE(2024,12,31))", 360},
		{"jan30_to_feb28", "DAYS360(DATE(2024,1,30), DATE(2024,2,28))", 28},
		{"half_year", "DAYS360(DATE(2024,1,1), DATE(2024,7,1))", 180},
		{"us_31_to_31", "DAYS360(DATE(2024,1,31), DATE(2024,3,31))", 60},
		{"european_31_to_31", "DAYS360(DATE(2024,1,31), DATE(2024,3,31), TRUE)", 60},
		{"negative_result", "DAYS360(DATE(2024,2,28), DATE(2024,1,1))", -57},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("DAYS360: got type %v, want number for %s", got.Type, tc.formula)
			}
			if got.Num != tc.want {
				t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
			}
		})
	}
}

func TestWORKDAYEdgeCases(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		expect  string
	}{
		// Negative workdays
		{"negative_five", "WORKDAY(DATE(2008,10,15), -5)", "DATE(2008,10,8)"},
		{"negative_ten", "WORKDAY(DATE(2008,10,15), -10)", "DATE(2008,10,1)"},
		// Crossing weekends backward
		{"negative_across_weekend", "WORKDAY(DATE(2008,10,6), -2)", "DATE(2008,10,2)"},

		// Start on weekend, zero workdays (returns start_date as-is)
		{"saturday_zero", "WORKDAY(DATE(2008,10,4), 0)", "DATE(2008,10,4)"},
		{"sunday_zero", "WORKDAY(DATE(2008,10,5), 0)", "DATE(2008,10,5)"},

		// Large number of workdays (roughly 1 year = ~260 workdays)
		{"one_year_workdays", "WORKDAY(DATE(2008,1,1), 260)", "DATE(2008,12,30)"},

		// Backward from Monday across multiple weekends
		{"backward_many_weeks", "WORKDAY(DATE(2008,10,20), -15)", "DATE(2008,9,29)"},

		// Forward from Friday across weekend
		{"friday_plus_two", "WORKDAY(DATE(2008,10,3), 2)", "DATE(2008,10,7)"},
		{"friday_plus_five", "WORKDAY(DATE(2008,10,3), 5)", "DATE(2008,10,10)"},

		// Backward from Monday
		{"monday_minus_five", "WORKDAY(DATE(2008,10,6), -5)", "DATE(2008,9,29)"},

		// Start on weekend going backward across another weekend
		{"sunday_minus_six", "WORKDAY(DATE(2008,10,5), -6)", "DATE(2008,9,26)"},

		// Very small number: 1 workday forward from midweek
		{"tuesday_plus_one", "WORKDAY(DATE(2008,9,30), 1)", "DATE(2008,10,1)"},

		// Negative one from a Friday gives Thursday
		{"friday_minus_one", "WORKDAY(DATE(2008,10,3), -1)", "DATE(2008,10,2)"},

		// Large backward count: 260 workdays back from Dec 31, 2008 (Wed) = Jan 2, 2008 (Wed)
		{"large_backward_260", "WORKDAY(DATE(2008,12,31), -260)", "DATE(2008,1,2)"},

		// Forward across multiple months: Oct 30 (Thu) + 10 = Nov 13 (Thu)
		{"across_multiple_months", "WORKDAY(DATE(2008,10,30), 10)", "DATE(2008,11,13)"},

		// Near epoch: Jan 1, 1900 (Mon) + 5 = Jan 8, 1900 (Mon)
		{"near_epoch_forward", "WORKDAY(DATE(1900,1,1), 5)", "DATE(1900,1,8)"},

		// Near epoch backward: Jan 8, 1900 (Mon) - 5 = Jan 1, 1900 (Mon)
		{"near_epoch_backward", "WORKDAY(DATE(1900,1,8), -5)", "DATE(1900,1,1)"},

		// Saturday + 5 workdays: Oct 4 (Sat) → Oct 10 (Fri)
		{"saturday_plus_five", "WORKDAY(DATE(2008,10,4), 5)", "DATE(2008,10,10)"},

		// === Zero serial (serial 0 = fictitious Jan 0, 1900 = Dec 31, 1899, Sunday) ===

		// Zero serial + 0 days → returns 0 unchanged
		{"zero_serial_zero_days", "WORKDAY(0, 0)", "DATE(1900,1,1)-1"},
		// Zero serial + 1: Sun Dec 31, 1899 → step to Mon Jan 1 = serial 1
		{"zero_serial_plus_one", "WORKDAY(0, 1)", "DATE(1900,1,1)"},
		// Zero serial + 5: skips to Fri Jan 5 = serial 5
		{"zero_serial_plus_five", "WORKDAY(0, 5)", "DATE(1900,1,5)"},

		// === Leap day boundary ===

		// Feb 29, 2024 (Thursday) + 0 = same day
		{"leap_day_same_day", "WORKDAY(DATE(2024,2,29), 0)", "DATE(2024,2,29)"},
		// Feb 29, 2024 (Thursday) + 1 = Fri March 1
		{"leap_day_plus_one", "WORKDAY(DATE(2024,2,29), 1)", "DATE(2024,3,1)"},
		// Feb 29, 2024 (Thursday) - 1 = Wed Feb 28
		{"leap_day_minus_one", "WORKDAY(DATE(2024,2,29), -1)", "DATE(2024,2,28)"},

		// === Sunday start with various day counts ===

		// Sunday + 1: Oct 5, 2008 (Sun) + 1 = Mon Oct 6
		{"sunday_plus_one", "WORKDAY(DATE(2008,10,5), 1)", "DATE(2008,10,6)"},
		// Sunday - 1: Oct 5, 2008 (Sun) - 1 = Fri Oct 3
		{"sunday_minus_one", "WORKDAY(DATE(2008,10,5), -1)", "DATE(2008,10,3)"},
		// Saturday + 1: Oct 4, 2008 (Sat) + 1 = Mon Oct 6
		{"saturday_plus_one", "WORKDAY(DATE(2008,10,4), 1)", "DATE(2008,10,6)"},
		// Saturday - 1: Oct 4, 2008 (Sat) - 1 = Fri Oct 3
		{"saturday_minus_one", "WORKDAY(DATE(2008,10,4), -1)", "DATE(2008,10,3)"},

		// === Very large workday counts ===

		// ~4 years ≈ 1040 workdays forward
		{"four_years_workdays", "WORKDAY(DATE(2008,1,1), 1040)", "DATE(2011,12,27)"},

		// === Far-future dates ===

		// Year 9999: Dec 30, 9999 (Thu) + 1 = Dec 31, 9999 (Fri)
		{"far_future_plus_one", "WORKDAY(DATE(9999,12,30), 1)", "DATE(9999,12,31)"},

		// === Leap year boundary with larger workday counts ===
		// Mon Jan 8, 2024 + 38 workdays:
		// +5/week: Jan 15, 22, 29, Feb 5, 12, 19, 26 (Mon) = +35
		// +36 = Feb 27 (Tue), +37 = Feb 28 (Wed), +38 = Feb 29 (Thu)
		{"leap_year_lands_on_feb29", "WORKDAY(DATE(2024,1,8), 38)", "DATE(2024,2,29)"},

		// Non-leap year: same start but 2023, Jan 9 is Monday
		// +35 = Feb 27 (Mon), +36 = Feb 28 (Tue), +37 = Mar 1 (Wed), +38 = Mar 2 (Thu)
		{"non_leap_year_skips_feb29", "WORKDAY(DATE(2023,1,9), 38)", "DATE(2023,3,2)"},

		// Saturday + 10 workdays: Oct 4, 2008 (Sat) → Oct 17 (Fri)
		{"saturday_plus_ten", "WORKDAY(DATE(2008,10,4), 10)", "DATE(2008,10,17)"},

		// Sunday backward 10 workdays: Oct 5, 2008 (Sun) → Sep 22 (Mon)
		{"sunday_minus_ten", "WORKDAY(DATE(2008,10,5), -10)", "DATE(2008,9,22)"},

		// Start mid-week with large negative: Oct 15, 2008 (Wed) - 20 = Sep 17 (Wed)
		{"wednesday_minus_twenty", "WORKDAY(DATE(2008,10,15), -20)", "DATE(2008,9,17)"},
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
		// Basic forward/backward operations
		{"same_day", "WORKDAY(DATE(2008,10,1), 0)", "DATE(2008,10,1)"},
		{"friday_plus_one", "WORKDAY(DATE(2008,10,3), 1)", "DATE(2008,10,6)"},
		{"monday_minus_one", "WORKDAY(DATE(2008,10,6), -1)", "DATE(2008,10,3)"},
		{"five_workdays", "WORKDAY(DATE(2008,10,1), 5)", "DATE(2008,10,8)"},
		{"ten_workdays", "WORKDAY(DATE(2008,10,1), 10)", "DATE(2008,10,15)"},
		{"skip_christmas_weekend", "WORKDAY(DATE(2024,12,23), 5)", "DATE(2024,12,30)"},

		// Documentation example: 151 workdays from Oct 1, 2008 = Apr 30, 2009
		{"doc_example_151_days", "WORKDAY(DATE(2008,10,1), 151)", "DATE(2009,4,30)"},

		// Starting on weekend days
		{"start_saturday_plus_one", "WORKDAY(DATE(2008,10,4), 1)", "DATE(2008,10,6)"},    // Sat → Mon
		{"start_sunday_plus_one", "WORKDAY(DATE(2008,10,5), 1)", "DATE(2008,10,6)"},      // Sun → Mon
		{"start_saturday_minus_one", "WORKDAY(DATE(2008,10,4), -1)", "DATE(2008,10,3)"},  // Sat → Fri
		{"start_sunday_minus_one", "WORKDAY(DATE(2008,10,5), -1)", "DATE(2008,10,3)"},    // Sun → Fri

		// Crossing multiple weekends
		{"twenty_workdays", "WORKDAY(DATE(2008,10,1), 20)", "DATE(2008,10,29)"},
		{"across_month_boundary", "WORKDAY(DATE(2008,10,29), 5)", "DATE(2008,11,5)"},

		// Fractional days should be truncated (per Excel docs)
		{"fractional_days_positive", "WORKDAY(DATE(2008,10,1), 1.9)", "DATE(2008,10,2)"},
		{"fractional_days_negative", "WORKDAY(DATE(2008,10,3), -1.7)", "DATE(2008,10,2)"},

		// Single workday with single holiday (third argument as single date)
		{"single_holiday", "WORKDAY(DATE(2008,10,1), 1, DATE(2008,10,2))", "DATE(2008,10,3)"},
		// Holiday on a weekend should have no effect
		{"holiday_on_weekend", "WORKDAY(DATE(2008,10,1), 1, DATE(2008,10,4))", "DATE(2008,10,2)"},

		// Year boundary crossing
		{"cross_year_forward", "WORKDAY(DATE(2008,12,29), 5)", "DATE(2009,1,5)"},
		{"cross_year_backward", "WORKDAY(DATE(2009,1,5), -5)", "DATE(2008,12,29)"},

		// Leap year: Feb 29, 2024 is a Thursday
		{"leap_year_crossing", "WORKDAY(DATE(2024,2,28), 1)", "DATE(2024,2,29)"},
		{"leap_year_crossing_2", "WORKDAY(DATE(2024,2,28), 2)", "DATE(2024,3,1)"},

		// One workday from Wednesday = Thursday
		{"wednesday_plus_one", "WORKDAY(DATE(2008,10,1), 1)", "DATE(2008,10,2)"},
		// Thursday + 1 = Friday
		{"thursday_plus_one", "WORKDAY(DATE(2008,10,2), 1)", "DATE(2008,10,3)"},

		// Type coercion: boolean as days (TRUE=1, FALSE=0)
		{"bool_true_as_days", "WORKDAY(DATE(2008,10,1), TRUE)", "DATE(2008,10,2)"},
		{"bool_false_as_days", "WORKDAY(DATE(2008,10,1), FALSE)", "DATE(2008,10,1)"},

		// Type coercion: numeric string as days
		{"string_numeric_days", `WORKDAY(DATE(2008,10,1), "5")`, "DATE(2008,10,8)"},

		// Fractional days: truncated (not rounded) per documentation "If days is not an integer, it is truncated."
		{"fractional_0_point_9_truncates_to_0", "WORKDAY(DATE(2008,10,1), 0.9)", "DATE(2008,10,1)"},
		{"fractional_0_point_1_truncates_to_0", "WORKDAY(DATE(2008,10,1), 0.1)", "DATE(2008,10,1)"},
		{"fractional_4_point_99_truncates_to_4", "WORKDAY(DATE(2008,10,1), 4.99)", "DATE(2008,10,7)"},
		{"fractional_neg_0_point_9_truncates_to_0", "WORKDAY(DATE(2008,10,3), -0.9)", "DATE(2008,10,3)"},

		// Y2K boundary crossing
		{"y2k_crossing", "WORKDAY(DATE(1999,12,31), 1)", "DATE(2000,1,3)"},

		// New year crossing 2023→2024 (Dec 29 is Friday)
		{"new_year_crossing_2024", "WORKDAY(DATE(2023,12,29), 2)", "DATE(2024,1,2)"},

		// Large forward count (2 years ≈ 520 workdays)
		{"two_years_workdays", "WORKDAY(DATE(2008,1,1), 520)", "DATE(2009,12,29)"},

		// === Tests locking in DATE floor-truncation fix (mismatch scenario) ===

		// Exact mismatch case: DATE(2024,-0.5,3.14...) uses Floor(-0.5)=-1 for month
		// and Floor(3.14)=3 for day → DATE(2024,-1,3) = Nov 3, 2023 (Fri).
		// Days=0.5 truncates to 0 → returns start_date unchanged.
		{"mismatch_neg_frac_month_frac_days", "WORKDAY(DATE(2024,-0.5,3.14159265358979), 0.5)", "DATE(2023,11,3)"},

		// DATE with month=0 → December of previous year; Dec 15, 2023 is Friday
		{"date_month_zero_forward", "WORKDAY(DATE(2024,0,15), 1)", "DATE(2023,12,18)"},

		// DATE with negative integer month: DATE(2024,-1,3) = Nov 3, 2023 (Fri) + 1 = Nov 6 (Mon)
		{"date_neg_int_month_forward", "WORKDAY(DATE(2024,-1,3), 1)", "DATE(2023,11,6)"},

		// DATE with positive fractional month/day: Floor(1.9)=1, Floor(15.7)=15
		// → Jan 15, 2024 (Mon) + 1 = Jan 16 (Tue)
		{"date_pos_frac_month_day", "WORKDAY(DATE(2024,1.9,15.7), 1)", "DATE(2024,1,16)"},

		// Fractional days = exactly 0.5, mid-week start (Mon Jan 15, 2024)
		{"fractional_half_day_midweek", "WORKDAY(DATE(2024,1,15), 0.5)", "DATE(2024,1,15)"},

		// Fractional days = -0.5 truncates to 0 → same as start
		{"fractional_neg_half_day", "WORKDAY(DATE(2024,1,15), -0.5)", "DATE(2024,1,15)"},

		// Raw serial number as start_date (39722 = Oct 1, 2008, Wed)
		{"raw_serial_start_date", "WORKDAY(39722, 1)", "DATE(2008,10,2)"},

		// Near-epoch: serial 1 = Jan 1, 1900 (Mon) + 1 = Jan 2, 1900 (Tue)
		{"serial_one_plus_one", "WORKDAY(1, 1)", "DATE(1900,1,2)"},

		// === Boundary-valid results near epoch (complement of result-goes-negative tests) ===

		// Jan 3 (Wed, serial 3) - 2 workdays: Wed→Tue(2)→Mon(1) = serial 1 (valid!)
		{"backward_to_serial_1", "WORKDAY(DATE(1900,1,3), -2)", "DATE(1900,1,1)"},
		// Jan 2 (Tue, serial 2) - 1 workday = Mon Jan 1 (serial 1, valid)
		{"backward_to_serial_1_from_tue", "WORKDAY(2, -1)", "DATE(1900,1,1)"},
		// Jan 5 (Fri, serial 5) - 4 workdays: Fri→Thu(4)→Wed(3)→Tue(2)→Mon(1) = valid
		{"backward_to_serial_1_from_fri", "WORKDAY(5, -4)", "DATE(1900,1,1)"},

		// === Additional type coercion tests ===

		// Boolean TRUE as start_date (TRUE=1 → serial 1 = Jan 1, 1900 Mon)
		{"bool_true_as_start_date", "WORKDAY(TRUE, 1)", "DATE(1900,1,2)"},
		// String numeric start_date
		{"string_serial_as_start_date", `WORKDAY("39722", 1)`, "DATE(2008,10,2)"},
		// Fractional serial as start_date (39722.75 = Oct 1, 2008 at 6pm)
		// WORKDAY preserves the time component, so result has fractional part too
		{"fractional_serial_start_date", "WORKDAY(39722.75, 1)", "DATE(2008,10,2)+0.75"},

		// === Documentation example verification ===

		// From docs: WORKDAY(DATE(2008,10,1), 151) = 4/30/2009
		// Already tested above as doc_example_151_days, but verify with raw serial too
		// DATE(2008,10,1) = serial 39722
		{"doc_example_raw_serial", "WORKDAY(39722, 151)", "DATE(2009,4,30)"},

		// === Additional coverage for negative serial fix ===

		// Boolean FALSE as start_date (FALSE=0 → serial 0) + 0 = serial 0
		{"bool_false_as_start_date_zero", "WORKDAY(FALSE, 0)", "DATE(1900,1,1)-1"},
		// Boolean FALSE as start_date + 1 = serial 1 (Mon Jan 1, 1900)
		{"bool_false_as_start_date_plus_one", "WORKDAY(FALSE, 1)", "DATE(1900,1,1)"},
		// String "0" as start_date (coerced to serial 0) + 1 = serial 1
		{"string_zero_as_start_date", `WORKDAY("0", 1)`, "DATE(1900,1,1)"},
		// String negative days: "-5" coerced to -5
		{"string_negative_days", `WORKDAY(DATE(2008,10,15), "-5")`, "DATE(2008,10,8)"},
		// Boolean TRUE as days from Friday: TRUE=1 → Fri + 1 = Mon
		{"bool_true_days_from_friday", "WORKDAY(DATE(2008,10,3), TRUE)", "DATE(2008,10,6)"},
		// Non-leap year crossing: Feb 28, 2023 (Tue) + 1 = Mar 1 (Wed), no Feb 29
		{"non_leap_year_feb_crossing", "WORKDAY(DATE(2023,2,28), 1)", "DATE(2023,3,1)"},
		// End of month: Mar 31, 2024 (Sun) + 1 = Mon Apr 1
		{"end_of_month_sunday", "WORKDAY(DATE(2024,3,31), 1)", "DATE(2024,4,1)"},
		// 1900 leap year bug: serial 59 = Feb 28, serial 60 = fictitious Feb 29, serial 61 = Mar 1
		// WORKDAY uses real dates internally, so Feb 28 + 1 workday = Mar 1 (skips fictitious Feb 29)
		{"excel_1900_serial_59_plus_1", "WORKDAY(59, 1)", "61"},

		// === Additional documentation-based coverage ===

		// Negative fractional days: -1.7 truncates to -1 (toward zero, not floor)
		{"negative_frac_truncates_toward_zero", "WORKDAY(DATE(2008,10,6), -1.7)", "DATE(2008,10,3)"},
		// Very small positive fractional truncates to 0
		{"tiny_positive_frac", "WORKDAY(DATE(2008,10,1), 0.0001)", "DATE(2008,10,1)"},

		// Start on Saturday going backward: Oct 4 (Sat) - 5 = Sep 29 (Mon)
		{"saturday_minus_five", "WORKDAY(DATE(2008,10,4), -5)", "DATE(2008,9,29)"},
		// Start on Sunday going forward across multiple weekends
		{"sunday_plus_ten", "WORKDAY(DATE(2008,10,5), 10)", "DATE(2008,10,17)"},

		// 0 days from weekday returns same date
		{"zero_days_tuesday", "WORKDAY(DATE(2008,10,7), 0)", "DATE(2008,10,7)"},

		// Crossing February in a non-leap year
		{"cross_feb_non_leap", "WORKDAY(DATE(2023,2,27), 2)", "DATE(2023,3,1)"},
		// Crossing February in a leap year
		{"cross_feb_leap", "WORKDAY(DATE(2024,2,27), 2)", "DATE(2024,2,29)"},

		// Large negative workday count (~2 years back)
		{"two_years_backward", "WORKDAY(DATE(2010,1,1), -520)", "DATE(2008,1,4)"},

		// Backward from end of year: Dec 31, 2025 (Wed) - 1 = Dec 30, 2025 (Tue)
		{"year_end_backward", "WORKDAY(DATE(2025,12,31), -1)", "DATE(2025,12,30)"},

		// Wednesday + exactly 5 = next Wednesday (skipping one weekend)
		{"wednesday_plus_five", "WORKDAY(DATE(2008,10,1), 5)", "DATE(2008,10,8)"},

		// Multiple of 5 workdays = exact weeks
		{"fifteen_workdays", "WORKDAY(DATE(2008,10,1), 15)", "DATE(2008,10,22)"},

		// Negative string days via coercion
		{"string_neg_frac_days", `WORKDAY(DATE(2008,10,6), "-1.9")`, "DATE(2008,10,3)"},

		// === Documentation example: reverse verification ===
		// WORKDAY(Apr 30 2009, -151) should return Oct 1 2008
		{"doc_example_151_backward", "WORKDAY(DATE(2009,4,30), -151)", "DATE(2008,10,1)"},

		// === Additional documentation-based coverage ===
		// 50 workdays forward from Oct 1 2008 (Wed)
		// 50 workdays = 10 weeks → Dec 10, 2008 (Wed)
		{"fifty_workdays_forward", "WORKDAY(DATE(2008,10,1), 50)", "DATE(2008,12,10)"},

		// WORKDAY preserves time component of start_date (fractional serial)
		// Serial 39722.5 = Oct 1, 2008 noon; +1 workday = Oct 2 noon
		{"preserves_time_component", "WORKDAY(39722.5, 1)", "DATE(2008,10,2)+0.5"},

		// Very small days value truncates to 0
		{"days_1e_minus_10_truncates_to_0", "WORKDAY(DATE(2008,10,1), 1E-10)", "DATE(2008,10,1)"},

		// Negative fractional days just below zero truncates to 0
		{"days_neg_1e_minus_10_truncates_to_0", "WORKDAY(DATE(2008,10,1), -1E-10)", "DATE(2008,10,1)"},

		// Day count that crosses multiple months: 65 workdays = 13 weeks
		// Oct 1 (Wed) + 13 weeks = Dec 31 (Wed)
		{"sixty_five_workdays", "WORKDAY(DATE(2008,10,1), 65)", "DATE(2008,12,31)"},

		// Wednesday + 3 = Monday (crosses one weekend)
		{"wednesday_plus_three", "WORKDAY(DATE(2008,10,1), 3)", "DATE(2008,10,6)"},

		// Thursday + 2 = Monday (crosses one weekend)
		{"thursday_plus_two", "WORKDAY(DATE(2008,10,2), 2)", "DATE(2008,10,6)"},

		// === Empty string coercion: empty string coerces to 0 ===
		// Empty string as start_date → serial 0 + 5 workdays = serial 5 (Jan 5, 1900 Fri)
		{"empty_string_start_date", `WORKDAY("", 5)`, "DATE(1900,1,5)"},
		// Empty string as days → 0 days, returns start_date unchanged
		{"empty_string_days", `WORKDAY(DATE(2008,10,1), "")`, "DATE(2008,10,1)"},

		// === Truncation property verified via self-referential formulas ===
		// Per docs: "If days is not an integer, it is truncated."
		// Large fractional parts should not affect result
		{"truncation_pos_self_check", "WORKDAY(DATE(2024,6,3), 25.7)", "WORKDAY(DATE(2024,6,3), 25)"},
		{"truncation_neg_self_check", "WORKDAY(DATE(2024,6,3), -25.3)", "WORKDAY(DATE(2024,6,3), -25)"},

		// === Round-trip symmetry: forward then backward returns to start ===
		// WORKDAY(WORKDAY(start, n), -n) == start for weekday starts
		{"round_trip_30_days", "WORKDAY(WORKDAY(DATE(2024,1,8), 30), -30)", "DATE(2024,1,8)"},
		{"round_trip_100_days", "WORKDAY(WORKDAY(DATE(2024,3,4), 100), -100)", "DATE(2024,3,4)"},

		// === Year boundary: Dec 31, 2024 (Tue) + 1 = Jan 1, 2025 (Wed) ===
		{"cross_year_2024_to_2025", "WORKDAY(DATE(2024,12,31), 1)", "DATE(2025,1,1)"},

		// === 25 workdays (5 full work-weeks) from Monday = Monday ===
		// Mon Jan 8, 2024 + 25 workdays = Mon Feb 12, 2024
		{"five_full_work_weeks", "WORKDAY(DATE(2024,1,8), 25)", "DATE(2024,2,12)"},

		// === Both args empty string: both coerce to 0 ===
		{"both_empty_strings", `WORKDAY("", "")`, "DATE(1900,1,1)-1"},

		// === Max serial boundary: valid cases using raw serials ===
		// Serial 2958465 = Dec 31, 9999, the last valid Excel date
		// 0 days → returns same serial unchanged
		{"max_serial_zero_days", "WORKDAY(2958465, 0)", "2958465"},

		// === Additional practical tests ===
		// String "1.5" as start → serial 1.5 (Jan 1, 1900 noon) + 1 = Jan 2 noon
		{"string_frac_start", `WORKDAY("1.5", 1)`, "DATE(1900,1,2)+0.5"},
		// Single inline numeric value as holiday (third arg as number, not range)
		// Oct 1 (Wed) + 1 workday with Oct 2 (serial 39723) as holiday → Oct 3 (Fri)
		{"inline_numeric_holiday", "WORKDAY(DATE(2008,10,1), 1, 39723)", "DATE(2008,10,3)"},
		// Fractional start (0.25 = 6am Jan 0, 1900) + 1 = serial 1.25
		{"fractional_start_quarter_day", "WORKDAY(0.25, 1)", "DATE(1900,1,1)+0.25"},
		// Monday + 1 = Tuesday (basic weekday progression)
		{"monday_plus_one", "WORKDAY(DATE(2008,10,6), 1)", "DATE(2008,10,7)"},
		// Tuesday + 1 = Wednesday
		{"tuesday_plus_one_explicit", "WORKDAY(DATE(2008,10,7), 1)", "DATE(2008,10,8)"},
		// Wednesday + 1 = Thursday (already tested above but this is explicit)
		// Friday + 1 = Monday (crosses weekend, already tested above)
		// Verify symmetry: WORKDAY(date, n) forward then -n backward = date
		// Using a date in a different year (2020) for diversity
		{"round_trip_15_days_2020", "WORKDAY(WORKDAY(DATE(2020,6,15), 15), -15)", "DATE(2020,6,15)"},
		// DST boundary crossing: March 8, 2020 (Sun) + 1 = Mon March 9
		{"dst_spring_forward_sunday", "WORKDAY(DATE(2020,3,8), 1)", "DATE(2020,3,9)"},
		// DST boundary: Nov 1, 2020 (Sun) + 1 = Mon Nov 2
		{"dst_fall_back_sunday", "WORKDAY(DATE(2020,11,1), 1)", "DATE(2020,11,2)"},

		// === Mismatch lock-in: DATE floor-truncation flowing into WORKDAY ===
		// DATE(2024,-0.5,3.14...) uses Floor(-0.5)=-1 for month →
		// DATE(2024,-1,3) = Nov 3, 2023 (Fri, serial 45233).
		// Verify WORKDAY uses the correct serial with various day counts.
		{"mismatch_serial_45233_zero_days", "WORKDAY(45233, 0)", "45233"},
		{"mismatch_serial_45233_plus_one", "WORKDAY(45233, 1)", "DATE(2023,11,6)"},
		{"mismatch_serial_45233_plus_five", "WORKDAY(45233, 5)", "DATE(2023,11,10)"},
		{"mismatch_serial_45233_minus_one", "WORKDAY(45233, -1)", "DATE(2023,11,2)"},
		// DATE with negative fractional month + positive workdays (not just truncated-to-zero days)
		{"date_neg_frac_month_plus_three", "WORKDAY(DATE(2024,-0.5,3.14159265358979), 3)", "DATE(2023,11,8)"},
		{"date_neg_frac_month_plus_five", "WORKDAY(DATE(2024,-0.5,3.14159265358979), 5)", "DATE(2023,11,10)"},
		{"date_neg_frac_month_minus_five", "WORKDAY(DATE(2024,-0.5,3.14159265358979), -5)", "DATE(2023,10,27)"},

		// === Time preservation across weekends ===
		// Fri at noon + 1 = Mon at noon (time component preserved)
		{"time_preserved_across_weekend_fwd", "WORKDAY(DATE(2008,10,3)+0.5, 1)", "DATE(2008,10,6)+0.5"},
		// Mon at noon - 1 = Fri at noon
		{"time_preserved_across_weekend_bwd", "WORKDAY(DATE(2008,10,6)+0.5, -1)", "DATE(2008,10,3)+0.5"},
		// Time preserved across multiple weekends
		{"time_preserved_multi_weekend", "WORKDAY(DATE(2008,10,3)+0.25, 10)", "DATE(2008,10,17)+0.25"},

		// === DATE with month floor causing year wrap ===
		// DATE(2024,0.9,15): Floor(0.9)=0 → Dec 15, 2023 (Fri)
		{"date_month_floor_to_zero", "WORKDAY(DATE(2024,0.9,15), 1)", "DATE(2023,12,18)"},
		// DATE(2024,-13,1): month=-13 → Nov 1, 2022 (Tue) + 1 = Nov 2, 2022 (Wed)
		{"date_large_neg_month_wrap", "WORKDAY(DATE(2024,-13,1), 1)", "DATE(2022,11,2)"},

		// === Days truncation verification (INT not FLOOR for negative) ===
		// -2.1 truncates to -2: Mon Oct 6 - 2 = Thu Oct 2
		{"days_neg_2_1_truncates_to_neg_2", "WORKDAY(DATE(2008,10,6), -2.1)", "DATE(2008,10,2)"},
		// -4.999 truncates to -4
		{"days_neg_4_999_truncates_to_neg_4", "WORKDAY(DATE(2008,10,8), -4.999)", "DATE(2008,10,2)"},
		// 999.99 truncates to 999 (large fractional)
		{"large_fractional_days", "WORKDAY(DATE(2008,10,1), 999.99)", "WORKDAY(DATE(2008,10,1), 999)"},

		// === Result at max serial boundary (using DATE for proper date handling) ===
		// Dec 29, 9999 (Wed) + 2 = Dec 31, 9999 (Fri)
		{"result_near_max_serial_via_date", "WORKDAY(DATE(9999,12,29), 2)", "DATE(9999,12,31)"},

		// === Every day of week as start_date (completeness) ===
		// Mon Oct 6 + 1 = Tue Oct 7
		{"start_monday", "WORKDAY(DATE(2008,10,6), 1)", "DATE(2008,10,7)"},
		// Tue Oct 7 + 1 = Wed Oct 8
		{"start_tuesday", "WORKDAY(DATE(2008,10,7), 1)", "DATE(2008,10,8)"},
		// Wed Oct 8 + 1 = Thu Oct 9
		{"start_wednesday", "WORKDAY(DATE(2008,10,8), 1)", "DATE(2008,10,9)"},
		// Thu Oct 9 + 1 = Fri Oct 10
		{"start_thursday", "WORKDAY(DATE(2008,10,9), 1)", "DATE(2008,10,10)"},
		// Fri Oct 10 + 1 = Mon Oct 13 (skips weekend)
		{"start_friday", "WORKDAY(DATE(2008,10,10), 1)", "DATE(2008,10,13)"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, nil, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tt.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("WORKDAY: got type %v (%v), want number for %s", got.Type, got, tt.formula)
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

func TestWORKDAYWithHolidays(t *testing.T) {
	// Test with holiday range via mockResolver
	// Set up holidays: Nov 26, Dec 4, Jan 21 (from documentation example)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// Holidays in A1:A3
			{Col: 1, Row: 1}: NumberVal(39778), // DATE(2008,11,26) = 39778
			{Col: 1, Row: 2}: NumberVal(39786), // DATE(2008,12,4) = 39786
			{Col: 1, Row: 3}: NumberVal(39834), // DATE(2009,1,21) = 39834
		},
	}

	// Documentation example: WORKDAY(DATE(2008,10,1), 151, holidays) = DATE(2009,5,5)
	t.Run("doc_example_with_holidays", func(t *testing.T) {
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 151, A1:A3)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2009,5,5)")
		want, err := Eval(cfExpect, nil, nil)
		if err != nil {
			t.Fatalf("Eval expect: %v", err)
		}
		if got.Num != want.Num {
			t.Errorf("WORKDAY with holidays = %g, want %g", got.Num, want.Num)
		}
	})

	// Holiday that falls exactly on what would be the result
	t.Run("holiday_shifts_result", func(t *testing.T) {
		// Oct 2, 2008 is Thursday; set it as holiday
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(39723), // DATE(2008,10,2) = serial
			},
		}
		// Actually compute the serial for Oct 2 2008
		cfOct2 := evalCompile(t, "DATE(2008,10,2)")
		oct2, _ := Eval(cfOct2, nil, nil)
		r.cells[CellAddr{Col: 1, Row: 1}] = NumberVal(oct2.Num)

		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 1, A1:A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,3)")
		want, err := Eval(cfExpect, nil, nil)
		if err != nil {
			t.Fatalf("Eval expect: %v", err)
		}
		if got.Num != want.Num {
			t.Errorf("WORKDAY skip holiday = %g, want %g", got.Num, want.Num)
		}
	})

	// Consecutive holidays
	t.Run("consecutive_holidays", func(t *testing.T) {
		cfOct2 := evalCompile(t, "DATE(2008,10,2)")
		oct2, _ := Eval(cfOct2, nil, nil)
		cfOct3 := evalCompile(t, "DATE(2008,10,3)")
		oct3, _ := Eval(cfOct3, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(oct2.Num),
				{Col: 1, Row: 2}: NumberVal(oct3.Num),
			},
		}
		// Wed Oct 1 + 1 workday, skipping Thu Oct 2 and Fri Oct 3 → Mon Oct 6
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 1, A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,6)")
		want, err := Eval(cfExpect, nil, nil)
		if err != nil {
			t.Fatalf("Eval expect: %v", err)
		}
		if got.Num != want.Num {
			t.Errorf("WORKDAY consecutive holidays = %g, want %g", got.Num, want.Num)
		}
	})

	// Backward counting with a holiday
	t.Run("backward_with_holiday", func(t *testing.T) {
		// Oct 7, 2008 (Tuesday) is a holiday
		cfOct7 := evalCompile(t, "DATE(2008,10,7)")
		oct7, _ := Eval(cfOct7, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(oct7.Num),
			},
		}
		// Oct 8 (Wed) - 1 workday, skipping Oct 7 (Tue, holiday) → Oct 6 (Mon)
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,8), -1, A1:A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,6)")
		want, err := Eval(cfExpect, nil, nil)
		if err != nil {
			t.Fatalf("Eval expect: %v", err)
		}
		if got.Num != want.Num {
			t.Errorf("WORKDAY backward with holiday = %g, want %g", got.Num, want.Num)
		}
	})

	// Holiday on weekend has no extra effect
	t.Run("holiday_on_weekend_no_effect", func(t *testing.T) {
		// Oct 4 (Sat) and Oct 7 (Tue) as holidays
		cfOct4 := evalCompile(t, "DATE(2008,10,4)")
		oct4, _ := Eval(cfOct4, nil, nil)
		cfOct7 := evalCompile(t, "DATE(2008,10,7)")
		oct7, _ := Eval(cfOct7, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(oct4.Num),
				{Col: 1, Row: 2}: NumberVal(oct7.Num),
			},
		}
		// Oct 1 (Wed) + 5 workdays:
		// Without holidays: Oct 2,3,6,7,8 → Oct 8
		// With Oct 4 (Sat, already weekend) and Oct 7 (Tue, holiday):
		// Oct 2,3,6,8,9 → Oct 9
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 5, A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,9)")
		want, err := Eval(cfExpect, nil, nil)
		if err != nil {
			t.Fatalf("Eval expect: %v", err)
		}
		if got.Num != want.Num {
			t.Errorf("WORKDAY with weekend holiday = %g, want %g", got.Num, want.Num)
		}
	})

	// Entire work week (Mon-Fri) as holidays forces jump over full week
	t.Run("entire_week_as_holidays", func(t *testing.T) {
		cfMon := evalCompile(t, "DATE(2008,10,6)")
		mon, _ := Eval(cfMon, nil, nil)
		cfTue := evalCompile(t, "DATE(2008,10,7)")
		tue, _ := Eval(cfTue, nil, nil)
		cfWed := evalCompile(t, "DATE(2008,10,8)")
		wed, _ := Eval(cfWed, nil, nil)
		cfThu := evalCompile(t, "DATE(2008,10,9)")
		thu, _ := Eval(cfThu, nil, nil)
		cfFri := evalCompile(t, "DATE(2008,10,10)")
		fri, _ := Eval(cfFri, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(mon.Num),
				{Col: 1, Row: 2}: NumberVal(tue.Num),
				{Col: 1, Row: 3}: NumberVal(wed.Num),
				{Col: 1, Row: 4}: NumberVal(thu.Num),
				{Col: 1, Row: 5}: NumberVal(fri.Num),
			},
		}
		// Oct 3 (Fri) + 1 workday, with Oct 6-10 all holidays
		// → skip weekend (Oct 4-5) and holiday week (Oct 6-10) → Oct 13 (Mon)
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,3), 1, A1:A5)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,13)")
		want, err := Eval(cfExpect, nil, nil)
		if err != nil {
			t.Fatalf("Eval expect: %v", err)
		}
		if got.Num != want.Num {
			t.Errorf("WORKDAY entire week holiday = %g, want %g", got.Num, want.Num)
		}
	})

	// Backward across a week of holidays
	t.Run("backward_across_holiday_week", func(t *testing.T) {
		cfMon := evalCompile(t, "DATE(2008,10,6)")
		mon, _ := Eval(cfMon, nil, nil)
		cfTue := evalCompile(t, "DATE(2008,10,7)")
		tue, _ := Eval(cfTue, nil, nil)
		cfWed := evalCompile(t, "DATE(2008,10,8)")
		wed, _ := Eval(cfWed, nil, nil)
		cfThu := evalCompile(t, "DATE(2008,10,9)")
		thu, _ := Eval(cfThu, nil, nil)
		cfFri := evalCompile(t, "DATE(2008,10,10)")
		fri, _ := Eval(cfFri, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(mon.Num),
				{Col: 1, Row: 2}: NumberVal(tue.Num),
				{Col: 1, Row: 3}: NumberVal(wed.Num),
				{Col: 1, Row: 4}: NumberVal(thu.Num),
				{Col: 1, Row: 5}: NumberVal(fri.Num),
			},
		}
		// Oct 13 (Mon) - 1 workday, with Oct 6-10 all holidays
		// → skip weekend (Oct 11-12) and holiday week (Oct 6-10) → Oct 3 (Fri)
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,13), -1, A1:A5)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,3)")
		want, err := Eval(cfExpect, nil, nil)
		if err != nil {
			t.Fatalf("Eval expect: %v", err)
		}
		if got.Num != want.Num {
			t.Errorf("WORKDAY backward across holiday week = %g, want %g", got.Num, want.Num)
		}
	})

	// Duplicate holidays should not double-count
	t.Run("duplicate_holidays", func(t *testing.T) {
		cfOct2 := evalCompile(t, "DATE(2008,10,2)")
		oct2, _ := Eval(cfOct2, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(oct2.Num),
				{Col: 1, Row: 2}: NumberVal(oct2.Num), // same holiday listed twice
			},
		}
		// Oct 1 (Wed) + 1 workday, Oct 2 is holiday (listed twice) → Oct 3 (Fri)
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 1, A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,3)")
		want, err := Eval(cfExpect, nil, nil)
		if err != nil {
			t.Fatalf("Eval expect: %v", err)
		}
		if got.Num != want.Num {
			t.Errorf("WORKDAY with duplicate holidays = %g, want %g", got.Num, want.Num)
		}
	})

	// Holiday on start_date (start_date itself is a holiday, days > 0)
	t.Run("holiday_on_start_date", func(t *testing.T) {
		cfOct1 := evalCompile(t, "DATE(2008,10,1)")
		oct1, _ := Eval(cfOct1, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(oct1.Num),
			},
		}
		// Oct 1 (Wed) + 0 workdays, Oct 1 is holiday → result is still Oct 1 (days=0 returns start_date)
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 0, A1:A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		if got.Num != oct1.Num {
			t.Errorf("WORKDAY with start=holiday, days=0 = %g, want %g", got.Num, oct1.Num)
		}
	})

	// Holiday that coincides with a weekend day has no extra effect (single inline holiday)
	t.Run("inline_holiday_on_saturday", func(t *testing.T) {
		// Oct 1 (Wed) + 5 with Saturday Oct 4 as holiday (already a weekend)
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 5, DATE(2008,10,4))")
		got, err := Eval(cf, nil, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,8)")
		want, err := Eval(cfExpect, nil, nil)
		if err != nil {
			t.Fatalf("Eval expect: %v", err)
		}
		if got.Num != want.Num {
			t.Errorf("WORKDAY with inline Saturday holiday = %g, want %g", got.Num, want.Num)
		}
	})

	// Multiple holidays spanning across a weekend
	t.Run("holidays_around_weekend", func(t *testing.T) {
		// Oct 3 (Fri) and Oct 6 (Mon) as holidays; Oct 4-5 is weekend
		cfFri := evalCompile(t, "DATE(2008,10,3)")
		fri, _ := Eval(cfFri, nil, nil)
		cfMon := evalCompile(t, "DATE(2008,10,6)")
		mon, _ := Eval(cfMon, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(fri.Num),
				{Col: 1, Row: 2}: NumberVal(mon.Num),
			},
		}
		// Oct 2 (Thu) + 1 workday: skip Fri (holiday), Sat/Sun (weekend), Mon (holiday) → Tue Oct 7
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,2), 1, A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,7)")
		want, err := Eval(cfExpect, nil, nil)
		if err != nil {
			t.Fatalf("Eval expect: %v", err)
		}
		if got.Num != want.Num {
			t.Errorf("WORKDAY holidays around weekend = %g, want %g", got.Num, want.Num)
		}
	})

	// Backward counting across holidays that span weekend
	t.Run("backward_across_holidays_around_weekend", func(t *testing.T) {
		// Oct 3 (Fri) and Oct 6 (Mon) as holidays
		cfFri := evalCompile(t, "DATE(2008,10,3)")
		fri, _ := Eval(cfFri, nil, nil)
		cfMon := evalCompile(t, "DATE(2008,10,6)")
		mon, _ := Eval(cfMon, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(fri.Num),
				{Col: 1, Row: 2}: NumberVal(mon.Num),
			},
		}
		// Oct 7 (Tue) - 1 workday: skip Mon (holiday), Sun/Sat (weekend), Fri (holiday) → Thu Oct 2
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,7), -1, A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,2)")
		want, err := Eval(cfExpect, nil, nil)
		if err != nil {
			t.Fatalf("Eval expect: %v", err)
		}
		if got.Num != want.Num {
			t.Errorf("WORKDAY backward across holidays = %g, want %g", got.Num, want.Num)
		}
	})

	// Empty cells in holiday range should be ignored
	t.Run("empty_cell_in_holiday_range", func(t *testing.T) {
		cfOct2 := evalCompile(t, "DATE(2008,10,2)")
		oct2, _ := Eval(cfOct2, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(oct2.Num),
				// A2 is not set → EmptyVal(), should be ignored
			},
		}
		// Oct 1 (Wed) + 1 workday, Oct 2 is holiday → Oct 3 (Fri)
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 1, A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,3)")
		want, err := Eval(cfExpect, nil, nil)
		if err != nil {
			t.Fatalf("Eval expect: %v", err)
		}
		if got.Num != want.Num {
			t.Errorf("WORKDAY with empty in holidays = %g, want %g", got.Num, want.Num)
		}
	})

	// Holidays in reverse chronological order should produce same result as sorted
	t.Run("holidays_reverse_order", func(t *testing.T) {
		// Same doc example holidays (Nov 26, Dec 4, Jan 21) but in reverse order
		cfNov26 := evalCompile(t, "DATE(2008,11,26)")
		nov26, _ := Eval(cfNov26, nil, nil)
		cfDec4 := evalCompile(t, "DATE(2008,12,4)")
		dec4, _ := Eval(cfDec4, nil, nil)
		cfJan21 := evalCompile(t, "DATE(2009,1,21)")
		jan21, _ := Eval(cfJan21, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(jan21.Num),  // Jan 21 first (reverse)
				{Col: 1, Row: 2}: NumberVal(dec4.Num),   // Dec 4 second
				{Col: 1, Row: 3}: NumberVal(nov26.Num),  // Nov 26 third
			},
		}
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 151, A1:A3)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		// Should produce same result as chronologically ordered holidays: DATE(2009,5,5)
		cfExpect := evalCompile(t, "DATE(2009,5,5)")
		want, _ := Eval(cfExpect, nil, nil)
		if got.Num != want.Num {
			t.Errorf("WORKDAY with reverse-ordered holidays = %g, want %g", got.Num, want.Num)
		}
	})

	// Multiple workdays forward with holiday on the landing date
	t.Run("holiday_on_landing_date_chains", func(t *testing.T) {
		// Oct 1 (Wed) + 2 workdays = Oct 3 (Fri). But if Oct 3 is a holiday,
		// result shifts to next workday = Oct 6 (Mon).
		cfOct3 := evalCompile(t, "DATE(2008,10,3)")
		oct3, _ := Eval(cfOct3, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(oct3.Num),
			},
		}
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 2, A1:A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,6)")
		want, _ := Eval(cfExpect, nil, nil)
		if got.Num != want.Num {
			t.Errorf("WORKDAY holiday on landing = %g, want %g", got.Num, want.Num)
		}
	})

	// Holiday on start_date with positive days: start moves to next workday first
	t.Run("holiday_on_start_plus_days", func(t *testing.T) {
		// Oct 1 (Wed) is a holiday, + 1 workday → Oct 2 (Thu)
		cfOct1 := evalCompile(t, "DATE(2008,10,1)")
		oct1, _ := Eval(cfOct1, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(oct1.Num),
			},
		}
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 1, A1:A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,2)")
		want, _ := Eval(cfExpect, nil, nil)
		if got.Num != want.Num {
			t.Errorf("WORKDAY start=holiday + 1 = %g, want %g", got.Num, want.Num)
		}
	})

	// Two weeks of consecutive holidays (10 weekday holidays)
	t.Run("two_weeks_consecutive_holidays", func(t *testing.T) {
		// Oct 6-10 and Oct 13-17 (Mon-Fri, Mon-Fri) are all holidays
		holidays := []string{
			"DATE(2008,10,6)", "DATE(2008,10,7)", "DATE(2008,10,8)", "DATE(2008,10,9)", "DATE(2008,10,10)",
			"DATE(2008,10,13)", "DATE(2008,10,14)", "DATE(2008,10,15)", "DATE(2008,10,16)", "DATE(2008,10,17)",
		}
		r := &mockResolver{cells: map[CellAddr]Value{}}
		for i, h := range holidays {
			cfH := evalCompile(t, h)
			v, _ := Eval(cfH, nil, nil)
			r.cells[CellAddr{Col: 1, Row: i + 1}] = NumberVal(v.Num)
		}
		// Oct 3 (Fri) + 1 workday with 2 weeks holidays → Oct 20 (Mon)
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,3), 1, A1:A10)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,20)")
		want, _ := Eval(cfExpect, nil, nil)
		if got.Num != want.Num {
			t.Errorf("WORKDAY with 2 weeks holidays = %g, want %g", got.Num, want.Num)
		}
	})

	// Error value in holiday range should be propagated
	t.Run("error_in_holiday_range_num", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNUM),
			},
		}
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 5, A1:A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("WORKDAY with #NUM! in holidays: got %v, want #NUM!", got)
		}
	})

	t.Run("error_in_holiday_range_value", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValVALUE),
			},
		}
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 5, A1:A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("WORKDAY with #VALUE! in holidays: got %v, want #VALUE!", got)
		}
	})

	// Error mixed among valid holidays: error takes precedence
	t.Run("error_mixed_with_valid_holidays", func(t *testing.T) {
		cfOct2 := evalCompile(t, "DATE(2008,10,2)")
		oct2, _ := Eval(cfOct2, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(oct2.Num),
				{Col: 1, Row: 2}: ErrorVal(ErrValDIV0),
			},
		}
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 5, A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("WORKDAY with error among holidays: got %v, want #DIV/0!", got)
		}
	})

	// Holiday on the corrected DATE result (mismatch lock-in)
	// DATE(2024,-0.5,3.14...) = Nov 3, 2023 (Fri). If Nov 3 is a holiday,
	// WORKDAY(Nov 3, 1) should skip to next workday = Nov 6 (Mon).
	t.Run("mismatch_date_result_with_holiday", func(t *testing.T) {
		cfNov3 := evalCompile(t, "DATE(2023,11,3)")
		nov3, _ := Eval(cfNov3, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(nov3.Num),
			},
		}
		// Start = DATE(2024,-0.5,3.14...) = Nov 3, 2023 (Fri).
		// Nov 3 is a holiday, so +1 workday: skip Sat/Sun → Mon Nov 6, then Nov 6 is workday #1
		cf := evalCompile(t, "WORKDAY(DATE(2024,-0.5,3.14159265358979), 1, A1:A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2023,11,6)")
		want, err := Eval(cfExpect, nil, nil)
		if err != nil {
			t.Fatalf("Eval expect: %v", err)
		}
		if got.Num != want.Num {
			t.Errorf("WORKDAY with holiday on mismatch date = %g, want %g", got.Num, want.Num)
		}
	})

	// Holidays far in the future have no effect on backward counting
	t.Run("future_holidays_no_effect_on_backward", func(t *testing.T) {
		cfDec31 := evalCompile(t, "DATE(2008,12,31)")
		dec31, _ := Eval(cfDec31, nil, nil)
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(dec31.Num),
			},
		}
		// Oct 3 (Fri) - 1 workday = Oct 2 (Thu), Dec 31 holiday is irrelevant
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,3), -1, A1:A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("got type %v (%v), want number", got.Type, got)
		}
		cfExpect := evalCompile(t, "DATE(2008,10,2)")
		want, _ := Eval(cfExpect, nil, nil)
		if got.Num != want.Num {
			t.Errorf("WORKDAY backward with future holiday = %g, want %g", got.Num, want.Num)
		}
	})
}

func TestWORKDAYErrors(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		errVal  ErrorValue
	}{
		// Error propagation: DATE with negative year returns #NUM!,
		// WORKDAY should propagate it (this was the original mismatch fix)
		{"num_error_from_date_neg_year", "WORKDAY(DATE(-1,-1,3), 0)", ErrValNUM},
		{"num_error_from_date_large_year", "WORKDAY(DATE(10000,1,1), 5)", ErrValNUM},

		// Invalid start_date: non-numeric string → #VALUE!
		{"value_error_string_start", `WORKDAY("hello", 5)`, ErrValVALUE},

		// Invalid days: non-numeric string → #VALUE!
		{"value_error_string_days", `WORKDAY(DATE(2008,10,1), "abc")`, ErrValVALUE},

		// Negative serial numbers → #NUM! (the core mismatch fix)
		// When start_date is a negative number, WORKDAY returns #NUM!
		{"negative_serial_minus_one", "WORKDAY(-1, 0)", ErrValNUM},
		{"negative_serial_minus_28", "WORKDAY(-28, 0)", ErrValNUM},
		{"negative_serial_large", "WORKDAY(-1000, 5)", ErrValNUM},
		{"negative_serial_small_fraction", "WORKDAY(-0.5, 0)", ErrValNUM},
		{"negative_serial_tiny", "WORKDAY(-1E-15, 0)", ErrValNUM},
		{"negative_serial_positive_days", "WORKDAY(-5, 10)", ErrValNUM},
		{"negative_serial_negative_days", "WORKDAY(-5, -10)", ErrValNUM},

		// DATE producing #NUM! that propagates through WORKDAY
		// DATE(0,0,3) → Dec 3, 1899 → serial < 0 → #NUM!
		{"date_month_zero_negative_serial", "WORKDAY(DATE(0, 0, 3), 0)", ErrValNUM},
		// Exact mismatch case from the fuzz test spec
		{"date_fractional_args_negative", "WORKDAY(DATE(0.5, 0.01, 3.14159265358979), 0.5)", ErrValNUM},
		// DATE with very negative day going before epoch
		{"date_negative_day_past_epoch", "WORKDAY(DATE(1900, 1, -40), 1)", ErrValNUM},

		// Non-numeric string as holiday → #VALUE!
		{"value_error_string_holiday", `WORKDAY(DATE(2008,10,1), 5, "hello")`, ErrValVALUE},

		// Both arguments invalid
		{"value_error_both_args_invalid", `WORKDAY("hello", "world")`, ErrValVALUE},

		// Result goes negative when counting backward past epoch → #NUM!
		// (per docs: "If start_date plus days yields an invalid date, WORKDAY returns the #NUM! error value.")
		// Jan 1, 1900 (Mon, serial 1) - 1 workday = Fri Dec 29, 1899 → negative serial
		{"result_negative_backward_from_epoch", "WORKDAY(DATE(1900,1,1), -1)", ErrValNUM},
		// Using raw serial 1
		{"result_negative_backward_serial_1", "WORKDAY(1, -1)", ErrValNUM},
		// Jan 3 (Wed, serial 3) - 3 workdays: Wed→Tue(2)→Mon(1)→Fri(Dec 29) → negative
		{"result_negative_backward_from_serial_3", "WORKDAY(3, -3)", ErrValNUM},
		// Jan 5 (Fri, serial 5) - 5 workdays: Fri→Thu(4)→Wed(3)→Tue(2)→Mon(1)→Fri(Dec 29) → negative
		{"result_negative_backward_from_serial_5", "WORKDAY(5, -5)", ErrValNUM},
		// Larger backward count from a date close to epoch
		{"result_negative_backward_from_jan_15", "WORKDAY(DATE(1900,1,15), -15)", ErrValNUM},
		// Serial 2 (Jan 2, Tue) - 2 workdays: Tue→Mon(1)→Fri(Dec 29) → negative
		{"result_negative_backward_from_serial_2", "WORKDAY(2, -2)", ErrValNUM},

		// === Exact mismatch lock-in from fuzz test spec ===
		// DATE(2024,-2147483648,-0.5): extreme negative month triggers overflow guard → #NUM!
		// WORKDAY(#NUM!, -1) propagates the #NUM! error
		{"mismatch_extreme_neg_month", "WORKDAY(DATE(2024,-2147483648,-0.5), -1)", ErrValNUM},
		// Same with INT32_MAX month
		{"mismatch_extreme_pos_month", "WORKDAY(DATE(2024,2147483647,-0.5), -1)", ErrValNUM},
		// Extreme positive month just above the overflow guard limit (>120000)
		{"extreme_month_above_limit", "WORKDAY(DATE(2024,120001,1), 0)", ErrValNUM},
		// Extreme negative day below the overflow guard limit (<-4000000)
		{"extreme_day_below_limit", "WORKDAY(DATE(2024,1,-4000001), 0)", ErrValNUM},
		// Extreme positive day above the overflow guard limit (>4000000)
		{"extreme_day_above_limit", "WORKDAY(DATE(2024,1,4000001), 0)", ErrValNUM},

		// === Exact mismatch lock-in: DATE with 1e15 month ===
		// DATE(2024,1000000000000000,3.14159265358979) → #NUM! (month overflow)
		// WORKDAY should propagate that #NUM! error
		{"mismatch_1e15_month_frac_day", "WORKDAY(DATE(2024,1000000000000000,3.14159265358979), 0.5)", ErrValNUM},
		{"mismatch_1e15_month_zero_days", "WORKDAY(DATE(2024,1000000000000000,1), 0)", ErrValNUM},
		{"mismatch_1e15_month_positive_days", "WORKDAY(DATE(2024,1000000000000000,1), 5)", ErrValNUM},
		{"mismatch_neg_1e15_month", "WORKDAY(DATE(2024,-1000000000000000,1), 0)", ErrValNUM},
		// Extreme day value (1e15) should also overflow
		{"extreme_day_1e15", "WORKDAY(DATE(2024,1,1000000000000000), 0)", ErrValNUM},
		// Both month and day extreme
		{"extreme_month_and_day_1e15", "WORKDAY(DATE(2024,1000000000000000,1000000000000000), 0)", ErrValNUM},
		// DATE(2024,1e15,pi) with negative days
		{"mismatch_1e15_month_neg_days", "WORKDAY(DATE(2024,1000000000000000,3.14159265358979), -1)", ErrValNUM},
		// Various large-but-not-int-max month values
		{"extreme_month_1e12", "WORKDAY(DATE(2024,1000000000000,1), 0)", ErrValNUM},
		{"extreme_month_1e9", "WORKDAY(DATE(2024,1000000000,1), 0)", ErrValNUM},
		{"extreme_neg_month_1e12", "WORKDAY(DATE(2024,-1000000000000,1), 0)", ErrValNUM},

		// === Mismatch lock-in from fuzz test spec (Sheet4!A9) ===
		// DATE(-999999999, 0.001, 3.14159265358979) → #NUM! from extreme negative year.
		// WORKDAY must propagate #NUM! (not #VALUE!) regardless of days value.
		{"mismatch_lockin_frac_days", "WORKDAY(DATE(-999999999, 0.001, 3.14159265358979), 0.5)", ErrValNUM},
		{"mismatch_lockin_int_days", "WORKDAY(DATE(-999999999, 0.001, 3.14159265358979), 5)", ErrValNUM},
		{"mismatch_lockin_neg_days", "WORKDAY(DATE(-999999999, 0.001, 3.14159265358979), -1)", ErrValNUM},
		{"mismatch_lockin_zero_days", "WORKDAY(DATE(-999999999, 0.001, 3.14159265358979), 0)", ErrValNUM},
		// Verify the same behavior with negative year directly
		{"num_error_neg_year_500", "WORKDAY(DATE(-500, 1, 1), 1)", ErrValNUM},
		{"num_error_neg_year_intmin", "WORKDAY(DATE(-2147483648, 1, 1), 0)", ErrValNUM},

		// === Argument count validation ===
		// Too few arguments (1 arg) → #VALUE!
		{"too_few_args_one", "WORKDAY(DATE(2008,10,1))", ErrValVALUE},
		// Too many arguments (4 args) → #VALUE!
		{"too_many_args_four", "WORKDAY(DATE(2008,10,1), 5, DATE(2008,10,2), 99)", ErrValVALUE},

		// === Zero serial backward → result goes negative ===
		// Serial 0 (Dec 31, 1899 Sun) - 1 workday = Fri Dec 29 → negative
		{"zero_serial_backward", "WORKDAY(0, -1)", ErrValNUM},

		// === Negative result from forward direction (shouldn't normally happen,
		// but verify guard works when start_date is already at boundary) ===
		// DATE that produces serial exactly at boundary then backward
		{"date_produces_small_serial_backward", "WORKDAY(DATE(1900,1,2), -2)", ErrValNUM},

		// === Exact mismatch lock-in from test spec: positive extreme year (999999999) ===
		// DATE(999999999, 0.001, 3.14159265358979) → #NUM! from extreme positive year.
		// WORKDAY must propagate #NUM! (not #VALUE!). This was the exact bug:
		// werkbook returned #VALUE! but local-excel returned #NUM!.
		{"mismatch_lockin_pos_999999999_frac_days", "WORKDAY(DATE(999999999, 0.001, 3.14159265358979), 0.5)", ErrValNUM},
		{"mismatch_lockin_pos_999999999_zero_days", "WORKDAY(DATE(999999999, 0.001, 3.14159265358979), 0)", ErrValNUM},
		{"mismatch_lockin_pos_999999999_int_days", "WORKDAY(DATE(999999999, 0.001, 3.14159265358979), 5)", ErrValNUM},
		{"mismatch_lockin_pos_999999999_neg_days", "WORKDAY(DATE(999999999, 0.001, 3.14159265358979), -1)", ErrValNUM},
		// Variations with different extreme positive years
		{"num_error_year_100000", "WORKDAY(DATE(100000,1,1), 5)", ErrValNUM},
		{"num_error_year_50000", "WORKDAY(DATE(50000,1,1), 0)", ErrValNUM},
		{"num_error_year_intmax", "WORKDAY(DATE(2147483647,1,1), 0)", ErrValNUM},
		// Positive extreme year with fractional month and day (matching fuzz test spec inputs)
		{"pos_extreme_year_frac_month", "WORKDAY(DATE(999999999, 0.5, 1), 0)", ErrValNUM},
		{"pos_extreme_year_frac_day", "WORKDAY(DATE(999999999, 1, 0.5), 0)", ErrValNUM},
		// Exact test spec scenario: C1=999999999, C2=0.001, C3=3.14159265358979, C5=0.5
		{"test_spec_sheet4_a9", "WORKDAY(DATE(999999999, 0.001, 3.14159265358979), 0.5)", ErrValNUM},

		// === Empty string as argument: coerced to 0 ===
		// (empty string coerces to 0, which is valid for both start_date and days)

		// === Exact spec mismatch lock-in: DATE(-1,-1,pi) → #NUM! propagation ===
		// The original bug: werkbook returned #VALUE! but Excel returns #NUM!.
		// DATE(-1,-1,3.14159265358979) produces #NUM! (negative year), and
		// WORKDAY must propagate that #NUM! regardless of the days argument.
		{"spec_exact_date_neg_args_pi_zero_days", "WORKDAY(DATE(-1,-1,3.14159265358979), 0)", ErrValNUM},
		{"spec_exact_date_neg_args_pi_pos_days", "WORKDAY(DATE(-1,-1,3.14159265358979), 10)", ErrValNUM},
		{"spec_exact_date_neg_args_pi_neg_days", "WORKDAY(DATE(-1,-1,3.14159265358979), -10)", ErrValNUM},
		{"spec_exact_date_neg_args_pi_frac_days", "WORKDAY(DATE(-1,-1,3.14159265358979), 2.5)", ErrValNUM},
		{"spec_exact_date_neg_args_pi_bool_days", "WORKDAY(DATE(-1,-1,3.14159265358979), TRUE)", ErrValNUM},

		// === Additional #NUM! propagation from DATE with various invalid inputs ===
		// DATE(0, -100, 1) → year 1900, month -100 → July 1891 → pre-epoch → #NUM!
		{"date_year0_very_neg_month", "WORKDAY(DATE(0,-100,1), 0)", ErrValNUM},
		// DATE with all negative args
		{"date_all_negative_args", "WORKDAY(DATE(-1,-1,-1), 5)", ErrValNUM},
		// DATE with fractional negative year just below zero
		{"date_tiny_neg_year", "WORKDAY(DATE(-0.001, 1, 1), 0)", ErrValNUM},

		// === Result goes negative: backward from small serials ===
		// Jan 10, 1900 (serial 10, Wed) - 10 workdays → past epoch → #NUM!
		{"backward_past_epoch_from_jan10", "WORKDAY(DATE(1900,1,10), -10)", ErrValNUM},
		// Serial 4 (Jan 4, Thu) - 4 workdays → past epoch → #NUM!
		{"backward_past_epoch_serial_4", "WORKDAY(4, -4)", ErrValNUM},

		// === #REF! error propagation ===
		{"ref_error_start_date", "WORKDAY(#REF!, 5)", ErrValREF},

		// === Max serial boundary: start serial exceeds maxExcelSerial (2958465) ===
		// Raw serial one above max → #NUM! even with 0 days
		{"start_above_max_serial", "WORKDAY(2958466, 0)", ErrValNUM},
		// Raw serial far above max → #NUM!
		{"start_far_above_max_serial", "WORKDAY(3000000, 5)", ErrValNUM},
		// Raw serial above max going backward → still #NUM! (invalid start)
		{"start_above_max_backward", "WORKDAY(2958466, -5)", ErrValNUM},
		// Raw serial exactly at max+1 with negative days → still #NUM! (invalid start)
		{"start_max_plus_one_backward", "WORKDAY(2958466, -1)", ErrValNUM},

		// === Type coercion edge cases producing errors ===
		// Numeric string coerced to negative number → hits negative serial check → #NUM!
		{"string_neg_serial_coerced", `WORKDAY("-5", 1)`, ErrValNUM},
		// String "-0.5" coerced to -0.5 → negative start serial → #NUM!
		{"string_neg_frac_serial_coerced", `WORKDAY("-0.5", 0)`, ErrValNUM},
		// Boolean FALSE (=0) going backward → result goes negative → #NUM!
		{"bool_false_backward", "WORKDAY(FALSE, -1)", ErrValNUM},
		// Boolean TRUE (=1, Mon Jan 1 1900) backward 1 → past epoch → #NUM!
		{"bool_true_backward_one", "WORKDAY(TRUE, -1)", ErrValNUM},

		// === Mismatch lock-in: DATE with negative fractional month producing different serial ===
		// Verify WORKDAY(DATE(2024,-0.5,pi), 0.5) does NOT produce serial 45263
		// (the old wrong value); it should produce 45233 (correct).
		// We can't directly test "not equal" in this error table, but we CAN verify
		// that WORKDAY with a truly invalid start propagates errors correctly.
		// DATE(0,-0.5,1): Floor(-0.5)=-1 → month=-1 → Nov 1899 → serial < 0 → #NUM!
		{"date_year0_neg_frac_month", "WORKDAY(DATE(0,-0.5,1), 0)", ErrValNUM},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if got.Type != ValueError {
				t.Fatalf("%s: got type %v (%v), want error", tc.formula, got.Type, got)
			}
			if got.Err != tc.errVal {
				t.Errorf("%s: got error %v, want %v", tc.formula, got.Err, tc.errVal)
			}
		})
	}
}

func TestWORKDAYErrorPropagation(t *testing.T) {
	// When start_date evaluates to an error, WORKDAY should propagate it.
	// This matches the mismatch fix: DATE(-1,-1,pi) → #NUM! and
	// WORKDAY(#NUM!, ...) → #NUM! (not #VALUE!)
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			// A1 contains an error value (simulating DATE(-1,-1,pi) result)
			{Col: 1, Row: 1}: ErrorVal(ErrValNUM),
			// B1 contains a valid number
			{Col: 2, Row: 1}: NumberVal(0),
		},
	}

	t.Run("propagate_num_error_from_cell", func(t *testing.T) {
		cf := evalCompile(t, "WORKDAY(A1, B1)")
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("WORKDAY(#NUM!, 0): got %v, want #NUM!", got)
		}
	})

	t.Run("propagate_value_error_from_cell", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValVALUE),
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, 5)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("WORKDAY(#VALUE!, 5): got %v, want #VALUE!", got)
		}
	})

	t.Run("propagate_div0_error_from_cell", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValDIV0),
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, 5)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValDIV0 {
			t.Errorf("WORKDAY(#DIV/0!, 5): got %v, want #DIV/0!", got)
		}
	})

	// Error in days argument should be propagated
	t.Run("propagate_error_from_days_arg", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(39722),       // valid start date
				{Col: 2, Row: 1}: ErrorVal(ErrValVALUE),  // error in days
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, B1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("WORKDAY(valid, #VALUE!): got %v, want #VALUE!", got)
		}
	})

	// Mismatch lock-in: #NUM! propagation with fractional days (0.5)
	// This is the exact scenario from the fuzz test spec (Sheet4!A9):
	// WORKDAY(DATE(-999999999,...), 0.5) must return #NUM!, not #VALUE!
	t.Run("propagate_num_error_with_fractional_days", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNUM),
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, 0.5)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("WORKDAY(#NUM!, 0.5): got %v, want #NUM!", got)
		}
	})

	// #NUM! propagation with positive integer days
	t.Run("propagate_num_error_with_positive_days", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNUM),
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, 5)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("WORKDAY(#NUM!, 5): got %v, want #NUM!", got)
		}
	})

	// #NUM! propagation with negative days
	t.Run("propagate_num_error_with_negative_days", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNUM),
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, -1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("WORKDAY(#NUM!, -1): got %v, want #NUM!", got)
		}
	})

	// #NA propagation from start_date
	t.Run("propagate_na_error_from_start", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNA),
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, 5)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNA {
			t.Errorf("WORKDAY(#N/A, 5): got %v, want #N/A", got)
		}
	})

	// Error in days argument with fractional start_date
	t.Run("propagate_error_from_days_fractional", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(39722),      // valid start date
				{Col: 2, Row: 1}: ErrorVal(ErrValNUM),   // error in days
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, B1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("WORKDAY(valid, #NUM!): got %v, want #NUM!", got)
		}
	})

	// Error in holiday range should be propagated
	t.Run("error_in_holiday_range", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNUM),
			},
		}
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 5, A1:A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("WORKDAY with error in holidays: got %v, want #NUM!", got)
		}
	})

	// === Exact test spec scenario lock-in (Sheet4!A9) ===
	// Simulates: A1 = DATE(999999999, 0.001, 3.14159265358979) → #NUM!
	// Then WORKDAY(A1, 0.5) must return #NUM!, not #VALUE!.
	// This is the cross-sheet scenario from the fuzz test spec.
	t.Run("test_spec_sheet4_a9_cell_ref", func(t *testing.T) {
		// A1 holds the #NUM! error that DATE(999999999,...) would produce
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNUM),
				{Col: 2, Row: 1}: NumberVal(0.5),
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, B1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("WORKDAY(#NUM!, 0.5) via cell ref: got %v, want #NUM!", got)
		}
	})

	// #VALUE! in holiday range should propagate as #VALUE!
	t.Run("value_error_in_holiday_range", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValVALUE),
			},
		}
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 5, A1:A1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("WORKDAY with #VALUE! in holidays: got %v, want #VALUE!", got)
		}
	})

	// Both start_date and days are errors — first error (start_date) takes precedence
	t.Run("both_args_are_errors", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNUM),
				{Col: 2, Row: 1}: ErrorVal(ErrValVALUE),
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, B1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Fatalf("WORKDAY(#NUM!, #VALUE!): got type %v, want error", got.Type)
		}
		// The first error encountered (start_date) should be propagated
		if got.Err != ErrValNUM {
			t.Errorf("WORKDAY(#NUM!, #VALUE!): got %v, want #NUM!", got.Err)
		}
	})

	// Empty cell as start_date: empty → 0 (serial 0) + 0 days = serial 0
	t.Run("empty_cell_as_start_date", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				// A1 not set → EmptyVal
				{Col: 2, Row: 1}: NumberVal(1),
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, B1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("WORKDAY(empty, 1): got type %v (%v), want number", got.Type, got)
		}
		// Empty cell → serial 0 (Sun Dec 31 1899) + 1 workday = serial 1 (Mon Jan 1 1900)
		if got.Num != 1 {
			t.Errorf("WORKDAY(empty, 1) = %g, want 1", got.Num)
		}
	})

	// Empty cell as days: empty → 0 days, returns start_date unchanged
	t.Run("empty_cell_as_days", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(39722), // Oct 1, 2008
				// B1 not set → EmptyVal
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, B1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueNumber {
			t.Fatalf("WORKDAY(date, empty): got type %v (%v), want number", got.Type, got)
		}
		if got.Num != 39722 {
			t.Errorf("WORKDAY(date, empty) = %g, want 39722", got.Num)
		}
	})

	// #REF! error propagation from start_date cell
	t.Run("propagate_ref_error_from_start", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValREF),
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, 5)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValREF {
			t.Errorf("WORKDAY(#REF!, 5): got %v, want #REF!", got)
		}
	})

	// #NUM! error propagation with zero days (exact mismatch scenario)
	// The spec had WORKDAY(DATE(-1,-1,pi), 0) which should return #NUM! not #VALUE!
	t.Run("propagate_num_error_zero_days", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNUM),
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, 0)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValNUM {
			t.Errorf("WORKDAY(#NUM!, 0): got %v, want #NUM!", got)
		}
	})

	// Error in holiday range with valid start and days: error propagates
	t.Run("error_in_holiday_range_with_valid_args", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: NumberVal(39722),       // valid holiday
				{Col: 1, Row: 2}: ErrorVal(ErrValVALUE),  // error in holiday range
			},
		}
		cf := evalCompile(t, "WORKDAY(DATE(2008,10,1), 5, A1:A2)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("WORKDAY with error mixed in holidays: got %v, want #VALUE!", got)
		}
	})

	// All three args are errors: start_date error takes precedence
	t.Run("all_three_args_are_errors", func(t *testing.T) {
		r := &mockResolver{
			cells: map[CellAddr]Value{
				{Col: 1, Row: 1}: ErrorVal(ErrValNUM),    // start_date
				{Col: 2, Row: 1}: ErrorVal(ErrValVALUE),  // days
				{Col: 3, Row: 1}: ErrorVal(ErrValDIV0),   // holiday
			},
		}
		cf := evalCompile(t, "WORKDAY(A1, B1, C1:C1)")
		got, err := Eval(cf, r, nil)
		if err != nil {
			t.Fatalf("Eval: %v", err)
		}
		if got.Type != ValueError {
			t.Fatalf("WORKDAY(#NUM!, #VALUE!, #DIV/0!): got type %v, want error", got.Type)
		}
		if got.Err != ErrValNUM {
			t.Errorf("WORKDAY(#NUM!, #VALUE!, #DIV/0!): got %v, want #NUM!", got.Err)
		}
	})
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
