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

func TestDAYS360(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Basic US method: Jan 1 to Feb 1 = 30 days
		{"us_jan1_feb1", "DAYS360(DATE(2025,1,1),DATE(2025,2,1),FALSE)", 30},
		// US method: Feb 28 (last day of Feb, non-leap) to Mar 31
		// Feb 28 → D1=30 (last-of-Feb rule), Mar 31 → D2=30 (D2==31 && D1>=30)
		// Result: (3-2)*30 + (30-30) = 30
		{"us_feb28_mar31", "DAYS360(45716,45747,FALSE)", 30},
		// US method: Jan 31 to Mar 31
		// Jan 31 → D1=30, Mar 31 → D2=30 (D2==31 && D1>=30)
		// Result: (3-1)*30 + (30-30) = 60
		{"us_jan31_mar31", "DAYS360(DATE(2025,1,31),DATE(2025,3,31),FALSE)", 60},
		// US method: both dates last day of Feb (leap year 2024)
		// Feb 29 2024 → D1=30, Feb 29 2024 → D2=30 (both last-of-Feb)
		// Result: 0
		{"us_both_feb_leap", "DAYS360(DATE(2024,2,29),DATE(2024,2,29),FALSE)", 0},
		// US method: Feb 29 (leap) to Mar 31
		// Feb 29 → D1=30, Mar 31 → D2=30
		// Result: (3-2)*30 + (30-30) = 30
		{"us_feb29_mar31_leap", "DAYS360(DATE(2024,2,29),DATE(2024,3,31),FALSE)", 30},
		// European method: same dates
		// European: D1=28, D2=31→30. (3-2)*30 + (30-28) = 32
		{"eu_feb28_mar31", "DAYS360(45716,45747,TRUE)", 32},
		// US method: regular dates, no adjustments needed
		{"us_jan15_mar15", "DAYS360(DATE(2025,1,15),DATE(2025,3,15),FALSE)", 60},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
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

func TestDATEDIF(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Basic units
		{"Y", `DATEDIF(45307,45736,"Y")`, 1},
		{"M", `DATEDIF(45307,45736,"M")`, 14},
		{"D", `DATEDIF(45307,45736,"D")`, 429},
		{"MD", `DATEDIF(45307,45736,"MD")`, 4},
		{"YM", `DATEDIF(45307,45736,"YM")`, 2},

		// YD unit — days ignoring year difference
		{"YD_cross_year", `DATEDIF(45307,45736,"YD")`, 64},
		{"YD_within_year", `DATEDIF(45307,45672,"YD")`, 365},
		{"YD_same_date", `DATEDIF(45307,45307,"YD")`, 0},

		// YD unit — leap year boundary cases (excel-audit edge cases)
		{"YD_leap_mar1_to_mar1", `DATEDIF(DATE(2000,3,1),DATE(2004,3,1),"YD")`, 0},
		{"YD_leap_jan1_to_jan1", `DATEDIF(DATE(2000,1,1),DATE(2008,1,1),"YD")`, 0},
		{"YD_leap_feb28_to_mar1", `DATEDIF(DATE(2000,2,28),DATE(2004,3,1),"YD")`, 2},
		{"YD_leap_jan1_to_jan2", `DATEDIF(DATE(2000,1,1),DATE(2004,1,2),"YD")`, 1},
		{"YD_leap_jul1_to_jul1", `DATEDIF(DATE(2023,7,1),DATE(2024,7,1),"YD")`, 0},
		{"YD_leap_mar1_to_mar2", `DATEDIF(DATE(2001,3,1),DATE(2004,3,2),"YD")`, 1},
		{"YD_leap_dec1_to_dec1", `DATEDIF(DATE(2000,12,1),DATE(2004,12,1),"YD")`, 0},
		{"YD_leap_feb29_to_feb29", `DATEDIF(DATE(2000,2,29),DATE(2004,2,29),"YD")`, 0},
		{"YD_leap_feb29_to_mar1", `DATEDIF(DATE(2000,2,29),DATE(2004,3,1),"YD")`, 1},
		{"YD_leap_century_jan1", `DATEDIF(DATE(1901,1,1),DATE(2001,1,1),"YD")`, 0},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
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

func TestSerial60Boundary(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		// Serial 60 = Excel's fictional Feb 29, 1900
		{"DAY_60", "DAY(60)", 29},
		{"MONTH_60", "MONTH(60)", 2},
		{"YEAR_60", "YEAR(60)", 1900},
		// Neighbors: serial 59 = Feb 28, serial 61 = Mar 1
		{"DAY_59", "DAY(59)", 28},
		{"MONTH_59", "MONTH(59)", 2},
		{"DAY_61", "DAY(61)", 1},
		{"MONTH_61", "MONTH(61)", 3},
		// DATE(1900,2,29) should return serial 60
		{"DATE_1900_2_29", "DATE(1900,2,29)", 60},
		// TIME(25,0,0) should wrap to fractional part only
		{"TIME_25_0_0", "TIME(25,0,0)", 1.0 / 24.0},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: got type %v, want number", tc.formula, got.Type)
			}
			if math.Abs(got.Num-tc.want) > 1e-12 {
				t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
			}
		})
	}
}
