package formula

import (
	"math"
	"testing"
	"time"
)

func excelDateSerial(year int, month time.Month, day int) float64 {
	return math.Floor(TimeToExcelSerial(time.Date(year, month, day, 0, 0, 0, 0, time.UTC)))
}

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

func TestDATEComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		// Basic dates
		{"basic_jan1_2023", "DATE(2023,1,1)", 44927, false, 0},
		{"basic_dec31_2023", "DATE(2023,12,31)", 45291, false, 0},
		{"basic_jul4_2000", "DATE(2000,7,4)", 36711, false, 0},

		// Excel doc examples
		{"doc_example_2008_1_2", "DATE(2008,1,2)", 39449, false, 0},
		{"doc_example_108_1_2", "DATE(108,1,2)", 39449, false, 0},         // 1900+108=2008
		{"doc_example_2008_14_2", "DATE(2008,14,2)", 39846, false, 0},     // Feb 2, 2009
		{"doc_example_2008_neg3_2", "DATE(2008,-3,2)", 39327, false, 0},   // Sep 2, 2007
		{"doc_example_2008_1_35", "DATE(2008,1,35)", 39482, false, 0},     // Feb 4, 2008
		{"doc_example_2008_1_neg15", "DATE(2008,1,-15)", 39432, false, 0}, // Dec 16, 2007

		// Year < 1900 behavior (0-1899 adds to 1900)
		{"year_0_jan1", "DATE(0,1,1)", 1, false, 0},               // 1900-01-01 = serial 1
		{"year_99_jan1", "DATE(99,1,1)", 36161, false, 0},         // 1999-01-01
		{"year_1899_dec31", "DATE(1899,12,31)", 693962, false, 0}, // 3799-12-31

		// Year 1900 (serial number 1 for Jan 1)
		{"year_1900_jan1", "DATE(1900,1,1)", 1, false, 0},
		{"year_1900_jan2", "DATE(1900,1,2)", 2, false, 0},
		{"year_1900_feb28", "DATE(1900,2,28)", 59, false, 0},
		{"year_1900_feb29_fictional", "DATE(1900,2,29)", 60, false, 0}, // Excel fictional leap day
		{"year_1900_mar1", "DATE(1900,3,1)", 61, false, 0},

		// Month overflow: DATE(2023,13,1) → Jan 2024
		{"month_overflow_13", "DATE(2023,13,1)", 45292, false, 0},
		{"month_overflow_14", "DATE(2023,14,1)", 45323, false, 0}, // Feb 1, 2024
		{"month_overflow_25", "DATE(2023,25,1)", 45658, false, 0}, // Jan 1, 2025

		// Month negative/zero: DATE(2023,0,1) → Dec 2022
		{"month_zero", "DATE(2023,0,1)", 44896, false, 0},    // Dec 1, 2022
		{"month_neg1", "DATE(2023,-1,1)", 44866, false, 0},   // Nov 1, 2022
		{"month_neg12", "DATE(2023,-12,1)", 44531, false, 0}, // Dec 1, 2021 (actually Jan 1 -12 months)

		// Day overflow: DATE(2023,1,32) → Feb 1
		{"day_overflow_32", "DATE(2023,1,32)", 44958, false, 0},   // Feb 1, 2023
		{"day_overflow_60", "DATE(2023,1,60)", 44986, false, 0},   // Mar 1, 2023
		{"day_overflow_365", "DATE(2023,1,365)", 45291, false, 0}, // Dec 31, 2023

		// Day negative/zero
		{"day_zero", "DATE(2023,1,0)", 44926, false, 0},    // Dec 31, 2022
		{"day_neg1", "DATE(2023,1,-1)", 44925, false, 0},   // Dec 30, 2022
		{"day_neg30", "DATE(2023,1,-30)", 44896, false, 0}, // Dec 1, 2022

		// Leap year: DATE(2024,2,29) — valid
		{"leap_year_2024_feb29", "DATE(2024,2,29)", 45351, false, 0},
		{"leap_year_2024_feb28", "DATE(2024,2,28)", 45350, false, 0},
		{"leap_year_2000_feb29", "DATE(2000,2,29)", 36585, false, 0},

		// Non-leap year: DATE(2023,2,29) → Mar 1
		{"non_leap_feb29", "DATE(2023,2,29)", 44986, false, 0},   // Mar 1, 2023
		{"non_leap_1900_feb29", "DATE(1900,2,29)", 60, false, 0}, // Excel fictional

		// Large year values (but within range)
		{"year_9999_dec31", "DATE(9999,12,31)", 2958465, false, 0}, // Max serial
		{"year_9999_jan1", "DATE(9999,1,1)", 2958101, false, 0},

		// Year out of range
		{"year_10000_out_of_range", "DATE(10000,1,1)", 0, true, ErrValNUM},
		{"year_negative_out_of_range", "DATE(-1,1,1)", 0, true, ErrValNUM},

		// String/boolean coercion (via CoerceNum)
		{"bool_true_year", "DATE(TRUE,1,1)", 367, false, 0}, // TRUE=1, 1+1900=1901, Jan 1

		// Too few/many args → error
		{"too_few_args_0", "DATE()", 0, true, ErrValVALUE},
		{"too_few_args_2", "DATE(2023,1)", 0, true, ErrValVALUE},
		{"too_many_args", "DATE(2023,1,1,1)", 0, true, ErrValVALUE},
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
				t.Fatalf("%s: got type %v (%v), want number", tc.formula, got.Type, got)
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

func TestYEARComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		// Basic: YEAR(DATE(2023,6,15)) → 2023
		{"basic_date_2023", "YEAR(DATE(2023,6,15))", 2023, false, 0},
		// Serial number 1 → 1900 (Jan 1, 1900)
		{"serial_1", "YEAR(1)", 1900, false, 0},
		// Serial number 367 → 1901 (Jan 2, 1901)
		{"serial_367", "YEAR(367)", 1901, false, 0},
		// Year 2000 (Y2K): Jan 1, 2000 = serial 36526
		{"year_2000", "YEAR(36526)", 2000, false, 0},
		// Modern date 2024: serial 45306 = Jan 15, 2024
		{"year_2024", "YEAR(45306)", 2024, false, 0},
		// Dec 31 of a year: DATE(2023,12,31)
		{"dec_31", "YEAR(DATE(2023,12,31))", 2023, false, 0},
		// Jan 1 of a year: DATE(2025,1,1)
		{"jan_1", "YEAR(DATE(2025,1,1))", 2025, false, 0},
		// Leap year date: Feb 29, 2024
		{"leap_year_2024", "YEAR(DATE(2024,2,29))", 2024, false, 0},
		// Serial 0 → 1900 (Excel's "January 0, 1900" sentinel)
		{"serial_0", "YEAR(0)", 1900, false, 0},
		// Serial 60 → 1900 (Excel's fictional Feb 29, 1900)
		{"serial_60", "YEAR(60)", 1900, false, 0},
		// String date input via DATEVALUE: YEAR(DATEVALUE("1/1/2023"))
		{"string_date_via_datevalue", `YEAR(DATEVALUE("1/1/2023"))`, 2023, false, 0},
		// String date input directly: CoerceNum cannot parse date strings → #VALUE!
		{"string_date_direct", `YEAR("1/1/2023")`, 0, true, ErrValVALUE},
		// Negative serial → #NUM!
		{"negative_serial", "YEAR(-1)", 0, true, ErrValNUM},
		// Too few args → error
		{"no_args", "YEAR()", 0, true, ErrValVALUE},
		// Too many args → error
		{"too_many_args", "YEAR(1,2)", 0, true, ErrValVALUE},
		// Error propagation: YEAR(#VALUE!) → #VALUE!
		{"error_propagation", `YEAR("abc")`, 0, true, ErrValVALUE},
		// Large serial number (far future): serial 2958465 = Dec 31, 9999
		{"max_serial", "YEAR(2958465)", 9999, false, 0},
		// Beyond max serial → #NUM!
		{"beyond_max_serial", "YEAR(2958466)", 0, true, ErrValNUM},
		// Excel doc examples via DATEVALUE
		{"excel_doc_2023", `YEAR(DATEVALUE("7/5/2023"))`, 2023, false, 0},
		{"excel_doc_2025", `YEAR(DATEVALUE("7/5/2025"))`, 2025, false, 0},
		// Fractional serial (should use integer part): 45306.75 → 2024
		{"fractional_serial", "YEAR(45306.75)", 2024, false, 0},
		// Last day of 1900: serial 366 = Dec 31, 1900
		{"last_day_1900", "YEAR(366)", 1900, false, 0},
		// Boolean TRUE coerced to 1 → 1900
		{"bool_true", "YEAR(TRUE)", 1900, false, 0},
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

func TestYEARMONTHDAY_Serial0(t *testing.T) {
	resolver := &mockResolver{}

	// Excel serial 0 is "January 0, 1900" — a special sentinel.
	// YEAR(0)=1900, MONTH(0)=1, DAY(0)=0.
	cf := evalCompile(t, "YEAR(0)")
	got, err := Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(YEAR(0)): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1900 {
		t.Errorf("YEAR(0) = %g, want 1900", got.Num)
	}

	cf = evalCompile(t, "MONTH(0)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(MONTH(0)): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 1 {
		t.Errorf("MONTH(0) = %g, want 1", got.Num)
	}

	cf = evalCompile(t, "DAY(0)")
	got, err = Eval(cf, resolver, nil)
	if err != nil {
		t.Fatalf("Eval(DAY(0)): %v", err)
	}
	if got.Type != ValueNumber || got.Num != 0 {
		t.Errorf("DAY(0) = %g, want 0", got.Num)
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

func TestDAYS(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		{"doc_example", `DAYS(DATEVALUE("3/15/2021"),DATEVALUE("2/1/2021"))`, 42, false, 0},
		{"same_year", "DAYS(DATE(2021,12,31),DATE(2021,1,1))", 364, false, 0},
		{"same_date", "DAYS(DATE(2025,1,1),DATE(2025,1,1))", 0, false, 0},
		{"negative_result", "DAYS(DATE(2025,2,1),DATE(2025,3,15))", -42, false, 0},
		{"leap_year_span", "DAYS(DATE(2024,3,1),DATE(2024,2,28))", 2, false, 0},
		{"non_leap_span", "DAYS(DATE(2025,3,1),DATE(2025,2,28))", 1, false, 0},
		{"ignores_time_components", "DAYS(DATE(2025,1,2)+TIME(23,59,59),DATE(2025,1,1)+TIME(12,0,0))", 1, false, 0},
		{"truncates_fractional_serials", "DAYS(10.9,8.1)", 2, false, 0},
		{"string_numbers", `DAYS("10","8")`, 2, false, 0},
		{"excel_1900_leap_bug_gap", "DAYS(DATE(1900,3,1),DATE(1900,2,28))", 2, false, 0},
		{"far_future", "DAYS(DATE(9999,12,31),DATE(9999,1,1))", 364, false, 0},
		{"no_args", "DAYS()", 0, true, ErrValVALUE},
		{"one_arg", "DAYS(1)", 0, true, ErrValVALUE},
		{"too_many_args", "DAYS(1,2,3)", 0, true, ErrValVALUE},
		{"non_numeric", `DAYS("abc",1)`, 0, true, ErrValVALUE},
		{"error_propagation_end", "DAYS(1/0,1)", 0, true, ErrValDIV0},
		{"error_propagation_start", "DAYS(1,1/0)", 0, true, ErrValDIV0},
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

func TestEDATE(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		{"doc_plus_one", "EDATE(DATE(2011,1,15),1)", excelDateSerial(2011, time.February, 15), false, 0},
		{"doc_minus_one", "EDATE(DATE(2011,1,15),-1)", excelDateSerial(2010, time.December, 15), false, 0},
		{"doc_plus_two", "EDATE(DATE(2011,1,15),2)", excelDateSerial(2011, time.March, 15), false, 0},
		{"month_end_non_leap", "EDATE(DATE(2025,1,31),1)", excelDateSerial(2025, time.February, 28), false, 0},
		{"month_end_leap", "EDATE(DATE(2024,1,31),1)", excelDateSerial(2024, time.February, 29), false, 0},
		{"cross_year_forward", "EDATE(DATE(2024,11,30),3)", excelDateSerial(2025, time.February, 28), false, 0},
		{"cross_year_backward", "EDATE(DATE(2025,1,31),-2)", excelDateSerial(2024, time.November, 30), false, 0},
		{"zero_months", "EDATE(DATE(2025,1,15),0)", excelDateSerial(2025, time.January, 15), false, 0},
		{"truncates_positive_fraction", "EDATE(DATE(2025,1,15),1.9)", excelDateSerial(2025, time.February, 15), false, 0},
		{"truncates_negative_fraction", "EDATE(DATE(2025,1,15),-1.9)", excelDateSerial(2024, time.December, 15), false, 0},
		{"ignores_time_component", "EDATE(DATE(2025,1,15)+TIME(12,0,0),1)", excelDateSerial(2025, time.February, 15), false, 0},
		{"leap_back_one_month", "EDATE(DATE(2024,3,31),-1)", excelDateSerial(2024, time.February, 29), false, 0},
		{"string_months", `EDATE(DATE(2025,1,15),"2")`, excelDateSerial(2025, time.March, 15), false, 0},
		{"too_few_args", "EDATE(DATE(2025,1,15))", 0, true, ErrValVALUE},
		{"too_many_args", "EDATE(DATE(2025,1,15),1,2)", 0, true, ErrValVALUE},
		{"invalid_start", `EDATE("abc",1)`, 0, true, ErrValVALUE},
		{"invalid_months", `EDATE(DATE(2025,1,15),"abc")`, 0, true, ErrValVALUE},
		{"error_propagation", "EDATE(1/0,1)", 0, true, ErrValDIV0},
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

func TestEOMONTH(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		{"doc_plus_one", "EOMONTH(DATE(2011,1,1),1)", excelDateSerial(2011, time.February, 28), false, 0},
		{"doc_minus_three", "EOMONTH(DATE(2011,1,1),-3)", excelDateSerial(2010, time.October, 31), false, 0},
		{"same_month", "EOMONTH(DATE(2025,1,15),0)", excelDateSerial(2025, time.January, 31), false, 0},
		{"leap_february", "EOMONTH(DATE(2024,1,15),1)", excelDateSerial(2024, time.February, 29), false, 0},
		{"month_end_self", "EOMONTH(DATE(2025,1,31),0)", excelDateSerial(2025, time.January, 31), false, 0},
		{"previous_month_end", "EOMONTH(DATE(2025,1,31),-1)", excelDateSerial(2024, time.December, 31), false, 0},
		{"cross_year_forward", "EOMONTH(DATE(2024,11,5),2)", excelDateSerial(2025, time.January, 31), false, 0},
		{"truncates_positive_fraction", "EOMONTH(DATE(2025,1,15),1.9)", excelDateSerial(2025, time.February, 28), false, 0},
		{"truncates_negative_fraction", "EOMONTH(DATE(2025,1,15),-1.9)", excelDateSerial(2024, time.December, 31), false, 0},
		{"ignores_time_component", "EOMONTH(DATE(2025,1,15)+TIME(18,30,0),1)", excelDateSerial(2025, time.February, 28), false, 0},
		{"leap_plus_twelve", "EOMONTH(DATE(2024,2,29),12)", excelDateSerial(2025, time.February, 28), false, 0},
		{"leap_minus_twelve", "EOMONTH(DATE(2024,2,29),-12)", excelDateSerial(2023, time.February, 28), false, 0},
		{"string_months", `EOMONTH(DATE(2025,1,15),"2")`, excelDateSerial(2025, time.March, 31), false, 0},
		{"too_few_args", "EOMONTH(DATE(2025,1,15))", 0, true, ErrValVALUE},
		{"too_many_args", "EOMONTH(DATE(2025,1,15),1,2)", 0, true, ErrValVALUE},
		{"invalid_months", `EOMONTH(DATE(2025,1,15),"abc")`, 0, true, ErrValVALUE},
		{"error_propagation", "EOMONTH(DATE(2025,1,15),1/0)", 0, true, ErrValDIV0},
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

func TestHOUR(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		{"doc_decimal", "HOUR(0.75)", 18, false, 0},
		{"doc_datetime", "HOUR(DATE(2011,7,18)+TIME(7,45,0))", 7, false, 0},
		{"doc_date_only", "HOUR(DATE(2012,4,21))", 0, false, 0},
		{"timevalue_pm", `HOUR(TIMEVALUE("6:45 PM"))`, 18, false, 0},
		{"midnight", "HOUR(0)", 0, false, 0},
		{"noon", "HOUR(0.5)", 12, false, 0},
		{"end_of_day", "HOUR(TIME(23,59,59))", 23, false, 0},
		{"whole_day_plus_fraction", "HOUR(1.75)", 18, false, 0},
		{"date_with_time", "HOUR(DATE(2025,1,1)+TIME(1,30,0))", 1, false, 0},
		{"seconds_only", "HOUR(TIME(0,0,59))", 0, false, 0},
		{"timevalue_midnight", `HOUR(TIMEVALUE("12:00 AM"))`, 0, false, 0},
		{"timevalue_noon", `HOUR(TIMEVALUE("12:00 PM"))`, 12, false, 0},
		{"no_args", "HOUR()", 0, true, ErrValVALUE},
		{"too_many_args", "HOUR(1,2)", 0, true, ErrValVALUE},
		{"non_numeric", `HOUR("abc")`, 0, true, ErrValVALUE},
		{"error_propagation", "HOUR(1/0)", 0, true, ErrValDIV0},
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

func TestMINUTE(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		{"doc_value", "MINUTE(TIME(12,45,0))", 45, false, 0},
		{"timevalue_pm", `MINUTE(TIMEVALUE("6:45 PM"))`, 45, false, 0},
		{"datetime_serial", "MINUTE(DATE(2025,1,1)+TIME(7,45,30))", 45, false, 0},
		{"date_only", "MINUTE(DATE(2012,4,21))", 0, false, 0},
		{"three_quarter_day", "MINUTE(0.75)", 0, false, 0},
		{"end_of_hour", "MINUTE(TIME(0,59,59))", 59, false, 0},
		{"one_minute", "MINUTE(TIME(0,1,0))", 1, false, 0},
		{"end_of_day", "MINUTE(TIME(23,59,59))", 59, false, 0},
		{"timevalue_seconds", `MINUTE(TIMEVALUE("1:30:45"))`, 30, false, 0},
		{"noon", "MINUTE(0.5)", 0, false, 0},
		{"timevalue_midnight", `MINUTE(TIMEVALUE("12:00 AM"))`, 0, false, 0},
		{"zero_minute", "MINUTE(TIME(10,0,1))", 0, false, 0},
		{"no_args", "MINUTE()", 0, true, ErrValVALUE},
		{"too_many_args", "MINUTE(1,2)", 0, true, ErrValVALUE},
		{"non_numeric", `MINUTE("abc")`, 0, true, ErrValVALUE},
		{"error_propagation", "MINUTE(1/0)", 0, true, ErrValDIV0},
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

func TestSECOND(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		{"doc_value", "SECOND(TIME(16,48,18))", 18, false, 0},
		{"doc_zero", "SECOND(TIME(16,48,0))", 0, false, 0},
		{"timevalue_seconds", `SECOND(TIMEVALUE("1:30:45"))`, 45, false, 0},
		{"datetime_serial", "SECOND(DATE(2025,1,1)+TIME(7,45,30))", 30, false, 0},
		{"date_only", "SECOND(DATE(2012,4,21))", 0, false, 0},
		{"one_second", "SECOND(TIME(0,0,1))", 1, false, 0},
		{"half_day", "SECOND(0.5)", 0, false, 0},
		{"end_of_day", "SECOND(TIME(23,59,59))", 59, false, 0},
		{"timevalue_noon", `SECOND(TIMEVALUE("12:00 PM"))`, 0, false, 0},
		{"timevalue_end_of_day", `SECOND(TIMEVALUE("23:59:59"))`, 59, false, 0},
		{"fraction_with_seconds", "SECOND(DATE(2025,1,1)+TIME(0,0,30))", 30, false, 0},
		{"whole_day_plus_one_second", "SECOND(1+TIME(0,0,1))", 1, false, 0},
		{"no_args", "SECOND()", 0, true, ErrValVALUE},
		{"too_many_args", "SECOND(1,2)", 0, true, ErrValVALUE},
		{"non_numeric", `SECOND("abc")`, 0, true, ErrValVALUE},
		{"error_propagation", "SECOND(1/0)", 0, true, ErrValDIV0},
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

func TestTIME(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("numeric", func(t *testing.T) {
		tests := []struct {
			name    string
			formula string
			want    float64
		}{
			// Basic cases
			{"midnight", "TIME(0,0,0)", 0},
			{"noon", "TIME(12,0,0)", 0.5},
			{"quarter_day", "TIME(6,0,0)", 0.25},
			{"three_quarter_day", "TIME(18,0,0)", 0.75},
			{"end_of_day", "TIME(23,59,59)", 0.999988425925926},

			// Minutes only
			{"30_minutes", "TIME(0,30,0)", 30.0 / 1440.0},
			{"1_minute", "TIME(0,1,0)", 1.0 / 1440.0},
			{"59_minutes", "TIME(0,59,0)", 59.0 / 1440.0},

			// Seconds only
			{"30_seconds", "TIME(0,0,30)", 30.0 / 86400.0},
			{"1_second", "TIME(0,0,1)", 1.0 / 86400.0},

			// Mixed
			{"16_48_10", "TIME(16,48,10)", 0.700115740740741},

			// Hour overflow (mod 24)
			{"hour_25", "TIME(25,0,0)", 1.0 / 24.0},
			{"hour_24", "TIME(24,0,0)", 0},
			{"hour_48", "TIME(48,0,0)", 0},
			{"hour_27", "TIME(27,0,0)", 3.0 / 24.0},

			// Minute overflow
			{"90_minutes", "TIME(0,90,0)", 1.5 / 24.0},
			{"minute_750", "TIME(0,750,0)", 0.520833333333333},
			{"minute_1440", "TIME(0,1440,0)", 0},

			// Second overflow
			{"3600_seconds", "TIME(0,0,3600)", 1.0 / 24.0},
			{"second_2000", "TIME(0,0,2000)", 0.023148148148148},
			{"second_7200", "TIME(0,0,7200)", 2.0 / 24.0},

			// Fractional args are truncated
			{"frac_hour", "TIME(12.9,0,0)", 0.5},
			{"frac_minute", "TIME(0,30.7,0)", 30.0 / 1440.0},
			{"frac_second", "TIME(0,0,30.9)", 30.0 / 86400.0},

			// String coercion
			{"string_hour", `TIME("12",0,0)`, 0.5},
			{"string_all", `TIME("6","30","0")`, 6.5 / 24.0},

			// Excel doc examples
			{"doc_example_1", "TIME(12,0,0)", 0.5},
			{"doc_example_2", "TIME(16,48,10)", 0.700115740740741},

			// Large valid values
			{"hour_32767", "TIME(32767,0,0)", float64(32767%24) * 3600 / 86400},
		}

		for _, tc := range tests {
			t.Run(tc.name, func(t *testing.T) {
				cf := evalCompile(t, tc.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%s): %v", tc.formula, err)
				}
				if got.Type != ValueNumber {
					t.Fatalf("%s: got type %v (%v), want number", tc.formula, got.Type, got)
				}
				if math.Abs(got.Num-tc.want) > 1e-9 {
					t.Errorf("%s = %.15g, want %.15g", tc.formula, got.Num, tc.want)
				}
			})
		}
	})

	t.Run("errors", func(t *testing.T) {
		tests := []struct {
			name    string
			formula string
			errVal  ErrorValue
		}{
			// Too few args
			{"no_args", "TIME()", ErrValVALUE},
			{"one_arg", "TIME(1)", ErrValVALUE},
			{"two_args", "TIME(1,2)", ErrValVALUE},

			// Too many args
			{"four_args", "TIME(1,2,3,4)", ErrValVALUE},

			// Non-numeric
			{"non_numeric_hour", `TIME("abc",0,0)`, ErrValVALUE},
			{"non_numeric_minute", `TIME(0,"xyz",0)`, ErrValVALUE},
			{"non_numeric_second", `TIME(0,0,"foo")`, ErrValVALUE},

			// Exceeds 32767
			{"hour_over_32767", "TIME(32768,0,0)", ErrValNUM},
			{"minute_over_32767", "TIME(0,32768,0)", ErrValNUM},
			{"second_over_32767", "TIME(0,0,32768)", ErrValNUM},

			// Negative total time
			{"negative_hour", "TIME(-1,0,0)", ErrValNUM},
			{"negative_total", "TIME(0,-1,0)", ErrValNUM},
		}

		for _, tc := range tests {
			t.Run(tc.name, func(t *testing.T) {
				cf := evalCompile(t, tc.formula)
				got, err := Eval(cf, resolver, nil)
				if err != nil {
					t.Fatalf("Eval(%s): unexpected Go error: %v", tc.formula, err)
				}
				if got.Type != ValueError {
					t.Fatalf("%s: got type %v (%v), want error", tc.formula, got.Type, got)
				}
				if got.Err != tc.errVal {
					t.Errorf("%s: got error %v, want %v", tc.formula, got.Err, tc.errVal)
				}
			})
		}
	})
}

func TestTIMEVALUE(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"noon", `TIMEVALUE("12:00")`, 0.5},
		{"6:30 PM", `TIMEVALUE("6:30 PM")`, 0.7708333333333334},
		{"midnight_0:00", `TIMEVALUE("0:00")`, 0},
		{"23:59:59", `TIMEVALUE("23:59:59")`, 0.999988425925926},
		{"1:30:45", `TIMEVALUE("1:30:45")`, 0.06302083333333333},
		{"12:00 AM", `TIMEVALUE("12:00 AM")`, 0},
		{"12:00 PM", `TIMEVALUE("12:00 PM")`, 0.5},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: got type %v (%v), want number", tc.formula, got.Type, got)
			}
			if math.Abs(got.Num-tc.want) > 1e-12 {
				t.Errorf("%s = %.18g, want %.18g", tc.formula, got.Num, tc.want)
			}
		})
	}
}

func TestDATEVALUE_extended(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
	}{
		{"two_digit_year", `DATEVALUE("03/04/25")`, 45720},
		{"date_with_time", `DATEVALUE("2025-03-04 12:00")`, 45720},
		{"month_day_only", `DATEVALUE("March 4")`, func() float64 {
			// Use current year
			now := time.Now()
			t := time.Date(now.Year(), 3, 4, 0, 0, 0, 0, time.UTC)
			return math.Floor(TimeToExcelSerial(t))
		}()},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tc.formula, err)
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: got type %v (%v), want number", tc.formula, got.Type, got)
			}
			if got.Num != tc.want {
				t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
			}
		})
	}
}

func TestDATEVALUEComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		// Various date formats
		{"slash_m_d_yyyy", `DATEVALUE("1/1/2023")`, 44927, false, 0},
		{"slash_mm_dd_yyyy", `DATEVALUE("01/01/2023")`, 44927, false, 0},
		{"month_name_full", `DATEVALUE("January 1, 2023")`, 44927, false, 0},
		{"iso_format", `DATEVALUE("2023-01-01")`, 44927, false, 0},
		{"slash_yyyy_mm_dd", `DATEVALUE("2023/01/01")`, 44927, false, 0},
		{"dash_d_mon_yyyy", `DATEVALUE("1-Jan-2023")`, 44927, false, 0},
		{"dash_dd_mon_yyyy", `DATEVALUE("01-Jan-2023")`, 44927, false, 0},

		// End of year
		{"dec_31_2023", `DATEVALUE("12/31/2023")`, 45291, false, 0},

		// Known serial: 1/1/1900 = 1
		{"jan_1_1900", `DATEVALUE("1/1/1900")`, 1, false, 0},

		// After the leap year bug: 3/1/1900 = 61
		{"mar_1_1900", `DATEVALUE("3/1/1900")`, 61, false, 0},

		// Leap year date
		{"leap_year_feb29_2024", `DATEVALUE("2/29/2024")`, 45351, false, 0},
		{"leap_year_feb29_2000", `DATEVALUE("2/29/2000")`, 36585, false, 0},

		// Excel doc examples
		{"doc_8_22_2011", `DATEVALUE("8/22/2011")`, 40777, false, 0},
		{"doc_22_may_2011", `DATEVALUE("22-May-2011")`, 40685, false, 0},
		{"doc_2011_02_23", `DATEVALUE("2011/02/23")`, 40597, false, 0},
		{"doc_jan1_2008", `DATEVALUE("1/1/2008")`, 39448, false, 0},

		// Two-digit years
		{"two_digit_year_25", `DATEVALUE("01/01/25")`, 45658, false, 0},
		{"two_digit_year_99", `DATEVALUE("12/31/99")`, 36525, false, 0},
		{"two_digit_year_00", `DATEVALUE("1/1/00")`, 36526, false, 0},

		// Date with time portion (time should be ignored)
		{"datetime_iso_with_time", `DATEVALUE("2023-06-15 14:30")`, 45092, false, 0},
		{"datetime_with_seconds", `DATEVALUE("2023-06-15 14:30:45")`, 45092, false, 0},

		// Mid-year dates
		{"jul_4_2000", `DATEVALUE("7/4/2000")`, 36711, false, 0},
		{"feb_28_1900", `DATEVALUE("2/28/1900")`, 59, false, 0},

		// Invalid date string → #VALUE!
		{"invalid_string", `DATEVALUE("not a date")`, 0, true, ErrValVALUE},
		{"invalid_gibberish", `DATEVALUE("abc123")`, 0, true, ErrValVALUE},

		// Empty string → #VALUE!
		{"empty_string", `DATEVALUE("")`, 0, true, ErrValVALUE},

		// Non-date string → #VALUE!
		{"non_date_hello", `DATEVALUE("hello world")`, 0, true, ErrValVALUE},

		// Number input: DATEVALUE coerces to string, which won't parse as a date format
		{"number_input", `DATEVALUE(12345)`, 0, true, ErrValVALUE},
		{"number_zero", `DATEVALUE(0)`, 0, true, ErrValVALUE},

		// Too few args → #VALUE!
		{"too_few_args", `DATEVALUE()`, 0, true, ErrValVALUE},

		// Too many args → #VALUE!
		{"too_many_args", `DATEVALUE("1/1/2023","extra")`, 0, true, ErrValVALUE},

		// Error propagation
		{"error_propagation_ref", `DATEVALUE(1/0)`, 0, true, ErrValDIV0},

		// Boolean input (TRUE → "1", FALSE → "0", neither parses as date)
		{"bool_true", `DATEVALUE(TRUE)`, 0, true, ErrValVALUE},
		{"bool_false", `DATEVALUE(FALSE)`, 0, true, ErrValVALUE},
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
				t.Fatalf("%s: got type %v (%v), want number", tc.formula, got.Type, got)
			}
			if got.Num != tc.want {
				t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
			}
		})
	}
}

func TestWORKDAY(t *testing.T) {
	resolver := &mockResolver{}

	// Key dates (2025) — serial numbers:
	// Jan 1  = 45658 (Wed), Jan 2  = 45659 (Thu), Jan 3  = 45660 (Fri)
	// Jan 4  = 45661 (Sat), Jan 5  = 45662 (Sun), Jan 6  = 45663 (Mon)
	// Jan 7  = 45664 (Tue), Jan 8  = 45665 (Wed), Jan 9  = 45666 (Thu)
	// Jan 10 = 45667 (Fri), Jan 13 = 45670 (Mon), Jan 15 = 45672 (Wed)
	//
	// Excel doc dates:
	// 2008-10-01 (Wed) = 39722 (start), 2009-04-30 (Thu) = 39933 (result no holidays)
	// 2009-05-05 (Tue) = 39938 (result with holidays)
	// Holidays: 2008-11-26 = 39778, 2008-12-04 = 39786, 2009-01-21 = 39834

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		// --- Basic: add workdays, result skips weekends ---
		// 1 workday from Wed Jan 1 → Thu Jan 2
		{"basic_1day_wed", "WORKDAY(45658,1)", 45659, false, 0},
		// 2 workdays from Wed Jan 1 → Fri Jan 3
		{"basic_2days_wed", "WORKDAY(45658,2)", 45660, false, 0},

		// --- Add 0 days → same date (if weekday) ---
		{"zero_days_weekday", "WORKDAY(45658,0)", 45658, false, 0},
		// Add 0 days on Saturday → returns Saturday serial (implementation returns start)
		{"zero_days_saturday", "WORKDAY(45661,0)", 45661, false, 0},

		// --- Add 1 day on Friday → Monday ---
		{"friday_plus_1", "WORKDAY(45660,1)", 45663, false, 0},

		// --- Add 5 days → one week later (skipping weekend) ---
		// Wed Jan 1 + 5 workdays: Thu(1), Fri(2), Mon(3), Tue(4), Wed(5) = Jan 8
		{"five_days_skip_weekend", "WORKDAY(45658,5)", 45665, false, 0},

		// --- Add 10 days → two weeks later ---
		// Wed Jan 1 + 10 workdays = Wed Jan 15
		{"ten_days", "WORKDAY(45658,10)", 45672, false, 0},

		// --- Negative days (go backward) ---
		// Mon Jan 6 - 1 workday → Fri Jan 3
		{"negative_1", "WORKDAY(45663,-1)", 45660, false, 0},

		// --- Negative days crossing weekend ---
		// Mon Jan 6 - 3 workdays: Fri(1), Thu(2), Wed(3) = Jan 1
		{"negative_cross_weekend", "WORKDAY(45663,-3)", 45658, false, 0},
		// Wed Jan 8 - 5 workdays: Tue(1), Mon(2), Fri(3), Thu(4), Wed(5) = Jan 1
		{"negative_5_cross_weekend", "WORKDAY(45665,-5)", 45658, false, 0},

		// --- With holidays (skips them) ---
		// Wed Jan 1 + 5, holiday on Thu Jan 2: Fri(1), Mon(2), Tue(3), Wed(4), Thu Jan 9(5) = 45666
		{"with_one_holiday", "WORKDAY(45658,5,45659)", 45666, false, 0},

		// --- Holiday on weekend (no extra skip) ---
		// Wed Jan 1 + 5, holiday on Sat Jan 4: same as no holidays = Jan 8
		{"holiday_on_weekend", "WORKDAY(45658,5,45661)", 45665, false, 0},

		// --- Multiple holidays ---
		// Wed Jan 1 + 5, holidays on Thu Jan 2 and Fri Jan 3:
		// Mon(1), Tue(2), Wed(3), Thu(4), Fri Jan 10(5) = 45667
		{"multiple_holidays", "WORKDAY(45658,5,{45659,45660})", 45667, false, 0},

		// --- Start on weekend ---
		// Sat Jan 4 + 1 workday → Mon Jan 6
		{"start_saturday_plus1", "WORKDAY(45661,1)", 45663, false, 0},
		// Sun Jan 5 + 1 workday → Mon Jan 6
		{"start_sunday_plus1", "WORKDAY(45662,1)", 45663, false, 0},
		// Sat Jan 4 - 1 workday → Fri Jan 3
		{"start_saturday_minus1", "WORKDAY(45661,-1)", 45660, false, 0},

		// --- Large number of days ---
		// 250 workdays from Wed Jan 1 2025 = Wed Dec 17 2025 = 46008
		{"large_positive", "WORKDAY(45658,250)", 46008, false, 0},

		// --- Too few args → error ---
		{"too_few_args_zero", "WORKDAY()", 0, true, ErrValVALUE},
		{"too_few_args_one", "WORKDAY(45658)", 0, true, ErrValVALUE},

		// --- Too many args → error ---
		{"too_many_args", "WORKDAY(45658,5,45659,1)", 0, true, ErrValVALUE},

		// --- Excel doc examples ---
		// WORKDAY(10/1/2008, 151) = 4/30/2009 = 39933
		{"excel_doc_no_holidays", "WORKDAY(39722,151)", 39933, false, 0},
		// WORKDAY(10/1/2008, 151, {11/26/2008,12/4/2008,1/21/2009}) = 5/5/2009 = 39938
		{"excel_doc_with_holidays", "WORKDAY(39722,151,{39778,39786,39834})", 39938, false, 0},
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
					t.Fatalf("%s: want error %v, got %v", tc.formula, tc.errVal, got)
				}
				return
			}
			if got.Type != ValueNumber {
				t.Fatalf("%s: got type %v (%v), want number", tc.formula, got.Type, got)
			}
			if got.Num != tc.want {
				t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
			}
		})
	}
}

func TestWORKDAY_INTL(t *testing.T) {
	resolver := &mockResolver{}

	// Reference dates:
	// DATE(2025,1,1) = 45658 (Wednesday)
	// DATE(2025,1,2) = 45659 (Thursday)
	// DATE(2025,1,3) = 45660 (Friday)
	// DATE(2025,1,4) = 45661 (Saturday)
	// DATE(2025,1,5) = 45662 (Sunday)
	// DATE(2025,1,6) = 45663 (Monday)
	// DATE(2025,1,7) = 45664 (Tuesday)
	// DATE(2025,1,8) = 45665 (Wednesday)
	// DATE(2025,1,9) = 45666 (Thursday)
	// DATE(2025,1,10) = 45667 (Friday)
	// DATE(2025,1,13) = 45670 (Monday)
	// DATE(2025,1,15) = 45672 (Wednesday)

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		// --- Default weekend (Sat/Sun, code 1) ---
		// 5 workdays from Wed Jan 1: Thu(1), Fri(2), Mon(3), Tue(4), Wed(5) = Jan 8
		{"default_5days", "WORKDAY.INTL(DATE(2025,1,1),5)", 45665, false, 0},
		// 1 workday from Wed Jan 1 → Thu Jan 2
		{"default_1day", "WORKDAY.INTL(DATE(2025,1,1),1)", 45659, false, 0},
		// 10 workdays from Wed Jan 1 → Wed Jan 15
		{"default_10days", "WORKDAY.INTL(DATE(2025,1,1),10)", 45672, false, 0},

		// --- Explicit weekend=1 (same as default) ---
		{"weekend1_5days", "WORKDAY.INTL(DATE(2025,1,1),5,1)", 45665, false, 0},

		// --- Weekend=2 (Sun, Mon) ---
		// From Wed Jan 1: Thu(1), Fri(2), Sat(3), Tue(4), Wed(5) = Jan 8
		{"weekend2_5days", "WORKDAY.INTL(DATE(2025,1,1),5,2)", 45665, false, 0},

		// --- Weekend=11 (Sunday only) ---
		// From Wed Jan 1: Thu(1), Fri(2), Sat(3), Mon(4), Tue(5) = Jan 7
		{"weekend11_5days", "WORKDAY.INTL(DATE(2025,1,1),5,11)", 45664, false, 0},

		// --- Weekend=17 (Saturday only) ---
		// From Wed Jan 1: Thu(1), Fri(2), Sun(3), Mon(4), Tue(5) = Jan 7
		{"weekend17_5days", "WORKDAY.INTL(DATE(2025,1,1),5,17)", 45664, false, 0},

		// --- Weekend=12 (Monday only) ---
		// From Wed Jan 1: Thu(1), Fri(2), Sat(3), Sun(4), Tue(5) = Jan 7
		{"weekend12_5days", "WORKDAY.INTL(DATE(2025,1,1),5,12)", 45664, false, 0},

		// --- String weekend: "0000011" (Sat+Sun off, same as default) ---
		{"string_satsun", `WORKDAY.INTL(DATE(2025,1,1),5,"0000011")`, 45665, false, 0},

		// --- String weekend: "1000001" (Mon+Sun off) ---
		// From Wed Jan 1: Thu(1), Fri(2), Sat(3), Tue(4), Wed(5) = Jan 8
		{"string_monsun", `WORKDAY.INTL(DATE(2025,1,1),5,"1000001")`, 45665, false, 0},

		// --- String weekend: "0000000" (no weekends, every day is a workday) ---
		// From Wed Jan 1: Thu(1), Fri(2), Sat(3), Sun(4), Mon(5) = Jan 6
		{"string_noweekend", `WORKDAY.INTL(DATE(2025,1,1),5,"0000000")`, 45663, false, 0},

		// --- String weekend: "1111100" (Mon-Fri off, Sat+Sun are workdays) ---
		// From Wed Jan 1: Sat Jan 4(1), Sun Jan 5(2) = 45662
		{"string_weekdays_off", `WORKDAY.INTL(DATE(2025,1,1),2,"1111100")`, 45662, false, 0},

		// --- With holidays ---
		// 5 workdays from Jan 1, holiday on Jan 2 (Thu):
		// Fri(1), Mon(2), Tue(3), Wed(4), Thu Jan 9(5) = 45666
		{"with_holiday", "WORKDAY.INTL(DATE(2025,1,1),5,1,DATE(2025,1,2))", 45666, false, 0},

		// Holiday that falls on a weekend should not matter
		// 5 workdays from Jan 1, holiday on Jan 4 (Sat — already weekend):
		// Same as default: Thu(1), Fri(2), Mon(3), Tue(4), Wed(5) = Jan 8
		{"holiday_on_weekend", "WORKDAY.INTL(DATE(2025,1,1),5,1,DATE(2025,1,4))", 45665, false, 0},

		// --- Negative days ---
		// -1 workday from Thu Jan 2 → Wed Jan 1
		{"negative_1", "WORKDAY.INTL(DATE(2025,1,2),-1)", 45658, false, 0},
		// -5 workdays from Wed Jan 8 → Wed Jan 1
		{"negative_5", "WORKDAY.INTL(DATE(2025,1,8),-5)", 45658, false, 0},
		// -5 workdays from Wed Jan 8 with weekend=11 (Sun only):
		// Tue(1), Mon(2), Sat(3), Fri(4), Thu(5) = Jan 2
		{"negative_weekend11", "WORKDAY.INTL(DATE(2025,1,8),-5,11)", 45659, false, 0},

		// --- Zero days ---
		{"zero_days", "WORKDAY.INTL(DATE(2025,1,1),0)", 45658, false, 0},

		// --- Invalid weekend: "1111111" (all days off) ---
		{"invalid_all_ones", `WORKDAY.INTL(DATE(2025,1,1),5,"1111111")`, 0, true, ErrValVALUE},

		// --- Invalid weekend: code 0 ---
		{"invalid_code_0", "WORKDAY.INTL(DATE(2025,1,1),5,0)", 0, true, ErrValVALUE},

		// --- Invalid weekend: code 8 ---
		{"invalid_code_8", "WORKDAY.INTL(DATE(2025,1,1),5,8)", 0, true, ErrValVALUE},

		// --- Invalid weekend: code 9 ---
		{"invalid_code_9", "WORKDAY.INTL(DATE(2025,1,1),5,9)", 0, true, ErrValVALUE},

		// --- Invalid weekend: code 10 ---
		{"invalid_code_10", "WORKDAY.INTL(DATE(2025,1,1),5,10)", 0, true, ErrValVALUE},

		// --- Invalid weekend: code 18 ---
		{"invalid_code_18", "WORKDAY.INTL(DATE(2025,1,1),5,18)", 0, true, ErrValVALUE},

		// --- Invalid weekend string: wrong length ---
		{"invalid_string_short", `WORKDAY.INTL(DATE(2025,1,1),5,"001")`, 0, true, ErrValVALUE},

		// --- Invalid weekend string: bad chars ---
		{"invalid_string_chars", `WORKDAY.INTL(DATE(2025,1,1),5,"000002a")`, 0, true, ErrValVALUE},

		// --- Large positive day count ---
		// 250 workdays from Jan 1 2025 (default weekend, no holidays)
		{"large_positive", "WORKDAY.INTL(DATE(2025,1,1),250)", 46008, false, 0},

		// --- Weekend=3 (Mon, Tue) ---
		// From Wed Jan 1: Thu(1), Fri(2), Sat(3), Sun(4), Wed(5) = Jan 8
		{"weekend3_5days", "WORKDAY.INTL(DATE(2025,1,1),5,3)", 45665, false, 0},

		// --- Weekend=7 (Fri, Sat) ---
		// From Wed Jan 1: Thu(1), Sun(2), Mon(3), Tue(4), Wed(5) = Jan 8
		{"weekend7_5days", "WORKDAY.INTL(DATE(2025,1,1),5,7)", 45665, false, 0},

		// --- Too few args ---
		{"too_few_args", "WORKDAY.INTL(DATE(2025,1,1))", 0, true, ErrValVALUE},

		// --- Weekend=4 (Tue, Wed) ---
		// From Wed Jan 1: Thu(1), Fri(2), Sat(3), Sun(4), Mon(5) = Jan 6
		{"weekend4_5days", "WORKDAY.INTL(DATE(2025,1,1),5,4)", 45663, false, 0},

		// --- Weekend=5 (Wed, Thu) ---
		// From Wed Jan 1: Fri(1), Sat(2), Sun(3), Mon(4), Tue(5) = Jan 7
		{"weekend5_5days", "WORKDAY.INTL(DATE(2025,1,1),5,5)", 45664, false, 0},

		// --- Weekend=6 (Thu, Fri) ---
		// From Wed Jan 1: Sat(1), Sun(2), Mon(3), Tue(4), Wed(5) = Jan 8
		{"weekend6_5days", "WORKDAY.INTL(DATE(2025,1,1),5,6)", 45665, false, 0},
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
				t.Fatalf("%s: got type %v (%v), want number", tc.formula, got.Type, got)
			}
			if got.Num != tc.want {
				t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
			}
		})
	}
}

func TestNETWORKDAYS(t *testing.T) {
	resolver := &mockResolver{}

	// Key dates (2025) — serial numbers per ExcelSerialToTime:
	// Jan 1  = 45658 (Wed), Jan 3  = 45660 (Fri), Jan 4  = 45661 (Sat)
	// Jan 5  = 45662 (Sun), Jan 6  = 45663 (Mon), Jan 7  = 45664 (Tue)
	// Jan 8  = 45665 (Wed), Jan 10 = 45667 (Fri), Jan 11 = 45668 (Sat)
	// Jan 12 = 45669 (Sun), Jan 13 = 45670 (Mon), Jan 31 = 45688 (Fri)
	//
	// Excel doc example dates (serial numbers):
	// 2012-10-01 (Mon) = 41183  (project start)
	// 2013-03-01 (Fri) = 41334  (project end)
	// 2012-11-22 (Thu) = 41235  (holiday 1)
	// 2012-12-04 (Tue) = 41247  (holiday 2)
	// 2013-01-21 (Mon) = 41295  (holiday 3)

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		// Basic: same week Monday–Friday
		{"mon_to_fri_same_week", "NETWORKDAYS(45663,45667)", 5, false, 0},

		// Cross-weekend: Friday to Monday
		{"fri_to_mon", "NETWORKDAYS(45660,45663)", 2, false, 0},

		// Cross-weekend: Monday to next Monday
		{"mon_to_next_mon", "NETWORKDAYS(45663,45670)", 6, false, 0},

		// Same day (weekday) → 1
		{"same_day_weekday_wed", "NETWORKDAYS(45658,45658)", 1, false, 0},
		{"same_day_weekday_mon", "NETWORKDAYS(45663,45663)", 1, false, 0},
		{"same_day_weekday_fri", "NETWORKDAYS(45667,45667)", 1, false, 0},

		// Same day (weekend) → 0
		{"same_day_saturday", "NETWORKDAYS(45661,45661)", 0, false, 0},
		{"same_day_sunday", "NETWORKDAYS(45662,45662)", 0, false, 0},

		// Start and end on weekend → count only weekdays between
		{"sat_to_sun_same_weekend", "NETWORKDAYS(45661,45662)", 0, false, 0},
		{"sat_to_next_sat", "NETWORKDAYS(45661,45668)", 5, false, 0},
		{"sun_to_next_sun", "NETWORKDAYS(45662,45669)", 5, false, 0},

		// One full week (Mon–Fri) → 5
		{"one_full_week", "NETWORKDAYS(45663,45667)", 5, false, 0},

		// One month: Jan 1 (Wed) to Jan 31 (Fri)
		{"one_month_jan", "NETWORKDAYS(45658,45688)", 23, false, 0},

		// With holidays: one holiday on a weekday
		{"one_holiday", "NETWORKDAYS(45663,45667,45665)", 4, false, 0},

		// Holiday on weekend (doesn't reduce count)
		{"holiday_on_saturday", "NETWORKDAYS(45663,45667,45661)", 5, false, 0},
		{"holiday_on_sunday", "NETWORKDAYS(45663,45667,45662)", 5, false, 0},

		// Multiple holidays
		{"two_holidays", "NETWORKDAYS(45663,45667,{45664,45665})", 3, false, 0},

		// Negative result (end before start)
		{"negative_range", "NETWORKDAYS(45667,45663)", -5, false, 0},
		{"negative_cross_weekend", "NETWORKDAYS(45670,45660)", -7, false, 0},

		// Excel doc examples:
		// NETWORKDAYS(10/1/2012, 3/1/2013) = 110
		{"excel_doc_no_holidays", "NETWORKDAYS(41183,41334)", 110, false, 0},
		// NETWORKDAYS(10/1/2012, 3/1/2013, 11/22/2012) = 109
		{"excel_doc_one_holiday", "NETWORKDAYS(41183,41334,41235)", 109, false, 0},
		// NETWORKDAYS(10/1/2012, 3/1/2013, {11/22/2012,12/4/2012,1/21/2013}) = 107
		{"excel_doc_three_holidays", "NETWORKDAYS(41183,41334,{41235,41247,41295})", 107, false, 0},

		// Too few args → error
		{"too_few_args_zero", "NETWORKDAYS()", 0, true, ErrValVALUE},
		{"too_few_args_one", "NETWORKDAYS(45663)", 0, true, ErrValVALUE},

		// Too many args → error
		{"too_many_args", "NETWORKDAYS(45663,45667,45665,1)", 0, true, ErrValVALUE},
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
				t.Fatalf("%s: got type %v (%v), want number", tc.formula, got.Type, got)
			}
			if got.Num != tc.want {
				t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
			}
		})
	}
}

func TestNETWORKDAYS_INTL(t *testing.T) {
	resolver := &mockResolver{}

	// Key dates (2025):
	// Jan 1 = 45658 (Wed), Jan 4 = 45661 (Sat), Jan 5 = 45662 (Sun)
	// Jan 10 = 45667 (Fri), Jan 31 = 45688 (Fri), Mar 15 = 45731 (Sat)

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		// --- Default weekend (Sat/Sun, code 1) ---
		{"default_basic", "NETWORKDAYS.INTL(45658,45667)", 8, false, 0},
		{"default_explicit_1", "NETWORKDAYS.INTL(45658,45667,1)", 8, false, 0},
		{"default_month", "NETWORKDAYS.INTL(45658,45688)", 23, false, 0},

		// --- Numeric weekend codes (two-day) ---
		{"weekend_2_sun_mon", "NETWORKDAYS.INTL(45658,45667,2)", 8, false, 0},
		{"weekend_3_mon_tue", "NETWORKDAYS.INTL(45658,45667,3)", 8, false, 0},
		{"weekend_4_tue_wed", "NETWORKDAYS.INTL(45658,45667,4)", 7, false, 0},
		{"weekend_5_wed_thu", "NETWORKDAYS.INTL(45658,45667,5)", 6, false, 0},
		{"weekend_6_thu_fri", "NETWORKDAYS.INTL(45658,45667,6)", 6, false, 0},
		{"weekend_7_fri_sat", "NETWORKDAYS.INTL(45658,45667,7)", 7, false, 0},

		// --- Numeric weekend codes (single-day) ---
		{"weekend_11_sun_only", "NETWORKDAYS.INTL(45658,45667,11)", 9, false, 0},
		{"weekend_12_mon_only", "NETWORKDAYS.INTL(45658,45667,12)", 9, false, 0},
		{"weekend_13_tue_only", "NETWORKDAYS.INTL(45658,45667,13)", 9, false, 0},
		{"weekend_14_wed_only", "NETWORKDAYS.INTL(45658,45667,14)", 8, false, 0},
		{"weekend_15_thu_only", "NETWORKDAYS.INTL(45658,45667,15)", 8, false, 0},
		{"weekend_16_fri_only", "NETWORKDAYS.INTL(45658,45667,16)", 8, false, 0},
		{"weekend_17_sat_only", "NETWORKDAYS.INTL(45658,45667,17)", 9, false, 0},

		// --- String weekend format ---
		{"string_0000011_sat_sun", `NETWORKDAYS.INTL(45658,45667,"0000011")`, 8, false, 0},
		{"string_1000001_sun_mon", `NETWORKDAYS.INTL(45658,45667,"1000001")`, 8, false, 0},
		{"string_0000000_no_weekends", `NETWORKDAYS.INTL(45658,45667,"0000000")`, 10, false, 0},
		{"string_0000110_fri_sat", `NETWORKDAYS.INTL(45658,45667,"0000110")`, 7, false, 0},

		// --- With holidays ---
		{"holiday_jan1", "NETWORKDAYS.INTL(45658,45667,1,45658)", 7, false, 0},
		{"holiday_jan1_jan2", "NETWORKDAYS.INTL(45658,45667,1,{45658,45659})", 6, false, 0},
		{"holiday_on_weekend", "NETWORKDAYS.INTL(45658,45667,1,45661)", 8, false, 0},
		{"holiday_custom_weekend", "NETWORKDAYS.INTL(45658,45667,2,45658)", 7, false, 0},

		// --- Negative range ---
		{"negative_range_default", "NETWORKDAYS.INTL(45667,45658)", -8, false, 0},
		{"negative_range_sun_only", "NETWORKDAYS.INTL(45667,45658,11)", -9, false, 0},
		{"negative_range_sat_only", "NETWORKDAYS.INTL(45667,45658,17)", -9, false, 0},

		// --- Same date ---
		{"same_date_workday", "NETWORKDAYS.INTL(45658,45658)", 1, false, 0},
		{"same_date_saturday", "NETWORKDAYS.INTL(45661,45661)", 0, false, 0},
		{"same_date_sunday", "NETWORKDAYS.INTL(45662,45662)", 0, false, 0},
		{"same_date_custom_workday", "NETWORKDAYS.INTL(45661,45661,11)", 1, false, 0},

		// --- End on weekend ---
		{"end_on_saturday", "NETWORKDAYS.INTL(45658,45731)", 53, false, 0},

		// --- Matches NETWORKDAYS behavior ---
		{"matches_networkdays", "NETWORKDAYS.INTL(45658,45667,1)", 8, false, 0},

		// --- Invalid weekend codes ---
		{"invalid_weekend_0", "NETWORKDAYS.INTL(45658,45667,0)", 0, true, ErrValVALUE},
		{"invalid_weekend_8", "NETWORKDAYS.INTL(45658,45667,8)", 0, true, ErrValVALUE},
		{"invalid_weekend_9", "NETWORKDAYS.INTL(45658,45667,9)", 0, true, ErrValVALUE},
		{"invalid_weekend_10", "NETWORKDAYS.INTL(45658,45667,10)", 0, true, ErrValVALUE},
		{"invalid_weekend_18", "NETWORKDAYS.INTL(45658,45667,18)", 0, true, ErrValVALUE},

		// --- Invalid string weekends ---
		{"invalid_string_all_ones", `NETWORKDAYS.INTL(45658,45667,"1111111")`, 0, true, ErrValVALUE},
		{"invalid_string_too_short", `NETWORKDAYS.INTL(45658,45667,"000011")`, 0, true, ErrValVALUE},
		{"invalid_string_too_long", `NETWORKDAYS.INTL(45658,45667,"00000110")`, 0, true, ErrValVALUE},
		{"invalid_string_bad_char", `NETWORKDAYS.INTL(45658,45667,"000001x")`, 0, true, ErrValVALUE},

		// --- Too few args ---
		{"too_few_args", "NETWORKDAYS.INTL(45658)", 0, true, ErrValVALUE},
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
				t.Fatalf("%s: got type %v (%v), want number", tc.formula, got.Type, got)
			}
			if got.Num != tc.want {
				t.Errorf("%s = %g, want %g", tc.formula, got.Num, tc.want)
			}
		})
	}
}
