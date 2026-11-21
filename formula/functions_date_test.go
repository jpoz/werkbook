package formula

import (
	"math"
	"testing"
	"time"
)

func dateSerial(year int, month time.Month, day int) float64 {
	return math.Floor(TimeToSerial(time.Date(year, month, day, 0, 0, 0, 0, time.UTC)))
}

func TestParseHolidays_FullColumnSparseRef(t *testing.T) {
	holidays, errVal := parseHolidays(sparseFullColumnRefValue(1, map[int]Value{
		1: NumberVal(45659),
		3: NumberVal(45661),
	}))
	if errVal != nil {
		t.Fatalf("parseHolidays: %v", *errVal)
	}
	if !holidays[45659] || !holidays[45661] {
		t.Fatalf("holidays = %#v, want entries for 45659 and 45661", holidays)
	}
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

func TestDateFunctionsWith1904System(t *testing.T) {
	resolver := &mockResolver{}
	ctx1904 := &EvalContext{Date1904: true}

	tests := []struct {
		name    string
		formula string
		want    Value
	}{
		{name: "DATE", formula: "DATE(2024,1,31)", want: NumberVal(43860)},
		{name: "DAY", formula: "DAY(43889)", want: NumberVal(29)},
		{name: "EDATE", formula: "EDATE(43860,1)", want: NumberVal(43889)},
		{name: "EOMONTH", formula: "EOMONTH(43860,0)", want: NumberVal(43860)},
		{name: "DATEVALUE", formula: `DATEVALUE("2024-02-29")`, want: NumberVal(43889)},
		{name: "WORKDAY.INTL", formula: "WORKDAY.INTL(43860,10,1)", want: NumberVal(43874)},
		{name: "WEEKDAY", formula: "WEEKDAY(43874,2)", want: NumberVal(3)},
		{name: "NETWORKDAYS.INTL", formula: "NETWORKDAYS.INTL(43860,43904,1)", want: NumberVal(33)},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, ctx1904)
			if err != nil {
				t.Fatalf("Eval(%q): %v", tt.formula, err)
			}
			if got.Type != tt.want.Type {
				t.Fatalf("Eval(%q) type = %v, want %v", tt.formula, got.Type, tt.want.Type)
			}
			switch got.Type {
			case ValueNumber:
				if got.Num != tt.want.Num {
					t.Fatalf("Eval(%q) = %v, want %v", tt.formula, got.Num, tt.want.Num)
				}
			default:
				t.Fatalf("unexpected test value type %v", got.Type)
			}
		})
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

		// Doc examples
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
		{"year_1900_feb29_fictional", "DATE(1900,2,29)", 60, false, 0}, // Fictional leap day
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
		{"non_leap_1900_feb29", "DATE(1900,2,29)", 60, false, 0}, // Fictional

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
		// Serial 0 → 1900 ("January 0, 1900" sentinel)
		{"serial_0", "YEAR(0)", 1900, false, 0},
		// Serial 60 → 1900 (fictional Feb 29, 1900)
		{"serial_60", "YEAR(60)", 1900, false, 0},
		// String date input via DATEVALUE: YEAR(DATEVALUE("1/1/2023"))
		{"string_date_via_datevalue", `YEAR(DATEVALUE("1/1/2023"))`, 2023, false, 0},
		// String date input directly: coerceDateNum parses date strings
		{"string_date_direct", `YEAR("1/1/2023")`, 2023, false, 0},
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
		// Doc examples via DATEVALUE
		{"doc_2023", `YEAR(DATEVALUE("7/5/2023"))`, 2023, false, 0},
		{"doc_2025", `YEAR(DATEVALUE("7/5/2025"))`, 2025, false, 0},
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

func TestMONTH_Comprehensive(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		// --- Basic happy path: each month Jan through Dec ---
		{"jan_date_func", "MONTH(DATE(2024,1,15))", 1, false, 0},
		{"feb_date_func", "MONTH(DATE(2024,2,10))", 2, false, 0},
		{"mar_date_func", "MONTH(DATE(2024,3,1))", 3, false, 0},
		{"apr_date_func", "MONTH(DATE(2024,4,30))", 4, false, 0},
		{"may_date_func", "MONTH(DATE(2023,5,23))", 5, false, 0},
		{"jun_date_func", "MONTH(DATE(2023,6,15))", 6, false, 0},
		{"jul_date_func", "MONTH(DATE(2023,7,4))", 7, false, 0},
		{"aug_date_func", "MONTH(DATE(2023,8,31))", 8, false, 0},
		{"sep_date_func", "MONTH(DATE(2025,9,1))", 9, false, 0},
		{"oct_date_func", "MONTH(DATE(2025,10,15))", 10, false, 0},
		{"nov_date_func", "MONTH(DATE(2025,11,30))", 11, false, 0},
		{"dec_date_func", "MONTH(DATE(2025,12,25))", 12, false, 0},

		// --- Serial number inputs ---
		{"serial_1_jan_1900", "MONTH(1)", 1, false, 0},         // Jan 1, 1900
		{"serial_45306_jan_2024", "MONTH(45306)", 1, false, 0}, // Jan 15, 2024
		{"serial_44927_jan_2023", "MONTH(44927)", 1, false, 0}, // Jan 1, 2023
		{"serial_32_feb_1_1900", "MONTH(32)", 2, false, 0},     // Feb 1, 1900
		{"serial_59_feb_28_1900", "MONTH(59)", 2, false, 0},    // Feb 28, 1900
		{"serial_61_mar_1_1900", "MONTH(61)", 3, false, 0},     // Mar 1, 1900

		// --- Boundary: serial 0 (Jan 0, 1900 sentinel) ---
		{"serial_0_jan", "MONTH(0)", 1, false, 0},

		// --- Boundary: serial 60 (fictional Feb 29, 1900 — Excel's leap year bug) ---
		{"serial_60_feb_29_1900", "MONTH(60)", 2, false, 0},

		// --- Fractional serial numbers (time component ignored) ---
		{"fractional_0_25", "MONTH(44927.25)", 1, false, 0},
		{"fractional_0_5", "MONTH(44927.5)", 1, false, 0},
		{"fractional_0_75", "MONTH(44927.75)", 1, false, 0},
		{"fractional_0_99", "MONTH(44927.99)", 1, false, 0},

		// --- Very large serial number (far future) ---
		{"max_serial_dec_9999", "MONTH(2958465)", 12, false, 0}, // Dec 31, 9999

		// --- Boolean inputs ---
		{"bool_true_serial_1", "MONTH(TRUE)", 1, false, 0},   // TRUE coerces to 1 = Jan 1, 1900
		{"bool_false_serial_0", "MONTH(FALSE)", 1, false, 0}, // FALSE coerces to 0 = Jan 0, 1900

		// --- Numeric string coercion ---
		{"string_numeric_44927", `MONTH("44927")`, 1, false, 0},
		{"string_numeric_1", `MONTH("1")`, 1, false, 0},

		// --- Error: negative serial numbers ---
		{"negative_serial", "MONTH(-1)", 0, true, ErrValNUM},

		// --- Error: beyond max serial ---
		{"beyond_max_serial", "MONTH(2958466)", 0, true, ErrValNUM},

		// --- Error: wrong argument count ---
		{"no_args", "MONTH()", 0, true, ErrValVALUE},
		{"two_args", "MONTH(1,2)", 0, true, ErrValVALUE},

		// --- Error: non-numeric string ---
		{"non_date_string", `MONTH("hello")`, 0, true, ErrValVALUE},
		{"empty_string", `MONTH("")`, 0, true, ErrValVALUE},

		// --- Error propagation ---
		{"div_by_zero", "MONTH(1/0)", 0, true, ErrValDIV0},

		// --- Date boundary: Dec 31 of a year, next day is Jan of next year ---
		{"dec_31_2023", "MONTH(DATE(2023,12,31))", 12, false, 0},
		{"jan_1_2024", "MONTH(DATE(2024,1,1))", 1, false, 0},

		// --- Leap year Feb 29 (real) ---
		{"leap_feb_29_2024", "MONTH(DATE(2024,2,29))", 2, false, 0},
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

func TestDAY_Comprehensive(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name         string
		formula      string
		want         float64
		wantErr      bool
		wantErrValue ErrorValue
	}{
		{name: "serial_example", formula: "DAY(45306)", want: 15},
		{name: "serial_zero", formula: "DAY(0)", want: 0},
		{name: "serial_59", formula: "DAY(59)", want: 28},
		{name: "serial_60", formula: "DAY(60)", want: 29},
		{name: "serial_61", formula: "DAY(61)", want: 1},
		{name: "fractional_serial", formula: "DAY(45306.75)", want: 15},
		{name: "max_serial", formula: "DAY(2958465)", want: 31},
		{name: "bool_true", formula: "DAY(TRUE)", want: 1},
		{name: "bool_false", formula: "DAY(FALSE)", want: 0},
		{name: "numeric_string", formula: `DAY("45306")`, want: 15},
		{name: "slash_date_string", formula: `DAY("1/15/2024")`, want: 15},
		{name: "dash_date_string", formula: `DAY("15-Jan-2024")`, want: 15},
		{name: "iso_date_string", formula: `DAY("2024-01-31")`, want: 31},
		{name: "long_date_string", formula: `DAY("January 2, 2024")`, want: 2},
		{name: "error_div0", formula: "DAY(1/0)", wantErr: true, wantErrValue: ErrValDIV0},
		{name: "negative_serial", formula: "DAY(-1)", wantErr: true, wantErrValue: ErrValNUM},
		{name: "beyond_max_serial", formula: "DAY(2958466)", wantErr: true, wantErrValue: ErrValNUM},
		{name: "text_error", formula: `DAY("not a date")`, wantErr: true, wantErrValue: ErrValVALUE},
		{name: "wrong_arg_count_zero", formula: "DAY()", wantErr: true, wantErrValue: ErrValVALUE},
		{name: "wrong_arg_count_two", formula: "DAY(1,2)", wantErr: true, wantErrValue: ErrValVALUE},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cf := evalCompile(t, tt.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): %v", tt.formula, err)
			}
			if tt.wantErr {
				if got.Type != ValueError || got.Err != tt.wantErrValue {
					t.Fatalf("%s = %v, want error %v", tt.formula, got, tt.wantErrValue)
				}
				return
			}
			if got.Type != ValueNumber || got.Num != tt.want {
				t.Fatalf("%s = %v, want %g", tt.formula, got, tt.want)
			}
		})
	}
}

func TestYEARMONTHDAY_Serial0(t *testing.T) {
	resolver := &mockResolver{}

	// Serial 0 is "January 0, 1900" — a special sentinel.
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

func TestNOW(t *testing.T) {
	resolver := &mockResolver{}

	// Helper: evaluate a formula and return the result value.
	eval := func(t *testing.T, formula string) Value {
		t.Helper()
		cf := evalCompile(t, formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", formula, err)
		}
		return got
	}

	// --- Basic return type ---

	t.Run("returns_number_type", func(t *testing.T) {
		got := eval(t, "NOW()")
		if got.Type != ValueNumber {
			t.Errorf("NOW() type = %v, want ValueNumber", got.Type)
		}
	})

	t.Run("no_args_works", func(t *testing.T) {
		got := eval(t, "NOW()")
		if got.Type == ValueError {
			t.Errorf("NOW() returned error: %v", got)
		}
	})

	// --- Wrong argument count ---

	t.Run("one_arg_error", func(t *testing.T) {
		got := eval(t, "NOW(1)")
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("NOW(1) = %v, want #VALUE! error", got)
		}
	})

	t.Run("string_arg_error", func(t *testing.T) {
		got := eval(t, `NOW("x")`)
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`NOW("x") = %v, want #VALUE! error`, got)
		}
	})

	// --- Value range checks ---

	t.Run("positive_value", func(t *testing.T) {
		got := eval(t, "NOW()")
		if got.Num <= 0 {
			t.Errorf("NOW() = %g, expected > 0", got.Num)
		}
	})

	t.Run("reasonable_serial_lower_bound", func(t *testing.T) {
		// Serial 44000 is approximately mid-2020
		got := eval(t, "NOW()")
		if got.Num < 44000 {
			t.Errorf("NOW() = %g, expected > 44000 (~ year 2020)", got.Num)
		}
	})

	t.Run("reasonable_serial_upper_bound", func(t *testing.T) {
		// Serial 55000 is approximately year 2050
		got := eval(t, "NOW()")
		if got.Num > 55000 {
			t.Errorf("NOW() = %g, expected < 55000 (~ year 2050)", got.Num)
		}
	})

	// --- Fractional part (NOW includes time, unlike TODAY) ---

	t.Run("fractional_part_valid_range", func(t *testing.T) {
		// The fractional part of NOW() should be >= 0 and < 1.
		// This is always true regardless of time of day.
		got := eval(t, "NOW()")
		frac := got.Num - math.Floor(got.Num)
		if frac < 0 || frac >= 1 {
			t.Errorf("NOW() fractional part = %g, expected 0 <= frac < 1", frac)
		}
	})

	// --- Relationship between NOW and TODAY ---

	t.Run("int_now_equals_today", func(t *testing.T) {
		intNow := eval(t, "INT(NOW())")
		today := eval(t, "TODAY()")
		if intNow.Num != today.Num {
			t.Errorf("INT(NOW()) = %g, TODAY() = %g, expected equal", intNow.Num, today.Num)
		}
	})

	// --- Arithmetic ---

	t.Run("now_plus_1_is_tomorrow_same_time", func(t *testing.T) {
		now := eval(t, "NOW()")
		nowPlus1 := eval(t, "NOW()+1")
		// The difference should be very close to 1 (may differ slightly due to evaluation time)
		diff := nowPlus1.Num - now.Num
		if diff < 0.999 || diff > 1.001 {
			t.Errorf("NOW()+1 - NOW() = %g, expected ~1.0", diff)
		}
	})

	t.Run("now_minus_now_approx_zero", func(t *testing.T) {
		got := eval(t, "NOW()-NOW()")
		// Both NOW() calls evaluate very close in time; allow small delta
		if got.Num < -0.001 || got.Num > 0.001 {
			t.Errorf("NOW()-NOW() = %g, expected ~0", got.Num)
		}
	})

	// --- Date component extraction ---

	t.Run("year_now_reasonable", func(t *testing.T) {
		got := eval(t, "YEAR(NOW())")
		if got.Type != ValueNumber {
			t.Fatalf("YEAR(NOW()) type = %v, want number", got.Type)
		}
		if got.Num < 2024 || got.Num > 2030 {
			t.Errorf("YEAR(NOW()) = %g, expected between 2024 and 2030", got.Num)
		}
	})

	t.Run("month_now_in_range", func(t *testing.T) {
		got := eval(t, "MONTH(NOW())")
		if got.Type != ValueNumber {
			t.Fatalf("MONTH(NOW()) type = %v, want number", got.Type)
		}
		if got.Num < 1 || got.Num > 12 {
			t.Errorf("MONTH(NOW()) = %g, expected 1-12", got.Num)
		}
	})

	t.Run("day_now_in_range", func(t *testing.T) {
		got := eval(t, "DAY(NOW())")
		if got.Type != ValueNumber {
			t.Fatalf("DAY(NOW()) type = %v, want number", got.Type)
		}
		if got.Num < 1 || got.Num > 31 {
			t.Errorf("DAY(NOW()) = %g, expected 1-31", got.Num)
		}
	})

	t.Run("hour_now_in_range", func(t *testing.T) {
		got := eval(t, "HOUR(NOW())")
		if got.Type != ValueNumber {
			t.Fatalf("HOUR(NOW()) type = %v, want number", got.Type)
		}
		if got.Num < 0 || got.Num > 23 {
			t.Errorf("HOUR(NOW()) = %g, expected 0-23", got.Num)
		}
	})

	t.Run("minute_now_in_range", func(t *testing.T) {
		got := eval(t, "MINUTE(NOW())")
		if got.Type != ValueNumber {
			t.Fatalf("MINUTE(NOW()) type = %v, want number", got.Type)
		}
		if got.Num < 0 || got.Num > 59 {
			t.Errorf("MINUTE(NOW()) = %g, expected 0-59", got.Num)
		}
	})

	t.Run("second_now_in_range", func(t *testing.T) {
		got := eval(t, "SECOND(NOW())")
		if got.Type != ValueNumber {
			t.Fatalf("SECOND(NOW()) type = %v, want number", got.Type)
		}
		if got.Num < 0 || got.Num > 59 {
			t.Errorf("SECOND(NOW()) = %g, expected 0-59", got.Num)
		}
	})

	// --- Type checking functions ---

	t.Run("type_now_is_1", func(t *testing.T) {
		got := eval(t, "TYPE(NOW())")
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("TYPE(NOW()) = %v, want 1 (number)", got)
		}
	})

	t.Run("isnumber_now_is_true", func(t *testing.T) {
		got := eval(t, "ISNUMBER(NOW())")
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(NOW()) = %v, want TRUE", got)
		}
	})

	t.Run("istext_now_is_false", func(t *testing.T) {
		got := eval(t, "ISTEXT(NOW())")
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISTEXT(NOW()) = %v, want FALSE", got)
		}
	})

	// --- Truthiness ---

	t.Run("now_in_if_condition_truthy", func(t *testing.T) {
		got := eval(t, `IF(NOW(),"yes","no")`)
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IF(NOW(),"yes","no") = %v, want "yes"`, got)
		}
	})

	// --- Boundary comparisons ---

	t.Run("now_gt_today_minus_1", func(t *testing.T) {
		got := eval(t, "NOW()>TODAY()-1")
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("NOW()>TODAY()-1 = %v, want TRUE", got)
		}
	})

	t.Run("now_lt_today_plus_2", func(t *testing.T) {
		got := eval(t, "NOW()<TODAY()+2")
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("NOW()<TODAY()+2 = %v, want TRUE", got)
		}
	})

	// --- WEEKDAY ---

	t.Run("weekday_now_in_range", func(t *testing.T) {
		got := eval(t, "WEEKDAY(NOW())")
		if got.Type != ValueNumber {
			t.Fatalf("WEEKDAY(NOW()) type = %v, want number", got.Type)
		}
		if got.Num < 1 || got.Num > 7 {
			t.Errorf("WEEKDAY(NOW()) = %g, expected 1-7", got.Num)
		}
	})

	// --- Cross-check with Go time ---

	t.Run("now_matches_go_time_now_approx", func(t *testing.T) {
		// Verify NOW() is close to the Go-computed serial for time.Now().
		goSerial := TimeToSerial(time.Now())
		got := eval(t, "NOW()")
		diff := got.Num - goSerial
		if diff < 0 {
			diff = -diff
		}
		// Allow up to ~1 second of difference (1/86400 of a day)
		if diff > 1.0/86400.0*2 {
			t.Errorf("NOW() = %g, Go TimeToSerial = %g, differ by %g (too large)", got.Num, goSerial, diff)
		}
	})

	// --- Excel doc examples ---

	t.Run("now_minus_half_is_12_hours_ago", func(t *testing.T) {
		now := eval(t, "NOW()")
		halfAgo := eval(t, "NOW()-0.5")
		diff := now.Num - halfAgo.Num
		if diff < 0.499 || diff > 0.501 {
			t.Errorf("NOW() - (NOW()-0.5) = %g, expected ~0.5", diff)
		}
	})

	t.Run("now_plus_7_is_one_week_ahead", func(t *testing.T) {
		now := eval(t, "NOW()")
		weekAhead := eval(t, "NOW()+7")
		diff := weekAhead.Num - now.Num
		if diff < 6.999 || diff > 7.001 {
			t.Errorf("NOW()+7 - NOW() = %g, expected ~7", diff)
		}
	})
}

func TestTODAY(t *testing.T) {
	resolver := &mockResolver{}

	// Helper: evaluate a formula and return the result value.
	eval := func(t *testing.T, formula string) Value {
		t.Helper()
		cf := evalCompile(t, formula)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("Eval(%s): %v", formula, err)
		}
		return got
	}

	// --- Structural and type tests (table-driven) ---

	t.Run("returns_number_type", func(t *testing.T) {
		got := eval(t, "TODAY()")
		if got.Type != ValueNumber {
			t.Errorf("TODAY() type = %v, want ValueNumber", got.Type)
		}
	})

	t.Run("returns_integer_no_fractional_part", func(t *testing.T) {
		got := eval(t, "TODAY()")
		if got.Num != math.Floor(got.Num) {
			t.Errorf("TODAY() = %g, expected integer (no fractional time)", got.Num)
		}
	})

	t.Run("int_today_equals_today", func(t *testing.T) {
		today := eval(t, "TODAY()")
		intToday := eval(t, "INT(TODAY())")
		if today.Num != intToday.Num {
			t.Errorf("TODAY() = %g, INT(TODAY()) = %g, expected equal", today.Num, intToday.Num)
		}
	})

	t.Run("no_args_works", func(t *testing.T) {
		got := eval(t, "TODAY()")
		if got.Type != ValueNumber {
			t.Errorf("TODAY() should work with no arguments, got type %v", got.Type)
		}
	})

	t.Run("one_arg_error", func(t *testing.T) {
		got := eval(t, "TODAY(1)")
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf("TODAY(1) = %v, want #VALUE! error", got)
		}
	})

	t.Run("string_arg_error", func(t *testing.T) {
		got := eval(t, `TODAY("x")`)
		if got.Type != ValueError || got.Err != ErrValVALUE {
			t.Errorf(`TODAY("x") = %v, want #VALUE! error`, got)
		}
	})

	t.Run("positive_value", func(t *testing.T) {
		got := eval(t, "TODAY()")
		if got.Num <= 0 {
			t.Errorf("TODAY() = %g, expected > 0", got.Num)
		}
	})

	t.Run("reasonable_serial_lower_bound", func(t *testing.T) {
		// Serial 44000 is approximately mid-2020
		got := eval(t, "TODAY()")
		if got.Num < 44000 {
			t.Errorf("TODAY() = %g, expected > 44000 (~ year 2020)", got.Num)
		}
	})

	t.Run("reasonable_serial_upper_bound", func(t *testing.T) {
		// Serial 55000 is approximately year 2050
		got := eval(t, "TODAY()")
		if got.Num > 55000 {
			t.Errorf("TODAY() = %g, expected < 55000 (~ year 2050)", got.Num)
		}
	})

	t.Run("today_equals_int_now", func(t *testing.T) {
		today := eval(t, "TODAY()")
		intNow := eval(t, "INT(NOW())")
		if today.Num != intNow.Num {
			t.Errorf("TODAY() = %g, INT(NOW()) = %g, expected equal", today.Num, intNow.Num)
		}
	})

	t.Run("tomorrow_is_today_plus_1", func(t *testing.T) {
		today := eval(t, "TODAY()")
		tomorrow := eval(t, "TODAY()+1")
		if tomorrow.Num != today.Num+1 {
			t.Errorf("TODAY()+1 = %g, TODAY() = %g, expected difference of 1", tomorrow.Num, today.Num)
		}
	})

	t.Run("yesterday_is_today_minus_1", func(t *testing.T) {
		today := eval(t, "TODAY()")
		yesterday := eval(t, "TODAY()-1")
		if yesterday.Num != today.Num-1 {
			t.Errorf("TODAY()-1 = %g, TODAY() = %g, expected difference of -1", yesterday.Num, today.Num)
		}
	})

	t.Run("today_minus_today_is_zero", func(t *testing.T) {
		got := eval(t, "TODAY()-TODAY()")
		if got.Num != 0 {
			t.Errorf("TODAY()-TODAY() = %g, want 0", got.Num)
		}
	})

	t.Run("year_today_reasonable", func(t *testing.T) {
		got := eval(t, "YEAR(TODAY())")
		if got.Type != ValueNumber {
			t.Fatalf("YEAR(TODAY()) type = %v, want number", got.Type)
		}
		if got.Num < 2024 || got.Num > 2030 {
			t.Errorf("YEAR(TODAY()) = %g, expected between 2024 and 2030", got.Num)
		}
	})

	t.Run("month_today_in_range", func(t *testing.T) {
		got := eval(t, "MONTH(TODAY())")
		if got.Type != ValueNumber {
			t.Fatalf("MONTH(TODAY()) type = %v, want number", got.Type)
		}
		if got.Num < 1 || got.Num > 12 {
			t.Errorf("MONTH(TODAY()) = %g, expected 1-12", got.Num)
		}
	})

	t.Run("day_today_in_range", func(t *testing.T) {
		got := eval(t, "DAY(TODAY())")
		if got.Type != ValueNumber {
			t.Fatalf("DAY(TODAY()) type = %v, want number", got.Type)
		}
		if got.Num < 1 || got.Num > 31 {
			t.Errorf("DAY(TODAY()) = %g, expected 1-31", got.Num)
		}
	})

	t.Run("today_gte_date_2024_1_1", func(t *testing.T) {
		got := eval(t, "TODAY()>=DATE(2024,1,1)")
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("TODAY()>=DATE(2024,1,1) = %v, expected TRUE", got)
		}
	})

	t.Run("type_today_is_1", func(t *testing.T) {
		got := eval(t, "TYPE(TODAY())")
		if got.Type != ValueNumber || got.Num != 1 {
			t.Errorf("TYPE(TODAY()) = %v, want 1 (number)", got)
		}
	})

	t.Run("isnumber_today_is_true", func(t *testing.T) {
		got := eval(t, "ISNUMBER(TODAY())")
		if got.Type != ValueBool || !got.Bool {
			t.Errorf("ISNUMBER(TODAY()) = %v, want TRUE", got)
		}
	})

	t.Run("istext_today_is_false", func(t *testing.T) {
		got := eval(t, "ISTEXT(TODAY())")
		if got.Type != ValueBool || got.Bool {
			t.Errorf("ISTEXT(TODAY()) = %v, want FALSE", got)
		}
	})

	t.Run("today_in_if_condition_truthy", func(t *testing.T) {
		got := eval(t, `IF(TODAY(),"yes","no")`)
		if got.Type != ValueString || got.Str != "yes" {
			t.Errorf(`IF(TODAY(),"yes","no") = %v, want "yes"`, got)
		}
	})

	t.Run("text_today_yyyy_is_4_chars", func(t *testing.T) {
		got := eval(t, `TEXT(TODAY(),"yyyy")`)
		if got.Type != ValueString {
			t.Fatalf(`TEXT(TODAY(),"yyyy") type = %v, want string`, got.Type)
		}
		if len(got.Str) != 4 {
			t.Errorf(`TEXT(TODAY(),"yyyy") = %q, expected a 4-character year string`, got.Str)
		}
	})

	t.Run("today_plus_half_has_fraction", func(t *testing.T) {
		got := eval(t, "TODAY()+0.5")
		frac := got.Num - math.Floor(got.Num)
		if math.Abs(frac-0.5) > 1e-9 {
			t.Errorf("TODAY()+0.5 fractional part = %g, expected 0.5", frac)
		}
	})

	t.Run("weekday_today_in_range", func(t *testing.T) {
		got := eval(t, "WEEKDAY(TODAY())")
		if got.Type != ValueNumber {
			t.Fatalf("WEEKDAY(TODAY()) type = %v, want number", got.Type)
		}
		if got.Num < 1 || got.Num > 7 {
			t.Errorf("WEEKDAY(TODAY()) = %g, expected 1-7", got.Num)
		}
	})

	t.Run("today_matches_go_time_now", func(t *testing.T) {
		// Verify TODAY() matches the Go-computed serial for today's date.
		now := time.Now()
		today := time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, time.UTC)
		expected := math.Floor(TimeToSerial(today))
		got := eval(t, "TODAY()")
		if got.Num != expected {
			t.Errorf("TODAY() = %g, expected %g (from Go time.Now())", got.Num, expected)
		}
	})

	t.Run("today_times_1_equals_today", func(t *testing.T) {
		today := eval(t, "TODAY()")
		product := eval(t, "TODAY()*1")
		if product.Num != today.Num {
			t.Errorf("TODAY()*1 = %g, TODAY() = %g, expected equal", product.Num, today.Num)
		}
	})

	t.Run("today_plus_7_minus_today_equals_7", func(t *testing.T) {
		got := eval(t, "(TODAY()+7)-TODAY()")
		if got.Num != 7 {
			t.Errorf("(TODAY()+7)-TODAY() = %g, want 7", got.Num)
		}
	})
}

func TestDAYS360(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		// ── Existing tests (basic US method) ───────────────────────────
		// Basic US method: Jan 1 to Feb 1 = 30 days
		{"us_jan1_feb1", "DAYS360(DATE(2025,1,1),DATE(2025,2,1),FALSE)", 30, false, 0},
		// US method: Feb 28 (last day of Feb, non-leap) to Mar 31
		// Feb 28 → D1=30 (last-of-Feb rule), Mar 31 → D2=30 (D2==31 && D1>=30)
		{"us_feb28_mar31", "DAYS360(45716,45747,FALSE)", 30, false, 0},
		// US method: Jan 31 to Mar 31
		// Jan 31 → D1=30, Mar 31 → D2=30 (D2==31 && D1>=30)
		{"us_jan31_mar31", "DAYS360(DATE(2025,1,31),DATE(2025,3,31),FALSE)", 60, false, 0},
		// US method: both dates last day of Feb (leap year 2024)
		// Feb 29 2024 → D1=30, Feb 29 2024 → D2=30 (both last-of-Feb)
		{"us_both_feb_leap", "DAYS360(DATE(2024,2,29),DATE(2024,2,29),FALSE)", 0, false, 0},
		// US method: Feb 29 (leap) to Mar 31
		// Feb 29 → D1=30, Mar 31 → D2=30
		{"us_feb29_mar31_leap", "DAYS360(DATE(2024,2,29),DATE(2024,3,31),FALSE)", 30, false, 0},
		// European method: same dates
		// European: D1=28, D2=31→30. (3-2)*30 + (30-28) = 32
		{"eu_feb28_mar31", "DAYS360(45716,45747,TRUE)", 32, false, 0},
		// US method: regular dates, no adjustments needed
		{"us_jan15_mar15", "DAYS360(DATE(2025,1,15),DATE(2025,3,15),FALSE)", 60, false, 0},

		// ── Same date → 0 ──────────────────────────────────────────────
		{"same_date_us", "DAYS360(DATE(2025,6,15),DATE(2025,6,15))", 0, false, 0},
		{"same_date_eu", "DAYS360(DATE(2025,6,15),DATE(2025,6,15),TRUE)", 0, false, 0},

		// ── Same month dates ───────────────────────────────────────────
		{"us_same_month", "DAYS360(DATE(2025,3,5),DATE(2025,3,20))", 15, false, 0},
		{"eu_same_month", "DAYS360(DATE(2025,3,5),DATE(2025,3,20),TRUE)", 15, false, 0},

		// ── Cross-month dates ──────────────────────────────────────────
		{"us_cross_month", "DAYS360(DATE(2025,1,20),DATE(2025,2,15))", 25, false, 0},

		// ── Cross-year dates ───────────────────────────────────────────
		{"us_cross_year", "DAYS360(DATE(2024,12,1),DATE(2025,1,15))", 44, false, 0},
		{"eu_cross_year", "DAYS360(DATE(2024,12,1),DATE(2025,1,15),TRUE)", 44, false, 0},

		// ── Start after end → negative result ──────────────────────────
		{"us_negative_result", "DAYS360(DATE(2025,3,15),DATE(2025,1,15),FALSE)", -60, false, 0},
		{"eu_negative_result", "DAYS360(DATE(2025,3,15),DATE(2025,1,15),TRUE)", -60, false, 0},

		// ── US method default (omitted method parameter) ───────────────
		{"us_default_no_method", "DAYS360(DATE(2025,1,1),DATE(2025,2,1))", 30, false, 0},

		// ── Full year: Jan 1 to Dec 31 ─────────────────────────────────
		// US: Jan 1 to Dec 31. D1=1, D2=31→stays 31 (D1<30). (12-1)*30+(31-1) = 360
		{"us_full_year", "DAYS360(DATE(2025,1,1),DATE(2025,12,31),FALSE)", 360, false, 0},
		// EU: Jan 1 to Dec 31. D1=1, D2=31→30. (12-1)*30+(30-1) = 359
		{"eu_full_year", "DAYS360(DATE(2025,1,1),DATE(2025,12,31),TRUE)", 359, false, 0},

		// ── Exactly 360 days: Jan 1 to Dec 30 ─────────────────────────
		{"us_exactly_360", "DAYS360(DATE(2025,1,1),DATE(2025,12,30),FALSE)", 359, false, 0},
		// Jan 30 to Jan 30 next year = 360
		{"us_jan30_to_jan30", "DAYS360(DATE(2025,1,30),DATE(2026,1,30),FALSE)", 360, false, 0},

		// ── US vs European method differences ──────────────────────────
		// Jan 31 to Apr 30: US: D1=31→30, D2=30. (4-1)*30+(30-30)=90
		// EU: D1=31→30, D2=30. Same result = 90
		{"us_jan31_apr30", "DAYS360(DATE(2025,1,31),DATE(2025,4,30),FALSE)", 90, false, 0},
		{"eu_jan31_apr30", "DAYS360(DATE(2025,1,31),DATE(2025,4,30),TRUE)", 90, false, 0},
		// Jan 30 to Mar 31: US: D1=30, D2=31→30 (D1>=30). (3-1)*30+(30-30)=60
		// EU: D1=30, D2=31→30. (3-1)*30+(30-30)=60
		{"us_jan30_mar31", "DAYS360(DATE(2025,1,30),DATE(2025,3,31),FALSE)", 60, false, 0},
		{"eu_jan30_mar31", "DAYS360(DATE(2025,1,30),DATE(2025,3,31),TRUE)", 60, false, 0},
		// Jan 15 to Mar 31: US: D1=15, D2=31 stays (D1<30). (3-1)*30+(31-15)=76
		// EU: D1=15, D2=31→30. (3-1)*30+(30-15)=75
		{"us_jan15_mar31", "DAYS360(DATE(2025,1,15),DATE(2025,3,31),FALSE)", 76, false, 0},
		{"eu_jan15_mar31", "DAYS360(DATE(2025,1,15),DATE(2025,3,31),TRUE)", 75, false, 0},

		// ── End of month handling ──────────────────────────────────────
		// Feb 28 non-leap to Feb 28: same date = 0
		{"us_feb28_to_feb28", "DAYS360(DATE(2025,2,28),DATE(2025,2,28),FALSE)", 0, false, 0},
		// Feb 28 in non-leap year to Mar 31
		// US: Feb 28 is last of Feb → D1=30. Mar 31 → D2=30. (3-2)*30+(30-30)=30
		{"us_feb28_mar31_nonleap", "DAYS360(DATE(2025,2,28),DATE(2025,3,31),FALSE)", 30, false, 0},
		// EU: Feb 28 → D1=28 (no adjustment). Mar 31 → D2=30. (3-2)*30+(30-28)=32
		{"eu_feb28_mar31_nonleap", "DAYS360(DATE(2025,2,28),DATE(2025,3,31),TRUE)", 32, false, 0},

		// ── Leap year: Feb 29 ──────────────────────────────────────────
		// Feb 29 to Mar 1 in leap year
		// US: Feb 29 is last of Feb → D1=30. Mar 1 → D2=1. (3-2)*30+(1-30)=1
		{"us_feb29_mar1_leap", "DAYS360(DATE(2024,2,29),DATE(2024,3,1),FALSE)", 1, false, 0},
		// EU: D1=29, D2=1. (3-2)*30+(1-29)=2
		{"eu_feb29_mar1_leap", "DAYS360(DATE(2024,2,29),DATE(2024,3,1),TRUE)", 2, false, 0},
		// Both Feb 29 in same leap year, European
		{"eu_both_feb29_leap", "DAYS360(DATE(2024,2,29),DATE(2024,2,29),TRUE)", 0, false, 0},

		// ── Jan 30 to Feb 28 (non-leap) ────────────────────────────────
		// US: D1=30, Feb 28 is last of Feb so both-last-of-Feb check fails (Jan 30 not last of Feb).
		//   D1=30, Feb 28 not last-of-Feb-D1 case. D2=28, D1=30. D2<31 so no adj.
		//   Actually: isLastDayOfFeb(sy=2025,sm=1,sd=30) is false. So no US Feb rules apply.
		//   D2=28, D1=30. (2-1)*30+(28-30)=28
		{"us_jan30_feb28", "DAYS360(DATE(2025,1,30),DATE(2025,2,28),FALSE)", 28, false, 0},
		// EU: same, no 31-day adjustments. (2-1)*30+(28-30)=28
		{"eu_jan30_feb28", "DAYS360(DATE(2025,1,30),DATE(2025,2,28),TRUE)", 28, false, 0},

		// ── Month-end to month-end ─────────────────────────────────────
		// Mar 31 to May 31: US: D1=31→30, D2=31→30 (D1>=30). (5-3)*30+(30-30)=60
		{"us_mar31_may31", "DAYS360(DATE(2025,3,31),DATE(2025,5,31),FALSE)", 60, false, 0},
		// EU: D1=31→30, D2=31→30. Same result = 60
		{"eu_mar31_may31", "DAYS360(DATE(2025,3,31),DATE(2025,5,31),TRUE)", 60, false, 0},
		// Apr 30 to Jun 30: no 31-adjustments needed. (6-4)*30+(30-30)=60
		{"us_apr30_jun30", "DAYS360(DATE(2025,4,30),DATE(2025,6,30),FALSE)", 60, false, 0},

		// ── Large date spans (multiple years) ──────────────────────────
		// 2020-Jan-1 to 2025-Jan-1 = 5*360 = 1800
		{"us_five_years", "DAYS360(DATE(2020,1,1),DATE(2025,1,1),FALSE)", 1800, false, 0},
		{"eu_five_years", "DAYS360(DATE(2020,1,1),DATE(2025,1,1),TRUE)", 1800, false, 0},
		// 10-year span
		{"us_ten_years", "DAYS360(DATE(2015,6,15),DATE(2025,6,15),FALSE)", 3600, false, 0},

		// ── Serial number inputs ───────────────────────────────────────
		// Serial 1 = Jan 1, 1900; Serial 31 = Jan 31, 1900. Difference = 30
		{"serial_numbers", "DAYS360(1,31,FALSE)", 30, false, 0},

		// ── Boolean coercion for method parameter ──────────────────────
		// TRUE = European, FALSE = US
		{"method_true_bool", "DAYS360(DATE(2025,1,15),DATE(2025,3,31),TRUE)", 75, false, 0},
		{"method_false_bool", "DAYS360(DATE(2025,1,15),DATE(2025,3,31),FALSE)", 76, false, 0},
		// Numeric: 0 = US (FALSE), 1 = European (TRUE)
		{"method_zero_is_us", "DAYS360(DATE(2025,1,15),DATE(2025,3,31),0)", 76, false, 0},
		{"method_one_is_eu", "DAYS360(DATE(2025,1,15),DATE(2025,3,31),1)", 75, false, 0},
		// Any non-zero number is European
		{"method_nonzero_is_eu", "DAYS360(DATE(2025,1,15),DATE(2025,3,31),5)", 75, false, 0},

		// ── US both last-day-of-Feb across different years ─────────────
		// Feb 28 2025 (non-leap, last of Feb) to Feb 29 2024 (leap, last of Feb)
		// start > end so negative. Both are last-of-Feb but D2 day (29) > D1 day (28),
		// so D2 is NOT adjusted to 30. D1=30, D2=29.
		// (2024-2025)*360+(2-2)*30+(29-30) = -361
		{"us_feb28_to_feb29_diff_years", "DAYS360(DATE(2025,2,28),DATE(2024,2,29),FALSE)", -361, false, 0},

		// ── Error handling ─────────────────────────────────────────────
		{"no_args", "DAYS360()", 0, true, ErrValVALUE},
		{"one_arg", "DAYS360(1)", 0, true, ErrValVALUE},
		{"too_many_args", "DAYS360(1,2,3,4)", 0, true, ErrValVALUE},
		{"non_numeric_start", `DAYS360("abc",1)`, 0, true, ErrValVALUE},
		{"non_numeric_end", `DAYS360(1,"abc")`, 0, true, ErrValVALUE},
		{"non_numeric_method", `DAYS360(1,2,"abc")`, 0, true, ErrValVALUE},

		// ── Error propagation ──────────────────────────────────────────
		{"error_prop_start", "DAYS360(1/0,1)", 0, true, ErrValDIV0},
		{"error_prop_end", "DAYS360(1,1/0)", 0, true, ErrValDIV0},
		{"error_prop_method", "DAYS360(1,2,1/0)", 0, true, ErrValDIV0},

		// ── Additional comprehensive tests ───────────────────────────

		// US: D2=31, D1<30 → D2 stays 31
		// Jan 15 to May 31: (5-1)*30+(31-15) = 136
		{"us_d2_31_d1_lt_30", "DAYS360(DATE(2025,1,15),DATE(2025,5,31),FALSE)", 136, false, 0},
		// EU: Jan 15 to May 31: D2=31→30. (5-1)*30+(30-15) = 135
		{"eu_d2_31_d1_lt_30", "DAYS360(DATE(2025,1,15),DATE(2025,5,31),TRUE)", 135, false, 0},

		// US: D1=31, D2=non-31
		// Mar 31 to Apr 15: D1=31→30. (4-3)*30+(15-30) = 15
		{"us_d1_31_d2_mid", "DAYS360(DATE(2025,3,31),DATE(2025,4,15),FALSE)", 15, false, 0},
		// EU: same. D1=31→30. (4-3)*30+(15-30) = 15
		{"eu_d1_31_d2_mid", "DAYS360(DATE(2025,3,31),DATE(2025,4,15),TRUE)", 15, false, 0},

		// Both Feb 28 in non-leap year (US: both last-of-Feb → D1=30,D2=30 → 0)
		{"us_both_feb28_nonleap", "DAYS360(DATE(2025,2,28),DATE(2025,2,28),FALSE)", 0, false, 0},

		// Feb 27 to Feb 28 non-leap year
		// US: Feb 27 not last-of-Feb, Feb 28 is last-of-Feb (but only D2 adj applies if both are last-of-Feb)
		// isLastDayOfFeb(2025,2,27)=false, so no US Feb rules. D1=27, D2=28. 28-27=1
		{"us_feb27_to_feb28_nonleap", "DAYS360(DATE(2025,2,27),DATE(2025,2,28),FALSE)", 1, false, 0},
		{"eu_feb27_to_feb28_nonleap", "DAYS360(DATE(2025,2,27),DATE(2025,2,28),TRUE)", 1, false, 0},

		// Feb 28 non-leap to Mar 1
		// US: Feb 28 is last-of-Feb → D1=30. Mar 1 → D2=1. (3-2)*30+(1-30)=1
		{"us_feb28_to_mar1_nonleap", "DAYS360(DATE(2025,2,28),DATE(2025,3,1),FALSE)", 1, false, 0},
		// EU: D1=28, D2=1. (3-2)*30+(1-28)=3
		{"eu_feb28_to_mar1_nonleap", "DAYS360(DATE(2025,2,28),DATE(2025,3,1),TRUE)", 3, false, 0},

		// Leap year: Feb 28 to Feb 29
		// US: Feb 28 not last-of-Feb in leap year. D1=28, D2=29. 29-28=1
		{"us_feb28_to_feb29_leap", "DAYS360(DATE(2024,2,28),DATE(2024,2,29),FALSE)", 1, false, 0},
		{"eu_feb28_to_feb29_leap", "DAYS360(DATE(2024,2,28),DATE(2024,2,29),TRUE)", 1, false, 0},

		// 30-day month vs 31-day month
		// Apr 30 to May 31: US: D1=30, D2=31→30 (D1>=30). (5-4)*30+(30-30)=30
		{"us_apr30_may31", "DAYS360(DATE(2025,4,30),DATE(2025,5,31),FALSE)", 30, false, 0},
		// EU: D1=30, D2=31→30. Same = 30
		{"eu_apr30_may31", "DAYS360(DATE(2025,4,30),DATE(2025,5,31),TRUE)", 30, false, 0},

		// Negative span with 31-day adjustments
		// US: Mar 31 to Jan 15 (reversed): D1=31→30, D2=15. (1-3)*30+(15-30)=-75
		{"us_negative_31_adj", "DAYS360(DATE(2025,3,31),DATE(2025,1,15),FALSE)", -75, false, 0},

		// Cross multiple years
		// Jan 1, 2020 to Jan 1, 2030 = 10*360 = 3600
		{"us_ten_years_jan_jan", "DAYS360(DATE(2020,1,1),DATE(2030,1,1),FALSE)", 3600, false, 0},
		{"eu_ten_years_jan_jan", "DAYS360(DATE(2020,1,1),DATE(2030,1,1),TRUE)", 3600, false, 0},

		// One day apart, mid-month
		{"one_day_mid_month", "DAYS360(DATE(2025,6,14),DATE(2025,6,15))", 1, false, 0},

		// US: end=31, start=29 (not 30 or 31) → D2 stays 31
		// Jan 29 to Mar 31: (3-1)*30+(31-29) = 62
		{"us_start_29_end_31", "DAYS360(DATE(2025,1,29),DATE(2025,3,31),FALSE)", 62, false, 0},
		// EU: D2=31→30. (3-1)*30+(30-29) = 61
		{"eu_start_29_end_31", "DAYS360(DATE(2025,1,29),DATE(2025,3,31),TRUE)", 61, false, 0},

		// Full year from mid-month (Jun 15 to Jun 15)
		{"us_full_year_mid", "DAYS360(DATE(2024,6,15),DATE(2025,6,15),FALSE)", 360, false, 0},
		{"eu_full_year_mid", "DAYS360(DATE(2024,6,15),DATE(2025,6,15),TRUE)", 360, false, 0},

		// Feb 29 leap to Feb 28 next year (non-leap)
		// US: Feb 29 is last-of-Feb (D1=30), Feb 28 next year is last-of-Feb (D2=30).
		// Both last-of-Feb rule: D2=30. Then D1 last-of-Feb: D1=30.
		// (2025-2024)*360+(2-2)*30+(30-30)=360
		{"us_feb29_to_feb28_next", "DAYS360(DATE(2024,2,29),DATE(2025,2,28),FALSE)", 360, false, 0},
		// EU: D1=29, D2=28. (2025-2024)*360+(2-2)*30+(28-29)=359
		{"eu_feb29_to_feb28_next", "DAYS360(DATE(2024,2,29),DATE(2025,2,28),TRUE)", 359, false, 0},
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
		{"doc_1900_leap_bug_gap", "DAYS(DATE(1900,3,1),DATE(1900,2,28))", 2, false, 0},
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
		{"doc_plus_one", "EDATE(DATE(2011,1,15),1)", dateSerial(2011, time.February, 15), false, 0},
		{"doc_minus_one", "EDATE(DATE(2011,1,15),-1)", dateSerial(2010, time.December, 15), false, 0},
		{"doc_plus_two", "EDATE(DATE(2011,1,15),2)", dateSerial(2011, time.March, 15), false, 0},
		{"month_end_non_leap", "EDATE(DATE(2025,1,31),1)", dateSerial(2025, time.February, 28), false, 0},
		{"month_end_leap", "EDATE(DATE(2024,1,31),1)", dateSerial(2024, time.February, 29), false, 0},
		{"cross_year_forward", "EDATE(DATE(2024,11,30),3)", dateSerial(2025, time.February, 28), false, 0},
		{"cross_year_backward", "EDATE(DATE(2025,1,31),-2)", dateSerial(2024, time.November, 30), false, 0},
		{"zero_months", "EDATE(DATE(2025,1,15),0)", dateSerial(2025, time.January, 15), false, 0},
		{"truncates_positive_fraction", "EDATE(DATE(2025,1,15),1.9)", dateSerial(2025, time.February, 15), false, 0},
		{"truncates_negative_fraction", "EDATE(DATE(2025,1,15),-1.9)", dateSerial(2024, time.December, 15), false, 0},
		{"ignores_time_component", "EDATE(DATE(2025,1,15)+TIME(12,0,0),1)", dateSerial(2025, time.February, 15), false, 0},
		{"leap_back_one_month", "EDATE(DATE(2024,3,31),-1)", dateSerial(2024, time.February, 29), false, 0},
		{"string_months", `EDATE(DATE(2025,1,15),"2")`, dateSerial(2025, time.March, 15), false, 0},
		{"too_few_args", "EDATE(DATE(2025,1,15))", 0, true, ErrValVALUE},
		{"too_many_args", "EDATE(DATE(2025,1,15),1,2)", 0, true, ErrValVALUE},
		{"invalid_start", `EDATE("abc",1)`, 0, true, ErrValVALUE},
		{"invalid_months", `EDATE(DATE(2025,1,15),"abc")`, 0, true, ErrValVALUE},
		{"error_propagation", "EDATE(1/0,1)", 0, true, ErrValDIV0},

		// --- Additional comprehensive EDATE tests ---

		// Basic: EDATE(Jan 15, 1) = Feb 15
		{"basic_jan_to_feb", "EDATE(DATE(2023,1,15),1)", dateSerial(2023, time.February, 15), false, 0},
		// Negative months: EDATE(Mar 15, -1) = Feb 15
		{"negative_mar_to_feb", "EDATE(DATE(2023,3,15),-1)", dateSerial(2023, time.February, 15), false, 0},
		// End of month: EDATE(Jan 31, 1) = Feb 28 (non-leap year)
		{"eom_jan31_to_feb28", "EDATE(DATE(2023,1,31),1)", dateSerial(2023, time.February, 28), false, 0},
		// End of month: EDATE(Jan 31, 1) = Feb 29 (leap year)
		{"eom_jan31_to_feb29_leap", "EDATE(DATE(2024,1,31),1)", dateSerial(2024, time.February, 29), false, 0},
		// EDATE(Mar 31, -1) = Feb 28 non-leap
		{"eom_mar31_back_to_feb28", "EDATE(DATE(2023,3,31),-1)", dateSerial(2023, time.February, 28), false, 0},
		// EDATE(Mar 31, -1) = Feb 29 leap
		{"eom_mar31_back_to_feb29_leap", "EDATE(DATE(2024,3,31),-1)", dateSerial(2024, time.February, 29), false, 0},
		// Leap year: EDATE(Feb 29 2020, 12) = Feb 28 2021
		{"leap_feb29_plus_12", "EDATE(DATE(2020,2,29),12)", dateSerial(2021, time.February, 28), false, 0},
		// Leap year: EDATE(Feb 29 2020, 48) = Feb 29 2024
		{"leap_feb29_plus_48", "EDATE(DATE(2020,2,29),48)", dateSerial(2024, time.February, 29), false, 0},
		// Zero months: EDATE(date, 0) = date
		{"zero_months_mid", "EDATE(DATE(2023,6,15),0)", dateSerial(2023, time.June, 15), false, 0},
		// Large month offsets (24 months = 2 years)
		{"large_offset_24", "EDATE(DATE(2020,3,15),24)", dateSerial(2022, time.March, 15), false, 0},
		// Large month offsets (60 months = 5 years)
		{"large_offset_60", "EDATE(DATE(2020,3,15),60)", dateSerial(2025, time.March, 15), false, 0},
		// Negative large offsets
		{"negative_large_24", "EDATE(DATE(2025,3,15),-24)", dateSerial(2023, time.March, 15), false, 0},
		{"negative_large_60", "EDATE(DATE(2025,3,15),-60)", dateSerial(2020, time.March, 15), false, 0},
		// Cross year boundary forward (Nov -> Feb)
		{"cross_year_nov_to_feb", "EDATE(DATE(2023,11,15),3)", dateSerial(2024, time.February, 15), false, 0},
		// December to January
		{"dec_to_jan", "EDATE(DATE(2023,12,15),1)", dateSerial(2024, time.January, 15), false, 0},
		// January to December (negative)
		{"jan_to_dec_neg", "EDATE(DATE(2024,1,15),-1)", dateSerial(2023, time.December, 15), false, 0},
		// EOM: Apr 30 + 1 = May 30 (not clamped since May has 31 days)
		{"apr30_plus1", "EDATE(DATE(2023,4,30),1)", dateSerial(2023, time.May, 30), false, 0},
		// EOM: May 31 + 1 = Jun 30
		{"may31_plus1", "EDATE(DATE(2023,5,31),1)", dateSerial(2023, time.June, 30), false, 0},
		// EOM chain: Jan 31 -> Feb 28 -> Mar 28 (not 31!)
		{"eom_chain_step1", "EDATE(DATE(2023,1,31),1)", dateSerial(2023, time.February, 28), false, 0},
		{"eom_chain_step2", "EDATE(EDATE(DATE(2023,1,31),1),1)", dateSerial(2023, time.March, 28), false, 0},
		// Error propagation for months arg
		{"error_propagation_months", "EDATE(DATE(2023,1,15),1/0)", 0, true, ErrValDIV0},
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
		{"doc_plus_one", "EOMONTH(DATE(2011,1,1),1)", dateSerial(2011, time.February, 28), false, 0},
		{"doc_minus_three", "EOMONTH(DATE(2011,1,1),-3)", dateSerial(2010, time.October, 31), false, 0},
		{"same_month", "EOMONTH(DATE(2025,1,15),0)", dateSerial(2025, time.January, 31), false, 0},
		{"leap_february", "EOMONTH(DATE(2024,1,15),1)", dateSerial(2024, time.February, 29), false, 0},
		{"month_end_self", "EOMONTH(DATE(2025,1,31),0)", dateSerial(2025, time.January, 31), false, 0},
		{"previous_month_end", "EOMONTH(DATE(2025,1,31),-1)", dateSerial(2024, time.December, 31), false, 0},
		{"cross_year_forward", "EOMONTH(DATE(2024,11,5),2)", dateSerial(2025, time.January, 31), false, 0},
		{"truncates_positive_fraction", "EOMONTH(DATE(2025,1,15),1.9)", dateSerial(2025, time.February, 28), false, 0},
		{"truncates_negative_fraction", "EOMONTH(DATE(2025,1,15),-1.9)", dateSerial(2024, time.December, 31), false, 0},
		{"ignores_time_component", "EOMONTH(DATE(2025,1,15)+TIME(18,30,0),1)", dateSerial(2025, time.February, 28), false, 0},
		{"leap_plus_twelve", "EOMONTH(DATE(2024,2,29),12)", dateSerial(2025, time.February, 28), false, 0},
		{"leap_minus_twelve", "EOMONTH(DATE(2024,2,29),-12)", dateSerial(2023, time.February, 28), false, 0},
		{"string_months", `EOMONTH(DATE(2025,1,15),"2")`, dateSerial(2025, time.March, 31), false, 0},
		{"too_few_args", "EOMONTH(DATE(2025,1,15))", 0, true, ErrValVALUE},
		{"too_many_args", "EOMONTH(DATE(2025,1,15),1,2)", 0, true, ErrValVALUE},
		{"invalid_months", `EOMONTH(DATE(2025,1,15),"abc")`, 0, true, ErrValVALUE},
		{"error_propagation", "EOMONTH(DATE(2025,1,15),1/0)", 0, true, ErrValDIV0},

		// --- Comprehensive additional EOMONTH tests ---

		// Basic: same month returns last day
		{"basic_jan15_zero", "EOMONTH(DATE(2025,1,15),0)", dateSerial(2025, time.January, 31), false, 0},

		// Forward by 1 month: Jan → Feb (non-leap year 2025)
		{"forward_1_jan_to_feb_nonleap", "EOMONTH(DATE(2025,1,15),1)", dateSerial(2025, time.February, 28), false, 0},

		// Forward by 1 month: Jan → Feb (leap year 2020)
		{"forward_1_jan_to_feb_leap_2020", "EOMONTH(DATE(2020,1,15),1)", dateSerial(2020, time.February, 29), false, 0},

		// Forward by 1 month: Jan → Feb (non-leap year 2021)
		{"forward_1_jan_to_feb_nonleap_2021", "EOMONTH(DATE(2021,1,15),1)", dateSerial(2021, time.February, 28), false, 0},

		// Backward: Mar 15 → Feb end (non-leap)
		{"backward_mar_to_feb_nonleap", "EOMONTH(DATE(2025,3,15),-1)", dateSerial(2025, time.February, 28), false, 0},

		// Backward: Mar 15 → Feb end (leap year 2020)
		{"backward_mar_to_feb_leap", "EOMONTH(DATE(2020,3,15),-1)", dateSerial(2020, time.February, 29), false, 0},

		// Start on last day: Jan 31 + 1 → Feb 28 (non-leap)
		{"start_last_day_jan31_plus1", "EOMONTH(DATE(2025,1,31),1)", dateSerial(2025, time.February, 28), false, 0},

		// Start on last day: Jan 31 + 1 → Feb 29 (leap year 2024)
		{"start_last_day_jan31_plus1_leap", "EOMONTH(DATE(2024,1,31),1)", dateSerial(2024, time.February, 29), false, 0},

		// Large offset: +12 months = same month next year last day
		{"large_offset_plus12", "EOMONTH(DATE(2025,3,15),12)", dateSerial(2026, time.March, 31), false, 0},

		// Negative large offset: -24 months = two years back
		{"large_offset_minus24", "EOMONTH(DATE(2025,6,10),-24)", dateSerial(2023, time.June, 30), false, 0},

		// December → January crossing (forward)
		{"dec_to_jan_forward", "EOMONTH(DATE(2025,12,15),1)", dateSerial(2026, time.January, 31), false, 0},

		// January → December crossing (backward)
		{"jan_to_dec_backward", "EOMONTH(DATE(2025,1,15),-1)", dateSerial(2024, time.December, 31), false, 0},

		// February start in leap year: Feb 1 2024, 0 months → Feb 29
		{"feb_start_leap_zero", "EOMONTH(DATE(2024,2,1),0)", dateSerial(2024, time.February, 29), false, 0},

		// February start in non-leap year: Feb 1 2025, 0 months → Feb 28
		{"feb_start_nonleap_zero", "EOMONTH(DATE(2025,2,1),0)", dateSerial(2025, time.February, 28), false, 0},

		// 30-day months: April
		{"thirty_day_april", "EOMONTH(DATE(2025,4,10),0)", dateSerial(2025, time.April, 30), false, 0},

		// 30-day months: June
		{"thirty_day_june", "EOMONTH(DATE(2025,6,1),0)", dateSerial(2025, time.June, 30), false, 0},

		// 30-day months: September
		{"thirty_day_september", "EOMONTH(DATE(2025,9,20),0)", dateSerial(2025, time.September, 30), false, 0},

		// 30-day months: November
		{"thirty_day_november", "EOMONTH(DATE(2025,11,5),0)", dateSerial(2025, time.November, 30), false, 0},

		// 31-day months: March
		{"thirtyone_day_march", "EOMONTH(DATE(2025,3,1),0)", dateSerial(2025, time.March, 31), false, 0},

		// 31-day months: May
		{"thirtyone_day_may", "EOMONTH(DATE(2025,5,15),0)", dateSerial(2025, time.May, 31), false, 0},

		// 31-day months: July
		{"thirtyone_day_july", "EOMONTH(DATE(2025,7,4),0)", dateSerial(2025, time.July, 31), false, 0},

		// 31-day months: August
		{"thirtyone_day_august", "EOMONTH(DATE(2025,8,20),0)", dateSerial(2025, time.August, 31), false, 0},

		// 31-day months: October
		{"thirtyone_day_october", "EOMONTH(DATE(2025,10,1),0)", dateSerial(2025, time.October, 31), false, 0},

		// 31-day months: December
		{"thirtyone_day_december", "EOMONTH(DATE(2025,12,25),0)", dateSerial(2025, time.December, 31), false, 0},

		// Multi-year forward: +36 months = 3 years
		{"multi_year_plus36", "EOMONTH(DATE(2020,2,15),36)", dateSerial(2023, time.February, 28), false, 0},

		// Leap year to leap year: Feb 2020 + 48 → Feb 2024
		{"leap_to_leap_plus48", "EOMONTH(DATE(2020,2,15),48)", dateSerial(2024, time.February, 29), false, 0},

		// String date coercion via DATEVALUE
		{"string_coercion_datevalue", `EOMONTH(DATEVALUE("1/15/2025"),1)`, dateSerial(2025, time.February, 28), false, 0},
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

		// YD unit — leap year boundary cases (audit edge cases)
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

func TestDATEDIFComprehensive(t *testing.T) {
	resolver := &mockResolver{}

	// Tests that expect a numeric result.
	numTests := []struct {
		name    string
		formula string
		want    float64
	}{
		// --- Each unit type with basic dates ---
		{"Y_basic", `DATEDIF(DATE(2020,1,15),DATE(2025,3,20),"Y")`, 5},
		{"M_basic", `DATEDIF(DATE(2020,1,15),DATE(2025,3,20),"M")`, 62},
		{"D_basic", `DATEDIF(DATE(2020,1,15),DATE(2025,3,20),"D")`, 1891},
		{"MD_basic", `DATEDIF(DATE(2020,1,15),DATE(2025,3,20),"MD")`, 5},
		{"YM_basic", `DATEDIF(DATE(2020,1,15),DATE(2025,3,20),"YM")`, 2},
		{"YD_basic", `DATEDIF(DATE(2020,1,15),DATE(2025,3,20),"YD")`, 65},

		// --- Exactly 1 year apart ---
		{"Y_exact_1yr", `DATEDIF(DATE(2020,6,15),DATE(2021,6,15),"Y")`, 1},
		{"M_exact_1yr", `DATEDIF(DATE(2020,6,15),DATE(2021,6,15),"M")`, 12},
		{"D_exact_1yr_non_leap", `DATEDIF(DATE(2021,6,15),DATE(2022,6,15),"D")`, 365},
		{"D_exact_1yr_leap", `DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"D")`, 366},

		// --- Same date: all units = 0 ---
		{"Y_same_date", `DATEDIF(DATE(2023,7,4),DATE(2023,7,4),"Y")`, 0},
		{"M_same_date", `DATEDIF(DATE(2023,7,4),DATE(2023,7,4),"M")`, 0},
		{"D_same_date", `DATEDIF(DATE(2023,7,4),DATE(2023,7,4),"D")`, 0},
		{"MD_same_date", `DATEDIF(DATE(2023,7,4),DATE(2023,7,4),"MD")`, 0},
		{"YM_same_date", `DATEDIF(DATE(2023,7,4),DATE(2023,7,4),"YM")`, 0},
		{"YD_same_date2", `DATEDIF(DATE(2023,7,4),DATE(2023,7,4),"YD")`, 0},

		// --- Leap year crossing: Feb 29 handling ---
		{"Y_leap_crossing", `DATEDIF(DATE(2020,2,29),DATE(2024,2,29),"Y")`, 4},
		{"M_leap_crossing", `DATEDIF(DATE(2020,2,29),DATE(2024,2,29),"M")`, 48},
		{"D_leap_crossing", `DATEDIF(DATE(2020,2,29),DATE(2024,2,29),"D")`, 1461},
		{"Y_leap_to_nonleap", `DATEDIF(DATE(2020,2,29),DATE(2021,2,28),"Y")`, 0},
		{"M_leap_to_nonleap", `DATEDIF(DATE(2020,2,29),DATE(2021,2,28),"M")`, 11},

		// --- End of month alignment (Jan 31 -> Feb 28) ---
		{"M_jan31_to_feb28", `DATEDIF(DATE(2023,1,31),DATE(2023,2,28),"M")`, 0},
		{"D_jan31_to_feb28", `DATEDIF(DATE(2023,1,31),DATE(2023,2,28),"D")`, 28},
		{"MD_jan31_to_feb28", `DATEDIF(DATE(2023,1,31),DATE(2023,2,28),"MD")`, 28},

		// --- Multi-year spans ---
		{"Y_5_years", `DATEDIF(DATE(2015,3,10),DATE(2020,3,10),"Y")`, 5},
		{"Y_10_years", `DATEDIF(DATE(2010,6,1),DATE(2020,6,1),"Y")`, 10},
		{"M_10_years", `DATEDIF(DATE(2010,6,1),DATE(2020,6,1),"M")`, 120},

		// --- Short spans ---
		{"D_1_day", `DATEDIF(DATE(2023,5,15),DATE(2023,5,16),"D")`, 1},
		{"D_1_week", `DATEDIF(DATE(2023,5,15),DATE(2023,5,22),"D")`, 7},
		{"Y_1_day", `DATEDIF(DATE(2023,5,15),DATE(2023,5,16),"Y")`, 0},
		{"M_1_day", `DATEDIF(DATE(2023,5,15),DATE(2023,5,16),"M")`, 0},

		// --- "MD" tricky cases ---
		{"MD_jan31_to_mar1", `DATEDIF(DATE(2023,1,31),DATE(2023,3,1),"MD")`, 1},
		{"MD_jan31_to_mar31", `DATEDIF(DATE(2023,1,31),DATE(2023,3,31),"MD")`, 0},
		{"MD_feb15_to_mar10", `DATEDIF(DATE(2023,2,15),DATE(2023,3,10),"MD")`, 23},
		{"MD_across_months", `DATEDIF(DATE(2023,3,25),DATE(2023,4,5),"MD")`, 11},

		// --- "YM" for 13 months = 1 ---
		{"YM_13_months", `DATEDIF(DATE(2022,1,15),DATE(2023,2,20),"YM")`, 1},
		{"YM_25_months", `DATEDIF(DATE(2020,6,10),DATE(2022,7,15),"YM")`, 1},
		{"YM_11_months", `DATEDIF(DATE(2023,1,15),DATE(2023,12,20),"YM")`, 11},

		// --- "YD" crossing year boundary ---
		{"YD_dec_to_jan", `DATEDIF(DATE(2022,12,15),DATE(2023,1,15),"YD")`, 31},
		{"YD_oct_to_feb", `DATEDIF(DATE(2022,10,1),DATE(2024,2,1),"YD")`, 123},

		// --- Case insensitive unit ---
		{"case_lower_y", `DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"y")`, 1},
		{"case_lower_m", `DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"m")`, 12},
		{"case_lower_d", `DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"d")`, 366},
		{"case_lower_md", `DATEDIF(DATE(2020,1,1),DATE(2020,2,15),"md")`, 14},
		{"case_lower_ym", `DATEDIF(DATE(2020,1,1),DATE(2021,3,1),"ym")`, 2},
		{"case_lower_yd", `DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"yd")`, 0},

		// --- Cross-check: DATEDIF("Y")*12 + DATEDIF("YM") = DATEDIF("M") ---
		// We verify both sides evaluate to the same number.
		{"crosscheck_M_via_Y_and_YM",
			`DATEDIF(DATE(2018,3,15),DATE(2025,7,22),"Y")*12+DATEDIF(DATE(2018,3,15),DATE(2025,7,22),"YM")`,
			88}, // 7*12+4 = 88, same as DATEDIF "M"
		{"crosscheck_M_direct",
			`DATEDIF(DATE(2018,3,15),DATE(2025,7,22),"M")`,
			88},

		// --- Known age calculation pattern ---
		{"age_calc", `DATEDIF(DATE(1990,5,20),DATE(2025,3,14),"Y")`, 34},
	}

	for _, tc := range numTests {
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

	// Tests that expect an error.
	errTests := []struct {
		name    string
		formula string
		errVal  ErrorValue
	}{
		{"start_after_end", `DATEDIF(DATE(2025,1,1),DATE(2020,1,1),"Y")`, ErrValNUM},
		{"invalid_unit_X", `DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"X")`, ErrValNUM},
		{"invalid_unit_empty", `DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"")`, ErrValNUM},
		{"error_propagation_start", `DATEDIF(1/0,DATE(2021,1,1),"Y")`, ErrValDIV0},
		{"error_propagation_end", `DATEDIF(DATE(2020,1,1),1/0,"Y")`, ErrValDIV0},
		{"too_few_args", `DATEDIF(DATE(2020,1,1),DATE(2021,1,1))`, ErrValVALUE},
		{"too_many_args", `DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"Y",1)`, ErrValVALUE},
	}

	for _, tc := range errTests {
		t.Run(tc.name, func(t *testing.T) {
			cf := evalCompile(t, tc.formula)
			got, err := Eval(cf, resolver, nil)
			if err != nil {
				t.Fatalf("Eval(%s): unexpected Go error: %v", tc.formula, err)
			}
			if got.Type != ValueError || got.Err != tc.errVal {
				t.Errorf("%s: got %v, want error %v", tc.formula, got, tc.errVal)
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
		// Serial 60 = fictional Feb 29, 1900
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

			// Doc examples
			{"doc_example_1", "TIME(12,0,0)", 0.5},
			{"doc_example_2", "TIME(16,48,10)", 0.700115740740741},

			// Large valid values
			{"hour_32767", "TIME(32767,0,0)", float64(32767%24) * 3600 / 86400},

			// Mixed negative args where total seconds is still positive
			{"mixed_neg_min", "TIME(1,-30,0)", 1800.0 / 86400.0}, // 0:30
			{"mixed_neg_sec", "TIME(0,30,-1)", 1799.0 / 86400.0}, // 0:29:59
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
			{"second_86400", "TIME(0,0,86400)", ErrValNUM},

			// Negative args producing negative total
			{"negative_hour", "TIME(-1,0,0)", ErrValNUM},
			{"negative_12_hour", "TIME(-12,0,0)", ErrValNUM},
			{"negative_minute", "TIME(0,-1,0)", ErrValNUM},
			{"negative_second", "TIME(0,0,-1)", ErrValNUM},
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

	t.Run("valid_times", func(t *testing.T) {
		tests := []struct {
			name    string
			formula string
			want    float64
		}{
			// Original tests
			{"noon_24h", `TIMEVALUE("12:00")`, 0.5},
			{"6:30 PM", `TIMEVALUE("6:30 PM")`, 0.7708333333333334},
			{"midnight_0:00", `TIMEVALUE("0:00")`, 0},
			{"23:59:59", `TIMEVALUE("23:59:59")`, 0.999988425925926},
			{"1:30:45_24h", `TIMEVALUE("1:30:45")`, 0.06302083333333333},
			{"12:00 AM", `TIMEVALUE("12:00 AM")`, 0},
			{"12:00 PM", `TIMEVALUE("12:00 PM")`, 0.5},

			// Basic quarter-day times
			{"6:00 AM", `TIMEVALUE("6:00 AM")`, 0.25},
			{"6:00 PM", `TIMEVALUE("6:00 PM")`, 0.75},
			{"3:00 AM", `TIMEVALUE("3:00 AM")`, 0.125},
			{"9:00 PM", `TIMEVALUE("9:00 PM")`, 0.875},

			// End of day
			{"11:59:59 PM", `TIMEVALUE("11:59:59 PM")`, 0.999988425925926},

			// 24-hour format
			{"13:00", `TIMEVALUE("13:00")`, 13.0 / 24.0},
			{"23:59", `TIMEVALUE("23:59")`, (23*3600 + 59*60) / 86400.0},
			{"6:00_24h", `TIMEVALUE("6:00")`, 0.25},
			{"18:00_24h", `TIMEVALUE("18:00")`, 0.75},

			// With seconds
			{"1:30:30 PM", `TIMEVALUE("1:30:30 PM")`, (13*3600 + 30*60 + 30) / 86400.0},
			{"12:00:00_24h", `TIMEVALUE("12:00:00")`, 0.5},
			{"0:00:01", `TIMEVALUE("0:00:01")`, 1.0 / 86400.0},

			// Minutes only (hour zero)
			{"0:30", `TIMEVALUE("0:30")`, 30.0 * 60.0 / 86400.0},
			{"0:01", `TIMEVALUE("0:01")`, 60.0 / 86400.0},
			{"0:59", `TIMEVALUE("0:59")`, 59.0 * 60.0 / 86400.0},

			// Noon variations
			{"12:00:00 PM", `TIMEVALUE("12:00:00 PM")`, 0.5},
			{"12:00:00_noon", `TIMEVALUE("12:00:00")`, 0.5},

			// Edge times around midnight
			{"12:00:01 AM", `TIMEVALUE("12:00:01 AM")`, 1.0 / 86400.0},

			// Edge times around noon
			{"11:59:59 AM", `TIMEVALUE("11:59:59 AM")`, (11*3600 + 59*60 + 59) / 86400.0},
			{"12:00:01 PM", `TIMEVALUE("12:00:01 PM")`, (12*3600 + 1) / 86400.0},

			// AM/PM variants
			{"1:00 AM", `TIMEVALUE("1:00 AM")`, 1.0 / 24.0},
			{"1:00 PM", `TIMEVALUE("1:00 PM")`, 13.0 / 24.0},
			{"11:00 AM", `TIMEVALUE("11:00 AM")`, 11.0 / 24.0},
			{"11:00 PM", `TIMEVALUE("11:00 PM")`, 23.0 / 24.0},

			// Specific Excel-known values
			{"2:00 AM", `TIMEVALUE("2:00 AM")`, 2.0 / 24.0},
			{"4:30 PM", `TIMEVALUE("4:30 PM")`, (16*3600 + 30*60) / 86400.0},
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
	})

	t.Run("errors", func(t *testing.T) {
		tests := []struct {
			name    string
			formula string
		}{
			// Wrong argument count
			{"no_args", `TIMEVALUE()`},
			{"two_args", `TIMEVALUE("12:00","13:00")`},

			// Non-time strings
			{"non_time_string", `TIMEVALUE("hello")`},
			{"empty_string", `TIMEVALUE("")`},
			{"date_only", `TIMEVALUE("2024-01-01")`},
			{"random_text", `TIMEVALUE("abc:def")`},

			// Boolean input (converted to "TRUE"/"FALSE" text, not parseable)
			{"boolean_true", `TIMEVALUE(TRUE)`},
			{"boolean_false", `TIMEVALUE(FALSE)`},

			// Numeric input (converted to text like "123", not parseable)
			{"numeric_input", `TIMEVALUE(123)`},
			{"numeric_zero", `TIMEVALUE(0)`},

			// Invalid time values
			{"invalid_minute_60", `TIMEVALUE("12:60")`},
			{"invalid_second_60", `TIMEVALUE("12:00:60")`},
			{"invalid_hour_13_am", `TIMEVALUE("13:00 AM")`},
			{"invalid_hour_0_am", `TIMEVALUE("0:00 AM")`},
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
				if got.Err != ErrValVALUE {
					t.Errorf("%s: got error %v, want %v", tc.formula, got.Err, ErrValVALUE)
				}
			})
		}
	})

	t.Run("error_propagation", func(t *testing.T) {
		cf := evalCompile(t, `TIMEVALUE(1/0)`)
		got, err := Eval(cf, resolver, nil)
		if err != nil {
			t.Fatalf("unexpected Go error: %v", err)
		}
		if got.Type != ValueError {
			t.Fatalf("got type %v (%v), want error", got.Type, got)
		}
		if got.Err != ErrValDIV0 {
			t.Errorf("got error %v, want %v", got.Err, ErrValDIV0)
		}
	})
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
			return math.Floor(TimeToSerial(t))
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

		// Doc examples
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
	// Doc dates:
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

		// --- Doc examples ---
		// WORKDAY(10/1/2008, 151) = 4/30/2009 = 39933
		{"doc_no_holidays", "WORKDAY(39722,151)", 39933, false, 0},
		// WORKDAY(10/1/2008, 151, {11/26/2008,12/4/2008,1/21/2009}) = 5/5/2009 = 39938
		{"doc_with_holidays", "WORKDAY(39722,151,{39778,39786,39834})", 39938, false, 0},

		// --- Comprehensive additional WORKDAY tests ---

		// Friday + 1 = next Monday (skips weekend)
		// DATE(2025,1,3) = 45660 (Fri), next Mon = 45663
		{"friday_plus1_skip_weekend", "WORKDAY(DATE(2025,1,3),1)", 45663, false, 0},

		// Monday - 1 = previous Friday
		// DATE(2025,1,6) = 45663 (Mon), prev Fri = 45660
		{"monday_minus1_prev_friday", "WORKDAY(DATE(2025,1,6),-1)", 45660, false, 0},

		// Saturday + 1 = Tuesday (skips to next working day after weekend)
		// DATE(2025,1,4) = 45661 (Sat), next workday = Mon 45663
		{"saturday_plus1_to_monday", "WORKDAY(DATE(2025,1,4),1)", 45663, false, 0},

		// Sunday + 1 = Monday
		// DATE(2025,1,5) = 45662 (Sun), next workday = Mon 45663
		{"sunday_plus1_to_monday", "WORKDAY(DATE(2025,1,5),1)", 45663, false, 0},

		// Zero days on weekday returns same date
		{"zero_days_on_monday", "WORKDAY(DATE(2025,1,6),0)", 45663, false, 0},
		{"zero_days_on_friday", "WORKDAY(DATE(2025,1,3),0)", 45660, false, 0},

		// Zero days on weekend returns weekend serial (per implementation)
		{"zero_days_on_sunday", "WORKDAY(DATE(2025,1,5),0)", 45662, false, 0},

		// 20 working days = 4 weeks (28 calendar days, crossing 4 weekends)
		// Wed Jan 1 + 20 workdays: 4 full weeks of workdays = Wed Jan 29
		{"twenty_days_four_weeks", "WORKDAY(DATE(2025,1,1),20)", dateSerial(2025, time.January, 29), false, 0},

		// Negative with holidays: Mon Jan 13 -5, holiday on Fri Jan 10
		// Without holiday: Fri(1), Thu(2), Wed(3), Tue(4), Mon Jan 6(5) = 45663
		// With holiday on Fri Jan 10 (45667): Thu(1), Wed(2), Tue(3), Mon(4), Fri Jan 3(5) = 45660
		{"negative_with_holiday", "WORKDAY(DATE(2025,1,13),-5,DATE(2025,1,10))", 45660, false, 0},

		// Holiday that falls on a weekend should not double-count
		// Wed Jan 1 + 5, holiday on Sat Jan 4 = same result as no holidays
		{"holiday_on_saturday_no_effect", "WORKDAY(DATE(2025,1,1),5,DATE(2025,1,4))", 45665, false, 0},
		{"holiday_on_sunday_no_effect", "WORKDAY(DATE(2025,1,1),5,DATE(2025,1,5))", 45665, false, 0},

		// Multiple consecutive holidays (Mon+Tue)
		// Wed Jan 1 + 3, holidays on Mon Jan 6 and Tue Jan 7:
		// Thu(1), Fri(2), skip Mon+Tue, Wed Jan 8(3) = 45665
		{"consecutive_holidays", "WORKDAY(DATE(2025,1,1),3,{45663,45664})", 45665, false, 0},

		// 5 work days per week cross-check: Mon Jan 6 + 5 workdays
		// Tue(1), Wed(2), Thu(3), Fri(4), Mon Jan 13(5) = 45670
		{"five_days_from_monday", "WORKDAY(DATE(2025,1,6),5)", dateSerial(2025, time.January, 13), false, 0},

		// Large negative offset
		// Wed Jan 15 - 10 workdays = Wed Jan 1
		{"large_negative_10", "WORKDAY(DATE(2025,1,15),-10)", dateSerial(2025, time.January, 1), false, 0},

		// Start Sunday - 1 = previous Friday
		// Sun Jan 5 - 1 workday → Fri Jan 3
		{"start_sunday_minus1", "WORKDAY(DATE(2025,1,5),-1)", 45660, false, 0},

		// String coercion for days argument
		{"string_days_coercion", `WORKDAY(DATE(2025,1,1),"5")`, 45665, false, 0},

		// Cross-month boundary
		// Fri Jan 31 2025 + 1 = Mon Feb 3 2025
		{"cross_month_jan_to_feb", "WORKDAY(DATE(2025,1,31),1)", dateSerial(2025, time.February, 3), false, 0},

		// Cross-year boundary
		// Wed Dec 31 2025 + 1 = Thu Jan 1 2026 (assuming Jan 1 is a workday)
		{"cross_year_dec_to_jan", "WORKDAY(DATE(2025,12,31),1)", dateSerial(2026, time.January, 1), false, 0},

		// Negative crossing year boundary
		// Thu Jan 1 2026 - 1 = Wed Dec 31 2025
		{"negative_cross_year", "WORKDAY(DATE(2026,1,1),-1)", dateSerial(2025, time.December, 31), false, 0},
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

	// Key dates (2025) — date serial numbers:
	// Jan 1  = 45658 (Wed), Jan 3  = 45660 (Fri), Jan 4  = 45661 (Sat)
	// Jan 5  = 45662 (Sun), Jan 6  = 45663 (Mon), Jan 7  = 45664 (Tue)
	// Jan 8  = 45665 (Wed), Jan 10 = 45667 (Fri), Jan 11 = 45668 (Sat)
	// Jan 12 = 45669 (Sun), Jan 13 = 45670 (Mon), Jan 31 = 45688 (Fri)
	//
	// Doc example dates (serial numbers):
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

		// Doc examples:
		// NETWORKDAYS(10/1/2012, 3/1/2013) = 110
		{"doc_no_holidays", "NETWORKDAYS(41183,41334)", 110, false, 0},
		// NETWORKDAYS(10/1/2012, 3/1/2013, 11/22/2012) = 109
		{"doc_one_holiday", "NETWORKDAYS(41183,41334,41235)", 109, false, 0},
		// NETWORKDAYS(10/1/2012, 3/1/2013, {11/22/2012,12/4/2012,1/21/2013}) = 107
		{"doc_three_holidays", "NETWORKDAYS(41183,41334,{41235,41247,41295})", 107, false, 0},

		// Too few args → error
		{"too_few_args_zero", "NETWORKDAYS()", 0, true, ErrValVALUE},
		{"too_few_args_one", "NETWORKDAYS(45663)", 0, true, ErrValVALUE},

		// Too many args → error
		{"too_many_args", "NETWORKDAYS(45663,45667,45665,1)", 0, true, ErrValVALUE},

		// ── DATE()-based readability tests ────────────────────────────
		// Mon 2024-01-01 to Fri 2024-01-05 = 5 working days
		{"date_mon_to_fri", "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5))", 5, false, 0},
		// Fri 2024-01-05 to Mon 2024-01-08 = 2 (Fri + Mon, skip Sat/Sun)
		{"date_fri_to_mon", "NETWORKDAYS(DATE(2024,1,5),DATE(2024,1,8))", 2, false, 0},
		// Same day weekday (Mon)
		{"date_same_day_weekday", "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,1))", 1, false, 0},
		// Same day weekend (Sat 2024-01-06)
		{"date_same_day_saturday", "NETWORKDAYS(DATE(2024,1,6),DATE(2024,1,6))", 0, false, 0},
		// Same day weekend (Sun 2024-01-07)
		{"date_same_day_sunday", "NETWORKDAYS(DATE(2024,1,7),DATE(2024,1,7))", 0, false, 0},

		// ── Start on weekend ──────────────────────────────────────────
		// Sat to following Fri = 5 weekdays
		{"start_saturday_to_fri", "NETWORKDAYS(DATE(2024,1,6),DATE(2024,1,12))", 5, false, 0},
		// Sun to following Fri = 5 weekdays
		{"start_sunday_to_fri", "NETWORKDAYS(DATE(2024,1,7),DATE(2024,1,12))", 5, false, 0},

		// ── End on weekend ────────────────────────────────────────────
		// Mon to Sat = 5 weekdays (Mon-Fri)
		{"end_on_saturday", "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,6))", 5, false, 0},
		// Mon to Sun = 5 weekdays
		{"end_on_sunday", "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7))", 5, false, 0},

		// ── Cross-month boundary ──────────────────────────────────────
		// Wed Jan 31 to Fri Feb 2, 2024 = 3 (Wed, Thu, Fri)
		{"cross_month_jan_feb", "NETWORKDAYS(DATE(2024,1,31),DATE(2024,2,2))", 3, false, 0},
		// Fri Feb 28 to Mon Mar 4 2025 = 3 (Fri, Mon Mar 3, Tue Mar 4)... wait
		// Feb 28 = Fri, Mar 3 = Mon, Mar 4 = Tue. Count: Feb28(Fri)=1, Mar3(Mon)=2, Mar4(Tue)=3
		{"cross_month_feb_mar", "NETWORKDAYS(DATE(2025,2,28),DATE(2025,3,4))", 3, false, 0},

		// ── Cross-year boundary ───────────────────────────────────────
		// Tue Dec 31, 2024 to Thu Jan 2, 2025 = 3 (Tue, Wed=Jan1, Thu=Jan2)
		{"cross_year_2024_2025", "NETWORKDAYS(DATE(2024,12,31),DATE(2025,1,2))", 3, false, 0},
		// Fri Dec 29, 2023 to Tue Jan 2, 2024
		// Dec29(Fri)=1, Dec30(Sat)=skip, Dec31(Sun)=skip, Jan1(Mon)=2, Jan2(Tue)=3
		{"cross_year_2023_2024", "NETWORKDAYS(DATE(2023,12,29),DATE(2024,1,2))", 3, false, 0},

		// ── Full year ≈ 261 working days ──────────────────────────────
		// 2024 is a leap year: Jan 1 (Mon) to Dec 31 (Tue)
		// 366 calendar days, 52 full weeks = 260 weekdays + 2 extra days (Mon, Tue) = 262
		{"full_year_2024", "NETWORKDAYS(DATE(2024,1,1),DATE(2024,12,31))", 262, false, 0},
		// 2025: Jan 1 (Wed) to Dec 31 (Wed)
		// 365 calendar days, 52 weeks + 1 day. 52*5=260 + 1 (Wed) = 261
		{"full_year_2025", "NETWORKDAYS(DATE(2025,1,1),DATE(2025,12,31))", 261, false, 0},

		// ── Large span with holidays ──────────────────────────────────
		// Full January 2024 with 1 holiday (Jan 15 Mon = MLK day)
		// Jan 1 (Mon) to Jan 31 (Wed): 23 weekdays - 1 holiday = 22
		{"jan_2024_with_mlk", "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),DATE(2024,1,15))", 22, false, 0},

		// ── Multiple holidays as array ────────────────────────────────
		// Mon-Fri with 3 holidays (Tue, Wed, Thu) = 2 remaining
		{"three_weekday_holidays", "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),{DATE(2024,1,2),DATE(2024,1,3),DATE(2024,1,4)})", 2, false, 0},

		// ── Holiday on weekend (no double-count) ──────────────────────
		// Mon to Fri with holiday on Sat = still 5
		{"holiday_on_sat_no_effect", "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),DATE(2024,1,6))", 5, false, 0},

		// ── All weekdays are holidays → 0 ─────────────────────────────
		// Mon-Fri, all 5 weekdays are holidays
		{"all_holidays_zero", "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),{DATE(2024,1,1),DATE(2024,1,2),DATE(2024,1,3),DATE(2024,1,4),DATE(2024,1,5)})", 0, false, 0},

		// ── Negative span (end before start) ──────────────────────────
		{"date_negative_span", "NETWORKDAYS(DATE(2024,1,5),DATE(2024,1,1))", -5, false, 0},
		{"date_negative_cross_year", "NETWORKDAYS(DATE(2025,1,2),DATE(2024,12,31))", -3, false, 0},

		// ── Error propagation ──────────────────────────────────────────
		{"error_prop_start", "NETWORKDAYS(1/0,45667)", 0, true, ErrValDIV0},
		{"error_prop_end", "NETWORKDAYS(45663,1/0)", 0, true, ErrValDIV0},
		{"error_prop_holiday", "NETWORKDAYS(45663,45667,1/0)", 0, true, ErrValDIV0},

		// ── Two-week span ─────────────────────────────────────────────
		// Mon Jan 1 to Fri Jan 12 = 10 weekdays
		{"two_weeks", "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,12))", 10, false, 0},

		// ── February in leap year ─────────────────────────────────────
		// Feb 1 (Thu) to Feb 29 (Thu) 2024 = 21 weekdays
		{"feb_leap_year", "NETWORKDAYS(DATE(2024,2,1),DATE(2024,2,29))", 21, false, 0},

		// ── Duplicate holidays (should not double-subtract) ───────────
		{"duplicate_holiday", "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5),{DATE(2024,1,3),DATE(2024,1,3)})", 4, false, 0},
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

func TestWEEKNUM(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		// Doc examples: Mar 9, 2012
		{"doc_example_default", "WEEKNUM(DATE(2012,3,9))", 10, false, 0},
		{"doc_example_rt2", "WEEKNUM(DATE(2012,3,9),2)", 11, false, 0},

		// Default return_type (1 = Sunday start)
		{"jan1_2023_default", "WEEKNUM(DATE(2023,1,1))", 1, false, 0},     // Jan 1, 2023 (Sunday)
		{"jan1_2024_default", "WEEKNUM(DATE(2024,1,1))", 1, false, 0},     // Jan 1, 2024 (Monday)
		{"dec31_2023_default", "WEEKNUM(DATE(2023,12,31))", 53, false, 0}, // Dec 31, 2023 (Sunday)
		{"dec31_2024_default", "WEEKNUM(DATE(2024,12,31))", 53, false, 0}, // Dec 31, 2024 (Tuesday)
		{"jun15_2023_default", "WEEKNUM(DATE(2023,6,15))", 24, false, 0},  // Jun 15, 2023 (Thursday)
		{"jan7_2023_sat_default", "WEEKNUM(DATE(2023,1,7))", 1, false, 0}, // Jan 7, 2023 (Saturday) - last day of week 1
		{"jan8_2023_sun_default", "WEEKNUM(DATE(2023,1,8))", 2, false, 0}, // Jan 8, 2023 (Sunday) - first day of week 2

		// return_type 2 (Monday start)
		{"jan1_2023_rt2", "WEEKNUM(DATE(2023,1,1),2)", 1, false, 0},    // Jan 1, 2023 (Sunday)
		{"jan2_2023_rt2", "WEEKNUM(DATE(2023,1,2),2)", 2, false, 0},    // Jan 2, 2023 (Monday) - new week
		{"jun15_2023_rt2", "WEEKNUM(DATE(2023,6,15),2)", 25, false, 0}, // Jun 15, 2023 (Thursday)
		{"jul4_2023_rt2", "WEEKNUM(DATE(2023,7,4),2)", 28, false, 0},   // Jul 4, 2023 (Tuesday)

		// return_type 21 (ISO week, Monday start, System 2)
		{"jan1_2023_iso", "WEEKNUM(DATE(2023,1,1),21)", 52, false, 0},    // Jan 1, 2023 (Sun) -> ISO week 52 of 2022
		{"jan1_2024_iso", "WEEKNUM(DATE(2024,1,1),21)", 1, false, 0},     // Jan 1, 2024 (Mon) -> ISO week 1
		{"dec31_2023_iso", "WEEKNUM(DATE(2023,12,31),21)", 52, false, 0}, // Dec 31, 2023 -> ISO week 52
		{"dec31_2024_iso", "WEEKNUM(DATE(2024,12,31),21)", 1, false, 0},  // Dec 31, 2024 (Tue) -> ISO week 1 of 2025
		{"jan1_2021_iso", "WEEKNUM(DATE(2021,1,1),21)", 53, false, 0},    // Jan 1, 2021 (Fri) -> ISO week 53 of 2020

		// Various return_type values (other week start days) - Mar 1, 2023 (Wednesday)
		{"mar1_2023_rt11", "WEEKNUM(DATE(2023,3,1),11)", 10, false, 0}, // rt 11 = Monday start
		{"mar1_2023_rt12", "WEEKNUM(DATE(2023,3,1),12)", 10, false, 0}, // rt 12 = Tuesday start
		{"mar1_2023_rt13", "WEEKNUM(DATE(2023,3,1),13)", 10, false, 0}, // rt 13 = Wednesday start
		{"mar1_2023_rt14", "WEEKNUM(DATE(2023,3,1),14)", 9, false, 0},  // rt 14 = Thursday start
		{"mar1_2023_rt15", "WEEKNUM(DATE(2023,3,1),15)", 9, false, 0},  // rt 15 = Friday start
		{"mar1_2023_rt16", "WEEKNUM(DATE(2023,3,1),16)", 9, false, 0},  // rt 16 = Saturday start
		{"mar1_2023_rt17", "WEEKNUM(DATE(2023,3,1),17)", 9, false, 0},  // rt 17 = Sunday start (same as 1)

		// Leap year date
		{"leap_day_2024", "WEEKNUM(DATE(2024,2,29))", 9, false, 0}, // Feb 29, 2024 (Thursday)

		// Dec 31 / year boundary edge cases
		{"dec31_2020_default", "WEEKNUM(DATE(2020,12,31))", 53, false, 0}, // Dec 31, 2020 (Thursday)
		{"dec31_2020_rt14", "WEEKNUM(DATE(2020,12,31),14)", 54, false, 0}, // rt 14 = Thursday start -> week 54
		{"jan1_2025_default", "WEEKNUM(DATE(2025,1,1))", 1, false, 0},     // Jan 1, 2025 (Wednesday)

		// Error cases: wrong argument count
		{"no_args", "WEEKNUM()", 0, true, ErrValVALUE},
		{"too_many_args", "WEEKNUM(44927,1,1)", 0, true, ErrValVALUE},

		// Error cases: invalid return_type
		{"invalid_rt_0", "WEEKNUM(44927,0)", 0, true, ErrValNUM},
		{"invalid_rt_3", "WEEKNUM(44927,3)", 0, true, ErrValNUM},
		{"invalid_rt_10", "WEEKNUM(44927,10)", 0, true, ErrValNUM},
		{"invalid_rt_18", "WEEKNUM(44927,18)", 0, true, ErrValNUM},
		{"invalid_rt_20", "WEEKNUM(44927,20)", 0, true, ErrValNUM},
		{"invalid_rt_22", "WEEKNUM(44927,22)", 0, true, ErrValNUM},

		// Error propagation
		{"error_in_serial", `WEEKNUM("abc")`, 0, true, ErrValVALUE},
		{"error_in_rt", `WEEKNUM(44927,"abc")`, 0, true, ErrValVALUE},
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

func TestISOWEEKNUM(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		// Documentation example: March 9, 2012 = ISO week 10
		{"doc_mar_9_2012", "ISOWEEKNUM(40977)", 10, false, 0},
		// Using DATE() to construct the same date
		{"doc_via_date", "ISOWEEKNUM(DATE(2012,3,9))", 10, false, 0},

		// Jan 1 that falls in ISO week 1 (Thu Jan 1, 2015)
		{"jan1_week1_2015", "ISOWEEKNUM(42005)", 1, false, 0},
		// Jan 1 that falls in ISO week 1 (Wed Jan 1, 2014)
		{"jan1_week1_2014", "ISOWEEKNUM(41640)", 1, false, 0},
		// Jan 1 that belongs to previous year's last week (Fri Jan 1, 2016 = week 53 of 2015)
		{"jan1_prev_year_2016", "ISOWEEKNUM(42370)", 53, false, 0},
		// Jan 1 that belongs to previous year's last week (Sun Jan 1, 2017 = week 52 of 2016)
		{"jan1_prev_year_2017", "ISOWEEKNUM(42736)", 52, false, 0},
		// Jan 1 that belongs to previous year's last week (Fri Jan 1, 2010 = week 53 of 2009)
		{"jan1_prev_year_2010", "ISOWEEKNUM(40179)", 53, false, 0},
		// Jan 1, 2021 (Fri) = week 53 of 2020
		{"jan1_prev_year_2021", "ISOWEEKNUM(44197)", 53, false, 0},

		// Dec 31 in a year with 53 ISO weeks (Thu Dec 31, 2015)
		{"dec31_week53_2015", "ISOWEEKNUM(42369)", 53, false, 0},
		// Dec 31 that falls in ISO week 1 of the next year (Wed Dec 31, 2014)
		{"dec31_week1_next_2014", "ISOWEEKNUM(42004)", 1, false, 0},
		// Dec 31 in a year with 53 ISO weeks (Thu Dec 31, 2009)
		{"dec31_week53_2009", "ISOWEEKNUM(40178)", 53, false, 0},
		// Dec 31 that falls in ISO week 1 of the next year (Mon Dec 31, 2012)
		{"dec31_week1_next_2012", "ISOWEEKNUM(41274)", 1, false, 0},
		// Dec 31, 2020 (Thu) = week 53
		{"dec31_week53_2020", "ISOWEEKNUM(44196)", 53, false, 0},

		// Mid-year dates
		{"mid_year_jun_15_2023", "ISOWEEKNUM(45092)", 24, false, 0},
		{"mid_year_jul_4_2023", "ISOWEEKNUM(45111)", 27, false, 0},
		{"mid_year_sep_1_2023", "ISOWEEKNUM(45170)", 35, false, 0},

		// Leap year date: Feb 29, 2024 = week 9
		{"leap_year_feb29_2024", "ISOWEEKNUM(45351)", 9, false, 0},

		// Early serial numbers
		// Serial 1 = Jan 1, 1900 (Monday) = ISO week 1
		{"serial_1_jan1_1900", "ISOWEEKNUM(1)", 1, false, 0},
		// Serial 7 = Jan 7, 1900 (Sunday) = ISO week 1
		{"serial_7_jan7_1900", "ISOWEEKNUM(7)", 1, false, 0},
		// Serial 0 = "Jan 0, 1900" mapped to Dec 31, 1899 = ISO week 52
		{"serial_0", "ISOWEEKNUM(0)", 52, false, 0},

		// Fractional serial: should use the date portion only
		{"fractional_serial", "ISOWEEKNUM(40977.75)", 10, false, 0},

		// Boolean TRUE coerced to 1 = Jan 1, 1900 = week 1
		{"bool_true", "ISOWEEKNUM(TRUE)", 1, false, 0},

		// Error cases
		// No arguments
		{"no_args", "ISOWEEKNUM()", 0, true, ErrValVALUE},
		// Too many arguments
		{"too_many_args", "ISOWEEKNUM(1,2)", 0, true, ErrValVALUE},
		// Non-numeric string
		{"non_numeric_string", `ISOWEEKNUM("abc")`, 0, true, ErrValVALUE},
		// Error propagation
		{"error_propagation", `ISOWEEKNUM("hello")`, 0, true, ErrValVALUE},
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

func TestYEARFRAC(t *testing.T) {
	resolver := &mockResolver{}

	t.Run("values", func(t *testing.T) {
		tests := []struct {
			name    string
			formula string
			want    float64
		}{
			// === Documentation examples ===
			// 1/1/2012 to 7/30/2012
			{"doc_basis0_default", "YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))", 0.58055555555555556},
			{"doc_basis1", "YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)", 0.57650273224043716},
			{"doc_basis3", "YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)", 0.57808219178082187},

			// === Same date returns 0 ===
			{"same_date_basis0", "YEARFRAC(DATE(2023,6,15),DATE(2023,6,15),0)", 0},
			{"same_date_basis1", "YEARFRAC(DATE(2023,6,15),DATE(2023,6,15),1)", 0},
			{"same_date_basis2", "YEARFRAC(DATE(2023,6,15),DATE(2023,6,15),2)", 0},
			{"same_date_basis3", "YEARFRAC(DATE(2023,6,15),DATE(2023,6,15),3)", 0},
			{"same_date_basis4", "YEARFRAC(DATE(2023,6,15),DATE(2023,6,15),4)", 0},

			// === Full year = 1.0 ===
			// Basis 0: 30/360 -> 360/360 = 1
			{"full_year_basis0", "YEARFRAC(DATE(2023,1,1),DATE(2024,1,1),0)", 1.0},
			// Basis 1: actual/actual, spans 2023-2024, avg=(365+366)/2=365.5, 365/365.5
			{"full_year_basis1_nonleap", "YEARFRAC(DATE(2023,1,1),DATE(2024,1,1),1)", 365.0 / 365.5},
			// Basis 3: actual/365 -> 365/365 = 1
			{"full_year_basis3_nonleap", "YEARFRAC(DATE(2023,1,1),DATE(2024,1,1),3)", 1.0},
			// Basis 4: European 30/360 -> 360/360 = 1
			{"full_year_basis4", "YEARFRAC(DATE(2023,1,1),DATE(2024,1,1),4)", 1.0},

			// === Full year in a leap year ===
			// Basis 1: actual/actual, spans 2024-2025, avg=(366+365)/2=365.5, 366/365.5
			{"full_year_basis1_leap", "YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),1)", 366.0 / 365.5},
			// Basis 2: actual/360, 366 days -> 366/360
			{"full_year_basis2_leap", "YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),2)", 366.0 / 360.0},
			// Basis 3: actual/365, 366 days -> 366/365
			{"full_year_basis3_leap", "YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),3)", 366.0 / 365.0},

			// === Half year ~ 0.5 ===
			// Basis 0: Jan 1 to Jul 1 -> 180/360 = 0.5
			{"half_year_basis0", "YEARFRAC(DATE(2023,1,1),DATE(2023,7,1),0)", 0.5},
			// Basis 3: Jan 1 to Jul 3 -> 183/365
			{"half_year_basis3", "YEARFRAC(DATE(2023,1,1),DATE(2023,7,3),3)", 183.0 / 365.0},
			// Basis 4: Jan 1 to Jul 1 -> 180/360 = 0.5 (European 30/360)
			{"half_year_basis4", "YEARFRAC(DATE(2023,1,1),DATE(2023,7,1),4)", 0.5},

			// === Reversed dates (end before start) should swap ===
			{"reversed_dates_basis0", "YEARFRAC(DATE(2024,1,1),DATE(2023,1,1),0)", 1.0},
			// Reversed, same as full_year_basis1_nonleap after swap: 365/365.5
			{"reversed_dates_basis1", "YEARFRAC(DATE(2024,1,1),DATE(2023,1,1),1)", 365.0 / 365.5},
			{"reversed_dates_basis3", "YEARFRAC(DATE(2023,7,3),DATE(2023,1,1),3)", 183.0 / 365.0},

			// === Basis 0 (US 30/360) specific cases ===
			// 30-day month adjustment: Jan 31 to Feb 28
			{"basis0_jan31_feb28", "YEARFRAC(DATE(2023,1,31),DATE(2023,2,28),0)", 28.0 / 360.0},
			// Both end-of-month dates
			{"basis0_eom_to_eom", "YEARFRAC(DATE(2023,1,31),DATE(2023,3,31),0)", 60.0 / 360.0},

			// === Basis 1 (Actual/actual) specific cases ===
			// Spanning 2023-2024 boundary: 366 actual days, avg=(365+366)/2=365.5
			{"basis1_span_leap", "YEARFRAC(DATE(2023,7,1),DATE(2024,7,1),1)", 366.0 / 365.5},

			// === Basis 2 (Actual/360) ===
			// 90 actual days / 360
			{"basis2_quarter", "YEARFRAC(DATE(2023,1,1),DATE(2023,4,1),2)", 90.0 / 360.0},

			// === Basis 4 (European 30/360) specific cases ===
			// Day 31 -> 30 adjustment for both dates
			{"basis4_31st_to_31st", "YEARFRAC(DATE(2023,1,31),DATE(2023,3,31),4)", 60.0 / 360.0},

			// === Default basis (omitted = 0) ===
			{"default_basis_omitted", "YEARFRAC(DATE(2023,1,1),DATE(2023,7,1))", 0.5},

			// === Multi-year spans ===
			{"multi_year_basis0", "YEARFRAC(DATE(2020,1,1),DATE(2025,1,1),0)", 5.0},
			// Years 2020-2025: 2 leap (2020,2024) + 4 non-leap = 2192 days total
			// avg = 2192/6 = 365.333..., actual days = 1827
			{"multi_year_basis1", "YEARFRAC(DATE(2020,1,1),DATE(2025,1,1),1)", 1827.0 / (2192.0 / 6.0)},
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
			{"no_args", "YEARFRAC()", ErrValVALUE},
			{"one_arg", "YEARFRAC(DATE(2023,1,1))", ErrValVALUE},

			// Too many args
			{"four_args", "YEARFRAC(DATE(2023,1,1),DATE(2024,1,1),0,1)", ErrValVALUE},

			// Invalid basis values
			{"basis_negative", "YEARFRAC(DATE(2023,1,1),DATE(2024,1,1),-1)", ErrValNUM},
			{"basis_5", "YEARFRAC(DATE(2023,1,1),DATE(2024,1,1),5)", ErrValNUM},
			{"basis_99", "YEARFRAC(DATE(2023,1,1),DATE(2024,1,1),99)", ErrValNUM},

			// Error propagation
			{"error_start", "YEARFRAC(1/0,DATE(2024,1,1),0)", ErrValDIV0},
			{"error_end", "YEARFRAC(DATE(2023,1,1),1/0,0)", ErrValDIV0},
			{"error_basis", "YEARFRAC(DATE(2023,1,1),DATE(2024,1,1),1/0)", ErrValDIV0},
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
	})
}

func TestWEEKDAY(t *testing.T) {
	resolver := &mockResolver{}

	tests := []struct {
		name    string
		formula string
		want    float64
		isErr   bool
		errVal  ErrorValue
	}{
		// === Basic usage with default return_type (1: Sun=1..Sat=7) ===
		// Feb 14, 2008 (serial 39492) is a Thursday
		{"thu_default", "WEEKDAY(39492)", 5, false, 0},
		{"thu_explicit_rt1", "WEEKDAY(39492,1)", 5, false, 0},

		// === All 7 days of the week, return_type 1 (Sun=1..Sat=7) ===
		// Feb 10-16, 2008: Sun through Sat
		{"sunday_rt1", "WEEKDAY(39488,1)", 1, false, 0},
		{"monday_rt1", "WEEKDAY(39489,1)", 2, false, 0},
		{"tuesday_rt1", "WEEKDAY(39490,1)", 3, false, 0},
		{"wednesday_rt1", "WEEKDAY(39491,1)", 4, false, 0},
		{"thursday_rt1", "WEEKDAY(39492,1)", 5, false, 0},
		{"friday_rt1", "WEEKDAY(39493,1)", 6, false, 0},
		{"saturday_rt1", "WEEKDAY(39494,1)", 7, false, 0},

		// === return_type 2 (Mon=1..Sun=7) ===
		{"sunday_rt2", "WEEKDAY(39488,2)", 7, false, 0},
		{"monday_rt2", "WEEKDAY(39489,2)", 1, false, 0},
		{"thursday_rt2", "WEEKDAY(39492,2)", 4, false, 0},
		{"saturday_rt2", "WEEKDAY(39494,2)", 6, false, 0},

		// === return_type 3 (Mon=0..Sun=6) ===
		{"sunday_rt3", "WEEKDAY(39488,3)", 6, false, 0},
		{"monday_rt3", "WEEKDAY(39489,3)", 0, false, 0},
		{"thursday_rt3", "WEEKDAY(39492,3)", 3, false, 0},
		{"saturday_rt3", "WEEKDAY(39494,3)", 5, false, 0},

		// === return_type 11 (Mon=1..Sun=7, same as rt2) ===
		{"sunday_rt11", "WEEKDAY(39488,11)", 7, false, 0},
		{"monday_rt11", "WEEKDAY(39489,11)", 1, false, 0},
		{"thursday_rt11", "WEEKDAY(39492,11)", 4, false, 0},

		// === return_type 12 (Tue=1..Mon=7) ===
		{"sunday_rt12", "WEEKDAY(39488,12)", 6, false, 0},
		{"monday_rt12", "WEEKDAY(39489,12)", 7, false, 0},
		{"tuesday_rt12", "WEEKDAY(39490,12)", 1, false, 0},
		{"thursday_rt12", "WEEKDAY(39492,12)", 3, false, 0},

		// === return_type 13 (Wed=1..Tue=7) ===
		{"sunday_rt13", "WEEKDAY(39488,13)", 5, false, 0},
		{"tuesday_rt13", "WEEKDAY(39490,13)", 7, false, 0},
		{"wednesday_rt13", "WEEKDAY(39491,13)", 1, false, 0},
		{"thursday_rt13", "WEEKDAY(39492,13)", 2, false, 0},

		// === return_type 14 (Thu=1..Wed=7) ===
		{"sunday_rt14", "WEEKDAY(39488,14)", 4, false, 0},
		{"wednesday_rt14", "WEEKDAY(39491,14)", 7, false, 0},
		{"thursday_rt14", "WEEKDAY(39492,14)", 1, false, 0},
		{"friday_rt14", "WEEKDAY(39493,14)", 2, false, 0},

		// === return_type 15 (Fri=1..Thu=7) ===
		{"sunday_rt15", "WEEKDAY(39488,15)", 3, false, 0},
		{"thursday_rt15", "WEEKDAY(39492,15)", 7, false, 0},
		{"friday_rt15", "WEEKDAY(39493,15)", 1, false, 0},
		{"saturday_rt15", "WEEKDAY(39494,15)", 2, false, 0},

		// === return_type 16 (Sat=1..Fri=7) ===
		{"sunday_rt16", "WEEKDAY(39488,16)", 2, false, 0},
		{"friday_rt16", "WEEKDAY(39493,16)", 7, false, 0},
		{"saturday_rt16", "WEEKDAY(39494,16)", 1, false, 0},

		// === return_type 17 (Sun=1..Sat=7, same as rt1) ===
		{"sunday_rt17", "WEEKDAY(39488,17)", 1, false, 0},
		{"thursday_rt17", "WEEKDAY(39492,17)", 5, false, 0},
		{"saturday_rt17", "WEEKDAY(39494,17)", 7, false, 0},

		// === DATE() input instead of raw serial ===
		{"date_func_thu", "WEEKDAY(DATE(2008,2,14),1)", 5, false, 0},
		{"date_func_sun", "WEEKDAY(DATE(2008,2,10),1)", 1, false, 0},

		// === Large serial number (Dec 31, 9999 = serial 2958465, Friday) ===
		{"max_serial", "WEEKDAY(2958465,1)", 6, false, 0},

		// === Fractional serial (time portion ignored for weekday) ===
		// 39492.75 = Feb 14, 2008 at 6pm, still Thursday
		{"fractional_serial", "WEEKDAY(39492.75,1)", 5, false, 0},

		// === Boolean coercion: TRUE -> serial 1 ===
		// Serial 1 maps to Jan 1, 1900 = Sunday in Excel's convention.
		// (Historically Monday, but Excel's 1900 leap-year bug shifts it.)
		{"bool_true", "WEEKDAY(TRUE,1)", 1, false, 0},

		// === Invalid return_type -> #NUM! ===
		{"invalid_rt_0", "WEEKDAY(39492,0)", 0, true, ErrValNUM},
		{"invalid_rt_4", "WEEKDAY(39492,4)", 0, true, ErrValNUM},
		{"invalid_rt_5", "WEEKDAY(39492,5)", 0, true, ErrValNUM},
		{"invalid_rt_10", "WEEKDAY(39492,10)", 0, true, ErrValNUM},
		{"invalid_rt_18", "WEEKDAY(39492,18)", 0, true, ErrValNUM},
		{"invalid_rt_99", "WEEKDAY(39492,99)", 0, true, ErrValNUM},
		{"invalid_rt_neg", "WEEKDAY(39492,-1)", 0, true, ErrValNUM},

		// === Error propagation ===
		{"error_text_serial", `WEEKDAY("hello")`, 0, true, ErrValVALUE},
		{"error_text_rt", `WEEKDAY(39492,"abc")`, 0, true, ErrValVALUE},

		// === Wrong argument count ===
		{"no_args", "WEEKDAY()", 0, true, ErrValVALUE},
		{"too_many_args", "WEEKDAY(39492,1,1)", 0, true, ErrValVALUE},
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
