package formula

import (
	"math"
	"testing"
	"time"
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
