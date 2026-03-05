package formula

import (
	"math"
	"testing"
)

func numArgs(vals ...float64) []Value {
	out := make([]Value, len(vals))
	for i, v := range vals {
		out[i] = NumberVal(v)
	}
	return out
}

func assertClose(t *testing.T, name string, got Value, want float64) {
	t.Helper()
	if got.Type != ValueNumber {
		t.Fatalf("%s: expected number, got type %v (str=%q)", name, got.Type, got.Str)
	}
	if math.Abs(got.Num-want) > 0.01 {
		t.Errorf("%s: got %f, want %f", name, got.Num, want)
	}
}

func assertError(t *testing.T, name string, got Value) {
	t.Helper()
	if got.Type != ValueError {
		t.Fatalf("%s: expected error, got type %v (num=%f)", name, got.Type, got.Num)
	}
}

// === PMT ===

func TestPMT_BasicLoan(t *testing.T) {
	// 5% annual rate / 12 months, 30 year mortgage, $200,000 loan
	// Excel: PMT(0.05/12, 360, 200000) = -1073.64
	v, err := fnPMT(numArgs(0.05/12, 360, 200000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PMT basic", v, -1073.64)
}

func TestPMT_ZeroRate(t *testing.T) {
	// PMT(0, 10, 1000) = -100
	v, err := fnPMT(numArgs(0, 10, 1000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PMT zero rate", v, -100)
}

func TestPMT_WithFV(t *testing.T) {
	// PMT(0.08/12, 120, 0, 100000) — saving for $100k
	// Excel: -546.61
	v, err := fnPMT(numArgs(0.08/12, 120, 0, 100000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PMT with FV", v, -546.61)
}

func TestPMT_ZeroNper(t *testing.T) {
	v, _ := fnPMT(numArgs(0.05, 0, 1000))
	assertError(t, "PMT zero nper", v)
}

// === FV ===

func TestFV_Savings(t *testing.T) {
	// FV(0.06/12, 120, -200, -5000) — $200/month, $5000 initial, 6% over 10 years
	// Excel: 41872.85
	v, err := fnFV(numArgs(0.06/12, 120, -200, -5000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "FV savings", v, 41872.85)
}

func TestFV_ZeroRate(t *testing.T) {
	// FV(0, 10, -100, -1000) = 2000
	v, err := fnFV(numArgs(0, 10, -100, -1000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "FV zero rate", v, 2000)
}

// === PV ===

func TestPV_Annuity(t *testing.T) {
	// PV(0.08/12, 240, -500) — 20 years of $500/month payments at 8%
	// Excel: 59777.15
	v, err := fnPV(numArgs(0.08/12, 240, -500))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PV annuity", v, 59777.15)
}

// === NPER ===

func TestNPER_Loan(t *testing.T) {
	// NPER(0.01, -100, 1000) — how many months to pay off $1000 at 1%/month
	// Excel: 10.58
	v, err := fnNPER(numArgs(0.01, -100, 1000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPER loan", v, 10.58)
}

func TestNPER_ZeroRate(t *testing.T) {
	// NPER(0, -100, 1000) = 10
	v, err := fnNPER(numArgs(0, -100, 1000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPER zero rate", v, 10)
}

// === RATE ===

func TestRATE_Loan(t *testing.T) {
	// RATE(360, -1073.64, 200000) — reverse of the PMT example
	// Excel: ~0.004167 (0.05/12)
	v, err := fnRATE(numArgs(360, -1073.64, 200000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "RATE loan", v, 0.05/12)
}

func TestRATE_NoSolution(t *testing.T) {
	// RATE(60, 0, 50000) — pmt=0, fv=0, pv≠0: no solution exists.
	// Excel returns #NUM!.
	v, err := fnRATE(numArgs(60, 0, 50000))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueError {
		t.Fatalf("expected #NUM! error, got %v", v)
	}
}

// === IPMT ===

func TestIPMT_FirstPayment(t *testing.T) {
	// IPMT(0.05/12, 1, 360, 200000) — interest portion of first payment
	// Excel: -833.33
	v, err := fnIPMT(numArgs(0.05/12, 1, 360, 200000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "IPMT first", v, -833.33)
}

// === PPMT ===

func TestPPMT_FirstPayment(t *testing.T) {
	// PPMT(0.05/12, 1, 360, 200000) — principal portion of first payment
	// Excel: -240.31
	v, err := fnPPMT(numArgs(0.05/12, 1, 360, 200000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PPMT first", v, -240.31)
}

func TestIPMT_Plus_PPMT_Equals_PMT(t *testing.T) {
	rate := 0.05 / 12
	nper := 360.0
	pv := 200000.0
	for _, per := range []float64{1, 12, 60, 180, 360} {
		ipmt, _ := fnIPMT(numArgs(rate, per, nper, pv))
		ppmt, _ := fnPPMT(numArgs(rate, per, nper, pv))
		pmt, _ := fnPMT(numArgs(rate, nper, pv))
		sum := ipmt.Num + ppmt.Num
		if math.Abs(sum-pmt.Num) > 0.01 {
			t.Errorf("per=%v: IPMT(%f) + PPMT(%f) = %f, PMT = %f", per, ipmt.Num, ppmt.Num, sum, pmt.Num)
		}
	}
}

// === NPV ===

func TestNPV_Basic(t *testing.T) {
	// NPV(0.1, -10000, 3000, 4200, 6800)
	// Excel: 1188.44
	v, err := fnNPV(append(numArgs(0.1), numArgs(-10000, 3000, 4200, 6800)...))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV basic", v, 1188.44)
}

// === IRR ===

func TestIRR_Basic(t *testing.T) {
	// IRR({-10000, 3000, 4200, 6800}) = ~0.1634 (16.34%)
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000), NumberVal(3000), NumberVal(4200), NumberVal(6800)},
		},
	}
	v, err := fnIRR([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "IRR basic", v, 0.1634)
}

func TestIRR_NoSolution(t *testing.T) {
	// All positive — no solution
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(100), NumberVal(200)},
		},
	}
	v, _ := fnIRR([]Value{arr})
	assertError(t, "IRR no solution", v)
}

// === SLN ===

func TestSLN_Basic(t *testing.T) {
	// SLN(30000, 7500, 10) = 2250
	v, err := fnSLN(numArgs(30000, 7500, 10))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SLN basic", v, 2250)
}

func TestSLN_ZeroLife(t *testing.T) {
	v, _ := fnSLN(numArgs(30000, 7500, 0))
	assertError(t, "SLN zero life", v)
}

// === XNPV ===

func TestXNPV_Basic(t *testing.T) {
	// XNPV(0.09, {-10000, 2750, 4250, 3250, 2750}, {39448, 39508, 39751, 39859, 39904})
	// dates as Excel serial numbers: 2008-01-01, 2008-03-01, 2008-10-30, 2009-02-15, 2009-04-01
	// Excel: ~2086.65
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000), NumberVal(2750), NumberVal(4250), NumberVal(3250), NumberVal(2750)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39508), NumberVal(39751), NumberVal(39859), NumberVal(39904)},
		},
	}
	v, err := fnXNPV([]Value{NumberVal(0.09), vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "XNPV basic", v, 2086.65)
}

// === XIRR ===

func TestXIRR_Basic(t *testing.T) {
	// XIRR({-10000, 2750, 4250, 3250, 2750}, {39448, 39508, 39751, 39859, 39904})
	// Excel: ~0.3734 (37.34%)
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000), NumberVal(2750), NumberVal(4250), NumberVal(3250), NumberVal(2750)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39508), NumberVal(39751), NumberVal(39859), NumberVal(39904)},
		},
	}
	v, err := fnXIRR([]Value{vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "XIRR basic", v, 0.3734)
}

func TestXIRR_DateStrings(t *testing.T) {
	// XIRR with date strings instead of serial numbers (fa.xlsx scenario).
	// Cash flows: -245000, 8196.57, 18829.92, 27082.48, 19123.73, 51711
	// Dates: 10/17/2017, 04/29/2022, 08/02/2024, 03/31/2025, 07/14/2025, 03/04/2026
	// Excel: ~-0.08442739386111497
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-245000), NumberVal(8196.57), NumberVal(18829.92), NumberVal(27082.48), NumberVal(19123.73), NumberVal(51711)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("10/17/2017"), StringVal("04/29/2022"), StringVal("08/02/2024"), StringVal("03/31/2025"), StringVal("07/14/2025"), StringVal("03/04/2026")},
		},
	}
	v, err := fnXIRR([]Value{vals, dates, NumberVal(-0.01)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XIRR date strings: expected number, got type %v (str=%q)", v.Type, v.Str)
	}
	if math.Abs(v.Num-(-0.08442739386111497)) > 0.001 {
		t.Errorf("XIRR date strings: got %f, want ~-0.0844", v.Num)
	}
}

// === DB ===

func TestDB_ExcelExample_Period1(t *testing.T) {
	// DB(1000000, 100000, 6, 1, 7) = 186083.33
	v, err := fnDB(numArgs(1000000, 100000, 6, 1, 7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB period 1", v, 186083.33)
}

func TestDB_ExcelExample_Period2(t *testing.T) {
	// DB(1000000, 100000, 6, 2, 7) = 259639.42
	v, err := fnDB(numArgs(1000000, 100000, 6, 2, 7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB period 2", v, 259639.42)
}

func TestDB_ExcelExample_Period3(t *testing.T) {
	// DB(1000000, 100000, 6, 3, 7) = 176814.44
	v, err := fnDB(numArgs(1000000, 100000, 6, 3, 7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB period 3", v, 176814.44)
}

func TestDB_ExcelExample_Period4(t *testing.T) {
	// DB(1000000, 100000, 6, 4, 7) = 120410.64
	v, err := fnDB(numArgs(1000000, 100000, 6, 4, 7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB period 4", v, 120410.64)
}

func TestDB_ExcelExample_Period5(t *testing.T) {
	// DB(1000000, 100000, 6, 5, 7) = 81999.64
	v, err := fnDB(numArgs(1000000, 100000, 6, 5, 7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB period 5", v, 81999.64)
}

func TestDB_ExcelExample_Period6(t *testing.T) {
	// DB(1000000, 100000, 6, 6, 7) = 55841.76
	v, err := fnDB(numArgs(1000000, 100000, 6, 6, 7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB period 6", v, 55841.76)
}

func TestDB_ExcelExample_Period7_LastFractional(t *testing.T) {
	// DB(1000000, 100000, 6, 7, 7) = 15845.10
	v, err := fnDB(numArgs(1000000, 100000, 6, 7, 7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB period 7 (last fractional)", v, 15845.10)
}

func TestDB_FullYear_DefaultMonth(t *testing.T) {
	// DB(1000000, 100000, 6, 1) — month defaults to 12
	// rate = 1 - (100000/1000000)^(1/6) = 1 - 0.1^(1/6) ≈ 0.319
	// period 1: 1000000 * 0.319 * 12/12 = 319000
	v, err := fnDB(numArgs(1000000, 100000, 6, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB default month", v, 319000)
}

func TestDB_FullYear_Period2(t *testing.T) {
	// DB(1000000, 100000, 6, 2) — month defaults to 12
	// period 2: (1000000 - 319000) * 0.319 = 681000 * 0.319 = 217239
	v, err := fnDB(numArgs(1000000, 100000, 6, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB default month period 2", v, 217239)
}

func TestDB_Life1(t *testing.T) {
	// DB(10000, 1000, 1, 1) — life of 1 year, month=12
	// rate = 1 - (1000/10000)^(1/1) = 1 - 0.1 = 0.9
	// period 1: 10000 * 0.9 * 12/12 = 9000
	v, err := fnDB(numArgs(10000, 1000, 1, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB life=1", v, 9000)
}

func TestDB_Life1_Month6(t *testing.T) {
	// DB(10000, 1000, 1, 1, 6) — life of 1, month=6
	// rate = 0.9, period 1: 10000 * 0.9 * 6/12 = 4500
	v, err := fnDB(numArgs(10000, 1000, 1, 1, 6))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB life=1 month=6 period 1", v, 4500)
}

func TestDB_Life1_Month6_Period2(t *testing.T) {
	// DB(10000, 1000, 1, 2, 6) — last fractional period
	// rate = 0.9, period 1 dep = 4500, remaining = 5500
	// period 2: 5500 * 0.9 * (12-6)/12 = 5500 * 0.9 * 0.5 = 2475
	v, err := fnDB(numArgs(10000, 1000, 1, 2, 6))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB life=1 month=6 period 2", v, 2475)
}

func TestDB_CostZero(t *testing.T) {
	// DB(0, 0, 5, 1) — cost is 0, depreciation is 0
	v, err := fnDB(numArgs(0, 0, 5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB cost=0", v, 0)
}

func TestDB_SalvageEqualsCost(t *testing.T) {
	// DB(10000, 10000, 5, 1) — no depreciation when salvage = cost
	v, err := fnDB(numArgs(10000, 10000, 5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB salvage=cost", v, 0)
}

func TestDB_SalvageZero(t *testing.T) {
	// DB(10000, 0, 5, 1) — salvage = 0
	// rate = 1 - (0/10000)^(1/5) = 1 - 0 = 1.0
	// period 1: 10000 * 1.0 * 12/12 = 10000
	v, err := fnDB(numArgs(10000, 0, 5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB salvage=0", v, 10000)
}

func TestDB_SalvageZero_Period2(t *testing.T) {
	// DB(10000, 0, 5, 2) — all depreciated in period 1
	// period 2: (10000 - 10000) * 1.0 = 0
	v, err := fnDB(numArgs(10000, 0, 5, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB salvage=0 period 2", v, 0)
}

func TestDB_LargeNumbers(t *testing.T) {
	// DB(50000000, 5000000, 10, 1) — large asset
	// rate = 1 - (5000000/50000000)^(1/10) = 1 - 0.1^0.1 ≈ 0.206
	// period 1: 50000000 * 0.206 = 10300000
	v, err := fnDB(numArgs(50000000, 5000000, 10, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB large numbers", v, 10300000)
}

func TestDB_SmallAsset(t *testing.T) {
	// DB(100, 10, 5, 1)
	// rate = 1 - (10/100)^(1/5) = 1 - 0.631 = 0.369
	// period 1: 100 * 0.369 = 36.9
	v, err := fnDB(numArgs(100, 10, 5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB small asset", v, 36.90)
}

func TestDB_ErrorTooFewArgs(t *testing.T) {
	v, _ := fnDB(numArgs(1000, 100, 5))
	assertError(t, "DB too few args", v)
}

func TestDB_ErrorTooManyArgs(t *testing.T) {
	v, _ := fnDB(numArgs(1000, 100, 5, 1, 12, 99))
	assertError(t, "DB too many args", v)
}

func TestDB_ErrorNegativeCost(t *testing.T) {
	v, _ := fnDB(numArgs(-1000, 100, 5, 1))
	assertError(t, "DB negative cost", v)
}

func TestDB_ErrorNegativeSalvage(t *testing.T) {
	v, _ := fnDB(numArgs(1000, -100, 5, 1))
	assertError(t, "DB negative salvage", v)
}

func TestDB_ErrorLifeZero(t *testing.T) {
	v, _ := fnDB(numArgs(1000, 100, 0, 1))
	assertError(t, "DB life=0", v)
}

func TestDB_ErrorPeriodZero(t *testing.T) {
	v, _ := fnDB(numArgs(1000, 100, 5, 0))
	assertError(t, "DB period=0", v)
}

func TestDB_ErrorPeriodExceedsLife(t *testing.T) {
	// month=12 so max period = life
	v, _ := fnDB(numArgs(1000, 100, 5, 6))
	assertError(t, "DB period > life (month=12)", v)
}

func TestDB_ErrorPeriodExceedsLifePlus1(t *testing.T) {
	// month=6 so max period = life+1 = 6, period 7 should error
	v, _ := fnDB(numArgs(1000, 100, 5, 7, 6))
	assertError(t, "DB period > life+1", v)
}

func TestDB_ErrorMonthZero(t *testing.T) {
	v, _ := fnDB(numArgs(1000, 100, 5, 1, 0))
	assertError(t, "DB month=0", v)
}

func TestDB_ErrorMonthOver12(t *testing.T) {
	v, _ := fnDB(numArgs(1000, 100, 5, 1, 13))
	assertError(t, "DB month>12", v)
}

func TestDB_Month1(t *testing.T) {
	// DB(1000000, 100000, 6, 1, 1)
	// rate = 0.319
	// period 1: 1000000 * 0.319 * 1/12 = 26583.33
	v, err := fnDB(numArgs(1000000, 100000, 6, 1, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB month=1", v, 26583.33)
}

// === DDB ===

func TestDDB_ExcelExample_FirstDayDepreciation(t *testing.T) {
	// DDB(2400, 300, 10*365, 1) = 1.32 (first day's depreciation)
	v, err := fnDDB(numArgs(2400, 300, 10*365, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB first day", v, 1.32)
}

func TestDDB_ExcelExample_FirstMonthDepreciation(t *testing.T) {
	// DDB(2400, 300, 10*12, 1, 2) = 40.00
	v, err := fnDDB(numArgs(2400, 300, 10*12, 1, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB first month", v, 40.00)
}

func TestDDB_ExcelExample_FirstYearDepreciation(t *testing.T) {
	// DDB(2400, 300, 10, 1, 2) = 480.00
	v, err := fnDDB(numArgs(2400, 300, 10, 1, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB first year", v, 480.00)
}

func TestDDB_ExcelExample_SecondYearFactor15(t *testing.T) {
	// DDB(2400, 300, 10, 2, 1.5) = 306.00
	v, err := fnDDB(numArgs(2400, 300, 10, 2, 1.5))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB second year factor 1.5", v, 306.00)
}

func TestDDB_ExcelExample_TenthYear(t *testing.T) {
	// DDB(2400, 300, 10, 10) = 22.12
	v, err := fnDDB(numArgs(2400, 300, 10, 10))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB tenth year", v, 22.12)
}

func TestDDB_Period2_DefaultFactor(t *testing.T) {
	// DDB(2400, 300, 10, 2)
	// period 1: 2400 * 2/10 = 480, book = 1920
	// period 2: 1920 * 2/10 = 384
	v, err := fnDDB(numArgs(2400, 300, 10, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB period 2", v, 384.00)
}

func TestDDB_Period3_DefaultFactor(t *testing.T) {
	// period 3: (2400-480-384) * 0.2 = 1536 * 0.2 = 307.20
	v, err := fnDDB(numArgs(2400, 300, 10, 3))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB period 3", v, 307.20)
}

func TestDDB_CostZero(t *testing.T) {
	v, err := fnDDB(numArgs(0, 0, 5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB cost=0", v, 0)
}

func TestDDB_SalvageEqualsCost(t *testing.T) {
	v, err := fnDDB(numArgs(10000, 10000, 5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB salvage=cost", v, 0)
}

func TestDDB_SalvageZero(t *testing.T) {
	// DDB(10000, 0, 5, 1) = 10000 * 2/5 = 4000
	v, err := fnDDB(numArgs(10000, 0, 5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB salvage=0", v, 4000)
}

func TestDDB_SalvageZero_Period2(t *testing.T) {
	// DDB(10000, 0, 5, 2) = (10000-4000) * 2/5 = 2400
	v, err := fnDDB(numArgs(10000, 0, 5, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB salvage=0 period 2", v, 2400)
}

func TestDDB_LargeAsset(t *testing.T) {
	// DDB(1000000, 100000, 10, 1) = 1000000 * 2/10 = 200000
	v, err := fnDDB(numArgs(1000000, 100000, 10, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB large asset", v, 200000)
}

func TestDDB_Factor1(t *testing.T) {
	// DDB(2400, 300, 10, 1, 1) = 2400 * 1/10 = 240
	v, err := fnDDB(numArgs(2400, 300, 10, 1, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB factor=1", v, 240)
}

func TestDDB_Factor3(t *testing.T) {
	// DDB(2400, 300, 10, 1, 3) = 2400 * 3/10 = 720
	v, err := fnDDB(numArgs(2400, 300, 10, 1, 3))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB factor=3", v, 720)
}

func TestDDB_DepreciationCappedBySalvage(t *testing.T) {
	// DDB(1000, 800, 5, 1) = min(1000 * 2/5, 1000 - 800) = min(400, 200) = 200
	v, err := fnDDB(numArgs(1000, 800, 5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB capped by salvage", v, 200)
}

func TestDDB_DepreciationCappedBySalvage_Period2(t *testing.T) {
	// After period 1 (dep=200), book = 800 = salvage, so period 2 dep = 0
	v, err := fnDDB(numArgs(1000, 800, 5, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB capped period 2", v, 0)
}

func TestDDB_Life1(t *testing.T) {
	// DDB(10000, 1000, 1, 1) = min(10000 * 2/1, 10000 - 1000) = min(20000, 9000) = 9000
	v, err := fnDDB(numArgs(10000, 1000, 1, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB life=1", v, 9000)
}

func TestDDB_SmallAsset(t *testing.T) {
	// DDB(100, 10, 5, 1) = 100 * 2/5 = 40
	v, err := fnDDB(numArgs(100, 10, 5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB small asset", v, 40)
}

func TestDDB_ErrorTooFewArgs(t *testing.T) {
	v, _ := fnDDB(numArgs(1000, 100, 5))
	assertError(t, "DDB too few args", v)
}

func TestDDB_ErrorTooManyArgs(t *testing.T) {
	v, _ := fnDDB(numArgs(1000, 100, 5, 1, 2, 99))
	assertError(t, "DDB too many args", v)
}

func TestDDB_ErrorNegativeCost(t *testing.T) {
	v, _ := fnDDB(numArgs(-1000, 100, 5, 1))
	assertError(t, "DDB negative cost", v)
}

func TestDDB_ErrorNegativeSalvage(t *testing.T) {
	v, _ := fnDDB(numArgs(1000, -100, 5, 1))
	assertError(t, "DDB negative salvage", v)
}

func TestDDB_ErrorLifeZero(t *testing.T) {
	v, _ := fnDDB(numArgs(1000, 100, 0, 1))
	assertError(t, "DDB life=0", v)
}

func TestDDB_ErrorPeriodZero(t *testing.T) {
	v, _ := fnDDB(numArgs(1000, 100, 5, 0))
	assertError(t, "DDB period=0", v)
}

func TestDDB_ErrorPeriodExceedsLife(t *testing.T) {
	v, _ := fnDDB(numArgs(1000, 100, 5, 6))
	assertError(t, "DDB period > life", v)
}

func TestDDB_ErrorNegativeFactor(t *testing.T) {
	v, _ := fnDDB(numArgs(1000, 100, 5, 1, -2))
	assertError(t, "DDB negative factor", v)
}

func TestDDB_ErrorFactorZero(t *testing.T) {
	v, _ := fnDDB(numArgs(1000, 100, 5, 1, 0))
	assertError(t, "DDB factor=0", v)
}

// === DOLLARDE ===

func TestDOLLARDE_ExcelExample1(t *testing.T) {
	// DOLLARDE(1.02, 16) = 1.125
	v, err := fnDOLLARDE(numArgs(1.02, 16))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARDE(1.02,16)", v, 1.125)
}

func TestDOLLARDE_ExcelExample2(t *testing.T) {
	// DOLLARDE(1.1, 32) = 1.3125
	v, err := fnDOLLARDE(numArgs(1.1, 32))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARDE(1.1,32)", v, 1.3125)
}

func TestDOLLARDE_WholeNumber(t *testing.T) {
	// DOLLARDE(5.0, 8) = 5.0
	v, err := fnDOLLARDE(numArgs(5.0, 8))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARDE(5.0,8)", v, 5.0)
}

func TestDOLLARDE_Fraction8(t *testing.T) {
	// DOLLARDE(1.4, 8) = 1.5 (4/8 = 0.5)
	v, err := fnDOLLARDE(numArgs(1.4, 8))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARDE(1.4,8)", v, 1.5)
}

func TestDOLLARDE_Fraction4(t *testing.T) {
	// DOLLARDE(1.1, 4) = 1.25 (1/4 = 0.25)
	v, err := fnDOLLARDE(numArgs(1.1, 4))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARDE(1.1,4)", v, 1.25)
}

func TestDOLLARDE_Fraction2(t *testing.T) {
	// DOLLARDE(1.1, 2) = 1.5 (1/2 = 0.5)
	v, err := fnDOLLARDE(numArgs(1.1, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARDE(1.1,2)", v, 1.5)
}

func TestDOLLARDE_Negative(t *testing.T) {
	// DOLLARDE(-1.02, 16) = -1.125
	v, err := fnDOLLARDE(numArgs(-1.02, 16))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARDE(-1.02,16)", v, -1.125)
}

func TestDOLLARDE_Zero(t *testing.T) {
	// DOLLARDE(0, 16) = 0
	v, err := fnDOLLARDE(numArgs(0, 16))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARDE(0,16)", v, 0)
}

func TestDOLLARDE_FractionTruncated(t *testing.T) {
	// DOLLARDE(1.02, 16.9) should truncate fraction to 16 => 1.125
	v, err := fnDOLLARDE(numArgs(1.02, 16.9))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARDE(1.02,16.9)", v, 1.125)
}

func TestDOLLARDE_Fraction1(t *testing.T) {
	// DOLLARDE(1.5, 1) = 1 + 5/1 = 6
	v, err := fnDOLLARDE(numArgs(1.5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARDE(1.5,1)", v, 6.0)
}

func TestDOLLARDE_LargeFraction(t *testing.T) {
	// DOLLARDE(1.001, 100) = 1 + 1/100 = 1.01
	v, err := fnDOLLARDE(numArgs(1.001, 100))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARDE(1.001,100)", v, 1.01)
}

func TestDOLLARDE_ErrorFractionNegative(t *testing.T) {
	v, _ := fnDOLLARDE(numArgs(1.02, -1))
	assertError(t, "DOLLARDE fraction<0", v)
}

func TestDOLLARDE_ErrorFractionZero(t *testing.T) {
	v, _ := fnDOLLARDE(numArgs(1.02, 0))
	assertError(t, "DOLLARDE fraction=0", v)
}

func TestDOLLARDE_ErrorFractionBetween0And1(t *testing.T) {
	// Fraction 0.5 truncates to 0 => #DIV/0!
	v, _ := fnDOLLARDE(numArgs(1.02, 0.5))
	assertError(t, "DOLLARDE fraction=0.5", v)
}

func TestDOLLARDE_ErrorTooFewArgs(t *testing.T) {
	v, _ := fnDOLLARDE(numArgs(1.02))
	assertError(t, "DOLLARDE too few args", v)
}

func TestDOLLARDE_ErrorTooManyArgs(t *testing.T) {
	v, _ := fnDOLLARDE(numArgs(1.02, 16, 99))
	assertError(t, "DOLLARDE too many args", v)
}

func TestDOLLARDE_NegativeWithFraction8(t *testing.T) {
	// DOLLARDE(-1.4, 8) = -1.5
	v, err := fnDOLLARDE(numArgs(-1.4, 8))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARDE(-1.4,8)", v, -1.5)
}

// === DOLLARFR ===

func TestDOLLARFR_ExcelExample1(t *testing.T) {
	// DOLLARFR(1.125, 16) = 1.02
	v, err := fnDOLLARFR(numArgs(1.125, 16))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARFR(1.125,16)", v, 1.02)
}

func TestDOLLARFR_ExcelExample2(t *testing.T) {
	// DOLLARFR(1.125, 32) = 1.04
	v, err := fnDOLLARFR(numArgs(1.125, 32))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARFR(1.125,32)", v, 1.04)
}

func TestDOLLARFR_WholeNumber(t *testing.T) {
	// DOLLARFR(5.0, 8) = 5.0
	v, err := fnDOLLARFR(numArgs(5.0, 8))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARFR(5.0,8)", v, 5.0)
}

func TestDOLLARFR_Fraction8(t *testing.T) {
	// DOLLARFR(1.5, 8) = 1.4 (0.5 * 8 = 4, placed as .4 since 8 is 1 digit)
	v, err := fnDOLLARFR(numArgs(1.5, 8))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARFR(1.5,8)", v, 1.4)
}

func TestDOLLARFR_Fraction4(t *testing.T) {
	// DOLLARFR(1.25, 4) = 1.1 (0.25 * 4 = 1, placed as .1)
	v, err := fnDOLLARFR(numArgs(1.25, 4))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARFR(1.25,4)", v, 1.1)
}

func TestDOLLARFR_Fraction2(t *testing.T) {
	// DOLLARFR(1.5, 2) = 1.1 (0.5 * 2 = 1, placed as .1)
	v, err := fnDOLLARFR(numArgs(1.5, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARFR(1.5,2)", v, 1.1)
}

func TestDOLLARFR_Negative(t *testing.T) {
	// DOLLARFR(-1.125, 16) = -1.02
	v, err := fnDOLLARFR(numArgs(-1.125, 16))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARFR(-1.125,16)", v, -1.02)
}

func TestDOLLARFR_Zero(t *testing.T) {
	// DOLLARFR(0, 16) = 0
	v, err := fnDOLLARFR(numArgs(0, 16))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARFR(0,16)", v, 0)
}

func TestDOLLARFR_FractionTruncated(t *testing.T) {
	// DOLLARFR(1.125, 16.9) should truncate to 16 => 1.02
	v, err := fnDOLLARFR(numArgs(1.125, 16.9))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARFR(1.125,16.9)", v, 1.02)
}

func TestDOLLARFR_Fraction1(t *testing.T) {
	// DOLLARFR(6.0, 1) = 1 + ? ... actually 6.0 integer=6, frac=0, so 6.0
	v, err := fnDOLLARFR(numArgs(6.0, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARFR(6.0,1)", v, 6.0)
}

func TestDOLLARFR_LargeFraction(t *testing.T) {
	// DOLLARFR(1.01, 100) = 1.001 (0.01 * 100 = 1, placed as .001)
	v, err := fnDOLLARFR(numArgs(1.01, 100))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARFR(1.01,100)", v, 1.001)
}

func TestDOLLARFR_ErrorFractionNegative(t *testing.T) {
	v, _ := fnDOLLARFR(numArgs(1.125, -1))
	assertError(t, "DOLLARFR fraction<0", v)
}

func TestDOLLARFR_ErrorFractionZero(t *testing.T) {
	v, _ := fnDOLLARFR(numArgs(1.125, 0))
	assertError(t, "DOLLARFR fraction=0", v)
}

func TestDOLLARFR_ErrorFractionBetween0And1(t *testing.T) {
	// Fraction 0.5 truncates to 0 => #DIV/0!
	v, _ := fnDOLLARFR(numArgs(1.125, 0.5))
	assertError(t, "DOLLARFR fraction=0.5", v)
}

func TestDOLLARFR_ErrorTooFewArgs(t *testing.T) {
	v, _ := fnDOLLARFR(numArgs(1.125))
	assertError(t, "DOLLARFR too few args", v)
}

func TestDOLLARFR_ErrorTooManyArgs(t *testing.T) {
	v, _ := fnDOLLARFR(numArgs(1.125, 16, 99))
	assertError(t, "DOLLARFR too many args", v)
}

func TestDOLLARFR_NegativeWithFraction8(t *testing.T) {
	// DOLLARFR(-1.5, 8) = -1.4
	v, err := fnDOLLARFR(numArgs(-1.5, 8))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARFR(-1.5,8)", v, -1.4)
}

func TestDOLLARDE_DOLLARFR_RoundTrip(t *testing.T) {
	// DOLLARDE(DOLLARFR(1.125, 16), 16) should equal 1.125
	fr, err := fnDOLLARFR(numArgs(1.125, 16))
	if err != nil {
		t.Fatal(err)
	}
	de, err := fnDOLLARDE(numArgs(fr.Num, 16))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "round-trip DOLLARDE(DOLLARFR(1.125,16),16)", de, 1.125)
}

// === EFFECT ===

func TestEFFECT_ExcelExample(t *testing.T) {
	// EFFECT(0.0525, 4) = 0.0535427
	v, err := fnEFFECT(numArgs(0.0525, 4))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "EFFECT excel example", v, 0.0535427)
}

func TestEFFECT_Monthly(t *testing.T) {
	// EFFECT(0.12, 12) = (1+0.01)^12 - 1 = 0.126825
	v, err := fnEFFECT(numArgs(0.12, 12))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "EFFECT monthly", v, 0.126825)
}

func TestEFFECT_Annual(t *testing.T) {
	// EFFECT(0.10, 1) = 0.10 (compounding once = nominal)
	v, err := fnEFFECT(numArgs(0.10, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "EFFECT annual", v, 0.10)
}

func TestEFFECT_SemiAnnual(t *testing.T) {
	// EFFECT(0.10, 2) = (1+0.05)^2 - 1 = 0.1025
	v, err := fnEFFECT(numArgs(0.10, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "EFFECT semi-annual", v, 0.1025)
}

func TestEFFECT_Quarterly(t *testing.T) {
	// EFFECT(0.08, 4) = (1+0.02)^4 - 1 = 0.08243216
	v, err := fnEFFECT(numArgs(0.08, 4))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "EFFECT quarterly", v, 0.08243216)
}

func TestEFFECT_Daily(t *testing.T) {
	// EFFECT(0.10, 365) ≈ 0.10515578
	v, err := fnEFFECT(numArgs(0.10, 365))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "EFFECT daily", v, 0.10515578)
}

func TestEFFECT_Weekly(t *testing.T) {
	// EFFECT(0.10, 52) ≈ 0.10506479
	v, err := fnEFFECT(numArgs(0.10, 52))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "EFFECT weekly", v, 0.10506479)
}

func TestEFFECT_HighRate(t *testing.T) {
	// EFFECT(1.0, 12) = (1+1/12)^12 - 1 ≈ 1.61303529
	v, err := fnEFFECT(numArgs(1.0, 12))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("EFFECT high rate: expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-1.61303529) > 0.01 {
		t.Errorf("EFFECT high rate: got %f, want 1.61303529", v.Num)
	}
}

func TestEFFECT_SmallRate(t *testing.T) {
	// EFFECT(0.001, 12) ≈ 0.001000416
	v, err := fnEFFECT(numArgs(0.001, 12))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("EFFECT small rate: expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-0.001000416) > 0.000001 {
		t.Errorf("EFFECT small rate: got %f, want ~0.001000416", v.Num)
	}
}

func TestEFFECT_NperyTruncated(t *testing.T) {
	// EFFECT(0.10, 4.9) should truncate npery to 4
	v1, _ := fnEFFECT(numArgs(0.10, 4))
	v2, _ := fnEFFECT(numArgs(0.10, 4.9))
	if v1.Num != v2.Num {
		t.Errorf("EFFECT npery truncation: EFFECT(0.10,4)=%f != EFFECT(0.10,4.9)=%f", v1.Num, v2.Num)
	}
}

func TestEFFECT_LargeNpery(t *testing.T) {
	// EFFECT(0.05, 1000) approaches e^0.05 - 1 ≈ 0.05127
	v, err := fnEFFECT(numArgs(0.05, 1000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "EFFECT large npery", v, 0.05127)
}

func TestEFFECT_ErrorZeroRate(t *testing.T) {
	v, _ := fnEFFECT(numArgs(0, 4))
	assertError(t, "EFFECT rate=0", v)
}

func TestEFFECT_ErrorNegativeRate(t *testing.T) {
	v, _ := fnEFFECT(numArgs(-0.05, 4))
	assertError(t, "EFFECT rate<0", v)
}

func TestEFFECT_ErrorNperyZero(t *testing.T) {
	v, _ := fnEFFECT(numArgs(0.10, 0))
	assertError(t, "EFFECT npery=0", v)
}

func TestEFFECT_ErrorNperyFraction(t *testing.T) {
	// npery=0.5 truncates to 0 => #NUM!
	v, _ := fnEFFECT(numArgs(0.10, 0.5))
	assertError(t, "EFFECT npery=0.5", v)
}

func TestEFFECT_ErrorTooFewArgs(t *testing.T) {
	v, _ := fnEFFECT(numArgs(0.10))
	assertError(t, "EFFECT too few args", v)
}

func TestEFFECT_ErrorTooManyArgs(t *testing.T) {
	v, _ := fnEFFECT(numArgs(0.10, 4, 1))
	assertError(t, "EFFECT too many args", v)
}

// === NOMINAL ===

func TestNOMINAL_ExcelExample(t *testing.T) {
	// NOMINAL(0.053543, 4) ≈ 0.0525003
	v, err := fnNOMINAL(numArgs(0.053543, 4))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NOMINAL excel example", v, 0.0525003)
}

func TestNOMINAL_Monthly(t *testing.T) {
	// NOMINAL(0.126825, 12) ≈ 0.12
	v, err := fnNOMINAL(numArgs(0.126825, 12))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NOMINAL monthly", v, 0.12)
}

func TestNOMINAL_Annual(t *testing.T) {
	// NOMINAL(0.10, 1) = 0.10
	v, err := fnNOMINAL(numArgs(0.10, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NOMINAL annual", v, 0.10)
}

func TestNOMINAL_SemiAnnual(t *testing.T) {
	// NOMINAL(0.1025, 2) ≈ 0.10
	v, err := fnNOMINAL(numArgs(0.1025, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NOMINAL semi-annual", v, 0.10)
}

func TestNOMINAL_Quarterly(t *testing.T) {
	// NOMINAL(0.08243216, 4) ≈ 0.08
	v, err := fnNOMINAL(numArgs(0.08243216, 4))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NOMINAL quarterly", v, 0.08)
}

func TestNOMINAL_Daily(t *testing.T) {
	// NOMINAL(0.10515578, 365) ≈ 0.10
	v, err := fnNOMINAL(numArgs(0.10515578, 365))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NOMINAL daily", v, 0.10)
}

func TestNOMINAL_Weekly(t *testing.T) {
	// NOMINAL(0.10506479, 52) ≈ 0.10
	v, err := fnNOMINAL(numArgs(0.10506479, 52))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NOMINAL weekly", v, 0.10)
}

func TestNOMINAL_HighRate(t *testing.T) {
	// NOMINAL(1.0, 12) — high effective rate
	v, err := fnNOMINAL(numArgs(1.0, 12))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("NOMINAL high rate: expected number, got %v", v.Type)
	}
	// Verify round-trip: EFFECT(NOMINAL(1.0, 12), 12) ≈ 1.0
	eff, _ := fnEFFECT(numArgs(v.Num, 12))
	if math.Abs(eff.Num-1.0) > 0.01 {
		t.Errorf("NOMINAL high rate round-trip: EFFECT(%f,12)=%f, want 1.0", v.Num, eff.Num)
	}
}

func TestNOMINAL_SmallRate(t *testing.T) {
	// NOMINAL(0.001, 12) ≈ 0.000999584
	v, err := fnNOMINAL(numArgs(0.001, 12))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("NOMINAL small rate: expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-0.000999584) > 0.000001 {
		t.Errorf("NOMINAL small rate: got %f, want ~0.000999584", v.Num)
	}
}

func TestNOMINAL_NperyTruncated(t *testing.T) {
	// NOMINAL(0.10, 4.9) should truncate npery to 4
	v1, _ := fnNOMINAL(numArgs(0.10, 4))
	v2, _ := fnNOMINAL(numArgs(0.10, 4.9))
	if v1.Num != v2.Num {
		t.Errorf("NOMINAL npery truncation: NOMINAL(0.10,4)=%f != NOMINAL(0.10,4.9)=%f", v1.Num, v2.Num)
	}
}

func TestNOMINAL_LargeNpery(t *testing.T) {
	// NOMINAL(0.05127, 1000) ≈ 0.05
	v, err := fnNOMINAL(numArgs(0.05127, 1000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NOMINAL large npery", v, 0.05)
}

func TestNOMINAL_ErrorZeroRate(t *testing.T) {
	v, _ := fnNOMINAL(numArgs(0, 4))
	assertError(t, "NOMINAL rate=0", v)
}

func TestNOMINAL_ErrorNegativeRate(t *testing.T) {
	v, _ := fnNOMINAL(numArgs(-0.05, 4))
	assertError(t, "NOMINAL rate<0", v)
}

func TestNOMINAL_ErrorNperyZero(t *testing.T) {
	v, _ := fnNOMINAL(numArgs(0.10, 0))
	assertError(t, "NOMINAL npery=0", v)
}

func TestNOMINAL_ErrorNperyFraction(t *testing.T) {
	// npery=0.5 truncates to 0 => #NUM!
	v, _ := fnNOMINAL(numArgs(0.10, 0.5))
	assertError(t, "NOMINAL npery=0.5", v)
}

func TestNOMINAL_ErrorTooFewArgs(t *testing.T) {
	v, _ := fnNOMINAL(numArgs(0.10))
	assertError(t, "NOMINAL too few args", v)
}

func TestNOMINAL_ErrorTooManyArgs(t *testing.T) {
	v, _ := fnNOMINAL(numArgs(0.10, 4, 1))
	assertError(t, "NOMINAL too many args", v)
}

func TestEFFECT_NOMINAL_RoundTrip(t *testing.T) {
	// NOMINAL(EFFECT(0.08, 4), 4) should give back 0.08
	eff, _ := fnEFFECT(numArgs(0.08, 4))
	nom, _ := fnNOMINAL(numArgs(eff.Num, 4))
	assertClose(t, "EFFECT->NOMINAL round-trip", nom, 0.08)

	// EFFECT(NOMINAL(0.10, 12), 12) should give back 0.10
	nom2, _ := fnNOMINAL(numArgs(0.10, 12))
	eff2, _ := fnEFFECT(numArgs(nom2.Num, 12))
	assertClose(t, "NOMINAL->EFFECT round-trip", eff2, 0.10)
}

func TestXIRR_NegativeRate(t *testing.T) {
	// XIRR with guess parameter and negative expected rate.
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-245000), NumberVal(8196.57), NumberVal(18829.92), NumberVal(27082.48), NumberVal(19123.73), NumberVal(51711)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(43025), NumberVal(44680), NumberVal(45505), NumberVal(45747), NumberVal(45852), NumberVal(46080)},
		},
	}
	v, err := fnXIRR([]Value{vals, dates, NumberVal(-0.01)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XIRR negative rate: expected number, got type %v (str=%q)", v.Type, v.Str)
	}
	if v.Num >= 0 {
		t.Errorf("XIRR negative rate: expected negative rate, got %f", v.Num)
	}
}
