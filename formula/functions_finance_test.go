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
