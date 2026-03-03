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
