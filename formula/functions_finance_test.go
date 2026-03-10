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
	// Expected: PMT(0.05/12, 360, 200000) = -1073.64
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
	// Expected: -546.61
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

func TestPMT_Comprehensive(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// --- Doc examples ---
		{
			name: "doc: monthly payment 8%/12, 10 months, $10000",
			args: numArgs(0.08/12, 10, 10000),
			want: -1037.03,
		},
		{
			name: "doc: payment at beginning of period (type=1)",
			args: numArgs(0.08/12, 10, 10000, 0, 1),
			want: -1030.16,
		},
		{
			name: "doc: saving for $50k over 18 years at 6%",
			args: numArgs(0.06/12, 18*12, 0, 50000),
			want: -129.08,
		},

		// --- Car loan ---
		{
			name: "5-year car loan at 6.5%",
			args: numArgs(0.065/12, 60, 25000),
			want: -489.15,
		},

		// --- 15-year mortgage ---
		{
			name: "15-year mortgage at 4%",
			args: numArgs(0.04/12, 180, 300000),
			want: -2219.06,
		},

		// --- Zero interest rate with fv ---
		{
			name: "zero rate with future value",
			args: numArgs(0, 12, 0, 6000),
			want: -500,
		},
		{
			name: "zero rate with pv and fv",
			args: numArgs(0, 24, 1000, 2000),
			want: -125,
		},

		// --- Future value: pv and fv combined ---
		{
			name: "pv and fv combined",
			args: numArgs(0.06/12, 60, 10000, 5000),
			want: -264.99,
		},

		// --- Type=0 vs type=1 ---
		{
			name: "type=0 (end of period)",
			args: numArgs(0.10/12, 120, 50000, 0, 0),
			want: -660.75,
		},
		{
			name: "type=1 (beginning of period)",
			args: numArgs(0.10/12, 120, 50000, 0, 1),
			want: -655.29,
		},

		// --- Negative pv (lender perspective) ---
		{
			name: "negative pv gives positive payment",
			args: numArgs(0.05/12, 360, -200000),
			want: 1073.64,
		},

		// --- Large nper ---
		{
			name: "large nper: 600 months (50 years)",
			args: numArgs(0.03/12, 600, 100000),
			want: -321.98,
		},

		// --- All zeros except nper ---
		{
			name: "all zeros: pv=0, fv=0, rate=0",
			args: numArgs(0, 10, 0),
			want: 0,
		},

		// --- Only pv, no fv ---
		{
			name: "only pv, fv defaults to 0",
			args: numArgs(0.12/12, 48, 20000),
			want: -526.68,
		},

		// --- Only fv, no pv ---
		{
			name: "only fv, pv=0",
			args: numArgs(0.05/12, 60, 0, 20000),
			want: -294.09,
		},

		// --- String coercion ---
		{
			name: "string coercion for rate",
			args: []Value{StringVal("0.05"), NumberVal(12), NumberVal(1200)},
			want: -135.39,
		},
		{
			name: "string coercion for all numeric strings",
			args: []Value{StringVal("0"), StringVal("10"), StringVal("1000")},
			want: -100,
		},

		// --- Savings with type=1 and fv ---
		{
			name: "savings: type=1 with fv",
			args: numArgs(0.06/12, 60, 0, 10000, 1),
			want: -142.61,
		},

		// --- High interest rate ---
		{
			name: "high annual rate 24%",
			args: numArgs(0.24/12, 36, 5000),
			want: -196.16,
		},

		// --- Error: too few arguments ---
		{
			name:    "error: too few args (2)",
			args:    numArgs(0.05, 12),
			wantErr: true,
		},
		{
			name:    "error: too few args (0)",
			args:    []Value{},
			wantErr: true,
		},

		// --- Error: too many arguments ---
		{
			name:    "error: too many args (6)",
			args:    numArgs(0.05, 12, 1000, 0, 0, 0),
			wantErr: true,
		},

		// --- Error: non-numeric ---
		{
			name:    "error: non-numeric rate",
			args:    []Value{StringVal("abc"), NumberVal(12), NumberVal(1000)},
			wantErr: true,
		},
		{
			name:    "error: non-numeric nper",
			args:    []Value{NumberVal(0.05), StringVal("xyz"), NumberVal(1000)},
			wantErr: true,
		},
		{
			name:    "error: non-numeric pv",
			args:    []Value{NumberVal(0.05), NumberVal(12), StringVal("bad")},
			wantErr: true,
		},
		{
			name:    "error: non-numeric fv",
			args:    []Value{NumberVal(0.05), NumberVal(12), NumberVal(1000), StringVal("no")},
			wantErr: true,
		},
		{
			name:    "error: non-numeric type",
			args:    []Value{NumberVal(0.05), NumberVal(12), NumberVal(1000), NumberVal(0), StringVal("x")},
			wantErr: true,
		},

		// --- Error propagation ---
		{
			name:    "error propagation in rate",
			args:    []Value{ErrorVal(ErrValNUM), NumberVal(12), NumberVal(1000)},
			wantErr: true,
		},
		{
			name:    "error propagation in fv",
			args:    []Value{NumberVal(0.05), NumberVal(12), NumberVal(1000), ErrorVal(ErrValREF)},
			wantErr: true,
		},
		{
			name:    "error propagation in nper",
			args:    []Value{NumberVal(0.05), ErrorVal(ErrValNA), NumberVal(1000)},
			wantErr: true,
		},
		{
			name:    "error propagation in pv",
			args:    []Value{NumberVal(0.05), NumberVal(12), ErrorVal(ErrValDIV0)},
			wantErr: true,
		},
		{
			name:    "error propagation in type",
			args:    []Value{NumberVal(0.05), NumberVal(12), NumberVal(1000), NumberVal(0), ErrorVal(ErrValVALUE)},
			wantErr: true,
		},

		// --- Error: empty string (not coercible to number) ---
		{
			name:    "error: empty string rate",
			args:    []Value{StringVal(""), NumberVal(12), NumberVal(1000)},
			wantErr: true,
		},

		// --- Boolean coercion ---
		{
			name: "bool coercion: TRUE=1 for all args",
			args: []Value{BoolVal(true), BoolVal(true), BoolVal(true)},
			want: -2.00,
		},
		{
			name: "bool coercion: FALSE=0 for rate",
			args: []Value{BoolVal(false), NumberVal(10), NumberVal(1000)},
			want: -100,
		},
		{
			name: "bool coercion: TRUE for type (beginning of period)",
			args: []Value{NumberVal(0.10 / 12), NumberVal(120), NumberVal(50000), NumberVal(0), BoolVal(true)},
			want: -655.29,
		},
		{
			name: "bool coercion: FALSE for type (end of period)",
			args: []Value{NumberVal(0.10 / 12), NumberVal(120), NumberVal(50000), NumberVal(0), BoolVal(false)},
			want: -660.75,
		},

		// --- Empty cell references (coerce to 0) ---
		{
			name: "empty val for rate (coerces to 0)",
			args: []Value{EmptyVal(), NumberVal(10), NumberVal(1000)},
			want: -100,
		},
		{
			name: "empty val for fv (coerces to 0)",
			args: []Value{NumberVal(0.05 / 12), NumberVal(60), NumberVal(10000), EmptyVal()},
			want: -188.71,
		},
		{
			name: "empty val for type (coerces to 0, end of period)",
			args: []Value{NumberVal(0.10 / 12), NumberVal(120), NumberVal(50000), NumberVal(0), EmptyVal()},
			want: -660.75,
		},

		// --- Negative interest rate ---
		{
			name: "negative rate (deflation scenario)",
			args: numArgs(-0.05/12, 60, 10000),
			want: -146.35,
		},

		// --- Single period ---
		{
			name: "single period nper=1",
			args: numArgs(0.10, 1, 1000),
			want: -1100.00,
		},

		// --- Very high interest rate ---
		{
			name: "very high rate 100% per period",
			args: numArgs(1.0, 12, 10000),
			want: -10002.44,
		},

		// --- Fractional nper ---
		{
			name: "fractional nper 30.5 months",
			args: numArgs(0.05/12, 30.5, 10000),
			want: -349.82,
		},

		// --- Negative fv (you owe money at end) ---
		{
			name: "negative fv gives positive payment",
			args: numArgs(0.06/12, 60, 0, -10000),
			want: 143.33,
		},

		// --- Both pv and fv negative ---
		{
			name: "both pv and fv negative (receiving money)",
			args: numArgs(0.05/12, 60, -10000, -5000),
			want: 262.24,
		},

		// --- Negative nper ---
		{
			name: "negative nper",
			args: numArgs(0.05/12, -60, 10000),
			want: 147.05,
		},

		// --- Very large present value ---
		{
			name: "large pv $10M mortgage",
			args: numArgs(0.04/12, 360, 10000000),
			want: -47741.53,
		},

		// --- Small nper (2 periods) ---
		{
			name: "small nper 2 periods",
			args: numArgs(0.10, 2, 1000),
			want: -576.19,
		},

		// --- Zero result: pv=0, fv=0, rate != 0 ---
		{
			name: "zero result: pv and fv both zero",
			args: numArgs(0.05, 12, 0),
			want: 0,
		},

		// --- Negative pv with fv and type=1 ---
		{
			name: "negative pv with fv and type=1 (investment receiving payments)",
			args: numArgs(0.06/12, 60, -10000, 5000, 1),
			want: 121.06,
		},

		// --- Annual payments (doc consistency remark) ---
		{
			name: "annual payments 12% over 4 years",
			args: numArgs(0.12, 4, 10000),
			want: -3292.34,
		},

		// --- Quarterly payments ---
		{
			name: "quarterly payments 8% annual",
			args: numArgs(0.08/4, 4, 5000),
			want: -1313.12,
		},

		// --- Very small rate ---
		{
			name: "very small rate 0.01% annual",
			args: numArgs(0.0001/12, 360, 200000),
			want: -556.39,
		},

		// --- type=0 explicit with fv ---
		{
			name: "explicit type=0 with pv and fv",
			args: numArgs(0.07/12, 84, 30000, 5000, 0),
			want: -499.08,
		},

		// --- type=1 with both pv and fv ---
		{
			name: "type=1 with pv and fv",
			args: numArgs(0.07/12, 84, 30000, 5000, 1),
			want: -496.18,
		},

		// --- Zero rate ignores type (no interest = timing irrelevant) ---
		{
			name: "zero rate with type=1",
			args: numArgs(0, 12, 6000, 0, 1),
			want: -500.00,
		},

		// --- type=5 treated as type=1 (any non-zero) ---
		{
			name: "type=5 treated as beginning of period",
			args: numArgs(0.10/12, 120, 50000, 0, 5),
			want: -655.29,
		},
		{
			name: "type=-1 treated as beginning of period",
			args: numArgs(0.10/12, 120, 50000, 0, -1),
			want: -655.29,
		},

		// --- One arg only (too few) ---
		{
			name:    "error: too few args (1)",
			args:    numArgs(0.05),
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			v, err := fnPMT(tt.args)
			if err != nil {
				t.Fatalf("unexpected Go error: %v", err)
			}
			if tt.wantErr {
				assertError(t, tt.name, v)
			} else {
				assertClose(t, tt.name, v, tt.want)
			}
		})
	}
}

// === FV ===

func TestFV_Savings(t *testing.T) {
	// FV(0.06/12, 120, -200, -5000) — $200/month, $5000 initial, 6% over 10 years
	// Expected: 41872.85
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

func TestFV_Comprehensive(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// --- Basic: monthly savings with interest ---
		{
			name: "monthly savings 5% annual over 5 years",
			args: numArgs(0.05/12, 60, -200),
			want: 13601.22,
		},
		// --- Zero interest rate ---
		{
			name: "zero rate pmt only",
			args: numArgs(0, 24, -500),
			want: 12000,
		},
		{
			name: "zero rate with pv",
			args: numArgs(0, 12, -100, -500),
			want: 1700,
		},
		// --- Zero payment ---
		{
			name: "zero pmt with pv only",
			args: numArgs(0.05/12, 120, 0, -10000),
			want: 16470.09,
		},
		// --- With present value (pv argument) ---
		{
			name: "pmt and pv combined",
			args: numArgs(0.08/12, 60, -300, -5000),
			want: 29492.29,
		},
		// --- Type=0 (end of period) vs Type=1 (beginning of period) ---
		{
			name: "type 0 end of period",
			args: numArgs(0.06/12, 12, -100, 0, 0),
			want: 1233.56,
		},
		{
			name: "type 1 beginning of period",
			args: numArgs(0.06/12, 12, -100, 0, 1),
			want: 1239.72,
		},
		// --- Negative pmt (paying out) ---
		{
			name: "negative pmt deposits",
			args: numArgs(0.10/12, 36, -500),
			want: 20890.91,
		},
		{
			name: "positive pmt withdrawals",
			args: numArgs(0.06/12, 24, 200, -10000),
			want: 6185.21,
		},
		// --- Large nper (many periods) ---
		{
			name: "large nper 480 months",
			args: numArgs(0.04/12, 480, -100),
			want: 118196.13,
		},
		// --- All zeros ---
		{
			name: "all zeros",
			args: numArgs(0, 0, 0),
			want: 0,
		},
		{
			name: "all zeros with pv and type",
			args: numArgs(0, 0, 0, 0, 0),
			want: 0,
		},
		// --- Only pv, no pmt ---
		{
			name: "only pv no pmt",
			args: numArgs(0.10, 5, 0, -1000),
			want: 1610.51,
		},
		// --- Only pmt, no pv ---
		{
			name: "only pmt no pv",
			args: numArgs(0.08/12, 120, -200),
			want: 36589.21,
		},
		// --- Doc example 1: FV(0.06/12, 10, -200, -500, 1) = 2581.40 ---
		{
			name: "doc example 1",
			args: numArgs(0.06/12, 10, -200, -500, 1),
			want: 2581.40,
		},
		// --- Doc example 2: FV(0.12/12, 12, -1000) = 12682.50 ---
		{
			name: "doc example 2",
			args: numArgs(0.12/12, 12, -1000),
			want: 12682.50,
		},
		// --- Doc example 3: FV(0.11/12, 35, -2000, 0, 1) = 82846.25 ---
		{
			name: "doc example 3",
			args: numArgs(0.11/12, 35, -2000, 0, 1),
			want: 82846.25,
		},
		// --- Doc example 4: FV(0.06/12, 12, -100, -1000, 1) = 2301.40 ---
		{
			name: "doc example 4",
			args: numArgs(0.06/12, 12, -100, -1000, 1),
			want: 2301.40,
		},
		// --- Monthly vs annual rates ---
		{
			name: "annual rate 10% over 3 years",
			args: numArgs(0.10, 3, -1000),
			want: 3310.00,
		},
		{
			name: "monthly rate equivalent",
			args: numArgs(0.10/12, 36, -1000),
			want: 41781.82,
		},
		// --- String coercion for args ---
		{
			name: "string coercion rate",
			args: []Value{StringVal("0.06"), NumberVal(12), NumberVal(-100)},
			want: 1686.99,
		},
		{
			name: "string coercion nper",
			args: []Value{NumberVal(0.05), StringVal("3"), NumberVal(-1000)},
			want: 3152.50,
		},
		{
			name: "string coercion pmt",
			args: []Value{NumberVal(0.05), NumberVal(3), StringVal("-1000")},
			want: 3152.50,
		},
		{
			name: "string coercion pv",
			args: []Value{NumberVal(0.05), NumberVal(3), NumberVal(0), StringVal("-1000")},
			want: 1157.63,
		},
		{
			name: "string coercion type",
			args: []Value{NumberVal(0.06 / 12), NumberVal(12), NumberVal(-100), NumberVal(0), StringVal("1")},
			want: 1239.72,
		},
		// --- Negative rate ---
		{
			name: "negative rate pmt only",
			args: numArgs(-0.05, 10, -100),
			want: 802.53,
		},
		{
			name: "negative rate with pv",
			args: numArgs(-0.02, 10, -100, -5000),
			want: 5000.00,
		},
		{
			name: "negative rate type 1",
			args: numArgs(-0.05, 10, -100, 0, 1),
			want: 762.40,
		},
		// --- Very high rate (100% per period) ---
		{
			name: "very high rate 100 pct",
			args: numArgs(1, 5, -100),
			want: 3100.00,
		},
		{
			name: "very high rate 100 pct with pv",
			args: numArgs(1, 5, -100, -1000),
			want: 35100.00,
		},
		{
			name: "very high rate 100 pct type 1",
			args: numArgs(1, 5, -100, 0, 1),
			want: 6200.00,
		},
		// --- Zero nper ---
		{
			name: "zero nper with pv",
			args: numArgs(0.05, 0, -100, -1000),
			want: 1000.00,
		},
		{
			name: "zero nper zero rate",
			args: numArgs(0, 0, -100, -1000),
			want: 1000.00,
		},
		// --- Large nper (360 months at 1%) ---
		{
			name: "large nper 360 at 1 pct",
			args: numArgs(0.01, 360, -100),
			want: 349496.41,
		},
		// --- Very small rate close to zero ---
		{
			name: "very small rate near zero",
			args: numArgs(0.000001, 12, -100),
			want: 1200.01,
		},
		// --- All positive args (borrowing scenario) ---
		{
			name: "positive pmt and pv borrowing",
			args: numArgs(0.05, 10, 100, 1000),
			want: -2886.68,
		},
		// --- Negative pmt with positive pv ---
		{
			name: "negative pmt positive pv",
			args: numArgs(0.05, 10, -200, 5000),
			want: -5628.89,
		},
		// --- Both pmt and pv negative ---
		{
			name: "both pmt and pv negative",
			args: numArgs(0.05, 10, -200, -5000),
			want: 10660.05,
		},
		// --- Fractional nper ---
		{
			name: "fractional nper 2.5 periods",
			args: numArgs(0.05, 2.5, -1000),
			want: 2594.53,
		},
		// --- Large pv compound only ---
		{
			name: "large pv compound growth",
			args: numArgs(0.08, 30, 0, -1000000),
			want: 10062656.89,
		},
		// --- Single period ---
		{
			name: "single period type 0",
			args: numArgs(0.10, 1, -1000),
			want: 1000.00,
		},
		{
			name: "single period type 1",
			args: numArgs(0.10, 1, -1000, 0, 1),
			want: 1100.00,
		},
		// --- Zero rate type 1 ---
		{
			name: "zero rate type 1 with pv",
			args: numArgs(0, 12, -100, -500, 1),
			want: 1700.00,
		},
		// --- Boolean coercion ---
		{
			name: "bool TRUE as rate coerced to 1",
			args: []Value{BoolVal(true), NumberVal(5), NumberVal(-100)},
			want: 3100.00,
		},
		{
			name: "bool FALSE as rate coerced to 0",
			args: []Value{BoolVal(false), NumberVal(12), NumberVal(-100)},
			want: 1200.00,
		},
		{
			name: "bool TRUE as type",
			args: []Value{NumberVal(0.06 / 12), NumberVal(12), NumberVal(-100), NumberVal(0), BoolVal(true)},
			want: 1239.72,
		},
		{
			name: "bool FALSE as type",
			args: []Value{NumberVal(0.06 / 12), NumberVal(12), NumberVal(-100), NumberVal(0), BoolVal(false)},
			want: 1233.56,
		},
		{
			name: "bool TRUE as nper",
			args: []Value{NumberVal(0.10), BoolVal(true), NumberVal(-1000)},
			want: 1000.00,
		},
		// --- Empty cell args treated as 0 ---
		{
			name: "empty rate treated as 0",
			args: []Value{EmptyVal(), NumberVal(12), NumberVal(-100)},
			want: 1200.00,
		},
		{
			name: "empty pv treated as 0",
			args: []Value{NumberVal(0.05), NumberVal(3), NumberVal(-1000), EmptyVal()},
			want: 3152.50,
		},
		{
			name: "empty type treated as 0",
			args: []Value{NumberVal(0.06 / 12), NumberVal(12), NumberVal(-100), NumberVal(0), EmptyVal()},
			want: 1233.56,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnFV(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestFV_Errors(t *testing.T) {
	tests := []struct {
		name string
		args []Value
	}{
		{
			name: "too few args",
			args: numArgs(0.05, 10),
		},
		{
			name: "too many args",
			args: numArgs(0.05, 10, -100, 0, 0, 99),
		},
		{
			name: "non-numeric rate",
			args: []Value{StringVal("abc"), NumberVal(10), NumberVal(-100)},
		},
		{
			name: "non-numeric nper",
			args: []Value{NumberVal(0.05), StringVal("xyz"), NumberVal(-100)},
		},
		{
			name: "non-numeric pmt",
			args: []Value{NumberVal(0.05), NumberVal(10), StringVal("bad")},
		},
		{
			name: "non-numeric pv",
			args: []Value{NumberVal(0.05), NumberVal(10), NumberVal(-100), StringVal("bad")},
		},
		{
			name: "non-numeric type",
			args: []Value{NumberVal(0.05), NumberVal(10), NumberVal(-100), NumberVal(0), StringVal("bad")},
		},
		{
			name: "zero args",
			args: []Value{},
		},
		{
			name: "one arg",
			args: numArgs(0.05),
		},
		{
			name: "error value propagation in rate",
			args: []Value{ErrorVal(ErrValNUM), NumberVal(10), NumberVal(-100)},
		},
		{
			name: "error value propagation in nper",
			args: []Value{NumberVal(0.05), ErrorVal(ErrValDIV0), NumberVal(-100)},
		},
		{
			name: "error value propagation in pmt",
			args: []Value{NumberVal(0.05), NumberVal(10), ErrorVal(ErrValREF)},
		},
		{
			name: "error value propagation in pv",
			args: []Value{NumberVal(0.05), NumberVal(10), NumberVal(-100), ErrorVal(ErrValNA)},
		},
		{
			name: "error value propagation in type",
			args: []Value{NumberVal(0.05), NumberVal(10), NumberVal(-100), NumberVal(0), ErrorVal(ErrValVALUE)},
		},
		{
			name: "empty string rate",
			args: []Value{StringVal(""), NumberVal(10), NumberVal(-100)},
		},
		{
			name: "empty string nper",
			args: []Value{NumberVal(0.05), StringVal(""), NumberVal(-100)},
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnFV(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			assertError(t, tc.name, v)
		})
	}
}

// === PV ===

func TestPV_Annuity(t *testing.T) {
	// PV(0.08/12, 240, -500) — 20 years of $500/month payments at 8%
	// Expected: 59777.15
	v, err := fnPV(numArgs(0.08/12, 240, -500))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PV annuity", v, 59777.15)
}

func TestPV_Comprehensive(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// --- Basic: present value of annuity ---
		{
			name: "basic annuity monthly payments",
			args: numArgs(0.08/12, 240, -500),
			want: 59777.15,
		},
		// --- Zero interest rate ---
		{
			name: "zero rate pmt only",
			args: numArgs(0, 24, -500),
			want: 12000,
		},
		{
			name: "zero rate with fv",
			args: numArgs(0, 12, -100, -500),
			want: 1700,
		},
		// --- Zero payment with fv ---
		{
			name: "zero pmt with fv only",
			args: numArgs(0.05/12, 120, 0, -10000),
			want: 6071.61,
		},
		// --- With future value (fv argument) ---
		{
			name: "pmt and fv combined",
			args: numArgs(0.08/12, 60, -300, -5000),
			want: 18151.58,
		},
		// --- Type=0 (end of period) vs Type=1 (beginning of period) ---
		{
			name: "type 0 end of period",
			args: numArgs(0.06/12, 12, -100, 0, 0),
			want: 1161.89,
		},
		{
			name: "type 1 beginning of period",
			args: numArgs(0.06/12, 12, -100, 0, 1),
			want: 1167.70,
		},
		// --- Negative pmt (paying out = positive PV) ---
		{
			name: "negative pmt deposits yield positive PV",
			args: numArgs(0.10/12, 36, -500),
			want: 15495.62,
		},
		{
			name: "positive pmt withdrawals yield negative PV",
			args: numArgs(0.06/12, 24, 200),
			want: -4512.57,
		},
		// --- Large nper (many periods) ---
		{
			name: "large nper 480 months",
			args: numArgs(0.04/12, 480, -100),
			want: 23926.97,
		},
		// --- All zeros ---
		{
			name: "all zeros",
			args: numArgs(0, 0, 0),
			want: 0,
		},
		{
			name: "all zeros with fv and type",
			args: numArgs(0, 0, 0, 0, 0),
			want: 0,
		},
		// --- Only fv, no pmt ---
		{
			name: "only fv no pmt",
			args: numArgs(0.10, 5, 0, -1000),
			want: 620.92,
		},
		// --- Only pmt, no fv ---
		{
			name: "only pmt no fv",
			args: numArgs(0.08/12, 120, -200),
			want: 16484.30,
		},
		// --- Doc example: PV(0.08/12, 12*20, 500, , 0) = -59777.15 ---
		// (sign convention: positive pmt = receiving money, so PV is negative)
		{
			name: "doc example annuity payout",
			args: numArgs(0.08/12, 12*20, 500, 0, 0),
			want: -59777.15,
		},
		// --- Monthly vs annual rates ---
		{
			name: "annual rate 10% over 3 years",
			args: numArgs(0.10, 3, -1000),
			want: 2486.85,
		},
		{
			name: "monthly rate equivalent",
			args: numArgs(0.10/12, 36, -1000),
			want: 30991.24,
		},
		// --- Mortgage present value ---
		{
			name: "mortgage 30yr 5%",
			args: numArgs(0.05/12, 360, -1073.64),
			want: 199999.40,
		},
		// --- String coercion for args ---
		{
			name: "string coercion rate",
			args: []Value{StringVal("0.06"), NumberVal(12), NumberVal(-100)},
			want: 838.38,
		},
		{
			name: "string coercion nper",
			args: []Value{NumberVal(0.05), StringVal("3"), NumberVal(-1000)},
			want: 2723.25,
		},
		{
			name: "string coercion pmt",
			args: []Value{NumberVal(0.05), NumberVal(3), StringVal("-1000")},
			want: 2723.25,
		},
		{
			name: "string coercion fv",
			args: []Value{NumberVal(0.05), NumberVal(3), NumberVal(0), StringVal("-1000")},
			want: 863.84,
		},
		{
			name: "string coercion type",
			args: []Value{NumberVal(0.06 / 12), NumberVal(12), NumberVal(-100), NumberVal(0), StringVal("1")},
			want: 1167.70,
		},
		// --- Single period ---
		{
			name: "single period nper=1 type=0",
			args: numArgs(0.10, 1, -1000),
			want: 909.09,
		},
		{
			name: "single period nper=1 type=1",
			args: numArgs(0.10, 1, -1000, 0, 1),
			want: 1000.0,
		},
		// --- Very high interest rate ---
		{
			name: "very high rate 100%",
			args: numArgs(1.0, 5, -100),
			want: 96.88,
		},
		// --- Very small interest rate ---
		{
			name: "very small rate 0.001%",
			args: numArgs(0.00001, 120, -500),
			want: 59963.71,
		},
		// --- Fractional nper ---
		{
			name: "fractional nper 6.5",
			args: numArgs(0.05, 6.5, -1000),
			want: 5435.37,
		},
		// --- Negative nper ---
		{
			name: "negative nper",
			args: numArgs(0.05, -5, -1000),
			want: -5525.63,
		},
		// --- Both pmt and fv with type=1 ---
		{
			name: "pmt and fv combined type=1",
			args: numArgs(0.06/12, 60, -200, -5000, 1),
			want: 14103.70,
		},
		// --- Boolean coercion for type ---
		{
			name: "bool TRUE for type",
			args: []Value{NumberVal(0.06 / 12), NumberVal(12), NumberVal(-100), NumberVal(0), BoolVal(true)},
			want: 1167.70,
		},
		{
			name: "bool FALSE for type",
			args: []Value{NumberVal(0.06 / 12), NumberVal(12), NumberVal(-100), NumberVal(0), BoolVal(false)},
			want: 1161.89,
		},
		// --- Empty cell references (EmptyVal coerces to 0) ---
		{
			name: "empty rate coerces to zero",
			args: []Value{EmptyVal(), NumberVal(10), NumberVal(-500)},
			want: 5000,
		},
		{
			name: "empty fv coerces to zero",
			args: []Value{NumberVal(0.05), NumberVal(3), NumberVal(-1000), EmptyVal()},
			want: 2723.25,
		},
		{
			name: "empty type coerces to zero",
			args: []Value{NumberVal(0.06 / 12), NumberVal(12), NumberVal(-100), NumberVal(0), EmptyVal()},
			want: 1161.89,
		},
		// --- Large fv with no pmt ---
		{
			name: "large fv no pmt 30yr",
			args: numArgs(0.03/12, 360, 0, -1000000),
			want: 407026.55,
		},
		// --- Classic mortgage question ---
		{
			name: "classic mortgage PV 8% 30yr $1000/mo",
			args: numArgs(0.08/12, 360, -1000),
			want: 136283.49,
		},
		// --- Zero pmt zero fv nonzero rate ---
		{
			name: "zero pmt zero fv",
			args: numArgs(0.05, 10, 0, 0),
			want: 0,
		},
		// --- Very large nper ---
		{
			name: "very large nper 1000 periods",
			args: numArgs(0.01, 1000, -1),
			want: 100.0,
		},
		// --- Negative rate ---
		{
			name: "negative rate",
			args: numArgs(-0.05, 10, -1000),
			want: 13403.65,
		},
		// --- Type non-zero non-one treated as 1 ---
		{
			name: "type=5 treated as beginning of period",
			args: numArgs(0.06/12, 12, -100, 0, 5),
			want: 1167.70,
		},
		// --- Zero rate with type=1 (type is ignored in zero-rate branch) ---
		{
			name: "zero rate with type=1",
			args: numArgs(0, 12, -100, 0, 1),
			want: 1200,
		},
		// --- fv only type=0 vs type=1 (pmt=0 so type doesn't change result) ---
		{
			name: "fv only type=0",
			args: numArgs(0.05, 10, 0, -1000, 0),
			want: 613.91,
		},
		{
			name: "fv only type=1",
			args: numArgs(0.05, 10, 0, -1000, 1),
			want: 613.91,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnPV(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestPV_Errors(t *testing.T) {
	tests := []struct {
		name string
		args []Value
	}{
		{
			name: "too few args",
			args: numArgs(0.05, 10),
		},
		{
			name: "too many args",
			args: numArgs(0.05, 10, -100, 0, 0, 99),
		},
		{
			name: "non-numeric rate",
			args: []Value{StringVal("abc"), NumberVal(10), NumberVal(-100)},
		},
		{
			name: "non-numeric nper",
			args: []Value{NumberVal(0.05), StringVal("xyz"), NumberVal(-100)},
		},
		{
			name: "non-numeric pmt",
			args: []Value{NumberVal(0.05), NumberVal(10), StringVal("bad")},
		},
		{
			name: "non-numeric fv",
			args: []Value{NumberVal(0.05), NumberVal(10), NumberVal(-100), StringVal("bad")},
		},
		{
			name: "non-numeric type",
			args: []Value{NumberVal(0.05), NumberVal(10), NumberVal(-100), NumberVal(0), StringVal("bad")},
		},
		{
			name: "nper zero with nonzero rate",
			args: numArgs(0.05, 0, -100),
		},
		// --- Error propagation: ErrorVal passed as argument ---
		{
			name: "error propagation rate",
			args: []Value{ErrorVal(ErrValNUM), NumberVal(10), NumberVal(-100)},
		},
		{
			name: "error propagation nper",
			args: []Value{NumberVal(0.05), ErrorVal(ErrValREF), NumberVal(-100)},
		},
		{
			name: "error propagation pmt",
			args: []Value{NumberVal(0.05), NumberVal(10), ErrorVal(ErrValDIV0)},
		},
		{
			name: "error propagation fv",
			args: []Value{NumberVal(0.05), NumberVal(10), NumberVal(-100), ErrorVal(ErrValNA)},
		},
		{
			name: "error propagation type",
			args: []Value{NumberVal(0.05), NumberVal(10), NumberVal(-100), NumberVal(0), ErrorVal(ErrValVALUE)},
		},
		// --- Empty string is #VALUE! ---
		{
			name: "empty string rate",
			args: []Value{StringVal(""), NumberVal(10), NumberVal(-100)},
		},
		// --- Single arg (too few) ---
		{
			name: "single arg",
			args: numArgs(0.05),
		},
		// --- No args (too few) ---
		{
			name: "no args",
			args: []Value{},
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnPV(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			assertError(t, tc.name, v)
		})
	}
}

// === NPER ===

func TestNPER_Loan(t *testing.T) {
	// NPER(0.01, -100, 1000) — how many months to pay off $1000 at 1%/month
	// Expected: 10.58
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

func TestNPER_Comprehensive(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// --- Doc examples ---
		{
			name: "doc: type=1, rate=12%/12, pmt=-100, pv=-1000, fv=10000",
			args: numArgs(0.12/12, -100, -1000, 10000, 1),
			want: 59.67,
		},
		{
			name: "doc: type=0 (default), rate=12%/12, pmt=-100, pv=-1000, fv=10000",
			args: numArgs(0.12/12, -100, -1000, 10000),
			want: 60.08,
		},
		{
			name: "doc: no fv, rate=12%/12, pmt=-100, pv=-1000",
			args: numArgs(0.12/12, -100, -1000),
			want: -9.58,
		},

		// --- Basic loan payoff ---
		{
			name: "basic loan: 1%/month, $100 payments, $1000 loan",
			args: numArgs(0.01, -100, 1000),
			want: 10.58,
		},

		// --- Zero rate ---
		{
			name: "zero rate: pmt only",
			args: numArgs(0, -100, 1000),
			want: 10,
		},
		{
			name: "zero rate with fv",
			args: numArgs(0, -200, 1000, 5000),
			want: 30,
		},
		{
			name:    "zero rate: pmt=0 should error",
			args:    numArgs(0, 0, 1000),
			wantErr: true,
		},

		// --- With future value ---
		{
			name: "loan with future value (balloon payment)",
			args: numArgs(0.05/12, -500, 20000, 5000),
			want: 53.67,
		},

		// --- Type=0 vs type=1 ---
		{
			name: "type=0 end of period",
			args: numArgs(0.06/12, -200, 10000, 0, 0),
			want: 57.68,
		},
		{
			name: "type=1 beginning of period (fewer periods)",
			args: numArgs(0.06/12, -200, 10000, 0, 1),
			want: 57.35,
		},

		// --- Saving scenario: positive pmt, negative fv ---
		{
			name: "saving: monthly $300 deposits to reach $50000",
			args: numArgs(0.04/12, -300, 0, 50000),
			want: 132.77,
		},

		// --- Large loan, small payment ---
		{
			name: "large loan small payment: 30yr mortgage check",
			args: numArgs(0.06/12, -1199.10, 200000),
			want: 360.00,
		},

		// --- Negative pmt (standard loan payments) ---
		{
			name: "negative pmt standard loan",
			args: numArgs(0.08/12, -1000, 50000),
			want: 61.02,
		},

		// --- Zero pmt with fv: result based on rate and pv/fv ---
		{
			name: "zero pmt with fv: compound growth only",
			args: numArgs(0.05, 0, -1000, 2000),
			want: 14.21,
		},

		// --- String coercion ---
		{
			name: "string coercion: all numeric strings",
			args: []Value{StringVal("0.01"), StringVal("-100"), StringVal("1000")},
			want: 10.58,
		},
		{
			name: "string coercion: rate as string",
			args: []Value{StringVal("0"), NumberVal(-250), NumberVal(1000)},
			want: 4,
		},

		// --- Error: too few args ---
		{
			name:    "error: too few args (2)",
			args:    numArgs(0.01, -100),
			wantErr: true,
		},
		{
			name:    "error: too few args (1)",
			args:    numArgs(0.01),
			wantErr: true,
		},

		// --- Error: too many args ---
		{
			name:    "error: too many args (6)",
			args:    numArgs(0.01, -100, 1000, 0, 0, 99),
			wantErr: true,
		},

		// --- Error: non-numeric ---
		{
			name:    "error: non-numeric rate",
			args:    []Value{StringVal("abc"), NumberVal(-100), NumberVal(1000)},
			wantErr: true,
		},
		{
			name:    "error: non-numeric pmt",
			args:    []Value{NumberVal(0.01), StringVal("xyz"), NumberVal(1000)},
			wantErr: true,
		},
		{
			name:    "error: non-numeric pv",
			args:    []Value{NumberVal(0.01), NumberVal(-100), StringVal("bad")},
			wantErr: true,
		},

		// --- Error: non-numeric fv ---
		{
			name:    "error: non-numeric fv",
			args:    []Value{NumberVal(0.01), NumberVal(-100), NumberVal(1000), StringVal("bad")},
			wantErr: true,
		},

		// --- Error: non-numeric type ---
		{
			name:    "error: non-numeric type",
			args:    []Value{NumberVal(0.01), NumberVal(-100), NumberVal(1000), NumberVal(0), StringVal("nope")},
			wantErr: true,
		},

		// --- Boolean coercion ---
		{
			name: "boolean coercion: TRUE as type (=1)",
			args: []Value{NumberVal(0.06 / 12), NumberVal(-200), NumberVal(10000), NumberVal(0), BoolVal(true)},
			want: 57.35,
		},
		{
			name: "boolean coercion: FALSE as type (=0)",
			args: []Value{NumberVal(0.06 / 12), NumberVal(-200), NumberVal(10000), NumberVal(0), BoolVal(false)},
			want: 57.68,
		},

		// --- Empty cell references ---
		{
			name: "empty fv treated as 0",
			args: []Value{NumberVal(0.01), NumberVal(-100), NumberVal(1000), EmptyVal()},
			want: 10.58,
		},
		{
			name: "empty type treated as 0",
			args: []Value{NumberVal(0.01), NumberVal(-100), NumberVal(1000), NumberVal(0), EmptyVal()},
			want: 10.58,
		},
		{
			name: "empty rate treated as 0",
			args: []Value{EmptyVal(), NumberVal(-100), NumberVal(1000)},
			want: 10,
		},

		// --- Very high interest rate ---
		{
			name:    "very high rate: 100% per period, pmt too small",
			args:    numArgs(1.0, -500, 1000),
			wantErr: true,
		},
		{
			name: "high rate: 50% per period, large payment",
			args: numArgs(0.50, -2000, 1000),
			want: 0.71,
		},

		// --- Very small interest rate ---
		{
			name: "very small rate: 0.001% per period",
			args: numArgs(0.00001, -100, 1000),
			want: 10.00,
		},

		// --- Both pv and fv specified ---
		{
			name: "both pv and fv: loan with residual",
			args: numArgs(0.005, -500, 20000, 5000),
			want: 54.52,
		},

		// --- Payment too small to cover interest (NUM error) ---
		{
			name:    "pmt too small: interest exceeds payment",
			args:    numArgs(0.10, -50, 1000),
			wantErr: true,
		},

		// --- Precise Excel doc values ---
		{
			name: "excel doc: type=1, precise value 59.6738657",
			args: numArgs(0.12/12, -100, -1000, 10000, 1),
			want: 59.6738657,
		},
		{
			name: "excel doc: type=0 default, precise value 60.0821229",
			args: numArgs(0.12/12, -100, -1000, 10000),
			want: 60.0821229,
		},
		{
			name: "excel doc: no fv, precise value -9.57859404",
			args: numArgs(0.12/12, -100, -1000),
			want: -9.57859404,
		},

		// --- Investment scenarios: positive pv, negative fv ---
		{
			name: "investment: grow $5000 to $10000 at 6%/yr",
			args: numArgs(0.06/12, 0, -5000, 10000),
			want: 138.98,
		},

		// --- Zero args edge ---
		{
			name:    "error: zero args",
			args:    []Value{},
			wantErr: true,
		},

		// --- Negative rate (deflation scenario) ---
		{
			name: "negative rate: -1% per period (deflation)",
			args: numArgs(-0.01, -100, 1000),
			want: 9.48,
		},

		// --- pmt=0, rate=0: both zero should error ---
		{
			name:    "pmt=0 rate=0: div by zero",
			args:    numArgs(0, 0, 1000, 5000),
			wantErr: true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnNPER(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

// === RATE ===

func TestRATE_Loan(t *testing.T) {
	// RATE(360, -1073.64, 200000) — reverse of the PMT example
	// Expected: ~0.004167 (0.05/12)
	v, err := fnRATE(numArgs(360, -1073.64, 200000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "RATE loan", v, 0.05/12)
}

func TestRATE_NoSolution(t *testing.T) {
	// RATE(60, 0, 50000) — pmt=0, fv=0, pv≠0: no solution exists.
	// Returns #NUM!.
	v, err := fnRATE(numArgs(60, 0, 50000))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueError {
		t.Fatalf("expected #NUM! error, got %v", v)
	}
}

func TestRATE_Comprehensive(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// --- Basic: loan rate calculation ---
		{
			name: "basic 30yr mortgage rate",
			// RATE(360, -1073.64, 200000) ≈ 0.05/12
			args: numArgs(360, -1073.64, 200000),
			want: 0.05 / 12,
		},
		// --- Doc example ---
		{
			name: "doc: 4yr loan monthly, RATE(48, -200, 8000)",
			// Expected ~1% per month
			args: numArgs(48, -200, 8000),
			want: 0.0077,
		},
		// --- With future value ---
		{
			name: "with future value, pv to fv with payments",
			// RATE(120, -100, 0, 20000) — saving $100/mo to reach $20000
			args: numArgs(120, -100, 0, 20000),
			want: 0.00596, // ~7.15% annual
		},
		// --- Type=0 vs type=1 ---
		{
			name: "type=0 end of period (default)",
			// RATE(60, -500, 25000, 0, 0)
			args: numArgs(60, -500, 25000, 0, 0),
			want: 0.003242,
		},
		{
			name: "type=1 beginning of period",
			// RATE(60, -500, 25000, 0, 1)
			args: numArgs(60, -500, 25000, 0, 1),
			want: 0.003231,
		},
		// --- With explicit guess ---
		{
			name: "explicit guess close to answer",
			// RATE(360, -1073.64, 200000, 0, 0, 0.004)
			args: numArgs(360, -1073.64, 200000, 0, 0, 0.004),
			want: 0.05 / 12,
		},
		{
			name: "explicit guess far from answer",
			// RATE(48, -200, 8000, 0, 0, 0.5)
			args: numArgs(48, -200, 8000, 0, 0, 0.5),
			want: 0.0077,
		},
		// --- Zero payment: pv to fv only ---
		{
			name: "zero payment pv to fv",
			// RATE(10, 0, -1000, 1500) — what rate turns 1000 into 1500 in 10 periods?
			// (1+r)^10 = 1.5 → r ≈ 0.04138
			args: numArgs(10, 0, -1000, 1500),
			want: 0.04138,
		},
		// --- Savings scenario ---
		{
			name: "savings monthly contributions",
			// RATE(240, -300, 0, 150000) — save $300/mo for 20yr to get $150k
			args: numArgs(240, -300, 0, 150000),
			want: 0.003055, // ~3.67% annual
		},
		// --- High rate result ---
		{
			name: "high rate result",
			// RATE(12, -1000, 5000) — short loan, large payments relative to principal
			args: numArgs(12, -1000, 5000),
			want: 0.16943,
		},
		// --- Low rate result ---
		{
			name: "low rate result",
			// RATE(360, -1000, 350000) — low rate mortgage
			args: numArgs(360, -1000, 350000),
			want: 0.000953,
		},
		// --- Negative nper → #NUM! ---
		{
			name:    "negative nper",
			args:    numArgs(-12, -100, 1000),
			want:    0,
			wantErr: true,
		},
		// --- Zero nper → #NUM! ---
		{
			name:    "zero nper",
			args:    numArgs(0, -100, 1000),
			want:    0,
			wantErr: true,
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnRATE(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestRATE_Errors(t *testing.T) {
	tests := []struct {
		name string
		args []Value
	}{
		{
			name: "too few args",
			args: numArgs(360, -1073.64),
		},
		{
			name: "too many args",
			args: numArgs(360, -1073.64, 200000, 0, 0, 0.1, 99),
		},
		{
			name: "non-numeric nper",
			args: []Value{StringVal("abc"), NumberVal(-1073.64), NumberVal(200000)},
		},
		{
			name: "non-numeric pmt",
			args: []Value{NumberVal(360), StringVal("abc"), NumberVal(200000)},
		},
		{
			name: "non-numeric pv",
			args: []Value{NumberVal(360), NumberVal(-1073.64), StringVal("abc")},
		},
		{
			name: "no convergence pmt=0 fv=0 pv!=0",
			// pmt=0, fv=0, pv≠0: no rate satisfies the equation
			args: numArgs(60, 0, 50000),
		},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnRATE(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			assertError(t, tc.name, v)
		})
	}
}

func TestRATE_PMT_RoundTrip(t *testing.T) {
	// Verify: PMT(RATE(nper, pmt, pv), nper, pv) ≈ pmt
	nper := 360.0
	pmt := -1073.64
	pv := 200000.0

	rateVal, err := fnRATE(numArgs(nper, pmt, pv))
	if err != nil {
		t.Fatal(err)
	}
	if rateVal.Type != ValueNumber {
		t.Fatalf("RATE returned non-number: %v", rateVal)
	}

	pmtVal, err := fnPMT(numArgs(rateVal.Num, nper, pv))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PMT(RATE(...),...) round-trip", pmtVal, pmt)
}

// === IPMT ===

func TestIPMT_FirstPayment(t *testing.T) {
	// IPMT(0.05/12, 1, 360, 200000) — interest portion of first payment
	// Expected: -833.33
	v, err := fnIPMT(numArgs(0.05/12, 1, 360, 200000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "IPMT first", v, -833.33)
}

func TestIPMT_Comprehensive(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool // expect ValueError
	}{
		// --- Basic cases ---
		{
			name: "last payment 30yr mortgage",
			// IPMT(0.05/12, 360, 360, 200000) — interest of last payment is small
			args: numArgs(0.05/12, 360, 360, 200000),
			want: -4.45,
		},
		{
			name: "middle period",
			// IPMT(0.05/12, 180, 360, 200000) — halfway through 30yr mortgage
			args: numArgs(0.05/12, 180, 360, 200000),
			want: -567.81,
		},
		{
			name: "zero rate returns zero interest",
			// IPMT(0, 5, 10, 10000) — no interest when rate=0
			args: numArgs(0, 5, 10, 10000),
			want: 0,
		},

		// --- With future value ---
		{
			name: "with future value",
			// IPMT(0.1/12, 1, 60, 50000, 10000) — loan with balloon payment
			args: numArgs(0.1/12, 1, 60, 50000, 10000),
			want: -416.67,
		},

		// --- Type=0 vs type=1 ---
		{
			name: "type 0 explicit",
			// IPMT(0.1/12, 1, 12, 10000, 0, 0)
			args: numArgs(0.1/12, 1, 12, 10000, 0, 0),
			want: -83.33,
		},
		{
			name: "type 1 period 1 returns 0",
			// IPMT(0.1/12, 1, 12, 10000, 0, 1) — beginning of period, first payment has no interest
			args: numArgs(0.1/12, 1, 12, 10000, 0, 1),
			want: 0,
		},
		{
			name: "type 1 period 2",
			// IPMT(0.1/12, 2, 12, 10000, 0, 1) — beginning of period, second payment
			args: numArgs(0.1/12, 2, 12, 10000, 0, 1),
			want: -76.07,
		},

		// --- Documentation examples ---
		{
			name: "doc example 1 monthly",
			// IPMT(0.1/12, 1, 36, 8000) — from docs
			args: numArgs(0.1/12, 1, 36, 8000),
			want: -66.67,
		},
		{
			name: "doc example 2 annual",
			// IPMT(0.1, 3, 3, 8000) — from docs, annual payments, last year
			args: numArgs(0.1, 3, 3, 8000),
			want: -292.45,
		},

		// --- Monthly vs annual ---
		{
			name: "annual rate 5yr loan first year",
			// IPMT(0.08, 1, 5, 25000)
			args: numArgs(0.08, 1, 5, 25000),
			want: -2000.00,
		},
		{
			name: "monthly equivalent first month",
			// IPMT(0.08/12, 1, 60, 25000)
			args: numArgs(0.08/12, 1, 60, 25000),
			want: -166.67,
		},

		// --- Large loan scenario ---
		{
			name: "large loan first payment",
			// IPMT(0.04/12, 1, 360, 1000000)
			args: numArgs(0.04/12, 1, 360, 1000000),
			want: -3333.33,
		},
		{
			name: "large loan last payment",
			// IPMT(0.04/12, 360, 360, 1000000)
			args: numArgs(0.04/12, 360, 360, 1000000),
			want: -15.86,
		},

		// --- Error cases ---
		{
			name:    "per=0 returns NUM error",
			args:    numArgs(0.05/12, 0, 360, 200000),
			wantErr: true,
		},
		{
			name:    "per > nper returns NUM error",
			args:    numArgs(0.05/12, 361, 360, 200000),
			wantErr: true,
		},
		{
			name:    "nper=0 returns NUM error",
			args:    numArgs(0.05/12, 1, 0, 200000),
			wantErr: true,
		},
		{
			name:    "too few args",
			args:    numArgs(0.05, 1, 12),
			wantErr: true,
		},
		{
			name:    "too many args",
			args:    numArgs(0.05, 1, 12, 10000, 0, 0, 99),
			wantErr: true,
		},
		{
			name:    "non-numeric rate",
			args:    []Value{StringVal("abc"), NumberVal(1), NumberVal(12), NumberVal(10000)},
			wantErr: true,
		},
		{
			name:    "non-numeric per",
			args:    []Value{NumberVal(0.05), StringVal("xyz"), NumberVal(12), NumberVal(10000)},
			wantErr: true,
		},

		// --- Negative pv (investment perspective) ---
		{
			name: "negative pv gives positive interest",
			// IPMT(0.05/12, 1, 360, -200000) — lender perspective
			args: numArgs(0.05/12, 1, 360, -200000),
			want: 833.33,
		},

		// --- Single period ---
		{
			name: "single period nper=1 type=0",
			// IPMT(0.10, 1, 1, 1000) — entire loan repaid in one period
			args: numArgs(0.10, 1, 1, 1000),
			want: -100.00,
		},
		{
			name: "single period nper=1 type=1",
			// IPMT(0.10, 1, 1, 1000, 0, 1) — beginning of period, no interest accrued
			args: numArgs(0.10, 1, 1, 1000, 0, 1),
			want: 0,
		},

		// --- Very high interest rate ---
		{
			name: "very high rate 100% per period",
			// IPMT(1.0, 1, 12, 10000) — extreme rate
			args: numArgs(1.0, 1, 12, 10000),
			want: -10000.00,
		},
		{
			name: "very high rate 50% middle period",
			// IPMT(0.5, 5, 10, 10000)
			args: numArgs(0.5, 5, 10, 10000),
			want: -4641.53,
		},

		// --- Very small interest rate ---
		{
			name: "very small rate 0.001% per period",
			// IPMT(0.00001, 1, 360, 200000)
			args: numArgs(0.00001, 1, 360, 200000),
			want: -2.00,
		},

		// --- Both pv and fv specified ---
		{
			name: "pv and fv both specified first period",
			// IPMT(0.06/12, 1, 60, 10000, 5000)
			args: numArgs(0.06/12, 1, 60, 10000, 5000),
			want: -50.00,
		},
		{
			name: "pv and fv both specified last period",
			// IPMT(0.06/12, 60, 60, 10000, 5000)
			args: numArgs(0.06/12, 60, 60, 10000, 5000),
			want: 23.56,
		},

		// --- fv only (saving with target future value, pv=0) ---
		{
			name: "saving with fv only pv=0 first period",
			// IPMT(0.05/12, 1, 60, 0, 10000)
			args: numArgs(0.05/12, 1, 60, 0, 10000),
			want: 0,
		},
		{
			name: "saving with fv only middle period",
			// IPMT(0.05/12, 30, 60, 0, 10000)
			args: numArgs(0.05/12, 30, 60, 0, 10000),
			want: 18.84,
		},

		// --- Large nper values ---
		{
			name: "large nper 600 months first payment",
			// IPMT(0.03/12, 1, 600, 100000) — 50 year mortgage first payment
			args: numArgs(0.03/12, 1, 600, 100000),
			want: -250.00,
		},
		{
			name: "large nper 600 months last payment",
			// IPMT(0.03/12, 600, 600, 100000) — 50 year mortgage last payment
			args: numArgs(0.03/12, 600, 600, 100000),
			want: -0.80,
		},

		// --- Type=1 with future value ---
		{
			name: "type=1 with fv period 1",
			// IPMT(0.06/12, 1, 60, 10000, 5000, 1) — beginning of period
			args: numArgs(0.06/12, 1, 60, 10000, 5000, 1),
			want: 0,
		},
		{
			name: "type=1 with fv period 2",
			// IPMT(0.06/12, 2, 60, 10000, 5000, 1)
			args: numArgs(0.06/12, 2, 60, 10000, 5000, 1),
			want: -48.68,
		},
		{
			name: "type=1 last period with fv",
			// IPMT(0.06/12, 60, 60, 10000, 5000, 1)
			args: numArgs(0.06/12, 60, 60, 10000, 5000, 1),
			want: 23.44,
		},

		// --- String coercion ---
		{
			name: "string coercion for rate",
			args: []Value{StringVal("0.1"), NumberVal(1), NumberVal(36), NumberVal(8000)},
			want: -800.00,
		},
		{
			name: "string coercion for per",
			args: []Value{NumberVal(0.1 / 12), StringVal("1"), NumberVal(36), NumberVal(8000)},
			want: -66.67,
		},
		{
			name: "string coercion for nper",
			args: []Value{NumberVal(0.1 / 12), NumberVal(1), StringVal("36"), NumberVal(8000)},
			want: -66.67,
		},
		{
			name: "string coercion for pv",
			args: []Value{NumberVal(0.1 / 12), NumberVal(1), NumberVal(36), StringVal("8000")},
			want: -66.67,
		},
		{
			name: "string coercion for fv",
			args: []Value{NumberVal(0.1 / 12), NumberVal(1), NumberVal(60), NumberVal(50000), StringVal("10000")},
			want: -416.67,
		},
		{
			name: "string coercion for type",
			args: []Value{NumberVal(0.1 / 12), NumberVal(1), NumberVal(12), NumberVal(10000), NumberVal(0), StringVal("1")},
			want: 0,
		},

		// --- Boolean coercion ---
		{
			name: "bool TRUE as rate coerced to 1",
			args: []Value{BoolVal(true), NumberVal(1), NumberVal(12), NumberVal(10000)},
			want: -10000.00,
		},
		{
			name: "bool FALSE as rate coerced to 0",
			args: []Value{BoolVal(false), NumberVal(5), NumberVal(10), NumberVal(10000)},
			want: 0,
		},
		{
			name: "bool TRUE as type (beginning of period)",
			args: []Value{NumberVal(0.1 / 12), NumberVal(1), NumberVal(12), NumberVal(10000), NumberVal(0), BoolVal(true)},
			want: 0,
		},
		{
			name: "bool FALSE as type (end of period)",
			args: []Value{NumberVal(0.1 / 12), NumberVal(1), NumberVal(12), NumberVal(10000), NumberVal(0), BoolVal(false)},
			want: -83.33,
		},
		{
			name: "bool TRUE as per coerced to 1",
			args: []Value{NumberVal(0.05 / 12), BoolVal(true), NumberVal(360), NumberVal(200000)},
			want: -833.33,
		},

		// --- Empty cell references (coerce to 0) ---
		{
			name: "empty val for rate coerces to 0",
			args: []Value{EmptyVal(), NumberVal(5), NumberVal(10), NumberVal(10000)},
			want: 0,
		},
		{
			name: "empty val for fv coerces to 0",
			args: []Value{NumberVal(0.1 / 12), NumberVal(1), NumberVal(36), NumberVal(8000), EmptyVal()},
			want: -66.67,
		},
		{
			name: "empty val for type coerces to 0 end of period",
			args: []Value{NumberVal(0.1 / 12), NumberVal(1), NumberVal(12), NumberVal(10000), NumberVal(0), EmptyVal()},
			want: -83.33,
		},

		// --- Error propagation ---
		{
			name:    "error propagation in rate",
			args:    []Value{ErrorVal(ErrValNUM), NumberVal(1), NumberVal(12), NumberVal(10000)},
			wantErr: true,
		},
		{
			name:    "error propagation in per",
			args:    []Value{NumberVal(0.05), ErrorVal(ErrValNA), NumberVal(12), NumberVal(10000)},
			wantErr: true,
		},
		{
			name:    "error propagation in nper",
			args:    []Value{NumberVal(0.05), NumberVal(1), ErrorVal(ErrValREF), NumberVal(10000)},
			wantErr: true,
		},
		{
			name:    "error propagation in pv",
			args:    []Value{NumberVal(0.05), NumberVal(1), NumberVal(12), ErrorVal(ErrValDIV0)},
			wantErr: true,
		},
		{
			name:    "error propagation in fv",
			args:    []Value{NumberVal(0.05), NumberVal(1), NumberVal(12), NumberVal(10000), ErrorVal(ErrValVALUE)},
			wantErr: true,
		},
		{
			name:    "error propagation in type",
			args:    []Value{NumberVal(0.05), NumberVal(1), NumberVal(12), NumberVal(10000), NumberVal(0), ErrorVal(ErrValNUM)},
			wantErr: true,
		},

		// --- Non-numeric errors for remaining args ---
		{
			name:    "non-numeric nper",
			args:    []Value{NumberVal(0.05), NumberVal(1), StringVal("bad"), NumberVal(10000)},
			wantErr: true,
		},
		{
			name:    "non-numeric pv",
			args:    []Value{NumberVal(0.05), NumberVal(1), NumberVal(12), StringVal("bad")},
			wantErr: true,
		},
		{
			name:    "non-numeric fv",
			args:    []Value{NumberVal(0.05), NumberVal(1), NumberVal(12), NumberVal(10000), StringVal("bad")},
			wantErr: true,
		},
		{
			name:    "non-numeric type",
			args:    []Value{NumberVal(0.05), NumberVal(1), NumberVal(12), NumberVal(10000), NumberVal(0), StringVal("bad")},
			wantErr: true,
		},
		{
			name:    "empty string rate not coercible",
			args:    []Value{StringVal(""), NumberVal(1), NumberVal(12), NumberVal(10000)},
			wantErr: true,
		},

		// --- per boundary errors ---
		{
			name:    "per < 1 fractional returns NUM error",
			args:    numArgs(0.05/12, 0.5, 360, 200000),
			wantErr: true,
		},
		{
			name:    "negative per returns NUM error",
			args:    numArgs(0.05/12, -1, 360, 200000),
			wantErr: true,
		},

		// --- Interest decreases over time (amortization) ---
		{
			name: "amortization: period 2 less interest than period 1",
			// IPMT(0.08, 2, 5, 25000) — second year interest less than first
			args: numArgs(0.08, 2, 5, 25000),
			want: -1659.09,
		},
		{
			name: "amortization: last period of 5yr loan",
			// IPMT(0.08, 5, 5, 25000) — last year
			args: numArgs(0.08, 5, 5, 25000),
			want: -463.81,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnIPMT(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
				return
			}
			assertClose(t, tc.name, v, tc.want)
		})
	}
}

// === PPMT ===

func TestPPMT_FirstPayment(t *testing.T) {
	// PPMT(0.05/12, 1, 360, 200000) — principal portion of first payment
	// Expected: -240.31
	v, err := fnPPMT(numArgs(0.05/12, 1, 360, 200000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PPMT first", v, -240.31)
}

func TestPPMT_Comprehensive(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// --- Basic cases ---
		{
			name: "principal of first payment 30yr mortgage",
			// PPMT(0.05/12, 1, 360, 200000)
			args: numArgs(0.05/12, 1, 360, 200000),
			want: -240.31,
		},
		{
			name: "principal of last payment 30yr mortgage",
			// PPMT(0.05/12, 360, 360, 200000) — last payment is mostly principal
			args: numArgs(0.05/12, 360, 360, 200000),
			want: -1069.19,
		},
		{
			name: "middle period 30yr mortgage",
			// PPMT(0.05/12, 180, 360, 200000) — halfway through
			args: numArgs(0.05/12, 180, 360, 200000),
			want: -505.83,
		},

		// --- Zero rate: equal principal payments ---
		{
			name: "zero rate equal principal",
			// PPMT(0, 5, 10, 10000) — each payment is pv/nper = -1000
			args: numArgs(0, 5, 10, 10000),
			want: -1000.00,
		},
		{
			name: "zero rate first period",
			// PPMT(0, 1, 10, 10000)
			args: numArgs(0, 1, 10, 10000),
			want: -1000.00,
		},

		// --- With future value ---
		{
			name: "with future value",
			// PPMT(0.1/12, 1, 60, 50000, 10000)
			args: numArgs(0.1/12, 1, 60, 50000, 10000),
			want: -774.82,
		},
		{
			name: "with future value last period",
			// PPMT(0.1/12, 60, 60, 50000, 10000)
			args: numArgs(0.1/12, 60, 60, 50000, 10000),
			want: -1264.29,
		},

		// --- Type=0 vs type=1 ---
		{
			name: "type 0 explicit",
			// PPMT(0.1/12, 1, 12, 10000, 0, 0)
			args: numArgs(0.1/12, 1, 12, 10000, 0, 0),
			want: -795.83,
		},
		{
			name: "type 1 period 1 equals full PMT",
			// PPMT(0.1/12, 1, 12, 10000, 0, 1) — at beginning, first period PPMT = PMT
			args: numArgs(0.1/12, 1, 12, 10000, 0, 1),
			want: -871.89,
		},
		{
			name: "type 1 period 2",
			// PPMT(0.1/12, 2, 12, 10000, 0, 1)
			args: numArgs(0.1/12, 2, 12, 10000, 0, 1),
			want: -795.83,
		},

		// --- Documentation examples ---
		{
			name: "doc example 1: 10% 2yr loan month 1",
			// PPMT(0.10/12, 1, 2*12, 2000) = -75.62
			args: numArgs(0.10/12, 1, 24, 2000),
			want: -75.62,
		},
		{
			name: "doc example 2: 8% 10yr loan year 10",
			// PPMT(0.08, 10, 10, 200000) = -27598.05
			args: numArgs(0.08, 10, 10, 200000),
			want: -27598.05,
		},

		// --- Monthly vs annual ---
		{
			name: "annual rate 5yr loan first year",
			// PPMT(0.08, 1, 5, 25000)
			args: numArgs(0.08, 1, 5, 25000),
			want: -4261.41,
		},
		{
			name: "monthly equivalent first month",
			// PPMT(0.08/12, 1, 60, 25000)
			args: numArgs(0.08/12, 1, 60, 25000),
			want: -340.24,
		},

		// --- Large loan ---
		{
			name: "large loan first payment",
			// PPMT(0.04/12, 1, 360, 1000000)
			args: numArgs(0.04/12, 1, 360, 1000000),
			want: -1440.82,
		},
		{
			name: "large loan last payment",
			// PPMT(0.04/12, 360, 360, 1000000)
			args: numArgs(0.04/12, 360, 360, 1000000),
			want: -4758.29,
		},

		// --- Error cases ---
		{
			name:    "per=0 returns NUM error",
			args:    numArgs(0.05/12, 0, 360, 200000),
			wantErr: true,
		},
		{
			name:    "per > nper returns NUM error",
			args:    numArgs(0.05/12, 361, 360, 200000),
			wantErr: true,
		},
		{
			name:    "negative per returns NUM error",
			args:    numArgs(0.05/12, -1, 360, 200000),
			wantErr: true,
		},
		{
			name:    "nper=0 returns NUM error",
			args:    numArgs(0.05/12, 1, 0, 200000),
			wantErr: true,
		},
		{
			name:    "too few args",
			args:    numArgs(0.05, 1, 12),
			wantErr: true,
		},
		{
			name:    "too many args",
			args:    numArgs(0.05, 1, 12, 10000, 0, 0, 99),
			wantErr: true,
		},
		{
			name:    "non-numeric rate",
			args:    []Value{StringVal("abc"), NumberVal(1), NumberVal(12), NumberVal(10000)},
			wantErr: true,
		},
		{
			name:    "non-numeric per",
			args:    []Value{NumberVal(0.05), StringVal("xyz"), NumberVal(12), NumberVal(10000)},
			wantErr: true,
		},

		// --- Negative pv (investment perspective) ---
		{
			name: "negative pv gives positive principal",
			// PPMT(0.05/12, 1, 360, -200000) — lender perspective
			args: numArgs(0.05/12, 1, 360, -200000),
			want: 240.31,
		},

		// --- Single period ---
		{
			name: "single period nper=1 type=0",
			// PPMT(0.10, 1, 1, 1000) — entire loan repaid in one period
			args: numArgs(0.10, 1, 1, 1000),
			want: -1000.00,
		},
		{
			name: "single period nper=1 type=1",
			// PPMT(0.10, 1, 1, 1000, 0, 1) — beginning of period, PPMT=PMT
			args: numArgs(0.10, 1, 1, 1000, 0, 1),
			want: -1000.00,
		},

		// --- Very high interest rate ---
		{
			name: "very high rate 100% per period",
			// PPMT(1.0, 1, 12, 10000) — extreme rate
			args: numArgs(1.0, 1, 12, 10000),
			want: -2.44,
		},
		{
			name: "very high rate 50% middle period",
			// PPMT(0.5, 5, 10, 10000)
			args: numArgs(0.5, 5, 10, 10000),
			want: -446.70,
		},

		// --- Very small interest rate ---
		{
			name: "very small rate 0.001% per period",
			// PPMT(0.00001, 1, 360, 200000)
			args: numArgs(0.00001, 1, 360, 200000),
			want: -554.56,
		},

		// --- Both pv and fv specified ---
		{
			name: "pv and fv both specified first period",
			// PPMT(0.06/12, 1, 60, 10000, 5000)
			args: numArgs(0.06/12, 1, 60, 10000, 5000),
			want: -214.99,
		},
		{
			name: "pv and fv both specified last period",
			// PPMT(0.06/12, 60, 60, 10000, 5000)
			args: numArgs(0.06/12, 60, 60, 10000, 5000),
			want: -288.55,
		},

		// --- fv only (saving with target future value, pv=0) ---
		{
			name: "saving with fv only pv=0 first period",
			// PPMT(0.05/12, 1, 60, 0, 10000)
			args: numArgs(0.05/12, 1, 60, 0, 10000),
			want: -147.05,
		},
		{
			name: "saving with fv only middle period",
			// PPMT(0.05/12, 30, 60, 0, 10000)
			args: numArgs(0.05/12, 30, 60, 0, 10000),
			want: -165.89,
		},

		// --- Large nper values ---
		{
			name: "large nper 600 months first payment",
			// PPMT(0.03/12, 1, 600, 100000) — 50 year mortgage first payment
			args: numArgs(0.03/12, 1, 600, 100000),
			want: -71.98,
		},
		{
			name: "large nper 600 months last payment",
			// PPMT(0.03/12, 600, 600, 100000) — 50 year mortgage last payment
			args: numArgs(0.03/12, 600, 600, 100000),
			want: -321.17,
		},

		// --- Type=1 with future value ---
		{
			name: "type=1 with fv period 1",
			// PPMT(0.06/12, 1, 60, 10000, 5000, 1) — beginning of period, PPMT=PMT
			args: numArgs(0.06/12, 1, 60, 10000, 5000, 1),
			want: -263.67,
		},
		{
			name: "type=1 with fv period 2",
			// PPMT(0.06/12, 2, 60, 10000, 5000, 1)
			args: numArgs(0.06/12, 2, 60, 10000, 5000, 1),
			want: -214.99,
		},
		{
			name: "type=1 last period with fv",
			// PPMT(0.06/12, 60, 60, 10000, 5000, 1)
			args: numArgs(0.06/12, 60, 60, 10000, 5000, 1),
			want: -287.11,
		},

		// --- String coercion ---
		{
			name: "string coercion for rate",
			// PPMT("0.1", 1, 12, 10000, 0, 0)
			args: []Value{StringVal("0.1"), NumberVal(1), NumberVal(12), NumberVal(10000), NumberVal(0), NumberVal(0)},
			want: -467.63,
		},
		{
			name: "string coercion for per",
			args: []Value{NumberVal(0.05 / 12), StringVal("1"), NumberVal(360), NumberVal(200000)},
			want: -240.31,
		},
		{
			name: "string coercion for nper",
			args: []Value{NumberVal(0.05 / 12), NumberVal(1), StringVal("360"), NumberVal(200000)},
			want: -240.31,
		},
		{
			name: "string coercion for pv",
			args: []Value{NumberVal(0.05 / 12), NumberVal(1), NumberVal(360), StringVal("200000")},
			want: -240.31,
		},
		{
			name: "string coercion for fv",
			args: []Value{NumberVal(0.1 / 12), NumberVal(1), NumberVal(60), NumberVal(50000), StringVal("10000")},
			want: -774.82,
		},
		{
			name: "string coercion for type",
			args: []Value{NumberVal(0.1 / 12), NumberVal(1), NumberVal(12), NumberVal(10000), NumberVal(0), StringVal("1")},
			want: -871.89,
		},

		// --- Boolean coercion ---
		{
			name: "bool TRUE as rate coerced to 1",
			args: []Value{BoolVal(true), NumberVal(1), NumberVal(12), NumberVal(10000)},
			want: -2.44,
		},
		{
			name: "bool FALSE as rate coerced to 0",
			args: []Value{BoolVal(false), NumberVal(5), NumberVal(10), NumberVal(10000)},
			want: -1000.00,
		},
		{
			name: "bool TRUE as type (beginning of period)",
			args: []Value{NumberVal(0.1 / 12), NumberVal(1), NumberVal(12), NumberVal(10000), NumberVal(0), BoolVal(true)},
			want: -871.89,
		},
		{
			name: "bool FALSE as type (end of period)",
			args: []Value{NumberVal(0.1 / 12), NumberVal(1), NumberVal(12), NumberVal(10000), NumberVal(0), BoolVal(false)},
			want: -795.83,
		},
		{
			name: "bool TRUE as per coerced to 1",
			args: []Value{NumberVal(0.05 / 12), BoolVal(true), NumberVal(360), NumberVal(200000)},
			want: -240.31,
		},

		// --- Empty cell references (coerce to 0) ---
		{
			name: "empty val for rate coerces to 0",
			args: []Value{EmptyVal(), NumberVal(5), NumberVal(10), NumberVal(10000)},
			want: -1000.00,
		},
		{
			name: "empty val for fv coerces to 0",
			args: []Value{NumberVal(0.05 / 12), NumberVal(1), NumberVal(360), NumberVal(200000), EmptyVal()},
			want: -240.31,
		},
		{
			name: "empty val for type coerces to 0 end of period",
			args: []Value{NumberVal(0.1 / 12), NumberVal(1), NumberVal(12), NumberVal(10000), NumberVal(0), EmptyVal()},
			want: -795.83,
		},

		// --- Error propagation ---
		{
			name:    "error propagation in rate",
			args:    []Value{ErrorVal(ErrValNUM), NumberVal(1), NumberVal(12), NumberVal(10000)},
			wantErr: true,
		},
		{
			name:    "error propagation in per",
			args:    []Value{NumberVal(0.05), ErrorVal(ErrValNA), NumberVal(12), NumberVal(10000)},
			wantErr: true,
		},
		{
			name:    "error propagation in nper",
			args:    []Value{NumberVal(0.05), NumberVal(1), ErrorVal(ErrValREF), NumberVal(10000)},
			wantErr: true,
		},
		{
			name:    "error propagation in pv",
			args:    []Value{NumberVal(0.05), NumberVal(1), NumberVal(12), ErrorVal(ErrValDIV0)},
			wantErr: true,
		},
		{
			name:    "error propagation in fv",
			args:    []Value{NumberVal(0.05), NumberVal(1), NumberVal(12), NumberVal(10000), ErrorVal(ErrValVALUE)},
			wantErr: true,
		},
		{
			name:    "error propagation in type",
			args:    []Value{NumberVal(0.05), NumberVal(1), NumberVal(12), NumberVal(10000), NumberVal(0), ErrorVal(ErrValNUM)},
			wantErr: true,
		},

		// --- Non-numeric errors for remaining args ---
		{
			name:    "non-numeric nper",
			args:    []Value{NumberVal(0.05), NumberVal(1), StringVal("bad"), NumberVal(10000)},
			wantErr: true,
		},
		{
			name:    "non-numeric pv",
			args:    []Value{NumberVal(0.05), NumberVal(1), NumberVal(12), StringVal("bad")},
			wantErr: true,
		},
		{
			name:    "non-numeric fv",
			args:    []Value{NumberVal(0.05), NumberVal(1), NumberVal(12), NumberVal(10000), StringVal("bad")},
			wantErr: true,
		},
		{
			name:    "non-numeric type",
			args:    []Value{NumberVal(0.05), NumberVal(1), NumberVal(12), NumberVal(10000), NumberVal(0), StringVal("bad")},
			wantErr: true,
		},
		{
			name:    "empty string rate not coercible",
			args:    []Value{StringVal(""), NumberVal(1), NumberVal(12), NumberVal(10000)},
			wantErr: true,
		},

		// --- per boundary errors ---
		{
			name:    "per < 1 fractional returns NUM error",
			args:    numArgs(0.05/12, 0.5, 360, 200000),
			wantErr: true,
		},

		// --- Amortization: principal increases over time ---
		{
			name: "amortization: period 2 more principal than period 1",
			// PPMT(0.08, 2, 5, 25000)
			args: numArgs(0.08, 2, 5, 25000),
			want: -4602.32,
		},
		{
			name: "amortization: last period of 5yr loan",
			// PPMT(0.08, 5, 5, 25000)
			args: numArgs(0.08, 5, 5, 25000),
			want: -5797.60,
		},

		// --- Zero rate with fv ---
		{
			name: "zero rate with fv",
			// PPMT(0, 3, 10, 10000, 5000) — principal = -(pv+fv)/nper = -1500
			args: numArgs(0, 3, 10, 10000, 5000),
			want: -1500.00,
		},

		// --- Zero rate with type=1 ---
		{
			name: "zero rate with type 1",
			// PPMT(0, 1, 10, 10000, 0, 1)
			args: numArgs(0, 1, 10, 10000, 0, 1),
			want: -1000.00,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnPPMT(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
				return
			}
			assertClose(t, tc.name, v, tc.want)
		})
	}
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

func TestIPMT_Plus_PPMT_Equals_PMT_WithFV(t *testing.T) {
	// Verify PPMT + IPMT = PMT with future value and type=1
	rate := 0.06 / 12
	nper := 120.0
	pv := 100000.0
	fv := 20000.0
	for _, payType := range []float64{0, 1} {
		for _, per := range []float64{1, 30, 60, 90, 120} {
			ipmt, _ := fnIPMT(numArgs(rate, per, nper, pv, fv, payType))
			ppmt, _ := fnPPMT(numArgs(rate, per, nper, pv, fv, payType))
			pmt, _ := fnPMT(numArgs(rate, nper, pv, fv, payType))
			sum := ipmt.Num + ppmt.Num
			if math.Abs(sum-pmt.Num) > 0.01 {
				t.Errorf("type=%v per=%v: IPMT(%f) + PPMT(%f) = %f, PMT = %f",
					payType, per, ipmt.Num, ppmt.Num, sum, pmt.Num)
			}
		}
	}
}

// === NPV ===

func TestNPV_Basic(t *testing.T) {
	// NPV(0.1, -10000, 3000, 4200, 6800)
	// Expected: 1188.44
	v, err := fnNPV(append(numArgs(0.1), numArgs(-10000, 3000, 4200, 6800)...))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV basic", v, 1188.44)
}

func TestNPV_PositiveCashFlows(t *testing.T) {
	// NPV(0.08, 1000, 2000, 3000)
	// = 1000/1.08 + 2000/1.08^2 + 3000/1.08^3 = 925.93 + 1714.68 + 2381.50 = 5022.11
	v, err := fnNPV(append(numArgs(0.08), numArgs(1000, 2000, 3000)...))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV positive cash flows", v, 5022.10)
}

func TestNPV_MixedCashFlows(t *testing.T) {
	// NPV(0.1, -5000, 2000, -1000, 4000, 3000)
	v, err := fnNPV(append(numArgs(0.1), numArgs(-5000, 2000, -1000, 4000, 3000)...))
	if err != nil {
		t.Fatal(err)
	}
	// -5000/1.1 + 2000/1.1^2 + -1000/1.1^3 + 4000/1.1^4 + 3000/1.1^5
	// = -4545.45 + 1652.89 + -751.31 + 2732.05 + 1862.76 = 950.94
	assertClose(t, "NPV mixed cash flows", v, 950.94)
}

func TestNPV_SingleCashFlow(t *testing.T) {
	// NPV(0.1, 1000) = 1000/1.1 = 909.09
	v, err := fnNPV(append(numArgs(0.1), numArgs(1000)...))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV single cash flow", v, 909.09)
}

func TestNPV_ZeroRate(t *testing.T) {
	// NPV(0, 1000, 2000, 3000) = 1000 + 2000 + 3000 = 6000
	v, err := fnNPV(append(numArgs(0), numArgs(1000, 2000, 3000)...))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV zero rate", v, 6000.00)
}

func TestNPV_HighRate(t *testing.T) {
	// NPV(1.0, 1000, 1000, 1000) = 1000/2 + 1000/4 + 1000/8 = 500 + 250 + 125 = 875
	v, err := fnNPV(append(numArgs(1.0), numArgs(1000, 1000, 1000)...))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV high rate", v, 875.00)
}

func TestNPV_NegativeRate(t *testing.T) {
	// NPV(-0.1, 1000, 1000, 1000)
	// = 1000/0.9 + 1000/0.81 + 1000/0.729 = 1111.11 + 1234.57 + 1371.74 = 3717.42
	v, err := fnNPV(append(numArgs(-0.1), numArgs(1000, 1000, 1000)...))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV negative rate", v, 3717.42)
}

func TestNPV_RateMinusOne(t *testing.T) {
	// NPV(-1, ...) → #DIV/0! because (1+rate)=0
	v, _ := fnNPV(append(numArgs(-1), numArgs(1000)...))
	assertError(t, "NPV rate=-1", v)
}

func TestNPV_ManyPeriods(t *testing.T) {
	// NPV(0.05, 100 repeated 20 times)
	// = sum of 100/1.05^i for i=1..20 = 100 * (1 - 1.05^-20) / 0.05 = 1246.22
	args := numArgs(0.05)
	for i := 0; i < 20; i++ {
		args = append(args, NumberVal(100))
	}
	v, err := fnNPV(args)
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV many periods", v, 1246.22)
}

func TestNPV_RangeInput(t *testing.T) {
	// NPV(0.1, {-10000, 3000, 4200, 6800}) — values passed as array
	arr := Value{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(-10000), NumberVal(3000), NumberVal(4200), NumberVal(6800)}},
	}
	v, err := fnNPV([]Value{NumberVal(0.1), arr})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV range input", v, 1188.44)
}

func TestNPV_AllNegativeCashFlows(t *testing.T) {
	// NPV(0.1, -1000, -2000, -3000)
	// = -1000/1.1 + -2000/1.21 + -3000/1.331 = -909.09 + -1652.89 + -2253.94 = -4815.93
	v, err := fnNPV(append(numArgs(0.1), numArgs(-1000, -2000, -3000)...))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV all negative", v, -4815.93)
}

func TestNPV_ZeroCashFlowsMixedIn(t *testing.T) {
	// NPV(0.1, 0, 1000, 0, 2000) = 0/1.1 + 1000/1.21 + 0/1.331 + 2000/1.4641
	// = 0 + 826.45 + 0 + 1366.03 = 2192.47
	v, err := fnNPV(append(numArgs(0.1), numArgs(0, 1000, 0, 2000)...))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV zero mixed in", v, 2192.47)
}

func TestNPV_TooFewArgs(t *testing.T) {
	// NPV(0.1) — no values → #VALUE!
	v, _ := fnNPV(numArgs(0.1))
	assertError(t, "NPV too few args", v)
}

func TestNPV_NoArgs(t *testing.T) {
	// NPV() — no args at all → #VALUE!
	v, _ := fnNPV([]Value{})
	assertError(t, "NPV no args", v)
}

func TestNPV_StringCoercionRate(t *testing.T) {
	// NPV("0.1", -10000, 3000, 4200, 6800) — rate as numeric string
	args := []Value{StringVal("0.1"), NumberVal(-10000), NumberVal(3000), NumberVal(4200), NumberVal(6800)}
	v, err := fnNPV(args)
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV string rate", v, 1188.44)
}

func TestNPV_StringCoercionValue(t *testing.T) {
	// NPV(0.1, "1000") — value as numeric string (direct arg, coerced)
	args := []Value{NumberVal(0.1), StringVal("1000")}
	v, err := fnNPV(args)
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV string value", v, 909.09)
}

func TestNPV_BooleanCoercion(t *testing.T) {
	// NPV(0.1, TRUE) — TRUE coerced to 1, so NPV = 1/1.1 = 0.909..
	args := []Value{NumberVal(0.1), BoolVal(true)}
	v, err := fnNPV(args)
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV bool TRUE", v, 0.91)
}

func TestNPV_ErrorPropagationRate(t *testing.T) {
	// NPV(#VALUE!, 1000) → error
	args := []Value{ErrorVal(ErrValVALUE), NumberVal(1000)}
	v, _ := fnNPV(args)
	assertError(t, "NPV error rate", v)
}

func TestNPV_ErrorPropagationValue(t *testing.T) {
	// NPV(0.1, #DIV/0!) → error
	args := []Value{NumberVal(0.1), ErrorVal(ErrValDIV0)}
	v, _ := fnNPV(args)
	assertError(t, "NPV error value", v)
}

func TestNPV_NonNumericRate(t *testing.T) {
	// NPV("abc", 1000) → #VALUE!
	args := []Value{StringVal("abc"), NumberVal(1000)}
	v, _ := fnNPV(args)
	assertError(t, "NPV non-numeric rate", v)
}

func TestNPV_NonNumericValue(t *testing.T) {
	// NPV(0.1, "abc") — non-numeric string as direct arg → #VALUE!
	args := []Value{NumberVal(0.1), StringVal("abc")}
	v, _ := fnNPV(args)
	assertError(t, "NPV non-numeric value", v)
}

func TestNPV_DocExample2(t *testing.T) {
	// Doc example 2: NPV(0.08, 8000, 9200, 10000, 12000, 14500) + (-40000)
	// = NPV(0.08, 8000, 9200, 10000, 12000, 14500) + (-40000) = 1922.06
	v, err := fnNPV(append(numArgs(0.08), numArgs(8000, 9200, 10000, 12000, 14500)...))
	if err != nil {
		t.Fatal(err)
	}
	// NPV alone is 41922.06, then + (-40000) = 1922.06
	assertClose(t, "NPV doc example 2", v, 41922.06)
}

func TestNPV_DocExample2WithLoss(t *testing.T) {
	// Doc example 2 with loss: NPV(0.08, 8000, 9200, 10000, 12000, 14500, -9000) + (-40000)
	// = -3749.47, so NPV alone = -3749.47 + 40000 = 36250.53
	v, err := fnNPV(append(numArgs(0.08), numArgs(8000, 9200, 10000, 12000, 14500, -9000)...))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NPV doc example 2 with loss", v, 36250.53)
}

func TestNPV_InitialInvestmentPattern(t *testing.T) {
	// Common pattern: initial cost + NPV of future cash flows
	// NPV(0.08, 8000, 9200, 10000, 12000, 14500) + (-40000) = 1922.06
	npv, err := fnNPV(append(numArgs(0.08), numArgs(8000, 9200, 10000, 12000, 14500)...))
	if err != nil {
		t.Fatal(err)
	}
	result := npv.Num + (-40000)
	if math.Abs(result-1922.06) > 0.01 {
		t.Errorf("NPV initial investment pattern: got %f, want 1922.06", result)
	}
}

func TestNPV_EmptyValueSkipped(t *testing.T) {
	// Empty values in array should be skipped (period not advanced)
	// NPV(0.1, {1000, <empty>, 2000}) should give same result as NPV(0.1, 1000, 2000)
	arr := Value{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(1000), EmptyVal(), NumberVal(2000)}},
	}
	v, err := fnNPV([]Value{NumberVal(0.1), arr})
	if err != nil {
		t.Fatal(err)
	}
	// 1000/1.1 + 2000/1.21 = 909.09 + 1652.89 = 2561.98
	v2, _ := fnNPV(append(numArgs(0.1), numArgs(1000, 2000)...))
	assertClose(t, "NPV empty skipped", v, v2.Num)
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

func TestIRR_AllNegative(t *testing.T) {
	// All negative values — no solution → #NUM!
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-100), NumberVal(-200), NumberVal(-300)},
		},
	}
	v, _ := fnIRR([]Value{arr})
	assertError(t, "IRR all negative", v)
}

func TestIRR_SingleValue(t *testing.T) {
	// Single value → #NUM! (need at least 2 cash flows)
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000)},
		},
	}
	v, _ := fnIRR([]Value{arr})
	assertError(t, "IRR single value", v)
}

func TestIRR_TwoValues(t *testing.T) {
	// Two values: invest -1000, return 1100 → IRR = 0.10
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-1000), NumberVal(1100)},
		},
	}
	v, err := fnIRR([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "IRR two values", v, 0.10)
}

func TestIRR_WithExplicitGuess(t *testing.T) {
	// Same as basic but with explicit guess of 0.05
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000), NumberVal(3000), NumberVal(4200), NumberVal(6800)},
		},
	}
	v, err := fnIRR([]Value{arr, NumberVal(0.05)})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "IRR with guess", v, 0.1634)
}

func TestIRR_DefaultGuess(t *testing.T) {
	// Without guess, default is 0.1. Same result as with explicit guess.
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000), NumberVal(3000), NumberVal(4200), NumberVal(6800)},
		},
	}
	v1, err := fnIRR([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	v2, err := fnIRR([]Value{arr, NumberVal(0.1)})
	if err != nil {
		t.Fatal(err)
	}
	if math.Abs(v1.Num-v2.Num) > 1e-9 {
		t.Errorf("IRR default guess: results differ: %f vs %f", v1.Num, v2.Num)
	}
}

func TestIRR_HighReturnRate(t *testing.T) {
	// Invest -100, get 500 back → IRR = 4.0 (400%)
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-100), NumberVal(500)},
		},
	}
	v, err := fnIRR([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "IRR high return", v, 4.0)
}

func TestIRR_LowReturnRate(t *testing.T) {
	// Invest -10000, get 10010 back in 10 periods → very low IRR
	flows := make([]Value, 11)
	flows[0] = NumberVal(-10000)
	for i := 1; i <= 10; i++ {
		flows[i] = NumberVal(1001)
	}
	arr := Value{
		Type:  ValueArray,
		Array: [][]Value{flows},
	}
	v, err := fnIRR([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	// IRR should be very close to 0, slightly positive
	if v.Type != ValueNumber {
		t.Fatalf("IRR low return: expected number, got %v", v.Type)
	}
	if v.Num < -0.01 || v.Num > 0.02 {
		t.Errorf("IRR low return: got %f, want near 0", v.Num)
	}
}

func TestIRR_BreakEven(t *testing.T) {
	// Invest -1000, return 1000 → IRR = 0
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-1000), NumberVal(1000)},
		},
	}
	v, err := fnIRR([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "IRR break even", v, 0.0)
}

func TestIRR_MonthlyCashFlows(t *testing.T) {
	// Monthly: invest -10000, receive 900/month for 12 months
	flows := make([]Value, 13)
	flows[0] = NumberVal(-10000)
	for i := 1; i <= 12; i++ {
		flows[i] = NumberVal(900)
	}
	arr := Value{
		Type:  ValueArray,
		Array: [][]Value{flows},
	}
	v, err := fnIRR([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	// Monthly IRR for these cash flows (~1.6% per month)
	if v.Type != ValueNumber {
		t.Fatalf("IRR monthly: expected number, got %v", v.Type)
	}
	if v.Num < 0.01 || v.Num > 0.05 {
		t.Errorf("IRR monthly: got %f, want between 0.01 and 0.05", v.Num)
	}
}

func TestIRR_IrregularCashFlows(t *testing.T) {
	// Irregular: negative initial, mix of positive and negative later
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-5000), NumberVal(1000), NumberVal(-500), NumberVal(3000), NumberVal(4000)},
		},
	}
	v, err := fnIRR([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("IRR irregular: expected number, got %v", v.Type)
	}
	// Just verify it returns a valid rate (positive for profitable flows)
	if v.Num < 0 || v.Num > 1 {
		t.Errorf("IRR irregular: got %f, want reasonable rate", v.Num)
	}
}

func TestIRR_LargeNumberOfPeriods(t *testing.T) {
	// 50 periods: invest -50000, receive 1200/period
	flows := make([]Value, 51)
	flows[0] = NumberVal(-50000)
	for i := 1; i <= 50; i++ {
		flows[i] = NumberVal(1200)
	}
	arr := Value{
		Type:  ValueArray,
		Array: [][]Value{flows},
	}
	v, err := fnIRR([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("IRR large periods: expected number, got %v", v.Type)
	}
	// Should be a small positive rate (~1.2%)
	if v.Num < 0 || v.Num > 0.05 {
		t.Errorf("IRR large periods: got %f, want small positive rate", v.Num)
	}
}

func TestIRR_TooFewArgs(t *testing.T) {
	v, _ := fnIRR([]Value{})
	assertError(t, "IRR too few args", v)
}

func TestIRR_TooManyArgs(t *testing.T) {
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-1000), NumberVal(1100)},
		},
	}
	v, _ := fnIRR([]Value{arr, NumberVal(0.1), NumberVal(0.2)})
	assertError(t, "IRR too many args", v)
}

func TestIRR_DocExample1(t *testing.T) {
	// Doc: IRR({-70000, 12000, 15000, 18000, 21000}) = -2.1% (-0.021)
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-70000), NumberVal(12000), NumberVal(15000), NumberVal(18000), NumberVal(21000)},
		},
	}
	v, err := fnIRR([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "IRR doc 4yr", v, -0.02)
}

func TestIRR_DocExample2(t *testing.T) {
	// Doc: IRR({-70000, 12000, 15000, 18000, 21000, 26000}) = 8.7% (0.087)
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-70000), NumberVal(12000), NumberVal(15000), NumberVal(18000), NumberVal(21000), NumberVal(26000)},
		},
	}
	v, err := fnIRR([]Value{arr})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "IRR doc 5yr", v, 0.087)
}

func TestIRR_DocExample3(t *testing.T) {
	// Doc: IRR({-70000, 12000, 15000}, -10%) = -44.4% (-0.444)
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-70000), NumberVal(12000), NumberVal(15000)},
		},
	}
	v, err := fnIRR([]Value{arr, NumberVal(-0.10)})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "IRR doc 2yr with guess", v, -0.44)
}

func TestIRR_EmptyGuessIgnored(t *testing.T) {
	// Empty guess should be treated as default (0.1)
	arr := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000), NumberVal(3000), NumberVal(4200), NumberVal(6800)},
		},
	}
	v, err := fnIRR([]Value{arr, {Type: ValueEmpty}})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "IRR empty guess", v, 0.1634)
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

func TestSLN_Comprehensive(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// Documentation example
		{"doc example", numArgs(30000, 7500, 10), 2250, false},
		// Zero salvage value (full depreciation)
		{"zero salvage", numArgs(30000, 0, 10), 3000, false},
		// Salvage equals cost → zero depreciation
		{"salvage equals cost", numArgs(10000, 10000, 5), 0, false},
		// Salvage > cost → negative depreciation
		{"salvage greater than cost", numArgs(5000, 8000, 10), -300, false},
		// Life = 1 (single period depreciation)
		{"life is 1", numArgs(10000, 2000, 1), 8000, false},
		// Large values (expensive assets)
		{"large values", numArgs(1e9, 1e6, 20), 4.995e7, false},
		// Small decimals (cheap assets)
		{"small decimals", numArgs(0.50, 0.10, 4), 0.10, false},
		// Zero cost, zero salvage
		{"zero cost", numArgs(0, 0, 10), 0, false},
		// Zero cost with salvage → negative
		{"zero cost with salvage", numArgs(0, 5000, 10), -500, false},
		// Negative cost
		{"negative cost", numArgs(-10000, 2000, 5), -2400, false},
		// Fractional life
		{"fractional life", numArgs(10000, 0, 2.5), 4000, false},
		// Large life (100+ periods)
		{"large life", numArgs(10000, 0, 1000), 10, false},
		// Negative salvage
		{"negative salvage", numArgs(10000, -2000, 5), 2400, false},

		// --- Additional cases ---

		// Very large life (200 periods)
		{"very large life 200", numArgs(50000, 5000, 200), 225, false},
		// Both cost and salvage negative
		{"both negative", numArgs(-5000, -2000, 10), -300, false},
		// Negative life (computes normally, just negative result sign flip)
		{"negative life", numArgs(10000, 2000, -5), -1600, false},
		// Fractional life less than 1
		{"fractional life less than 1", numArgs(1000, 0, 0.5), 2000, false},
		// Fractional cost and salvage
		{"fractional cost and salvage", numArgs(1234.56, 234.56, 5), 200, false},
		// Very small positive life
		{"very small life", numArgs(1000, 0, 0.001), 1e6, false},
		// Cost much larger than salvage
		{"cost much larger", numArgs(1e10, 0, 10), 1e9, false},
		// Salvage much larger than cost
		{"salvage much larger", numArgs(0, 1e10, 10), -1e9, false},
		// Life with many decimal places
		{"life many decimals", numArgs(10000, 0, 3.333), 3000.30003000, false},
		// Typical car depreciation
		{"car depreciation", numArgs(25000, 5000, 5), 4000, false},
		// Typical equipment depreciation
		{"equipment depreciation", numArgs(50000, 10000, 7), 5714.285714, false},
		// Penny values
		{"penny values", numArgs(0.01, 0.001, 1), 0.009, false},
		// Life = cost (unusual but valid)
		{"life equals cost", numArgs(100, 0, 100), 1, false},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnSLN(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestSLN_Errors(t *testing.T) {
	tests := []struct {
		name string
		args []Value
	}{
		// Life = 0 → #DIV/0!
		{"life zero", numArgs(30000, 7500, 0)},
		// Too few args
		{"too few args 0", []Value{}},
		{"too few args 1", numArgs(30000)},
		{"too few args 2", numArgs(30000, 7500)},
		// Too many args
		{"too many args", numArgs(30000, 7500, 10, 1)},
		// Non-numeric string
		{"non-numeric cost", []Value{StringVal("abc"), NumberVal(7500), NumberVal(10)}},
		{"non-numeric salvage", []Value{NumberVal(30000), StringVal("xyz"), NumberVal(10)}},
		{"non-numeric life", []Value{NumberVal(30000), NumberVal(7500), StringVal("abc")}},
		// Error propagation
		{"error in cost", []Value{ErrorVal(ErrValNUM), NumberVal(7500), NumberVal(10)}},
		{"error in salvage", []Value{NumberVal(30000), ErrorVal(ErrValDIV0), NumberVal(10)}},
		{"error in life", []Value{NumberVal(30000), NumberVal(7500), ErrorVal(ErrValVALUE)}},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnSLN(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			assertError(t, tc.name, v)
		})
	}
}

func TestSLN_StringCoercion(t *testing.T) {
	// Numeric strings should be coerced
	v, err := fnSLN([]Value{StringVal("30000"), StringVal("7500"), StringVal("10")})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SLN string coercion", v, 2250)
}

func TestSLN_BoolCoercion(t *testing.T) {
	// TRUE=1, FALSE=0: SLN(TRUE, FALSE, TRUE) = SLN(1,0,1) = 1
	v, err := fnSLN([]Value{boolArg(true), boolArg(false), boolArg(true)})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SLN bool coercion", v, 1)
}

func TestSLN_MixedCoercion(t *testing.T) {
	tests := []struct {
		name string
		args []Value
		want float64
	}{
		// Mix of number and string args
		{"number and string mix", []Value{NumberVal(10000), StringVal("2000"), NumberVal(5)}, 1600},
		// Bool as cost, number as others
		{"bool cost true", []Value{boolArg(true), NumberVal(0), NumberVal(1)}, 1},
		// Bool as salvage
		{"bool salvage false", []Value{NumberVal(10), boolArg(false), NumberVal(2)}, 5},
		// Bool as life (TRUE = 1)
		{"bool life true", []Value{NumberVal(100), NumberVal(0), boolArg(true)}, 100},
		// String with decimal
		{"string decimal cost", []Value{StringVal("1500.50"), NumberVal(500.50), NumberVal(10)}, 100},
		// All bools
		{"all true", []Value{boolArg(true), boolArg(true), boolArg(true)}, 0},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnSLN(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			assertClose(t, tc.name, v, tc.want)
		})
	}
}

func TestSLN_BoolLifeZero(t *testing.T) {
	// FALSE as life → life=0 → #DIV/0!
	v, err := fnSLN([]Value{NumberVal(1000), NumberVal(0), boolArg(false)})
	if err != nil {
		t.Fatal(err)
	}
	assertError(t, "SLN bool life false div0", v)
}

func TestSLN_NonNumericStrings(t *testing.T) {
	tests := []struct {
		name string
		args []Value
	}{
		{"empty string cost", []Value{StringVal(""), NumberVal(0), NumberVal(10)}},
		{"empty string salvage", []Value{NumberVal(1000), StringVal(""), NumberVal(10)}},
		{"empty string life", []Value{NumberVal(1000), NumberVal(0), StringVal("")}},
		{"alpha cost", []Value{StringVal("cost"), NumberVal(0), NumberVal(10)}},
		{"alpha salvage", []Value{NumberVal(1000), StringVal("salvage"), NumberVal(10)}},
		{"alpha life", []Value{NumberVal(1000), NumberVal(0), StringVal("years")}},
		{"special chars", []Value{StringVal("#$%"), NumberVal(0), NumberVal(10)}},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnSLN(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			assertError(t, tc.name, v)
		})
	}
}

// === XNPV ===

func TestXNPV_Basic(t *testing.T) {
	// XNPV(0.09, {-10000, 2750, 4250, 3250, 2750}, {39448, 39508, 39751, 39859, 39904})
	// dates as date serial numbers: 2008-01-01, 2008-03-01, 2008-10-30, 2009-02-15, 2009-04-01
	// Expected: ~2086.65
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

func TestXNPV_DocExample(t *testing.T) {
	// From documentation: XNPV(0.09, {-10000, 2750, 4250, 3250, 2750}, dates) = $2,086.65
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
	assertClose(t, "XNPV doc", v, 2086.65)
}

func TestXNPV_ZeroRate(t *testing.T) {
	// With rate=0, XNPV should equal the sum of all values
	// sum = -10000 + 2750 + 4250 + 3250 + 2750 = 3000
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
	v, err := fnXNPV([]Value{NumberVal(0), vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "XNPV zero rate", v, 3000.0)
}

func TestXNPV_HighRate(t *testing.T) {
	// High discount rate reduces future cash flows significantly
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
	v, err := fnXNPV([]Value{NumberVal(5.0), vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	// At rate=5 (500%), future values are heavily discounted
	if v.Type != ValueNumber {
		t.Fatalf("XNPV high rate: expected number, got type %v", v.Type)
	}
	// Result should be negative since future cash flows are nearly worthless
	if v.Num >= 0 {
		t.Fatalf("XNPV high rate: expected negative result, got %f", v.Num)
	}
}

func TestXNPV_MultipleCashFlows(t *testing.T) {
	// 7 cash flows: investment + 6 returns over 2 years
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-50000), NumberVal(8000), NumberVal(9500), NumberVal(10000), NumberVal(11000), NumberVal(12000), NumberVal(13000)},
		},
	}
	// Quarterly payments starting 2020-01-01 (serial 43831)
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(43831), NumberVal(43922), NumberVal(44013), NumberVal(44105), NumberVal(44197), NumberVal(44287), NumberVal(44378)},
		},
	}
	v, err := fnXNPV([]Value{NumberVal(0.10), vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XNPV multiple cash flows: expected number, got type %v", v.Type)
	}
}

func TestXNPV_AllPositiveValues(t *testing.T) {
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(1000), NumberVal(2000), NumberVal(3000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39813), NumberVal(40179)},
		},
	}
	v, err := fnXNPV([]Value{NumberVal(0.05), vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XNPV all positive: expected number, got type %v", v.Type)
	}
	if v.Num <= 0 {
		t.Fatalf("XNPV all positive: expected positive, got %f", v.Num)
	}
}

func TestXNPV_AllNegativeValues(t *testing.T) {
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-1000), NumberVal(-2000), NumberVal(-3000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39813), NumberVal(40179)},
		},
	}
	v, err := fnXNPV([]Value{NumberVal(0.05), vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XNPV all negative: expected number, got type %v", v.Type)
	}
	if v.Num >= 0 {
		t.Fatalf("XNPV all negative: expected negative, got %f", v.Num)
	}
}

func TestXNPV_MixedCashFlows(t *testing.T) {
	// Alternating positive and negative
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-5000), NumberVal(3000), NumberVal(-1000), NumberVal(4000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39630), NumberVal(39813), NumberVal(39995)},
		},
	}
	v, err := fnXNPV([]Value{NumberVal(0.08), vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XNPV mixed: expected number, got type %v", v.Type)
	}
}

func TestXNPV_MonthlyCashFlows(t *testing.T) {
	// Monthly payments for 6 months starting 2008-01-01
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-6000), NumberVal(1100), NumberVal(1100), NumberVal(1100), NumberVal(1100), NumberVal(1100), NumberVal(1100)},
		},
	}
	// ~monthly intervals from serial 39448
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39479), NumberVal(39508), NumberVal(39539), NumberVal(39569), NumberVal(39600), NumberVal(39630)},
		},
	}
	v, err := fnXNPV([]Value{NumberVal(0.10), vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XNPV monthly: expected number, got type %v", v.Type)
	}
	// With 6 payments of 1100 = 6600 total vs 6000 investment, result should be positive
	if v.Num <= 0 {
		t.Fatalf("XNPV monthly: expected positive, got %f", v.Num)
	}
}

func TestXNPV_SingleCashFlow(t *testing.T) {
	// Single cash flow: XNPV = value / (1+rate)^0 = value
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(5000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448)},
		},
	}
	v, err := fnXNPV([]Value{NumberVal(0.10), vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "XNPV single cash flow", v, 5000.0)
}

func TestXNPV_RateNeg1(t *testing.T) {
	// rate = -1 → #NUM! (division by zero in denominator)
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-1000), NumberVal(2000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39813)},
		},
	}
	v, _ := fnXNPV([]Value{NumberVal(-1), vals, dates})
	assertError(t, "XNPV rate=-1", v)
}

func TestXNPV_RateBelowNeg1(t *testing.T) {
	// rate < -1 → #NUM!
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-1000), NumberVal(2000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39813)},
		},
	}
	v, _ := fnXNPV([]Value{NumberVal(-2), vals, dates})
	assertError(t, "XNPV rate<-1", v)
}

func TestXNPV_MismatchedArrays(t *testing.T) {
	// Different number of values and dates → #NUM!
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-1000), NumberVal(2000), NumberVal(3000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39813)},
		},
	}
	v, _ := fnXNPV([]Value{NumberVal(0.05), vals, dates})
	assertError(t, "XNPV mismatched arrays", v)
}

func TestXNPV_TooFewArgs(t *testing.T) {
	v, _ := fnXNPV([]Value{NumberVal(0.05), NumberVal(100)})
	assertError(t, "XNPV too few args", v)
}

func TestXNPV_TooManyArgs(t *testing.T) {
	v, _ := fnXNPV([]Value{NumberVal(0.05), NumberVal(100), NumberVal(39448), NumberVal(99)})
	assertError(t, "XNPV too many args", v)
}

func TestXNPV_DatesNotInOrder(t *testing.T) {
	// Dates not in chronological order — should still work per spec
	// "All other dates must be later than this date, but they may occur in any order."
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000), NumberVal(4250), NumberVal(2750), NumberVal(2750), NumberVal(3250)},
		},
	}
	// Shuffled dates (but all after first date)
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39751), NumberVal(39508), NumberVal(39904), NumberVal(39859)},
		},
	}
	v, err := fnXNPV([]Value{NumberVal(0.09), vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	// Same cash flows as basic test, just reordered — result should match
	assertClose(t, "XNPV dates not in order", v, 2086.65)
}

func TestXNPV_NegativeRate(t *testing.T) {
	// Negative rate between -1 and 0: future values are worth MORE than present
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
	v, err := fnXNPV([]Value{NumberVal(-0.05), vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XNPV negative rate: expected number, got type %v", v.Type)
	}
	// With negative rate, NPV should be higher than the zero-rate sum of 3000
	if v.Num <= 3000 {
		t.Fatalf("XNPV negative rate: expected > 3000, got %f", v.Num)
	}
}

func TestXNPV_EmptyArrays(t *testing.T) {
	// Empty arrays → #NUM!
	vals := Value{
		Type:  ValueArray,
		Array: [][]Value{{}},
	}
	dates := Value{
		Type:  ValueArray,
		Array: [][]Value{{}},
	}
	v, _ := fnXNPV([]Value{NumberVal(0.05), vals, dates})
	assertError(t, "XNPV empty arrays", v)
}

// === XIRR ===

func TestXIRR_Basic(t *testing.T) {
	// XIRR({-10000, 2750, 4250, 3250, 2750}, {39448, 39508, 39751, 39859, 39904})
	// Expected: ~0.3734 (37.34%)
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
	// Expected: ~-0.08442739386111497
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

func TestDB_DocExample_Period1(t *testing.T) {
	// DB(1000000, 100000, 6, 1, 7) = 186083.33
	v, err := fnDB(numArgs(1000000, 100000, 6, 1, 7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB period 1", v, 186083.33)
}

func TestDB_DocExample_Period2(t *testing.T) {
	// DB(1000000, 100000, 6, 2, 7) = 259639.42
	v, err := fnDB(numArgs(1000000, 100000, 6, 2, 7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB period 2", v, 259639.42)
}

func TestDB_DocExample_Period3(t *testing.T) {
	// DB(1000000, 100000, 6, 3, 7) = 176814.44
	v, err := fnDB(numArgs(1000000, 100000, 6, 3, 7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB period 3", v, 176814.44)
}

func TestDB_DocExample_Period4(t *testing.T) {
	// DB(1000000, 100000, 6, 4, 7) = 120410.64
	v, err := fnDB(numArgs(1000000, 100000, 6, 4, 7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB period 4", v, 120410.64)
}

func TestDB_DocExample_Period5(t *testing.T) {
	// DB(1000000, 100000, 6, 5, 7) = 81999.64
	v, err := fnDB(numArgs(1000000, 100000, 6, 5, 7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB period 5", v, 81999.64)
}

func TestDB_DocExample_Period6(t *testing.T) {
	// DB(1000000, 100000, 6, 6, 7) = 55841.76
	v, err := fnDB(numArgs(1000000, 100000, 6, 6, 7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DB period 6", v, 55841.76)
}

func TestDB_DocExample_Period7_LastFractional(t *testing.T) {
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

func TestDDB_DocExample_FirstDayDepreciation(t *testing.T) {
	// DDB(2400, 300, 10*365, 1) = 1.32 (first day's depreciation)
	v, err := fnDDB(numArgs(2400, 300, 10*365, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB first day", v, 1.32)
}

func TestDDB_DocExample_FirstMonthDepreciation(t *testing.T) {
	// DDB(2400, 300, 10*12, 1, 2) = 40.00
	v, err := fnDDB(numArgs(2400, 300, 10*12, 1, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB first month", v, 40.00)
}

func TestDDB_DocExample_FirstYearDepreciation(t *testing.T) {
	// DDB(2400, 300, 10, 1, 2) = 480.00
	v, err := fnDDB(numArgs(2400, 300, 10, 1, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB first year", v, 480.00)
}

func TestDDB_DocExample_SecondYearFactor15(t *testing.T) {
	// DDB(2400, 300, 10, 2, 1.5) = 306.00
	v, err := fnDDB(numArgs(2400, 300, 10, 2, 1.5))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DDB second year factor 1.5", v, 306.00)
}

func TestDDB_DocExample_TenthYear(t *testing.T) {
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

func TestDOLLARDE_DocExample1(t *testing.T) {
	// DOLLARDE(1.02, 16) = 1.125
	v, err := fnDOLLARDE(numArgs(1.02, 16))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARDE(1.02,16)", v, 1.125)
}

func TestDOLLARDE_DocExample2(t *testing.T) {
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

func TestDOLLARFR_DocExample1(t *testing.T) {
	// DOLLARFR(1.125, 16) = 1.02
	v, err := fnDOLLARFR(numArgs(1.125, 16))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "DOLLARFR(1.125,16)", v, 1.02)
}

func TestDOLLARFR_DocExample2(t *testing.T) {
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

func TestEFFECT_DocExample(t *testing.T) {
	// EFFECT(0.0525, 4) = 0.0535427
	v, err := fnEFFECT(numArgs(0.0525, 4))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "EFFECT doc example", v, 0.0535427)
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

func TestNOMINAL_DocExample(t *testing.T) {
	// NOMINAL(0.053543, 4) ≈ 0.0525003
	v, err := fnNOMINAL(numArgs(0.053543, 4))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "NOMINAL doc example", v, 0.0525003)
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

// === CUMIPMT ===

func TestCUMIPMT_DocExample_SecondYear(t *testing.T) {
	// From docs: CUMIPMT(0.09/12, 30*12, 125000, 13, 24, 0) = -11135.23
	v, err := fnCumipmt(numArgs(0.09/12, 360, 125000, 13, 24, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMIPMT doc second year", v, -11135.23)
}

func TestCUMIPMT_DocExample_FirstMonth(t *testing.T) {
	// From docs: CUMIPMT(0.09/12, 30*12, 125000, 1, 1, 0) = -937.50
	v, err := fnCumipmt(numArgs(0.09/12, 360, 125000, 1, 1, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMIPMT doc first month", v, -937.50)
}

func TestCUMIPMT_TotalInterest30YrMortgage(t *testing.T) {
	// CUMIPMT(0.1/12, 360, 100000, 1, 360, 0) — total interest on 30yr mortgage at 10%
	v, err := fnCumipmt(numArgs(0.1/12, 360, 100000, 1, 360, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMIPMT total 30yr", v, -215925.77)
}

func TestCUMIPMT_FirstYear(t *testing.T) {
	// CUMIPMT(0.1/12, 360, 100000, 1, 12, 0) — first year interest
	v, err := fnCumipmt(numArgs(0.1/12, 360, 100000, 1, 12, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMIPMT first year", v, -9974.98)
}

func TestCUMIPMT_SinglePeriod(t *testing.T) {
	// CUMIPMT(0.1/12, 360, 100000, 1, 1, 0) — first month only
	// Expected: -833.33
	v, err := fnCumipmt(numArgs(0.1/12, 360, 100000, 1, 1, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMIPMT single period", v, -833.33)
}

func TestCUMIPMT_MiddlePeriods(t *testing.T) {
	// CUMIPMT(0.1/12, 360, 100000, 13, 24, 0) — second year
	v, err := fnCumipmt(numArgs(0.1/12, 360, 100000, 13, 24, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMIPMT second year", v, -9916.77)
}

func TestCUMIPMT_Type1_FirstYear(t *testing.T) {
	// CUMIPMT(0.1/12, 360, 100000, 1, 12, 1) — type=1, first year
	v, err := fnCumipmt(numArgs(0.1/12, 360, 100000, 1, 12, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMIPMT type1 first year", v, -9066.10)
}

func TestCUMIPMT_Type1_SingleFirst(t *testing.T) {
	// CUMIPMT(0.1/12, 360, 100000, 1, 1, 1) — type=1, first period only
	// Interest for period 1 with type=1 is 0
	v, err := fnCumipmt(numArgs(0.1/12, 360, 100000, 1, 1, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMIPMT type1 period 1", v, 0)
}

func TestCUMIPMT_SimpleLoan(t *testing.T) {
	// CUMIPMT(0.05, 3, 1000, 1, 3, 0) — simple 3-period loan at 5%
	v, err := fnCumipmt(numArgs(0.05, 3, 1000, 1, 3, 0))
	if err != nil {
		t.Fatal(err)
	}
	// PMT = -367.21, total paid = 1101.63, interest = 1101.63 - 1000 = -101.63
	assertClose(t, "CUMIPMT simple loan", v, -101.63)
}

func TestCUMIPMT_SimpleLoan_Period1(t *testing.T) {
	// CUMIPMT(0.05, 3, 1000, 1, 1, 0)
	// First period interest = 1000 * 0.05 = -50
	v, err := fnCumipmt(numArgs(0.05, 3, 1000, 1, 1, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMIPMT simple period 1", v, -50.00)
}

func TestCUMIPMT_SimpleLoan_Period2(t *testing.T) {
	// CUMIPMT(0.05, 3, 1000, 2, 2, 0)
	v, err := fnCumipmt(numArgs(0.05, 3, 1000, 2, 2, 0))
	if err != nil {
		t.Fatal(err)
	}
	// After period 1: balance = 1000*1.05 + pmt = 1050 - 367.21 = 682.79
	// Interest period 2: 682.79 * 0.05 = -34.14
	assertClose(t, "CUMIPMT simple period 2", v, -34.14)
}

func TestCUMIPMT_SimpleLoan_Period3(t *testing.T) {
	// CUMIPMT(0.05, 3, 1000, 3, 3, 0)
	v, err := fnCumipmt(numArgs(0.05, 3, 1000, 3, 3, 0))
	if err != nil {
		t.Fatal(err)
	}
	// Interest period 3: remaining balance * 0.05
	assertClose(t, "CUMIPMT simple period 3", v, -17.49)
}

func TestCUMIPMT_HighRate(t *testing.T) {
	// CUMIPMT(0.5, 10, 10000, 1, 10, 0)
	v, err := fnCumipmt(numArgs(0.5, 10, 10000, 1, 10, 0))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber || v.Num >= 0 {
		t.Errorf("CUMIPMT high rate: expected negative number, got %v", v)
	}
}

func TestCUMIPMT_LastPeriodOnly(t *testing.T) {
	// CUMIPMT(0.1/12, 360, 100000, 360, 360, 0) — last period interest
	v, err := fnCumipmt(numArgs(0.1/12, 360, 100000, 360, 360, 0))
	if err != nil {
		t.Fatal(err)
	}
	// Last period interest should be small
	if v.Type != ValueNumber || v.Num >= 0 {
		t.Errorf("CUMIPMT last period: expected negative number, got %v", v)
	}
	assertClose(t, "CUMIPMT last period", v, -7.25)
}

func TestCUMIPMT_Type1_Total(t *testing.T) {
	// CUMIPMT(0.1/12, 360, 100000, 1, 360, 1) — type=1 total interest
	v, err := fnCumipmt(numArgs(0.1/12, 360, 100000, 1, 360, 1))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber || v.Num >= 0 {
		t.Errorf("CUMIPMT type1 total: expected negative number, got %v", v)
	}
}

func TestCUMIPMT_LargerLoan(t *testing.T) {
	// CUMIPMT(0.06/12, 180, 500000, 1, 180, 0)
	v, err := fnCumipmt(numArgs(0.06/12, 180, 500000, 1, 180, 0))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber || v.Num >= 0 {
		t.Errorf("CUMIPMT larger loan: expected negative number, got %v", v)
	}
}

func TestCUMIPMT_SumPartsEqualsTotal(t *testing.T) {
	// Sum of first half + second half should equal total
	rate := 0.1 / 12
	nper := 360.0
	pvVal := 100000.0
	total, _ := fnCumipmt(numArgs(rate, nper, pvVal, 1, 360, 0))
	first, _ := fnCumipmt(numArgs(rate, nper, pvVal, 1, 180, 0))
	second, _ := fnCumipmt(numArgs(rate, nper, pvVal, 181, 360, 0))
	sum := first.Num + second.Num
	if math.Abs(sum-total.Num) > 0.01 {
		t.Errorf("CUMIPMT parts: first(%f) + second(%f) = %f, total = %f", first.Num, second.Num, sum, total.Num)
	}
}

func TestCUMIPMT_MatchesIPMTSum(t *testing.T) {
	// CUMIPMT should equal sum of individual IPMT calls
	rate := 0.05 / 12
	nper := 60.0
	pvVal := 20000.0
	cum, _ := fnCumipmt(numArgs(rate, nper, pvVal, 1, 60, 0))
	ipmtSum := 0.0
	for i := 1; i <= 60; i++ {
		ipmt, _ := fnIPMT(numArgs(rate, float64(i), nper, pvVal))
		ipmtSum += ipmt.Num
	}
	if math.Abs(cum.Num-ipmtSum) > 0.01 {
		t.Errorf("CUMIPMT vs IPMT sum: CUMIPMT=%f, sum(IPMT)=%f", cum.Num, ipmtSum)
	}
}

// --- Validation errors ---

func TestCUMIPMT_ErrorRateZero(t *testing.T) {
	v, _ := fnCumipmt(numArgs(0, 360, 100000, 1, 360, 0))
	assertError(t, "CUMIPMT rate=0", v)
}

func TestCUMIPMT_ErrorRateNegative(t *testing.T) {
	v, _ := fnCumipmt(numArgs(-0.01, 360, 100000, 1, 360, 0))
	assertError(t, "CUMIPMT rate<0", v)
}

func TestCUMIPMT_ErrorNperZero(t *testing.T) {
	v, _ := fnCumipmt(numArgs(0.01, 0, 100000, 1, 1, 0))
	assertError(t, "CUMIPMT nper=0", v)
}

func TestCUMIPMT_ErrorNperNegative(t *testing.T) {
	v, _ := fnCumipmt(numArgs(0.01, -12, 100000, 1, 1, 0))
	assertError(t, "CUMIPMT nper<0", v)
}

func TestCUMIPMT_ErrorPvZero(t *testing.T) {
	v, _ := fnCumipmt(numArgs(0.01, 360, 0, 1, 360, 0))
	assertError(t, "CUMIPMT pv=0", v)
}

func TestCUMIPMT_ErrorPvNegative(t *testing.T) {
	v, _ := fnCumipmt(numArgs(0.01, 360, -100000, 1, 360, 0))
	assertError(t, "CUMIPMT pv<0", v)
}

func TestCUMIPMT_ErrorStartZero(t *testing.T) {
	v, _ := fnCumipmt(numArgs(0.01, 360, 100000, 0, 360, 0))
	assertError(t, "CUMIPMT start=0", v)
}

func TestCUMIPMT_ErrorStartNegative(t *testing.T) {
	v, _ := fnCumipmt(numArgs(0.01, 360, 100000, -1, 360, 0))
	assertError(t, "CUMIPMT start<0", v)
}

func TestCUMIPMT_ErrorEndLessThanStart(t *testing.T) {
	v, _ := fnCumipmt(numArgs(0.01, 360, 100000, 10, 5, 0))
	assertError(t, "CUMIPMT end<start", v)
}

func TestCUMIPMT_ErrorEndExceedsNper(t *testing.T) {
	v, _ := fnCumipmt(numArgs(0.01, 360, 100000, 1, 361, 0))
	assertError(t, "CUMIPMT end>nper", v)
}

func TestCUMIPMT_ErrorTypeInvalid(t *testing.T) {
	v, _ := fnCumipmt(numArgs(0.01, 360, 100000, 1, 360, 2))
	assertError(t, "CUMIPMT type=2", v)
}

func TestCUMIPMT_ErrorTypeNegative(t *testing.T) {
	v, _ := fnCumipmt(numArgs(0.01, 360, 100000, 1, 360, -1))
	assertError(t, "CUMIPMT type=-1", v)
}

func TestCUMIPMT_ErrorTooFewArgs(t *testing.T) {
	v, _ := fnCumipmt(numArgs(0.01, 360, 100000, 1, 360))
	assertError(t, "CUMIPMT too few args", v)
}

func TestCUMIPMT_ErrorTooManyArgs(t *testing.T) {
	v, _ := fnCumipmt(numArgs(0.01, 360, 100000, 1, 360, 0, 1))
	assertError(t, "CUMIPMT too many args", v)
}

func TestCUMIPMT_ErrorPropagation(t *testing.T) {
	// Pass an error value as an argument
	args := []Value{
		ErrorVal(ErrValNUM),
		NumberVal(360),
		NumberVal(100000),
		NumberVal(1),
		NumberVal(360),
		NumberVal(0),
	}
	v, _ := fnCumipmt(args)
	assertError(t, "CUMIPMT error propagation", v)
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

func TestXIRR_DocExample(t *testing.T) {
	// Documentation example: XIRR(A3:A7, B3:B7, 0.1) = 0.373362535
	// Values: -10000, 2750, 4250, 3250, 2750
	// Dates: 1-Jan-08, 1-Mar-08, 30-Oct-08, 15-Feb-09, 1-Apr-09
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
	v, err := fnXIRR([]Value{vals, dates, NumberVal(0.1)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XIRR doc: expected number, got type %v", v.Type)
	}
	if math.Abs(v.Num-0.373362535) > 0.0001 {
		t.Errorf("XIRR doc: got %f, want ~0.3734", v.Num)
	}
}

func TestXIRR_WithoutGuess(t *testing.T) {
	// Same as basic but without guess; default 0.1 should be used.
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
	assertClose(t, "XIRR without guess", v, 0.3734)
}

func TestXIRR_WithExplicitGuess(t *testing.T) {
	// Provide guess=0.5, should still converge to same answer.
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
	v, err := fnXIRR([]Value{vals, dates, NumberVal(0.5)})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "XIRR with guess 0.5", v, 0.3734)
}

func TestXIRR_AllPositive(t *testing.T) {
	// All positive cash flows should return #NUM!
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(1000), NumberVal(2000), NumberVal(3000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39813), NumberVal(40179)},
		},
	}
	v, _ := fnXIRR([]Value{vals, dates})
	assertError(t, "XIRR all positive", v)
}

func TestXIRR_AllNegative(t *testing.T) {
	// All negative cash flows should return #NUM!
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-1000), NumberVal(-2000), NumberVal(-3000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39813), NumberVal(40179)},
		},
	}
	v, _ := fnXIRR([]Value{vals, dates})
	assertError(t, "XIRR all negative", v)
}

func TestXIRR_SingleCashFlow(t *testing.T) {
	// Single cash flow (len < 2) should return #NUM!
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448)},
		},
	}
	v, _ := fnXIRR([]Value{vals, dates})
	assertError(t, "XIRR single cash flow", v)
}

func TestXIRR_DatesOutOfOrder(t *testing.T) {
	// Dates not in chronological order should still work.
	// Same data as basic but with shuffled order.
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(4250), NumberVal(-10000), NumberVal(2750), NumberVal(2750), NumberVal(3250)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39751), NumberVal(39448), NumberVal(39508), NumberVal(39904), NumberVal(39859)},
		},
	}
	v, err := fnXIRR([]Value{vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XIRR dates out of order: expected number, got type %v", v.Type)
	}
	// Should converge to approximately the same result as the sorted case.
	if math.Abs(v.Num-0.3734) > 0.01 {
		t.Errorf("XIRR dates out of order: got %f, want ~0.3734", v.Num)
	}
}

func TestXIRR_HighReturn(t *testing.T) {
	// Investment that doubles in 6 months: ~300% annualized.
	// -1000 on day 0, +2000 on day 182 (~6 months).
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-1000), NumberVal(2000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39630)},
		},
	}
	v, err := fnXIRR([]Value{vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XIRR high return: expected number, got type %v", v.Type)
	}
	// Annualized return for doubling in ~182 days is very high (>100%).
	if v.Num <= 1.0 {
		t.Errorf("XIRR high return: expected rate > 1.0, got %f", v.Num)
	}
}

func TestXIRR_LowReturn(t *testing.T) {
	// Small return: -10000 invested, 10100 returned after 1 year (~1%).
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000), NumberVal(10100)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39813)},
		},
	}
	v, err := fnXIRR([]Value{vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XIRR low return: expected number, got type %v", v.Type)
	}
	if math.Abs(v.Num-0.01) > 0.005 {
		t.Errorf("XIRR low return: got %f, want ~0.01", v.Num)
	}
}

func TestXIRR_MonthlyCashFlows(t *testing.T) {
	// Monthly cash flows over 12 months.
	// -12000 initial investment, then 1100 each month for 12 months.
	// Dates are approximately monthly starting from serial 39448 (Jan 1, 2008).
	cfVals := []Value{NumberVal(-12000)}
	cfDates := []Value{NumberVal(39448)}
	for i := 1; i <= 12; i++ {
		cfVals = append(cfVals, NumberVal(1100))
		cfDates = append(cfDates, NumberVal(39448+float64(i*30)))
	}
	vals := Value{
		Type:  ValueArray,
		Array: [][]Value{cfVals},
	}
	dates := Value{
		Type:  ValueArray,
		Array: [][]Value{cfDates},
	}
	v, err := fnXIRR([]Value{vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XIRR monthly: expected number, got type %v", v.Type)
	}
	// 12 payments of 1100 = 13200 on 12000 investment; positive return expected.
	if v.Num <= 0 {
		t.Errorf("XIRR monthly: expected positive rate, got %f", v.Num)
	}
}

func TestXIRR_TooFewArgs(t *testing.T) {
	// Only one argument should return #VALUE!
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000), NumberVal(5000)},
		},
	}
	v, _ := fnXIRR([]Value{vals})
	assertError(t, "XIRR too few args", v)
}

func TestXIRR_TooManyArgs(t *testing.T) {
	// Four arguments should return #VALUE!
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000), NumberVal(5000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39813)},
		},
	}
	v, _ := fnXIRR([]Value{vals, dates, NumberVal(0.1), NumberVal(0.2)})
	assertError(t, "XIRR too many args", v)
}

func TestXIRR_MismatchedArraySizes(t *testing.T) {
	// Values has 3 elements, dates has 2 → #NUM!
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000), NumberVal(2750), NumberVal(4250)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39508)},
		},
	}
	v, _ := fnXIRR([]Value{vals, dates})
	assertError(t, "XIRR mismatched arrays", v)
}

func TestXIRR_MultipleCashFlows(t *testing.T) {
	// Multiple irregular cash flows.
	// -5000 initial, then small returns over 2 years.
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-5000), NumberVal(500), NumberVal(700), NumberVal(800), NumberVal(1000), NumberVal(1200), NumberVal(1500)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39539), NumberVal(39630), NumberVal(39722), NumberVal(39813), NumberVal(39904), NumberVal(39995)},
		},
	}
	v, err := fnXIRR([]Value{vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XIRR multiple cash flows: expected number, got type %v", v.Type)
	}
	// Total returns = 5700 on 5000, should be a positive rate.
	if v.Num <= 0 {
		t.Errorf("XIRR multiple cash flows: expected positive rate, got %f", v.Num)
	}
}

func TestXIRR_BreakEven(t *testing.T) {
	// Invest -10000, get back exactly 10000 after 1 year → rate ≈ 0.
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000), NumberVal(10000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39813)},
		},
	}
	v, err := fnXIRR([]Value{vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XIRR break even: expected number, got type %v", v.Type)
	}
	if math.Abs(v.Num) > 0.01 {
		t.Errorf("XIRR break even: got %f, want ~0.0", v.Num)
	}
}

func TestXIRR_ZeroCashFlow(t *testing.T) {
	// Include a zero cash flow in the middle; should still work.
	vals := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(-10000), NumberVal(0), NumberVal(12000)},
		},
	}
	dates := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(39448), NumberVal(39630), NumberVal(39813)},
		},
	}
	v, err := fnXIRR([]Value{vals, dates})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("XIRR zero cash flow: expected number, got type %v", v.Type)
	}
	// 12000 return on 10000 in 1 year → ~20% return.
	if math.Abs(v.Num-0.20) > 0.02 {
		t.Errorf("XIRR zero cash flow: got %f, want ~0.20", v.Num)
	}
}

// === CUMPRINC ===

func TestCUMPRINC_FullLife30YrMortgage(t *testing.T) {
	// Total principal over full life of a 30-year mortgage should equal -PV.
	// CUMPRINC(0.1/12, 360, 100000, 1, 360, 0) = -100000
	v, err := fnCumprinc(numArgs(0.1/12, 360, 100000, 1, 360, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMPRINC full life", v, -100000)
}

func TestCUMPRINC_FirstYear(t *testing.T) {
	// CUMPRINC(0.1/12, 360, 100000, 1, 12, 0)
	// Expected: -555.88 (approx)
	v, err := fnCumprinc(numArgs(0.1/12, 360, 100000, 1, 12, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMPRINC first year", v, -555.88)
}

func TestCUMPRINC_SinglePeriod(t *testing.T) {
	// CUMPRINC(0.1/12, 360, 100000, 1, 1, 0)
	// First principal payment = PMT - interest on full balance
	// PMT = -877.57, interest = -100000*0.1/12 = -833.33, principal = -877.57 - (-833.33) = -44.24
	v, err := fnCumprinc(numArgs(0.1/12, 360, 100000, 1, 1, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMPRINC single period", v, -44.24)
}

func TestCUMPRINC_Type1(t *testing.T) {
	// CUMPRINC(0.1/12, 360, 100000, 1, 12, 1) — beginning of period payments
	v, err := fnCumprinc(numArgs(0.1/12, 360, 100000, 1, 12, 1))
	if err != nil {
		t.Fatal(err)
	}
	// With type=1, first period has 0 interest, so more principal is paid.
	if v.Type != ValueNumber {
		t.Fatalf("CUMPRINC type=1: expected number, got type %v", v.Type)
	}
	if v.Num >= 0 {
		t.Errorf("CUMPRINC type=1: expected negative, got %f", v.Num)
	}
}

func TestCUMPRINC_SimpleLoan(t *testing.T) {
	// CUMPRINC(0.05, 3, 1000, 1, 3, 0) — full life of simple 3-period loan
	// Total principal should equal -1000
	v, err := fnCumprinc(numArgs(0.05, 3, 1000, 1, 3, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMPRINC simple loan", v, -1000)
}

func TestCUMPRINC_MiddlePeriods(t *testing.T) {
	// CUMPRINC(0.1/12, 360, 100000, 13, 24, 0) — second year
	v, err := fnCumprinc(numArgs(0.1/12, 360, 100000, 13, 24, 0))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("CUMPRINC middle periods: expected number, got type %v", v.Type)
	}
	// Second year principal should be more than first year (in absolute value)
	firstYear, _ := fnCumprinc(numArgs(0.1/12, 360, 100000, 1, 12, 0))
	if v.Num >= firstYear.Num {
		t.Errorf("CUMPRINC: second year principal (%f) should be more negative than first year (%f)", v.Num, firstYear.Num)
	}
}

func TestCUMPRINC_LastPeriod(t *testing.T) {
	// CUMPRINC(0.1/12, 360, 100000, 360, 360, 0) — last payment
	v, err := fnCumprinc(numArgs(0.1/12, 360, 100000, 360, 360, 0))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("CUMPRINC last period: expected number, got type %v", v.Type)
	}
	if v.Num >= 0 {
		t.Errorf("CUMPRINC last period: expected negative, got %f", v.Num)
	}
}

func TestCUMPRINC_Type1_FullLife(t *testing.T) {
	// CUMPRINC(0.05, 3, 1000, 1, 3, 1) — full life with type=1
	// Total principal should still equal -PV = -1000
	v, err := fnCumprinc(numArgs(0.05, 3, 1000, 1, 3, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMPRINC type=1 full life", v, -1000)
}

func TestCUMPRINC_Type1_SingleFirstPeriod(t *testing.T) {
	// CUMPRINC(0.1/12, 360, 100000, 1, 1, 1) — type=1, first period
	// With beginning-of-period payment, interest for period 1 is 0,
	// so principal = PMT
	v, err := fnCumprinc(numArgs(0.1/12, 360, 100000, 1, 1, 1))
	if err != nil {
		t.Fatal(err)
	}
	pmt := pmtCore(0.1/12, 360, 100000, 0, 1)
	assertClose(t, "CUMPRINC type=1 first period", v, pmt)
}

func TestCUMPRINC_HighRate(t *testing.T) {
	// CUMPRINC(0.5, 10, 10000, 1, 10, 0) — 50% rate
	v, err := fnCumprinc(numArgs(0.5, 10, 10000, 1, 10, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMPRINC high rate", v, -10000)
}

func TestCUMPRINC_LowRate(t *testing.T) {
	// CUMPRINC(0.001, 12, 5000, 1, 12, 0) — low rate, full life
	v, err := fnCumprinc(numArgs(0.001, 12, 5000, 1, 12, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMPRINC low rate", v, -5000)
}

func TestCUMPRINC_RelationshipWithCUMIPMT(t *testing.T) {
	// CUMIPMT + CUMPRINC over full life should equal total payments (PMT * nper)
	rate := 0.1 / 12
	nper := 360.0
	pv := 100000.0
	cumI, _ := fnCumipmt(numArgs(rate, nper, pv, 1, nper, 0))
	cumP, _ := fnCumprinc(numArgs(rate, nper, pv, 1, nper, 0))
	pmt := pmtCore(rate, nper, pv, 0, 0)
	totalPayments := pmt * nper
	sum := cumI.Num + cumP.Num
	if math.Abs(sum-totalPayments) > 0.01 {
		t.Errorf("CUMIPMT(%f) + CUMPRINC(%f) = %f, total payments = %f", cumI.Num, cumP.Num, sum, totalPayments)
	}
}

func TestCUMPRINC_RelationshipWithCUMIPMT_Type1(t *testing.T) {
	// Same relationship with type=1
	rate := 0.08 / 12
	nper := 120.0
	pv := 50000.0
	cumI, _ := fnCumipmt(numArgs(rate, nper, pv, 1, nper, 1))
	cumP, _ := fnCumprinc(numArgs(rate, nper, pv, 1, nper, 1))
	pmt := pmtCore(rate, nper, pv, 0, 1)
	totalPayments := pmt * nper
	sum := cumI.Num + cumP.Num
	if math.Abs(sum-totalPayments) > 0.01 {
		t.Errorf("Type1: CUMIPMT(%f) + CUMPRINC(%f) = %f, total payments = %f", cumI.Num, cumP.Num, sum, totalPayments)
	}
}

func TestCUMPRINC_RelationshipPartialRange(t *testing.T) {
	// CUMIPMT + CUMPRINC for a partial range should equal PMT * number of periods
	rate := 0.06 / 12
	nper := 240.0
	pv := 200000.0
	cumI, _ := fnCumipmt(numArgs(rate, nper, pv, 13, 24, 0))
	cumP, _ := fnCumprinc(numArgs(rate, nper, pv, 13, 24, 0))
	pmt := pmtCore(rate, nper, pv, 0, 0)
	totalPayments := pmt * 12
	sum := cumI.Num + cumP.Num
	if math.Abs(sum-totalPayments) > 0.01 {
		t.Errorf("Partial: CUMIPMT(%f) + CUMPRINC(%f) = %f, expected %f", cumI.Num, cumP.Num, sum, totalPayments)
	}
}

func TestCUMPRINC_SinglePeriodMatchesPPMT(t *testing.T) {
	// CUMPRINC for a single period should equal PPMT for that period
	rate := 0.05 / 12
	nper := 360.0
	pv := 200000.0
	for _, per := range []float64{1, 12, 60, 180, 360} {
		cumP, _ := fnCumprinc(numArgs(rate, nper, pv, per, per, 0))
		ppmt, _ := fnPPMT(numArgs(rate, per, nper, pv))
		if math.Abs(cumP.Num-ppmt.Num) > 0.01 {
			t.Errorf("per=%v: CUMPRINC(%f) != PPMT(%f)", per, cumP.Num, ppmt.Num)
		}
	}
}

func TestCUMPRINC_LargeLoan(t *testing.T) {
	// CUMPRINC(0.04/12, 360, 1000000, 1, 360, 0) = -1000000
	v, err := fnCumprinc(numArgs(0.04/12, 360, 1000000, 1, 360, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMPRINC large loan", v, -1000000)
}

func TestCUMPRINC_ShortLoan(t *testing.T) {
	// CUMPRINC(0.06/12, 12, 10000, 1, 12, 0) = -10000
	v, err := fnCumprinc(numArgs(0.06/12, 12, 10000, 1, 12, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "CUMPRINC short loan", v, -10000)
}

func TestCUMPRINC_SumPartsEqualsTotal(t *testing.T) {
	// Sum of first half + second half should equal total
	rate := 0.1 / 12
	nper := 360.0
	pvVal := 100000.0
	total, _ := fnCumprinc(numArgs(rate, nper, pvVal, 1, 360, 0))
	first, _ := fnCumprinc(numArgs(rate, nper, pvVal, 1, 180, 0))
	second, _ := fnCumprinc(numArgs(rate, nper, pvVal, 181, 360, 0))
	sum := first.Num + second.Num
	if math.Abs(sum-total.Num) > 0.01 {
		t.Errorf("CUMPRINC parts: first(%f) + second(%f) = %f, total = %f", first.Num, second.Num, sum, total.Num)
	}
}

func TestCUMPRINC_ErrorTooFewArgs(t *testing.T) {
	v, _ := fnCumprinc(numArgs(0.1, 12, 1000, 1, 12))
	assertError(t, "CUMPRINC too few args", v)
}

func TestCUMPRINC_ErrorTooManyArgs(t *testing.T) {
	v, _ := fnCumprinc(numArgs(0.1, 12, 1000, 1, 12, 0, 1))
	assertError(t, "CUMPRINC too many args", v)
}

func TestCUMPRINC_ErrorRateZero(t *testing.T) {
	v, _ := fnCumprinc(numArgs(0, 12, 1000, 1, 12, 0))
	assertError(t, "CUMPRINC rate=0", v)
}

func TestCUMPRINC_ErrorRateNegative(t *testing.T) {
	v, _ := fnCumprinc(numArgs(-0.1, 12, 1000, 1, 12, 0))
	assertError(t, "CUMPRINC rate<0", v)
}

func TestCUMPRINC_ErrorNperZero(t *testing.T) {
	v, _ := fnCumprinc(numArgs(0.1, 0, 1000, 1, 1, 0))
	assertError(t, "CUMPRINC nper=0", v)
}

func TestCUMPRINC_ErrorPvZero(t *testing.T) {
	v, _ := fnCumprinc(numArgs(0.1, 12, 0, 1, 12, 0))
	assertError(t, "CUMPRINC pv=0", v)
}

func TestCUMPRINC_ErrorStartPeriodZero(t *testing.T) {
	v, _ := fnCumprinc(numArgs(0.1, 12, 1000, 0, 12, 0))
	assertError(t, "CUMPRINC start=0", v)
}

func TestCUMPRINC_ErrorEndBeforeStart(t *testing.T) {
	v, _ := fnCumprinc(numArgs(0.1, 12, 1000, 6, 3, 0))
	assertError(t, "CUMPRINC end<start", v)
}

func TestCUMPRINC_ErrorEndExceedsNper(t *testing.T) {
	v, _ := fnCumprinc(numArgs(0.1, 12, 1000, 1, 13, 0))
	assertError(t, "CUMPRINC end>nper", v)
}

func TestCUMPRINC_ErrorInvalidType(t *testing.T) {
	v, _ := fnCumprinc(numArgs(0.1, 12, 1000, 1, 12, 2))
	assertError(t, "CUMPRINC type=2", v)
}

func TestCUMPRINC_ErrorNegativeType(t *testing.T) {
	v, _ := fnCumprinc(numArgs(0.1, 12, 1000, 1, 12, -1))
	assertError(t, "CUMPRINC type=-1", v)
}

// === MIRR ===

func mirrArray(vals ...float64) Value {
	row := make([]Value, len(vals))
	for i, v := range vals {
		row[i] = NumberVal(v)
	}
	return Value{Type: ValueArray, Array: [][]Value{row}}
}

func TestMIRR_DocExample(t *testing.T) {
	// MIRR({-120000,39000,30000,21000,37000,46000}, 0.10, 0.12) ≈ 0.126094
	v, err := fnMirr([]Value{mirrArray(-120000, 39000, 30000, 21000, 37000, 46000), NumberVal(0.10), NumberVal(0.12)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR doc example: expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-0.126094) > 0.0001 {
		t.Errorf("MIRR doc example: got %f, want ~0.126094", v.Num)
	}
}

func TestMIRR_AllNegative(t *testing.T) {
	// Returns -1 when all cash flows are negative (FV of positives is 0).
	v, _ := fnMirr([]Value{mirrArray(-1, -2, -3), NumberVal(0.1), NumberVal(0.1)})
	assertClose(t, "MIRR all negative", v, -1.0)
}

func TestMIRR_AllPositive(t *testing.T) {
	v, _ := fnMirr([]Value{mirrArray(1, 2, 3), NumberVal(0.1), NumberVal(0.1)})
	assertError(t, "MIRR all positive", v)
}

func TestMIRR_TwoValues(t *testing.T) {
	// MIRR({-100,110}, 0.1, 0.1) = 0.10
	v, err := fnMirr([]Value{mirrArray(-100, 110), NumberVal(0.1), NumberVal(0.1)})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "MIRR two values", v, 0.10)
}

func TestMIRR_ZeroFinanceRate(t *testing.T) {
	// MIRR({-100,50,60}, 0, 0.1)
	v, err := fnMirr([]Value{mirrArray(-100, 50, 60), NumberVal(0), NumberVal(0.1)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR zero finance rate: expected number, got %v", v.Type)
	}
	// With finance_rate=0, PV of negatives = -100
	// FV of positives at 0.1: 50*1.1 + 60 = 55 + 60 = 115
	// MIRR = (115/100)^(1/2) - 1 ≈ 0.07238
	if math.Abs(v.Num-0.07238) > 0.001 {
		t.Errorf("MIRR zero finance rate: got %f, want ~0.07238", v.Num)
	}
}

func TestMIRR_ZeroReinvestRate(t *testing.T) {
	// MIRR({-100,50,60}, 0.1, 0)
	v, err := fnMirr([]Value{mirrArray(-100, 50, 60), NumberVal(0.1), NumberVal(0)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR zero reinvest rate: expected number, got %v", v.Type)
	}
	// With reinvest_rate=0, FV of positives = 50 + 60 = 110
	// PV of negatives at 0.1: -100 / 1.0 = -100
	// MIRR = (110/100)^(1/2) - 1 ≈ 0.04881
	if math.Abs(v.Num-0.04881) > 0.001 {
		t.Errorf("MIRR zero reinvest rate: got %f, want ~0.04881", v.Num)
	}
}

func TestMIRR_LargeValues(t *testing.T) {
	v, err := fnMirr([]Value{mirrArray(-1000000, 300000, 400000, 500000), NumberVal(0.05), NumberVal(0.08)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR large values: expected number, got %v", v.Type)
	}
	// Verify it returns a reasonable rate
	if v.Num < -1 || v.Num > 1 {
		t.Errorf("MIRR large values: got unreasonable rate %f", v.Num)
	}
}

func TestMIRR_WrongArgCount_TooFew(t *testing.T) {
	v, _ := fnMirr([]Value{mirrArray(-100, 110), NumberVal(0.1)})
	assertError(t, "MIRR too few args", v)
}

func TestMIRR_WrongArgCount_TooMany(t *testing.T) {
	v, _ := fnMirr([]Value{mirrArray(-100, 110), NumberVal(0.1), NumberVal(0.1), NumberVal(0.1)})
	assertError(t, "MIRR too many args", v)
}

func TestMIRR_ErrorPropagation_FinanceRate(t *testing.T) {
	v, _ := fnMirr([]Value{mirrArray(-100, 110), ErrorVal(ErrValVALUE), NumberVal(0.1)})
	assertError(t, "MIRR error in finance_rate", v)
}

func TestMIRR_ErrorPropagation_ReinvestRate(t *testing.T) {
	v, _ := fnMirr([]Value{mirrArray(-100, 110), NumberVal(0.1), ErrorVal(ErrValVALUE)})
	assertError(t, "MIRR error in reinvest_rate", v)
}

func TestMIRR_ErrorInValues(t *testing.T) {
	arr := Value{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(-100), ErrorVal(ErrValVALUE), NumberVal(50)}},
	}
	v, _ := fnMirr([]Value{arr, NumberVal(0.1), NumberVal(0.1)})
	assertError(t, "MIRR error in values", v)
}

func TestMIRR_SingleValue(t *testing.T) {
	// Only one cash flow — not enough
	v, _ := fnMirr([]Value{mirrArray(-100), NumberVal(0.1), NumberVal(0.1)})
	assertError(t, "MIRR single value", v)
}

func TestMIRR_EqualRates(t *testing.T) {
	// MIRR({-100,50,60}, 0.1, 0.1)
	v, err := fnMirr([]Value{mirrArray(-100, 50, 60), NumberVal(0.1), NumberVal(0.1)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR equal rates: expected number, got %v", v.Type)
	}
}

func TestMIRR_NegativeRates(t *testing.T) {
	// Negative rates are valid inputs
	v, err := fnMirr([]Value{mirrArray(-100, 50, 60), NumberVal(-0.05), NumberVal(-0.05)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR negative rates: expected number, got %v", v.Type)
	}
}

func TestMIRR_ZeroCashFlowsIncluded(t *testing.T) {
	// Zero values are neither positive nor negative; should still work if there are both pos and neg
	v, err := fnMirr([]Value{mirrArray(-100, 0, 0, 110), NumberVal(0.1), NumberVal(0.1)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR with zeros: expected number, got %v", v.Type)
	}
}

func TestMIRR_AllZeros(t *testing.T) {
	// All zeros — no positive or negative
	v, _ := fnMirr([]Value{mirrArray(0, 0, 0), NumberVal(0.1), NumberVal(0.1)})
	assertError(t, "MIRR all zeros", v)
}

func TestMIRR_HighRates(t *testing.T) {
	v, err := fnMirr([]Value{mirrArray(-100, 200, 300), NumberVal(0.5), NumberVal(0.5)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR high rates: expected number, got %v", v.Type)
	}
}

func TestMIRR_ManyPeriods(t *testing.T) {
	// 10 periods
	v, err := fnMirr([]Value{mirrArray(-500, 50, 60, 70, 80, 90, 100, 110, 120, 130), NumberVal(0.08), NumberVal(0.10)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR many periods: expected number, got %v", v.Type)
	}
}

func TestMIRR_MultipleNegativeCashFlows(t *testing.T) {
	// Multiple negative cash flows
	v, err := fnMirr([]Value{mirrArray(-100, 50, -20, 80, 60), NumberVal(0.10), NumberVal(0.12)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR multiple negatives: expected number, got %v", v.Type)
	}
}

func TestMIRR_SmallValues(t *testing.T) {
	v, err := fnMirr([]Value{mirrArray(-0.01, 0.005, 0.006), NumberVal(0.1), NumberVal(0.1)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR small values: expected number, got %v", v.Type)
	}
}

func TestMIRR_DocExample2(t *testing.T) {
	// MIRR({-120000,39000,30000,21000,37000,46000}, 0.10, 0.14) ≈ 0.134759
	v, err := fnMirr([]Value{mirrArray(-120000, 39000, 30000, 21000, 37000, 46000), NumberVal(0.10), NumberVal(0.14)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR doc example 2: expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-0.134759) > 0.0001 {
		t.Errorf("MIRR doc example 2: got %f, want ~0.134759", v.Num)
	}
}

func TestMIRR_BothRatesZero(t *testing.T) {
	// MIRR({-100,50,60}, 0, 0)
	v, err := fnMirr([]Value{mirrArray(-100, 50, 60), NumberVal(0), NumberVal(0)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR both rates zero: expected number, got %v", v.Type)
	}
	// FV of positives at 0: 50 + 60 = 110
	// PV of negatives at 0: -100
	// MIRR = (110/100)^(1/2) - 1 ≈ 0.04881
	if math.Abs(v.Num-0.04881) > 0.001 {
		t.Errorf("MIRR both rates zero: got %f, want ~0.04881", v.Num)
	}
}

func TestMIRR_EmptyValuesSkipped(t *testing.T) {
	arr := Value{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(-100), {Type: ValueEmpty}, NumberVal(110)}},
	}
	v, err := fnMirr([]Value{arr, NumberVal(0.1), NumberVal(0.1)})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("MIRR empty skipped: expected number, got %v", v.Type)
	}
	assertClose(t, "MIRR empty skipped", v, 0.10)
}

// === PDURATION ===

func TestPDURATION_Basic(t *testing.T) {
	// PDURATION(0.025, 2000, 2200) ≈ 3.859
	v, err := fnPduration(numArgs(0.025, 2000, 2200))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PDURATION basic", v, 3.86)
}

func TestPDURATION_DoublingAt10Percent(t *testing.T) {
	v, err := fnPduration(numArgs(0.1, 1000, 2000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PDURATION doubling 10%", v, 7.27)
}

func TestPDURATION_DoublingAt1Percent(t *testing.T) {
	v, err := fnPduration(numArgs(0.01, 100, 200))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PDURATION doubling 1%", v, 69.66)
}

func TestPDURATION_TripleAt5Percent(t *testing.T) {
	v, err := fnPduration(numArgs(0.05, 1000, 3000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PDURATION triple 5%", v, 22.52)
}

func TestPDURATION_SmallRate(t *testing.T) {
	v, err := fnPduration(numArgs(0.001, 500, 600))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PDURATION small rate", v, 182.41)
}

func TestPDURATION_LargeRate(t *testing.T) {
	v, err := fnPduration(numArgs(1.0, 100, 1000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PDURATION large rate", v, 3.32)
}

func TestPDURATION_SmallGrowth(t *testing.T) {
	v, err := fnPduration(numArgs(0.05, 1000, 1001))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PDURATION small growth", v, 0.02)
}

func TestPDURATION_LargeValues(t *testing.T) {
	v, err := fnPduration(numArgs(0.08, 1000000, 2000000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PDURATION large values", v, 9.01)
}

func TestPDURATION_FractionalRate(t *testing.T) {
	v, err := fnPduration(numArgs(0.0375, 5000, 7500))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PDURATION fractional rate", v, 11.01)
}

func TestPDURATION_PVEqualsFV(t *testing.T) {
	v, err := fnPduration(numArgs(0.05, 1000, 1000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "PDURATION pv=fv", v, 0.0)
}

func TestPDURATION_FVLessThanPV(t *testing.T) {
	v, err := fnPduration(numArgs(0.05, 2000, 1000))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber || v.Num >= 0 {
		t.Errorf("PDURATION fv<pv: expected negative number, got %v", v)
	}
}

func TestPDURATION_ErrorRateZero(t *testing.T) {
	v, _ := fnPduration(numArgs(0, 1000, 2000))
	assertError(t, "PDURATION rate=0", v)
}

func TestPDURATION_ErrorRateNegative(t *testing.T) {
	v, _ := fnPduration(numArgs(-0.05, 1000, 2000))
	assertError(t, "PDURATION rate<0", v)
}

func TestPDURATION_ErrorPVZero(t *testing.T) {
	v, _ := fnPduration(numArgs(0.05, 0, 2000))
	assertError(t, "PDURATION pv=0", v)
}

func TestPDURATION_ErrorPVNegative(t *testing.T) {
	v, _ := fnPduration(numArgs(0.05, -1000, 2000))
	assertError(t, "PDURATION pv<0", v)
}

func TestPDURATION_ErrorFVZero(t *testing.T) {
	v, _ := fnPduration(numArgs(0.05, 1000, 0))
	assertError(t, "PDURATION fv=0", v)
}

func TestPDURATION_ErrorFVNegative(t *testing.T) {
	v, _ := fnPduration(numArgs(0.05, 1000, -2000))
	assertError(t, "PDURATION fv<0", v)
}

func TestPDURATION_ErrorTooFewArgs(t *testing.T) {
	v, _ := fnPduration(numArgs(0.05, 1000))
	assertError(t, "PDURATION too few args", v)
}

func TestPDURATION_ErrorTooManyArgs(t *testing.T) {
	v, _ := fnPduration(numArgs(0.05, 1000, 2000, 1))
	assertError(t, "PDURATION too many args", v)
}

func TestPDURATION_ErrorStringArg(t *testing.T) {
	v, _ := fnPduration([]Value{StringVal("abc"), NumberVal(1000), NumberVal(2000)})
	assertError(t, "PDURATION string arg", v)
}

func TestPDURATION_ErrorNoArgs(t *testing.T) {
	v, _ := fnPduration([]Value{})
	assertError(t, "PDURATION no args", v)
}

// === RRI ===

func TestRRI_Basic(t *testing.T) {
	v, err := fnRri(numArgs(96, 10000, 11000))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("RRI basic: expected number, got %v", v)
	}
	if math.Abs(v.Num-0.000988) > 0.001 {
		t.Errorf("RRI basic: got %f, want ~0.000988", v.Num)
	}
}

func TestRRI_SinglePeriod(t *testing.T) {
	v, err := fnRri(numArgs(1, 100, 110))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "RRI single period", v, 0.10)
}

func TestRRI_AnnualDoublingRate(t *testing.T) {
	v, err := fnRri(numArgs(12, 1000, 2000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "RRI annual doubling", v, 0.06)
}

func TestRRI_NegativeGrowth(t *testing.T) {
	v, err := fnRri(numArgs(10, 1000, 500))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber || v.Num >= 0 {
		t.Errorf("RRI negative growth: expected negative number, got %v", v)
	}
}

func TestRRI_NoGrowth(t *testing.T) {
	v, err := fnRri(numArgs(10, 1000, 1000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "RRI no growth", v, 0.0)
}

func TestRRI_LargeNper(t *testing.T) {
	v, err := fnRri(numArgs(360, 100000, 200000))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("RRI large nper: expected number, got %v", v)
	}
	if math.Abs(v.Num-0.001928) > 0.001 {
		t.Errorf("RRI large nper: got %f, want ~0.001928", v.Num)
	}
}

func TestRRI_SmallValues(t *testing.T) {
	v, err := fnRri(numArgs(5, 1, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "RRI small values", v, 0.15)
}

func TestRRI_LargeGrowth(t *testing.T) {
	v, err := fnRri(numArgs(10, 100, 10000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "RRI large growth", v, 0.58)
}

func TestRRI_FVZero(t *testing.T) {
	v, err := fnRri(numArgs(10, 1000, 0))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "RRI fv=0", v, -1.0)
}

func TestRRI_NegativePV(t *testing.T) {
	v, err := fnRri(numArgs(10, -1000, -2000))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("RRI negative pv: expected number, got %v", v)
	}
}

func TestRRI_NegativeFV(t *testing.T) {
	v, err := fnRri(numArgs(1, 100, -50))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "RRI negative fv single period", v, -1.5)
}

func TestRRI_ErrorNperZero(t *testing.T) {
	v, _ := fnRri(numArgs(0, 1000, 2000))
	assertError(t, "RRI nper=0", v)
}

func TestRRI_ErrorNperNegative(t *testing.T) {
	v, _ := fnRri(numArgs(-5, 1000, 2000))
	assertError(t, "RRI nper<0", v)
}

func TestRRI_ErrorPVZero(t *testing.T) {
	v, _ := fnRri(numArgs(10, 0, 2000))
	assertError(t, "RRI pv=0", v)
}

func TestRRI_ErrorTooFewArgs(t *testing.T) {
	v, _ := fnRri(numArgs(10, 1000))
	assertError(t, "RRI too few args", v)
}

func TestRRI_ErrorTooManyArgs(t *testing.T) {
	v, _ := fnRri(numArgs(10, 1000, 2000, 1))
	assertError(t, "RRI too many args", v)
}

func TestRRI_ErrorStringArg(t *testing.T) {
	v, _ := fnRri([]Value{StringVal("abc"), NumberVal(1000), NumberVal(2000)})
	assertError(t, "RRI string arg", v)
}

func TestRRI_ErrorNoArgs(t *testing.T) {
	v, _ := fnRri([]Value{})
	assertError(t, "RRI no args", v)
}

func TestRRI_FractionalNper(t *testing.T) {
	v, err := fnRri(numArgs(0.5, 1000, 1100))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "RRI fractional nper", v, 0.21)
}

func TestRRI_VerySmallNper(t *testing.T) {
	v, err := fnRri(numArgs(0.01, 1000, 1010))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("RRI very small nper: expected number, got %v", v)
	}
}

// === VDB ===

func boolArg(b bool) Value {
	return Value{Type: ValueBool, Bool: b}
}

func TestVDB_DocExample_FirstDayDepreciation(t *testing.T) {
	// VDB(2400, 300, 10*365, 0, 1) = 1.32
	v, err := fnVdb(numArgs(2400, 300, 10*365, 0, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB first day", v, 1.32)
}

func TestVDB_DocExample_FirstMonthDepreciation(t *testing.T) {
	// VDB(2400, 300, 10*12, 0, 1) = 40.00
	v, err := fnVdb(numArgs(2400, 300, 10*12, 0, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB first month", v, 40.00)
}

func TestVDB_DocExample_FirstYearDepreciation(t *testing.T) {
	// VDB(2400, 300, 10, 0, 1) = 480.00
	v, err := fnVdb(numArgs(2400, 300, 10, 0, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB first year", v, 480.00)
}

func TestVDB_DocExample_Month6To18(t *testing.T) {
	// VDB(2400, 300, 10*12, 6, 18) = 396.31
	v, err := fnVdb(numArgs(2400, 300, 10*12, 6, 18))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB month 6-18", v, 396.31)
}

func TestVDB_DocExample_Month6To18_Factor15(t *testing.T) {
	// VDB(2400, 300, 10*12, 6, 18, 1.5) = 311.81
	v, err := fnVdb(numArgs(2400, 300, 10*12, 6, 18, 1.5))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB month 6-18 factor 1.5", v, 311.81)
}

func TestVDB_DocExample_FractionalPeriod(t *testing.T) {
	// VDB(2400, 300, 10, 0, 0.875, 1.5) = 315.00
	v, err := fnVdb(numArgs(2400, 300, 10, 0, 0.875, 1.5))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB fractional period", v, 315.00)
}

func TestVDB_FirstYear_10000(t *testing.T) {
	// VDB(10000, 1000, 5, 0, 1) — first year DDB
	// DDB rate = 2/5 = 0.4, dep = 10000 * 0.4 = 4000
	v, err := fnVdb(numArgs(10000, 1000, 5, 0, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB first year 10000", v, 4000.00)
}

func TestVDB_SecondYear_10000(t *testing.T) {
	// VDB(10000, 1000, 5, 1, 2)
	// After year 1: book = 6000, dep = 6000 * 0.4 = 2400
	v, err := fnVdb(numArgs(10000, 1000, 5, 1, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB second year 10000", v, 2400.00)
}

func TestVDB_FullLife_EqualsCostMinusSalvage(t *testing.T) {
	// VDB(10000, 1000, 5, 0, 5) = 9000 (cost - salvage)
	v, err := fnVdb(numArgs(10000, 1000, 5, 0, 5))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB full life", v, 9000.00)
}

func TestVDB_Factor15(t *testing.T) {
	// VDB(10000, 1000, 5, 0, 1, 1.5)
	// dep = 10000 * 1.5/5 = 3000
	v, err := fnVdb(numArgs(10000, 1000, 5, 0, 1, 1.5))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB factor 1.5", v, 3000.00)
}

func TestVDB_NoSwitch_True(t *testing.T) {
	// VDB(10000, 1000, 5, 0, 1, 2, TRUE)
	// Same as DDB for first period: 10000 * 2/5 = 4000
	args := append(numArgs(10000, 1000, 5, 0, 1, 2), boolArg(true))
	v, err := fnVdb(args)
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB no_switch=TRUE first year", v, 4000.00)
}

func TestVDB_NoSwitch_False(t *testing.T) {
	// VDB(10000, 1000, 5, 0, 1, 2, FALSE)
	// Same as default (no_switch=FALSE)
	args := append(numArgs(10000, 1000, 5, 0, 1, 2), boolArg(false))
	v, err := fnVdb(args)
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB no_switch=FALSE first year", v, 4000.00)
}

func TestVDB_FractionalStart(t *testing.T) {
	// VDB(10000, 1000, 5, 0, 0.5) — half of first year
	// First year full dep = 4000, half = 2000
	v, err := fnVdb(numArgs(10000, 1000, 5, 0, 0.5))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB fractional half year", v, 2000.00)
}

func TestVDB_FractionalEnd(t *testing.T) {
	// VDB(10000, 1000, 5, 0, 1.5)
	// Year 1: 4000, half of year 2 (2400*0.5=1200) = 5200
	v, err := fnVdb(numArgs(10000, 1000, 5, 0, 1.5))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB fractional end", v, 5200.00)
}

func TestVDB_MiddlePeriod(t *testing.T) {
	// VDB(10000, 1000, 5, 2, 3)
	// After y1: book=6000, after y2: book=3600, y3: SL=(3600-1000)/2=1300 vs DDB=3600*0.4=1440 => 1440
	v, err := fnVdb(numArgs(10000, 1000, 5, 2, 3))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB middle period", v, 1440.00)
}

func TestVDB_SwitchToSL(t *testing.T) {
	// VDB(10000, 1000, 5, 3, 4)
	// y0: book=10000, DDB=4000, SL=1800 -> 4000, book=6000
	// y1: book=6000, DDB=2400, SL=1250 -> 2400, book=3600
	// y2: book=3600, DDB=1440, SL=866.67 -> 1440, book=2160
	// y3: book=2160, DDB=864, SL=580 -> 864
	v, err := fnVdb(numArgs(10000, 1000, 5, 3, 4))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB period 3", v, 864.00)
}

func TestVDB_SwitchToSL_Period4(t *testing.T) {
	// VDB(10000, 1000, 5, 4, 5)
	// y4: book=1296, DDB=518.40, SL=296 -> capped to 296
	v, err := fnVdb(numArgs(10000, 1000, 5, 4, 5))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB period 4 capped", v, 296.00)
}

func TestVDB_SwitchToSL_WithSwitch(t *testing.T) {
	// VDB(10000, 1000, 10, 0, 10) with default factor=2 should equal 9000
	// With 10 periods, switch from DDB to SL will occur
	v, err := fnVdb(numArgs(10000, 1000, 10, 0, 10))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB 10yr full life", v, 9000.00)
}

func TestVDB_NoSwitch_FullLife_LessThanCostMinusSalvage(t *testing.T) {
	// With no_switch=TRUE, total depreciation over full life may be less than cost-salvage
	// because DDB alone doesn't fully depreciate.
	args := append(numArgs(10000, 1000, 5, 0, 5, 2), boolArg(true))
	v, err := fnVdb(args)
	if err != nil {
		t.Fatal(err)
	}
	// DDB only: y1=4000, y2=2400, y3=1440, y4=864, y5=518.4 (capped to 296 since 1296-1000=296)
	// Total = 4000+2400+1440+864+296 = 9000... wait, let me recalc.
	// y1: 10000*0.4=4000, book=6000
	// y2: 6000*0.4=2400, book=3600
	// y3: 3600*0.4=1440, book=2160
	// y4: 2160*0.4=864, book=1296
	// y5: 1296*0.4=518.4, but book-salvage=296, so capped to 296
	// Total: 4000+2400+1440+864+296 = 9000
	// Actually with no_switch the salvage cap still applies, so it happens to equal 9000.
	assertClose(t, "VDB no_switch full life", v, 9000.00)
}

func TestVDB_CostZero(t *testing.T) {
	v, err := fnVdb(numArgs(0, 0, 5, 0, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB cost=0", v, 0)
}

func TestVDB_SalvageEqualsCost(t *testing.T) {
	v, err := fnVdb(numArgs(10000, 10000, 5, 0, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB salvage=cost", v, 0)
}

func TestVDB_SalvageZero(t *testing.T) {
	// VDB(10000, 0, 5, 0, 1) = 10000 * 2/5 = 4000
	v, err := fnVdb(numArgs(10000, 0, 5, 0, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB salvage=0", v, 4000)
}

func TestVDB_StartEqualsEnd(t *testing.T) {
	// VDB(10000, 1000, 5, 1, 1) = 0 (no period range)
	v, err := fnVdb(numArgs(10000, 1000, 5, 1, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB start=end", v, 0)
}

func TestVDB_LargeAsset(t *testing.T) {
	// VDB(1000000, 100000, 10, 0, 1) = 1000000 * 2/10 = 200000
	v, err := fnVdb(numArgs(1000000, 100000, 10, 0, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "VDB large asset", v, 200000)
}

func TestVDB_ErrorTooFewArgs(t *testing.T) {
	v, _ := fnVdb(numArgs(1000, 100, 5, 0))
	assertError(t, "VDB too few args", v)
}

func TestVDB_ErrorTooManyArgs(t *testing.T) {
	args := append(numArgs(1000, 100, 5, 0, 1, 2), boolArg(false), NumberVal(99))
	v, _ := fnVdb(args)
	assertError(t, "VDB too many args", v)
}

func TestVDB_ErrorLifeZero(t *testing.T) {
	v, _ := fnVdb(numArgs(1000, 100, 0, 0, 1))
	assertError(t, "VDB life=0", v)
}

func TestVDB_ErrorStartNegative(t *testing.T) {
	v, _ := fnVdb(numArgs(1000, 100, 5, -1, 1))
	assertError(t, "VDB start<0", v)
}

func TestVDB_ErrorEndLessThanStart(t *testing.T) {
	v, _ := fnVdb(numArgs(1000, 100, 5, 3, 2))
	assertError(t, "VDB end<start", v)
}

func TestVDB_ErrorEndExceedsLife(t *testing.T) {
	v, _ := fnVdb(numArgs(1000, 100, 5, 0, 6))
	assertError(t, "VDB end>life", v)
}

func TestVDB_ErrorNegativeCost(t *testing.T) {
	v, _ := fnVdb(numArgs(-1000, 100, 5, 0, 1))
	assertError(t, "VDB negative cost", v)
}

func TestVDB_ErrorNegativeSalvage(t *testing.T) {
	v, _ := fnVdb(numArgs(1000, -100, 5, 0, 1))
	assertError(t, "VDB negative salvage", v)
}

func TestVDB_ErrorFactorZero(t *testing.T) {
	v, _ := fnVdb(numArgs(1000, 100, 5, 0, 1, 0))
	assertError(t, "VDB factor=0", v)
}

// === SYD ===

func TestSYD_DocExample_Period1(t *testing.T) {
	// SYD(30000, 7500, 10, 1) = 4090.909090...
	v, err := fnSYD(numArgs(30000, 7500, 10, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD period 1", v, 4090.91)
}

func TestSYD_DocExample_Period10(t *testing.T) {
	// SYD(30000, 7500, 10, 10) = 409.090909...
	v, err := fnSYD(numArgs(30000, 7500, 10, 10))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD period 10", v, 409.09)
}

func TestSYD_Period5(t *testing.T) {
	// SYD(30000, 7500, 10, 5) = (22500) * 6 / 55 = 2454.545454...
	v, err := fnSYD(numArgs(30000, 7500, 10, 5))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD period 5", v, 2454.55)
}

func TestSYD_Period2(t *testing.T) {
	// SYD(30000, 7500, 10, 2) = 22500 * 9 / 55 = 3681.818181...
	v, err := fnSYD(numArgs(30000, 7500, 10, 2))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD period 2", v, 3681.82)
}

func TestSYD_Period3(t *testing.T) {
	// SYD(30000, 7500, 10, 3) = 22500 * 8 / 55 = 3272.727272...
	v, err := fnSYD(numArgs(30000, 7500, 10, 3))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD period 3", v, 3272.73)
}

func TestSYD_ZeroSalvage(t *testing.T) {
	// SYD(10000, 0, 5, 1) = 10000 * 5 / 15 = 3333.33
	v, err := fnSYD(numArgs(10000, 0, 5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD zero salvage", v, 3333.33)
}

func TestSYD_ZeroSalvage_Period5(t *testing.T) {
	// SYD(10000, 0, 5, 5) = 10000 * 1 / 15 = 666.67
	v, err := fnSYD(numArgs(10000, 0, 5, 5))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD zero salvage period 5", v, 666.67)
}

func TestSYD_SalvageEqualsCost(t *testing.T) {
	// SYD(10000, 10000, 5, 1) = 0
	v, err := fnSYD(numArgs(10000, 10000, 5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD salvage=cost", v, 0)
}

func TestSYD_Life1(t *testing.T) {
	// SYD(10000, 1000, 1, 1) = 9000 * 1 / 1 = 9000
	v, err := fnSYD(numArgs(10000, 1000, 1, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD life=1", v, 9000)
}

func TestSYD_PerEqualsLife(t *testing.T) {
	// SYD(10000, 2000, 4, 4) = 8000 * 1 / 10 = 800
	v, err := fnSYD(numArgs(10000, 2000, 4, 4))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD per=life", v, 800)
}

func TestSYD_LargeValues(t *testing.T) {
	// SYD(50000000, 5000000, 10, 1) = 45000000 * 10 / 55 = 8181818.18
	v, err := fnSYD(numArgs(50000000, 5000000, 10, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD large values", v, 8181818.18)
}

func TestSYD_SmallAsset(t *testing.T) {
	// SYD(100, 10, 5, 1) = 90 * 5 / 15 = 30
	v, err := fnSYD(numArgs(100, 10, 5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD small asset", v, 30)
}

func TestSYD_FractionalLifeTruncated(t *testing.T) {
	// SYD(10000, 1000, 5.9, 1) — life truncated to 5
	// = 9000 * 5 / 15 = 3000
	v, err := fnSYD(numArgs(10000, 1000, 5.9, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD fractional life", v, 3000)
}

func TestSYD_FractionalPerTruncated(t *testing.T) {
	// SYD(10000, 1000, 5, 2.7) — per truncated to 2
	// = 9000 * 4 / 15 = 2400
	v, err := fnSYD(numArgs(10000, 1000, 5, 2.7))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD fractional per", v, 2400)
}

func TestSYD_NegativeDepreciableAmount(t *testing.T) {
	// SYD(1000, 5000, 5, 1) — salvage > cost, negative depreciation
	// = (1000-5000) * 5 / 15 = -1333.33
	v, err := fnSYD(numArgs(1000, 5000, 5, 1))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD negative depreciable", v, -1333.33)
}

func TestSYD_ErrorTooFewArgs(t *testing.T) {
	v, _ := fnSYD(numArgs(30000, 7500, 10))
	assertError(t, "SYD too few args", v)
}

func TestSYD_ErrorTooManyArgs(t *testing.T) {
	v, _ := fnSYD(numArgs(30000, 7500, 10, 1, 99))
	assertError(t, "SYD too many args", v)
}

func TestSYD_ErrorNonNumeric(t *testing.T) {
	args := []Value{StringVal("abc"), NumberVal(7500), NumberVal(10), NumberVal(1)}
	v, _ := fnSYD(args)
	assertError(t, "SYD non-numeric cost", v)
}

func TestSYD_ErrorLifeZero(t *testing.T) {
	v, _ := fnSYD(numArgs(30000, 7500, 0, 1))
	assertError(t, "SYD life=0", v)
}

func TestSYD_ErrorLifeNegative(t *testing.T) {
	v, _ := fnSYD(numArgs(30000, 7500, -5, 1))
	assertError(t, "SYD life negative", v)
}

func TestSYD_ErrorPerZero(t *testing.T) {
	v, _ := fnSYD(numArgs(30000, 7500, 10, 0))
	assertError(t, "SYD per=0", v)
}

func TestSYD_ErrorPerNegative(t *testing.T) {
	v, _ := fnSYD(numArgs(30000, 7500, 10, -1))
	assertError(t, "SYD per negative", v)
}

func TestSYD_ErrorPerGreaterThanLife(t *testing.T) {
	v, _ := fnSYD(numArgs(30000, 7500, 10, 11))
	assertError(t, "SYD per > life", v)
}

func TestSYD_ViaEval_ThreeArgError(t *testing.T) {
	// SYD with 3 args should error; IFERROR should catch it and return "err".
	cf := evalCompile(t, `IFERROR(SYD(10000,1000,5),"err")`)
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueString || v.Str != "err" {
		t.Fatalf("IFERROR(SYD(3 args)) = %#v, want string:err", v)
	}
}

func TestSYD_ViaEval_FiveArgError(t *testing.T) {
	// SYD with 5 args should error; IFERROR should catch it and return "err".
	cf := evalCompile(t, `IFERROR(SYD(10000,1000,5,1,1),"err")`)
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueString || v.Str != "err" {
		t.Fatalf("IFERROR(SYD(5 args)) = %#v, want string:err", v)
	}
}

func TestSYD_ViaEval_SumAllPeriods(t *testing.T) {
	// Sum of SYD over all periods should equal cost - salvage = 10000 - 2000 = 8000.
	cf := evalCompile(t, "SYD(10000,2000,5,1)+SYD(10000,2000,5,2)+SYD(10000,2000,5,3)+SYD(10000,2000,5,4)+SYD(10000,2000,5,5)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "SYD sum all periods", v, 8000)
}

// === ISPMT ===

func TestISPMT_DocExample(t *testing.T) {
	// ISPMT(0.1/12, 1, 36, 8000000) ≈ -64814.81
	// Monthly payment on 8M loan at 10% annual for 3 years, period 1
	v, err := fnIspmt(numArgs(0.1/12, 1, 36, 8000000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT doc", v, -64814.81)
}

func TestISPMT_FirstPeriod(t *testing.T) {
	// per=0 (first period): ISPMT = pv * rate * (0/nper - 1) = pv * rate * (-1) = -pv*rate
	// ISPMT(0.1/12, 0, 36, 8000000) = 8000000 * 0.1/12 * (0/36 - 1) = -66666.67
	v, err := fnIspmt(numArgs(0.1/12, 0, 36, 8000000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT first period", v, -66666.67)
}

func TestISPMT_LastPeriod(t *testing.T) {
	// per=35 (last period, 0-based): ISPMT(0.1/12, 35, 36, 8000000)
	// = 8000000 * (0.1/12) * (35/36 - 1) = 8000000 * 0.008333... * (-0.02778)
	// = -1851.85
	v, err := fnIspmt(numArgs(0.1/12, 35, 36, 8000000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT last period", v, -1851.85)
}

func TestISPMT_MiddlePeriod(t *testing.T) {
	// per=18 (middle): ISPMT(0.1/12, 18, 36, 8000000)
	// = 8000000 * (0.1/12) * (18/36 - 1) = 8000000 * 0.008333... * (-0.5) = -33333.33
	v, err := fnIspmt(numArgs(0.1/12, 18, 36, 8000000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT middle period", v, -33333.33)
}

func TestISPMT_ZeroRate(t *testing.T) {
	// rate=0: ISPMT = pv * 0 * (per/nper - 1) = 0
	v, err := fnIspmt(numArgs(0, 5, 12, 100000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT zero rate", v, 0)
}

func TestISPMT_NegativeRate(t *testing.T) {
	// Negative rate: ISPMT(-0.05/12, 1, 24, 50000)
	// = 50000 * (-0.05/12) * (1/24 - 1) = 50000 * (-0.004167) * (-0.95833) = 199.65
	v, err := fnIspmt(numArgs(-0.05/12, 1, 24, 50000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT negative rate", v, 199.65)
}

func TestISPMT_AnnualPayments(t *testing.T) {
	// Annual: ISPMT(0.08, 1, 5, 100000)
	// = 100000 * 0.08 * (1/5 - 1) = 100000 * 0.08 * (-0.8) = -6400
	v, err := fnIspmt(numArgs(0.08, 1, 5, 100000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT annual", v, -6400)
}

func TestISPMT_AnnualFirstPeriod(t *testing.T) {
	// ISPMT(0.08, 0, 5, 100000) = 100000 * 0.08 * (0/5 - 1) = -8000
	v, err := fnIspmt(numArgs(0.08, 0, 5, 100000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT annual first", v, -8000)
}

func TestISPMT_AnnualLastPeriod(t *testing.T) {
	// ISPMT(0.08, 4, 5, 100000) = 100000 * 0.08 * (4/5 - 1) = 100000 * 0.08 * (-0.2) = -1600
	v, err := fnIspmt(numArgs(0.08, 4, 5, 100000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT annual last", v, -1600)
}

func TestISPMT_LargeLoanAmount(t *testing.T) {
	// ISPMT(0.05/12, 6, 360, 500000000)
	// = 500000000 * (0.05/12) * (6/360 - 1) = 500000000 * 0.004167 * (-0.98333)
	// = -2048611.11
	v, err := fnIspmt(numArgs(0.05/12, 6, 360, 500000000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT large loan", v, -2048611.11)
}

func TestISPMT_SmallLoanAmount(t *testing.T) {
	// ISPMT(0.06/12, 3, 12, 1000)
	// = 1000 * (0.06/12) * (3/12 - 1) = 1000 * 0.005 * (-0.75) = -3.75
	v, err := fnIspmt(numArgs(0.06/12, 3, 12, 1000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT small loan", v, -3.75)
}

func TestISPMT_NegativePV(t *testing.T) {
	// Negative PV (investment): ISPMT(0.1/12, 1, 36, -8000000)
	// = -8000000 * (0.1/12) * (1/36 - 1) = 64814.81
	v, err := fnIspmt(numArgs(0.1/12, 1, 36, -8000000))
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT negative pv", v, 64814.81)
}

func TestISPMT_NperZero(t *testing.T) {
	// nper=0 → #DIV/0!
	v, _ := fnIspmt(numArgs(0.1, 1, 0, 10000))
	assertError(t, "ISPMT nper zero", v)
	if v.Err != ErrValDIV0 {
		t.Errorf("expected #DIV/0!, got err=%v", v.Err)
	}
}

func TestISPMT_Comprehensive(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// Doc example
		{"doc example", numArgs(0.1/12, 1, 36, 8000000), -64814.81, false},
		// First period (per=0)
		{"per=0", numArgs(0.1/12, 0, 36, 8000000), -66666.67, false},
		// Last period (per=nper-1)
		{"per=nper-1", numArgs(0.1/12, 35, 36, 8000000), -1851.85, false},
		// Middle period
		{"middle period", numArgs(0.1/12, 18, 36, 8000000), -33333.33, false},
		// Zero rate
		{"zero rate", numArgs(0, 5, 12, 100000), 0, false},
		// Negative rate
		{"negative rate", numArgs(-0.05/12, 1, 24, 50000), 199.65, false},
		// Annual payments
		{"annual 8% per=1", numArgs(0.08, 1, 5, 100000), -6400, false},
		// Negative PV (investment)
		{"negative pv", numArgs(0.1/12, 1, 36, -8000000), 64814.81, false},
		// Small loan
		{"small loan", numArgs(0.06/12, 3, 12, 1000), -3.75, false},
		// Very small rate
		{"very small rate", numArgs(0.001, 0, 10, 50000), -50, false},
		// Large rate
		{"large rate", numArgs(1.0, 0, 4, 10000), -10000, false},
		// nper=1, per=0
		{"nper=1 per=0", numArgs(0.1, 0, 1, 10000), -1000, false},
		// Per equals nper (beyond last 0-based period)
		{"per equals nper", numArgs(0.1, 36, 36, 8000000), 0, false},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnIspmt(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestISPMT_Errors(t *testing.T) {
	tests := []struct {
		name string
		args []Value
	}{
		// nper=0 → #DIV/0!
		{"nper zero", numArgs(0.1, 1, 0, 10000)},
		// Too few args
		{"too few args 0", []Value{}},
		{"too few args 1", numArgs(0.1)},
		{"too few args 2", numArgs(0.1, 1)},
		{"too few args 3", numArgs(0.1, 1, 36)},
		// Too many args
		{"too many args", numArgs(0.1, 1, 36, 8000000, 99)},
		// Non-numeric string
		{"non-numeric rate", []Value{StringVal("abc"), NumberVal(1), NumberVal(36), NumberVal(8000000)}},
		{"non-numeric per", []Value{NumberVal(0.1), StringVal("xyz"), NumberVal(36), NumberVal(8000000)}},
		{"non-numeric nper", []Value{NumberVal(0.1), NumberVal(1), StringVal("abc"), NumberVal(8000000)}},
		{"non-numeric pv", []Value{NumberVal(0.1), NumberVal(1), NumberVal(36), StringVal("abc")}},
		// Error propagation
		{"error in rate", []Value{ErrorVal(ErrValNUM), NumberVal(1), NumberVal(36), NumberVal(8000000)}},
		{"error in per", []Value{NumberVal(0.1), ErrorVal(ErrValDIV0), NumberVal(36), NumberVal(8000000)}},
		{"error in nper", []Value{NumberVal(0.1), NumberVal(1), ErrorVal(ErrValVALUE), NumberVal(8000000)}},
		{"error in pv", []Value{NumberVal(0.1), NumberVal(1), NumberVal(36), ErrorVal(ErrValNUM)}},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnIspmt(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			assertError(t, tc.name, v)
		})
	}
}

func TestISPMT_StringCoercion(t *testing.T) {
	// Numeric strings should be coerced: ISPMT("0.08", "1", "5", "100000") = -6400
	v, err := fnIspmt([]Value{StringVal("0.08"), StringVal("1"), StringVal("5"), StringVal("100000")})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT string coercion", v, -6400)
}

func TestISPMT_BoolCoercion(t *testing.T) {
	// TRUE=1, FALSE=0: ISPMT(TRUE, FALSE, TRUE, TRUE) = ISPMT(1, 0, 1, 1)
	// = 1 * 1 * (0/1 - 1) = -1
	v, err := fnIspmt([]Value{boolArg(true), boolArg(false), boolArg(true), boolArg(true)})
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "ISPMT bool coercion", v, -1)
}

// === TBILLPRICE ===

// Date serial number constants for tests:
// 2008-01-01 => 39448
// 2008-03-30 => 39537
// 2008-03-31 => 39538
// 2008-04-01 => 39539
// 2008-04-30 => 39568
// 2008-05-01 => 39569
// 2008-06-01 => 39600
// 2008-07-01 => 39630
// 2008-09-30 => 39721
// 2008-12-30 => 39812
// 2008-12-31 => 39813
// 2009-01-01 => 39814
// 2009-03-30 => 39902
// 2009-03-31 => 39903
// 2009-04-01 => 39904

func TestTBILLPRICE_Comprehensive(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// Doc example: settlement=3/31/2008, maturity=6/1/2008, discount=9%
		// DSM=62, price = 100*(1-0.09*62/360) = 98.45
		{
			name: "doc example",
			args: numArgs(39538, 39600, 0.09),
			want: 98.45,
		},
		// 30-day T-bill, 5% discount: DSM=30, price = 100*(1-0.05*30/360) = 99.5833...
		{
			name: "30 day 5%",
			args: numArgs(39538, 39568, 0.05),
			want: 99.5833,
		},
		// 90-day T-bill, 3% discount: DSM=92 (4/1 to 7/1), price = 100*(1-0.03*92/360) = 99.2333...
		{
			name: "92 day 3%",
			args: numArgs(39539, 39630, 0.03),
			want: 99.2333,
		},
		// 182-day T-bill, 4% discount: DSM=182 (4/1 to 9/30), price = 100*(1-0.04*182/360) = 97.9778
		{
			name: "182 day 4%",
			args: numArgs(39539, 39721, 0.04),
			want: 97.9778,
		},
		// 364-day T-bill, 2% discount: DSM=364 (1/1/2008 to 12/30/2008), price = 100*(1-0.02*364/360) = 97.9778
		{
			name: "364 day 2%",
			args: numArgs(39448, 39812, 0.02),
			want: 97.9778,
		},
		// 1-day T-bill: DSM=1, discount=10%, price = 100*(1-0.10*1/360) = 99.9722
		{
			name: "1 day 10%",
			args: numArgs(39538, 39539, 0.10),
			want: 99.9722,
		},
		// Very small discount: 0.001 (0.1%), DSM=62
		// price = 100*(1-0.001*62/360) = 99.9828
		{
			name: "very small discount",
			args: numArgs(39538, 39600, 0.001),
			want: 99.9828,
		},
		// Large discount: 50%, DSM=62
		// price = 100*(1-0.50*62/360) = 91.3889
		{
			name: "large discount 50%",
			args: numArgs(39538, 39600, 0.50),
			want: 91.3889,
		},
		// Very large discount: 100%, DSM=62
		// price = 100*(1-1.0*62/360) = 82.7778
		{
			name: "very large discount 100%",
			args: numArgs(39538, 39600, 1.0),
			want: 82.7778,
		},
		// Exactly 365 days (one year): 3/31/2008 to 3/31/2009, DSM=365
		// price = 100*(1-0.05*365/360) = 94.9306
		{
			name: "exactly one year",
			args: numArgs(39538, 39903, 0.05),
			want: 94.9306,
		},
		// 2% discount, DSM=31 (3/31 to 4/30)
		// price = 100*(1-0.02*31/360) = 99.8278
		{
			name: "31 day 2%",
			args: numArgs(39538, 39569, 0.02),
			want: 99.8278,
		},
		// High precision check: 7.5% discount, DSM=62
		// price = 100*(1-0.075*62/360) = 98.7083...
		{
			name: "7.5% discount 62 day",
			args: numArgs(39538, 39600, 0.075),
			want: 98.7083,
		},
		// Fractional serial numbers should be truncated
		// 39538.7 => 39538, 39600.9 => 39600, DSM=62
		{
			name: "fractional dates truncated",
			args: numArgs(39538.7, 39600.9, 0.09),
			want: 98.45,
		},
		// DSM=275 (1/1/2008 to 9/30/2008): 0.06 discount
		// price = 100*(1-0.06*273/360) = 95.45
		{
			name: "273 day 6%",
			args: numArgs(39448, 39721, 0.06),
			want: 95.45,
		},
		// --- Error cases ---
		// settlement == maturity
		{
			name:    "settlement equals maturity",
			args:    numArgs(39538, 39538, 0.05),
			wantErr: true,
		},
		// settlement > maturity
		{
			name:    "settlement after maturity",
			args:    numArgs(39600, 39538, 0.05),
			wantErr: true,
		},
		// maturity more than one year after settlement
		{
			name:    "maturity more than one year",
			args:    numArgs(39538, 39904, 0.05),
			wantErr: true,
		},
		// discount <= 0
		{
			name:    "discount zero",
			args:    numArgs(39538, 39600, 0),
			wantErr: true,
		},
		{
			name:    "discount negative",
			args:    numArgs(39538, 39600, -0.05),
			wantErr: true,
		},
		// wrong number of arguments
		{
			name:    "too few args",
			args:    numArgs(39538, 39600),
			wantErr: true,
		},
		{
			name:    "too many args",
			args:    numArgs(39538, 39600, 0.09, 1),
			wantErr: true,
		},
		// non-numeric args
		{
			name:    "non-numeric settlement",
			args:    []Value{StringVal("abc"), NumberVal(39600), NumberVal(0.09)},
			wantErr: true,
		},
		{
			name:    "non-numeric maturity",
			args:    []Value{NumberVal(39538), StringVal("xyz"), NumberVal(0.09)},
			wantErr: true,
		},
		{
			name:    "non-numeric discount",
			args:    []Value{NumberVal(39538), NumberVal(39600), StringVal("abc")},
			wantErr: true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnTbillPrice(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

// === TBILLYIELD ===

func TestTBILLYIELD_Comprehensive(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// Doc example: settlement=3/31/2008, maturity=6/1/2008, pr=$98.45
		// DSM=62, yield = ((100-98.45)/98.45)*(360/62) = 0.09141...
		{
			name: "doc example",
			args: numArgs(39538, 39600, 98.45),
			want: 0.09141,
		},
		// 30-day T-bill, pr=99.50: DSM=30
		// yield = ((100-99.50)/99.50)*(360/30) = 0.005025*12 = 0.06030
		{
			name: "30 day pr 99.50",
			args: numArgs(39538, 39568, 99.50),
			want: 0.06030,
		},
		// 92-day T-bill, pr=99.00: DSM=91 (4/1 to 7/1)
		// yield = ((100-99)/99)*(360/91) = 0.01010*3.95604 = 0.03996
		{
			name: "91 day pr 99.00",
			args: numArgs(39539, 39630, 99.00),
			want: 0.03957,
		},
		// 183-day T-bill, pr=97.50: DSM=182 (4/1 to 9/30)
		// yield = ((100-97.50)/97.50)*(360/183) = 0.02564*1.9672 = 0.05043
		{
			name: "182 day pr 97.50",
			args: numArgs(39539, 39721, 97.50),
			want: 0.05043,
		},
		// 364-day T-bill, pr=95.00: DSM=364 (1/1/2008 to 12/30/2008)
		// yield = ((100-95)/95)*(360/364) = 0.05263*0.9890 = 0.05205
		{
			name: "364 day pr 95",
			args: numArgs(39448, 39812, 95),
			want: 0.05205,
		},
		// 1-day T-bill, pr=99.99: DSM=1
		// yield = ((100-99.99)/99.99)*(360/1) = 0.0001*360 = 0.03601
		{
			name: "1 day pr 99.99",
			args: numArgs(39538, 39539, 99.99),
			want: 0.03601,
		},
		// pr=50 (very low): DSM=62
		// yield = ((100-50)/50)*(360/62) = 1.0*5.80645 = 5.80645
		{
			name: "very low price",
			args: numArgs(39538, 39600, 50),
			want: 5.80645,
		},
		// pr=99.99 (very high): DSM=62
		// yield = ((100-99.99)/99.99)*(360/62) = 0.0001*5.80645 = 0.000581
		{
			name: "very high price",
			args: numArgs(39538, 39600, 99.99),
			want: 0.000581,
		},
		// pr=1 (extremely low): DSM=62
		// yield = ((100-1)/1)*(360/62) = 99*5.80645 = 574.84
		{
			name: "extremely low price",
			args: numArgs(39538, 39600, 1),
			want: 574.84,
		},
		// pr=0.01 (near zero): DSM=62
		// yield = ((100-0.01)/0.01)*(360/62) = 9999*5.80645 = 58058.71
		{
			name: "near zero price",
			args: numArgs(39538, 39600, 0.01),
			want: 58058.71,
		},
		// Exactly one year (365 days): pr=95
		// yield = ((100-95)/95)*(360/365) = 0.05263*0.98630 = 0.05191
		{
			name: "exactly one year",
			args: numArgs(39538, 39903, 95),
			want: 0.05191,
		},
		// Fractional serial numbers truncated
		{
			name: "fractional dates truncated",
			args: numArgs(39538.5, 39600.9, 98.45),
			want: 0.09141,
		},
		// pr=100 (par): DSM=62
		// yield = ((100-100)/100)*(360/62) = 0
		{
			name: "price at par",
			args: numArgs(39538, 39600, 100),
			want: 0,
		},
		// pr > 100 (premium): DSM=62
		// yield = ((100-105)/105)*(360/62) = -0.04762*5.80645 = -0.27650
		{
			name: "price above par (negative yield)",
			args: numArgs(39538, 39600, 105),
			want: -0.27650,
		},
		// --- Error cases ---
		// settlement == maturity
		{
			name:    "settlement equals maturity",
			args:    numArgs(39538, 39538, 98.45),
			wantErr: true,
		},
		// settlement > maturity
		{
			name:    "settlement after maturity",
			args:    numArgs(39600, 39538, 98.45),
			wantErr: true,
		},
		// maturity more than one year after settlement
		{
			name:    "maturity more than one year",
			args:    numArgs(39538, 39904, 98.45),
			wantErr: true,
		},
		// pr <= 0
		{
			name:    "price zero",
			args:    numArgs(39538, 39600, 0),
			wantErr: true,
		},
		{
			name:    "price negative",
			args:    numArgs(39538, 39600, -10),
			wantErr: true,
		},
		// wrong number of arguments
		{
			name:    "too few args",
			args:    numArgs(39538, 39600),
			wantErr: true,
		},
		{
			name:    "too many args",
			args:    numArgs(39538, 39600, 98.45, 1),
			wantErr: true,
		},
		// non-numeric args
		{
			name:    "non-numeric settlement",
			args:    []Value{StringVal("abc"), NumberVal(39600), NumberVal(98.45)},
			wantErr: true,
		},
		{
			name:    "non-numeric maturity",
			args:    []Value{NumberVal(39538), StringVal("xyz"), NumberVal(98.45)},
			wantErr: true,
		},
		{
			name:    "non-numeric price",
			args:    []Value{NumberVal(39538), NumberVal(39600), StringVal("abc")},
			wantErr: true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnTbillYield(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

// === TBILLEQ ===

func TestTBILLEQ_Comprehensive(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// Doc example: settlement=3/31/2008, maturity=6/1/2008, discount=9.14%
		// DSM=62, TBILLEQ = (365*0.0914)/(360-0.0914*62) = 33.361/(360-5.6668) = 33.361/354.3332 = 0.09415
		{
			name: "doc example",
			args: numArgs(39538, 39600, 0.0914),
			want: 0.09415,
		},
		// 30-day, 5% discount: DSM=30
		// TBILLEQ = (365*0.05)/(360-0.05*30) = 18.25/(360-1.5) = 18.25/358.5 = 0.05091
		{
			name: "30 day 5%",
			args: numArgs(39538, 39568, 0.05),
			want: 0.05091,
		},
		// 91-day, 3% discount: DSM=91 (4/1 to 7/1)
		// TBILLEQ = (365*0.03)/(360-0.03*91) = 10.95/(360-2.73) = 10.95/357.27 = 0.03065
		{
			name: "91 day 3%",
			args: numArgs(39539, 39630, 0.03),
			want: 0.03065,
		},
		// 182-day, 4% discount: DSM=182 (4/1 to 9/30)
		// TBILLEQ = (365*0.04)/(360-0.04*182) = 14.6/(360-7.28) = 14.6/352.72 = 0.04139
		{
			name: "182 day 4%",
			args: numArgs(39539, 39721, 0.04),
			want: 0.04139,
		},
		// 364-day, 2% discount: DSM=364 (1/1/2008 to 12/30/2008)
		// TBILLEQ = (365*0.02)/(360-0.02*364) = 7.3/(360-7.28) = 7.3/352.72 = 0.02070
		{
			name: "364 day 2%",
			args: numArgs(39448, 39812, 0.02),
			want: 0.02070,
		},
		// 1-day, 10% discount: DSM=1
		// TBILLEQ = (365*0.10)/(360-0.10*1) = 36.5/359.9 = 0.10142
		{
			name: "1 day 10%",
			args: numArgs(39538, 39539, 0.10),
			want: 0.10142,
		},
		// Very small discount: 0.001, DSM=62
		// TBILLEQ = (365*0.001)/(360-0.001*62) = 0.365/359.938 = 0.001014
		{
			name: "very small discount",
			args: numArgs(39538, 39600, 0.001),
			want: 0.001014,
		},
		// Large discount: 50%, DSM=62
		// TBILLEQ = (365*0.50)/(360-0.50*62) = 182.5/(360-31) = 182.5/329 = 0.55471
		{
			name: "large discount 50%",
			args: numArgs(39538, 39600, 0.50),
			want: 0.55471,
		},
		// Very large discount: 100%, DSM=62
		// TBILLEQ = (365*1.0)/(360-1.0*62) = 365/298 = 1.22483
		{
			name: "very large discount 100%",
			args: numArgs(39538, 39600, 1.0),
			want: 1.22483,
		},
		// Exactly one year (365 days): 5% discount
		// TBILLEQ = (365*0.05)/(360-0.05*365) = 18.25/(360-18.25) = 18.25/341.75 = 0.05339
		{
			name: "exactly one year 5%",
			args: numArgs(39538, 39903, 0.05),
			want: 0.05339,
		},
		// 12% discount, DSM=62
		// TBILLEQ = (365*0.12)/(360-0.12*62) = 43.8/(360-7.44) = 43.8/352.56 = 0.12425
		{
			name: "12% discount 62 day",
			args: numArgs(39538, 39600, 0.12),
			want: 0.12425,
		},
		// Fractional dates truncated: same as doc example
		{
			name: "fractional dates truncated",
			args: numArgs(39538.7, 39600.9, 0.0914),
			want: 0.09415,
		},
		// settlement == maturity (allowed for TBILLEQ, not strict less)
		// DSM=0, TBILLEQ = (365*0.05)/(360-0.05*0) = 18.25/360 = 0.05069
		{
			name: "settlement equals maturity",
			args: numArgs(39538, 39538, 0.05),
			want: 0.05069,
		},
		// 15% discount, DSM=62
		// TBILLEQ = (365*0.15)/(360-0.15*62) = 54.75/(360-9.3) = 54.75/350.7 = 0.15611
		{
			name: "15% discount",
			args: numArgs(39538, 39600, 0.15),
			want: 0.15611,
		},
		// 0.5% discount, DSM=62
		// TBILLEQ = (365*0.005)/(360-0.005*62) = 1.825/(360-0.31) = 1.825/359.69 = 0.005074
		{
			name: "0.5% discount",
			args: numArgs(39538, 39600, 0.005),
			want: 0.005074,
		},
		// --- Error cases ---
		// settlement > maturity (strictly greater)
		{
			name:    "settlement after maturity",
			args:    numArgs(39600, 39538, 0.05),
			wantErr: true,
		},
		// maturity more than one year after settlement
		{
			name:    "maturity more than one year",
			args:    numArgs(39538, 39904, 0.05),
			wantErr: true,
		},
		// discount <= 0
		{
			name:    "discount zero",
			args:    numArgs(39538, 39600, 0),
			wantErr: true,
		},
		{
			name:    "discount negative",
			args:    numArgs(39538, 39600, -0.05),
			wantErr: true,
		},
		// wrong number of arguments
		{
			name:    "too few args",
			args:    numArgs(39538, 39600),
			wantErr: true,
		},
		{
			name:    "too many args",
			args:    numArgs(39538, 39600, 0.0914, 1),
			wantErr: true,
		},
		// non-numeric args
		{
			name:    "non-numeric settlement",
			args:    []Value{StringVal("abc"), NumberVal(39600), NumberVal(0.0914)},
			wantErr: true,
		},
		{
			name:    "non-numeric maturity",
			args:    []Value{NumberVal(39538), StringVal("xyz"), NumberVal(0.0914)},
			wantErr: true,
		},
		{
			name:    "non-numeric discount",
			args:    []Value{NumberVal(39538), NumberVal(39600), StringVal("abc")},
			wantErr: true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnTbillEq(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

// === TBILLPRICE / TBILLYIELD / TBILLEQ via formula evaluation ===

func TestTBILLPRICE_ViaEval(t *testing.T) {
	// TBILLPRICE(DATE(2008,3,31), DATE(2008,6,1), 0.09)
	cf := evalCompile(t, "TBILLPRICE(DATE(2008,3,31), DATE(2008,6,1), 0.09)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "TBILLPRICE via eval", v, 98.45)
}

func TestTBILLYIELD_ViaEval(t *testing.T) {
	// TBILLYIELD(DATE(2008,3,31), DATE(2008,6,1), 98.45)
	cf := evalCompile(t, "TBILLYIELD(DATE(2008,3,31), DATE(2008,6,1), 98.45)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "TBILLYIELD via eval", v, 0.09141)
}

func TestTBILLEQ_ViaEval(t *testing.T) {
	// TBILLEQ(DATE(2008,3,31), DATE(2008,6,1), 0.0914)
	cf := evalCompile(t, "TBILLEQ(DATE(2008,3,31), DATE(2008,6,1), 0.0914)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	assertClose(t, "TBILLEQ via eval", v, 0.09415)
}

// === DISC ===

func TestDISC_Comprehensive(t *testing.T) {
	// Serial numbers for reference dates:
	// 7/1/2018 = 43282, 1/1/2038 = 50406
	// 2/15/2008 = 39493, 5/15/2008 = 39583
	// 1/1/2023 = 44927, 12/31/2023 = 45291
	// 1/1/2024 = 45292, 7/1/2024 = 45474
	// 1/31/2023 = 44957

	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
		tol     float64 // custom tolerance; 0 means use default 0.01
	}{
		// --- Doc example ---
		// DISC(7/1/2018, 1/1/2038, 97.975, 100, 1) = 0.001038
		{
			name: "doc example basis 1",
			args: numArgs(43282, 50406, 97.975, 100, 1),
			want: 0.001038,
			tol:  0.000001,
		},
		// --- All 5 basis types with same dates (2/15/2008 to 5/15/2008, pr=97, redemption=100) ---
		// basis 0: DSM=days360(US)=90, B=360 => DISC=(3/100)*(360/90)=0.12
		{
			name: "basis 0 US 30/360",
			args: numArgs(39493, 39583, 97, 100, 0),
			want: 0.12,
		},
		// basis 1: DSM=90 actual, B=366 (2008 is leap) => DISC=(3/100)*(366/90)=0.12200
		{
			name: "basis 1 actual/actual",
			args: numArgs(39493, 39583, 97, 100, 1),
			want: 0.12200,
			tol:  0.001,
		},
		// basis 2: DSM=90 actual, B=360 => DISC=(3/100)*(360/90)=0.12
		{
			name: "basis 2 actual/360",
			args: numArgs(39493, 39583, 97, 100, 2),
			want: 0.12,
		},
		// basis 3: DSM=90 actual, B=365 => DISC=(3/100)*(365/90)=0.12167
		{
			name: "basis 3 actual/365",
			args: numArgs(39493, 39583, 97, 100, 3),
			want: 0.12167,
			tol:  0.001,
		},
		// basis 4: European 30/360: DSM=days360(EU)=90, B=360 => DISC=(3/100)*(360/90)=0.12
		{
			name: "basis 4 European 30/360",
			args: numArgs(39493, 39583, 97, 100, 4),
			want: 0.12,
		},
		// --- Default basis (omitted = 0) ---
		{
			name: "default basis omitted",
			args: numArgs(39493, 39583, 97, 100),
			want: 0.12,
		},
		// --- Various date ranges ---
		// Short range: 30 days, basis 0
		// 1/1/2023 to 1/31/2023, pr=99, redemption=100
		// DSM=30, B=360 => DISC=(1/100)*(360/30)=0.12
		{
			name: "short 30 day range",
			args: numArgs(44927, 44957, 99, 100, 0),
			want: 0.12,
		},
		// Long range: full year, basis 0
		// 1/1/2023 to 12/31/2023, pr=95, redemption=100
		// DSM=360 (US 30/360), B=360 => DISC=(5/100)*(360/360)=0.05
		{
			name: "full year range",
			args: numArgs(44927, 45291, 95, 100, 0),
			want: 0.05,
		},
		// Half year, basis 2 (actual/360)
		// 1/1/2024 to 7/1/2024, pr=98, redemption=100
		// DSM=182 actual, B=360 => DISC=(2/100)*(360/182)=0.03956
		{
			name: "half year basis 2",
			args: numArgs(45292, 45474, 98, 100, 2),
			want: 0.03956,
			tol:  0.001,
		},
		// --- Edge cases: very small discount ---
		// pr=99.999, redemption=100
		{
			name: "very small discount",
			args: numArgs(39493, 39583, 99.999, 100, 0),
			want: 0.00004,
			tol:  0.00001,
		},
		// --- Edge cases: very large discount ---
		// pr=1, redemption=100
		// DSM=90, B=360 => DISC=(99/100)*(360/90)=3.96
		{
			name: "very large discount",
			args: numArgs(39493, 39583, 1, 100, 0),
			want: 3.96,
		},
		// pr > redemption (negative result)
		// pr=105, redemption=100
		// DISC=(100-105)/100*(360/90) = -0.2
		{
			name: "pr exceeds redemption negative result",
			args: numArgs(39493, 39583, 105, 100, 0),
			want: -0.2,
		},
		// Fractional dates get truncated
		{
			name: "fractional dates truncated",
			args: numArgs(39493.7, 39583.9, 97, 100, 0),
			want: 0.12,
		},
		// Large redemption value
		{
			name: "large redemption",
			args: numArgs(39493, 39583, 97000, 100000, 0),
			want: 0.12,
		},
		// --- Error cases ---
		{
			name:    "pr zero",
			args:    numArgs(39493, 39583, 0, 100, 0),
			wantErr: true,
		},
		{
			name:    "pr negative",
			args:    numArgs(39493, 39583, -10, 100, 0),
			wantErr: true,
		},
		{
			name:    "redemption zero",
			args:    numArgs(39493, 39583, 97, 0, 0),
			wantErr: true,
		},
		{
			name:    "redemption negative",
			args:    numArgs(39493, 39583, 97, -100, 0),
			wantErr: true,
		},
		{
			name:    "settlement equals maturity",
			args:    numArgs(39493, 39493, 97, 100, 0),
			wantErr: true,
		},
		{
			name:    "settlement after maturity",
			args:    numArgs(39583, 39493, 97, 100, 0),
			wantErr: true,
		},
		{
			name:    "basis negative",
			args:    numArgs(39493, 39583, 97, 100, -1),
			wantErr: true,
		},
		{
			name:    "basis 5",
			args:    numArgs(39493, 39583, 97, 100, 5),
			wantErr: true,
		},
		{
			name:    "basis 99",
			args:    numArgs(39493, 39583, 97, 100, 99),
			wantErr: true,
		},
		{
			name:    "too few args",
			args:    numArgs(39493, 39583, 97),
			wantErr: true,
		},
		{
			name:    "too many args",
			args:    numArgs(39493, 39583, 97, 100, 0, 1),
			wantErr: true,
		},
		{
			name:    "non-numeric settlement",
			args:    []Value{StringVal("abc"), NumberVal(39583), NumberVal(97), NumberVal(100), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric maturity",
			args:    []Value{NumberVal(39493), StringVal("xyz"), NumberVal(97), NumberVal(100), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric pr",
			args:    []Value{NumberVal(39493), NumberVal(39583), StringVal("abc"), NumberVal(100), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric redemption",
			args:    []Value{NumberVal(39493), NumberVal(39583), NumberVal(97), StringVal("abc"), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric basis",
			args:    []Value{NumberVal(39493), NumberVal(39583), NumberVal(97), NumberVal(100), StringVal("abc")},
			wantErr: true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnDisc(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				tol := tc.tol
				if tol == 0 {
					tol = 0.01
				}
				if v.Type != ValueNumber {
					t.Fatalf("%s: expected number, got type %v (str=%q)", tc.name, v.Type, v.Str)
				}
				if math.Abs(v.Num-tc.want) > tol {
					t.Errorf("%s: got %f, want %f (tol=%g)", tc.name, v.Num, tc.want, tol)
				}
			}
		})
	}
}

func TestDISC_ViaEval(t *testing.T) {
	// DISC(DATE(2018,7,1), DATE(2038,1,1), 97.975, 100, 1) = 0.001038
	cf := evalCompile(t, "DISC(DATE(2018,7,1), DATE(2038,1,1), 97.975, 100, 1)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-0.001038) > 0.000001 {
		t.Errorf("got %f, want 0.001038", v.Num)
	}
}

// === INTRATE ===

func TestINTRATE_Comprehensive(t *testing.T) {
	// Serial numbers:
	// 2/15/2008 = 39493, 5/15/2008 = 39583
	// 1/1/2023 = 44927, 12/31/2023 = 45291
	// 1/1/2024 = 45292, 7/1/2024 = 45474

	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
		tol     float64
	}{
		// --- Doc example ---
		// INTRATE(2/15/2008, 5/15/2008, 1000000, 1014420, 2) = 0.05768
		{
			name: "doc example basis 2",
			args: numArgs(39493, 39583, 1000000, 1014420, 2),
			want: 0.05768,
			tol:  0.0001,
		},
		// --- All 5 basis types (2/15/2008 to 5/15/2008, inv=1000, red=1010) ---
		// basis 0: DIM=days360(US)=90, B=360 => INTRATE=(10/1000)*(360/90)=0.04
		{
			name: "basis 0 US 30/360",
			args: numArgs(39493, 39583, 1000, 1010, 0),
			want: 0.04,
		},
		// basis 1: DIM=90 actual, B=366 (2008 is leap) => INTRATE=(10/1000)*(366/90)=0.04067
		{
			name: "basis 1 actual/actual",
			args: numArgs(39493, 39583, 1000, 1010, 1),
			want: 0.04067,
			tol:  0.001,
		},
		// basis 2: DIM=90 actual, B=360 => INTRATE=(10/1000)*(360/90)=0.04
		{
			name: "basis 2 actual/360",
			args: numArgs(39493, 39583, 1000, 1010, 2),
			want: 0.04,
		},
		// basis 3: DIM=90 actual, B=365 => INTRATE=(10/1000)*(365/90)=0.04056
		{
			name: "basis 3 actual/365",
			args: numArgs(39493, 39583, 1000, 1010, 3),
			want: 0.04056,
			tol:  0.001,
		},
		// basis 4: DIM=days360(EU)=90, B=360 => INTRATE=(10/1000)*(360/90)=0.04
		{
			name: "basis 4 European 30/360",
			args: numArgs(39493, 39583, 1000, 1010, 4),
			want: 0.04,
		},
		// --- Default basis (omitted = 0) ---
		{
			name: "default basis omitted",
			args: numArgs(39493, 39583, 1000, 1010),
			want: 0.04,
		},
		// --- Various date ranges ---
		// Short range: 30 days
		{
			name: "short 30 day range",
			args: numArgs(44927, 44957, 10000, 10100, 0),
			want: 0.12,
		},
		// Full year
		{
			name: "full year basis 0",
			args: numArgs(44927, 45291, 10000, 10500, 0),
			want: 0.05,
		},
		// Half year, basis 3
		// 1/1/2024 to 7/1/2024: DIM=182 actual, B=365
		// INTRATE=(200/10000)*(365/182)=0.02*2.00549=0.04011
		{
			name: "half year basis 3",
			args: numArgs(45292, 45474, 10000, 10200, 3),
			want: 0.04011,
			tol:  0.001,
		},
		// --- Edge cases ---
		// Very small return
		{
			name: "very small return",
			args: numArgs(39493, 39583, 1000000, 1000001, 0),
			want: 0.000004,
			tol:  0.000001,
		},
		// Very large return
		{
			name: "very large return",
			args: numArgs(39493, 39583, 100, 10000, 0),
			want: 396.0,
		},
		// Negative return (redemption < investment)
		{
			name: "negative return",
			args: numArgs(39493, 39583, 1000, 900, 0),
			want: -0.4,
		},
		// Large investment values
		{
			name: "large values",
			args: numArgs(39493, 39583, 50000000, 51000000, 2),
			want: 0.08,
		},
		// Fractional dates truncated
		{
			name: "fractional dates truncated",
			args: numArgs(39493.8, 39583.2, 1000, 1010, 0),
			want: 0.04,
		},
		// --- Error cases ---
		{
			name:    "investment zero",
			args:    numArgs(39493, 39583, 0, 1010, 0),
			wantErr: true,
		},
		{
			name:    "investment negative",
			args:    numArgs(39493, 39583, -1000, 1010, 0),
			wantErr: true,
		},
		{
			name:    "redemption zero",
			args:    numArgs(39493, 39583, 1000, 0, 0),
			wantErr: true,
		},
		{
			name:    "redemption negative",
			args:    numArgs(39493, 39583, 1000, -1010, 0),
			wantErr: true,
		},
		{
			name:    "settlement equals maturity",
			args:    numArgs(39493, 39493, 1000, 1010, 0),
			wantErr: true,
		},
		{
			name:    "settlement after maturity",
			args:    numArgs(39583, 39493, 1000, 1010, 0),
			wantErr: true,
		},
		{
			name:    "basis negative",
			args:    numArgs(39493, 39583, 1000, 1010, -1),
			wantErr: true,
		},
		{
			name:    "basis 5",
			args:    numArgs(39493, 39583, 1000, 1010, 5),
			wantErr: true,
		},
		{
			name:    "basis 99",
			args:    numArgs(39493, 39583, 1000, 1010, 99),
			wantErr: true,
		},
		{
			name:    "too few args",
			args:    numArgs(39493, 39583, 1000),
			wantErr: true,
		},
		{
			name:    "too many args",
			args:    numArgs(39493, 39583, 1000, 1010, 0, 1),
			wantErr: true,
		},
		{
			name:    "non-numeric settlement",
			args:    []Value{StringVal("abc"), NumberVal(39583), NumberVal(1000), NumberVal(1010), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric maturity",
			args:    []Value{NumberVal(39493), StringVal("xyz"), NumberVal(1000), NumberVal(1010), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric investment",
			args:    []Value{NumberVal(39493), NumberVal(39583), StringVal("abc"), NumberVal(1010), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric redemption",
			args:    []Value{NumberVal(39493), NumberVal(39583), NumberVal(1000), StringVal("abc"), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric basis",
			args:    []Value{NumberVal(39493), NumberVal(39583), NumberVal(1000), NumberVal(1010), StringVal("abc")},
			wantErr: true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnIntrate(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				tol := tc.tol
				if tol == 0 {
					tol = 0.01
				}
				if v.Type != ValueNumber {
					t.Fatalf("%s: expected number, got type %v (str=%q)", tc.name, v.Type, v.Str)
				}
				if math.Abs(v.Num-tc.want) > tol {
					t.Errorf("%s: got %f, want %f (tol=%g)", tc.name, v.Num, tc.want, tol)
				}
			}
		})
	}
}

func TestINTRATE_ViaEval(t *testing.T) {
	// INTRATE(DATE(2008,2,15), DATE(2008,5,15), 1000000, 1014420, 2) = 0.05768
	cf := evalCompile(t, "INTRATE(DATE(2008,2,15), DATE(2008,5,15), 1000000, 1014420, 2)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-0.05768) > 0.0001 {
		t.Errorf("got %f, want 0.05768", v.Num)
	}
}

// === RECEIVED ===

func TestRECEIVED_Comprehensive(t *testing.T) {
	// Serial numbers:
	// 2/15/2008 = 39493, 5/15/2008 = 39583
	// 1/1/2023 = 44927, 12/31/2023 = 45291
	// 1/1/2024 = 45292, 7/1/2024 = 45474

	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
		tol     float64
	}{
		// --- Doc example ---
		// RECEIVED(2/15/2008, 5/15/2008, 1000000, 0.0575, 2) = 1014584.65
		{
			name: "doc example basis 2",
			args: numArgs(39493, 39583, 1000000, 0.0575, 2),
			want: 1014584.65,
			tol:  1.0,
		},
		// --- All 5 basis types (2/15/2008 to 5/15/2008, inv=10000, disc=0.10) ---
		// basis 0: DIM=days360(US)=90, B=360 => RECEIVED=10000/(1-0.10*90/360)=10000/0.975=10256.41
		{
			name: "basis 0 US 30/360",
			args: numArgs(39493, 39583, 10000, 0.10, 0),
			want: 10256.41,
			tol:  1.0,
		},
		// basis 1: DIM=90 actual, B=366 (2008 leap) => RECEIVED=10000/(1-0.10*90/366)=10000/0.97541=10252.18
		{
			name: "basis 1 actual/actual",
			args: numArgs(39493, 39583, 10000, 0.10, 1),
			want: 10252.18,
			tol:  1.0,
		},
		// basis 2: DIM=90 actual, B=360 => RECEIVED=10000/(1-0.10*90/360)=10000/0.975=10256.41
		{
			name: "basis 2 actual/360",
			args: numArgs(39493, 39583, 10000, 0.10, 2),
			want: 10256.41,
			tol:  1.0,
		},
		// basis 3: DIM=90 actual, B=365 => RECEIVED=10000/(1-0.10*90/365)=10000/0.97534=10252.86
		{
			name: "basis 3 actual/365",
			args: numArgs(39493, 39583, 10000, 0.10, 3),
			want: 10252.86,
			tol:  1.0,
		},
		// basis 4: DIM=days360(EU)=90, B=360 => RECEIVED=10000/(1-0.10*90/360)=10000/0.975=10256.41
		{
			name: "basis 4 European 30/360",
			args: numArgs(39493, 39583, 10000, 0.10, 4),
			want: 10256.41,
			tol:  1.0,
		},
		// --- Default basis (omitted = 0) ---
		{
			name: "default basis omitted",
			args: numArgs(39493, 39583, 10000, 0.10),
			want: 10256.41,
			tol:  1.0,
		},
		// --- Various date ranges ---
		// Short range: 30 days, basis 0
		// 1/1/2023 to 1/31/2023: DIM=30 (US 30/360), B=360
		// RECEIVED=10000/(1-0.05*30/360)=10000/0.99583=10041.84
		{
			name: "short 30 day range",
			args: numArgs(44927, 44957, 10000, 0.05, 0),
			want: 10041.84,
			tol:  1.0,
		},
		// Full year, basis 0
		// 1/1/2023 to 12/31/2023: DIM=360 (US 30/360), B=360
		// RECEIVED=10000/(1-0.05*360/360)=10000/0.95=10526.32
		{
			name: "full year basis 0",
			args: numArgs(44927, 45291, 10000, 0.05, 0),
			want: 10526.32,
			tol:  1.0,
		},
		// Half year, basis 3
		// 1/1/2024 to 7/1/2024: DIM=182 actual, B=365
		// RECEIVED=10000/(1-0.06*182/365)=10000/(1-0.02992)=10000/0.97008=10308.40
		{
			name: "half year basis 3",
			args: numArgs(45292, 45474, 10000, 0.06, 3),
			want: 10308.40,
			tol:  1.0,
		},
		// --- Edge cases ---
		// Very small discount
		{
			name: "very small discount",
			args: numArgs(39493, 39583, 10000, 0.0001, 0),
			want: 10000.25,
			tol:  1.0,
		},
		// Large investment
		{
			name: "large investment",
			args: numArgs(39493, 39583, 50000000, 0.05, 0),
			want: 50632911.39,
			tol:  100.0,
		},
		// Small investment
		{
			name: "small investment",
			args: numArgs(39493, 39583, 100, 0.08, 0),
			want: 102.04,
			tol:  0.1,
		},
		// Fractional dates truncated
		{
			name: "fractional dates truncated",
			args: numArgs(39493.7, 39583.9, 10000, 0.10, 0),
			want: 10256.41,
			tol:  1.0,
		},
		// High discount rate
		// DIM=90, B=360, discount=0.50 => RECEIVED=10000/(1-0.50*90/360)=10000/0.875=11428.57
		{
			name: "high discount rate",
			args: numArgs(39493, 39583, 10000, 0.50, 0),
			want: 11428.57,
			tol:  1.0,
		},
		// --- Error cases ---
		{
			name:    "investment zero",
			args:    numArgs(39493, 39583, 0, 0.05, 0),
			wantErr: true,
		},
		{
			name:    "investment negative",
			args:    numArgs(39493, 39583, -10000, 0.05, 0),
			wantErr: true,
		},
		{
			name:    "discount zero",
			args:    numArgs(39493, 39583, 10000, 0, 0),
			wantErr: true,
		},
		{
			name:    "discount negative",
			args:    numArgs(39493, 39583, 10000, -0.05, 0),
			wantErr: true,
		},
		{
			name:    "settlement equals maturity",
			args:    numArgs(39493, 39493, 10000, 0.05, 0),
			wantErr: true,
		},
		{
			name:    "settlement after maturity",
			args:    numArgs(39583, 39493, 10000, 0.05, 0),
			wantErr: true,
		},
		{
			name:    "basis negative",
			args:    numArgs(39493, 39583, 10000, 0.05, -1),
			wantErr: true,
		},
		{
			name:    "basis 5",
			args:    numArgs(39493, 39583, 10000, 0.05, 5),
			wantErr: true,
		},
		{
			name:    "basis 99",
			args:    numArgs(39493, 39583, 10000, 0.05, 99),
			wantErr: true,
		},
		{
			name:    "too few args",
			args:    numArgs(39493, 39583, 10000),
			wantErr: true,
		},
		{
			name:    "too many args",
			args:    numArgs(39493, 39583, 10000, 0.05, 0, 1),
			wantErr: true,
		},
		{
			name:    "non-numeric settlement",
			args:    []Value{StringVal("abc"), NumberVal(39583), NumberVal(10000), NumberVal(0.05), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric maturity",
			args:    []Value{NumberVal(39493), StringVal("xyz"), NumberVal(10000), NumberVal(0.05), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric investment",
			args:    []Value{NumberVal(39493), NumberVal(39583), StringVal("abc"), NumberVal(0.05), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric discount",
			args:    []Value{NumberVal(39493), NumberVal(39583), NumberVal(10000), StringVal("abc"), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric basis",
			args:    []Value{NumberVal(39493), NumberVal(39583), NumberVal(10000), NumberVal(0.05), StringVal("abc")},
			wantErr: true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnReceived(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				tol := tc.tol
				if tol == 0 {
					tol = 0.01
				}
				if v.Type != ValueNumber {
					t.Fatalf("%s: expected number, got type %v (str=%q)", tc.name, v.Type, v.Str)
				}
				if math.Abs(v.Num-tc.want) > tol {
					t.Errorf("%s: got %f, want %f (tol=%g)", tc.name, v.Num, tc.want, tol)
				}
			}
		})
	}
}

func TestRECEIVED_ViaEval(t *testing.T) {
	// RECEIVED(DATE(2008,2,15), DATE(2008,5,15), 1000000, 0.0575, 2) = 1014584.65
	cf := evalCompile(t, "RECEIVED(DATE(2008,2,15), DATE(2008,5,15), 1000000, 0.0575, 2)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-1014584.65) > 1.0 {
		t.Errorf("got %f, want 1014584.65", v.Num)
	}
}

// === ACCRINT ===

func TestACCRINT_Comprehensive(t *testing.T) {
	// Serial numbers for reference dates:
	// DATE(2008,3,1) = 39508, DATE(2008,8,31) = 39691, DATE(2008,5,1) = 39569
	// DATE(2008,3,5) = 39512, DATE(2008,4,5) = 39543
	// DATE(2008,1,15) = 39462, DATE(2008,7,15) = 39644, DATE(2008,4,15) = 39553
	// DATE(2009,1,15) = 39828, DATE(2008,3,15) = 39522
	// DATE(2007,1,1) = 39083, DATE(2007,7,1) = 39264, DATE(2007,4,1) = 39173
	// DATE(2008,4,1) = 39539, DATE(2008,1,1) = 39448, DATE(2008,7,1) = 39630
	// DATE(2008,6,15) = 39614

	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
		tol     float64 // custom tolerance; 0 means use default 0.01
	}{
		// --- Doc examples ---
		// ACCRINT(39508,39691,39569,0.1,1000,2,0) = 16.666667
		// issue=2008-03-01, first_interest=2008-08-31, settlement=2008-05-01
		{
			name: "doc example 1",
			args: numArgs(39508, 39691, 39569, 0.1, 1000, 2, 0),
			want: 16.666667,
			tol:  0.000001,
		},
		// ACCRINT(DATE(2008,3,5),39691,39569,0.1,1000,2,0,FALSE) = 15.555556
		{
			name: "doc example 2 calc_method FALSE",
			args: numArgs(39512, 39691, 39569, 0.1, 1000, 2, 0, 0),
			want: 15.555556,
			tol:  0.000001,
		},
		// ACCRINT(DATE(2008,4,5),39691,39569,0.1,1000,2,0,TRUE) = 7.222222
		{
			name: "doc example 3 calc_method TRUE",
			args: numArgs(39543, 39691, 39569, 0.1, 1000, 2, 0, 1),
			want: 7.222222,
			tol:  0.000001,
		},

		// --- All 5 basis types ---
		// issue=2008-01-15, first_interest=2008-07-15, settlement=2008-04-15
		// rate=0.08, par=1000, freq=2
		// basis 0: 30/360 US, 90 days / 180 = 0.5, 1000*0.08/2*0.5 = 20.0
		{
			name: "basis 0 US 30/360",
			args: numArgs(39462, 39644, 39553, 0.08, 1000, 2, 0),
			want: 20.0,
			tol:  0.0001,
		},
		// basis 1: actual/actual, 91 actual days / 182 period days = 20.0
		{
			name: "basis 1 actual/actual",
			args: numArgs(39462, 39644, 39553, 0.08, 1000, 2, 1),
			want: 20.0,
			tol:  0.0001,
		},
		// basis 2: actual/360, 91 actual days / 180 = 20.2222
		{
			name: "basis 2 actual/360",
			args: numArgs(39462, 39644, 39553, 0.08, 1000, 2, 2),
			want: 20.2222,
			tol:  0.001,
		},
		// basis 3: actual/365, 91 actual days / 182.5 = 19.9452
		{
			name: "basis 3 actual/365",
			args: numArgs(39462, 39644, 39553, 0.08, 1000, 2, 3),
			want: 19.9452,
			tol:  0.001,
		},
		// basis 4: European 30/360, 90 days / 180 = 20.0
		{
			name: "basis 4 European 30/360",
			args: numArgs(39462, 39644, 39553, 0.08, 1000, 2, 4),
			want: 20.0,
			tol:  0.0001,
		},

		// --- All 3 frequencies ---
		// issue=2008-01-15, settlement=2008-03-15, rate=0.12, par=1000, basis=0

		// freq=1 (annual), fi=2009-01-15
		// 30/360 from 1/15 to 3/15 = 60 days, NL=360. 1000*0.12/1*60/360 = 20.0
		{
			name: "freq 1 annual",
			args: numArgs(39462, 39828, 39522, 0.12, 1000, 1, 0),
			want: 20.0,
			tol:  0.0001,
		},
		// freq=2 (semiannual), fi=2008-07-15
		// 30/360 from 1/15 to 3/15 = 60 days, NL=180. 1000*0.12/2*60/180 = 20.0
		{
			name: "freq 2 semiannual",
			args: numArgs(39462, 39644, 39522, 0.12, 1000, 2, 0),
			want: 20.0,
			tol:  0.0001,
		},
		// freq=4 (quarterly), fi=2008-04-15
		// 30/360 from 1/15 to 3/15 = 60 days, NL=90. 1000*0.12/4*60/90 = 20.0
		{
			name: "freq 4 quarterly",
			args: numArgs(39462, 39553, 39522, 0.12, 1000, 4, 0),
			want: 20.0,
			tol:  0.0001,
		},

		// --- calc_method TRUE vs FALSE ---
		// issue=2007-01-01, fi=2007-07-01, settlement=2008-04-01
		// rate=0.1, par=1000, freq=2, basis=0

		// TRUE: accrual from issue (2007-01-01) spanning multiple periods
		// 3 full periods (150) - partial leftover gives 125.0
		{
			name: "calc_method TRUE multi-period",
			args: numArgs(39083, 39264, 39539, 0.1, 1000, 2, 0, 1),
			want: 125.0,
			tol:  0.0001,
		},
		// FALSE: accrual only from PCD before settlement
		// PCD=2008-01-01, 30/360 from 1/1 to 4/1 = 90 days. 50*90/180=25.0
		{
			name: "calc_method FALSE multi-period",
			args: numArgs(39083, 39264, 39539, 0.1, 1000, 2, 0, 0),
			want: 25.0,
			tol:  0.0001,
		},

		// --- Default calc_method (should be TRUE) ---
		{
			name: "default calc_method equals TRUE",
			args: numArgs(39083, 39264, 39539, 0.1, 1000, 2, 0),
			want: 125.0,
			tol:  0.0001,
		},

		// --- Settlement == first_interest ---
		// issue=2007-01-01, fi=2007-07-01, settlement=2007-07-01
		// Full half-year period: 1000*0.1/2*180/180 = 50.0
		{
			name: "settlement equals first_interest",
			args: numArgs(39083, 39264, 39264, 0.1, 1000, 2, 0),
			want: 50.0,
			tol:  0.0001,
		},

		// --- Settlement before first_interest ---
		// issue=2007-01-01, fi=2007-07-01, settlement=2007-04-01
		// 30/360 from 1/1 to 4/1 = 90 days, NL=180. 50*90/180=25.0
		{
			name: "settlement before first_interest",
			args: numArgs(39083, 39264, 39173, 0.1, 1000, 2, 0),
			want: 25.0,
			tol:  0.0001,
		},

		// --- Settlement after first_interest (multiple periods) ---
		// issue=2007-01-01, fi=2007-07-01, settlement=2008-06-15
		// 30/360: issue to fi (180 days, full period=50) + fi to 2008-01-01 (180 days=50)
		//   + 2008-01-01 to 2008-06-15: (6-1)*30+(15-1)=164 days. 50*164/180=45.5556
		// Total: 50+50+45.5556 = 145.5556
		{
			name: "settlement after first_interest multiple periods",
			args: numArgs(39083, 39264, 39614, 0.1, 1000, 2, 0),
			want: 145.5556,
			tol:  0.001,
		},

		// --- Very small rate ---
		// issue=2008-01-01, fi=2008-07-01, settlement=2008-04-01
		// 30/360 from 1/1 to 4/1 = 90 days. 1000*0.001/2*90/180=0.25
		{
			name: "very small rate",
			args: numArgs(39448, 39630, 39539, 0.001, 1000, 2, 0),
			want: 0.25,
			tol:  0.0001,
		},

		// --- Large par value ---
		// issue=2008-01-01, fi=2008-07-01, settlement=2008-04-01
		// 30/360 from 1/1 to 4/1 = 90 days. 1000000*0.05/2*90/180=12500.0
		{
			name: "large par value",
			args: numArgs(39448, 39630, 39539, 0.05, 1000000, 2, 0),
			want: 12500.0,
			tol:  0.01,
		},

		// --- Fractional dates get truncated ---
		{
			name: "fractional dates truncated",
			args: numArgs(39508.7, 39691.3, 39569.9, 0.1, 1000, 2, 0),
			want: 16.666667,
			tol:  0.000001,
		},

		// --- Error cases ---
		{
			name:    "rate zero",
			args:    numArgs(39462, 39644, 39553, 0, 1000, 2, 0),
			wantErr: true,
		},
		{
			name:    "rate negative",
			args:    numArgs(39462, 39644, 39553, -0.05, 1000, 2, 0),
			wantErr: true,
		},
		{
			name:    "par zero",
			args:    numArgs(39462, 39644, 39553, 0.05, 0, 2, 0),
			wantErr: true,
		},
		{
			name:    "par negative",
			args:    numArgs(39462, 39644, 39553, 0.05, -1000, 2, 0),
			wantErr: true,
		},
		{
			name:    "bad frequency 3",
			args:    numArgs(39462, 39644, 39553, 0.05, 1000, 3, 0),
			wantErr: true,
		},
		{
			name:    "bad frequency 0",
			args:    numArgs(39462, 39644, 39553, 0.05, 1000, 0, 0),
			wantErr: true,
		},
		{
			name:    "bad frequency 6",
			args:    numArgs(39462, 39644, 39553, 0.05, 1000, 6, 0),
			wantErr: true,
		},
		{
			name:    "basis negative",
			args:    numArgs(39462, 39644, 39553, 0.05, 1000, 2, -1),
			wantErr: true,
		},
		{
			name:    "basis 5",
			args:    numArgs(39462, 39644, 39553, 0.05, 1000, 2, 5),
			wantErr: true,
		},
		{
			name:    "issue equals settlement",
			args:    numArgs(39553, 39644, 39553, 0.05, 1000, 2, 0),
			wantErr: true,
		},
		{
			name:    "issue after settlement",
			args:    numArgs(39644, 39644, 39553, 0.05, 1000, 2, 0),
			wantErr: true,
		},
		{
			name:    "too few args 5",
			args:    numArgs(39462, 39644, 39553, 0.05, 1000),
			wantErr: true,
		},
		{
			name:    "too many args 9",
			args:    numArgs(39462, 39644, 39553, 0.05, 1000, 2, 0, 1, 99),
			wantErr: true,
		},
		{
			name:    "non-numeric issue",
			args:    []Value{StringVal("abc"), NumberVal(39644), NumberVal(39553), NumberVal(0.05), NumberVal(1000), NumberVal(2), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric settlement",
			args:    []Value{NumberVal(39462), NumberVal(39644), StringVal("xyz"), NumberVal(0.05), NumberVal(1000), NumberVal(2), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric rate",
			args:    []Value{NumberVal(39462), NumberVal(39644), NumberVal(39553), StringVal("abc"), NumberVal(1000), NumberVal(2), NumberVal(0)},
			wantErr: true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnAccrint(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				tol := tc.tol
				if tol == 0 {
					tol = 0.01
				}
				if v.Type != ValueNumber {
					t.Fatalf("%s: expected number, got type %v (str=%q)", tc.name, v.Type, v.Str)
				}
				if math.Abs(v.Num-tc.want) > tol {
					t.Errorf("%s: got %f, want %f (tol=%g)", tc.name, v.Num, tc.want, tol)
				}
			}
		})
	}
}

func TestACCRINT_ViaEval(t *testing.T) {
	// ACCRINT(39508,39691,39569,0.1,1000,2,0) = 16.666667
	cf := evalCompile(t, "ACCRINT(39508,39691,39569,0.1,1000,2,0)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-16.666667) > 0.000001 {
		t.Errorf("got %f, want 16.666667", v.Num)
	}
}

// === ACCRINTM ===

func TestACCRINTM_Comprehensive(t *testing.T) {
	// Serial numbers for reference dates:
	// DATE(2008,4,1) = 39539, DATE(2008,6,15) = 39614
	// DATE(2008,2,15) = 39493, DATE(2008,5,15) = 39583
	// DATE(2023,1,1) = 44927, DATE(2023,1,31) = 44957
	// DATE(2023,4,1) = 45017, DATE(2023,6,30) = 45107
	// DATE(2023,12,31) = 45291
	// DATE(2024,1,1) = 45292, DATE(2024,7,1) = 45474

	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
		tol     float64 // custom tolerance; 0 means use default 0.01
	}{
		// --- Doc example ---
		// ACCRINTM(DATE(2008,4,1), DATE(2008,6,15), 0.1, 1000, 3) = 20.54794521
		// A = 75 actual days, D = 365 => 1000 * 0.1 * 75 / 365
		{
			name: "doc example basis 3",
			args: numArgs(39539, 39614, 0.1, 1000, 3),
			want: 20.54794521,
			tol:  0.00001,
		},
		// --- All 5 basis types with same dates (2/15/2008 to 5/15/2008, rate=0.05, par=1000) ---
		// basis 0: US 30/360, A=90, D=360 => 1000 * 0.05 * 90/360 = 12.5
		{
			name: "basis 0 US 30/360",
			args: numArgs(39493, 39583, 0.05, 1000, 0),
			want: 12.5,
		},
		// basis 1: actual/actual, A=90, D=366 (2008 is leap) => 1000 * 0.05 * 90/366 = 12.29508
		{
			name: "basis 1 actual/actual",
			args: numArgs(39493, 39583, 0.05, 1000, 1),
			want: 12.29508,
			tol:  0.001,
		},
		// basis 2: actual/360, A=90, D=360 => 1000 * 0.05 * 90/360 = 12.5
		{
			name: "basis 2 actual/360",
			args: numArgs(39493, 39583, 0.05, 1000, 2),
			want: 12.5,
		},
		// basis 3: actual/365, A=90, D=365 => 1000 * 0.05 * 90/365 = 12.32877
		{
			name: "basis 3 actual/365",
			args: numArgs(39493, 39583, 0.05, 1000, 3),
			want: 12.32877,
			tol:  0.001,
		},
		// basis 4: European 30/360, A=90, D=360 => 1000 * 0.05 * 90/360 = 12.5
		{
			name: "basis 4 European 30/360",
			args: numArgs(39493, 39583, 0.05, 1000, 4),
			want: 12.5,
		},
		// --- Default basis (omitted = 0) ---
		{
			name: "default basis omitted 4 args",
			args: numArgs(39493, 39583, 0.05, 1000),
			want: 12.5,
		},
		// --- Various date ranges ---
		// Short range: 30 days, basis 0
		// 1/1/2023 to 1/31/2023, rate=0.08, par=1000
		// A=30 (US 30/360), D=360 => 1000 * 0.08 * 30/360 = 6.6667
		{
			name: "short 30 day range",
			args: numArgs(44927, 44957, 0.08, 1000, 0),
			want: 6.6667,
			tol:  0.001,
		},
		// 90 day range: 4/1/2023 to 6/30/2023, basis 2 (actual/360)
		// A=90 actual days, D=360 => 5000 * 0.06 * 90/360 = 75.0
		{
			name: "90 day range basis 2",
			args: numArgs(45017, 45107, 0.06, 5000, 2),
			want: 75.0,
		},
		// 180 day range: 1/1/2024 to 7/1/2024, basis 3 (actual/365)
		// A=182 actual days, D=365 => 2000 * 0.12 * 182/365 = 119.6712
		{
			name: "180 day range basis 3",
			args: numArgs(45292, 45474, 0.12, 2000, 3),
			want: 119.6712,
			tol:  0.01,
		},
		// Full year: 1/1/2023 to 12/31/2023, basis 1 (actual/actual)
		// A=364 actual days, D=365 (2023 is not leap) => 10000 * 0.05 * 364/365 = 498.6301
		{
			name: "full year basis 1",
			args: numArgs(44927, 45291, 0.05, 10000, 1),
			want: 498.6301,
			tol:  0.01,
		},
		// --- Large par value ---
		// par=1000000, rate=0.03, basis 0
		// 2/15/2008 to 5/15/2008: A=90, D=360 => 1000000 * 0.03 * 90/360 = 7500.0
		{
			name: "large par value",
			args: numArgs(39493, 39583, 0.03, 1000000, 0),
			want: 7500.0,
		},
		// --- Small par value ---
		// par=1, rate=0.1, basis 0
		// 2/15/2008 to 5/15/2008: A=90, D=360 => 1 * 0.1 * 90/360 = 0.025
		{
			name: "small par value",
			args: numArgs(39493, 39583, 0.1, 1, 0),
			want: 0.025,
			tol:  0.001,
		},
		// --- Small rate ---
		// rate=0.001, par=1000, basis 0
		// A=90, D=360 => 1000 * 0.001 * 90/360 = 0.25
		{
			name: "small rate",
			args: numArgs(39493, 39583, 0.001, 1000, 0),
			want: 0.25,
			tol:  0.001,
		},
		// --- Large rate ---
		// rate=2.0, par=1000, basis 0
		// A=90, D=360 => 1000 * 2.0 * 90/360 = 500.0
		{
			name: "large rate",
			args: numArgs(39493, 39583, 2.0, 1000, 0),
			want: 500.0,
		},
		// --- Fractional dates get truncated ---
		{
			name: "fractional dates truncated",
			args: numArgs(39539.7, 39614.9, 0.1, 1000, 3),
			want: 20.54794521,
			tol:  0.00001,
		},
		// --- Error cases ---
		{
			name:    "rate zero",
			args:    numArgs(39493, 39583, 0, 1000, 0),
			wantErr: true,
		},
		{
			name:    "rate negative",
			args:    numArgs(39493, 39583, -0.05, 1000, 0),
			wantErr: true,
		},
		{
			name:    "par zero",
			args:    numArgs(39493, 39583, 0.05, 0, 0),
			wantErr: true,
		},
		{
			name:    "par negative",
			args:    numArgs(39493, 39583, 0.05, -1000, 0),
			wantErr: true,
		},
		{
			name:    "issue equals settlement",
			args:    numArgs(39493, 39493, 0.05, 1000, 0),
			wantErr: true,
		},
		{
			name:    "issue after settlement",
			args:    numArgs(39583, 39493, 0.05, 1000, 0),
			wantErr: true,
		},
		{
			name:    "basis negative",
			args:    numArgs(39493, 39583, 0.05, 1000, -1),
			wantErr: true,
		},
		{
			name:    "basis 5",
			args:    numArgs(39493, 39583, 0.05, 1000, 5),
			wantErr: true,
		},
		{
			name:    "basis 99",
			args:    numArgs(39493, 39583, 0.05, 1000, 99),
			wantErr: true,
		},
		{
			name:    "too few args 3",
			args:    numArgs(39493, 39583, 0.05),
			wantErr: true,
		},
		{
			name:    "too many args 6",
			args:    numArgs(39493, 39583, 0.05, 1000, 0, 1),
			wantErr: true,
		},
		{
			name:    "non-numeric issue",
			args:    []Value{StringVal("abc"), NumberVal(39583), NumberVal(0.05), NumberVal(1000), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric settlement",
			args:    []Value{NumberVal(39493), StringVal("xyz"), NumberVal(0.05), NumberVal(1000), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric rate",
			args:    []Value{NumberVal(39493), NumberVal(39583), StringVal("abc"), NumberVal(1000), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric par",
			args:    []Value{NumberVal(39493), NumberVal(39583), NumberVal(0.05), StringVal("abc"), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric basis",
			args:    []Value{NumberVal(39493), NumberVal(39583), NumberVal(0.05), NumberVal(1000), StringVal("abc")},
			wantErr: true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnAccrintm(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				tol := tc.tol
				if tol == 0 {
					tol = 0.01
				}
				if v.Type != ValueNumber {
					t.Fatalf("%s: expected number, got type %v (str=%q)", tc.name, v.Type, v.Str)
				}
				if math.Abs(v.Num-tc.want) > tol {
					t.Errorf("%s: got %f, want %f (tol=%g)", tc.name, v.Num, tc.want, tol)
				}
			}
		})
	}
}

func TestACCRINTM_ViaEval(t *testing.T) {
	// ACCRINTM(DATE(2008,4,1), DATE(2008,6,15), 0.1, 1000, 3) = 20.54794521
	cf := evalCompile(t, "ACCRINTM(DATE(2008,4,1), DATE(2008,6,15), 0.1, 1000, 3)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-20.54794521) > 0.00001 {
		t.Errorf("got %f, want 20.54794521", v.Num)
	}
}

// === PRICEDISC ===

func TestPRICEDISC_Comprehensive(t *testing.T) {
	// Serial numbers for reference dates:
	// 2/16/2008 = 39494, 3/1/2008 = 39508
	// 2/15/2008 = 39493, 5/15/2008 = 39583
	// 1/1/2023 = 44927, 12/31/2023 = 45291
	// 1/1/2024 = 45292, 7/1/2024 = 45474
	// 1/31/2023 = 44957
	// 7/1/2018 = 43282, 1/1/2038 = 50406

	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
		tol     float64 // custom tolerance; 0 means use default 0.01
	}{
		// --- Doc example ---
		// PRICEDISC(2/16/2008, 3/1/2008, 0.0525, 100, 2) = 99.79583333
		// basis 2: DSM=14, B=360 => price = 100 - 0.0525*(14/360)*100 = 99.79583333
		{
			name: "doc example basis 2",
			args: numArgs(39494, 39508, 0.0525, 100, 2),
			want: 99.79583333,
			tol:  0.00001,
		},
		// --- All 5 basis types with same dates (2/15/2008 to 5/15/2008, discount=0.05, redemption=100) ---
		// basis 0: DSM=days360(US)=90, B=360 => price = 100 - 0.05*(90/360)*100 = 98.75
		{
			name: "basis 0 US 30/360",
			args: numArgs(39493, 39583, 0.05, 100, 0),
			want: 98.75,
			tol:  0.001,
		},
		// basis 1: DSM=90 actual, B=366 (2008 is leap) => price = 100 - 0.05*(90/366)*100 = 98.77049
		{
			name: "basis 1 actual/actual",
			args: numArgs(39493, 39583, 0.05, 100, 1),
			want: 98.77049,
			tol:  0.001,
		},
		// basis 2: DSM=90 actual, B=360 => price = 100 - 0.05*(90/360)*100 = 98.75
		{
			name: "basis 2 actual/360",
			args: numArgs(39493, 39583, 0.05, 100, 2),
			want: 98.75,
			tol:  0.001,
		},
		// basis 3: DSM=90 actual, B=365 => price = 100 - 0.05*(90/365)*100 = 98.76712
		{
			name: "basis 3 actual/365",
			args: numArgs(39493, 39583, 0.05, 100, 3),
			want: 98.76712,
			tol:  0.001,
		},
		// basis 4: European 30/360: DSM=days360(EU)=90, B=360 => price = 100 - 0.05*(90/360)*100 = 98.75
		{
			name: "basis 4 European 30/360",
			args: numArgs(39493, 39583, 0.05, 100, 4),
			want: 98.75,
			tol:  0.001,
		},
		// --- Default basis (omitted = 0) ---
		{
			name: "default basis omitted",
			args: numArgs(39493, 39583, 0.05, 100),
			want: 98.75,
			tol:  0.001,
		},
		// --- Various date ranges ---
		// Short range: 30 days, basis 0
		// 1/1/2023 to 1/31/2023, discount=0.10, redemption=100
		// DSM=30, B=360 => price = 100 - 0.10*(30/360)*100 = 99.16667
		{
			name: "short 30 day range",
			args: numArgs(44927, 44957, 0.10, 100, 0),
			want: 99.16667,
			tol:  0.001,
		},
		// Full year, basis 0
		// 1/1/2023 to 12/31/2023, discount=0.05, redemption=100
		// DSM=360 (US 30/360), B=360 => price = 100 - 0.05*(360/360)*100 = 95.0
		{
			name: "full year range",
			args: numArgs(44927, 45291, 0.05, 100, 0),
			want: 95.0,
			tol:  0.001,
		},
		// Half year, basis 2 (actual/360)
		// 1/1/2024 to 7/1/2024, discount=0.04, redemption=100
		// DSM=182 actual, B=360 => price = 100 - 0.04*(182/360)*100 = 97.97778
		{
			name: "half year basis 2",
			args: numArgs(45292, 45474, 0.04, 100, 2),
			want: 97.97778,
			tol:  0.001,
		},
		// Long range: 20 years, basis 1
		// 7/1/2018 to 1/1/2038, discount=0.03, redemption=100
		{
			name: "long range 20 years basis 1",
			args: numArgs(43282, 50406, 0.03, 100, 1),
			want: 41.4847,
			tol:  0.01,
		},
		// --- Non-100 redemption ---
		// Redemption=1000, discount=0.05, basis 0
		// DSM=90, B=360 => price = 1000 - 0.05*(90/360)*1000 = 987.5
		{
			name: "redemption 1000",
			args: numArgs(39493, 39583, 0.05, 1000, 0),
			want: 987.5,
			tol:  0.001,
		},
		// --- Small discount ---
		{
			name: "very small discount",
			args: numArgs(39493, 39583, 0.001, 100, 0),
			want: 99.975,
			tol:  0.001,
		},
		// --- Large discount ---
		// discount=0.50, redemption=100, basis 0
		// price = 100 - 0.50*(90/360)*100 = 87.5
		{
			name: "large discount",
			args: numArgs(39493, 39583, 0.50, 100, 0),
			want: 87.5,
			tol:  0.001,
		},
		// Fractional dates get truncated
		{
			name: "fractional dates truncated",
			args: numArgs(39494.7, 39508.9, 0.0525, 100, 2),
			want: 99.79583333,
			tol:  0.00001,
		},
		// --- Error cases ---
		{
			name:    "discount zero",
			args:    numArgs(39493, 39583, 0, 100, 0),
			wantErr: true,
		},
		{
			name:    "discount negative",
			args:    numArgs(39493, 39583, -0.05, 100, 0),
			wantErr: true,
		},
		{
			name:    "redemption zero",
			args:    numArgs(39493, 39583, 0.05, 0, 0),
			wantErr: true,
		},
		{
			name:    "redemption negative",
			args:    numArgs(39493, 39583, 0.05, -100, 0),
			wantErr: true,
		},
		{
			name:    "settlement equals maturity",
			args:    numArgs(39493, 39493, 0.05, 100, 0),
			wantErr: true,
		},
		{
			name:    "settlement after maturity",
			args:    numArgs(39583, 39493, 0.05, 100, 0),
			wantErr: true,
		},
		{
			name:    "basis negative",
			args:    numArgs(39493, 39583, 0.05, 100, -1),
			wantErr: true,
		},
		{
			name:    "basis 5",
			args:    numArgs(39493, 39583, 0.05, 100, 5),
			wantErr: true,
		},
		{
			name:    "basis 99",
			args:    numArgs(39493, 39583, 0.05, 100, 99),
			wantErr: true,
		},
		{
			name:    "too few args",
			args:    numArgs(39493, 39583, 0.05),
			wantErr: true,
		},
		{
			name:    "too many args",
			args:    numArgs(39493, 39583, 0.05, 100, 0, 1),
			wantErr: true,
		},
		{
			name:    "non-numeric settlement",
			args:    []Value{StringVal("abc"), NumberVal(39583), NumberVal(0.05), NumberVal(100), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric maturity",
			args:    []Value{NumberVal(39493), StringVal("xyz"), NumberVal(0.05), NumberVal(100), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric discount",
			args:    []Value{NumberVal(39493), NumberVal(39583), StringVal("abc"), NumberVal(100), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric redemption",
			args:    []Value{NumberVal(39493), NumberVal(39583), NumberVal(0.05), StringVal("abc"), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric basis",
			args:    []Value{NumberVal(39493), NumberVal(39583), NumberVal(0.05), NumberVal(100), StringVal("abc")},
			wantErr: true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnPricedisc(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				tol := tc.tol
				if tol == 0 {
					tol = 0.01
				}
				if v.Type != ValueNumber {
					t.Fatalf("%s: expected number, got type %v (str=%q)", tc.name, v.Type, v.Str)
				}
				if math.Abs(v.Num-tc.want) > tol {
					t.Errorf("%s: got %f, want %f (tol=%g)", tc.name, v.Num, tc.want, tol)
				}
			}
		})
	}
}

func TestPRICEDISC_ViaEval(t *testing.T) {
	// PRICEDISC(DATE(2008,2,16), DATE(2008,3,1), 0.0525, 100, 2) = 99.79583333
	cf := evalCompile(t, "PRICEDISC(DATE(2008,2,16), DATE(2008,3,1), 0.0525, 100, 2)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-99.79583333) > 0.00001 {
		t.Errorf("got %f, want 99.79583333", v.Num)
	}
}

// === YIELDDISC ===

func TestYIELDDISC_Comprehensive(t *testing.T) {
	// Serial numbers for reference dates:
	// 2/16/2008 = 39494, 3/1/2008 = 39508
	// 2/15/2008 = 39493, 5/15/2008 = 39583
	// 1/1/2023 = 44927, 12/31/2023 = 45291
	// 1/1/2024 = 45292, 7/1/2024 = 45474
	// 1/31/2023 = 44957

	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
		tol     float64 // custom tolerance; 0 means use default 0.01
	}{
		// --- Doc example ---
		// YIELDDISC(2/16/2008, 3/1/2008, 99.795, 100, 2) = 0.052823
		// basis 2: DSM=14, B=360 => yield = (100-99.795)/99.795 * (360/14) = 0.052823
		{
			name: "doc example basis 2",
			args: numArgs(39494, 39508, 99.795, 100, 2),
			want: 0.052823,
			tol:  0.0001,
		},
		// --- All 5 basis types with same dates (2/15/2008 to 5/15/2008, pr=98, redemption=100) ---
		// basis 0: DSM=days360(US)=90, B=360 => yield = (2/98)*(360/90) = 0.08163
		{
			name: "basis 0 US 30/360",
			args: numArgs(39493, 39583, 98, 100, 0),
			want: 0.08163,
			tol:  0.001,
		},
		// basis 1: DSM=90 actual, B=366 (2008 is leap) => yield = (2/98)*(366/90) = 0.08299
		{
			name: "basis 1 actual/actual",
			args: numArgs(39493, 39583, 98, 100, 1),
			want: 0.08299,
			tol:  0.001,
		},
		// basis 2: DSM=90 actual, B=360 => yield = (2/98)*(360/90) = 0.08163
		{
			name: "basis 2 actual/360",
			args: numArgs(39493, 39583, 98, 100, 2),
			want: 0.08163,
			tol:  0.001,
		},
		// basis 3: DSM=90 actual, B=365 => yield = (2/98)*(365/90) = 0.08277
		{
			name: "basis 3 actual/365",
			args: numArgs(39493, 39583, 98, 100, 3),
			want: 0.08277,
			tol:  0.001,
		},
		// basis 4: European 30/360: DSM=days360(EU)=90, B=360 => yield = (2/98)*(360/90) = 0.08163
		{
			name: "basis 4 European 30/360",
			args: numArgs(39493, 39583, 98, 100, 4),
			want: 0.08163,
			tol:  0.001,
		},
		// --- Default basis (omitted = 0) ---
		{
			name: "default basis omitted",
			args: numArgs(39493, 39583, 98, 100),
			want: 0.08163,
			tol:  0.001,
		},
		// --- Various date ranges ---
		// Short range: 30 days, basis 0
		// 1/1/2023 to 1/31/2023, pr=99, redemption=100
		// DSM=30, B=360 => yield = (1/99)*(360/30) = 0.12121
		{
			name: "short 30 day range",
			args: numArgs(44927, 44957, 99, 100, 0),
			want: 0.12121,
			tol:  0.001,
		},
		// Full year, basis 0
		// 1/1/2023 to 12/31/2023, pr=95, redemption=100
		// DSM=360 (US 30/360), B=360 => yield = (5/95)*(360/360) = 0.05263
		{
			name: "full year range",
			args: numArgs(44927, 45291, 95, 100, 0),
			want: 0.05263,
			tol:  0.001,
		},
		// Half year, basis 2 (actual/360)
		// 1/1/2024 to 7/1/2024, pr=98, redemption=100
		// DSM=182 actual, B=360 => yield = (2/98)*(360/182) = 0.04039
		{
			name: "half year basis 2",
			args: numArgs(45292, 45474, 98, 100, 2),
			want: 0.04039,
			tol:  0.001,
		},
		// --- Larger spread ---
		// pr=90, redemption=100
		// DSM=90, B=360 => yield = (10/90)*(360/90) = 0.44444
		{
			name: "large spread",
			args: numArgs(39493, 39583, 90, 100, 0),
			want: 0.44444,
			tol:  0.001,
		},
		// --- Very small spread ---
		// pr=99.999, redemption=100
		// DSM=90, B=360 => yield = (0.001/99.999)*(360/90) = 0.00004
		{
			name: "very small spread",
			args: numArgs(39493, 39583, 99.999, 100, 0),
			want: 0.00004,
			tol:  0.00001,
		},
		// pr > redemption (negative yield)
		// pr=105, redemption=100
		// DSM=90, B=360 => yield = (100-105)/105 * (360/90) = -0.19048
		{
			name: "pr exceeds redemption negative yield",
			args: numArgs(39493, 39583, 105, 100, 0),
			want: -0.19048,
			tol:  0.001,
		},
		// Non-100 redemption
		// pr=980, redemption=1000, DSM=90, B=360
		// yield = (20/980)*(360/90) = 0.08163
		{
			name: "non-100 redemption",
			args: numArgs(39493, 39583, 980, 1000, 0),
			want: 0.08163,
			tol:  0.001,
		},
		// Fractional dates get truncated
		{
			name: "fractional dates truncated",
			args: numArgs(39494.7, 39508.9, 99.795, 100, 2),
			want: 0.052823,
			tol:  0.0001,
		},
		// Long range: basis 3
		// 1/1/2023 to 12/31/2023, pr=95, redemption=100
		// DSM=364 actual, B=365 => yield = (5/95)*(365/364) = 0.05277
		{
			name: "full year basis 3",
			args: numArgs(44927, 45291, 95, 100, 3),
			want: 0.05277,
			tol:  0.001,
		},
		// --- Error cases ---
		{
			name:    "pr zero",
			args:    numArgs(39493, 39583, 0, 100, 0),
			wantErr: true,
		},
		{
			name:    "pr negative",
			args:    numArgs(39493, 39583, -10, 100, 0),
			wantErr: true,
		},
		{
			name:    "redemption zero",
			args:    numArgs(39493, 39583, 98, 0, 0),
			wantErr: true,
		},
		{
			name:    "redemption negative",
			args:    numArgs(39493, 39583, 98, -100, 0),
			wantErr: true,
		},
		{
			name:    "settlement equals maturity",
			args:    numArgs(39493, 39493, 98, 100, 0),
			wantErr: true,
		},
		{
			name:    "settlement after maturity",
			args:    numArgs(39583, 39493, 98, 100, 0),
			wantErr: true,
		},
		{
			name:    "basis negative",
			args:    numArgs(39493, 39583, 98, 100, -1),
			wantErr: true,
		},
		{
			name:    "basis 5",
			args:    numArgs(39493, 39583, 98, 100, 5),
			wantErr: true,
		},
		{
			name:    "basis 99",
			args:    numArgs(39493, 39583, 98, 100, 99),
			wantErr: true,
		},
		{
			name:    "too few args",
			args:    numArgs(39493, 39583, 98),
			wantErr: true,
		},
		{
			name:    "too many args",
			args:    numArgs(39493, 39583, 98, 100, 0, 1),
			wantErr: true,
		},
		{
			name:    "non-numeric settlement",
			args:    []Value{StringVal("abc"), NumberVal(39583), NumberVal(98), NumberVal(100), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric maturity",
			args:    []Value{NumberVal(39493), StringVal("xyz"), NumberVal(98), NumberVal(100), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric pr",
			args:    []Value{NumberVal(39493), NumberVal(39583), StringVal("abc"), NumberVal(100), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric redemption",
			args:    []Value{NumberVal(39493), NumberVal(39583), NumberVal(98), StringVal("abc"), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric basis",
			args:    []Value{NumberVal(39493), NumberVal(39583), NumberVal(98), NumberVal(100), StringVal("abc")},
			wantErr: true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnYielddisc(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				tol := tc.tol
				if tol == 0 {
					tol = 0.01
				}
				if v.Type != ValueNumber {
					t.Fatalf("%s: expected number, got type %v (str=%q)", tc.name, v.Type, v.Str)
				}
				if math.Abs(v.Num-tc.want) > tol {
					t.Errorf("%s: got %f, want %f (tol=%g)", tc.name, v.Num, tc.want, tol)
				}
			}
		})
	}
}

func TestYIELDDISC_ViaEval(t *testing.T) {
	// YIELDDISC(DATE(2008,2,16), DATE(2008,3,1), 99.795, 100, 2) = 0.052823
	cf := evalCompile(t, "YIELDDISC(DATE(2008,2,16), DATE(2008,3,1), 99.795, 100, 2)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-0.052823) > 0.0001 {
		t.Errorf("got %f, want 0.052823", v.Num)
	}
}

// === PRICEMAT ===

func TestPRICEMAT_Comprehensive(t *testing.T) {
	// Serial numbers for reference dates:
	// 2/15/2008 = 39493, 4/13/2008 = 39551, 11/11/2007 = 39397
	// 3/15/2008 = 39522, 11/3/2008 = 39755, 11/8/2007 = 39394
	// 1/1/2023 = 44927, 6/30/2023 = 45107, 12/31/2023 = 45291
	// 1/1/2024 = 45292, 7/1/2024 = 45474
	// 3/15/2023 = 45000, 6/1/2022 = 44713

	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
		tol     float64 // custom tolerance; 0 means use default 0.01
	}{
		// --- Doc example ---
		// PRICEMAT(2/15/2008, 4/13/2008, 11/11/2007, 0.061, 0.061, 0) = 99.98449...
		{
			name: "doc example basis 0",
			args: numArgs(39493, 39551, 39397, 0.061, 0.061, 0),
			want: 99.98449888,
			tol:  0.0001,
		},
		// --- All 5 basis types with doc example dates ---
		{
			name: "basis 1 actual/actual",
			args: numArgs(39493, 39551, 39397, 0.061, 0.061, 1),
			want: 99.98468141,
			tol:  0.0001,
		},
		{
			name: "basis 2 actual/360",
			args: numArgs(39493, 39551, 39397, 0.061, 0.061, 2),
			want: 99.98416906,
			tol:  0.0001,
		},
		{
			name: "basis 3 actual/365",
			args: numArgs(39493, 39551, 39397, 0.061, 0.061, 3),
			want: 99.98459776,
			tol:  0.0001,
		},
		{
			name: "basis 4 European 30/360",
			args: numArgs(39493, 39551, 39397, 0.061, 0.061, 4),
			want: 99.98449888,
			tol:  0.0001,
		},
		// --- Default basis (omitted = 0) ---
		{
			name: "default basis omitted",
			args: numArgs(39493, 39551, 39397, 0.061, 0.061),
			want: 99.98449888,
			tol:  0.0001,
		},
		// --- rate=0 edge case ---
		// With rate=0: price = 100/(1+DSM/B*yld)
		{
			name: "rate zero",
			args: numArgs(39493, 39551, 39397, 0, 0.05, 0),
			want: 99.20088179,
			tol:  0.0001,
		},
		// --- yld=0 edge case ---
		// With yld=0: denom=1, price = 100*(1+DIM/B*rate) - 100*A/B*rate
		{
			name: "yld zero",
			args: numArgs(39493, 39551, 39397, 0.05, 0, 0),
			want: 100.80555556,
			tol:  0.0001,
		},
		// --- rate=0, yld=0 ---
		{
			name: "rate zero yld zero",
			args: numArgs(39493, 39551, 39397, 0, 0, 0),
			want: 100.0,
			tol:  0.0001,
		},
		// --- Long range: settle=1/1/2023, mat=12/31/2023, issue=6/1/2022 ---
		{
			name: "long range basis 0",
			args: numArgs(44927, 45291, 44713, 0.05, 0.06, 0),
			want: 98.89150943,
			tol:  0.0001,
		},
		{
			name: "long range basis 1",
			args: numArgs(44927, 45291, 44713, 0.05, 0.06, 1),
			want: 98.89353710,
			tol:  0.0001,
		},
		{
			name: "long range basis 2",
			args: numArgs(44927, 45291, 44713, 0.05, 0.06, 2),
			want: 98.87671974,
			tol:  0.0001,
		},
		{
			name: "long range basis 3",
			args: numArgs(44927, 45291, 44713, 0.05, 0.06, 3),
			want: 98.89353710,
			tol:  0.0001,
		},
		{
			name: "long range basis 4",
			args: numArgs(44927, 45291, 44713, 0.05, 0.06, 4),
			want: 98.89441474,
			tol:  0.0001,
		},
		// --- Short range: settle=3/15/2023, mat=6/30/2023, issue=1/1/2023 ---
		{
			name: "short range basis 0",
			args: numArgs(45000, 45107, 44927, 0.04, 0.05, 0),
			want: 99.70070728,
			tol:  0.0001,
		},
		{
			name: "short range basis 2",
			args: numArgs(45000, 45107, 44927, 0.04, 0.05, 2),
			want: 99.69525265,
			tol:  0.0001,
		},
		// --- High rate/yield ---
		{
			name: "high rate and yield",
			args: numArgs(39493, 39551, 39397, 0.20, 0.15, 0),
			want: 100.66332158,
			tol:  0.0001,
		},
		// --- Fractional dates get truncated ---
		{
			name: "fractional dates truncated",
			args: numArgs(39493.7, 39551.9, 39397.3, 0.061, 0.061, 0),
			want: 99.98449888,
			tol:  0.0001,
		},
		// --- Error cases ---
		{
			name:    "rate negative",
			args:    numArgs(39493, 39551, 39397, -0.01, 0.061, 0),
			wantErr: true,
		},
		{
			name:    "yld negative",
			args:    numArgs(39493, 39551, 39397, 0.061, -0.01, 0),
			wantErr: true,
		},
		{
			name:    "settlement equals maturity",
			args:    numArgs(39493, 39493, 39397, 0.061, 0.061, 0),
			wantErr: true,
		},
		{
			name:    "settlement after maturity",
			args:    numArgs(39551, 39493, 39397, 0.061, 0.061, 0),
			wantErr: true,
		},
		{
			name:    "basis negative",
			args:    numArgs(39493, 39551, 39397, 0.061, 0.061, -1),
			wantErr: true,
		},
		{
			name:    "basis 5",
			args:    numArgs(39493, 39551, 39397, 0.061, 0.061, 5),
			wantErr: true,
		},
		{
			name:    "basis 99",
			args:    numArgs(39493, 39551, 39397, 0.061, 0.061, 99),
			wantErr: true,
		},
		{
			name:    "too few args",
			args:    numArgs(39493, 39551, 39397, 0.061),
			wantErr: true,
		},
		{
			name:    "too many args",
			args:    numArgs(39493, 39551, 39397, 0.061, 0.061, 0, 1),
			wantErr: true,
		},
		{
			name:    "non-numeric settlement",
			args:    []Value{StringVal("abc"), NumberVal(39551), NumberVal(39397), NumberVal(0.061), NumberVal(0.061), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric maturity",
			args:    []Value{NumberVal(39493), StringVal("xyz"), NumberVal(39397), NumberVal(0.061), NumberVal(0.061), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric issue",
			args:    []Value{NumberVal(39493), NumberVal(39551), StringVal("abc"), NumberVal(0.061), NumberVal(0.061), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric rate",
			args:    []Value{NumberVal(39493), NumberVal(39551), NumberVal(39397), StringVal("abc"), NumberVal(0.061), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric yld",
			args:    []Value{NumberVal(39493), NumberVal(39551), NumberVal(39397), NumberVal(0.061), StringVal("abc"), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric basis",
			args:    []Value{NumberVal(39493), NumberVal(39551), NumberVal(39397), NumberVal(0.061), NumberVal(0.061), StringVal("abc")},
			wantErr: true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnPricemat(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				tol := tc.tol
				if tol == 0 {
					tol = 0.01
				}
				if v.Type != ValueNumber {
					t.Fatalf("%s: expected number, got type %v (str=%q)", tc.name, v.Type, v.Str)
				}
				if math.Abs(v.Num-tc.want) > tol {
					t.Errorf("%s: got %f, want %f (tol=%g)", tc.name, v.Num, tc.want, tol)
				}
			}
		})
	}
}

func TestPRICEMAT_ViaEval(t *testing.T) {
	// PRICEMAT(DATE(2008,2,15), DATE(2008,4,13), DATE(2007,11,11), 0.061, 0.061, 0) = 99.98449...
	cf := evalCompile(t, "PRICEMAT(DATE(2008,2,15), DATE(2008,4,13), DATE(2007,11,11), 0.061, 0.061, 0)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-99.98449888) > 0.001 {
		t.Errorf("got %f, want 99.98449888", v.Num)
	}
}

// === YIELDMAT ===

func TestYIELDMAT_Comprehensive(t *testing.T) {
	// Serial numbers for reference dates:
	// 3/15/2008 = 39522, 11/3/2008 = 39755, 11/8/2007 = 39394
	// 2/15/2008 = 39493, 4/13/2008 = 39551, 11/11/2007 = 39397
	// 1/1/2023 = 44927, 6/30/2023 = 45107, 12/31/2023 = 45291
	// 3/15/2023 = 45000, 6/1/2022 = 44713

	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
		tol     float64 // custom tolerance; 0 means use default 0.01
	}{
		// --- Doc example ---
		// YIELDMAT(3/15/2008, 11/3/2008, 11/8/2007, 0.0625, 100.0123, 0) = 0.060954
		{
			name: "doc example basis 0",
			args: numArgs(39522, 39755, 39394, 0.0625, 100.0123, 0),
			want: 0.06095433,
			tol:  0.0001,
		},
		// --- All 5 basis types ---
		{
			name: "basis 1 actual/actual",
			args: numArgs(39522, 39755, 39394, 0.0625, 100.0123, 1),
			want: 0.06096363,
			tol:  0.0001,
		},
		{
			name: "basis 2 actual/360",
			args: numArgs(39522, 39755, 39394, 0.0625, 100.0123, 2),
			want: 0.06094806,
			tol:  0.0001,
		},
		{
			name: "basis 3 actual/365",
			args: numArgs(39522, 39755, 39394, 0.0625, 100.0123, 3),
			want: 0.06096363,
			tol:  0.0001,
		},
		{
			name: "basis 4 European 30/360",
			args: numArgs(39522, 39755, 39394, 0.0625, 100.0123, 4),
			want: 0.06095433,
			tol:  0.0001,
		},
		// --- Default basis (omitted = 0) ---
		{
			name: "default basis omitted",
			args: numArgs(39522, 39755, 39394, 0.0625, 100.0123),
			want: 0.06095433,
			tol:  0.0001,
		},
		// --- rate=0 edge case ---
		{
			name: "rate zero",
			args: numArgs(39522, 39755, 39394, 0, 100.0123, 0),
			want: -0.00019419,
			tol:  0.0001,
		},
		// --- pr small (well below par) ---
		{
			name: "pr equals 50",
			args: numArgs(39522, 39755, 39394, 0.0625, 50, 0),
			want: 1.63198152,
			tol:  0.0001,
		},
		// --- pr large (above par) ---
		{
			name: "pr equals 150",
			args: numArgs(39522, 39755, 39394, 0.0625, 150, 0),
			want: -0.47762843,
			tol:  0.0001,
		},
		// --- Long range: settle=1/1/2023, mat=12/31/2023, issue=6/1/2022 ---
		{
			name: "long range basis 0",
			args: numArgs(44927, 45291, 44713, 0.05, 100, 0),
			want: 0.04858300,
			tol:  0.0001,
		},
		{
			name: "long range basis 1",
			args: numArgs(44927, 45291, 44713, 0.05, 100, 1),
			want: 0.04857599,
			tol:  0.0001,
		},
		{
			name: "long range basis 2",
			args: numArgs(44927, 45291, 44713, 0.05, 100, 2),
			want: 0.04855678,
			tol:  0.0001,
		},
		{
			name: "long range basis 3",
			args: numArgs(44927, 45291, 44713, 0.05, 100, 3),
			want: 0.04857599,
			tol:  0.0001,
		},
		{
			name: "long range basis 4",
			args: numArgs(44927, 45291, 44713, 0.05, 100, 4),
			want: 0.04858300,
			tol:  0.0001,
		},
		// --- Short range: settle=3/15/2023, mat=6/30/2023, issue=1/1/2023 ---
		{
			name: "short range basis 0",
			args: numArgs(45000, 45107, 44927, 0.04, 99.95, 0),
			want: 0.04139463,
			tol:  0.0001,
		},
		{
			name: "short range basis 3",
			args: numArgs(45000, 45107, 44927, 0.04, 99.95, 3),
			want: 0.04139514,
			tol:  0.0001,
		},
		// --- High rate ---
		{
			name: "high rate pr 105",
			args: numArgs(39522, 39755, 39394, 0.20, 105, 0),
			want: 0.10802912,
			tol:  0.0001,
		},
		// --- Fractional dates get truncated ---
		{
			name: "fractional dates truncated",
			args: numArgs(39522.7, 39755.9, 39394.3, 0.0625, 100.0123, 0),
			want: 0.06095433,
			tol:  0.0001,
		},
		// --- pr exactly at par ---
		{
			name: "pr exactly 100 rate 0.05",
			args: numArgs(44927, 45291, 44713, 0.05, 100, 0),
			want: 0.04858300,
			tol:  0.0001,
		},
		// --- Error cases ---
		{
			name:    "rate negative",
			args:    numArgs(39522, 39755, 39394, -0.01, 100.0123, 0),
			wantErr: true,
		},
		{
			name:    "pr zero",
			args:    numArgs(39522, 39755, 39394, 0.0625, 0, 0),
			wantErr: true,
		},
		{
			name:    "pr negative",
			args:    numArgs(39522, 39755, 39394, 0.0625, -10, 0),
			wantErr: true,
		},
		{
			name:    "settlement equals maturity",
			args:    numArgs(39522, 39522, 39394, 0.0625, 100.0123, 0),
			wantErr: true,
		},
		{
			name:    "settlement after maturity",
			args:    numArgs(39755, 39522, 39394, 0.0625, 100.0123, 0),
			wantErr: true,
		},
		{
			name:    "basis negative",
			args:    numArgs(39522, 39755, 39394, 0.0625, 100.0123, -1),
			wantErr: true,
		},
		{
			name:    "basis 5",
			args:    numArgs(39522, 39755, 39394, 0.0625, 100.0123, 5),
			wantErr: true,
		},
		{
			name:    "basis 99",
			args:    numArgs(39522, 39755, 39394, 0.0625, 100.0123, 99),
			wantErr: true,
		},
		{
			name:    "too few args",
			args:    numArgs(39522, 39755, 39394, 0.0625),
			wantErr: true,
		},
		{
			name:    "too many args",
			args:    numArgs(39522, 39755, 39394, 0.0625, 100.0123, 0, 1),
			wantErr: true,
		},
		{
			name:    "non-numeric settlement",
			args:    []Value{StringVal("abc"), NumberVal(39755), NumberVal(39394), NumberVal(0.0625), NumberVal(100.0123), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric maturity",
			args:    []Value{NumberVal(39522), StringVal("xyz"), NumberVal(39394), NumberVal(0.0625), NumberVal(100.0123), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric issue",
			args:    []Value{NumberVal(39522), NumberVal(39755), StringVal("abc"), NumberVal(0.0625), NumberVal(100.0123), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric rate",
			args:    []Value{NumberVal(39522), NumberVal(39755), NumberVal(39394), StringVal("abc"), NumberVal(100.0123), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric pr",
			args:    []Value{NumberVal(39522), NumberVal(39755), NumberVal(39394), NumberVal(0.0625), StringVal("abc"), NumberVal(0)},
			wantErr: true,
		},
		{
			name:    "non-numeric basis",
			args:    []Value{NumberVal(39522), NumberVal(39755), NumberVal(39394), NumberVal(0.0625), NumberVal(100.0123), StringVal("abc")},
			wantErr: true,
		},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnYieldmat(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				tol := tc.tol
				if tol == 0 {
					tol = 0.01
				}
				if v.Type != ValueNumber {
					t.Fatalf("%s: expected number, got type %v (str=%q)", tc.name, v.Type, v.Str)
				}
				if math.Abs(v.Num-tc.want) > tol {
					t.Errorf("%s: got %f, want %f (tol=%g)", tc.name, v.Num, tc.want, tol)
				}
			}
		})
	}
}

func TestYIELDMAT_ViaEval(t *testing.T) {
	// YIELDMAT(DATE(2008,3,15), DATE(2008,11,3), DATE(2007,11,8), 0.0625, 100.0123, 0) = 0.060954
	cf := evalCompile(t, "YIELDMAT(DATE(2008,3,15), DATE(2008,11,3), DATE(2007,11,8), 0.0625, 100.0123, 0)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-0.060954) > 0.001 {
		t.Errorf("got %f, want 0.060954", v.Num)
	}
}

// ---------------------------------------------------------------------------
// COUPPCD tests
// ---------------------------------------------------------------------------

func TestCOUPPCD(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// Doc example: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=2, basis=1 => Nov 15 2010 (40497)
		{name: "doc example", args: numArgs(40568, 40862, 2, 1), want: 40497},
		// Semiannual: settlement=Jan 25 2007, maturity=Nov 15 2008, freq=2, basis=0
		// Coupon schedule: ...May 15 2006, Nov 15 2006, May 15 2007, Nov 15 2007, May 15 2008, Nov 15 2008
		// PCD for Jan 25 2007 = Nov 15 2006
		{name: "semiannual 2007", args: numArgs(39107, 39767, 2, 0), want: 39036},
		// Annual: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=1, basis=0
		// Coupon schedule: ...Nov 15 2009, Nov 15 2010, Nov 15 2011
		// PCD = Nov 15 2010
		{name: "annual", args: numArgs(40568, 40862, 1, 0), want: 40497},
		// Quarterly: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=4, basis=0
		// Coupon schedule: ...Aug 15 2010, Nov 15 2010, Feb 15 2011, May 15 2011, Aug 15 2011, Nov 15 2011
		// PCD for Jan 25 2011 = Nov 15 2010
		{name: "quarterly", args: numArgs(40568, 40862, 4, 0), want: 40497},
		// Settlement on a coupon date: settlement=Nov 15 2010, maturity=Nov 15 2011, freq=2
		// PCD = Nov 15 2010 (the settlement date itself)
		{name: "settlement on coupon date", args: numArgs(40497, 40862, 2, 0), want: 40497},
		// Settlement one day after coupon date: settlement=Nov 16 2010, maturity=Nov 15 2011, freq=2
		// PCD = Nov 15 2010
		{name: "one day after coupon", args: numArgs(40498, 40862, 2, 0), want: 40497},
		// Settlement one day before maturity: settlement=Nov 14 2011, maturity=Nov 15 2011, freq=2
		// PCD = May 15 2011
		{name: "one day before maturity", args: numArgs(40861, 40862, 2, 0), want: 40678},
		// Basis=2
		{name: "basis 2", args: numArgs(40568, 40862, 2, 2), want: 40497},
		// Basis=3
		{name: "basis 3", args: numArgs(40568, 40862, 2, 3), want: 40497},
		// Basis=4
		{name: "basis 4", args: numArgs(40568, 40862, 2, 4), want: 40497},
		// Long-term bond: settlement=Jan 25 2007, maturity=Nov 15 2011, freq=2
		{name: "long-term bond", args: numArgs(39107, 40862, 2, 0), want: 39036},
		// End-of-month maturity: maturity=Aug 31 2012, freq=2
		// Coupon schedule includes Feb 29 2012 (leap year), Aug 31 2011, Feb 28 2011
		// Settlement=Jan 25 2012: PCD = Aug 31 2011
		{name: "EOM maturity Aug 31", args: numArgs(40933, 41152, 2, 0), want: 40786},
		// End-of-month maturity: maturity=Feb 29 2012, freq=2 (leap year)
		// Coupon schedule: ...Aug 29 2011, Feb 29 2012
		// Settlement=Jan 25 2012: PCD = Aug 29 2011
		{name: "EOM maturity Feb 29 leap", args: numArgs(40933, 40968, 2, 0), want: 40784},
		// --- Error cases ---
		{name: "too few args", args: numArgs(40568, 40862), wantErr: true},
		{name: "too many args", args: numArgs(40568, 40862, 2, 1, 99), wantErr: true},
		{name: "settlement >= maturity", args: numArgs(40862, 40568, 2, 0), wantErr: true},
		{name: "settlement == maturity", args: numArgs(40862, 40862, 2, 0), wantErr: true},
		{name: "bad frequency 3", args: numArgs(40568, 40862, 3, 0), wantErr: true},
		{name: "bad frequency 0", args: numArgs(40568, 40862, 0, 0), wantErr: true},
		{name: "bad basis -1", args: numArgs(40568, 40862, 2, -1), wantErr: true},
		{name: "bad basis 5", args: numArgs(40568, 40862, 2, 5), wantErr: true},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnCouppcd(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestCOUPPCD_ViaEval(t *testing.T) {
	cf := evalCompile(t, "COUPPCD(DATE(2011,1,25), DATE(2011,11,15), 2, 1)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	// Nov 15 2010 = serial 40497
	if math.Abs(v.Num-40497) > 1 {
		t.Errorf("got %f, want 40497", v.Num)
	}
}

// ---------------------------------------------------------------------------
// COUPNCD tests
// ---------------------------------------------------------------------------

func TestCOUPNCD(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// Doc example: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=2, basis=1 => May 15 2011 (40678)
		{name: "doc example", args: numArgs(40568, 40862, 2, 1), want: 40678},
		// Semiannual: settlement=Jan 25 2007, maturity=Nov 15 2008, freq=2
		// NCD = May 15 2007 (39217)
		{name: "semiannual 2007", args: numArgs(39107, 39767, 2, 0), want: 39217},
		// Annual: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=1
		// NCD = Nov 15 2011 (maturity itself)
		{name: "annual", args: numArgs(40568, 40862, 1, 0), want: 40862},
		// Quarterly: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=4
		// NCD = Feb 15 2011 (40589)
		{name: "quarterly", args: numArgs(40568, 40862, 4, 0), want: 40589},
		// Settlement on a coupon date: settlement=Nov 15 2010, maturity=Nov 15 2011, freq=2
		// NCD = May 15 2011 (40678)
		{name: "settlement on coupon date", args: numArgs(40497, 40862, 2, 0), want: 40678},
		// Settlement one day after coupon date: settlement=Nov 16 2010, maturity=Nov 15 2011, freq=2
		// NCD = May 15 2011 (40678)
		{name: "one day after coupon", args: numArgs(40498, 40862, 2, 0), want: 40678},
		// Settlement one day before maturity: settlement=Nov 14 2011, maturity=Nov 15 2011, freq=2
		// NCD = Nov 15 2011 (40862)
		{name: "one day before maturity", args: numArgs(40861, 40862, 2, 0), want: 40862},
		// All basis values give same coupon dates
		{name: "basis 0", args: numArgs(40568, 40862, 2, 0), want: 40678},
		{name: "basis 2", args: numArgs(40568, 40862, 2, 2), want: 40678},
		{name: "basis 3", args: numArgs(40568, 40862, 2, 3), want: 40678},
		{name: "basis 4", args: numArgs(40568, 40862, 2, 4), want: 40678},
		// Long-term bond: settlement=Jan 25 2007, maturity=Nov 15 2011, freq=2
		// NCD = May 15 2007
		{name: "long-term bond", args: numArgs(39107, 40862, 2, 0), want: 39217},
		// End-of-month maturity: maturity=Aug 31 2012, freq=2
		// Settlement=Jan 25 2012: PCD=Aug 31 2011, NCD=Feb 29 2012 (leap year)
		{name: "EOM maturity Aug 31 NCD", args: numArgs(40933, 41152, 2, 0), want: 40968},
		// End-of-month maturity: maturity=Feb 29 2012, freq=2
		// Settlement=Jan 25 2012: NCD = Feb 29 2012
		{name: "EOM maturity Feb 29 NCD", args: numArgs(40933, 40968, 2, 0), want: 40968},
		// --- Error cases ---
		{name: "too few args", args: numArgs(40568, 40862), wantErr: true},
		{name: "too many args", args: numArgs(40568, 40862, 2, 1, 99), wantErr: true},
		{name: "settlement >= maturity", args: numArgs(40862, 40568, 2, 0), wantErr: true},
		{name: "bad frequency 5", args: numArgs(40568, 40862, 5, 0), wantErr: true},
		{name: "bad basis 6", args: numArgs(40568, 40862, 2, 6), wantErr: true},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnCoupncd(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestCOUPNCD_ViaEval(t *testing.T) {
	cf := evalCompile(t, "COUPNCD(DATE(2011,1,25), DATE(2011,11,15), 2, 1)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	// May 15 2011 = serial 40678
	if math.Abs(v.Num-40678) > 1 {
		t.Errorf("got %f, want 40678", v.Num)
	}
}

// ---------------------------------------------------------------------------
// COUPNUM tests
// ---------------------------------------------------------------------------

func TestCOUPNUM(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// Doc example: settlement=Jan 25 2007, maturity=Nov 15 2008, freq=2, basis=1 => 4
		{name: "doc example", args: numArgs(39107, 39767, 2, 1), want: 4},
		// Semiannual: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=2 => 2
		// Coupons after settlement: May 15 2011, Nov 15 2011
		{name: "semiannual 2 periods", args: numArgs(40568, 40862, 2, 0), want: 2},
		// Annual: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=1 => 1
		{name: "annual 1 period", args: numArgs(40568, 40862, 1, 0), want: 1},
		// Quarterly: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=4 => 4
		// Coupons after settlement: Feb 15, May 15, Aug 15, Nov 15
		{name: "quarterly 4 periods", args: numArgs(40568, 40862, 4, 0), want: 4},
		// Settlement on coupon date: settlement=May 15 2011, maturity=Nov 15 2011, freq=2 => 1
		// Only Nov 15 2011 is strictly after settlement
		{name: "settlement on coupon", args: numArgs(40678, 40862, 2, 0), want: 1},
		// Settlement one day after coupon: settlement=May 16 2011, maturity=Nov 15 2011, freq=2 => 1
		{name: "one day after coupon", args: numArgs(40679, 40862, 2, 0), want: 1},
		// Settlement one day before maturity: settlement=Nov 14 2011, maturity=Nov 15 2011, freq=2 => 1
		{name: "one day before maturity", args: numArgs(40861, 40862, 2, 0), want: 1},
		// Long-term: settlement=Jan 25 2007, maturity=Nov 15 2011, freq=2 => 10
		{name: "long-term 10 periods", args: numArgs(39107, 40862, 2, 0), want: 10},
		// Long-term annual: settlement=Jan 25 2007, maturity=Nov 15 2011, freq=1 => 5
		{name: "long-term annual 5 periods", args: numArgs(39107, 40862, 1, 0), want: 5},
		// Long-term quarterly: settlement=Jan 25 2007, maturity=Nov 15 2008, freq=4 => 8
		{name: "quarterly 8 periods", args: numArgs(39107, 39767, 4, 0), want: 8},
		// Different basis values should not affect count
		{name: "basis 1", args: numArgs(39107, 39767, 2, 1), want: 4},
		{name: "basis 2", args: numArgs(39107, 39767, 2, 2), want: 4},
		{name: "basis 3", args: numArgs(39107, 39767, 2, 3), want: 4},
		{name: "basis 4", args: numArgs(39107, 39767, 2, 4), want: 4},
		// --- Error cases ---
		{name: "too few args", args: numArgs(39107, 39767), wantErr: true},
		{name: "too many args", args: numArgs(39107, 39767, 2, 1, 1), wantErr: true},
		{name: "settlement >= maturity", args: numArgs(39767, 39107, 2, 0), wantErr: true},
		{name: "bad frequency", args: numArgs(39107, 39767, 6, 0), wantErr: true},
		{name: "bad basis", args: numArgs(39107, 39767, 2, 7), wantErr: true},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnCoupnum(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestCOUPNUM_ViaEval(t *testing.T) {
	cf := evalCompile(t, "COUPNUM(DATE(2007,1,25), DATE(2008,11,15), 2, 1)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-4) > 0.01 {
		t.Errorf("got %f, want 4", v.Num)
	}
}

// ---------------------------------------------------------------------------
// COUPDAYBS tests
// ---------------------------------------------------------------------------

func TestCOUPDAYBS(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// Doc example: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=2, basis=1 => 71
		// PCD = Nov 15 2010. Actual days Nov 15 2010 to Jan 25 2011 = 71
		{name: "doc example basis 1", args: numArgs(40568, 40862, 2, 1), want: 71},
		// Basis 0 (30/360): Nov 15 2010 to Jan 25 2011
		// 30/360: (2011-2010)*360 + (1-11)*30 + (25-15) = 360 - 300 + 10 = 70
		{name: "basis 0", args: numArgs(40568, 40862, 2, 0), want: 70},
		// Basis 2 (actual/360): same actual days as basis 1 = 71
		{name: "basis 2", args: numArgs(40568, 40862, 2, 2), want: 71},
		// Basis 3 (actual/365): same actual days as basis 1 = 71
		{name: "basis 3", args: numArgs(40568, 40862, 2, 3), want: 71},
		// Basis 4 (European 30/360): same as US for these dates = 70
		{name: "basis 4", args: numArgs(40568, 40862, 2, 4), want: 70},
		// Annual: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=1, basis=1
		// PCD = Nov 15 2010. Same = 71 actual days
		{name: "annual basis 1", args: numArgs(40568, 40862, 1, 1), want: 71},
		// Quarterly: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=4, basis=1
		// PCD = Nov 15 2010. Same = 71 actual days
		{name: "quarterly basis 1", args: numArgs(40568, 40862, 4, 1), want: 71},
		// Settlement on coupon date: COUPDAYBS = 0
		{name: "settlement on coupon", args: numArgs(40497, 40862, 2, 1), want: 0},
		// Settlement one day after coupon: COUPDAYBS = 1
		{name: "one day after coupon", args: numArgs(40498, 40862, 2, 1), want: 1},
		// Settlement one day before maturity: PCD = May 15 2011 (40678)
		// Actual days from May 15 to Nov 14 = 183
		{name: "one day before maturity basis 1", args: numArgs(40861, 40862, 2, 1), want: 183},
		// Long-term: settlement=Jan 25 2007, maturity=Nov 15 2008, freq=2, basis=1
		// PCD = Nov 15 2006 (serial 39036). Actual days Nov 15 2006 to Jan 25 2007 = 71
		{name: "long-term basis 1", args: numArgs(39107, 39767, 2, 1), want: 71},
		// Long-term basis 0: 30/360 = 70
		{name: "long-term basis 0", args: numArgs(39107, 39767, 2, 0), want: 70},
		// --- Error cases ---
		{name: "too few args", args: numArgs(40568, 40862), wantErr: true},
		{name: "too many args", args: numArgs(40568, 40862, 2, 1, 1), wantErr: true},
		{name: "settlement >= maturity", args: numArgs(40862, 40568, 2, 0), wantErr: true},
		{name: "bad frequency", args: numArgs(40568, 40862, 3, 0), wantErr: true},
		{name: "bad basis", args: numArgs(40568, 40862, 2, -1), wantErr: true},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnCoupdaybs(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestCOUPDAYBS_ViaEval(t *testing.T) {
	cf := evalCompile(t, "COUPDAYBS(DATE(2011,1,25), DATE(2011,11,15), 2, 1)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-71) > 0.01 {
		t.Errorf("got %f, want 71", v.Num)
	}
}

// ---------------------------------------------------------------------------
// COUPDAYS tests
// ---------------------------------------------------------------------------

func TestCOUPDAYS(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// Doc example: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=2, basis=1 => 181
		// PCD = Nov 15 2010, NCD = May 15 2011. Actual days = 181
		{name: "doc example basis 1", args: numArgs(40568, 40862, 2, 1), want: 181},
		// Basis 0 (30/360): 360/2 = 180
		{name: "basis 0 semiannual", args: numArgs(40568, 40862, 2, 0), want: 180},
		// Basis 2 (actual/360): 360/2 = 180
		{name: "basis 2 semiannual", args: numArgs(40568, 40862, 2, 2), want: 180},
		// Basis 3 (actual/365): 365/2 = 182.5
		{name: "basis 3 semiannual", args: numArgs(40568, 40862, 2, 3), want: 182.5},
		// Basis 4 (European 30/360): 360/2 = 180
		{name: "basis 4 semiannual", args: numArgs(40568, 40862, 2, 4), want: 180},
		// Annual basis 0: 360/1 = 360
		{name: "annual basis 0", args: numArgs(40568, 40862, 1, 0), want: 360},
		// Quarterly basis 0: 360/4 = 90
		{name: "quarterly basis 0", args: numArgs(40568, 40862, 4, 0), want: 90},
		// Annual basis 3: 365/1 = 365
		{name: "annual basis 3", args: numArgs(40568, 40862, 1, 3), want: 365},
		// Quarterly basis 3: 365/4 = 91.25
		{name: "quarterly basis 3", args: numArgs(40568, 40862, 4, 3), want: 91.25},
		// Basis 1 (actual/actual) with different coupon period:
		// settlement=May 16 2011, maturity=Nov 15 2011, freq=2
		// PCD = May 15 2011, NCD = Nov 15 2011. Actual days = 184
		{name: "basis 1 May-Nov period", args: numArgs(40679, 40862, 2, 1), want: 184},
		// Settlement on coupon: PCD=Nov 15 2010, NCD=May 15 2011 = 181
		{name: "settlement on coupon basis 1", args: numArgs(40497, 40862, 2, 1), want: 181},
		// Long-term bond: PCD=Nov 15 2006, NCD=May 15 2007 = 181 actual days
		{name: "long-term basis 1", args: numArgs(39107, 39767, 2, 1), want: 181},
		// Long-term basis 0: always 180 for semiannual
		{name: "long-term basis 0", args: numArgs(39107, 39767, 2, 0), want: 180},
		// --- Error cases ---
		{name: "too few args", args: numArgs(40568, 40862), wantErr: true},
		{name: "too many args", args: numArgs(40568, 40862, 2, 1, 1), wantErr: true},
		{name: "settlement >= maturity", args: numArgs(40862, 40568, 2, 0), wantErr: true},
		{name: "bad frequency", args: numArgs(40568, 40862, 7, 0), wantErr: true},
		{name: "bad basis", args: numArgs(40568, 40862, 2, 5), wantErr: true},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnCoupdays(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestCOUPDAYS_ViaEval(t *testing.T) {
	cf := evalCompile(t, "COUPDAYS(DATE(2011,1,25), DATE(2011,11,15), 2, 1)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-181) > 0.01 {
		t.Errorf("got %f, want 181", v.Num)
	}
}

// ---------------------------------------------------------------------------
// COUPDAYSNC tests
// ---------------------------------------------------------------------------

func TestCOUPDAYSNC(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// Doc example: settlement=Jan 25 2011, maturity=Nov 15 2011, freq=2, basis=1 => 110
		// NCD = May 15 2011. Actual days Jan 25 to May 15 = 110
		{name: "doc example basis 1", args: numArgs(40568, 40862, 2, 1), want: 110},
		// Basis 0 (30/360): COUPDAYS(180) - COUPDAYBS(70) = 110
		{name: "basis 0", args: numArgs(40568, 40862, 2, 0), want: 110},
		// Basis 2 (actual/360): actual days = 110
		{name: "basis 2", args: numArgs(40568, 40862, 2, 2), want: 110},
		// Basis 3 (actual/365): actual days = 110
		{name: "basis 3", args: numArgs(40568, 40862, 2, 3), want: 110},
		// Basis 4 (European 30/360): COUPDAYS(180) - COUPDAYBS(70) = 110
		{name: "basis 4", args: numArgs(40568, 40862, 2, 4), want: 110},
		// Annual: NCD = Nov 15 2011. Actual days Jan 25 to Nov 15 = 294
		{name: "annual basis 1", args: numArgs(40568, 40862, 1, 1), want: 294},
		// Quarterly: NCD = Feb 15 2011. Actual days Jan 25 to Feb 15 = 21
		{name: "quarterly basis 1", args: numArgs(40568, 40862, 4, 1), want: 21},
		// Settlement on coupon: days to next coupon
		// Settlement=Nov 15 2010, NCD=May 15 2011 = 181 actual days
		{name: "settlement on coupon basis 1", args: numArgs(40497, 40862, 2, 1), want: 181},
		// Settlement one day after coupon: NCD=May 15 2011, days from Nov 16 = 180
		{name: "one day after coupon basis 1", args: numArgs(40498, 40862, 2, 1), want: 180},
		// Settlement one day before maturity: NCD=Nov 15 2011, days = 1
		{name: "one day before maturity", args: numArgs(40861, 40862, 2, 1), want: 1},
		// Long-term: settlement=Jan 25 2007, NCD=May 15 2007
		// Actual days Jan 25 to May 15 = 110
		{name: "long-term basis 1", args: numArgs(39107, 39767, 2, 1), want: 110},
		// Long-term basis 0: 180 - 70 = 110
		{name: "long-term basis 0", args: numArgs(39107, 39767, 2, 0), want: 110},
		// Settlement on coupon basis 0: 180 - 0 = 180
		{name: "settlement on coupon basis 0", args: numArgs(40497, 40862, 2, 0), want: 180},
		// --- Error cases ---
		{name: "too few args", args: numArgs(40568, 40862), wantErr: true},
		{name: "too many args", args: numArgs(40568, 40862, 2, 1, 1), wantErr: true},
		{name: "settlement >= maturity", args: numArgs(40862, 40568, 2, 0), wantErr: true},
		{name: "bad frequency", args: numArgs(40568, 40862, 8, 0), wantErr: true},
		{name: "bad basis", args: numArgs(40568, 40862, 2, 9), wantErr: true},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnCoupdaysnc(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestCOUPDAYSNC_ViaEval(t *testing.T) {
	cf := evalCompile(t, "COUPDAYSNC(DATE(2011,1,25), DATE(2011,11,15), 2, 1)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-110) > 0.01 {
		t.Errorf("got %f, want 110", v.Num)
	}
}

// ---------------------------------------------------------------------------
// DURATION tests
// ---------------------------------------------------------------------------

func TestDURATION(t *testing.T) {
	// Serial date constants used in tests:
	// Jul 1 2018  = 43282    Jan 1 2048  = 54058
	// Jan 1 2008  = 39448    Jan 1 2016  = 42370
	// Jan 1 2020  = 43831    Jan 1 2025  = 45658
	// Jan 1 2030  = 47484    Jun 15 2020 = 43997
	// Dec 15 2020 = 44180    Jun 15 2021 = 44362
	// Jan 15 2020 = 43845    Jul 15 2020 = 44027
	// Jan 15 2021 = 44211    Jan 15 2050 = 54803
	// Jul 15 2025 = 45853    Jan 15 2025 = 45672
	// Jan 15 2030 = 47498

	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// --- Doc example ---
		// settlement=7/1/2018, maturity=1/1/2048, coupon=0.08, yld=0.09, freq=2, basis=1
		{name: "doc example", args: numArgs(43282, 54058, 0.08, 0.09, 2, 1), want: 10.9191},

		// --- All 3 frequencies ---
		// Annual (freq=1): settlement=1/1/2020, maturity=1/1/2030, coupon=0.05, yld=0.06, basis=0
		{name: "annual freq=1", args: numArgs(43831, 47484, 0.05, 0.06, 1, 0), want: 8.0225},
		// Semiannual (freq=2): same dates
		{name: "semiannual freq=2", args: numArgs(43831, 47484, 0.05, 0.06, 2, 0), want: 7.8950},
		// Quarterly (freq=4): same dates
		{name: "quarterly freq=4", args: numArgs(43831, 47484, 0.05, 0.06, 4, 0), want: 7.8304},

		// --- All 5 basis types ---
		// basis=0 (US 30/360)
		{name: "basis 0", args: numArgs(43282, 54058, 0.08, 0.09, 2, 0), want: 10.9137},
		// basis=1 (actual/actual) - doc example
		{name: "basis 1", args: numArgs(43282, 54058, 0.08, 0.09, 2, 1), want: 10.9191},
		// basis=2 (actual/360)
		{name: "basis 2", args: numArgs(43282, 54058, 0.08, 0.09, 2, 2), want: 10.9191},
		// basis=3 (actual/365)
		{name: "basis 3", args: numArgs(43282, 54058, 0.08, 0.09, 2, 3), want: 10.9191},
		// basis=4 (European 30/360)
		{name: "basis 4", args: numArgs(43282, 54058, 0.08, 0.09, 2, 4), want: 10.9137},

		// --- Zero coupon bond ---
		// coupon=0: duration equals time to maturity
		{name: "zero coupon", args: numArgs(43831, 47484, 0, 0.05, 2, 0), want: 10.0000},

		// --- Zero coupon and zero yield ---
		// coupon=0, yld=0: duration = maturity time in years
		{name: "zero coupon zero yield", args: numArgs(43831, 47484, 0, 0, 2, 0), want: 10.0},

		// --- Short-term bond (1 period) ---
		// settlement=Jun 15 2020, maturity=Dec 15 2020, freq=2
		{name: "one period", args: numArgs(43997, 44180, 0.06, 0.05, 2, 0), want: 0.5},

		// --- Long-term bond (30 years) ---
		{name: "long-term 30yr", args: numArgs(43845, 54803, 0.06, 0.07, 2, 1), want: 13.2630},

		// --- High yield ---
		{name: "high yield 20%", args: numArgs(43831, 47484, 0.05, 0.20, 2, 0), want: 6.3224},

		// --- Low yield ---
		{name: "low yield 0.5%", args: numArgs(43831, 47484, 0.05, 0.005, 2, 0), want: 8.3774},

		// --- High coupon ---
		{name: "high coupon 15%", args: numArgs(43831, 47484, 0.15, 0.06, 2, 0), want: 6.4988},

		// --- Default basis (5 args, no basis) ---
		{name: "default basis", args: numArgs(43831, 47484, 0.05, 0.06, 2), want: 7.8950},

		// --- Short maturity, quarterly ---
		// settlement=Jan 15 2025, maturity=Jul 15 2025, freq=4
		{name: "short maturity quarterly", args: numArgs(45672, 45853, 0.04, 0.03, 4, 0), want: 0.4951},

		// --- Multi-period, annual, basis=3 ---
		{name: "annual basis=3", args: numArgs(43831, 47484, 0.07, 0.08, 1, 3), want: 7.4205},

		// --- Error cases ---
		{name: "too few args", args: numArgs(43831, 47484, 0.05, 0.06), wantErr: true},
		{name: "too many args", args: numArgs(43831, 47484, 0.05, 0.06, 2, 0, 99), wantErr: true},
		{name: "negative coupon", args: numArgs(43831, 47484, -0.05, 0.06, 2, 0), wantErr: true},
		{name: "negative yield", args: numArgs(43831, 47484, 0.05, -0.06, 2, 0), wantErr: true},
		{name: "settlement >= maturity", args: numArgs(47484, 43831, 0.05, 0.06, 2, 0), wantErr: true},
		{name: "settlement == maturity", args: numArgs(43831, 43831, 0.05, 0.06, 2, 0), wantErr: true},
		{name: "bad frequency 3", args: numArgs(43831, 47484, 0.05, 0.06, 3, 0), wantErr: true},
		{name: "bad frequency 0", args: numArgs(43831, 47484, 0.05, 0.06, 0, 0), wantErr: true},
		{name: "bad basis -1", args: numArgs(43831, 47484, 0.05, 0.06, 2, -1), wantErr: true},
		{name: "bad basis 5", args: numArgs(43831, 47484, 0.05, 0.06, 2, 5), wantErr: true},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnDuration(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestDURATION_ViaEval(t *testing.T) {
	cf := evalCompile(t, "DURATION(DATE(2018,7,1), DATE(2048,1,1), 0.08, 0.09, 2, 1)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-10.9191) > 0.01 {
		t.Errorf("got %f, want ~10.9191", v.Num)
	}
}

// ---------------------------------------------------------------------------
// MDURATION tests
// ---------------------------------------------------------------------------

func TestMDURATION(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// --- Doc example ---
		// settlement=1/1/2008, maturity=1/1/2016, coupon=0.08, yld=0.09, freq=2, basis=1
		{name: "doc example", args: numArgs(39448, 42370, 0.08, 0.09, 2, 1), want: 5.7357},

		// --- All 3 frequencies ---
		// Annual (freq=1): settlement=1/1/2020, maturity=1/1/2030, coupon=0.05, yld=0.06
		{name: "annual freq=1", args: numArgs(43831, 47484, 0.05, 0.06, 1, 0), want: 7.5684},
		// Semiannual (freq=2): same dates
		{name: "semiannual freq=2", args: numArgs(43831, 47484, 0.05, 0.06, 2, 0), want: 7.6650},
		// Quarterly (freq=4): same dates
		{name: "quarterly freq=4", args: numArgs(43831, 47484, 0.05, 0.06, 4, 0), want: 7.7146},

		// --- All 5 basis types ---
		// settlement=7/1/2018, maturity=1/1/2048, coupon=0.08, yld=0.09, freq=2
		{name: "basis 0", args: numArgs(43282, 54058, 0.08, 0.09, 2, 0), want: 10.4438},
		{name: "basis 1", args: numArgs(43282, 54058, 0.08, 0.09, 2, 1), want: 10.4490},
		{name: "basis 2", args: numArgs(43282, 54058, 0.08, 0.09, 2, 2), want: 10.4490},
		{name: "basis 3", args: numArgs(43282, 54058, 0.08, 0.09, 2, 3), want: 10.4490},
		{name: "basis 4", args: numArgs(43282, 54058, 0.08, 0.09, 2, 4), want: 10.4438},

		// --- Zero coupon bond ---
		{name: "zero coupon", args: numArgs(43831, 47484, 0, 0.05, 2, 0), want: 9.7561},

		// --- Zero coupon and zero yield ---
		{name: "zero coupon zero yield", args: numArgs(43831, 47484, 0, 0, 2, 0), want: 10.0},

		// --- Short-term bond (1 period) ---
		{name: "one period", args: numArgs(43997, 44180, 0.06, 0.05, 2, 0), want: 0.4878},

		// --- Long-term bond (30 years) ---
		{name: "long-term 30yr", args: numArgs(43845, 54803, 0.06, 0.07, 2, 1), want: 12.8145},

		// --- High yield ---
		{name: "high yield 20%", args: numArgs(43831, 47484, 0.05, 0.20, 2, 0), want: 5.7476},

		// --- Low yield ---
		{name: "low yield 0.5%", args: numArgs(43831, 47484, 0.05, 0.005, 2, 0), want: 8.3565},

		// --- High coupon ---
		{name: "high coupon 15%", args: numArgs(43831, 47484, 0.15, 0.06, 2, 0), want: 6.3095},

		// --- Default basis (5 args, no basis) ---
		{name: "default basis", args: numArgs(43831, 47484, 0.05, 0.06, 2), want: 7.6650},

		// --- Short maturity, quarterly ---
		{name: "short maturity quarterly", args: numArgs(45672, 45853, 0.04, 0.03, 4, 0), want: 0.4914},

		// --- Multi-period, annual, basis=3 ---
		{name: "annual basis=3", args: numArgs(43831, 47484, 0.07, 0.08, 1, 3), want: 6.8708},

		// --- MDURATION = DURATION / (1 + yld/freq) verified ---
		// Use the doc example values: DURATION=10.9191, yld=0.09, freq=2
		// MDURATION = 10.9191 / (1 + 0.045) = 10.4490
		{name: "formula verification", args: numArgs(43282, 54058, 0.08, 0.09, 2, 1), want: 10.4490},

		// --- Error cases ---
		{name: "too few args", args: numArgs(43831, 47484, 0.05, 0.06), wantErr: true},
		{name: "too many args", args: numArgs(43831, 47484, 0.05, 0.06, 2, 0, 99), wantErr: true},
		{name: "negative coupon", args: numArgs(43831, 47484, -0.05, 0.06, 2, 0), wantErr: true},
		{name: "negative yield", args: numArgs(43831, 47484, 0.05, -0.06, 2, 0), wantErr: true},
		{name: "settlement >= maturity", args: numArgs(47484, 43831, 0.05, 0.06, 2, 0), wantErr: true},
		{name: "settlement == maturity", args: numArgs(43831, 43831, 0.05, 0.06, 2, 0), wantErr: true},
		{name: "bad frequency 3", args: numArgs(43831, 47484, 0.05, 0.06, 3, 0), wantErr: true},
		{name: "bad frequency 0", args: numArgs(43831, 47484, 0.05, 0.06, 0, 0), wantErr: true},
		{name: "bad basis -1", args: numArgs(43831, 47484, 0.05, 0.06, 2, -1), wantErr: true},
		{name: "bad basis 5", args: numArgs(43831, 47484, 0.05, 0.06, 2, 5), wantErr: true},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnMduration(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestMDURATION_ViaEval(t *testing.T) {
	cf := evalCompile(t, "MDURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 2, 1)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-5.7357) > 0.01 {
		t.Errorf("got %f, want ~5.7357", v.Num)
	}
}

// ---------------------------------------------------------------------------
// PRICE tests
// ---------------------------------------------------------------------------

func TestPRICE(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// --- Doc example ---
		// settlement=2/15/2008 (39494), maturity=11/15/2017 (43055)
		// rate=0.0575, yld=0.065, redemption=100, freq=2, basis=0
		{name: "doc example", args: numArgs(39494, 43055, 0.0575, 0.065, 100, 2, 0), want: 94.6344},

		// --- All 3 frequencies ---
		// settlement=1/1/2020 (43832), maturity=1/1/2030 (47485)
		{name: "annual freq=1", args: numArgs(43832, 47485, 0.05, 0.06, 100, 1, 0), want: 92.6399},
		{name: "semiannual freq=2", args: numArgs(43832, 47485, 0.05, 0.06, 100, 2, 0), want: 92.5613},
		{name: "quarterly freq=4", args: numArgs(43832, 47485, 0.05, 0.06, 100, 4, 0), want: 92.5210},

		// --- Multiple basis values ---
		{name: "basis 0 (US 30/360)", args: numArgs(39494, 43055, 0.0575, 0.065, 100, 2, 0), want: 94.6344},
		{name: "basis 1 (actual/actual)", args: numArgs(39494, 43055, 0.0575, 0.065, 100, 2, 1), want: 94.6354},
		{name: "basis 2 (actual/360)", args: numArgs(39494, 43055, 0.0575, 0.065, 100, 2, 2), want: 94.6024},
		{name: "basis 3 (actual/365)", args: numArgs(39494, 43055, 0.0575, 0.065, 100, 2, 3), want: 94.6436},
		{name: "basis 4 (European 30/360)", args: numArgs(39494, 43055, 0.0575, 0.065, 100, 2, 4), want: 94.6344},

		// --- Single coupon period (N=1) ---
		// settlement=6/1/2020 (43984), maturity=12/1/2020 (44167), freq=2
		{name: "single period", args: numArgs(43984, 44167, 0.06, 0.05, 100, 2, 0), want: 100.4878},

		// --- Short quarterly single period ---
		// settlement=3/1/2025 (45718), maturity=6/1/2025 (45810), freq=4
		{name: "short quarterly N=1", args: numArgs(45718, 45810, 0.04, 0.03, 100, 4, 0), want: 100.2481},

		// --- Zero coupon ---
		{name: "zero coupon", args: numArgs(43832, 47485, 0, 0.06, 100, 2, 0), want: 55.3676},

		// --- Zero coupon zero yield ---
		{name: "zero rate zero yield", args: numArgs(43832, 47485, 0, 0, 100, 2, 0), want: 100.00},

		// --- Premium bond (high coupon, price > 100) ---
		{name: "premium bond", args: numArgs(39494, 43055, 0.10, 0.065, 100, 2, 0), want: 124.9660},

		// --- Discount bond (low coupon, price < 100) ---
		{name: "discount bond", args: numArgs(39494, 43055, 0.03, 0.065, 100, 2, 0), want: 75.0080},

		// --- Par bond (rate == yield => price ~ 100) ---
		{name: "par bond", args: numArgs(43832, 47485, 0.06, 0.06, 100, 2, 0), want: 100.0000},

		// --- High yield ---
		{name: "high yield 20%", args: numArgs(39494, 43055, 0.0575, 0.20, 100, 2, 0), want: 39.8235},

		// --- Low yield ---
		{name: "low yield 0.5%", args: numArgs(39494, 43055, 0.0575, 0.005, 100, 2, 0), want: 149.8981},

		// --- Zero yield ---
		{name: "zero yield", args: numArgs(39494, 43055, 0.0575, 0, 100, 2, 0), want: 156.0625},

		// --- Non-100 redemption ---
		{name: "redemption 110", args: numArgs(39494, 43055, 0.0575, 0.065, 110, 2, 0), want: 99.9941},

		// --- Long bond 20 years annual ---
		// settlement=1/1/2025 (45659), maturity=1/1/2045 (52972), rate=0.07, yld=0.08
		{name: "long bond 20yr annual", args: numArgs(45659, 52972, 0.07, 0.08, 100, 1, 1), want: 90.1715},

		// --- Default basis (6 args, no basis) ---
		{name: "default basis", args: numArgs(39494, 43055, 0.0575, 0.065, 100, 2), want: 94.6344},

		// --- Error cases ---
		{name: "too few args", args: numArgs(39494, 43055, 0.0575, 0.065, 100), wantErr: true},
		{name: "too many args", args: numArgs(39494, 43055, 0.0575, 0.065, 100, 2, 0, 99), wantErr: true},
		{name: "negative rate", args: numArgs(39494, 43055, -0.05, 0.065, 100, 2, 0), wantErr: true},
		{name: "negative yield", args: numArgs(39494, 43055, 0.0575, -0.01, 100, 2, 0), wantErr: true},
		{name: "zero redemption", args: numArgs(39494, 43055, 0.0575, 0.065, 0, 2, 0), wantErr: true},
		{name: "negative redemption", args: numArgs(39494, 43055, 0.0575, 0.065, -100, 2, 0), wantErr: true},
		{name: "settlement >= maturity", args: numArgs(43055, 39494, 0.0575, 0.065, 100, 2, 0), wantErr: true},
		{name: "settlement == maturity", args: numArgs(39494, 39494, 0.0575, 0.065, 100, 2, 0), wantErr: true},
		{name: "bad frequency 3", args: numArgs(39494, 43055, 0.0575, 0.065, 100, 3, 0), wantErr: true},
		{name: "bad frequency 0", args: numArgs(39494, 43055, 0.0575, 0.065, 100, 0, 0), wantErr: true},
		{name: "bad basis -1", args: numArgs(39494, 43055, 0.0575, 0.065, 100, 2, -1), wantErr: true},
		{name: "bad basis 5", args: numArgs(39494, 43055, 0.0575, 0.065, 100, 2, 5), wantErr: true},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnPrice(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestPRICE_ViaEval(t *testing.T) {
	cf := evalCompile(t, "PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, 100, 2, 0)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-94.6344) > 0.01 {
		t.Errorf("got %f, want ~94.6344", v.Num)
	}
}

// ---------------------------------------------------------------------------
// YIELD tests
// ---------------------------------------------------------------------------

func TestYIELD(t *testing.T) {
	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// --- Doc example ---
		// settlement=2/15/2008 (39494), maturity=11/15/2016 (42690)
		// rate=0.0575, pr=95.04287, redemption=100, freq=2, basis=0
		{name: "doc example", args: numArgs(39494, 42690, 0.0575, 95.04287, 100, 2, 0), want: 0.0650},

		// --- All 3 frequencies ---
		// Prices computed from PRICE(settlement, maturity, 0.05, 0.06, 100, freq, 0)
		{name: "annual freq=1", args: numArgs(43832, 47485, 0.05, 92.6399, 100, 1, 0), want: 0.0600},
		{name: "semiannual freq=2", args: numArgs(43832, 47485, 0.05, 92.5613, 100, 2, 0), want: 0.0600},
		{name: "quarterly freq=4", args: numArgs(43832, 47485, 0.05, 92.5210, 100, 4, 0), want: 0.0600},

		// --- Multiple basis values ---
		{name: "basis 0", args: numArgs(39494, 43055, 0.0575, 94.6344, 100, 2, 0), want: 0.0650},
		{name: "basis 1", args: numArgs(39494, 43055, 0.0575, 94.6354, 100, 2, 1), want: 0.0650},
		{name: "basis 2", args: numArgs(39494, 43055, 0.0575, 94.6024, 100, 2, 2), want: 0.0650},
		{name: "basis 3", args: numArgs(39494, 43055, 0.0575, 94.6436, 100, 2, 3), want: 0.0650},
		{name: "basis 4", args: numArgs(39494, 43055, 0.0575, 94.6344, 100, 2, 4), want: 0.0650},

		// --- Single coupon period (N=1) ---
		{name: "single period", args: numArgs(43984, 44167, 0.06, 100.4878, 100, 2, 0), want: 0.0500},

		// --- Zero coupon (multi-period) ---
		{name: "zero coupon", args: numArgs(43832, 47485, 0, 55.3676, 100, 2, 0), want: 0.0600},

		// --- Premium bond (pr > 100, low yield) ---
		{name: "premium bond pr>100", args: numArgs(39494, 43055, 0.10, 124.9660, 100, 2, 0), want: 0.0650},

		// --- Discount bond (pr < 100, high yield) ---
		{name: "discount bond pr<100", args: numArgs(39494, 43055, 0.03, 75.0080, 100, 2, 0), want: 0.0650},

		// --- Par bond (pr = 100, yield ≈ rate) ---
		{name: "par bond", args: numArgs(43832, 47485, 0.06, 100.0, 100, 2, 0), want: 0.0600},

		// --- High yield ---
		{name: "high yield 20%", args: numArgs(39494, 43055, 0.0575, 39.8235, 100, 2, 0), want: 0.2000},

		// --- Low yield ---
		{name: "low yield 0.5%", args: numArgs(39494, 43055, 0.0575, 149.8981, 100, 2, 0), want: 0.0050},

		// --- Non-100 redemption ---
		{name: "redemption 110", args: numArgs(39494, 43055, 0.0575, 99.9941, 110, 2, 0), want: 0.0650},

		// --- Default basis (6 args, no basis) ---
		{name: "default basis", args: numArgs(39494, 42690, 0.0575, 95.04287, 100, 2), want: 0.0650},

		// --- Error cases ---
		{name: "too few args", args: numArgs(39494, 43055, 0.0575, 95, 100), wantErr: true},
		{name: "too many args", args: numArgs(39494, 43055, 0.0575, 95, 100, 2, 0, 99), wantErr: true},
		{name: "negative rate", args: numArgs(39494, 43055, -0.05, 95, 100, 2, 0), wantErr: true},
		{name: "zero price", args: numArgs(39494, 43055, 0.0575, 0, 100, 2, 0), wantErr: true},
		{name: "negative price", args: numArgs(39494, 43055, 0.0575, -10, 100, 2, 0), wantErr: true},
		{name: "zero redemption", args: numArgs(39494, 43055, 0.0575, 95, 0, 2, 0), wantErr: true},
		{name: "negative redemption", args: numArgs(39494, 43055, 0.0575, 95, -100, 2, 0), wantErr: true},
		{name: "settlement >= maturity", args: numArgs(43055, 39494, 0.0575, 95, 100, 2, 0), wantErr: true},
		{name: "settlement == maturity", args: numArgs(39494, 39494, 0.0575, 95, 100, 2, 0), wantErr: true},
		{name: "bad frequency 3", args: numArgs(39494, 43055, 0.0575, 95, 100, 3, 0), wantErr: true},
		{name: "bad frequency 0", args: numArgs(39494, 43055, 0.0575, 95, 100, 0, 0), wantErr: true},
		{name: "bad basis -1", args: numArgs(39494, 43055, 0.0575, 95, 100, 2, -1), wantErr: true},
		{name: "bad basis 5", args: numArgs(39494, 43055, 0.0575, 95, 100, 2, 5), wantErr: true},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			v, err := fnYield(tc.args)
			if err != nil {
				t.Fatal(err)
			}
			if tc.wantErr {
				assertError(t, tc.name, v)
			} else {
				assertClose(t, tc.name, v, tc.want)
			}
		})
	}
}

func TestYIELD_ViaEval(t *testing.T) {
	cf := evalCompile(t, "YIELD(DATE(2008,2,15), DATE(2016,11,15), 0.0575, 95.04287, 100, 2, 0)")
	v, err := Eval(cf, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	if math.Abs(v.Num-0.065) > 0.01 {
		t.Errorf("got %f, want ~0.065", v.Num)
	}
}

// TestPRICE_YIELD_RoundTrip verifies that YIELD(PRICE(yld)) ≈ yld for various inputs.
func TestPRICE_YIELD_RoundTrip(t *testing.T) {
	type roundTrip struct {
		name                  string
		settlement, maturity  float64
		rate, yld, redemption float64
		freq, basis           int
	}

	trips := []roundTrip{
		{"doc example semi b0", 39494, 43055, 0.0575, 0.065, 100, 2, 0},
		{"annual b0", 43832, 47485, 0.05, 0.06, 100, 1, 0},
		{"semiannual b0", 43832, 47485, 0.05, 0.06, 100, 2, 0},
		{"quarterly b0", 43832, 47485, 0.05, 0.06, 100, 4, 0},
		{"basis 1", 39494, 43055, 0.0575, 0.065, 100, 2, 1},
		{"single period", 43984, 44167, 0.06, 0.05, 100, 2, 0},
		{"premium bond", 39494, 43055, 0.10, 0.065, 100, 2, 0},
		{"discount bond", 39494, 43055, 0.03, 0.065, 100, 2, 0},
		{"zero coupon", 43832, 47485, 0, 0.06, 100, 2, 0},
		{"high yield", 39494, 43055, 0.0575, 0.20, 100, 2, 0},
		{"low yield", 39494, 43055, 0.0575, 0.005, 100, 2, 0},
		{"redemption 110", 39494, 43055, 0.0575, 0.065, 110, 2, 0},
	}

	for _, tr := range trips {
		t.Run(tr.name, func(t *testing.T) {
			// Compute price from yield.
			pv, err := fnPrice(numArgs(tr.settlement, tr.maturity, tr.rate, tr.yld, tr.redemption, float64(tr.freq), float64(tr.basis)))
			if err != nil {
				t.Fatal(err)
			}
			if pv.Type != ValueNumber {
				t.Fatalf("PRICE: expected number, got %v (%s)", pv.Type, pv.Str)
			}

			// Compute yield from that price.
			yv, err := fnYield(numArgs(tr.settlement, tr.maturity, tr.rate, pv.Num, tr.redemption, float64(tr.freq), float64(tr.basis)))
			if err != nil {
				t.Fatal(err)
			}
			if yv.Type != ValueNumber {
				t.Fatalf("YIELD: expected number, got %v (%s)", yv.Type, yv.Str)
			}

			if math.Abs(yv.Num-tr.yld) > 1e-6 {
				t.Errorf("round-trip: got yield %f, want %f (price was %f)", yv.Num, tr.yld, pv.Num)
			}
		})
	}
}

// === FVSCHEDULE ===

func TestFVSchedule_Comprehensive(t *testing.T) {
	mkArr := func(vals ...Value) Value {
		return Value{Type: ValueArray, Array: [][]Value{vals}}
	}

	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// Basic usage — documentation example
		{
			name: "doc example: FVSCHEDULE(1,{0.09,0.11,0.1})",
			args: []Value{NumberVal(1), mkArr(NumberVal(0.09), NumberVal(0.11), NumberVal(0.1))},
			want: 1.33089,
		},
		// Single rate
		{
			name: "single rate 10%",
			args: []Value{NumberVal(1000), mkArr(NumberVal(0.10))},
			want: 1100.0,
		},
		// Multiple rates
		{
			name: "multiple rates 5%, 10%, 15%",
			args: []Value{NumberVal(1000), mkArr(NumberVal(0.05), NumberVal(0.10), NumberVal(0.15))},
			want: 1000 * 1.05 * 1.10 * 1.15,
		},
		// Zero rates
		{
			name: "all zero rates",
			args: []Value{NumberVal(500), mkArr(NumberVal(0), NumberVal(0), NumberVal(0))},
			want: 500.0,
		},
		// Negative rates
		{
			name: "negative rate -10%",
			args: []Value{NumberVal(1000), mkArr(NumberVal(-0.10))},
			want: 900.0,
		},
		{
			name: "negative rate -50%",
			args: []Value{NumberVal(1000), mkArr(NumberVal(-0.50))},
			want: 500.0,
		},
		// Principal of 0
		{
			name: "principal zero",
			args: []Value{NumberVal(0), mkArr(NumberVal(0.1), NumberVal(0.2))},
			want: 0.0,
		},
		// Negative principal
		{
			name: "negative principal",
			args: []Value{NumberVal(-1000), mkArr(NumberVal(0.10))},
			want: -1100.0,
		},
		// Large principal
		{
			name: "large principal",
			args: []Value{NumberVal(1e9), mkArr(NumberVal(0.05), NumberVal(0.03))},
			want: 1e9 * 1.05 * 1.03,
		},
		// Empty cells in schedule treated as 0
		{
			name: "empty cell in schedule treated as zero",
			args: []Value{NumberVal(1000), mkArr(NumberVal(0.10), EmptyVal(), NumberVal(0.05))},
			want: 1000 * 1.10 * 1.0 * 1.05,
		},
		{
			name: "all empty cells in schedule",
			args: []Value{NumberVal(1000), mkArr(EmptyVal(), EmptyVal())},
			want: 1000.0,
		},
		// Single element array
		{
			name: "single element array",
			args: []Value{NumberVal(100), mkArr(NumberVal(0.5))},
			want: 150.0,
		},
		// Mixed positive and negative rates
		{
			name: "mixed positive and negative rates",
			args: []Value{NumberVal(1000), mkArr(NumberVal(0.10), NumberVal(-0.05), NumberVal(0.20))},
			want: 1000 * 1.10 * 0.95 * 1.20,
		},
		// Very small rates
		{
			name: "very small rate 0.001%",
			args: []Value{NumberVal(10000), mkArr(NumberVal(0.00001))},
			want: 10000 * 1.00001,
		},
		// Very large rate
		{
			name: "very large rate 1000%",
			args: []Value{NumberVal(1), mkArr(NumberVal(10.0))},
			want: 11.0,
		},
		// Multiple periods with same rate
		{
			name: "three periods same rate 5%",
			args: []Value{NumberVal(1000), mkArr(NumberVal(0.05), NumberVal(0.05), NumberVal(0.05))},
			want: 1000 * 1.05 * 1.05 * 1.05,
		},
		// Rate of -100% (complete loss)
		{
			name: "rate of -100% (complete loss)",
			args: []Value{NumberVal(1000), mkArr(NumberVal(-1.0))},
			want: 0.0,
		},
		// Boolean FALSE in schedule → #VALUE! (spreadsheet rejects booleans)
		{
			name:    "boolean FALSE in schedule is #VALUE!",
			args:    []Value{NumberVal(1000), mkArr(BoolVal(false), NumberVal(0.10))},
			wantErr: true,
		},
		// Boolean TRUE in schedule → #VALUE! (spreadsheet rejects booleans)
		{
			name:    "boolean TRUE in schedule is #VALUE!",
			args:    []Value{NumberVal(1000), mkArr(BoolVal(true))},
			wantErr: true,
		},
		// Fractional principal
		{
			name: "fractional principal 0.5",
			args: []Value{NumberVal(0.5), mkArr(NumberVal(0.10), NumberVal(0.20))},
			want: 0.5 * 1.10 * 1.20,
		},
		// Many rates
		{
			name: "five rates",
			args: []Value{NumberVal(100), mkArr(NumberVal(0.01), NumberVal(0.02), NumberVal(0.03), NumberVal(0.04), NumberVal(0.05))},
			want: 100 * 1.01 * 1.02 * 1.03 * 1.04 * 1.05,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			v, err := fnFVSchedule(tt.args)
			if err != nil {
				t.Fatal(err)
			}
			if tt.wantErr {
				assertError(t, tt.name, v)
				return
			}
			assertClose(t, tt.name, v, tt.want)
		})
	}
}

func TestFVSchedule_ErrorCases(t *testing.T) {
	mkArr := func(vals ...Value) Value {
		return Value{Type: ValueArray, Array: [][]Value{vals}}
	}

	t.Run("wrong number of args: too few", func(t *testing.T) {
		v, err := fnFVSchedule([]Value{NumberVal(1)})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "too few args", v)
	})

	t.Run("wrong number of args: too many", func(t *testing.T) {
		v, err := fnFVSchedule([]Value{NumberVal(1), mkArr(NumberVal(0.1)), NumberVal(2)})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "too many args", v)
	})

	t.Run("wrong number of args: zero", func(t *testing.T) {
		v, err := fnFVSchedule([]Value{})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "zero args", v)
	})

	t.Run("non-numeric principal string", func(t *testing.T) {
		v, err := fnFVSchedule([]Value{StringVal("abc"), mkArr(NumberVal(0.1))})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "string principal", v)
	})

	t.Run("string value in schedule", func(t *testing.T) {
		v, err := fnFVSchedule([]Value{NumberVal(1000), mkArr(NumberVal(0.1), StringVal("abc"))})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "string in schedule", v)
	})

	t.Run("error value in principal", func(t *testing.T) {
		v, err := fnFVSchedule([]Value{ErrorVal(ErrValNUM), mkArr(NumberVal(0.1))})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "error principal", v)
	})

	t.Run("error value in schedule", func(t *testing.T) {
		v, err := fnFVSchedule([]Value{NumberVal(1000), mkArr(NumberVal(0.1), ErrorVal(ErrValDIV0))})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "error in schedule", v)
	})

	t.Run("empty string in schedule is VALUE error", func(t *testing.T) {
		v, err := fnFVSchedule([]Value{NumberVal(1000), mkArr(StringVal(""))})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "empty string in schedule", v)
	})
}

// === AMORDEGRC ===

func TestAMORDEGRC_Comprehensive(t *testing.T) {
	// Serial numbers for reference dates:
	// DATE(2008,8,18)  = 39679
	// DATE(2008,12,30) = 39813
	// DATE(2010,1,1)   = 40179
	// DATE(2010,6,30)  = 40359
	// DATE(2010,12,31) = 40543
	// DATE(2011,1,1)   = 40544
	// DATE(2011,12,31) = 40908
	// DATE(2012,1,1)   = 40909
	// DATE(2012,6,30)  = 41090

	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// --- Doc example ---
		// AMORDEGRC(2400, 39679, 39813, 300, 1, 0.15, 1) = 776
		{
			name: "doc example period 1 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 1, 0.15, 1),
			want: 776,
		},
		// --- Period 0 (prorated first period) ---
		{
			name: "doc example period 0 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 0, 0.15, 1),
			want: 330,
		},
		// --- Multiple periods (0, 1, 2, 3, ...) ---
		// Period 2: remaining = 2400 - 330 - 776 = 1294, dep = round(1294 * 0.375) = 485
		{
			name: "doc example period 2 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 2, 0.15, 1),
			want: 485,
		},
		// Period 3: remaining = 1294 - 485 = 809, dep = round(809 * 0.375) = 303
		// But remaining-salvage = 809 - 300 = 509, dep 303 < 509 so normal
		{
			name: "doc example period 3 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 3, 0.15, 1),
			want: 303,
		},
		// Period 4: remaining = 809 - 303 = 506, dep = round(506 * 0.375) = 190
		// remaining-salvage = 506 - 300 = 206, dep 190 < 206 so normal
		{
			name: "doc example period 4 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 4, 0.15, 1),
			want: 190,
		},
		// Period 5 (second-to-last): remaining = 316, round(316 * 0.5) = 158
		{
			name: "doc example period 5 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 5, 0.15, 1),
			want: 158,
		},
		// Period 6 (last): remaining 316-158 = 158, but 158 < salvage 300 → 0
		{
			name: "doc example period 6 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 6, 0.15, 1),
			want: 0,
		},
		// Period 7: beyond asset life, return 0
		{
			name: "doc example period 7 beyond life",
			args: numArgs(2400, 39679, 39813, 300, 7, 0.15, 1),
			want: 0,
		},
		// --- Rate 0.10 (life=10, coeff=2.5) regression: period 3 rounding ---
		// cost=50000, datePurchased=40000, firstPeriod=40365, salvage=5000, rate=0.10, basis=1
		// adjustedRate=0.25, yearFrac=1.0
		// dep0 = round(50000*0.25*1.0) = 12500
		// dep1 = round(37500*0.25) = 9375
		// dep2 = round(28125*0.25) = round(7031.25) = 7031
		// dep3 = round(21094*0.25) = round(5273.5) = 5273 (half toward zero, not 5274)
		{
			name: "rate 0.10 coeff 2.5 period 0",
			args: numArgs(50000, 40000, 40365, 5000, 0, 0.10, 1),
			want: 12500,
		},
		{
			name: "rate 0.10 coeff 2.5 period 1",
			args: numArgs(50000, 40000, 40365, 5000, 1, 0.10, 1),
			want: 9375,
		},
		{
			name: "rate 0.10 coeff 2.5 period 2",
			args: numArgs(50000, 40000, 40365, 5000, 2, 0.10, 1),
			want: 7031,
		},
		{
			name: "rate 0.10 coeff 2.5 period 3 half toward zero",
			args: numArgs(50000, 40000, 40365, 5000, 3, 0.10, 1),
			want: 5273,
		},
		// --- Different rate: life 3-4 years, coeff=1.5 ---
		// rate=0.3, life=3.33, coeff=1.5, adjusted_rate=0.45
		// basis 0: 30/360 day count
		// DATE(2010,1,1)=40179, DATE(2010,6,30)=40359
		// days360 NASD: 6*30 - 1 = 179? Let's use basis 3 for simplicity.
		// basis 3: dsm = 40359-40179 = 180, bYear = 365
		// yearFrac = 180/365 = 0.493151
		// dep0 = round(10000 * 0.45 * 0.493151) = round(2219.18) = 2219
		{
			name: "rate 0.3 life 3.33 coeff 1.5 period 0 basis 3",
			args: numArgs(10000, 40179, 40359, 500, 0, 0.3, 3),
			want: 2219,
		},
		// --- Different rate: life 5-6 years, coeff=2 ---
		// rate=0.2, life=5, coeff=2, adjusted_rate=0.4
		// basis 3: dsm = 180, bYear = 365
		// dep0 = round(10000 * 0.4 * 180/365) = round(1972.60) = 1973
		{
			name: "rate 0.2 life 5 coeff 2 period 0 basis 3",
			args: numArgs(10000, 40179, 40359, 500, 0, 0.2, 3),
			want: 1973,
		},
		// --- Rate giving life < 1 (#NUM!) ---
		{
			name:    "rate 2.0 life 0.5 NUM error",
			args:    numArgs(2400, 39679, 39813, 300, 0, 2.0, 1),
			wantErr: true,
		},
		// --- Rate giving life 1-2 (#NUM!) ---
		{
			name:    "rate 0.7 life 1.43 NUM error",
			args:    numArgs(2400, 39679, 39813, 300, 0, 0.7, 1),
			wantErr: true,
		},
		// --- Rate giving life = 2 exactly (#NUM!) ---
		{
			name:    "rate 0.5 life 2.0 NUM error",
			args:    numArgs(2400, 39679, 39813, 300, 0, 0.5, 1),
			wantErr: true,
		},
		// --- Rate giving life 2.5 (accepted, coeff=1.5) ---
		// dep0 = round(2400 * 0.4 * 1.5 * 134/366) = round(527.21) = 527
		{
			name: "rate 0.4 life 2.5 coeff 1.5",
			args: numArgs(2400, 39679, 39813, 300, 0, 0.4, 1),
			want: 527,
		},
		// --- Basis 2 not supported (#NUM!) ---
		{
			name:    "basis 2 not supported NUM error",
			args:    numArgs(2400, 39679, 39813, 300, 1, 0.15, 2),
			wantErr: true,
		},
		// --- Basis 0 (US NASD 30/360) ---
		// AMORDEGRC(2400, 39679, 39813, 300, 0, 0.15, 0)
		// days360 NASD from 8/18 to 12/30: (12-8)*30 + (30-18) = 120+12 = 132
		// bYear=360, yearFrac = 132/360 = 0.366667
		// dep0 = round(2400 * 0.375 * 0.366667) = round(330) = 330
		{
			name: "basis 0 US NASD 30/360",
			args: numArgs(2400, 39679, 39813, 300, 0, 0.15, 0),
			want: 330,
		},
		// --- Basis 3 (actual/365) ---
		// dsm = 134, bYear = 365
		// yearFrac = 134/365 = 0.367123
		// dep0 = round(2400 * 0.375 * 0.367123) = round(330.41) = 330
		{
			name: "basis 3 actual/365",
			args: numArgs(2400, 39679, 39813, 300, 0, 0.15, 3),
			want: 330,
		},
		// --- Basis 4 (European 30/360) ---
		// 39679=2008-08-19, 39813=2008-12-31. European: ed 31→30.
		// days360 = (12-8)*30 + (30-19) = 131, bYear=360
		// yearFrac = 131/360, dep0 = round(2400 * 0.375 * 131/360) = 328
		{
			name: "basis 4 European 30/360",
			args: numArgs(2400, 39679, 39813, 300, 0, 0.15, 4),
			want: 328,
		},
		// --- Default basis (omitted, should default to 0) ---
		{
			name: "default basis omitted 6 args",
			args: numArgs(2400, 39679, 39813, 300, 0, 0.15),
			want: 330,
		},
		// --- Salvage = 0 ---
		// AMORDEGRC(2400, 39679, 39813, 0, 0, 0.15, 1)
		// Same as before but salvage=0
		// dep0 = 330
		{
			name: "salvage zero period 0",
			args: numArgs(2400, 39679, 39813, 0, 0, 0.15, 1),
			want: 330,
		},
		// --- Negative cost (#NUM!) ---
		{
			name:    "negative cost NUM error",
			args:    numArgs(-2400, 39679, 39813, 300, 1, 0.15, 1),
			wantErr: true,
		},
		// --- Negative salvage (#NUM!) ---
		{
			name:    "negative salvage NUM error",
			args:    numArgs(2400, 39679, 39813, -300, 1, 0.15, 1),
			wantErr: true,
		},
		// --- Salvage > cost (#NUM!) ---
		{
			name:    "salvage exceeds cost NUM error",
			args:    numArgs(2400, 39679, 39813, 3000, 1, 0.15, 1),
			wantErr: true,
		},
		// --- Negative period (#NUM!) ---
		{
			name:    "negative period NUM error",
			args:    numArgs(2400, 39679, 39813, 300, -1, 0.15, 1),
			wantErr: true,
		},
		// --- Negative rate (#NUM!) ---
		{
			name:    "negative rate NUM error",
			args:    numArgs(2400, 39679, 39813, 300, 1, -0.15, 1),
			wantErr: true,
		},
		// --- Zero rate (#NUM!) ---
		{
			name:    "zero rate NUM error",
			args:    numArgs(2400, 39679, 39813, 300, 1, 0, 1),
			wantErr: true,
		},
		// --- Period beyond asset life returns 0 ---
		{
			name: "period 10 beyond life returns 0",
			args: numArgs(2400, 39679, 39813, 300, 10, 0.15, 1),
			want: 0,
		},
		// --- Invalid basis 5 (#NUM!) ---
		{
			name:    "basis 5 invalid NUM error",
			args:    numArgs(2400, 39679, 39813, 300, 1, 0.15, 5),
			wantErr: true,
		},
		// --- Invalid basis -1 (#NUM!) ---
		{
			name:    "basis negative invalid NUM error",
			args:    numArgs(2400, 39679, 39813, 300, 1, 0.15, -1),
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			v, err := fnAmordegrc(tt.args)
			if err != nil {
				t.Fatal(err)
			}
			if tt.wantErr {
				assertError(t, tt.name, v)
				return
			}
			if v.Type != ValueNumber {
				t.Fatalf("%s: expected number, got type %v (str=%q)", tt.name, v.Type, v.Str)
			}
			if v.Num != tt.want {
				t.Errorf("%s: got %f, want %f", tt.name, v.Num, tt.want)
			}
		})
	}
}

func TestAMORDEGRC_WrongArgCount(t *testing.T) {
	t.Run("too few args", func(t *testing.T) {
		v, err := fnAmordegrc(numArgs(2400, 39679, 39813, 300, 1))
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "too few args", v)
	})

	t.Run("too many args", func(t *testing.T) {
		v, err := fnAmordegrc(numArgs(2400, 39679, 39813, 300, 1, 0.15, 1, 99))
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "too many args", v)
	})
}

func TestAMORDEGRC_ErrorPropagation(t *testing.T) {
	errVal := ErrorVal(ErrValDIV0)

	t.Run("error in cost", func(t *testing.T) {
		v, err := fnAmordegrc([]Value{errVal, NumberVal(39679), NumberVal(39813), NumberVal(300), NumberVal(1), NumberVal(0.15), NumberVal(1)})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "error cost", v)
	})

	t.Run("error in rate", func(t *testing.T) {
		v, err := fnAmordegrc([]Value{NumberVal(2400), NumberVal(39679), NumberVal(39813), NumberVal(300), NumberVal(1), errVal, NumberVal(1)})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "error rate", v)
	})

	t.Run("error in basis", func(t *testing.T) {
		v, err := fnAmordegrc([]Value{NumberVal(2400), NumberVal(39679), NumberVal(39813), NumberVal(300), NumberVal(1), NumberVal(0.15), errVal})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "error basis", v)
	})
}

// === AMORLINC ===

func TestAMORLINC_Comprehensive(t *testing.T) {
	// Serial numbers for reference dates:
	// DATE(2008,8,18)  = 39679
	// DATE(2008,12,30) = 39813
	// DATE(2010,1,1)   = 40179
	// DATE(2010,6,30)  = 40359

	tests := []struct {
		name    string
		args    []Value
		want    float64
		wantErr bool
	}{
		// --- Doc example: AMORLINC(2400, 39679, 39813, 300, 1, 0.15, 1) = 360 ---
		{
			name: "doc example period 1 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 1, 0.15, 1),
			want: 360,
		},
		// --- Period 0 (prorated first period) ---
		// basis 1: dsm=134, bYear=366, yearFrac=134/366
		// dep0 = 2400 * 0.15 * 134/366 = 131.803279...
		{
			name: "doc example period 0 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 0, 0.15, 1),
			want: 131.80327868852460,
		},
		// --- Multiple periods ---
		// nper = ceil(1/0.15) = 7
		// Periods 1..6: normalDep = 2400 * 0.15 = 360
		{
			name: "doc example period 2 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 2, 0.15, 1),
			want: 360,
		},
		{
			name: "doc example period 3 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 3, 0.15, 1),
			want: 360,
		},
		{
			name: "doc example period 4 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 4, 0.15, 1),
			want: 360,
		},
		{
			name: "doc example period 5 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 5, 0.15, 1),
			want: 360,
		},
		{
			name: "doc example period 6 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 6, 0.15, 1),
			want: 168.19672131147536,
		},
		// Period 7: beyond depreciable amount, returns 0
		{
			name: "doc example period 7 last period",
			args: numArgs(2400, 39679, 39813, 300, 7, 0.15, 1),
			want: 0,
		},
		// Period 8: beyond asset life, return 0
		{
			name: "doc example period 8 beyond life",
			args: numArgs(2400, 39679, 39813, 300, 8, 0.15, 1),
			want: 0,
		},
		// --- Basis 0 (US NASD 30/360) ---
		// dsm=132, bYear=360, dep0 = 2400 * 0.15 * 132/360 = 132
		{
			name: "basis 0 period 0",
			args: numArgs(2400, 39679, 39813, 300, 0, 0.15, 0),
			want: 132,
		},
		{
			name: "basis 0 period 1",
			args: numArgs(2400, 39679, 39813, 300, 1, 0.15, 0),
			want: 360,
		},
		// --- Basis 3 (actual/365) ---
		// dsm=134, bYear=365, dep0 = 2400 * 0.15 * 134/365 = 132.164383...
		{
			name: "basis 3 period 0",
			args: numArgs(2400, 39679, 39813, 300, 0, 0.15, 3),
			want: 132.16438356164384,
		},
		// --- Basis 4 (European 30/360) ---
		// dsm=131, bYear=360, dep0 = 2400 * 0.15 * 131/360 = 131.0
		{
			name: "basis 4 period 0",
			args: numArgs(2400, 39679, 39813, 300, 0, 0.15, 4),
			want: 131.0,
		},
		// --- Default basis (omitted, should default to 0) ---
		{
			name: "default basis omitted 6 args",
			args: numArgs(2400, 39679, 39813, 300, 0, 0.15),
			want: 132,
		},
		// --- Rate = 0.25, cost=10000, salvage=500, basis=3 ---
		// datePurchased=40179 (2010-01-01), firstPeriod=40359 (2010-06-30)
		// nper = ceil(1/0.25) = 4
		// dsm = 180, bYear = 365
		// dep0 = 10000 * 0.25 * 180/365 = 1232.876712...
		// normalDep = 10000 * 0.25 = 2500
		// period 4 (last): 10000 - 500 - 1232.876712 - 3*2500 = 767.123288...
		{
			name: "rate 0.25 period 0 basis 3",
			args: numArgs(10000, 40179, 40359, 500, 0, 0.25, 3),
			want: 1232.8767123287671,
		},
		{
			name: "rate 0.25 period 1 basis 3",
			args: numArgs(10000, 40179, 40359, 500, 1, 0.25, 3),
			want: 2500,
		},
		{
			name: "rate 0.25 period 3 basis 3",
			args: numArgs(10000, 40179, 40359, 500, 3, 0.25, 3),
			want: 2500,
		},
		{
			name: "rate 0.25 period 4 last basis 3",
			args: numArgs(10000, 40179, 40359, 500, 4, 0.25, 3),
			want: 767.1232876712329,
		},
		{
			name: "rate 0.25 period 5 beyond life",
			args: numArgs(10000, 40179, 40359, 500, 5, 0.25, 3),
			want: 0,
		},
		// --- Rate = 0.5, cost=2400, salvage=300, basis=1 ---
		// nper = ceil(1/0.5) = 2
		// dep0 = 2400 * 0.5 * 134/366 = 439.344262...
		// normalDep = 2400 * 0.5 = 1200
		// period 2 (last): 2400 - 300 - 439.344 - 1*1200 = 460.655738...
		{
			name: "rate 0.5 period 0 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 0, 0.5, 1),
			want: 439.34426229508197,
		},
		{
			name: "rate 0.5 period 1 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 1, 0.5, 1),
			want: 1200,
		},
		{
			name: "rate 0.5 period 2 last basis 1",
			args: numArgs(2400, 39679, 39813, 300, 2, 0.5, 1),
			want: 460.65573770491803,
		},
		// --- Rate = 1.0 ---
		// nper = ceil(1/1) = 1
		// dep0 = 2400 * 1.0 * 134/366 = 878.688525...
		// period 1 (last): 2400 - 300 - 878.688525 - 0*2400 = 1221.311475...
		{
			name: "rate 1.0 period 0 basis 1",
			args: numArgs(2400, 39679, 39813, 300, 0, 1.0, 1),
			want: 878.68852459016393,
		},
		{
			name: "rate 1.0 period 1 last basis 1",
			args: numArgs(2400, 39679, 39813, 300, 1, 1.0, 1),
			want: 1221.3114754098361,
		},
		// --- Zero cost → #NUM! ---
		{
			name:    "zero cost period 0",
			args:    numArgs(0, 39679, 39813, 0, 0, 0.15, 1),
			wantErr: true,
		},
		{
			name:    "zero cost period 1",
			args:    numArgs(0, 39679, 39813, 0, 1, 0.15, 1),
			wantErr: true,
		},
		// --- Salvage = cost → nothing to depreciate, always 0 ---
		{
			name: "salvage equals cost period 0",
			args: numArgs(2400, 39679, 39813, 2400, 0, 0.15, 1),
			want: 0,
		},
		{
			name: "salvage equals cost period 1",
			args: numArgs(2400, 39679, 39813, 2400, 1, 0.15, 1),
			want: 0,
		},
		// --- Error cases ---
		// Negative cost
		{
			name:    "negative cost NUM error",
			args:    numArgs(-2400, 39679, 39813, 300, 1, 0.15, 1),
			wantErr: true,
		},
		// Negative salvage
		{
			name:    "negative salvage NUM error",
			args:    numArgs(2400, 39679, 39813, -300, 1, 0.15, 1),
			wantErr: true,
		},
		// Salvage > cost
		{
			name:    "salvage exceeds cost NUM error",
			args:    numArgs(2400, 39679, 39813, 3000, 1, 0.15, 1),
			wantErr: true,
		},
		// Negative period
		{
			name:    "negative period NUM error",
			args:    numArgs(2400, 39679, 39813, 300, -1, 0.15, 1),
			wantErr: true,
		},
		// Negative rate
		{
			name:    "negative rate NUM error",
			args:    numArgs(2400, 39679, 39813, 300, 1, -0.15, 1),
			wantErr: true,
		},
		// Zero rate
		{
			name:    "zero rate NUM error",
			args:    numArgs(2400, 39679, 39813, 300, 1, 0, 1),
			wantErr: true,
		},
		// Invalid basis 2
		{
			name:    "basis 2 invalid NUM error",
			args:    numArgs(2400, 39679, 39813, 300, 1, 0.15, 2),
			wantErr: true,
		},
		// Invalid basis 5
		{
			name:    "basis 5 invalid NUM error",
			args:    numArgs(2400, 39679, 39813, 300, 1, 0.15, 5),
			wantErr: true,
		},
		// Invalid basis -1
		{
			name:    "basis negative invalid NUM error",
			args:    numArgs(2400, 39679, 39813, 300, 1, 0.15, -1),
			wantErr: true,
		},
		// Period far beyond life
		{
			name: "period 100 beyond life returns 0",
			args: numArgs(2400, 39679, 39813, 300, 100, 0.15, 1),
			want: 0,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			v, err := fnAmorlinc(tt.args)
			if err != nil {
				t.Fatal(err)
			}
			if tt.wantErr {
				assertError(t, tt.name, v)
				return
			}
			if v.Type != ValueNumber {
				t.Fatalf("%s: expected number, got type %v (str=%q)", tt.name, v.Type, v.Str)
			}
			if math.Abs(v.Num-tt.want) > 1e-9 {
				t.Errorf("%s: got %.15f, want %.15f", tt.name, v.Num, tt.want)
			}
		})
	}
}

func TestAMORLINC_WrongArgCount(t *testing.T) {
	t.Run("too few args", func(t *testing.T) {
		v, err := fnAmorlinc(numArgs(2400, 39679, 39813, 300, 1))
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "too few args", v)
	})

	t.Run("too many args", func(t *testing.T) {
		v, err := fnAmorlinc(numArgs(2400, 39679, 39813, 300, 1, 0.15, 1, 99))
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "too many args", v)
	})
}

func TestAMORLINC_ErrorPropagation(t *testing.T) {
	errVal := ErrorVal(ErrValDIV0)

	t.Run("error in cost", func(t *testing.T) {
		v, err := fnAmorlinc([]Value{errVal, NumberVal(39679), NumberVal(39813), NumberVal(300), NumberVal(1), NumberVal(0.15), NumberVal(1)})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "error cost", v)
	})

	t.Run("error in rate", func(t *testing.T) {
		v, err := fnAmorlinc([]Value{NumberVal(2400), NumberVal(39679), NumberVal(39813), NumberVal(300), NumberVal(1), errVal, NumberVal(1)})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "error rate", v)
	})

	t.Run("error in basis", func(t *testing.T) {
		v, err := fnAmorlinc([]Value{NumberVal(2400), NumberVal(39679), NumberVal(39813), NumberVal(300), NumberVal(1), NumberVal(0.15), errVal})
		if err != nil {
			t.Fatal(err)
		}
		assertError(t, "error basis", v)
	})
}

func TestAMORLINC_StringCoercion(t *testing.T) {
	// String "2400" should coerce to number 2400.
	v, err := fnAmorlinc([]Value{
		StringVal("2400"),
		NumberVal(39679), NumberVal(39813),
		NumberVal(300), NumberVal(1), NumberVal(0.15), NumberVal(1),
	})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got type %v", v.Type)
	}
	if math.Abs(v.Num-360) > 1e-9 {
		t.Errorf("got %f, want 360", v.Num)
	}
}

func TestAMORLINC_BoolCoercion(t *testing.T) {
	// TRUE coerces to 1 for cost, so cost=1, rate=0.15
	// normalDep = 1 * 0.15 = 0.15
	v, err := fnAmorlinc([]Value{
		BoolVal(true),
		NumberVal(39679), NumberVal(39813),
		NumberVal(0), NumberVal(1), NumberVal(0.15), NumberVal(1),
	})
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != ValueNumber {
		t.Fatalf("expected number, got type %v", v.Type)
	}
	if math.Abs(v.Num-0.15) > 1e-9 {
		t.Errorf("got %f, want 0.15", v.Num)
	}
}
