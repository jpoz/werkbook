package formula

import (
	"fmt"
	"math"
	"strings"
	"time"
)

func init() {
	Register("FV", NoCtx(fnFV))
	Register("IPMT", NoCtx(fnIPMT))
	Register("IRR", NoCtx(fnIRR))
	Register("NPER", NoCtx(fnNPER))
	Register("NPV", NoCtx(fnNPV))
	Register("PMT", NoCtx(fnPMT))
	Register("PPMT", NoCtx(fnPPMT))
	Register("PV", NoCtx(fnPV))
	Register("RATE", NoCtx(fnRATE))
	Register("DB", NoCtx(fnDB))
	Register("DDB", NoCtx(fnDDB))
	Register("SLN", NoCtx(fnSLN))
	Register("XIRR", NoCtx(fnXIRR))
	Register("XNPV", NoCtx(fnXNPV))
	Register("DOLLARDE", NoCtx(fnDOLLARDE))
	Register("DOLLARFR", NoCtx(fnDOLLARFR))
	Register("EFFECT", NoCtx(fnEFFECT))
	Register("NOMINAL", NoCtx(fnNOMINAL))
	Register("CUMIPMT", NoCtx(fnCumipmt))
	Register("CUMPRINC", NoCtx(fnCumprinc))
	Register("MIRR", NoCtx(fnMirr))
	Register("PDURATION", NoCtx(fnPduration))
	Register("RRI", NoCtx(fnRri))
}

// flattenValues extracts all numeric values from an arg that may be a scalar or array (range).
func flattenValues(arg Value) []Value {
	if arg.Type == ValueArray {
		total := 0
		for _, row := range arg.Array {
			total += len(row)
		}
		out := make([]Value, 0, total)
		for _, row := range arg.Array {
			out = append(out, row...)
		}
		return out
	}
	return []Value{arg}
}

// pmtCore computes the payment for a loan based on constant payments and a constant interest rate.
// This is shared by PMT, IPMT, PPMT, FV, and other financial functions.
func pmtCore(rate, nper, pv, fv float64, payType int) float64 {
	if rate == 0 {
		return -(pv + fv) / nper
	}
	term := math.Pow(1+rate, nper)
	if payType == 1 {
		return -(pv*term + fv) / ((1 + rate) * (term - 1) / rate)
	}
	return -(pv*term + fv) / ((term - 1) / rate)
}

// fvCore computes the future value.
func fvCore(rate, nper, pmt, pv float64, payType int) float64 {
	if rate == 0 {
		return -pv - pmt*nper
	}
	term := math.Pow(1+rate, nper)
	if payType == 1 {
		return -pv*term - pmt*(1+rate)*(term-1)/rate
	}
	return -pv*term - pmt*(term-1)/rate
}

// fnPMT implements PMT(rate, nper, pv, [fv], [type]).
func fnPMT(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	rate, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	nper, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pv, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	if nper == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	fv := 0.0
	if len(args) >= 4 {
		fv, e = CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
	}
	payType := 0
	if len(args) == 5 {
		pt, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		if pt != 0 {
			payType = 1
		}
	}
	return NumberVal(pmtCore(rate, nper, pv, fv, payType)), nil
}

// fnFV implements FV(rate, nper, pmt, [pv], [type]).
func fnFV(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	rate, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	nper, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pmt, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	pv := 0.0
	if len(args) >= 4 {
		pv, e = CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
	}
	payType := 0
	if len(args) == 5 {
		pt, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		if pt != 0 {
			payType = 1
		}
	}
	return NumberVal(fvCore(rate, nper, pmt, pv, payType)), nil
}

// fnPV implements PV(rate, nper, pmt, [fv], [type]).
func fnPV(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	rate, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	nper, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pmt, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	if nper == 0 && rate != 0 {
		return ErrorVal(ErrValNUM), nil
	}
	fv := 0.0
	if len(args) >= 4 {
		fv, e = CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
	}
	payType := 0
	if len(args) == 5 {
		pt, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		if pt != 0 {
			payType = 1
		}
	}
	if rate == 0 {
		return NumberVal(-fv - pmt*nper), nil
	}
	term := math.Pow(1+rate, nper)
	if payType == 1 {
		return NumberVal((-fv - pmt*(1+rate)*(term-1)/rate) / term), nil
	}
	return NumberVal((-fv - pmt*(term-1)/rate) / term), nil
}

// fnNPER implements NPER(rate, pmt, pv, [fv], [type]).
func fnNPER(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	rate, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	pmt, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pv, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	fv := 0.0
	if len(args) >= 4 {
		fv, e = CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
	}
	payType := 0
	if len(args) == 5 {
		pt, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		if pt != 0 {
			payType = 1
		}
	}
	if rate == 0 {
		if pmt == 0 {
			return ErrorVal(ErrValNUM), nil
		}
		return NumberVal(-(pv + fv) / pmt), nil
	}
	// NPER = log((pmt*(1+rate*type) - fv*rate) / (pmt*(1+rate*type) + pv*rate)) / log(1+rate)
	pmtAdj := pmt * (1 + rate*float64(payType))
	num := pmtAdj - fv*rate
	den := pmtAdj + pv*rate
	if den == 0 || num/den <= 0 || (1+rate) <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(math.Log(num/den) / math.Log(1+rate)), nil
}

// fnIPMT implements IPMT(rate, per, nper, pv, [fv], [type]).
func fnIPMT(args []Value) (Value, error) {
	if len(args) < 4 || len(args) > 6 {
		return ErrorVal(ErrValVALUE), nil
	}
	rate, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	per, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	nper, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	pv, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	if nper == 0 || per < 1 || per > nper {
		return ErrorVal(ErrValNUM), nil
	}
	fv := 0.0
	if len(args) >= 5 {
		fv, e = CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
	}
	payType := 0
	if len(args) == 6 {
		pt, e := CoerceNum(args[5])
		if e != nil {
			return *e, nil
		}
		if pt != 0 {
			payType = 1
		}
	}
	pmt := pmtCore(rate, nper, pv, fv, payType)
	if rate == 0 {
		return NumberVal(0), nil
	}
	// Interest portion = remaining balance at start of period * rate.
	// Balance after (per-1) payments:
	var ipmt float64
	if payType == 1 {
		if per == 1 {
			return NumberVal(0), nil
		}
		ipmt = (fvCore(rate, per-2, pmt, pv, 1) - pmt) * rate
	} else {
		ipmt = fvCore(rate, per-1, pmt, pv, 0) * rate
	}
	return NumberVal(ipmt), nil
}

// fnPPMT implements PPMT(rate, per, nper, pv, [fv], [type]).
func fnPPMT(args []Value) (Value, error) {
	if len(args) < 4 || len(args) > 6 {
		return ErrorVal(ErrValVALUE), nil
	}
	rate, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	per, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	nper, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	pv, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	if nper == 0 || per < 1 || per > nper {
		return ErrorVal(ErrValNUM), nil
	}
	fv := 0.0
	if len(args) >= 5 {
		fv, e = CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
	}
	payType := 0
	if len(args) == 6 {
		pt, e := CoerceNum(args[5])
		if e != nil {
			return *e, nil
		}
		if pt != 0 {
			payType = 1
		}
	}
	pmt := pmtCore(rate, nper, pv, fv, payType)
	// PPMT = PMT - IPMT
	if rate == 0 {
		return NumberVal(pmt), nil
	}
	var ipmt float64
	if payType == 1 {
		if per == 1 {
			return NumberVal(pmt), nil
		}
		ipmt = (fvCore(rate, per-2, pmt, pv, 1) - pmt) * rate
	} else {
		ipmt = fvCore(rate, per-1, pmt, pv, 0) * rate
	}
	return NumberVal(pmt - ipmt), nil
}

// fnNPV implements NPV(rate, value1, [value2], ...).
func fnNPV(args []Value) (Value, error) {
	if len(args) < 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	rate, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if rate == -1 {
		return ErrorVal(ErrValDIV0), nil
	}
	npv := 0.0
	period := 1
	for _, arg := range args[1:] {
		vals := flattenValues(arg)
		for _, v := range vals {
			if v.Type == ValueEmpty {
				continue
			}
			n, ev := CoerceNum(v)
			if ev != nil {
				return *ev, nil
			}
			npv += n / math.Pow(1+rate, float64(period))
			period++
		}
	}
	return NumberVal(npv), nil
}

// fnIRR implements IRR(values, [guess]).
// Uses Newton's method to find the rate where NPV = 0.
func fnIRR(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	values := flattenValues(args[0])
	var cashFlows []float64
	for _, v := range values {
		if v.Type == ValueEmpty {
			continue
		}
		n, e := CoerceNum(v)
		if e != nil {
			return *e, nil
		}
		cashFlows = append(cashFlows, n)
	}
	if len(cashFlows) < 2 {
		return ErrorVal(ErrValNUM), nil
	}
	// Must have at least one positive and one negative value.
	hasPos, hasNeg := false, false
	for _, cf := range cashFlows {
		if cf > 0 {
			hasPos = true
		}
		if cf < 0 {
			hasNeg = true
		}
	}
	if !hasPos || !hasNeg {
		return ErrorVal(ErrValNUM), nil
	}

	guess := 0.1
	if len(args) == 2 && args[1].Type != ValueEmpty {
		g, e := CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		guess = g
	}

	rate := guess
	for iter := 0; iter < 100; iter++ {
		npv := 0.0
		dnpv := 0.0
		for i, cf := range cashFlows {
			denom := math.Pow(1+rate, float64(i))
			if denom == 0 {
				return ErrorVal(ErrValNUM), nil
			}
			npv += cf / denom
			if i > 0 {
				dnpv -= float64(i) * cf / (denom * (1 + rate))
			}
		}
		if math.Abs(npv) < 1e-10 {
			return NumberVal(rate), nil
		}
		if dnpv == 0 {
			return ErrorVal(ErrValNUM), nil
		}
		newRate := rate - npv/dnpv
		if math.IsNaN(newRate) || math.IsInf(newRate, 0) {
			return ErrorVal(ErrValNUM), nil
		}
		if math.Abs(newRate-rate) < 1e-10 {
			return NumberVal(newRate), nil
		}
		rate = newRate
	}
	return ErrorVal(ErrValNUM), nil
}

// fnRATE implements RATE(nper, pmt, pv, [fv], [type], [guess]).
// Uses Newton's method.
func fnRATE(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 6 {
		return ErrorVal(ErrValVALUE), nil
	}
	nper, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	pmt, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pv, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	if nper <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	fv := 0.0
	if len(args) >= 4 {
		fv, e = CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
	}
	payType := 0
	if len(args) >= 5 {
		pt, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		if pt != 0 {
			payType = 1
		}
	}
	guess := 0.1
	if len(args) == 6 && args[5].Type != ValueEmpty {
		guess, e = CoerceNum(args[5])
		if e != nil {
			return *e, nil
		}
	}

	// When pmt=0 and fv=0 the equation reduces to pv*(1+r)^nper = 0.
	// If pv != 0 there is no rate that satisfies this, so return #NUM!.
	if pmt == 0 && fv == 0 && pv != 0 {
		return ErrorVal(ErrValNUM), nil
	}

	rate := guess
	for iter := 0; iter < 100; iter++ {
		if rate <= -1 {
			rate = -0.99
		}
		term := math.Pow(1+rate, nper)
		// f(rate) = pv*term + pmt*(1+rate*type)*(term-1)/rate + fv
		// When rate is very small, use the limit.
		var f, df float64
		if math.Abs(rate) < 1e-10 {
			f = pv + pmt*nper + fv
			df = pv*nper + pmt*nper*(nper-1)/2
		} else {
			pmtAdj := pmt * (1 + rate*float64(payType))
			f = pv*term + pmtAdj*(term-1)/rate + fv
			// Derivative with respect to rate.
			dterm := nper * math.Pow(1+rate, nper-1)
			df = pv*dterm + pmt*float64(payType)*(term-1)/rate + pmtAdj*(dterm*rate-(term-1))/(rate*rate)
		}
		if math.Abs(f) < 1e-10 {
			return NumberVal(rate), nil
		}
		if df == 0 {
			return ErrorVal(ErrValNUM), nil
		}
		newRate := rate - f/df
		if math.IsNaN(newRate) || math.IsInf(newRate, 0) {
			return ErrorVal(ErrValNUM), nil
		}
		if math.Abs(newRate-rate) < 1e-10 {
			return NumberVal(newRate), nil
		}
		rate = newRate
	}
	return ErrorVal(ErrValNUM), nil
}

// fnDB implements DB(cost, salvage, life, period, [month]).
// Returns the depreciation of an asset for a given period using the
// fixed-declining balance method.
func fnDB(args []Value) (Value, error) {
	if len(args) < 4 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	cost, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	salvage, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	life, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	period, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	month := 12.0
	if len(args) == 5 {
		month, e = CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
	}

	// Validate inputs.
	if cost < 0 || salvage < 0 || life <= 0 || period < 1 || month < 1 || month > 12 {
		return ErrorVal(ErrValNUM), nil
	}
	// period must be <= life, or life+1 when month < 12.
	maxPeriod := life
	if month < 12 {
		maxPeriod = life + 1
	}
	if period > maxPeriod {
		return ErrorVal(ErrValNUM), nil
	}

	// Special case: cost is 0, depreciation is always 0.
	if cost == 0 {
		return NumberVal(0), nil
	}

	// If salvage >= cost, no depreciation.
	if salvage >= cost {
		return NumberVal(0), nil
	}

	// rate = 1 - ((salvage/cost)^(1/life)), rounded to 3 decimal places.
	rate := 1 - math.Pow(salvage/cost, 1.0/life)
	rate = math.Round(rate*1000) / 1000

	// Compute depreciation for each period up to the requested one.
	totalDepreciation := 0.0
	var dep float64
	for p := 1.0; p <= period; p++ {
		if p == 1 {
			// First period: cost * rate * month / 12
			dep = cost * rate * month / 12
		} else if p == life+1 {
			// Last fractional period (only when month < 12):
			// (cost - totalDepreciation) * rate * (12 - month) / 12
			dep = (cost - totalDepreciation) * rate * (12 - month) / 12
		} else {
			// Intermediate periods:
			// (cost - totalDepreciation) * rate
			dep = (cost - totalDepreciation) * rate
		}
		totalDepreciation += dep
	}

	return NumberVal(dep), nil
}

// fnDDB implements DDB(cost, salvage, life, period, [factor]).
// Returns the depreciation of an asset for a given period using the
// double-declining balance method or another specified factor.
func fnDDB(args []Value) (Value, error) {
	if len(args) < 4 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	cost, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	salvage, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	life, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	period, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	factor := 2.0
	if len(args) == 5 {
		factor, e = CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
	}

	// All arguments must be positive; life and period must be > 0.
	if cost < 0 || salvage < 0 || life <= 0 || period <= 0 || factor <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if period > life {
		return ErrorVal(ErrValNUM), nil
	}

	// Special case: cost is 0, depreciation is always 0.
	if cost == 0 {
		return NumberVal(0), nil
	}

	// If salvage >= cost, no depreciation.
	if salvage >= cost {
		return NumberVal(0), nil
	}

	// Compute depreciation for each period up to the requested one.
	// DDB formula: dep = min(bookValue * (factor/life), bookValue - salvage)
	totalDepreciation := 0.0
	var dep float64
	// Period can be fractional in the sense that we iterate integer periods
	// up to math.Floor(period), but Excel DDB treats period as integer-based.
	// We iterate through whole periods.
	intPeriod := int(math.Floor(period))
	for p := 1; p <= intPeriod; p++ {
		bookValue := cost - totalDepreciation
		dep = bookValue * (factor / life)
		// Cannot depreciate below salvage value.
		if bookValue-dep < salvage {
			dep = bookValue - salvage
		}
		if dep < 0 {
			dep = 0
		}
		totalDepreciation += dep
	}

	// Handle fractional period: if period has a fractional part,
	// compute pro-rata depreciation for the remaining fraction.
	frac := period - float64(intPeriod)
	if frac > 0 {
		bookValue := cost - totalDepreciation
		fullDep := bookValue * (factor / life)
		if bookValue-fullDep < salvage {
			fullDep = bookValue - salvage
		}
		if fullDep < 0 {
			fullDep = 0
		}
		dep = fullDep * frac
		totalDepreciation += dep
	}

	return NumberVal(dep), nil
}

// fnSLN implements SLN(cost, salvage, life).
func fnSLN(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	cost, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	salvage, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	life, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	if life == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal((cost - salvage) / life), nil
}

// fnXNPV implements XNPV(rate, values, dates).
func fnXNPV(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	rate, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if rate <= -1 {
		return ErrorVal(ErrValNUM), nil
	}
	values := flattenValues(args[1])
	dates := flattenValues(args[2])
	if len(values) != len(dates) || len(values) == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	d0, e := coerceDateNum(dates[0])
	if e != nil {
		return *e, nil
	}
	xnpv := 0.0
	for i := range values {
		v, ev := CoerceNum(values[i])
		if ev != nil {
			return *ev, nil
		}
		di, ed := coerceDateNum(dates[i])
		if ed != nil {
			return *ed, nil
		}
		years := (di - d0) / 365.0
		denom := math.Pow(1+rate, years)
		if denom == 0 || math.IsInf(denom, 0) {
			return ErrorVal(ErrValNUM), nil
		}
		xnpv += v / denom
	}
	return NumberVal(xnpv), nil
}

// coerceDateNum converts a Value to a float64, treating date strings as Excel
// serial numbers. It first tries CoerceNum, and if that fails on a string
// value, it attempts to parse the string as a date using common date formats.
func coerceDateNum(v Value) (float64, *Value) {
	n, e := CoerceNum(v)
	if e == nil {
		return n, nil
	}
	// If the value is a string that couldn't be parsed as a number, try date parsing.
	if v.Type == ValueString {
		text := strings.TrimSpace(v.Str)
		layouts := []string{
			"1/2/2006",
			"01/02/2006",
			"2-Jan-2006",
			"02-Jan-2006",
			"2006/01/02",
			"2006-01-02",
			"January 2, 2006",
		}
		for _, layout := range layouts {
			t, err := time.Parse(layout, text)
			if err == nil {
				return math.Floor(TimeToExcelSerial(t)), nil
			}
		}
	}
	return 0, e
}

// fnDOLLARDE implements DOLLARDE(fractional_dollar, fraction).
// Converts a dollar price expressed as a fraction into a decimal number.
func fnDOLLARDE(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	fractionalDollar, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	fraction, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	// Truncate fraction to integer.
	fraction = math.Trunc(fraction)
	if fraction < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if fraction < 1 {
		return ErrorVal(ErrValDIV0), nil
	}

	// Determine sign and work with absolute value.
	sign := 1.0
	if fractionalDollar < 0 {
		sign = -1.0
		fractionalDollar = -fractionalDollar
	}

	intPart := math.Floor(fractionalDollar)
	fracPart := fractionalDollar - intPart

	// The number of digits in fraction determines how many decimal digits to extract.
	// e.g. fraction=16 has 2 digits, so we extract 2 decimal digits from fracPart.
	nDigits := len(fmt.Sprintf("%d", int(fraction)))
	divisor := math.Pow(10, float64(nDigits))

	// Extract the fractional portion: fracPart * divisor gives the numerator.
	numerator := fracPart * divisor
	result := intPart + numerator/fraction

	return NumberVal(sign * result), nil
}

// fnDOLLARFR implements DOLLARFR(decimal_dollar, fraction).
// Converts a decimal dollar price into a fractional dollar number.
func fnDOLLARFR(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	decimalDollar, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	fraction, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	// Truncate fraction to integer.
	fraction = math.Trunc(fraction)
	if fraction < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if fraction < 1 {
		return ErrorVal(ErrValDIV0), nil
	}

	// Determine sign and work with absolute value.
	sign := 1.0
	if decimalDollar < 0 {
		sign = -1.0
		decimalDollar = -decimalDollar
	}

	intPart := math.Floor(decimalDollar)
	fracPart := decimalDollar - intPart

	// The number of digits in fraction determines positioning.
	nDigits := len(fmt.Sprintf("%d", int(fraction)))
	divisor := math.Pow(10, float64(nDigits))

	// Multiply fractional part by fraction, then place in decimal position.
	numerator := fracPart * fraction
	result := intPart + numerator/divisor

	return NumberVal(sign * result), nil
}

// fnEFFECT implements EFFECT(nominal_rate, npery).
// Returns the effective annual interest rate.
func fnEFFECT(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	nominalRate, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	npery, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	npery = math.Trunc(npery)
	if nominalRate <= 0 || npery < 1 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(math.Pow(1+nominalRate/npery, npery) - 1), nil
}

// fnNOMINAL implements NOMINAL(effect_rate, npery).
// Returns the nominal annual interest rate.
func fnNOMINAL(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	effectRate, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	npery, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	npery = math.Trunc(npery)
	if effectRate <= 0 || npery < 1 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(npery * (math.Pow(1+effectRate, 1/npery) - 1)), nil
}

// fnCumipmt implements CUMIPMT(rate, nper, pv, start_period, end_period, type).
// Returns the cumulative interest paid on a loan between start_period and end_period.
func fnCumipmt(args []Value) (Value, error) {
	if len(args) != 6 {
		return ErrorVal(ErrValVALUE), nil
	}
	rate, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	nper, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pv, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	startPeriod, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	endPeriod, e := CoerceNum(args[4])
	if e != nil {
		return *e, nil
	}
	payTypeVal, e := CoerceNum(args[5])
	if e != nil {
		return *e, nil
	}

	// Truncate periods to integers.
	startPeriod = math.Floor(startPeriod)
	endPeriod = math.Floor(endPeriod)

	// Validate inputs.
	if rate <= 0 || nper <= 0 || pv <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if startPeriod < 1 || endPeriod < startPeriod {
		return ErrorVal(ErrValNUM), nil
	}
	if endPeriod > nper {
		return ErrorVal(ErrValNUM), nil
	}
	payType := 0
	if payTypeVal == 1 {
		payType = 1
	} else if payTypeVal != 0 {
		return ErrorVal(ErrValNUM), nil
	}

	pmt := pmtCore(rate, nper, pv, 0, payType)
	cumInterest := 0.0

	for i := int(startPeriod); i <= int(endPeriod); i++ {
		var ipmt float64
		if payType == 1 {
			if i == 1 {
				ipmt = 0
			} else {
				ipmt = (fvCore(rate, float64(i-2), pmt, pv, 1) - pmt) * rate
			}
		} else {
			ipmt = fvCore(rate, float64(i-1), pmt, pv, 0) * rate
		}
		cumInterest += ipmt
	}

	return NumberVal(cumInterest), nil
}

// fnCumprinc implements CUMPRINC(rate, nper, pv, start_period, end_period, type).
// Returns the cumulative principal paid on a loan between start_period and end_period.
func fnCumprinc(args []Value) (Value, error) {
	if len(args) != 6 {
		return ErrorVal(ErrValVALUE), nil
	}
	rate, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	nper, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pv, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	startPeriod, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	endPeriod, e := CoerceNum(args[4])
	if e != nil {
		return *e, nil
	}
	payTypeVal, e := CoerceNum(args[5])
	if e != nil {
		return *e, nil
	}

	// Truncate periods to integers.
	startPeriod = math.Floor(startPeriod)
	endPeriod = math.Floor(endPeriod)

	// Validate inputs.
	if rate <= 0 || nper <= 0 || pv <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if startPeriod < 1 || endPeriod < startPeriod {
		return ErrorVal(ErrValNUM), nil
	}
	if endPeriod > nper {
		return ErrorVal(ErrValNUM), nil
	}
	payType := 0
	if payTypeVal == 1 {
		payType = 1
	} else if payTypeVal != 0 {
		return ErrorVal(ErrValNUM), nil
	}

	pmt := pmtCore(rate, nper, pv, 0, payType)
	cumPrincipal := 0.0

	for i := int(startPeriod); i <= int(endPeriod); i++ {
		var ipmt float64
		if payType == 1 {
			if i == 1 {
				ipmt = 0
			} else {
				ipmt = (fvCore(rate, float64(i-2), pmt, pv, 1) - pmt) * rate
			}
		} else {
			ipmt = fvCore(rate, float64(i-1), pmt, pv, 0) * rate
		}
		cumPrincipal += pmt - ipmt
	}

	return NumberVal(cumPrincipal), nil
}

// fnMirr implements MIRR(values, finance_rate, reinvest_rate).
// Returns the modified internal rate of return for a series of periodic cash flows.
func fnMirr(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	values := flattenValues(args[0])
	financeRate, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	reinvestRate, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	var cashFlows []float64
	for _, v := range values {
		if v.Type == ValueEmpty {
			continue
		}
		n, ev := CoerceNum(v)
		if ev != nil {
			return *ev, nil
		}
		cashFlows = append(cashFlows, n)
	}

	n := len(cashFlows)
	if n < 2 {
		return ErrorVal(ErrValDIV0), nil
	}

	// Must have at least one positive and one negative value.
	hasPos, hasNeg := false, false
	for _, cf := range cashFlows {
		if cf > 0 {
			hasPos = true
		}
		if cf < 0 {
			hasNeg = true
		}
	}
	if !hasPos || !hasNeg {
		return ErrorVal(ErrValDIV0), nil
	}

	// FV of positive cash flows at reinvest_rate.
	fvPositive := 0.0
	for i, cf := range cashFlows {
		if cf > 0 {
			fvPositive += cf * math.Pow(1+reinvestRate, float64(n-i-1))
		}
	}

	// PV of negative cash flows at finance_rate.
	pvNegative := 0.0
	for i, cf := range cashFlows {
		if cf < 0 {
			pvNegative += cf / math.Pow(1+financeRate, float64(i))
		}
	}

	// MIRR = (FV_positive / (-PV_negative))^(1/(n-1)) - 1
	ratio := fvPositive / (-pvNegative)
	mirr := math.Pow(ratio, 1.0/float64(n-1)) - 1

	if math.IsNaN(mirr) || math.IsInf(mirr, 0) {
		return ErrorVal(ErrValDIV0), nil
	}

	return NumberVal(mirr), nil
}

// fnXIRR implements XIRR(values, dates, [guess]).
// Uses Newton's method to find the rate where XNPV = 0.
func fnXIRR(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	values := flattenValues(args[0])
	dates := flattenValues(args[1])
	if len(values) != len(dates) || len(values) < 2 {
		return ErrorVal(ErrValNUM), nil
	}

	var cashFlows []float64
	var dayOffsets []float64

	d0, e := coerceDateNum(dates[0])
	if e != nil {
		return *e, nil
	}

	hasPos, hasNeg := false, false
	for i := range values {
		v, ev := CoerceNum(values[i])
		if ev != nil {
			return *ev, nil
		}
		di, ed := coerceDateNum(dates[i])
		if ed != nil {
			return *ed, nil
		}
		cashFlows = append(cashFlows, v)
		dayOffsets = append(dayOffsets, (di-d0)/365.0)
		if v > 0 {
			hasPos = true
		}
		if v < 0 {
			hasNeg = true
		}
	}
	if !hasPos || !hasNeg {
		return ErrorVal(ErrValNUM), nil
	}

	guess := 0.1
	if len(args) == 3 && args[2].Type != ValueEmpty {
		g, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		guess = g
	}

	rate := guess
	for iter := 0; iter < 300; iter++ {
		if rate <= -1 {
			rate = -0.99
		}
		xnpv := 0.0
		dxnpv := 0.0
		for i, cf := range cashFlows {
			t := dayOffsets[i]
			denom := math.Pow(1+rate, t)
			if denom == 0 || math.IsInf(denom, 0) {
				return ErrorVal(ErrValNUM), nil
			}
			xnpv += cf / denom
			if t != 0 {
				dxnpv -= t * cf / (denom * (1 + rate))
			}
		}
		if math.Abs(xnpv) < 1e-14 {
			return NumberVal(rate), nil
		}
		if dxnpv == 0 {
			return ErrorVal(ErrValNUM), nil
		}
		newRate := rate - xnpv/dxnpv
		if math.IsNaN(newRate) || math.IsInf(newRate, 0) {
			return ErrorVal(ErrValNUM), nil
		}
		if math.Abs(newRate-rate) < 1e-12 {
			return NumberVal(newRate), nil
		}
		rate = newRate
	}
	return ErrorVal(ErrValNUM), nil
}

// fnPduration implements PDURATION(rate, pv, fv).
// Returns the number of periods required for an investment to reach a specified value.
func fnPduration(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	rate, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	pv, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	fv, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	if rate <= 0 || pv <= 0 || fv <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal((math.Log(fv) - math.Log(pv)) / math.Log(1+rate)), nil
}

// fnRri implements RRI(nper, pv, fv).
// Returns an equivalent interest rate for the growth of an investment.
func fnRri(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	nper, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	pv, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	fv, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	if nper <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if pv == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	ratio := fv / pv
	if ratio < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(math.Pow(ratio, 1.0/nper) - 1), nil
}
