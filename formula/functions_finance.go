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
	Register("SYD", NoCtx(fnSYD))
	Register("VDB", NoCtx(fnVdb))
	Register("ISPMT", NoCtx(fnIspmt))
	Register("TBILLPRICE", NoCtx(fnTbillPrice))
	Register("TBILLYIELD", NoCtx(fnTbillYield))
	Register("TBILLEQ", NoCtx(fnTbillEq))
	Register("DISC", NoCtx(fnDisc))
	Register("INTRATE", NoCtx(fnIntrate))
	Register("RECEIVED", NoCtx(fnReceived))
	Register("ACCRINT", NoCtx(fnAccrint))
	Register("ACCRINTM", NoCtx(fnAccrintm))
	Register("PRICEDISC", NoCtx(fnPricedisc))
	Register("YIELDDISC", NoCtx(fnYielddisc))
	Register("PRICEMAT", NoCtx(fnPricemat))
	Register("YIELDMAT", NoCtx(fnYieldmat))
	Register("COUPDAYBS", NoCtx(fnCoupdaybs))
	Register("COUPDAYS", NoCtx(fnCoupdays))
	Register("COUPDAYSNC", NoCtx(fnCoupdaysnc))
	Register("COUPNCD", NoCtx(fnCoupncd))
	Register("COUPNUM", NoCtx(fnCoupnum))
	Register("COUPPCD", NoCtx(fnCouppcd))
	Register("DURATION", NoCtx(fnDuration))
	Register("MDURATION", NoCtx(fnMduration))
	Register("PRICE", NoCtx(fnPrice))
	Register("YIELD", NoCtx(fnYield))
	Register("FVSCHEDULE", NoCtx(fnFVSchedule))
	Register("AMORDEGRC", NoCtx(fnAmordegrc))
	Register("AMORLINC", NoCtx(fnAmorlinc))
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
	if term == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	result := 0.0
	if payType == 1 {
		result = (-fv - pmt*(1+rate)*(term-1)/rate) / term
	} else {
		result = (-fv - pmt*(term-1)/rate) / term
	}
	if math.IsInf(result, 0) || math.IsNaN(result) {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(result), nil
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
	// up to math.Floor(period), but DDB treats period as integer-based.
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

// coerceDateNum converts a Value to a float64, treating date strings as date
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
				return math.Floor(TimeToSerial(t)), nil
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

	// Must have at least one negative value (otherwise PV of negatives is 0 → division by zero).
	// All-negative is valid (returns -1 because FV of positives is 0).
	hasNeg := false
	for _, cf := range cashFlows {
		if cf < 0 {
			hasNeg = true
			break
		}
	}
	if !hasNeg {
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
	result := math.Pow(ratio, 1.0/nper)
	if math.IsNaN(result) {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(result - 1), nil
}

// fnSYD implements SYD(cost, salvage, life, per).
// Returns the sum-of-years-digits depreciation of an asset for a given period.
func fnSYD(args []Value) (Value, error) {
	if len(args) != 4 {
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
	per, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}

	// Truncate life and per to integers.
	life = math.Trunc(life)
	per = math.Trunc(per)

	if life <= 0 || per < 1 || per > life {
		return ErrorVal(ErrValNUM), nil
	}

	// SYD = (cost - salvage) * (life - per + 1) / (life * (life + 1) / 2)
	return NumberVal((cost - salvage) * (life - per + 1) / (life * (life + 1) / 2)), nil
}

// vdbCalcOneperiod calculates the depreciation for a single period using
// DDB with optional switch to straight-line.
func vdbCalcOnePeriod(cost, salvage, life, period, factor float64, noSwitch bool) float64 {
	// Compute book value at the start of this period by accumulating
	// depreciation for all prior periods.
	bookValue := cost
	for i := 0.0; i < period; i++ {
		ddb := bookValue * (factor / life)
		if !noSwitch {
			sl := 0.0
			remaining := life - i
			if remaining > 0 {
				sl = (bookValue - salvage) / remaining
			}
			if sl > ddb {
				ddb = sl
			}
		}
		if bookValue-ddb < salvage {
			ddb = bookValue - salvage
		}
		if ddb < 0 {
			ddb = 0
		}
		bookValue -= ddb
	}

	// Now compute depreciation for the requested period.
	ddb := bookValue * (factor / life)
	if !noSwitch {
		sl := 0.0
		remaining := life - period
		if remaining > 0 {
			sl = (bookValue - salvage) / remaining
		}
		if sl > ddb {
			ddb = sl
		}
	}
	if bookValue-ddb < salvage {
		ddb = bookValue - salvage
	}
	if ddb < 0 {
		ddb = 0
	}
	return ddb
}

// fnVdb implements VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch]).
// Returns the depreciation of an asset for any specified period using the
// variable declining balance method.
func fnVdb(args []Value) (Value, error) {
	if len(args) < 5 || len(args) > 7 {
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
	startPeriod, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	endPeriod, e := CoerceNum(args[4])
	if e != nil {
		return *e, nil
	}
	factor := 2.0
	if len(args) >= 6 {
		factor, e = CoerceNum(args[5])
		if e != nil {
			return *e, nil
		}
	}
	noSwitch := false
	if len(args) == 7 {
		ns, e := CoerceNum(args[6])
		if e != nil {
			// Try bool coercion.
			if args[6].Type == ValueBool {
				noSwitch = args[6].Bool
			} else {
				return *e, nil
			}
		} else {
			noSwitch = ns != 0
		}
	}

	// Validate inputs.
	if cost < 0 || salvage < 0 || life <= 0 || startPeriod < 0 || endPeriod < startPeriod || factor <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if endPeriod > life {
		return ErrorVal(ErrValNUM), nil
	}

	// Special cases.
	if cost == 0 || salvage >= cost {
		return NumberVal(0), nil
	}

	// Accumulate depreciation across the period range [startPeriod, endPeriod].
	// VDB uses 0-based period intervals: period i covers interval [i, i+1).
	// VDB(cost,salvage,life,0,1) = depreciation for period 0.
	// VDB(cost,salvage,life,0,5) = total depreciation for periods 0-4.
	// Fractional start/end are prorated.
	totalDep := 0.0

	// Walk through each integer period that overlaps [startPeriod, endPeriod].
	firstPeriod := int(math.Floor(startPeriod))
	lastPeriod := int(math.Ceil(endPeriod)) - 1
	if endPeriod == math.Floor(endPeriod) {
		lastPeriod = int(endPeriod) - 1
	}

	for p := firstPeriod; p <= lastPeriod; p++ {
		dep := vdbCalcOnePeriod(cost, salvage, life, float64(p), factor, noSwitch)

		// Determine the fraction of this period that falls within [startPeriod, endPeriod].
		pStart := float64(p)
		pEnd := float64(p + 1)
		if startPeriod > pStart {
			pStart = startPeriod
		}
		if endPeriod < pEnd {
			pEnd = endPeriod
		}
		frac := pEnd - pStart
		totalDep += dep * frac
	}

	return NumberVal(totalDep), nil
}

// fnIspmt implements ISPMT(rate, per, nper, pv).
// Calculates the interest paid (or received) for the specified period of a
// loan (or investment) with even principal payments.
// Formula: pv * rate * (per/nper - 1)
func fnIspmt(args []Value) (Value, error) {
	if len(args) != 4 {
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
	if nper == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(pv * rate * (per/nper - 1)), nil
}

// tbillValidate is a shared helper for TBILLPRICE, TBILLYIELD, and TBILLEQ.
// It parses settlement and maturity as date serial numbers (truncated to int),
// validates that settlement < maturity (or settlement <= maturity for strictLess=false),
// that maturity is not more than one calendar year after settlement, and returns
// DSM (days from settlement to maturity).
func tbillValidate(args []Value, strictLess bool) (dsm float64, thirdArg float64, ev *Value) {
	if len(args) != 3 {
		ev := ErrorVal(ErrValVALUE)
		return 0, 0, &ev
	}
	settlementRaw, e := CoerceNum(args[0])
	if e != nil {
		return 0, 0, e
	}
	maturityRaw, e := CoerceNum(args[1])
	if e != nil {
		return 0, 0, e
	}
	third, e := CoerceNum(args[2])
	if e != nil {
		return 0, 0, e
	}

	// Truncate to integers (date serial numbers).
	settlement := math.Trunc(settlementRaw)
	maturity := math.Trunc(maturityRaw)

	// Validate third argument (discount or price) > 0.
	if third <= 0 {
		ev := ErrorVal(ErrValNUM)
		return 0, 0, &ev
	}

	// Validate settlement vs maturity.
	if strictLess {
		// TBILLPRICE and TBILLYIELD: settlement must be strictly less than maturity.
		if settlement >= maturity {
			ev := ErrorVal(ErrValNUM)
			return 0, 0, &ev
		}
	} else {
		// TBILLEQ: settlement must not be greater than maturity (settlement == maturity is allowed... but yields 0 DSM).
		if settlement > maturity {
			ev := ErrorVal(ErrValNUM)
			return 0, 0, &ev
		}
	}

	// Convert serial numbers to time.Time and check "more than one year" rule.
	settlementTime := SerialToTime(settlement)
	maturityTime := SerialToTime(maturity)

	// Maturity must not be more than one calendar year after settlement.
	oneYearLater := settlementTime.AddDate(1, 0, 0)
	if maturityTime.After(oneYearLater) {
		ev := ErrorVal(ErrValNUM)
		return 0, 0, &ev
	}

	// DSM = actual number of days between settlement and maturity.
	dsm = maturity - settlement

	return dsm, third, nil
}

// fnTbillPrice implements TBILLPRICE(settlement, maturity, discount).
// Returns the price per $100 face value for a Treasury bill.
// Formula: TBILLPRICE = 100 * (1 - discount * DSM / 360)
func fnTbillPrice(args []Value) (Value, error) {
	dsm, discount, ev := tbillValidate(args, true)
	if ev != nil {
		return *ev, nil
	}
	price := 100.0 * (1.0 - discount*dsm/360.0)
	return NumberVal(price), nil
}

// fnTbillYield implements TBILLYIELD(settlement, maturity, pr).
// Returns the yield for a Treasury bill.
// Formula: TBILLYIELD = ((100 - pr) / pr) * (360 / DSM)
func fnTbillYield(args []Value) (Value, error) {
	dsm, pr, ev := tbillValidate(args, true)
	if ev != nil {
		return *ev, nil
	}
	if dsm == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	yield := ((100.0 - pr) / pr) * (360.0 / dsm)
	return NumberVal(yield), nil
}

// fnTbillEq implements TBILLEQ(settlement, maturity, discount).
// Returns the bond-equivalent yield for a Treasury bill.
//
// For DSM <= 182 (half-year or less):
//
//	TBILLEQ = (365 * discount) / (360 - discount * DSM)
//
// For DSM > 182 (more than half-year), uses a semi-annual
// compounding formula derived from the US Treasury's investment rate
// calculation:
//
//	price = 100 * (1 - discount * DSM / 360)
//	b     = DSM / 365
//	TBILLEQ = (-2*b + 2*sqrt(b*b - (2*b - 1)*(1 - 100/price))) / (2*b - 1)
func fnTbillEq(args []Value) (Value, error) {
	dsm, discount, ev := tbillValidate(args, false)
	if ev != nil {
		return *ev, nil
	}

	if dsm <= 182 {
		// Short-term: simple bond-equivalent yield.
		denom := 360.0 - discount*dsm
		if denom == 0 {
			return ErrorVal(ErrValDIV0), nil
		}
		return NumberVal((365.0 * discount) / denom), nil
	}

	// Long-term (DSM > 182): semi-annual compounding formula.
	price := 100.0 * (1.0 - discount*dsm/360.0)
	if price <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	b := dsm / 365.0
	term := 2.0*b - 1.0
	discriminant := b*b - term*(1.0-100.0/price)
	if discriminant < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	result := (-2.0*b + 2.0*math.Sqrt(discriminant)) / term
	return NumberVal(result), nil
}

// dayCountBasis computes the number of days between two dates (DSM) and the
// number of days in a year (B) according to the specified day count basis.
// Both settlement and maturity should be truncated serial numbers.
// Basis values: 0=US 30/360, 1=Actual/actual, 2=Actual/360, 3=Actual/365, 4=European 30/360.
func dayCountBasis(settlement, maturity float64, basis int) (dsm, bYear float64) {
	st := SerialToTime(settlement)
	mt := SerialToTime(maturity)
	sy, sm, sd := st.Year(), int(st.Month()), st.Day()
	ey, em, ed := mt.Year(), int(mt.Month()), mt.Day()

	switch basis {
	case 0: // US (NASD) 30/360
		dsm = days360Calc(sy, sm, sd, ey, em, ed, false)
		bYear = 360
	case 1: // Actual/actual
		dsm = maturity - settlement
		if sy == ey {
			// Same calendar year: use that year's length.
			bYear = 365
			if isLeapYear(sy) {
				bYear = 366
			}
		} else if ey-sy == 1 {
			// Adjacent years: check whether Feb 29 falls within the range.
			bYear = 365
			for y := sy; y <= ey; y++ {
				if isLeapYear(y) {
					feb29 := TimeToSerial(time.Date(y, 2, 29, 0, 0, 0, 0, time.UTC))
					if feb29 >= settlement && feb29 <= maturity {
						bYear = 366
						break
					}
				}
			}
		} else {
			// Multi-year period: average days per year across the full range.
			totalYearDays := 0.0
			for y := sy; y <= ey; y++ {
				if isLeapYear(y) {
					totalYearDays += 366
				} else {
					totalYearDays += 365
				}
			}
			bYear = totalYearDays / float64(ey-sy+1)
		}
	case 2: // Actual/360
		dsm = maturity - settlement
		bYear = 360
	case 3: // Actual/365
		dsm = maturity - settlement
		bYear = 365
	case 4: // European 30/360
		dsm = days360Calc(sy, sm, sd, ey, em, ed, true)
		bYear = 360
	}
	return dsm, bYear
}

// fnDisc implements DISC(settlement, maturity, pr, redemption, [basis]).
// Returns the discount rate for a security.
// Formula: DISC = (redemption - pr) / redemption * (B / DSM)
func fnDisc(args []Value) (Value, error) {
	if len(args) < 4 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	settlementRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	maturityRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pr, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	redemption, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	basis := 0
	if len(args) == 5 {
		b, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		basis = int(b)
	}

	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if pr <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if redemption <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	settlement := math.Trunc(settlementRaw)
	maturity := math.Trunc(maturityRaw)
	if settlement >= maturity {
		return ErrorVal(ErrValNUM), nil
	}

	dsm, bYear := dayCountBasis(settlement, maturity, basis)
	if dsm == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	disc := (redemption - pr) / redemption * (bYear / dsm)
	return NumberVal(disc), nil
}

// fnIntrate implements INTRATE(settlement, maturity, investment, redemption, [basis]).
// Returns the interest rate for a fully invested security.
// Formula: INTRATE = (redemption - investment) / investment * (B / DIM)
func fnIntrate(args []Value) (Value, error) {
	if len(args) < 4 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	settlementRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	maturityRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	investment, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	redemption, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	basis := 0
	if len(args) == 5 {
		b, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		basis = int(b)
	}

	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if investment <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if redemption <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	settlement := math.Trunc(settlementRaw)
	maturity := math.Trunc(maturityRaw)
	if settlement >= maturity {
		return ErrorVal(ErrValNUM), nil
	}

	dim, bYear := dayCountBasis(settlement, maturity, basis)
	if dim == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	rate := (redemption - investment) / investment * (bYear / dim)
	return NumberVal(rate), nil
}

// fnReceived implements RECEIVED(settlement, maturity, investment, discount, [basis]).
// Returns the amount received at maturity for a fully invested security.
// Formula: RECEIVED = investment / (1 - discount * DIM / B)
func fnReceived(args []Value) (Value, error) {
	if len(args) < 4 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	settlementRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	maturityRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	investment, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	discount, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	basis := 0
	if len(args) == 5 {
		b, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		basis = int(b)
	}

	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if investment <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if discount <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	settlement := math.Trunc(settlementRaw)
	maturity := math.Trunc(maturityRaw)
	if settlement >= maturity {
		return ErrorVal(ErrValNUM), nil
	}

	dim, bYear := dayCountBasis(settlement, maturity, basis)
	denom := 1.0 - discount*dim/bYear
	if denom == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	received := investment / denom
	return NumberVal(received), nil
}

// fnAccrintm implements ACCRINTM(issue, settlement, rate, par, [basis]).
// Returns the accrued interest for a security that pays interest at maturity.
// Formula: ACCRINTM = par * rate * A / D
// where A = accrued days from issue to settlement, D = annual basis days.
// fnAccrint implements ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method]).
// Returns the accrued interest for a security that pays periodic interest.
func fnAccrint(args []Value) (Value, error) {
	if len(args) < 6 || len(args) > 8 {
		return ErrorVal(ErrValVALUE), nil
	}

	issueRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	firstInterestRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	settlementRaw, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	rate, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	par, e := CoerceNum(args[4])
	if e != nil {
		return *e, nil
	}
	freqRaw, e := CoerceNum(args[5])
	if e != nil {
		return *e, nil
	}

	basis := 0
	if len(args) >= 7 {
		b, e := CoerceNum(args[6])
		if e != nil {
			return *e, nil
		}
		basis = int(b)
	}

	calcMethod := true
	if len(args) == 8 {
		cm, e := CoerceNum(args[7])
		if e != nil {
			return *e, nil
		}
		calcMethod = cm != 0
	}

	issue := math.Trunc(issueRaw)
	firstInterest := math.Trunc(firstInterestRaw)
	settlement := math.Trunc(settlementRaw)
	freq := int(math.Trunc(freqRaw))

	// Validate inputs.
	if rate <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if par <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if freq != 1 && freq != 2 && freq != 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if issue >= settlement {
		return ErrorVal(ErrValNUM), nil
	}

	fiTime := SerialToTime(firstInterest)
	monthStep := 12 / freq
	fiD := fiTime.Day()

	// Determine the accrual start date.
	// calc_method=TRUE (default): accrue from issue to settlement.
	// calc_method=FALSE: accrue only from the previous coupon date (on or after issue) to settlement.
	startDate := issue

	if !calcMethod {
		// Find the most recent quasi-coupon date on or before settlement.
		settTime := SerialToTime(settlement)
		pcdOfSettlement := couppcd(settTime, fiTime, freq)
		pcdSer := TimeToSerial(pcdOfSettlement)
		// Use the later of issue and pcdOfSettlement.
		if pcdSer > issue {
			startDate = pcdSer
		}
	}

	// Find the quasi-coupon date on or before startDate.
	startTime := SerialToTime(startDate)
	pcd := couppcd(startTime, fiTime, freq)
	pcdSerial := TimeToSerial(pcd)

	// If pcd is after startDate (edge case), step back further.
	for pcdSerial > startDate {
		pcdY, pcdM := pcd.Year(), int(pcd.Month())
		prevMonth := pcdM - monthStep
		prevYear := pcdY
		for prevMonth <= 0 {
			prevMonth += 12
			prevYear--
		}
		pcd = clampDate(prevYear, prevMonth, fiD)
		pcdSerial = TimeToSerial(pcd)
	}

	// Walk forward through quasi-coupon periods, accumulating accrued interest.
	var result float64

	for pcdSerial < settlement {
		// Compute the next coupon date.
		ncdMonth := int(pcd.Month()) + monthStep
		ncdYear := pcd.Year()
		for ncdMonth > 12 {
			ncdMonth -= 12
			ncdYear++
		}
		ncd := clampDate(ncdYear, ncdMonth, fiD)
		ncdSerial := TimeToSerial(ncd)

		// Determine the portion of this quasi-coupon period that falls within [startDate, settlement].
		periodStart := pcdSerial
		if periodStart < startDate {
			periodStart = startDate
		}
		periodEnd := ncdSerial
		if periodEnd > settlement {
			periodEnd = settlement
		}

		if periodEnd > periodStart {
			// Compute A_i (accrued days in this quasi-coupon period).
			var ai float64
			psTime := SerialToTime(periodStart)
			peTime := SerialToTime(periodEnd)
			switch basis {
			case 0: // US 30/360
				ai = days360Calc(
					psTime.Year(), int(psTime.Month()), psTime.Day(),
					peTime.Year(), int(peTime.Month()), peTime.Day(),
					false,
				)
			case 4: // European 30/360
				ai = days360Calc(
					psTime.Year(), int(psTime.Month()), psTime.Day(),
					peTime.Year(), int(peTime.Month()), peTime.Day(),
					true,
				)
			default: // basis 1, 2, 3: actual days
				ai = periodEnd - periodStart
			}

			// Compute NL_i (normal length of quasi-coupon period).
			var nli float64
			switch basis {
			case 0, 4: // 30/360
				nli = 360.0 / float64(freq)
			case 1: // actual/actual
				nli = ncdSerial - pcdSerial
			case 2: // actual/360
				nli = 360.0 / float64(freq)
			case 3: // actual/365
				nli = 365.0 / float64(freq)
			}

			if nli > 0 {
				result += par * rate / float64(freq) * ai / nli
			}
		}

		// Move to the next quasi-coupon period.
		pcd = ncd
		pcdSerial = ncdSerial
	}

	return NumberVal(result), nil
}

func fnAccrintm(args []Value) (Value, error) {
	if len(args) < 4 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	issueRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	settlementRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	rate, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	par, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	basis := 0
	if len(args) == 5 {
		b, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		basis = int(b)
	}

	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if rate <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if par <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	issue := math.Trunc(issueRaw)
	settlement := math.Trunc(settlementRaw)
	if issue >= settlement {
		return ErrorVal(ErrValNUM), nil
	}

	a, d := dayCountBasis(issue, settlement, basis)
	if d == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	result := par * rate * a / d
	return NumberVal(result), nil
}

// fnPricedisc implements PRICEDISC(settlement, maturity, discount, redemption, [basis]).
// Returns the price per $100 face value of a discounted security.
// Formula: PRICEDISC = redemption - discount * (DSM / B) * redemption
func fnPricedisc(args []Value) (Value, error) {
	if len(args) < 4 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	settlementRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	maturityRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	discount, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	redemption, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	basis := 0
	if len(args) == 5 {
		b, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		basis = int(b)
	}

	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if discount <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if redemption <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	settlement := math.Trunc(settlementRaw)
	maturity := math.Trunc(maturityRaw)
	if settlement >= maturity {
		return ErrorVal(ErrValNUM), nil
	}

	dsm, bYear := dayCountBasis(settlement, maturity, basis)
	if bYear == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	price := redemption - discount*(dsm/bYear)*redemption
	return NumberVal(price), nil
}

// fnYielddisc implements YIELDDISC(settlement, maturity, pr, redemption, [basis]).
// Returns the annual yield for a discounted security.
// Formula: YIELDDISC = (redemption - pr) / pr * (B / DSM)
func fnYielddisc(args []Value) (Value, error) {
	if len(args) < 4 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	settlementRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	maturityRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pr, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	redemption, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	basis := 0
	if len(args) == 5 {
		b, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		basis = int(b)
	}

	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if pr <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if redemption <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	settlement := math.Trunc(settlementRaw)
	maturity := math.Trunc(maturityRaw)
	if settlement >= maturity {
		return ErrorVal(ErrValNUM), nil
	}

	dsm, bYear := dayCountBasis(settlement, maturity, basis)
	if dsm == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	yield := (redemption - pr) / pr * (bYear / dsm)
	return NumberVal(yield), nil
}

// fnPricemat implements PRICEMAT(settlement, maturity, issue, rate, yld, [basis]).
// Returns the price per $100 face value of a security that pays interest at maturity.
// Formula: PRICEMAT = 100*(1 + DIM/B*rate) / (1 + DSM/B*yld) - 100*A/B*rate
// where DSM = days from settlement to maturity, DIM = days from issue to maturity,
// A = days from issue to settlement, B = annual basis days.
func fnPricemat(args []Value) (Value, error) {
	if len(args) < 5 || len(args) > 6 {
		return ErrorVal(ErrValVALUE), nil
	}
	settlementRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	maturityRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	issueRaw, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	rate, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	yld, e := CoerceNum(args[4])
	if e != nil {
		return *e, nil
	}
	basis := 0
	if len(args) == 6 {
		b, e := CoerceNum(args[5])
		if e != nil {
			return *e, nil
		}
		basis = int(b)
	}

	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if rate < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if yld < 0 {
		return ErrorVal(ErrValNUM), nil
	}

	settlement := math.Trunc(settlementRaw)
	maturity := math.Trunc(maturityRaw)
	issue := math.Trunc(issueRaw)
	if settlement >= maturity {
		return ErrorVal(ErrValNUM), nil
	}

	dsm, _ := dayCountBasis(settlement, maturity, basis)
	dim, _ := dayCountBasis(issue, maturity, basis)
	a, bYear := dayCountBasis(issue, settlement, basis)

	if bYear == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	denom := 1 + dsm/bYear*yld
	if denom == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	price := 100*(1+dim/bYear*rate)/denom - 100*a/bYear*rate
	return NumberVal(price), nil
}

// fnYieldmat implements YIELDMAT(settlement, maturity, issue, rate, pr, [basis]).
// Returns the annual yield of a security that pays interest at maturity.
// Formula: YIELDMAT = ((1 + DIM/B*rate) / (pr/100 + A/B*rate) - 1) * (B / DSM)
// where DSM = days from settlement to maturity, DIM = days from issue to maturity,
// A = days from issue to settlement, B = annual basis days.
func fnYieldmat(args []Value) (Value, error) {
	if len(args) < 5 || len(args) > 6 {
		return ErrorVal(ErrValVALUE), nil
	}
	settlementRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	maturityRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	issueRaw, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	rate, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	pr, e := CoerceNum(args[4])
	if e != nil {
		return *e, nil
	}
	basis := 0
	if len(args) == 6 {
		b, e := CoerceNum(args[5])
		if e != nil {
			return *e, nil
		}
		basis = int(b)
	}

	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if rate < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if pr <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	settlement := math.Trunc(settlementRaw)
	maturity := math.Trunc(maturityRaw)
	issue := math.Trunc(issueRaw)
	if settlement >= maturity {
		return ErrorVal(ErrValNUM), nil
	}

	dsm, _ := dayCountBasis(settlement, maturity, basis)
	dim, _ := dayCountBasis(issue, maturity, basis)
	a, bYear := dayCountBasis(issue, settlement, basis)

	if bYear == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	if dsm == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	denom := pr/100 + a/bYear*rate
	if denom == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	yield := ((1+dim/bYear*rate)/denom - 1) * (bYear / dsm)
	return NumberVal(yield), nil
}

// ---------------------------------------------------------------------------
// Coupon period functions: COUPDAYBS, COUPDAYS, COUPDAYSNC, COUPNCD, COUPNUM, COUPPCD
// ---------------------------------------------------------------------------

// coupValidate parses and validates the common arguments shared by all COUP* functions.
// Signature: FUNC(settlement, maturity, frequency, [basis])
// Returns settlement serial, maturity serial, frequency, basis, and an optional error Value.
func coupValidate(args []Value) (float64, float64, int, int, *Value) {
	if len(args) < 3 || len(args) > 4 {
		ev := ErrorVal(ErrValVALUE)
		return 0, 0, 0, 0, &ev
	}
	settlementRaw, e := CoerceNum(args[0])
	if e != nil {
		return 0, 0, 0, 0, e
	}
	maturityRaw, e := CoerceNum(args[1])
	if e != nil {
		return 0, 0, 0, 0, e
	}
	freqRaw, e := CoerceNum(args[2])
	if e != nil {
		return 0, 0, 0, 0, e
	}
	basis := 0
	if len(args) == 4 {
		b, e := CoerceNum(args[3])
		if e != nil {
			return 0, 0, 0, 0, e
		}
		basis = int(b)
	}

	settlement := math.Trunc(settlementRaw)
	maturity := math.Trunc(maturityRaw)
	freq := int(math.Trunc(freqRaw))

	if settlement >= maturity {
		ev := ErrorVal(ErrValNUM)
		return 0, 0, 0, 0, &ev
	}
	if freq != 1 && freq != 2 && freq != 4 {
		ev := ErrorVal(ErrValNUM)
		return 0, 0, 0, 0, &ev
	}
	if basis < 0 || basis > 4 {
		ev := ErrorVal(ErrValNUM)
		return 0, 0, 0, 0, &ev
	}

	return settlement, maturity, freq, basis, nil
}

// couppcd computes the previous coupon date on or before settlement.
// Coupon dates form a schedule based on the maturity date's month and day,
// stepping backwards by 12/freq months at a time.
func couppcd(settlement, maturity time.Time, freq int) time.Time {
	monthStep := 12 / freq
	matY, matM, matD := maturity.Year(), int(maturity.Month()), maturity.Day()

	// Estimate how many periods back from maturity to reach settlement.
	settY, settM := settlement.Year(), int(settlement.Month())

	totalMonthsDiff := (matY-settY)*12 + (matM - settM)
	periodsBack := totalMonthsDiff / monthStep
	if periodsBack < 0 {
		periodsBack = 0
	}

	// Start from an estimated position and adjust.
	candidate := coupDateAtOffset(matY, matM, matD, -periodsBack*monthStep)

	// If candidate is after settlement, step back more.
	for candidate.After(settlement) {
		periodsBack++
		candidate = coupDateAtOffset(matY, matM, matD, -periodsBack*monthStep)
	}
	// If we can step forward and still be <= settlement, do so.
	for {
		next := coupDateAtOffset(matY, matM, matD, -(periodsBack-1)*monthStep)
		if next.After(settlement) {
			break
		}
		periodsBack--
		candidate = next
	}
	return candidate
}

// coupncd computes the next coupon date strictly after settlement.
func coupncd(settlement, maturity time.Time, freq int) time.Time {
	pcd := couppcd(settlement, maturity, freq)
	monthStep := 12 / freq
	matD := maturity.Day()

	// NCD is one period after PCD, clamped to the maturity day.
	ncdMonth := int(pcd.Month()) + monthStep
	ncdYear := pcd.Year()
	for ncdMonth > 12 {
		ncdMonth -= 12
		ncdYear++
	}
	ncd := clampDate(ncdYear, ncdMonth, matD)
	return ncd
}

// coupDateAtOffset returns a coupon date that is offsetMonths months from
// the maturity's year/month, using the maturity day clamped to the target month's length.
func coupDateAtOffset(matY, matM, matD, offsetMonths int) time.Time {
	totalMonths := (matY-1)*12 + (matM - 1) + offsetMonths
	year := totalMonths/12 + 1
	month := totalMonths%12 + 1
	if month <= 0 {
		month += 12
		year--
	}
	return clampDate(year, month, matD)
}

// clampDate creates a date with the given year, month, and day, clamping the
// day to the last day of the month if it exceeds the month's length.
func clampDate(year, month, day int) time.Time {
	// Find the last day of the target month.
	lastDay := daysInMonth(year, month)
	if day > lastDay {
		day = lastDay
	}
	return time.Date(year, time.Month(month), day, 0, 0, 0, 0, time.UTC)
}

// daysInMonth returns the number of days in the given month.
func daysInMonth(year, month int) int {
	// Use time.Date to go to the first of the next month, then subtract a day.
	return time.Date(year, time.Month(month+1), 0, 0, 0, 0, 0, time.UTC).Day()
}

// fnCOUPPCD implements COUPPCD(settlement, maturity, frequency, [basis]).
// Returns the previous coupon date before settlement as a serial date number.
func fnCouppcd(args []Value) (Value, error) {
	settlement, maturity, freq, _, ev := coupValidate(args)
	if ev != nil {
		return *ev, nil
	}
	st := SerialToTime(settlement)
	mt := SerialToTime(maturity)
	pcd := couppcd(st, mt, freq)
	return NumberVal(TimeToSerial(pcd)), nil
}

// fnCOUPNCD implements COUPNCD(settlement, maturity, frequency, [basis]).
// Returns the next coupon date after settlement as a serial date number.
func fnCoupncd(args []Value) (Value, error) {
	settlement, maturity, freq, _, ev := coupValidate(args)
	if ev != nil {
		return *ev, nil
	}
	st := SerialToTime(settlement)
	mt := SerialToTime(maturity)
	ncd := coupncd(st, mt, freq)
	return NumberVal(TimeToSerial(ncd)), nil
}

// fnCOUPNUM implements COUPNUM(settlement, maturity, frequency, [basis]).
// Returns the number of coupon dates between settlement and maturity.
func fnCoupnum(args []Value) (Value, error) {
	settlement, maturity, freq, _, ev := coupValidate(args)
	if ev != nil {
		return *ev, nil
	}
	st := SerialToTime(settlement)
	mt := SerialToTime(maturity)

	// Count coupon dates strictly after settlement and <= maturity.
	count := 0
	ncd := coupncd(st, mt, freq)
	for !ncd.After(mt) {
		count++
		monthStep := 12 / freq
		nextMonth := int(ncd.Month()) + monthStep
		nextYear := ncd.Year()
		for nextMonth > 12 {
			nextMonth -= 12
			nextYear++
		}
		ncd = clampDate(nextYear, nextMonth, mt.Day())
	}
	return NumberVal(float64(count)), nil
}

// fnCOUPDAYBS implements COUPDAYBS(settlement, maturity, frequency, [basis]).
// Returns the number of days from the beginning of the coupon period to settlement.
func fnCoupdaybs(args []Value) (Value, error) {
	settlement, maturity, freq, basis, ev := coupValidate(args)
	if ev != nil {
		return *ev, nil
	}
	st := SerialToTime(settlement)
	mt := SerialToTime(maturity)
	pcd := couppcd(st, mt, freq)

	switch basis {
	case 0, 4: // 30/360 day count
		sy, sm, sd := pcd.Year(), int(pcd.Month()), pcd.Day()
		ey, em, ed := st.Year(), int(st.Month()), st.Day()
		european := basis == 4
		return NumberVal(days360Calc(sy, sm, sd, ey, em, ed, european)), nil
	default: // basis 1, 2, 3: actual days
		pcdSerial := TimeToSerial(pcd)
		return NumberVal(settlement - pcdSerial), nil
	}
}

// fnCOUPDAYS implements COUPDAYS(settlement, maturity, frequency, [basis]).
// Returns the number of days in the coupon period containing settlement.
func fnCoupdays(args []Value) (Value, error) {
	settlement, maturity, freq, basis, ev := coupValidate(args)
	if ev != nil {
		return *ev, nil
	}

	switch basis {
	case 0, 4: // 30/360
		return NumberVal(360.0 / float64(freq)), nil
	case 1: // actual/actual
		st := SerialToTime(settlement)
		mt := SerialToTime(maturity)
		pcd := couppcd(st, mt, freq)
		ncd := coupncd(st, mt, freq)
		pcdSerial := TimeToSerial(pcd)
		ncdSerial := TimeToSerial(ncd)
		return NumberVal(ncdSerial - pcdSerial), nil
	case 2: // actual/360
		return NumberVal(360.0 / float64(freq)), nil
	case 3: // actual/365
		return NumberVal(365.0 / float64(freq)), nil
	}
	// unreachable (basis validated)
	return ErrorVal(ErrValNUM), nil
}

// fnCOUPDAYSNC implements COUPDAYSNC(settlement, maturity, frequency, [basis]).
// Returns the number of days from settlement to the next coupon date.
func fnCoupdaysnc(args []Value) (Value, error) {
	settlement, maturity, freq, basis, ev := coupValidate(args)
	if ev != nil {
		return *ev, nil
	}
	st := SerialToTime(settlement)
	mt := SerialToTime(maturity)
	ncd := coupncd(st, mt, freq)

	switch basis {
	case 0, 4: // 30/360: COUPDAYS - COUPDAYBS
		pcd := couppcd(st, mt, freq)
		sy, sm, sd := pcd.Year(), int(pcd.Month()), pcd.Day()
		ey, em, ed := st.Year(), int(st.Month()), st.Day()
		european := basis == 4
		daybs := days360Calc(sy, sm, sd, ey, em, ed, european)
		cdays := 360.0 / float64(freq)
		return NumberVal(cdays - daybs), nil
	default: // basis 1, 2, 3: actual days
		ncdSerial := TimeToSerial(ncd)
		return NumberVal(ncdSerial - settlement), nil
	}
}

// ---------------------------------------------------------------------------
// Internal helpers for computing coupon metrics from raw values.
// These avoid the Value boxing/unboxing overhead of the fnCoup* wrappers.
// ---------------------------------------------------------------------------

// coupnumRaw returns the number of remaining coupon periods.
func coupnumRaw(st, mt time.Time, freq int) int {
	count := 0
	ncd := coupncd(st, mt, freq)
	monthStep := 12 / freq
	for !ncd.After(mt) {
		count++
		nextMonth := int(ncd.Month()) + monthStep
		nextYear := ncd.Year()
		for nextMonth > 12 {
			nextMonth -= 12
			nextYear++
		}
		ncd = clampDate(nextYear, nextMonth, mt.Day())
	}
	return count
}

// coupdaybsRaw returns the number of days from the previous coupon date to settlement.
func coupdaybsRaw(settlement float64, st, mt time.Time, freq, basis int) float64 {
	pcd := couppcd(st, mt, freq)
	switch basis {
	case 0, 4:
		sy, sm, sd := pcd.Year(), int(pcd.Month()), pcd.Day()
		ey, em, ed := st.Year(), int(st.Month()), st.Day()
		return days360Calc(sy, sm, sd, ey, em, ed, basis == 4)
	default:
		return settlement - TimeToSerial(pcd)
	}
}

// coupdaysRaw returns the number of days in the coupon period containing settlement.
func coupdaysRaw(settlement float64, st, mt time.Time, freq, basis int) float64 {
	switch basis {
	case 0, 4:
		return 360.0 / float64(freq)
	case 1:
		pcd := couppcd(st, mt, freq)
		ncd := coupncd(st, mt, freq)
		return TimeToSerial(ncd) - TimeToSerial(pcd)
	case 2:
		return 360.0 / float64(freq)
	case 3:
		return 365.0 / float64(freq)
	}
	return 0
}

// coupdaysncRaw returns the number of days from settlement to the next coupon date.
func coupdaysncRaw(settlement float64, st, mt time.Time, freq, basis int) float64 {
	switch basis {
	case 0, 4:
		pcd := couppcd(st, mt, freq)
		sy, sm, sd := pcd.Year(), int(pcd.Month()), pcd.Day()
		ey, em, ed := st.Year(), int(st.Month()), st.Day()
		daybs := days360Calc(sy, sm, sd, ey, em, ed, basis == 4)
		return 360.0/float64(freq) - daybs
	default:
		ncd := coupncd(st, mt, freq)
		return TimeToSerial(ncd) - settlement
	}
}

// ---------------------------------------------------------------------------
// DURATION / MDURATION
// ---------------------------------------------------------------------------

// fnDuration implements DURATION(settlement, maturity, coupon, yld, frequency, [basis]).
// Returns the Macaulay duration for a security with an assumed par value of $100.
func fnDuration(args []Value) (Value, error) {
	if len(args) < 5 || len(args) > 6 {
		return ErrorVal(ErrValVALUE), nil
	}
	settlementRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	maturityRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	coupon, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	yld, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	freqRaw, e := CoerceNum(args[4])
	if e != nil {
		return *e, nil
	}
	basis := 0
	if len(args) == 6 {
		b, e := CoerceNum(args[5])
		if e != nil {
			return *e, nil
		}
		basis = int(b)
	}

	settlement := math.Trunc(settlementRaw)
	maturity := math.Trunc(maturityRaw)
	freq := int(math.Trunc(freqRaw))

	// Validate inputs.
	if coupon < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if yld < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if settlement >= maturity {
		return ErrorVal(ErrValNUM), nil
	}
	if freq != 1 && freq != 2 && freq != 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}

	dur := durationCalc(settlement, maturity, coupon, yld, freq, basis)
	return NumberVal(dur), nil
}

// durationCalc computes the Macaulay duration.
func durationCalc(settlement, maturity, coupon, yld float64, freq, basis int) float64 {
	st := SerialToTime(settlement)
	mt := SerialToTime(maturity)

	n := coupnumRaw(st, mt, freq)
	dsc := coupdaysncRaw(settlement, st, mt, freq, basis)
	e := coupdaysRaw(settlement, st, mt, freq, basis)

	// For basis 2 (actual/360) and 3 (actual/365), coupdaysncRaw returns
	// actual days but coupdaysRaw returns a fixed value (360/freq or 365/freq).
	// The DSC/E fraction must use a consistent convention so the ratio
	// stays in [0,1]. Use actual period days for E in these cases.
	if basis == 2 || basis == 3 {
		e = coupdaysRaw(settlement, st, mt, freq, 1) // actual/actual
	}

	// Fraction of the first coupon period remaining.
	accruedFrac := dsc / e

	yf := yld / float64(freq)    // yield per period
	cf := coupon / float64(freq) // coupon per period (as rate)

	var numerator, denominator float64

	for k := 1; k <= n; k++ {
		// Time in periods from settlement, expressed in years.
		tk := (float64(k-1) + accruedFrac) / float64(freq)

		// Cash flow at period k: coupon payment, plus par at maturity.
		cashFlow := cf * 100.0
		if k == n {
			cashFlow += 100.0
		}

		// Discount factor.
		discount := math.Pow(1.0+yf, float64(k-1)+accruedFrac)

		pv := cashFlow / discount
		numerator += tk * pv
		denominator += pv
	}

	if denominator == 0 {
		return 0
	}
	return numerator / denominator
}

// fnMduration implements MDURATION(settlement, maturity, coupon, yld, frequency, [basis]).
// Returns the modified Macaulay duration: DURATION / (1 + yld/frequency).
func fnMduration(args []Value) (Value, error) {
	if len(args) < 5 || len(args) > 6 {
		return ErrorVal(ErrValVALUE), nil
	}
	settlementRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	maturityRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	coupon, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	yld, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	freqRaw, e := CoerceNum(args[4])
	if e != nil {
		return *e, nil
	}
	basis := 0
	if len(args) == 6 {
		b, e := CoerceNum(args[5])
		if e != nil {
			return *e, nil
		}
		basis = int(b)
	}

	settlement := math.Trunc(settlementRaw)
	maturity := math.Trunc(maturityRaw)
	freq := int(math.Trunc(freqRaw))

	// Validate inputs.
	if coupon < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if yld < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if settlement >= maturity {
		return ErrorVal(ErrValNUM), nil
	}
	if freq != 1 && freq != 2 && freq != 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}

	dur := durationCalc(settlement, maturity, coupon, yld, freq, basis)
	mdur := dur / (1.0 + yld/float64(freq))
	return NumberVal(mdur), nil
}

// ---------------------------------------------------------------------------
// PRICE / YIELD
// ---------------------------------------------------------------------------

// priceCalc computes the clean price per $100 face value of a security that
// pays periodic interest. This is the core calculation shared by fnPrice and
// the Newton solver in fnYield.
func priceCalc(settlement, maturity, rate, yld float64, freq, basis int) float64 {
	st := SerialToTime(settlement)
	mt := SerialToTime(maturity)

	n := coupnumRaw(st, mt, freq)
	dsc := coupdaysncRaw(settlement, st, mt, freq, basis)
	e := coupdaysRaw(settlement, st, mt, freq, basis)
	a := coupdaybsRaw(settlement, st, mt, freq, basis)

	couponPmt := 100.0 * rate / float64(freq) // coupon payment per period

	if n == 1 {
		// Last (or only) coupon period: closed-form formula.
		dsr := dsc // days from settlement to redemption = days to next coupon (which is maturity)
		t1 := 100.0 + couponPmt
		t2 := yld / float64(freq) * dsr / e
		return (t1 / (1.0 + t2)) - (couponPmt * a / e)
	}

	// Multiple coupon periods: summation formula.
	frac := dsc / e
	yf := yld / float64(freq)

	// Redemption discounted to settlement.
	price := 100.0 / math.Pow(1.0+yf, float64(n-1)+frac)

	// Sum of discounted coupon payments.
	for k := 1; k <= n; k++ {
		price += couponPmt / math.Pow(1.0+yf, float64(k-1)+frac)
	}

	// Subtract accrued interest.
	price -= couponPmt * a / e

	return price
}

// fnPrice implements PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis]).
// Returns the price per $100 face value of a security that pays periodic interest.
func fnPrice(args []Value) (Value, error) {
	if len(args) < 6 || len(args) > 7 {
		return ErrorVal(ErrValVALUE), nil
	}
	settlementRaw, ev := CoerceNum(args[0])
	if ev != nil {
		return *ev, nil
	}
	maturityRaw, ev := CoerceNum(args[1])
	if ev != nil {
		return *ev, nil
	}
	rate, ev := CoerceNum(args[2])
	if ev != nil {
		return *ev, nil
	}
	yld, ev := CoerceNum(args[3])
	if ev != nil {
		return *ev, nil
	}
	redemption, ev := CoerceNum(args[4])
	if ev != nil {
		return *ev, nil
	}
	freqRaw, ev := CoerceNum(args[5])
	if ev != nil {
		return *ev, nil
	}
	basis := 0
	if len(args) == 7 {
		b, ev := CoerceNum(args[6])
		if ev != nil {
			return *ev, nil
		}
		basis = int(b)
	}

	settlement := math.Trunc(settlementRaw)
	maturity := math.Trunc(maturityRaw)
	freq := int(math.Trunc(freqRaw))

	// Validate inputs.
	if rate < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if yld < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if redemption <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if settlement >= maturity {
		return ErrorVal(ErrValNUM), nil
	}
	if freq != 1 && freq != 2 && freq != 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}

	// Use the general priceCalc which assumes redemption=100, then scale.
	price := priceCalc(settlement, maturity, rate, yld, freq, basis)
	// Adjust for non-100 redemption: the redemption component scales linearly.
	if redemption != 100.0 {
		// Recompute with explicit redemption handling.
		price = priceCalcRedemption(settlement, maturity, rate, yld, redemption, freq, basis)
	}

	return NumberVal(price), nil
}

// priceCalcRedemption computes the clean price with an arbitrary redemption value.
func priceCalcRedemption(settlement, maturity, rate, yld, redemption float64, freq, basis int) float64 {
	st := SerialToTime(settlement)
	mt := SerialToTime(maturity)

	n := coupnumRaw(st, mt, freq)
	dsc := coupdaysncRaw(settlement, st, mt, freq, basis)
	e := coupdaysRaw(settlement, st, mt, freq, basis)
	a := coupdaybsRaw(settlement, st, mt, freq, basis)

	couponPmt := 100.0 * rate / float64(freq)

	if n == 1 {
		dsr := dsc
		t1 := redemption + couponPmt
		t2 := yld / float64(freq) * dsr / e
		return (t1 / (1.0 + t2)) - (couponPmt * a / e)
	}

	frac := dsc / e
	yf := yld / float64(freq)

	price := redemption / math.Pow(1.0+yf, float64(n-1)+frac)

	for k := 1; k <= n; k++ {
		price += couponPmt / math.Pow(1.0+yf, float64(k-1)+frac)
	}

	price -= couponPmt * a / e

	return price
}

// fnYield implements YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis]).
// Returns the yield on a security that pays periodic interest.
func fnYield(args []Value) (Value, error) {
	if len(args) < 6 || len(args) > 7 {
		return ErrorVal(ErrValVALUE), nil
	}
	settlementRaw, ev := CoerceNum(args[0])
	if ev != nil {
		return *ev, nil
	}
	maturityRaw, ev := CoerceNum(args[1])
	if ev != nil {
		return *ev, nil
	}
	rate, ev := CoerceNum(args[2])
	if ev != nil {
		return *ev, nil
	}
	pr, ev := CoerceNum(args[3])
	if ev != nil {
		return *ev, nil
	}
	redemption, ev := CoerceNum(args[4])
	if ev != nil {
		return *ev, nil
	}
	freqRaw, ev := CoerceNum(args[5])
	if ev != nil {
		return *ev, nil
	}
	basis := 0
	if len(args) == 7 {
		b, ev := CoerceNum(args[6])
		if ev != nil {
			return *ev, nil
		}
		basis = int(b)
	}

	settlement := math.Trunc(settlementRaw)
	maturity := math.Trunc(maturityRaw)
	freq := int(math.Trunc(freqRaw))

	// Validate inputs.
	if rate < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if pr <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if redemption <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if settlement >= maturity {
		return ErrorVal(ErrValNUM), nil
	}
	if freq != 1 && freq != 2 && freq != 4 {
		return ErrorVal(ErrValNUM), nil
	}
	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}

	st := SerialToTime(settlement)
	mt := SerialToTime(maturity)
	n := coupnumRaw(st, mt, freq)

	if n == 1 {
		// Closed-form for single coupon period.
		dsc := coupdaysncRaw(settlement, st, mt, freq, basis)
		e := coupdaysRaw(settlement, st, mt, freq, basis)
		a := coupdaybsRaw(settlement, st, mt, freq, basis)

		couponPmt := rate / float64(freq)
		dsr := dsc

		num := (redemption/100.0 + couponPmt) - (pr/100.0 + a/e*couponPmt)
		den := pr/100.0 + a/e*couponPmt
		yield := (num / den) * (float64(freq) * e / dsr)
		return NumberVal(yield), nil
	}

	// Multiple coupon periods: use Newton's method to find yield where
	// priceCalcRedemption(yield) = pr.
	priceFn := func(y float64) float64 {
		return priceCalcRedemption(settlement, maturity, rate, y, redemption, freq, basis)
	}

	// Initial guess.
	guess := rate
	if guess <= 0 {
		guess = 0.1
	}

	// Newton's method with numerical derivative.
	const maxIter = 100
	const tol = 1e-12
	const dx = 1e-10

	y := guess
	for i := 0; i < maxIter; i++ {
		p := priceFn(y)
		diff := p - pr
		if math.Abs(diff) < tol {
			return NumberVal(y), nil
		}

		// Numerical derivative dp/dy.
		p2 := priceFn(y + dx)
		dpdy := (p2 - p) / dx
		if dpdy == 0 {
			break
		}

		y = y - diff/dpdy
	}

	// Check if we converged close enough.
	if math.Abs(priceFn(y)-pr) < 1e-7 {
		return NumberVal(y), nil
	}

	return ErrorVal(ErrValNUM), nil
}

// fnFVSchedule implements FVSCHEDULE(principal, schedule).
// Returns the future value of an initial principal after applying a series of
// compound interest rates: principal * (1+rate1) * (1+rate2) * ... * (1+rateN).
func fnFVSchedule(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	principal, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	rates := flattenValues(args[1])
	for _, v := range rates {
		switch v.Type {
		case ValueEmpty:
			// Blank cells are treated as zero (no interest), so multiply by 1.
			continue
		case ValueNumber:
			principal *= 1 + v.Num
		default:
			// spreadsheet treats booleans, strings, and any other non-numeric type
			// in the schedule as #VALUE! errors.
			return ErrorVal(ErrValVALUE), nil
		}
	}
	return NumberVal(principal), nil
}

// amordegrcRound rounds a value to the nearest integer, breaking exact ties
// toward zero.  Excel's AMORDEGRC uses this convention for depreciation loop
// iterations (periods 1+), where an intermediate product that lands exactly on
// 0.5 is rounded toward zero rather than away from zero.  For non-half values
// it behaves identically to math.Round.
//
// Example: amordegrcRound(5273.5) = 5273 (not 5274).
func amordegrcRound(x float64) float64 {
	truncated := math.Trunc(x)
	frac := math.Abs(x - truncated)
	if math.Abs(frac-0.5) < 1e-10 {
		return truncated
	}
	return math.Round(x)
}

// fnAmordegrc implements AMORDEGRC(cost, date_purchased, first_period, salvage, period, rate, [basis]).
// Returns the depreciation for each accounting period using the French degressive method.
func fnAmordegrc(args []Value) (Value, error) {
	if len(args) < 6 || len(args) > 7 {
		return ErrorVal(ErrValVALUE), nil
	}

	cost, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	datePurchasedRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	firstPeriodRaw, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	salvage, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	periodRaw, e := CoerceNum(args[4])
	if e != nil {
		return *e, nil
	}
	rate, e := CoerceNum(args[5])
	if e != nil {
		return *e, nil
	}

	basis := 0
	if len(args) == 7 {
		b, e := CoerceNum(args[6])
		if e != nil {
			return *e, nil
		}
		basis = int(b)
	}

	datePurchased := math.Trunc(datePurchasedRaw)
	firstPeriod := math.Trunc(firstPeriodRaw)
	period := int(math.Trunc(periodRaw))

	// Validate inputs.
	if cost < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if salvage < 0 || salvage > cost {
		return ErrorVal(ErrValNUM), nil
	}
	if period < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if rate <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	// Basis 2 is not supported for AMORDEGRC.
	if basis != 0 && basis != 1 && basis != 3 && basis != 4 {
		return ErrorVal(ErrValNUM), nil
	}

	// Compute life = 1/rate and determine the coefficient.
	// Empirical testing against Excel shows these coefficient brackets:
	//   life <= 2:  #NUM! error
	//   life <= 4:  1.5
	//   life <= 6:  2.0
	//   life > 6:   2.5
	// Note: the Microsoft documentation claims additional error ranges
	// (0-1, 1-2, 2-3, 4-5) but actual Excel behavior only errors for
	// life <= 2.  The 4-5 "gap" does not exist in practice.
	life := 1.0 / rate
	if life <= 2 {
		return ErrorVal(ErrValNUM), nil
	}
	var coeff float64
	switch {
	case life <= 4:
		coeff = 1.5
	case life <= 6:
		coeff = 2
	default:
		coeff = 2.5
	}

	adjustedRate := rate * coeff

	// Compute year fraction for the prorated first period.
	dsm, bYear := dayCountBasis(datePurchased, firstPeriod, basis)
	if bYear == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	yearFrac := dsm / bYear

	// Total number of depreciation periods = ceil(life).
	// Periods are numbered 0 .. nper-1.
	// The last two periods use the 50%/rest rule:
	//   period nper-2 (second-to-last): round(remaining_cost * 0.5)
	//   period nper-1 (last): remaining_cost - half
	// If the "rest" in the last period would be less than salvage, it is 0.
	nper := int(math.Ceil(life))

	// Period 0: prorated depreciation.
	// Period 0 uses standard rounding (half away from zero) to match
	// Excel's behavior for the initial prorated period.
	dep0 := math.Round(cost * adjustedRate * yearFrac)

	if period == 0 {
		return NumberVal(dep0), nil
	}
	if period >= nper {
		return NumberVal(0), nil
	}

	// Build the depreciation schedule up to the requested period.
	fCost := cost - dep0 // remaining cost after period 0

	for p := 1; p <= period; p++ {
		switch {
		case p == nper-2:
			// Second-to-last period: 50% of remaining cost.
			half := amordegrcRound(fCost * 0.5)
			if p == period {
				return NumberVal(half), nil
			}
			fCost -= half

		case p == nper-1:
			// Last period: the rest of remaining cost.
			// If the rest is less than salvage, return 0.
			if fCost <= salvage {
				return NumberVal(0), nil
			}
			return NumberVal(fCost), nil

		default:
			// Normal period: degressive depreciation.
			dep := amordegrcRound(adjustedRate * fCost)
			if p == period {
				return NumberVal(dep), nil
			}
			fCost -= dep
		}
	}

	// Should not be reached, but safeguard.
	return NumberVal(0), nil
}

// fnAmorlinc implements AMORLINC(cost, date_purchased, first_period, salvage, period, rate, [basis]).
// Returns the depreciation for each accounting period using straight-line (linear) depreciation,
// prorated for the first period. Used in the French accounting system.
func fnAmorlinc(args []Value) (Value, error) {
	if len(args) < 6 || len(args) > 7 {
		return ErrorVal(ErrValVALUE), nil
	}

	cost, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	datePurchasedRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	firstPeriodRaw, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	salvage, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	periodRaw, e := CoerceNum(args[4])
	if e != nil {
		return *e, nil
	}
	rate, e := CoerceNum(args[5])
	if e != nil {
		return *e, nil
	}

	basis := 0
	if len(args) == 7 {
		b, e := CoerceNum(args[6])
		if e != nil {
			return *e, nil
		}
		basis = int(b)
	}

	datePurchased := math.Trunc(datePurchasedRaw)
	firstPeriod := math.Trunc(firstPeriodRaw)
	period := int(math.Trunc(periodRaw))

	// Validate inputs.
	if cost <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if salvage < 0 || salvage > cost {
		return ErrorVal(ErrValNUM), nil
	}
	if period < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if rate <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	// Basis 2 is not supported for AMORLINC.
	if basis != 0 && basis != 1 && basis != 3 && basis != 4 {
		return ErrorVal(ErrValNUM), nil
	}

	// When salvage equals cost, nothing to depreciate.
	if salvage == cost {
		return NumberVal(0), nil
	}

	// Compute year fraction for the prorated first period.
	dsm, bYear := dayCountBasis(datePurchased, firstPeriod, basis)
	if bYear == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	yearFrac := dsm / bYear

	// Normal period depreciation (flat amount).
	normalDep := cost * rate

	// Period 0: prorated depreciation.
	dep0 := cost * rate * yearFrac

	// Total depreciable amount.
	depreciable := cost - salvage

	if period == 0 {
		if dep0 > depreciable {
			return NumberVal(depreciable), nil
		}
		return NumberVal(dep0), nil
	}

	// Accumulate depreciation through prior periods to cap at depreciable amount.
	accum := dep0
	for p := 1; p <= period; p++ {
		remaining := depreciable - accum
		if remaining <= 0 {
			return NumberVal(0), nil
		}
		if p == period {
			if normalDep > remaining {
				return NumberVal(remaining), nil
			}
			return NumberVal(normalDep), nil
		}
		accum += normalDep
	}

	return NumberVal(0), nil
}
