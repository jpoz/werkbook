package formula

import "math"

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
	Register("SLN", NoCtx(fnSLN))
	Register("XIRR", NoCtx(fnXIRR))
	Register("XNPV", NoCtx(fnXNPV))
}

// flattenValues extracts all numeric values from an arg that may be a scalar or array (range).
func flattenValues(arg Value) []Value {
	if arg.Type == ValueArray {
		var out []Value
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
				dnpv -= float64(i) * cf / math.Pow(1+rate, float64(i+1))
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
	d0, e := CoerceNum(dates[0])
	if e != nil {
		return *e, nil
	}
	xnpv := 0.0
	for i := range values {
		v, ev := CoerceNum(values[i])
		if ev != nil {
			return *ev, nil
		}
		di, ed := CoerceNum(dates[i])
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

	d0, e := CoerceNum(dates[0])
	if e != nil {
		return *e, nil
	}

	hasPos, hasNeg := false, false
	for i := range values {
		v, ev := CoerceNum(values[i])
		if ev != nil {
			return *ev, nil
		}
		di, ed := CoerceNum(dates[i])
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
	for iter := 0; iter < 100; iter++ {
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
				dxnpv -= t * cf / math.Pow(1+rate, t+1)
			}
		}
		if math.Abs(xnpv) < 1e-10 {
			return NumberVal(rate), nil
		}
		if dxnpv == 0 {
			return ErrorVal(ErrValNUM), nil
		}
		newRate := rate - xnpv/dxnpv
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
