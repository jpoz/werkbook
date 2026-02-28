package formula

import (
	"math"
	"math/rand"
)

func fnABS(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueArray {
		return liftUnary(args[0], func(v Value) Value {
			n, e := coerceNum(v)
			if e != nil {
				return *e
			}
			return NumberVal(math.Abs(n))
		}), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(math.Abs(n)), nil
}

func fnACOS(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n < -1 || n > 1 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(math.Acos(n)), nil
}

func fnASIN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n < -1 || n > 1 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(math.Asin(n)), nil
}

func fnATAN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(math.Atan(n)), nil
}

// fnATAN2 implements Excel's ATAN2(x, y) — note the reversed argument order vs Go's math.Atan2(y, x).
func fnATAN2(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	y, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	if x == 0 && y == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(math.Atan2(y, x)), nil
}

func fnCEILING(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	sig, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	if sig == 0 {
		return NumberVal(0), nil
	}
	return NumberVal(math.Ceil(n/sig) * sig), nil
}

func fnCOMBIN(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	nf, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	kf, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	n := int(math.Trunc(nf))
	k := int(math.Trunc(kf))
	if n < 0 || k < 0 || n < k {
		return ErrorVal(ErrValNUM), nil
	}
	// Multiplicative formula: C(n,k) = ∏(i=1..k) (n-k+i)/i
	result := 1.0
	for i := 1; i <= k; i++ {
		result = result * float64(n-k+i) / float64(i)
	}
	return NumberVal(result), nil
}

func fnDEGREES(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(n * 180 / math.Pi), nil
}

func fnCOS(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(math.Cos(n)), nil
}

func fnEVEN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n == 0 {
		return NumberVal(0), nil
	}
	if n > 0 {
		return NumberVal(math.Ceil(n/2) * 2), nil
	}
	return NumberVal(math.Floor(n/2) * 2), nil
}

func fnODD(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n == 0 {
		return NumberVal(1), nil
	}
	if n > 0 {
		n2 := math.Ceil(n)
		if int(n2)%2 == 0 {
			n2++
		}
		return NumberVal(n2), nil
	}
	// negative: round away from zero (toward more negative)
	n2 := math.Floor(n)
	if int(n2)%2 == 0 {
		n2--
	}
	return NumberVal(n2), nil
}

func fnEXP(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(math.Exp(n)), nil
}

func fnFACT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	n = math.Trunc(n)
	if n < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	result := 1.0
	for i := 2.0; i <= n; i++ {
		result *= i
	}
	return NumberVal(result), nil
}

func fnFLOOR(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	sig, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	if sig == 0 {
		return NumberVal(0), nil
	}
	return NumberVal(math.Floor(n/sig) * sig), nil
}

func fnGCD(args []Value) (Value, error) {
	if len(args) < 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	result := int64(0)
	for _, arg := range args {
		n, e := coerceNum(arg)
		if e != nil {
			return *e, nil
		}
		n = math.Trunc(n)
		if n < 0 {
			return ErrorVal(ErrValNUM), nil
		}
		v := int64(n)
		// Euclidean algorithm
		a, b := result, v
		for b != 0 {
			a, b = b, a%b
		}
		result = a
	}
	return NumberVal(float64(result)), nil
}

func fnLCM(args []Value) (Value, error) {
	if len(args) < 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	result := int64(1)
	hasZero := false
	for _, arg := range args {
		n, e := coerceNum(arg)
		if e != nil {
			return *e, nil
		}
		n = math.Trunc(n)
		if n < 0 {
			return ErrorVal(ErrValNUM), nil
		}
		v := int64(n)
		if v == 0 {
			hasZero = true
			continue
		}
		// lcm(a, b) = a / gcd(a, b) * b
		a, b := result, v
		for b != 0 {
			a, b = b, a%b
		}
		result = result / a * v
	}
	if hasZero {
		return NumberVal(0), nil
	}
	return NumberVal(float64(result)), nil
}

func fnINT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(math.Floor(n)), nil
}

func fnLN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(math.Log(n)), nil
}

func fnLOG(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	base := 10.0
	if len(args) == 2 {
		base, e = coerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		if base <= 0 || base == 1 {
			return ErrorVal(ErrValNUM), nil
		}
	}
	return NumberVal(math.Log(n) / math.Log(base)), nil
}

func fnLOG10(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(math.Log10(n)), nil
}

func fnMOD(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	d, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	if d == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	q := n / d
	// When |n/d| is very large, math.Floor loses precision and the result is
	// floating-point noise. Excel returns #NUM! in this case.
	if math.Abs(q) > 1<<49 {
		return ErrorVal(ErrValNUM), nil
	}
	result := n - d*math.Floor(q)
	return NumberVal(result), nil
}

func fnMROUND(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	multiple, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	if multiple == 0 {
		return NumberVal(0), nil
	}
	if (n > 0 && multiple < 0) || (n < 0 && multiple > 0) {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(math.Round(n/multiple) * multiple), nil
}

func fnPERMUT(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	nf, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	kf, e2 := coerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}
	if nf <= 0 || kf < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	n := int(nf)
	k := int(kf)
	if n < k {
		return ErrorVal(ErrValNUM), nil
	}
	result := 1.0
	for i := 0; i < k; i++ {
		result *= float64(n - i)
	}
	return NumberVal(result), nil
}

func fnPI(args []Value) (Value, error) {
	if len(args) != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return NumberVal(math.Pi), nil
}

func fnPOWER(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	base, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	exp, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	result := math.Pow(base, exp)
	if math.IsNaN(result) || math.IsInf(result, 0) {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(result), nil
}

func fnPRODUCT(args []Value) (Value, error) {
	product := 1.0
	count := 0
	if e := iterateNumeric(args, func(n float64) { product *= n; count++ }); e != nil {
		return *e, nil
	}
	if count == 0 {
		return NumberVal(0), nil
	}
	return NumberVal(product), nil
}

func fnQUOTIENT(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	num, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	den, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	if den == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(math.Trunc(num / den)), nil
}

func fnRADIANS(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(n * math.Pi / 180), nil
}

func fnRAND(args []Value) (Value, error) {
	if len(args) != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return NumberVal(rand.Float64()), nil
}

func fnRANDBETWEEN(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	bottom, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	top, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	lo := int(math.Ceil(bottom))
	hi := int(math.Floor(top))
	if lo > hi {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(float64(lo + rand.Intn(hi-lo+1))), nil
}

func fnROUND(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	digits, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pow := math.Pow(10, math.Floor(digits))
	return NumberVal(math.Round(n*pow) / pow), nil
}

func fnROUNDDOWN(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	digits, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pow := math.Pow(10, math.Floor(digits))
	if n >= 0 {
		return NumberVal(math.Floor(n*pow) / pow), nil
	}
	return NumberVal(math.Ceil(n*pow) / pow), nil
}

func fnROUNDUP(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	digits, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pow := math.Pow(10, math.Floor(digits))
	if n >= 0 {
		return NumberVal(math.Ceil(n*pow) / pow), nil
	}
	return NumberVal(math.Floor(n*pow) / pow), nil
}

func fnSIGN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	switch {
	case n > 0:
		return NumberVal(1), nil
	case n < 0:
		return NumberVal(-1), nil
	default:
		return NumberVal(0), nil
	}
}

func fnSIN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(math.Sin(n)), nil
}

func fnSQRT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(math.Sqrt(n)), nil
}

func fnTRUNC(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	digits := 0.0
	if len(args) == 2 {
		digits, e = coerceNum(args[1])
		if e != nil {
			return *e, nil
		}
	}
	pow := math.Pow(10, math.Floor(digits))
	return NumberVal(math.Trunc(n*pow) / pow), nil
}

func fnSUBTOTAL(args []Value) (Value, error) {
	if len(args) < 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	fnNum, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	fn := int(fnNum)
	// Normalize 101-111 to 1-11
	if fn >= 101 && fn <= 111 {
		fn -= 100
	}
	if fn < 1 || fn > 11 {
		return ErrorVal(ErrValVALUE), nil
	}
	rest := args[1:]
	switch fn {
	case 1:
		return fnAVERAGE(rest)
	case 2:
		return fnCOUNT(rest)
	case 3:
		return fnCOUNTA(rest)
	case 4:
		return fnMAX(rest)
	case 5:
		return fnMIN(rest)
	case 6:
		return fnPRODUCT(rest)
	case 7:
		return fnSTDEV(rest)
	case 8:
		return fnSTDEVP(rest)
	case 9:
		return fnSUM(rest)
	case 10:
		return fnVAR(rest)
	case 11:
		return fnVARP(rest)
	default:
		return ErrorVal(ErrValVALUE), nil
	}
}

func fnTAN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(math.Tan(n)), nil
}
