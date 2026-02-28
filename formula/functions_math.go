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
	result := n - d*math.Floor(n/d)
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
