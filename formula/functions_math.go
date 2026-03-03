package formula

import (
	"math"
	"math/rand"
)

func init() {
	Register("ABS", NoCtx(fnABS))
	Register("CEILING", NoCtx(fnCEILING))
	Register("FLOOR", NoCtx(fnFLOOR))
	Register("INT", NoCtx(fnINT))
	Register("MOD", NoCtx(fnMOD))
	Register("POWER", NoCtx(fnPOWER))
	Register("RAND", NoCtx(fnRAND))
	Register("RANDBETWEEN", NoCtx(fnRANDBETWEEN))
	Register("ROUND", NoCtx(fnROUND))
	Register("ROUNDDOWN", NoCtx(fnROUNDDOWN))
	Register("ROUNDUP", NoCtx(fnROUNDUP))
	Register("SQRT", NoCtx(fnSQRT))
}

func fnABS(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			n, e := CoerceNum(v)
			if e != nil {
				return *e
			}
			return NumberVal(math.Abs(n))
		}), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(math.Abs(n)), nil
}

func fnCEILING(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	sig, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	if sig == 0 {
		return NumberVal(0), nil
	}
	if (n > 0 && sig < 0) || (n < 0 && sig > 0) {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(math.Ceil(n/sig) * sig), nil
}

func fnFLOOR(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	sig, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	if sig == 0 {
		return NumberVal(0), nil
	}
	if (n > 0 && sig < 0) || (n < 0 && sig > 0) {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(math.Floor(n/sig) * sig), nil
}

func fnINT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(math.Floor(n)), nil
}

func fnMOD(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	d, e := CoerceNum(args[1])
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

func fnPOWER(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	base, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	exp, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	result := math.Pow(base, exp)
	if math.IsNaN(result) || math.IsInf(result, 0) {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(result), nil
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
	bottom, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	top, e := CoerceNum(args[1])
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
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	digits, e := CoerceNum(args[1])
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
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	digits, e := CoerceNum(args[1])
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
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	digits, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	pow := math.Pow(10, math.Floor(digits))
	if n >= 0 {
		return NumberVal(math.Ceil(n*pow) / pow), nil
	}
	return NumberVal(math.Floor(n*pow) / pow), nil
}

func fnSQRT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(math.Sqrt(n)), nil
}
