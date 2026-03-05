package formula

import (
	"fmt"
	"math"
	"strings"
)

func init() {
	Register("DELTA", NoCtx(fnDELTA))
	Register("DEC2BIN", NoCtx(fnDec2Bin))
	Register("GESTEP", NoCtx(fnGESTEP))
}

// fnDELTA implements the Excel DELTA function.
// DELTA(number1, [number2]) — returns 1 if number1 == number2, else 0.
// number2 defaults to 0. Non-numeric arguments produce #VALUE!.
func fnDELTA(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	n1, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	var n2 float64
	if len(args) == 2 {
		n2, e = CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
	}

	if n1 == n2 {
		return NumberVal(1), nil
	}
	return NumberVal(0), nil
}

// fnDec2Bin implements the Excel DEC2BIN function.
// DEC2BIN(number, [places]) — converts a decimal number to binary.
// number must be between -512 and 511 (inclusive). Non-integer values are truncated.
// Negative numbers use two's complement (10-bit). places specifies minimum digits (1–10).
func fnDec2Bin(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	num, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	// Truncate to integer.
	num = math.Trunc(num)

	// Range check: -512 to 511.
	if num < -512 || num > 511 {
		return ErrorVal(ErrValNUM), nil
	}

	n := int(num)

	var result string
	if n < 0 {
		// Two's complement: add 1024 to get 10-bit representation.
		result = fmt.Sprintf("%b", n+1024)
	} else {
		result = fmt.Sprintf("%b", n)
	}

	if len(args) == 2 {
		places, e := CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		places = math.Trunc(places)
		if places <= 0 || places > 10 {
			return ErrorVal(ErrValNUM), nil
		}
		p := int(places)
		if len(result) > p {
			return ErrorVal(ErrValNUM), nil
		}
		if n >= 0 {
			result = strings.Repeat("0", p-len(result)) + result
		}
	}

	return StringVal(result), nil
}

// fnGESTEP implements the Excel GESTEP function.
// GESTEP(number, [step]) — returns 1 if number >= step, else 0.
// step defaults to 0. Non-numeric arguments produce #VALUE!.
func fnGESTEP(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	num, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	var step float64
	if len(args) == 2 {
		step, e = CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
	}

	if num >= step {
		return NumberVal(1), nil
	}
	return NumberVal(0), nil
}
