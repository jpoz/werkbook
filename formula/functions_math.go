package formula

import (
	"math"
	"math/rand"
	"strconv"
	"strings"
)

// subtotalFuncNames maps SUBTOTAL function_num (1-11) to the aggregate function name.
var subtotalFuncNames = [12]string{
	1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN",
	6: "PRODUCT", 7: "STDEV", 8: "STDEVP", 9: "SUM", 10: "VAR", 11: "VARP",
}

func init() {
	Register("ABS", NoCtx(fnABS))
	Register("ACOS", NoCtx(fnACOS))
	Register("ACOSH", NoCtx(fnACOSH))
	Register("ACOT", NoCtx(fnACOT))
	Register("ACOTH", NoCtx(fnACOTH))
	Register("ARABIC", NoCtx(fnARABIC))
	Register("ASIN", NoCtx(fnASIN))
	Register("ASINH", NoCtx(fnASINH))
	Register("ATAN", NoCtx(fnATAN))
	Register("ATAN2", NoCtx(fnATAN2))
	Register("ATANH", NoCtx(fnATANH))
	Register("BASE", NoCtx(fnBASE))
	Register("BITAND", NoCtx(fnBITAND))
	Register("BITLSHIFT", NoCtx(fnBITLSHIFT))
	Register("BITOR", NoCtx(fnBITOR))
	Register("BITRSHIFT", NoCtx(fnBITRSHIFT))
	Register("BITXOR", NoCtx(fnBITXOR))
	Register("CEILING", NoCtx(fnCEILING))
	Register("CEILING.MATH", NoCtx(fnCEILINGMATH))
	Register("CEILING.PRECISE", NoCtx(fnCEILINGPRECISE))
	Register("COMBIN", NoCtx(fnCOMBIN))
	Register("COMBINA", NoCtx(fnCOMBINA))
	Register("COS", NoCtx(fnCOS))
	Register("COSH", NoCtx(fnCOSH))
	Register("COT", NoCtx(fnCOT))
	Register("COTH", NoCtx(fnCOTH))
	Register("CSC", NoCtx(fnCSC))
	Register("CSCH", NoCtx(fnCSCH))
	Register("DECIMAL", NoCtx(fnDECIMAL))
	Register("DEGREES", NoCtx(fnDEGREES))
	Register("ERF", NoCtx(fnERF))
	Register("ERF.PRECISE", NoCtx(fnERFPRECISE))
	Register("ERFC", NoCtx(fnERFC))
	Register("ERFC.PRECISE", NoCtx(fnERFCPRECISE))
	Register("EVEN", NoCtx(fnEVEN))
	Register("EXP", NoCtx(fnEXP))
	Register("FACT", NoCtx(fnFACT))
	Register("FACTDOUBLE", NoCtx(fnFACTDOUBLE))
	Register("FLOOR", NoCtx(fnFLOOR))
	Register("FLOOR.MATH", NoCtx(fnFLOORMATH))
	Register("FLOOR.PRECISE", NoCtx(fnFLOORPRECISE))
	Register("GCD", NoCtx(fnGCD))
	Register("INT", NoCtx(fnINT))
	Register("LCM", NoCtx(fnLCM))
	Register("LN", NoCtx(fnLN))
	Register("LOG", NoCtx(fnLOG))
	Register("LOG10", NoCtx(fnLOG10))
	Register("MOD", NoCtx(fnMOD))
	Register("MDETERM", NoCtx(fnMDETERM))
	Register("MINVERSE", NoCtx(fnMINVERSE))
	Register("MMULT", NoCtx(fnMMULT))
	Register("MROUND", NoCtx(fnMROUND))
	Register("MUNIT", NoCtx(fnMUNIT))
	Register("MULTINOMIAL", NoCtx(fnMULTINOMIAL))
	Register("ODD", NoCtx(fnODD))
	Register("PERMUT", NoCtx(fnPERMUT))
	Register("PI", NoCtx(fnPI))
	Register("POWER", NoCtx(fnPOWER))
	Register("PRODUCT", NoCtx(fnPRODUCT))
	Register("QUOTIENT", NoCtx(fnQUOTIENT))
	Register("RADIANS", NoCtx(fnRADIANS))
	Register("RAND", NoCtx(fnRAND))
	Register("RANDBETWEEN", NoCtx(fnRANDBETWEEN))
	Register("ROUND", NoCtx(fnROUND))
	Register("ROUNDDOWN", NoCtx(fnROUNDDOWN))
	Register("ROUNDUP", NoCtx(fnROUNDUP))
	Register("SEC", NoCtx(fnSEC))
	Register("SECH", NoCtx(fnSECH))
	Register("SEQUENCE", NoCtx(fnSEQUENCE))
	Register("SERIESSUM", NoCtx(fnSERIESSUM))
	Register("SIGN", NoCtx(fnSIGN))
	Register("SIN", NoCtx(fnSIN))
	Register("SINH", NoCtx(fnSINH))
	Register("SQRT", NoCtx(fnSQRT))
	Register("SQRTPI", NoCtx(fnSQRTPI))
	Register("SUBTOTAL", fnSUBTOTALCtx)
	Register("TAN", NoCtx(fnTAN))
	Register("TANH", NoCtx(fnTANH))
	Register("TRUNC", NoCtx(fnTRUNC))
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
	// Excel: positive number with negative significance is an error.
	if n > 0 && sig < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	// Excel: negative number with positive significance rounds toward zero.
	if n < 0 && sig > 0 {
		return NumberVal(-math.Floor(math.Abs(n)/sig) * sig), nil
	}
	return NumberVal(math.Ceil(n/sig) * sig), nil
}

func fnCEILINGMATH(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	// Default significance: +1 for positive numbers, -1 for negative.
	sig := 1.0
	if n < 0 {
		sig = -1.0
	}
	if len(args) >= 2 {
		sig, e = CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
	}
	mode := 0.0
	if len(args) == 3 {
		mode, e = CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
	}
	if sig == 0 {
		return NumberVal(0), nil
	}
	// Use absolute significance for the computation — the sign of significance
	// does not affect the result in CEILING.MATH (unlike CEILING).
	absSig := math.Abs(sig)
	if n >= 0 {
		// Positive numbers: round up (toward +infinity).
		return NumberVal(math.Ceil(n/absSig) * absSig), nil
	}
	// Negative numbers:
	if mode == 0 {
		// mode=0: round toward +infinity (toward zero).
		return NumberVal(-math.Floor(math.Abs(n)/absSig) * absSig), nil
	}
	// mode≠0: round away from zero (toward -infinity).
	return NumberVal(-math.Ceil(math.Abs(n)/absSig) * absSig), nil
}

func fnCEILINGPRECISE(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	sig := 1.0
	if len(args) == 2 {
		sig, e = CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
	}
	if sig == 0 {
		return NumberVal(0), nil
	}
	// Always use absolute value of significance.
	absSig := math.Abs(sig)
	// Always round toward +infinity.
	return NumberVal(math.Ceil(n/absSig) * absSig), nil
}

func fnFLOORPRECISE(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	sig := 1.0
	if len(args) == 2 {
		sig, e = CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
	}
	if sig == 0 {
		return NumberVal(0), nil
	}
	// Always use absolute value of significance.
	absSig := math.Abs(sig)
	// Always round toward -infinity.
	return NumberVal(math.Floor(n/absSig) * absSig), nil
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
	// Special case: when number is 0, result is always 0 regardless of significance.
	if n == 0 {
		return NumberVal(0), nil
	}
	if sig == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	// Excel: positive number with negative significance is an error.
	if n > 0 && sig < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	// Excel: negative number with positive significance rounds away from zero.
	if n < 0 && sig > 0 {
		return NumberVal(-math.Ceil(math.Abs(n)/sig) * sig), nil
	}
	return NumberVal(math.Floor(n/sig) * sig), nil
}

func fnFLOORMATH(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	// Default significance: +1 for positive numbers, -1 for negative.
	sig := 1.0
	if n < 0 {
		sig = -1.0
	}
	if len(args) >= 2 {
		sig, e = CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
	}
	mode := 0.0
	if len(args) == 3 {
		mode, e = CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
	}
	if sig == 0 {
		return NumberVal(0), nil
	}
	// Use absolute significance for the computation — the sign of significance
	// does not affect the result in FLOOR.MATH (unlike FLOOR).
	absSig := math.Abs(sig)
	if n >= 0 {
		// Positive numbers: round down (toward zero / toward -infinity).
		return NumberVal(math.Floor(n/absSig) * absSig), nil
	}
	// Negative numbers:
	if mode == 0 {
		// mode=0: round toward -infinity (away from zero).
		return NumberVal(-math.Ceil(math.Abs(n)/absSig) * absSig), nil
	}
	// mode≠0: round toward zero.
	return NumberVal(-math.Floor(math.Abs(n)/absSig) * absSig), nil
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
	// floating-point noise. Excel returns #NUM! in this case. The threshold
	// aligns with Excel's behavior: once the quotient exceeds ~10^13, the
	// intermediate INT(n/d)*d multiplication cannot recover n accurately
	// within float64's ~15 significant digits.
	if math.Abs(q) >= 1e13 {
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
	// Excel: POWER(0,0) returns #NUM!
	if base == 0 && exp == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	// Excel: POWER(0, negative) returns #DIV/0!
	if base == 0 && exp < 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	// Handle negative base with fractional exponent: detect unit fractions 1/n
	// where n is an odd integer, which allows computing real-valued roots.
	if base < 0 && exp != math.Floor(exp) {
		recip := 1.0 / exp
		rounded := math.Round(recip)
		if math.Abs(recip-rounded) < 1e-9 && int(rounded)%2 != 0 {
			// Odd root of negative number: -|base|^(1/n) * sign
			root := math.Pow(math.Abs(base), exp)
			if math.IsNaN(root) || math.IsInf(root, 0) {
				return ErrorVal(ErrValNUM), nil
			}
			return NumberVal(-root), nil
		}
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

func fnACOS(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n < -1 || n > 1 { return ErrorVal(ErrValNUM), nil }
	return NumberVal(math.Acos(n)), nil
}
func fnACOSH(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n < 1 { return ErrorVal(ErrValNUM), nil }
	return NumberVal(math.Acosh(n)), nil
}
func fnACOT(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(math.Pi/2 - math.Atan(n)), nil
}
func fnACOTH(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n >= -1 && n <= 1 { return ErrorVal(ErrValNUM), nil }
	return NumberVal(0.5 * math.Log((n+1)/(n-1))), nil
}
func fnARABIC(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	var text string
	switch args[0].Type {
	case ValueString: text = strings.TrimSpace(strings.ToUpper(args[0].Str))
	case ValueError: return args[0], nil
	case ValueEmpty: return NumberVal(0), nil
	default: return ErrorVal(ErrValVALUE), nil
	}
	if text == "" { return NumberVal(0), nil }
	negative := false
	if text[0] == '-' { negative = true; text = text[1:] }
	romanValues := map[byte]int{'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}
	result := 0
	for i := 0; i < len(text); i++ {
		val, ok := romanValues[text[i]]
		if !ok { return ErrorVal(ErrValVALUE), nil }
		if i+1 < len(text) {
			nextVal, ok2 := romanValues[text[i+1]]
			if !ok2 { return ErrorVal(ErrValVALUE), nil }
			if val < nextVal { result -= val } else { result += val }
		} else { result += val }
	}
	if negative { result = -result }
	return NumberVal(float64(result)), nil
}
func fnASIN(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n < -1 || n > 1 { return ErrorVal(ErrValNUM), nil }
	return NumberVal(math.Asin(n)), nil
}
func fnASINH(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(math.Asinh(n)), nil
}
func fnATAN(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(math.Atan(n)), nil
}
func fnATAN2(args []Value) (Value, error) {
	if len(args) != 2 { return ErrorVal(ErrValVALUE), nil }
	x, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	y, e := CoerceNum(args[1]); if e != nil { return *e, nil }
	if x == 0 && y == 0 { return ErrorVal(ErrValDIV0), nil }
	return NumberVal(math.Atan2(y, x)), nil
}
func fnATANH(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n <= -1 || n >= 1 { return ErrorVal(ErrValNUM), nil }
	return NumberVal(math.Atanh(n)), nil
}
func fnBASE(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 { return ErrorVal(ErrValVALUE), nil }
	nf, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	n := int64(math.Trunc(nf))
	if n < 0 { return ErrorVal(ErrValNUM), nil }
	rf, e := CoerceNum(args[1]); if e != nil { return *e, nil }
	radix := int(math.Trunc(rf))
	if radix < 2 || radix > 36 { return ErrorVal(ErrValNUM), nil }
	minLen := 0
	if len(args) == 3 {
		ml, e := CoerceNum(args[2]); if e != nil { return *e, nil }
		minLen = int(math.Trunc(ml))
		if minLen < 0 { return ErrorVal(ErrValNUM), nil }
	}
	result := strings.ToUpper(strconv.FormatInt(n, radix))
	for len(result) < minLen { result = "0" + result }
	return StringVal(result), nil
}
func fnCOMBIN(args []Value) (Value, error) {
	if len(args) != 2 { return ErrorVal(ErrValVALUE), nil }
	nf, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	kf, e := CoerceNum(args[1]); if e != nil { return *e, nil }
	n := int(math.Trunc(nf)); k := int(math.Trunc(kf))
	if n < 0 || k < 0 || n < k { return ErrorVal(ErrValNUM), nil }
	result := 1.0
	for i := 1; i <= k; i++ { result = result * float64(n-k+i) / float64(i) }
	return NumberVal(result), nil
}
func fnCOMBINA(args []Value) (Value, error) {
	if len(args) != 2 { return ErrorVal(ErrValVALUE), nil }
	nf, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	kf, e := CoerceNum(args[1]); if e != nil { return *e, nil }
	n := int(math.Trunc(nf)); k := int(math.Trunc(kf))
	if n < 0 || k < 0 { return ErrorVal(ErrValNUM), nil }
	total := n + k - 1
	if k == 0 { return NumberVal(1), nil }
	if total < 0 { return ErrorVal(ErrValNUM), nil }
	result := 1.0
	for i := 1; i <= k; i++ { result = result * float64(total-k+i) / float64(i) }
	return NumberVal(result), nil
}
func fnDECIMAL(args []Value) (Value, error) {
	if len(args) != 2 { return ErrorVal(ErrValVALUE), nil }
	text := ""
	switch args[0].Type {
	case ValueString: text = args[0].Str
	case ValueNumber: text = strconv.FormatFloat(args[0].Num, 'f', -1, 64)
	case ValueError: return args[0], nil
	default: return ErrorVal(ErrValVALUE), nil
	}
	radix, e := CoerceNum(args[1]); if e != nil { return *e, nil }
	r := int(math.Trunc(radix))
	if r < 2 || r > 36 { return ErrorVal(ErrValNUM), nil }
	text = strings.TrimSpace(text)
	if text == "" { return ErrorVal(ErrValNUM), nil }
	result, err := strconv.ParseInt(text, r, 64)
	if err != nil { return ErrorVal(ErrValNUM), nil }
	return NumberVal(float64(result)), nil
}
func fnDEGREES(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(n * 180 / math.Pi), nil
}
func fnCOS(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(math.Cos(n)), nil
}
func fnCOSH(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	result := math.Cosh(n)
	if math.IsInf(result, 0) { return ErrorVal(ErrValNUM), nil }
	return NumberVal(result), nil
}
func fnCOT(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n == 0 { return ErrorVal(ErrValDIV0), nil }
	return NumberVal(1 / math.Tan(n)), nil
}
func fnCOTH(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n == 0 { return ErrorVal(ErrValDIV0), nil }
	return NumberVal(1 / math.Tanh(n)), nil
}
func fnCSC(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n == 0 { return ErrorVal(ErrValDIV0), nil }
	return NumberVal(1 / math.Sin(n)), nil
}
func fnCSCH(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n == 0 { return ErrorVal(ErrValDIV0), nil }
	return NumberVal(1 / math.Sinh(n)), nil
}
func fnSEC(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(1 / math.Cos(n)), nil
}
func fnSECH(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(1 / math.Cosh(n)), nil
}
func fnERF(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 { return ErrorVal(ErrValVALUE), nil }
	lower, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if len(args) == 1 {
		return NumberVal(math.Erf(lower)), nil
	}
	upper, e := CoerceNum(args[1]); if e != nil { return *e, nil }
	return NumberVal(math.Erf(upper) - math.Erf(lower)), nil
}
func fnERFPRECISE(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(math.Erf(n)), nil
}
func fnERFC(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(math.Erfc(n)), nil
}
func fnERFCPRECISE(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(math.Erfc(n)), nil
}
func fnEVEN(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n == 0 { return NumberVal(0), nil }
	if n > 0 { return NumberVal(math.Ceil(n/2) * 2), nil }
	return NumberVal(math.Floor(n/2) * 2), nil
}
func fnODD(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n == 0 { return NumberVal(1), nil }
	if n > 0 {
		n2 := math.Ceil(n); if int(n2)%2 == 0 { n2++ }
		return NumberVal(n2), nil
	}
	n2 := math.Floor(n); if int(n2)%2 == 0 { n2-- }
	return NumberVal(n2), nil
}
func fnEXP(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(math.Exp(n)), nil
}
func fnFACT(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	n = math.Trunc(n)
	if n < 0 { return ErrorVal(ErrValNUM), nil }
	result := 1.0
	for i := 2.0; i <= n; i++ { result *= i }
	return NumberVal(result), nil
}
func fnFACTDOUBLE(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	nf, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	n := int(math.Trunc(nf))
	if n < -1 { return ErrorVal(ErrValNUM), nil }
	if n <= 0 { return NumberVal(1), nil }
	result := 1.0
	for i := n; i >= 2; i -= 2 { result *= float64(i) }
	return NumberVal(result), nil
}
func fnGCD(args []Value) (Value, error) {
	if len(args) < 1 { return ErrorVal(ErrValVALUE), nil }
	result := int64(0)
	for _, arg := range args {
		n, e := CoerceNum(arg); if e != nil { return *e, nil }
		n = math.Trunc(n)
		if n < 0 { return ErrorVal(ErrValNUM), nil }
		v := int64(n)
		a, b := result, v
		for b != 0 { a, b = b, a%b }
		result = a
	}
	return NumberVal(float64(result)), nil
}
func fnLCM(args []Value) (Value, error) {
	if len(args) < 1 { return ErrorVal(ErrValVALUE), nil }
	result := int64(1); hasZero := false
	for _, arg := range args {
		n, e := CoerceNum(arg); if e != nil { return *e, nil }
		n = math.Trunc(n)
		if n < 0 { return ErrorVal(ErrValNUM), nil }
		v := int64(n)
		if v == 0 { hasZero = true; continue }
		a, b := result, v
		for b != 0 { a, b = b, a%b }
		result = result / a * v
	}
	if hasZero { return NumberVal(0), nil }
	return NumberVal(float64(result)), nil
}
func fnLN(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n <= 0 { return ErrorVal(ErrValNUM), nil }
	return NumberVal(math.Log(n)), nil
}
func fnLOG(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n <= 0 { return ErrorVal(ErrValNUM), nil }
	base := 10.0
	if len(args) == 2 {
		base, e = CoerceNum(args[1]); if e != nil { return *e, nil }
		if base <= 0 { return ErrorVal(ErrValNUM), nil }
		if base == 1 { return ErrorVal(ErrValDIV0), nil }
	}
	return NumberVal(math.Log(n) / math.Log(base)), nil
}
func fnLOG10(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n <= 0 { return ErrorVal(ErrValNUM), nil }
	return NumberVal(math.Log10(n)), nil
}
func fnMMULT(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Coerce both arguments to 2D arrays.
	a1 := args[0]
	a2 := args[1]

	// Convert scalar/non-array values to 1x1 arrays.
	if a1.Type != ValueArray {
		if a1.Type == ValueError {
			return a1, nil
		}
		a1 = Value{Type: ValueArray, Array: [][]Value{{a1}}}
	}
	if a2.Type != ValueArray {
		if a2.Type == ValueError {
			return a2, nil
		}
		a2 = Value{Type: ValueArray, Array: [][]Value{{a2}}}
	}

	m := len(a1.Array) // rows of array1
	if m == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	n := len(a1.Array[0]) // cols of array1
	if n == 0 {
		return ErrorVal(ErrValVALUE), nil
	}

	n2 := len(a2.Array) // rows of array2
	if n2 == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	p := len(a2.Array[0]) // cols of array2
	if p == 0 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Columns of array1 must equal rows of array2.
	if n != n2 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Validate all values are numeric and extract into float64 slices.
	mat1 := make([][]float64, m)
	for r := 0; r < m; r++ {
		if len(a1.Array[r]) != n {
			return ErrorVal(ErrValVALUE), nil
		}
		mat1[r] = make([]float64, n)
		for c := 0; c < n; c++ {
			v := a1.Array[r][c]
			if v.Type == ValueEmpty || v.Type == ValueString {
				return ErrorVal(ErrValVALUE), nil
			}
			num, e := CoerceNum(v)
			if e != nil {
				return *e, nil
			}
			mat1[r][c] = num
		}
	}

	mat2 := make([][]float64, n2)
	for r := 0; r < n2; r++ {
		if len(a2.Array[r]) != p {
			return ErrorVal(ErrValVALUE), nil
		}
		mat2[r] = make([]float64, p)
		for c := 0; c < p; c++ {
			v := a2.Array[r][c]
			if v.Type == ValueEmpty || v.Type == ValueString {
				return ErrorVal(ErrValVALUE), nil
			}
			num, e := CoerceNum(v)
			if e != nil {
				return *e, nil
			}
			mat2[r][c] = num
		}
	}

	// Compute matrix product: result[i][j] = sum(mat1[i][k] * mat2[k][j]).
	result := make([][]Value, m)
	for i := 0; i < m; i++ {
		row := make([]Value, p)
		for j := 0; j < p; j++ {
			var sum float64
			for k := 0; k < n; k++ {
				sum += mat1[i][k] * mat2[k][j]
			}
			row[j] = NumberVal(sum)
		}
		result[i] = row
	}

	return Value{Type: ValueArray, Array: result}, nil
}

func fnMROUND(args []Value) (Value, error) {
	if len(args) != 2 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	multiple, e := CoerceNum(args[1]); if e != nil { return *e, nil }
	if multiple == 0 { return NumberVal(0), nil }
	if (n > 0 && multiple < 0) || (n < 0 && multiple > 0) { return ErrorVal(ErrValNUM), nil }
	return NumberVal(math.Round(n/multiple) * multiple), nil
}
func fnMULTINOMIAL(args []Value) (Value, error) {
	if len(args) < 1 { return ErrorVal(ErrValVALUE), nil }
	sum := 0; vals := make([]int, 0, len(args))
	for _, arg := range args {
		n, e := CoerceNum(arg); if e != nil { return *e, nil }
		ni := int(math.Trunc(n))
		if ni < 0 { return ErrorVal(ErrValNUM), nil }
		vals = append(vals, ni); sum += ni
	}
	num := 1.0
	for i := 2; i <= sum; i++ { num *= float64(i) }
	den := 1.0
	for _, v := range vals { for i := 2; i <= v; i++ { den *= float64(i) } }
	return NumberVal(num / den), nil
}
func fnPERMUT(args []Value) (Value, error) {
	if len(args) != 2 { return ErrorVal(ErrValVALUE), nil }
	nf, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	kf, e2 := CoerceNum(args[1]); if e2 != nil { return *e2, nil }
	if nf <= 0 || kf < 0 { return ErrorVal(ErrValNUM), nil }
	n := int(nf); k := int(kf)
	if n < k { return ErrorVal(ErrValNUM), nil }
	result := 1.0
	for i := 0; i < k; i++ { result *= float64(n - i) }
	return NumberVal(result), nil
}
func fnPI(args []Value) (Value, error) {
	if len(args) != 0 { return ErrorVal(ErrValVALUE), nil }
	return NumberVal(math.Pi), nil
}
func fnPRODUCT(args []Value) (Value, error) {
	product := 1.0; count := 0
	if e := IterateNumeric(args, func(n float64) { product *= n; count++ }); e != nil { return *e, nil }
	if count == 0 { return NumberVal(0), nil }
	return NumberVal(product), nil
}
func fnQUOTIENT(args []Value) (Value, error) {
	if len(args) != 2 { return ErrorVal(ErrValVALUE), nil }
	num, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	den, e := CoerceNum(args[1]); if e != nil { return *e, nil }
	if den == 0 { return ErrorVal(ErrValDIV0), nil }
	return NumberVal(math.Trunc(num / den)), nil
}
func fnRADIANS(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(n * math.Pi / 180), nil
}
func fnSIGN(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	switch { case n > 0: return NumberVal(1), nil; case n < 0: return NumberVal(-1), nil }
	return NumberVal(0), nil
}
func fnSIN(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(math.Sin(n)), nil
}
func fnSINH(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	result := math.Sinh(n)
	if math.IsInf(result, 0) { return ErrorVal(ErrValNUM), nil }
	return NumberVal(result), nil
}
func fnSQRTPI(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n < 0 { return ErrorVal(ErrValNUM), nil }
	return NumberVal(math.Sqrt(n * math.Pi)), nil
}
func fnTRUNC(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	digits := 0.0
	if len(args) == 2 { digits, e = CoerceNum(args[1]); if e != nil { return *e, nil } }
	pow := math.Pow(10, math.Floor(digits))
	return NumberVal(math.Trunc(n*pow) / pow), nil
}

// fnSUBTOTALCtx implements the Excel SUBTOTAL function. It applies an aggregate
// function (SUM, AVERAGE, etc.) to one or more ranges, but excludes cells that
// themselves contain SUBTOTAL formulas to prevent double-counting of nested
// subtotals.
func fnSUBTOTALCtx(args []Value, ctx *EvalContext) (Value, error) {
	if len(args) < 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	fnNum, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	fn := int(fnNum)
	excludeAllHidden := false
	if fn >= 101 && fn <= 111 {
		fn -= 100
		excludeAllHidden = true
	}
	if fn < 1 || fn > 11 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Filter out cells containing SUBTOTAL formulas from array arguments.
	// For function numbers 101-111, also exclude ALL hidden rows.
	// For function numbers 1-11, exclude rows hidden by autoFilter only.
	filtered := subtotalFilterArgs(args[1:], ctx, excludeAllHidden)

	name := subtotalFuncNames[fn]
	id := LookupFunc(name)
	if id < 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return CallFunc(id, filtered, ctx)
}

// subtotalFilterArgs processes SUBTOTAL's range arguments. For each ValueArray
// that has a RangeOrigin (came from a worksheet range, not a literal), it
// replaces cells that contain SUBTOTAL formulas with empty values so the
// aggregate function skips them. When excludeAllHidden is true (function
// numbers 101-111), ALL hidden rows are excluded. When false (function numbers
// 1-11), only rows hidden by an active autoFilter are excluded.
func subtotalFilterArgs(args []Value, ctx *EvalContext, excludeAllHidden bool) []Value {
	// Obtain the SubtotalChecker and HiddenRowChecker from the context's resolver.
	var checker SubtotalChecker
	var hiddenChecker HiddenRowChecker
	if ctx != nil && ctx.Resolver != nil {
		if sc, ok := ctx.Resolver.(SubtotalChecker); ok {
			checker = sc
		}
		if hc, ok := ctx.Resolver.(HiddenRowChecker); ok {
			hiddenChecker = hc
		}
	}
	if checker == nil && hiddenChecker == nil {
		// No checker available — return args unchanged (no filtering possible).
		return args
	}

	var out []Value
	for i, arg := range args {
		if arg.Type != ValueArray || arg.RangeOrigin == nil {
			if out != nil {
				out[i] = arg
			}
			continue
		}
		origin := arg.RangeOrigin
		sheet := origin.Sheet
		if sheet == "" && ctx != nil {
			sheet = ctx.CurrentSheet
		}

		// Build a filtered copy of the array, replacing SUBTOTAL cells
		// and hidden-row cells with empty.
		var filtered [][]Value
		initFiltered := func(upTo int) {
			filtered = make([][]Value, len(arg.Array))
			for k := 0; k < upTo; k++ {
				filtered[k] = arg.Array[k]
			}
		}
		for ri, row := range arg.Array {
			rowNum := origin.FromRow + ri
			// Determine if this row should be excluded based on hidden state.
			var rowExcluded bool
			if hiddenChecker != nil {
				if excludeAllHidden {
					rowExcluded = hiddenChecker.IsRowHidden(sheet, rowNum)
				} else {
					rowExcluded = hiddenChecker.IsRowFilteredByAutoFilter(sheet, rowNum)
				}
			}
			if rowExcluded {
				if filtered == nil {
					initFiltered(ri)
				}
				filtered[ri] = make([]Value, len(row))
				continue
			}
			// Check individual cells for SUBTOTAL formulas.
			var filteredRow []Value
			if checker != nil {
				for ci := range row {
					colNum := origin.FromCol + ci
					if checker.IsSubtotalCell(sheet, colNum, rowNum) {
						if filteredRow == nil {
							filteredRow = make([]Value, len(row))
							copy(filteredRow, row)
						}
						filteredRow[ci] = EmptyVal()
					}
				}
			}
			if filteredRow != nil {
				if filtered == nil {
					initFiltered(ri)
				}
				filtered[ri] = filteredRow
			} else if filtered != nil {
				filtered[ri] = row
			}
		}
		if filtered != nil {
			if out == nil {
				out = make([]Value, len(args))
				copy(out[:i], args[:i])
			}
			out[i] = Value{Type: ValueArray, Array: filtered}
		} else if out != nil {
			out[i] = arg
		}
	}
	if out == nil {
		return args
	}
	return out
}

func fnSEQUENCE(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	rowsF, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	rows := int(math.Trunc(rowsF))
	cols := 1
	if len(args) >= 2 {
		colsF, e := CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		cols = int(math.Trunc(colsF))
	}
	start := 1.0
	if len(args) >= 3 {
		s, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		start = s
	}
	step := 1.0
	if len(args) >= 4 {
		s, e := CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
		step = s
	}
	if rows <= 0 || cols <= 0 {
		return ErrorVal(ErrValCALC), nil
	}
	// Single cell: return scalar number.
	if rows == 1 && cols == 1 {
		return NumberVal(start), nil
	}
	cur := start
	result := make([][]Value, rows)
	for r := 0; r < rows; r++ {
		row := make([]Value, cols)
		for c := 0; c < cols; c++ {
			row[c] = NumberVal(cur)
			cur += step
		}
		result[r] = row
	}
	return Value{Type: ValueArray, Array: result}, nil
}

func fnTAN(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(math.Tan(n)), nil
}
func fnTANH(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	return NumberVal(math.Tanh(n)), nil
}

const bitMaxVal = (1 << 48) - 1 // 2^48 - 1

func validateBitArg(v Value) (int64, *Value) {
	n, e := CoerceNum(v)
	if e != nil { return 0, e }
	if n < 0 || math.Floor(n) != n || n > bitMaxVal {
		ev := ErrorVal(ErrValNUM); return 0, &ev
	}
	return int64(n), nil
}

func fnBITAND(args []Value) (Value, error) {
	if len(args) != 2 { return ErrorVal(ErrValVALUE), nil }
	n1, e := validateBitArg(args[0]); if e != nil { return *e, nil }
	n2, e := validateBitArg(args[1]); if e != nil { return *e, nil }
	return NumberVal(float64(n1 & n2)), nil
}
func fnBITOR(args []Value) (Value, error) {
	if len(args) != 2 { return ErrorVal(ErrValVALUE), nil }
	n1, e := validateBitArg(args[0]); if e != nil { return *e, nil }
	n2, e := validateBitArg(args[1]); if e != nil { return *e, nil }
	return NumberVal(float64(n1 | n2)), nil
}
func fnBITXOR(args []Value) (Value, error) {
	if len(args) != 2 { return ErrorVal(ErrValVALUE), nil }
	n1, e := validateBitArg(args[0]); if e != nil { return *e, nil }
	n2, e := validateBitArg(args[1]); if e != nil { return *e, nil }
	return NumberVal(float64(n1 ^ n2)), nil
}
func fnBITLSHIFT(args []Value) (Value, error) {
	if len(args) != 2 { return ErrorVal(ErrValVALUE), nil }
	n, e := validateBitArg(args[0]); if e != nil { return *e, nil }
	sf, ce := CoerceNum(args[1]); if ce != nil { return *ce, nil }
	if math.Floor(sf) != sf { return ErrorVal(ErrValNUM), nil }
	shift := int(sf)
	var result int64
	if shift >= 0 {
		result = n << uint(shift)
	} else {
		result = n >> uint(-shift)
	}
	if result < 0 || result > bitMaxVal { return ErrorVal(ErrValNUM), nil }
	return NumberVal(float64(result)), nil
}

func fnSERIESSUM(args []Value) (Value, error) {
	if len(args) != 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	n, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	m, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	// Flatten coefficients from arg[3].
	var coeffs []Value
	if args[3].Type == ValueArray {
		for _, row := range args[3].Array {
			coeffs = append(coeffs, row...)
		}
	} else {
		coeffs = []Value{args[3]}
	}

	var sum float64
	for i, cv := range coeffs {
		c, ce := CoerceNum(cv)
		if ce != nil {
			return *ce, nil
		}
		exp := n + float64(i)*m
		// Excel returns #NUM! for 0^0 and 0^negative.
		if x == 0 && exp <= 0 {
			return ErrorVal(ErrValNUM), nil
		}
		sum += c * math.Pow(x, exp)
	}
	return NumberVal(sum), nil
}

func fnMDETERM(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	a := args[0]

	// Propagate errors.
	if a.Type == ValueError {
		return a, nil
	}

	// Coerce scalar to 1x1 array.
	if a.Type != ValueArray {
		n, e := CoerceNum(a)
		if e != nil {
			return *e, nil
		}
		return NumberVal(n), nil
	}

	rows := len(a.Array)
	if rows == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	cols := len(a.Array[0])
	if cols == 0 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Must be square.
	if rows != cols {
		return ErrorVal(ErrValVALUE), nil
	}

	n := rows

	// Extract numeric matrix, rejecting empty/string values.
	mat := make([][]float64, n)
	for r := 0; r < n; r++ {
		if len(a.Array[r]) != cols {
			return ErrorVal(ErrValVALUE), nil
		}
		mat[r] = make([]float64, n)
		for c := 0; c < n; c++ {
			v := a.Array[r][c]
			if v.Type == ValueEmpty || v.Type == ValueString {
				return ErrorVal(ErrValVALUE), nil
			}
			num, e := CoerceNum(v)
			if e != nil {
				return *e, nil
			}
			mat[r][c] = num
		}
	}

	// Compute determinant via LU decomposition with partial pivoting.
	det := luDet(mat, n)
	return NumberVal(det), nil
}

// luDet computes the determinant of an n×n matrix using LU decomposition
// with partial pivoting.
func luDet(mat [][]float64, n int) float64 {
	if n == 1 {
		return mat[0][0]
	}
	if n == 2 {
		return mat[0][0]*mat[1][1] - mat[0][1]*mat[1][0]
	}

	// Work on a copy so we don't mutate the input.
	a := make([][]float64, n)
	for i := range a {
		a[i] = make([]float64, n)
		copy(a[i], mat[i])
	}

	sign := 1.0
	for col := 0; col < n; col++ {
		// Partial pivoting: find the row with the largest absolute value.
		maxVal := math.Abs(a[col][col])
		maxRow := col
		for row := col + 1; row < n; row++ {
			if v := math.Abs(a[row][col]); v > maxVal {
				maxVal = v
				maxRow = row
			}
		}
		if maxVal == 0 {
			return 0 // Singular matrix.
		}
		if maxRow != col {
			a[col], a[maxRow] = a[maxRow], a[col]
			sign = -sign
		}
		pivot := a[col][col]
		for row := col + 1; row < n; row++ {
			factor := a[row][col] / pivot
			for j := col + 1; j < n; j++ {
				a[row][j] -= factor * a[col][j]
			}
			a[row][col] = 0
		}
	}

	det := sign
	for i := 0; i < n; i++ {
		det *= a[i][i]
	}
	return det
}

func fnMINVERSE(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	a := args[0]

	// Propagate errors.
	if a.Type == ValueError {
		return a, nil
	}

	// Coerce scalar to 1x1 array.
	if a.Type != ValueArray {
		n, e := CoerceNum(a)
		if e != nil {
			return *e, nil
		}
		if n == 0 {
			return ErrorVal(ErrValNUM), nil
		}
		return NumberVal(1 / n), nil
	}

	rows := len(a.Array)
	if rows == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	cols := len(a.Array[0])
	if cols == 0 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Must be square.
	if rows != cols {
		return ErrorVal(ErrValVALUE), nil
	}

	n := rows

	// Extract numeric matrix, rejecting empty/string values.
	mat := make([][]float64, n)
	for r := 0; r < n; r++ {
		if len(a.Array[r]) != cols {
			return ErrorVal(ErrValVALUE), nil
		}
		mat[r] = make([]float64, n)
		for c := 0; c < n; c++ {
			v := a.Array[r][c]
			if v.Type == ValueEmpty || v.Type == ValueString {
				return ErrorVal(ErrValVALUE), nil
			}
			num, e := CoerceNum(v)
			if e != nil {
				return *e, nil
			}
			mat[r][c] = num
		}
	}

	// Compute inverse using Gauss-Jordan elimination with partial pivoting.
	// Augment with identity matrix.
	aug := make([][]float64, n)
	for i := 0; i < n; i++ {
		aug[i] = make([]float64, 2*n)
		copy(aug[i], mat[i])
		aug[i][n+i] = 1
	}

	for col := 0; col < n; col++ {
		// Partial pivoting: find the row with the largest absolute value.
		maxVal := math.Abs(aug[col][col])
		maxRow := col
		for row := col + 1; row < n; row++ {
			if v := math.Abs(aug[row][col]); v > maxVal {
				maxVal = v
				maxRow = row
			}
		}
		if maxVal < 1e-15 {
			return ErrorVal(ErrValNUM), nil // Singular matrix.
		}
		if maxRow != col {
			aug[col], aug[maxRow] = aug[maxRow], aug[col]
		}

		// Scale pivot row.
		pivot := aug[col][col]
		for j := 0; j < 2*n; j++ {
			aug[col][j] /= pivot
		}

		// Eliminate column in all other rows.
		for row := 0; row < n; row++ {
			if row == col {
				continue
			}
			factor := aug[row][col]
			for j := 0; j < 2*n; j++ {
				aug[row][j] -= factor * aug[col][j]
			}
		}
	}

	// Extract inverse from augmented matrix.
	result := make([][]Value, n)
	for i := 0; i < n; i++ {
		row := make([]Value, n)
		for j := 0; j < n; j++ {
			row[j] = NumberVal(aug[i][n+j])
		}
		result[i] = row
	}

	return Value{Type: ValueArray, Array: result}, nil
}

func fnMUNIT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	dim := int(n) // truncate toward zero
	if dim <= 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	result := make([][]Value, dim)
	for i := 0; i < dim; i++ {
		row := make([]Value, dim)
		for j := 0; j < dim; j++ {
			if i == j {
				row[j] = NumberVal(1)
			} else {
				row[j] = NumberVal(0)
			}
		}
		result[i] = row
	}
	return Value{Type: ValueArray, Array: result}, nil
}

func fnBITRSHIFT(args []Value) (Value, error) {
	if len(args) != 2 { return ErrorVal(ErrValVALUE), nil }
	n, e := validateBitArg(args[0]); if e != nil { return *e, nil }
	sf, ce := CoerceNum(args[1]); if ce != nil { return *ce, nil }
	if math.Floor(sf) != sf { return ErrorVal(ErrValNUM), nil }
	shift := int(sf)
	var result int64
	if shift >= 0 {
		result = n >> uint(shift)
	} else {
		result = n << uint(-shift)
	}
	if result < 0 || result > bitMaxVal { return ErrorVal(ErrValNUM), nil }
	return NumberVal(float64(result)), nil
}
