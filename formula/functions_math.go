package formula

import (
	"math"
	"math/rand"
	"strconv"
	"strings"
)

func init() {
	Register("ABS", NoCtx(fnABS))
	Register("ACOS", NoCtx(fnACOS))
	Register("ACOSH", NoCtx(fnACOSH))
	Register("ARABIC", NoCtx(fnARABIC))
	Register("ASIN", NoCtx(fnASIN))
	Register("ASINH", NoCtx(fnASINH))
	Register("ATAN", NoCtx(fnATAN))
	Register("ATAN2", NoCtx(fnATAN2))
	Register("ATANH", NoCtx(fnATANH))
	Register("BASE", NoCtx(fnBASE))
	Register("CEILING", NoCtx(fnCEILING))
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
	Register("EVEN", NoCtx(fnEVEN))
	Register("EXP", NoCtx(fnEXP))
	Register("FACT", NoCtx(fnFACT))
	Register("FACTDOUBLE", NoCtx(fnFACTDOUBLE))
	Register("FLOOR", NoCtx(fnFLOOR))
	Register("GCD", NoCtx(fnGCD))
	Register("INT", NoCtx(fnINT))
	Register("LCM", NoCtx(fnLCM))
	Register("LN", NoCtx(fnLN))
	Register("LOG", NoCtx(fnLOG))
	Register("LOG10", NoCtx(fnLOG10))
	Register("MOD", NoCtx(fnMOD))
	Register("MROUND", NoCtx(fnMROUND))
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
		if base <= 0 || base == 1 { return ErrorVal(ErrValNUM), nil }
	}
	return NumberVal(math.Log(n) / math.Log(base)), nil
}
func fnLOG10(args []Value) (Value, error) {
	if len(args) != 1 { return ErrorVal(ErrValVALUE), nil }
	n, e := CoerceNum(args[0]); if e != nil { return *e, nil }
	if n <= 0 { return ErrorVal(ErrValNUM), nil }
	return NumberVal(math.Log10(n)), nil
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
	if fn >= 101 && fn <= 111 {
		fn -= 100
	}
	if fn < 1 || fn > 11 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Filter out cells containing SUBTOTAL formulas from array arguments.
	filtered := subtotalFilterArgs(args[1:], ctx)

	names := map[int]string{
		1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN",
		6: "PRODUCT", 7: "STDEV", 8: "STDEVP", 9: "SUM", 10: "VAR", 11: "VARP",
	}
	name := names[fn]
	id := LookupFunc(name)
	if id < 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return CallFunc(id, filtered, ctx)
}

// subtotalFilterArgs processes SUBTOTAL's range arguments. For each ValueArray
// that has a RangeOrigin (came from a worksheet range, not a literal), it
// replaces cells that contain SUBTOTAL formulas with empty values so the
// aggregate function skips them.
func subtotalFilterArgs(args []Value, ctx *EvalContext) []Value {
	// Obtain the SubtotalChecker from the context's resolver.
	var checker SubtotalChecker
	if ctx != nil && ctx.Resolver != nil {
		if sc, ok := ctx.Resolver.(SubtotalChecker); ok {
			checker = sc
		}
	}
	if checker == nil {
		// No checker available — return args unchanged (no filtering possible).
		return args
	}

	out := make([]Value, len(args))
	for i, arg := range args {
		if arg.Type != ValueArray || arg.RangeOrigin == nil {
			out[i] = arg
			continue
		}
		origin := arg.RangeOrigin
		sheet := origin.Sheet
		if sheet == "" && ctx != nil {
			sheet = ctx.CurrentSheet
		}

		// Build a filtered copy of the array, replacing SUBTOTAL cells with empty.
		filtered := make([][]Value, len(arg.Array))
		changed := false
		for ri, row := range arg.Array {
			rowNum := origin.FromRow + ri
			filteredRow := make([]Value, len(row))
			copy(filteredRow, row)
			for ci := range row {
				colNum := origin.FromCol + ci
				if checker.IsSubtotalCell(sheet, colNum, rowNum) {
					filteredRow[ci] = EmptyVal()
					changed = true
				}
			}
			filtered[ri] = filteredRow
		}
		if changed {
			out[i] = Value{Type: ValueArray, Array: filtered}
		} else {
			out[i] = arg
		}
	}
	return out
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
