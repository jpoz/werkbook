package formula

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

func init() {
	Register("BESSELI", NoCtx(fnBesselI))
	Register("BESSELJ", NoCtx(fnBesselJ))
	Register("BESSELK", NoCtx(fnBesselK))
	Register("BESSELY", NoCtx(fnBesselY))
	Register("BIN2DEC", NoCtx(fnBin2Dec))
	Register("BIN2HEX", NoCtx(fnBin2Hex))
	Register("BIN2OCT", NoCtx(fnBin2Oct))
	Register("COMPLEX", NoCtx(fnComplex))
	Register("CONVERT", NoCtx(fnConvert))
	Register("DELTA", NoCtx(fnDELTA))
	Register("DEC2BIN", NoCtx(fnDec2Bin))
	Register("DEC2HEX", NoCtx(fnDec2Hex))
	Register("DEC2OCT", NoCtx(fnDec2Oct))
	Register("GESTEP", NoCtx(fnGESTEP))
	Register("HEX2BIN", NoCtx(fnHex2Bin))
	Register("HEX2DEC", NoCtx(fnHex2Dec))
	Register("HEX2OCT", NoCtx(fnHex2Oct))
	Register("IMABS", NoCtx(fnImabs))
	Register("IMAGINARY", NoCtx(fnImaginary))
	Register("IMARGUMENT", NoCtx(fnImargument))
	Register("IMCONJUGATE", NoCtx(fnImconjugate))
	Register("IMCOS", NoCtx(fnImcos))
	Register("IMDIV", NoCtx(fnImdiv))
	Register("IMEXP", NoCtx(fnImexp))
	Register("IMLN", NoCtx(fnImln))
	Register("IMLOG10", NoCtx(fnImlog10))
	Register("IMLOG2", NoCtx(fnImlog2))
	Register("IMPOWER", NoCtx(fnImpower))
	Register("IMPRODUCT", NoCtx(fnImproduct))
	Register("IMREAL", NoCtx(fnImreal))
	Register("IMCOSH", NoCtx(fnImcosh))
	Register("IMCOT", NoCtx(fnImcot))
	Register("IMCSC", NoCtx(fnImcsc))
	Register("IMCSCH", NoCtx(fnImcsch))
	Register("IMSEC", NoCtx(fnImsec))
	Register("IMSECH", NoCtx(fnImsech))
	Register("IMSIN", NoCtx(fnImsin))
	Register("IMSINH", NoCtx(fnImsinh))
	Register("IMSQRT", NoCtx(fnImsqrt))
	Register("IMSUB", NoCtx(fnImsub))
	Register("IMSUM", NoCtx(fnImsum))
	Register("IMTAN", NoCtx(fnImtan))
	Register("OCT2BIN", NoCtx(fnOct2Bin))
	Register("OCT2DEC", NoCtx(fnOct2Dec))
	Register("OCT2HEX", NoCtx(fnOct2Hex))
}

// fnBesselI implements the BESSELI function.
// BESSELI(X, N) — returns the modified Bessel function of the first kind, I_n(x).
// N is truncated to an integer. If N < 0, returns #NUM!.
func fnBesselI(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValNA), nil
	}

	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	nf, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	// Truncate n to integer.
	n := int(math.Trunc(nf))
	if n < 0 {
		return ErrorVal(ErrValNUM), nil
	}

	// Handle x = 0 specially.
	if x == 0 {
		if n == 0 {
			return NumberVal(1), nil
		}
		return NumberVal(0), nil
	}

	// Handle negative x: I_n(-x) = (-1)^n * I_n(x) for integer n.
	sign := 1.0
	if x < 0 {
		if n%2 != 0 {
			sign = -1
		}
		x = -x
	}

	result := besselI(n, x)
	return NumberVal(sign * result), nil
}

// besselI0 computes I_0(x) for x >= 0 using polynomial approximations
// from Abramowitz & Stegun (sections 9.8.1 and 9.8.2).
func besselI0(x float64) float64 {
	if x <= 3.75 {
		t := x / 3.75
		t2 := t * t
		return 1.0 +
			t2*(3.5156229+
				t2*(3.0899424+
					t2*(1.2067492+
						t2*(0.2659732+
							t2*(0.0360768+
								t2*0.0045813)))))
	}
	t := 3.75 / x
	return (math.Exp(x) / math.Sqrt(x)) *
		(0.39894228 +
			t*(0.01328592+
				t*(0.00225319+
					t*(-0.00157565+
						t*(0.00916281+
							t*(-0.02057706+
								t*(0.02635537+
									t*(-0.01647633+
										t*0.00392377))))))))
}

// besselI1 computes I_1(x) for x >= 0 using polynomial approximations
// from Abramowitz & Stegun (sections 9.8.3 and 9.8.4).
func besselI1(x float64) float64 {
	if x <= 3.75 {
		t := x / 3.75
		t2 := t * t
		return x * (0.5 +
			t2*(0.87890594+
				t2*(0.51498869+
					t2*(0.15084934+
						t2*(0.02658733+
							t2*(0.00301532+
								t2*0.00032411))))))
	}
	t := 3.75 / x
	return (math.Exp(x) / math.Sqrt(x)) *
		(0.39894228 +
			t*(-0.03988024+
				t*(-0.00362018+
					t*(0.00163801+
						t*(-0.01031555+
							t*(0.02282967+
								t*(-0.02895312+
									t*(0.01787654+
										t*(-0.00420059)))))))))
}

// besselI computes I_n(x) for integer n >= 0 and x > 0.
// For n=0 and n=1 it uses direct polynomial approximations (A&S).
// For n >= 2 it uses forward recurrence from I_0 and I_1 when stable,
// or Miller's backward recurrence otherwise.
func besselI(n int, x float64) float64 {
	if n == 0 {
		return besselI0(x)
	}
	if n == 1 {
		return besselI1(x)
	}

	if x == 0 {
		return 0
	}

	// For I_n(x), forward recurrence from I_0,I_1 is stable when x > n
	// because the ratio I_{n+1}/I_n < 1. For x <= n it is unstable
	// and we must use backward recurrence.
	//
	// Forward recurrence: I_{k+1}(x) = I_{k-1}(x) - (2k/x)*I_k(x)
	// Note: this is actually I_{k+1} = -(2k/x)*I_k + I_{k-1}
	// The standard recurrence is: I_{n-1}(x) - (2n/x)*I_n(x) = I_{n+1}(x)
	// Rearranged for forward: I_{n+1}(x) = I_{n-1}(x) - (2n/x)*I_n(x)
	// But actually for modified Bessel: I_{n+1}(x) = -(2n/x)*I_n(x) + I_{n-1}(x)
	// is UNSTABLE in the forward direction for I.
	//
	// The correct recurrence for modified Bessel of 1st kind is:
	//   I_{n-1}(x) - I_{n+1}(x) = (2n/x)*I_n(x)
	// or equivalently:
	//   I_{n+1}(x) = I_{n-1}(x) - (2n/x)*I_n(x)
	//
	// For modified Bessel I, forward recurrence is UNSTABLE for all x
	// (I_n grows, not decreases). Use Miller's backward recurrence.

	bi0 := besselI0(x)

	// Miller's backward recurrence with higher accuracy.
	// Start from a sufficiently high order where I_m ~ 0.
	const iacc = 40
	bigno := 1e10
	bigni := 1e-10

	// Starting index: must be well above n and large enough for convergence.
	m := 2 * ((n + int(math.Sqrt(float64(iacc*n)))) / 2)
	if m < 40 {
		m = 40
	}
	// For large x, we need even more terms.
	if extra := int(x) + 10; m < n+extra {
		m = n + extra
		// Round up to even.
		if m%2 != 0 {
			m++
		}
	}

	var bip, bi, bim float64
	var result float64
	tox := 2.0 / x

	bip = 0
	bi = 1.0
	for j := m; j >= 1; j-- {
		bim = bip + float64(j)*tox*bi
		bip = bi
		bi = bim
		// Renormalize to prevent overflow.
		if math.Abs(bi) > bigno {
			result *= bigni
			bi *= bigni
			bip *= bigni
		}
		if j == n {
			result = bip
		}
	}
	// Normalize: at j=0 we have bi ≈ I_0(x) * scale, so result/bi * I_0(x) gives I_n(x).
	result *= bi0 / bi
	return result
}

// fnBesselJ implements the BESSELJ function.
// BESSELJ(X, N) — returns the Bessel function of the first kind, J_n(x).
// N is truncated to an integer. If N < 0, returns #NUM!.
func fnBesselJ(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValNA), nil
	}

	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	nf, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	// Truncate n to integer.
	n := int(math.Trunc(nf))
	if n < 0 {
		return ErrorVal(ErrValNUM), nil
	}

	return NumberVal(math.Jn(n, x)), nil
}

// fnBesselK implements the BESSELK function.
// BESSELK(X, N) — returns the modified Bessel function of the second kind, K_n(x).
// N is truncated to an integer. If N < 0, returns #NUM!. If X <= 0, returns #NUM!.
func fnBesselK(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValNA), nil
	}

	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	nf, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	// Truncate n to integer.
	n := int(math.Trunc(nf))
	if n < 0 {
		return ErrorVal(ErrValNUM), nil
	}

	// K_n(x) is undefined for x <= 0.
	if x <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	result := besselK(n, x)
	return NumberVal(result), nil
}

// besselK0 computes K_0(x) for x > 0 using polynomial approximations
// from Abramowitz & Stegun (sections 9.8.5 and 9.8.6).
func besselK0(x float64) float64 {
	if x <= 2 {
		t := x / 2
		t2 := t * t
		return -math.Log(t)*besselI0(x) +
			(-0.57721566 +
				t2*(0.42278420+
					t2*(0.23069756+
						t2*(0.03488590+
							t2*(0.00262698+
								t2*(0.00010750+
									t2*0.00000740))))))
	}
	t := 2.0 / x
	return (math.Exp(-x) / math.Sqrt(x)) *
		(1.25331414 +
			t*(-0.07832358+
				t*(0.02189568+
					t*(-0.01062446+
						t*(0.00587872+
							t*(-0.00251540+
								t*0.00053208))))))
}

// besselK1 computes K_1(x) for x > 0 using polynomial approximations
// from Abramowitz & Stegun (sections 9.8.7 and 9.8.8).
func besselK1(x float64) float64 {
	if x <= 2 {
		t := x / 2
		t2 := t * t
		return math.Log(t)*besselI1(x) +
			(1.0/x)*
				(1.0+
					t2*(0.15443144+
						t2*(-0.67278579+
							t2*(-0.18156897+
								t2*(-0.01919402+
									t2*(-0.00110404+
										t2*(-0.00004686)))))))
	}
	t := 2.0 / x
	return (math.Exp(-x) / math.Sqrt(x)) *
		(1.25331414 +
			t*(0.23498619+
				t*(-0.03655620+
					t*(0.01504268+
						t*(-0.00780353+
							t*(0.00325614+
								t*(-0.00068245)))))))
}

// besselK computes K_n(x) for integer n >= 0 and x > 0.
// For n=0 and n=1 it uses direct polynomial approximations (A&S).
// For n >= 2 it uses forward recurrence: K_{n+1}(x) = K_{n-1}(x) + (2n/x)*K_n(x).
func besselK(n int, x float64) float64 {
	if n == 0 {
		return besselK0(x)
	}
	if n == 1 {
		return besselK1(x)
	}

	// Forward recurrence is stable for K_n (K_n grows with n).
	k0 := besselK0(x)
	k1 := besselK1(x)
	tox := 2.0 / x
	for i := 1; i < n; i++ {
		k2 := k0 + float64(i)*tox*k1
		k0 = k1
		k1 = k2
	}
	return k1
}

// fnBesselY implements the BESSELY function.
// BESSELY(X, N) — returns the Bessel function of the second kind, Y_n(x).
// N is truncated to an integer. If N < 0, returns #NUM!. If X <= 0, returns #NUM!.
func fnBesselY(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValNA), nil
	}

	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	nf, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	// Truncate n to integer.
	n := int(math.Trunc(nf))
	if n < 0 {
		return ErrorVal(ErrValNUM), nil
	}

	// Y_n(x) is undefined for x <= 0.
	if x <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	result := math.Yn(n, x)
	if math.IsInf(result, 0) || math.IsNaN(result) {
		return ErrorVal(ErrValNUM), nil
	}

	return NumberVal(result), nil
}

// fnBin2Dec implements the BIN2DEC function.
// BIN2DEC(number) — converts a binary number string to decimal.
// Input must contain only 0s and 1s, max 10 digits.
// 10-digit numbers starting with 1 are negative (two's complement).
func fnBin2Dec(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValNA), nil
	}

	// Engineering functions reject bare booleans with #VALUE!.
	if args[0].Type == ValueBool {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Coerce input to string.
	var s string
	switch args[0].Type {
	case ValueNumber:
		// Format as integer string (e.g., 1001.0 → "1001").
		s = strconv.FormatInt(int64(math.Trunc(args[0].Num)), 10)
	case ValueString:
		s = strings.TrimSpace(args[0].Str)
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Validate length: max 10 binary digits.
	if len(s) == 0 || len(s) > 10 {
		return ErrorVal(ErrValNUM), nil
	}

	// Validate characters: only 0 and 1 allowed.
	for _, c := range s {
		if c != '0' && c != '1' {
			return ErrorVal(ErrValNUM), nil
		}
	}

	// Parse as unsigned binary.
	v, err := strconv.ParseUint(s, 2, 64)
	if err != nil {
		return ErrorVal(ErrValNUM), nil
	}

	// Two's complement: 10-digit number starting with '1' is negative.
	var result float64
	if len(s) == 10 && s[0] == '1' {
		result = float64(int64(v) - 1024)
	} else {
		result = float64(v)
	}

	return NumberVal(result), nil
}

// parseBinToInt64 parses a binary string (max 10 digits, 0/1 only) to int64,
// handling two's complement for 10-digit numbers starting with 1.
// Returns the parsed value and an error Value if validation fails.
func parseBinToInt64(args []Value) (int64, *Value) {
	// Engineering functions reject bare booleans with #VALUE!.
	if args[0].Type == ValueBool {
		v := ErrorVal(ErrValVALUE)
		return 0, &v
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return 0, &args[0]
	}

	// Coerce input to string.
	var s string
	switch args[0].Type {
	case ValueNumber:
		s = strconv.FormatInt(int64(math.Trunc(args[0].Num)), 10)
	case ValueString:
		s = strings.TrimSpace(args[0].Str)
	default:
		v := ErrorVal(ErrValVALUE)
		return 0, &v
	}

	// Validate length: max 10 binary digits.
	if len(s) == 0 || len(s) > 10 {
		v := ErrorVal(ErrValNUM)
		return 0, &v
	}

	// Validate characters: only 0 and 1 allowed.
	for _, c := range s {
		if c != '0' && c != '1' {
			v := ErrorVal(ErrValNUM)
			return 0, &v
		}
	}

	// Parse as unsigned binary.
	u, err := strconv.ParseUint(s, 2, 64)
	if err != nil {
		v := ErrorVal(ErrValNUM)
		return 0, &v
	}

	// Two's complement: 10-digit number starting with '1' is negative.
	if len(s) == 10 && s[0] == '1' {
		return int64(u) - 1024, nil
	}
	return int64(u), nil
}

// fnBin2Hex implements the BIN2HEX function.
// BIN2HEX(number, [places]) — converts a binary number string to hexadecimal.
// Input must contain only 0s and 1s, max 10 digits.
// 10-digit numbers starting with 1 are negative (two's complement).
// Negative numbers produce 10-digit hex result (40-bit two's complement).
func fnBin2Hex(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	n, errVal := parseBinToInt64(args)
	if errVal != nil {
		return *errVal, nil
	}

	var result string
	if n < 0 {
		// Two's complement: add 2^40 to get 10-digit hex representation.
		result = fmt.Sprintf("%X", n+1099511627776)
	} else {
		result = fmt.Sprintf("%X", n)
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

// fnBin2Oct implements the BIN2OCT function.
// BIN2OCT(number, [places]) — converts a binary number string to octal.
// Input must contain only 0s and 1s, max 10 digits.
// 10-digit numbers starting with 1 are negative (two's complement).
// Negative numbers produce 10-digit octal result (30-bit two's complement).
func fnBin2Oct(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	n, errVal := parseBinToInt64(args)
	if errVal != nil {
		return *errVal, nil
	}

	var result string
	if n < 0 {
		// Two's complement: add 2^30 to get 10-digit octal representation.
		result = fmt.Sprintf("%o", n+1073741824)
	} else {
		result = fmt.Sprintf("%o", n)
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

// formatComplexNum formats a float64 for use in COMPLEX output.
// Integers display without decimals (e.g. 3, not 3.0).
func formatComplexNum(f float64) string {
	if f == math.Trunc(f) && !math.IsInf(f, 0) && !math.IsNaN(f) {
		return strconv.FormatFloat(f, 'f', 0, 64)
	}
	return strconv.FormatFloat(f, 'f', -1, 64)
}

// fnComplex implements the COMPLEX function.
// COMPLEX(real_num, i_num, [suffix]) — converts real and imaginary
// coefficients into a complex number string of the form x+yi or x+yj.
func fnComplex(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}

	realNum, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	iNum, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	suffix := "i"
	if len(args) == 3 {
		// Propagate errors from 3rd arg.
		if args[2].Type == ValueError {
			return args[2], nil
		}
		switch args[2].Type {
		case ValueString:
			suffix = args[2].Str
		case ValueBool:
			// Booleans not accepted as suffix.
			return ErrorVal(ErrValVALUE), nil
		default:
			return ErrorVal(ErrValVALUE), nil
		}
		if suffix != "i" && suffix != "j" {
			return ErrorVal(ErrValVALUE), nil
		}
	}

	// Both zero: just "0".
	if realNum == 0 && iNum == 0 {
		return StringVal("0"), nil
	}

	var result string

	// Build real part.
	if realNum != 0 {
		result = formatComplexNum(realNum)
	}

	// Build imaginary part.
	if iNum != 0 {
		if realNum != 0 {
			// Need a sign separator.
			if iNum > 0 {
				result += "+"
			}
			// For iNum < 0, the minus sign comes from formatting.
		}

		if iNum == 1 {
			result += suffix
		} else if iNum == -1 {
			result += "-" + suffix
		} else {
			result += formatComplexNum(iNum) + suffix
		}
	}

	return StringVal(result), nil
}

// fnDELTA implements the DELTA function.
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

// fnDec2Bin implements the DEC2BIN function.
// DEC2BIN(number, [places]) — converts a decimal number to binary.
// number must be between -512 and 511 (inclusive). Non-integer values are truncated.
// Negative numbers use two's complement (10-bit). places specifies minimum digits (1–10).
func fnDec2Bin(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Engineering functions reject bare booleans with #VALUE!.
	if args[0].Type == ValueBool {
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

// fnDec2Hex implements the DEC2HEX function.
// DEC2HEX(number, [places]) — converts a decimal number to hexadecimal.
// number must be between -549755813888 and 549755813887 (inclusive, -2^39 to 2^39-1).
// Non-integer values are truncated. Negative numbers use two's complement (10 hex digits = 40 bits).
// places specifies minimum digits (1–10).
func fnDec2Hex(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Engineering functions reject bare booleans with #VALUE!.
	if args[0].Type == ValueBool {
		return ErrorVal(ErrValVALUE), nil
	}

	num, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	// Truncate to integer.
	num = math.Trunc(num)

	// Range check: -549755813888 to 549755813887.
	if num < -549755813888 || num > 549755813887 {
		return ErrorVal(ErrValNUM), nil
	}

	n := int64(num)

	var result string
	if n < 0 {
		// Two's complement: add 2^40 (1099511627776) to get 10-digit hex representation.
		result = fmt.Sprintf("%X", n+1099511627776)
	} else {
		result = fmt.Sprintf("%X", n)
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

// fnDec2Oct implements the DEC2OCT function.
// DEC2OCT(number, [places]) — converts a decimal number to octal.
// number must be between -536870912 and 536870911 (inclusive, -2^29 to 2^29-1).
// Non-integer values are truncated. Negative numbers use two's complement (10 octal digits = 30 bits).
// places specifies minimum digits (1–10).
func fnDec2Oct(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Engineering functions reject bare booleans with #VALUE!.
	if args[0].Type == ValueBool {
		return ErrorVal(ErrValVALUE), nil
	}

	num, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	// Truncate to integer.
	num = math.Trunc(num)

	// Range check: -536870912 to 536870911.
	if num < -536870912 || num > 536870911 {
		return ErrorVal(ErrValNUM), nil
	}

	n := int64(num)

	var result string
	if n < 0 {
		// Two's complement: add 2^30 (1073741824) to get 10-digit octal representation.
		result = fmt.Sprintf("%o", n+1073741824)
	} else {
		result = fmt.Sprintf("%o", n)
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

// fnHex2Dec implements the HEX2DEC function.
// HEX2DEC(number) — converts a hexadecimal number string to decimal.
// Input must contain only hex chars (0-9, A-F, a-f), max 10 digits.
// 10-digit numbers starting with 8-F are negative (two's complement).
func fnHex2Dec(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValNA), nil
	}

	// Engineering functions reject bare booleans with #VALUE!.
	if args[0].Type == ValueBool {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Coerce input to string.
	var s string
	switch args[0].Type {
	case ValueNumber:
		// Format as integer string (e.g., 100.0 → "100").
		s = strconv.FormatInt(int64(math.Trunc(args[0].Num)), 10)
	case ValueString:
		s = strings.TrimSpace(args[0].Str)
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Validate length: max 10 hex digits, not empty.
	if len(s) == 0 || len(s) > 10 {
		return ErrorVal(ErrValNUM), nil
	}

	// Validate characters: only hex chars allowed.
	for _, c := range s {
		if !((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f')) {
			return ErrorVal(ErrValNUM), nil
		}
	}

	// Parse as unsigned hex.
	v, err := strconv.ParseUint(s, 16, 64)
	if err != nil {
		return ErrorVal(ErrValNUM), nil
	}

	// Two's complement: 10-digit number with first digit >= 8 is negative.
	var result float64
	if len(s) == 10 && s[0] >= '8' {
		result = float64(int64(v) - 1099511627776) // subtract 2^40
	} else {
		result = float64(v)
	}

	return NumberVal(result), nil
}

// fnOct2Dec implements the OCT2DEC function.
// OCT2DEC(number) — converts an octal number string to decimal.
// Input must contain only octal chars (0-7), max 10 digits.
// 10-digit numbers starting with 4-7 are negative (two's complement, 30-bit).
func fnOct2Dec(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValNA), nil
	}

	// Engineering functions reject bare booleans with #VALUE!.
	if args[0].Type == ValueBool {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Coerce input to string.
	var s string
	switch args[0].Type {
	case ValueNumber:
		// Format as integer string (e.g., 144.0 → "144").
		s = strconv.FormatInt(int64(math.Trunc(args[0].Num)), 10)
	case ValueString:
		s = strings.TrimSpace(args[0].Str)
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Validate length: max 10 octal digits, not empty.
	if len(s) == 0 || len(s) > 10 {
		return ErrorVal(ErrValNUM), nil
	}

	// Validate characters: only 0-7 allowed.
	for _, c := range s {
		if c < '0' || c > '7' {
			return ErrorVal(ErrValNUM), nil
		}
	}

	// Parse as unsigned octal.
	v, err := strconv.ParseUint(s, 8, 64)
	if err != nil {
		return ErrorVal(ErrValNUM), nil
	}

	// Two's complement: 10-digit number with first digit >= 4 is negative.
	var result float64
	if len(s) == 10 && s[0] >= '4' {
		result = float64(int64(v) - 1073741824) // subtract 2^30
	} else {
		result = float64(v)
	}

	return NumberVal(result), nil
}

// parseHexToInt64 parses a hex string (max 10 digits, 0-9/A-F/a-f only) to int64,
// handling two's complement for 10-digit numbers starting with 8-F.
// Returns the parsed value and an error Value if validation fails.
func parseHexToInt64(args []Value) (int64, *Value) {
	// Engineering functions reject bare booleans with #VALUE!.
	if args[0].Type == ValueBool {
		v := ErrorVal(ErrValVALUE)
		return 0, &v
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return 0, &args[0]
	}

	// Coerce input to string.
	var s string
	switch args[0].Type {
	case ValueNumber:
		s = strconv.FormatInt(int64(math.Trunc(args[0].Num)), 10)
	case ValueString:
		s = strings.TrimSpace(args[0].Str)
	default:
		v := ErrorVal(ErrValVALUE)
		return 0, &v
	}

	// Validate length: max 10 hex digits, not empty.
	if len(s) == 0 || len(s) > 10 {
		v := ErrorVal(ErrValNUM)
		return 0, &v
	}

	// Validate characters: only hex chars allowed.
	for _, c := range s {
		if !((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f')) {
			v := ErrorVal(ErrValNUM)
			return 0, &v
		}
	}

	// Parse as unsigned hex.
	u, err := strconv.ParseUint(s, 16, 64)
	if err != nil {
		v := ErrorVal(ErrValNUM)
		return 0, &v
	}

	// Two's complement: 10-digit number with first digit >= 8 is negative.
	if len(s) == 10 && s[0] >= '8' {
		return int64(u) - 1099511627776, nil // subtract 2^40
	}
	return int64(u), nil
}

// parseOctToInt64 parses an octal string (max 10 digits, 0-7 only) to int64,
// handling two's complement for 10-digit numbers starting with 4-7.
// Returns the parsed value and an error Value if validation fails.
func parseOctToInt64(args []Value) (int64, *Value) {
	// Engineering functions reject bare booleans with #VALUE!.
	if args[0].Type == ValueBool {
		v := ErrorVal(ErrValVALUE)
		return 0, &v
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return 0, &args[0]
	}

	// Coerce input to string.
	var s string
	switch args[0].Type {
	case ValueNumber:
		s = strconv.FormatInt(int64(math.Trunc(args[0].Num)), 10)
	case ValueString:
		s = strings.TrimSpace(args[0].Str)
	default:
		v := ErrorVal(ErrValVALUE)
		return 0, &v
	}

	// Validate length: max 10 octal digits, not empty.
	if len(s) == 0 || len(s) > 10 {
		v := ErrorVal(ErrValNUM)
		return 0, &v
	}

	// Validate characters: only 0-7 allowed.
	for _, c := range s {
		if c < '0' || c > '7' {
			v := ErrorVal(ErrValNUM)
			return 0, &v
		}
	}

	// Parse as unsigned octal.
	u, err := strconv.ParseUint(s, 8, 64)
	if err != nil {
		v := ErrorVal(ErrValNUM)
		return 0, &v
	}

	// Two's complement: 10-digit number with first digit >= 4 is negative.
	if len(s) == 10 && s[0] >= '4' {
		return int64(u) - 1073741824, nil // subtract 2^30
	}
	return int64(u), nil
}

// fnHex2Bin implements the HEX2BIN function.
// HEX2BIN(number, [places]) — converts a hexadecimal number string to binary.
// Input must contain only hex chars (0-9, A-F, a-f), max 10 digits.
// Output must be in range -512 to 511 (10-bit binary).
func fnHex2Bin(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	n, errVal := parseHexToInt64(args)
	if errVal != nil {
		return *errVal, nil
	}

	// Output range check: must fit in 10-bit binary (-512 to 511).
	if n < -512 || n > 511 {
		return ErrorVal(ErrValNUM), nil
	}

	var result string
	if n < 0 {
		// Two's complement: add 1024 to get 10-bit binary representation.
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

// fnHex2Oct implements the HEX2OCT function.
// HEX2OCT(number, [places]) — converts a hexadecimal number string to octal.
// Input must contain only hex chars (0-9, A-F, a-f), max 10 digits.
// Output must be in range -536870912 to 536870911 (30-bit octal).
func fnHex2Oct(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	n, errVal := parseHexToInt64(args)
	if errVal != nil {
		return *errVal, nil
	}

	// Output range check: must fit in 30-bit octal (-536870912 to 536870911).
	if n < -536870912 || n > 536870911 {
		return ErrorVal(ErrValNUM), nil
	}

	var result string
	if n < 0 {
		// Two's complement: add 2^30 to get 10-digit octal representation.
		result = fmt.Sprintf("%o", n+1073741824)
	} else {
		result = fmt.Sprintf("%o", n)
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

// fnOct2Bin implements the OCT2BIN function.
// OCT2BIN(number, [places]) — converts an octal number string to binary.
// Input must contain only octal chars (0-7), max 10 digits.
// Output must be in range -512 to 511 (10-bit binary).
func fnOct2Bin(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	n, errVal := parseOctToInt64(args)
	if errVal != nil {
		return *errVal, nil
	}

	// Output range check: must fit in 10-bit binary (-512 to 511).
	if n < -512 || n > 511 {
		return ErrorVal(ErrValNUM), nil
	}

	var result string
	if n < 0 {
		// Two's complement: add 1024 to get 10-bit binary representation.
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

// fnOct2Hex implements the OCT2HEX function.
// OCT2HEX(number, [places]) — converts an octal number string to hexadecimal.
// Input must contain only octal chars (0-7), max 10 digits.
// Output is uppercase hex. Negative numbers use 40-bit two's complement.
func fnOct2Hex(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	n, errVal := parseOctToInt64(args)
	if errVal != nil {
		return *errVal, nil
	}

	var result string
	if n < 0 {
		// Two's complement: add 2^40 to get 10-digit hex representation.
		result = fmt.Sprintf("%X", n+1099511627776)
	} else {
		result = fmt.Sprintf("%X", n)
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

// ---------------------------------------------------------------------------
// CONVERT — unit conversion tables and logic
// ---------------------------------------------------------------------------

// convertCategory groups a set of units that can be converted between each other.
// For non-temperature categories the factor field gives the multiplier to the
// category's base unit. Temperature uses special-case formulas.
type convertCategory struct {
	name  string
	units map[string]float64 // unit name -> factor relative to base
}

// SI metric prefixes that can be prepended to applicable base units.
// The map value is the multiplier.
var siPrefixes = map[string]float64{
	"Y":  1e24,
	"Z":  1e21,
	"E":  1e18,
	"P":  1e15,
	"T":  1e12,
	"G":  1e9,
	"M":  1e6,
	"k":  1e3,
	"h":  1e2,
	"da": 1e1,
	"d":  1e-1,
	"c":  1e-2,
	"m":  1e-3,
	"u":  1e-6,
	"n":  1e-9,
	"p":  1e-12,
	"f":  1e-15,
	"a":  1e-18,
	"z":  1e-21,
	"y":  1e-24,
}

// siPrefixOrder is used to try longer prefixes first (e.g. "da" before "d").
var siPrefixOrder = []string{
	"Y", "Z", "E", "P", "T", "G", "M", "k", "h", "da",
	"d", "c", "m", "u", "n", "p", "f", "a", "z", "y",
}

// Units that accept SI prefixes.
var siEligible = map[string]bool{
	"g": true, "m": true, "l": true, "lt": true, "N": true, "Pa": true,
	"J": true, "W": true, "T": true, "e": true, "eV": true, "Wh": true,
	"cal": true, "sec": true, "bit": true, "byte": true, "dyn": true,
	"pond": true, "ang": true,
}

var convertMass = convertCategory{
	name: "mass",
	units: map[string]float64{
		"g":     1,
		"kg":    1e3,
		"lbm":   453.59237,
		"ozm":   28.349523125,
		"stone": 6350.29318,
		"ton":   907184.74,
		"sg":    14593.9029372064,
		"u":     1.66053906660e-24,
	},
}

var convertDistance = convertCategory{
	name: "distance",
	units: map[string]float64{
		"m":         1,
		"km":        1e3,
		"mi":        1609.344,
		"yd":        0.9144,
		"ft":        0.3048,
		"in":        0.0254,
		"Nmi":       1852,
		"ang":       1e-10,
		"Pica":      0.0254 / 72 * 6, // 1 Pica = 1/6 inch (PostScript pica)
		"ell":       1.143,
		"ly":        9.46073047258e15,
		"survey_mi": 1609.3472186944373,
	},
}

var convertTime = convertCategory{
	name: "time",
	units: map[string]float64{
		"sec": 1,
		"mn":  60,
		"hr":  3600,
		"day": 86400,
		"yr":  365.25 * 86400,
	},
}

var convertPressure = convertCategory{
	name: "pressure",
	units: map[string]float64{
		"Pa":   1,
		"atm":  101325,
		"mmHg": 133.322,
		"psi":  6894.76,
		"Torr": 133.3223684211,
	},
}

var convertForce = convertCategory{
	name: "force",
	units: map[string]float64{
		"N":    1,
		"dyn":  1e-5,
		"lbf":  4.4482216152605,
		"pond": 9.80665e-3,
	},
}

var convertEnergy = convertCategory{
	name: "energy",
	units: map[string]float64{
		"J":   1,
		"e":   1e-7,
		"c":   4.1868,
		"cal": 4.1868,
		"eV":  1.602176634e-19,
		"HPh": 2684519.537696173,
		"Wh":  3600,
		"flb": 1.3558179483314,
		"BTU": 1055.05585262,
	},
}

var convertPower = convertCategory{
	name: "power",
	units: map[string]float64{
		"W":  1,
		"HP": 745.69987158227022,
		"PS": 735.49875,
	},
}

var convertMagnetism = convertCategory{
	name: "magnetism",
	units: map[string]float64{
		"T":  1,
		"ga": 1e-4,
	},
}

var convertVolume = convertCategory{
	name: "volume",
	units: map[string]float64{
		"l":      1,
		"lt":     1,
		"tsp":    0.00492892159375,
		"tbs":    0.01478676478125,
		"oz":     0.0295735295625,
		"cup":    0.236588236500,
		"pt":     0.473176473,
		"qt":     0.946352946,
		"gal":    3.785411784,
		"m3":     1000,
		"mi3":    4.168181825440579584e12,
		"yd3":    764.554857984,
		"ft3":    28.316846592,
		"in3":    0.016387064,
		"ang3":   1e-27,
		"Pica3":  2.2316552567e-8, // (Pica)^3
		"barrel": 158.987294928,
		"bushel": 35.23907016688,
		"regton": 2831.6846592,
		"Nmi3":   6.352182208e12,
		"ly3":    8.46786664623715e47,
	},
}

var convertArea = convertCategory{
	name: "area",
	units: map[string]float64{
		"m2":      1,
		"mi2":     2589988.110336,
		"yd2":     0.83612736,
		"ft2":     0.09290304,
		"in2":     6.4516e-4,
		"ang2":    1e-20,
		"Pica2":   1.76369644444e-7,
		"Nmi2":    3429904,
		"Morgen":  2500,
		"ar":      100,
		"ha":      10000,
		"us_acre": 4046.8564224,
	},
}

var convertInformation = convertCategory{
	name: "information",
	units: map[string]float64{
		"bit":  1,
		"byte": 8,
	},
}

var convertSpeed = convertCategory{
	name: "speed",
	units: map[string]float64{
		"m/s":   1,
		"m/h":   1.0 / 3600.0,
		"mph":   0.44704,
		"kn":    0.514444444444444,
		"admkn": 0.514773333333333,
	},
}

// allCategories collects every non-temperature category.
var allCategories = []convertCategory{
	convertMass,
	convertDistance,
	convertTime,
	convertPressure,
	convertForce,
	convertEnergy,
	convertPower,
	convertMagnetism,
	convertVolume,
	convertArea,
	convertInformation,
	convertSpeed,
}

// Temperature unit names (handled specially).
var tempUnits = map[string]bool{
	"C": true, "F": true, "K": true, "Rank": true, "Reau": true,
}

// resolveUnit looks up a unit string, trying exact match first, then SI-prefix
// + base. Returns (category-name, factor, ok). For temperature units the factor
// is unused and category-name is "temperature".
func resolveUnit(unit string) (category string, factor float64, ok bool) {
	// Temperature — exact match only, no SI prefixes.
	if tempUnits[unit] {
		return "temperature", 0, true
	}

	// Exact match in non-temperature categories.
	for _, cat := range allCategories {
		if f, exists := cat.units[unit]; exists {
			return cat.name, f, true
		}
	}

	// Try SI prefix + base unit.
	for _, pfx := range siPrefixOrder {
		if strings.HasPrefix(unit, pfx) {
			base := unit[len(pfx):]
			if !siEligible[base] {
				continue
			}
			for _, cat := range allCategories {
				if baseFactor, exists := cat.units[base]; exists {
					return cat.name, baseFactor * siPrefixes[pfx], true
				}
			}
		}
	}

	return "", 0, false
}

// convertTemperature converts a temperature value from one unit to another.
func convertTemperature(val float64, from, to string) (float64, bool) {
	if from == to {
		return val, true
	}
	// Convert from -> Celsius first, then Celsius -> to.
	var celsius float64
	switch from {
	case "C":
		celsius = val
	case "F":
		celsius = (val - 32) * 5.0 / 9.0
	case "K":
		celsius = val - 273.15
	case "Rank":
		celsius = (val - 491.67) * 5.0 / 9.0
	case "Reau":
		celsius = val * 5.0 / 4.0
	default:
		return 0, false
	}

	var result float64
	switch to {
	case "C":
		result = celsius
	case "F":
		result = celsius*9.0/5.0 + 32
	case "K":
		result = celsius + 273.15
	case "Rank":
		result = celsius*9.0/5.0 + 491.67
	case "Reau":
		result = celsius * 4.0 / 5.0
	default:
		return 0, false
	}
	return result, true
}

// fnConvert implements the CONVERT function.
// CONVERT(number, from_unit, to_unit) — converts a number from one measurement
// unit to another.
func fnConvert(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}

	// First arg: number.
	num, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	// Second and third args: unit strings.
	if args[1].Type == ValueError {
		return args[1], nil
	}
	if args[2].Type == ValueError {
		return args[2], nil
	}

	fromUnit := ""
	switch args[1].Type {
	case ValueString:
		fromUnit = args[1].Str
	case ValueNumber:
		fromUnit = strconv.FormatFloat(args[1].Num, 'f', -1, 64)
	default:
		return ErrorVal(ErrValNA), nil
	}

	toUnit := ""
	switch args[2].Type {
	case ValueString:
		toUnit = args[2].Str
	case ValueNumber:
		toUnit = strconv.FormatFloat(args[2].Num, 'f', -1, 64)
	default:
		return ErrorVal(ErrValNA), nil
	}

	// Same unit: short-circuit.
	if fromUnit == toUnit {
		return NumberVal(num), nil
	}

	// Resolve units.
	fromCat, fromFactor, fromOk := resolveUnit(fromUnit)
	toCat, toFactor, toOk := resolveUnit(toUnit)

	if !fromOk || !toOk {
		return ErrorVal(ErrValNA), nil
	}
	if fromCat != toCat {
		return ErrorVal(ErrValNA), nil
	}

	// Temperature: special formulas.
	if fromCat == "temperature" {
		result, ok := convertTemperature(num, fromUnit, toUnit)
		if !ok {
			return ErrorVal(ErrValNA), nil
		}
		return NumberVal(result), nil
	}

	// Factor-based conversion.
	result := num * fromFactor / toFactor
	return NumberVal(result), nil
}

// parseComplex parses a complex number string (e.g. "3+4i",
// "-3-4j", "i", "3", "-i") and returns the real and imaginary coefficients.
// The third return value is true if the string is not a valid complex number.
func parseComplex(s string) (real, imag float64, fail bool) {
	if len(s) == 0 {
		return 0, 0, true
	}

	// Check for i/j suffix to determine if there's an imaginary part.
	suffix := s[len(s)-1]
	if suffix != 'i' && suffix != 'j' {
		// No imaginary suffix — must be a pure real number.
		r, err := strconv.ParseFloat(s, 64)
		if err != nil {
			return 0, 0, true
		}
		return r, 0, false
	}

	// Strip the i/j suffix.
	s = s[:len(s)-1]

	// Bare "i" or "j" → 0+1i.
	if len(s) == 0 {
		return 0, 1, false
	}

	// Just a sign: "-" → 0-1i, "+" → 0+1i.
	if s == "-" {
		return 0, -1, false
	}
	if s == "+" {
		return 0, 1, false
	}

	// Find the last '+' or '-' that separates real and imaginary parts.
	// We skip index 0 because the first character may be a sign for the
	// real (or pure-imaginary) part.
	splitIdx := -1
	for i := len(s) - 1; i >= 1; i-- {
		if s[i] == '+' || s[i] == '-' {
			// Make sure this is not part of an exponent (e.g. "1e+2").
			if i > 0 && (s[i-1] == 'e' || s[i-1] == 'E') {
				continue
			}
			splitIdx = i
			break
		}
	}

	if splitIdx == -1 {
		// No separator found — this is a pure imaginary number (e.g. "4i", "-3.5i").
		imCoeff, err := strconv.ParseFloat(s, 64)
		if err != nil {
			return 0, 0, true
		}
		return 0, imCoeff, false
	}

	// Split into real and imaginary parts.
	realStr := s[:splitIdx]
	imagStr := s[splitIdx:] // includes the sign

	r, err := strconv.ParseFloat(realStr, 64)
	if err != nil {
		return 0, 0, true
	}

	// imagStr may be just "+" or "-" meaning coefficient of 1 or -1.
	var im float64
	if imagStr == "+" {
		im = 1
	} else if imagStr == "-" {
		im = -1
	} else {
		im, err = strconv.ParseFloat(imagStr, 64)
		if err != nil {
			return 0, 0, true
		}
	}

	return r, im, false
}

// fnImabs implements the IMABS function.
// IMABS(inumber) — returns the absolute value (modulus) of a complex number.
// The modulus is sqrt(real² + imag²).
func fnImabs(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImabs([]Value{v})
			return r
		}), nil
	}

	// Numeric input: treat as real number with 0 imaginary part.
	if args[0].Type == ValueNumber {
		return NumberVal(math.Abs(args[0].Num)), nil
	}

	// Boolean: TRUE=1, FALSE=0, both are real numbers.
	if args[0].Type == ValueBool {
		if args[0].Bool {
			return NumberVal(1), nil
		}
		return NumberVal(0), nil
	}

	if args[0].Type != ValueString {
		return ErrorVal(ErrValVALUE), nil
	}

	real, imag, fail := parseComplex(args[0].Str)
	if fail {
		return ErrorVal(ErrValNUM), nil
	}

	return NumberVal(math.Sqrt(real*real + imag*imag)), nil
}

// fnImaginary implements the IMAGINARY function.
// IMAGINARY(inumber) — returns the imaginary coefficient of a complex number.
func fnImaginary(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImaginary([]Value{v})
			return r
		}), nil
	}

	// Numeric input: treat as real number with 0 imaginary part.
	if args[0].Type == ValueNumber {
		return NumberVal(0), nil
	}

	// Boolean: TRUE=1, FALSE=0, both are real numbers.
	if args[0].Type == ValueBool {
		return NumberVal(0), nil
	}

	if args[0].Type != ValueString {
		return ErrorVal(ErrValVALUE), nil
	}

	_, imag, fail := parseComplex(args[0].Str)
	if fail {
		return ErrorVal(ErrValNUM), nil
	}

	return NumberVal(imag), nil
}

// fnImreal implements the IMREAL function.
// IMREAL(inumber) — returns the real coefficient of a complex number.
func fnImreal(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImreal([]Value{v})
			return r
		}), nil
	}

	// Numeric input: the number itself is the real part.
	if args[0].Type == ValueNumber {
		return NumberVal(args[0].Num), nil
	}

	// Boolean: TRUE=1, FALSE=0.
	if args[0].Type == ValueBool {
		if args[0].Bool {
			return NumberVal(1), nil
		}
		return NumberVal(0), nil
	}

	if args[0].Type != ValueString {
		return ErrorVal(ErrValVALUE), nil
	}

	real, _, fail := parseComplex(args[0].Str)
	if fail {
		return ErrorVal(ErrValNUM), nil
	}

	return NumberVal(real), nil
}

// fnImargument implements the IMARGUMENT function.
// IMARGUMENT(inumber) — returns the argument (theta/angle in radians) of a complex number.
// The argument of zero is undefined and returns #DIV/0!.
func fnImargument(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImargument([]Value{v})
			return r
		}), nil
	}

	// Numeric input: treat as real number with 0 imaginary part.
	if args[0].Type == ValueNumber {
		if args[0].Num == 0 {
			return ErrorVal(ErrValDIV0), nil
		}
		return NumberVal(math.Atan2(0, args[0].Num)), nil
	}

	// Boolean: TRUE=1, FALSE=0, both are real numbers.
	if args[0].Type == ValueBool {
		if args[0].Bool {
			return NumberVal(0), nil
		}
		return ErrorVal(ErrValDIV0), nil
	}

	if args[0].Type != ValueString {
		return ErrorVal(ErrValVALUE), nil
	}

	real, imag, fail := parseComplex(args[0].Str)
	if fail {
		return ErrorVal(ErrValNUM), nil
	}

	if real == 0 && imag == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	return NumberVal(math.Atan2(imag, real)), nil
}

// fnImconjugate implements the IMCONJUGATE function.
// IMCONJUGATE(inumber) — returns the complex conjugate of a complex number.
// The conjugate of a+bi is a-bi (the imaginary part is negated).
func fnImconjugate(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImconjugate([]Value{v})
			return r
		}), nil
	}

	// Numeric input: treat as real number, conjugate is itself.
	if args[0].Type == ValueNumber {
		return StringVal(formatComplex(args[0].Num, 0, "i")), nil
	}

	// Boolean: TRUE=1, FALSE=0, both are real numbers.
	if args[0].Type == ValueBool {
		if args[0].Bool {
			return StringVal("1"), nil
		}
		return StringVal("0"), nil
	}

	if args[0].Type != ValueString {
		return ErrorVal(ErrValVALUE), nil
	}

	real, imag, suffix, fail := parseComplexWithSuffix(args[0].Str)
	if fail {
		return ErrorVal(ErrValNUM), nil
	}

	return StringVal(formatComplex(real, -imag, suffix)), nil
}

// parseComplexWithSuffix is like parseComplex but also returns the suffix
// character ("i", "j", or "" for pure real numbers).
func parseComplexWithSuffix(s string) (real, imag float64, suffix string, fail bool) {
	if len(s) == 0 {
		return 0, 0, "", true
	}

	// Check for i/j suffix.
	last := s[len(s)-1]
	if last != 'i' && last != 'j' {
		// No imaginary suffix — must be a pure real number.
		r, err := strconv.ParseFloat(s, 64)
		if err != nil {
			return 0, 0, "", true
		}
		return r, 0, "", false
	}

	suffix = string(last)
	r, im, bad := parseComplex(s)
	if bad {
		return 0, 0, "", true
	}
	return r, im, suffix, false
}

// formatComplex formats a complex number as a formatted string
// using the same formatting rules as the COMPLEX function.
func formatComplex(real, imag float64, suffix string) string {
	// Both zero: just "0".
	if real == 0 && imag == 0 {
		return "0"
	}

	var result string

	// Build real part.
	if real != 0 {
		result = formatComplexNum(real)
	}

	// Build imaginary part.
	if imag != 0 {
		if real != 0 {
			if imag > 0 {
				result += "+"
			}
		}

		if imag == 1 {
			result += suffix
		} else if imag == -1 {
			result += "-" + suffix
		} else {
			result += formatComplexNum(imag) + suffix
		}
	}

	return result
}

// fnImdiv implements the IMDIV function.
// IMDIV(inumber1, inumber2) — returns the quotient of two complex numbers.
// Both arguments must use the same suffix (i or j). Returns #NUM! for invalid inputs.
// Division by zero (both real and imag of divisor are 0) returns #NUM!.
func fnImdiv(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	type parsed struct {
		real, imag float64
		suffix     string
	}

	var parts [2]parsed
	for i := 0; i < 2; i++ {
		arg := args[i]

		// Propagate errors.
		if arg.Type == ValueError {
			return arg, nil
		}

		switch arg.Type {
		case ValueNumber:
			parts[i].real = arg.Num
			parts[i].imag = 0
			parts[i].suffix = ""
		case ValueString:
			var fail bool
			parts[i].real, parts[i].imag, parts[i].suffix, fail = parseComplexWithSuffix(arg.Str)
			if fail {
				return ErrorVal(ErrValNUM), nil
			}
		case ValueBool:
			return ErrorVal(ErrValVALUE), nil
		default:
			return ErrorVal(ErrValVALUE), nil
		}
	}

	// Check suffix consistency.
	suffix := ""
	suffixSet := false
	for _, p := range parts {
		if p.suffix != "" {
			if !suffixSet {
				suffix = p.suffix
				suffixSet = true
			} else if suffix != p.suffix {
				return ErrorVal(ErrValNUM), nil
			}
		}
	}
	if !suffixSet {
		suffix = "i"
	}

	// Division by zero check.
	c, d := parts[1].real, parts[1].imag
	if c == 0 && d == 0 {
		return ErrorVal(ErrValNUM), nil
	}

	// (a+bi)/(c+di) = ((ac+bd) + (bc-ad)i) / (c²+d²)
	a, b := parts[0].real, parts[0].imag
	denom := c*c + d*d
	realResult := (a*c + b*d) / denom
	imagResult := (b*c - a*d) / denom

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImproduct implements the IMPRODUCT function.
// IMPRODUCT(inumber1, [inumber2], ...) — returns the product of 1 to 255 complex numbers.
// All arguments must use the same suffix (i or j). Returns #NUM! for invalid inputs.
func fnImproduct(args []Value) (Value, error) {
	if len(args) < 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Start with multiplicative identity.
	totalReal := 1.0
	totalImag := 0.0
	suffix := "" // track the resolved suffix across all args
	suffixSet := false

	for _, arg := range args {
		// Propagate errors.
		if arg.Type == ValueError {
			return arg, nil
		}

		var r, im float64
		var argSuffix string

		switch arg.Type {
		case ValueNumber:
			r = arg.Num
			im = 0
			argSuffix = ""
		case ValueString:
			var fail bool
			r, im, argSuffix, fail = parseComplexWithSuffix(arg.Str)
			if fail {
				return ErrorVal(ErrValNUM), nil
			}
		case ValueBool:
			return ErrorVal(ErrValVALUE), nil
		default:
			return ErrorVal(ErrValVALUE), nil
		}

		// Check suffix consistency.
		if argSuffix != "" {
			if !suffixSet {
				suffix = argSuffix
				suffixSet = true
			} else if suffix != argSuffix {
				return ErrorVal(ErrValNUM), nil
			}
		}

		// Complex multiplication: (a+bi)(c+di) = (ac-bd) + (ad+bc)i
		newReal := totalReal*r - totalImag*im
		newImag := totalReal*im + totalImag*r
		totalReal = newReal
		totalImag = newImag
	}

	// Default suffix if none was set.
	if !suffixSet {
		suffix = "i"
	}

	return StringVal(formatComplex(totalReal, totalImag, suffix)), nil
}

// fnImsum implements the IMSUM function.
// IMSUM(inumber1, [inumber2], ...) — returns the sum of two or more complex numbers.
// All arguments must use the same suffix (i or j). Returns #NUM! for invalid inputs.
func fnImsum(args []Value) (Value, error) {
	if len(args) < 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	var totalReal, totalImag float64
	suffix := "" // track the resolved suffix across all args
	suffixSet := false

	for _, arg := range args {
		// Propagate errors.
		if arg.Type == ValueError {
			return arg, nil
		}

		var r, im float64
		var argSuffix string

		switch arg.Type {
		case ValueNumber:
			r = arg.Num
			im = 0
			argSuffix = ""
		case ValueString:
			var fail bool
			r, im, argSuffix, fail = parseComplexWithSuffix(arg.Str)
			if fail {
				return ErrorVal(ErrValNUM), nil
			}
		case ValueBool:
			return ErrorVal(ErrValVALUE), nil
		default:
			return ErrorVal(ErrValVALUE), nil
		}

		// Check suffix consistency.
		if argSuffix != "" {
			if !suffixSet {
				suffix = argSuffix
				suffixSet = true
			} else if suffix != argSuffix {
				return ErrorVal(ErrValNUM), nil
			}
		}

		totalReal += r
		totalImag += im
	}

	// Default suffix if none was set.
	if !suffixSet {
		suffix = "i"
	}

	return StringVal(formatComplex(totalReal, totalImag, suffix)), nil
}

// fnImsub implements the IMSUB function.
// IMSUB(inumber1, inumber2) — returns the difference of two complex numbers.
// Both arguments must use the same suffix (i or j). Returns #NUM! for invalid inputs.
func fnImsub(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	type parsed struct {
		real, imag float64
		suffix     string
	}

	var parts [2]parsed
	for i := 0; i < 2; i++ {
		arg := args[i]

		// Propagate errors.
		if arg.Type == ValueError {
			return arg, nil
		}

		switch arg.Type {
		case ValueNumber:
			parts[i].real = arg.Num
			parts[i].imag = 0
			parts[i].suffix = ""
		case ValueString:
			var fail bool
			parts[i].real, parts[i].imag, parts[i].suffix, fail = parseComplexWithSuffix(arg.Str)
			if fail {
				return ErrorVal(ErrValNUM), nil
			}
		case ValueBool:
			return ErrorVal(ErrValVALUE), nil
		default:
			return ErrorVal(ErrValVALUE), nil
		}
	}

	// Check suffix consistency.
	suffix := ""
	suffixSet := false
	for _, p := range parts {
		if p.suffix != "" {
			if !suffixSet {
				suffix = p.suffix
				suffixSet = true
			} else if suffix != p.suffix {
				return ErrorVal(ErrValNUM), nil
			}
		}
	}
	if !suffixSet {
		suffix = "i"
	}

	realResult := parts[0].real - parts[1].real
	imagResult := parts[0].imag - parts[1].imag

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// cleanFloat rounds a floating-point value to the nearest integer when it is
// extremely close (within 1e-12), eliminating polar-form round-trip noise.
// Values that are not near an integer are returned unchanged.
func cleanFloat(v float64) float64 {
	rounded := math.Round(v)
	if math.Abs(v-rounded) < 1e-12 {
		return rounded
	}
	return v
}

// fnImsqrt implements the IMSQRT function.
// IMSQRT(inumber) — returns the square root of a complex number.
// Uses polar form: r=sqrt(x²+y²), θ=atan2(y,x), result=sqrt(r)*(cos(θ/2)+sin(θ/2)*i).
func fnImsqrt(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImsqrt([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	// Special case: zero.
	if x == 0 && y == 0 {
		return StringVal("0"), nil
	}

	// Convert to polar form and compute square root.
	r := math.Hypot(x, y)
	theta := math.Atan2(y, x)
	newR := math.Sqrt(r)
	newTheta := theta / 2

	realResult := cleanFloat(newR * math.Cos(newTheta))
	imagResult := cleanFloat(newR * math.Sin(newTheta))

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImpower implements the IMPOWER function.
// IMPOWER(inumber, number) — returns a complex number raised to a power.
// Uses polar form: r^n * (cos(nθ) + sin(nθ)*i).
func fnImpower(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}
	if args[1].Type == ValueError {
		return args[1], nil
	}

	// Second arg must be numeric (not a complex string).
	n, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	r := math.Hypot(x, y)

	// Special cases involving zero base.
	if r == 0 {
		if n < 0 {
			// Division by zero.
			return ErrorVal(ErrValNUM), nil
		}
		if n == 0 {
			// 0^0 = 1 by convention.
			return StringVal("1"), nil
		}
		// 0^positive = 0.
		return StringVal("0"), nil
	}

	// Special case: n=0, any non-zero base → 1.
	if n == 0 {
		return StringVal("1"), nil
	}

	theta := math.Atan2(y, x)
	newR := math.Pow(r, n)
	newTheta := n * theta

	realResult := cleanFloat(newR * math.Cos(newTheta))
	imagResult := cleanFloat(newR * math.Sin(newTheta))

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImexp implements the IMEXP function.
// IMEXP(inumber) — returns the exponential of a complex number.
// e^(x+yi) = e^x * (cos(y) + sin(y)*i).
func fnImexp(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImexp([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	// e^(x+yi) = e^x * (cos(y) + sin(y)*i)
	ex := math.Exp(x)
	realResult := cleanFloat(ex * math.Cos(y))
	imagResult := cleanFloat(ex * math.Sin(y))

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImln implements the IMLN function.
// IMLN(inumber) — returns the natural logarithm of a complex number.
// ln(x+yi) = ln(|z|) + atan2(y,x)*i.
func fnImln(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImln([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	// Log of zero is undefined → #NUM!
	r := math.Hypot(x, y)
	if r == 0 {
		return ErrorVal(ErrValNUM), nil
	}

	theta := math.Atan2(y, x)
	realResult := cleanFloat(math.Log(r))
	imagResult := cleanFloat(theta)

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImlog2 implements the IMLOG2 function.
// IMLOG2(inumber) — returns the base-2 logarithm of a complex number.
// log2(z) = ln(z) / ln(2).
func fnImlog2(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImlog2([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	// Log of zero is undefined → #NUM!
	r := math.Hypot(x, y)
	if r == 0 {
		return ErrorVal(ErrValNUM), nil
	}

	theta := math.Atan2(y, x)
	ln2 := math.Ln2
	realResult := cleanFloat(math.Log(r) / ln2)
	imagResult := cleanFloat(theta / ln2)

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImlog10 implements the IMLOG10 function.
// IMLOG10(inumber) — returns the base-10 logarithm of a complex number.
// log10(z) = ln(z) / ln(10).
func fnImlog10(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImlog10([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	// Log of zero is undefined → #NUM!
	r := math.Hypot(x, y)
	if r == 0 {
		return ErrorVal(ErrValNUM), nil
	}

	theta := math.Atan2(y, x)
	ln10 := math.Log(10)
	realResult := cleanFloat(math.Log(r) / ln10)
	imagResult := cleanFloat(theta / ln10)

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImsin implements the IMSIN function.
// IMSIN(inumber) — returns the sine of a complex number.
// sin(x+yi) = sin(x)*cosh(y) + cos(x)*sinh(y)*i.
func fnImsin(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImsin([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	realResult := cleanFloat(math.Sin(x) * math.Cosh(y))
	imagResult := cleanFloat(math.Cos(x) * math.Sinh(y))

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImcos implements the IMCOS function.
// IMCOS(inumber) — returns the cosine of a complex number.
// cos(x+yi) = cos(x)*cosh(y) - sin(x)*sinh(y)*i.
func fnImcos(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImcos([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	realResult := cleanFloat(math.Cos(x) * math.Cosh(y))
	imagResult := cleanFloat(-math.Sin(x) * math.Sinh(y))

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImtan implements the IMTAN function.
// IMTAN(inumber) — returns the tangent of a complex number.
// tan(z) = sin(z)/cos(z).
// real part = sin(2x)/(cos(2x)+cosh(2y))
// imag part = sinh(2y)/(cos(2x)+cosh(2y))
func fnImtan(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImtan([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	denom := math.Cos(2*x) + math.Cosh(2*y)

	// Check for overflow or zero denominator (practically impossible for finite inputs
	// since cosh(2y) >= 1, but guard against overflow with very large values).
	if denom == 0 || math.IsInf(denom, 0) || math.IsNaN(denom) {
		return ErrorVal(ErrValNUM), nil
	}

	realResult := cleanFloat(math.Sin(2*x) / denom)
	imagResult := cleanFloat(math.Sinh(2*y) / denom)

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImsinh implements the IMSINH function.
// IMSINH(inumber) — returns the hyperbolic sine of a complex number.
// sinh(x+yi) = sinh(x)*cos(y) + cosh(x)*sin(y)*i.
func fnImsinh(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImsinh([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	realResult := cleanFloat(math.Sinh(x) * math.Cos(y))
	imagResult := cleanFloat(math.Cosh(x) * math.Sin(y))

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImcosh implements the IMCOSH function.
// IMCOSH(inumber) — returns the hyperbolic cosine of a complex number.
// cosh(x+yi) = cosh(x)*cos(y) + sinh(x)*sin(y)*i.
func fnImcosh(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImcosh([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	realResult := cleanFloat(math.Cosh(x) * math.Cos(y))
	imagResult := cleanFloat(math.Sinh(x) * math.Sin(y))

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImsech implements the IMSECH function.
// IMSECH(inumber) — returns the hyperbolic secant of a complex number.
// sech(z) = 1/cosh(z).
func fnImsec(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImsec([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	// cos(x+yi) = cos(x)*cosh(y) - sin(x)*sinh(y)*i
	cr := math.Cos(x) * math.Cosh(y)
	ci := -math.Sin(x) * math.Sinh(y)

	// 1/(cr+ci*i): multiply by conjugate → (cr-ci*i)/(cr²+ci²)
	denom := cr*cr + ci*ci
	if denom < 1e-24 {
		return ErrorVal(ErrValNUM), nil
	}

	realResult := cleanFloat(cr / denom)
	imagResult := cleanFloat(-ci / denom)

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

func fnImsech(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImsech([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	// cosh(x+yi) = cosh(x)*cos(y) + sinh(x)*sin(y)*i
	cr := math.Cosh(x) * math.Cos(y)
	ci := math.Sinh(x) * math.Sin(y)

	// 1/(cr+ci*i): multiply by conjugate → (cr-ci*i)/(cr²+ci²)
	denom := cr*cr + ci*ci
	if denom < 1e-24 {
		return ErrorVal(ErrValNUM), nil
	}

	realResult := cleanFloat(cr / denom)
	imagResult := cleanFloat(-ci / denom)

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImcsc implements the IMCSC function.
// IMCSC(inumber) — returns the cosecant of a complex number.
// csc(z) = 1/sin(z).
func fnImcsc(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImcsc([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	// sin(x+yi) = sin(x)*cosh(y) + cos(x)*sinh(y)*i
	sr := math.Sin(x) * math.Cosh(y)
	si := math.Cos(x) * math.Sinh(y)

	// 1/(sr+si*i): multiply by conjugate → (sr-si*i)/(sr²+si²)
	denom := sr*sr + si*si
	if denom < 1e-24 {
		return ErrorVal(ErrValNUM), nil
	}

	realResult := cleanFloat(sr / denom)
	imagResult := cleanFloat(-si / denom)

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImcot implements the IMCOT function.
// IMCOT(inumber) — returns the cotangent of a complex number.
// cot(z) = cos(z)/sin(z).
func fnImcot(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImcot([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	// cos(x+yi) = cos(x)*cosh(y) - sin(x)*sinh(y)*i
	cr := math.Cos(x) * math.Cosh(y)
	ci := -math.Sin(x) * math.Sinh(y)

	// sin(x+yi) = sin(x)*cosh(y) + cos(x)*sinh(y)*i
	sr := math.Sin(x) * math.Cosh(y)
	si := math.Cos(x) * math.Sinh(y)

	// (cr+ci*i)/(sr+si*i): multiply by conjugate of denominator
	// = ((cr*sr+ci*si) + (ci*sr-cr*si)*i) / (sr²+si²)
	denom := sr*sr + si*si
	if denom < 1e-24 {
		return ErrorVal(ErrValNUM), nil
	}

	realResult := cleanFloat((cr*sr + ci*si) / denom)
	imagResult := cleanFloat((ci*sr - cr*si) / denom)

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnImcsch implements the IMCSCH function.
// IMCSCH(inumber) — returns the hyperbolic cosecant of a complex number.
// csch(z) = 1/sinh(z).
func fnImcsch(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors.
	if args[0].Type == ValueError {
		return args[0], nil
	}

	// Handle arrays.
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			r, _ := fnImcsch([]Value{v})
			return r
		}), nil
	}

	var x, y float64
	var suffix string

	switch args[0].Type {
	case ValueNumber:
		x = args[0].Num
		y = 0
		suffix = ""
	case ValueString:
		var fail bool
		x, y, suffix, fail = parseComplexWithSuffix(args[0].Str)
		if fail {
			return ErrorVal(ErrValNUM), nil
		}
	case ValueBool:
		return ErrorVal(ErrValVALUE), nil
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Default suffix if none was set.
	if suffix == "" {
		suffix = "i"
	}

	// sinh(x+yi) = sinh(x)*cos(y) + cosh(x)*sin(y)*i
	sr := math.Sinh(x) * math.Cos(y)
	si := math.Cosh(x) * math.Sin(y)

	// 1/(sr+si*i): multiply by conjugate → (sr-si*i)/(sr²+si²)
	denom := sr*sr + si*si
	if denom < 1e-24 {
		return ErrorVal(ErrValNUM), nil
	}

	realResult := cleanFloat(sr / denom)
	imagResult := cleanFloat(-si / denom)

	return StringVal(formatComplex(realResult, imagResult, suffix)), nil
}

// fnGESTEP implements the GESTEP function.
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
