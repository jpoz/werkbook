package formula

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

func init() {
	Register("BIN2DEC", NoCtx(fnBin2Dec))
	Register("BIN2HEX", NoCtx(fnBin2Hex))
	Register("BIN2OCT", NoCtx(fnBin2Oct))
	Register("CONVERT", NoCtx(fnConvert))
	Register("DELTA", NoCtx(fnDELTA))
	Register("DEC2BIN", NoCtx(fnDec2Bin))
	Register("DEC2HEX", NoCtx(fnDec2Hex))
	Register("DEC2OCT", NoCtx(fnDec2Oct))
	Register("GESTEP", NoCtx(fnGESTEP))
	Register("HEX2BIN", NoCtx(fnHex2Bin))
	Register("HEX2DEC", NoCtx(fnHex2Dec))
	Register("HEX2OCT", NoCtx(fnHex2Oct))
	Register("OCT2BIN", NoCtx(fnOct2Bin))
	Register("OCT2DEC", NoCtx(fnOct2Dec))
	Register("OCT2HEX", NoCtx(fnOct2Hex))
}

// fnBin2Dec implements the Excel BIN2DEC function.
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

// fnBin2Hex implements the Excel BIN2HEX function.
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

// fnBin2Oct implements the Excel BIN2OCT function.
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

	// Excel engineering functions reject bare booleans with #VALUE!.
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

// fnDec2Hex implements the Excel DEC2HEX function.
// DEC2HEX(number, [places]) — converts a decimal number to hexadecimal.
// number must be between -549755813888 and 549755813887 (inclusive, -2^39 to 2^39-1).
// Non-integer values are truncated. Negative numbers use two's complement (10 hex digits = 40 bits).
// places specifies minimum digits (1–10).
func fnDec2Hex(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Excel engineering functions reject bare booleans with #VALUE!.
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

// fnDec2Oct implements the Excel DEC2OCT function.
// DEC2OCT(number, [places]) — converts a decimal number to octal.
// number must be between -536870912 and 536870911 (inclusive, -2^29 to 2^29-1).
// Non-integer values are truncated. Negative numbers use two's complement (10 octal digits = 30 bits).
// places specifies minimum digits (1–10).
func fnDec2Oct(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Excel engineering functions reject bare booleans with #VALUE!.
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

// fnHex2Dec implements the Excel HEX2DEC function.
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

// fnOct2Dec implements the Excel OCT2DEC function.
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

// fnHex2Bin implements the Excel HEX2BIN function.
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

// fnHex2Oct implements the Excel HEX2OCT function.
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

// fnOct2Bin implements the Excel OCT2BIN function.
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

// fnOct2Hex implements the Excel OCT2HEX function.
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
		"m":    1,
		"km":   1e3,
		"mi":   1609.344,
		"yd":   0.9144,
		"ft":   0.3048,
		"in":   0.0254,
		"Nmi":  1852,
		"ang":  1e-10,
		"Pica": 0.0254 / 72 * 6, // 1 Pica = 1/6 inch (PostScript pica)
		"ell":  1.143,
		"ly":   9.46073047258e15,
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

// fnConvert implements the Excel CONVERT function.
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
