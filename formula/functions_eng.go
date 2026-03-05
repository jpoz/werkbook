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
