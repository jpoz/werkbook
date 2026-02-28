package formula

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"unicode"
	"unicode/utf8"
)

func fnCHAR(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	code := int(n)
	if code < 1 || code > 255 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(string(rune(code))), nil
}

func fnCHOOSE(args []Value) (Value, error) {
	if len(args) < 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	idx, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	i := int(idx)
	if i < 1 || i > len(args)-1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return args[i], nil
}

func fnCLEAN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	var b strings.Builder
	for _, r := range s {
		if r >= 32 {
			b.WriteRune(r)
		}
	}
	return StringVal(b.String()), nil
}

func fnCODE(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	if len(s) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	r, _ := utf8.DecodeRuneInString(s)
	return NumberVal(float64(r)), nil
}

func fnCONCATENATE(args []Value) (Value, error) {
	var b strings.Builder
	for _, arg := range args {
		if arg.Type == ValueError {
			return arg, nil
		}
		b.WriteString(valueToString(arg))
	}
	return StringVal(b.String()), nil
}

func fnEXACT(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(valueToString(args[0]) == valueToString(args[1])), nil
}

func fnFIND(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	findText := valueToString(args[0])
	withinText := valueToString(args[1])
	startNum := 1
	if len(args) == 3 {
		sn, e := coerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		startNum = int(sn)
	}
	if startNum < 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	runes := []rune(withinText)
	findRunes := []rune(findText)
	start := startNum - 1
	if start > len(runes) {
		return ErrorVal(ErrValVALUE), nil
	}

	for i := start; i <= len(runes)-len(findRunes); i++ {
		if string(runes[i:i+len(findRunes)]) == findText {
			return NumberVal(float64(i + 1)), nil
		}
	}
	return ErrorVal(ErrValVALUE), nil
}

func fnFIXED(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	decimals := 2
	if len(args) >= 2 {
		d, e := coerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		decimals = int(d)
	}

	noCommas := false
	if len(args) >= 3 {
		noCommas = isTruthy(args[2])
	}

	// Handle negative decimals: round to the left of the decimal point
	if decimals < 0 {
		factor := math.Pow(10, float64(-decimals))
		n = math.Round(n/factor) * factor
		decimals = 0
	} else {
		factor := math.Pow(10, float64(decimals))
		n = math.Round(n*factor) / factor
	}

	if noCommas {
		return StringVal(fmt.Sprintf("%.*f", decimals, n)), nil
	}
	return StringVal(formatWithCommas(n, decimals)), nil
}

func fnLEFT(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	n := 1
	if len(args) == 2 {
		num, e := coerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		n = int(num)
	}
	runes := []rune(s)
	if n > len(runes) {
		n = len(runes)
	}
	if n < 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(string(runes[:n])), nil
}

func fnLEN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	return NumberVal(float64(utf8.RuneCountInString(s))), nil
}

func fnLOWER(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	return StringVal(strings.ToLower(valueToString(args[0]))), nil
}

func fnMID(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	startNum, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	numChars, e := coerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	start := int(startNum) - 1
	length := int(numChars)
	if start < 0 || length < 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	runes := []rune(s)
	if start >= len(runes) {
		return StringVal(""), nil
	}
	end := start + length
	if end > len(runes) {
		end = len(runes)
	}
	return StringVal(string(runes[start:end])), nil
}

func fnNUMBERVALUE(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	text := valueToString(args[0])

	decSep := "."
	grpSep := ","
	if len(args) >= 2 {
		ds := valueToString(args[1])
		if len(ds) > 0 {
			decSep = string(ds[0])
		}
	}
	if len(args) >= 3 {
		gs := valueToString(args[2])
		if len(gs) > 0 {
			grpSep = string(gs[0])
		}
	}

	// Strip all whitespace (spaces, tabs, etc.)
	text = strings.Map(func(r rune) rune {
		if unicode.IsSpace(r) {
			return -1
		}
		return r
	}, text)

	if text == "" {
		return NumberVal(0), nil
	}

	// Count and remove percent signs
	percentCount := strings.Count(text, "%")
	text = strings.ReplaceAll(text, "%", "")

	// Check: group separator must not appear after decimal separator
	decIdx := strings.Index(text, decSep)
	if decIdx >= 0 {
		after := text[decIdx+len(decSep):]
		if strings.Contains(after, grpSep) {
			return ErrorVal(ErrValVALUE), nil
		}
	}

	// Remove group separators
	text = strings.ReplaceAll(text, grpSep, "")

	// Check for multiple decimal separators
	if strings.Count(text, decSep) > 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Replace decimal separator with "." for Go parsing
	text = strings.Replace(text, decSep, ".", 1)

	num, err := strconv.ParseFloat(text, 64)
	if err != nil {
		return ErrorVal(ErrValVALUE), nil
	}

	// Apply percent scaling
	for i := 0; i < percentCount; i++ {
		num /= 100
	}

	return NumberVal(num), nil
}

func fnPROPER(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	var b strings.Builder
	capitalizeNext := true
	for _, r := range s {
		if unicode.IsLetter(r) {
			if capitalizeNext {
				b.WriteRune(unicode.ToUpper(r))
				capitalizeNext = false
			} else {
				b.WriteRune(unicode.ToLower(r))
			}
		} else {
			b.WriteRune(r)
			capitalizeNext = true
		}
	}
	return StringVal(b.String()), nil
}

func fnREPLACE(args []Value) (Value, error) {
	if len(args) != 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	oldText := valueToString(args[0])
	startNum, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	numChars, e := coerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	newText := valueToString(args[3])

	runes := []rune(oldText)
	start := int(startNum) - 1
	length := int(numChars)
	if start < 0 || length < 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	if start > len(runes) {
		start = len(runes)
	}
	end := start + length
	if end > len(runes) {
		end = len(runes)
	}
	return StringVal(string(runes[:start]) + newText + string(runes[end:])), nil
}

func fnREPT(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	n, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	count := int(n)
	if count < 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(strings.Repeat(s, count)), nil
}

func fnRIGHT(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	n := 1
	if len(args) == 2 {
		num, e := coerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		n = int(num)
	}
	runes := []rune(s)
	if n > len(runes) {
		n = len(runes)
	}
	if n < 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(string(runes[len(runes)-n:])), nil
}

func fnSEARCH(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	findText := strings.ToLower(valueToString(args[0]))
	withinText := strings.ToLower(valueToString(args[1]))
	startNum := 1
	if len(args) == 3 {
		sn, e := coerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		startNum = int(sn)
	}
	if startNum < 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	runes := []rune(withinText)
	start := startNum - 1
	if start > len(runes) {
		return ErrorVal(ErrValVALUE), nil
	}
	remaining := string(runes[start:])

	idx := strings.Index(remaining, findText)
	if idx < 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	runeIdx := utf8.RuneCountInString(remaining[:idx])
	return NumberVal(float64(start + runeIdx + 1)), nil
}

func fnSUBSTITUTE(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	text := valueToString(args[0])
	oldText := valueToString(args[1])
	newText := valueToString(args[2])

	if len(args) == 4 {
		instanceNum, e := coerceNum(args[3])
		if e != nil {
			return *e, nil
		}
		n := int(instanceNum)
		if n < 1 {
			return ErrorVal(ErrValVALUE), nil
		}
		count := 0
		result := text
		idx := 0
		for {
			pos := strings.Index(result[idx:], oldText)
			if pos < 0 {
				break
			}
			count++
			if count == n {
				result = result[:idx+pos] + newText + result[idx+pos+len(oldText):]
				break
			}
			idx += pos + len(oldText)
		}
		return StringVal(result), nil
	}

	return StringVal(strings.ReplaceAll(text, oldText, newText)), nil
}

func fnT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueString {
		return args[0], nil
	}
	return StringVal(""), nil
}

func fnTEXT(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	format := valueToString(args[1])
	return StringVal(formatNumber(n, format)), nil
}

func formatNumber(n float64, format string) string {
	upper := strings.ToUpper(format)

	switch {
	case strings.Contains(format, "%"):
		decimals := strings.Count(format, "0") - 1
		if decimals < 0 {
			decimals = 0
		}
		return fmt.Sprintf("%.*f%%", decimals, n*100)

	case upper == "YYYY-MM-DD" || upper == "YYYY/MM/DD":
		t := excelSerialToTime(n)
		return t.Format("2006-01-02")

	case upper == "MM/DD/YYYY":
		t := excelSerialToTime(n)
		return t.Format("01/02/2006")

	case upper == "HH:MM:SS" || upper == "H:MM:SS":
		t := excelSerialToTime(n)
		return t.Format("15:04:05")

	case upper == "HH:MM" || upper == "H:MM":
		t := excelSerialToTime(n)
		return t.Format("15:04")

	case strings.Contains(format, "#,##0") || strings.Contains(format, "#,###"):
		decimals := 0
		if dotIdx := strings.Index(format, "."); dotIdx >= 0 {
			decimals = len(format) - dotIdx - 1
		}
		return formatWithCommas(n, decimals)

	case strings.Contains(format, "0."):
		decimals := strings.Count(format[strings.Index(format, "."):], "0")
		return fmt.Sprintf("%.*f", decimals, n)

	case format == "0":
		return fmt.Sprintf("%.0f", n)

	default:
		return fmt.Sprintf("%g", n)
	}
}

func formatWithCommas(n float64, decimals int) string {
	s := fmt.Sprintf("%.*f", decimals, n)
	parts := strings.SplitN(s, ".", 2)
	intPart := parts[0]
	negative := false
	if strings.HasPrefix(intPart, "-") {
		negative = true
		intPart = intPart[1:]
	}

	var result strings.Builder
	for i, ch := range intPart {
		if i > 0 && (len(intPart)-i)%3 == 0 {
			result.WriteByte(',')
		}
		result.WriteRune(ch)
	}

	s = result.String()
	if negative {
		s = "-" + s
	}
	if len(parts) == 2 {
		s += "." + parts[1]
	}
	return s
}

func fnTEXTJOIN(args []Value) (Value, error) {
	if len(args) < 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	delimiter := valueToString(args[0])
	ignoreEmpty := isTruthy(args[1])

	var parts []string
	for _, arg := range args[2:] {
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if ignoreEmpty && (cell.Type == ValueEmpty || (cell.Type == ValueString && cell.Str == "")) {
						continue
					}
					parts = append(parts, valueToString(cell))
				}
			}
		} else {
			if ignoreEmpty && (arg.Type == ValueEmpty || (arg.Type == ValueString && arg.Str == "")) {
				continue
			}
			parts = append(parts, valueToString(arg))
		}
	}

	result := strings.Join(parts, delimiter)
	if len(result) > 32767 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(result), nil
}

func fnTRIM(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := valueToString(args[0])
	fields := strings.Fields(s)
	return StringVal(strings.Join(fields, " ")), nil
}

func fnUPPER(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(strings.ToUpper(valueToString(args[0]))), nil
}

func fnVALUEFn(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueNumber {
		return args[0], nil
	}
	s := valueToString(args[0])
	s = strings.TrimSpace(s)
	s = strings.ReplaceAll(s, ",", "")
	s = strings.ReplaceAll(s, "$", "")

	if strings.HasSuffix(s, "%") {
		s = strings.TrimSuffix(s, "%")
		num, err := strconv.ParseFloat(s, 64)
		if err != nil {
			return ErrorVal(ErrValVALUE), nil
		}
		return NumberVal(num / 100), nil
	}

	num, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return ErrorVal(ErrValVALUE), nil
	}
	return NumberVal(num), nil
}
