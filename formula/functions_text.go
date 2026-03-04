package formula

import (
	"fmt"
	"math"
	"regexp"
	"strconv"
	"strings"
	"unicode"
	"unicode/utf8"
)

func init() {
	Register("CHAR", NoCtx(fnCHAR))
	Register("CHOOSE", NoCtx(fnCHOOSE))
	Register("CLEAN", NoCtx(fnCLEAN))
	Register("CODE", NoCtx(fnCODE))
	Register("CONCAT", NoCtx(fnCONCAT))
	Register("CONCATENATE", NoCtx(fnCONCATENATE))
	Register("EXACT", NoCtx(fnEXACT))
	Register("FIND", NoCtx(fnFIND))
	Register("FIXED", NoCtx(fnFIXED))
	Register("LEFT", NoCtx(fnLEFT))
	Register("LEN", NoCtx(fnLEN))
	Register("LOWER", NoCtx(fnLOWER))
	Register("MID", NoCtx(fnMID))
	Register("NUMBERVALUE", NoCtx(fnNUMBERVALUE))
	Register("PROPER", NoCtx(fnPROPER))
	Register("REPLACE", NoCtx(fnREPLACE))
	Register("REPT", NoCtx(fnREPT))
	Register("RIGHT", NoCtx(fnRIGHT))
	Register("SEARCH", NoCtx(fnSEARCH))
	Register("SUBSTITUTE", NoCtx(fnSUBSTITUTE))
	Register("T", NoCtx(fnT))
	Register("TEXT", fnTEXTCtx)
	Register("TEXTJOIN", NoCtx(fnTEXTJOIN))
	Register("TRIM", NoCtx(fnTRIM))
	Register("UPPER", NoCtx(fnUPPER))
	Register("VALUE", NoCtx(fnVALUEFn))
}

func fnCHOOSE(args []Value) (Value, error) {
	if len(args) < 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	idx, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	i := int(idx)
	if i < 1 || i > len(args)-1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return args[i], nil
}

func fnCONCAT(args []Value) (Value, error) {
	var b strings.Builder
	for _, arg := range args {
		if arg.Type == ValueError {
			return arg, nil
		}
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return cell, nil
					}
					b.WriteString(ValueToString(cell))
				}
			}
		} else {
			b.WriteString(ValueToString(arg))
		}
	}
	return StringVal(b.String()), nil
}

func fnCONCATENATE(args []Value) (Value, error) {
	var b strings.Builder
	for _, arg := range args {
		if arg.Type == ValueError {
			return arg, nil
		}
		b.WriteString(ValueToString(arg))
	}
	return StringVal(b.String()), nil
}

func fnFIND(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	findText := ValueToString(args[0])
	withinText := ValueToString(args[1])
	startNum := 1
	if len(args) == 3 {
		sn, e := CoerceNum(args[2])
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

func fnLEFT(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := ValueToString(args[0])
	n := 1
	if len(args) == 2 {
		num, e := CoerceNum(args[1])
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
	s := ValueToString(args[0])
	return NumberVal(float64(utf8.RuneCountInString(s))), nil
}

func fnLOWER(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	return StringVal(strings.ToLower(ValueToString(args[0]))), nil
}

func fnMID(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := ValueToString(args[0])
	startNum, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	numChars, e := CoerceNum(args[2])
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

func fnRIGHT(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := ValueToString(args[0])
	n := 1
	if len(args) == 2 {
		num, e := CoerceNum(args[1])
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

func fnSUBSTITUTE(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	text := ValueToString(args[0])
	oldText := ValueToString(args[1])
	newText := ValueToString(args[2])

	if len(args) == 4 {
		instanceNum, e := CoerceNum(args[3])
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

	if oldText == "" {
		return StringVal(text), nil
	}
	return StringVal(strings.ReplaceAll(text, oldText, newText)), nil
}

// fnTEXTCtx is the context-aware TEXT function that respects the 1904 date system.
func fnTEXTCtx(args []Value, ctx *EvalContext) (Value, error) {
	d1904 := ctx != nil && ctx.Date1904
	return fnTEXTWith1904(args, d1904)
}

func fnTEXTWith1904(args []Value, date1904 bool) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	format := ValueToString(args[1])
	v := args[0]

	// Excel rejects format strings containing lowercase 'e+' or 'e-'
	// (only uppercase 'E' triggers scientific notation).
	if hasInvalidLowercaseE(format) {
		return ErrorVal(ErrValVALUE), nil
	}

	// Check if the format has a text section (4th section).
	sections := splitFormatSections(format)

	// For non-numeric string values, use the text section if available,
	// or the @ placeholder in the format string.
	if v.Type == ValueString && v.Str != "" {
		n, e := CoerceNum(v)
		if e != nil {
			// Can't coerce to number — use text section if available.
			if len(sections) >= 4 {
				return StringVal(formatTextSection(v.Str, sections[3])), nil
			}
			// Check if any section contains the @ text placeholder.
			for _, sec := range sections {
				if sectionContainsAt(sec) {
					return StringVal(formatTextSection(v.Str, sec)), nil
				}
			}
			return *e, nil
		}
		return StringVal(formatExcelNumber(n, format, date1904)), nil
	}

	// Booleans: Excel's TEXT function always returns "TRUE" or "FALSE"
	// regardless of the format string — booleans are never formatted numerically.
	if v.Type == ValueBool {
		text := "TRUE"
		if !v.Bool {
			text = "FALSE"
		}
		return StringVal(text), nil
	}

	n, e := CoerceNum(v)
	if e != nil {
		return *e, nil
	}

	// If no section contains any number format codes (0, #, ?, E) or
	// date/time codes, the format is invalid for numeric values.
	if !strings.EqualFold(format, "General") && !anyNumberFormatCodes(sections) {
		return ErrorVal(ErrValVALUE), nil
	}

	return StringVal(formatExcelNumber(n, format, date1904)), nil
}

func FormatWithCommas(n float64, decimals int) string {
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

func fnTRIM(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := ValueToString(args[0])
	fields := strings.Fields(s)
	return StringVal(strings.Join(fields, " ")), nil
}

func fnUPPER(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(strings.ToUpper(ValueToString(args[0]))), nil
}

func fnCHAR(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	code := int(n)
	if code < 1 || code > 255 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(string(rune(code))), nil
}

func fnCLEAN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := ValueToString(args[0])
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
	s := ValueToString(args[0])
	if len(s) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	r, _ := utf8.DecodeRuneInString(s)
	return NumberVal(float64(r)), nil
}

func fnEXACT(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(ValueToString(args[0]) == ValueToString(args[1])), nil
}

func fnFIXED(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	decimals := 2
	if len(args) >= 2 {
		d, e := CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		decimals = int(d)
	}

	noCommas := false
	if len(args) >= 3 {
		noCommas = IsTruthy(args[2])
	}

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
	return StringVal(FormatWithCommas(n, decimals)), nil
}

func fnNUMBERVALUE(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	text := ValueToString(args[0])

	decSep := "."
	grpSep := ","
	if len(args) >= 2 {
		ds := ValueToString(args[1])
		if len(ds) > 0 {
			decSep = string(ds[0])
		}
	}
	if len(args) >= 3 {
		gs := ValueToString(args[2])
		if len(gs) > 0 {
			grpSep = string(gs[0])
		}
	}

	text = strings.Map(func(r rune) rune {
		if unicode.IsSpace(r) {
			return -1
		}
		return r
	}, text)

	if text == "" {
		return NumberVal(0), nil
	}

	percentCount := strings.Count(text, "%")
	text = strings.ReplaceAll(text, "%", "")

	decIdx := strings.Index(text, decSep)
	if decIdx >= 0 {
		after := text[decIdx+len(decSep):]
		if strings.Contains(after, grpSep) {
			return ErrorVal(ErrValVALUE), nil
		}
	}

	text = strings.ReplaceAll(text, grpSep, "")

	if strings.Count(text, decSep) > 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	text = strings.Replace(text, decSep, ".", 1)

	num, err := strconv.ParseFloat(text, 64)
	if err != nil {
		return ErrorVal(ErrValVALUE), nil
	}

	for i := 0; i < percentCount; i++ {
		num /= 100
	}

	return NumberVal(num), nil
}

func fnPROPER(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	s := ValueToString(args[0])
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
	oldText := ValueToString(args[0])
	startNum, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	numChars, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	newText := ValueToString(args[3])

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
	s := ValueToString(args[0])
	n, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	count := int(n)
	if count < 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(strings.Repeat(s, count)), nil
}

func fnSEARCH(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	findText := strings.ToLower(ValueToString(args[0]))
	withinText := strings.ToLower(ValueToString(args[1]))
	startNum := 1
	if len(args) == 3 {
		sn, e := CoerceNum(args[2])
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

	// Check if the pattern contains wildcards or tilde escapes.
	hasSpecial := strings.ContainsAny(findText, "*?~")

	if !hasSpecial {
		idx := strings.Index(remaining, findText)
		if idx < 0 {
			return ErrorVal(ErrValVALUE), nil
		}
		runeIdx := utf8.RuneCountInString(remaining[:idx])
		return NumberVal(float64(start + runeIdx + 1)), nil
	}

	// Convert Excel wildcard pattern to a regexp.
	re, err := excelPatternToRegexp(findText)
	if err != nil {
		return ErrorVal(ErrValVALUE), nil
	}
	loc := re.FindStringIndex(remaining)
	if loc == nil {
		return ErrorVal(ErrValVALUE), nil
	}
	runeIdx := utf8.RuneCountInString(remaining[:loc[0]])
	return NumberVal(float64(start + runeIdx + 1)), nil
}

// excelPatternToRegexp converts an Excel SEARCH wildcard pattern to a Go regexp.
// * matches any sequence of characters, ? matches exactly one character.
// ~* and ~? match literal * and ? respectively. ~~ matches a literal ~.
func excelPatternToRegexp(pattern string) (*regexp.Regexp, error) {
	var b strings.Builder
	runes := []rune(pattern)
	for i := 0; i < len(runes); i++ {
		ch := runes[i]
		if ch == '~' && i+1 < len(runes) {
			next := runes[i+1]
			if next == '*' || next == '?' || next == '~' {
				b.WriteString(regexp.QuoteMeta(string(next)))
				i++
				continue
			}
		}
		switch ch {
		case '*':
			b.WriteString(".*")
		case '?':
			b.WriteString(".")
		default:
			b.WriteString(regexp.QuoteMeta(string(ch)))
		}
	}
	return regexp.Compile(b.String())
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

func fnTEXTJOIN(args []Value) (Value, error) {
	if len(args) < 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	delimiter := ValueToString(args[0])
	ignoreEmpty := IsTruthy(args[1])

	var parts []string
	for _, arg := range args[2:] {
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if ignoreEmpty && (cell.Type == ValueEmpty || (cell.Type == ValueString && cell.Str == "")) {
						continue
					}
					parts = append(parts, ValueToString(cell))
				}
			}
		} else {
			if ignoreEmpty && (arg.Type == ValueEmpty || (arg.Type == ValueString && arg.Str == "")) {
				continue
			}
			parts = append(parts, ValueToString(arg))
		}
	}

	result := strings.Join(parts, delimiter)
	if len(result) > 32767 {
		return ErrorVal(ErrValVALUE), nil
	}
	return StringVal(result), nil
}

func fnVALUEFn(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueNumber {
		return args[0], nil
	}
	s := ValueToString(args[0])
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
