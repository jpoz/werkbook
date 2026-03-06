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
	Register("DOLLAR", NoCtx(fnDOLLAR))
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
	Register("ROMAN", NoCtx(fnROMAN))
	Register("SEARCH", NoCtx(fnSEARCH))
	Register("SUBSTITUTE", NoCtx(fnSUBSTITUTE))
	Register("T", NoCtx(fnT))
	Register("TEXT", fnTEXTCtx)
	Register("TEXTAFTER", NoCtx(fnTextAfter))
	Register("TEXTBEFORE", NoCtx(fnTextBefore))
	Register("TEXTJOIN", NoCtx(fnTEXTJOIN))
	Register("TEXTSPLIT", NoCtx(fnTextSplit))
	Register("TRIM", NoCtx(fnTRIM))
	Register("UPPER", NoCtx(fnUPPER))
	Register("VALUE", NoCtx(fnVALUEFn))
	Register("VALUETOTEXT", NoCtx(fnValueToText))
	Register("ARRAYTOTEXT", NoCtx(fnArrayToText))
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

func fnDOLLAR(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	decimals := 2
	if len(args) == 2 {
		d, e := CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		decimals = int(d)
	}

	// For negative decimals, round to the left of the decimal point.
	if decimals < 0 {
		factor := math.Pow(10, float64(-decimals))
		n = math.Round(n/factor) * factor
		decimals = 0
	} else {
		factor := math.Pow(10, float64(decimals))
		n = math.Round(n*factor) / factor
	}

	// Handle negative zero: after rounding, n may be -0.
	negative := n < 0
	if negative {
		n = -n
	}
	// Ensure -0.0 becomes +0.0.
	n = n + 0

	formatted := FormatWithCommas(n, decimals)
	if negative {
		return StringVal("($" + formatted + ")"), nil
	}
	return StringVal("$" + formatted), nil
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

	// Booleans: Excel's TEXT function returns "TRUE" or "FALSE" for numeric
	// formats, but uses the text section (4th section) if the format has one.
	if v.Type == ValueBool {
		text := "TRUE"
		if !v.Bool {
			text = "FALSE"
		}
		if len(sections) >= 4 {
			return StringVal(formatTextSection(text, sections[3])), nil
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

	// Excel rejects number format strings that contain unquoted alphabetic
	// characters (outside of date/time codes, scientific E+/E-, color codes,
	// etc.).  For example "Value: 0" is invalid because "Value" is not quoted.
	if !strings.EqualFold(format, "General") {
		for _, sec := range sections {
			stripped := stripColorCodes(sec)
			if !isDateTimeFormat(stripped) && !isElapsedTimeFormat(stripped) &&
				sectionHasNumberCodes(stripped) && sectionHasUnquotedLetters(stripped) {
				return ErrorVal(ErrValVALUE), nil
			}
		}
	}

	// "@" format with a numeric value: convert number to string using
	// the text section (@ is the text placeholder).
	if sectionContainsAt(format) && !strings.ContainsAny(format, "0#?") {
		return StringVal(formatTextSection(excelNumberToString(n), format)), nil
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
	if args[0].Type == ValueError {
		return args[0], nil
	}
	return StringVal(strings.ToUpper(ValueToString(args[0]))), nil
}

// win1252ToUnicode maps Windows-1252 byte values 0x80–0x9F to their Unicode
// code points. Bytes outside this range map 1:1 with Unicode.
var win1252ToUnicode = [32]rune{
	0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, // 80-87
	0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x017D, 0x008F, // 88-8F
	0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, // 90-97
	0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178, // 98-9F
}

// unicodeToWin1252 is the reverse mapping from Unicode code points back to
// Windows-1252 byte values for the 0x80–0x9F range.
var unicodeToWin1252 map[rune]byte

func init() {
	unicodeToWin1252 = make(map[rune]byte, 32)
	for i, r := range win1252ToUnicode {
		unicodeToWin1252[r] = byte(0x80 + i)
	}
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
	// Excel CHAR uses Windows-1252 encoding. Bytes 0x80–0x9F map to
	// different Unicode code points than their byte value.
	if code >= 0x80 && code <= 0x9F {
		return StringVal(string(win1252ToUnicode[code-0x80])), nil
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
	// Excel CODE uses Windows-1252 encoding.
	// Characters in the Windows-1252 0x80–0x9F range have different
	// Unicode code points, so we need to map them back.
	if b, ok := unicodeToWin1252[r]; ok {
		return NumberVal(float64(b)), nil
	}
	// Characters 0x00–0x7F and 0xA0–0xFF are identical in Windows-1252
	// and Unicode (Latin-1). Characters outside the Windows-1252 range
	// (e.g. CJK) get mapped to '_' (95), matching Excel's behavior
	// which uses underscore as the default substitution character for
	// unmappable characters.
	if r > 0xFF {
		return NumberVal(95), nil // '_'
	}
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
	if args[0].Type == ValueError {
		return args[0], nil
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

// romanPair maps an integer value to its Roman numeral representation.
type romanPair struct {
	val int
	sym string
}

// romanTables contains the value-to-symbol tables for each ROMAN form level (0-4).
// Form 0 is classic Roman numerals; higher forms allow increasingly non-standard
// subtractive pairs for more compact output.
var romanTables = [5][]romanPair{
	// Form 0: Classic
	{
		{1000, "M"}, {900, "CM"}, {500, "D"}, {400, "CD"},
		{100, "C"}, {90, "XC"}, {50, "L"}, {40, "XL"},
		{10, "X"}, {9, "IX"}, {5, "V"}, {4, "IV"}, {1, "I"},
	},
	// Form 1
	{
		{1000, "M"}, {950, "LM"}, {900, "CM"}, {500, "D"}, {450, "LD"}, {400, "CD"},
		{100, "C"}, {95, "VC"}, {90, "XC"}, {50, "L"}, {45, "VL"}, {40, "XL"},
		{10, "X"}, {9, "IX"}, {5, "V"}, {4, "IV"}, {1, "I"},
	},
	// Form 2
	{
		{1000, "M"}, {990, "XM"}, {950, "LM"}, {900, "CM"},
		{500, "D"}, {490, "XD"}, {450, "LD"}, {400, "CD"},
		{100, "C"}, {99, "IC"}, {95, "VC"}, {90, "XC"},
		{50, "L"}, {49, "IL"}, {45, "VL"}, {40, "XL"},
		{10, "X"}, {9, "IX"}, {5, "V"}, {4, "IV"}, {1, "I"},
	},
	// Form 3
	{
		{1000, "M"}, {995, "VM"}, {990, "XM"}, {950, "LM"}, {900, "CM"},
		{500, "D"}, {495, "VD"}, {490, "XD"}, {450, "LD"}, {400, "CD"},
		{100, "C"}, {99, "IC"}, {95, "VC"}, {90, "XC"},
		{50, "L"}, {49, "IL"}, {45, "VL"}, {40, "XL"},
		{10, "X"}, {9, "IX"}, {5, "V"}, {4, "IV"}, {1, "I"},
	},
	// Form 4: Simplified (most concise)
	{
		{1000, "M"}, {999, "IM"}, {995, "VM"}, {990, "XM"}, {950, "LM"}, {900, "CM"},
		{500, "D"}, {499, "ID"}, {495, "VD"}, {490, "XD"}, {450, "LD"}, {400, "CD"},
		{100, "C"}, {99, "IC"}, {95, "VC"}, {90, "XC"},
		{50, "L"}, {49, "IL"}, {45, "VL"}, {40, "XL"},
		{10, "X"}, {9, "IX"}, {5, "V"}, {4, "IV"}, {1, "I"},
	},
}

func fnROMAN(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	number := int(n)

	if number < 0 || number > 3999 {
		return ErrorVal(ErrValVALUE), nil
	}
	if number == 0 {
		return StringVal(""), nil
	}

	form := 0
	if len(args) == 2 {
		// TRUE -> 0 (Classic), FALSE -> 4 (Simplified)
		if args[1].Type == ValueBool {
			if args[1].Bool {
				form = 0
			} else {
				form = 4
			}
		} else {
			f, e := CoerceNum(args[1])
			if e != nil {
				return *e, nil
			}
			form = int(f)
		}
	}
	if form < 0 {
		form = 0
	}
	if form > 4 {
		form = 4
	}

	table := romanTables[form]
	var b strings.Builder
	rem := number
	for _, p := range table {
		for rem >= p.val {
			b.WriteString(p.sym)
			rem -= p.val
		}
	}
	return StringVal(b.String()), nil
}

// valueToTextFormat formats a single value for VALUETOTEXT / ARRAYTOTEXT.
// When strict is true, strings are wrapped in escaped double-quotes.
func valueToTextFormat(v Value, strict bool) string {
	switch v.Type {
	case ValueNumber:
		return excelNumberToString(v.Num)
	case ValueString:
		if strict {
			return "\"" + v.Str + "\""
		}
		return v.Str
	case ValueBool:
		if v.Bool {
			return "TRUE"
		}
		return "FALSE"
	case ValueEmpty:
		return ""
	default:
		return ""
	}
}

func fnValueToText(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	v := args[0]
	if v.Type == ValueError {
		return v, nil
	}

	format := 0
	if len(args) == 2 {
		f, e := CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		format = int(f)
		if format != 0 && format != 1 {
			return ErrorVal(ErrValVALUE), nil
		}
	}

	return StringVal(valueToTextFormat(v, format == 1)), nil
}

func fnArrayToText(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	format := 0
	if len(args) == 2 {
		f, e := CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		format = int(f)
		if format != 0 && format != 1 {
			return ErrorVal(ErrValVALUE), nil
		}
	}

	strict := format == 1

	arg := args[0]
	if arg.Type == ValueError {
		return arg, nil
	}

	if strict {
		// Strict: {v1,v2;v3,v4} with rows separated by ";" and columns by ","
		var rows []string
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				var cols []string
				for _, cell := range row {
					cols = append(cols, valueToTextFormat(cell, true))
				}
				rows = append(rows, strings.Join(cols, ","))
			}
		} else {
			rows = append(rows, valueToTextFormat(arg, true))
		}
		return StringVal("{" + strings.Join(rows, ";") + "}"), nil
	}

	// Concise: join all values with ", "
	var vals []Value
	if arg.Type == ValueArray {
		for _, row := range arg.Array {
			vals = append(vals, row...)
		}
	} else {
		vals = append(vals, arg)
	}
	var parts []string
	for _, v := range vals {
		parts = append(parts, valueToTextFormat(v, false))
	}
	return StringVal(strings.Join(parts, ", ")), nil
}

// textFindAllOccurrences finds all starting byte positions of delimiter in text.
// If caseInsensitive is true, matching is case-insensitive.
func textFindAllOccurrences(text, delimiter string, caseInsensitive bool) []int {
	if delimiter == "" {
		return nil
	}
	searchText := text
	searchDelim := delimiter
	if caseInsensitive {
		searchText = strings.ToLower(text)
		searchDelim = strings.ToLower(delimiter)
	}
	var positions []int
	start := 0
	for {
		idx := strings.Index(searchText[start:], searchDelim)
		if idx < 0 {
			break
		}
		positions = append(positions, start+idx)
		start += idx + len(searchDelim)
	}
	return positions
}

func fnTextBefore(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 6 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors from text arg.
	if args[0].Type == ValueError {
		return args[0], nil
	}
	text := ValueToString(args[0])

	// Propagate errors from delimiter arg.
	if args[1].Type == ValueError {
		return args[1], nil
	}
	delimiter := ValueToString(args[1])

	instanceNum := 1
	if len(args) >= 3 {
		n, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		instanceNum = int(n)
	}
	if instanceNum == 0 {
		return ErrorVal(ErrValVALUE), nil
	}

	matchMode := 0
	if len(args) >= 4 {
		m, e := CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
		matchMode = int(m)
	}

	matchEnd := 0
	if len(args) >= 5 {
		me, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		matchEnd = int(me)
	}

	var ifNotFound *Value
	if len(args) >= 6 {
		v := args[5]
		ifNotFound = &v
	}

	caseInsensitive := matchMode == 1

	// Empty delimiter: TEXTBEFORE returns "" for instance 1, and handles
	// negative instances as counting empty positions from the end.
	if delimiter == "" {
		if instanceNum == 1 || instanceNum == -1 {
			return StringVal(""), nil
		}
		runes := []rune(text)
		if instanceNum > 0 {
			// instance 2 means after 1st empty boundary = position 1, etc.
			pos := instanceNum - 1
			if pos > len(runes) {
				if ifNotFound != nil {
					return *ifNotFound, nil
				}
				return ErrorVal(ErrValNA), nil
			}
			return StringVal(string(runes[:pos])), nil
		}
		// Negative: count from end. -2 means 2nd from end.
		pos := len(runes) + instanceNum + 1
		if pos < 0 || pos > len(runes) {
			if ifNotFound != nil {
				return *ifNotFound, nil
			}
			return ErrorVal(ErrValNA), nil
		}
		return StringVal(string(runes[:pos])), nil
	}

	positions := textFindAllOccurrences(text, delimiter, caseInsensitive)

	if len(positions) == 0 {
		// Not found.
		if matchEnd == 1 {
			return StringVal(text), nil
		}
		if ifNotFound != nil {
			return *ifNotFound, nil
		}
		return ErrorVal(ErrValNA), nil
	}

	var pos int
	if instanceNum > 0 {
		if instanceNum > len(positions) {
			if matchEnd == 1 && instanceNum == len(positions)+1 {
				return StringVal(text), nil
			}
			if ifNotFound != nil {
				return *ifNotFound, nil
			}
			return ErrorVal(ErrValNA), nil
		}
		pos = positions[instanceNum-1]
	} else {
		// Negative: count from end. -1 = last occurrence.
		idx := len(positions) + instanceNum
		if idx < 0 {
			if ifNotFound != nil {
				return *ifNotFound, nil
			}
			return ErrorVal(ErrValNA), nil
		}
		pos = positions[idx]
	}

	return StringVal(text[:pos]), nil
}

func fnTextAfter(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 6 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Propagate errors from text arg.
	if args[0].Type == ValueError {
		return args[0], nil
	}
	text := ValueToString(args[0])

	// Propagate errors from delimiter arg.
	if args[1].Type == ValueError {
		return args[1], nil
	}
	delimiter := ValueToString(args[1])

	instanceNum := 1
	if len(args) >= 3 {
		n, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		instanceNum = int(n)
	}
	if instanceNum == 0 {
		return ErrorVal(ErrValVALUE), nil
	}

	matchMode := 0
	if len(args) >= 4 {
		m, e := CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
		matchMode = int(m)
	}

	matchEnd := 0
	if len(args) >= 5 {
		me, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		matchEnd = int(me)
	}

	var ifNotFound *Value
	if len(args) >= 6 {
		v := args[5]
		ifNotFound = &v
	}

	caseInsensitive := matchMode == 1

	// Empty delimiter: TEXTAFTER returns full text for instance 1,
	// and handles other instances similarly to TEXTBEFORE.
	if delimiter == "" {
		if instanceNum == 1 || instanceNum == -1 {
			return StringVal(text), nil
		}
		runes := []rune(text)
		if instanceNum > 0 {
			pos := instanceNum - 1
			if pos > len(runes) {
				if ifNotFound != nil {
					return *ifNotFound, nil
				}
				return ErrorVal(ErrValNA), nil
			}
			return StringVal(string(runes[pos:])), nil
		}
		// Negative
		pos := len(runes) + instanceNum + 1
		if pos < 0 || pos > len(runes) {
			if ifNotFound != nil {
				return *ifNotFound, nil
			}
			return ErrorVal(ErrValNA), nil
		}
		return StringVal(string(runes[pos:])), nil
	}

	positions := textFindAllOccurrences(text, delimiter, caseInsensitive)

	if len(positions) == 0 {
		// Not found.
		if matchEnd == 1 {
			return StringVal(""), nil
		}
		if ifNotFound != nil {
			return *ifNotFound, nil
		}
		return ErrorVal(ErrValNA), nil
	}

	var pos int
	if instanceNum > 0 {
		if instanceNum > len(positions) {
			if matchEnd == 1 && instanceNum == len(positions)+1 {
				return StringVal(""), nil
			}
			if ifNotFound != nil {
				return *ifNotFound, nil
			}
			return ErrorVal(ErrValNA), nil
		}
		pos = positions[instanceNum-1]
	} else {
		// Negative: count from end. -1 = last occurrence.
		idx := len(positions) + instanceNum
		if idx < 0 {
			if ifNotFound != nil {
				return *ifNotFound, nil
			}
			return ErrorVal(ErrValNA), nil
		}
		pos = positions[idx]
	}

	return StringVal(text[pos+len(delimiter):]), nil
}

// collectDelimiters extracts a list of delimiter strings from a Value.
// If the value is an array, all non-empty string elements are collected.
// Otherwise the single string value is returned.
func collectDelimiters(v Value) []string {
	if v.Type == ValueArray {
		var delims []string
		for _, row := range v.Array {
			for _, cell := range row {
				delims = append(delims, ValueToString(cell))
			}
		}
		return delims
	}
	return []string{ValueToString(v)}
}

// textSplitByDelimiters splits text by any of the given delimiters.
// When caseInsensitive is true, matching is case-insensitive but the
// original text segments are preserved. Returns the split parts.
func textSplitByDelimiters(text string, delimiters []string, caseInsensitive bool) []string {
	if len(delimiters) == 0 {
		return []string{text}
	}

	searchText := text
	searchDelims := delimiters
	if caseInsensitive {
		searchText = strings.ToLower(text)
		searchDelims = make([]string, len(delimiters))
		for i, d := range delimiters {
			searchDelims[i] = strings.ToLower(d)
		}
	}

	var parts []string
	start := 0
	for start <= len(text) {
		// Find the earliest match among all delimiters.
		bestIdx := -1
		bestLen := 0
		for i, d := range searchDelims {
			if d == "" {
				continue
			}
			idx := strings.Index(searchText[start:], d)
			if idx >= 0 && (bestIdx == -1 || idx < bestIdx || (idx == bestIdx && len(delimiters[i]) > bestLen)) {
				bestIdx = idx
				bestLen = len(delimiters[i])
			}
		}
		if bestIdx < 0 {
			// No more delimiters found; take the rest.
			parts = append(parts, text[start:])
			break
		}
		parts = append(parts, text[start:start+bestIdx])
		start += bestIdx + bestLen
		// If delimiter is at the very end, append a trailing empty part.
		if start == len(text) {
			parts = append(parts, "")
			break
		}
	}
	return parts
}

func fnTextSplit(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 6 {
		return ErrorVal(ErrValVALUE), nil
	}

	// arg 0: text
	if args[0].Type == ValueError {
		return args[0], nil
	}
	text := ValueToString(args[0])

	// arg 1: col_delimiter
	if args[1].Type == ValueError {
		return args[1], nil
	}
	colDelims := collectDelimiters(args[1])

	// arg 2: row_delimiter (optional)
	var rowDelims []string
	hasRowDelim := false
	if len(args) >= 3 && args[2].Type != ValueEmpty {
		if args[2].Type == ValueError {
			return args[2], nil
		}
		rowDelims = collectDelimiters(args[2])
		hasRowDelim = true
	}

	// arg 3: ignore_empty (optional, default FALSE)
	ignoreEmpty := false
	if len(args) >= 4 && args[3].Type != ValueEmpty {
		if args[3].Type == ValueError {
			return args[3], nil
		}
		ignoreEmpty = IsTruthy(args[3])
	}

	// arg 4: match_mode (optional, default 0 = case-sensitive)
	caseInsensitive := false
	if len(args) >= 5 && args[4].Type != ValueEmpty {
		if args[4].Type == ValueError {
			return args[4], nil
		}
		m, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		caseInsensitive = int(m) == 1
	}

	// arg 5: pad_with (optional, default #N/A)
	padWith := ErrorVal(ErrValNA)
	if len(args) >= 6 && args[5].Type != ValueEmpty {
		padWith = args[5]
	}

	// Split into rows first, then columns.
	var rowTexts []string
	if hasRowDelim {
		rowTexts = textSplitByDelimiters(text, rowDelims, caseInsensitive)
	} else {
		rowTexts = []string{text}
	}

	// Filter empty row texts if ignore_empty.
	if ignoreEmpty {
		filtered := rowTexts[:0]
		for _, rt := range rowTexts {
			if rt != "" {
				filtered = append(filtered, rt)
			}
		}
		rowTexts = filtered
	}

	// Split each row by col_delimiter.
	var rows [][]string
	for _, rt := range rowTexts {
		cols := textSplitByDelimiters(rt, colDelims, caseInsensitive)
		if ignoreEmpty {
			filtered := cols[:0]
			for _, c := range cols {
				if c != "" {
					filtered = append(filtered, c)
				}
			}
			cols = filtered
		}
		rows = append(rows, cols)
	}

	// If everything was filtered out, return empty string.
	if len(rows) == 0 {
		return StringVal(""), nil
	}

	// Find the maximum column count for padding.
	maxCols := 0
	for _, r := range rows {
		if len(r) > maxCols {
			maxCols = len(r)
		}
	}
	if maxCols == 0 {
		return StringVal(""), nil
	}

	// Build the result array.
	result := make([][]Value, len(rows))
	for i, r := range rows {
		result[i] = make([]Value, maxCols)
		for j := 0; j < maxCols; j++ {
			if j < len(r) {
				result[i][j] = StringVal(r[j])
			} else {
				result[i][j] = padWith
			}
		}
	}

	// If the result is 1x1, return the scalar.
	if len(result) == 1 && len(result[0]) == 1 {
		return result[0][0], nil
	}

	return Value{Type: ValueArray, Array: result}, nil
}
