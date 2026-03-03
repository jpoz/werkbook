package formula

import (
	"fmt"
	"strings"
	"unicode/utf8"
)

func init() {
	Register("CHOOSE", NoCtx(fnCHOOSE))
	Register("CONCAT", NoCtx(fnCONCATENATE))
	Register("CONCATENATE", NoCtx(fnCONCATENATE))
	Register("FIND", NoCtx(fnFIND))
	Register("LEFT", NoCtx(fnLEFT))
	Register("LEN", NoCtx(fnLEN))
	Register("LOWER", NoCtx(fnLOWER))
	Register("MID", NoCtx(fnMID))
	Register("RIGHT", NoCtx(fnRIGHT))
	Register("SUBSTITUTE", NoCtx(fnSUBSTITUTE))
	Register("TEXT", NoCtx(fnTEXT))
	Register("TRIM", NoCtx(fnTRIM))
	Register("UPPER", NoCtx(fnUPPER))
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

func fnTEXT(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	format := ValueToString(args[1])
	return StringVal(formatExcelNumber(n, format)), nil
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
