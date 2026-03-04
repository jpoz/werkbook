package formula

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"
)

// ---------------------------------------------------------------------------
// Excel number format engine
//
// Supports the format codes used by TEXT(), cell number formats, etc.
// Reference: https://support.microsoft.com/en-us/office/number-format-codes
//
// Features implemented:
//   - Section separator ; (positive;negative;zero;text)
//   - Literal text: "quoted", \escaped, and passthrough of $ - + / ( ) : ! ^ & ' ~ { } = < > space
//   - Date/time codes: d dd ddd dddd m mm mmm mmmm yy yyyy h hh m mm s ss AM/PM
//   - Elapsed time: [h] [m] [s] with optional decimal seconds
//   - Number codes: 0 # . , % E+/E-
//   - Fraction codes: # #/# etc
//   - General format
// ---------------------------------------------------------------------------

// monthNames for mmm/mmmm codes.
var shortMonths = [13]string{"", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
var longMonths = [13]string{"", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
var shortDays = [8]string{"", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"}
var longDays = [8]string{"", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"}

// formatExcelNumber formats a number using an Excel format string.
// This is the main entry point used by the TEXT() function.
func formatExcelNumber(n float64, format string) string {
	if strings.EqualFold(format, "General") {
		return formatGeneral(n)
	}

	// Split by unquoted semicolons into sections.
	sections := splitFormatSections(format)

	// Select the appropriate section based on the value's sign.
	var section string
	switch len(sections) {
	case 1:
		section = sections[0]
	case 2:
		// positive/zero ; negative
		if n < 0 {
			section = sections[1]
			n = -n // section handles the sign via literal or implicit
		} else {
			section = sections[0]
		}
	case 3:
		// positive ; negative ; zero
		if n > 0 {
			section = sections[0]
		} else if n < 0 {
			section = sections[1]
			n = -n
		} else {
			section = sections[2]
		}
	default:
		// 4+ sections: positive ; negative ; zero ; text
		// TEXT() always passes a number, so the 4th section is never reached here.
		if n > 0 {
			section = sections[0]
		} else if n < 0 {
			section = sections[1]
			n = -n
		} else {
			section = sections[2]
		}
	}

	if section == "" {
		return ""
	}

	// Check for elapsed time format [h], [m], [s] — must be checked before
	// general date/time because [h]:mm:ss also contains h/m/s tokens.
	if isElapsedTimeFormat(section) {
		return formatElapsedTime(n, section)
	}

	// Check if this is a date/time format.
	if isDateTimeFormat(section) {
		return formatDateTime(n, section)
	}

	// Check for fraction format.
	if isFractionFormat(section) {
		return formatFraction(n, section)
	}

	// Number format.
	return formatNumberSection(n, section)
}

// formatTextSection formats a text value using the text section (4th section)
// of an Excel format string. The '@' placeholder is replaced with the text value.
func formatTextSection(text string, section string) string {
	if section == "" {
		return text
	}
	var result strings.Builder
	for i := 0; i < len(section); i++ {
		ch := section[i]
		switch {
		case ch == '@':
			result.WriteString(text)
		case ch == '"':
			i++
			for i < len(section) && section[i] != '"' {
				result.WriteByte(section[i])
				i++
			}
		case ch == '\\' && i+1 < len(section):
			i++
			result.WriteByte(section[i])
		default:
			result.WriteByte(ch)
		}
	}
	return result.String()
}

// formatGeneral returns the "General" format representation.
func formatGeneral(n float64) string {
	if n == math.Trunc(n) && math.Abs(n) < 1e15 {
		return strconv.FormatFloat(n, 'f', -1, 64)
	}
	// Use %g-like formatting but with up to 10 significant digits like Excel.
	s := strconv.FormatFloat(n, 'G', 10, 64)
	return s
}

// splitFormatSections splits a format string by unquoted, unescaped semicolons.
func splitFormatSections(format string) []string {
	var sections []string
	var current strings.Builder
	inQuote := false

	for i := 0; i < len(format); i++ {
		ch := format[i]
		if ch == '"' {
			inQuote = !inQuote
			current.WriteByte(ch)
		} else if ch == '\\' && i+1 < len(format) && !inQuote {
			current.WriteByte(ch)
			i++
			current.WriteByte(format[i])
		} else if ch == ';' && !inQuote {
			sections = append(sections, current.String())
			current.Reset()
		} else {
			current.WriteByte(ch)
		}
	}
	sections = append(sections, current.String())
	return sections
}

// ---------------------------------------------------------------------------
// Date/Time detection and formatting
// ---------------------------------------------------------------------------

// dateTokens is the set of characters that indicate a date/time format.
// We check the stripped (no literal) version of the format.
func isDateTimeFormat(format string) bool {
	stripped := stripLiterals(format)
	upper := strings.ToUpper(stripped)

	// Remove elapsed time markers for this check.
	upper = strings.ReplaceAll(upper, "[H]", "")
	upper = strings.ReplaceAll(upper, "[M]", "")
	upper = strings.ReplaceAll(upper, "[S]", "")
	upper = strings.ReplaceAll(upper, "[HH]", "")
	upper = strings.ReplaceAll(upper, "[MM]", "")
	upper = strings.ReplaceAll(upper, "[SS]", "")

	// Look for date/time-specific tokens.
	for i := 0; i < len(upper); i++ {
		ch := upper[i]
		switch ch {
		case 'Y', 'D':
			return true
		case 'H', 'S':
			return true
		case 'M':
			// m is ambiguous (month or minute). If there's a y, d, h, or s
			// elsewhere in the format, it's datetime.
			for j := 0; j < len(upper); j++ {
				if j == i {
					continue
				}
				switch upper[j] {
				case 'Y', 'D', 'H', 'S':
					return true
				}
			}
			// Standalone "m" or "mm" without other date/time codes: could be month.
			// In Excel, a standalone "m" format IS a date format (month of a date).
			return true
		case 'A':
			// Check for AM/PM.
			if i+3 < len(upper) && upper[i:i+4] == "AM/P" {
				return true
			}
			if i+1 < len(upper) && (upper[i:i+2] == "AM" || upper[i:i+2] == "A/") {
				return true
			}
		}
	}
	return false
}

// isElapsedTimeFormat checks if the format contains [h], [m], or [s] elapsed time codes.
func isElapsedTimeFormat(format string) bool {
	// We need to check the raw format for bracket codes, but skip anything inside quotes.
	upper := strings.ToUpper(format)
	inQuote := false
	for i := 0; i < len(upper); i++ {
		if upper[i] == '"' {
			inQuote = !inQuote
			continue
		}
		if inQuote {
			continue
		}
		if upper[i] == '\\' && i+1 < len(upper) {
			i++
			continue
		}
		if upper[i] == '[' {
			end := strings.Index(upper[i:], "]")
			if end > 0 {
				code := upper[i+1 : i+end]
				switch code {
				case "H", "HH", "M", "MM", "S", "SS":
					return true
				}
			}
		}
	}
	return false
}

// stripLiterals removes quoted strings and backslash-escaped chars from a format
// to make token detection easier.
func stripLiterals(format string) string {
	var b strings.Builder
	inQuote := false
	for i := 0; i < len(format); i++ {
		ch := format[i]
		if ch == '"' {
			inQuote = !inQuote
			continue
		}
		if inQuote {
			continue
		}
		if ch == '\\' && i+1 < len(format) {
			i++ // skip escaped char
			continue
		}
		b.WriteByte(ch)
	}
	return b.String()
}

// formatDateTime formats an Excel serial number as a date/time string.
func formatDateTime(serial float64, format string) string {
	t := ExcelSerialToTime(serial)

	// Determine if there's an AM/PM marker.
	stripped := stripLiterals(format)
	upperStripped := strings.ToUpper(stripped)
	hasAMPM := strings.Contains(upperStripped, "AM/PM") || strings.Contains(upperStripped, "A/P")

	hour := t.Hour()
	hour12 := hour % 12
	if hour12 == 0 {
		hour12 = 12
	}
	ampm := "AM"
	if hour >= 12 {
		ampm = "PM"
	}
	ap := "A"
	if hour >= 12 {
		ap = "P"
	}

	minute := t.Minute()
	second := t.Second()
	// Fractional seconds from the serial number.
	frac := serial - math.Floor(serial)
	totalSeconds := frac * 86400
	fracSeconds := totalSeconds - math.Floor(totalSeconds)

	day := t.Day()
	month := int(t.Month())
	year := t.Year()
	weekday := int(t.Weekday())
	if weekday == 0 {
		weekday = 7 // Sunday = 7 for our array indexing
	}

	// Parse the format string token by token and build the result.
	var result strings.Builder
	upper := strings.ToUpper(format)

	// Track whether we've seen 'h' or 's' to disambiguate 'm' as minute vs month.
	// We need to pre-scan to determine this context for each 'm'.
	mContexts := computeMContexts(format)
	mIndex := 0

	i := 0
	for i < len(format) {
		ch := format[i]

		// Handle quoted literals.
		if ch == '"' {
			i++
			for i < len(format) && format[i] != '"' {
				result.WriteByte(format[i])
				i++
			}
			if i < len(format) {
				i++ // skip closing quote
			}
			continue
		}

		// Handle backslash escape.
		if ch == '\\' && i+1 < len(format) {
			i++
			result.WriteByte(format[i])
			i++
			continue
		}

		uCh := upper[i]

		// Date/time tokens.
		switch {
		case uCh == 'Y':
			count := countRun(upper, i, 'Y')
			if count >= 3 {
				result.WriteString(fmt.Sprintf("%04d", year))
			} else {
				result.WriteString(fmt.Sprintf("%02d", year%100))
			}
			i += count
			continue

		case uCh == 'M':
			count := countRun(upper, i, 'M')
			isMinute := false
			if mIndex < len(mContexts) {
				isMinute = mContexts[mIndex]
				mIndex++
			}
			if isMinute {
				if count >= 2 {
					result.WriteString(fmt.Sprintf("%02d", minute))
				} else {
					result.WriteString(strconv.Itoa(minute))
				}
			} else {
				switch count {
				case 1:
					result.WriteString(strconv.Itoa(month))
				case 2:
					result.WriteString(fmt.Sprintf("%02d", month))
				case 3:
					if month >= 1 && month <= 12 {
						result.WriteString(shortMonths[month])
					}
				default: // 4+
					if month >= 1 && month <= 12 {
						result.WriteString(longMonths[month])
					}
				}
			}
			i += count
			continue

		case uCh == 'D':
			count := countRun(upper, i, 'D')
			switch count {
			case 1:
				result.WriteString(strconv.Itoa(day))
			case 2:
				result.WriteString(fmt.Sprintf("%02d", day))
			case 3:
				if weekday >= 1 && weekday <= 7 {
					result.WriteString(shortDays[weekday])
				}
			default: // 4+
				if weekday >= 1 && weekday <= 7 {
					result.WriteString(longDays[weekday])
				}
			}
			i += count
			continue

		case uCh == 'H':
			count := countRun(upper, i, 'H')
			h := hour
			if hasAMPM {
				h = hour12
			}
			if count >= 2 {
				result.WriteString(fmt.Sprintf("%02d", h))
			} else {
				result.WriteString(strconv.Itoa(h))
			}
			i += count
			continue

		case uCh == 'S':
			count := countRun(upper, i, 'S')
			if count >= 2 {
				result.WriteString(fmt.Sprintf("%02d", second))
			} else {
				result.WriteString(strconv.Itoa(second))
			}
			// Check for fractional seconds: s.00 or ss.00
			if i+count < len(format) && format[i+count] == '.' {
				dotPos := i + count
				zeroCount := 0
				j := dotPos + 1
				for j < len(format) && format[j] == '0' {
					zeroCount++
					j++
				}
				if zeroCount > 0 {
					// Format fractional seconds.
					fracStr := fmt.Sprintf("%.*f", zeroCount, fracSeconds)
					// fracStr is like "0.12" — we want ".12"
					if dotIdx := strings.Index(fracStr, "."); dotIdx >= 0 {
						result.WriteString(fracStr[dotIdx:])
					}
					i = j
					continue
				}
			}
			i += count
			continue

		case uCh == 'A':
			// AM/PM or A/P marker.
			if i+4 <= len(upper) && upper[i:i+4] == "AM/P" && i+5 <= len(upper) && upper[i+5-1] == 'M' {
				result.WriteString(ampm)
				i += 5
				continue
			}
			if i+3 <= len(upper) && upper[i:i+3] == "A/P" {
				result.WriteString(ap)
				i += 3
				continue
			}
			// Literal A.
			result.WriteByte(ch)
			i++
			continue

		case uCh == '0' || uCh == '#':
			// Number-like codes in a date format (unusual but possible).
			// E.g. ".00" for fractional seconds when preceded by s — already handled above.
			// For other cases, just pass them through.
			result.WriteByte(ch)
			i++
			continue

		default:
			// Literal passthrough for common characters.
			if isLiteralPassthrough(ch) {
				result.WriteByte(ch)
				i++
				continue
			}
			result.WriteByte(ch)
			i++
			continue
		}
	}

	return result.String()
}

// computeMContexts pre-scans the format to determine for each 'm'/'M' run
// whether it represents minutes (true) or months (false).
// In Excel, 'm' after 'h' and before 's' means minutes; otherwise months.
func computeMContexts(format string) []bool {
	upper := strings.ToUpper(format)
	var results []bool

	// First, find positions of all token runs.
	type tokenRun struct {
		pos   int
		char  byte // 'H', 'M', 'S', 'D', 'Y'
		count int
	}
	var tokens []tokenRun
	inQuote := false
	for i := 0; i < len(upper); i++ {
		ch := upper[i]
		if ch == '"' {
			inQuote = !inQuote
			continue
		}
		if inQuote {
			continue
		}
		if ch == '\\' && i+1 < len(upper) {
			i++
			continue
		}
		switch ch {
		case 'H', 'M', 'S', 'D', 'Y':
			count := countRun(upper, i, ch)
			tokens = append(tokens, tokenRun{pos: i, char: ch, count: count})
			i += count - 1
		}
	}

	// For each M token, determine if preceded by H or followed by S.
	for ti, tok := range tokens {
		if tok.char != 'M' {
			continue
		}
		isMinute := false
		// Look backward for H.
		for j := ti - 1; j >= 0; j-- {
			switch tokens[j].char {
			case 'H':
				isMinute = true
			case 'Y', 'D':
				// A date token between H and M breaks the connection — but
				// in practice Excel still treats m as minute if h appeared before.
			}
			if tokens[j].char == 'H' {
				break
			}
		}
		// Look forward for S.
		for j := ti + 1; j < len(tokens); j++ {
			if tokens[j].char == 'S' {
				isMinute = true
				break
			}
			if tokens[j].char == 'H' || tokens[j].char == 'Y' || tokens[j].char == 'D' {
				break
			}
		}
		results = append(results, isMinute)
	}
	return results
}

// countRun counts how many consecutive occurrences of ch appear at position i.
func countRun(s string, i int, ch byte) int {
	count := 0
	for i+count < len(s) && s[i+count] == ch {
		count++
	}
	return count
}

// isLiteralPassthrough returns true for characters that Excel passes through as-is
// in format strings without needing quotes or backslash.
func isLiteralPassthrough(ch byte) bool {
	switch ch {
	case ' ', '$', '-', '+', '/', '(', ')', ':', '!', '^', '&', '\'', '~',
		'{', '}', '=', '<', '>', ',', '.':
		return true
	}
	return false
}

// ---------------------------------------------------------------------------
// Elapsed time formatting [h]:mm:ss
// ---------------------------------------------------------------------------

func formatElapsedTime(serial float64, format string) string {
	totalSeconds := serial * 86400.0
	negative := totalSeconds < 0
	if negative {
		totalSeconds = -totalSeconds
	}

	totalHours := int(totalSeconds / 3600)
	remaining := totalSeconds - float64(totalHours)*3600
	totalMinutes := int(totalSeconds / 60)
	minutes := int(remaining / 60)
	seconds := remaining - float64(minutes)*60

	upper := strings.ToUpper(format)

	// Pre-scan to determine the primary elapsed bracket code.
	// This affects how bare time codes are interpreted.
	elapsedUnit := detectElapsedUnit(upper)

	var result strings.Builder
	i := 0
	for i < len(format) {
		ch := format[i]
		uCh := upper[i]

		if ch == '"' {
			i++
			for i < len(format) && format[i] != '"' {
				result.WriteByte(format[i])
				i++
			}
			if i < len(format) {
				i++
			}
			continue
		}
		if ch == '\\' && i+1 < len(format) {
			i++
			result.WriteByte(format[i])
			i++
			continue
		}

		if uCh == '[' {
			// Elapsed time code.
			end := strings.Index(upper[i:], "]")
			if end < 0 {
				result.WriteByte(ch)
				i++
				continue
			}
			code := upper[i+1 : i+end]
			i += end + 1
			switch code {
			case "H", "HH":
				result.WriteString(strconv.Itoa(totalHours))
			case "M", "MM":
				result.WriteString(strconv.Itoa(totalMinutes))
			case "S", "SS":
				result.WriteString(strconv.Itoa(int(totalSeconds)))
				// Handle fractional seconds after [s]/[ss].
				i = writeElapsedFracSeconds(format, i, totalSeconds, &result)
			}
			continue
		}

		if uCh == 'M' {
			count := countRun(upper, i, 'M')
			// When [s] is the elapsed unit, bare m shows total minutes;
			// when [h] or [m] is the unit, m shows minutes-within-hour.
			mv := minutes
			if elapsedUnit == 'S' {
				mv = totalMinutes
			}
			if count >= 2 {
				result.WriteString(fmt.Sprintf("%02d", mv))
			} else {
				result.WriteString(strconv.Itoa(mv))
			}
			i += count
			continue
		}

		if uCh == 'S' {
			count := countRun(upper, i, 'S')
			// When [s] is the elapsed unit, bare s also shows total seconds.
			var sv int
			var fracSec float64
			if elapsedUnit == 'S' {
				sv = int(totalSeconds)
				fracSec = totalSeconds - float64(sv)
			} else {
				sv = int(seconds)
				fracSec = seconds - float64(sv)
			}
			if count >= 2 {
				result.WriteString(fmt.Sprintf("%02d", sv))
			} else {
				result.WriteString(strconv.Itoa(sv))
			}
			// Fractional seconds.
			if i+count < len(format) && format[i+count] == '.' {
				dotPos := i + count
				zeroCount := 0
				j := dotPos + 1
				for j < len(format) && format[j] == '0' {
					zeroCount++
					j++
				}
				if zeroCount > 0 {
					fracStr := fmt.Sprintf("%.*f", zeroCount, fracSec)
					if dotIdx := strings.Index(fracStr, "."); dotIdx >= 0 {
						result.WriteString(fracStr[dotIdx:])
					}
					i = j
					continue
				}
			}
			i += count
			continue
		}

		if uCh == 'H' {
			count := countRun(upper, i, 'H')
			if count >= 2 {
				result.WriteString(fmt.Sprintf("%02d", totalHours))
			} else {
				result.WriteString(strconv.Itoa(totalHours))
			}
			i += count
			continue
		}

		result.WriteByte(ch)
		i++
	}

	s := result.String()
	if negative {
		s = "-" + s
	}
	return s
}

// detectElapsedUnit scans the upper-case format for the first elapsed
// bracket code and returns 'H', 'M', or 'S'. Returns 0 if none found.
func detectElapsedUnit(upper string) byte {
	inQuote := false
	for i := 0; i < len(upper); i++ {
		if upper[i] == '"' {
			inQuote = !inQuote
			continue
		}
		if inQuote {
			continue
		}
		if upper[i] == '\\' && i+1 < len(upper) {
			i++
			continue
		}
		if upper[i] == '[' {
			end := strings.Index(upper[i:], "]")
			if end > 0 {
				code := upper[i+1 : i+end]
				switch code {
				case "H", "HH":
					return 'H'
				case "M", "MM":
					return 'M'
				case "S", "SS":
					return 'S'
				}
			}
		}
	}
	return 0
}

// writeElapsedFracSeconds checks for .000 after an elapsed [s]/[ss] bracket
// code and writes the fractional seconds. Returns the new position.
func writeElapsedFracSeconds(format string, i int, totalSeconds float64, result *strings.Builder) int {
	if i < len(format) && format[i] == '.' {
		zeroCount := 0
		j := i + 1
		for j < len(format) && format[j] == '0' {
			zeroCount++
			j++
		}
		if zeroCount > 0 {
			fracSec := totalSeconds - float64(int(totalSeconds))
			fracStr := fmt.Sprintf("%.*f", zeroCount, fracSec)
			if dotIdx := strings.Index(fracStr, "."); dotIdx >= 0 {
				result.WriteString(fracStr[dotIdx:])
			}
			return j
		}
	}
	return i
}

// ---------------------------------------------------------------------------
// Fraction formatting (# #/# etc.)
// ---------------------------------------------------------------------------

func isFractionFormat(format string) bool {
	stripped := stripLiterals(format)
	// A fraction format contains a '/' surrounded by digit placeholders.
	return strings.Contains(stripped, "/") && !isDateTimeFormat(format) && !isElapsedTimeFormat(format)
}

func formatFraction(n float64, format string) string {
	negative := n < 0
	if negative {
		n = -n
	}

	tokens := tokenizeNumberFormat(format)

	// Find the '/' literal token that separates numerator from denominator.
	slashTokIdx := -1
	for i, tok := range tokens {
		if tok.kind == tokLiteral && tok.value == "/" {
			slashTokIdx = i
			break
		}
	}
	if slashTokIdx < 0 {
		return fmt.Sprintf("%g", n)
	}

	isDigitTok := func(k numFmtTokenKind) bool {
		return k == tokDigit || k == tokDigitOpt || k == tokDigitSpace
	}

	// Identify digit groups before the slash (separated by non-digit tokens).
	type digitGroup struct {
		start, end int // token indices [start, end)
		count      int
	}
	var beforeGroups []digitGroup
	inGroup := false
	var cur digitGroup
	for i := 0; i < slashTokIdx; i++ {
		if isDigitTok(tokens[i].kind) {
			if !inGroup {
				cur = digitGroup{start: i}
				inGroup = true
			}
			cur.count++
			cur.end = i + 1
		} else if inGroup {
			beforeGroups = append(beforeGroups, cur)
			inGroup = false
		}
	}
	if inGroup {
		beforeGroups = append(beforeGroups, cur)
	}

	hasWhole := len(beforeGroups) >= 2

	// Count denominator digit placeholders and detect fixed denominator.
	denomDigits := 0
	var denomFixedParts []byte
	for i := slashTokIdx + 1; i < len(tokens); i++ {
		tok := tokens[i]
		if isDigitTok(tok.kind) {
			denomDigits++
			if tok.kind == tokDigit && tok.value[0] >= '1' && tok.value[0] <= '9' {
				denomFixedParts = append(denomFixedParts, tok.value[0])
			} else {
				denomFixedParts = nil // not a fixed denominator
			}
		}
	}
	if denomDigits == 0 {
		denomDigits = 1
	}

	denom := 0
	fixedDenom := false
	if len(denomFixedParts) > 0 {
		if d, err := strconv.Atoi(string(denomFixedParts)); err == nil && d > 0 {
			denom = d
			fixedDenom = true
		}
	}
	// Fallback: try from stripped format.
	if !fixedDenom {
		stripped := stripLiterals(format)
		if si := strings.Index(stripped, "/"); si >= 0 {
			ds := strings.TrimSpace(stripped[si+1:])
			if d, err := strconv.Atoi(strings.TrimRight(ds, " #0?")); err == nil && d > 0 {
				denom = d
				fixedDenom = true
			}
		}
	}

	// Compute whole and fractional parts.
	var wholePart int
	var fracPart float64
	if hasWhole {
		wholePart = int(n)
		fracPart = n - float64(wholePart)
	} else {
		fracPart = n
	}

	// Find best fraction.
	var bestNum, bestDen int
	if fixedDenom {
		bestDen = denom
		bestNum = int(math.Round(fracPart * float64(denom)))
	} else {
		maxDen := 1
		for i := 0; i < denomDigits; i++ {
			maxDen *= 10
		}
		maxDen--
		if maxDen < 1 {
			maxDen = 9
		}
		bestNum, bestDen = bestFraction(fracPart, maxDen)
	}

	// Handle carry-over.
	if bestDen > 0 && bestNum >= bestDen {
		wholePart += bestNum / bestDen
		bestNum = bestNum % bestDen
	}

	// Classify token regions:
	//   prefix | whole-digits | middle | numerator-digits | pre-slash | / | post-slash | denom-digits | suffix
	// If !hasWhole, whole-digits and middle are absent.

	// Find region boundaries.
	prefixEnd := 0 // tokens[0..prefixEnd) are prefix literals
	if len(beforeGroups) > 0 {
		prefixEnd = beforeGroups[0].start
	} else {
		prefixEnd = slashTokIdx
	}

	var wholeEnd int               // whole digit group end
	var middleStart, middleEnd int // literals between whole and numerator
	var numStart, numEnd int       // numerator digit group
	if hasWhole {
		_ = beforeGroups[0].start
		wholeEnd = beforeGroups[0].end
		middleStart = wholeEnd
		numStart = beforeGroups[len(beforeGroups)-1].start
		middleEnd = numStart
		numEnd = beforeGroups[len(beforeGroups)-1].end
	} else if len(beforeGroups) > 0 {
		numStart = beforeGroups[0].start
		numEnd = beforeGroups[0].end
	}

	preSlashStart := numEnd // literals between numerator and /
	// postSlash: between / and first denom digit
	postSlashEnd := len(tokens) // suffix starts after last denom digit
	denomStart := -1
	denomEnd := -1
	for i := slashTokIdx + 1; i < len(tokens); i++ {
		if isDigitTok(tokens[i].kind) {
			if denomStart < 0 {
				denomStart = i
			}
			denomEnd = i + 1
		}
	}
	if denomEnd > 0 {
		postSlashEnd = denomStart
	}

	// Find suffix start (after last denom digit).
	suffixStart := len(tokens)
	if denomEnd > 0 {
		suffixStart = denomEnd
	}

	// Helper to collect literal values from a token range.
	collectLiterals := func(from, to int) string {
		var b strings.Builder
		for i := from; i < to; i++ {
			if tokens[i].kind == tokLiteral {
				b.WriteString(tokens[i].value)
			}
		}
		return b.String()
	}

	prefix := collectLiterals(0, prefixEnd)
	middle := collectLiterals(middleStart, middleEnd)
	preSlash := collectLiterals(preSlashStart, slashTokIdx)
	postSlash := collectLiterals(slashTokIdx+1, postSlashEnd)
	suffix := collectLiterals(suffixStart, len(tokens))

	// Count ? placeholders in numerator and denominator for space padding.
	numPlaces := 0
	for i := numStart; i < numEnd; i++ {
		if isDigitTok(tokens[i].kind) {
			numPlaces++
		}
	}
	denPlaces := denomDigits

	padFrac := func(s string, places int, hasQMark bool) string {
		if !hasQMark || len(s) >= places {
			return s
		}
		return strings.Repeat(" ", places-len(s)) + s
	}

	// Check if numerator or denominator have ? placeholders.
	numHasQ := false
	for i := numStart; i < numEnd; i++ {
		if tokens[i].kind == tokDigitSpace {
			numHasQ = true
			break
		}
	}
	denHasQ := false
	for i := slashTokIdx + 1; i < len(tokens); i++ {
		if tokens[i].kind == tokDigitSpace {
			denHasQ = true
			break
		}
	}

	// Build output.
	var result strings.Builder
	if negative {
		result.WriteByte('-')
	}

	if hasWhole && bestNum == 0 {
		// Fraction is zero — show just the whole number with surrounding literals.
		// If the format has '?' placeholders, the fraction section becomes spaces
		// to maintain consistent width. With '#' placeholders, it's suppressed.
		result.WriteString(prefix)
		result.WriteString(strconv.Itoa(wholePart))
		if numHasQ || denHasQ {
			// Replace middle + numerator + pre-slash + slash + post-slash + denominator with spaces.
			fracWidth := len(middle) + numPlaces + len(preSlash) + 1 + len(postSlash) + denPlaces
			result.WriteString(strings.Repeat(" ", fracWidth))
		}
		result.WriteString(suffix)
	} else if hasWhole {
		result.WriteString(prefix)
		if wholePart != 0 {
			result.WriteString(strconv.Itoa(wholePart))
			result.WriteString(middle)
		}
		result.WriteString(padFrac(strconv.Itoa(bestNum), numPlaces, numHasQ))
		result.WriteString(preSlash)
		result.WriteByte('/')
		result.WriteString(postSlash)
		result.WriteString(padFrac(strconv.Itoa(bestDen), denPlaces, denHasQ))
		result.WriteString(suffix)
	} else {
		totalNum := wholePart*bestDen + bestNum
		result.WriteString(prefix)
		result.WriteString(padFrac(strconv.Itoa(totalNum), numPlaces, numHasQ))
		result.WriteString(preSlash)
		result.WriteByte('/')
		result.WriteString(postSlash)
		result.WriteString(padFrac(strconv.Itoa(bestDen), denPlaces, denHasQ))
		result.WriteString(suffix)
	}
	return result.String()
}

// bestFraction finds the best rational approximation p/q for x with q <= maxDen.
// Uses the Stern-Brocot / mediant method.
func bestFraction(x float64, maxDen int) (int, int) {
	if x < 0 {
		x = -x
	}
	if x == 0 {
		return 0, 1
	}

	bestP, bestQ := 0, 1
	bestErr := x

	// Simple brute-force for small denominators (fast enough for typical use).
	for q := 1; q <= maxDen; q++ {
		p := int(math.Round(x * float64(q)))
		err := math.Abs(x - float64(p)/float64(q))
		if err < bestErr {
			bestP = p
			bestQ = q
			bestErr = err
			if err == 0 {
				break
			}
		}
	}
	return bestP, bestQ
}

// ---------------------------------------------------------------------------
// Number formatting (0, #, commas, %, E+, currency, literals)
// ---------------------------------------------------------------------------

func formatNumberSection(n float64, format string) string {
	// Parse the format into: prefix literals, number format, suffix literals.
	// Also detect percentage, scientific, and comma grouping.

	tokens := tokenizeNumberFormat(format)

	// Determine format properties.
	percentCount := 0
	hasScientific := false
	hasCommaGrouping := false
	sciIdx := -1

	for i, tok := range tokens {
		switch tok.kind {
		case tokPercent:
			percentCount++
		case tokExponent:
			hasScientific = true
			sciIdx = i
		case tokComma:
			// Comma adjacent to digit placeholders = grouping.
			hasCommaGrouping = true
		}
	}

	// Apply percentage: each % multiplies by 100.
	for pc := 0; pc < percentCount; pc++ {
		n *= 100
	}
	// Snap to 15 significant digits after percentage scaling to eliminate
	// floating-point noise (e.g. 0.00035*100 = 0.034999… instead of 0.035).
	if percentCount > 0 && n != 0 {
		s := strconv.FormatFloat(n, 'g', 15, 64)
		n, _ = strconv.ParseFloat(s, 64)
	}

	// Determine decimal places from digit tokens.
	if hasScientific {
		return formatScientific(n, tokens, sciIdx)
	}

	// Find decimal point position in tokens.
	decIdx := -1
	for i, tok := range tokens {
		if tok.kind == tokDecimal {
			decIdx = i
			break
		}
	}

	// Count integer and decimal digit placeholders.
	intZeros := 0   // minimum integer digits (from '0')
	intDigits := 0  // total integer placeholders (from '0', '#', '?')
	intSpaces := 0  // integer '?' count (space-padded positions)
	decZeros := 0   // decimal '0' count
	decHashes := 0  // decimal '#' count
	decSpaces := 0  // decimal '?' count

	inDecimal := false
	for _, tok := range tokens {
		if tok.kind == tokDecimal {
			inDecimal = true
			continue
		}
		if tok.kind == tokDigit || tok.kind == tokDigitOpt || tok.kind == tokDigitSpace {
			if inDecimal {
				if tok.kind == tokDigit {
					decZeros++
				} else if tok.kind == tokDigitSpace {
					decSpaces++
				} else {
					decHashes++
				}
			} else {
				intDigits++
				if tok.kind == tokDigit {
					intZeros++
				} else if tok.kind == tokDigitSpace {
					intSpaces++
				}
			}
		}
	}
	_ = decSpaces

	totalDecPlaces := decZeros + decHashes
	_ = decIdx

	// Check for trailing commas (scaling): commas at end of digit sequence divide by 1000.
	trailingCommas := countTrailingCommas(tokens)
	for tc := 0; tc < trailingCommas; tc++ {
		n /= 1000
	}

	// Format the number.
	negative := n < 0
	if negative {
		n = -n
	}

	// Round to the number of decimal places.
	rounded := roundToPlaces(n, totalDecPlaces)

	// Split into integer and decimal parts.
	intPart, decPart := splitNumber(rounded, totalDecPlaces)

	// Format integer part with zero-padding.
	intStr := intPart
	if len(intStr) < intZeros {
		intStr = strings.Repeat("0", intZeros-len(intStr)) + intStr
	}
	if intStr == "" {
		intStr = "0"
	}

	// Apply comma grouping.
	if hasCommaGrouping && trailingCommas == 0 {
		intStr = addCommaGrouping(intStr)
	}

	// Format decimal part.
	decStr := ""
	if totalDecPlaces > 0 {
		decStr = decPart
		// Pad with zeros to meet minimum.
		for len(decStr) < totalDecPlaces {
			decStr += "0"
		}
		// Trim trailing zeros for '#' positions.
		if decHashes > 0 {
			minLen := decZeros
			for len(decStr) > minLen && decStr[len(decStr)-1] == '0' {
				decStr = decStr[:len(decStr)-1]
			}
		}
	}

	// Build the result using the token stream to preserve literals.
	var result strings.Builder
	if negative {
		result.WriteByte('-')
	}

	intWritten := false
	decWritten := false

	// If the format has '?' integer placeholders, pad the integer with leading
	// spaces so that the total digit count matches the placeholder count.
	if intSpaces > 0 {
		rawDigits := strings.ReplaceAll(intStr, ",", "")
		if len(rawDigits) < intDigits {
			padded := strings.Repeat("0", intDigits-len(rawDigits)) + rawDigits
			if hasCommaGrouping && trailingCommas == 0 {
				padded = addCommaGrouping(padded)
			}
			// Replace leading zeros (and their adjacent commas) with spaces.
			buf := []byte(padded)
			for i := 0; i < len(buf); i++ {
				if buf[i] == '0' {
					buf[i] = ' '
				} else if buf[i] == ',' {
					buf[i] = ' '
				} else {
					break
				}
			}
			intStr = string(buf)
		}
	}

	for _, tok := range tokens {
		switch tok.kind {
		case tokLiteral:
			result.WriteString(tok.value)
		case tokDigit, tokDigitOpt, tokDigitSpace:
			if !intWritten {
				result.WriteString(intStr)
				intWritten = true
			}
			// Skip subsequent digit tokens as intStr was already written.
		case tokDecimal:
			if totalDecPlaces > 0 && decStr != "" {
				result.WriteByte('.')
				result.WriteString(decStr)
			} else if decZeros > 0 {
				result.WriteByte('.')
				result.WriteString(decStr)
			}
			decWritten = true
			_ = decWritten
		case tokComma:
			// Already handled via comma grouping; skip.
		case tokPercent:
			result.WriteByte('%')
		case tokExponent:
			// Handled in formatScientific.
		}
	}

	// If no digit tokens existed and there are actual digit placeholders in
	// the format, write the number. Skip if the format is all literals (e.g.
	// the "zero" section of a multi-section format).
	if !intWritten && intDigits > 0 {
		result.WriteString(intStr)
		if totalDecPlaces > 0 {
			result.WriteByte('.')
			result.WriteString(decStr)
		}
	}

	return result.String()
}

// Token types for number format parsing.
type numFmtTokenKind byte

const (
	tokLiteral    numFmtTokenKind = iota
	tokDigit                       // '0' — required digit
	tokDigitOpt                    // '#' — optional digit
	tokDigitSpace                  // '?' — digit padded with space
	tokDecimal                  // '.'
	tokComma                    // ','
	tokPercent                  // '%'
	tokExponent                 // 'E+' or 'E-'
)

type numFmtToken struct {
	kind  numFmtTokenKind
	value string
}

// tokenizeNumberFormat breaks a number format string into tokens.
func tokenizeNumberFormat(format string) []numFmtToken {
	var tokens []numFmtToken
	upper := strings.ToUpper(format)
	i := 0

	for i < len(format) {
		ch := format[i]

		// Quoted string.
		if ch == '"' {
			var lit strings.Builder
			i++
			for i < len(format) && format[i] != '"' {
				lit.WriteByte(format[i])
				i++
			}
			if i < len(format) {
				i++ // skip closing quote
			}
			tokens = append(tokens, numFmtToken{kind: tokLiteral, value: lit.String()})
			continue
		}

		// Backslash escape.
		if ch == '\\' && i+1 < len(format) {
			i++
			tokens = append(tokens, numFmtToken{kind: tokLiteral, value: string(format[i])})
			i++
			continue
		}

		// Underscore (space placeholder in Excel) — skip next char, emit space.
		if ch == '_' && i+1 < len(format) {
			tokens = append(tokens, numFmtToken{kind: tokLiteral, value: " "})
			i += 2
			continue
		}

		// Asterisk (repeat fill char in Excel) — skip next char.
		if ch == '*' && i+1 < len(format) {
			i += 2
			continue
		}

		switch ch {
		case '0':
			tokens = append(tokens, numFmtToken{kind: tokDigit, value: "0"})
			i++
		case '#':
			tokens = append(tokens, numFmtToken{kind: tokDigitOpt, value: "#"})
			i++
		case '?':
			tokens = append(tokens, numFmtToken{kind: tokDigitSpace, value: "?"})
			i++
		case '.':
			tokens = append(tokens, numFmtToken{kind: tokDecimal, value: "."})
			i++
		case ',':
			tokens = append(tokens, numFmtToken{kind: tokComma, value: ","})
			i++
		case '%':
			tokens = append(tokens, numFmtToken{kind: tokPercent, value: "%"})
			i++
		case 'E', 'e':
			// Scientific notation: E+ or E-.
			if i+1 < len(format) && (format[i+1] == '+' || format[i+1] == '-') {
				tokens = append(tokens, numFmtToken{kind: tokExponent, value: format[i : i+2]})
				i += 2
				// Emit the exponent digit placeholders as tokens so formatScientific can count them.
				for i < len(format) && (format[i] == '0' || format[i] == '#') {
					if format[i] == '0' {
						tokens = append(tokens, numFmtToken{kind: tokDigit, value: "0"})
					} else {
						tokens = append(tokens, numFmtToken{kind: tokDigitOpt, value: "#"})
					}
					i++
				}
			} else {
				tokens = append(tokens, numFmtToken{kind: tokLiteral, value: string(ch)})
				i++
			}
		default:
			// Check for common literal characters.
			if isFormatLiteral(ch, upper, i) {
				tokens = append(tokens, numFmtToken{kind: tokLiteral, value: string(ch)})
				i++
			} else {
				// Unknown char — treat as literal.
				tokens = append(tokens, numFmtToken{kind: tokLiteral, value: string(ch)})
				i++
			}
		}
	}

	return tokens
}

// isFormatLiteral determines if a character should be treated as a literal in a number format.
func isFormatLiteral(ch byte, upper string, i int) bool {
	switch ch {
	case '$', '-', '+', '/', '(', ')', ':', '!', '^', '&', '\'', '~',
		'{', '}', '=', '<', '>', ' ', '@':
		return true
	}
	// Check for letters that aren't format codes.
	if ch >= 'a' && ch <= 'z' || ch >= 'A' && ch <= 'Z' {
		// In a pure number format, letters other than E are literals.
		return true
	}
	return false
}

// countTrailingCommas counts commas at the end of the digit sequence (scaling commas).
func countTrailingCommas(tokens []numFmtToken) int {
	// Find the last digit/decimal token, then count consecutive commas after it.
	lastDigitIdx := -1
	for i, tok := range tokens {
		if tok.kind == tokDigit || tok.kind == tokDigitOpt || tok.kind == tokDigitSpace || tok.kind == tokDecimal {
			lastDigitIdx = i
		}
	}
	if lastDigitIdx < 0 {
		return 0
	}

	count := 0
	for i := lastDigitIdx + 1; i < len(tokens); i++ {
		if tokens[i].kind == tokComma {
			count++
		} else if tokens[i].kind == tokPercent || tokens[i].kind == tokLiteral {
			break
		} else {
			break
		}
	}
	return count
}

// roundToPlaces rounds n to the given number of decimal places.
func roundToPlaces(n float64, places int) float64 {
	if places <= 0 {
		return math.Round(n)
	}
	factor := math.Pow(10, float64(places))
	return math.Round(n*factor) / factor
}

// splitNumber splits a non-negative number into integer and decimal string parts.
func splitNumber(n float64, decPlaces int) (string, string) {
	s := fmt.Sprintf("%.*f", decPlaces, n)
	if decPlaces == 0 {
		return s, ""
	}
	parts := strings.SplitN(s, ".", 2)
	if len(parts) == 2 {
		return parts[0], parts[1]
	}
	return parts[0], ""
}

// addCommaGrouping inserts commas into an integer string.
func addCommaGrouping(s string) string {
	if len(s) <= 3 {
		return s
	}
	var b strings.Builder
	start := len(s) % 3
	if start > 0 {
		b.WriteString(s[:start])
	}
	for i := start; i < len(s); i += 3 {
		if i > 0 {
			b.WriteByte(',')
		}
		b.WriteString(s[i : i+3])
	}
	return b.String()
}

// formatScientific formats a number in scientific notation based on the format tokens.
func formatScientific(n float64, tokens []numFmtToken, sciIdx int) string {
	// Count mantissa decimal places.
	decPlaces := 0
	inDecimal := false
	for i, tok := range tokens {
		if i >= sciIdx {
			break
		}
		if tok.kind == tokDecimal {
			inDecimal = true
			continue
		}
		if inDecimal && (tok.kind == tokDigit || tok.kind == tokDigitOpt || tok.kind == tokDigitSpace) {
			decPlaces++
		}
	}

	// Count exponent digit placeholders (after the E+/E- token).
	expDigits := 0
	for i := sciIdx + 1; i < len(tokens); i++ {
		if tokens[i].kind == tokDigit || tokens[i].kind == tokDigitOpt || tokens[i].kind == tokDigitSpace {
			expDigits++
		} else {
			break
		}
	}
	if expDigits == 0 {
		expDigits = 1
	}

	// Get the sign of E.
	expSign := "+"
	if sciIdx < len(tokens) && len(tokens[sciIdx].value) >= 2 {
		expSign = string(tokens[sciIdx].value[1])
	}

	negative := n < 0
	if negative {
		n = -n
	}

	// Calculate exponent.
	exp := 0
	mantissa := n
	if mantissa != 0 {
		exp = int(math.Floor(math.Log10(mantissa)))
		mantissa = mantissa / math.Pow(10, float64(exp))
	}

	// Round mantissa.
	mantissa = roundToPlaces(mantissa, decPlaces)

	// Handle rounding that pushes mantissa to 10.
	if mantissa >= 10 {
		mantissa /= 10
		exp++
	}

	var result strings.Builder
	if negative {
		result.WriteByte('-')
	}

	// Build prefix literals.
	for i := 0; i < sciIdx; i++ {
		tok := tokens[i]
		if tok.kind == tokLiteral {
			result.WriteString(tok.value)
		}
	}

	// Format mantissa.
	mStr := fmt.Sprintf("%.*f", decPlaces, mantissa)
	result.WriteString(mStr)

	// Format exponent.
	result.WriteByte('E')
	if exp >= 0 {
		result.WriteString(expSign)
	} else {
		result.WriteByte('-')
		exp = -exp
	}
	expStr := strconv.Itoa(exp)
	for len(expStr) < expDigits {
		expStr = "0" + expStr
	}
	result.WriteString(expStr)

	// Suffix literals.
	pastExpDigits := false
	for i := sciIdx + 1; i < len(tokens); i++ {
		tok := tokens[i]
		if !pastExpDigits && (tok.kind == tokDigit || tok.kind == tokDigitOpt || tok.kind == tokDigitSpace) {
			continue // skip exponent digit placeholders
		}
		pastExpDigits = true
		if tok.kind == tokLiteral {
			result.WriteString(tok.value)
		} else if tok.kind == tokPercent {
			result.WriteByte('%')
		}
	}

	return result.String()
}

// ---------------------------------------------------------------------------
// Time helpers
// ---------------------------------------------------------------------------

// daysSinceEpoch converts a time.Time to Excel days (for internal elapsed time).
func daysSinceEpoch(t time.Time) float64 {
	duration := t.Sub(ExcelEpoch)
	return duration.Hours() / 24
}
