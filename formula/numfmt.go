package formula

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"
)

// ---------------------------------------------------------------------------
// Number format engine
//
// Supports the format codes used by TEXT(), cell number formats, etc.
// Reference: https://support.microsoft.com/en-us/office/number-format-codes
//
// Features implemented:
//   - Section separator ; (positive;negative;zero;text)
//   - Literal text: "quoted", \escaped, and passthrough of $ - + / ( ) : ! ^ & ' ~ { } = < > space
//   - Date/time codes: d dd ddd dddd m mm mmm mmmm mmmmm yy yyyy h hh m mm s ss AM/PM
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

// formatNumber formats a number using a number format string.
// This is the main entry point used by the TEXT() function.
// If date1904 is true, date serial numbers use the 1904 date system.
func formatNumber(n float64, format string, date1904 bool) string {
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

	// Strip color codes like [Red], [Blue], [Green], etc. — they are display
	// hints only and do not affect the formatted text output.
	section = stripColorCodes(section)

	// Check for elapsed time format [h], [m], [s] — must be checked before
	// general date/time because [h]:mm:ss also contains h/m/s tokens.
	if isElapsedTimeFormat(section) {
		return formatElapsedTime(n, section)
	}

	// Check if this is a date/time format.
	if isDateTimeFormat(section) {
		return formatDateTime(n, section, date1904)
	}

	// Check for fraction format.
	if isFractionFormat(section) {
		return formatFraction(n, section)
	}

	// Number format.
	return formatNumberSection(n, section)
}

// anyNumberFormatCodes returns true if at least one section contains a number,
// date/time, or fraction format code. Sections that are purely literal text
// (like "pos", "neg", "zero") do not count.
func anyNumberFormatCodes(sections []string) bool {
	for _, sec := range sections {
		if sectionHasFormatCodes(sec) {
			return true
		}
	}
	return false
}

// sectionHasFormatCodes returns true if the section contains number format
// codes (0, #, ?, %, E+/E-), the @ placeholder, elapsed time codes ([h], [m], [s]),
// or date/time tokens (y, d, h, m, s), outside of quoted strings, escape
// sequences, and bracketed color codes.
func sectionHasFormatCodes(section string) bool {
	stripped := stripColorCodes(section)
	// Check for elapsed time [h], [m], [s].
	if isElapsedTimeFormat(stripped) {
		return true
	}
	// Check for fraction format.
	if isFractionFormat(stripped) {
		return true
	}
	// Scan for unquoted, unescaped format codes.
	inQuote := false
	for i := 0; i < len(stripped); i++ {
		ch := stripped[i]
		if ch == '"' {
			inQuote = !inQuote
			continue
		}
		if inQuote {
			continue
		}
		if ch == '\\' && i+1 < len(stripped) {
			i++ // skip escaped char
			continue
		}
		// Skip bracketed expressions.
		if ch == '[' {
			for i < len(stripped) && stripped[i] != ']' {
				i++
			}
			continue
		}
		upper := ch
		if upper >= 'a' && upper <= 'z' {
			upper -= 32
		}
		switch upper {
		case '0', '#', '?', '%', '@':
			return true
		case 'E':
			if i+1 < len(stripped) && (stripped[i+1] == '+' || stripped[i+1] == '-') {
				return true
			}
		// Date/time codes — Y, D, H are unambiguous; M could be month or minute
		// but is always a date/time code in number formats.
		case 'Y', 'D', 'H', 'M':
			return true
		}
	}
	return false
}

// sectionHasNumberCodes returns true if the section contains number digit
// placeholders (0, #, ?) outside of quoted strings and escape sequences.
func sectionHasNumberCodes(section string) bool {
	inQuote := false
	for i := 0; i < len(section); i++ {
		ch := section[i]
		if ch == '"' {
			inQuote = !inQuote
			continue
		}
		if inQuote {
			continue
		}
		if ch == '\\' && i+1 < len(section) {
			i++
			continue
		}
		if ch == '[' {
			for i < len(section) && section[i] != ']' {
				i++
			}
			continue
		}
		if ch == '0' || ch == '#' || ch == '?' {
			return true
		}
	}
	return false
}

// sectionHasUnquotedLetters returns true if the section contains unquoted,
// unescaped alphabetic characters (a-z, A-Z) that are not part of recognised
// format sequences like E+/E-.  This is used to detect invalid number format
// strings such as "Value: 0" where alphabetic text should be quoted.
func sectionHasUnquotedLetters(section string) bool {
	inQuote := false
	for i := 0; i < len(section); i++ {
		ch := section[i]
		if ch == '"' {
			inQuote = !inQuote
			continue
		}
		if inQuote {
			continue
		}
		if ch == '\\' && i+1 < len(section) {
			i++ // skip escaped char
			continue
		}
		if ch == '[' {
			for i < len(section) && section[i] != ']' {
				i++
			}
			continue
		}
		if ch == '_' && i+1 < len(section) {
			i++ // skip underscore + next char
			continue
		}
		if ch == '*' && i+1 < len(section) {
			i++ // skip asterisk + next char
			continue
		}
		// E followed by + or - is scientific notation, not a literal letter.
		if (ch == 'E' || ch == 'e') && i+1 < len(section) && (section[i+1] == '+' || section[i+1] == '-') {
			i++ // skip the +/-
			continue
		}
		if (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') {
			return true
		}
	}
	return false
}

// colorCodes is the set of color names recognised inside square brackets.
var colorCodes = map[string]bool{
	"BLACK":   true,
	"BLUE":    true,
	"CYAN":    true,
	"GREEN":   true,
	"MAGENTA": true,
	"RED":     true,
	"WHITE":   true,
	"YELLOW":  true,
}

// stripColorCodes removes color codes like [Red], [Blue], [Color 3], etc.
// from a format section. These are display hints and do not affect text output.
func stripColorCodes(format string) string {
	var result strings.Builder
	i := 0
	for i < len(format) {
		if format[i] == '[' {
			// Find the closing bracket.
			j := i + 1
			for j < len(format) && format[j] != ']' {
				j++
			}
			if j < len(format) {
				inner := format[i+1 : j]
				upperInner := strings.ToUpper(strings.TrimSpace(inner))
				// Check for named colors or "Color N" syntax.
				if colorCodes[upperInner] || strings.HasPrefix(upperInner, "COLOR ") || strings.HasPrefix(upperInner, "COLOR") && len(upperInner) > 5 && upperInner[5] >= '0' && upperInner[5] <= '9' {
					i = j + 1 // skip past ']'
					continue
				}
			}
			result.WriteByte(format[i])
			i++
		} else {
			result.WriteByte(format[i])
			i++
		}
	}
	return result.String()
}

// formatTextSection formats a text value using the text section (4th section)
// of a format string. The '@' placeholder is replaced with the text value.
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

// sectionContainsAt returns true if the format section contains an unquoted,
// unescaped '@' placeholder.
func sectionContainsAt(section string) bool {
	for i := 0; i < len(section); i++ {
		ch := section[i]
		if ch == '"' {
			i++
			for i < len(section) && section[i] != '"' {
				i++
			}
		} else if ch == '\\' && i+1 < len(section) {
			i++
		} else if ch == '@' {
			return true
		}
	}
	return false
}

// formatGeneral returns the "General" format representation.
// General format uses scientific notation (uppercase E) for:
//   - numbers with absolute value >= 1e11
//   - very small nonzero numbers with absolute value < 1e-4
//
// It displays at most ~11 significant digits.
func formatGeneral(n float64) string {
	abs := math.Abs(n)
	// Very large or very small numbers use scientific notation.
	// Excel General format uses decimal for numbers that fit within ~11
	// characters, otherwise scientific. Values >= 1e-9 are shown as decimal
	// (e.g. 0.000000001 = 11 chars). Below 1e-9, scientific notation is used.
	if abs >= 1e11 || (abs > 0 && abs < 1e-9) {
		// Excel General format uses 5 digits after the decimal point
		// (6 significant digits total) for scientific notation.
		s := strconv.FormatFloat(n, 'E', 5, 64)
		// Trim trailing zeros after the decimal point in the coefficient.
		if idx := strings.IndexByte(s, 'E'); idx >= 0 {
			coeff := s[:idx]
			expPart := s[idx:]
			if dotIdx := strings.IndexByte(coeff, '.'); dotIdx >= 0 {
				coeff = strings.TrimRight(coeff, "0")
				coeff = strings.TrimRight(coeff, ".")
			}
			// Strip leading zeros from the exponent and ensure sign is present.
			sign := expPart[1] // '+' or '-'
			exp := strings.TrimLeft(expPart[2:], "0")
			if exp == "" {
				exp = "0"
			}
			s = coeff + "E" + string(sign) + exp
		}
		return s
	}
	if n == math.Trunc(n) && abs < 1e15 {
		return strconv.FormatFloat(n, 'f', -1, 64)
	}
	// Use %g-like formatting but with up to 10 significant digits.
	s := strconv.FormatFloat(n, 'G', 10, 64)
	// For small numbers (1e-9 to 1e-4) Go's G format may produce scientific
	// notation, but Excel General format uses decimal notation for these.
	// Convert to decimal if that happened.
	if abs < 1e-4 && abs >= 1e-9 && (strings.ContainsRune(s, 'E') || strings.ContainsRune(s, 'e')) {
		s = strconv.FormatFloat(n, 'f', -1, 64)
		// Trim to 10 significant digits.
		if len(s) > 0 {
			neg := s[0] == '-'
			digits := s
			if neg {
				digits = s[1:]
			}
			// Count significant digits (skip leading "0." and zeros).
			sigCount := 0
			pastLeading := false
			cutIdx := -1
			for i, ch := range digits {
				if ch == '.' {
					continue
				}
				if ch != '0' {
					pastLeading = true
				}
				if pastLeading {
					sigCount++
					if sigCount > 10 {
						cutIdx = i
						break
					}
				}
			}
			if cutIdx >= 0 {
				digits = digits[:cutIdx]
				digits = strings.TrimRight(digits, "0")
				if neg {
					s = "-" + digits
				} else {
					s = digits
				}
			}
		}
	}
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
			// A standalone "m" format is a date format (month of a date).
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

// formatDateTime formats a date serial number as a date/time string.
func formatDateTime(serial float64, format string, date1904 bool) string {
	// Determine if there's an AM/PM marker.
	stripped := stripLiterals(format)
	upperStripped := strings.ToUpper(stripped)
	hasAMPM := strings.Contains(upperStripped, "AM/PM") || strings.Contains(upperStripped, "A/P")

	// Compute time components from the fractional day part. We round the
	// total seconds to the displayed precision (whole seconds when no
	// sub-second display, or N decimal places for ss.000-style formats)
	// to avoid truncation artifacts like 52526.999... showing as :26.
	// We compute h:m:s directly instead of via time.Time because the
	// float64 round-trip loses precision at high serial magnitudes.
	fracSecDigits := countFractionalSecondDigits(upperStripped)
	frac := serial - math.Floor(serial)
	if frac < 0 {
		frac += 1
	}
	totalSeconds := frac * 86400
	// Round to the display precision.
	var roundedTotal float64
	if fracSecDigits == 0 {
		roundedTotal = math.Round(totalSeconds)
	} else {
		p := math.Pow(10, float64(fracSecDigits))
		roundedTotal = math.Round(totalSeconds*p) / p
	}
	wholeSec := int(roundedTotal)
	fracSeconds := roundedTotal - float64(wholeSec)
	if wholeSec >= 86400 {
		wholeSec = 0
		fracSeconds = 0
	}
	hour := wholeSec / 3600
	minute := (wholeSec % 3600) / 60
	second := wholeSec % 60

	// We still need time.Time for weekday. Year/month/day come from
	// serialDatePartsForDateSystem so the 1900-system serial 60 renders as
	// "1900-02-29" (Excel's intentional leap-year bug) rather than
	// "1900-03-01".
	var t time.Time
	if date1904 {
		t = SerialToTime1904(serial)
	} else {
		t = SerialToTime(serial)
	}
	yearDP, monthDP, dayDP := serialDatePartsForDateSystem(serial, date1904)

	hour12 := hour % 12
	if hour12 == 0 {
		hour12 = 12
	}
	// Determine the case of AM/PM from the original format string.
	ampmLower := isAMPMLowercase(format)
	ampm := "AM"
	if hour >= 12 {
		ampm = "PM"
	}
	if ampmLower {
		ampm = strings.ToLower(ampm)
	}
	ap := "A"
	if hour >= 12 {
		ap = "P"
	}
	if ampmLower {
		ap = strings.ToLower(ap)
	}

	day := dayDP
	month := int(monthDP)
	year := yearDP
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
				case 4:
					if month >= 1 && month <= 12 {
						result.WriteString(longMonths[month])
					}
				default: // 5+ ("MMMMM" → first letter of month name)
					if month >= 1 && month <= 12 {
						result.WriteByte(longMonths[month][0])
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

// countFractionalSecondDigits returns how many fractional-second digits the
// format displays (e.g. "SS.000" returns 3, "SS.00" returns 2). Returns 0
// if the format has no sub-second display. Expects upper-cased, literal-stripped input.
func countFractionalSecondDigits(upper string) int {
	for i := 0; i < len(upper); i++ {
		if upper[i] == 'S' {
			// Skip the run of S characters.
			j := i
			for j < len(upper) && upper[j] == 'S' {
				j++
			}
			// Check for '.' followed by '0's.
			if j < len(upper) && upper[j] == '.' {
				k := j + 1
				for k < len(upper) && upper[k] == '0' {
					k++
				}
				if k > j+1 {
					return k - j - 1
				}
			}
			i = j - 1
		}
	}
	return 0
}

// computeMContexts pre-scans the format to determine for each 'm'/'M' run
// whether it represents minutes (true) or months (false).
// 'm' after 'h' and before 's' means minutes; otherwise months.
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
				// in practice m is still treated as minute if h appeared before.
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

// isAMPMLowercase scans the raw format string (skipping quoted/escaped regions)
// for an AM/PM or A/P marker and returns true when that marker is lowercase.
// Excel rules: AM/PM (any case) always produces uppercase output.
// Only the short form a/p (lowercase) produces lowercase output; A/P is uppercase.
func isAMPMLowercase(format string) bool {
	inQuote := false
	upper := strings.ToUpper(format)
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
			i++
			continue
		}
		uCh := upper[i]
		if uCh == 'A' {
			if i+4 < len(upper) && upper[i:i+5] == "AM/PM" {
				// AM/PM always produces uppercase regardless of case in format.
				return false
			}
			if i+2 < len(upper) && upper[i:i+3] == "A/P" {
				return format[i] == 'a'
			}
		}
	}
	return false
}

// isLiteralPassthrough returns true for characters that are passed through as-is
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
			case "H":
				result.WriteString(strconv.Itoa(totalHours))
			case "HH":
				result.WriteString(fmt.Sprintf("%02d", totalHours))
			case "M":
				result.WriteString(strconv.Itoa(totalMinutes))
			case "MM":
				result.WriteString(fmt.Sprintf("%02d", totalMinutes))
			case "S":
				result.WriteString(strconv.Itoa(int(totalSeconds)))
				// Handle fractional seconds after [s].
				i = writeElapsedFracSeconds(format, i, totalSeconds, &result)
			case "SS":
				result.WriteString(fmt.Sprintf("%02d", int(totalSeconds)))
				// Handle fractional seconds after [ss].
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

	var wholeStart, wholeEnd int   // whole digit section (all groups except last)
	var middleStart, middleEnd int // literals between whole and numerator
	var numStart, numEnd int       // numerator digit group
	if hasWhole {
		wholeStart = beforeGroups[0].start
		lastWholeIdx := len(beforeGroups) - 2
		wholeEnd = beforeGroups[lastWholeIdx].end
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

	// Count whole-part digit placeholders.
	wholePlaces := 0
	if hasWhole {
		for i := wholeStart; i < wholeEnd; i++ {
			if isDigitTok(tokens[i].kind) {
				wholePlaces++
			}
		}
	}

	// Count ? placeholders in numerator and denominator for space padding.
	numPlaces := 0
	for i := numStart; i < numEnd; i++ {
		if isDigitTok(tokens[i].kind) {
			numPlaces++
		}
	}
	denPlaces := denomDigits

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

	// formatWholeDigits writes the whole-part section digit-by-digit.
	// Each digit placeholder (#, 0, ?) is matched to a digit of wholePart
	// (right-aligned), and interleaved literals are always emitted.
	// When forceZero is true, the last digit position always shows '0' even
	// for # or ? placeholders (used when the entire value is exactly zero).
	formatWholeDigits := func(buf *strings.Builder, forceZero bool) {
		wholeStr := strconv.Itoa(wholePart)
		for len(wholeStr) < wholePlaces {
			wholeStr = "0" + wholeStr
		}
		// Right-align: if more digits than placeholders, leading digits go
		// into the first placeholder position.
		digitIdx := 0
		overflow := len(wholeStr) - wholePlaces
		for i := wholeStart; i < wholeEnd; i++ {
			tok := tokens[i]
			if isDigitTok(tok.kind) {
				isLast := digitIdx == wholePlaces-1
				if digitIdx == 0 && overflow > 0 {
					// First placeholder absorbs overflow digits.
					chunk := wholeStr[:overflow+1]
					allZero := true
					for _, c := range chunk {
						if c != '0' {
							allZero = false
							break
						}
					}
					switch tok.kind {
					case tokDigitOpt:
						if !allZero {
							buf.WriteString(chunk)
						}
					case tokDigitSpace:
						if allZero {
							buf.WriteString(strings.Repeat(" ", len(chunk)))
						} else {
							buf.WriteString(chunk)
						}
					default:
						buf.WriteString(chunk)
					}
					digitIdx += overflow + 1
				} else {
					ch := wholeStr[overflow+digitIdx]
					forceThis := forceZero && isLast && ch == '0'
					switch tok.kind {
					case tokDigitOpt:
						if ch != '0' || forceThis {
							buf.WriteByte(ch)
						}
					case tokDigitSpace:
						if ch == '0' && !forceThis {
							buf.WriteByte(' ')
						} else {
							buf.WriteByte(ch)
						}
					default:
						buf.WriteByte(ch)
					}
					digitIdx++
				}
			} else if tok.kind == tokLiteral {
				buf.WriteString(tok.value)
			}
		}
	}

	// Check if numerator has any '0' (mandatory) digit placeholders.
	numHasZero := false
	for i := numStart; i < numEnd; i++ {
		if tokens[i].kind == tokDigit {
			numHasZero = true
			break
		}
	}

	// formatFracDigits formats a number across the digit placeholders in
	// the token range [from, to), respecting each placeholder type.
	// When forceZero is true, the last digit shows '0' even for # or ? placeholders.
	// When leftAlign is true, the number is left-justified (trailing spaces for ?),
	// which is the correct behavior for fraction denominators.
	formatFracDigits := func(n, from, to int, forceZero, leftAlign bool) string {
		places := 0
		for i := from; i < to; i++ {
			if isDigitTok(tokens[i].kind) {
				places++
			}
		}
		s := strconv.Itoa(n)
		for len(s) < places {
			s = "0" + s
		}
		var buf strings.Builder
		digitIdx := 0
		overflow := len(s) - places
		for i := from; i < to; i++ {
			tok := tokens[i]
			if isDigitTok(tok.kind) {
				isLast := digitIdx == places-1
				if digitIdx == 0 && overflow > 0 {
					// First placeholder absorbs overflow digits.
					chunk := s[:overflow+1]
					allZero := true
					for _, c := range chunk {
						if c != '0' {
							allZero = false
							break
						}
					}
					switch tok.kind {
					case tokDigitOpt:
						if !allZero {
							buf.WriteString(chunk)
						}
					case tokDigitSpace:
						if allZero {
							buf.WriteString(strings.Repeat(" ", len(chunk)))
						} else {
							buf.WriteString(chunk)
						}
					default:
						buf.WriteString(chunk)
					}
					digitIdx += overflow + 1
				} else {
					ch := s[overflow+digitIdx]
					forceThis := forceZero && isLast && ch == '0'
					switch tok.kind {
					case tokDigitOpt: // #
						if ch != '0' || forceThis {
							buf.WriteByte(ch)
						}
					case tokDigitSpace: // ?
						if ch == '0' && !forceThis {
							buf.WriteByte(' ')
						} else {
							buf.WriteByte(ch)
						}
					default: // 0
						buf.WriteByte(ch)
					}
					digitIdx++
				}
			} else if tok.kind == tokLiteral {
				buf.WriteString(tok.value)
			}
		}
		result := buf.String()
		// For left-aligned mode (denominators), move leading spaces to the end
		// so that the number is left-justified within its field width.
		if leftAlign {
			trimmed := strings.TrimLeft(result, " ")
			leadingSpaces := len(result) - len(trimmed)
			if leadingSpaces > 0 {
				result = trimmed + strings.Repeat(" ", leadingSpaces)
			}
		}
		return result
	}

	// Detect the whole-part placeholder type (first digit token in whole range).
	// This controls behavior when wholePart is zero:
	//   tokDigitOpt (#):   suppress digit AND suppress middle literals entirely
	//   tokDigitSpace (?): show space for digit, replace each middle literal char with a space
	//   tokDigit (0):      show '0', show middle literals as-is
	wholeType := tokDigit // default to '0' behavior
	wholeHasMandatory := false
	if hasWhole {
		first := true
		for i := wholeStart; i < wholeEnd; i++ {
			if isDigitTok(tokens[i].kind) {
				if first {
					wholeType = tokens[i].kind
					first = false
				}
				if tokens[i].kind == tokDigit || tokens[i].kind == tokDigitSpace {
					wholeHasMandatory = true
				}
			}
		}
	}

	// Build output.
	var result strings.Builder
	if negative {
		result.WriteByte('-')
	}

	writeFrac := func(num int, forceNum bool) {
		result.WriteString(formatFracDigits(num, numStart, numEnd, forceNum, false))
		result.WriteString(preSlash)
		result.WriteByte('/')
		result.WriteString(postSlash)
		if denomEnd > 0 {
			result.WriteString(formatFracDigits(bestDen, denomStart, denomEnd, true, true))
		} else {
			result.WriteString(strconv.Itoa(bestDen))
		}
	}

	if hasWhole && bestNum == 0 {
		// Fraction is zero: show whole number, handle fraction area based on
		// placeholder types.
		result.WriteString(prefix)
		// For pure-# whole with wholePart=0, don't force zero on the whole
		// digit (the numerator's 0 placeholder will show the zero instead).
		// But if the whole section also has mandatory placeholders (0 or ?),
		// force zero so those placeholders display correctly.
		if wholePart == 0 && wholeType == tokDigitOpt && numHasZero && !wholeHasMandatory {
			formatWholeDigits(&result, false)
		} else {
			formatWholeDigits(&result, true)
		}
		if numHasZero {
			// Numerator has mandatory '0' digits: show the fraction with zeros.
			// When forceZero causes the whole digit to display (? or 0 types),
			// the middle literal should show as-is. When # suppresses, middle
			// is also suppressed.
			if wholePart == 0 && wholeType == tokDigitOpt && !wholeHasMandatory {
				// Pure # whole suppresses middle entirely.
			} else {
				result.WriteString(middle)
			}
			writeFrac(0, true)
		} else if wholeType == tokDigitSpace || numHasQ || denHasQ {
			// Whole is ? type, or fraction uses ? placeholders: pad fraction area with spaces.
			fracWidth := len(middle) + numPlaces + len(preSlash) + 1 + len(postSlash) + denPlaces
			result.WriteString(strings.Repeat(" ", fracWidth))
		}
		result.WriteString(suffix)
	} else if hasWhole {
		result.WriteString(prefix)
		if wholePart == 0 {
			switch wholeType {
			case tokDigitOpt: // # whole suppressed
				formatWholeDigits(&result, false)
				// If numerator has ? placeholders, replace middle with spaces
				// (keeps alignment). Otherwise suppress middle entirely.
				if numHasQ {
					result.WriteString(strings.Repeat(" ", len(middle)))
				}
			case tokDigitSpace: // ? whole: space-pad digit and replace middle chars with spaces
				formatWholeDigits(&result, false)
				result.WriteString(strings.Repeat(" ", len(middle)))
			default: // 0 whole: show '0' and show middle as-is
				formatWholeDigits(&result, false)
				result.WriteString(middle)
			}
		} else {
			formatWholeDigits(&result, false)
			result.WriteString(middle)
		}
		writeFrac(bestNum, false)
		result.WriteString(suffix)
	} else {
		totalNum := wholePart*bestDen + bestNum
		result.WriteString(prefix)
		writeFrac(totalNum, totalNum == 0)
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
	sciIdx := -1

	for i, tok := range tokens {
		switch tok.kind {
		case tokPercent:
			percentCount++
		case tokExponent:
			hasScientific = true
			sciIdx = i
		}
	}

	// Classify commas: grouping, scaling, or literal.
	commaClasses := classifyCommas(tokens)
	hasCommaGrouping := hasGroupingCommas(commaClasses)

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
	intZeros := 0  // minimum integer digits (from '0')
	intDigits := 0 // total integer placeholders (from '0', '#', '?')
	intSpaces := 0 // integer '?' count (space-padded positions)
	decZeros := 0  // decimal '0' count
	decHashes := 0 // decimal '#' count
	decSpaces := 0 // decimal '?' count

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
	totalDecPlaces := decZeros + decHashes + decSpaces
	_ = decIdx

	// Apply scaling commas: each scaling comma divides by 1000.
	scalingCommas := countScalingCommas(commaClasses)
	for tc := 0; tc < scalingCommas; tc++ {
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

	// Apply comma grouping if any commas were classified as grouping commas.
	if hasCommaGrouping {
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
		// Trim trailing zeros for '#' positions, then replace trailing
		// zeros with spaces for '?' positions.
		if decHashes > 0 {
			minLen := decZeros + decSpaces
			for len(decStr) > minLen && decStr[len(decStr)-1] == '0' {
				decStr = decStr[:len(decStr)-1]
			}
		}
		if decSpaces > 0 {
			buf := []byte(decStr)
			for i := len(buf) - 1; i >= decZeros && i >= len(buf)-decSpaces; i-- {
				if buf[i] == '0' {
					buf[i] = ' '
				} else {
					break
				}
			}
			decStr = string(buf)
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
			if hasCommaGrouping {
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

	// Check if literals are interleaved between integer digit placeholders
	// (e.g. phone format "(###) ###-####" or SSN "000-00-0000").
	// In this case, distribute digits right-to-left across placeholders.
	hasInterleavedLiterals := false
	if !hasCommaGrouping && totalDecPlaces == 0 {
		seenDigit := false
		seenLiteralAfterDigit := false
		for _, tok := range tokens {
			if tok.kind == tokDecimal {
				break
			}
			switch tok.kind {
			case tokDigit, tokDigitOpt, tokDigitSpace:
				if seenLiteralAfterDigit {
					hasInterleavedLiterals = true
				}
				seenDigit = true
			case tokLiteral:
				if seenDigit {
					// Only non-empty non-space-only literals count for interleaving detection.
					// Actually, any literal (including space) between digit groups counts.
					seenLiteralAfterDigit = true
				}
			}
		}
	}

	if hasInterleavedLiterals {
		// Distribute digits right-to-left across individual placeholder positions.
		// First, collect the integer digit positions in the token stream.
		rawDigits := strings.ReplaceAll(intStr, ",", "")
		// Pad with leading zeros if we have more placeholders than digits.
		for len(rawDigits) < intZeros {
			rawDigits = "0" + rawDigits
		}

		// Build a list of token indices that are integer digit placeholders.
		var digitTokenIdxs []int
		for ti, tok := range tokens {
			if tok.kind == tokDigit || tok.kind == tokDigitOpt || tok.kind == tokDigitSpace {
				digitTokenIdxs = append(digitTokenIdxs, ti)
			}
		}

		// Map digits right-to-left onto placeholder positions.
		digitChars := make([]byte, len(digitTokenIdxs))
		di := len(rawDigits) - 1
		for pi := len(digitTokenIdxs) - 1; pi >= 0; pi-- {
			if di >= 0 {
				digitChars[pi] = rawDigits[di]
				di--
			} else {
				// No more digits; use '0' for tokDigit, skip for tokDigitOpt.
				tok := tokens[digitTokenIdxs[pi]]
				if tok.kind == tokDigit {
					digitChars[pi] = '0'
				} else if tok.kind == tokDigitSpace {
					digitChars[pi] = ' '
				} else {
					digitChars[pi] = 0 // skip
				}
			}
		}

		// If there are extra digits that don't fit in placeholders, prepend them.
		var overflow string
		if di >= 0 {
			overflow = rawDigits[:di+1]
		}

		// Now build result from tokens.
		placeholderIdx := 0
		for ti, tok := range tokens {
			switch tok.kind {
			case tokLiteral:
				result.WriteString(tok.value)
			case tokDigit, tokDigitOpt, tokDigitSpace:
				if placeholderIdx == 0 && overflow != "" {
					result.WriteString(overflow)
				}
				ch := digitChars[placeholderIdx]
				if ch != 0 {
					result.WriteByte(ch)
				}
				placeholderIdx++
			case tokComma:
				if isLiteralComma(commaClasses, ti) {
					result.WriteByte(',')
				}
				// grouping and scaling commas are handled elsewhere; skip.
			case tokPercent:
				result.WriteByte('%')
			case tokDecimal:
				// should not happen in interleaved int-only format
			case tokExponent:
				// should not happen
			}
		}
		intWritten = true
	}

	if !hasInterleavedLiterals {
		for ti, tok := range tokens {
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
				} else {
					// Trailing dot with no decimal placeholders (e.g. "0.").
					result.WriteByte('.')
				}
				decWritten = true
				_ = decWritten
			case tokComma:
				if isLiteralComma(commaClasses, ti) {
					result.WriteByte(',')
				}
				// grouping and scaling commas are handled elsewhere; skip.
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
	}

	return result.String()
}

// Token types for number format parsing.
type numFmtTokenKind byte

const (
	tokLiteral    numFmtTokenKind = iota
	tokDigit                      // '0' — required digit
	tokDigitOpt                   // '#' — optional digit
	tokDigitSpace                 // '?' — digit padded with space
	tokDecimal                    // '.'
	tokComma                      // ','
	tokPercent                    // '%'
	tokExponent                   // 'E+' or 'E-'
)

type numFmtToken struct {
	kind  numFmtTokenKind
	value string
}

// tokenizeNumberFormat breaks a number format string into tokens.
func tokenizeNumberFormat(format string) []numFmtToken {
	var tokens []numFmtToken
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
			_, size := utf8.DecodeRuneInString(format[i:])
			tokens = append(tokens, numFmtToken{kind: tokLiteral, value: format[i : i+size]})
			i += size
			continue
		}

		// Underscore (space placeholder) — skip next char, emit space.
		if ch == '_' && i+1 < len(format) {
			_, size := utf8.DecodeRuneInString(format[i+1:])
			tokens = append(tokens, numFmtToken{kind: tokLiteral, value: " "})
			i += 1 + size
			continue
		}

		// Asterisk (repeat fill char) — skip next char.
		if ch == '*' && i+1 < len(format) {
			_, size := utf8.DecodeRuneInString(format[i+1:])
			i += 1 + size
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
		case 'E':
			// Scientific notation: E+, E-, or E followed directly by digit
			// placeholders (uppercase only; lowercase 'e' is treated as literal).
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
			} else if i+1 < len(format) && (format[i+1] == '0' || format[i+1] == '#') {
				// E followed directly by digit placeholders (no sign): behaves
				// like E- (sign shown only for negative exponents).
				tokens = append(tokens, numFmtToken{kind: tokExponent, value: "E-"})
				i++
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
			// Multi-byte UTF-8 characters (currency symbols like £, ¥, etc.)
			// must be consumed as a full rune to avoid double-encoding.
			if ch >= 0x80 {
				_, size := utf8.DecodeRuneInString(format[i:])
				tokens = append(tokens, numFmtToken{kind: tokLiteral, value: format[i : i+size]})
				i += size
			} else if isFormatLiteral(ch) {
				tokens = append(tokens, numFmtToken{kind: tokLiteral, value: string(rune(ch))})
				i++
			} else {
				// Unknown ASCII char — treat as literal.
				tokens = append(tokens, numFmtToken{kind: tokLiteral, value: string(rune(ch))})
				i++
			}
		}
	}

	return tokens
}

// isFormatLiteral determines if a character should be treated as a literal in a number format.
func isFormatLiteral(ch byte) bool {
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

// hasInvalidLowercaseE checks whether a format string contains lowercase 'e'
// followed by '+' or '-' outside of quoted strings or escape sequences.
// Only uppercase 'E' is recognised for scientific notation; lowercase 'e'
// in this position makes the entire format invalid (#VALUE!).
func hasInvalidLowercaseE(format string) bool {
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
			i++ // skip escaped character
			continue
		}
		if ch == 'e' && i+1 < len(format) && (format[i+1] == '+' || format[i+1] == '-') {
			return true
		}
	}
	return false
}

// commaClass classifies a comma token's role in a number format.
type commaClass byte

const (
	commaGrouping commaClass = iota // between digit placeholders: triggers thousands grouping
	commaScaling                    // trailing or adjacent to decimal point: divides by 1000
	commaLiteral                    // before first digit placeholder: emitted as literal text
)

// classifyCommas determines the role of each tokComma in the token stream.
// It returns a map from token index to commaClass.
//
// Excel rules:
//   - A comma BEFORE the first digit placeholder is a literal comma.
//   - A comma immediately before the decimal point (with no digit placeholder
//     between the comma and the decimal) is a scaling comma.
//   - A comma after the last digit placeholder (trailing) is a scaling comma.
//   - A comma between digit placeholders (digit on both sides) is a grouping comma.
func classifyCommas(tokens []numFmtToken) map[int]commaClass {
	result := make(map[int]commaClass)

	// Find the index of the first and last digit placeholder, and the decimal point.
	firstDigitIdx := -1
	lastDigitIdx := -1
	decimalIdx := -1
	for i, tok := range tokens {
		switch tok.kind {
		case tokDigit, tokDigitOpt, tokDigitSpace:
			if firstDigitIdx < 0 {
				firstDigitIdx = i
			}
			lastDigitIdx = i
		case tokDecimal:
			if decimalIdx < 0 {
				decimalIdx = i
			}
		}
	}

	for i, tok := range tokens {
		if tok.kind != tokComma {
			continue
		}

		// Rule 1: comma before the first digit placeholder is literal.
		if firstDigitIdx < 0 || i < firstDigitIdx {
			result[i] = commaLiteral
			continue
		}

		// Rule 2: comma immediately before the decimal point (no digit
		// placeholder between this comma and the decimal) is scaling.
		if decimalIdx >= 0 && i < decimalIdx {
			hasDigitBetween := false
			for j := i + 1; j < decimalIdx; j++ {
				if tokens[j].kind == tokDigit || tokens[j].kind == tokDigitOpt || tokens[j].kind == tokDigitSpace {
					hasDigitBetween = true
					break
				}
			}
			if !hasDigitBetween {
				result[i] = commaScaling
				continue
			}
		}

		// Rule 3: comma after the last digit placeholder is scaling (trailing).
		if i > lastDigitIdx {
			result[i] = commaScaling
			continue
		}

		// Rule 4: comma has digit placeholders on both sides = grouping.
		hasDigitBefore := false
		for j := i - 1; j >= 0; j-- {
			if tokens[j].kind == tokDigit || tokens[j].kind == tokDigitOpt || tokens[j].kind == tokDigitSpace {
				hasDigitBefore = true
				break
			}
			if tokens[j].kind == tokDecimal {
				break
			}
		}
		hasDigitAfter := false
		for j := i + 1; j < len(tokens); j++ {
			if tokens[j].kind == tokDigit || tokens[j].kind == tokDigitOpt || tokens[j].kind == tokDigitSpace {
				hasDigitAfter = true
				break
			}
			if tokens[j].kind == tokDecimal {
				break
			}
		}

		if hasDigitBefore && hasDigitAfter {
			result[i] = commaGrouping
		} else {
			// Fallback: treat as literal if it doesn't clearly fit another role.
			result[i] = commaLiteral
		}
	}

	return result
}

// countScalingCommas counts commas that act as scaling divisors (divide by 1000 each).
func countScalingCommas(commaClasses map[int]commaClass) int {
	count := 0
	for _, cls := range commaClasses {
		if cls == commaScaling {
			count++
		}
	}
	return count
}

// hasGroupingCommas returns true if any comma is classified as a grouping comma.
func hasGroupingCommas(commaClasses map[int]commaClass) bool {
	for _, cls := range commaClasses {
		if cls == commaGrouping {
			return true
		}
	}
	return false
}

// isLiteralComma returns true if the comma at the given token index is a literal.
func isLiteralComma(commaClasses map[int]commaClass, idx int) bool {
	cls, ok := commaClasses[idx]
	return ok && cls == commaLiteral
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
	// Count mantissa decimal places and required (non-optional) decimal places.
	decPlaces := 0
	minDecPlaces := 0
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
			if tok.kind == tokDigit {
				minDecPlaces = decPlaces // required '0' placeholder
			}
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

	// Partition pre-E tokens into: prefix literals, mantissa digits, middle literals.
	// Find the first and last digit/decimal token before E.
	firstDigit, lastDigit := -1, -1
	for i := 0; i < sciIdx; i++ {
		k := tokens[i].kind
		if k == tokDigit || k == tokDigitOpt || k == tokDigitSpace || k == tokDecimal {
			if firstDigit < 0 {
				firstDigit = i
			}
			lastDigit = i
		}
	}

	// Prefix literals: before the first coefficient digit.
	for i := 0; i < firstDigit; i++ {
		if tokens[i].kind == tokLiteral {
			result.WriteString(tokens[i].value)
		}
	}

	// Format mantissa.
	mStr := fmt.Sprintf("%.*f", decPlaces, mantissa)
	// Strip optional trailing zeros: '#' placeholders suppress trailing zeros,
	// but '0' placeholders require them. Strip down to minDecPlaces.
	if decPlaces > minDecPlaces {
		if dotIdx := strings.IndexByte(mStr, '.'); dotIdx >= 0 {
			fracPart := mStr[dotIdx+1:]
			trimmed := strings.TrimRight(fracPart, "0")
			if len(trimmed) < minDecPlaces {
				trimmed = fracPart[:minDecPlaces]
			}
			if len(trimmed) == 0 {
				mStr = mStr[:dotIdx+1] // keep trailing dot
			} else {
				mStr = mStr[:dotIdx+1] + trimmed
			}
		}
	}
	// If the format has a decimal point but no decimal digit placeholders
	// (e.g. "0.E0"), append the trailing dot that Sprintf omits.
	if decPlaces == 0 && inDecimal && !strings.ContainsRune(mStr, '.') {
		mStr += "."
	}
	result.WriteString(mStr)

	// Middle literals: after the last coefficient digit, before E.
	for i := lastDigit + 1; i < sciIdx; i++ {
		if tokens[i].kind == tokLiteral {
			result.WriteString(tokens[i].value)
		}
	}

	// Format exponent: E, then any pre-exponent literals, then sign, then digits.
	result.WriteByte('E')

	// Pre-exponent literals: between E token and first exponent digit.
	firstExpDigit := -1
	for i := sciIdx + 1; i < len(tokens); i++ {
		k := tokens[i].kind
		if k == tokDigit || k == tokDigitOpt || k == tokDigitSpace {
			firstExpDigit = i
			break
		}
	}
	if firstExpDigit > sciIdx+1 {
		for i := sciIdx + 1; i < firstExpDigit; i++ {
			if tokens[i].kind == tokLiteral {
				result.WriteString(tokens[i].value)
			}
		}
	}

	// Exponent sign: E- shows sign only if negative; E+ always shows sign.
	if exp >= 0 {
		if expSign == "+" {
			result.WriteByte('+')
		}
	} else {
		result.WriteByte('-')
		exp = -exp
	}

	expStr := strconv.Itoa(exp)
	for len(expStr) < expDigits {
		expStr = "0" + expStr
	}
	result.WriteString(expStr)

	// Suffix literals: after the last exponent digit placeholder.
	lastExpDigit := -1
	for i := sciIdx + 1; i < len(tokens); i++ {
		k := tokens[i].kind
		if k == tokDigit || k == tokDigitOpt || k == tokDigitSpace {
			lastExpDigit = i
		}
	}
	for i := lastExpDigit + 1; i < len(tokens); i++ {
		if tokens[i].kind == tokLiteral {
			result.WriteString(tokens[i].value)
		} else if tokens[i].kind == tokPercent {
			result.WriteByte('%')
		}
	}

	return result.String()
}

// ---------------------------------------------------------------------------
// Time helpers
// ---------------------------------------------------------------------------

// daysSinceEpoch converts a time.Time to serial days (for internal elapsed time).
func daysSinceEpoch(t time.Time) float64 {
	duration := t.Sub(Epoch1900)
	return duration.Hours() / 24
}
