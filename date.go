package werkbook

import (
	"strings"
	"time"
)

// excelEpoch is "January 0, 1900" = December 31, 1899.
// Excel serial 1 = January 1, 1900 = epoch + 1 day.
var excelEpoch = time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)

// excel1904Epoch is January 1, 1904.
// In the 1904 date system, serial number 0 = January 1, 1904.
// Unlike the 1900 system, there is no phantom "January 0" offset.
var excel1904Epoch = time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)

// timeToExcelSerial converts a time.Time to an Excel serial date number.
func timeToExcelSerial(t time.Time) float64 {
	duration := t.Sub(excelEpoch)
	days := duration.Hours() / 24
	// Excel 1900 leap year bug: Excel thinks Feb 29, 1900 exists (serial 60).
	// Dates on or after March 1, 1900 (real day 60) need serial incremented by 1.
	if days >= 60 {
		days++
	}
	return days
}

// ExcelSerialToTime converts an Excel serial date number to a time.Time.
func ExcelSerialToTime(serial float64) time.Time {
	return excelSerialToTime(serial)
}

// excelSerialToTime converts an Excel serial date number to a time.Time.
func excelSerialToTime(serial float64) time.Time {
	// Excel 1900 leap year bug: serial 60 is the phantom Feb 29.
	// For serial > 60, subtract 1 to get the real day count.
	if serial > 60 {
		serial--
	}
	days := int(serial)
	frac := serial - float64(days)
	t := excelEpoch.AddDate(0, 0, days)
	t = t.Add(time.Duration(frac * 24 * float64(time.Hour)))
	return t
}

// excelSerialToTime1904 converts an Excel serial date number to a time.Time
// using the 1904 date system (Mac Excel). No leap-year bug adjustment is needed.
func excelSerialToTime1904(serial float64) time.Time {
	days := int(serial)
	frac := serial - float64(days)
	t := excel1904Epoch.AddDate(0, 0, days)
	t = t.Add(time.Duration(frac * 24 * float64(time.Hour)))
	return t
}

// IsDateFormat reports whether a number format (identified by format string
// and/or built-in ID) represents a date or time value.
func IsDateFormat(numFmt string, numFmtID int) bool {
	// Check well-known built-in date format IDs.
	switch numFmtID {
	case 14, 15, 16, 17, 18, 19, 20, 21, 22,
		27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
		45, 46, 47,
		50, 51, 52, 53, 54, 55, 56, 57, 58:
		return true
	}

	if numFmt == "" {
		return false
	}

	// Scan the format string for date/time tokens after stripping literals.
	stripped := stripDateLiterals(numFmt)
	upper := strings.ToUpper(stripped)

	// Remove elapsed time markers.
	for _, marker := range []string{"[H]", "[M]", "[S]", "[HH]", "[MM]", "[SS]"} {
		upper = strings.ReplaceAll(upper, marker, "")
	}

	for i := 0; i < len(upper); i++ {
		switch upper[i] {
		case 'Y', 'D', 'H', 'S':
			return true
		case 'M':
			return true
		case 'A':
			if i+4 <= len(upper) && upper[i:i+4] == "AM/P" {
				return true
			}
			if i+2 <= len(upper) && (upper[i:i+2] == "AM" || upper[i:i+2] == "A/") {
				return true
			}
		}
	}
	return false
}

// stripDateLiterals removes quoted and escaped characters from a format string.
func stripDateLiterals(format string) string {
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
			i++
			continue
		}
		b.WriteByte(ch)
	}
	return b.String()
}
