package formula

import (
	"math"
	"strings"
	"time"
	"unicode"
)

// parseDateTimeString parses a wide variety of text date and date-time
// formats and returns the Excel 1900-system serial value.
//
// Returns (serial, hasTime, ok):
//   - serial is the date serial, with the fractional component populated
//     when the input carried a parseable time portion;
//   - hasTime is true when the string contained a time component;
//   - ok is false when no layout matched.
//
// The parser is intentionally liberal about case and separators so that
// DATEVALUE, TIMEVALUE, and implicit text→number coercion can all share
// a single implementation that tracks Excel behaviour.
func parseDateTimeString(text string) (serial float64, hasTime bool, ok bool) {
	text = strings.TrimSpace(text)
	if text == "" {
		return 0, false, false
	}

	// Excel's fictional 1900-02-29 leap day — time.Parse can never accept it
	// because the date doesn't exist, so short-circuit before touching the
	// layout list.
	if s, ok := parseExcelLeapBug(text); ok {
		return s, false, true
	}

	// Normalise case on alphabetic runs so "jan" and "JANUARY" match the
	// Title-cased layouts below. A trailing AM/PM marker must stay
	// upper-case or time.Parse won't recognise it.
	normalized := titleCaseAlpha(text)
	normalized = strings.ReplaceAll(normalized, " Am", " AM")
	normalized = strings.ReplaceAll(normalized, " Pm", " PM")

	for _, layout := range dateTimeLayouts {
		t, err := time.Parse(layout, normalized)
		if err == nil {
			return TimeToSerial(t), true, true
		}
	}

	for _, layout := range dateOnlyLayouts {
		t, err := time.Parse(layout, normalized)
		if err == nil {
			return math.Floor(TimeToSerial(t)), false, true
		}
	}

	// "Month Day" without a year — use the current year.
	if m, d, ok := parseMonthDay(normalized); ok {
		now := time.Now()
		t := time.Date(now.Year(), m, d, 0, 0, 0, 0, time.UTC)
		return math.Floor(TimeToSerial(t)), false, true
	}

	// "M/D" without a year — use the current year.
	if m, d, ok := parseNumericMonthDay(normalized); ok {
		now := time.Now()
		t := time.Date(now.Year(), m, d, 0, 0, 0, 0, time.UTC)
		return math.Floor(TimeToSerial(t)), false, true
	}

	return 0, false, false
}

// dateTimeLayouts are Go layouts that include a time component. They must be
// tried before the date-only set so "2024-01-15 12:30:00" doesn't silently
// lose its fractional-day portion.
//
// Excel's DATEVALUE rejects the ISO 8601 "T" separator and the trailing "Z"
// suffix, so we deliberately omit those layouts — accepting them would make
// parseDateTimeString succeed on inputs Excel treats as #VALUE!.
var dateTimeLayouts = []string{
	"2006-01-02 15:04:05",
	"2006-01-02 15:04",
	"2006-01-02 3:04:05 PM",
	"2006-01-02 3:04 PM",
	"1/2/2006 15:04:05",
	"1/2/2006 15:04",
	"1/2/2006 3:04:05 PM",
	"1/2/2006 3:04 PM",
}

// dateOnlyLayouts cover the formats Excel will coerce without a time part.
// Order matters because several patterns are structurally similar — put the
// more specific layouts first so ambiguous inputs resolve the way Excel does
// in en-US.
var dateOnlyLayouts = []string{
	"1/2/2006",
	"01/02/2006",
	"1-2-2006",
	"01-02-2006",
	"2-Jan-2006",
	"02-Jan-2006",
	"2-January-2006",
	"02-January-2006",
	"2006/01/02",
	"2006-01-02",
	"January 2, 2006",
	"Jan 2, 2006",
	"2 January 2006",
	"2 Jan 2006",
	"1/2/06",
	"01/02/06",
	"January 2006",
	"Jan 2006",
}

// parseExcelLeapBug recognises the fictional 1900-02-29 that Excel exposes as
// serial 60 (and the equivalent US-style variants).
func parseExcelLeapBug(text string) (float64, bool) {
	switch strings.ToLower(text) {
	case "1900-02-29", "02/29/1900", "2/29/1900", "1900/02/29":
		return 60, true
	}
	return 0, false
}

// titleCaseAlpha returns s with each run of ASCII letters Title-Cased: the
// first letter becomes upper-case and subsequent letters become lower-case.
// Non-letter characters pass through unchanged. This lets case-insensitive
// month names match Go's case-sensitive time.Parse layouts.
func titleCaseAlpha(s string) string {
	var b strings.Builder
	b.Grow(len(s))
	startOfRun := true
	for _, r := range s {
		if unicode.IsLetter(r) {
			if startOfRun {
				b.WriteRune(unicode.ToUpper(r))
				startOfRun = false
			} else {
				b.WriteRune(unicode.ToLower(r))
			}
		} else {
			b.WriteRune(r)
			startOfRun = true
		}
	}
	return b.String()
}

// parseNumericMonthDay parses "M/D" (e.g. "5/5") with no year component.
func parseNumericMonthDay(s string) (time.Month, int, bool) {
	parts := strings.Split(s, "/")
	if len(parts) != 2 {
		return 0, 0, false
	}
	m, err := parseInt(parts[0])
	if err != nil || m < 1 || m > 12 {
		return 0, 0, false
	}
	d, err := parseInt(parts[1])
	if err != nil || d < 1 || d > 31 {
		return 0, 0, false
	}
	return time.Month(m), d, true
}
