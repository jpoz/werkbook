package formula

import (
	"math"
	"strings"
	"time"
)

func init() {
	Register("DATE", NoCtx(fnDATE))
	Register("DATEDIF", NoCtx(fnDATEDIF))
	Register("DATEVALUE", NoCtx(fnDATEVALUE))
	Register("DAY", NoCtx(fnDAY))
	Register("DAYS", NoCtx(fnDAYS))
	Register("DAYS360", NoCtx(fnDAYS360))
	Register("EDATE", NoCtx(fnEDATE))
	Register("EOMONTH", NoCtx(fnEOMONTH))
	Register("HOUR", NoCtx(fnHOUR))
	Register("ISOWEEKNUM", NoCtx(fnISOWEEKNUM))
	Register("MINUTE", NoCtx(fnMINUTE))
	Register("MONTH", NoCtx(fnMONTH))
	Register("NETWORKDAYS", NoCtx(fnNETWORKDAYS))
	Register("NETWORKDAYS.INTL", NoCtx(fnNetworkdaysIntl))
	Register("NOW", NoCtx(fnNOW))
	Register("SECOND", NoCtx(fnSECOND))
	Register("TIME", NoCtx(fnTIME))
	Register("TIMEVALUE", NoCtx(fnTIMEVALUE))
	Register("TODAY", NoCtx(fnTODAY))
	Register("WEEKDAY", NoCtx(fnWEEKDAY))
	Register("WEEKNUM", NoCtx(fnWEEKNUM))
	Register("WORKDAY", NoCtx(fnWORKDAY))
	Register("WORKDAY.INTL", NoCtx(fnWorkdayIntl))
	Register("YEAR", NoCtx(fnYEAR))
	Register("YEARFRAC", NoCtx(fnYEARFRAC))
}

// Serial date helpers — duplicated from werkbook/date.go to avoid circular imports.
var ExcelEpoch = time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)

// Excel1904Epoch is January 1, 1904.
// In the 1904 date system, serial number 0 = January 1, 1904, and there is
// no leap-year bug to compensate for.
var Excel1904Epoch = time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)

// MaxExcelSerial is the largest valid Excel serial date (Dec 31, 9999).
const MaxExcelSerial = 2958465

func TimeToExcelSerial(t time.Time) float64 {
	// Calculate the number of days between the Excel epoch and the given time.
	// We cannot use t.Sub(ExcelEpoch) because time.Duration is an int64 of
	// nanoseconds, which overflows for dates more than ~292 years from the epoch.
	// Instead, compute whole days via calendar difference and fractional day
	// from the time-of-day components.
	y1, m1, d1 := ExcelEpoch.Date()
	y2, m2, d2 := t.Date()
	epochDays := julianDayNumber(y1, int(m1), d1)
	tDays := julianDayNumber(y2, int(m2), d2)
	days := float64(tDays - epochDays)

	// Add fractional day from time-of-day.
	h, min, sec := t.Clock()
	days += (float64(h)*3600 + float64(min)*60 + float64(sec)) / 86400.0

	if days >= 60 {
		days++
	}
	return days
}

// julianDayNumber returns a Julian Day Number for the given date, useful for
// computing the difference in days between two dates without overflow.
func julianDayNumber(year, month, day int) int {
	// Algorithm from Meeus, "Astronomical Algorithms".
	a := (14 - month) / 12
	y := year + 4800 - a
	m := month + 12*a - 3
	return day + (153*m+2)/5 + 365*y + y/4 - y/100 + y/400 - 32045
}

func ExcelSerialToTime(serial float64) time.Time {
	if serial > 60 {
		serial--
	}
	days := int(serial)
	frac := serial - float64(days)
	t := ExcelEpoch.AddDate(0, 0, days)
	t = t.Add(time.Duration(frac * 24 * float64(time.Hour)))
	return t
}

// excelSerialDateParts returns the year, month, day for an Excel serial number,
// correctly handling serial 60 (Excel's fictional Feb 29, 1900).
func excelSerialDateParts(serial float64) (int, time.Month, int) {
	s := int(serial)
	if s == 0 {
		// Serial 0 is Excel's "January 0, 1900" sentinel value.
		return 1900, time.January, 0
	}
	if s == 60 {
		return 1900, time.February, 29
	}
	t := ExcelSerialToTime(serial)
	return t.Year(), t.Month(), t.Day()
}

// ExcelSerialToTime1904 converts an Excel serial date number to a time.Time
// using the 1904 date system (Mac Excel). No leap-year bug adjustment is needed.
func ExcelSerialToTime1904(serial float64) time.Time {
	days := int(serial)
	frac := serial - float64(days)
	t := Excel1904Epoch.AddDate(0, 0, days)
	t = t.Add(time.Duration(frac * 24 * float64(time.Hour)))
	return t
}

func fnDATE(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	year, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	month, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	day, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	// Excel checks the raw float BEFORE truncation: negative values → #NUM!
	// e.g. int(-0.5) truncates to 0 in Go, but Excel sees -0.5 < 0 → #NUM!
	if year < 0 || year >= 10000 {
		return ErrorVal(ErrValNUM), nil
	}

	y := int(year)

	// Excel adds 1900 to years in the range 0–1899.
	// e.g. DATE(108,1,2) → year 2008.
	if y >= 0 && y <= 1899 {
		y += 1900
	}

	// After normalization the year must still be in range.
	if y < 0 || y >= 10000 {
		return ErrorVal(ErrValNUM), nil
	}

	// Excel uses INT (floor) semantics to truncate month and day to integers,
	// not TRUNC (toward zero). E.g. INT(-0.5) = -1, not 0.
	m := int(math.Floor(month))
	d := int(math.Floor(day))

	// Guard against extreme month/day values that would overflow time.Duration
	// (max ≈ 292 years in nanoseconds). Excel's valid range is 1/1/1900–12/31/9999,
	// so values that shift the year far outside always produce #NUM!.
	if m < -120000 || m > 120000 || d < -4000000 || d > 4000000 {
		return ErrorVal(ErrValNUM), nil
	}

	// Special-case: DATE(1900,2,29) should return serial 60 (Excel's fictional
	// leap day). Go's time.Date normalizes Feb 29, 1900 to Mar 1, 1900, which
	// would produce serial 61 via TimeToExcelSerial, so we intercept it here.
	if y == 1900 && m == 2 && d == 29 {
		return NumberVal(60), nil
	}

	t := time.Date(y, time.Month(m), d, 0, 0, 0, 0, time.UTC)
	serial := TimeToExcelSerial(t)
	if serial < 0 || serial > MaxExcelSerial {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(serial), nil
}

func fnDAY(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n < 0 || n > MaxExcelSerial {
		return ErrorVal(ErrValNUM), nil
	}
	_, _, day := excelSerialDateParts(n)
	return NumberVal(float64(day)), nil
}

func fnMONTH(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n < 0 || n > MaxExcelSerial {
		return ErrorVal(ErrValNUM), nil
	}
	_, month, _ := excelSerialDateParts(n)
	return NumberVal(float64(month)), nil
}

func fnNOW(args []Value) (Value, error) {
	if len(args) != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return NumberVal(TimeToExcelSerial(time.Now())), nil
}

func fnTODAY(args []Value) (Value, error) {
	if len(args) != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	now := time.Now()
	today := time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, time.UTC)
	return NumberVal(math.Floor(TimeToExcelSerial(today))), nil
}

func fnYEAR(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n < 0 || n > MaxExcelSerial {
		return ErrorVal(ErrValNUM), nil
	}
	year, _, _ := excelSerialDateParts(n)
	return NumberVal(float64(year)), nil
}

func isLeapYear(year int) bool {
	return year%4 == 0 && (year%100 != 0 || year%400 == 0)
}

func fnDATEDIF(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	startSerial, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	endSerial, e2 := CoerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}
	unitStr := strings.ToUpper(ValueToString(args[2]))

	if startSerial > endSerial {
		return ErrorVal(ErrValNUM), nil
	}

	start := ExcelSerialToTime(startSerial)
	end := ExcelSerialToTime(endSerial)

	switch unitStr {
	case "D":
		days := int(endSerial - startSerial)
		return NumberVal(float64(days)), nil
	case "M":
		months := (end.Year()-start.Year())*12 + int(end.Month()) - int(start.Month())
		if end.Day() < start.Day() {
			months--
		}
		return NumberVal(float64(months)), nil
	case "Y":
		years := end.Year() - start.Year()
		if end.Month() < start.Month() || (end.Month() == start.Month() && end.Day() < start.Day()) {
			years--
		}
		return NumberVal(float64(years)), nil
	case "MD":
		d := end.Day() - start.Day()
		if d < 0 {
			prevMonth := time.Date(end.Year(), end.Month(), 0, 0, 0, 0, 0, time.UTC)
			d = prevMonth.Day() - start.Day() + end.Day()
		}
		return NumberVal(float64(d)), nil
	case "YM":
		m := int(end.Month()) - int(start.Month())
		if end.Day() < start.Day() {
			m--
		}
		if m < 0 {
			m += 12
		}
		return NumberVal(float64(m)), nil
	case "YD":
		// Excel's YD algorithm: move end's month/day into start's year,
		// then compute the day difference from start to the adjusted date.
		endMonth := end.Month()
		endDay := end.Day()

		// Handle Feb 29: if end is Feb 29 and start's year is not a leap year,
		// treat as Feb 28.
		adjYear := start.Year()
		if endMonth == 2 && endDay == 29 && !isLeapYear(adjYear) {
			endDay = 28
		}

		adjusted := time.Date(adjYear, endMonth, endDay, 0, 0, 0, 0, time.UTC)
		if adjusted.Before(start) {
			// Move to next year
			adjYear++
			// Re-check Feb 29 for the new year
			if end.Month() == 2 && end.Day() == 29 && !isLeapYear(adjYear) {
				adjusted = time.Date(adjYear, 2, 28, 0, 0, 0, 0, time.UTC)
			} else {
				adjusted = time.Date(adjYear, end.Month(), end.Day(), 0, 0, 0, 0, time.UTC)
			}
		}

		days := int(adjusted.Sub(start).Hours() / 24)
		return NumberVal(float64(days)), nil
	default:
		return ErrorVal(ErrValNUM), nil
	}
}

func fnDAYS(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	end, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	start, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	return NumberVal(math.Trunc(end) - math.Trunc(start)), nil
}

// isLastDayOfFeb reports whether the given date is the last day of February.
func isLastDayOfFeb(year, month, day int) bool {
	if month != 2 {
		return false
	}
	if isLeapYear(year) {
		return day == 29
	}
	return day == 28
}

// days360Calc computes the number of days between two dates using the 30/360 convention.
func days360Calc(sy, sm, sd, ey, em, ed int, european bool) float64 {
	if european {
		if sd == 31 {
			sd = 30
		}
		if ed == 31 {
			ed = 30
		}
	} else {
		// US (NASD) method — order of checks matters:
		// 1. If both dates are last day of February, set D2 to 30.
		if isLastDayOfFeb(sy, sm, sd) && isLastDayOfFeb(ey, em, ed) {
			ed = 30
		}
		// 2. If start date is last day of February, set D1 to 30.
		if isLastDayOfFeb(sy, sm, sd) {
			sd = 30
		}
		// 3. If D2 is 31 and D1 is 30 or 31, set D2 to 30.
		if ed == 31 && sd >= 30 {
			ed = 30
		}
		// 4. If D1 is 31, set D1 to 30.
		if sd == 31 {
			sd = 30
		}
	}
	return float64((ey-sy)*360 + (em-sm)*30 + (ed - sd))
}

func fnDAYS360(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	startSerial, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	endSerial, e2 := CoerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}

	european := false
	if len(args) == 3 {
		m, e3 := CoerceNum(args[2])
		if e3 != nil {
			return *e3, nil
		}
		european = m != 0
	}

	start := ExcelSerialToTime(math.Trunc(startSerial))
	end := ExcelSerialToTime(math.Trunc(endSerial))

	sy, sm, sd := start.Year(), int(start.Month()), start.Day()
	ey, em, ed := end.Year(), int(end.Month()), end.Day()

	return NumberVal(days360Calc(sy, sm, sd, ey, em, ed, european)), nil
}

func fnHOUR(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	hour, _, _ := excelSerialTimeParts(n)
	return NumberVal(float64(hour)), nil
}

func fnMINUTE(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	_, minute, _ := excelSerialTimeParts(n)
	return NumberVal(float64(minute)), nil
}

func fnSECOND(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	_, _, second := excelSerialTimeParts(n)
	return NumberVal(float64(second)), nil
}

func excelSerialTimeParts(serial float64) (int, int, int) {
	frac := serial - math.Floor(serial)
	totalSeconds := int(math.Round(frac * 86400))
	if totalSeconds == 86400 {
		totalSeconds = 0
	}
	hour := totalSeconds / 3600
	minute := (totalSeconds % 3600) / 60
	second := totalSeconds % 60
	return hour, minute, second
}

func fnTIME(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	hour, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	minute, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	second, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	// Excel truncates arguments to integers.
	hour = math.Trunc(hour)
	minute = math.Trunc(minute)
	second = math.Trunc(second)

	// Excel returns #NUM! if any argument exceeds 32767.
	if hour > 32767 || minute > 32767 || second > 32767 {
		return ErrorVal(ErrValNUM), nil
	}

	totalSeconds := hour*3600 + minute*60 + second
	// Excel returns #NUM! if the total time is negative.
	if totalSeconds < 0 {
		return ErrorVal(ErrValNUM), nil
	}

	result := totalSeconds / 86400
	// TIME returns only the fractional part (time of day).
	// Values >= 1.0 wrap; e.g. TIME(25,0,0) = 0.04167 (just the 1-hour fraction).
	result = result - math.Floor(result)
	return NumberVal(result), nil
}

func fnWEEKDAY(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	serial, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	returnType := 1.0
	if len(args) == 2 {
		returnType, e = CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
	}

	t := ExcelSerialToTime(serial)
	wd := int(t.Weekday())

	rt := int(returnType)
	var result int
	switch rt {
	case 1, 17:
		result = wd + 1
	case 2, 11:
		result = (wd+6)%7 + 1
	case 3:
		result = (wd + 6) % 7
	case 12:
		result = (wd-2+7)%7 + 1
	case 13:
		result = (wd-3+7)%7 + 1
	case 14:
		result = (wd-4+7)%7 + 1
	case 15:
		result = (wd-5+7)%7 + 1
	case 16:
		result = (wd-6+7)%7 + 1
	default:
		return ErrorVal(ErrValNUM), nil
	}

	return NumberVal(float64(result)), nil
}

func fnISOWEEKNUM(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	serial, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	t := ExcelSerialToTime(serial)
	_, week := t.ISOWeek()
	return NumberVal(float64(week)), nil
}

func fnDATEVALUE(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	text := strings.TrimSpace(ValueToString(args[0]))

	// Strip time portion from date-time strings like "2025-03-04 12:00".
	// Only strip if the part after the space looks like a time (contains a colon).
	dateOnly := text
	if idx := strings.Index(text, " "); idx > 0 {
		rest := strings.TrimSpace(text[idx+1:])
		if strings.Contains(rest, ":") {
			dateOnly = text[:idx]
		}
	}

	layouts := []string{
		"1/2/2006",
		"01/02/2006",
		"2-Jan-2006",
		"02-Jan-2006",
		"2006/01/02",
		"2006-01-02",
		"January 2, 2006",
	}

	for _, layout := range layouts {
		t, err := time.Parse(layout, dateOnly)
		if err == nil {
			return NumberVal(math.Floor(TimeToExcelSerial(t))), nil
		}
	}

	// Try 2-digit year formats: MM/DD/YY
	twoDigitLayouts := []string{
		"1/2/06",
		"01/02/06",
	}
	for _, layout := range twoDigitLayouts {
		t, err := time.Parse(layout, dateOnly)
		if err == nil {
			return NumberVal(math.Floor(TimeToExcelSerial(t))), nil
		}
	}

	// Try "Month Day" without year — use current year.
	if m, d, ok := parseMonthDay(dateOnly); ok {
		now := time.Now()
		t := time.Date(now.Year(), m, d, 0, 0, 0, 0, time.UTC)
		return NumberVal(math.Floor(TimeToExcelSerial(t))), nil
	}

	return ErrorVal(ErrValVALUE), nil
}

func fnTIMEVALUE(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	text := strings.TrimSpace(ValueToString(args[0]))

	t, ok := parseTimeString(text)
	if !ok {
		return ErrorVal(ErrValVALUE), nil
	}

	return NumberVal(t), nil
}

// parseTimeString parses a time string and returns the fractional day value.
// Supported formats: "H:MM", "H:MM:SS", "H:MM AM/PM", "H:MM:SS AM/PM"
func parseTimeString(text string) (float64, bool) {
	upper := strings.ToUpper(strings.TrimSpace(text))

	// Check for AM/PM suffix.
	isPM := false
	hasAMPM := false
	if strings.HasSuffix(upper, " AM") {
		hasAMPM = true
		upper = strings.TrimSpace(upper[:len(upper)-3])
	} else if strings.HasSuffix(upper, " PM") {
		hasAMPM = true
		isPM = true
		upper = strings.TrimSpace(upper[:len(upper)-3])
	}

	// Split on ":"
	parts := strings.Split(upper, ":")
	if len(parts) < 2 || len(parts) > 3 {
		return 0, false
	}

	var hour, minute, second int
	var err error

	// Use fmt.Sscanf-like parsing with strconv
	hour, err = parseInt(parts[0])
	if err != nil {
		return 0, false
	}
	minute, err = parseInt(parts[1])
	if err != nil {
		return 0, false
	}
	if len(parts) == 3 {
		second, err = parseInt(parts[2])
		if err != nil {
			return 0, false
		}
	}

	if hasAMPM {
		if hour < 1 || hour > 12 {
			return 0, false
		}
		if hour == 12 {
			hour = 0
		}
		if isPM {
			hour += 12
		}
	}

	if minute < 0 || minute > 59 || second < 0 || second > 59 {
		return 0, false
	}

	totalSeconds := float64(hour)*3600 + float64(minute)*60 + float64(second)
	return totalSeconds / 86400.0, true
}

// parseInt parses a string as a non-negative integer, trimming whitespace.
func parseInt(s string) (int, error) {
	s = strings.TrimSpace(s)
	n := 0
	if len(s) == 0 {
		return 0, errParseInt
	}
	for _, c := range s {
		if c < '0' || c > '9' {
			return 0, errParseInt
		}
		n = n*10 + int(c-'0')
	}
	return n, nil
}

var errParseInt = errorString("invalid integer")

type errorString string

func (e errorString) Error() string { return string(e) }

// monthNames maps lowercased full and abbreviated month names to their month number.
var monthNames = map[string]time.Month{
	"january": time.January, "jan": time.January,
	"february": time.February, "feb": time.February,
	"march": time.March, "mar": time.March,
	"april": time.April, "apr": time.April,
	"may":  time.May,
	"june": time.June, "jun": time.June,
	"july": time.July, "jul": time.July,
	"august": time.August, "aug": time.August,
	"september": time.September, "sep": time.September,
	"october": time.October, "oct": time.October,
	"november": time.November, "nov": time.November,
	"december": time.December, "dec": time.December,
}

// parseMonthDay parses strings like "March 4" or "Jan 15" and returns month and day.
func parseMonthDay(s string) (time.Month, int, bool) {
	parts := strings.Fields(s)
	if len(parts) != 2 {
		return 0, 0, false
	}
	m, ok := monthNames[strings.ToLower(parts[0])]
	if !ok {
		return 0, 0, false
	}
	d, err := parseInt(parts[1])
	if err != nil || d < 1 || d > 31 {
		return 0, 0, false
	}
	return m, d, true
}

func fnEDATE(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	serial, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	months, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	t := ExcelSerialToTime(serial)
	y, mo, d := t.Date()
	m := int(math.Trunc(months))
	targetMonth := time.Month(int(mo) + m)
	targetYear := y
	for targetMonth > 12 {
		targetMonth -= 12
		targetYear++
	}
	for targetMonth < 1 {
		targetMonth += 12
		targetYear--
	}
	result := time.Date(targetYear, targetMonth, d, 0, 0, 0, 0, time.UTC)
	if result.Month() != targetMonth {
		result = time.Date(targetYear, targetMonth+1, 0, 0, 0, 0, 0, time.UTC)
	}
	return NumberVal(math.Floor(TimeToExcelSerial(result))), nil
}

func fnEOMONTH(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	serial, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	months, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	t := ExcelSerialToTime(serial)
	y, m, _ := t.Date()
	last := time.Date(y, m+time.Month(int(math.Trunc(months))+1), 0, 0, 0, 0, 0, time.UTC)
	return NumberVal(math.Floor(TimeToExcelSerial(last))), nil
}

// parseHolidays extracts a set of truncated serial dates from an optional holiday argument.
func parseHolidays(arg Value) (map[float64]bool, *Value) {
	holidays := make(map[float64]bool)
	if arg.Type == ValueArray {
		for _, row := range arg.Array {
			for _, cell := range row {
				if cell.Type == ValueError {
					return nil, &cell
				}
				n, ce := CoerceNum(cell)
				if ce != nil {
					return nil, ce
				}
				holidays[math.Trunc(n)] = true
			}
		}
	} else if arg.Type != ValueEmpty {
		n, ce := CoerceNum(arg)
		if ce != nil {
			return nil, ce
		}
		holidays[math.Trunc(n)] = true
	}
	return holidays, nil
}

func fnNETWORKDAYS(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	startSerial, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	endSerial, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	startSerial = math.Trunc(startSerial)
	endSerial = math.Trunc(endSerial)

	holidays := make(map[float64]bool)
	if len(args) == 3 {
		var ev *Value
		holidays, ev = parseHolidays(args[2])
		if ev != nil {
			return *ev, nil
		}
	}

	negate := false
	from := startSerial
	to := endSerial
	if from > to {
		from, to = to, from
		negate = true
	}

	count := 0.0
	for d := from; d <= to; d++ {
		t := ExcelSerialToTime(d)
		wd := t.Weekday()
		if wd != time.Saturday && wd != time.Sunday && !holidays[d] {
			count++
		}
	}

	if negate {
		count = -count
	}
	return NumberVal(count), nil
}

func fnNetworkdaysIntl(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	startSerial, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	endSerial, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	startSerial = math.Trunc(startSerial)
	endSerial = math.Trunc(endSerial)

	// Default weekend: Saturday and Sunday (code 1).
	weekend := [7]bool{true, false, false, false, false, false, true}
	if len(args) >= 3 && args[2].Type != ValueEmpty {
		var ev *Value
		weekend, ev = parseWeekendParam(args[2])
		if ev != nil {
			return *ev, nil
		}
	}

	holidays := make(map[float64]bool)
	if len(args) == 4 {
		var ev *Value
		holidays, ev = parseHolidays(args[3])
		if ev != nil {
			return *ev, nil
		}
	}

	negate := false
	from := startSerial
	to := endSerial
	if from > to {
		from, to = to, from
		negate = true
	}

	count := 0.0
	for d := from; d <= to; d++ {
		t := ExcelSerialToTime(d)
		wd := int(t.Weekday())
		if !weekend[wd] && !holidays[d] {
			count++
		}
	}

	if negate {
		count = -count
	}
	return NumberVal(count), nil
}

func fnWEEKNUM(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	serial, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	returnType := 1.0
	if len(args) == 2 {
		returnType, e = CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
	}

	rt := int(returnType)
	t := ExcelSerialToTime(serial)

	if rt == 21 {
		_, isoWeek := t.ISOWeek()
		return NumberVal(float64(isoWeek)), nil
	}

	var weekStart time.Weekday
	switch rt {
	case 1, 17:
		weekStart = time.Sunday
	case 2, 11:
		weekStart = time.Monday
	case 12:
		weekStart = time.Tuesday
	case 13:
		weekStart = time.Wednesday
	case 14:
		weekStart = time.Thursday
	case 15:
		weekStart = time.Friday
	case 16:
		weekStart = time.Saturday
	default:
		return ErrorVal(ErrValNUM), nil
	}

	jan1 := time.Date(t.Year(), 1, 1, 0, 0, 0, 0, time.UTC)
	jan1Wd := jan1.Weekday()

	offset := int(jan1Wd-weekStart+7) % 7
	dayOfYear := t.YearDay()
	weekNum := (dayOfYear+offset-1)/7 + 1
	return NumberVal(float64(weekNum)), nil
}

func fnYEARFRAC(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	startSerial, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	endSerial, e2 := CoerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}
	basis := 0
	if len(args) == 3 {
		b, e3 := CoerceNum(args[2])
		if e3 != nil {
			return *e3, nil
		}
		basis = int(b)
	}
	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}

	if startSerial > endSerial {
		startSerial, endSerial = endSerial, startSerial
	}

	start := ExcelSerialToTime(startSerial)
	end := ExcelSerialToTime(endSerial)

	sy, sm, sd := start.Year(), int(start.Month()), start.Day()
	ey, em, ed := end.Year(), int(end.Month()), end.Day()

	switch basis {
	case 0: // US (NASD) 30/360
		return NumberVal(days360Calc(sy, sm, sd, ey, em, ed, false) / 360), nil

	case 1: // Actual/actual
		actualDays := endSerial - startSerial
		if sy == ey {
			daysInYear := 365.0
			if isLeapYear(sy) {
				daysInYear = 366.0
			}
			return NumberVal(actualDays / daysInYear), nil
		}
		totalYearDays := 0.0
		for y := sy; y <= ey; y++ {
			if isLeapYear(y) {
				totalYearDays += 366
			} else {
				totalYearDays += 365
			}
		}
		avgYear := totalYearDays / float64(ey-sy+1)
		return NumberVal(actualDays / avgYear), nil

	case 2: // Actual/360
		actualDays := endSerial - startSerial
		return NumberVal(actualDays / 360), nil

	case 3: // Actual/365
		actualDays := endSerial - startSerial
		return NumberVal(actualDays / 365), nil

	case 4: // European 30/360
		return NumberVal(days360Calc(sy, sm, sd, ey, em, ed, true) / 360), nil
	}

	return ErrorVal(ErrValNUM), nil
}

// parseWeekendParam interprets the weekend parameter used by WORKDAY.INTL and
// NETWORKDAYS.INTL. It returns a [7]bool where index 0=Sunday, 1=Monday, ...
// 6=Saturday, with true meaning the day is a weekend (non-working) day.
// Returns nil error on success, or a Value error for invalid input.
func parseWeekendParam(arg Value) ([7]bool, *Value) {
	// If the argument is a string, interpret as a 7-character mask (Mon-Sun).
	if arg.Type == ValueString {
		s := arg.Str
		if len(s) != 7 {
			ev := ErrorVal(ErrValVALUE)
			return [7]bool{}, &ev
		}
		// Check all weekend — "1111111" is invalid.
		allOnes := true
		for i := 0; i < 7; i++ {
			if s[i] != '0' && s[i] != '1' {
				ev := ErrorVal(ErrValVALUE)
				return [7]bool{}, &ev
			}
			if s[i] == '0' {
				allOnes = false
			}
		}
		if allOnes {
			ev := ErrorVal(ErrValVALUE)
			return [7]bool{}, &ev
		}
		// String positions: 0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri, 5=Sat, 6=Sun
		// Array indices:    0=Sun, 1=Mon, 2=Tue, 3=Wed, 4=Thu, 5=Fri, 6=Sat
		var result [7]bool
		for i := 0; i < 7; i++ {
			if s[i] == '1' {
				// Map string index i (Mon=0..Sun=6) to Weekday (Sun=0..Sat=6)
				wd := (i + 1) % 7
				result[wd] = true
			}
		}
		return result, nil
	}

	// Numeric weekend codes.
	n, e := CoerceNum(arg)
	if e != nil {
		return [7]bool{}, e
	}
	code := int(n)

	// Codes 1-7: two-day weekends
	// Codes 11-17: single-day weekends
	type weekendDef struct {
		days [7]bool // index 0=Sunday
	}
	defs := map[int]weekendDef{
		1:  {days: [7]bool{true, false, false, false, false, false, true}},  // Sat, Sun
		2:  {days: [7]bool{true, true, false, false, false, false, false}},  // Sun, Mon
		3:  {days: [7]bool{false, true, true, false, false, false, false}},  // Mon, Tue
		4:  {days: [7]bool{false, false, true, true, false, false, false}},  // Tue, Wed
		5:  {days: [7]bool{false, false, false, true, true, false, false}},  // Wed, Thu
		6:  {days: [7]bool{false, false, false, false, true, true, false}},  // Thu, Fri
		7:  {days: [7]bool{false, false, false, false, false, true, true}},  // Fri, Sat
		11: {days: [7]bool{true, false, false, false, false, false, false}}, // Sun only
		12: {days: [7]bool{false, true, false, false, false, false, false}}, // Mon only
		13: {days: [7]bool{false, false, true, false, false, false, false}}, // Tue only
		14: {days: [7]bool{false, false, false, true, false, false, false}}, // Wed only
		15: {days: [7]bool{false, false, false, false, true, false, false}}, // Thu only
		16: {days: [7]bool{false, false, false, false, false, true, false}}, // Fri only
		17: {days: [7]bool{false, false, false, false, false, false, true}}, // Sat only
	}

	def, ok := defs[code]
	if !ok {
		ev := ErrorVal(ErrValVALUE)
		return [7]bool{}, &ev
	}
	return def.days, nil
}

func fnWorkdayIntl(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	startSerial, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if startSerial < 0 || startSerial > MaxExcelSerial {
		return ErrorVal(ErrValNUM), nil
	}
	daysF, e2 := CoerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}
	days := int(daysF)

	// Default weekend: Saturday and Sunday (code 1).
	weekend := [7]bool{true, false, false, false, false, false, true}
	if len(args) >= 3 && args[2].Type != ValueEmpty {
		var ev *Value
		weekend, ev = parseWeekendParam(args[2])
		if ev != nil {
			return *ev, nil
		}
	}

	holidays := make(map[float64]bool)
	if len(args) == 4 {
		var ev *Value
		holidays, ev = parseHolidays(args[3])
		if ev != nil {
			return *ev, nil
		}
	}

	t := ExcelSerialToTime(startSerial)

	if days == 0 {
		return NumberVal(startSerial), nil
	}

	step := 1
	if days < 0 {
		step = -1
		days = -days
	}

	for days > 0 {
		t = t.AddDate(0, 0, step)
		serial := math.Trunc(TimeToExcelSerial(t))
		wd := int(t.Weekday())
		if weekend[wd] {
			continue
		}
		if holidays[serial] {
			continue
		}
		days--
	}

	result := TimeToExcelSerial(t)
	if result < 0 || result > MaxExcelSerial {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(result), nil
}

func fnWORKDAY(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	startSerial, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if startSerial < 0 || startSerial > MaxExcelSerial {
		return ErrorVal(ErrValNUM), nil
	}
	daysF, e2 := CoerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}
	days := int(daysF)

	holidays := make(map[float64]bool)
	if len(args) == 3 {
		var ev *Value
		holidays, ev = parseHolidays(args[2])
		if ev != nil {
			return *ev, nil
		}
	}

	t := ExcelSerialToTime(startSerial)

	if days == 0 {
		return NumberVal(startSerial), nil
	}

	step := 1
	if days < 0 {
		step = -1
		days = -days
	}

	for days > 0 {
		t = t.AddDate(0, 0, step)
		serial := math.Trunc(TimeToExcelSerial(t))
		wd := t.Weekday()
		if wd == time.Saturday || wd == time.Sunday {
			continue
		}
		if holidays[serial] {
			continue
		}
		days--
	}

	result := TimeToExcelSerial(t)
	if result < 0 || result > MaxExcelSerial {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(result), nil
}
