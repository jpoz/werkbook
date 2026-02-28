package formula

import (
	"math"
	"strings"
	"time"
)

// Serial date helpers — duplicated from werkbook/date.go to avoid circular imports.
var excelEpoch = time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)

func timeToExcelSerial(t time.Time) float64 {
	duration := t.Sub(excelEpoch)
	days := duration.Hours() / 24
	if days >= 60 {
		days++
	}
	return days
}

func excelSerialToTime(serial float64) time.Time {
	if serial > 60 {
		serial--
	}
	days := int(serial)
	frac := serial - float64(days)
	t := excelEpoch.AddDate(0, 0, days)
	t = t.Add(time.Duration(frac * 24 * float64(time.Hour)))
	return t
}

func fnDATE(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	year, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	month, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	day, e := coerceNum(args[2])
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

	t := time.Date(y, time.Month(m), d, 0, 0, 0, 0, time.UTC)
	serial := timeToExcelSerial(t)
	if serial < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(serial), nil
}

func fnDATEDIF(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	startSerial, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	endSerial, e2 := coerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}
	unitStr := strings.ToUpper(valueToString(args[2]))

	if startSerial > endSerial {
		return ErrorVal(ErrValNUM), nil
	}

	start := excelSerialToTime(startSerial)
	end := excelSerialToTime(endSerial)

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
			// Get days in the previous month
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
		// Set start to same year as end
		startInEndYear := time.Date(end.Year(), start.Month(), start.Day(), 0, 0, 0, 0, time.UTC)
		days := int(end.Sub(startInEndYear).Hours() / 24)
		if days < 0 {
			startInEndYear = time.Date(end.Year()-1, start.Month(), start.Day(), 0, 0, 0, 0, time.UTC)
			days = int(end.Sub(startInEndYear).Hours() / 24)
		}
		return NumberVal(float64(days)), nil
	default:
		return ErrorVal(ErrValNUM), nil
	}
}

func fnDAY(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	t := excelSerialToTime(n)
	return NumberVal(float64(t.Day())), nil
}

func fnDAYS(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	end, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	start, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	return NumberVal(math.Trunc(end) - math.Trunc(start)), nil
}

func fnHOUR(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	t := excelSerialToTime(n)
	return NumberVal(float64(t.Hour())), nil
}

func fnMINUTE(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	t := excelSerialToTime(n)
	return NumberVal(float64(t.Minute())), nil
}

func fnMONTH(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	t := excelSerialToTime(n)
	return NumberVal(float64(t.Month())), nil
}

func fnNOW(args []Value) (Value, error) {
	if len(args) != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return NumberVal(timeToExcelSerial(time.Now())), nil
}

func fnSECOND(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	t := excelSerialToTime(n)
	return NumberVal(float64(t.Second())), nil
}

func fnTIME(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	hour, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	minute, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	second, e := coerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	return NumberVal((hour*3600 + minute*60 + second) / 86400), nil
}

func fnTODAY(args []Value) (Value, error) {
	if len(args) != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	now := time.Now()
	today := time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, time.UTC)
	return NumberVal(math.Floor(timeToExcelSerial(today))), nil
}

func fnWEEKDAY(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	serial, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	returnType := 1.0
	if len(args) == 2 {
		returnType, e = coerceNum(args[1])
		if e != nil {
			return *e, nil
		}
	}

	t := excelSerialToTime(serial)
	wd := int(t.Weekday()) // 0=Sunday, 1=Monday, ..., 6=Saturday

	rt := int(returnType)
	var result int
	switch rt {
	case 1, 17:
		// Sunday=1 through Saturday=7
		result = wd + 1
	case 2, 11:
		// Monday=1 through Sunday=7
		result = (wd+6)%7 + 1
	case 3:
		// Monday=0 through Sunday=6
		result = (wd + 6) % 7
	case 12:
		// Tuesday=1 through Monday=7
		result = (wd-2+7)%7 + 1
	case 13:
		// Wednesday=1 through Tuesday=7
		result = (wd-3+7)%7 + 1
	case 14:
		// Thursday=1 through Wednesday=7
		result = (wd-4+7)%7 + 1
	case 15:
		// Friday=1 through Thursday=7
		result = (wd-5+7)%7 + 1
	case 16:
		// Saturday=1 through Friday=7
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
	serial, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	t := excelSerialToTime(serial)
	_, week := t.ISOWeek()
	return NumberVal(float64(week)), nil
}

func fnYEAR(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if n < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	t := excelSerialToTime(n)
	return NumberVal(float64(t.Year())), nil
}

// DATEVALUE(date_text) — converts a date string to a serial number.
func fnDATEVALUE(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	text := strings.TrimSpace(valueToString(args[0]))

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
		t, err := time.Parse(layout, text)
		if err == nil {
			return NumberVal(math.Floor(timeToExcelSerial(t))), nil
		}
	}
	return ErrorVal(ErrValVALUE), nil
}

// EDATE(start_date, months) — returns date serial shifted by N months.
func fnEDATE(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	serial, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	months, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	t := excelSerialToTime(serial)
	y, mo, d := t.Date()
	m := int(math.Trunc(months))
	// Compute target year/month by adding months to the month component
	targetMonth := time.Month(int(mo) + m)
	targetYear := y
	// Normalize month (Go's time.Date handles this, but we need the actual month for clamping)
	for targetMonth > 12 {
		targetMonth -= 12
		targetYear++
	}
	for targetMonth < 1 {
		targetMonth += 12
		targetYear--
	}
	// Try original day; if it overflows, clamp to last day of target month
	result := time.Date(targetYear, targetMonth, d, 0, 0, 0, 0, time.UTC)
	if result.Month() != targetMonth {
		// Day overflowed; use last day of target month
		result = time.Date(targetYear, targetMonth+1, 0, 0, 0, 0, 0, time.UTC)
	}
	return NumberVal(math.Floor(timeToExcelSerial(result))), nil
}

// EOMONTH(start_date, months) — returns serial for last day of month N months away.
func fnEOMONTH(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	serial, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	months, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	t := excelSerialToTime(serial)
	y, m, _ := t.Date()
	// Day 0 of month X is the last day of month X-1 in Go.
	// So day 0 of (m + months + 1) is last day of (m + months).
	last := time.Date(y, m+time.Month(int(math.Trunc(months))+1), 0, 0, 0, 0, 0, time.UTC)
	return NumberVal(math.Floor(timeToExcelSerial(last))), nil
}

// NETWORKDAYS(start_date, end_date, [holidays]) — count working days between two dates.
func fnNETWORKDAYS(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	startSerial, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	endSerial, e := coerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	startSerial = math.Trunc(startSerial)
	endSerial = math.Trunc(endSerial)

	// Build holiday set
	holidays := make(map[float64]bool)
	if len(args) == 3 {
		arg := args[2]
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return cell, nil
					}
					n, ce := coerceNum(cell)
					if ce != nil {
						return *ce, nil
					}
					holidays[math.Trunc(n)] = true
				}
			}
		} else if arg.Type != ValueEmpty {
			n, ce := coerceNum(arg)
			if ce != nil {
				return *ce, nil
			}
			holidays[math.Trunc(n)] = true
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
		t := excelSerialToTime(d)
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

// WEEKNUM(serial_number, [return_type]) — returns the week number of a date.
func fnWEEKNUM(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	serial, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}

	returnType := 1.0
	if len(args) == 2 {
		returnType, e = coerceNum(args[1])
		if e != nil {
			return *e, nil
		}
	}

	rt := int(returnType)
	t := excelSerialToTime(serial)

	// return_type 21 = ISO 8601 week number
	if rt == 21 {
		_, isoWeek := t.ISOWeek()
		return NumberVal(float64(isoWeek)), nil
	}

	// Determine which weekday starts the week based on return_type.
	// The start day is expressed as time.Weekday (0=Sunday..6=Saturday).
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

	// Find Jan 1 of the date's year.
	jan1 := time.Date(t.Year(), 1, 1, 0, 0, 0, 0, time.UTC)
	jan1Wd := jan1.Weekday()

	// Offset = how many days Jan 1 is past the week start day.
	// This tells us how far into the first partial week Jan 1 falls.
	offset := int(jan1Wd-weekStart+7) % 7

	// Day of year (1-based).
	dayOfYear := t.YearDay()

	// Week number: Jan 1 is always in week 1.
	weekNum := (dayOfYear + offset - 1) / 7 + 1
	return NumberVal(float64(weekNum)), nil
}

// YEARFRAC(start_date, end_date, [basis]) — calculates the fraction of the year
// represented by the number of whole days between two dates.
func fnYEARFRAC(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	startSerial, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	endSerial, e2 := coerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}
	basis := 0
	if len(args) == 3 {
		b, e3 := coerceNum(args[2])
		if e3 != nil {
			return *e3, nil
		}
		basis = int(b)
	}
	if basis < 0 || basis > 4 {
		return ErrorVal(ErrValNUM), nil
	}

	// Swap if start > end
	if startSerial > endSerial {
		startSerial, endSerial = endSerial, startSerial
	}

	start := excelSerialToTime(startSerial)
	end := excelSerialToTime(endSerial)

	sy, sm, sd := start.Year(), int(start.Month()), start.Day()
	ey, em, ed := end.Year(), int(end.Month()), end.Day()

	switch basis {
	case 0: // US (NASD) 30/360
		if sd == 31 {
			sd = 30
		}
		if ed == 31 && sd >= 30 {
			ed = 30
		}
		num := float64((ey-sy)*360 + (em-sm)*30 + (ed - sd))
		return NumberVal(num / 360), nil

	case 1: // Actual/actual
		actualDays := endSerial - startSerial
		// Average year length across the span
		if sy == ey {
			daysInYear := 365.0
			if isLeapYear(sy) {
				daysInYear = 366.0
			}
			return NumberVal(actualDays / daysInYear), nil
		}
		// Multi-year span: compute average days per year
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
		if sd == 31 {
			sd = 30
		}
		if ed == 31 {
			ed = 30
		}
		num := float64((ey-sy)*360 + (em-sm)*30 + (ed - sd))
		return NumberVal(num / 360), nil
	}

	return ErrorVal(ErrValNUM), nil
}

func isLeapYear(year int) bool {
	return year%4 == 0 && (year%100 != 0 || year%400 == 0)
}

// WORKDAY(start_date, days, [holidays]) — returns the serial number of the date
// that is the indicated number of working days before or after a start date.
// Working days exclude weekends (Saturday and Sunday) and specified holidays.
func fnWORKDAY(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	startSerial, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if startSerial < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	daysF, e2 := coerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}
	days := int(daysF)

	// Collect holidays if provided
	holidays := make(map[int]bool)
	if len(args) == 3 {
		arg := args[2]
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return cell, nil
					}
					if cell.Type == ValueNumber {
						holidays[int(cell.Num)] = true
					}
				}
			}
		} else if arg.Type == ValueNumber {
			holidays[int(arg.Num)] = true
		} else if arg.Type != ValueEmpty {
			n, ce := coerceNum(arg)
			if ce != nil {
				return *ce, nil
			}
			holidays[int(n)] = true
		}
	}

	t := excelSerialToTime(startSerial)

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
		serial := timeToExcelSerial(t)
		wd := t.Weekday()
		if wd == time.Saturday || wd == time.Sunday {
			continue
		}
		if holidays[int(serial)] {
			continue
		}
		days--
	}

	result := timeToExcelSerial(t)
	if result < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(result), nil
}
