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
	t := time.Date(int(year), time.Month(int(month)), int(day), 0, 0, 0, 0, time.UTC)
	return NumberVal(timeToExcelSerial(t)), nil
}

func fnDAY(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
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

func fnYEAR(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := coerceNum(args[0])
	if e != nil {
		return *e, nil
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
