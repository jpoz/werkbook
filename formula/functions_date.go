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
	Register("NOW", NoCtx(fnNOW))
	Register("SECOND", NoCtx(fnSECOND))
	Register("TIME", NoCtx(fnTIME))
	Register("TODAY", NoCtx(fnTODAY))
	Register("WEEKDAY", NoCtx(fnWEEKDAY))
	Register("WEEKNUM", NoCtx(fnWEEKNUM))
	Register("WORKDAY", NoCtx(fnWORKDAY))
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
	duration := t.Sub(ExcelEpoch)
	days := duration.Hours() / 24
	if days >= 60 {
		days++
	}
	return days
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
	t := ExcelSerialToTime(n)
	return NumberVal(float64(t.Day())), nil
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
	t := ExcelSerialToTime(n)
	return NumberVal(float64(t.Month())), nil
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
	t := ExcelSerialToTime(n)
	return NumberVal(float64(t.Year())), nil
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
		if sd == 31 {
			sd = 30
		}
		if ed == 31 && sd >= 30 {
			ed = 30
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
	t := ExcelSerialToTime(n)
	return NumberVal(float64(t.Hour())), nil
}

func fnMINUTE(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	t := ExcelSerialToTime(n)
	return NumberVal(float64(t.Minute())), nil
}

func fnSECOND(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	t := ExcelSerialToTime(n)
	return NumberVal(float64(t.Second())), nil
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
	return NumberVal((hour*3600 + minute*60 + second) / 86400), nil
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
			return NumberVal(math.Floor(TimeToExcelSerial(t))), nil
		}
	}
	return ErrorVal(ErrValVALUE), nil
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
	weekNum := (dayOfYear + offset - 1) / 7 + 1
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
