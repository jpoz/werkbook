package formula

import (
	"math"
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
