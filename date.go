package werkbook

import "time"

// excelEpoch is "January 0, 1900" = December 31, 1899.
// Excel serial 1 = January 1, 1900 = epoch + 1 day.
var excelEpoch = time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)

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
