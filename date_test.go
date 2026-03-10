package werkbook

import (
	"math"
	"testing"
	"time"
)

func TestTimeToSerial(t *testing.T) {
	tests := []struct {
		name string
		t    time.Time
		want float64
	}{
		{"Jan 1 1900", time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC), 1},
		{"Feb 28 1900", time.Date(1900, 2, 28, 0, 0, 0, 0, time.UTC), 59},
		// 1900 leap year bug: Feb 29, 1900 does not exist but serial 60 is reserved for it.
		// March 1, 1900 should be serial 61.
		{"Mar 1 1900", time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC), 61},
		{"Jan 1 2000", time.Date(2000, 1, 1, 0, 0, 0, 0, time.UTC), 36526},
		{"Jan 1 2024", time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC), 45292},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := timeToSerial(tt.t)
			if math.Abs(got-tt.want) > 0.0001 {
				t.Errorf("timeToSerial(%v) = %f, want %f", tt.t, got, tt.want)
			}
		})
	}
}

func TestSerialToTime(t *testing.T) {
	tests := []struct {
		name   string
		serial float64
		want   time.Time
	}{
		{"serial 1", 1, time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC)},
		{"serial 59", 59, time.Date(1900, 2, 28, 0, 0, 0, 0, time.UTC)},
		{"serial 61", 61, time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC)},
		{"serial 36526", 36526, time.Date(2000, 1, 1, 0, 0, 0, 0, time.UTC)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := serialToTime(tt.serial)
			if !got.Equal(tt.want) {
				t.Errorf("serialToTime(%f) = %v, want %v", tt.serial, got, tt.want)
			}
		})
	}
}

func TestSerialToTimeExported(t *testing.T) {
	got := SerialToTime(36526)
	want := time.Date(2000, 1, 1, 0, 0, 0, 0, time.UTC)
	if !got.Equal(want) {
		t.Errorf("SerialToTime(36526) = %v, want %v", got, want)
	}
}

func TestIsDateFormat(t *testing.T) {
	tests := []struct {
		name     string
		numFmt   string
		numFmtID int
		want     bool
	}{
		{"built-in 14", "", 14, true},
		{"built-in 22", "", 22, true},
		{"built-in 0", "", 0, false},
		{"built-in 1", "", 1, false},
		{"custom date", "yyyy-mm-dd", 0, true},
		{"custom time", "h:mm:ss", 0, true},
		{"custom datetime", "yyyy-mm-dd h:mm", 0, true},
		{"number format", "#,##0.00", 0, false},
		{"percentage", "0.00%", 0, false},
		{"empty", "", 0, false},
		{"custom with quotes", `"Date: "yyyy-mm-dd`, 0, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := IsDateFormat(tt.numFmt, tt.numFmtID)
			if got != tt.want {
				t.Errorf("IsDateFormat(%q, %d) = %v, want %v", tt.numFmt, tt.numFmtID, got, tt.want)
			}
		})
	}
}

func TestSerialToTime1904(t *testing.T) {
	tests := []struct {
		name   string
		serial float64
		want   time.Time
	}{
		// In the 1904 system, serial 0 = Jan 1, 1904.
		{"serial 0", 0, time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)},
		{"serial 1", 1, time.Date(1904, 1, 2, 0, 0, 0, 0, time.UTC)},
		// No leap year bug in 1904 system.
		{"serial 59", 59, time.Date(1904, 2, 29, 0, 0, 0, 0, time.UTC)},
		{"serial 60", 60, time.Date(1904, 3, 1, 0, 0, 0, 0, time.UTC)},
		// Verified: serial 17816 in 1904 = Oct 11, 1952.
		{"serial 17816", 17816, time.Date(1952, 10, 11, 0, 0, 0, 0, time.UTC)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := serialToTime1904(tt.serial)
			if !got.Equal(tt.want) {
				t.Errorf("serialToTime1904(%f) = %v, want %v", tt.serial, got, tt.want)
			}
		})
	}
}

func TestDateRoundTrip(t *testing.T) {
	dates := []time.Time{
		time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC),
		time.Date(2024, 6, 15, 0, 0, 0, 0, time.UTC),
		time.Date(2000, 12, 31, 0, 0, 0, 0, time.UTC),
	}
	for _, d := range dates {
		serial := timeToSerial(d)
		got := serialToTime(serial)
		if !got.Equal(d) {
			t.Errorf("round-trip failed: %v -> %f -> %v", d, serial, got)
		}
	}
}
