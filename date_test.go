package werkbook

import (
	"math"
	"testing"
	"time"
)

func TestTimeToExcelSerial(t *testing.T) {
	tests := []struct {
		name string
		t    time.Time
		want float64
	}{
		{"Jan 1 1900", time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC), 1},
		{"Feb 28 1900", time.Date(1900, 2, 28, 0, 0, 0, 0, time.UTC), 59},
		// Excel has a bug where it thinks Feb 29, 1900 exists.
		// March 1, 1900 should be serial 61 in Excel.
		{"Mar 1 1900", time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC), 61},
		{"Jan 1 2000", time.Date(2000, 1, 1, 0, 0, 0, 0, time.UTC), 36526},
		{"Jan 1 2024", time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC), 45292},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := timeToExcelSerial(tt.t)
			if math.Abs(got-tt.want) > 0.0001 {
				t.Errorf("timeToExcelSerial(%v) = %f, want %f", tt.t, got, tt.want)
			}
		})
	}
}

func TestExcelSerialToTime(t *testing.T) {
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
			got := excelSerialToTime(tt.serial)
			if !got.Equal(tt.want) {
				t.Errorf("excelSerialToTime(%f) = %v, want %v", tt.serial, got, tt.want)
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
		serial := timeToExcelSerial(d)
		got := excelSerialToTime(serial)
		if !got.Equal(d) {
			t.Errorf("round-trip failed: %v -> %f -> %v", d, serial, got)
		}
	}
}
