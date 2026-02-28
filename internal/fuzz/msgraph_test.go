package fuzz

import (
	"encoding/json"
	"testing"
	"time"
)

func TestParseUsedRangeResponse(t *testing.T) {
	body := []byte(`{
		"address": "Sheet1!A1:C3",
		"columnIndex": 0,
		"rowIndex": 0,
		"values": [
			[10, "hello", true],
			[20, "", false],
			[30, "#VALUE!", null]
		]
	}`)

	result, err := parseUsedRangeResponse(body)
	if err != nil {
		t.Fatalf("parseUsedRangeResponse: %v", err)
	}

	tests := []struct {
		ref  string
		want string
	}{
		{"A1", "10"},
		{"B1", "hello"},
		{"C1", "TRUE"},
		{"A2", "20"},
		// B2 is empty string, should be skipped
		{"C2", "FALSE"},
		{"A3", "30"},
		{"B3", "#VALUE!"},
		// C3 is null, should be skipped
	}

	for _, tt := range tests {
		got, ok := result[tt.ref]
		if !ok {
			t.Errorf("missing ref %s", tt.ref)
			continue
		}
		if got != tt.want {
			t.Errorf("ref %s: got %q, want %q", tt.ref, got, tt.want)
		}
	}

	// Verify empty/null cells are not present.
	if _, ok := result["B2"]; ok {
		t.Error("B2 (empty string) should not be in result")
	}
	if _, ok := result["C3"]; ok {
		t.Error("C3 (null) should not be in result")
	}
}

func TestParseUsedRangeResponseOffset(t *testing.T) {
	body := []byte(`{
		"address": "Sheet1!B2:C3",
		"columnIndex": 1,
		"rowIndex": 1,
		"values": [
			[100, 200],
			[300, 400]
		]
	}`)

	result, err := parseUsedRangeResponse(body)
	if err != nil {
		t.Fatalf("parseUsedRangeResponse: %v", err)
	}

	tests := []struct {
		ref  string
		want string
	}{
		{"B2", "100"},
		{"C2", "200"},
		{"B3", "300"},
		{"C3", "400"},
	}

	for _, tt := range tests {
		got, ok := result[tt.ref]
		if !ok {
			t.Errorf("missing ref %s", tt.ref)
			continue
		}
		if got != tt.want {
			t.Errorf("ref %s: got %q, want %q", tt.ref, got, tt.want)
		}
	}
}

func TestParseRangeStart(t *testing.T) {
	tests := []struct {
		address  string
		wantCol  int
		wantRow  int
	}{
		{"Sheet1!A1:D5", 1, 1},
		{"B2:C3", 2, 2},
		{"AA10:AB20", 27, 10},
		{"A1", 1, 1},
		{"Sheet1!Z100:AA200", 26, 100},
	}

	for _, tt := range tests {
		col, row := parseRangeStart(tt.address)
		if col != tt.wantCol || row != tt.wantRow {
			t.Errorf("parseRangeStart(%q): got (%d, %d), want (%d, %d)",
				tt.address, col, row, tt.wantCol, tt.wantRow)
		}
	}
}

func TestFormatGraphValue(t *testing.T) {
	tests := []struct {
		input any
		want  string
	}{
		{nil, ""},
		{"hello", "hello"},
		{float64(10), "10"},
		{float64(3.14), "3.14"},
		{float64(0), "0"},
		{float64(-5), "-5"},
		{float64(1.5e10), "15000000000"},
		{true, "TRUE"},
		{false, "FALSE"},
		{"#VALUE!", "#VALUE!"},
		{"#REF!", "#REF!"},
	}

	for _, tt := range tests {
		got := formatGraphValue(tt.input)
		if got != tt.want {
			t.Errorf("formatGraphValue(%v): got %q, want %q", tt.input, got, tt.want)
		}
	}
}

func TestGuessExcelType(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"", "empty"},
		{"TRUE", "bool"},
		{"FALSE", "bool"},
		{"true", "bool"},
		{"#VALUE!", "error"},
		{"#REF!", "error"},
		{"#N/A", "error"},
		{"42", "number"},
		{"hello", "number"},
	}

	for _, tt := range tests {
		got := guessExcelType(tt.input)
		if got != tt.want {
			t.Errorf("guessExcelType(%q): got %q, want %q", tt.input, got, tt.want)
		}
	}
}

func TestTokenCacheSerialization(t *testing.T) {
	tc := &tokenCache{
		AccessToken:  "access-token-123",
		RefreshToken: "refresh-token-456",
		ExpiresAt:    time.Date(2025, 6, 15, 12, 0, 0, 0, time.UTC),
	}

	data, err := json.MarshalIndent(tc, "", "  ")
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var tc2 tokenCache
	if err := json.Unmarshal(data, &tc2); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	if tc2.AccessToken != tc.AccessToken {
		t.Errorf("access token: got %q, want %q", tc2.AccessToken, tc.AccessToken)
	}
	if tc2.RefreshToken != tc.RefreshToken {
		t.Errorf("refresh token: got %q, want %q", tc2.RefreshToken, tc.RefreshToken)
	}
	if !tc2.ExpiresAt.Equal(tc.ExpiresAt) {
		t.Errorf("expires at: got %v, want %v", tc2.ExpiresAt, tc.ExpiresAt)
	}
}

func TestTruncateBytes(t *testing.T) {
	tests := []struct {
		input  string
		maxLen int
		want   string
	}{
		{"hello", 10, "hello"},
		{"hello world", 5, "hello..."},
		{"", 5, ""},
	}

	for _, tt := range tests {
		got := truncateBytes([]byte(tt.input), tt.maxLen)
		if got != tt.want {
			t.Errorf("truncateBytes(%q, %d): got %q, want %q", tt.input, tt.maxLen, got, tt.want)
		}
	}
}
