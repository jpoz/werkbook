package werkbook

import (
	"math"
	"testing"
	"time"

	"github.com/jpoz/werkbook/ooxml"
)

func TestToValue(t *testing.T) {
	tests := []struct {
		name    string
		input   any
		want    Value
		wantErr bool
	}{
		{"nil", nil, Value{Type: TypeEmpty}, false},
		{"string", "hello", Value{Type: TypeString, String: "hello"}, false},
		{"empty string", "", Value{Type: TypeString, String: ""}, false},
		{"bool true", true, Value{Type: TypeBool, Bool: true}, false},
		{"bool false", false, Value{Type: TypeBool, Bool: false}, false},
		{"int", 42, Value{Type: TypeNumber, Number: 42}, false},
		{"int64", int64(100), Value{Type: TypeNumber, Number: 100}, false},
		{"float64", 3.14, Value{Type: TypeNumber, Number: 3.14}, false},
		{"float32", float32(1.5), Value{Type: TypeNumber, Number: 1.5}, false},
		{"uint", uint(7), Value{Type: TypeNumber, Number: 7}, false},
		{"NaN", math.NaN(), Value{}, true},
		{"Inf", math.Inf(1), Value{}, true},
		{"unsupported", []int{1}, Value{}, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := toValue(tt.input)
			if (err != nil) != tt.wantErr {
				t.Fatalf("toValue(%v) err = %v, wantErr %v", tt.input, err, tt.wantErr)
			}
			if !tt.wantErr && got != tt.want {
				t.Errorf("toValue(%v) = %#v, want %#v", tt.input, got, tt.want)
			}
		})
	}
}

func TestCellDataToValue_SharedStringNumeric(t *testing.T) {
	// Styles slice: index 0 = default, index 1 = text format (@).
	textStyle := &Style{NumFmtID: 49}
	styles := []*Style{nil, textStyle}

	tests := []struct {
		name     string
		cd       ooxml.CellData
		styles   []*Style
		wantType ValueType
		wantNum  float64
		wantStr  string
	}{
		{
			name:     "shared string numeric stays string",
			cd:       ooxml.CellData{Type: "s", Value: "-8086931554011838357"},
			wantType: TypeString,
			wantStr:  "-8086931554011838357",
		},
		{
			name:     "shared string positive integer stays string",
			cd:       ooxml.CellData{Type: "s", Value: "42"},
			wantType: TypeString,
			wantStr:  "42",
		},
		{
			name:     "shared string float stays string",
			cd:       ooxml.CellData{Type: "s", Value: "3.14"},
			wantType: TypeString,
			wantStr:  "3.14",
		},
		{
			name:     "shared string with non-numeric text",
			cd:       ooxml.CellData{Type: "s", Value: "hello"},
			wantType: TypeString,
			wantStr:  "hello",
		},
		{
			name:     "shared string numeric with text format stays string",
			cd:       ooxml.CellData{Type: "s", Value: "42", StyleIdx: 1},
			styles:   styles,
			wantType: TypeString,
			wantStr:  "42",
		},
		{
			name:     "shared string float with text format stays string",
			cd:       ooxml.CellData{Type: "s", Value: "3.14", StyleIdx: 1},
			styles:   styles,
			wantType: TypeString,
			wantStr:  "3.14",
		},
		{
			name:     "str type stays string even if numeric",
			cd:       ooxml.CellData{Type: "str", Value: "42"},
			wantType: TypeString,
			wantStr:  "42",
		},
		{
			name:     "inlineStr type stays string even if numeric",
			cd:       ooxml.CellData{Type: "inlineStr", Value: "100"},
			wantType: TypeString,
			wantStr:  "100",
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := cellDataToValue(tt.cd, tt.styles, false)
			if got.Type != tt.wantType {
				t.Fatalf("cellDataToValue(%+v).Type = %v, want %v", tt.cd, got.Type, tt.wantType)
			}
			if tt.wantType == TypeNumber && got.Number != tt.wantNum {
				t.Errorf("cellDataToValue(%+v).Number = %v, want %v", tt.cd, got.Number, tt.wantNum)
			}
			if tt.wantType == TypeString && got.String != tt.wantStr {
				t.Errorf("cellDataToValue(%+v).String = %q, want %q", tt.cd, got.String, tt.wantStr)
			}
		})
	}
}

func TestCellDataToValue_DateCell(t *testing.T) {
	tests := []struct {
		name     string
		cd       ooxml.CellData
		date1904 bool
		want     float64
	}{
		{
			name:     "1900 date system",
			cd:       ooxml.CellData{Type: "d", Value: "2024-06-15"},
			date1904: false,
			want:     timeToExcelSerial(time.Date(2024, 6, 15, 0, 0, 0, 0, time.UTC)),
		},
		{
			name:     "1904 date system",
			cd:       ooxml.CellData{Type: "d", Value: "1904-01-01T00:00:00Z"},
			date1904: true,
			want:     0,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := cellDataToValue(tt.cd, nil, tt.date1904)
			if got.Type != TypeNumber {
				t.Fatalf("cellDataToValue(%+v).Type = %v, want %v", tt.cd, got.Type, TypeNumber)
			}
			if got.Number != tt.want {
				t.Fatalf("cellDataToValue(%+v).Number = %v, want %v", tt.cd, got.Number, tt.want)
			}
		})
	}
}
