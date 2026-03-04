package ooxml

import "testing"

func TestParseCellData_NumericSharedString(t *testing.T) {
	sst := []string{"-8086931554011838357", "hello", "3.14", "1e10"}

	tests := []struct {
		name     string
		cell     xlsxC
		wantType string
		wantVal  string
	}{
		{
			name:     "large negative integer SST is numeric",
			cell:     xlsxC{R: "A1", T: "s", V: "0"},
			wantType: "",
			wantVal:  "-8086931554011838357",
		},
		{
			name:     "text SST stays string",
			cell:     xlsxC{R: "A2", T: "s", V: "1"},
			wantType: "s",
			wantVal:  "hello",
		},
		{
			name:     "decimal SST is numeric",
			cell:     xlsxC{R: "A3", T: "s", V: "2"},
			wantType: "",
			wantVal:  "3.14",
		},
		{
			name:     "scientific notation SST is numeric",
			cell:     xlsxC{R: "A4", T: "s", V: "3"},
			wantType: "",
			wantVal:  "1e10",
		},
		{
			name:     "plain number cell stays numeric",
			cell:     xlsxC{R: "A5", V: "42"},
			wantType: "",
			wantVal:  "42",
		},
		{
			name:     "bool cell stays bool",
			cell:     xlsxC{R: "A6", T: "b", V: "1"},
			wantType: "b",
			wantVal:  "1",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cd := parseCellData(tt.cell, sst)
			if cd.Type != tt.wantType {
				t.Errorf("Type = %q, want %q", cd.Type, tt.wantType)
			}
			if cd.Value != tt.wantVal {
				t.Errorf("Value = %q, want %q", cd.Value, tt.wantVal)
			}
		})
	}
}
