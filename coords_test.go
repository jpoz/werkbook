package werkbook

import "testing"

func TestColumnNameToNumber(t *testing.T) {
	tests := []struct {
		name string
		want int
	}{
		{"A", 1},
		{"B", 2},
		{"Z", 26},
		{"AA", 27},
		{"AZ", 52},
		{"BA", 53},
		{"ZZ", 702},
		{"AAA", 703},
		{"XFD", 16384},
		// Case insensitive
		{"a", 1},
		{"xfd", 16384},
		{"Ab", 28},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := ColumnNameToNumber(tt.name)
			if err != nil {
				t.Fatalf("ColumnNameToNumber(%q): %v", tt.name, err)
			}
			if got != tt.want {
				t.Errorf("ColumnNameToNumber(%q) = %d, want %d", tt.name, got, tt.want)
			}
		})
	}
}

func TestColumnNumberToName(t *testing.T) {
	tests := []struct {
		col  int
		want string
	}{
		{1, "A"},
		{2, "B"},
		{26, "Z"},
		{27, "AA"},
		{52, "AZ"},
		{53, "BA"},
		{702, "ZZ"},
		{703, "AAA"},
		{16384, "XFD"},
	}
	for _, tt := range tests {
		t.Run(tt.want, func(t *testing.T) {
			got := ColumnNumberToName(tt.col)
			if got != tt.want {
				t.Errorf("ColumnNumberToName(%d) = %q, want %q", tt.col, got, tt.want)
			}
		})
	}
}

func TestCellNameToCoordinates(t *testing.T) {
	tests := []struct {
		cell    string
		col     int
		row     int
		wantErr bool
	}{
		{"A1", 1, 1, false},
		{"B3", 2, 3, false},
		{"Z100", 26, 100, false},
		{"AA1", 27, 1, false},
		{"XFD1048576", 16384, 1048576, false},
		// Case insensitive
		{"a1", 1, 1, false},
		{"xfd1", 16384, 1, false},
		// Errors
		{"", 0, 0, true},
		{"1", 0, 0, true},
		{"A", 0, 0, true},
		{"A0", 0, 0, true},
		{"AAAA1", 0, 0, true}, // exceeds max column
	}
	for _, tt := range tests {
		t.Run(tt.cell, func(t *testing.T) {
			col, row, err := CellNameToCoordinates(tt.cell)
			if (err != nil) != tt.wantErr {
				t.Fatalf("CellNameToCoordinates(%q) err = %v, wantErr %v", tt.cell, err, tt.wantErr)
			}
			if err == nil {
				if col != tt.col || row != tt.row {
					t.Errorf("CellNameToCoordinates(%q) = (%d, %d), want (%d, %d)", tt.cell, col, row, tt.col, tt.row)
				}
			}
		})
	}
}

func TestCoordinatesToCellName(t *testing.T) {
	tests := []struct {
		col     int
		row     int
		want    string
		wantErr bool
	}{
		{1, 1, "A1", false},
		{2, 3, "B3", false},
		{16384, 1048576, "XFD1048576", false},
		{0, 1, "", true},
		{1, 0, "", true},
		{16385, 1, "", true},
		{1, 1048577, "", true},
	}
	for _, tt := range tests {
		t.Run(tt.want, func(t *testing.T) {
			got, err := CoordinatesToCellName(tt.col, tt.row)
			if (err != nil) != tt.wantErr {
				t.Fatalf("CoordinatesToCellName(%d, %d) err = %v, wantErr %v", tt.col, tt.row, err, tt.wantErr)
			}
			if got != tt.want {
				t.Errorf("CoordinatesToCellName(%d, %d) = %q, want %q", tt.col, tt.row, got, tt.want)
			}
		})
	}
}

func TestRangeToCoordinates(t *testing.T) {
	tests := []struct {
		ref     string
		col1    int
		row1    int
		col2    int
		row2    int
		wantErr bool
	}{
		{"A1:C5", 1, 1, 3, 5, false},
		{"B2:B2", 2, 2, 2, 2, false},
		// Single cell (no colon).
		{"D10", 4, 10, 4, 10, false},
		// Reversed range should normalize.
		{"C5:A1", 1, 1, 3, 5, false},
		{"Z10:A1", 1, 1, 26, 10, false},
		// Errors.
		{"", 0, 0, 0, 0, true},
		{"A1:", 0, 0, 0, 0, true},
		{":B2", 0, 0, 0, 0, true},
		{"!!!:B2", 0, 0, 0, 0, true},
		{"A1:!!!", 0, 0, 0, 0, true},
	}
	for _, tt := range tests {
		t.Run(tt.ref, func(t *testing.T) {
			col1, row1, col2, row2, err := RangeToCoordinates(tt.ref)
			if (err != nil) != tt.wantErr {
				t.Fatalf("RangeToCoordinates(%q) err = %v, wantErr %v", tt.ref, err, tt.wantErr)
			}
			if err == nil {
				if col1 != tt.col1 || row1 != tt.row1 || col2 != tt.col2 || row2 != tt.row2 {
					t.Errorf("RangeToCoordinates(%q) = (%d,%d,%d,%d), want (%d,%d,%d,%d)",
						tt.ref, col1, row1, col2, row2, tt.col1, tt.row1, tt.col2, tt.row2)
				}
			}
		})
	}
}

func TestRoundTripColumnNumber(t *testing.T) {
	for i := 1; i <= MaxColumns; i++ {
		name := ColumnNumberToName(i)
		got, err := ColumnNameToNumber(name)
		if err != nil {
			t.Fatalf("ColumnNameToNumber(%q): %v", name, err)
		}
		if got != i {
			t.Fatalf("round-trip failed: %d -> %q -> %d", i, name, got)
		}
	}
}
