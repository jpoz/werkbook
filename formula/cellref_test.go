package formula

import "testing"

func TestParseCellRefToken(t *testing.T) {
	tests := []struct {
		input string
		want  CellRef
	}{
		// Bare references.
		{"A1", CellRef{Col: 1, Row: 1}},
		{"B2", CellRef{Col: 2, Row: 2}},
		{"Z9", CellRef{Col: 26, Row: 9}},
		{"AA100", CellRef{Col: 27, Row: 100}},
		{"XFD1048576", CellRef{Col: 16384, Row: 1048576}},

		// Absolute references.
		{"$A$1", CellRef{Col: 1, Row: 1, AbsCol: true, AbsRow: true}},
		{"$A1", CellRef{Col: 1, Row: 1, AbsCol: true}},
		{"A$1", CellRef{Col: 1, Row: 1, AbsRow: true}},
		{"$XFD$1048576", CellRef{Col: 16384, Row: 1048576, AbsCol: true, AbsRow: true}},

		// Unquoted sheet.
		{"Sheet1!A1", CellRef{Sheet: "Sheet1", Col: 1, Row: 1}},
		{"Sheet1!$A$1", CellRef{Sheet: "Sheet1", Col: 1, Row: 1, AbsCol: true, AbsRow: true}},
		{"Data!C10", CellRef{Sheet: "Data", Col: 3, Row: 10}},
		{"Sheet2!$B5", CellRef{Sheet: "Sheet2", Col: 2, Row: 5, AbsCol: true}},

		// Quoted sheet.
		{"'Sheet Name'!A1", CellRef{Sheet: "Sheet Name", Col: 1, Row: 1}},
		{"'Sheet Name'!$A$1", CellRef{Sheet: "Sheet Name", Col: 1, Row: 1, AbsCol: true, AbsRow: true}},
		{"'Q1 Data'!B5", CellRef{Sheet: "Q1 Data", Col: 2, Row: 5}},

		// Quoted sheet with escaped quotes.
		{"'It''s a sheet'!B2", CellRef{Sheet: "It's a sheet", Col: 2, Row: 2}},
		{"'Tom''s Q1'!$C$3", CellRef{Sheet: "Tom's Q1", Col: 3, Row: 3, AbsCol: true, AbsRow: true}},

		// Case insensitive column letters.
		{"a1", CellRef{Col: 1, Row: 1}},
		{"aA1", CellRef{Col: 27, Row: 1}},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := parseCellRefToken(tt.input)
			if err != nil {
				t.Fatalf("parseCellRefToken(%q) error: %v", tt.input, err)
			}
			if *got != tt.want {
				t.Errorf("parseCellRefToken(%q)\n  got:  %+v\n  want: %+v", tt.input, *got, tt.want)
			}
		})
	}
}

func TestParseCellRefTokenColumnOnly(t *testing.T) {
	tests := []struct {
		input string
		want  CellRef
	}{
		// Column-only references have Row=0.
		{"F", CellRef{Col: 6, Row: 0}},
		{"$A", CellRef{Col: 1, Row: 0, AbsCol: true}},
		{"AA", CellRef{Col: 27, Row: 0}},
		{"XFD", CellRef{Col: 16384, Row: 0}},

		// Sheet-qualified column-only refs.
		{"Sheet1!F", CellRef{Sheet: "Sheet1", Col: 6, Row: 0}},
		{"Ledger!F", CellRef{Sheet: "Ledger", Col: 6, Row: 0}},
		{"'Sheet Name'!B", CellRef{Sheet: "Sheet Name", Col: 2, Row: 0}},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, err := parseCellRefToken(tt.input)
			if err != nil {
				t.Fatalf("parseCellRefToken(%q) error: %v", tt.input, err)
			}
			if *got != tt.want {
				t.Errorf("parseCellRefToken(%q)\n  got:  %+v\n  want: %+v", tt.input, *got, tt.want)
			}
		})
	}
}

func TestParseCellRefTokenErrors(t *testing.T) {
	tests := []string{
		"",
		"$",
		"123",
		"!A1",
		"'unterminated!A1",
		"'Sheet'X1", // missing ! after quoted name
	}

	for _, input := range tests {
		t.Run(input, func(t *testing.T) {
			_, err := parseCellRefToken(input)
			if err == nil {
				t.Errorf("parseCellRefToken(%q) expected error, got nil", input)
			}
		})
	}
}

func TestParseCellRefToken3DRef(t *testing.T) {
	// 3D sheet references (multi-sheet ranges) should parse successfully.
	tests := []struct {
		name  string
		input string
		want  CellRef
	}{
		{"quoted_3d", "'Sheet2:Sheet5'!A11", CellRef{Sheet: "Sheet2", SheetEnd: "Sheet5", Col: 1, Row: 11}},
		{"quoted_3d_spaces", "'My Sheet:Other Sheet'!B2", CellRef{Sheet: "My Sheet", SheetEnd: "Other Sheet", Col: 2, Row: 2}},
		{"unquoted_3d", "Sheet2:Sheet5!A11", CellRef{Sheet: "Sheet2", SheetEnd: "Sheet5", Col: 1, Row: 11}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := parseCellRefToken(tt.input)
			if err != nil {
				t.Fatalf("parseCellRefToken(%q) error: %v", tt.input, err)
			}
			if *got != tt.want {
				t.Errorf("parseCellRefToken(%q)\n  got:  %+v\n  want: %+v", tt.input, *got, tt.want)
			}
		})
	}
}

func TestParseCellRefTokenOutOfRange(t *testing.T) {
	// Column or row numbers exceeding limits must return an error.
	tests := []struct {
		name  string
		input string
	}{
		{"col_too_large", "AAAA1"},           // 4 letters → col > 16384
		{"col_beyond_xfd", "XFE1"},           // one past XFD
		{"row_too_large", "A1048577"},        // one past max row
		{"row_way_too_large", "A9999999999"}, // very large row
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, err := parseCellRefToken(tt.input)
			if err == nil {
				t.Fatalf("parseCellRefToken(%q) expected error for out-of-range ref, got nil", tt.input)
			}
			t.Logf("parseCellRefToken(%q) correctly returned error: %v", tt.input, err)
		})
	}
}

func TestColLettersToNumber(t *testing.T) {
	tests := []struct {
		input string
		want  int
	}{
		{"A", 1},
		{"B", 2},
		{"Z", 26},
		{"AA", 27},
		{"AZ", 52},
		{"BA", 53},
		{"XFD", 16384},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got := colLettersToNumber(tt.input)
			if got != tt.want {
				t.Errorf("colLettersToNumber(%q) = %d, want %d", tt.input, got, tt.want)
			}
		})
	}
}
