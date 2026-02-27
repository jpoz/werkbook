package werkbook

import (
	"fmt"
	"strings"
)

const (
	// MaxRows is the maximum number of rows in a worksheet (1-based).
	MaxRows = 1048576
	// MaxColumns is the maximum number of columns in a worksheet (1-based).
	MaxColumns = 16384 // XFD
)

// CellNameToCoordinates converts a cell reference like "B3" to (col, row) 1-based coordinates.
// For example, "B3" returns (2, 3).
func CellNameToCoordinates(cell string) (col, row int, err error) {
	cell = strings.TrimSpace(cell)
	if cell == "" {
		return 0, 0, fmt.Errorf("empty cell reference")
	}

	// Split into column letters and row digits.
	i := 0
	for i < len(cell) && isAlpha(cell[i]) {
		i++
	}
	if i == 0 || i == len(cell) {
		return 0, 0, fmt.Errorf("invalid cell reference %q", cell)
	}

	colStr := cell[:i]
	rowStr := cell[i:]

	col, err = ColumnNameToNumber(colStr)
	if err != nil {
		return 0, 0, err
	}

	row, err = parseRow(rowStr)
	if err != nil {
		return 0, 0, fmt.Errorf("invalid cell reference %q: %w", cell, err)
	}

	return col, row, nil
}

// CoordinatesToCellName converts 1-based (col, row) coordinates to a cell reference like "B3".
func CoordinatesToCellName(col, row int) (string, error) {
	if col < 1 || col > MaxColumns {
		return "", fmt.Errorf("column %d out of range [1, %d]", col, MaxColumns)
	}
	if row < 1 || row > MaxRows {
		return "", fmt.Errorf("row %d out of range [1, %d]", row, MaxRows)
	}
	colName := ColumnNumberToName(col)
	return fmt.Sprintf("%s%d", colName, row), nil
}

// ColumnNameToNumber converts a column name like "B" or "XFD" to a 1-based column number.
func ColumnNameToNumber(name string) (int, error) {
	if name == "" {
		return 0, fmt.Errorf("empty column name")
	}
	col := 0
	for _, c := range strings.ToUpper(name) {
		if c < 'A' || c > 'Z' {
			return 0, fmt.Errorf("invalid column name %q", name)
		}
		col = col*26 + int(c-'A') + 1
	}
	if col > MaxColumns {
		return 0, fmt.Errorf("column %q exceeds maximum (%d)", name, MaxColumns)
	}
	return col, nil
}

// ColumnNumberToName converts a 1-based column number to a column name like "A", "Z", "AA", "XFD".
func ColumnNumberToName(col int) string {
	var buf [3]byte // max 3 letters for XFD
	i := len(buf)
	for col > 0 {
		col-- // adjust to 0-based
		i--
		buf[i] = byte('A' + col%26)
		col /= 26
	}
	return string(buf[i:])
}

// RangeToCoordinates converts a range reference like "A1:C5" to
// (col1, row1, col2, row2) 1-based coordinates. It normalizes reversed
// ranges so that col1 <= col2 and row1 <= row2. A single-cell reference
// like "A1" (no colon) returns the same cell for both corners.
func RangeToCoordinates(ref string) (col1, row1, col2, row2 int, err error) {
	ref = strings.TrimSpace(ref)
	parts := strings.SplitN(ref, ":", 2)
	col1, row1, err = CellNameToCoordinates(parts[0])
	if err != nil {
		return 0, 0, 0, 0, fmt.Errorf("invalid range %q: %w", ref, err)
	}
	if len(parts) == 1 {
		return col1, row1, col1, row1, nil
	}
	col2, row2, err = CellNameToCoordinates(parts[1])
	if err != nil {
		return 0, 0, 0, 0, fmt.Errorf("invalid range %q: %w", ref, err)
	}
	// Normalize so that (col1,row1) is the top-left corner.
	if col1 > col2 {
		col1, col2 = col2, col1
	}
	if row1 > row2 {
		row1, row2 = row2, row1
	}
	return col1, row1, col2, row2, nil
}

func isAlpha(b byte) bool {
	return (b >= 'A' && b <= 'Z') || (b >= 'a' && b <= 'z')
}

func parseRow(s string) (int, error) {
	if s == "" {
		return 0, fmt.Errorf("empty row number")
	}
	n := 0
	for _, c := range s {
		if c < '0' || c > '9' {
			return 0, fmt.Errorf("invalid row number %q", s)
		}
		n = n*10 + int(c-'0')
		if n > MaxRows {
			return 0, fmt.Errorf("row %d exceeds maximum (%d)", n, MaxRows)
		}
	}
	if n == 0 {
		return 0, fmt.Errorf("row number must be >= 1")
	}
	return n, nil
}
