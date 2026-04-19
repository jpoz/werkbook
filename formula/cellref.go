package formula

import (
	"fmt"
	"strings"
)

// parseCellRefToken parses a TokCellRef value string into a CellRef node.
// It handles bare refs (A1, $A$1), unquoted sheet refs (Sheet1!A1),
// and quoted sheet refs ('Sheet Name'!A1, 'It”s a sheet'!B2).
func parseCellRefToken(raw string) (*CellRef, error) {
	ref := &CellRef{}
	s := raw

	// Check for quoted sheet: 'Sheet Name'!...
	if len(s) > 0 && s[0] == '\'' {
		i := 1
		for i < len(s) {
			if s[i] == '\'' {
				if i+1 < len(s) && s[i+1] == '\'' {
					i += 2 // skip escaped quote
					continue
				}
				break // end of quoted name
			}
			i++
		}
		if i >= len(s) || s[i] != '\'' {
			return nil, fmt.Errorf("unterminated quoted sheet name in %q", raw)
		}
		// Extract sheet name, un-escaping doubled quotes.
		sheetName := strings.ReplaceAll(s[1:i], "''", "'")
		// Handle 3D sheet references like 'Sheet2:Sheet5'!A1.
		if colonIdx := strings.IndexByte(sheetName, ':'); colonIdx > 0 {
			ref.Sheet = sheetName[:colonIdx]
			ref.SheetEnd = sheetName[colonIdx+1:]
		} else {
			ref.Sheet = sheetName
		}
		if i+1 >= len(s) || s[i+1] != '!' {
			return nil, fmt.Errorf("expected '!' after quoted sheet name in %q", raw)
		}
		s = s[i+2:] // skip past closing quote and !
	} else if idx := strings.IndexByte(s, '!'); idx > 0 {
		// Unquoted sheet: Sheet1!A1 or 3D: Sheet2:Sheet5!A1
		sheetPart := s[:idx]
		if colonIdx := strings.IndexByte(sheetPart, ':'); colonIdx > 0 {
			ref.Sheet = sheetPart[:colonIdx]
			ref.SheetEnd = sheetPart[colonIdx+1:]
		} else {
			ref.Sheet = sheetPart
		}
		s = s[idx+1:]
	} else if idx := findDotSheetSeparator(s); idx > 0 {
		// Dot notation: Sheet1.A1 (LibreOffice style; returns #NAME? in standard mode)
		ref.Sheet = s[:idx]
		ref.DotNotation = true
		s = s[idx+1:]
	} else if len(s) > 0 && s[0] == '!' {
		return nil, fmt.Errorf("empty sheet name in %q", raw)
	}

	// Parse the cell part: optional $, letters, optional $, digits.
	i := 0

	if i < len(s) && s[i] == '$' {
		ref.AbsCol = true
		i++
	}

	colStart := i
	for i < len(s) && ((s[i] >= 'A' && s[i] <= 'Z') || (s[i] >= 'a' && s[i] <= 'z')) {
		i++
	}
	if i == colStart {
		return nil, fmt.Errorf("expected column letters in %q", raw)
	}
	ref.Col = ColLettersToNumber(s[colStart:i])
	if ref.Col <= 0 {
		return nil, fmt.Errorf("column out of range in %q", raw)
	}

	if i < len(s) && s[i] == '$' {
		ref.AbsRow = true
		i++
	}

	rowStart := i
	for i < len(s) && s[i] >= '0' && s[i] <= '9' {
		i++
	}
	if i == rowStart {
		// No row number: this is a column-only reference (e.g. "F" or "Ledger!F").
		// Row=0 is the sentinel for column-only refs; the parser expands these
		// into full-column ranges when it sees the colon operator (F:F).
		ref.Row = 0
	} else {
		row := 0
		for _, c := range s[rowStart:i] {
			row = row*10 + int(c-'0')
		}
		if row > 1048576 { // max row
			return nil, fmt.Errorf("row out of range in %q", raw)
		}
		ref.Row = row
	}

	if i != len(s) {
		// Fall back to a bare-identifier CellRef when the input looks like an
		// identifier (letters, digits, underscores, dots) rather than a cell
		// reference. LET and LAMBDA use this for parameter names like
		// `running_sum` that can't be encoded as column+row. Only accept the
		// fallback when the raw input has no sheet qualifier, no '$', and
		// starts with a letter or underscore.
		if isBareIdentifier(raw) && ref.Sheet == "" && ref.SheetEnd == "" && !ref.AbsCol && !ref.AbsRow && !ref.DotNotation {
			return &CellRef{Name: raw}, nil
		}
		return nil, fmt.Errorf("unexpected trailing characters in cell ref %q", raw)
	}

	return ref, nil
}

// isBareIdentifier reports whether s is a legal Excel-style name identifier:
// starts with a letter or underscore, followed by letters, digits, underscores
// or dots. Intended for falling back from a failed cell-ref parse to a
// named-identifier interpretation.
func isBareIdentifier(s string) bool {
	if s == "" {
		return false
	}
	c := s[0]
	if !((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_') {
		return false
	}
	for i := 1; i < len(s); i++ {
		c := s[i]
		if !((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') ||
			(c >= '0' && c <= '9') || c == '_' || c == '.') {
			return false
		}
	}
	return true
}

// findDotSheetSeparator finds a '.' that separates a sheet name from a cell
// reference (e.g. "Sheet1.A1"). Returns the index of the dot, or -1 if not found.
// The text after the dot must start with an optional '$' then a letter (column).
func findDotSheetSeparator(s string) int {
	for i := 0; i < len(s); i++ {
		if s[i] == '.' {
			rest := s[i+1:]
			if len(rest) == 0 {
				continue
			}
			j := 0
			if j < len(rest) && rest[j] == '$' {
				j++
			}
			if j < len(rest) && ((rest[j] >= 'A' && rest[j] <= 'Z') || (rest[j] >= 'a' && rest[j] <= 'z')) {
				return i
			}
		}
	}
	return -1
}

// ColLettersToNumber converts column letters (e.g. "A", "AA", "XFD") to a 1-based column number.
func ColLettersToNumber(s string) int {
	col := 0
	for _, c := range strings.ToUpper(s) {
		col = col*26 + int(c-'A') + 1
	}
	return col
}
