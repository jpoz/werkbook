package formula

import (
	"strconv"
	"strings"
)

// TableInfo holds the resolved geometry of a table, used to expand
// structured references in formulas before parsing.
type TableInfo struct {
	Name            string   // display name (case-insensitive match)
	SheetName       string   // sheet where the table lives
	Columns         []string // column names in order
	FirstCol        int      // 1-based first column of the table range
	FirstRow        int      // 1-based first row (header row if present)
	LastCol         int      // 1-based last column
	LastRow         int      // 1-based last row (including totals if present)
	HeaderRows      int      // number of header rows (usually 1)
	TotalRows       int      // number of total rows (usually 0)
	HasActiveFilter bool     // true if this table has an autoFilter with active filter columns
}

// DataFirstRow returns the 1-based first data row (after headers).
func (t *TableInfo) DataFirstRow() int {
	return t.FirstRow + t.HeaderRows
}

// DataLastRow returns the 1-based last data row (before totals).
func (t *TableInfo) DataLastRow() int {
	return t.LastRow - t.TotalRows
}

// ColumnIndex returns the 0-based index of the named column, or -1 if not found.
// Matching is case-insensitive and normalizes whitespace (CR, LF, CRLF are equivalent).
func (t *TableInfo) ColumnIndex(name string) int {
	norm := normalizeWS(strings.ToLower(name))
	for i, col := range t.Columns {
		if normalizeWS(strings.ToLower(col)) == norm {
			return i
		}
	}
	return -1
}

// normalizeWS collapses \r\n, \r, and \n to a single \n for comparison.
func normalizeWS(s string) string {
	s = strings.ReplaceAll(s, "\r\n", "\n")
	s = strings.ReplaceAll(s, "\r", "\n")
	return s
}

// ExpandTableRefs expands table structured references in a formula string.
// tables is the set of tables available in the workbook.
// currentRow is the 1-based row of the cell containing the formula.
// Returns the formula with structured references replaced by cell ranges.
func ExpandTableRefs(formula string, tables []TableInfo, currentRow int) string {
	if len(tables) == 0 || !containsTableRef(formula) {
		return formula
	}

	var result strings.Builder
	i := 0
	for i < len(formula) {
		// Skip string literals.
		if formula[i] == '"' {
			j := i + 1
			for j < len(formula) {
				if formula[j] == '"' {
					j++
					if j < len(formula) && formula[j] == '"' {
						j++ // escaped quote
						continue
					}
					break
				}
				j++
			}
			result.WriteString(formula[i:j])
			i = j
			continue
		}

		// Skip quoted sheet names like 'my-sheet'!A1. Without this guard, a
		// table whose name appears as a substring of the sheet name and is
		// (implausibly but legally) followed by '[' would trigger spurious
		// structured-reference expansion inside the quotes. Mirrors the guard
		// in ExpandDefinedNamesBounded.
		if formula[i] == '\'' {
			j := i + 1
			for j < len(formula) {
				if formula[j] == '\'' {
					if j+1 < len(formula) && formula[j+1] == '\'' {
						j += 2 // escaped quote inside sheet name
						continue
					}
					j++
					break
				}
				j++
			}
			result.WriteString(formula[i:j])
			i = j
			continue
		}

		// Look for TableName[ pattern.
		if isIdentStartByte(formula[i]) {
			nameStart := i
			for i < len(formula) && isIdentOrDotByte(formula[i]) {
				i++
			}
			name := formula[nameStart:i]

			if i < len(formula) && formula[i] == '[' {
				// Check if this name matches a table.
				table := findTable(tables, name)
				if table != nil {
					// Parse the structured reference bracket expression.
					bracketStart := i
					expanded, end := expandStructRef(formula, bracketStart, table, currentRow)
					if expanded != "" {
						result.WriteString(expanded)
						i = end
						continue
					}
				}
			}
			result.WriteString(name)
			continue
		}

		result.WriteByte(formula[i])
		i++
	}
	return result.String()
}

// containsTableRef quickly checks if the formula might contain a table reference.
func containsTableRef(formula string) bool {
	return strings.ContainsRune(formula, '[')
}

// findTable looks up a table by display name (case-insensitive).
func findTable(tables []TableInfo, name string) *TableInfo {
	lower := strings.ToLower(name)
	for i := range tables {
		if strings.ToLower(tables[i].Name) == lower {
			return &tables[i]
		}
	}
	return nil
}

// expandStructRef parses a structured reference starting at formula[pos] which is '['.
// Returns the expanded cell reference string and the position after the parsed reference,
// or ("", pos) if parsing fails.
func expandStructRef(formula string, pos int, table *TableInfo, currentRow int) (string, int) {
	// Scan for the matching close bracket, handling nested brackets.
	depth := 0
	end := pos
	for end < len(formula) {
		if formula[end] == '[' {
			depth++
		} else if formula[end] == ']' {
			depth--
			if depth == 0 {
				end++
				break
			}
		}
		end++
	}
	if depth != 0 {
		return "", pos
	}

	inner := formula[pos+1 : end-1]

	// Case 1: Simple column reference: Table[ColumnName]
	if !strings.Contains(inner, "[") {
		ref := expandSimpleColumn(table, inner)
		if ref == "" {
			return "", pos
		}
		return ref, end
	}

	// Case 2: Compound reference with specifiers: Table[[#This Row],[Column]]
	ref := expandCompoundRef(table, inner, currentRow)
	if ref == "" {
		return "", pos
	}
	return ref, end
}

// expandSimpleColumn handles Table[ColumnName] -> Sheet!C2:C100
func expandSimpleColumn(table *TableInfo, colName string) string {
	colIdx := table.ColumnIndex(colName)
	if colIdx < 0 {
		return ""
	}

	absCol := table.FirstCol + colIdx
	colLetters := ColNumberToLetters(absCol)
	dataFirst := table.DataFirstRow()
	dataLast := table.DataLastRow()

	var b strings.Builder
	writeSheetPrefix(&b, table.SheetName)
	b.WriteString(colLetters)
	b.WriteString(itoa(dataFirst))
	b.WriteByte(':')
	b.WriteString(colLetters)
	b.WriteString(itoa(dataLast))

	return b.String()
}

// expandCompoundRef handles Table[[specifier],[Column]] patterns.
func expandCompoundRef(table *TableInfo, inner string, currentRow int) string {
	parts := splitStructRef(inner)
	if len(parts) == 0 {
		return ""
	}

	specifier := ""
	var colNames []string

	for _, part := range parts {
		p := strings.TrimSpace(part)
		if len(p) >= 2 && p[0] == '[' && p[len(p)-1] == ']' {
			p = p[1 : len(p)-1]
		}
		if strings.HasPrefix(p, "#") {
			specifier = strings.ToLower(p)
		} else {
			colNames = append(colNames, p)
		}
	}

	if len(colNames) == 0 {
		return ""
	}

	switch specifier {
	case "#this row":
		return expandThisRow(table, colNames, currentRow)
	case "#headers":
		return expandHeaders(table, colNames)
	case "#totals":
		return expandTotals(table, colNames)
	case "#data", "":
		return expandData(table, colNames)
	default:
		return ""
	}
}

// expandThisRow handles [[#This Row],[Column]]
func expandThisRow(table *TableInfo, colNames []string, currentRow int) string {
	if len(colNames) == 1 {
		colIdx := table.ColumnIndex(colNames[0])
		if colIdx < 0 {
			return ""
		}
		absCol := table.FirstCol + colIdx
		var b strings.Builder
		writeSheetPrefix(&b, table.SheetName)
		b.WriteString(ColNumberToLetters(absCol))
		b.WriteString(itoa(currentRow))
		return b.String()
	}
	if len(colNames) == 2 {
		idx1 := table.ColumnIndex(colNames[0])
		idx2 := table.ColumnIndex(colNames[1])
		if idx1 < 0 || idx2 < 0 {
			return ""
		}
		col1 := table.FirstCol + idx1
		col2 := table.FirstCol + idx2
		var b strings.Builder
		writeSheetPrefix(&b, table.SheetName)
		b.WriteString(ColNumberToLetters(col1))
		b.WriteString(itoa(currentRow))
		b.WriteByte(':')
		b.WriteString(ColNumberToLetters(col2))
		b.WriteString(itoa(currentRow))
		return b.String()
	}
	return ""
}

// expandHeaders handles [[#Headers],[Column]]
func expandHeaders(table *TableInfo, colNames []string) string {
	if table.HeaderRows == 0 {
		return ""
	}
	headerRow := table.FirstRow
	if len(colNames) == 1 {
		colIdx := table.ColumnIndex(colNames[0])
		if colIdx < 0 {
			return ""
		}
		absCol := table.FirstCol + colIdx
		var b strings.Builder
		writeSheetPrefix(&b, table.SheetName)
		b.WriteString(ColNumberToLetters(absCol))
		b.WriteString(itoa(headerRow))
		return b.String()
	}
	return ""
}

// expandTotals handles [[#Totals],[Column]]
func expandTotals(table *TableInfo, colNames []string) string {
	if table.TotalRows == 0 {
		return ""
	}
	totalsRow := table.LastRow
	if len(colNames) == 1 {
		colIdx := table.ColumnIndex(colNames[0])
		if colIdx < 0 {
			return ""
		}
		absCol := table.FirstCol + colIdx
		var b strings.Builder
		writeSheetPrefix(&b, table.SheetName)
		b.WriteString(ColNumberToLetters(absCol))
		b.WriteString(itoa(totalsRow))
		return b.String()
	}
	return ""
}

// expandData handles [[Column]] or [[#Data],[Column]]
func expandData(table *TableInfo, colNames []string) string {
	if len(colNames) == 1 {
		return expandSimpleColumn(table, colNames[0])
	}
	if len(colNames) == 2 {
		idx1 := table.ColumnIndex(colNames[0])
		idx2 := table.ColumnIndex(colNames[1])
		if idx1 < 0 || idx2 < 0 {
			return ""
		}
		col1 := table.FirstCol + idx1
		col2 := table.FirstCol + idx2
		dataFirst := table.DataFirstRow()
		dataLast := table.DataLastRow()
		var b strings.Builder
		writeSheetPrefix(&b, table.SheetName)
		b.WriteString(ColNumberToLetters(col1))
		b.WriteString(itoa(dataFirst))
		b.WriteByte(':')
		b.WriteString(ColNumberToLetters(col2))
		b.WriteString(itoa(dataLast))
		return b.String()
	}
	return ""
}

// splitStructRef splits the inner part of a compound structured reference
// on commas that are not inside brackets.
func splitStructRef(s string) []string {
	var parts []string
	depth := 0
	start := 0
	for i := 0; i < len(s); i++ {
		switch s[i] {
		case '[':
			depth++
		case ']':
			depth--
		case ',':
			if depth == 0 {
				parts = append(parts, s[start:i])
				start = i + 1
			}
		}
	}
	parts = append(parts, s[start:])
	return parts
}

func writeSheetPrefix(b *strings.Builder, sheet string) {
	if sheet == "" {
		return
	}
	if needsQuoting(sheet) {
		b.WriteByte('\'')
		b.WriteString(strings.ReplaceAll(sheet, "'", "''"))
		b.WriteByte('\'')
	} else {
		b.WriteString(sheet)
	}
	b.WriteByte('!')
}

func isIdentStartByte(c byte) bool {
	return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_'
}

func itoa(n int) string {
	return strconv.Itoa(n)
}
