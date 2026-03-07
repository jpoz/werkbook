package werkbook

import (
	"fmt"
	"strconv"
	"strings"
	"unicode"

	"github.com/jpoz/werkbook/formula"
	"github.com/jpoz/werkbook/ooxml"
)

// Table models a worksheet table (ListObject).
type Table struct {
	SheetName      string
	Name           string
	DisplayName    string
	Ref            string
	Columns        []string
	HeaderRowCount int
	TotalsRowCount int
	AutoFilter     bool
	Style          *TableStyle
}

// TableStyle models <tableStyleInfo>.
type TableStyle struct {
	Name              string
	ShowFirstColumn   bool
	ShowLastColumn    bool
	ShowRowStripes    bool
	ShowColumnStripes bool
}

// Tables returns the workbook's tables in workbook order.
func (f *File) Tables() []Table {
	out := make([]Table, len(f.tableDefs))
	for i, td := range f.tableDefs {
		out[i] = td.clone()
	}
	return out
}

// Tables returns the tables defined on this sheet.
func (s *Sheet) Tables() []Table {
	var out []Table
	for _, td := range s.file.tableDefs {
		if td.SheetName == s.name {
			out = append(out, td.clone())
		}
	}
	return out
}

// AddTable registers a table on the sheet and writes it as a proper OOXML table part.
func (s *Sheet) AddTable(td Table) error {
	norm, info, err := s.normalizeTable(td)
	if err != nil {
		return err
	}
	s.file.tableDefs = append(s.file.tableDefs, norm)
	s.file.tables = append(s.file.tables, info)
	s.file.rebuildFormulaState()
	return nil
}

func (s *Sheet) normalizeTable(td Table) (Table, formula.TableInfo, error) {
	if td.SheetName != "" && td.SheetName != s.name {
		return Table{}, formula.TableInfo{}, fmt.Errorf("table sheet %q does not match sheet %q", td.SheetName, s.name)
	}
	col1, row1, col2, row2, err := RangeToCoordinates(td.Ref)
	if err != nil {
		return Table{}, formula.TableInfo{}, fmt.Errorf("invalid table ref %q: %w", td.Ref, err)
	}
	width := col2 - col1 + 1
	height := row2 - row1 + 1

	headerRows := td.HeaderRowCount
	if headerRows == 0 {
		headerRows = 1
	}
	if headerRows < 0 || td.TotalsRowCount < 0 {
		return Table{}, formula.TableInfo{}, fmt.Errorf("table row counts must be non-negative")
	}
	if headerRows+td.TotalsRowCount > height {
		return Table{}, formula.TableInfo{}, fmt.Errorf("table ref %q is too small for header/totals rows", td.Ref)
	}

	name := strings.TrimSpace(td.Name)
	displayName := strings.TrimSpace(td.DisplayName)
	if name == "" && displayName == "" {
		name = s.file.nextTableName()
		displayName = name
	} else if name == "" {
		name = displayName
	} else if displayName == "" {
		displayName = name
	}
	if !isValidTableName(name) {
		return Table{}, formula.TableInfo{}, fmt.Errorf("invalid table name %q", name)
	}
	if !isValidTableName(displayName) {
		return Table{}, formula.TableInfo{}, fmt.Errorf("invalid table display name %q", displayName)
	}
	for _, existing := range s.file.tableDefs {
		if strings.EqualFold(existing.Name, name) || strings.EqualFold(existing.DisplayName, displayName) {
			return Table{}, formula.TableInfo{}, fmt.Errorf("table name %q already exists", displayName)
		}
		if existing.SheetName == s.name && refsOverlap(existing.Ref, td.Ref) {
			return Table{}, formula.TableInfo{}, fmt.Errorf("table ref %q overlaps existing table %q", td.Ref, existing.DisplayName)
		}
	}

	cols, err := s.normalizeTableColumns(td.Columns, col1, col2, row1, headerRows)
	if err != nil {
		return Table{}, formula.TableInfo{}, err
	}
	if len(cols) != width {
		return Table{}, formula.TableInfo{}, fmt.Errorf("table ref %q spans %d columns, but %d columns were provided", td.Ref, width, len(cols))
	}

	norm := Table{
		SheetName:      s.name,
		Name:           name,
		DisplayName:    displayName,
		Ref:            td.Ref,
		Columns:        cols,
		HeaderRowCount: headerRows,
		TotalsRowCount: td.TotalsRowCount,
		AutoFilter:     td.AutoFilter || headerRows > 0,
		Style:          copyTableStyle(td.Style),
	}
	info := formula.TableInfo{
		Name:       displayName,
		SheetName:  s.name,
		Columns:    append([]string(nil), cols...),
		FirstCol:   col1,
		FirstRow:   row1,
		LastCol:    col2,
		LastRow:    row2,
		HeaderRows: headerRows,
		TotalRows:  td.TotalsRowCount,
	}
	return norm, info, nil
}

func (s *Sheet) normalizeTableColumns(cols []string, firstCol, lastCol, headerRow, headerRows int) ([]string, error) {
	if len(cols) == 0 {
		cols = make([]string, lastCol-firstCol+1)
		if headerRows > 0 {
			for col := firstCol; col <= lastCol; col++ {
				cols[col-firstCol] = s.tableHeaderValue(col, headerRow)
			}
		}
	}
	out := make([]string, len(cols))
	seen := make(map[string]int)
	for i, col := range cols {
		name := strings.TrimSpace(col)
		if name == "" {
			name = "Column" + strconv.Itoa(i+1)
		}
		base := name
		lower := strings.ToLower(base)
		for seen[lower] > 0 {
			name = fmt.Sprintf("%s_%d", base, seen[lower]+1)
			lower = strings.ToLower(name)
		}
		seen[strings.ToLower(name)]++
		out[i] = name
	}
	return out, nil
}

func (s *Sheet) tableHeaderValue(col, row int) string {
	r, ok := s.rows[row]
	if !ok {
		return ""
	}
	c, ok := r.cells[col]
	if !ok {
		return ""
	}
	s.resolveCell(c, col, row)
	switch c.value.Type {
	case TypeString:
		return c.value.String
	case TypeNumber:
		return strconv.FormatFloat(c.value.Number, 'f', -1, 64)
	case TypeBool:
		if c.value.Bool {
			return "TRUE"
		}
		return "FALSE"
	case TypeError:
		return c.value.String
	default:
		return ""
	}
}

func refsOverlap(a, b string) bool {
	a1, ar1, a2, ar2, err := RangeToCoordinates(a)
	if err != nil {
		return false
	}
	b1, br1, b2, br2, err := RangeToCoordinates(b)
	if err != nil {
		return false
	}
	return a1 <= b2 && a2 >= b1 && ar1 <= br2 && ar2 >= br1
}

func isValidTableName(name string) bool {
	if name == "" {
		return false
	}
	if _, _, err := CellNameToCoordinates(name); err == nil {
		return false
	}
	for i, r := range name {
		if i == 0 {
			if !unicode.IsLetter(r) && r != '_' {
				return false
			}
			continue
		}
		if !unicode.IsLetter(r) && !unicode.IsDigit(r) && r != '_' {
			return false
		}
	}
	return true
}

func (f *File) nextTableName() string {
	for i := 1; ; i++ {
		name := "Table" + strconv.Itoa(i)
		used := false
		for _, td := range f.tableDefs {
			if strings.EqualFold(td.Name, name) || strings.EqualFold(td.DisplayName, name) {
				used = true
				break
			}
		}
		if !used {
			return name
		}
	}
}

func (t Table) clone() Table {
	out := t
	out.Columns = append([]string(nil), t.Columns...)
	out.Style = copyTableStyle(t.Style)
	return out
}

func copyTableStyle(style *TableStyle) *TableStyle {
	if style == nil {
		return nil
	}
	cp := *style
	return &cp
}

func tableFromData(td ooxml.TableDef, sheetName string) Table {
	return Table{
		SheetName:      sheetName,
		Name:           td.Name,
		DisplayName:    td.DisplayName,
		Ref:            td.Ref,
		Columns:        append([]string(nil), td.Columns...),
		HeaderRowCount: td.HeaderRowCount,
		TotalsRowCount: td.TotalsRowCount,
		AutoFilter:     td.HasAutoFilter,
		Style:          tableStyleFromData(td.Style),
	}
}

func tableStyleFromData(style *ooxml.TableStyleData) *TableStyle {
	if style == nil {
		return nil
	}
	return &TableStyle{
		Name:              style.Name,
		ShowFirstColumn:   style.ShowFirstColumn,
		ShowLastColumn:    style.ShowLastColumn,
		ShowRowStripes:    style.ShowRowStripes,
		ShowColumnStripes: style.ShowColumnStripes,
	}
}

func (t Table) toData(sheetIndex int) ooxml.TableDef {
	var style *ooxml.TableStyleData
	if t.Style != nil {
		style = &ooxml.TableStyleData{
			Name:              t.Style.Name,
			ShowFirstColumn:   t.Style.ShowFirstColumn,
			ShowLastColumn:    t.Style.ShowLastColumn,
			ShowRowStripes:    t.Style.ShowRowStripes,
			ShowColumnStripes: t.Style.ShowColumnStripes,
		}
	}
	return ooxml.TableDef{
		Name:            t.Name,
		DisplayName:     t.DisplayName,
		Ref:             t.Ref,
		SheetIndex:      sheetIndex,
		Columns:         append([]string(nil), t.Columns...),
		HeaderRowCount:  t.HeaderRowCount,
		TotalsRowCount:  t.TotalsRowCount,
		HasAutoFilter:   t.AutoFilter,
		HasActiveFilter: false,
		Style:           style,
	}
}
