package werkbook

import (
	"fmt"
	"io"
	"iter"
	"sort"
	"strings"

	"github.com/jpoz/werkbook/formula"
	"github.com/jpoz/werkbook/ooxml"
)

// Sheet represents a single worksheet in the workbook.
type Sheet struct {
	file      *File
	name      string
	rows      map[int]*Row
	colWidths map[int]float64
}

func newSheet(name string, file *File) *Sheet {
	return &Sheet{
		file:      file,
		name:      name,
		rows:      make(map[int]*Row),
		colWidths: make(map[int]float64),
	}
}

// Name returns the sheet name.
func (s *Sheet) Name() string {
	return s.name
}

// SetValue sets the value of a cell by reference (e.g. "A1").
// Supported types: string, bool, int*, uint*, float32, float64, nil.
func (s *Sheet) SetValue(cell string, v any) error {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}

	val, err := toValue(v)
	if err != nil {
		return err
	}

	r := s.ensureRow(row)
	c := r.ensureCell(col)
	// Unregister old formula if any.
	if c.formula != "" {
		s.file.deps.Unregister(formula.QualifiedCell{Sheet: s.name, Col: col, Row: row})
	}
	c.value = val
	c.formula = ""
	c.compiled = nil
	c.cachedGen = 0
	s.file.calcGen++
	s.file.invalidateDependents(s.name, col, row)
	return nil
}

// SetFormula sets a formula on a cell by reference (e.g. "A1").
// The formula should not include the leading '=' sign.
func (s *Sheet) SetFormula(cell string, f string) error {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	r := s.ensureRow(row)
	c := r.ensureCell(col)
	// Unregister old formula if any.
	qc := formula.QualifiedCell{Sheet: s.name, Col: col, Row: row}
	if c.formula != "" {
		s.file.deps.Unregister(qc)
	}
	c.formula = f
	c.compiled = nil
	c.value = Value{}
	c.cachedGen = 0
	c.dirty = true
	s.file.calcGen++
	// Compile and register in dep graph.
	src := s.file.expandFormula(f, s.name, row)
	node, parseErr := formula.Parse(src)
	if parseErr == nil {
		cf, compErr := formula.Compile(src, node)
		if compErr == nil {
			c.compiled = cf
			s.file.deps.Register(qc, s.name, cf.Refs, cf.Ranges)
		}
	}
	s.file.invalidateDependents(s.name, col, row)
	return nil
}

// SetStyle sets the style of a cell by reference (e.g. "A1").
func (s *Sheet) SetStyle(cell string, style *Style) error {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	r := s.ensureRow(row)
	c := r.ensureCell(col)
	c.style = style
	return nil
}

// GetStyle returns the style of a cell by reference (e.g. "A1").
// Returns nil for default-styled or nonexistent cells.
func (s *Sheet) GetStyle(cell string) (*Style, error) {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return nil, fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	r, ok := s.rows[row]
	if !ok {
		return nil, nil
	}
	c, ok := r.cells[col]
	if !ok {
		return nil, nil
	}
	return c.style, nil
}

// SetColumnWidth sets the width of a column by name (e.g. "A").
func (s *Sheet) SetColumnWidth(col string, width float64) error {
	num, err := ColumnNameToNumber(col)
	if err != nil {
		return fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	s.colWidths[num] = width
	return nil
}

// GetColumnWidth returns the width of a column by name, or 0 if not set.
func (s *Sheet) GetColumnWidth(col string) (float64, error) {
	num, err := ColumnNameToNumber(col)
	if err != nil {
		return 0, fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	return s.colWidths[num], nil
}

// SetRowHeight sets the height of a row by 1-based row number.
func (s *Sheet) SetRowHeight(row int, height float64) error {
	if row < 1 || row > MaxRows {
		return fmt.Errorf("%w: row %d out of range [1, %d]", ErrInvalidCellRef, row, MaxRows)
	}
	r := s.ensureRow(row)
	r.height = height
	return nil
}

// GetRowHeight returns the height of a row, or 0 if not set.
func (s *Sheet) GetRowHeight(row int) (float64, error) {
	if row < 1 || row > MaxRows {
		return 0, fmt.Errorf("%w: row %d out of range [1, %d]", ErrInvalidCellRef, row, MaxRows)
	}
	r, ok := s.rows[row]
	if !ok {
		return 0, nil
	}
	return r.height, nil
}

// SetRangeStyle applies the given style to every cell in the range (e.g. "A1:C5").
// Cells that do not yet exist are created.
func (s *Sheet) SetRangeStyle(rangeRef string, style *Style) error {
	col1, row1, col2, row2, err := RangeToCoordinates(rangeRef)
	if err != nil {
		return fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	for r := row1; r <= row2; r++ {
		row := s.ensureRow(r)
		for c := col1; c <= col2; c++ {
			cell := row.ensureCell(c)
			cell.style = style
		}
	}
	return nil
}

// GetFormula returns the formula for a cell, or "" if none.
func (s *Sheet) GetFormula(cell string) (string, error) {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return "", fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	r, ok := s.rows[row]
	if !ok {
		return "", nil
	}
	c, ok := r.cells[col]
	if !ok {
		return "", nil
	}
	return c.formula, nil
}

// GetValue returns the value of a cell by reference (e.g. "A1").
func (s *Sheet) GetValue(cell string) (Value, error) {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return Value{}, fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}

	r, ok := s.rows[row]
	if !ok {
		return Value{Type: TypeEmpty}, nil
	}
	c, ok := r.cells[col]
	if !ok {
		return Value{Type: TypeEmpty}, nil
	}

	s.resolveCell(c, col, row)
	return c.value, nil
}

func (s *Sheet) ensureRow(num int) *Row {
	r, ok := s.rows[num]
	if !ok {
		r = &Row{num: num, cells: make(map[int]*Cell)}
		s.rows[num] = r
	}
	return r
}

func (r *Row) ensureCell(col int) *Cell {
	c, ok := r.cells[col]
	if !ok {
		c = &Cell{col: col}
		r.cells[col] = c
	}
	return c
}

// Rows returns an iterator over all non-empty rows in ascending order.
// Rows with a custom height but no cells are also included.
func (s *Sheet) Rows() iter.Seq[*Row] {
	return func(yield func(*Row) bool) {
		rowNums := make([]int, 0, len(s.rows))
		for n := range s.rows {
			rowNums = append(rowNums, n)
		}
		sort.Ints(rowNums)
		for _, n := range rowNums {
			r := s.rows[n]
			if len(r.cells) == 0 && r.height == 0 {
				continue
			}
			if !yield(r) {
				return
			}
		}
	}
}

// MaxRow returns the highest 1-based row number with data, or 0 if empty.
func (s *Sheet) MaxRow() int {
	max := 0
	for n := range s.rows {
		if n > max {
			max = n
		}
	}
	return max
}

// MaxCol returns the highest 1-based column number with data across all rows, or 0 if empty.
func (s *Sheet) MaxCol() int {
	max := 0
	for _, r := range s.rows {
		for c := range r.cells {
			if c > max {
				max = c
			}
		}
	}
	return max
}

// PrintTo writes a human-readable table of all cell values to w.
func (s *Sheet) PrintTo(w io.Writer) {
	maxCol := s.MaxCol()
	if maxCol == 0 {
		return
	}

	colWidths := make([]int, maxCol)
	var grid [][]string

	for row := range s.Rows() {
		vals := make([]string, maxCol)
		for _, c := range row.Cells() {
			ref, _ := CoordinatesToCellName(c.Col(), row.Num())
			v, _ := s.GetValue(ref)
			var text string
			switch v.Type {
			case TypeNumber:
				if v.Number == float64(int64(v.Number)) {
					text = fmt.Sprintf("%d", int64(v.Number))
				} else {
					text = fmt.Sprintf("%.2f", v.Number)
				}
			case TypeString:
				text = v.String
			case TypeBool:
				if v.Bool {
					text = "TRUE"
				} else {
					text = "FALSE"
				}
			case TypeError:
				text = v.String
			}
			idx := c.Col() - 1
			vals[idx] = text
			if len(text) > colWidths[idx] {
				colWidths[idx] = len(text)
			}
		}
		grid = append(grid, vals)
	}

	for _, vals := range grid {
		for c, text := range vals {
			if c > 0 {
				fmt.Fprint(w, "  ")
			}
			fmt.Fprintf(w, "%-*s", colWidths[c], text)
		}
		fmt.Fprintln(w)
	}
}

// toSheetData converts the sheet to the ooxml intermediate representation.
// styleMap maps style keys to indices in the WorkbookData.Styles slice.
// styles collects all unique StyleData values; both are mutated in place.
func (s *Sheet) toSheetData(styleMap map[string]int, styles *[]ooxml.StyleData) ooxml.SheetData {
	sd := ooxml.SheetData{Name: s.name}

	// Convert column widths map to ColWidthData slice.
	if len(s.colWidths) > 0 {
		colNums := make([]int, 0, len(s.colWidths))
		for c := range s.colWidths {
			colNums = append(colNums, c)
		}
		sort.Ints(colNums)
		for _, c := range colNums {
			sd.ColWidths = append(sd.ColWidths, ooxml.ColWidthData{
				Min: c, Max: c, Width: s.colWidths[c],
			})
		}
	}

	// Sort rows by number.
	rowNums := make([]int, 0, len(s.rows))
	for n := range s.rows {
		rowNums = append(rowNums, n)
	}
	sort.Ints(rowNums)

	for _, rn := range rowNums {
		r := s.rows[rn]
		if len(r.cells) == 0 && r.height == 0 {
			continue
		}
		rd := ooxml.RowData{Num: rn, Height: r.height}

		// Sort cells by column.
		colNums := make([]int, 0, len(r.cells))
		for c := range r.cells {
			colNums = append(colNums, c)
		}
		sort.Ints(colNums)

		for _, cn := range colNums {
			c := r.cells[cn]
			// Resolve dirty/stale formulas before serializing.
			s.resolveCell(c, cn, rn)
			if c.value.Type == TypeEmpty && c.formula == "" && c.style == nil {
				continue
			}
			ref, _ := CoordinatesToCellName(cn, rn)
			cd := cellToData(ref, c.value, c.formula)

			if c.style != nil {
				stData := styleToStyleData(c.style)
				key := styleKey(stData)
				idx, ok := styleMap[key]
				if !ok {
					idx = len(*styles)
					styleMap[key] = idx
					*styles = append(*styles, stData)
				}
				cd.StyleIdx = idx
			}

			rd.Cells = append(rd.Cells, cd)
		}
		if len(rd.Cells) > 0 || rd.Height != 0 {
			sd.Rows = append(sd.Rows, rd)
		}
	}
	return sd
}

// resolveCell evaluates the cell's formula if it is dirty or stale.
// dirty is the primary signal from the dep graph; cachedGen is a safety net
// for formulas not yet registered in the dep graph.
func (s *Sheet) resolveCell(c *Cell, col, row int) {
	if c.formula != "" && (c.dirty || c.cachedGen < s.file.calcGen) {
		c.value = s.evaluateFormula(c, col, row)
		c.cachedGen = s.file.calcGen
		c.dirty = false
	}
}

// evaluateFormula parses, compiles, and executes the formula on the given cell.
func (s *Sheet) evaluateFormula(c *Cell, col, row int) Value {
	f := s.file
	if f.evaluating == nil {
		f.evaluating = make(map[cellKey]bool)
	}
	key := cellKey{sheet: s.name, col: col, row: row}
	if f.evaluating[key] {
		// Circular reference
		return Value{Type: TypeError, String: "#REF!"}
	}
	f.evaluating[key] = true
	defer delete(f.evaluating, key)

	cf := c.compiled
	if cf == nil {
		// Expand table structured references and defined names before parsing.
		src := f.expandFormula(c.formula, s.name, row)
		node, err := formula.Parse(src)
		if err != nil {
			return Value{Type: TypeError, String: "#NAME?"}
		}
		compiled, err := formula.Compile(src, node)
		if err != nil {
			return Value{Type: TypeError, String: "#NAME?"}
		}
		c.compiled = compiled
		cf = compiled
		// Register in dep graph on first compilation.
		qc := formula.QualifiedCell{Sheet: s.name, Col: col, Row: row}
		f.deps.Register(qc, s.name, cf.Refs, cf.Ranges)
	}

	resolver := &fileResolver{file: f, currentSheet: s.name}
	ctx := &formula.EvalContext{
		CurrentCol:     col,
		CurrentRow:     row,
		CurrentSheet:   s.name,
		IsArrayFormula: c.isArrayFormula,
		Date1904:       f.date1904,
		Resolver:       resolver,
	}
	result, err := formula.Eval(cf, resolver, ctx)
	if err != nil {
		return Value{Type: TypeError, String: err.Error()}
	}

	return formulaValueToValue(result)
}

// formulaValueToValue converts a formula.Value to a werkbook Value.
// Excel coerces empty formula results to 0 (a cell containing =EmptyRef
// displays and caches 0, not blank), so ValueEmpty maps to TypeNumber 0.
func formulaValueToValue(fv formula.Value) Value {
	switch fv.Type {
	case formula.ValueNumber:
		return Value{Type: TypeNumber, Number: fv.Num}
	case formula.ValueString:
		return Value{Type: TypeString, String: fv.Str}
	case formula.ValueBool:
		return Value{Type: TypeBool, Bool: fv.Bool}
	case formula.ValueError:
		return Value{Type: TypeError, String: fv.Err.String()}
	default:
		// Excel treats empty formula results as numeric 0.
		return Value{Type: TypeNumber, Number: 0}
	}
}

// fileResolver implements formula.CellResolver with cross-sheet support.
type fileResolver struct {
	file         *File
	currentSheet string // sheet name for resolving unqualified refs
}

func (fr *fileResolver) resolveSheet(name string) *Sheet {
	if name == "" {
		name = fr.currentSheet
	}
	return fr.file.Sheet(name)
}

func (fr *fileResolver) GetCellValue(addr formula.CellAddr) formula.Value {
	s := fr.resolveSheet(addr.Sheet)
	if s == nil {
		return formula.ErrorVal(formula.ErrValREF)
	}

	r, ok := s.rows[addr.Row]
	if !ok {
		return formula.EmptyVal()
	}
	c, ok := r.cells[addr.Col]
	if !ok {
		return formula.EmptyVal()
	}

	s.resolveCell(c, addr.Col, addr.Row)
	return valueToFormulaValue(c.value)
}

func (fr *fileResolver) GetRangeValues(addr formula.RangeAddr) [][]formula.Value {
	s := fr.resolveSheet(addr.Sheet)
	if s == nil {
		// For unresolvable sheets, return a single-row error rather than
		// potentially allocating millions of rows for full-column refs.
		nRows := addr.ToRow - addr.FromRow + 1
		if nRows > 1000 {
			nRows = 1
		}
		rows := make([][]formula.Value, nRows)
		for i := range rows {
			row := make([]formula.Value, addr.ToCol-addr.FromCol+1)
			for j := range row {
				row[j] = formula.ErrorVal(formula.ErrValREF)
			}
			rows[i] = row
		}
		return rows
	}

	// Clamp the row range to the sheet's actual data extent so that
	// full-column references (e.g. F:F → F1:F1048576) don't allocate
	// a million rows.
	toRow := addr.ToRow
	if maxRow := s.MaxRow(); maxRow > 0 && toRow > maxRow {
		toRow = maxRow
	}
	if toRow < addr.FromRow {
		toRow = addr.FromRow
	}

	rows := make([][]formula.Value, toRow-addr.FromRow+1)
	for r := addr.FromRow; r <= toRow; r++ {
		row := make([]formula.Value, addr.ToCol-addr.FromCol+1)
		for col := addr.FromCol; col <= addr.ToCol; col++ {
			row[col-addr.FromCol] = fr.GetCellValue(formula.CellAddr{
				Sheet: addr.Sheet,
				Col:   col,
				Row:   r,
			})
		}
		rows[r-addr.FromRow] = row
	}
	return rows
}

// GetSheetNames returns the ordered list of all sheet names in the workbook.
func (fr *fileResolver) GetSheetNames() []string {
	return fr.file.SheetNames()
}

// IsSubtotalCell reports whether the cell at (sheet, col, row) contains a formula
// whose outermost function call is SUBTOTAL. This is used by the SUBTOTAL function
// to skip nested SUBTOTAL results and avoid double-counting.
func (fr *fileResolver) IsSubtotalCell(sheet string, col, row int) bool {
	s := fr.resolveSheet(sheet)
	if s == nil {
		return false
	}
	r, ok := s.rows[row]
	if !ok {
		return false
	}
	c, ok := r.cells[col]
	if !ok {
		return false
	}
	return isSubtotalFormula(c.formula)
}

// isSubtotalFormula returns true if the formula string starts with "SUBTOTAL("
// (case-insensitive), with optional leading whitespace. This matches both
// "SUBTOTAL(...)" and "_xlfn.SUBTOTAL(...)".
func isSubtotalFormula(f string) bool {
	if f == "" {
		return false
	}
	upper := strings.ToUpper(strings.TrimSpace(f))
	return strings.HasPrefix(upper, "SUBTOTAL(") || strings.HasPrefix(upper, "_XLFN.SUBTOTAL(")
}

// valueToFormulaValue converts a werkbook Value to a formula.Value.
func valueToFormulaValue(v Value) formula.Value {
	switch v.Type {
	case TypeNumber:
		return formula.NumberVal(v.Number)
	case TypeString:
		return formula.StringVal(v.String)
	case TypeBool:
		return formula.BoolVal(v.Bool)
	case TypeError:
		return formula.ErrorVal(formula.ErrorValueFromString(v.String))
	default:
		return formula.EmptyVal()
	}
}

func cellToData(ref string, v Value, f string) ooxml.CellData {
	var cd ooxml.CellData
	switch v.Type {
	case TypeString:
		if f != "" {
			cd = ooxml.CellData{Ref: ref, Type: "str", Value: v.String}
		} else {
			cd = ooxml.CellData{Ref: ref, Type: "s", Value: v.String}
		}
	case TypeNumber:
		cd = ooxml.CellData{Ref: ref, Value: fmt.Sprintf("%g", v.Number)}
	case TypeBool:
		val := "0"
		if v.Bool {
			val = "1"
		}
		cd = ooxml.CellData{Ref: ref, Type: "b", Value: val}
	case TypeError:
		cd = ooxml.CellData{Ref: ref, Type: "e", Value: v.String}
	default:
		cd = ooxml.CellData{Ref: ref}
	}
	cd.Formula = formula.AddXlfnPrefixes(f)
	return cd
}
