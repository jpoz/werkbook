package werkbook

import (
	"fmt"
	"io"
	"iter"
	"sort"

	"github.com/jpoz/werkbook/formula"
	"github.com/jpoz/werkbook/ooxml"
)

// Sheet represents a single worksheet in the workbook.
type Sheet struct {
	file *File
	name string
	rows map[int]*Row
}

func newSheet(name string, file *File) *Sheet {
	return &Sheet{
		file: file,
		name: name,
		rows: make(map[int]*Row),
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
	node, parseErr := formula.Parse(f)
	if parseErr == nil {
		cf, compErr := formula.Compile(f, node)
		if compErr == nil {
			c.compiled = cf
			s.file.deps.Register(qc, s.name, cf.Refs, cf.Ranges)
		}
	}
	s.file.invalidateDependents(s.name, col, row)
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
func (s *Sheet) Rows() iter.Seq[*Row] {
	return func(yield func(*Row) bool) {
		rowNums := make([]int, 0, len(s.rows))
		for n := range s.rows {
			rowNums = append(rowNums, n)
		}
		sort.Ints(rowNums)
		for _, n := range rowNums {
			r := s.rows[n]
			if len(r.cells) == 0 {
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
func (s *Sheet) toSheetData() ooxml.SheetData {
	sd := ooxml.SheetData{Name: s.name}

	// Sort rows by number.
	rowNums := make([]int, 0, len(s.rows))
	for n := range s.rows {
		rowNums = append(rowNums, n)
	}
	sort.Ints(rowNums)

	for _, rn := range rowNums {
		r := s.rows[rn]
		if len(r.cells) == 0 {
			continue
		}
		rd := ooxml.RowData{Num: rn}

		// Sort cells by column.
		colNums := make([]int, 0, len(r.cells))
		for c := range r.cells {
			colNums = append(colNums, c)
		}
		sort.Ints(colNums)

		for _, cn := range colNums {
			c := r.cells[cn]
			if c.value.Type == TypeEmpty && c.formula == "" {
				continue
			}
			ref, _ := CoordinatesToCellName(cn, rn)
			cd := cellToData(ref, c.value, c.formula)
			rd.Cells = append(rd.Cells, cd)
		}
		if len(rd.Cells) > 0 {
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
		node, err := formula.Parse(c.formula)
		if err != nil {
			return Value{Type: TypeError, String: "#NAME?"}
		}
		compiled, err := formula.Compile(c.formula, node)
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
		CurrentCol:   col,
		CurrentRow:   row,
		CurrentSheet: s.name,
	}
	result, err := formula.Eval(cf, resolver, ctx)
	if err != nil {
		return Value{Type: TypeError, String: err.Error()}
	}

	return formulaValueToValue(result)
}

// formulaValueToValue converts a formula.Value to a werkbook Value.
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
		return Value{Type: TypeEmpty}
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
		rows := make([][]formula.Value, addr.ToRow-addr.FromRow+1)
		for i := range rows {
			row := make([]formula.Value, addr.ToCol-addr.FromCol+1)
			for j := range row {
				row[j] = formula.ErrorVal(formula.ErrValREF)
			}
			rows[i] = row
		}
		return rows
	}

	rows := make([][]formula.Value, addr.ToRow-addr.FromRow+1)
	for r := addr.FromRow; r <= addr.ToRow; r++ {
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
		return formula.ErrorVal(formula.ErrValVALUE)
	default:
		return formula.EmptyVal()
	}
}

func cellToData(ref string, v Value, formula string) ooxml.CellData {
	var cd ooxml.CellData
	switch v.Type {
	case TypeString:
		cd = ooxml.CellData{Ref: ref, Type: "s", Value: v.String}
	case TypeNumber:
		cd = ooxml.CellData{Ref: ref, Value: fmt.Sprintf("%g", v.Number)}
	case TypeBool:
		val := "0"
		if v.Bool {
			val = "1"
		}
		cd = ooxml.CellData{Ref: ref, Type: "b", Value: val}
	default:
		cd = ooxml.CellData{Ref: ref}
	}
	cd.Formula = formula
	return cd
}
