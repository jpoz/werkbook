package werkbook

import (
	"fmt"
	"iter"
	"sort"

	"github.com/jpoz/werkbook/formula"
	"github.com/jpoz/werkbook/ooxml"
)

// Sheet represents a single worksheet in the workbook.
type Sheet struct {
	name       string
	rows       map[int]*Row
	evaluating map[[2]int]bool // tracks cells currently being evaluated (circular ref detection)
}

func newSheet(name string) *Sheet {
	return &Sheet{
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
	c.value = val
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
	if c.formula != "" {
		c.value = Value{}
	}
	c.formula = f
	c.compiled = nil
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

	if c.formula != "" && c.value.Type == TypeEmpty {
		v := s.evaluateFormula(c, col, row)
		c.value = v
	}

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

// evaluateFormula parses, compiles, and executes the formula on the given cell.
func (s *Sheet) evaluateFormula(c *Cell, col, row int) Value {
	if s.evaluating == nil {
		s.evaluating = make(map[[2]int]bool)
	}
	key := [2]int{col, row}
	if s.evaluating[key] {
		// Circular reference
		return Value{Type: TypeError, String: "#REF!"}
	}
	s.evaluating[key] = true
	defer delete(s.evaluating, key)

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
	}

	resolver := &sheetResolver{sheet: s}
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
		return Value{Type: TypeError, String: fv.Str}
	default:
		return Value{Type: TypeEmpty}
	}
}

// sheetResolver implements formula.CellResolver backed by a Sheet.
type sheetResolver struct {
	sheet *Sheet
}

func (sr *sheetResolver) GetCellValue(addr formula.CellAddr) formula.Value {
	r, ok := sr.sheet.rows[addr.Row]
	if !ok {
		return formula.EmptyVal()
	}
	c, ok := r.cells[addr.Col]
	if !ok {
		return formula.EmptyVal()
	}

	// If this cell has a formula that hasn't been evaluated, evaluate it now.
	if c.formula != "" && c.value.Type == TypeEmpty {
		v := sr.sheet.evaluateFormula(c, addr.Col, addr.Row)
		c.value = v
	}

	return valueToFormulaValue(c.value)
}

func (sr *sheetResolver) GetRangeValues(addr formula.RangeAddr) [][]formula.Value {
	rows := make([][]formula.Value, addr.ToRow-addr.FromRow+1)
	for r := addr.FromRow; r <= addr.ToRow; r++ {
		row := make([]formula.Value, addr.ToCol-addr.FromCol+1)
		for col := addr.FromCol; col <= addr.ToCol; col++ {
			row[col-addr.FromCol] = sr.GetCellValue(formula.CellAddr{
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
