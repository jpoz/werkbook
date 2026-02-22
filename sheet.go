package werkbook

import (
	"fmt"
	"iter"
	"sort"

	"github.com/jpoz/werkbook/ooxml"
)

// Sheet represents a single worksheet in the workbook.
type Sheet struct {
	name string
	rows map[int]*Row
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
			if c.value.Type == TypeEmpty {
				continue
			}
			ref, _ := CoordinatesToCellName(cn, rn)
			cd := cellToData(ref, c.value)
			rd.Cells = append(rd.Cells, cd)
		}
		if len(rd.Cells) > 0 {
			sd.Rows = append(sd.Rows, rd)
		}
	}
	return sd
}

func cellToData(ref string, v Value) ooxml.CellData {
	switch v.Type {
	case TypeString:
		// Type "s" signals the writer to use the shared string table.
		// The Value field holds the raw string; the writer replaces it with the SST index.
		return ooxml.CellData{Ref: ref, Type: "s", Value: v.String}
	case TypeNumber:
		return ooxml.CellData{Ref: ref, Value: fmt.Sprintf("%g", v.Number)}
	case TypeBool:
		val := "0"
		if v.Bool {
			val = "1"
		}
		return ooxml.CellData{Ref: ref, Type: "b", Value: val}
	default:
		return ooxml.CellData{Ref: ref}
	}
}
