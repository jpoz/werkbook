package werkbook

// Row represents a row in a worksheet.
type Row struct {
	num    int
	cells  map[int]*Cell
	height float64
}

// Num returns the 1-based row number.
func (r *Row) Num() int {
	return r.num
}

// Height returns the custom row height, or 0 if not set.
func (r *Row) Height() float64 {
	return r.height
}

// Cells returns all cells in this row, sorted by column.
func (r *Row) Cells() []*Cell {
	if len(r.cells) == 0 {
		return nil
	}
	maxCol := 0
	for c := range r.cells {
		if c > maxCol {
			maxCol = c
		}
	}
	cells := make([]*Cell, 0, len(r.cells))
	for col := 1; col <= maxCol; col++ {
		if c, ok := r.cells[col]; ok {
			cells = append(cells, c)
		}
	}
	return cells
}
