package werkbook

// Cell represents a single cell in a worksheet.
type Cell struct {
	col   int
	value Value
}

// Col returns the 1-based column number.
func (c *Cell) Col() int {
	return c.col
}

// Value returns the cell's value.
func (c *Cell) Value() Value {
	return c.value
}
