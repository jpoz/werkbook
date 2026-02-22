package werkbook

// Cell represents a single cell in a worksheet.
type Cell struct {
	col     int
	value   Value
	formula string
}

// Col returns the 1-based column number.
func (c *Cell) Col() int {
	return c.col
}

// Value returns the cell's value.
func (c *Cell) Value() Value {
	return c.value
}

// Formula returns the cell's formula text, or "" if none.
func (c *Cell) Formula() string {
	return c.formula
}
