package werkbook

import "github.com/jpoz/werkbook/formula"

// Cell represents a single cell in a worksheet.
type Cell struct {
	col       int
	value     Value
	formula   string
	compiled  *formula.CompiledFormula
	cachedGen uint64 // file.calcGen when value was last computed from formula
	dirty     bool   // needs recalculation via dependency graph
	style     *Style
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

// Style returns the cell's style, or nil for default styling.
func (c *Cell) Style() *Style {
	return c.style
}
