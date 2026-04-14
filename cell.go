package werkbook

import "github.com/jpoz/werkbook/formula"

// Cell represents a single cell in a worksheet.
type Cell struct {
	col            int
	value          Value
	formula        string
	isArrayFormula bool // CSE (Ctrl+Shift+Enter) array formula
	// dynamicArraySpill controls whether this cell's dynamic-array result
	// should spill into neighboring cells. Imported formulas without OOXML
	// dynamic-array metadata do not spill in Excel and should stay scalar.
	dynamicArraySpill bool
	compiled          *formula.CompiledFormula
	rawValue          formula.Value
	cachedGen         uint64 // file.calcGen when value was last computed from formula
	rawCachedGen      uint64 // file.calcGen when rawValue was last computed from formula
	dirty             bool   // needs recalculation via dependency graph
	style             *Style
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

// IsDynamicArraySpill returns true if this cell is a dynamic-array spill anchor.
func (c *Cell) IsDynamicArraySpill() bool {
	return c.dynamicArraySpill
}

// Style returns the cell's style, or nil for default styling.
func (c *Cell) Style() *Style {
	return c.style
}
