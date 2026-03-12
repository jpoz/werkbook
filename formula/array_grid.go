package formula

// arrayGrid provides lazy access to a 2D value grid whose effective size may
// be larger than the materialized cells when it originated from a trimmed
// full-row or full-column worksheet reference.
type arrayGrid struct {
	cells       [][]Value
	rowCount    int
	colCount    int
	rangeOrigin *RangeAddr
}

// effectiveArrayBounds returns the logical size of an array value, preferring
// RangeOrigin when present so trimmed worksheet ranges still report their full
// extent.
func effectiveArrayBounds(v Value) (rows, cols int) {
	rows = len(v.Array)
	cols = materializedArrayCols(v.Array)
	if v.RangeOrigin == nil {
		return rows, cols
	}
	originRows := v.RangeOrigin.ToRow - v.RangeOrigin.FromRow + 1
	originCols := v.RangeOrigin.ToCol - v.RangeOrigin.FromCol + 1
	if originRows > rows {
		rows = originRows
	}
	if originCols > cols {
		cols = originCols
	}
	return rows, cols
}

func materializedArrayCols(rows [][]Value) int {
	cols := 0
	for _, row := range rows {
		if len(row) > cols {
			cols = len(row)
		}
	}
	return cols
}

func newValueMatrix(rowCount, colCount int) [][]Value {
	if rowCount <= 0 || colCount <= 0 {
		return nil
	}
	rows := make([][]Value, rowCount)
	cells := make([]Value, rowCount*colCount)
	for i := range rows {
		start := i * colCount
		rows[i] = cells[start : start+colCount]
	}
	return rows
}

func usesFullSheetAxisRange(v Value) bool {
	if v.RangeOrigin == nil {
		return false
	}
	return v.RangeOrigin.ToRow-v.RangeOrigin.FromRow+1 >= maxRows ||
		v.RangeOrigin.ToCol-v.RangeOrigin.FromCol+1 >= maxCols
}

// arrayOpBounds returns the dimensions evaluator-style array operations should
// materialize. Small trimmed ranges use their logical RangeOrigin size, while
// full-row/full-column refs keep the populated extent to avoid exploding
// intermediate arrays for expressions like A:A="x".
func arrayOpBounds(v Value) (rows, cols int) {
	if v.Type != ValueArray {
		return 1, 1
	}
	rows = len(v.Array)
	cols = materializedArrayCols(v.Array)
	if v.RangeOrigin == nil || usesFullSheetAxisRange(v) {
		return rows, cols
	}
	return effectiveArrayBounds(v)
}

func arrayOpRangeOrigin(v Value, rows, cols int) *RangeAddr {
	if v.Type != ValueArray || v.RangeOrigin == nil || usesFullSheetAxisRange(v) {
		return nil
	}
	wantRows, wantCols := effectiveArrayBounds(v)
	if wantRows != rows || wantCols != cols {
		return nil
	}
	origin := *v.RangeOrigin
	return &origin
}

func combinedArrayOpRangeOrigin(rows, cols int, values ...Value) *RangeAddr {
	var origin *RangeAddr
	for _, v := range values {
		candidate := arrayOpRangeOrigin(v, rows, cols)
		if candidate == nil {
			continue
		}
		if origin == nil {
			origin = candidate
			continue
		}
		if *origin != *candidate {
			return nil
		}
	}
	return origin
}

func normalizeToArrayGrid(v Value) (arrayGrid, *Value) {
	switch v.Type {
	case ValueArray:
		rows, cols := effectiveArrayBounds(v)
		if rows == 0 || cols == 0 {
			errVal := ErrorVal(ErrValVALUE)
			return arrayGrid{}, &errVal
		}
		return arrayGrid{
			cells:       v.Array,
			rowCount:    rows,
			colCount:    cols,
			rangeOrigin: v.RangeOrigin,
		}, nil
	case ValueError:
		return arrayGrid{}, &v
	default:
		return arrayGrid{
			cells:    [][]Value{{v}},
			rowCount: 1,
			colCount: 1,
		}, nil
	}
}

func (g arrayGrid) cell(row, col int) Value {
	if row >= 0 && row < len(g.cells) && col >= 0 && col < len(g.cells[row]) {
		return g.cells[row][col]
	}
	return EmptyVal()
}

func (g arrayGrid) hasMaterializedCell(row, col int) bool {
	return row >= 0 && row < len(g.cells) && col >= 0 && col < len(g.cells[row])
}

func (g arrayGrid) row(row int) []Value {
	values := make([]Value, g.colCount)
	for col := 0; col < g.colCount; col++ {
		values[col] = g.cell(row, col)
	}
	return values
}

func (g arrayGrid) subgrid(rowStart, rowEnd, colStart, colEnd int) [][]Value {
	if rowStart >= rowEnd || colStart >= colEnd {
		return nil
	}
	rows := make([][]Value, rowEnd-rowStart)
	for row := rowStart; row < rowEnd; row++ {
		values := make([]Value, colEnd-colStart)
		for col := colStart; col < colEnd; col++ {
			values[col-colStart] = g.cell(row, col)
		}
		rows[row-rowStart] = values
	}
	return rows
}

func (g arrayGrid) matrix() [][]Value {
	return g.subgrid(0, g.rowCount, 0, g.colCount)
}
