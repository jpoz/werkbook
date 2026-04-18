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

// arrayElementDirect returns element [i][j] from an array Value using
// precomputed bounds, avoiding repeated O(rows) scans of materializedArrayCols.
func arrayElementDirect(v Value, rows, cols, i, j int) Value {
	if v.Type != ValueArray {
		return v
	}
	if i < 0 || j < 0 || i >= rows || j >= cols {
		return ErrorVal(ErrValNA)
	}
	if i < len(v.Array) && j < len(v.Array[i]) {
		return v.Array[i][j]
	}
	if v.RangeOrigin != nil {
		return EmptyVal()
	}
	return ErrorVal(ErrValNA)
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

// arrayOpBoundsOrScalar returns precomputed array bounds, or (1,1) for scalars.
// Used to avoid repeated O(rows) scans in hot element-access loops.
func arrayOpBoundsOrScalar(v Value) (rows, cols int) {
	if v.Type != ValueArray {
		return 1, 1
	}
	return arrayOpBounds(v)
}

func usesFullSheetAxisRange(v Value) bool {
	if v.RangeOrigin == nil {
		return false
	}
	// Check whether the range reaches the maximum row or column boundary,
	// regardless of where it starts.  Ranges like A2:A1048576 are
	// semantically equivalent to full-column references (A:A) for the
	// purpose of keeping intermediate arrays small; the previous check
	// (ToRow-FromRow+1 >= maxRows) missed ranges starting after row 1.
	return v.RangeOrigin.ToRow >= maxRows ||
		v.RangeOrigin.ToCol >= maxCols
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
