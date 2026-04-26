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

func liveRefGrid(v Value) (Grid, bool) {
	if v.evalRef == nil {
		return nil, false
	}
	grid := v.evalRef.Materialized
	if grid == nil {
		grid = emptyRefGrid{
			rows: v.evalRef.ToRow - v.evalRef.FromRow + 1,
			cols: v.evalRef.ToCol - v.evalRef.FromCol + 1,
		}
	}
	return grid, true
}

// effectiveArrayBounds returns the logical size of an array value, preferring
// RangeOrigin when present so trimmed worksheet ranges still report their full
// extent.
func effectiveArrayBounds(v Value) (rows, cols int) {
	if grid, ok := liveRefGrid(v); ok {
		return grid.Rows(), grid.Cols()
	}
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
	if grid, ok := liveRefGrid(v); ok {
		if i < 0 || j < 0 || i >= rows || j >= cols {
			return ErrorVal(ErrValNA)
		}
		return EvalValueToValue(grid.Cell(i, j))
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

func arrayTopLeft(v Value) Value {
	if v.Type != ValueArray {
		return v
	}
	rows, cols := effectiveArrayBounds(v)
	if rows == 0 || cols == 0 {
		return ErrorVal(ErrValVALUE)
	}
	return arrayElementDirect(v, rows, cols, 0, 0)
}

func valueCellCount(v Value) int {
	rows, cols := arrayOpBoundsOrScalar(v)
	return rows * cols
}

func iterateValueElements(v Value, fn func(Value) bool) {
	if v.Type != ValueArray {
		fn(v)
		return
	}
	rows, cols := arrayOpBounds(v)
	for row := 0; row < rows; row++ {
		for col := 0; col < cols; col++ {
			if !fn(arrayElementDirect(v, rows, cols, row, col)) {
				return
			}
		}
	}
}

func valueFlattenRowMajor(v Value) []Value {
	values := make([]Value, 0, valueCellCount(v))
	iterateValueElements(v, func(cell Value) bool {
		values = append(values, cell)
		return true
	})
	return values
}

func valueProjectRowArray(v Value, row int) Value {
	if v.Type != ValueArray {
		if row != 0 {
			return ErrorVal(ErrValNA)
		}
		return Value{Type: ValueArray, Array: [][]Value{{v}}}
	}
	rows, cols := arrayOpBounds(v)
	if row < 0 || row >= rows {
		return ErrorVal(ErrValNA)
	}
	out := make([]Value, cols)
	for col := 0; col < cols; col++ {
		out[col] = arrayElementDirect(v, rows, cols, row, col)
	}
	return Value{Type: ValueArray, Array: [][]Value{out}}
}

func valueProjectColArray(v Value, col int) Value {
	if v.Type != ValueArray {
		if col != 0 {
			return ErrorVal(ErrValNA)
		}
		return Value{Type: ValueArray, Array: [][]Value{{v}}}
	}
	rows, cols := arrayOpBounds(v)
	if col < 0 || col >= cols {
		return ErrorVal(ErrValNA)
	}
	out := make([][]Value, rows)
	for row := 0; row < rows; row++ {
		out[row] = []Value{arrayElementDirect(v, rows, cols, row, col)}
	}
	return Value{Type: ValueArray, Array: out}
}

func iterateAlignedArgs(args []Value, fn func([]Value) bool) *Value {
	if len(args) == 0 {
		return nil
	}
	type arrayBounds struct {
		rows int
		cols int
	}
	bounds := make([]arrayBounds, len(args))
	rows, cols := arrayOpBoundsOrScalar(args[0])
	bounds[0] = arrayBounds{rows: rows, cols: cols}
	for i := 1; i < len(args); i++ {
		argRows, argCols := arrayOpBoundsOrScalar(args[i])
		if argRows != rows || argCols != cols {
			err := ErrorVal(ErrValVALUE)
			return &err
		}
		bounds[i] = arrayBounds{rows: argRows, cols: argCols}
	}
	cells := make([]Value, len(args))
	for row := 0; row < rows; row++ {
		for col := 0; col < cols; col++ {
			for i, arg := range args {
				cells[i] = arrayElementDirect(arg, bounds[i].rows, bounds[i].cols, row, col)
			}
			if !fn(cells) {
				return nil
			}
		}
	}
	return nil
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
	if grid, ok := liveRefGrid(v); ok {
		return grid.Rows(), grid.Cols()
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
		if grid, ok := liveRefGrid(v); ok {
			rows, cols := grid.Rows(), grid.Cols()
			if rows == 0 || cols == 0 {
				errVal := ErrorVal(ErrValVALUE)
				return arrayGrid{}, &errVal
			}
			return arrayGrid{
				cells:       materializeGridBounds(grid, rows, cols),
				rowCount:    rows,
				colCount:    cols,
				rangeOrigin: v.RangeOrigin,
			}, nil
		}
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

func (g arrayGrid) col(col int) []Value {
	values := make([]Value, g.rowCount)
	for row := 0; row < g.rowCount; row++ {
		values[row] = g.cell(row, col)
	}
	return values
}

func (g arrayGrid) flattenRowMajor() []Value {
	values := make([]Value, 0, g.rowCount*g.colCount)
	for row := 0; row < g.rowCount; row++ {
		values = append(values, g.row(row)...)
	}
	return values
}

func (g arrayGrid) projectRow(row int) Value {
	if row < 0 || row >= g.rowCount {
		return ErrorVal(ErrValNA)
	}
	if g.colCount <= 1 {
		return g.cell(row, 0)
	}
	return Value{Type: ValueArray, Array: [][]Value{g.row(row)}}
}

func (g arrayGrid) projectCol(col int) Value {
	if col < 0 || col >= g.colCount {
		return ErrorVal(ErrValNA)
	}
	if g.rowCount <= 1 {
		return g.cell(0, col)
	}
	values := g.col(col)
	rows := make([][]Value, len(values))
	for i, v := range values {
		rows[i] = []Value{v}
	}
	return Value{Type: ValueArray, Array: rows}
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
