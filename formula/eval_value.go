package formula

// EvalKind classifies the internal evaluator categories. It intentionally
// separates scalars, arrays, and refs while the public/test-facing surface
// continues to use Value.
type EvalKind uint8

const (
	EvalScalar EvalKind = iota
	EvalArray
	EvalRef
	EvalKindError
)

// SpillClass classifies whether a computed array is spillable and whether its
// extent is known up front.
type SpillClass uint8

const (
	SpillNone SpillClass = iota
	SpillBounded
	SpillUnbounded
	SpillScalarOnly
)

// EvalValue is the internal runtime envelope.
type EvalValue struct {
	Kind   EvalKind
	Scalar Value
	Array  *ArrayValue
	Ref    *RefValue
	Err    ErrorValue
}

// ArrayOrigin tracks provenance without claiming the array is still a live
// worksheet reference.
type ArrayOrigin struct {
	Range *RangeAddr
	Cell  *CellAddr
}

// Grid is the shared array access interface for evaluator paths.
type Grid interface {
	Rows() int
	Cols() int
	Cell(r, c int) EvalValue
	Iterate(fn func(r, c int, v EvalValue) bool)
}

// ArrayValue is the evaluator's first-class computed array representation.
type ArrayValue struct {
	Rows       int
	Cols       int
	Grid       Grid
	Origin     *ArrayOrigin
	SpillClass SpillClass
}

func evalError(err ErrorValue) EvalValue {
	return EvalValue{Kind: EvalKindError, Err: err}
}

func evalScalar(v Value) EvalValue {
	if v.Type == ValueError {
		return evalError(v.Err)
	}
	return EvalValue{Kind: EvalScalar, Scalar: v}
}

func evalArray(rows [][]Value, spillClass SpillClass) EvalValue {
	rowCount := len(rows)
	colCount := materializedArrayCols(rows)
	return EvalValue{
		Kind: EvalArray,
		Array: &ArrayValue{
			Rows:       rowCount,
			Cols:       colCount,
			Grid:       newLegacyValueGrid(rows),
			SpillClass: spillClass,
		},
	}
}

// RefValue is the evaluator's first-class worksheet reference representation.
type RefValue struct {
	Sheet        string
	SheetEnd     string
	FromCol      int
	FromRow      int
	ToCol        int
	ToRow        int
	Materialized Grid
	Legacy       *RefLegacyBoundary
}

// RefLegacyBoundary controls the two compatibility cases that still matter
// when a ref-backed EvalValue must cross the legacy Value boundary:
// preserving single-cell identity for ref-aware scalar consumers, and
// preserving placeholder shape for full-axis refs that cannot materialize at
// legacy-array size.
type RefLegacyBoundary struct {
	SingleCellValue Value
	PlaceholderRows int
	PlaceholderCols int
	UseEmptyArray   bool
}

func (r *RefValue) Bounds() RangeAddr {
	if r == nil {
		return RangeAddr{}
	}
	return RangeAddr{
		Sheet:    r.Sheet,
		SheetEnd: r.SheetEnd,
		FromCol:  r.FromCol,
		FromRow:  r.FromRow,
		ToCol:    r.ToCol,
		ToRow:    r.ToRow,
	}
}

func (r *RefValue) IterateCells(fn func(r, c int, v EvalValue) bool) {
	if r == nil {
		return
	}
	grid := r.Materialized
	if grid == nil {
		grid = emptyRefGrid{
			rows: r.ToRow - r.FromRow + 1,
			cols: r.ToCol - r.FromCol + 1,
		}
	}
	grid.Iterate(fn)
}

// ValueToEvalValue adapts the public Value surface into the internal evaluator
// categories.
func ValueToEvalValue(v Value) EvalValue {
	return valueToEvalValueWithResolver(v, nil)
}

func valueToEvalValueWithResolver(v Value, resolver CellResolver) EvalValue {
	if v.evalRef != nil {
		return EvalValue{Kind: EvalRef, Ref: cloneRefValue(v.evalRef)}
	}
	switch v.Type {
	case ValueError:
		return EvalValue{Kind: EvalKindError, Err: v.Err}
	case ValueArray:
		if v.RangeOrigin != nil && !v.NoSpill {
			ro := v.RangeOrigin
			return EvalValue{
				Kind: EvalRef,
				Ref: &RefValue{
					Sheet:    ro.Sheet,
					SheetEnd: ro.SheetEnd,
					FromCol:  ro.FromCol,
					FromRow:  ro.FromRow,
					ToCol:    ro.ToCol,
					ToRow:    ro.ToRow,
					// Keep direct-range arguments in ref form while allowing
					// migrated reducers to fall back to the legacy resolver.
					Materialized: newResolverRangeGrid(*ro, v.Array, resolver),
				},
			}
		}
		rows, cols := effectiveArrayBounds(v)
		return EvalValue{
			Kind: EvalArray,
			Array: &ArrayValue{
				Rows:       rows,
				Cols:       cols,
				Grid:       newLegacyValueGrid(v.Array),
				Origin:     arrayOriginFromLegacy(v),
				SpillClass: spillClassFromLegacy(v),
			},
		}
	case ValueRef:
		col := int(v.Num) % 100_000
		row := int(v.Num) / 100_000
		return EvalValue{
			Kind: EvalRef,
			Ref: &RefValue{
				Sheet:   v.Str,
				FromCol: col,
				FromRow: row,
				ToCol:   col,
				ToRow:   row,
			},
		}
	default:
		if v.CellOrigin != nil {
			cell := *v.CellOrigin
			return EvalValue{
				Kind: EvalRef,
				Ref: &RefValue{
					Sheet:        cell.Sheet,
					SheetEnd:     cell.SheetEnd,
					FromCol:      cell.Col,
					FromRow:      cell.Row,
					ToCol:        cell.Col,
					ToRow:        cell.Row,
					Materialized: newLegacyValueGrid([][]Value{{stripRefMetadata(v)}}),
					Legacy: &RefLegacyBoundary{
						SingleCellValue: v,
					},
				},
			}
		}
		return EvalValue{Kind: EvalScalar, Scalar: v}
	}
}

func legacyValueRef(v Value) (*RefValue, bool) {
	ev := ValueToEvalValue(v)
	if ev.Kind != EvalRef || ev.Ref == nil {
		return nil, false
	}
	return ev.Ref, true
}

func legacyArrayRef(v Value) (*RefValue, bool) {
	if v.Type != ValueArray {
		return nil, false
	}
	return legacyValueRef(v)
}

func legacyRefCellValue(ref *RefValue, rowOffset, colOffset int) Value {
	if ref == nil || rowOffset < 0 || colOffset < 0 {
		return ErrorVal(ErrValVALUE)
	}
	if ref.Materialized != nil {
		if rowOffset >= ref.Materialized.Rows() || colOffset >= ref.Materialized.Cols() {
			return ErrorVal(ErrValVALUE)
		}
		return EvalValueToValue(ref.Materialized.Cell(rowOffset, colOffset))
	}
	if rowOffset == 0 && colOffset == 0 && ref.Legacy != nil {
		return ref.Legacy.SingleCellValue
	}
	return ErrorVal(ErrValVALUE)
}

// EvalValueToValue adapts internal evaluator categories back to the Value
// surface used by the workbook, tests, and public APIs.
func EvalValueToValue(v EvalValue) Value {
	switch v.Kind {
	case EvalKindError:
		return ErrorVal(v.Err)
	case EvalScalar:
		return v.Scalar
	case EvalRef:
		if v.Ref == nil {
			return EmptyVal()
		}
		if v.Ref.SheetEnd == "" && v.Ref.FromCol == v.Ref.ToCol && v.Ref.FromRow == v.Ref.ToRow {
			if v.Ref.Legacy != nil && v.Ref.Legacy.SingleCellValue.CellOrigin != nil {
				out := v.Ref.Legacy.SingleCellValue
				out.evalRef = cloneRefValue(v.Ref)
				return out
			}
			out := Value{
				Type: ValueRef,
				Num:  float64(v.Ref.FromCol + v.Ref.FromRow*100_000),
				Str:  v.Ref.Sheet,
			}
			out.evalRef = cloneRefValue(v.Ref)
			return out
		}
		rows := v.Ref.ToRow - v.Ref.FromRow + 1
		cols := v.Ref.ToCol - v.Ref.FromCol + 1
		useEmptyPlaceholder := false
		if v.Ref.Legacy != nil {
			if v.Ref.Legacy.PlaceholderRows > 0 {
				rows = v.Ref.Legacy.PlaceholderRows
			}
			if v.Ref.Legacy.PlaceholderCols > 0 {
				cols = v.Ref.Legacy.PlaceholderCols
			}
			useEmptyPlaceholder = v.Ref.Legacy.UseEmptyArray
		}
		out := Value{
			Type:        ValueArray,
			RangeOrigin: ptrRange(v.Ref.Bounds()),
		}
		if useEmptyPlaceholder {
			out.Array = newValueMatrix(rows, cols)
		} else {
			grid := v.Ref.Materialized
			if grid == nil {
				grid = emptyRefGrid{rows: rows, cols: cols}
			}
			out.Array = materializeGridBounds(grid, rows, cols)
		}
		out.evalRef = cloneRefValue(v.Ref)
		return out
	case EvalArray:
		if v.Array == nil {
			return EmptyVal()
		}
		out := Value{
			Type:  ValueArray,
			Array: materializeGrid(v.Array.Grid),
		}
		if v.Array.Origin != nil {
			if v.Array.Origin.Range != nil {
				out.RangeOrigin = ptrRange(*v.Array.Origin.Range)
			}
			if v.Array.Origin.Cell != nil {
				out.CellOrigin = ptrCell(*v.Array.Origin.Cell)
			}
		}
		out.NoSpill = v.Array.SpillClass == SpillScalarOnly
		return out
	default:
		return EmptyVal()
	}
}

type legacyValueGrid struct {
	rows [][]Value
	r    int
	c    int
}

func newLegacyValueGrid(rows [][]Value) legacyValueGrid {
	maxCols := 0
	for _, row := range rows {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}
	return legacyValueGrid{rows: rows, r: len(rows), c: maxCols}
}

func (g legacyValueGrid) legacyRows() ([][]Value, int, int) {
	return g.rows, g.r, g.c
}

func (g legacyValueGrid) Rows() int { return g.r }
func (g legacyValueGrid) Cols() int { return g.c }
func (g legacyValueGrid) Cell(r, c int) EvalValue {
	if r < 0 || c < 0 || r >= g.r || c >= g.c {
		return EvalValue{Kind: EvalKindError, Err: ErrValNA}
	}
	if r < len(g.rows) && c < len(g.rows[r]) {
		return ValueToEvalValue(g.rows[r][c])
	}
	return EvalValue{Kind: EvalScalar, Scalar: EmptyVal()}
}
func (g legacyValueGrid) Iterate(fn func(r, c int, v EvalValue) bool) {
	for r := 0; r < g.r; r++ {
		for c := 0; c < g.c; c++ {
			if !fn(r, c, g.Cell(r, c)) {
				return
			}
		}
	}
}

type emptyRefGrid struct {
	rows int
	cols int
}

func (g emptyRefGrid) Rows() int { return g.rows }
func (g emptyRefGrid) Cols() int { return g.cols }
func (g emptyRefGrid) Cell(r, c int) EvalValue {
	if r < 0 || c < 0 || r >= g.rows || c >= g.cols {
		return EvalValue{Kind: EvalKindError, Err: ErrValNA}
	}
	return EvalValue{Kind: EvalScalar, Scalar: EmptyVal()}
}
func (g emptyRefGrid) Iterate(fn func(r, c int, v EvalValue) bool) {
	for r := 0; r < g.rows; r++ {
		for c := 0; c < g.cols; c++ {
			if !fn(r, c, EvalValue{Kind: EvalScalar, Scalar: EmptyVal()}) {
				return
			}
		}
	}
}

type resolverRangeGrid struct {
	bounds   RangeAddr
	resolver CellResolver
	rows     [][]Value
	loaded   bool
	r        int
	c        int
}

func newResolverRangeGrid(bounds RangeAddr, rows [][]Value, resolver CellResolver) Grid {
	g := &resolverRangeGrid{
		bounds:   bounds,
		resolver: resolver,
	}
	if rows != nil {
		g.setRows(rows)
		g.loaded = true
	}
	return g
}

func (g *resolverRangeGrid) legacyRows() ([][]Value, int, int) {
	if g == nil {
		return nil, 0, 0
	}
	g.ensureLoaded()
	return g.rows, g.r, g.c
}

func (g *resolverRangeGrid) Rows() int {
	g.ensureLoaded()
	return g.r
}

func (g *resolverRangeGrid) Cols() int {
	g.ensureLoaded()
	return g.c
}

func (g *resolverRangeGrid) Cell(r, c int) EvalValue {
	g.ensureLoaded()
	if r < 0 || c < 0 || r >= g.r || c >= g.c {
		return EvalValue{Kind: EvalKindError, Err: ErrValNA}
	}
	if r < len(g.rows) && c < len(g.rows[r]) {
		return ValueToEvalValue(g.rows[r][c])
	}
	return EvalValue{Kind: EvalScalar, Scalar: EmptyVal()}
}

func (g *resolverRangeGrid) Iterate(fn func(r, c int, v EvalValue) bool) {
	g.ensureLoaded()
	for r := 0; r < g.r; r++ {
		for c := 0; c < g.c; c++ {
			if !fn(r, c, g.Cell(r, c)) {
				return
			}
		}
	}
}

func (g *resolverRangeGrid) ensureLoaded() {
	if g.loaded {
		return
	}
	g.loaded = true
	if g.resolver == nil {
		return
	}
	g.setRows(g.resolver.GetRangeValues(g.bounds))
}

func (g *resolverRangeGrid) setRows(rows [][]Value) {
	g.rows = normalizeResolverRangeRows(g.bounds, rows)
	g.r = len(g.rows)
	g.c = materializedArrayCols(g.rows)
}

func normalizeResolverRangeRows(bounds RangeAddr, rows [][]Value) [][]Value {
	isFullCol := bounds.FromRow == 1 && bounds.ToRow >= maxRows
	isFullRow := bounds.FromCol == 1 && bounds.ToCol >= maxCols
	reachesMaxAxis := bounds.ToRow >= maxRows || bounds.ToCol >= maxCols
	if isFullCol || isFullRow || reachesMaxAxis {
		return rows
	}
	expectedRows := bounds.ToRow - bounds.FromRow + 1
	cols := bounds.ToCol - bounds.FromCol + 1
	if rows == nil {
		rows = make([][]Value, 0, expectedRows)
	}
	for len(rows) < expectedRows {
		emptyRow := make([]Value, cols)
		for j := range emptyRow {
			emptyRow[j] = EmptyVal()
		}
		rows = append(rows, emptyRow)
	}
	return rows
}

func arrayOriginFromLegacy(v Value) *ArrayOrigin {
	if v.RangeOrigin == nil && v.CellOrigin == nil {
		return nil
	}
	out := &ArrayOrigin{}
	if v.RangeOrigin != nil {
		out.Range = ptrRange(*v.RangeOrigin)
	}
	if v.CellOrigin != nil {
		out.Cell = ptrCell(*v.CellOrigin)
	}
	return out
}

func spillClassFromLegacy(v Value) SpillClass {
	if v.Type != ValueArray {
		return SpillNone
	}
	if v.NoSpill {
		return SpillScalarOnly
	}
	if v.RangeOrigin != nil {
		fullRows := v.RangeOrigin.FromRow == 1 && v.RangeOrigin.ToRow >= maxRows
		fullCols := v.RangeOrigin.FromCol == 1 && v.RangeOrigin.ToCol >= maxCols
		if fullRows || fullCols {
			return SpillUnbounded
		}
	}
	return SpillBounded
}

func materializeGrid(grid Grid) [][]Value {
	if grid == nil || grid.Rows() == 0 || grid.Cols() == 0 {
		return nil
	}
	return materializeGridBounds(grid, grid.Rows(), grid.Cols())
}

func materializeGridBounds(grid Grid, rows, cols int) [][]Value {
	if grid == nil || rows == 0 || cols == 0 {
		return nil
	}
	out := make([][]Value, rows)
	for r := 0; r < rows; r++ {
		row := make([]Value, cols)
		for c := 0; c < cols; c++ {
			row[c] = EvalValueToValue(grid.Cell(r, c))
		}
		out[r] = row
	}
	return out
}

func ptrRange(v RangeAddr) *RangeAddr {
	cp := v
	return &cp
}

func ptrCell(v CellAddr) *CellAddr {
	cp := v
	return &cp
}

func cloneRefValue(r *RefValue) *RefValue {
	if r == nil {
		return nil
	}
	cp := *r
	if r.Legacy != nil {
		legacy := *r.Legacy
		cp.Legacy = &legacy
	}
	return &cp
}

func stripRefMetadata(v Value) Value {
	out := v
	out.RangeOrigin = nil
	out.CellOrigin = nil
	out.FromCell = false
	out.evalRef = nil
	return out
}
