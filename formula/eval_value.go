package formula

// EvalKind classifies the internal evaluator categories described in the v2
// architecture plan. It intentionally separates scalars, arrays, and refs even
// while the public/test-facing surface continues to use Value.
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

// EvalValue is the internal runtime envelope used by the v2 migration.
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

// Grid is the shared array access interface for the new evaluator model.
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

// RefValue is the evaluator's first-class worksheet reference representation.
type RefValue struct {
	Sheet        string
	FromCol      int
	FromRow      int
	ToCol        int
	ToRow        int
	Materialized Grid
}

func (r *RefValue) Bounds() RangeAddr {
	if r == nil {
		return RangeAddr{}
	}
	return RangeAddr{
		Sheet:   r.Sheet,
		FromCol: r.FromCol,
		FromRow: r.FromRow,
		ToCol:   r.ToCol,
		ToRow:   r.ToRow,
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

// ValueToEvalValue adapts the legacy Value model into the new internal v2
// categories so migration slices can be tested without a flag day.
func ValueToEvalValue(v Value) EvalValue {
	return valueToEvalValueWithResolver(v, nil)
}

func valueToEvalValueWithResolver(v Value, resolver CellResolver) EvalValue {
	switch v.Type {
	case ValueError:
		return EvalValue{Kind: EvalKindError, Err: v.Err}
	case ValueArray:
		if v.RangeOrigin != nil {
			ro := v.RangeOrigin
			return EvalValue{
				Kind: EvalRef,
				Ref: &RefValue{
					Sheet:   ro.Sheet,
					FromCol: ro.FromCol,
					FromRow: ro.FromRow,
					ToCol:   ro.ToCol,
					ToRow:   ro.ToRow,
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
		return EvalValue{Kind: EvalScalar, Scalar: v}
	}
}

// EvalValueToValue adapts the new internal v2 categories back to the legacy
// Value surface used by today's workbook, tests, and public APIs.
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
		if v.Ref.FromCol == v.Ref.ToCol && v.Ref.FromRow == v.Ref.ToRow {
			return Value{
				Type: ValueRef,
				Num:  float64(v.Ref.FromCol + v.Ref.FromRow*100_000),
				Str:  v.Ref.Sheet,
			}
		}
		rows := v.Ref.ToRow - v.Ref.FromRow + 1
		cols := v.Ref.ToCol - v.Ref.FromCol + 1
		grid := v.Ref.Materialized
		if grid == nil {
			grid = emptyRefGrid{rows: rows, cols: cols}
		}
		out := Value{
			Type:        ValueArray,
			Array:       materializeGrid(grid),
			RangeOrigin: ptrRange(v.Ref.Bounds()),
		}
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
	if rows == nil {
		return nil
	}
	isFullCol := bounds.FromRow == 1 && bounds.ToRow >= maxRows
	isFullRow := bounds.FromCol == 1 && bounds.ToCol >= maxCols
	reachesMaxAxis := bounds.ToRow >= maxRows || bounds.ToCol >= maxCols
	if isFullCol || isFullRow || reachesMaxAxis {
		return rows
	}
	expectedRows := bounds.ToRow - bounds.FromRow + 1
	cols := bounds.ToCol - bounds.FromCol + 1
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
	rows := make([][]Value, grid.Rows())
	for r := 0; r < grid.Rows(); r++ {
		row := make([]Value, grid.Cols())
		for c := 0; c < grid.Cols(); c++ {
			row[c] = EvalValueToValue(grid.Cell(r, c))
		}
		rows[r] = row
	}
	return rows
}

func ptrRange(v RangeAddr) *RangeAddr {
	cp := v
	return &cp
}

func ptrCell(v CellAddr) *CellAddr {
	cp := v
	return &cp
}
