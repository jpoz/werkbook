package formula

// QualifiedCell is a fully-qualified cell address (sheet is never empty).
type QualifiedCell struct {
	Sheet string
	Col   int
	Row   int
}

// rangeSub records that formulaCell depends on every cell in the given range.
type rangeSub struct {
	formulaCell QualifiedCell
	rng         RangeAddr // Sheet is always qualified
}

type dynamicRangeSub struct {
	formulaCell QualifiedCell
	kind        DynamicRangeKind
	rng         RangeAddr // Sheet is always qualified
}

// DynamicRangeKind distinguishes dynamic dependency edge families.
type DynamicRangeKind uint8

const (
	// DynamicRangeKindSpillBlocker tracks the attempted spill rectangle for
	// a dynamic array anchor so blockers can dirty it.
	DynamicRangeKindSpillBlocker DynamicRangeKind = iota
	// DynamicRangeKindMaterialized tracks grown ranges discovered while
	// materializing array results.
	DynamicRangeKindMaterialized
)

// DepGraph tracks formula-to-cell dependencies for incremental recalculation.
type DepGraph struct {
	// forward edges: formula cell → set of cells it reads
	dependsOn map[QualifiedCell]map[QualifiedCell]bool
	// reverse edges: data cell → set of formula cells that read it
	dependents map[QualifiedCell]map[QualifiedCell]bool
	// range subscriptions for containment checks
	rangeSubs []rangeSub
	// dynamic range subscriptions grouped by formula cell and kind.
	dynamicRanges map[QualifiedCell]map[DynamicRangeKind][]RangeAddr
	// dynamic range subscriptions indexed by the range's sheet for faster
	// point invalidation lookups.
	dynamicRangeSubsBySheet map[string][]dynamicRangeSub
}

// NewDepGraph creates an empty dependency graph.
func NewDepGraph() *DepGraph {
	return &DepGraph{
		dependsOn:               make(map[QualifiedCell]map[QualifiedCell]bool),
		dependents:              make(map[QualifiedCell]map[QualifiedCell]bool),
		dynamicRanges:           make(map[QualifiedCell]map[DynamicRangeKind][]RangeAddr),
		dynamicRangeSubsBySheet: make(map[string][]dynamicRangeSub),
	}
}

// Register records that formulaCell (on owningSheet) depends on the given
// cell refs and ranges. Unqualified refs (Sheet == "") are resolved to
// owningSheet. Calling Register again for the same formulaCell replaces
// the previous edges.
func (g *DepGraph) Register(formulaCell QualifiedCell, owningSheet string, refs []CellAddr, ranges []RangeAddr) {
	// Remove old edges first.
	g.Unregister(formulaCell)

	deps := make(map[QualifiedCell]bool)
	g.dependsOn[formulaCell] = deps

	// Point cell refs.
	for _, ref := range refs {
		sheet := ref.Sheet
		if sheet == "" {
			sheet = owningSheet
		}
		target := QualifiedCell{Sheet: sheet, Col: ref.Col, Row: ref.Row}
		deps[target] = true
		if g.dependents[target] == nil {
			g.dependents[target] = make(map[QualifiedCell]bool)
		}
		g.dependents[target][formulaCell] = true
	}

	// Range subscriptions.
	for _, rng := range ranges {
		sheet := rng.Sheet
		if sheet == "" {
			sheet = owningSheet
		}
		qualifiedRange := RangeAddr{
			Sheet:   sheet,
			FromCol: rng.FromCol,
			FromRow: rng.FromRow,
			ToCol:   rng.ToCol,
			ToRow:   rng.ToRow,
		}
		g.rangeSubs = append(g.rangeSubs, rangeSub{
			formulaCell: formulaCell,
			rng:         qualifiedRange,
		})
	}
}

// SetDynamicRanges replaces the dynamic ranges for formulaCell/kind without
// disturbing static dependencies or other dynamic kinds.
func (g *DepGraph) SetDynamicRanges(formulaCell QualifiedCell, kind DynamicRangeKind, ranges []RangeAddr) {
	if g.dynamicRanges == nil {
		g.dynamicRanges = make(map[QualifiedCell]map[DynamicRangeKind][]RangeAddr)
	}
	if g.dynamicRangeSubsBySheet == nil {
		g.dynamicRangeSubsBySheet = make(map[string][]dynamicRangeSub)
	}

	byKind := g.dynamicRanges[formulaCell]
	existing := byKind[kind]
	if len(ranges) == 0 {
		if len(existing) != 0 {
			g.removeDynamicRangeSubs(formulaCell, kind, existing)
		}
		if byKind != nil {
			delete(byKind, kind)
			if len(byKind) == 0 {
				delete(g.dynamicRanges, formulaCell)
			}
		}
		return
	}
	if byKind != nil && rangeAddrsEqual(existing, ranges) {
		return
	}
	if len(existing) != 0 {
		g.removeDynamicRangeSubs(formulaCell, kind, existing)
	}

	if byKind == nil {
		byKind = make(map[DynamicRangeKind][]RangeAddr)
		g.dynamicRanges[formulaCell] = byKind
	}
	copied := make([]RangeAddr, len(ranges))
	copy(copied, ranges)
	byKind[kind] = copied
	g.addDynamicRangeSubs(formulaCell, kind, copied)
}

func rangeAddrsEqual(a, b []RangeAddr) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

func (g *DepGraph) addDynamicRangeSubs(formulaCell QualifiedCell, kind DynamicRangeKind, ranges []RangeAddr) {
	if g.dynamicRangeSubsBySheet == nil {
		g.dynamicRangeSubsBySheet = make(map[string][]dynamicRangeSub)
	}
	for _, rng := range ranges {
		g.dynamicRangeSubsBySheet[rng.Sheet] = append(g.dynamicRangeSubsBySheet[rng.Sheet], dynamicRangeSub{
			formulaCell: formulaCell,
			kind:        kind,
			rng:         rng,
		})
	}
}

func (g *DepGraph) removeDynamicRangeSubs(formulaCell QualifiedCell, kind DynamicRangeKind, ranges []RangeAddr) {
	if g.dynamicRangeSubsBySheet == nil || len(ranges) == 0 {
		return
	}

	removeBySheet := make(map[string]map[RangeAddr]bool)
	for _, rng := range ranges {
		set := removeBySheet[rng.Sheet]
		if set == nil {
			set = make(map[RangeAddr]bool)
			removeBySheet[rng.Sheet] = set
		}
		set[rng] = true
	}

	for sheet, removeSet := range removeBySheet {
		subs := g.dynamicRangeSubsBySheet[sheet]
		if len(subs) == 0 {
			continue
		}
		n := 0
		for _, sub := range subs {
			if sub.formulaCell == formulaCell && sub.kind == kind && removeSet[sub.rng] {
				continue
			}
			subs[n] = sub
			n++
		}
		if n == 0 {
			delete(g.dynamicRangeSubsBySheet, sheet)
			continue
		}
		g.dynamicRangeSubsBySheet[sheet] = subs[:n]
	}
}

// Unregister removes all edges from formulaCell.
func (g *DepGraph) Unregister(formulaCell QualifiedCell) {
	// Remove point dependencies.
	if deps, ok := g.dependsOn[formulaCell]; ok {
		for target := range deps {
			if rev := g.dependents[target]; rev != nil {
				delete(rev, formulaCell)
				if len(rev) == 0 {
					delete(g.dependents, target)
				}
			}
		}
		delete(g.dependsOn, formulaCell)
	}

	// Remove range subscriptions.
	n := 0
	for _, rs := range g.rangeSubs {
		if rs.formulaCell != formulaCell {
			g.rangeSubs[n] = rs
			n++
		}
	}
	g.rangeSubs = g.rangeSubs[:n]

	if byKind := g.dynamicRanges[formulaCell]; byKind != nil {
		for kind, ranges := range byKind {
			g.removeDynamicRangeSubs(formulaCell, kind, ranges)
		}
	}
	delete(g.dynamicRanges, formulaCell)
}

// DirectDependents returns the immediate formula cells that depend on cell,
// either via a direct cell reference or because cell falls within a subscribed range.
func (g *DepGraph) DirectDependents(cell QualifiedCell) []QualifiedCell {
	seen := make(map[QualifiedCell]bool)
	var result []QualifiedCell

	// Point dependents.
	for dep := range g.dependents[cell] {
		if !seen[dep] {
			seen[dep] = true
			result = append(result, dep)
		}
	}

	// Range containment.
	for _, rs := range g.rangeSubs {
		if seen[rs.formulaCell] {
			continue
		}
		if rangeContainsCell(rs.rng, cell) {
			seen[rs.formulaCell] = true
			result = append(result, rs.formulaCell)
		}
	}

	for _, rs := range g.dynamicRangeSubsBySheet[cell.Sheet] {
		if seen[rs.formulaCell] {
			continue
		}
		if !rangeContainsCell(rs.rng, cell) {
			continue
		}
		seen[rs.formulaCell] = true
		result = append(result, rs.formulaCell)
	}

	return result
}

// DependsOn returns the point cells and ranges that cell's formula reads from.
// Returns nil slices if the cell has no registered dependencies.
func (g *DepGraph) DependsOn(cell QualifiedCell) (points []QualifiedCell, ranges []RangeAddr) {
	for target := range g.dependsOn[cell] {
		points = append(points, target)
	}
	for _, rs := range g.rangeSubs {
		if rs.formulaCell == cell {
			ranges = append(ranges, rs.rng)
		}
	}
	return points, ranges
}

func rangeContainsCell(rng RangeAddr, cell QualifiedCell) bool {
	return cell.Sheet == rng.Sheet &&
		cell.Col >= rng.FromCol && cell.Col <= rng.ToCol &&
		cell.Row >= rng.FromRow && cell.Row <= rng.ToRow
}

// Dependents returns all formula cells that transitively depend on changed,
// in BFS order (topological: dependencies before their dependents).
func (g *DepGraph) Dependents(changed QualifiedCell) []QualifiedCell {
	visited := make(map[QualifiedCell]bool)
	var result []QualifiedCell
	queue := []QualifiedCell{changed}
	visited[changed] = true

	for len(queue) > 0 {
		cur := queue[0]
		queue = queue[1:]

		for _, dep := range g.DirectDependents(cur) {
			if !visited[dep] {
				visited[dep] = true
				result = append(result, dep)
				queue = append(queue, dep)
			}
		}
	}

	return result
}
