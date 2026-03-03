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

// DepGraph tracks formula-to-cell dependencies for incremental recalculation.
type DepGraph struct {
	// forward edges: formula cell → set of cells it reads
	dependsOn map[QualifiedCell]map[QualifiedCell]bool
	// reverse edges: data cell → set of formula cells that read it
	dependents map[QualifiedCell]map[QualifiedCell]bool
	// range subscriptions for containment checks
	rangeSubs []rangeSub
}

// NewDepGraph creates an empty dependency graph.
func NewDepGraph() *DepGraph {
	return &DepGraph{
		dependsOn:  make(map[QualifiedCell]map[QualifiedCell]bool),
		dependents: make(map[QualifiedCell]map[QualifiedCell]bool),
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
		r := rs.rng
		if cell.Sheet == r.Sheet &&
			cell.Col >= r.FromCol && cell.Col <= r.ToCol &&
			cell.Row >= r.FromRow && cell.Row <= r.ToRow {
			seen[rs.formulaCell] = true
			result = append(result, rs.formulaCell)
		}
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
