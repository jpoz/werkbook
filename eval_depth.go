package werkbook

import "github.com/jpoz/werkbook/formula"

// maxEvalRecursionDepth is the number of nested formula evaluations that the
// engine permits on the Go stack. Evaluation is recursive: a formula reads a
// precedent cell, and a stale precedent is evaluated in a nested call. Each
// level uses tens of kilobytes of stack, so a long dependency chain (for
// example a running total where each row reads the row above) can exceed the
// 1 GB goroutine stack limit. That failure is a fatal error that a caller
// cannot recover from.
//
// At this depth the engine stops the recursion and resolves the remaining
// stale precedents with an explicit stack. See resolvePrecedentsIteratively.
const maxEvalRecursionDepth = 256

// pendingEval is one entry on the explicit evaluation stack.
type pendingEval struct {
	sheet    *Sheet
	cell     *Cell
	col, row int
	// expanded is true after the stale precedents of the cell are on the
	// stack. When the entry comes back to the top, all of them are resolved.
	expanded bool
}

// needsEval reports whether the cell has a formula whose cached value is not
// current for the given calculation generation.
func (c *Cell) needsEval(calcGen uint64) bool {
	return c.formula != "" && (c.dirty || c.cachedGen < calcGen)
}

// resolvePrecedentsIteratively evaluates the stale static precedents of the
// given formula cell, and their precedents, in dependency order. It uses a
// heap-allocated stack, so the length of the dependency chain does not
// increase the depth of the Go stack. When it returns, the caller can
// evaluate the cell and find its static precedents cached.
//
// Only precedents known at compile time (cell references and ranges) are
// resolved here. References that a formula computes at run time, such as
// INDIRECT or OFFSET, are still evaluated by recursion.
//
// The caller must have the root cell in f.evaluating. Cells that are in
// f.evaluating or that are already expanded on the explicit stack are part of
// a reference cycle. They are skipped, and the usual recursive evaluation
// detects the cycle.
func (f *File) resolvePrecedentsIteratively(s *Sheet, c *Cell, col, row int) {
	expanded := make(map[cellKey]bool)
	stack := f.pushStalePrecedents(nil, s, c, col, row, expanded)

	for len(stack) > 0 {
		top := &stack[len(stack)-1]
		if !top.cell.needsEval(f.calcGen) {
			// Another dependent resolved this cell first.
			stack = stack[:len(stack)-1]
			continue
		}
		key := cellKey{sheet: top.sheet.name, col: top.col, row: top.row}
		if !top.expanded {
			top.expanded = true
			expanded[key] = true
			// The append can move the backing array, so top is not valid
			// after this call.
			stack = f.pushStalePrecedents(stack, top.sheet, top.cell, top.col, top.row, expanded)
			continue
		}
		entry := *top
		stack = stack[:len(stack)-1]
		delete(expanded, key)
		entry.sheet.resolveCell(entry.cell, entry.col, entry.row)
	}
}

// pushStalePrecedents appends the stale formula cells that the given cell
// reads directly. It skips cells that are in a reference cycle with the cells
// now in evaluation.
func (f *File) pushStalePrecedents(
	stack []pendingEval,
	s *Sheet,
	c *Cell,
	col, row int,
	expanded map[cellKey]bool,
) []pendingEval {
	cf, err := s.compileCellFormula(c, col, row)
	if err != nil {
		// The evaluation of this cell reports the error.
		return stack
	}

	push := func(ps *Sheet, pc *Cell, pcol, prow int) {
		if !pc.needsEval(f.calcGen) {
			return
		}
		key := cellKey{sheet: ps.name, col: pcol, row: prow}
		if expanded[key] || f.evaluating[key] {
			return
		}
		stack = append(stack, pendingEval{sheet: ps, cell: pc, col: pcol, row: prow})
	}

	for _, ref := range cf.Refs {
		ps := f.precedentSheet(s, ref.Sheet, ref.SheetEnd)
		if ps == nil {
			continue
		}
		r, ok := ps.rows[ref.Row]
		if !ok {
			continue
		}
		if pc, ok := r.cells[ref.Col]; ok {
			push(ps, pc, ref.Col, ref.Row)
		}
	}
	for _, rng := range cf.Ranges {
		ps := f.precedentSheet(s, rng.Sheet, rng.SheetEnd)
		if ps == nil {
			continue
		}
		ps.forEachCellInRange(rng, func(pc *Cell, pcol, prow int) {
			push(ps, pc, pcol, prow)
		})
	}
	return stack
}

// precedentSheet returns the sheet for a compiled reference. An empty name
// refers to the sheet of the formula. 3D references are not resolved here;
// the recursive evaluation handles them.
func (f *File) precedentSheet(current *Sheet, name, sheetEnd string) *Sheet {
	if sheetEnd != "" {
		return nil
	}
	if name == "" {
		return current
	}
	return f.Sheet(name)
}

// forEachCellInRange calls fn for each stored cell in the range. It iterates
// the smaller of the range and the stored rows or cells, so a full-column or
// full-row range costs no more than a scan of the stored data.
func (s *Sheet) forEachCellInRange(rng formula.RangeAddr, fn func(c *Cell, col, row int)) {
	visitRow := func(r *Row) {
		if rng.ToCol-rng.FromCol < len(r.cells) {
			for col := rng.FromCol; col <= rng.ToCol; col++ {
				if c, ok := r.cells[col]; ok {
					fn(c, col, r.num)
				}
			}
			return
		}
		for col, c := range r.cells {
			if col >= rng.FromCol && col <= rng.ToCol {
				fn(c, col, r.num)
			}
		}
	}

	if rng.ToRow-rng.FromRow < len(s.rows) {
		for row := rng.FromRow; row <= rng.ToRow; row++ {
			if r, ok := s.rows[row]; ok {
				visitRow(r)
			}
		}
		return
	}
	for row, r := range s.rows {
		if row >= rng.FromRow && row <= rng.ToRow {
			visitRow(r)
		}
	}
}
