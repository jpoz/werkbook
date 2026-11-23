package bridge

import (
	werkbook "github.com/jpoz/werkbook"
)

// ClearResult is the result of ClearRange.
type ClearResult struct {
	History HistoryState `json:"history"`
	// Changed is the number of cells that had a value or a formula.
	Changed int `json:"changed"`
	// Skipped is the number of read-only spill cells in the range.
	Skipped int `json:"skipped"`
}

// ClearRange removes the values and formulas of a range as one undo step.
// Styles stay. A spill follower is read-only: it is skipped, unless its anchor
// is in the range too.
func (s *Session) ClearRange(sheet string, rg Range) (ClearResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return ClearResult{}, err
	}
	rg, err = normRange(rg)
	if err != nil {
		return ClearResult{}, err
	}
	s.ensureCalculated()
	spills := s.spillRects(sh)

	var res ClearResult
	err = s.mutate(labelClear, func(rec *recorder) error {
		rec.setSheet(sheet)
		rec.setRange(rg)
		var err error
		res.Changed, res.Skipped, err = clearCells(rec, sh, rg, spills)
		return err
	})
	if err != nil {
		return ClearResult{}, err
	}
	res.History = s.History()
	return res, nil
}

// clearCells removes values and formulas in rg. It touches only stored cells,
// so a full-row or full-column range is cheap.
func clearCells(rec *recorder, sh *werkbook.Sheet, rg Range, spills []rect) (changed, skipped int, err error) {
	var targets []cellPos
	for row := range sh.Rows() {
		if row.Num() < rg.Row1 {
			continue
		}
		if row.Num() > rg.Row2 {
			break
		}
		for _, c := range row.Cells() {
			if c.Col() < rg.Col1 || c.Col() > rg.Col2 {
				continue
			}
			if c.Formula() == "" && c.Value().IsEmpty() {
				continue
			}
			targets = append(targets, cellPos{col: c.Col(), row: row.Num()})
		}
	}
	for _, pos := range targets {
		if err := rec.write(sh, pos.col, pos.row, sh.SnapshotCell(pos.col, pos.row).WithFormula("")); err != nil {
			return 0, 0, err
		}
	}
	return len(targets), countReadOnlyFollowers(spills, rg), nil
}

// countReadOnlyFollowers counts the spill followers in rg whose anchor is not
// in rg.
func countReadOnlyFollowers(spills []rect, rg Range) int {
	n := 0
	for _, sp := range spills {
		if rg.contains(sp.row1, sp.col1) {
			continue
		}
		rows := min(sp.row2, rg.Row2) - max(sp.row1, rg.Row1) + 1
		cols := min(sp.col2, rg.Col2) - max(sp.col1, rg.Col1) + 1
		if rows > 0 && cols > 0 {
			n += rows * cols
		}
	}
	return n
}

// followerAnchor returns the anchor of the published spill that has
// (row, col) as a follower.
func followerAnchor(spills []rect, row, col int) (anchor cellPos, ok bool) {
	for _, sp := range spills {
		if sp.contains(row, col) && (row != sp.row1 || col != sp.col1) {
			return cellPos{col: sp.col1, row: sp.row1}, true
		}
	}
	return cellPos{}, false
}

// sheetMerges returns the merged ranges of a sheet.
func sheetMerges(sh *werkbook.Sheet) []Range {
	var out []Range
	for _, mr := range sh.MergeCells() {
		col1, row1, col2, row2, err := werkbook.RangeToCoordinates(mr.Start + ":" + mr.End)
		if err != nil {
			continue
		}
		out = append(out, Range{Row1: row1, Col1: col1, Row2: row2, Col2: col2})
	}
	return out
}

func intersectsMerge(merges []Range, rg Range) bool {
	for _, m := range merges {
		if m.intersects(rg) {
			return true
		}
	}
	return false
}

// usedRange returns the last used row and column of a sheet, with the cells
// that dynamic array formulas spill into. The caller must call
// ensureCalculated first.
func (s *Session) usedRange(sh *werkbook.Sheet) (maxRow, maxCol int) {
	maxCol = sh.MaxCol()
	maxRow = lastRowWithCells(sh, maxCol)
	for _, sp := range s.spillRects(sh) {
		maxRow = max(maxRow, sp.row2)
		maxCol = max(maxCol, sp.col2)
	}
	return maxRow, maxCol
}

// lastRowWithCells returns the last row that has a stored cell. A row that
// has only a custom height is not a used row: Sheet.MaxRow counts it, but it
// must not make the used range larger (the limit for a copy of full columns,
// for formats and for Ctrl+End).
func lastRowWithCells(sh *werkbook.Sheet, maxCol int) int {
	last := sh.MaxRow()
	if last == 0 || maxCol == 0 {
		return 0
	}
	// The usual case: the last row has a cell.
	for col := 1; col <= maxCol; col++ {
		if sh.SnapshotCell(col, last).Exists() {
			return last
		}
	}
	var rows []*werkbook.Row
	for row := range sh.Rows() {
		rows = append(rows, row)
	}
	for i := len(rows) - 1; i >= 0; i-- {
		if len(rows[i].Cells()) > 0 {
			return rows[i].Num()
		}
	}
	return 0
}

// sourceSnapshot returns the cell state that a copy or a fill takes from a
// source cell. A spill follower has no stored value; it gives its shown
// value, with the style of its stored placeholder cell.
func sourceSnapshot(sh *werkbook.Sheet, spills []rect, row, col int) werkbook.CellSnapshot {
	snap := sh.SnapshotCell(col, row)
	if !snap.IsEmpty() {
		return snap
	}
	if _, ok := followerAnchor(spills, row, col); !ok {
		return snap
	}
	ref, err := werkbook.CoordinatesToCellName(col, row)
	if err != nil {
		return snap
	}
	v, err := sh.GetValue(ref)
	if err != nil || v.IsEmpty() {
		return snap
	}
	return snap.WithValue(v)
}
