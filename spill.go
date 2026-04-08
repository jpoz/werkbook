package werkbook

import "github.com/jpoz/werkbook/formula"

type spillOverlay struct {
	gen     uint64
	anchors []spillAnchorRef
}

type spillAnchorRef struct {
	col  int
	row  int
	cell *Cell
}

func isDynamicArrayAnchor(c *Cell) bool {
	return c != nil && c.formula != "" && !c.isArrayFormula && c.dynamicArraySpill
}

func (f *File) isEvaluatingCell(sheet string, col, row int) bool {
	if f.evaluating == nil {
		return false
	}
	return f.evaluating[cellKey{sheet: sheet, col: col, row: row}]
}

func spillArrayRect(raw formula.Value, anchorCol, anchorRow int) (toCol, toRow int, ok bool) {
	if raw.Type != formula.ValueArray || raw.NoSpill {
		return 0, 0, false
	}
	spillCols := 0
	for _, row := range raw.Array {
		if len(row) > spillCols {
			spillCols = len(row)
		}
	}
	if len(raw.Array) == 0 || spillCols == 0 {
		return 0, 0, false
	}
	return anchorCol + spillCols - 1, anchorRow + len(raw.Array) - 1, true
}

func (s *Sheet) invalidateSpillOverlay() {
	s.spill.gen = 0
	s.spill.anchors = nil
}

func (s *Sheet) ensureSpillOverlay() *spillOverlay {
	if s.spill.gen == s.file.calcGen {
		return &s.spill
	}
	s.rebuildSpillOverlay()
	return &s.spill
}

func (s *Sheet) rebuildSpillOverlay() {
	next := spillOverlay{gen: s.file.calcGen}
	for anchorRow, sheetRow := range s.rows {
		for anchorCol, cell := range sheetRow.cells {
			if isDynamicArrayAnchor(cell) {
				next.anchors = append(next.anchors, spillAnchorRef{
					col:  anchorCol,
					row:  anchorRow,
					cell: cell,
				})
			}
		}
	}
	s.spill = next
}

func (s *Sheet) refreshSpillState(c *Cell, col, row int) {
	if !isDynamicArrayAnchor(c) {
		s.clearSpillState(c)
		return
	}
	if s.file.isEvaluatingCell(s.name, col, row) {
		return
	}
	s.resolveCell(c, col, row)
	raw := c.rawValue
	if c.rawCachedGen != s.file.calcGen {
		savedValue := c.value
		savedCachedGen := c.cachedGen
		savedDirty := c.dirty
		raw = s.evaluateFormulaRaw(c, col, row)
		c.value = savedValue
		c.cachedGen = savedCachedGen
		c.dirty = savedDirty
	}
	blocked := c.value.Type == TypeError && c.value.String == "#SPILL!"
	s.publishSpillState(c, col, row, raw, blocked)
}

func (s *Sheet) publishSpillState(c *Cell, col, row int, raw formula.Value, blocked bool) {
	if c == nil {
		return
	}
	attemptedToCol, attemptedToRow := col, row
	if toCol, toRow, ok := spillArrayRect(raw, col, row); ok {
		attemptedToCol, attemptedToRow = toCol, toRow
	}
	publishedToCol, publishedToRow := col, row
	if !blocked {
		if toCol, toRow, ok := spillArrayRect(raw, col, row); ok {
			publishedToCol, publishedToRow = toCol, toRow
		}
	}
	if c.spillStateGen == s.file.calcGen &&
		c.spillAttemptedToCol == attemptedToCol &&
		c.spillAttemptedToRow == attemptedToRow &&
		c.spillPublishedToCol == publishedToCol &&
		c.spillPublishedToRow == publishedToRow {
		return
	}
	c.spillStateGen = s.file.calcGen
	c.spillAttemptedToCol = attemptedToCol
	c.spillAttemptedToRow = attemptedToRow
	c.spillPublishedToCol = publishedToCol
	c.spillPublishedToRow = publishedToRow
	s.invalidateSpillOverlay()
}

func (s *Sheet) clearSpillState(c *Cell) {
	if c == nil {
		return
	}
	if c.spillStateGen == 0 &&
		c.spillAttemptedToCol == 0 &&
		c.spillAttemptedToRow == 0 &&
		c.spillPublishedToCol == 0 &&
		c.spillPublishedToRow == 0 {
		return
	}
	c.spillStateGen = 0
	c.spillAttemptedToCol = 0
	c.spillAttemptedToRow = 0
	c.spillPublishedToCol = 0
	c.spillPublishedToRow = 0
	s.invalidateSpillOverlay()
}

func cellHasPublishedSpill(c *Cell, anchorCol, anchorRow int, gen uint64) bool {
	if c == nil || c.spillStateGen != gen {
		return false
	}
	return c.spillPublishedToCol > anchorCol || c.spillPublishedToRow > anchorRow
}

func cellSpillFormulaRef(c *Cell, anchorCol, anchorRow int) (string, bool) {
	if !cellHasPublishedSpill(c, anchorCol, anchorRow, c.spillStateGen) {
		return "", false
	}
	endRef, err := CoordinatesToCellName(c.spillPublishedToCol, c.spillPublishedToRow)
	if err != nil {
		return "", false
	}
	anchorRef, err := CoordinatesToCellName(anchorCol, anchorRow)
	if err != nil {
		return "", false
	}
	return anchorRef + ":" + endRef, true
}

func (s *Sheet) spillAttemptedRect(c *Cell, anchorCol, anchorRow int) (toCol, toRow int, ok bool) {
	if c == nil {
		return 0, 0, false
	}
	if c.spillStateGen != 0 && (c.spillAttemptedToCol > anchorCol || c.spillAttemptedToRow > anchorRow) {
		return c.spillAttemptedToCol, c.spillAttemptedToRow, true
	}
	if c.formulaRef == "" {
		return 0, 0, false
	}
	_, _, toCol, toRow, err := RangeToCoordinates(c.formulaRef)
	if err != nil {
		return 0, 0, false
	}
	if toCol <= anchorCol && toRow <= anchorRow {
		return 0, 0, false
	}
	return toCol, toRow, true
}

func (s *Sheet) markOverlappingSpillAnchorsDirty(col, row int) {
	for anchorRow, sheetRow := range s.rows {
		for anchorCol, cell := range sheetRow.cells {
			if !isDynamicArrayAnchor(cell) {
				continue
			}
			if anchorCol == col && anchorRow == row {
				continue
			}
			toCol, toRow, ok := s.spillAttemptedRect(cell, anchorCol, anchorRow)
			if !ok {
				continue
			}
			if row < anchorRow || row > toRow || col < anchorCol || col > toCol {
				continue
			}
			cell.dirty = true
		}
	}
}
