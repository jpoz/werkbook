package werkbook

import (
	"sort"

	"github.com/jpoz/werkbook/formula"
)

type spillOverlay struct {
	gen            uint64
	pendingGen     uint64
	anchors        map[cellKey]*spillAnchorState
	pendingRows    map[int][]pendingSpillAnchor
	pendingRowNums []int
	index          spillLookupIndex
}

type spillAnchorState struct {
	attemptedToCol, attemptedToRow int
	publishedToCol, publishedToRow int
	blocked                        bool
	gen                            uint64

	// Imported or last-known OOXML formula ref. The raw array value remains
	// formula-cache state on Cell; the overlay owns bounds and writer fallback.
	formulaRef string
}

type pendingSpillAnchor struct {
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
	rows, ok := spillArrayRows(raw)
	if !ok {
		return 0, 0, false
	}
	spillCols := 0
	for _, row := range rows {
		if len(row) > spillCols {
			spillCols = len(row)
		}
	}
	if len(rows) == 0 || spillCols == 0 {
		return 0, 0, false
	}
	return anchorCol + spillCols - 1, anchorRow + len(rows) - 1, true
}

func spillArrayRows(raw formula.Value) ([][]formula.Value, bool) {
	if raw.Type != formula.ValueArray {
		return nil, false
	}
	if raw.NoSpill {
		return nil, false
	}
	return raw.Array, true
}

func spillArrayCell(raw formula.Value, rowOffset, colOffset int) (formula.Value, bool) {
	rows, ok := spillArrayRows(raw)
	if !ok {
		return formula.Value{}, false
	}
	if rowOffset < 0 || rowOffset >= len(rows) || colOffset < 0 {
		return formula.Value{}, false
	}
	if colOffset >= len(rows[rowOffset]) {
		return formula.Value{}, false
	}
	return rows[rowOffset][colOffset], true
}

func (s *Sheet) invalidateSpillOverlay() {
	s.spill.gen = 0
	s.spill.pendingGen = 0
}

func (s *Sheet) ensureSpillOverlay() *spillOverlay {
	if s.spill.gen == s.file.calcGen {
		return &s.spill
	}
	s.rebuildSpillOverlay()
	return &s.spill
}

func (s *Sheet) ensurePendingSpillAnchors() {
	if s.spill.pendingGen == s.file.calcGen {
		return
	}
	if !s.hasPendingSpillAnchors() {
		s.spill.pendingRows = nil
		s.spill.pendingRowNums = nil
		s.spill.pendingGen = s.file.calcGen
		return
	}

	pendingRows := make(map[int][]pendingSpillAnchor)
	for anchorRow, sheetRow := range s.rows {
		for anchorCol, cell := range sheetRow.cells {
			if !isDynamicArrayAnchor(cell) {
				continue
			}
			if state, ok := s.spillState(anchorCol, anchorRow); ok && state.gen == s.file.calcGen {
				continue
			}
			pendingRows[anchorRow] = append(pendingRows[anchorRow], pendingSpillAnchor{
				col:  anchorCol,
				row:  anchorRow,
				cell: cell,
			})
		}
	}
	for row := range pendingRows {
		anchors := pendingRows[row]
		sort.Slice(anchors, func(i, j int) bool {
			return anchors[i].col < anchors[j].col
		})
		pendingRows[row] = anchors
	}
	pendingRowNums := make([]int, 0, len(pendingRows))
	for row := range pendingRows {
		pendingRowNums = append(pendingRowNums, row)
	}
	sort.Ints(pendingRowNums)
	s.spill.pendingRows = pendingRows
	s.spill.pendingRowNums = pendingRowNums
	s.spill.pendingGen = s.file.calcGen
}

func (s *Sheet) hasPendingSpillAnchors() bool {
	for anchorRow, sheetRow := range s.rows {
		for anchorCol, cell := range sheetRow.cells {
			if !isDynamicArrayAnchor(cell) {
				continue
			}
			if state, ok := s.spillState(anchorCol, anchorRow); ok && state.gen == s.file.calcGen {
				continue
			}
			return true
		}
	}
	return false
}

func (s *Sheet) markSpillAnchorResolved(col, row int) {
	if s.spill.pendingGen != s.file.calcGen {
		return
	}
	anchors := s.spill.pendingRows[row]
	if len(anchors) == 0 {
		return
	}
	idx := sort.Search(len(anchors), func(i int) bool {
		return anchors[i].col >= col
	})
	if idx >= len(anchors) || anchors[idx].col != col {
		return
	}
	copy(anchors[idx:], anchors[idx+1:])
	anchors = anchors[:len(anchors)-1]
	if len(anchors) == 0 {
		delete(s.spill.pendingRows, row)
		idx := sort.SearchInts(s.spill.pendingRowNums, row)
		if idx < len(s.spill.pendingRowNums) && s.spill.pendingRowNums[idx] == row {
			copy(s.spill.pendingRowNums[idx:], s.spill.pendingRowNums[idx+1:])
			s.spill.pendingRowNums = s.spill.pendingRowNums[:len(s.spill.pendingRowNums)-1]
		}
		return
	}
	s.spill.pendingRows[row] = anchors
}

func (s *Sheet) rebuildSpillOverlay() {
	if s.spill.anchors == nil {
		s.spill.anchors = make(map[cellKey]*spillAnchorState)
	}
	next := spillLookupIndex{}
	for anchorRow, sheetRow := range s.rows {
		for anchorCol, cell := range sheetRow.cells {
			if isDynamicArrayAnchor(cell) {
				state, ok := s.spillState(anchorCol, anchorRow)
				if !ok || !spillStateHasPublishedSpill(state, anchorCol, anchorRow, s.file.calcGen) {
					continue
				}
				if cell.rawCachedGen != s.file.calcGen {
					continue
				}
				raw := cell.rawValue
				if _, ok := spillArrayRows(raw); !ok {
					continue
				}
				anchor := spillLookupAnchor{
					col:   anchorCol,
					row:   anchorRow,
					cell:  cell,
					state: state,
				}
				next.anchors = append(next.anchors, rangeSpillAnchor{
					col:   anchorCol,
					row:   anchorRow,
					toCol: state.publishedToCol,
					toRow: state.publishedToRow,
					raw:   raw,
				})
				next.spans = append(next.spans, spillLookupSpan{
					fromCol: anchorCol,
					fromRow: anchorRow,
					toCol:   state.publishedToCol,
					toRow:   state.publishedToRow,
					anchor:  anchor,
				})
			}
		}
	}
	sort.Slice(next.spans, func(i, j int) bool {
		if next.spans[i].fromRow == next.spans[j].fromRow {
			return next.spans[i].fromCol < next.spans[j].fromCol
		}
		return next.spans[i].fromRow < next.spans[j].fromRow
	})
	s.spill.index = next
	s.spill.gen = s.file.calcGen
}

func (s *Sheet) ensureSpillState(col, row int) *spillAnchorState {
	if s.spill.anchors == nil {
		s.spill.anchors = make(map[cellKey]*spillAnchorState)
	}
	key := cellKey{sheet: s.name, col: col, row: row}
	state := s.spill.anchors[key]
	if state == nil {
		state = &spillAnchorState{}
		s.spill.anchors[key] = state
	}
	return state
}

func (s *Sheet) spillState(col, row int) (*spillAnchorState, bool) {
	if s.spill.anchors == nil {
		return nil, false
	}
	state, ok := s.spill.anchors[cellKey{sheet: s.name, col: col, row: row}]
	return state, ok
}

func (s *Sheet) setSpillFormulaRef(col, row int, formulaRef string) {
	if formulaRef == "" {
		return
	}
	state := s.ensureSpillState(col, row)
	state.formulaRef = formulaRef
}

func (s *Sheet) refreshSpillState(c *Cell, col, row int) {
	if !isDynamicArrayAnchor(c) {
		s.clearSpillState(col, row)
		return
	}
	if s.file.isEvaluatingCell(s.name, col, row) {
		return
	}
	s.resolveCell(c, col, row)
	var plan *SpillPlan
	if c.rawCachedGen != s.file.calcGen {
		outcome, err := s.evalCellOutcome(c, col, row)
		if err != nil {
			s.publishSpillState(col, row, nil)
			s.markSpillAnchorResolved(col, row)
			return
		}
		savedValue := c.value
		savedCachedGen := c.cachedGen
		savedDirty := c.dirty
		s.applyCellOutcome(c, col, row, outcome)
		plan = outcome.Spill
		c.value = savedValue
		c.cachedGen = savedCachedGen
		c.dirty = savedDirty
	} else {
		plan = newSpillPlan(c.rawValue, col, row)
		if c.value.Type == TypeError && c.value.String == "#SPILL!" && plan != nil {
			plan.Blocked = true
			plan.PublishedToCol = col
			plan.PublishedToRow = row
		}
	}
	s.publishSpillState(col, row, plan)
	s.markSpillAnchorResolved(col, row)
}

func (s *Sheet) refreshSpillAnchorsForPoint(col, row int) bool {
	s.ensurePendingSpillAnchors()
	refreshed := false
	s.forPendingSpillAnchorsThrough(col, row, func(anchor pendingSpillAnchor) {
		s.refreshSpillState(anchor.cell, anchor.col, anchor.row)
		refreshed = true
	})
	return refreshed
}

func (s *Sheet) refreshSpillAnchorsForRange(req rangeMaterializationRequest, logicalToCol int) bool {
	s.ensurePendingSpillAnchors()
	refreshed := false
	s.forPendingSpillAnchorsThrough(logicalToCol, req.toRow, func(anchor pendingSpillAnchor) {
		s.refreshSpillState(anchor.cell, anchor.col, anchor.row)
		refreshed = true
	})
	return refreshed
}

func (s *Sheet) forPendingSpillAnchorsThrough(maxCol, maxRow int, fn func(pendingSpillAnchor)) {
	if fn == nil || len(s.spill.pendingRowNums) == 0 {
		return
	}
	rowLimit := sort.Search(len(s.spill.pendingRowNums), func(i int) bool {
		return s.spill.pendingRowNums[i] > maxRow
	})
	for _, anchorRow := range s.spill.pendingRowNums[:rowLimit] {
		anchors := s.spill.pendingRows[anchorRow]
		colLimit := sort.Search(len(anchors), func(i int) bool {
			return anchors[i].col > maxCol
		})
		for _, anchor := range anchors[:colLimit] {
			fn(anchor)
		}
	}
}

func (s *Sheet) publishSpillState(col, row int, plan *SpillPlan) {
	state := s.ensureSpillState(col, row)
	attemptedToCol, attemptedToRow := col, row
	publishedToCol, publishedToRow := col, row
	blocked := false
	if plan != nil {
		attemptedToCol, attemptedToRow = plan.AttemptedToCol, plan.AttemptedToRow
		publishedToCol, publishedToRow = plan.PublishedToCol, plan.PublishedToRow
		blocked = plan.Blocked
	}
	if state.gen == s.file.calcGen &&
		state.attemptedToCol == attemptedToCol &&
		state.attemptedToRow == attemptedToRow &&
		state.publishedToCol == publishedToCol &&
		state.publishedToRow == publishedToRow &&
		state.blocked == blocked {
		s.setSpillBlockerDynamicRanges(col, row, state)
		return
	}
	state.gen = s.file.calcGen
	state.attemptedToCol = attemptedToCol
	state.attemptedToRow = attemptedToRow
	state.publishedToCol = publishedToCol
	state.publishedToRow = publishedToRow
	state.blocked = blocked
	s.setSpillBlockerDynamicRanges(col, row, state)
	s.invalidateSpillOverlay()
}

func (s *Sheet) clearSpillState(col, row int) {
	key := cellKey{sheet: s.name, col: col, row: row}
	if s.spill.anchors == nil {
		return
	}
	if _, ok := s.spill.anchors[key]; !ok {
		return
	}
	delete(s.spill.anchors, key)
	s.file.deps.SetDynamicRanges(
		formula.QualifiedCell{Sheet: s.name, Col: col, Row: row},
		formula.DynamicRangeKindSpillBlocker,
		nil,
	)
	s.markSpillAnchorResolved(col, row)
	s.invalidateSpillOverlay()
}

func (s *Sheet) setSpillBlockerDynamicRanges(anchorCol, anchorRow int, state *spillAnchorState) {
	qc := formula.QualifiedCell{Sheet: s.name, Col: anchorCol, Row: anchorRow}
	if state == nil || state.attemptedToCol <= anchorCol && state.attemptedToRow <= anchorRow {
		s.file.deps.SetDynamicRanges(qc, formula.DynamicRangeKindSpillBlocker, nil)
		return
	}
	s.file.deps.SetDynamicRanges(qc, formula.DynamicRangeKindSpillBlocker, []formula.RangeAddr{{
		Sheet:   s.name,
		FromCol: anchorCol,
		FromRow: anchorRow,
		ToCol:   state.attemptedToCol,
		ToRow:   state.attemptedToRow,
	}})
}

func spillStateHasPublishedSpill(state *spillAnchorState, anchorCol, anchorRow int, gen uint64) bool {
	if state == nil || state.gen != gen || state.blocked {
		return false
	}
	return state.publishedToCol > anchorCol || state.publishedToRow > anchorRow
}

func spillStateFormulaRef(state *spillAnchorState, anchorCol, anchorRow int) (string, bool) {
	if state == nil {
		return "", false
	}
	if !spillStateHasPublishedSpill(state, anchorCol, anchorRow, state.gen) {
		return "", false
	}
	endRef, err := CoordinatesToCellName(state.publishedToCol, state.publishedToRow)
	if err != nil {
		return "", false
	}
	anchorRef, err := CoordinatesToCellName(anchorCol, anchorRow)
	if err != nil {
		return "", false
	}
	return anchorRef + ":" + endRef, true
}

func (s *Sheet) spillAttemptedRect(anchorCol, anchorRow int) (toCol, toRow int, ok bool) {
	state, ok := s.spillState(anchorCol, anchorRow)
	if !ok || state == nil {
		return 0, 0, false
	}
	if state.gen != 0 && (state.attemptedToCol > anchorCol || state.attemptedToRow > anchorRow) {
		return state.attemptedToCol, state.attemptedToRow, true
	}
	if state.formulaRef == "" {
		return 0, 0, false
	}
	_, _, toCol, toRow, err := RangeToCoordinates(state.formulaRef)
	if err != nil {
		return 0, 0, false
	}
	if toCol <= anchorCol && toRow <= anchorRow {
		return 0, 0, false
	}
	return toCol, toRow, true
}
