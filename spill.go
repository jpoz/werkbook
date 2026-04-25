package werkbook

import (
	"sort"

	"github.com/jpoz/werkbook/formula"
)

type spillOverlay struct {
	gen     uint64
	anchors map[cellKey]*spillAnchorState
	index   spillLookupIndex
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
}

func (s *Sheet) ensureSpillOverlay() *spillOverlay {
	if s.spill.gen == s.file.calcGen {
		return &s.spill
	}
	s.rebuildSpillOverlay()
	return &s.spill
}

func (s *Sheet) rebuildSpillOverlay() {
	if s.spill.anchors == nil {
		s.spill.anchors = make(map[cellKey]*spillAnchorState)
	}
	next := spillLookupIndex{
		rows: make(map[int][]spillLookupSpan),
	}
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
				if raw.Type != formula.ValueArray || raw.NoSpill {
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
				for rowOffset, spillRow := range raw.Array {
					if len(spillRow) == 0 {
						continue
					}
					rowNum := anchorRow + rowOffset
					fromCol := anchorCol
					if rowOffset == 0 {
						fromCol++
					}
					toCol := anchorCol + len(spillRow) - 1
					if fromCol > toCol {
						continue
					}
					next.rows[rowNum] = append(next.rows[rowNum], spillLookupSpan{
						fromCol: fromCol,
						toCol:   toCol,
						anchor:  anchor,
					})
				}
			}
		}
	}
	for row := range next.rows {
		spans := next.rows[row]
		sort.Slice(spans, func(i, j int) bool {
			if spans[i].fromCol == spans[j].fromCol {
				return spans[i].toCol < spans[j].toCol
			}
			return spans[i].fromCol < spans[j].fromCol
		})
		next.rows[row] = spans
	}
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
}

func (s *Sheet) refreshSpillAnchorsForPoint(col, row int) bool {
	refreshed := false
	for anchorRow, sheetRow := range s.rows {
		if anchorRow > row {
			continue
		}
		for anchorCol, cell := range sheetRow.cells {
			if anchorCol > col || !isDynamicArrayAnchor(cell) {
				continue
			}
			if state, ok := s.spillState(anchorCol, anchorRow); ok {
				if spillStateHasPublishedSpill(state, anchorCol, anchorRow, s.file.calcGen) &&
					row <= state.publishedToRow && col <= state.publishedToCol {
					continue
				}
				if toCol, toRow, ok := s.spillAttemptedRect(anchorCol, anchorRow); ok &&
					(row > toRow || col > toCol) {
					continue
				}
			}
			s.refreshSpillState(cell, anchorCol, anchorRow)
			refreshed = true
		}
	}
	return refreshed
}

func (s *Sheet) refreshSpillAnchorsForRange(req rangeMaterializationRequest, logicalToCol int) bool {
	refreshed := false
	for anchorRow, sheetRow := range s.rows {
		if anchorRow > req.toRow {
			continue
		}
		for anchorCol, cell := range sheetRow.cells {
			if anchorCol > logicalToCol || !isDynamicArrayAnchor(cell) {
				continue
			}
			if state, ok := s.spillState(anchorCol, anchorRow); ok {
				if spillStateHasPublishedSpill(state, anchorCol, anchorRow, s.file.calcGen) {
					if state.publishedToRow < req.fromRow || state.publishedToCol < req.fromCol {
						continue
					}
					continue
				}
				if toCol, toRow, ok := s.spillAttemptedRect(anchorCol, anchorRow); ok &&
					(toRow < req.fromRow || toCol < req.fromCol) {
					continue
				}
			}
			s.refreshSpillState(cell, anchorCol, anchorRow)
			refreshed = true
		}
	}
	return refreshed
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
