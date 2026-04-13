package werkbook

import "github.com/jpoz/werkbook/formula"

type spillOverlay struct {
	gen     uint64
	anchors map[cellKey]*spillAnchorState
	refs    []spillAnchorRef
}

type spillAnchorRef struct {
	col   int
	row   int
	cell  *Cell
	state *spillAnchorState
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
	next := spillOverlay{
		gen:     s.file.calcGen,
		anchors: make(map[cellKey]*spillAnchorState),
	}
	for anchorRow, sheetRow := range s.rows {
		for anchorCol, cell := range sheetRow.cells {
			if isDynamicArrayAnchor(cell) {
				key := cellKey{sheet: s.name, col: anchorCol, row: anchorRow}
				var state *spillAnchorState
				if existing, ok := s.spill.anchors[key]; ok {
					cp := *existing
					state = &cp
					next.anchors[key] = state
				} else {
					state = &spillAnchorState{}
					next.anchors[key] = state
				}
				next.refs = append(next.refs, spillAnchorRef{
					col:   anchorCol,
					row:   anchorRow,
					cell:  cell,
					state: state,
				})
			}
		}
	}
	s.spill = next
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
	state := s.ensureSpillState(col, row)
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
