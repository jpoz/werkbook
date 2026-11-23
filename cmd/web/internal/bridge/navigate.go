package bridge

import (
	"errors"

	werkbook "github.com/jpoz/werkbook"
)

// EdgeResult is the result of FindEdge.
type EdgeResult struct {
	Row int `json:"row"`
	Col int `json:"col"`
}

// RegionResult is the result of CurrentRegion.
type RegionResult struct {
	Range Range `json:"range"`
	// Header is true when the first row of the region looks like a header.
	Header bool `json:"header"`
}

// occupancy answers "is this cell empty" for navigation. A cell is not empty
// when it has a value or a formula (also a formula that gives ""), or when it
// is a spill follower with a value. A cell with only a style is empty. A
// merged range is one cell: all its cells have the state of its top-left
// cell.
type occupancy struct {
	sh             *werkbook.Sheet
	spills         []rect
	merges         []Range
	maxRow, maxCol int
}

func (s *Session) occupancy(sh *werkbook.Sheet) occupancy {
	s.ensureCalculated()
	o := occupancy{sh: sh, spills: s.spillRects(sh), merges: sheetMerges(sh)}
	o.maxRow, o.maxCol = s.usedRange(sh)
	return o
}

// block returns the merged range that contains the cell, or the cell itself.
func (o occupancy) block(row, col int) Range {
	for _, m := range o.merges {
		if m.contains(row, col) {
			return m
		}
	}
	return Range{Row1: row, Col1: col, Row2: row, Col2: col}
}

// filled reports whether the cell is not empty. For a cell in a merged range
// it reports the state of the top-left cell of the range.
func (o occupancy) filled(row, col int) bool {
	b := o.block(row, col)
	row, col = b.Row1, b.Col1
	if row > o.maxRow || col > o.maxCol {
		return false // out of the used range
	}
	if !o.sh.SnapshotCell(col, row).IsEmpty() {
		return true
	}
	if _, ok := followerAnchor(o.spills, row, col); !ok {
		return false
	}
	ref, err := werkbook.CoordinatesToCellName(col, row)
	if err != nil {
		return false
	}
	v, err := o.sh.GetValue(ref)
	return err == nil && !v.IsEmpty()
}

// FindEdge returns the cell that Ctrl+arrow goes to from (row, col). Exactly
// one of dRow and dCol is -1 or 1. See cmd/web/README.md for the rules.
func (s *Session) FindEdge(sheet string, row, col, dRow, dCol int) (EdgeResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return EdgeResult{}, err
	}
	if _, err := cellRef(row, col); err != nil {
		return EdgeResult{}, err
	}
	vertical := (dRow == 1 || dRow == -1) && dCol == 0
	horizontal := (dCol == 1 || dCol == -1) && dRow == 0
	if !vertical && !horizontal {
		return EdgeResult{}, errors.New("findEdge: exactly one of dRow and dCol must be -1 or 1")
	}
	o := s.occupancy(sh)

	// The scan moves from block to block along the row or column of the
	// start cell. next returns the first cell after a block, in the direction
	// of travel.
	next := func(b Range) (r, c int, ok bool) {
		switch {
		case dRow > 0:
			r, c = b.Row2+1, col
		case dRow < 0:
			r, c = b.Row1-1, col
		case dCol > 0:
			r, c = row, b.Col2+1
		default:
			r, c = row, b.Col1-1
		}
		return r, c, r >= 1 && r <= werkbook.MaxRows && c >= 1 && c <= werkbook.MaxColumns
	}
	result := func(b Range) EdgeResult { return EdgeResult{Row: b.Row1, Col: b.Col1} }

	cur := o.block(row, col)
	r, c, ok := next(cur)
	if !ok {
		return EdgeResult{Row: row, Col: col}, nil // already at the edge of the sheet
	}

	// Rule 1: in a run of non-empty cells, go to its last cell.
	if o.filled(row, col) && o.filled(r, c) {
		for ok && o.filled(r, c) {
			cur = o.block(r, c)
			r, c, ok = next(cur)
		}
		return result(cur), nil
	}

	// Rule 2: go to the next non-empty cell.
	for ok {
		if o.filled(r, c) {
			return result(o.block(r, c)), nil
		}
		// All cells after the used range are empty: stop a scan that leaves
		// it, and start a scan toward it at its border.
		switch {
		case dRow > 0 && r > o.maxRow, dCol > 0 && c > o.maxCol:
			ok = false
			continue
		case dRow < 0 && r > o.maxRow+1:
			r = o.maxRow + 1
		case dCol < 0 && c > o.maxCol+1:
			c = o.maxCol + 1
		}
		r, c, ok = next(o.block(r, c))
	}

	// Rule 3: the edge of the sheet.
	switch {
	case dRow > 0:
		return EdgeResult{Row: werkbook.MaxRows, Col: col}, nil
	case dRow < 0:
		return EdgeResult{Row: 1, Col: col}, nil
	case dCol > 0:
		return EdgeResult{Row: row, Col: werkbook.MaxColumns}, nil
	default:
		return EdgeResult{Row: row, Col: 1}, nil
	}
}

// CurrentRegion returns the smallest rectangle around a cell that has an
// empty border on all sides (the Ctrl+A region of a spreadsheet), and whether
// its first row looks like a header.
func (s *Session) CurrentRegion(sheet string, row, col int) (RegionResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return RegionResult{}, err
	}
	if _, err := cellRef(row, col); err != nil {
		return RegionResult{}, err
	}
	o := s.occupancy(sh)
	rg := o.block(row, col)

	// anyFilled reports whether a row or column strip next to the range has
	// a non-empty cell. The strip includes the corners.
	rowFilled := func(r int) bool {
		if r < 1 || r > werkbook.MaxRows {
			return false
		}
		for c := max(1, rg.Col1-1); c <= min(werkbook.MaxColumns, rg.Col2+1); c++ {
			if o.filled(r, c) {
				return true
			}
		}
		return false
	}
	colFilled := func(c int) bool {
		if c < 1 || c > werkbook.MaxColumns {
			return false
		}
		for r := max(1, rg.Row1-1); r <= min(werkbook.MaxRows, rg.Row2+1); r++ {
			if o.filled(r, c) {
				return true
			}
		}
		return false
	}
	for grown := true; grown; {
		grown = false
		for rowFilled(rg.Row1 - 1) {
			rg.Row1--
			grown = true
		}
		for rowFilled(rg.Row2 + 1) {
			rg.Row2++
			grown = true
		}
		for colFilled(rg.Col1 - 1) {
			rg.Col1--
			grown = true
		}
		for colFilled(rg.Col2 + 1) {
			rg.Col2++
			grown = true
		}
	}

	return RegionResult{Range: rg, Header: o.hasHeader(rg)}, nil
}

// hasHeader reports whether the first row of a range looks like a header:
// it has text and only text, and the second row has a cell that is not text.
func (o occupancy) hasHeader(rg Range) bool {
	if rg.rows() < 2 {
		return false
	}
	valueType := func(row, col int) werkbook.ValueType {
		ref, err := werkbook.CoordinatesToCellName(col, row)
		if err != nil {
			return werkbook.TypeEmpty
		}
		v, err := o.sh.GetValue(ref)
		if err != nil {
			return werkbook.TypeEmpty
		}
		return v.Type
	}
	text := 0
	for col := rg.Col1; col <= rg.Col2; col++ {
		switch valueType(rg.Row1, col) {
		case werkbook.TypeString:
			text++
		case werkbook.TypeEmpty:
		default:
			return false
		}
	}
	if text == 0 {
		return false
	}
	for col := rg.Col1; col <= rg.Col2; col++ {
		if t := valueType(rg.Row1+1, col); t != werkbook.TypeEmpty && t != werkbook.TypeString {
			return true
		}
	}
	return false
}
