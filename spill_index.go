package werkbook

import (
	"sort"

	"github.com/jpoz/werkbook/formula"
)

type spillLookupIndex struct {
	anchors []rangeSpillAnchor
	spans   []spillLookupSpan
}

type spillLookupSpan struct {
	fromCol int
	fromRow int
	toCol   int
	toRow   int
	anchor  spillLookupAnchor
}

type spillLookupAnchor struct {
	col   int
	row   int
	cell  *Cell
	state *spillAnchorState
}

func (idx *spillLookupIndex) lookup(col, row int) *spillLookupSpan {
	if idx == nil || len(idx.spans) == 0 {
		return nil
	}
	limit := sort.Search(len(idx.spans), func(i int) bool {
		return idx.spans[i].fromRow > row
	})
	for i := 0; i < limit; i++ {
		span := &idx.spans[i]
		if row < span.fromRow || row > span.toRow {
			continue
		}
		if col < span.fromCol || col > span.toCol {
			continue
		}
		if col == span.anchor.col && row == span.anchor.row {
			continue
		}
		raw := span.anchor.cell.rawValue
		if raw.Type != formula.ValueArray || raw.NoSpill {
			continue
		}
		rowOffset := row - span.anchor.row
		colOffset := col - span.anchor.col
		if rowOffset < 0 || rowOffset >= len(raw.Array) || colOffset < 0 {
			continue
		}
		if colOffset >= len(raw.Array[rowOffset]) {
			continue
		}
		return span
	}
	return nil
}

func (idx *spillLookupIndex) Anchors(string) []rangeSpillAnchor {
	if idx == nil {
		return nil
	}
	return idx.anchors
}
