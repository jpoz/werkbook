package werkbook

import "sort"

type spillLookupIndex struct {
	anchors []rangeSpillAnchor
	rows    map[int][]spillLookupSpan
}

type spillLookupSpan struct {
	fromCol int
	toCol   int
	anchor  spillLookupAnchor
}

type spillLookupAnchor struct {
	col   int
	row   int
	cell  *Cell
	state *spillAnchorState
}

func (idx *spillLookupIndex) lookup(col, row int) *spillLookupSpan {
	if idx == nil || len(idx.rows) == 0 {
		return nil
	}
	spans := idx.rows[row]
	if len(spans) == 0 {
		return nil
	}
	i := sort.Search(len(spans), func(i int) bool {
		return spans[i].toCol >= col
	})
	if i >= len(spans) {
		return nil
	}
	span := &spans[i]
	if col < span.fromCol || col > span.toCol {
		return nil
	}
	return span
}
