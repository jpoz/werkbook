package bridge

import (
	"errors"
	"fmt"
	"sort"

	werkbook "github.com/jpoz/werkbook"
	"github.com/jpoz/werkbook/formula"
)

const labelSort = "Sort"

// SortKey is one sort key: a sheet column of the range and a direction.
type SortKey struct {
	Col  int  `json:"col"`
	Desc bool `json:"desc"`
}

// SortResult is the result of SortRange.
type SortResult struct {
	History HistoryState `json:"history"`
	// Changed is the number of cells whose content changed.
	Changed int `json:"changed"`
	// Range is the range that was sorted, with the header row.
	Range Range `json:"range"`
}

// Type ranks of the sort order of a spreadsheet: numbers, text, booleans,
// errors. Empty cells are not ranked: they are always last.
const (
	sortRankNumber = iota
	sortRankText
	sortRankBool
	sortRankError
)

func sortRank(v werkbook.Value) int {
	switch v.Type {
	case werkbook.TypeNumber:
		return sortRankNumber
	case werkbook.TypeString:
		return sortRankText
	case werkbook.TypeBool:
		return sortRankBool
	default:
		return sortRankError
	}
}

// compareSortValues compares two key values for an ascending or descending
// key. Empty cells are last in both directions. Errors are equal to each
// other, so that the stable sort keeps their sequence.
func compareSortValues(a, b werkbook.Value, desc bool) int {
	switch {
	case a.IsEmpty() && b.IsEmpty():
		return 0
	case a.IsEmpty():
		return 1
	case b.IsEmpty():
		return -1
	}
	c := 0
	ra, rb := sortRank(a), sortRank(b)
	switch {
	case ra != rb:
		c = ra - rb
	case ra == sortRankNumber:
		switch {
		case a.Number < b.Number:
			c = -1
		case a.Number > b.Number:
			c = 1
		}
	case ra == sortRankText:
		// Letter case is not important, and a letter with an accent sorts
		// with its base letter. Equal text keeps its sequence.
		c = compareText(a.String, b.String)
	case ra == sortRankBool:
		switch {
		case !a.Bool && b.Bool:
			c = -1
		case a.Bool && !b.Bool:
			c = 1
		}
	}
	if desc {
		c = -c
	}
	return c
}

// SortRange sorts the rows of a range by key columns as one undo step. See
// the rules in cmd/web/README.md.
func (s *Session) SortRange(sheet string, rg Range, keys []SortKey, header bool) (SortResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return SortResult{}, err
	}
	rg, err = normRange(rg)
	if err != nil {
		return SortResult{}, err
	}
	if len(keys) == 0 {
		return SortResult{}, errors.New("sortRange: no sort key")
	}
	for i, k := range keys {
		if k.Col < rg.Col1 || k.Col > rg.Col2 {
			return SortResult{}, fmt.Errorf("sortRange: key %d: column %d is not in the range", i, k.Col)
		}
	}

	s.ensureCalculated()
	// A range of full rows or full columns is the used part of it.
	maxRow, maxCol := s.usedRange(sh)
	if rg.Col1 == 1 && rg.Col2 == werkbook.MaxColumns {
		rg.Col2 = max(rg.Col1, min(rg.Col2, maxCol))
	}
	if rg.Row1 == 1 && rg.Row2 == werkbook.MaxRows {
		rg.Row2 = max(rg.Row1, min(rg.Row2, maxRow))
	}
	for i, k := range keys {
		if k.Col > rg.Col2 {
			return SortResult{}, fmt.Errorf("sortRange: key %d: column %d is not in the used range", i, k.Col)
		}
	}
	res := SortResult{Range: rg}
	if rg.cells() > maxOpCells {
		return SortResult{}, fmt.Errorf("sortRange: the range has more than %d cells", maxOpCells)
	}
	if intersectsMerge(sheetMerges(sh), rg) {
		return SortResult{}, errors.New("sortRange: the range has merged cells")
	}
	for _, sp := range s.spillRects(sh) {
		if (Range{Row1: sp.row1, Col1: sp.col1, Row2: sp.row2, Col2: sp.col2}).intersects(rg) {
			return SortResult{}, errors.New("sortRange: the range has a dynamic array (spill) formula")
		}
	}

	first := rg.Row1
	if header {
		first++
	}
	n := rg.Row2 - first + 1
	if n < 2 {
		res.History = s.History()
		return res, nil
	}

	// Read all rows before the first write: the stored state of each cell
	// and the shown value of each key cell.
	width := rg.cols()
	rows := make([][]werkbook.CellSnapshot, n)
	keyValues := make([][]werkbook.Value, n)
	for i := 0; i < n; i++ {
		row := first + i
		rows[i] = make([]werkbook.CellSnapshot, width)
		for j := 0; j < width; j++ {
			snap := sh.SnapshotCell(rg.Col1+j, row)
			if snap.IsArrayFormula() {
				return SortResult{}, errors.New("sortRange: the range has an array formula")
			}
			if anchorLike(snap) {
				return SortResult{}, errors.New("sortRange: the range has a dynamic array (spill) formula")
			}
			rows[i][j] = snap
		}
		keyValues[i] = make([]werkbook.Value, len(keys))
		for k, key := range keys {
			ref, err := werkbook.CoordinatesToCellName(key.Col, row)
			if err != nil {
				return SortResult{}, err
			}
			v, err := sh.GetValue(ref)
			if err != nil {
				return SortResult{}, err
			}
			keyValues[i][k] = v
		}
	}

	// order[i] is the index of the old row that goes to position i.
	order := make([]int, n)
	for i := range order {
		order[i] = i
	}
	sort.SliceStable(order, func(a, b int) bool {
		for k, key := range keys {
			if c := compareSortValues(keyValues[order[a]][k], keyValues[order[b]][k], key.Desc); c != 0 {
				return c < 0
			}
		}
		return false
	})

	err = s.mutate(labelSort, func(rec *recorder) error {
		rec.setSheet(sheet)
		rec.setRange(rg)
		for i, from := range order {
			if from == i {
				continue // the row stays
			}
			row := first + i
			for j := 0; j < width; j++ {
				snap := rows[from][j]
				if f := snap.Formula(); f != "" {
					// A formula moves as a copy does: its relative
					// references shift by the row distance.
					snap = snap.WithFormula(formula.ShiftRefs(f, 0, i-from))
				}
				if sameContent(snap, rows[i][j]) {
					continue
				}
				if err := rec.write(sh, rg.Col1+j, row, snap); err != nil {
					return err
				}
				res.Changed++
			}
		}
		return nil
	})
	if err != nil {
		return SortResult{}, err
	}
	res.History = s.History()
	return res, nil
}

// anchorLike reports whether a snapshot is a formula that would spill (a
// dynamic array anchor). Such a cell cannot move in a sort: its spill range
// would cross the rows of the range.
func anchorLike(snap werkbook.CellSnapshot) bool {
	f := snap.Formula()
	return f != "" && formula.IsDynamicArrayFormula(f)
}

// sameContent reports whether two snapshots give the same cell: the same
// value or formula text, and the same style object.
func sameContent(a, b werkbook.CellSnapshot) bool {
	if a.Style() != b.Style() || a.Formula() != b.Formula() {
		return false
	}
	if a.Formula() != "" {
		return true
	}
	if a.IsEmpty() && b.IsEmpty() {
		return true
	}
	return a.Value() == b.Value()
}
