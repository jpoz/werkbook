package werkbook

import (
	"strconv"
	"strings"

	"github.com/jpoz/werkbook/formula"
)

type externalBook struct {
	sheets map[string]*externalSheet
}

type externalSheet struct {
	rows   map[int]map[int]formula.Value
	maxRow int
	maxCol int
}

func externalSheetKey(name string) string {
	return strings.ToLower(name)
}

func parseExternalSheetRef(name string) (bookIndex int, sheetName string, ok bool) {
	if len(name) < 4 || name[0] != '[' {
		return 0, "", false
	}
	end := strings.IndexByte(name, ']')
	if end <= 1 || end >= len(name)-1 {
		return 0, "", false
	}
	n, err := strconv.Atoi(name[1:end])
	if err != nil || n < 1 {
		return 0, "", false
	}
	return n - 1, name[end+1:], true
}

func (s *externalSheet) cellValue(col, row int) formula.Value {
	if s == nil {
		return formula.ErrorVal(formula.ErrValREF)
	}
	if rowVals, ok := s.rows[row]; ok {
		if v, ok := rowVals[col]; ok {
			return v
		}
	}
	return formula.EmptyVal()
}

func (s *externalSheet) rangeValues(addr formula.RangeAddr) [][]formula.Value {
	rangeOverflow := func() [][]formula.Value {
		return [][]formula.Value{{{
			Type:          formula.ValueError,
			Err:           formula.ErrValREF,
			RangeOverflow: true,
		}}}
	}

	toRow := addr.ToRow
	maxRow := s.maxRow
	if maxRow < addr.FromRow {
		maxRow = addr.FromRow
	}
	if toRow > maxRow {
		toRow = maxRow
	}
	if toRow < addr.FromRow {
		toRow = addr.FromRow
	}

	toCol := addr.ToCol
	if addr.FromCol == 1 && addr.ToCol >= MaxColumns {
		maxCol := s.maxCol
		if maxCol < addr.FromCol {
			maxCol = addr.FromCol
		}
		if toCol > maxCol {
			toCol = maxCol
		}
		if toCol < addr.FromCol {
			toCol = addr.FromCol
		}
	}

	nRows := toRow - addr.FromRow + 1
	nCols := toCol - addr.FromCol + 1
	if formula.RangeCellCountExceedsLimit(nRows, nCols) {
		return rangeOverflow()
	}

	rows := newFormulaValueMatrix(nRows, nCols)
	for rowNum, rowVals := range s.rows {
		if rowNum < addr.FromRow || rowNum > toRow {
			continue
		}
		row := rows[rowNum-addr.FromRow]
		for colNum, val := range rowVals {
			if colNum < addr.FromCol || colNum > toCol {
				continue
			}
			row[colNum-addr.FromCol] = val
		}
	}
	return rows
}
