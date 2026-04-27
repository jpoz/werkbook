package werkbook

import "github.com/jpoz/werkbook/formula"

type rangeMaterializationRequest struct {
	sheet   string
	fromCol int
	fromRow int
	toCol   int
	toRow   int
}

type rangeMaterializationResult struct {
	cells          [][]formula.Value
	bounds         formula.RangeAddr
	discoveredDeps []formula.RangeAddr
	overflow       bool
}

type rangeGridReader interface {
	MaxRow(sheet string) int
	MaxCol(sheet string) int
	ForEachCell(
		sheet string,
		fromCol, fromRow, toCol, toRow int,
		fn func(col, row int, value formula.Value, occupies bool),
	)
}

type rangeSpillReader interface {
	Anchors(sheet string) []rangeSpillAnchor
}

type rangeSpillAnchor struct {
	col   int
	row   int
	toCol int
	toRow int
	raw   formula.Value
}

type rangeMaterializationGrid struct {
	maxRow int
	maxCol int
	cells  []rangeMaterializationCell
}

type sheetRangeGrid struct {
	sheet  *Sheet
	maxRow int
	maxCol int
}

type rangeMaterializationCell struct {
	col      int
	row      int
	value    formula.Value
	occupies bool
}

type rangeMaterializationSpills struct {
	anchors []rangeSpillAnchor
}

func (g *rangeMaterializationGrid) MaxRow(string) int {
	if g == nil {
		return 0
	}
	return g.maxRow
}

func (g *rangeMaterializationGrid) MaxCol(string) int {
	if g == nil {
		return 0
	}
	return g.maxCol
}

func (g *rangeMaterializationGrid) ForEachCell(
	_ string,
	fromCol, fromRow, toCol, toRow int,
	fn func(col, row int, value formula.Value, occupies bool),
) {
	if g == nil || fn == nil {
		return
	}
	for _, cell := range g.cells {
		if cell.col < fromCol || cell.col > toCol {
			continue
		}
		if cell.row < fromRow || cell.row > toRow {
			continue
		}
		fn(cell.col, cell.row, cell.value, cell.occupies)
	}
}

func (g sheetRangeGrid) MaxRow(string) int {
	return g.maxRow
}

func (g sheetRangeGrid) MaxCol(string) int {
	return g.maxCol
}

func (g sheetRangeGrid) ForEachCell(
	_ string,
	fromCol, fromRow, toCol, toRow int,
	fn func(col, row int, value formula.Value, occupies bool),
) {
	if g.sheet == nil || fn == nil {
		return
	}
	for rowNum, sheetRow := range g.sheet.rows {
		if rowNum < fromRow || rowNum > toRow {
			continue
		}
		for colNum, cell := range sheetRow.cells {
			if colNum < fromCol || colNum > toCol {
				continue
			}
			g.sheet.resolveCell(cell, colNum, rowNum)
			fn(colNum, rowNum, valueToFormulaValue(cell.value), cellOccupiesSpillSlot(cell))
		}
	}
}

func (s *rangeMaterializationSpills) Anchors(string) []rangeSpillAnchor {
	if s == nil {
		return nil
	}
	return s.anchors
}

func clampMaterializedRange(req rangeMaterializationRequest, maxRow, maxCol int) (toRow, toCol int) {
	toRow = req.toRow
	if maxRow < req.fromRow {
		maxRow = req.fromRow
	}
	if toRow > maxRow {
		toRow = maxRow
	}
	if toRow < req.fromRow {
		toRow = req.fromRow
	}

	toCol = req.toCol
	if req.fromCol == 1 && req.toCol >= MaxColumns {
		if maxCol < req.fromCol {
			maxCol = req.fromCol
		}
		if toCol > maxCol {
			toCol = maxCol
		}
		if toCol < req.fromCol {
			toCol = req.fromCol
		}
	}

	return toRow, toCol
}

func rangeOverflowMatrix() [][]formula.Value {
	return [][]formula.Value{{{
		Type:          formula.ValueError,
		Err:           formula.ErrValREF,
		RangeOverflow: true,
	}}}
}

func newBoolMatrix(nRows, nCols int) [][]bool {
	if nRows <= 0 || nCols <= 0 {
		return nil
	}
	rows := make([][]bool, nRows)
	cells := make([]bool, nRows*nCols)
	for i := range rows {
		start := i * nCols
		rows[i] = cells[start : start+nCols]
	}
	return rows
}

func materializeRange(req rangeMaterializationRequest, grid rangeGridReader, spills rangeSpillReader) rangeMaterializationResult {
	res := rangeMaterializationResult{
		bounds: formula.RangeAddr{
			Sheet:   req.sheet,
			FromCol: req.fromCol,
			FromRow: req.fromRow,
			ToCol:   req.toCol,
			ToRow:   req.toRow,
		},
	}

	nCols := req.toCol - req.fromCol + 1
	if grid == nil {
		nRows := req.toRow - req.fromRow + 1
		if formula.RangeCellCountExceedsLimit(nRows, nCols) {
			res.overflow = true
			res.cells = rangeOverflowMatrix()
			return res
		}
		rows := newFormulaValueMatrix(nRows, nCols)
		if rows == nil {
			return res
		}
		refErr := formula.ErrorVal(formula.ErrValREF)
		for i := range rows {
			for j := range rows[i] {
				rows[i][j] = refErr
			}
		}
		res.cells = rows
		return res
	}

	toRow, toCol := clampMaterializedRange(req, grid.MaxRow(req.sheet), grid.MaxCol(req.sheet))
	nRows := toRow - req.fromRow + 1
	nCols = toCol - req.fromCol + 1
	if nRows <= 0 || nCols <= 0 {
		return res
	}
	if formula.RangeCellCountExceedsLimit(nRows, nCols) {
		res.overflow = true
		res.cells = rangeOverflowMatrix()
		return res
	}

	rows := newFormulaValueMatrix(nRows, nCols)
	occupied := newBoolMatrix(nRows, nCols)

	grid.ForEachCell(req.sheet, req.fromCol, req.fromRow, toCol, toRow, func(col, row int, value formula.Value, occupies bool) {
		rowIdx := row - req.fromRow
		colIdx := col - req.fromCol
		if rowIdx < 0 || rowIdx >= len(rows) || colIdx < 0 || colIdx >= len(rows[rowIdx]) {
			return
		}
		rows[rowIdx][colIdx] = value
		occupied[rowIdx][colIdx] = occupies
	})

	if spills != nil {
		logicalToCol := req.toCol
		for _, anchor := range spills.Anchors(req.sheet) {
			if anchor.row > toRow || anchor.col > logicalToCol {
				continue
			}
			if anchor.toRow < req.fromRow || anchor.toCol < req.fromCol {
				continue
			}

			nextToRow := toRow
			nextToCol := toCol
			if anchor.toRow > nextToRow {
				nextToRow = anchor.toRow
			}
			if anchor.toCol > nextToCol {
				nextToCol = anchor.toCol
			}
			if nextToRow > toRow || nextToCol > toCol {
				if nextToRow < req.fromRow {
					nextToRow = req.fromRow
				}
				if nextToCol < req.fromCol {
					nextToCol = req.fromCol
				}
				if nextToRow > req.toRow {
					nextToRow = req.toRow
				}
				if nextToCol > logicalToCol {
					nextToCol = logicalToCol
				}
				nextRows := nextToRow - req.fromRow + 1
				nextCols := nextToCol - req.fromCol + 1
				if formula.RangeCellCountExceedsLimit(nextRows, nextCols) {
					res.overflow = true
					res.cells = rangeOverflowMatrix()
					return res
				}

				grownRows := newFormulaValueMatrix(nextRows, nextCols)
				grownOccupied := newBoolMatrix(nextRows, nextCols)
				for i := range rows {
					copy(grownRows[i], rows[i])
					copy(grownOccupied[i], occupied[i])
				}
				rows = grownRows
				occupied = grownOccupied
				toRow = nextToRow
				toCol = nextToCol
				res.discoveredDeps = append(res.discoveredDeps, formula.RangeAddr{
					Sheet:   req.sheet,
					FromCol: anchor.col,
					FromRow: anchor.row,
					ToCol:   anchor.toCol,
					ToRow:   anchor.toRow,
				})
			}

			raw := anchor.raw
			if raw.Type != formula.ValueArray || raw.NoSpill {
				continue
			}
			for rowOffset, spillRow := range raw.Array {
				rowNum := anchor.row + rowOffset
				if rowNum < req.fromRow || rowNum > toRow {
					continue
				}
				rowIdx := rowNum - req.fromRow
				row := rows[rowIdx]
				for colOffset, spillVal := range spillRow {
					colNum := anchor.col + colOffset
					if colNum < req.fromCol || colNum > toCol {
						continue
					}
					colIdx := colNum - req.fromCol
					if occupied[rowIdx][colIdx] {
						continue
					}
					row[colIdx] = spillVal
				}
			}
		}
	}

	res.cells = rows
	res.bounds.ToRow = toRow
	res.bounds.ToCol = toCol
	return res
}
