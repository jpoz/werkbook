package werkbook

import (
	"fmt"
	"io"
	"iter"
	"sort"
	"strconv"
	"strings"

	"github.com/jpoz/werkbook/formula"
	"github.com/jpoz/werkbook/ooxml"
)

// Sheet represents a single worksheet in the workbook.
type Sheet struct {
	file      *File
	name      string
	visible   bool
	rows      map[int]*Row
	colWidths map[int]float64
	merges    []MergeRange
}

func newSheet(name string, file *File) *Sheet {
	return &Sheet{
		file:      file,
		name:      name,
		visible:   true,
		rows:      make(map[int]*Row),
		colWidths: make(map[int]float64),
	}
}

// MergeRange represents a merged cell range.
type MergeRange struct {
	Start string
	End   string
}

// Name returns the sheet name.
func (s *Sheet) Name() string {
	return s.name
}

// Visible reports whether the sheet is visible.
func (s *Sheet) Visible() bool {
	return s.visible
}

// SetValue sets the value of a cell by reference (e.g. "A1").
// Supported types: string, bool, int*, uint*, float32, float64, nil.
func (s *Sheet) SetValue(cell string, v any) error {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}

	val, err := toValueWithDate1904(v, s.file.date1904)
	if err != nil {
		return err
	}

	r := s.ensureRow(row)
	c := r.ensureCell(col)
	s.prepareCellForWrite(c, col, row)
	r = s.ensureRow(row)
	c = r.ensureCell(col)
	// Unregister old formula if any.
	if c.formula != "" {
		s.file.deps.Unregister(formula.QualifiedCell{Sheet: s.name, Col: col, Row: row})
	}
	c.value = val
	c.formula = ""
	c.isArrayFormula = false
	c.isDynamicArray = false
	c.formulaRef = ""
	c.spillParentCol = 0
	c.spillParentRow = 0
	c.compiled = nil
	c.cachedGen = 0
	s.file.calcGen++
	s.file.invalidateDependents(s.name, col, row)
	s.markDynamicArrayAnchorsDirty(col, row)
	return nil
}

// SetFormula sets a formula on a cell by reference (e.g. "A1").
// The formula should not include the leading '=' sign.
func (s *Sheet) SetFormula(cell string, f string) error {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	src, err := s.file.expandFormula(f, s.name, row)
	if err != nil {
		return err
	}
	r := s.ensureRow(row)
	c := r.ensureCell(col)
	s.prepareCellForWrite(c, col, row)
	r = s.ensureRow(row)
	c = r.ensureCell(col)
	// Unregister old formula if any.
	qc := formula.QualifiedCell{Sheet: s.name, Col: col, Row: row}
	if c.formula != "" {
		s.file.deps.Unregister(qc)
	}
	c.formula = f
	c.isArrayFormula = false
	c.isDynamicArray = formula.IsDynamicArrayFormula(f)
	c.formulaRef = ""
	c.spillParentCol = 0
	c.spillParentRow = 0
	c.compiled = nil
	c.value = Value{}
	c.cachedGen = 0
	c.dirty = true
	s.file.calcGen++
	// Compile and register in dep graph.
	node, parseErr := formula.Parse(src)
	if parseErr == nil {
		cf, compErr := formula.Compile(src, node)
		if compErr == nil {
			c.compiled = cf
			s.file.deps.Register(qc, s.name, cf.Refs, cf.Ranges)
		}
	}
	s.file.invalidateDependents(s.name, col, row)
	s.markDynamicArrayAnchorsDirty(col, row)
	if c.isDynamicArray {
		s.resolveCell(c, col, row)
	}
	return nil
}

// SetStyle sets the style of a cell by reference (e.g. "A1").
func (s *Sheet) SetStyle(cell string, style *Style) error {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	r := s.ensureRow(row)
	c := r.ensureCell(col)
	c.style = style
	return nil
}

// GetStyle returns the style of a cell by reference (e.g. "A1").
// Returns nil for default-styled or nonexistent cells.
func (s *Sheet) GetStyle(cell string) (*Style, error) {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return nil, fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	r, ok := s.rows[row]
	if !ok {
		return nil, nil
	}
	c, ok := r.cells[col]
	if !ok {
		return nil, nil
	}
	return c.style, nil
}

// SetColumnWidth sets the width of a column by name (e.g. "A").
func (s *Sheet) SetColumnWidth(col string, width float64) error {
	num, err := ColumnNameToNumber(col)
	if err != nil {
		return fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	s.colWidths[num] = width
	return nil
}

// GetColumnWidth returns the width of a column by name, or 0 if not set.
func (s *Sheet) GetColumnWidth(col string) (float64, error) {
	num, err := ColumnNameToNumber(col)
	if err != nil {
		return 0, fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	return s.colWidths[num], nil
}

// SetRowHeight sets the height of a row by 1-based row number.
func (s *Sheet) SetRowHeight(row int, height float64) error {
	if row < 1 || row > MaxRows {
		return fmt.Errorf("%w: row %d out of range [1, %d]", ErrInvalidCellRef, row, MaxRows)
	}
	r := s.ensureRow(row)
	r.height = height
	return nil
}

// GetRowHeight returns the height of a row, or 0 if not set.
func (s *Sheet) GetRowHeight(row int) (float64, error) {
	if row < 1 || row > MaxRows {
		return 0, fmt.Errorf("%w: row %d out of range [1, %d]", ErrInvalidCellRef, row, MaxRows)
	}
	r, ok := s.rows[row]
	if !ok {
		return 0, nil
	}
	return r.height, nil
}

// RemoveRow removes a row and shifts all following rows up by one.
func (s *Sheet) RemoveRow(row int) error {
	if row < 1 || row > MaxRows {
		return fmt.Errorf("%w: row %d out of range [1, %d]", ErrInvalidCellRef, row, MaxRows)
	}

	newRows := make(map[int]*Row, len(s.rows))
	for rn, r := range s.rows {
		switch {
		case rn < row:
			newRows[rn] = r
		case rn == row:
			continue
		default:
			r.num = rn - 1
			newRows[rn-1] = r
		}
	}
	s.rows = newRows
	s.adjustMergedRows(row)
	s.file.rebuildFormulaState()
	return nil
}

// SetRangeStyle applies the given style to every cell in the range (e.g. "A1:C5").
// Cells that do not yet exist are created.
func (s *Sheet) SetRangeStyle(rangeRef string, style *Style) error {
	col1, row1, col2, row2, err := RangeToCoordinates(rangeRef)
	if err != nil {
		return fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	for r := row1; r <= row2; r++ {
		row := s.ensureRow(r)
		for c := col1; c <= col2; c++ {
			cell := row.ensureCell(c)
			cell.style = style
		}
	}
	return nil
}

// MergeCell merges the rectangular range bounded by start and end.
func (s *Sheet) MergeCell(start, end string) error {
	col1, row1, col2, row2, err := RangeToCoordinates(start + ":" + end)
	if err != nil {
		return fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	startRef, err := CoordinatesToCellName(col1, row1)
	if err != nil {
		return err
	}
	endRef, err := CoordinatesToCellName(col2, row2)
	if err != nil {
		return err
	}
	for _, mr := range s.merges {
		if mr.Start == startRef && mr.End == endRef {
			return nil
		}
	}
	s.merges = append(s.merges, MergeRange{Start: startRef, End: endRef})
	return nil
}

// MergeCells returns the merged ranges on the sheet.
func (s *Sheet) MergeCells() []MergeRange {
	out := make([]MergeRange, len(s.merges))
	copy(out, s.merges)
	return out
}

// GetFormula returns the formula for a cell, or "" if none.
func (s *Sheet) GetFormula(cell string) (string, error) {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return "", fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	r, ok := s.rows[row]
	if !ok {
		return "", nil
	}
	c, ok := r.cells[col]
	if !ok {
		return "", nil
	}
	return c.formula, nil
}

// GetValue returns the value of a cell by reference (e.g. "A1").
func (s *Sheet) GetValue(cell string) (Value, error) {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return Value{}, fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}

	r, ok := s.rows[row]
	if !ok {
		return Value{Type: TypeEmpty}, nil
	}
	c, ok := r.cells[col]
	if !ok {
		return Value{Type: TypeEmpty}, nil
	}
	if c.isSpillFollower() {
		if anchor := s.cellAt(c.spillParentCol, c.spillParentRow); anchor != nil {
			s.resolveCell(anchor, c.spillParentCol, c.spillParentRow)
			r = s.rows[row]
			if r == nil {
				return Value{Type: TypeEmpty}, nil
			}
			c = r.cells[col]
			if c == nil {
				return Value{Type: TypeEmpty}, nil
			}
		}
	}

	s.resolveCell(c, col, row)
	return c.value, nil
}

type sheetCoord struct {
	col int
	row int
}

func (s *Sheet) cellAt(col, row int) *Cell {
	r, ok := s.rows[row]
	if !ok {
		return nil
	}
	return r.cells[col]
}

func spillRangeRef(col, row, rows, cols int) string {
	start, err := CoordinatesToCellName(col, row)
	if err != nil {
		return ""
	}
	if rows <= 1 && cols <= 1 {
		return start
	}
	end, err := CoordinatesToCellName(col+cols-1, row+rows-1)
	if err != nil {
		return start
	}
	return start + ":" + end
}

func formulaArraySize(array [][]formula.Value) (rows, cols int) {
	rows = len(array)
	for _, row := range array {
		if len(row) > cols {
			cols = len(row)
		}
	}
	return rows, cols
}

func formulaArrayValueAt(array [][]formula.Value, row, col int) formula.Value {
	if row >= 0 && row < len(array) && col >= 0 && col < len(array[row]) {
		return array[row][col]
	}
	return formula.EmptyVal()
}

func (s *Sheet) prepareCellForWrite(c *Cell, col, row int) {
	if c == nil {
		return
	}
	if c.isDynamicArray {
		s.clearSpillFollowers(col, row)
	}
	if !c.isSpillFollower() {
		return
	}
	anchorCol := c.spillParentCol
	anchorRow := c.spillParentRow
	s.clearSpillFollowers(anchorCol, anchorRow)
	if anchor := s.cellAt(anchorCol, anchorRow); anchor != nil && anchor.formula != "" {
		anchor.dirty = true
		anchor.cachedGen = 0
		s.file.invalidateDependents(s.name, anchorCol, anchorRow)
	}
}

func (s *Sheet) clearSpillFollowers(anchorCol, anchorRow int) {
	if anchorCol <= 0 || anchorRow <= 0 {
		return
	}

	rowNums := make([]int, 0, len(s.rows))
	for rowNum := range s.rows {
		rowNums = append(rowNums, rowNum)
	}
	sort.Ints(rowNums)

	for _, rowNum := range rowNums {
		r := s.rows[rowNum]
		if r == nil {
			continue
		}

		colNums := make([]int, 0, len(r.cells))
		for col := range r.cells {
			colNums = append(colNums, col)
		}
		sort.Ints(colNums)

		for _, col := range colNums {
			c := r.cells[col]
			if c == nil || c.spillParentCol != anchorCol || c.spillParentRow != anchorRow {
				continue
			}
			c.spillParentCol = 0
			c.spillParentRow = 0
			c.value = Value{}
			c.cachedGen = 0
			c.dirty = false
			if c.formula == "" && c.style == nil && c.value.Type == TypeEmpty {
				delete(r.cells, col)
			}
			s.file.invalidateDependents(s.name, col, rowNum)
		}

		if len(r.cells) == 0 && r.height == 0 && !r.hidden {
			delete(s.rows, rowNum)
		}
	}

	if anchor := s.cellAt(anchorCol, anchorRow); anchor != nil {
		anchor.formulaRef = spillRangeRef(anchorCol, anchorRow, 1, 1)
	}
}

func (s *Sheet) canSpillInto(anchorCol, anchorRow, rows, cols int) bool {
	if rows <= 0 || cols <= 0 {
		return true
	}
	if anchorRow+rows-1 > MaxRows || anchorCol+cols-1 > MaxColumns {
		return false
	}

	for rowOff := 0; rowOff < rows; rowOff++ {
		for colOff := 0; colOff < cols; colOff++ {
			if rowOff == 0 && colOff == 0 {
				continue
			}
			target := s.cellAt(anchorCol+colOff, anchorRow+rowOff)
			if target == nil {
				continue
			}
			if target.spillParentCol == anchorCol && target.spillParentRow == anchorRow {
				continue
			}
			if target.formula != "" || target.isSpillFollower() {
				return false
			}
			if target.value.Type != TypeEmpty {
				return false
			}
		}
	}
	return true
}

func (s *Sheet) materializeDynamicArray(c *Cell, col, row int, result formula.Value) Value {
	c.formulaRef = spillRangeRef(col, row, 1, 1)
	s.clearSpillFollowers(col, row)

	if result.Type != formula.ValueArray || result.NoSpill {
		return formulaValueToValue(result, c.isArrayFormula)
	}

	nRows, nCols := formulaArraySize(result.Array)
	if nRows <= 0 || nCols <= 0 {
		return formulaValueToValue(result, c.isArrayFormula)
	}
	if nRows == 1 && nCols == 1 {
		return formulaValueToValue(result.Array[0][0], c.isArrayFormula)
	}
	if !s.canSpillInto(col, row, nRows, nCols) {
		return Value{Type: TypeError, String: "#SPILL!"}
	}

	c.formulaRef = spillRangeRef(col, row, nRows, nCols)
	for rowOff := 0; rowOff < nRows; rowOff++ {
		for colOff := 0; colOff < nCols; colOff++ {
			if rowOff == 0 && colOff == 0 {
				continue
			}
			targetRow := row + rowOff
			targetCol := col + colOff
			targetRowData := s.ensureRow(targetRow)
			target := targetRowData.ensureCell(targetCol)
			target.value = formulaValueToValue(formulaArrayValueAt(result.Array, rowOff, colOff), false)
			target.formula = ""
			target.isArrayFormula = false
			target.isDynamicArray = false
			target.formulaRef = ""
			target.spillParentCol = col
			target.spillParentRow = row
			target.compiled = nil
			target.cachedGen = 0
			target.dirty = false
			s.file.invalidateDependents(s.name, targetCol, targetRow)
		}
	}
	return formulaValueToValue(result.Array[0][0], c.isArrayFormula)
}

func (s *Sheet) restoreDynamicSpillFollowers() {
	rowNums := make([]int, 0, len(s.rows))
	for rowNum := range s.rows {
		rowNums = append(rowNums, rowNum)
	}
	sort.Ints(rowNums)

	for _, rowNum := range rowNums {
		r := s.rows[rowNum]
		if r == nil {
			continue
		}

		colNums := make([]int, 0, len(r.cells))
		for col := range r.cells {
			colNums = append(colNums, col)
		}
		sort.Ints(colNums)

		for _, col := range colNums {
			c := r.cells[col]
			if c == nil || !c.isDynamicArray || c.formulaRef == "" {
				continue
			}
			fromCol, fromRow, toCol, toRow, err := RangeToCoordinates(c.formulaRef)
			if err != nil {
				continue
			}
			for spillRow := fromRow; spillRow <= toRow; spillRow++ {
				spillCells := s.rows[spillRow]
				if spillCells == nil {
					continue
				}
				for spillCol := fromCol; spillCol <= toCol; spillCol++ {
					if spillCol == col && spillRow == rowNum {
						continue
					}
					spillCell := spillCells.cells[spillCol]
					if spillCell == nil {
						continue
					}
					spillCell.spillParentCol = col
					spillCell.spillParentRow = rowNum
				}
			}
		}
	}
}

func (s *Sheet) spillDependents(anchorCol, anchorRow int, g *formula.DepGraph) []formula.QualifiedCell {
	if anchorCol <= 0 || anchorRow <= 0 {
		return nil
	}

	seen := make(map[formula.QualifiedCell]bool)
	var deps []formula.QualifiedCell
	for rowNum, r := range s.rows {
		for col, c := range r.cells {
			if c == nil || c.spillParentCol != anchorCol || c.spillParentRow != anchorRow {
				continue
			}
			for _, dep := range g.DirectDependents(formula.QualifiedCell{Sheet: s.name, Col: col, Row: rowNum}) {
				if seen[dep] {
					continue
				}
				seen[dep] = true
				deps = append(deps, dep)
			}
		}
	}
	return deps
}

func (s *Sheet) markDynamicArrayAnchorsDirty(excludeCol, excludeRow int) {
	for rowNum, r := range s.rows {
		for col, c := range r.cells {
			if c == nil || !c.isDynamicArray || c.formula == "" {
				continue
			}
			if col == excludeCol && rowNum == excludeRow {
				continue
			}
			c.dirty = true
			c.cachedGen = 0
		}
	}
}

func (s *Sheet) ensureRow(num int) *Row {
	r, ok := s.rows[num]
	if !ok {
		r = &Row{num: num, cells: make(map[int]*Cell)}
		s.rows[num] = r
	}
	return r
}

func (r *Row) ensureCell(col int) *Cell {
	c, ok := r.cells[col]
	if !ok {
		c = &Cell{col: col}
		r.cells[col] = c
	}
	return c
}

// Rows returns an iterator over all non-empty rows in ascending order.
// Rows with a custom height but no cells are also included.
func (s *Sheet) Rows() iter.Seq[*Row] {
	return func(yield func(*Row) bool) {
		rowNums := make([]int, 0, len(s.rows))
		for n := range s.rows {
			rowNums = append(rowNums, n)
		}
		sort.Ints(rowNums)
		for _, n := range rowNums {
			r := s.rows[n]
			if len(r.cells) == 0 && r.height == 0 {
				continue
			}
			if !yield(r) {
				return
			}
		}
	}
}

// MaxRow returns the highest 1-based row number with data, or 0 if empty.
func (s *Sheet) MaxRow() int {
	max := 0
	for n := range s.rows {
		if n > max {
			max = n
		}
	}
	return max
}

// MaxCol returns the highest 1-based column number with data across all rows, or 0 if empty.
func (s *Sheet) MaxCol() int {
	max := 0
	for _, r := range s.rows {
		for c := range r.cells {
			if c > max {
				max = c
			}
		}
	}
	return max
}

// PrintTo writes a human-readable table of all cell values to w.
func (s *Sheet) PrintTo(w io.Writer) {
	maxCol := s.MaxCol()
	if maxCol == 0 {
		return
	}

	colWidths := make([]int, maxCol)
	var grid [][]string

	for row := range s.Rows() {
		vals := make([]string, maxCol)
		for _, c := range row.Cells() {
			ref, _ := CoordinatesToCellName(c.Col(), row.Num())
			v, _ := s.GetValue(ref)
			var text string
			switch v.Type {
			case TypeNumber:
				if v.Number == float64(int64(v.Number)) {
					text = fmt.Sprintf("%d", int64(v.Number))
				} else {
					text = fmt.Sprintf("%.2f", v.Number)
				}
			case TypeString:
				text = v.String
			case TypeBool:
				if v.Bool {
					text = "TRUE"
				} else {
					text = "FALSE"
				}
			case TypeError:
				text = v.String
			}
			idx := c.Col() - 1
			vals[idx] = text
			if len(text) > colWidths[idx] {
				colWidths[idx] = len(text)
			}
		}
		grid = append(grid, vals)
	}

	for _, vals := range grid {
		for c, text := range vals {
			if c > 0 {
				fmt.Fprint(w, "  ")
			}
			fmt.Fprintf(w, "%-*s", colWidths[c], text)
		}
		fmt.Fprintln(w)
	}
}

// toSheetData converts the sheet to the ooxml intermediate representation.
// styleMap maps style keys to indices in the WorkbookData.Styles slice.
// styles collects all unique StyleData values; both are mutated in place.
func (s *Sheet) toSheetData(styleMap map[string]int, styles *[]ooxml.StyleData) ooxml.SheetData {
	sd := ooxml.SheetData{Name: s.name}
	if !s.visible {
		sd.State = "hidden"
	}

	// Convert column widths map to ColWidthData slice.
	if len(s.colWidths) > 0 {
		colNums := make([]int, 0, len(s.colWidths))
		for c := range s.colWidths {
			colNums = append(colNums, c)
		}
		sort.Ints(colNums)
		for _, c := range colNums {
			sd.ColWidths = append(sd.ColWidths, ooxml.ColWidthData{
				Min: c, Max: c, Width: s.colWidths[c],
			})
		}
	}

	// Resolve formula cells up front so dynamic-array spill followers are
	// materialized before we snapshot the sparse row/cell structure.
	var formulaCells []sheetCoord
	for rowNum, r := range s.rows {
		for col, c := range r.cells {
			if c.formula == "" {
				continue
			}
			formulaCells = append(formulaCells, sheetCoord{col: col, row: rowNum})
		}
	}
	for _, coord := range formulaCells {
		r := s.rows[coord.row]
		if r == nil {
			continue
		}
		c := r.cells[coord.col]
		if c == nil {
			continue
		}
		s.resolveCell(c, coord.col, coord.row)
	}

	// Sort rows by number.
	rowNums := make([]int, 0, len(s.rows))
	for n := range s.rows {
		rowNums = append(rowNums, n)
	}
	sort.Ints(rowNums)

	for _, rn := range rowNums {
		r := s.rows[rn]
		if len(r.cells) == 0 && r.height == 0 && !r.hidden {
			continue
		}
		rd := ooxml.RowData{Num: rn, Height: r.height, Hidden: r.hidden}

		// Sort cells by column.
		colNums := make([]int, 0, len(r.cells))
		for c := range r.cells {
			colNums = append(colNums, c)
		}
		sort.Ints(colNums)

		for _, cn := range colNums {
			c := r.cells[cn]
			// Resolve dirty/stale formulas before serializing.
			s.resolveCell(c, cn, rn)
			if c.value.Type == TypeEmpty && c.formula == "" && c.style == nil {
				continue
			}
			ref, _ := CoordinatesToCellName(cn, rn)
			cd := cellToData(ref, c.value, c.formula, c.isArrayFormula, c.isDynamicArray, c.formulaRef)

			if c.style != nil {
				stData := styleToStyleData(c.style)
				key := styleKey(stData)
				idx, ok := styleMap[key]
				if !ok {
					idx = len(*styles)
					styleMap[key] = idx
					*styles = append(*styles, stData)
				}
				cd.StyleIdx = idx
			}

			rd.Cells = append(rd.Cells, cd)
		}
		if len(rd.Cells) > 0 || rd.Height != 0 || rd.Hidden {
			sd.Rows = append(sd.Rows, rd)
		}
	}
	if len(s.merges) > 0 {
		sd.MergeCells = make([]ooxml.MergeCellData, len(s.merges))
		for i, mr := range s.merges {
			sd.MergeCells[i] = ooxml.MergeCellData{
				StartAxis: mr.Start,
				EndAxis:   mr.End,
			}
		}
	}
	return sd
}

func (s *Sheet) adjustMergedRows(deletedRow int) {
	if len(s.merges) == 0 {
		return
	}

	adjusted := s.merges[:0]
	for _, mr := range s.merges {
		col1, row1, col2, row2, err := RangeToCoordinates(mr.Start + ":" + mr.End)
		if err != nil {
			continue
		}

		switch {
		case deletedRow < row1:
			row1--
			row2--
		case deletedRow > row2:
			// unchanged
		case row1 == row2:
			continue
		default:
			row2--
			if deletedRow < row1 {
				row1--
			}
		}

		startRef, err := CoordinatesToCellName(col1, row1)
		if err != nil {
			continue
		}
		endRef, err := CoordinatesToCellName(col2, row2)
		if err != nil {
			continue
		}
		if startRef == endRef {
			continue
		}
		adjusted = append(adjusted, MergeRange{Start: startRef, End: endRef})
	}
	s.merges = adjusted
}

// resolveCell evaluates the cell's formula if it is dirty or stale.
// dirty is the primary signal from the dep graph; cachedGen is a safety net
// for formulas not yet registered in the dep graph.
func (s *Sheet) resolveCell(c *Cell, col, row int) {
	if c.formula != "" && (c.dirty || c.cachedGen < s.file.calcGen) {
		c.value = s.evaluateFormula(c, col, row)
		c.cachedGen = s.file.calcGen
		c.dirty = false
	}
}

// evaluateFormula parses, compiles, and executes the formula on the given cell.
func (s *Sheet) evaluateFormula(c *Cell, col, row int) Value {
	f := s.file
	if f.evaluating == nil {
		f.evaluating = make(map[cellKey]bool)
	}
	key := cellKey{sheet: s.name, col: col, row: row}
	if f.evaluating[key] {
		// Circular reference
		return Value{Type: TypeError, String: "#REF!"}
	}
	f.evaluating[key] = true
	defer delete(f.evaluating, key)

	cf := c.compiled
	if cf == nil {
		// Expand table structured references and defined names before parsing.
		src, err := f.expandFormula(c.formula, s.name, row)
		if err != nil {
			if c.value.Type == TypeString {
				return Value{Type: TypeString, String: "#NAME?"}
			}
			return Value{Type: TypeError, String: "#NAME?"}
		}
		node, err := formula.Parse(src)
		if err != nil {
			if c.value.Type == TypeString {
				return Value{Type: TypeString, String: "#NAME?"}
			}
			return Value{Type: TypeError, String: "#NAME?"}
		}
		compiled, err := formula.Compile(src, node)
		if err != nil {
			if c.value.Type == TypeString {
				return Value{Type: TypeString, String: "#NAME?"}
			}
			return Value{Type: TypeError, String: "#NAME?"}
		}
		c.compiled = compiled
		cf = compiled
		// Register in dep graph on first compilation.
		qc := formula.QualifiedCell{Sheet: s.name, Col: col, Row: row}
		f.deps.Register(qc, s.name, cf.Refs, cf.Ranges)
	}

	resolver := &fileResolver{file: f, currentSheet: s.name}
	ctx := &formula.EvalContext{
		CurrentCol:     col,
		CurrentRow:     row,
		CurrentSheet:   s.name,
		IsArrayFormula: c.isArrayFormula,
		Date1904:       f.date1904,
		Resolver:       resolver,
	}
	result, err := formula.Eval(cf, resolver, ctx)
	if err != nil {
		return Value{Type: TypeError, String: err.Error()}
	}
	if c.isDynamicArray {
		return s.materializeDynamicArray(c, col, row, result)
	}
	return formulaValueToValue(result, c.isArrayFormula)
}

// evaluateFormulaRaw is like evaluateFormula but returns the raw formula.Value
// without converting arrays to scalars. This is used by ANCHORARRAY to obtain
// the full dynamic array result from a cell's formula.
func (s *Sheet) evaluateFormulaRaw(c *Cell, col, row int) formula.Value {
	f := s.file
	if f.evaluating == nil {
		f.evaluating = make(map[cellKey]bool)
	}
	key := cellKey{sheet: s.name, col: col, row: row}
	if f.evaluating[key] {
		return formula.ErrorVal(formula.ErrValREF)
	}
	f.evaluating[key] = true
	defer delete(f.evaluating, key)

	cf := c.compiled
	if cf == nil {
		src, err := f.expandFormula(c.formula, s.name, row)
		if err != nil {
			return formula.ErrorVal(formula.ErrValNAME)
		}
		node, err := formula.Parse(src)
		if err != nil {
			return formula.ErrorVal(formula.ErrValNAME)
		}
		compiled, err := formula.Compile(src, node)
		if err != nil {
			return formula.ErrorVal(formula.ErrValNAME)
		}
		c.compiled = compiled
		cf = compiled
		qc := formula.QualifiedCell{Sheet: s.name, Col: col, Row: row}
		f.deps.Register(qc, s.name, cf.Refs, cf.Ranges)
	}

	resolver := &fileResolver{file: f, currentSheet: s.name}
	ctx := &formula.EvalContext{
		CurrentCol:     col,
		CurrentRow:     row,
		CurrentSheet:   s.name,
		IsArrayFormula: c.isArrayFormula,
		Date1904:       f.date1904,
		Resolver:       resolver,
	}
	result, err := formula.Eval(cf, resolver, ctx)
	if err != nil {
		return formula.ErrorVal(formula.ErrValVALUE)
	}
	return result
}

// formulaValueToValue converts a formula.Value to a werkbook Value.
// Empty formula results are coerced to 0 (a cell containing =EmptyRef
// displays and caches 0, not blank), so ValueEmpty maps to TypeNumber 0.
// isArrayFormula indicates whether the originating cell is a CSE array formula.
func formulaValueToValue(fv formula.Value, isArrayFormula bool) Value {
	switch fv.Type {
	case formula.ValueNumber:
		return Value{Type: TypeNumber, Number: fv.Num}
	case formula.ValueString:
		return Value{Type: TypeString, String: fv.Str}
	case formula.ValueBool:
		return Value{Type: TypeBool, Bool: fv.Bool}
	case formula.ValueError:
		return Value{Type: TypeError, String: fv.Err.String()}
	case formula.ValueArray:
		// Arrays marked NoSpill (e.g. INDEX with row_num=0) cannot be
		// displayed in a single non-array cell; returns #VALUE!.
		if fv.NoSpill && !isArrayFormula {
			return Value{Type: TypeError, String: "#VALUE!"}
		}
		// When an array result is reduced to a single cell, expose its
		// top-left element.
		if len(fv.Array) > 0 && len(fv.Array[0]) > 0 {
			return formulaValueToValue(fv.Array[0][0], isArrayFormula)
		}
		// Empty array — treat as numeric 0.
		return Value{Type: TypeNumber, Number: 0}
	default:
		// Empty formula results are treated as numeric 0.
		return Value{Type: TypeNumber, Number: 0}
	}
}

// fileResolver implements formula.CellResolver with cross-sheet support.
type fileResolver struct {
	file         *File
	currentSheet string // sheet name for resolving unqualified refs
}

func (fr *fileResolver) resolveSheet(name string) *Sheet {
	if name == "" {
		name = fr.currentSheet
	}
	return fr.file.Sheet(name)
}

func (fr *fileResolver) resolveExternalSheet(name string) *externalSheet {
	if name == "" {
		return nil
	}
	bookIndex, sheetName, ok := parseExternalSheetRef(name)
	if !ok || bookIndex < 0 || bookIndex >= len(fr.file.externalBooks) {
		return nil
	}
	return fr.file.externalBooks[bookIndex].sheets[externalSheetKey(sheetName)]
}

func (fr *fileResolver) GetCellValue(addr formula.CellAddr) formula.Value {
	s := fr.resolveSheet(addr.Sheet)
	if s != nil {
		r, ok := s.rows[addr.Row]
		if !ok {
			return formula.EmptyVal()
		}
		c, ok := r.cells[addr.Col]
		if !ok {
			return formula.EmptyVal()
		}
		if c.isSpillFollower() {
			if anchor := s.cellAt(c.spillParentCol, c.spillParentRow); anchor != nil {
				s.resolveCell(anchor, c.spillParentCol, c.spillParentRow)
				r = s.rows[addr.Row]
				if r == nil {
					return formula.EmptyVal()
				}
				c = r.cells[addr.Col]
				if c == nil {
					return formula.EmptyVal()
				}
			}
		}

		s.resolveCell(c, addr.Col, addr.Row)
		return valueToFormulaValue(c.value)
	}

	if es := fr.resolveExternalSheet(addr.Sheet); es != nil {
		return es.cellValue(addr.Col, addr.Row)
	}

	return formula.ErrorVal(formula.ErrValREF)
}

func newFormulaValueMatrix(nRows, nCols int) [][]formula.Value {
	if nRows <= 0 || nCols <= 0 {
		return nil
	}
	rows := make([][]formula.Value, nRows)
	cells := make([]formula.Value, nRows*nCols)
	for i := range rows {
		start := i * nCols
		rows[i] = cells[start : start+nCols]
	}
	return rows
}

func (fr *fileResolver) GetRangeValues(addr formula.RangeAddr) [][]formula.Value {
	rangeOverflow := func() [][]formula.Value {
		return [][]formula.Value{{{
			Type:          formula.ValueError,
			Err:           formula.ErrValREF,
			RangeOverflow: true,
		}}}
	}

	nCols := addr.ToCol - addr.FromCol + 1
	s := fr.resolveSheet(addr.Sheet)
	if s == nil {
		if es := fr.resolveExternalSheet(addr.Sheet); es != nil {
			return es.rangeValues(addr)
		}
		nRows := addr.ToRow - addr.FromRow + 1
		if formula.RangeCellCountExceedsLimit(nRows, nCols) {
			return rangeOverflow()
		}
		rows := newFormulaValueMatrix(nRows, nCols)
		refErr := formula.ErrorVal(formula.ErrValREF)
		for i := range rows {
			for j := range rows[i] {
				rows[i][j] = refErr
			}
		}
		return rows
	}

	// Clamp the row range to the sheet's actual data extent so that
	// full-column references (e.g. F:F → F1:F1048576) don't allocate
	// rows far beyond the populated extent. Empty sheets fall back to the
	// range's start row so full-column references still evaluate as blanks.
	toRow := addr.ToRow
	maxRow := s.MaxRow()
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
		maxCol := s.MaxCol()
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
	nCols = toCol - addr.FromCol + 1
	if formula.RangeCellCountExceedsLimit(nRows, nCols) {
		return rangeOverflow()
	}

	rows := newFormulaValueMatrix(nRows, nCols)
	for rowNum, sheetRow := range s.rows {
		if rowNum < addr.FromRow || rowNum > toRow {
			continue
		}
		row := rows[rowNum-addr.FromRow]
		for colNum, cell := range sheetRow.cells {
			if colNum < addr.FromCol || colNum > toCol {
				continue
			}
			s.resolveCell(cell, colNum, rowNum)
			row[colNum-addr.FromCol] = valueToFormulaValue(cell.value)
		}
	}
	return rows
}

// GetSheetNames returns the ordered list of all sheet names in the workbook.
func (fr *fileResolver) GetSheetNames() []string {
	return fr.file.SheetNames()
}

// IsSubtotalCell reports whether the cell at (sheet, col, row) contains a formula
// whose outermost function call is SUBTOTAL. This is used by the SUBTOTAL function
// to skip nested SUBTOTAL results and avoid double-counting.
func (fr *fileResolver) IsSubtotalCell(sheet string, col, row int) bool {
	s := fr.resolveSheet(sheet)
	if s == nil {
		return false
	}
	r, ok := s.rows[row]
	if !ok {
		return false
	}
	c, ok := r.cells[col]
	if !ok {
		return false
	}
	return isSubtotalFormula(c.formula)
}

// IsRowHidden reports whether the given row on the given sheet is hidden.
func (fr *fileResolver) IsRowHidden(sheet string, row int) bool {
	s := fr.resolveSheet(sheet)
	if s == nil {
		return false
	}
	r, ok := s.rows[row]
	if !ok {
		return false
	}
	return r.hidden
}

// IsRowFilteredByAutoFilter reports whether the given row is hidden AND falls
// within a table that has an active autoFilter (with filterColumn elements).
// This is used by SUBTOTAL(1-11) which excludes filtered rows but not manually
// hidden rows.
func (fr *fileResolver) IsRowFilteredByAutoFilter(sheet string, row int) bool {
	if !fr.IsRowHidden(sheet, row) {
		return false
	}
	// Check if the row falls within any table on this sheet that has an active filter.
	sheetLower := strings.ToLower(sheet)
	for _, t := range fr.file.tables {
		if strings.ToLower(t.SheetName) != sheetLower {
			continue
		}
		if !t.HasActiveFilter {
			continue
		}
		dataFirst := t.DataFirstRow()
		dataLast := t.DataLastRow()
		if row >= dataFirst && row <= dataLast {
			return true
		}
	}
	return false
}

// EvalCellFormula evaluates the formula in the cell at (sheet, col, row) and
// returns the full formula.Value result. For dynamic array formulas this
// returns the complete ValueArray. If the cell has no formula, it returns the
// cell's scalar value.
func (fr *fileResolver) EvalCellFormula(sheet string, col, row int) formula.Value {
	s := fr.resolveSheet(sheet)
	if s == nil {
		return formula.ErrorVal(formula.ErrValREF)
	}
	r, ok := s.rows[row]
	if !ok {
		return formula.EmptyVal()
	}
	c, ok := r.cells[col]
	if !ok {
		return formula.EmptyVal()
	}
	if c.formula == "" {
		return valueToFormulaValue(c.value)
	}
	return s.evaluateFormulaRaw(c, col, row)
}

// HasFormula reports whether the cell at (sheet, col, row) contains a formula.
func (fr *fileResolver) HasFormula(sheet string, col, row int) bool {
	s := fr.resolveSheet(sheet)
	if s == nil {
		return false
	}
	r, ok := s.rows[row]
	if !ok {
		return false
	}
	c, ok := r.cells[col]
	if !ok {
		return false
	}
	return c.formula != ""
}

// GetFormulaText returns the formula text for the cell at (sheet, col, row),
// or "" if the cell has no formula. The returned text does not include the
// leading '=' sign.
func (fr *fileResolver) GetFormulaText(sheet string, col, row int) string {
	s := fr.resolveSheet(sheet)
	if s == nil {
		return ""
	}
	r, ok := s.rows[row]
	if !ok {
		return ""
	}
	c, ok := r.cells[col]
	if !ok {
		return ""
	}
	return c.formula
}

// isSubtotalFormula returns true if the formula string starts with "SUBTOTAL("
// (case-insensitive), with optional leading whitespace. This matches both
// "SUBTOTAL(...)" and "_xlfn.SUBTOTAL(...)".
func isSubtotalFormula(f string) bool {
	if f == "" {
		return false
	}
	upper := strings.ToUpper(strings.TrimSpace(f))
	return strings.HasPrefix(upper, "SUBTOTAL(") || strings.HasPrefix(upper, "_XLFN.SUBTOTAL(")
}

// valueToFormulaValue converts a werkbook Value to a formula.Value.
func valueToFormulaValue(v Value) formula.Value {
	switch v.Type {
	case TypeNumber:
		return formula.NumberVal(v.Number)
	case TypeString:
		return formula.StringVal(v.String)
	case TypeBool:
		return formula.BoolVal(v.Bool)
	case TypeError:
		return formula.ErrorVal(formula.ErrorValueFromString(v.String))
	default:
		return formula.EmptyVal()
	}
}

func cellToData(ref string, v Value, f string, isArrayFormula bool, isDynamicArray bool, formulaRef string) ooxml.CellData {
	cd := ooxml.CellData{Ref: ref}
	if isDynamicArray {
		cd.FormulaType = "array"
		if formulaRef != "" {
			cd.FormulaRef = formulaRef
		} else {
			cd.FormulaRef = ref
		}
		cd.IsDynamicArray = true
	} else if isArrayFormula {
		cd.FormulaType = "array"
		if formulaRef != "" {
			cd.FormulaRef = formulaRef
		} else {
			cd.FormulaRef = ref
		}
		cd.IsArrayFormula = true
	}

	switch v.Type {
	case TypeString:
		if f != "" {
			cd.Type = "str"
			cd.Value = v.String
		} else {
			cd.Type = "s"
			cd.Value = v.String
		}
	case TypeNumber:
		cd.Value = strconv.FormatFloat(v.Number, 'f', -1, 64)
	case TypeBool:
		val := "0"
		if v.Bool {
			val = "1"
		}
		cd.Type = "b"
		cd.Value = val
	case TypeError:
		if f != "" && (isDynamicArray || isArrayFormula || !isLegacyFormulaError(v.String)) {
			cd.Type = "str"
		} else {
			cd.Type = "e"
		}
		cd.Value = v.String
	}
	cd.Formula = formula.AddXlfnPrefixes(f)
	return cd
}

func isLegacyFormulaError(err string) bool {
	switch err {
	case "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!":
		return true
	default:
		return false
	}
}
