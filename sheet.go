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
	spill     spillOverlay
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
	s.clearSpillState(col, row)
	// Unregister old formula if any.
	if c.formula != "" {
		s.file.deps.Unregister(formula.QualifiedCell{Sheet: s.name, Col: col, Row: row})
	}
	c.value = val
	c.formula = ""
	c.isArrayFormula = false
	c.dynamicArraySpill = false
	c.compiled = nil
	c.rawValue = formula.Value{}
	c.cachedGen = 0
	c.rawCachedGen = 0
	s.file.calcGen++
	s.file.invalidateDependents(s.name, col, row)
	return nil
}

// SetFormula sets a formula on a cell by reference (e.g. "A1").
// A single leading '=' (as users often type in Excel) is tolerated and stripped;
// leaving it in would nest inside the OOXML <f> element and save as a #NAME? cell.
func (s *Sheet) SetFormula(cell string, f string) error {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}
	f = strings.TrimPrefix(f, "=")
	src, err := s.file.expandFormula(f, s.name, row)
	if err != nil {
		return err
	}
	r := s.ensureRow(row)
	c := r.ensureCell(col)
	s.clearSpillState(col, row)
	// Unregister old formula if any.
	qc := formula.QualifiedCell{Sheet: s.name, Col: col, Row: row}
	if c.formula != "" {
		s.file.deps.Unregister(qc)
	}
	c.formula = f
	c.isArrayFormula = false
	c.dynamicArraySpill = formula.IsDynamicArrayFormula(f)
	c.compiled = nil
	c.value = Value{}
	c.rawValue = formula.Value{}
	c.cachedGen = 0
	c.rawCachedGen = 0
	c.dirty = true
	s.file.calcGen++
	// Compile and register in dep graph.
	node, parseErr := formula.Parse(src)
	if parseErr == nil {
		cf, compErr := formula.Compile(src, node)
		if compErr == nil {
			c.compiled = cf
			c.dynamicArraySpill = formulaNeedsSpillAnchor(c.formula, cf)
			s.file.deps.Register(qc, s.name, cf.Refs, cf.Ranges)
		}
	}
	s.file.invalidateDependents(s.name, col, row)
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

// SpillBounds returns the spill range for a dynamic-array anchor cell. After
// recalculation it returns the published bounds; before recalculation it falls
// back to the OOXML formula ref imported from the file. Returns zeroes and
// false when the cell is not a spill anchor or has no spill range.
func (s *Sheet) SpillBounds(col, row int) (toCol, toRow int, ok bool) {
	state, exists := s.spillState(col, row)
	if !exists {
		return 0, 0, false
	}
	if spillStateHasPublishedSpill(state, col, row, s.file.calcGen) {
		return state.publishedToCol, state.publishedToRow, true
	}
	// Fall back to imported OOXML formula ref (pre-recalculation).
	if state.formulaRef == "" {
		return 0, 0, false
	}
	_, _, toCol, toRow, err := RangeToCoordinates(state.formulaRef)
	if err != nil {
		return 0, 0, false
	}
	if toCol <= col && toRow <= row {
		return 0, 0, false
	}
	return toCol, toRow, true
}

// GetValue returns the value of a cell by reference (e.g. "A1").
func (s *Sheet) GetValue(cell string) (Value, error) {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return Value{}, fmt.Errorf("%w: %v", ErrInvalidCellRef, err)
	}

	if v, ok := s.valueAt(col, row); ok {
		return v, nil
	}
	return Value{Type: TypeEmpty}, nil
}

func (s *Sheet) valueAt(col, row int) (Value, bool) {
	if r, ok := s.rows[row]; ok {
		if c, ok := r.cells[col]; ok {
			s.resolveCell(c, col, row)
			if !cellOccupiesSpillSlot(c) {
				if v, ok := s.spillValueAt(col, row); ok {
					return v, true
				}
			}
			return c.value, true
		}
	}
	if v, ok := s.spillValueAt(col, row); ok {
		return v, true
	}
	return Value{}, false
}

func (s *Sheet) spillValueAt(col, row int) (Value, bool) {
	fv, ok := s.spillFormulaValueAt(col, row)
	if !ok {
		return Value{}, false
	}
	return formulaValueToValue(fv, false), true
}

func (s *Sheet) spillFormulaValueAt(col, row int) (formula.Value, bool) {
	span := s.ensureSpillOverlay().index.lookup(col, row)
	if span == nil {
		if !s.refreshSpillAnchorsForPoint(col, row) {
			return formula.Value{}, false
		}
		span = s.ensureSpillOverlay().index.lookup(col, row)
		if span == nil {
			return formula.Value{}, false
		}
	}
	raw := span.anchor.cell.rawValue
	rowOffset := row - span.anchor.row
	colOffset := col - span.anchor.col
	return spillArrayCell(raw, rowOffset, colOffset)
}

// hasSpillConflict checks whether any cell in the proposed spill range
// (excluding the anchor at anchorCol,anchorRow) is occupied by a non-empty
// value or formula. Returns true if at least one cell would block the spill.
func (s *Sheet) hasSpillConflict(anchorCol, anchorRow int, arr [][]formula.Value) bool {
	for rowOff, arrRow := range arr {
		for colOff := range arrRow {
			if rowOff == 0 && colOff == 0 {
				continue // skip anchor
			}
			r, ok := s.rows[anchorRow+rowOff]
			if !ok {
				continue
			}
			c, ok := r.cells[anchorCol+colOff]
			if !ok {
				continue
			}
			if cellOccupiesSpillSlot(c) {
				return true
			}
		}
	}
	return false
}

// cellOccupiesSpillSlot reports whether a physical cell should block
// dynamic-array spill semantics. Imported OOXML can contain empty placeholder
// cells inside a spill range; those cells must behave like absent cells so
// direct references and range materialization can still see the spill result.
func cellOccupiesSpillSlot(c *Cell) bool {
	if c == nil {
		return false
	}
	return c.formula != "" || c.value.Type != TypeEmpty
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

	// Cached spill follower values, indexed by row then column. These are
	// emitted as physical "shadow" cells alongside the regular cell stream
	// so that consumers which cannot evaluate dynamic-array functions
	// (notably Apple Numbers) can still display each spilled row.
	shadowsByRow := s.collectSpillShadowsByRow()

	// Build the union of row numbers from physical rows and shadow-only rows.
	rowSet := make(map[int]struct{}, len(s.rows)+len(shadowsByRow))
	for n := range s.rows {
		rowSet[n] = struct{}{}
	}
	for n := range shadowsByRow {
		rowSet[n] = struct{}{}
	}
	rowNums := make([]int, 0, len(rowSet))
	for n := range rowSet {
		rowNums = append(rowNums, n)
	}
	sort.Ints(rowNums)

	for _, rn := range rowNums {
		r, hasRow := s.rows[rn]
		rowShadows := shadowsByRow[rn]
		if !hasRow && len(rowShadows) == 0 {
			continue
		}
		var (
			rowHeight float64
			rowHidden bool
			rowCells  map[int]*Cell
		)
		if hasRow {
			rowHeight = r.height
			rowHidden = r.hidden
			rowCells = r.cells
		}
		if len(rowCells) == 0 && len(rowShadows) == 0 && rowHeight == 0 && !rowHidden {
			continue
		}
		rd := ooxml.RowData{Num: rn, Height: rowHeight, Hidden: rowHidden}

		// Build the union of column numbers from physical cells and
		// shadow cells in this row.
		colSet := make(map[int]struct{}, len(rowCells)+len(rowShadows))
		for c := range rowCells {
			colSet[c] = struct{}{}
		}
		for c := range rowShadows {
			colSet[c] = struct{}{}
		}
		colNums := make([]int, 0, len(colSet))
		for c := range colSet {
			colNums = append(colNums, c)
		}
		sort.Ints(colNums)

		for _, cn := range colNums {
			c, hasCell := rowCells[cn]
			shadowVal, hasShadow := rowShadows[cn]
			if !hasCell {
				// Shadow-only cell: emit a value-bearing follower with no
				// formula text but mark it as a spill follower so the writer
				// emits an empty <f ca="1"/> alongside the cached value.
				ref, _ := CoordinatesToCellName(cn, rn)
				cd := cellToData(ref, shadowVal, "", false, "", false)
				cd.IsSpillFollower = true
				rd.Cells = append(rd.Cells, cd)
				continue
			}
			// Resolve dirty/stale formulas before serializing.
			s.resolveCell(c, cn, rn)
			if c.value.Type == TypeEmpty && c.formula == "" && c.style == nil && !hasShadow {
				continue
			}
			ref, _ := CoordinatesToCellName(cn, rn)
			formulaRef := ""
			if state, ok := s.spillState(cn, rn); ok {
				formulaRef = state.formulaRef
			}
			saveValue := c.value
			injectShadow := false
			if c.dynamicArraySpill && !c.isArrayFormula {
				if c.value.Type == TypeError && c.value.String == "#SPILL!" {
					// #SPILL! means the spill failed. Write as plain formula
					// with no cached value — Excel re-detects and recomputes.
					formulaRef = ""
					saveValue = Value{}
				} else {
					formulaRef = s.dynamicArrayFormulaRef(ref, cn, rn, c)
				}
			} else if hasShadow && c.formula == "" && c.value.Type == TypeEmpty {
				// Empty placeholder cell (typically a styled stub from the
				// template) sitting inside a published spill range. Inject
				// the cached spill value so downstream readers can display
				// the spilled row.
				saveValue = shadowVal
				injectShadow = true
			}
			isDynamicArray := false
			if c.dynamicArraySpill && !c.isArrayFormula {
				isDynamicArray = formula.IsDynamicArrayFormula(c.formula) || formulaRef != ""
				if c.value.Type == TypeError && c.value.String == "#SPILL!" {
					isDynamicArray = true
				}
			}
			cd := cellToData(ref, saveValue, c.formula, c.isArrayFormula, formulaRef, isDynamicArray)
			if injectShadow {
				cd.IsSpillFollower = true
			}

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

// collectSpillShadowsByRow returns the cached spill follower values for
// every dynamic-array anchor on this sheet, indexed by row then column.
// The anchor cell itself is excluded; the returned map only contains
// follower positions.
//
// Spill follower cells are not stored as physical cells in werkbook —
// they live in the spill overlay and are recomputed on demand. Excel
// re-evaluates the anchor's formula on load so it does not need follower
// values persisted, but other consumers (notably Apple Numbers) cannot
// evaluate `_xlfn._xlws.FILTER`/SORT/etc. and rely on these cached values
// to render each spilled row.
func (s *Sheet) collectSpillShadowsByRow() map[int]map[int]Value {
	overlay := s.ensureSpillOverlay()
	if overlay == nil || len(overlay.index.spans) == 0 {
		return nil
	}
	shadows := make(map[int]map[int]Value)
	for i := range overlay.index.spans {
		span := &overlay.index.spans[i]
		raw := span.anchor.cell.rawValue
		rows, ok := spillArrayRows(raw)
		if !ok {
			continue
		}
		for rowOff, arrRow := range rows {
			row := span.anchor.row + rowOff
			if row > span.toRow {
				break
			}
			for colOff := range arrRow {
				col := span.anchor.col + colOff
				if col > span.toCol {
					break
				}
				if col == span.anchor.col && row == span.anchor.row {
					continue
				}
				fv, ok := spillArrayCell(raw, rowOff, colOff)
				if !ok {
					continue
				}
				rowShadows := shadows[row]
				if rowShadows == nil {
					rowShadows = make(map[int]Value)
					shadows[row] = rowShadows
				}
				rowShadows[col] = formulaValueToValue(fv, false)
			}
		}
	}
	return shadows
}

func (s *Sheet) dynamicArrayFormulaRef(anchorRef string, anchorCol, anchorRow int, c *Cell) string {
	if c == nil {
		return anchorRef
	}
	s.refreshSpillState(c, anchorCol, anchorRow)
	// #SPILL! means the spill failed — don't claim a spill range.
	if c.value.Type == TypeError && c.value.String == "#SPILL!" {
		return ""
	}
	state, _ := s.spillState(anchorCol, anchorRow)
	if ref, ok := spillStateFormulaRef(state, anchorCol, anchorRow); ok {
		return ref
	}
	// Imported workbooks can carry valid dynamic-array metadata even when the
	// current engine cannot derive a fresh spill result. Preserve that metadata
	// as a fallback rather than discarding it on save.
	if state != nil && state.formulaRef != "" {
		return state.formulaRef
	}
	if formula.IsDynamicArrayFormula(c.formula) {
		return anchorRef
	}
	return ""
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

func formulaNeedsSpillAnchor(f string, cf *formula.CompiledFormula) bool {
	return formula.IsDynamicArrayFormula(f) || (cf != nil && cf.TopLevelArray != nil)
}

func (s *Sheet) compileCellFormula(c *Cell, col, row int) (*formula.CompiledFormula, error) {
	if c.compiled != nil {
		return c.compiled, nil
	}
	f := s.file
	src, err := f.expandFormula(c.formula, s.name, row)
	if err != nil {
		return nil, err
	}
	node, err := formula.Parse(src)
	if err != nil {
		return nil, err
	}
	compiled, err := formula.Compile(src, node)
	if err != nil {
		return nil, err
	}
	c.compiled = compiled
	qc := formula.QualifiedCell{Sheet: s.name, Col: col, Row: row}
	f.deps.Register(qc, s.name, compiled.Refs, compiled.Ranges)
	return compiled, nil
}

func (s *Sheet) evalCellFormula(c *Cell, col, row int, topLevelArray bool) (formula.Value, error) {
	f := s.file
	cf, err := s.compileCellFormula(c, col, row)
	if err != nil {
		// compileCellFormula wraps expandFormula, Parse, and Compile;
		// classify as VALUE for size overflow and NAME for parse/compile.
		return formula.Value{}, formula.WrapEvalError(formula.ErrorValueFromErr(err), err)
	}
	if topLevelArray && cf.TopLevelArray != nil {
		cf = cf.TopLevelArray
	}

	qc := formula.QualifiedCell{Sheet: s.name, Col: col, Row: row}
	resolver := &fileResolver{
		file:             f,
		currentSheet:     s.name,
		currentCell:      qc,
		trackDynamicDeps: true,
	}
	ctx := &formula.EvalContext{
		CurrentCol:     col,
		CurrentRow:     row,
		CurrentSheet:   s.name,
		IsArrayFormula: c.isArrayFormula,
		Date1904:       f.date1904,
		Resolver:       resolver,
	}
	result, err := formula.Eval(cf, resolver, ctx)
	f.deps.SetDynamicRanges(qc, formula.DynamicRangeKindMaterialized, resolver.materializedDeps)
	// Runtime evaluation failures are engine bugs or resource exhaustion
	// (stack underflow, bad opcode); Excel would surface these as #VALUE!.
	return result, formula.WrapEvalError(formula.ErrValVALUE, err)
}

func (s *Sheet) evalCellOutcome(c *Cell, col, row int) (CellEvalOutcome, error) {
	displayRaw, err := s.evalCellFormula(c, col, row, false)
	if err != nil {
		return CellEvalOutcome{}, err
	}
	raw := displayRaw
	var spill *SpillPlan
	if c.dynamicArraySpill && !c.isArrayFormula {
		if !canPublishSpill(displayRaw) && c.compiled != nil && c.compiled.TopLevelArray != nil {
			if probe, probeErr := s.evalCellFormula(c, col, row, true); probeErr == nil && canPublishSpill(probe) {
				raw = probe
			}
		}
		spill = newSpillPlan(raw, col, row)
		if spillRows, ok := spillArrayRows(raw); spill != nil && ok && s.hasSpillConflict(col, row, spillRows) {
			spill.Blocked = true
			spill.PublishedToCol = col
			spill.PublishedToRow = row
			return newCellEvalOutcome(raw, formula.ErrorVal(formula.ErrValSPILL), spill), nil
		}
	}
	displaySource := displayRaw
	if c.dynamicArraySpill && !c.isArrayFormula {
		displaySource = raw
	}
	// When the cell carries a spill anchor but the formula did not actually
	// publish a spillable array (e.g. INDEX(range,0,col) collapsing through
	// implicit intersection), fall back to legacy implicit intersection at
	// the formula cell so the displayed value matches Excel's cached scalar.
	intersectRangeOrigin := !c.dynamicArraySpill || spill == nil
	return newCellEvalOutcome(
		raw,
		formulaDisplayValueAt(displaySource, c.isArrayFormula, intersectRangeOrigin, col, row),
		spill,
	), nil
}

func canPublishSpill(raw formula.Value) bool {
	_, ok := spillArrayRows(raw)
	return ok
}

func (s *Sheet) applyCellOutcome(c *Cell, col, row int, outcome CellEvalOutcome) {
	raw := outcome.RawValue()
	c.rawValue = raw
	c.rawCachedGen = s.file.calcGen
	c.value = formulaValueToValue(outcome.Display, false)
	c.cachedGen = s.file.calcGen
	c.dirty = false
	if c.dynamicArraySpill && !c.isArrayFormula {
		s.publishSpillState(col, row, outcome.Spill)
	} else {
		s.clearSpillState(col, row)
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

	outcome, err := s.evalCellOutcome(c, col, row)
	if err != nil {
		// Classify the failure (parse/compile → #NAME?; expansion overflow
		// or runtime engine failure → #VALUE!) instead of collapsing every
		// error to #NAME?. The underlying message stays reachable via
		// errors.Is / errors.As on an *formula.EvalError.
		code := formula.ErrorValueFromErr(err).String()
		if c.value.Type == TypeString {
			return Value{Type: TypeString, String: code}
		}
		return Value{Type: TypeError, String: code}
	}
	s.applyCellOutcome(c, col, row, outcome)
	return c.value
}

// evaluateFormulaRaw is like evaluateFormula but returns the raw formula.Value
// without converting arrays to scalars. This is used by ANCHORARRAY to obtain
// the full dynamic array result from a cell's formula.
func (s *Sheet) evaluateFormulaRaw(c *Cell, col, row int) formula.Value {
	f := s.file
	if !c.dirty && c.rawCachedGen == f.calcGen {
		return c.rawValue
	}
	if f.evaluating == nil {
		f.evaluating = make(map[cellKey]bool)
	}
	key := cellKey{sheet: s.name, col: col, row: row}
	if f.evaluating[key] {
		return formula.ErrorVal(formula.ErrValREF)
	}
	f.evaluating[key] = true
	defer delete(f.evaluating, key)

	outcome, err := s.evalCellOutcome(c, col, row)
	if err != nil {
		return formula.ErrorVal(formula.ErrValVALUE)
	}
	s.applyCellOutcome(c, col, row, outcome)
	return c.rawValue
}

// formulaValueToValue converts a formula.Value to a werkbook Value.
// Empty formula results are coerced to 0 (a cell containing =EmptyRef
// displays and caches 0, not blank), so ValueEmpty maps to TypeNumber 0.
// isArrayFormula indicates whether the originating cell is a CSE array formula.
func formulaValueToValue(fv formula.Value, isArrayFormula bool) Value {
	fv = formulaDisplayValue(fv, isArrayFormula)
	switch fv.Type {
	case formula.ValueNumber:
		return Value{Type: TypeNumber, Number: fv.Num}
	case formula.ValueString:
		return Value{Type: TypeString, String: fv.Str}
	case formula.ValueBool:
		return Value{Type: TypeBool, Bool: fv.Bool}
	case formula.ValueError:
		return Value{Type: TypeError, String: fv.Err.String()}
	default:
		// Empty formula results are treated as numeric 0.
		return Value{Type: TypeNumber, Number: 0}
	}
}

// fileResolver implements formula.CellResolver with cross-sheet support.
type fileResolver struct {
	file             *File
	currentSheet     string // sheet name for resolving unqualified refs
	currentCell      formula.QualifiedCell
	trackDynamicDeps bool
	materializedDeps []formula.RangeAddr
}

func (fr *fileResolver) resolveSheet(name string) *Sheet {
	if name == "" {
		name = fr.currentSheet
	}
	return fr.file.Sheet(name)
}

func (fr *fileResolver) GetCellValue(addr formula.CellAddr) formula.Value {
	s := fr.resolveSheet(addr.Sheet)
	if s == nil {
		return formula.ErrorVal(formula.ErrValREF)
	}

	r, ok := s.rows[addr.Row]
	if !ok {
		if v, ok := s.spillFormulaValueAt(addr.Col, addr.Row); ok {
			return v
		}
		return formula.EmptyVal()
	}
	c, ok := r.cells[addr.Col]
	if !ok {
		if v, ok := s.spillFormulaValueAt(addr.Col, addr.Row); ok {
			return v
		}
		return formula.EmptyVal()
	}

	s.resolveCell(c, addr.Col, addr.Row)
	if !cellOccupiesSpillSlot(c) {
		if v, ok := s.spillFormulaValueAt(addr.Col, addr.Row); ok {
			return v
		}
	}
	return valueToFormulaValue(c.value)
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
	s := fr.resolveSheet(addr.Sheet)
	if s == nil {
		res := materializeRange(rangeMaterializationRequest{
			sheet:   addr.Sheet,
			fromCol: addr.FromCol,
			fromRow: addr.FromRow,
			toCol:   addr.ToCol,
			toRow:   addr.ToRow,
		}, nil, nil)
		return res.cells
	}

	req := rangeMaterializationRequest{
		sheet:   addr.Sheet,
		fromCol: addr.FromCol,
		fromRow: addr.FromRow,
		toCol:   addr.ToCol,
		toRow:   addr.ToRow,
	}
	maxRow := s.MaxRow()
	maxCol := s.MaxCol()
	grid := sheetRangeGrid{
		sheet:  s,
		maxRow: maxRow,
		maxCol: maxCol,
	}

	s.refreshSpillAnchorsForRange(req, req.toCol)
	overlay := s.ensureSpillOverlay()
	var spills rangeSpillReader
	if len(overlay.index.anchors) > 0 {
		spills = &overlay.index
	}

	res := materializeRange(req, grid, spills)
	fr.appendMaterializedDeps(res.discoveredDeps)
	return res.cells
}

func (fr *fileResolver) ResolveDefinedNameValue(name, scopeSheet string) (formula.Value, bool) {
	sheetIdx := -1
	if scopeSheet != "" {
		sheetIdx = fr.file.SheetIndex(scopeSheet)
	}
	ref, err := fr.file.lookupDefinedName(name, sheetIdx)
	if err != nil {
		return formula.Value{}, false
	}

	sheetName, cellRef, err := parseDefinedNameRef(ref)
	if err != nil {
		return formula.ErrorVal(formula.ErrValREF), true
	}
	area, err := parseDefinedNameArea(sheetName, cellRef)
	if err != nil {
		return formula.ErrorVal(formula.ErrValREF), true
	}
	if area.isRange {
		rows, err := fr.file.resolveDefinedNameRange(area.rangeAddr)
		if err != nil {
			return formula.ErrorVal(formula.ErrValREF), true
		}
		formulaRows := make([][]formula.Value, len(rows))
		for i, row := range rows {
			formulaRows[i] = make([]formula.Value, len(row))
			for j, cell := range row {
				formulaRows[i][j] = valueToFormulaValue(cell)
			}
		}
		origin := area.rangeAddr
		return formula.Value{Type: formula.ValueArray, Array: formulaRows, RangeOrigin: &origin}, true
	}

	s := fr.file.Sheet(sheetName)
	if s == nil {
		return formula.ErrorVal(formula.ErrValREF), true
	}
	cellRefText, err := CoordinatesToCellName(area.cellCol, area.cellRow)
	if err != nil {
		return formula.ErrorVal(formula.ErrValREF), true
	}
	v, err := s.GetValue(cellRefText)
	if err != nil {
		return formula.ErrorVal(formula.ErrValREF), true
	}
	out := valueToFormulaValue(v)
	out.CellOrigin = &formula.CellAddr{Sheet: sheetName, Col: area.cellCol, Row: area.cellRow}
	return out, true
}

func (fr *fileResolver) appendMaterializedDeps(ranges []formula.RangeAddr) {
	if !fr.trackDynamicDeps || len(ranges) == 0 {
		return
	}
	for _, rng := range ranges {
		if rng.Sheet == "" {
			rng.Sheet = fr.currentSheet
		}
		fr.materializedDeps = append(fr.materializedDeps, rng)
	}
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

func cellToData(ref string, v Value, f string, isArrayFormula bool, formulaRef string, isDynamicArray bool) ooxml.CellData {
	cd := ooxml.CellData{Ref: ref}
	if isArrayFormula {
		cd.FormulaType = "array"
		if formulaRef != "" {
			cd.FormulaRef = formulaRef
		} else {
			cd.FormulaRef = ref
		}
		cd.IsArrayFormula = true
	}
	if !isArrayFormula && isDynamicArray {
		cd.IsDynamicArray = true
		if formulaRef != "" {
			cd.FormulaType = "array"
			cd.FormulaRef = formulaRef
		}
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
		if f != "" && (isArrayFormula || !isLegacyFormulaError(v.String)) {
			cd.Type = "str"
		} else {
			cd.Type = "e"
		}
		cd.Value = v.String
	}
	// Future-function OOXML uses two layers of prefixes. We first add the
	// function prefix (_xlfn./_xlfn._xlws.), then decorate LET/LAMBDA parameter
	// identifiers so LET(x,5,x+1) becomes _xlfn.LET(_xlpm.x,5,_xlpm.x+1).
	cd.Formula = formula.AddXlpmPrefixes(formula.AddXlfnPrefixes(f))
	return cd
}

func isLegacyFormulaError(err string) bool {
	switch err {
	case "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#SPILL!":
		return true
	default:
		return false
	}
}
