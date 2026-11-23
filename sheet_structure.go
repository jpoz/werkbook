package werkbook

import (
	"errors"
	"fmt"
	"strconv"
	"strings"

	"github.com/jpoz/werkbook/formula"
)

// This file holds insert and delete of rows and columns. One function,
// applyAxisEdit, does all four operations. The rule for what happens to a
// span of rows or columns is formula.StructuralEdit.ShiftSpan; cells, merged
// ranges, tables, spill ranges and formula references all use it, so they
// cannot disagree.
//
// Not adjusted in this stage: the passthrough XML of the sheet (conditional
// formats, data validation, hyperlinks, autoFilter, drawings). See
// Sheet.PassthroughElementNames.

// StructuralUndo reverses one InsertRows, DeleteRows, InsertCols or
// DeleteCols call. It holds what the edit removed or changed in a way that
// the inverse edit cannot calculate: removed rows and cells, the old text of
// formulas, and the old merged ranges, column widths, defined names, tables
// and spill ranges.
//
// Undo is valid only when the workbook has had no other change after the
// edit (a strict last-in, first-out order). Undo checks this with the change
// mark of the workbook (File.ChangeMark) and returns an error when the
// workbook changed. A caller that keeps a history must undo later changes
// first, and must give the workbook its older mark again
// (File.RestoreChangeMark) when it reversed those changes by other means than
// a StructuralUndo.
type StructuralUndo struct {
	sheet *Sheet
	edit  formula.StructuralEdit
	done  bool

	// removedRows holds the rows that a row edit removed, by their old row
	// number: the rows of a deleted band, or the rows that an insert pushed
	// off the sheet.
	removedRows map[int]*Row
	// removedCells holds the cells that a column edit removed.
	removedCells []removedCell
	// addedCells holds the cells that the edit made (header cells of new
	// table columns).
	addedCells []cellKey

	formulas     []formulaChange
	merges       []MergeRange
	colWidths    map[int]float64 // column edits only
	definedNames []formula.DefinedNameInfo
	tables       []formula.TableInfo
	tableDefs    []Table
	spillAnchors map[cellKey]spillAnchorState

	frozenRows, frozenCols int

	// markBefore and markAfter are the change marks of the workbook before
	// and after the edit. Undo runs only at markAfter, and gives the
	// workbook markBefore again.
	markBefore, markAfter uint64
}

type removedCell struct {
	row, col int
	cell     *Cell
}

type formulaChange struct {
	cell *Cell
	old  string
}

// Cost returns the size of the undo data: the number of removed cells plus
// the number of formulas whose text changed. A history can use it as the
// memory cost of the step.
func (u *StructuralUndo) Cost() int {
	n := len(u.removedCells) + len(u.formulas)
	for _, r := range u.removedRows {
		n += len(r.cells)
	}
	return n
}

// InsertRows inserts count empty rows at row. The rows that were at row and
// below move down. Formulas in all sheets, defined names, merged ranges,
// tables and spill ranges follow. It fails, with no change, when the insert
// would push a non-empty cell, a merged range or a table off the sheet.
func (s *Sheet) InsertRows(row, count int) (*StructuralUndo, error) {
	return s.applyAxisEdit(false, row, count, false)
}

// DeleteRows removes the rows row .. row+count-1. The rows below move up.
// References to deleted cells become #REF!. It fails, with no change, when
// the band has the header row of a table and not the full table.
func (s *Sheet) DeleteRows(row, count int) (*StructuralUndo, error) {
	return s.applyAxisEdit(false, row, count, true)
}

// InsertCols inserts count empty columns at col. See InsertRows.
func (s *Sheet) InsertCols(col, count int) (*StructuralUndo, error) {
	return s.applyAxisEdit(true, col, count, false)
}

// DeleteCols removes the columns col .. col+count-1. See DeleteRows.
func (s *Sheet) DeleteCols(col, count int) (*StructuralUndo, error) {
	return s.applyAxisEdit(true, col, count, true)
}

// RemoveRow removes a row and shifts all following rows up by one. It is
// DeleteRows(row, 1): formulas, defined names, merged ranges and tables
// follow.
func (s *Sheet) RemoveRow(row int) error {
	_, err := s.DeleteRows(row, 1)
	return err
}

func axisName(cols bool) string {
	if cols {
		return "column"
	}
	return "row"
}

func axisLast(cols bool) int {
	if cols {
		return MaxColumns
	}
	return MaxRows
}

// applyAxisEdit does an insert or a delete of rows or columns. All checks
// run before the first change, so an error leaves the workbook unchanged.
func (s *Sheet) applyAxisEdit(cols bool, at, count int, del bool) (*StructuralUndo, error) {
	last := axisLast(cols)
	if at < 1 || at > last {
		return nil, fmt.Errorf("%w: %s %d out of range [1, %d]", ErrInvalidCellRef, axisName(cols), at, last)
	}
	if count < 1 || at+count-1 > last {
		return nil, fmt.Errorf("%w: %d %ss at %d do not fit in [1, %d]", ErrInvalidCellRef, count, axisName(cols), at, last)
	}
	f := s.file
	e := formula.StructuralEdit{
		Sheet:      s.name,
		Cols:       cols,
		At:         at,
		Count:      count,
		Delete:     del,
		SheetNames: f.SheetNames(),
	}

	// Checks and the new state of the parts that can fail.
	if !del {
		if err := s.checkInsertFits(e); err != nil {
			return nil, err
		}
	}
	merges, err := s.shiftedMerges(e)
	if err != nil {
		return nil, err
	}
	tableDefs, tables, newHeaders, err := s.shiftedTables(e)
	if err != nil {
		return nil, err
	}
	lostColumns := s.removedTableColumns(e, tableDefs)

	u := &StructuralUndo{
		sheet:        s,
		edit:         e,
		merges:       s.merges,
		definedNames: append([]formula.DefinedNameInfo(nil), f.definedNames...),
		tables:       f.tables,
		tableDefs:    f.tableDefs,
		spillAnchors: s.copySpillAnchors(),
		frozenRows:   s.frozenRows,
		frozenCols:   s.frozenCols,
		markBefore:   f.changes,
	}
	if cols {
		u.colWidths = s.colWidths
	}

	// (a) Cells and rows, with row heights and hidden flags.
	s.shiftCells(e, u)

	// Formula text in all sheets. The cells of the edited sheet are at
	// their new positions; the rewrite does not depend on the position.
	for _, sh := range f.sheets {
		for _, r := range sh.rows {
			for _, c := range r.cells {
				if c.formula == "" {
					continue
				}
				out, _ := formula.ApplyStructuralEdit(c.formula, sh.name, e)
				// A structured reference to a table column that the edit
				// deleted becomes #REF!, as a cell reference does.
				for _, lost := range lostColumns {
					inTable := sh == s && lost.rect.contains(c.col, r.num)
					out, _ = formula.InvalidateTableColumnRefs(out, lost.table, lost.columns, inTable)
				}
				if out != c.formula {
					u.formulas = append(u.formulas, formulaChange{cell: c, old: c.formula})
					c.formula = out
				}
			}
		}
	}

	// (b) Merged ranges.
	s.merges = merges

	// (c) Column widths.
	if cols {
		widths := make(map[int]float64, len(s.colWidths))
		for col, w := range s.colWidths {
			if col2, _, gone := e.ShiftSpan(col, col); !gone {
				widths[col2] = w
			}
		}
		s.colWidths = widths
	}

	// (d) Defined names. A reference with no sheet name in a name that has
	// a sheet scope refers to the scope sheet.
	for i := range f.definedNames {
		dn := &f.definedNames[i]
		scope := ""
		if dn.LocalSheetID >= 0 && dn.LocalSheetID < len(f.sheetNames) {
			scope = f.sheetNames[dn.LocalSheetID]
		}
		if out, changed := formula.ApplyStructuralEdit(dn.Value, scope, e); changed {
			dn.Value = out
		}
	}

	// (e) Tables. A new table column gets its name as header text, as the
	// table part and the header cell must agree.
	f.tableDefs, f.tables = tableDefs, tables
	for _, h := range newHeaders {
		c := s.ensureRow(h.row).ensureCell(h.col)
		c.value = Value{Type: TypeString, String: h.name}
		u.addedCells = append(u.addedCells, cellKey{sheet: s.name, col: h.col, row: h.row})
	}

	// (f) Spill ranges move with their anchors.
	s.shiftSpillAnchors(e)

	// The frozen pane is the span 1..n of rows or columns: an insert in it
	// makes it larger and a delete in it makes it smaller.
	if cols {
		s.frozenCols = shiftedFrozenCount(e, s.frozenCols)
	} else {
		s.frozenRows = shiftedFrozenCount(e, s.frozenRows)
	}

	f.rebuildFormulaState()
	u.markAfter = f.changes
	return u, nil
}

// Undo restores the state from before the edit. See the type comment for
// the condition: no other change after the edit.
func (u *StructuralUndo) Undo() error {
	if u == nil || u.sheet == nil {
		return errors.New("structural undo: no edit")
	}
	if u.done {
		return errors.New("structural undo: the edit is already undone")
	}
	s := u.sheet
	f := s.file
	if f.Sheet(s.name) != s {
		return fmt.Errorf("structural undo: sheet %q is not in the workbook", s.name)
	}
	// The undo data is for the state immediately after the edit. On a
	// workbook that changed after it, the inverse move would put rows, old
	// formula text and the old lists of merges, names and tables in wrong
	// positions, with no sign of an error. Refuse it.
	if f.changes != u.markAfter {
		return errors.New("structural undo: the workbook changed after the edit; undo the later changes first")
	}
	u.done = true

	// Cells that the edit made go first: the inverse move must not see them.
	for _, key := range u.addedCells {
		if r, ok := s.rows[key.row]; ok {
			delete(r.cells, key.col)
			if len(r.cells) == 0 && r.height == 0 && !r.hidden {
				delete(s.rows, key.row)
			}
		}
	}

	// The inverse move puts the cells that stayed at their old positions.
	// The lines that an insert made are empty, so the inverse delete removes
	// nothing that must stay.
	inverse := u.edit
	inverse.Delete = !u.edit.Delete
	s.shiftCells(inverse, nil)

	for n, r := range u.removedRows {
		r.num = n
		s.rows[n] = r
	}
	for _, rc := range u.removedCells {
		rc.cell.col = rc.col
		s.ensureRow(rc.row).cells[rc.col] = rc.cell
	}
	for _, fc := range u.formulas {
		fc.cell.formula = fc.old
	}
	s.merges = u.merges
	if u.colWidths != nil {
		s.colWidths = u.colWidths
	}
	s.frozenRows, s.frozenCols = u.frozenRows, u.frozenCols
	f.definedNames = u.definedNames
	f.tables = u.tables
	f.tableDefs = u.tableDefs
	s.spill.anchors = make(map[cellKey]*spillAnchorState, len(u.spillAnchors))
	for key, state := range u.spillAnchors {
		st := state
		s.spill.anchors[key] = &st
	}
	s.invalidateSpillOverlay()

	f.rebuildFormulaState()
	// The workbook is in the state from before the edit again.
	f.changes = u.markBefore
	return nil
}

// shiftedFrozenCount returns the number of frozen rows or columns after the
// edit. The frozen lines are the span 1..n; its new last line is the new
// count (an insert at line 1 moves the span, and the new lines are frozen
// too, as in Excel).
func shiftedFrozenCount(e formula.StructuralEdit, n int) int {
	if n <= 0 {
		return 0
	}
	_, hi, gone := e.ShiftSpan(1, n)
	if gone {
		return 0
	}
	// A freeze must leave one line that scrolls.
	return min(hi, axisLast(e.Cols)-1)
}

// checkInsertFits fails when an insert would push a non-empty cell off the
// sheet. Empty cells (style only) and empty rows at the end are dropped.
func (s *Sheet) checkInsertFits(e formula.StructuralEdit) error {
	limit := axisLast(e.Cols) - e.Count // the last line that stays on the sheet
	for n, r := range s.rows {
		if !e.Cols && n <= limit {
			continue
		}
		for col, c := range r.cells {
			if e.Cols && col <= limit {
				continue
			}
			if c.formula != "" || c.value.Type != TypeEmpty {
				ref, _ := CoordinatesToCellName(col, n)
				return fmt.Errorf("cannot insert %d %ss: cell %s is not empty and would move off the sheet", e.Count, axisName(e.Cols), ref)
			}
		}
	}
	return nil
}

// shiftCells moves the rows (row edit) or the cells of each row (column
// edit) to their new positions. Lines that the edit removes are stored in u
// when u is not nil.
func (s *Sheet) shiftCells(e formula.StructuralEdit, u *StructuralUndo) {
	if !e.Cols {
		rows := make(map[int]*Row, len(s.rows))
		for n, r := range s.rows {
			n2, _, gone := e.ShiftSpan(n, n)
			if gone {
				if u != nil {
					if u.removedRows == nil {
						u.removedRows = make(map[int]*Row)
					}
					u.removedRows[n] = r
				}
				continue
			}
			r.num = n2
			rows[n2] = r
		}
		s.rows = rows
		return
	}

	for n, r := range s.rows {
		moved := false
		for col := range r.cells {
			if col >= e.At {
				moved = true
				break
			}
		}
		if !moved {
			continue
		}
		cells := make(map[int]*Cell, len(r.cells))
		for col, c := range r.cells {
			col2, _, gone := e.ShiftSpan(col, col)
			if gone {
				if u != nil {
					u.removedCells = append(u.removedCells, removedCell{row: n, col: col, cell: c})
				}
				continue
			}
			c.col = col2
			cells[col2] = c
		}
		r.cells = cells
		// A row that lost all its cells and has no height is not a row of
		// the sheet any more.
		if len(cells) == 0 && r.height == 0 && !r.hidden {
			delete(s.rows, n)
		}
	}
}

// shiftedMerges returns the merged ranges after the edit. A range that
// shrinks to one cell or to nothing is removed. An insert that would push a
// merged range off the sheet is an error.
func (s *Sheet) shiftedMerges(e formula.StructuralEdit) ([]MergeRange, error) {
	if len(s.merges) == 0 {
		return s.merges, nil
	}
	out := make([]MergeRange, 0, len(s.merges))
	for _, mr := range s.merges {
		col1, row1, col2, row2, err := RangeToCoordinates(mr.Start + ":" + mr.End)
		if err != nil {
			out = append(out, mr)
			continue
		}
		lo, hi := row1, row2
		if e.Cols {
			lo, hi = col1, col2
		}
		lo2, hi2, gone := e.ShiftSpan(lo, hi)
		if gone {
			if !e.Delete {
				return nil, fmt.Errorf("cannot insert %d %ss: the merged range %s:%s would move off the sheet", e.Count, axisName(e.Cols), mr.Start, mr.End)
			}
			continue
		}
		if insertCutsSpan(e, lo, hi, lo2, hi2) {
			return nil, fmt.Errorf("cannot insert %d %ss: the merged range %s:%s would move off the sheet", e.Count, axisName(e.Cols), mr.Start, mr.End)
		}
		if e.Cols {
			col1, col2 = lo2, hi2
		} else {
			row1, row2 = lo2, hi2
		}
		if col1 == col2 && row1 == row2 {
			continue
		}
		start, err1 := CoordinatesToCellName(col1, row1)
		end, err2 := CoordinatesToCellName(col2, row2)
		if err1 != nil || err2 != nil {
			continue
		}
		out = append(out, MergeRange{Start: start, End: end})
	}
	return out, nil
}

// insertCutsSpan reports whether an insert cut the span lo..hi at the end of
// the sheet: its new form lo2..hi2 is shorter than the moved or grown span.
// A reference can be cut; a merged range and a table cannot.
func insertCutsSpan(e formula.StructuralEdit, lo, hi, lo2, hi2 int) bool {
	if e.Delete || (lo <= 1 && hi >= axisLast(e.Cols)) {
		return false
	}
	want := hi - lo
	if lo < e.At && hi >= e.At {
		want += e.Count
	}
	return hi2-lo2 < want
}

type tableHeaderCell struct {
	row, col int
	name     string
}

// shiftedTables returns the table lists of the workbook after the edit, and
// the header cells of the table columns that a column insert adds.
func (s *Sheet) shiftedTables(e formula.StructuralEdit) ([]Table, []formula.TableInfo, []tableHeaderCell, error) {
	f := s.file
	onSheet := false
	for _, td := range f.tableDefs {
		if td.SheetName == s.name {
			onSheet = true
			break
		}
	}
	if !onSheet {
		return f.tableDefs, f.tables, nil, nil
	}

	// The info list is parallel to the definitions by name.
	infoByName := make(map[string]formula.TableInfo, len(f.tables))
	for _, info := range f.tables {
		infoByName[strings.ToLower(info.Name)] = info
	}

	var (
		defs    []Table
		infos   []formula.TableInfo
		headers []tableHeaderCell
	)
	for _, td := range f.tableDefs {
		info, hasInfo := infoByName[strings.ToLower(td.DisplayName)]
		if td.SheetName != s.name {
			defs = append(defs, td)
			if hasInfo {
				infos = append(infos, info)
			}
			continue
		}
		col1, row1, col2, row2, err := RangeToCoordinates(td.Ref)
		if err != nil {
			defs = append(defs, td)
			if hasInfo {
				infos = append(infos, info)
			}
			continue
		}
		td = td.clone()
		lo, hi := row1, row2
		if e.Cols {
			lo, hi = col1, col2
		}
		lo2, hi2, gone := e.ShiftSpan(lo, hi)
		if gone {
			if !e.Delete {
				return nil, nil, nil, fmt.Errorf("cannot insert %d %ss: table %q would move off the sheet", e.Count, axisName(e.Cols), td.DisplayName)
			}
			continue // the full table is in the deleted band
		}
		if insertCutsSpan(e, lo, hi, lo2, hi2) {
			return nil, nil, nil, fmt.Errorf("cannot insert %d %ss: table %q would move off the sheet", e.Count, axisName(e.Cols), td.DisplayName)
		}

		switch {
		case e.Cols && e.Delete:
			// Columns in the band leave the column list.
			kept := td.Columns[:0:0]
			for i, name := range td.Columns {
				if c := col1 + i; c < e.At || c > e.At+e.Count-1 {
					kept = append(kept, name)
				}
			}
			td.Columns = kept
		case e.Cols && lo < e.At && hi >= e.At:
			// New columns in the table get unique names.
			idx := e.At - col1
			names := make([]string, 0, len(td.Columns)+e.Count)
			names = append(names, td.Columns[:min(idx, len(td.Columns))]...)
			for i := 0; i < e.Count; i++ {
				name := nextTableColumnName(append(names, td.Columns...))
				names = append(names, name)
				if td.HeaderRowCount > 0 {
					headers = append(headers, tableHeaderCell{row: row1, col: e.At + i, name: name})
				}
			}
			if idx < len(td.Columns) {
				names = append(names, td.Columns[idx:]...)
			}
			td.Columns = names
		case !e.Cols && e.Delete:
			first, end := e.At, e.At+e.Count-1
			if td.HeaderRowCount > 0 && first <= row1+td.HeaderRowCount-1 && end >= row1 {
				return nil, nil, nil, fmt.Errorf("cannot delete the header row of table %q", td.DisplayName)
			}
			// A deleted totals row leaves the table.
			if td.TotalsRowCount > 0 && end >= row2-td.TotalsRowCount+1 && first <= row2 {
				td.TotalsRowCount = 0
			}
			if hi2-lo2+1 < td.HeaderRowCount+td.TotalsRowCount+1 {
				return nil, nil, nil, fmt.Errorf("cannot delete all data rows of table %q", td.DisplayName)
			}
		}

		if e.Cols {
			col1, col2 = lo2, hi2
		} else {
			row1, row2 = lo2, hi2
		}
		start, err1 := CoordinatesToCellName(col1, row1)
		end, err2 := CoordinatesToCellName(col2, row2)
		if err1 != nil || err2 != nil {
			return nil, nil, nil, fmt.Errorf("table %q has no valid range after the edit", td.DisplayName)
		}
		td.Ref = start + ":" + end
		defs = append(defs, td)
		if hasInfo {
			info.Columns = append([]string(nil), td.Columns...)
			info.FirstCol, info.FirstRow, info.LastCol, info.LastRow = col1, row1, col2, row2
			info.TotalRows = td.TotalsRowCount
			infos = append(infos, info)
		}
	}
	return defs, infos, headers, nil
}

// lostTableColumns names the columns that a column delete removed from one
// table, and the cells of the table after the edit.
type lostTableColumns struct {
	table   string
	columns []string
	rect    cellRect
}

type cellRect struct{ col1, row1, col2, row2 int }

func (r cellRect) contains(col, row int) bool {
	return col >= r.col1 && col <= r.col2 && row >= r.row1 && row <= r.row2
}

// removedTableColumns returns, for each table of the sheet that a column
// delete made narrower, the names of its deleted columns. newDefs is the
// table list after the edit. A table that the edit removed as a whole is not
// in the result: its references have no table to refer to any more.
func (s *Sheet) removedTableColumns(e formula.StructuralEdit, newDefs []Table) []lostTableColumns {
	if !e.Cols || !e.Delete {
		return nil
	}
	var out []lostTableColumns
	for _, old := range s.file.tableDefs {
		if old.SheetName != s.name {
			continue
		}
		col1, _, _, _, err := RangeToCoordinates(old.Ref)
		if err != nil {
			continue
		}
		var lost []string
		for i, name := range old.Columns {
			if c := col1 + i; c >= e.At && c <= e.At+e.Count-1 {
				lost = append(lost, name)
			}
		}
		if len(lost) == 0 {
			continue
		}
		for _, td := range newDefs {
			if td.SheetName != s.name || !strings.EqualFold(td.DisplayName, old.DisplayName) {
				continue
			}
			c1, r1, c2, r2, err := RangeToCoordinates(td.Ref)
			if err != nil {
				break
			}
			out = append(out, lostTableColumns{
				table:   old.DisplayName,
				columns: lost,
				rect:    cellRect{col1: c1, row1: r1, col2: c2, row2: r2},
			})
			break
		}
	}
	return out
}

// nextTableColumnName returns the first "ColumnN" that is not in names.
func nextTableColumnName(names []string) string {
	used := make(map[string]bool, len(names))
	for _, n := range names {
		used[strings.ToLower(n)] = true
	}
	for i := 1; ; i++ {
		name := "Column" + strconv.Itoa(i)
		if !used[strings.ToLower(name)] {
			return name
		}
	}
}

func (s *Sheet) copySpillAnchors() map[cellKey]spillAnchorState {
	out := make(map[cellKey]spillAnchorState, len(s.spill.anchors))
	for key, state := range s.spill.anchors {
		if state != nil {
			out[key] = *state
		}
	}
	return out
}

// shiftSpillAnchors moves the spill state of each anchor with its cell, and
// shifts the imported formula ref (the fallback of the writer). The
// published bounds are derived again at the next evaluation.
func (s *Sheet) shiftSpillAnchors(e formula.StructuralEdit) {
	if len(s.spill.anchors) == 0 {
		return
	}
	anchors := make(map[cellKey]*spillAnchorState, len(s.spill.anchors))
	for key, state := range s.spill.anchors {
		if state == nil {
			continue
		}
		pos := key.row
		if e.Cols {
			pos = key.col
		}
		pos2, _, gone := e.ShiftSpan(pos, pos)
		if gone {
			continue
		}
		if e.Cols {
			key.col = pos2
		} else {
			key.row = pos2
		}
		st := *state
		st.gen = 0 // not current: the bounds are for the old position
		st.formulaRef = shiftRangeRef(st.formulaRef, e)
		anchors[key] = &st
	}
	s.spill.anchors = anchors
	s.invalidateSpillOverlay()
}

// shiftRangeRef applies the edit to an A1 range text such as "A1:A3". It
// returns "" when the range is gone or is not a range.
func shiftRangeRef(ref string, e formula.StructuralEdit) string {
	if ref == "" {
		return ""
	}
	col1, row1, col2, row2, err := RangeToCoordinates(ref)
	if err != nil {
		return ""
	}
	lo, hi := row1, row2
	if e.Cols {
		lo, hi = col1, col2
	}
	lo2, hi2, gone := e.ShiftSpan(lo, hi)
	if gone {
		return ""
	}
	if e.Cols {
		col1, col2 = lo2, hi2
	} else {
		row1, row2 = lo2, hi2
	}
	start, err1 := CoordinatesToCellName(col1, row1)
	end, err2 := CoordinatesToCellName(col2, row2)
	if err1 != nil || err2 != nil {
		return ""
	}
	return start + ":" + end
}
