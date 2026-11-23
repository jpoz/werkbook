package bridge

import (
	"errors"
	"fmt"

	werkbook "github.com/jpoz/werkbook"
)

// History limits. See cmd/web/PLAN.md section 2.1.
const (
	// maxOpCells is the largest number of cells that one call can change.
	maxOpCells = 200_000
	// maxHistorySteps is the largest number of undo steps.
	maxHistorySteps = 100
	// maxHistoryCost is the largest total cost of all steps. One unit is one
	// cell delta.
	maxHistoryCost = 500_000
)

// Step labels. The UI shows them as "Undo <label>".
const (
	labelEditCell    = "Edit cell"
	labelEditCells   = "Edit cells"
	labelClear       = "Clear"
	labelFill        = "Fill"
	labelCut         = "Cut"
	labelPaste       = "Paste"
	labelAddSheet    = "Add sheet"
	labelRenameSheet = "Rename sheet"
)

// Range is an inclusive cell rectangle. All coordinates are 1-based.
type Range struct {
	Row1 int `json:"row1"`
	Col1 int `json:"col1"`
	Row2 int `json:"row2"`
	Col2 int `json:"col2"`
}

func (r Range) rows() int  { return r.Row2 - r.Row1 + 1 }
func (r Range) cols() int  { return r.Col2 - r.Col1 + 1 }
func (r Range) cells() int { return r.rows() * r.cols() }

func (r Range) contains(row, col int) bool {
	return row >= r.Row1 && row <= r.Row2 && col >= r.Col1 && col <= r.Col2
}

func (r Range) containsRange(o Range) bool {
	return o.Row1 >= r.Row1 && o.Row2 <= r.Row2 && o.Col1 >= r.Col1 && o.Col2 <= r.Col2
}

func (r Range) intersects(o Range) bool {
	return r.Row1 <= o.Row2 && o.Row1 <= r.Row2 && r.Col1 <= o.Col2 && o.Col1 <= r.Col2
}

// normRange validates the corners and returns the range with ordered corners.
func normRange(r Range) (Range, error) {
	if _, err := cellRef(r.Row1, r.Col1); err != nil {
		return Range{}, err
	}
	if _, err := cellRef(r.Row2, r.Col2); err != nil {
		return Range{}, err
	}
	if r.Row1 > r.Row2 {
		r.Row1, r.Row2 = r.Row2, r.Row1
	}
	if r.Col1 > r.Col2 {
		r.Col1, r.Col2 = r.Col2, r.Col1
	}
	return r, nil
}

// HistoryState is the state of the undo and redo stacks.
type HistoryState struct {
	CanUndo   bool   `json:"canUndo"`
	CanRedo   bool   `json:"canRedo"`
	UndoLabel string `json:"undoLabel"`
	RedoLabel string `json:"redoLabel"`
	Depth     int    `json:"depth"`
	// Cleared is true when the call cleared the history because its step was
	// above the cost limit.
	Cleared bool `json:"cleared,omitempty"`
}

// HistoryResult is the result of Undo and Redo.
type HistoryResult struct {
	History HistoryState `json:"history"`
	// Sheet is the sheet that the step changed, with its name after the undo
	// or redo. For an undo of "Add sheet" it is the last sheet that stays.
	Sheet string `json:"sheet"`
	// Select is the range of the step; nil for a step with no range.
	Select *Range      `json:"select"`
	Sheets []SheetInfo `json:"sheets"`
	Label  string      `json:"label"`
}

// change is one reversible part of a step.
type change interface {
	// apply makes the before-state (redo false) or the after-state (redo
	// true) current.
	apply(s *Session, redo bool) error
	cost() int
	// recalc reports whether the change can change a formula result. A
	// change of a column width or a row height cannot; it must not start a
	// recalculation (a volatile formula such as RAND would get a new value).
	recalc() bool
}

func anyRecalc(changes []change) bool {
	for _, c := range changes {
		if c.recalc() {
			return true
		}
	}
	return false
}

type cellPos struct{ col, row int }

type cellDelta struct {
	pos           cellPos
	before, after werkbook.CellSnapshot
}

// cellsChange holds the cells of one sheet that a step changed.
type cellsChange struct {
	sheet string
	cells []cellDelta
}

func (c *cellsChange) cost() int    { return len(c.cells) }
func (c *cellsChange) recalc() bool { return true }

func (c *cellsChange) apply(s *Session, redo bool) error {
	sh := s.file.Sheet(c.sheet)
	if sh == nil {
		return fmt.Errorf("sheet %q not found", c.sheet)
	}
	s.file.BeginBatchWrite()
	defer s.file.EndBatchWrite()
	var first error
	for _, d := range c.cells {
		snap := d.before
		if redo {
			snap = d.after
		}
		if err := sh.RestoreCell(d.pos.col, d.pos.row, snap); err != nil && first == nil {
			first = err
		}
	}
	return first
}

type sheetAddChange struct{ name string }

func (c *sheetAddChange) cost() int    { return 1 }
func (c *sheetAddChange) recalc() bool { return true }

func (c *sheetAddChange) apply(s *Session, redo bool) error {
	if redo {
		_, err := s.file.NewSheet(c.name)
		return err
	}
	s.clipboardSheetGone(c.name)
	return s.file.DeleteSheet(c.name)
}

type sheetRenameChange struct{ from, to string }

func (c *sheetRenameChange) cost() int    { return 1 }
func (c *sheetRenameChange) recalc() bool { return true }

func (c *sheetRenameChange) apply(s *Session, redo bool) error {
	from, to := c.to, c.from
	if redo {
		from, to = c.from, c.to
	}
	if err := s.file.SetSheetName(from, to); err != nil {
		return err
	}
	s.clipboardSheetRenamed(from, to)
	return nil
}

// step is one undo step: the changes of one mutating call.
type step struct {
	// mark is the change mark of the workbook before the step ran (the
	// last time: a redo runs the step again).
	mark  uint64
	label string
	// undoSheet and redoSheet are the sheet to show after an undo and after
	// a redo. They differ for a rename. An empty name means "the last sheet".
	undoSheet, redoSheet string
	sel                  *Range
	changes              []change
	cost                 int
}

// recorder collects the changes of one mutating call.
type recorder struct {
	s *Session

	// Cell changes. One call changes the cells of one sheet.
	sheet  *werkbook.Sheet
	order  []cellPos
	before map[cellPos]werkbook.CellSnapshot

	// Changes that are not cell changes, in the sequence of their execution.
	other []change

	sel                  *Range
	undoSheet, redoSheet string
}

// setSheet gives the sheet that the step shows after an undo and a redo.
func (r *recorder) setSheet(name string) {
	r.undoSheet, r.redoSheet = name, name
}

// setRange gives the range that the UI selects after an undo or redo.
func (r *recorder) setRange(rg Range) { r.sel = &rg }

// write makes the stored state of a cell equal to snap. It takes the
// before-snapshot at the first touch of the cell in this call. A snapshot of
// an empty cell with the default style removes the cell, so that the used
// range of the sheet does not grow from cleared cells.
func (r *recorder) write(sh *werkbook.Sheet, col, row int, snap werkbook.CellSnapshot) error {
	if r.sheet == nil {
		r.sheet = sh
		r.before = make(map[cellPos]werkbook.CellSnapshot)
	} else if r.sheet != sh {
		return errors.New("internal error: one step changed cells on two sheets")
	}
	if snap.IsEmpty() && snap.Style() == nil {
		snap = werkbook.CellSnapshot{}
	}
	pos := cellPos{col: col, row: row}
	if _, ok := r.before[pos]; !ok {
		if len(r.order) >= maxOpCells {
			return fmt.Errorf("the operation changes more than %d cells", maxOpCells)
		}
		r.before[pos] = sh.SnapshotCell(col, row)
		r.order = append(r.order, pos)
	}
	return sh.RestoreCell(col, row, snap)
}

// rollback restores the state from before the call.
func (r *recorder) rollback() {
	for i := len(r.other) - 1; i >= 0; i-- {
		_ = r.other[i].apply(r.s, false)
	}
	for i := len(r.order) - 1; i >= 0; i-- {
		pos := r.order[i]
		_ = r.sheet.RestoreCell(pos.col, pos.row, r.before[pos])
	}
}

// mutate runs one mutating call as one undo step. If fn returns an error, all
// changes of the call are reversed and the history is unchanged, so each
// contract function is atomic. A call that changes nothing makes no step.
//
// A panic in fn (for example in the formula compiler) is handled as an
// error: the changes are reversed, the batch state of the file is closed, and
// the caller gets an error. The WASM wrapper would recover the panic too, but
// then the workbook would keep the half-done changes.
func (s *Session) mutate(label string, fn func(rec *recorder) error) error {
	rec := &recorder{s: s}
	s.lastCleared = false
	// The library gives the workbook a new change mark for each write. An
	// undo of this step gives the old mark back (applyStep), so that the
	// undo data of an older structural step is valid again.
	mark := s.file.ChangeMark()

	s.file.BeginBatchWrite()
	err := protect(func() error { return fn(rec) })
	if err != nil {
		if rbErr := protect(func() error { rec.rollback(); return nil }); rbErr != nil {
			// The state of the workbook is not known. Steps that were
			// recorded before can be wrong for it.
			s.clearHistory()
			err = fmt.Errorf("%w (the changes could not be reversed: %v; the history was cleared)", err, rbErr)
		}
	}
	s.file.EndBatchWrite()
	if len(rec.order) > 0 || anyRecalc(rec.other) {
		s.stale = true
	}
	if err != nil {
		// All changes are reversed.
		s.file.RestoreChangeMark(mark)
		return err
	}

	st := &step{
		mark:      mark,
		label:     label,
		undoSheet: rec.undoSheet,
		redoSheet: rec.redoSheet,
		sel:       rec.sel,
		changes:   rec.other,
	}
	if len(rec.order) > 0 {
		cc := &cellsChange{sheet: rec.sheet.Name()}
		for _, pos := range rec.order {
			before := rec.before[pos]
			after := rec.sheet.SnapshotCell(pos.col, pos.row)
			if before.Equal(after) {
				continue
			}
			cc.cells = append(cc.cells, cellDelta{pos: pos, before: before, after: after})
		}
		if len(cc.cells) > 0 {
			st.changes = append(st.changes, cc)
		}
	}
	if len(st.changes) == 0 {
		// The call wrote cells but each cell has its old state: the
		// workbook did not change.
		s.file.RestoreChangeMark(mark)
		return nil
	}
	for _, c := range st.changes {
		st.cost += c.cost()
	}
	s.pushStep(st)
	return nil
}

// protect runs fn and returns a panic of fn as an error.
func protect(fn func() error) (err error) {
	defer func() {
		if r := recover(); r != nil {
			err = fmt.Errorf("internal error: %v", r)
		}
	}()
	return fn()
}

func (s *Session) pushStep(st *step) {
	// One step above the total limit: the operation ran, but its undo data
	// is too large to keep. Older steps cannot stay either, because they are
	// valid only for the state before this step.
	if st.cost > maxHistoryCost {
		s.clearHistory()
		s.lastCleared = true
		return
	}
	s.redoStack = nil
	s.undoStack = append(s.undoStack, st)
	total := 0
	for _, x := range s.undoStack {
		total += x.cost
	}
	// Drop the oldest steps until the stack is in its limits. The new step
	// always stays: a cell-level call cannot be above the cost limit.
	drop := 0
	for len(s.undoStack)-drop > 1 &&
		(len(s.undoStack)-drop > maxHistorySteps || total > maxHistoryCost) {
		total -= s.undoStack[drop].cost
		drop++
	}
	if drop > 0 {
		s.undoStack = append([]*step(nil), s.undoStack[drop:]...)
	}
}

func (s *Session) clearHistory() {
	s.undoStack = nil
	s.redoStack = nil
}

// History returns the state of the undo and redo stacks.
func (s *Session) History() HistoryState {
	h := HistoryState{
		CanUndo: len(s.undoStack) > 0,
		CanRedo: len(s.redoStack) > 0,
		Depth:   len(s.undoStack),
	}
	if h.CanUndo {
		h.UndoLabel = s.undoStack[len(s.undoStack)-1].label
	}
	if h.CanRedo {
		h.RedoLabel = s.redoStack[len(s.redoStack)-1].label
	}
	return h
}

// historyOfLastMutation returns the history state for the result of a
// mutating call. Cleared is true when that call cleared the history because
// its step was above the cost limit.
func (s *Session) historyOfLastMutation() HistoryState {
	h := s.History()
	h.Cleared = s.lastCleared
	return h
}

// Undo reverses the last step.
func (s *Session) Undo() (HistoryResult, error) {
	if len(s.undoStack) == 0 {
		return HistoryResult{}, errors.New("nothing to undo")
	}
	st := s.undoStack[len(s.undoStack)-1]
	s.undoStack = s.undoStack[:len(s.undoStack)-1]
	s.redoStack = append(s.redoStack, st)
	return s.applyStep(st, false)
}

// Redo applies the last step that was undone.
func (s *Session) Redo() (HistoryResult, error) {
	if len(s.redoStack) == 0 {
		return HistoryResult{}, errors.New("nothing to redo")
	}
	st := s.redoStack[len(s.redoStack)-1]
	s.redoStack = s.redoStack[:len(s.redoStack)-1]
	s.undoStack = append(s.undoStack, st)
	return s.applyStep(st, true)
}

func (s *Session) applyStep(st *step, redo bool) (HistoryResult, error) {
	if anyRecalc(st.changes) {
		s.stale = true
	}
	var first error
	if redo {
		st.mark = s.file.ChangeMark()
		for _, c := range st.changes {
			if err := c.apply(s, true); err != nil && first == nil {
				first = err
			}
		}
	} else {
		for i := len(st.changes) - 1; i >= 0; i-- {
			if err := st.changes[i].apply(s, false); err != nil && first == nil {
				first = err
			}
		}
	}
	if first != nil {
		// The stacks hold steps in strict sequence. A step that does not
		// apply makes all other steps invalid.
		s.clearHistory()
		return HistoryResult{}, fmt.Errorf("%s: %w (the history was cleared)", st.label, first)
	}
	if !redo {
		// The workbook is in the state from before the step again.
		s.file.RestoreChangeMark(st.mark)
	}

	sheet := st.undoSheet
	if redo {
		sheet = st.redoSheet
	}
	if s.file.Sheet(sheet) == nil {
		names := s.file.SheetNames()
		sheet = names[len(names)-1]
	}
	return HistoryResult{
		History: s.History(),
		Sheet:   sheet,
		Select:  st.sel,
		Sheets:  s.Sheets(),
		Label:   st.label,
	}, nil
}
