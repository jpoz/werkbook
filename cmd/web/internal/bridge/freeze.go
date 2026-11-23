package bridge

import (
	"fmt"

	werkbook "github.com/jpoz/werkbook"
)

const labelFreeze = "Freeze panes"

// Freeze is the frozen pane of a sheet: Rows rows at the top and Cols
// columns at the left stay in view. 0 means "not frozen" on that axis.
type Freeze struct {
	Rows int `json:"rows"`
	Cols int `json:"cols"`
}

// FreezeResult is the result of SetFreeze.
type FreezeResult struct {
	History HistoryState `json:"history"`
}

// freezeChange holds the old and the new frozen pane of a sheet.
type freezeChange struct {
	sheet         string
	before, after Freeze
}

func (c *freezeChange) cost() int { return 1 }

// A frozen pane is a view setting: it cannot change a formula result.
func (c *freezeChange) recalc() bool { return false }

func (c *freezeChange) apply(s *Session, redo bool) error {
	sh := s.file.Sheet(c.sheet)
	if sh == nil {
		return fmt.Errorf("sheet %q not found", c.sheet)
	}
	fz := c.before
	if redo {
		fz = c.after
	}
	return sh.SetFreezePanes(fz.Rows, fz.Cols)
}

// SetFreeze freezes rows rows at the top and cols columns at the left of a
// sheet as one undo step. SetFreeze(sheet, 0, 0) removes the freeze.
func (s *Session) SetFreeze(sheet string, rows, cols int) (FreezeResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return FreezeResult{}, err
	}
	if rows < 0 || rows >= werkbook.MaxRows {
		return FreezeResult{}, fmt.Errorf("rows %d is out of range [0, %d]", rows, werkbook.MaxRows-1)
	}
	if cols < 0 || cols >= werkbook.MaxColumns {
		return FreezeResult{}, fmt.Errorf("cols %d is out of range [0, %d]", cols, werkbook.MaxColumns-1)
	}
	oldRows, oldCols := sh.FreezePanes()
	change := &freezeChange{
		sheet:  sheet,
		before: Freeze{Rows: oldRows, Cols: oldCols},
		after:  Freeze{Rows: rows, Cols: cols},
	}
	if change.before == change.after {
		return FreezeResult{History: s.History()}, nil
	}
	err = s.mutate(labelFreeze, func(rec *recorder) error {
		rec.setSheet(sheet)
		rec.other = append(rec.other, change)
		return change.apply(s, true)
	})
	if err != nil {
		return FreezeResult{}, err
	}
	return FreezeResult{History: s.History()}, nil
}
