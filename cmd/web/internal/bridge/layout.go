package bridge

import (
	"fmt"
	"math"

	werkbook "github.com/jpoz/werkbook"
)

// Size limits of the contract, in CSS pixels.
const (
	minColWidthPx  = 12
	maxColWidthPx  = 2000
	minRowHeightPx = 8
	maxRowHeightPx = 1000
	// maxLayoutRows is the largest number of rows that one SetRowHeight
	// call gives a height.
	maxLayoutRows = maxOpCells
)

const (
	labelColWidth  = "Column width"
	labelRowHeight = "Row height"
)

// LayoutResult is the result of SetColWidth and SetRowHeight.
type LayoutResult struct {
	History HistoryState `json:"history"`
	// Range is the full columns or full rows that the call covers.
	Range Range `json:"range"`
	// Clamped is true when SetRowHeight limited a range of more than
	// 200,000 rows to the used range of the sheet.
	Clamped bool `json:"clamped"`
}

// Pixel conversions. A size that the contract sets comes back from SheetMeta
// as the same pixel number.

func colWidthToPx(w float64) int { return int(math.Round(w*colWidthCharPx + colWidthPadPx)) }
func pxToColWidth(px int) float64 {
	return (float64(px) - colWidthPadPx) / colWidthCharPx
}
func rowHeightToPx(h float64) int  { return int(math.Round(h * pointsToPx)) }
func pxToRowHeight(px int) float64 { return float64(px) / pointsToPx }

// sizeChange holds the old and the new size of columns or rows. A size of 0
// is "no custom size".
type sizeChange struct {
	sheet         string
	cols          bool
	before, after map[int]float64
}

func (c *sizeChange) cost() int { return len(c.before) }

// A size cannot change a formula result.
func (c *sizeChange) recalc() bool { return false }

func (c *sizeChange) apply(s *Session, redo bool) error {
	sh := s.file.Sheet(c.sheet)
	if sh == nil {
		return fmt.Errorf("sheet %q not found", c.sheet)
	}
	sizes := c.before
	if redo {
		sizes = c.after
	}
	for idx, size := range sizes {
		var err error
		if c.cols {
			err = sh.SetColumnWidth(werkbook.ColumnNumberToName(idx), size)
		} else {
			err = sh.SetRowHeight(idx, size)
		}
		if err != nil {
			return err
		}
	}
	return nil
}

// SetColWidth gives the columns col1..col2 the width px (CSS pixels) as one
// undo step. px <= 0 removes the custom width; else px is limited to
// 12..2000.
func (s *Session) SetColWidth(sheet string, col1, col2, px int) (LayoutResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return LayoutResult{}, err
	}
	if col1 > col2 {
		col1, col2 = col2, col1
	}
	if col1 < 1 || col2 > werkbook.MaxColumns {
		return LayoutResult{}, fmt.Errorf("column range %d..%d is out of range [1, %d]", col1, col2, werkbook.MaxColumns)
	}
	width := 0.0
	if px > 0 {
		width = pxToColWidth(max(minColWidthPx, min(px, maxColWidthPx)))
	}

	widths := sh.ColumnWidths()
	change := &sizeChange{sheet: sheet, cols: true, before: map[int]float64{}, after: map[int]float64{}}
	for col := col1; col <= col2; col++ {
		if widths[col] != width {
			change.before[col] = widths[col]
			change.after[col] = width
		}
	}
	sel := Range{Row1: 1, Col1: col1, Row2: werkbook.MaxRows, Col2: col2}
	return s.applySizeChange(labelColWidth, change, sel, false)
}

// SetRowHeight gives the rows row1..row2 the height px (CSS pixels) as one
// undo step. px <= 0 removes the custom height; else px is limited to
// 8..1000.
//
// A removal touches only the rows that have a height, so it has no limit. A
// height for more than 200,000 rows (a selection of all rows) is limited to
// the used range of the sheet; the result then has Clamped true and the
// range that got the height.
func (s *Session) SetRowHeight(sheet string, row1, row2, px int) (LayoutResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return LayoutResult{}, err
	}
	if row1 > row2 {
		row1, row2 = row2, row1
	}
	if row1 < 1 || row2 > werkbook.MaxRows {
		return LayoutResult{}, fmt.Errorf("row range %d..%d is out of range [1, %d]", row1, row2, werkbook.MaxRows)
	}
	change := &sizeChange{sheet: sheet, before: map[int]float64{}, after: map[int]float64{}}

	if px <= 0 {
		for row := range sh.Rows() {
			if n := row.Num(); n >= row1 && n <= row2 && row.Height() != 0 {
				change.before[n] = row.Height()
				change.after[n] = 0
			}
		}
		sel := Range{Row1: row1, Col1: 1, Row2: row2, Col2: werkbook.MaxColumns}
		return s.applySizeChange(labelRowHeight, change, sel, false)
	}

	clamped := false
	if row2-row1+1 > maxLayoutRows {
		s.ensureCalculated()
		maxRow, _ := s.usedRange(sh)
		row2 = min(row2, max(row1, maxRow))
		clamped = true
		if row2-row1+1 > maxLayoutRows {
			return LayoutResult{}, fmt.Errorf("the call changes more than %d rows", maxLayoutRows)
		}
	}
	height := pxToRowHeight(max(minRowHeightPx, min(px, maxRowHeightPx)))
	for row := row1; row <= row2; row++ {
		old, err := sh.GetRowHeight(row)
		if err != nil {
			return LayoutResult{}, err
		}
		if old != height {
			change.before[row] = old
			change.after[row] = height
		}
	}
	sel := Range{Row1: row1, Col1: 1, Row2: row2, Col2: werkbook.MaxColumns}
	return s.applySizeChange(labelRowHeight, change, sel, clamped)
}

func (s *Session) applySizeChange(label string, change *sizeChange, sel Range, clamped bool) (LayoutResult, error) {
	res := LayoutResult{Range: sel, Clamped: clamped}
	if len(change.after) == 0 {
		// No size changes: no step.
		res.History = s.History()
		return res, nil
	}
	err := s.mutate(label, func(rec *recorder) error {
		rec.setSheet(change.sheet)
		rec.setRange(sel)
		// Record first: a rollback then restores the sizes that a failed
		// apply changed.
		rec.other = append(rec.other, change)
		return change.apply(s, true)
	})
	if err != nil {
		return LayoutResult{}, err
	}
	res.History = s.History()
	return res, nil
}
