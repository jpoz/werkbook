package bridge

import (
	"errors"
	"fmt"
	"math"
	"strconv"

	werkbook "github.com/jpoz/werkbook"
	"github.com/jpoz/werkbook/formula"
)

// FillResult is the result of Fill.
type FillResult struct {
	History HistoryState `json:"history"`
	// Range is the destination range.
	Range Range `json:"range"`
	// Changed is the number of destination cells that were written.
	Changed int `json:"changed"`
	// Skipped is the number of read-only spill cells that were not written.
	Skipped int `json:"skipped"`
}

// lane is the source cells of one column (vertical fill) or one row
// (horizontal fill), with the rule that continues them.
type lane struct {
	cells []werkbook.CellSnapshot
	// linear is true when the lane continues a linear trend
	// value(p) = base + step*p, where p is the offset from the first source
	// cell. It also covers the date rule (one date cell, step 1).
	linear     bool
	base, step float64
}

// Fill extends the source range into the destination range as one undo step.
// dst must contain src and extend it on exactly one axis. See the fill rules
// in cmd/web/README.md.
func (s *Session) Fill(sheet string, src, dst Range, series bool) (FillResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return FillResult{}, err
	}
	if src, err = normRange(src); err != nil {
		return FillResult{}, fmt.Errorf("src: %w", err)
	}
	if dst, err = normRange(dst); err != nil {
		return FillResult{}, fmt.Errorf("dst: %w", err)
	}
	if !dst.containsRange(src) {
		return FillResult{}, errors.New("fill: the destination range must contain the source range")
	}
	res := FillResult{Range: dst}
	if dst == src {
		res.History = s.History()
		return res, nil
	}
	vertical := dst.Col1 == src.Col1 && dst.Col2 == src.Col2
	horizontal := dst.Row1 == src.Row1 && dst.Row2 == src.Row2
	switch {
	case vertical && (dst.Row1 == src.Row1 || dst.Row2 == src.Row2):
	case horizontal && (dst.Col1 == src.Col1 || dst.Col2 == src.Col2):
	default:
		return FillResult{}, errors.New("fill: the destination range must extend the source range in one direction")
	}
	if dst.cells()-src.cells() > maxOpCells {
		return FillResult{}, fmt.Errorf("fill: the destination has more than %d cells", maxOpCells)
	}
	if intersectsMerge(sheetMerges(sh), dst) {
		return FillResult{}, errors.New("fill: the range has merged cells")
	}

	s.ensureCalculated()
	spills := s.spillRects(sh)

	// source returns the state that the fill takes from a source cell. A
	// follower of an anchor that is in the source too gives an absent cell:
	// the copy of the anchor formula makes the followers again, and a copy
	// of their values would block that spill.
	source := func(row, col int) werkbook.CellSnapshot {
		if anchor, ok := followerAnchor(spills, row, col); ok && src.contains(anchor.row, anchor.col) {
			return sh.SnapshotCell(col, row).WithFormula("")
		}
		return sourceSnapshot(sh, spills, row, col)
	}

	// Read all lanes before the first write.
	var lanes []lane
	if vertical {
		for col := src.Col1; col <= src.Col2; col++ {
			cells := make([]werkbook.CellSnapshot, 0, src.rows())
			for row := src.Row1; row <= src.Row2; row++ {
				cells = append(cells, source(row, col))
			}
			lanes = append(lanes, newLane(cells, series))
		}
	} else {
		for row := src.Row1; row <= src.Row2; row++ {
			cells := make([]werkbook.CellSnapshot, 0, src.cols())
			for col := src.Col1; col <= src.Col2; col++ {
				cells = append(cells, source(row, col))
			}
			lanes = append(lanes, newLane(cells, series))
		}
	}

	err = s.mutate(labelFill, func(rec *recorder) error {
		rec.setSheet(sheet)
		rec.setRange(dst)
		for row := dst.Row1; row <= dst.Row2; row++ {
			for col := dst.Col1; col <= dst.Col2; col++ {
				if src.contains(row, col) {
					continue
				}
				// A follower is read-only unless its anchor is written too.
				if anchor, ok := followerAnchor(spills, row, col); ok {
					if !dst.contains(anchor.row, anchor.col) || src.contains(anchor.row, anchor.col) {
						res.Skipped++
						continue
					}
				}
				var snap werkbook.CellSnapshot
				if vertical {
					snap = lanes[col-src.Col1].at(row-src.Row1, false)
				} else {
					snap = lanes[row-src.Row1].at(col-src.Col1, true)
				}
				if err := rec.write(sh, col, row, snap); err != nil {
					return err
				}
				res.Changed++
			}
		}
		return nil
	})
	if err != nil {
		return FillResult{}, err
	}
	res.History = s.History()
	return res, nil
}

// newLane selects the rule of a lane.
func newLane(cells []werkbook.CellSnapshot, series bool) lane {
	ln := lane{cells: cells}
	if !series {
		return ln
	}
	for _, c := range cells {
		if c.Formula() != "" || c.Value().Type != werkbook.TypeNumber {
			return ln
		}
	}
	switch {
	case len(cells) >= 2:
		// Least-squares line through (0, y0), (1, y1), ... This is the trend
		// that Excel continues; an arithmetic sequence continues with its step.
		n := float64(len(cells))
		var sumX, sumY, sumXY, sumXX float64
		for i, c := range cells {
			x, y := float64(i), c.Value().Number
			sumX += x
			sumY += y
			sumXY += x * y
			sumXX += x * x
		}
		ln.step = (n*sumXY - sumX*sumY) / (n*sumXX - sumX*sumX)
		ln.base = (sumY - ln.step*sumX) / n
		ln.linear = true
	case len(cells) == 1:
		// One number copies, but one date counts days.
		if st := cells[0].Style(); st != nil && werkbook.IsDateFormat(st.NumFmt, st.NumFmtID) {
			ln.base, ln.step, ln.linear = cells[0].Value().Number, 1, true
		}
	}
	return ln
}

// at returns the cell state for the offset p from the first source cell. p is
// negative for a fill up or left. The source cells repeat as a pattern.
func (ln lane) at(p int, horizontal bool) werkbook.CellSnapshot {
	m := len(ln.cells)
	k := ((p % m) + m) % m
	srcCell := ln.cells[k]
	if ln.linear {
		v := roundSeries(ln.base + ln.step*float64(p))
		return srcCell.WithValue(werkbook.Value{Type: werkbook.TypeNumber, Number: v})
	}
	if f := srcCell.Formula(); f != "" {
		dCol, dRow := 0, p-k
		if horizontal {
			dCol, dRow = p-k, 0
		}
		return srcCell.WithFormula(formula.ShiftRefs(f, dCol, dRow))
	}
	// A value, or an empty cell: the destination gets the same state, so an
	// empty source cell clears the destination and gives it its style.
	return srcCell
}

// roundSeries removes the binary noise of a trend value: 0.1, 0.2 continues
// with 0.3 and not 0.30000000000000004. The value with 15 significant digits
// is used only when it is a neighbor float of the calculated value (1 ULP),
// which is the size of that noise. A value such as 16/3 has no short form and
// is stored as calculated.
func roundSeries(v float64) float64 {
	r, err := strconv.ParseFloat(strconv.FormatFloat(v, 'g', 15, 64), 64)
	if err != nil || r == v {
		return v
	}
	if r == math.Nextafter(v, math.Inf(1)) || r == math.Nextafter(v, math.Inf(-1)) {
		return r
	}
	return v
}
