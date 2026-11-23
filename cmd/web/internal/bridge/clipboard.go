package bridge

import (
	"crypto/rand"
	"encoding/hex"
	"fmt"

	werkbook "github.com/jpoz/werkbook"
	"github.com/jpoz/werkbook/formula"
)

// maxCopyCells is the largest range that Copy makes text for.
const maxCopyCells = 2_000_000

// clipboard is the internal clipboard of the session: a snapshot of the cells
// of the last Copy. The system clipboard holds the text and, when the browser
// permits a custom clipboard type, the id. Paste finds the snapshot again by
// the id; the compare of the text is the fallback.
type clipboard struct {
	// id is random, so that the clipboards of two sessions (two browser
	// tabs) never have the same id.
	id    string
	sheet string
	rg    Range
	// cells holds the snapshot of each source cell that is not absent, by
	// position on the source sheet. A spill follower whose anchor is not in
	// the range is stored as its value. A follower whose anchor is in the
	// range is not stored: a paste of the anchor formula makes it again.
	cells map[cellPos]werkbook.CellSnapshot
	text  string // normalized TSV text
	cut   bool
	// merged is true when the source range had merged cells at copy time.
	merged bool
}

// CopyResult is the result of Copy.
type CopyResult struct {
	// Text is the TSV text for the system clipboard.
	Text string `json:"text"`
	// ID identifies the internal clipboard of this copy. It is "" when
	// Internal is false.
	ID string `json:"id"`
	// Cells is the number of cells in the range that are not empty.
	Cells int `json:"cells"`
	// Internal is true when the session kept a snapshot of the cells, so a
	// paste copies formulas and styles.
	Internal bool         `json:"internal"`
	History  HistoryState `json:"history"`
}

// PasteResult is the result of Paste.
type PasteResult struct {
	History HistoryState `json:"history"`
	// Range is the range that was written.
	Range Range `json:"range"`
	// Changed is the number of destination cells that were written.
	Changed int `json:"changed"`
	// Skipped is the number of read-only spill cells that were not written.
	Skipped int `json:"skipped"`
	// Clipped is true when a part of the data did not fit on the sheet.
	Clipped bool `json:"clipped"`
	// Mode is "internal" (formulas and styles from the internal clipboard)
	// or "text" (TSV text through the input rules).
	Mode string `json:"mode"`
}

// clipboardSheetGone clears the internal clipboard when its source sheet is
// removed.
func (s *Session) clipboardSheetGone(name string) {
	if s.clip != nil && s.clip.sheet == name {
		s.clip = nil
	}
}

// clipboardSheetRenamed keeps the source sheet name of the clipboard current.
func (s *Session) clipboardSheetRenamed(from, to string) {
	if s.clip != nil && s.clip.sheet == from {
		s.clip.sheet = to
	}
}

// cutMarkChange keeps the cut mark of the clipboard in step with the "Cut"
// step. After an undo of the cut, the source cells are back, so a paste is a
// copy (formulas shift). A redo makes the clipboard a cut again.
type cutMarkChange struct{ id string }

func (c *cutMarkChange) cost() int    { return 0 }
func (c *cutMarkChange) recalc() bool { return false }

func (c *cutMarkChange) apply(s *Session, redo bool) error {
	if s.clip != nil && s.clip.id == c.id {
		s.clip.cut = redo
	}
	return nil
}

func newClipboardID() string {
	var b [12]byte
	if _, err := rand.Read(b[:]); err != nil {
		// No random source: the text compare of Paste still finds the
		// clipboard, so an id that is not unique across sessions is the
		// only loss.
		return fmt.Sprintf("local-%p", &b)
	}
	return hex.EncodeToString(b[:])
}

// Copy returns the TSV text of a range and keeps a snapshot of its cells as
// the internal clipboard. With cut, it also clears the range as one undo
// step. A full-row or full-column range is limited to the used range.
func (s *Session) Copy(sheet string, rg Range, cut bool) (CopyResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return CopyResult{}, err
	}
	rg, err = normRange(rg)
	if err != nil {
		return CopyResult{}, err
	}
	s.ensureCalculated()
	spills := s.spillRects(sh)

	// Limit full rows and full columns to the used range.
	maxRow, maxCol := s.usedRange(sh)
	if rg.Col1 == 1 && rg.Col2 == werkbook.MaxColumns {
		rg.Col2 = max(rg.Col1, min(rg.Col2, maxCol))
	}
	if rg.Row1 == 1 && rg.Row2 == werkbook.MaxRows {
		rg.Row2 = max(rg.Row1, min(rg.Row2, maxRow))
	}
	if rg.cells() > maxCopyCells {
		return CopyResult{}, fmt.Errorf("copy: the range has more than %d cells", maxCopyCells)
	}

	// Text. A formula cell and a spill cell give their shown value, so that
	// other programs get values. A constant gives its editor text: that is
	// the text that the input rules read back to the same value and type
	// (a formatted number "50%" or "1,234.50" would come back as text).
	shown, err := s.GetRange(sheet, rg.Row1, rg.Col1, rg.Row2, rg.Col2)
	if err != nil {
		return CopyResult{}, err
	}
	grid := make([][]string, rg.rows())
	for i := range grid {
		grid[i] = make([]string, rg.cols())
	}
	res := CopyResult{}
	for _, c := range shown {
		if c.Type == "empty" {
			continue
		}
		text := c.Input
		if c.Formula || c.Spill {
			text = c.Display
		}
		grid[c.Row-rg.Row1][c.Col-rg.Col1] = text
		res.Cells++
	}
	res.Text = writeTSV(grid)

	// Snapshot.
	s.clip = nil
	if res.Cells <= maxOpCells {
		clip := &clipboard{
			id:     newClipboardID(),
			sheet:  sheet,
			rg:     rg,
			cells:  make(map[cellPos]werkbook.CellSnapshot),
			text:   normalizeClipboardText(res.Text),
			cut:    cut,
			merged: intersectsMerge(sheetMerges(sh), rg),
		}
		for row := range sh.Rows() {
			if row.Num() < rg.Row1 {
				continue
			}
			if row.Num() > rg.Row2 {
				break
			}
			for _, c := range row.Cells() {
				if c.Col() >= rg.Col1 && c.Col() <= rg.Col2 {
					clip.cells[cellPos{col: c.Col(), row: row.Num()}] = sh.SnapshotCell(c.Col(), row.Num())
				}
			}
		}
		// Followers with no anchor in the range are copied as values.
		for _, sp := range spills {
			if rg.contains(sp.row1, sp.col1) {
				continue
			}
			for row := max(sp.row1, rg.Row1); row <= min(sp.row2, rg.Row2); row++ {
				for col := max(sp.col1, rg.Col1); col <= min(sp.col2, rg.Col2); col++ {
					snap := sourceSnapshot(sh, spills, row, col)
					if snap.Exists() {
						clip.cells[cellPos{col: col, row: row}] = snap
					}
				}
			}
		}
		s.clip = clip
		res.Internal = true
		res.ID = clip.id
	}

	if cut {
		err := s.mutate(labelCut, func(rec *recorder) error {
			rec.setSheet(sheet)
			rec.setRange(rg)
			changed, _, err := clearCells(rec, sh, rg, spills)
			// A cut of empty cells makes no step, so it gets no mark.
			if err == nil && changed > 0 && s.clip != nil {
				rec.other = append(rec.other, &cutMarkChange{id: s.clip.id})
			}
			return err
		})
		if err != nil {
			s.clip = nil
			return CopyResult{}, err
		}
	}
	res.History = s.History()
	return res, nil
}

// Paste writes clipboard text to a sheet as one undo step.
//
// clipID is the id that the system clipboard holds next to the text: the id
// of a Copy result, or "" when the clipboard has no id (the text comes from a
// different program). The internal clipboard is used only when the id is its
// id. A nil clipID means that the caller cannot read the id; the internal
// clipboard is then used when text is equal to its text.
//
// With the internal clipboard, formulas and styles are copied
// (mode "internal"). Else the text is read as TSV (mode "text"). See the
// rules in cmd/web/README.md.
func (s *Session) Paste(sheet string, dst Range, text string, clipID *string) (PasteResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return PasteResult{}, err
	}
	dst, err = normRange(dst)
	if err != nil {
		return PasteResult{}, err
	}

	clip := s.clip
	internal := false
	switch {
	case clip == nil:
	case clipID != nil:
		internal = *clipID != "" && *clipID == clip.id
	default:
		internal = normalizeClipboardText(text) == clip.text
	}

	// targetFor returns the range that data of h rows and w columns is
	// written to. The data repeats when the destination is a whole multiple
	// of its size on both axes. Else it is pasted one time at the top-left
	// cell.
	clipped := false
	targetFor := func(h, w int) Range {
		target := Range{Row1: dst.Row1, Col1: dst.Col1, Row2: dst.Row1 + h - 1, Col2: dst.Col1 + w - 1}
		if dst.rows()%h == 0 && dst.cols()%w == 0 {
			target = dst
		}
		clipped = false
		if target.Row2 > werkbook.MaxRows {
			target.Row2 = werkbook.MaxRows
			clipped = true
		}
		if target.Col2 > werkbook.MaxColumns {
			target.Col2 = werkbook.MaxColumns
			clipped = true
		}
		return target
	}

	merges := sheetMerges(sh)
	var (
		target Range
		h, w   int
		grid   [][]string
	)
	if internal {
		h, w = clip.rg.rows(), clip.rg.cols()
		target = targetFor(h, w)
		// Merged cells have no cell-by-cell copy. The values are still
		// pasted, through the text path.
		if clip.merged || intersectsMerge(merges, target) {
			internal = false
		}
	}
	if !internal {
		// Empty text is "nothing to paste", not "clear the cell".
		if normalizeClipboardText(text) == "" {
			return PasteResult{
				History: s.History(),
				Range:   Range{Row1: dst.Row1, Col1: dst.Col1, Row2: dst.Row1, Col2: dst.Col1},
				Mode:    "text",
			}, nil
		}
		grid = parseTSV(text)
		h, w = len(grid), 0
		for _, row := range grid {
			w = max(w, len(row))
		}
		target = targetFor(h, w)
	}
	res := PasteResult{Mode: "text", Range: target, Clipped: clipped}
	if internal {
		res.Mode = "internal"
	}
	if target.cells() > maxOpCells {
		return PasteResult{}, fmt.Errorf("paste: the destination has more than %d cells", maxOpCells)
	}

	s.ensureCalculated()
	spills := s.spillRects(sh)

	err = s.mutate(labelPaste, func(rec *recorder) error {
		rec.setSheet(sheet)
		rec.setRange(target)
		for row := target.Row1; row <= target.Row2; row++ {
			for col := target.Col1; col <= target.Col2; col++ {
				// A follower is read-only unless its anchor is written too.
				if anchor, ok := followerAnchor(spills, row, col); ok && !target.contains(anchor.row, anchor.col) {
					res.Skipped++
					continue
				}
				i, j := (row-target.Row1)%h, (col-target.Col1)%w

				var snap werkbook.CellSnapshot
				if internal {
					srcPos := cellPos{col: clip.rg.Col1 + j, row: clip.rg.Row1 + i}
					snap = clip.cells[srcPos]
					if f := snap.Formula(); f != "" {
						snap = snap.WithFormula(pastedFormula(f, clip, sheet, col-srcPos.col, row-srcPos.row))
					}
				} else {
					// A merged range takes text in its top-left cell only.
					if coveredByMerge(merges, row, col) {
						continue
					}
					field := ""
					if j < len(grid[i]) {
						field = grid[i][j]
					}
					snap = inputSnapshot(sh.SnapshotCell(col, row), field)
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
		return PasteResult{}, err
	}
	res.History = s.History()
	return res, nil
}

// pastedFormula returns the formula text for an internal paste on dstSheet,
// dCol columns and dRow rows from the source cell.
//
// A copy shifts the relative references, and a reference with no sheet name
// then reads the destination sheet. A cut is a move: the formula must read
// the same cells as before. Its references do not shift, and on a different
// sheet each reference with no sheet name gets the name of the source sheet.
func pastedFormula(f string, clip *clipboard, dstSheet string, dCol, dRow int) string {
	if !clip.cut {
		return formula.ShiftRefs(f, dCol, dRow)
	}
	if dstSheet == clip.sheet {
		return f
	}
	return formula.RewriteRefs(f, func(site formula.RefSite) (formula.RefSite, formula.RefAction) {
		if site.Sheet != "" || site.SheetEnd != "" {
			return site, formula.RefKeep
		}
		site.Sheet = clip.sheet
		return site, formula.RefReplace
	})
}

// coveredByMerge reports whether a cell is in a merged range and is not its
// top-left cell.
func coveredByMerge(merges []Range, row, col int) bool {
	for _, m := range merges {
		if m.contains(row, col) && (row != m.Row1 || col != m.Col1) {
			return true
		}
	}
	return false
}
