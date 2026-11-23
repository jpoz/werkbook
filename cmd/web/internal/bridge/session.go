// Package bridge holds the session logic of the in-browser editor.
//
// It is pure Go (no syscall/js), so the usual toolchain can test it. The
// cmd/web/wasm package is a thin wrapper that exposes a Session to JavaScript.
// The function set and the result shapes follow the contract in
// cmd/web/README.md.
package bridge

import (
	"bytes"
	"errors"
	"fmt"
	"sort"

	werkbook "github.com/jpoz/werkbook"
)

// Version is the version string that the version() function reports. Set it
// at build time with -ldflags "-X .../bridge.Version=...".
var Version = "dev"

// Conversion factors from OOXML units to CSS pixels.
const (
	// colWidthCharPx is the width of one character of the default font
	// (Calibri 11) and colWidthPadPx is the cell padding that Excel adds.
	colWidthCharPx = 7.0
	colWidthPadPx  = 5.0
	// pointsToPx converts row heights, which OOXML stores in points.
	pointsToPx = 96.0 / 72.0
)

// SheetInfo describes one sheet of the workbook.
type SheetInfo struct {
	Name    string `json:"name"`
	MaxRow  int    `json:"maxRow"`
	MaxCol  int    `json:"maxCol"`
	Visible bool   `json:"visible"`
}

// Merge is one merged range. All coordinates are 1-based and inclusive.
type Merge struct {
	Row1 int `json:"row1"`
	Col1 int `json:"col1"`
	Row2 int `json:"row2"`
	Col2 int `json:"col2"`
}

// SheetMeta holds the layout data of a sheet. Sizes are in CSS pixels. The
// maps contain only the columns and rows that have a custom size.
type SheetMeta struct {
	ColWidths  map[int]int `json:"colWidths"`
	RowHeights map[int]int `json:"rowHeights"`
	Merges     []Merge     `json:"merges"`
	// Warnings lists the kinds of content of the sheet that the library
	// keeps unchanged and that do not move with rows and columns.
	Warnings []string `json:"warnings"`
	// Freeze is the frozen pane of the sheet.
	Freeze Freeze `json:"freeze"`
}

// Edit is one cell write for SetCells.
type Edit struct {
	Row   int    `json:"row"`
	Col   int    `json:"col"`
	Input string `json:"input"`
}

// rect is an inclusive cell rectangle.
type rect struct {
	row1, col1, row2, col2 int
}

func (r rect) contains(row, col int) bool {
	return row >= r.row1 && row <= r.row2 && col >= r.col1 && col <= r.col2
}

// Session holds one workbook. It is not safe for concurrent use; the WASM
// module calls it from the single JavaScript thread.
type Session struct {
	file *werkbook.File

	// stale is true when a write occurred after the last recalculation.
	stale bool
	// spills caches the spill rectangles of each sheet (anchor at row1/col1).
	// A recalculation clears it.
	spills map[string][]rect

	// undoStack and redoStack hold the history. The last element is the
	// next step to undo or redo. See history.go.
	undoStack, redoStack []*step
	// lastCleared is true when the last mutating call cleared the history
	// because its step was above the cost limit.
	lastCleared bool

	// clip is the internal clipboard, or nil. See clipboard.go.
	clip *clipboard
}

// NewSession returns a session that holds a new workbook with one empty sheet.
func NewSession() *Session {
	s := &Session{}
	s.New()
	return s
}

// New replaces the workbook with a new one that has one empty sheet. It
// clears the history and the internal clipboard.
func (s *Session) New() []SheetInfo {
	s.setFile(werkbook.New())
	return s.Sheets()
}

// Open replaces the workbook with the XLSX file in data, and clears the
// history and the internal clipboard. If the file cannot be read, the session
// keeps the workbook and the history it had.
func (s *Session) Open(data []byte) ([]SheetInfo, error) {
	if len(data) == 0 {
		return nil, errors.New("open: the file is empty")
	}
	f, err := werkbook.OpenReaderAt(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		if errors.Is(err, werkbook.ErrEncryptedFile) {
			return nil, errors.New("open: the file is encrypted or password-protected")
		}
		return nil, fmt.Errorf("open: %w", err)
	}
	s.setFile(f)
	return s.Sheets(), nil
}

// Save serializes the workbook to XLSX bytes.
func (s *Session) Save() ([]byte, error) {
	s.ensureCalculated()
	var buf bytes.Buffer
	if _, err := s.file.WriteTo(&buf); err != nil {
		return nil, fmt.Errorf("save: %w", err)
	}
	return buf.Bytes(), nil
}

// Sheets returns the sheets of the workbook in workbook order. MaxRow and
// MaxCol are the used range: the stored cells and the cells that dynamic
// array formulas spill into. A row that has only a custom height does not
// count.
func (s *Session) Sheets() []SheetInfo {
	s.ensureCalculated()
	names := s.file.SheetNames()
	out := make([]SheetInfo, 0, len(names))
	for _, name := range names {
		sh := s.file.Sheet(name)
		if sh == nil {
			continue
		}
		info := SheetInfo{Name: name, Visible: sh.Visible()}
		info.MaxRow, info.MaxCol = s.usedRange(sh)
		out = append(out, info)
	}
	return out
}

// SheetMeta returns the custom column widths, custom row heights and merged
// ranges of a sheet.
func (s *Session) SheetMeta(sheet string) (SheetMeta, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return SheetMeta{}, err
	}
	meta := SheetMeta{
		ColWidths:  map[int]int{},
		RowHeights: map[int]int{},
		Merges:     []Merge{},
		Warnings:   sheetWarnings(sh),
	}
	meta.Freeze.Rows, meta.Freeze.Cols = sh.FreezePanes()
	for col, w := range sh.ColumnWidths() {
		meta.ColWidths[col] = colWidthToPx(w)
	}
	for row := range sh.Rows() {
		if h := row.Height(); h > 0 {
			meta.RowHeights[row.Num()] = rowHeightToPx(h)
		}
	}
	for _, mr := range sh.MergeCells() {
		col1, row1, col2, row2, err := werkbook.RangeToCoordinates(mr.Start + ":" + mr.End)
		if err != nil {
			continue
		}
		meta.Merges = append(meta.Merges, Merge{Row1: row1, Col1: col1, Row2: row2, Col2: col2})
	}
	return meta, nil
}

// SetCell writes one cell as one undo step. See the input rules in
// cmd/web/README.md.
func (s *Session) SetCell(sheet string, row, col int, input string) error {
	sh, err := s.sheet(sheet)
	if err != nil {
		return err
	}
	if _, err := cellRef(row, col); err != nil {
		return err
	}
	return s.mutate(labelEditCell, func(rec *recorder) error {
		rec.setSheet(sheet)
		rec.setRange(Range{Row1: row, Col1: col, Row2: row, Col2: col})
		return rec.write(sh, col, row, inputSnapshot(sh.SnapshotCell(col, row), input))
	})
}

// SetCells writes many cells as one batch, one recalculation and one undo
// step. The call is atomic: an invalid edit leaves the workbook unchanged.
func (s *Session) SetCells(sheet string, edits []Edit) error {
	sh, err := s.sheet(sheet)
	if err != nil {
		return err
	}
	var box Range
	for i, e := range edits {
		if _, err := cellRef(e.Row, e.Col); err != nil {
			return fmt.Errorf("edit %d: %w", i, err)
		}
		if i == 0 {
			box = Range{Row1: e.Row, Col1: e.Col, Row2: e.Row, Col2: e.Col}
			continue
		}
		box.Row1, box.Col1 = min(box.Row1, e.Row), min(box.Col1, e.Col)
		box.Row2, box.Col2 = max(box.Row2, e.Row), max(box.Col2, e.Col)
	}
	if len(edits) == 0 {
		return nil
	}
	return s.mutate(labelEditCells, func(rec *recorder) error {
		rec.setSheet(sheet)
		rec.setRange(box)
		for i, e := range edits {
			// The base is the state at this point of the batch, so a later
			// edit of the same cell wins.
			snap := inputSnapshot(sh.SnapshotCell(e.Col, e.Row), e.Input)
			if err := rec.write(sh, e.Col, e.Row, snap); err != nil {
				return fmt.Errorf("edit %d: %w", i, err)
			}
		}
		return nil
	})
}

// AddSheet adds an empty sheet at the end of the workbook as one undo step.
func (s *Session) AddSheet(name string) ([]SheetInfo, error) {
	if name == "" {
		return nil, errors.New("sheet name is empty")
	}
	err := s.mutate(labelAddSheet, func(rec *recorder) error {
		if _, err := s.file.NewSheet(name); err != nil {
			return err
		}
		rec.other = append(rec.other, &sheetAddChange{name: name})
		// After an undo the sheet is gone; the step then shows the last sheet.
		rec.undoSheet, rec.redoSheet = "", name
		return nil
	})
	if err != nil {
		return nil, err
	}
	return s.Sheets(), nil
}

// RenameSheet renames a sheet as one undo step. Formulas that refer to the
// sheet follow.
func (s *Session) RenameSheet(oldName, newName string) ([]SheetInfo, error) {
	if _, err := s.sheet(oldName); err != nil {
		return nil, err
	}
	if newName == "" {
		return nil, errors.New("sheet name is empty")
	}
	if oldName == newName {
		return s.Sheets(), nil
	}
	err := s.mutate(labelRenameSheet, func(rec *recorder) error {
		if err := s.file.SetSheetName(oldName, newName); err != nil {
			return err
		}
		s.clipboardSheetRenamed(oldName, newName)
		rec.other = append(rec.other, &sheetRenameChange{from: oldName, to: newName})
		rec.undoSheet, rec.redoSheet = oldName, newName
		return nil
	})
	if err != nil {
		return nil, err
	}
	return s.Sheets(), nil
}

// DeleteSheet removes a sheet. The last sheet of a workbook cannot be
// removed. The call cannot be undone: it clears the history.
func (s *Session) DeleteSheet(name string) ([]SheetInfo, error) {
	if _, err := s.sheet(name); err != nil {
		return nil, err
	}
	if err := s.file.DeleteSheet(name); err != nil {
		return nil, err
	}
	s.stale = true
	s.clearHistory()
	s.clipboardSheetGone(name)
	return s.Sheets(), nil
}

func (s *Session) setFile(f *werkbook.File) {
	s.file = f
	s.stale = true
	s.spills = nil
	s.clip = nil
	s.clearHistory()
}

// ensureCalculated brings all formula values and spill ranges up to date.
// Reads call it, so a sequence of writes costs one recalculation.
func (s *Session) ensureCalculated() {
	if !s.stale {
		return
	}
	s.file.Recalculate()
	s.stale = false
	s.spills = nil
}

func (s *Session) sheet(name string) (*werkbook.Sheet, error) {
	sh := s.file.Sheet(name)
	if sh == nil {
		return nil, fmt.Errorf("sheet %q not found", name)
	}
	return sh, nil
}

// spillRects returns the published spill rectangles of a sheet, sorted by
// anchor in row-major order. Blocked spills (#SPILL!) have no rectangle.
// The caller must call ensureCalculated first.
func (s *Session) spillRects(sh *werkbook.Sheet) []rect {
	if rects, ok := s.spills[sh.Name()]; ok {
		return rects
	}
	rects := []rect{}
	for row := range sh.Rows() {
		for _, c := range row.Cells() {
			if !c.IsDynamicArraySpill() {
				continue
			}
			toCol, toRow, ok := sh.SpillBounds(c.Col(), row.Num())
			if !ok {
				continue
			}
			rects = append(rects, rect{row1: row.Num(), col1: c.Col(), row2: toRow, col2: toCol})
		}
	}
	sort.Slice(rects, func(i, j int) bool {
		if rects[i].row1 != rects[j].row1 {
			return rects[i].row1 < rects[j].row1
		}
		return rects[i].col1 < rects[j].col1
	})
	if s.spills == nil {
		s.spills = make(map[string][]rect)
	}
	s.spills[sh.Name()] = rects
	return rects
}

// cellRef validates 1-based coordinates and returns the A1 reference.
func cellRef(row, col int) (string, error) {
	if row < 1 || row > werkbook.MaxRows {
		return "", fmt.Errorf("row %d is out of range [1, %d]", row, werkbook.MaxRows)
	}
	if col < 1 || col > werkbook.MaxColumns {
		return "", fmt.Errorf("column %d is out of range [1, %d]", col, werkbook.MaxColumns)
	}
	return werkbook.CoordinatesToCellName(col, row)
}
