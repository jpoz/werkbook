package bridge

import (
	"fmt"
	"sort"

	werkbook "github.com/jpoz/werkbook"
)

const (
	labelInsertRows = "Insert rows"
	labelDeleteRows = "Delete rows"
	labelInsertCols = "Insert columns"
	labelDeleteCols = "Delete columns"
)

// StructureResult is the result of InsertRows, DeleteRows, InsertCols and
// DeleteCols.
type StructureResult struct {
	History HistoryState `json:"history"`
	Sheets  []SheetInfo  `json:"sheets"`
}

// structureChange is one insert or delete of rows or columns. The library
// gives the data to reverse it (werkbook.StructuralUndo). That data is valid
// only in strict last-in, first-out order, which the history keeps: a step is
// undone only after all later steps are undone.
type structureChange struct {
	sheet  string
	cols   bool
	at, n  int
	delete bool
	undo   *werkbook.StructuralUndo
}

func (c *structureChange) recalc() bool { return true }

func (c *structureChange) cost() int {
	if c.undo == nil {
		return 1
	}
	return max(1, c.undo.Cost())
}

func (c *structureChange) run(s *Session) error {
	sh := s.file.Sheet(c.sheet)
	if sh == nil {
		return fmt.Errorf("sheet %q not found", c.sheet)
	}
	var (
		u   *werkbook.StructuralUndo
		err error
	)
	switch {
	case c.cols && c.delete:
		u, err = sh.DeleteCols(c.at, c.n)
	case c.cols:
		u, err = sh.InsertCols(c.at, c.n)
	case c.delete:
		u, err = sh.DeleteRows(c.at, c.n)
	default:
		u, err = sh.InsertRows(c.at, c.n)
	}
	if err != nil {
		return err
	}
	c.undo = u
	return nil
}

func (c *structureChange) apply(s *Session, redo bool) error {
	if redo {
		// The edit runs again and gives new undo data.
		return c.run(s)
	}
	if c.undo == nil {
		return fmt.Errorf("internal error: structural step has no undo data")
	}
	err := c.undo.Undo()
	c.undo = nil
	return err
}

// InsertRows inserts count empty rows at row as one undo step.
func (s *Session) InsertRows(sheet string, row, count int) (StructureResult, error) {
	return s.structure(labelInsertRows, &structureChange{sheet: sheet, at: row, n: count})
}

// DeleteRows removes the rows row .. row+count-1 as one undo step.
func (s *Session) DeleteRows(sheet string, row, count int) (StructureResult, error) {
	return s.structure(labelDeleteRows, &structureChange{sheet: sheet, at: row, n: count, delete: true})
}

// InsertCols inserts count empty columns at col as one undo step.
func (s *Session) InsertCols(sheet string, col, count int) (StructureResult, error) {
	return s.structure(labelInsertCols, &structureChange{sheet: sheet, cols: true, at: col, n: count})
}

// DeleteCols removes the columns col .. col+count-1 as one undo step.
func (s *Session) DeleteCols(sheet string, col, count int) (StructureResult, error) {
	return s.structure(labelDeleteCols, &structureChange{sheet: sheet, cols: true, at: col, n: count, delete: true})
}

func (s *Session) structure(label string, c *structureChange) (StructureResult, error) {
	if _, err := s.sheet(c.sheet); err != nil {
		return StructureResult{}, err
	}
	// The library does the same checks; these give the messages of the
	// contract before the library is called.
	last, axis := werkbook.MaxRows, "row"
	if c.cols {
		last, axis = werkbook.MaxColumns, "column"
	}
	if c.at < 1 || c.at > last {
		return StructureResult{}, fmt.Errorf("%s %d is out of range [1, %d]", axis, c.at, last)
	}
	if c.n < 1 || c.at+c.n-1 > last {
		return StructureResult{}, fmt.Errorf("count %d is out of range: %ss %d and after must be in [1, %d]", c.n, axis, c.at, last)
	}

	err := s.mutate(label, func(rec *recorder) error {
		if err := c.run(s); err != nil {
			return err
		}
		rec.other = append(rec.other, c)
		rec.setSheet(c.sheet)
		if c.cols {
			rec.setRange(Range{Row1: 1, Col1: c.at, Row2: werkbook.MaxRows, Col2: c.at + c.n - 1})
		} else {
			rec.setRange(Range{Row1: c.at, Col1: 1, Row2: c.at + c.n - 1, Col2: werkbook.MaxColumns})
		}
		return nil
	})
	if err != nil {
		return StructureResult{}, err
	}
	return StructureResult{History: s.historyOfLastMutation(), Sheets: s.Sheets()}, nil
}

// passthroughWarnings maps the worksheet elements that the library keeps
// unchanged, and that hold cell addresses or objects at cell positions, to
// the text of a warning. A structural edit does not move them.
var passthroughWarnings = map[string]string{
	"conditionalFormatting": "conditional formats",
	"dataValidations":       "data validation",
	"hyperlinks":            "hyperlinks",
	"autoFilter":            "filters",
	"sortState":             "filters",
	"drawing":               "drawings and charts",
	"legacyDrawing":         "comments and form controls",
	"protectedRanges":       "protected ranges",
	"rowBreaks":             "page breaks",
	"colBreaks":             "page breaks",
	"extLst":                "extended features (for example sparklines)",
}

// sheetWarnings returns the kinds of content of a sheet that do not move
// with rows and columns, sorted, with no duplicates.
func sheetWarnings(sh *werkbook.Sheet) []string {
	seen := map[string]bool{}
	out := []string{}
	for _, name := range sh.PassthroughElementNames() {
		if text, ok := passthroughWarnings[name]; ok && !seen[text] {
			seen[text] = true
			out = append(out, text)
		}
	}
	sort.Strings(out)
	return out
}
