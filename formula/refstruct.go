package formula

import "strings"

// StructuralEdit describes an insert or a delete of rows or columns on one
// sheet.
type StructuralEdit struct {
	// Sheet is the name of the sheet that is edited.
	Sheet string
	// Cols is true for columns and false for rows.
	Cols bool
	// At is the first row or column of the edit (1-based) and Count is the
	// number of rows or columns. An insert puts Count new lines at At; the
	// line that was at At moves to At+Count. A delete removes the lines
	// At .. At+Count-1.
	At, Count int
	Delete    bool
	// SheetNames is the sheet order of the workbook. It is used only for 3D
	// references (Sheet1:Sheet3!A1), which follow the edit when the edited
	// sheet is in their span.
	SheetNames []string
}

// ShiftSpan applies the edit to the inclusive span lo..hi on the axis of the
// edit. It is the single place for the range rule: formula references,
// merged ranges, tables and spill ranges all use it.
//
// Insert: a span that starts at or after At moves by Count. A span that
// contains At (and starts before it) grows by Count. A span before At is
// unchanged.
//
// Delete: a span that is all in the deleted band is gone. A span that is
// partly in the band shrinks. A span after the band moves back by Count.
//
// A span that covers the full axis (1 .. the last row or column) means "all
// lines" and never changes. A span that an insert moves off the sheet is
// gone; a span that an insert grows past the end is cut at the end.
func (e StructuralEdit) ShiftSpan(lo, hi int) (lo2, hi2 int, gone bool) {
	last := refMaxRow
	if e.Cols {
		last = refMaxCol
	}
	if e.Count <= 0 || (lo <= 1 && hi >= last) {
		return lo, hi, false
	}
	if !e.Delete {
		switch {
		case lo >= e.At:
			lo, hi = lo+e.Count, hi+e.Count
		case hi >= e.At:
			hi += e.Count
		}
		if lo > last {
			return 0, 0, true
		}
		return lo, min(hi, last), false
	}

	first, end := e.At, e.At+e.Count-1 // the deleted band
	switch {
	case hi < first:
		// Before the band.
	case lo > end:
		lo, hi = lo-e.Count, hi-e.Count
	case lo >= first && hi <= end:
		return 0, 0, true
	case lo < first && hi <= end:
		hi = first - 1 // the end of the span is deleted
	case lo >= first:
		lo, hi = first, hi-e.Count // the start of the span is deleted
	default:
		hi -= e.Count // the band is in the span
	}
	return lo, hi, false
}

// appliesTo reports whether a reference with the given sheet qualifier, in a
// formula on formulaSheet, refers to the edited sheet.
func (e StructuralEdit) appliesTo(site RefSite, formulaSheet string) bool {
	switch {
	case site.Sheet == "" && site.SheetEnd == "":
		return strings.EqualFold(formulaSheet, e.Sheet)
	case site.SheetEnd == "":
		return strings.EqualFold(site.Sheet, e.Sheet)
	}
	// A 3D reference: the edited sheet must be in the span of the two
	// sheets, in workbook order.
	from, to, edited := -1, -1, -1
	for i, name := range e.SheetNames {
		if strings.EqualFold(name, site.Sheet) {
			from = i
		}
		if strings.EqualFold(name, site.SheetEnd) {
			to = i
		}
		if strings.EqualFold(name, e.Sheet) {
			edited = i
		}
	}
	if from < 0 || to < 0 || edited < 0 {
		return false
	}
	if from > to {
		from, to = to, from
	}
	return edited >= from && edited <= to
}

// ApplyStructuralEdit returns the formula text after an insert or a delete
// of rows or columns. formulaSheet is the sheet of the formula (or the scope
// sheet of a defined name); references with no sheet name refer to it.
//
// The references follow the cells that they refer to (see ShiftSpan).
// Absolute references are adjusted in the same way as relative references. A
// reference to deleted cells becomes #REF! with no sheet name. A column-only
// reference does not change in a row edit, and a row-only reference does not
// change in a column edit. All other text of the formula is unchanged.
//
// src is the formula without the leading '='. changed reports whether the
// text is different.
func ApplyStructuralEdit(src, formulaSheet string, e StructuralEdit) (out string, changed bool) {
	if e.Count <= 0 {
		return src, false
	}
	out = RewriteRefs(src, func(site RefSite) (RefSite, RefAction) {
		// A reference with no part on the axis of the edit.
		if (e.Cols && site.Kind == RefRows) || (!e.Cols && site.Kind == RefCols) {
			return site, RefKeep
		}
		if !e.appliesTo(site, formulaSheet) {
			return site, RefKeep
		}

		from, to := site.From.Row, site.To.Row
		if e.Cols {
			from, to = site.From.Col, site.To.Col
		}
		// The corners of a range are in text order, which can be reversed.
		swapped := from > to
		lo, hi := min(from, to), max(from, to)
		lo2, hi2, gone := e.ShiftSpan(lo, hi)
		if gone {
			return site, RefInvalid
		}
		if lo2 == lo && hi2 == hi {
			return site, RefKeep
		}
		if swapped {
			lo2, hi2 = hi2, lo2
		}
		if e.Cols {
			site.From.Col, site.To.Col = lo2, hi2
		} else {
			site.From.Row, site.To.Row = lo2, hi2
		}
		return site, RefReplace
	})
	return out, out != src
}
