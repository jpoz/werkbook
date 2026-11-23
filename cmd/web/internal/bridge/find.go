package bridge

import (
	"errors"
	"fmt"
	"strings"

	werkbook "github.com/jpoz/werkbook"
)

const (
	labelReplace    = "Replace"
	labelReplaceAll = "Replace all"
	// maxFindCount is the largest match count that Find reports.
	maxFindCount = 10_000
)

// CellPos is a cell position in a contract argument.
type CellPos struct {
	Row int `json:"row"`
	Col int `json:"col"`
}

// FindOptions are the options of Find, Replace and ReplaceAll.
type FindOptions struct {
	MatchCase bool `json:"matchCase"`
	WholeCell bool `json:"wholeCell"`
	// LookIn is "input" (default: the editor text, which is the formula for
	// a formula cell) or "display" (the text that the grid shows).
	LookIn string `json:"lookIn"`
	// From is the cell that the search starts after. Nil: before A1.
	From     *CellPos `json:"from"`
	Backward bool     `json:"backward"`
}

func (o FindOptions) validate() error {
	switch o.LookIn {
	case "", "input", "display":
	default:
		return fmt.Errorf(`lookIn %q is not "input" or "display"`, o.LookIn)
	}
	if o.From != nil {
		if o.From.Row < 0 || o.From.Row > werkbook.MaxRows || o.From.Col < 0 || o.From.Col > werkbook.MaxColumns {
			return fmt.Errorf("from %d,%d is out of range", o.From.Row, o.From.Col)
		}
	}
	return nil
}

// FindResult is the result of Find.
type FindResult struct {
	Found bool `json:"found"`
	Row   int  `json:"row"`
	Col   int  `json:"col"`
	// Count is the number of matching cells on the sheet, at most 10,000.
	Count int `json:"count"`
	// Wrapped is true when the search went past the end (or the start) of
	// the sheet to find the cell.
	Wrapped bool `json:"wrapped"`
}

// ReplaceResult is the result of Replace.
type ReplaceResult struct {
	History  HistoryState `json:"history"`
	Replaced bool         `json:"replaced"`
}

// ReplaceAllResult is the result of ReplaceAll.
type ReplaceAllResult struct {
	History  HistoryState `json:"history"`
	Replaced int          `json:"replaced"`
	// Skipped is the number of matching cells that are read-only spill
	// followers.
	Skipped int `json:"skipped"`
}

// searchText returns the text of a cell that a search reads.
//
// For the editor text ("input") a formula gives its text with the '='. A
// text cell gives its text without the apostrophe that the editor shows for
// text such as '007: the apostrophe is not content.
func searchText(c Cell, lookIn string) string {
	if lookIn == "display" {
		return c.Display
	}
	if !c.Formula && c.Type == "string" {
		return c.Display // the raw text
	}
	return c.Input
}

// matcher tests and replaces text by the options of a search.
type matcher struct {
	query     string
	lower     string
	matchCase bool
	wholeCell bool
}

func newMatcher(query string, opts FindOptions) matcher {
	return matcher{query: query, lower: strings.ToLower(query), matchCase: opts.MatchCase, wholeCell: opts.WholeCell}
}

func (m matcher) matches(text string) bool {
	if m.matchCase {
		if m.wholeCell {
			return text == m.query
		}
		return strings.Contains(text, m.query)
	}
	if m.wholeCell {
		return strings.EqualFold(text, m.query)
	}
	return strings.Contains(strings.ToLower(text), m.lower)
}

// replace returns text with each match replaced.
func (m matcher) replace(text, replacement string) string {
	if m.wholeCell {
		if m.matches(text) {
			return replacement
		}
		return text
	}
	if m.matchCase {
		return strings.ReplaceAll(text, m.query, replacement)
	}
	// Without letter case: find the matches in the lower-case form and cut
	// the same byte spans from the text. That is valid only when the
	// lower-case forms have the byte lengths of their sources (not so for a
	// few characters, for example "İ"); else only exact-case matches are
	// replaced.
	lowerText := strings.ToLower(text)
	if len(lowerText) != len(text) || len(m.lower) != len(m.query) {
		return strings.ReplaceAll(text, m.query, replacement)
	}
	var b strings.Builder
	i := 0
	for {
		j := strings.Index(lowerText[i:], m.lower)
		if j < 0 {
			break
		}
		b.WriteString(text[i : i+j])
		b.WriteString(replacement)
		i += j + len(m.lower)
	}
	b.WriteString(text[i:])
	return b.String()
}

// searchCells returns the cells of a sheet that a search visits, in
// row-major order: the cells that are not empty, without the cells that a
// merged range covers.
func (s *Session) searchCells(sh *werkbook.Sheet) ([]Cell, error) {
	s.ensureCalculated()
	maxRow, maxCol := s.usedRange(sh)
	if maxRow == 0 || maxCol == 0 {
		return nil, nil
	}
	cells, err := s.GetRange(sh.Name(), 1, 1, maxRow, maxCol)
	if err != nil {
		return nil, err
	}
	merges := sheetMerges(sh)
	out := cells[:0]
	for _, c := range cells {
		if c.Type == "empty" || coveredByMerge(merges, c.Row, c.Col) {
			continue
		}
		out = append(out, c)
	}
	return out, nil
}

// Find returns the next cell that matches the query, and the number of
// matching cells on the sheet. See the rules in cmd/web/README.md.
func (s *Session) Find(sheet, query string, opts FindOptions) (FindResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return FindResult{}, err
	}
	if query == "" {
		return FindResult{}, errors.New("find: the query is empty")
	}
	if err := opts.validate(); err != nil {
		return FindResult{}, err
	}
	cells, err := s.searchCells(sh)
	if err != nil {
		return FindResult{}, err
	}
	m := newMatcher(query, opts)
	var hits []CellPos
	for _, c := range cells {
		if m.matches(searchText(c, opts.LookIn)) {
			hits = append(hits, CellPos{Row: c.Row, Col: c.Col})
		}
	}
	res := FindResult{Count: min(len(hits), maxFindCount)}
	if len(hits) == 0 {
		return res, nil
	}

	var from CellPos
	if opts.From != nil {
		from = *opts.From
	}
	after := func(p CellPos) bool { return p.Row > from.Row || (p.Row == from.Row && p.Col > from.Col) }
	before := func(p CellPos) bool { return p.Row < from.Row || (p.Row == from.Row && p.Col < from.Col) }

	hit := hits[0]
	res.Wrapped = true
	if opts.Backward {
		hit = hits[len(hits)-1]
		for i := len(hits) - 1; i >= 0; i-- {
			if before(hits[i]) {
				hit, res.Wrapped = hits[i], false
				break
			}
		}
	} else {
		for _, p := range hits {
			if after(p) {
				hit, res.Wrapped = p, false
				break
			}
		}
	}
	res.Found, res.Row, res.Col = true, hit.Row, hit.Col
	return res, nil
}

// replacement holds one cell that a replace changes.
type replacement struct {
	pos   cellPos
	input string
}

// replacementFor returns the new editor text of a cell, and whether the
// cell matches and its text changes.
func replacementFor(c Cell, m matcher, with string) (input string, ok bool) {
	text := searchText(c, "input")
	if !m.matches(text) {
		return "", false
	}
	replaced := m.replace(text, with)
	if replaced == text {
		return "", false
	}
	// A text cell whose editor text has the apostrophe ('007) keeps it, so
	// that the cell stays text. All other new text goes through the input
	// rules, as an edit of the cell does.
	if !c.Formula && c.Type == "string" && strings.HasPrefix(c.Input, "'") && c.Input[1:] == text {
		return "'" + replaced, true
	}
	return replaced, true
}

// Replace replaces the query in one cell as one undo step. The search reads
// the editor text of the cell.
func (s *Session) Replace(sheet string, row, col int, query, with string, opts FindOptions) (ReplaceResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return ReplaceResult{}, err
	}
	if _, err := cellRef(row, col); err != nil {
		return ReplaceResult{}, err
	}
	if query == "" {
		return ReplaceResult{}, errors.New("replace: the query is empty")
	}
	if err := opts.validate(); err != nil {
		return ReplaceResult{}, err
	}
	// A cell that a merged range covers is searched through the top-left
	// cell of the range.
	if coveredByMerge(sheetMerges(sh), row, col) {
		return ReplaceResult{History: s.History()}, nil
	}
	cells, err := s.GetRange(sheet, row, col, row, col)
	if err != nil {
		return ReplaceResult{}, err
	}
	res := ReplaceResult{}
	if len(cells) == 1 && cells[0].Type != "empty" && !cells[0].Spill {
		if input, ok := replacementFor(cells[0], newMatcher(query, opts), with); ok {
			err := s.mutate(labelReplace, func(rec *recorder) error {
				rec.setSheet(sheet)
				rec.setRange(Range{Row1: row, Col1: col, Row2: row, Col2: col})
				return rec.write(sh, col, row, inputSnapshot(sh.SnapshotCell(col, row), input))
			})
			if err != nil {
				return ReplaceResult{}, err
			}
			res.Replaced = true
		}
	}
	res.History = s.History()
	return res, nil
}

// ReplaceAll replaces the query in all cells of a sheet as one undo step.
func (s *Session) ReplaceAll(sheet, query, with string, opts FindOptions) (ReplaceAllResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return ReplaceAllResult{}, err
	}
	if query == "" {
		return ReplaceAllResult{}, errors.New("replaceAll: the query is empty")
	}
	if err := opts.validate(); err != nil {
		return ReplaceAllResult{}, err
	}
	cells, err := s.searchCells(sh)
	if err != nil {
		return ReplaceAllResult{}, err
	}
	m := newMatcher(query, opts)
	res := ReplaceAllResult{}
	var todo []replacement
	var box Range
	for _, c := range cells {
		if c.Spill {
			if m.matches(searchText(c, "input")) {
				res.Skipped++
			}
			continue
		}
		input, ok := replacementFor(c, m, with)
		if !ok {
			continue
		}
		if len(todo) == 0 {
			box = Range{Row1: c.Row, Col1: c.Col, Row2: c.Row, Col2: c.Col}
		}
		box.Row1, box.Col1 = min(box.Row1, c.Row), min(box.Col1, c.Col)
		box.Row2, box.Col2 = max(box.Row2, c.Row), max(box.Col2, c.Col)
		todo = append(todo, replacement{pos: cellPos{col: c.Col, row: c.Row}, input: input})
	}
	if len(todo) > 0 {
		err := s.mutate(labelReplaceAll, func(rec *recorder) error {
			rec.setSheet(sheet)
			rec.setRange(box)
			for _, r := range todo {
				snap := inputSnapshot(sh.SnapshotCell(r.pos.col, r.pos.row), r.input)
				if err := rec.write(sh, r.pos.col, r.pos.row, snap); err != nil {
					return err
				}
			}
			return nil
		})
		if err != nil {
			return ReplaceAllResult{}, err
		}
		res.Replaced = len(todo)
	}
	res.History = s.History()
	return res, nil
}
