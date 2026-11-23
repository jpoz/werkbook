package bridge

import (
	"bytes"
	"fmt"
	"strings"
	"testing"

	werkbook "github.com/jpoz/werkbook"
	"github.com/jpoz/werkbook/ooxml"
)

// legacyArrayWorkbook returns an XLSX with a legacy (CSE) array formula in
// B1. The public API has no setter for one; files have them.
func legacyArrayWorkbook(t *testing.T) []byte {
	t.Helper()
	data := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{{
			Name: sheet1,
			Rows: []ooxml.RowData{
				{Num: 1, Cells: []ooxml.CellData{
					{Ref: "A1", Value: "2"},
					{Ref: "B1", Value: "6", Formula: "SUM(A1:A2*2)", FormulaType: "array", FormulaRef: "B1", IsArrayFormula: true},
				}},
				{Num: 2, Cells: []ooxml.CellData{{Ref: "A2", Value: "1"}, {Ref: "B2", Value: "9"}}},
			},
		}},
	}
	var buf bytes.Buffer
	if err := ooxml.WriteWorkbook(&buf, data); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func mustSort(t *testing.T, s *Session, rg string, header bool, keys ...SortKey) SortResult {
	t.Helper()
	res, err := s.SortRange(sheet1, rangeOf(t, rg), keys, header)
	if err != nil {
		t.Fatalf("SortRange(%s): %v", rg, err)
	}
	return res
}

// column returns the display text of the cells row1..row2 of a column, with
// "" for an empty cell.
func column(t *testing.T, s *Session, col string, row1, row2 int) []string {
	t.Helper()
	c, err := werkbook.ColumnNameToNumber(col)
	if err != nil {
		t.Fatal(err)
	}
	cells := cellsByRef(t, s, sheet1, row1, c, row2, c)
	var out []string
	for r := row1; r <= row2; r++ {
		out = append(out, cells[fmt.Sprintf("%s%d", col, r)].Display)
	}
	return out
}

func TestSortOrderOfTypes(t *testing.T) {
	// All types in one column, with a row id in column B.
	values := []string{"banana", "10", "=1/0", "", "TRUE", "Apple", "-5", "FALSE", "=NA()", "apple", "2.5", "'10", ""}
	setup := func(t *testing.T) *Session {
		s := NewSession()
		var pairs []string
		for i, v := range values {
			pairs = append(pairs, fmt.Sprintf("A%d", i+1), v, fmt.Sprintf("B%d", i+1), fmt.Sprintf("r%d", i+1))
		}
		mustSet(t, s, sheet1, pairs...)
		return s
	}
	last := len(values)

	t.Run("ascending", func(t *testing.T) {
		s := setup(t)
		before := dump(t, s, sheet1)
		res := mustSort(t, s, fmt.Sprintf("A1:B%d", last), false, SortKey{Col: 1})
		if res.Range != rangeOf(t, fmt.Sprintf("A1:B%d", last)) || res.History.UndoLabel != "Sort" {
			t.Errorf("result = %+v", res)
		}
		// Numbers, text (no letter case; equal text keeps its sequence),
		// FALSE, TRUE, errors in their sequence, empty cells last.
		wantA := []string{"-5", "2.5", "10", "10", "Apple", "apple", "banana", "FALSE", "TRUE", "#DIV/0!", "#N/A", "", ""}
		wantB := []string{"r7", "r11", "r2", "r12", "r6", "r10", "r1", "r8", "r5", "r3", "r9", "r4", "r13"}
		if got := column(t, s, "A", 1, last); strings.Join(got, "|") != strings.Join(wantA, "|") {
			t.Errorf("column A = %v\nwant       %v", got, wantA)
		}
		if got := column(t, s, "B", 1, last); strings.Join(got, "|") != strings.Join(wantB, "|") {
			t.Errorf("column B = %v\nwant       %v", got, wantB)
		}
		// The text "10" is text: it is after the numbers.
		if got := cellsByRef(t, s, sheet1, 4, 1, 4, 1)["A4"]; got.Type != "string" || got.Input != "'10" {
			t.Errorf("A4 = %+v, want the text 10", got)
		}
		checkUndoRedo(t, s, sheet1, before)
		if undo := mustUndo(t, s); undo.Label != "Sort" || *undo.Select != rangeOf(t, fmt.Sprintf("A1:B%d", last)) {
			t.Errorf("undo = %+v %v", undo, undo.Select)
		}
	})

	t.Run("descending keeps empty cells last", func(t *testing.T) {
		s := setup(t)
		mustSort(t, s, fmt.Sprintf("A1:B%d", last), false, SortKey{Col: 1, Desc: true})
		// The reverse of the type order; equal keys keep their sequence.
		wantA := []string{"#DIV/0!", "#N/A", "TRUE", "FALSE", "banana", "Apple", "apple", "10", "10", "2.5", "-5", "", ""}
		wantB := []string{"r3", "r9", "r5", "r8", "r1", "r6", "r10", "r12", "r2", "r11", "r7", "r4", "r13"}
		if got := column(t, s, "A", 1, last); strings.Join(got, "|") != strings.Join(wantA, "|") {
			t.Errorf("column A = %v\nwant       %v", got, wantA)
		}
		if got := column(t, s, "B", 1, last); strings.Join(got, "|") != strings.Join(wantB, "|") {
			t.Errorf("column B = %v\nwant       %v", got, wantB)
		}
	})
}

// Text sorts as a reader expects, not by bytes: a letter with an accent
// sorts with its base letter, digits are before symbols, symbols are before
// letters, and other scripts are after the Latin letters.
func TestSortTextOrder(t *testing.T) {
	sorted := func(t *testing.T, desc bool, values ...string) []string {
		t.Helper()
		s := NewSession()
		edits := make([]Edit, len(values))
		for i, v := range values {
			edits[i] = Edit{Row: i + 1, Col: 1, Input: "'" + v}
		}
		if err := s.SetCells(sheet1, edits); err != nil {
			t.Fatal(err)
		}
		mustSort(t, s, fmt.Sprintf("A1:A%d", len(values)), false, SortKey{Col: 1, Desc: desc})
		return column(t, s, "A", 1, len(values))
	}

	// The case of the review.
	got := sorted(t, false, "b", "Zeta", "alpha", "Émile", "a-b", "ab", "_x", "10", "9")
	want := []string{"10", "9", "_x", "a-b", "ab", "alpha", "b", "Émile", "Zeta"}
	if strings.Join(got, "|") != strings.Join(want, "|") {
		t.Errorf("ascending = %v\nwant        %v", got, want)
	}
	got = sorted(t, true, "b", "Zeta", "alpha", "Émile", "a-b", "ab", "_x", "10", "9")
	want = []string{"Zeta", "Émile", "b", "alpha", "ab", "a-b", "_x", "9", "10"}
	if strings.Join(got, "|") != strings.Join(want, "|") {
		t.Errorf("descending = %v\nwant         %v", got, want)
	}

	tests := []struct {
		name   string
		values []string
		want   []string
	}{
		{"accents sort with their base letter",
			[]string{"zebra", "Öl", "Über", "Ärger", "apple", "ober", "unter", "Çelik", "cat", "ñandú", "nube", "Ångström"},
			[]string{"Ångström", "apple", "Ärger", "cat", "Çelik", "ñandú", "nube", "ober", "Öl", "Über", "unter", "zebra"}},
		{"text that differs only in an accent has one order",
			[]string{"résumé", "resume", "Resume", "résume"},
			[]string{"resume", "Resume", "résume", "résumé"}},
		{"letter case is not important and keeps the sequence",
			[]string{"b", "B", "a", "A", "b"},
			[]string{"a", "A", "b", "B", "b"}},
		{"ligatures and sharp s", []string{"strasse", "straße", "strand", "œuvre", "oeil", "of"},
			[]string{"oeil", "œuvre", "of", "strand", "strasse", "straße"}},
		{"Central European letters", []string{"Łódź", "Lodz", "Zürich", "Žilina", "Šibenik", "Sibiu", "Ďumbier", "Dublin"},
			[]string{"Dublin", "Ďumbier", "Lodz", "Łódź", "Šibenik", "Sibiu", "Žilina", "Zürich"}},
		{"other scripts are after the Latin letters",
			[]string{"яблоко", "zeta", "Ωmega", "алма", "alpha", "βeta", "日本"},
			[]string{"alpha", "zeta", "βeta", "Ωmega", "алма", "яблоко", "日本"}},
		{"digits, then symbols, then letters",
			[]string{"a1", "#tag", "1a", " space", "A", "!x", "2", "_y"},
			[]string{"1a", "2", " space", "!x", "#tag", "_y", "A", "a1"}},
		{"a shorter text is before a longer text with the same start",
			[]string{"abc", "ab", "a", "abcd"},
			[]string{"a", "ab", "abc", "abcd"}},
		{"a combining mark is the same as the composed letter",
			[]string{"zz", "éa", "eb", "éc"},
			[]string{"éa", "eb", "éc", "zz"}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := sorted(t, false, tt.values...); strings.Join(got, "|") != strings.Join(tt.want, "|") {
				t.Errorf("sorted = %q\nwant     %q", got, tt.want)
			}
		})
	}
}

func TestCompareText(t *testing.T) {
	tests := []struct {
		a, b string
		want int
	}{
		{"a", "b", -1}, {"b", "a", 1}, {"a", "a", 0},
		{"Apple", "apple", 0}, {"APPLE", "apple", 0},
		{"Émile", "Zeta", -1}, {"Émile", "b", 1}, {"emile", "Émile", -1},
		{"ab", "abc", -1}, {"", "a", -1}, {"", "", 0},
		{"10", "9", -1}, // text: by characters, not by value
		{"9", "a", -1}, {"_", "a", -1}, {"9", "_", -1},
		{"z", "я", -1}, {"ß", "ss", 1}, {"ss", "st", -1}, {"ß", "st", -1},
	}
	for _, tt := range tests {
		got := compareText(tt.a, tt.b)
		if (got < 0) != (tt.want < 0) || (got > 0) != (tt.want > 0) {
			t.Errorf("compareText(%q, %q) = %d, want sign %d", tt.a, tt.b, got, tt.want)
		}
		// The compare is antisymmetric.
		back := compareText(tt.b, tt.a)
		if (got < 0) != (back > 0) || (got == 0) != (back == 0) {
			t.Errorf("compareText(%q, %q) = %d but the reverse is %d", tt.a, tt.b, got, back)
		}
	}
}

func TestSortStableHeaderAndKeys(t *testing.T) {
	setup := func(t *testing.T) *Session {
		s := NewSession()
		mustSet(t, s, sheet1,
			"A1", "Name", "B1", "Dept", "C1", "Pay",
			"A2", "dan", "B2", "ops", "C2", "30",
			"A3", "amy", "B3", "dev", "C3", "50",
			"A4", "bob", "B4", "ops", "C4", "40",
			"A5", "cat", "B5", "dev", "C5", "50",
			"A6", "eve", "B6", "ops", "C6", "30",
		)
		return s
	}

	t.Run("header row stays", func(t *testing.T) {
		s := setup(t)
		res := mustSort(t, s, "A1:C6", true, SortKey{Col: 1})
		if got := column(t, s, "A", 1, 6); strings.Join(got, " ") != "Name amy bob cat dan eve" {
			t.Errorf("column A = %v", got)
		}
		if got := column(t, s, "C", 1, 6); strings.Join(got, " ") != "Pay 50 40 50 30 30" {
			t.Errorf("column C = %v; full rows must move", got)
		}
		if res.Changed == 0 {
			t.Errorf("changed = 0")
		}
	})

	t.Run("without header the first row is data", func(t *testing.T) {
		s := setup(t)
		mustSort(t, s, "A1:C6", false, SortKey{Col: 1})
		// "Name" is text and sorts with the names.
		if got := column(t, s, "A", 1, 6); strings.Join(got, " ") != "amy bob cat dan eve Name" {
			t.Errorf("column A = %v", got)
		}
	})

	t.Run("stable: equal keys keep their sequence", func(t *testing.T) {
		s := setup(t)
		mustSort(t, s, "A1:C6", true, SortKey{Col: 2})
		// dev: amy, cat; ops: dan, bob, eve (the sequence of the sheet).
		if got := column(t, s, "A", 2, 6); strings.Join(got, " ") != "amy cat dan bob eve" {
			t.Errorf("column A = %v", got)
		}
	})

	t.Run("two keys", func(t *testing.T) {
		s := setup(t)
		mustSort(t, s, "A1:C6", true, SortKey{Col: 3, Desc: true}, SortKey{Col: 1, Desc: true})
		// Pay 50: cat, amy; 40: bob; 30: eve, dan.
		if got := column(t, s, "A", 2, 6); strings.Join(got, " ") != "cat amy bob eve dan" {
			t.Errorf("column A = %v", got)
		}
	})

	t.Run("a range that is in order makes no step", func(t *testing.T) {
		s := setup(t)
		mustSort(t, s, "A1:C6", true, SortKey{Col: 1})
		depth := s.History().Depth
		res := mustSort(t, s, "A1:C6", true, SortKey{Col: 1})
		if res.Changed != 0 || res.History.Depth != depth {
			t.Errorf("result = %+v, want no change", res)
		}
	})

	t.Run("cells out of the range do not move", func(t *testing.T) {
		s := setup(t)
		mustSet(t, s, sheet1, "D2", "x2", "D6", "x6")
		mustSort(t, s, "A2:C6", false, SortKey{Col: 1})
		if got := column(t, s, "D", 2, 6); strings.Join(got, "|") != "x2||||x6" {
			t.Errorf("column D = %v", got)
		}
	})

	t.Run("one data row and an empty range change nothing", func(t *testing.T) {
		s := setup(t)
		depth := s.History().Depth
		for _, c := range []struct {
			rg     string
			header bool
		}{{"A1:C2", true}, {"A1:C1", false}, {"A1:C1", true}, {"H1:J9", false}} {
			if res := mustSort(t, s, c.rg, c.header, SortKey{Col: rangeOf(t, c.rg).Col1}); res.Changed != 0 || res.History.Depth != depth {
				t.Errorf("%s: result = %+v", c.rg, res)
			}
		}
	})

	t.Run("full columns are limited to the used range", func(t *testing.T) {
		s := setup(t)
		res, err := s.SortRange(sheet1, Range{1, 1, werkbook.MaxRows, 3}, []SortKey{{Col: 1}}, true)
		if err != nil {
			t.Fatal(err)
		}
		if res.Range != rangeOf(t, "A1:C6") {
			t.Errorf("range = %+v, want A1:C6", res.Range)
		}
		if got := column(t, s, "A", 1, 6); strings.Join(got, " ") != "Name amy bob cat dan eve" {
			t.Errorf("column A = %v", got)
		}
	})
}

func TestSortMovesFormulasAndStyles(t *testing.T) {
	s := NewSession()
	sh := s.file.Sheet(sheet1)
	bold := &werkbook.Style{Font: &werkbook.Font{Bold: true}}
	if err := sh.SetStyle("A1", bold); err != nil {
		t.Fatal(err)
	}
	// Column B is a formula on the row; column C reads an absolute cell and
	// the row above.
	mustSet(t, s, sheet1,
		"A1", "30", "B1", "=A1*2", "C1", "=$E$1+A1",
		"A2", "10", "B2", "=A2*2", "C2", "=$E$1+A2",
		"A3", "20", "B3", "=A3*2", "C3", "=$E$1+A3",
		"E1", "100",
	)
	before := dump(t, s, sheet1)
	mustSort(t, s, "A1:C3", false, SortKey{Col: 1})

	// A formula moves as a copy does: relative references follow the row.
	wantInputs(t, s, sheet1, map[string]string{
		"A1": "10", "B1": "=A1*2", "C1": "=$E$1+A1",
		"A2": "20", "B2": "=A2*2", "C2": "=$E$1+A2",
		"A3": "30", "B3": "=A3*2", "C3": "=$E$1+A3",
		"E1": "100",
	})
	if got := column(t, s, "B", 1, 3); strings.Join(got, " ") != "20 40 60" {
		t.Errorf("column B = %v", got)
	}
	if got := column(t, s, "C", 1, 3); strings.Join(got, " ") != "110 120 130" {
		t.Errorf("column C = %v", got)
	}
	// The style moved with its row (30 is now on row 3).
	if st, _ := sh.GetStyle("A3"); st != bold {
		t.Errorf("A3 style = %+v, want the style of the old A1", st)
	}
	if st, _ := sh.GetStyle("A1"); st != nil {
		t.Errorf("A1 style = %+v, want none", st)
	}
	checkUndoRedo(t, s, sheet1, before)

	// A sort by a formula column uses the shown values.
	mustSet(t, s, sheet1, "G1", "=10-H1", "H1", "1", "G2", "=10-H2", "H2", "5", "G3", "=10-H3", "H3", "3")
	mustSort(t, s, "G1:H3", false, SortKey{Col: 7})
	if got := column(t, s, "H", 1, 3); strings.Join(got, " ") != "5 3 1" {
		t.Errorf("column H after a sort by the formula column = %v", got)
	}
	if got := column(t, s, "G", 1, 3); strings.Join(got, " ") != "5 7 9" {
		t.Errorf("column G = %v", got)
	}

	// A formula that reads a different row: the reference shifts with the
	// row, and one that would leave the sheet becomes #REF!.
	s = NewSession()
	mustSet(t, s, sheet1, "A1", "2", "B1", "=A3", "A2", "1", "B2", "=A1", "A3", "3", "B3", "=A2")
	mustSort(t, s, "A1:B3", false, SortKey{Col: 1})
	wantInputs(t, s, sheet1, map[string]string{
		"A1": "1", "B1": "=#REF!", // was row 2 with =A1; one row up
		"A2": "2", "B2": "=A4", // was row 1 with =A3; one row down
		"A3": "3", "B3": "=A2", // the row did not move
	})
}

func TestSortErrors(t *testing.T) {
	s := NewSession()
	sh := s.file.Sheet(sheet1)
	mustSet(t, s, sheet1, "A1", "3", "A2", "1", "A3", "2", "B1", "x", "B2", "y", "B3", "z",
		"D1", "=SEQUENCE(3)", "F1", "5", "F2", "4", "G1", "=SEQUENCE(1,2)", "H1", "block")
	if err := sh.MergeCell("J1", "K2"); err != nil {
		t.Fatal(err)
	}
	before := dump(t, s, sheet1)
	history := s.History()
	R := func(a1 string) Range { return rangeOf(t, a1) }
	one := []SortKey{{Col: 1}}

	tests := []struct {
		name  string
		sheet string
		rg    Range
		keys  []SortKey
		want  string
	}{
		{"unknown sheet", "Nope", R("A1:B3"), one, "not found"},
		{"range out of the sheet", sheet1, Range{0, 1, 3, 2}, one, "out of range"},
		{"no keys", sheet1, R("A1:B3"), nil, "no sort key"},
		{"key column left of the range", sheet1, R("B1:C3"), one, "not in the range"},
		{"key column right of the range", sheet1, R("A1:B3"), []SortKey{{Col: 3}}, "not in the range"},
		{"second key out of the range", sheet1, R("A1:B3"), []SortKey{{Col: 1}, {Col: 9}}, "key 1"},
		{"merged range", sheet1, R("J1:L3"), []SortKey{{Col: 10}}, "merged"},
		{"spill anchor and followers", sheet1, R("D1:D3"), []SortKey{{Col: 4}}, "spill"},
		{"spill followers only", sheet1, R("D2:E3"), []SortKey{{Col: 4}}, "spill"},
		{"blocked spill anchor", sheet1, R("F1:G2"), []SortKey{{Col: 6}}, "spill"},
		{"more than the cell limit", sheet1, Range{1, 1, 100_001, 2}, one, "200000"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, err := s.SortRange(tt.sheet, tt.rg, tt.keys, false)
			if err == nil || !strings.Contains(err.Error(), tt.want) {
				t.Fatalf("error = %v, want it to contain %q", err, tt.want)
			}
			if s.History() != history {
				t.Errorf("history changed")
			}
			if got := dump(t, s, sheet1); got != before {
				t.Errorf("the workbook changed")
			}
		})
	}

	t.Run("legacy array formula", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "2", "A2", "1", "B1", "=A1", "B2", "=A2")
		// The public API has no setter for a legacy array formula; make the
		// stored state through a snapshot of an opened file in history_test.
		src := legacyArrayWorkbook(t)
		if _, err := s.Open(src); err != nil {
			t.Fatal(err)
		}
		if _, err := s.SortRange(sheet1, R("A1:B2"), one, false); err == nil || !strings.Contains(err.Error(), "array formula") {
			t.Errorf("error = %v, want an array formula error", err)
		}
	})
}
