package bridge

import (
	"encoding/json"
	"fmt"
	"strings"
	"testing"

	werkbook "github.com/jpoz/werkbook"
)

// newFindSession builds a sheet with text, numbers, formulas and a spill:
//
//	   A        B           C
//	1  apple    Apple pie   =A1&"!"
//	2  banana   100         =B2*2
//	3  APPLE    '100        (empty)
//	4  grape    =SEQUENCE(2)
//	5  pineapple  (spill 2)
func newFindSession(t *testing.T) *Session {
	t.Helper()
	s := NewSession()
	mustSet(t, s, sheet1,
		"A1", "apple", "B1", "Apple pie", "C1", `=A1&"!"`,
		"A2", "banana", "B2", "100", "C2", "=B2*2",
		"A3", "APPLE", "B3", "'100",
		"A4", "grape", "B4", "=SEQUENCE(2)",
		"A5", "pineapple",
	)
	return s
}

func at(row, col int) *CellPos { return &CellPos{Row: row, Col: col} }

func found(t *testing.T, res FindResult) string {
	t.Helper()
	if !res.Found {
		return "-"
	}
	ref, err := werkbook.CoordinatesToCellName(res.Col, res.Row)
	if err != nil {
		t.Fatal(err)
	}
	return ref
}

func TestFindOptions(t *testing.T) {
	s := newFindSession(t)
	tests := []struct {
		name      string
		query     string
		opts      FindOptions
		want      string // the found cell, "-" for none
		wantCount int
		wrapped   bool
	}{
		// Defaults: no letter case, part of the text, the editor text.
		{"first match from the start", "apple", FindOptions{}, "A1", 4, false},
		{"next match", "apple", FindOptions{From: at(1, 1)}, "B1", 4, false},
		{"next match on a later row", "apple", FindOptions{From: at(1, 2)}, "A3", 4, false},
		{"last match", "apple", FindOptions{From: at(3, 1)}, "A5", 4, false},
		{"wrap to the first match", "apple", FindOptions{From: at(5, 1)}, "A1", 4, true},
		{"from a cell that is not a match", "apple", FindOptions{From: at(2, 3)}, "A3", 4, false},
		{"from past the data", "apple", FindOptions{From: at(900, 5)}, "A1", 4, true},

		// Backward.
		{"backward from a match", "apple", FindOptions{From: at(3, 1), Backward: true}, "B1", 4, false},
		{"backward wraps to the last match", "apple", FindOptions{From: at(1, 1), Backward: true}, "A5", 4, true},
		{"backward with no start wraps to the last match", "apple", FindOptions{Backward: true}, "A5", 4, true},
		{"backward in one row", "apple", FindOptions{From: at(1, 2), Backward: true}, "A1", 4, false},

		// Options.
		{"match case", "apple", FindOptions{MatchCase: true}, "A1", 2, false},
		{"match case, upper", "APPLE", FindOptions{MatchCase: true}, "A3", 1, false},
		{"whole cell", "apple", FindOptions{WholeCell: true}, "A1", 2, false},
		{"whole cell and match case", "APPLE", FindOptions{WholeCell: true, MatchCase: true}, "A3", 1, false},
		{"whole cell with no match", "appl", FindOptions{WholeCell: true}, "-", 0, false},

		// The editor text: a formula is searched as its text.
		{"formula text", "B2*", FindOptions{}, "C2", 1, false},
		{"formula text with the equals sign", "=a1", FindOptions{}, "C1", 1, false},
		{"function name", "sequence", FindOptions{}, "B4", 1, false},
		{"the result of a formula is not in the editor text", "200", FindOptions{}, "-", 0, false},
		{"the result of a formula is in the display text", "200", FindOptions{LookIn: "display"}, "C2", 1, false},
		{"display text of a formula", "apple!", FindOptions{LookIn: "display"}, "C1", 1, false},
		{"formula text is not in the display text", "B2*", FindOptions{LookIn: "display"}, "-", 0, false},
		{"explicit input", "B2*", FindOptions{LookIn: "input"}, "C2", 1, false},

		// Numbers and text numbers; the apostrophe is not content.
		{"number and text number", "100", FindOptions{}, "B2", 2, false},
		{"whole cell number", "100", FindOptions{WholeCell: true}, "B2", 2, false},
		{"the apostrophe of a text number is not searched", "'100", FindOptions{}, "-", 0, false},
		{"an apostrophe is not found in text cells", "'", FindOptions{}, "-", 0, false},

		// Spill followers can be found: B5 has the value 2.
		{"spill follower", "2", FindOptions{WholeCell: true}, "B5", 1, false},

		// No match.
		{"no match", "zebra", FindOptions{}, "-", 0, false},
		{"no match backward", "zebra", FindOptions{Backward: true}, "-", 0, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			res, err := s.Find(sheet1, tt.query, tt.opts)
			if err != nil {
				t.Fatalf("Find: %v", err)
			}
			if got := found(t, res); got != tt.want || res.Count != tt.wantCount || res.Wrapped != tt.wrapped {
				t.Errorf("Find(%q, %+v) = %s count %d wrapped %v, want %s count %d wrapped %v",
					tt.query, tt.opts, got, res.Count, res.Wrapped, tt.want, tt.wantCount, tt.wrapped)
			}
			if !res.Found && (res.Row != 0 || res.Col != 0) {
				t.Errorf("no match but row %d col %d", res.Row, res.Col)
			}
		})
	}

	// A walk with find-next visits each match one time, then wraps.
	var walk []string
	from := (*CellPos)(nil)
	for i := 0; i < 5; i++ {
		res, err := s.Find(sheet1, "apple", FindOptions{From: from})
		if err != nil {
			t.Fatal(err)
		}
		walk = append(walk, found(t, res))
		from = at(res.Row, res.Col)
	}
	if got := strings.Join(walk, " "); got != "A1 B1 A3 A5 A1" {
		t.Errorf("find-next walk = %s", got)
	}

	// One match on the sheet: find-next from it gives it again, wrapped.
	res, err := s.Find(sheet1, "banana", FindOptions{From: at(2, 1)})
	if err != nil || found(t, res) != "A2" || !res.Wrapped || res.Count != 1 {
		t.Errorf("one match from itself = %+v, %v", res, err)
	}
	// A find makes no undo step.
	if got := s.History().Depth; got != 11 {
		t.Errorf("history depth after finds = %d, want 11", got)
	}
}

func TestFindEdgeCases(t *testing.T) {
	t.Run("empty sheet", func(t *testing.T) {
		s := NewSession()
		res, err := s.Find(sheet1, "x", FindOptions{})
		if err != nil || res.Found || res.Count != 0 {
			t.Errorf("result = %+v, %v", res, err)
		}
	})

	t.Run("errors", func(t *testing.T) {
		s := newFindSession(t)
		bad := []struct {
			sheet, query string
			opts         FindOptions
		}{
			{"Nope", "x", FindOptions{}},
			{sheet1, "", FindOptions{}},
			{sheet1, "x", FindOptions{LookIn: "formula"}},
			{sheet1, "x", FindOptions{From: at(-1, 1)}},
			{sheet1, "x", FindOptions{From: at(1, werkbook.MaxColumns+1)}},
		}
		for _, b := range bad {
			if _, err := s.Find(b.sheet, b.query, b.opts); err == nil {
				t.Errorf("Find(%q, %q, %+v) gave no error", b.sheet, b.query, b.opts)
			}
		}
	})

	t.Run("count limit", func(t *testing.T) {
		s := NewSession()
		edits := make([]Edit, 0, maxFindCount+50)
		for i := 0; i < maxFindCount+50; i++ {
			edits = append(edits, Edit{Row: i/50 + 1, Col: i%50 + 1, Input: "hit"})
		}
		if err := s.SetCells(sheet1, edits); err != nil {
			t.Fatal(err)
		}
		res, err := s.Find(sheet1, "hit", FindOptions{})
		if err != nil || res.Count != maxFindCount || found(t, res) != "A1" {
			t.Errorf("result = %+v, %v; want the count limit %d", res, err, maxFindCount)
		}
	})

	t.Run("merged ranges are searched in their top-left cell", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "x", "B2", "x", "C3", "x")
		// B2 and C3 become covered cells of the merge A1:C3.
		if err := s.file.Sheet(sheet1).MergeCell("A1", "C3"); err != nil {
			t.Fatal(err)
		}
		res, err := s.Find(sheet1, "x", FindOptions{})
		if err != nil || res.Count != 1 || found(t, res) != "A1" {
			t.Errorf("result = %+v, %v", res, err)
		}
		// Replace on a covered cell changes nothing.
		rep, err := s.Replace(sheet1, 2, 2, "x", "y", FindOptions{})
		if err != nil || rep.Replaced {
			t.Errorf("replace on a covered cell = %+v, %v", rep, err)
		}
	})

	t.Run("style-only cells and text that is not ASCII", func(t *testing.T) {
		s := NewSession()
		if err := s.file.Sheet(sheet1).SetStyle("A1", &werkbook.Style{Fill: &werkbook.Fill{Color: "FF0000"}}); err != nil {
			t.Fatal(err)
		}
		mustSet(t, s, sheet1, "B1", "Größe", "B2", "GRÖSSE", "B3", "日本語")
		res, err := s.Find(sheet1, "größe", FindOptions{})
		if err != nil || found(t, res) != "B1" || res.Count != 1 {
			t.Errorf("result = %+v, %v", res, err)
		}
		res, err = s.Find(sheet1, "本", FindOptions{})
		if err != nil || found(t, res) != "B3" {
			t.Errorf("result = %+v, %v", res, err)
		}
	})
}

func TestReplace(t *testing.T) {
	tests := []struct {
		name        string
		cell        string
		query, with string
		opts        FindOptions
		replaced    bool
		wantInput   string
		wantType    string
	}{
		{"text", "A1", "apple", "pear", FindOptions{}, true, "pear", "string"},
		{"part of the text", "B1", "pie", "tart", FindOptions{}, true, "Apple tart", "string"},
		{"no letter case", "A3", "apple", "pear", FindOptions{}, true, "pear", "string"},
		{"each occurrence", "A2", "an", "AN", FindOptions{}, true, "bANANa", "string"},
		{"match case that does not match", "A3", "apple", "pear", FindOptions{MatchCase: true}, false, "APPLE", "string"},
		{"whole cell", "A5", "pineapple", "kiwi", FindOptions{WholeCell: true}, true, "kiwi", "string"},
		{"whole cell that is only a part", "A5", "apple", "kiwi", FindOptions{WholeCell: true}, false, "pineapple", "string"},
		{"cell that does not match", "A4", "apple", "pear", FindOptions{}, false, "grape", "string"},

		// A formula stays a formula.
		// B3 is the text 100; the engine reads it as a number in a product.
		{"in formula text", "C2", "B2", "B3", FindOptions{}, true, "=B3*2", "number"},
		{"operator in a formula", "C2", "*2", "+1", FindOptions{}, true, "=B2+1", "number"},
		{"the result of a formula is not replaced", "C2", "200", "5", FindOptions{}, false, "=B2*2", "number"},
		{"lookIn display is ignored: the editor text is replaced", "C2", "B2", "A2", FindOptions{LookIn: "display"}, true, "=A2*2", "error"},
		{"a replace that makes an invalid formula keeps the text", "C2", "*2", "*(", FindOptions{}, true, "=B2*(", "error"},
		{"a formula that loses its equals sign is text", "C2", "=", "", FindOptions{}, true, "B2*2", "string"},

		// The new text goes through the input rules.
		{"number stays a number", "B2", "100", "250", FindOptions{}, true, "250", "number"},
		{"digit in a number", "B2", "1", "9", FindOptions{}, true, "900", "number"},
		{"number becomes text", "B2", "100", "n/a", FindOptions{}, true, "n/a", "string"},
		{"text becomes a number", "A4", "grape", "42", FindOptions{}, true, "42", "number"},
		{"text becomes a boolean", "A4", "grape", "true", FindOptions{}, true, "TRUE", "bool"},
		{"text becomes empty: the cell is cleared", "A4", "grape", "", FindOptions{}, true, "", ""},
		{"text becomes a formula", "A4", "grape", "=1+1", FindOptions{}, true, "=1+1", "number"},

		// A text number keeps its apostrophe, so it stays text.
		{"text number stays text", "B3", "1", "9", FindOptions{}, true, "'900", "string"},
		{"text number, whole cell", "B3", "100", "007", FindOptions{WholeCell: true}, true, "'007", "string"},

		// Spill cells.
		{"spill follower is read-only", "B5", "2", "9", FindOptions{}, false, "2", "number"},
		{"spill anchor formula", "B4", "2", "3", FindOptions{}, true, "=SEQUENCE(3)", "number"},
		{"empty cell", "E9", "x", "y", FindOptions{}, false, "", ""},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := newFindSession(t)
			before := dump(t, s, sheet1)
			depth := s.History().Depth
			col, row, err := werkbook.CellNameToCoordinates(tt.cell)
			if err != nil {
				t.Fatal(err)
			}
			res, err := s.Replace(sheet1, row, col, tt.query, tt.with, tt.opts)
			if err != nil {
				t.Fatalf("Replace: %v", err)
			}
			if res.Replaced != tt.replaced {
				t.Fatalf("replaced = %v, want %v", res.Replaced, tt.replaced)
			}
			got := cellsByRef(t, s, sheet1, row, col, row, col)[tt.cell]
			if got.Input != tt.wantInput || got.Type != tt.wantType {
				t.Errorf("%s = %q (%s), want %q (%s)", tt.cell, got.Input, got.Type, tt.wantInput, tt.wantType)
			}
			if !tt.replaced {
				if res.History.Depth != depth || dump(t, s, sheet1) != before {
					t.Errorf("a replace with no match changed the workbook or the history")
				}
				return
			}
			if res.History.Depth != depth+1 || res.History.UndoLabel != "Replace" {
				t.Errorf("history = %+v", res.History)
			}
			undo := mustUndo(t, s)
			if *undo.Select != (Range{row, col, row, col}) {
				t.Errorf("undo select = %v", undo.Select)
			}
			if got := dump(t, s, sheet1); got != before {
				t.Errorf("state after undo:\n%s\nwant:\n%s", got, before)
			}
		})
	}

	t.Run("errors", func(t *testing.T) {
		s := newFindSession(t)
		h := s.History()
		if _, err := s.Replace("Nope", 1, 1, "a", "b", FindOptions{}); err == nil {
			t.Error("unknown sheet gave no error")
		}
		if _, err := s.Replace(sheet1, 0, 1, "a", "b", FindOptions{}); err == nil {
			t.Error("row 0 gave no error")
		}
		if _, err := s.Replace(sheet1, 1, 1, "", "b", FindOptions{}); err == nil {
			t.Error("an empty query gave no error")
		}
		if _, err := s.Replace(sheet1, 1, 1, "a", "b", FindOptions{LookIn: "x"}); err == nil {
			t.Error("a bad lookIn gave no error")
		}
		if s.History() != h {
			t.Errorf("history changed")
		}
	})
}

func TestReplaceAll(t *testing.T) {
	s := newFindSession(t)
	before := dump(t, s, sheet1)
	depth := s.History().Depth

	res, err := s.ReplaceAll(sheet1, "apple", "pear", FindOptions{})
	if err != nil {
		t.Fatal(err)
	}
	if res.Replaced != 4 || res.Skipped != 0 {
		t.Errorf("result = %+v, want 4 replaced", res)
	}
	if res.History.Depth != depth+1 || res.History.UndoLabel != "Replace all" {
		t.Errorf("history = %+v", res.History)
	}
	want := map[string]string{"A1": "pear", "B1": "pear pie", "A3": "pear", "A5": "pinepear"}
	got := inputsOf(t, s, sheet1)
	for ref, w := range want {
		if got[ref] != w {
			t.Errorf("%s = %q, want %q", ref, got[ref], w)
		}
	}
	// The formula C1 reads A1: its value follows.
	if c1 := cellsByRef(t, s, sheet1, 1, 3, 1, 3)["C1"]; c1.Display != "pear!" || c1.Input != `=A1&"!"` {
		t.Errorf("C1 = %+v", c1)
	}
	// One undo step for all cells; its range is the bounding box.
	undo := mustUndo(t, s)
	if undo.Label != "Replace all" || *undo.Select != (Range{1, 1, 5, 2}) {
		t.Errorf("undo = %+v %v", undo, undo.Select)
	}
	if got := dump(t, s, sheet1); got != before {
		t.Errorf("state after undo:\n%s\nwant:\n%s", got, before)
	}

	tests := []struct {
		name        string
		query, with string
		opts        FindOptions
		replaced    int
		skipped     int
		check       map[string]string
	}{
		{"match case", "apple", "X", FindOptions{MatchCase: true}, 2, 0, map[string]string{"A1": "X", "A3": "APPLE", "A5": "pineX", "B1": "Apple pie"}},
		{"whole cell", "apple", "X", FindOptions{WholeCell: true}, 2, 0, map[string]string{"A1": "X", "A3": "X", "A5": "pineapple"}},
		{"in formulas", "2", "3", FindOptions{}, 2, 1, map[string]string{"C2": "=B3*3", "B4": "=SEQUENCE(3)"}},
		{"numbers and text numbers keep their types", "100", "55", FindOptions{}, 2, 0, map[string]string{"B2": "55", "B3": "'55"}},
		{"from and backward are ignored", "banana", "kiwi", FindOptions{From: at(9, 9), Backward: true}, 1, 0, map[string]string{"A2": "kiwi"}},
		{"no match", "zebra", "x", FindOptions{}, 0, 0, map[string]string{"A1": "apple"}},
		{"replacement equal to the match changes nothing", "banana", "banana", FindOptions{}, 0, 0, map[string]string{"A2": "banana"}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := newFindSession(t)
			before := dump(t, s, sheet1)
			depth := s.History().Depth
			res, err := s.ReplaceAll(sheet1, tt.query, tt.with, tt.opts)
			if err != nil {
				t.Fatal(err)
			}
			if res.Replaced != tt.replaced || res.Skipped != tt.skipped {
				t.Errorf("result = %+v, want %d replaced and %d skipped", res, tt.replaced, tt.skipped)
			}
			got := inputsOf(t, s, sheet1)
			for ref, w := range tt.check {
				if got[ref] != w {
					t.Errorf("%s = %q, want %q", ref, got[ref], w)
				}
			}
			wantDepth := depth
			if tt.replaced > 0 {
				wantDepth++
			}
			if res.History.Depth != wantDepth {
				t.Errorf("depth = %d, want %d", res.History.Depth, wantDepth)
			}
			if tt.replaced > 0 {
				mustUndo(t, s)
			}
			if got := dump(t, s, sheet1); got != before {
				t.Errorf("state after undo is not the start state")
			}
		})
	}

	t.Run("only spill followers match", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(3)*11")
		res, err := s.ReplaceAll(sheet1, "22", "x", FindOptions{WholeCell: true})
		if err != nil || res.Replaced != 0 || res.Skipped != 1 || res.History.Depth != 1 {
			t.Errorf("result = %+v, %v", res, err)
		}
	})

	t.Run("errors", func(t *testing.T) {
		s := newFindSession(t)
		h := s.History()
		if _, err := s.ReplaceAll("Nope", "a", "b", FindOptions{}); err == nil {
			t.Error("unknown sheet gave no error")
		}
		if _, err := s.ReplaceAll(sheet1, "", "b", FindOptions{}); err == nil {
			t.Error("an empty query gave no error")
		}
		if s.History() != h {
			t.Errorf("history changed")
		}
	})
}

func TestMatcherReplace(t *testing.T) {
	tests := []struct {
		text, query, with string
		opts              FindOptions
		want              string
	}{
		{"aXbxc", "x", "-", FindOptions{}, "a-b-c"},
		{"aXbxc", "x", "-", FindOptions{MatchCase: true}, "aXb-c"},
		{"xxxx", "xx", "y", FindOptions{}, "yy"},
		{"abc", "abc", "", FindOptions{}, ""},
		{"abc", "b", "bb", FindOptions{}, "abbc"},
		{"ABC", "abc", "x", FindOptions{WholeCell: true}, "x"},
		{"ABCD", "abc", "x", FindOptions{WholeCell: true}, "ABCD"},
		{"Größe und GRÖSSE", "größe", "Size", FindOptions{}, "Size und GRÖSSE"},
		{"日本語の日本", "日本", "JP", FindOptions{}, "JP語のJP"},
		// A character whose lower-case form has a different byte length:
		// only exact-case matches are replaced, and no text is damaged.
		{"İstanbul x X", "x", "y", FindOptions{}, "İstanbul y X"},
	}
	for _, tt := range tests {
		m := newMatcher(tt.query, tt.opts)
		if got := m.replace(tt.text, tt.with); got != tt.want {
			t.Errorf("replace(%q, %q -> %q, %+v) = %q, want %q", tt.text, tt.query, tt.with, tt.opts, got, tt.want)
		}
	}
}

func TestFindJSONShapes(t *testing.T) {
	var opts FindOptions
	dec := json.NewDecoder(strings.NewReader(`{"matchCase":true,"wholeCell":true,"lookIn":"display","from":{"row":2,"col":3},"backward":true}`))
	dec.DisallowUnknownFields()
	if err := dec.Decode(&opts); err != nil {
		t.Fatal(err)
	}
	if !opts.MatchCase || !opts.WholeCell || opts.LookIn != "display" || *opts.From != (CellPos{2, 3}) || !opts.Backward {
		t.Errorf("opts = %+v", opts)
	}
	for _, text := range []string{`{"regex":true}`, `{"from":{"row":1,"col":1,"sheet":"x"}}`, `{"matchCase":"yes"}`} {
		dec := json.NewDecoder(strings.NewReader(text))
		dec.DisallowUnknownFields()
		if err := dec.Decode(new(FindOptions)); err == nil {
			t.Errorf("options %s gave no error", text)
		}
	}
	checks := []struct {
		v    any
		want string
	}{
		{FindResult{Found: true, Row: 2, Col: 3, Count: 4, Wrapped: true}, `{"found":true,"row":2,"col":3,"count":4,"wrapped":true}`},
		{FindResult{}, `{"found":false,"row":0,"col":0,"count":0,"wrapped":false}`},
		{ReplaceResult{Replaced: true}, `{"history":{"canUndo":false,"canRedo":false,"undoLabel":"","redoLabel":"","depth":0},"replaced":true}`},
		{ReplaceAllResult{Replaced: 3, Skipped: 1}, `{"history":{"canUndo":false,"canRedo":false,"undoLabel":"","redoLabel":"","depth":0},"replaced":3,"skipped":1}`},
		{SortResult{Changed: 2, Range: Range{1, 1, 2, 2}}, `{"history":{"canUndo":false,"canRedo":false,"undoLabel":"","redoLabel":"","depth":0},"changed":2,"range":{"row1":1,"col1":1,"row2":2,"col2":2}}`},
	}
	for _, c := range checks {
		got, err := json.Marshal(c.v)
		if err != nil || string(got) != c.want {
			t.Errorf("JSON = %s (%v)\nwant   %s", got, err, c.want)
		}
	}
	_ = fmt.Sprint()
}
