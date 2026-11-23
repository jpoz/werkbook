package bridge

import (
	"fmt"
	"strings"
	"testing"

	werkbook "github.com/jpoz/werkbook"
)

func mustCopy(t *testing.T, s *Session, sheet, rg string, cut bool) CopyResult {
	t.Helper()
	res, err := s.Copy(sheet, rangeOf(t, rg), cut)
	if err != nil {
		t.Fatalf("Copy(%s): %v", rg, err)
	}
	return res
}

func mustPaste(t *testing.T, s *Session, sheet, dst, text string) PasteResult {
	t.Helper()
	// No id: the internal clipboard is found by its text.
	res, err := s.Paste(sheet, rangeOf(t, dst), text, nil)
	if err != nil {
		t.Fatalf("Paste(%s): %v", dst, err)
	}
	return res
}

// mustPasteID pastes with the clipboard id that the system clipboard holds
// ("" = the clipboard has no id).
func mustPasteID(t *testing.T, s *Session, sheet, dst, text, id string) PasteResult {
	t.Helper()
	res, err := s.Paste(sheet, rangeOf(t, dst), text, &id)
	if err != nil {
		t.Fatalf("Paste(%s): %v", dst, err)
	}
	return res
}

func TestCopyText(t *testing.T) {
	s := NewSession()
	if err := s.file.Sheet(sheet1).SetStyle("B2", &werkbook.Style{NumFmt: "0.00"}); err != nil {
		t.Fatal(err)
	}
	if err := s.file.Sheet(sheet1).SetStyle("C3", &werkbook.Style{Fill: &werkbook.Fill{Color: "FF0000"}}); err != nil {
		t.Fatal(err)
	}
	mustSet(t, s, sheet1,
		"A1", "1", "B1", "=A1*2", "C1", "tab\there",
		"A2", "say \"hi\"", "B2", "1.5", "C2", "TRUE",
		"A3", "=1/0", "B3", "'007",
	)
	depth := s.History().Depth

	res := mustCopy(t, s, sheet1, "A1:C3", false)
	// Formula cells give their shown value. Constants give their editor
	// text: B2 is 1.5 (not the formatted 1.50) and B3 keeps its apostrophe.
	// The style-only cell C3 is empty.
	want := "1\t2\t\"tab\there\"\n\"say \"\"hi\"\"\"\t1.5\tTRUE\n#DIV/0!\t'007\t"
	if res.Text != want {
		t.Errorf("text = %q\nwant   %q", res.Text, want)
	}
	if res.Cells != 8 || !res.Internal {
		t.Errorf("cells = %d internal = %v, want 8 and true", res.Cells, res.Internal)
	}
	if len(res.ID) < 16 {
		t.Errorf("id = %q, want a random id", res.ID)
	}
	if again := mustCopy(t, s, sheet1, "A1:C3", false); again.ID == res.ID {
		t.Errorf("two copies have the same id %q", res.ID)
	}
	if res.History.Depth != depth {
		t.Errorf("a copy changed the history: %+v", res.History)
	}

	// One cell, one empty cell, and a range with only empty cells.
	if got := mustCopy(t, s, sheet1, "B1", false); got.Text != "2" || got.Cells != 1 {
		t.Errorf("copy of B1 = %+v", got)
	}
	if got := mustCopy(t, s, sheet1, "E5", false); got.Text != "" || got.Cells != 0 || !got.Internal {
		t.Errorf("copy of an empty cell = %+v", got)
	}
	if got := mustCopy(t, s, sheet1, "E5:F6", false); got.Text != "\t\n\t" {
		t.Errorf("copy of an empty range = %q", got.Text)
	}
	// Corners in any order.
	if got, err := s.Copy(sheet1, Range{Row1: 1, Col1: 2, Row2: 1, Col2: 1}, false); err != nil || got.Text != "1\t2" {
		t.Errorf("copy with reversed corners = %+v, %v", got, err)
	}
}

func TestCopyFullRowsAndColumnsUseTheUsedRange(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "a", "B2", "b", "C3", "=SEQUENCE(2)")
	// Column B, full height: limited to rows 1..4 (C4 is a spill follower).
	res, err := s.Copy(sheet1, Range{Row1: 1, Col1: 2, Row2: werkbook.MaxRows, Col2: 2}, false)
	if err != nil {
		t.Fatal(err)
	}
	if res.Text != "\nb\n\n" {
		t.Errorf("column text = %q", res.Text)
	}
	// Row 2, full width: limited to columns A..C.
	res, err = s.Copy(sheet1, Range{Row1: 2, Col1: 1, Row2: 2, Col2: werkbook.MaxColumns}, false)
	if err != nil {
		t.Fatal(err)
	}
	if res.Text != "\tb\t" {
		t.Errorf("row text = %q", res.Text)
	}
	// Select all on an empty sheet gives one empty cell.
	s.New()
	res, err = s.Copy(sheet1, Range{Row1: 1, Col1: 1, Row2: werkbook.MaxRows, Col2: werkbook.MaxColumns}, false)
	if err != nil {
		t.Fatal(err)
	}
	if res.Text != "" || res.Cells != 0 {
		t.Errorf("select-all copy of an empty sheet = %+v", res)
	}
	// A large range that is not full rows or columns is above the text limit.
	if _, err := s.Copy(sheet1, Range{Row1: 1, Col1: 1, Row2: 2000, Col2: 1001}, false); err == nil || !strings.Contains(err.Error(), "2000000") {
		t.Errorf("copy above the text limit: error = %v", err)
	}
}

func TestInternalPasteShiftsFormulasAndCopiesStyles(t *testing.T) {
	s := NewSession()
	sh := s.file.Sheet(sheet1)
	bold := &werkbook.Style{Font: &werkbook.Font{Bold: true}}
	if err := sh.SetStyle("B1", bold); err != nil {
		t.Fatal(err)
	}
	mustSet(t, s, sheet1, "A1", "5", "B1", "=A1*2", "C1", "=$A$1+A$1+$A1", "A2", "text")
	before := dump(t, s, sheet1)

	cp := mustCopy(t, s, sheet1, "A1:C2", false)
	depth := s.History().Depth
	res := mustPaste(t, s, sheet1, "D4", cp.Text)
	if res.Mode != "internal" || res.Range != rangeOf(t, "D4:F5") || res.Changed != 6 || res.Skipped != 0 || res.Clipped {
		t.Errorf("result = %+v", res)
	}
	if res.History.Depth != depth+1 || res.History.UndoLabel != "Paste" {
		t.Errorf("history = %+v", res.History)
	}
	wantInputs(t, s, sheet1, map[string]string{
		"A1": "5", "B1": "=A1*2", "C1": "=$A$1+A$1+$A1", "A2": "text",
		"D4": "5", "E4": "=D4*2", "F4": "=$A$1+D$1+$A4", "D5": "text",
	})
	if st, _ := sh.GetStyle("E4"); st != bold {
		t.Errorf("E4 style = %+v, want the pointer of B1", st)
	}
	if got := cellsByRef(t, s, sheet1, 4, 5, 4, 5)["E4"].Display; got != "10" {
		t.Errorf("E4 = %q, want 10", got)
	}
	checkUndoRedo(t, s, sheet1, before)
	if undo := mustUndo(t, s); *undo.Select != rangeOf(t, "D4:F5") || undo.Label != "Paste" {
		t.Errorf("undo = %+v %v", undo, undo.Select)
	}
}

func TestInternalPasteEdgeCases(t *testing.T) {
	t.Run("reference that leaves the sheet becomes #REF!", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "B2", "=A1+B2")
		cp := mustCopy(t, s, sheet1, "B2", false)
		mustPaste(t, s, sheet1, "A1", cp.Text)
		wantInputs(t, s, sheet1, map[string]string{"B2": "=A1+B2", "A1": "=#REF!+A1"})
	})

	t.Run("line ends of the system clipboard do not change the mode", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=B1", "A2", "=B2")
		cp := mustCopy(t, s, sheet1, "A1:A2", false)
		for _, text := range []string{
			strings.ReplaceAll(cp.Text, "\n", "\r\n"),
			strings.ReplaceAll(cp.Text, "\n", "\r"),
			cp.Text + "\n",
			strings.ReplaceAll(cp.Text, "\n", "\r\n") + "\r\n",
		} {
			if res := mustPaste(t, s, sheet1, "D1", text); res.Mode != "internal" {
				t.Errorf("paste of %q: mode = %q, want internal", text, res.Mode)
			}
		}
		if got := inputsOf(t, s, sheet1)["D2"]; got != "=E2" {
			t.Errorf("D2 = %q, want =E2", got)
		}
	})

	t.Run("text mismatch gives a text paste", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=1+1")
		mustCopy(t, s, sheet1, "A1", false)
		res := mustPaste(t, s, sheet1, "C1", "3")
		if res.Mode != "text" {
			t.Errorf("mode = %q, want text", res.Mode)
		}
		wantInputs(t, s, sheet1, map[string]string{"A1": "=1+1", "C1": "3"})
		// The clipboard text itself is the value 2, so it is internal.
		if res := mustPaste(t, s, sheet1, "D1", "2"); res.Mode != "internal" || inputsOf(t, s, sheet1)["D1"] != "=1+1" {
			t.Errorf("paste of the clipboard text: mode %q, D1 %q", res.Mode, inputsOf(t, s, sheet1)["D1"])
		}
	})

	t.Run("the clipboard is a snapshot from the time of the copy", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=B1+1")
		cp := mustCopy(t, s, sheet1, "A1", false)
		mustSet(t, s, sheet1, "A1", "changed")
		mustPaste(t, s, sheet1, "A3", cp.Text)
		wantInputs(t, s, sheet1, map[string]string{"A1": "changed", "A3": "=B3+1"})
	})

	t.Run("destination that is a whole multiple is filled", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=C1", "B1", "x")
		cp := mustCopy(t, s, sheet1, "A1:B1", false)
		res := mustPaste(t, s, sheet1, "A3:D4", cp.Text)
		if res.Range != rangeOf(t, "A3:D4") || res.Changed != 8 {
			t.Errorf("result = %+v", res)
		}
		wantInputs(t, s, sheet1, map[string]string{
			"A1": "=C1", "B1": "x",
			"A3": "=C3", "B3": "x", "C3": "=E3", "D3": "x",
			"A4": "=C4", "B4": "x", "C4": "=E4", "D4": "x",
		})
	})

	t.Run("destination that is not a multiple gets one copy at its top-left cell", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "1", "B1", "2")
		cp := mustCopy(t, s, sheet1, "A1:B1", false)
		res := mustPaste(t, s, sheet1, "A3:C5", cp.Text)
		if res.Range != rangeOf(t, "A3:B3") || res.Changed != 2 {
			t.Errorf("result = %+v", res)
		}
		wantInputs(t, s, sheet1, map[string]string{"A1": "1", "B1": "2", "A3": "1", "B3": "2"})
	})

	t.Run("one cell fills a range", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=B1")
		cp := mustCopy(t, s, sheet1, "A1", false)
		res := mustPaste(t, s, sheet1, "A3:B4", cp.Text)
		if res.Changed != 4 {
			t.Errorf("result = %+v", res)
		}
		wantInputs(t, s, sheet1, map[string]string{"A1": "=B1", "A3": "=B3", "B3": "=C3", "A4": "=B4", "B4": "=C4"})
	})

	t.Run("empty clipboard cells clear the destination and its style", func(t *testing.T) {
		s := NewSession()
		sh := s.file.Sheet(sheet1)
		fill := &werkbook.Style{Fill: &werkbook.Fill{Color: "00FF00"}}
		old := &werkbook.Style{Font: &werkbook.Font{Italic: true}}
		for ref, st := range map[string]*werkbook.Style{"B1": fill, "A3": old, "B3": old, "C3": old} {
			if err := sh.SetStyle(ref, st); err != nil {
				t.Fatal(err)
			}
		}
		mustSet(t, s, sheet1, "C1", "v", "A3", "old", "B3", "old", "C3", "old")
		before := dump(t, s, sheet1)
		// A1 is absent, B1 is a styled empty cell, C1 has a value.
		cp := mustCopy(t, s, sheet1, "A1:C1", false)
		mustPaste(t, s, sheet1, "A3", cp.Text)
		wantInputs(t, s, sheet1, map[string]string{"C1": "v", "C3": "v"})
		if st, _ := sh.GetStyle("A3"); st != nil {
			t.Errorf("A3 style = %+v, want none (the source cell has none)", st)
		}
		if st, _ := sh.GetStyle("B3"); st != fill {
			t.Errorf("B3 style = %+v, want the fill of B1", st)
		}
		if st, _ := sh.GetStyle("C3"); st != nil {
			t.Errorf("C3 style = %+v, want none", st)
		}
		checkUndoRedo(t, s, sheet1, before)
	})

	t.Run("paste to a different sheet keeps the reference text", func(t *testing.T) {
		s := NewSession()
		if _, err := s.AddSheet("Two"); err != nil {
			t.Fatal(err)
		}
		mustSet(t, s, sheet1, "A1", "10", "B1", "=A1+1", "C1", "=Sheet1!A1+Two!A1")
		mustSet(t, s, "Two", "A1", "100")
		cp := mustCopy(t, s, sheet1, "B1:C1", false)
		res := mustPaste(t, s, "Two", "B1", cp.Text)
		if res.Mode != "internal" {
			t.Fatalf("mode = %q", res.Mode)
		}
		// =A1+1 now reads Two!A1.
		cells := cellsByRef(t, s, "Two", 1, 1, 1, 3)
		if got := cells["B1"]; got.Input != "=A1+1" || got.Display != "101" {
			t.Errorf("Two!B1 = %+v, want =A1+1 with 101", got)
		}
		if got := cells["C1"]; got.Input != "=Sheet1!A1+Two!A1" || got.Display != "110" {
			t.Errorf("Two!C1 = %+v", got)
		}
		// The undo step is on the destination sheet.
		if undo := mustUndo(t, s); undo.Sheet != "Two" {
			t.Errorf("undo sheet = %q, want Two", undo.Sheet)
		}
	})

	t.Run("data that does not fit on the sheet is cut", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "1", "B1", "2", "A2", "3", "B2", "4")
		cp := mustCopy(t, s, sheet1, "A1:B2", false)
		res, err := s.Paste(sheet1, Range{werkbook.MaxRows, werkbook.MaxColumns, werkbook.MaxRows, werkbook.MaxColumns}, cp.Text, nil)
		if err != nil {
			t.Fatal(err)
		}
		last := Range{werkbook.MaxRows, werkbook.MaxColumns, werkbook.MaxRows, werkbook.MaxColumns}
		if !res.Clipped || res.Range != last || res.Changed != 1 {
			t.Errorf("result = %+v, want one cell and clipped", res)
		}
		cells, err := s.GetRange(sheet1, werkbook.MaxRows, werkbook.MaxColumns, werkbook.MaxRows, werkbook.MaxColumns)
		if err != nil || len(cells) != 1 || cells[0].Display != "1" {
			t.Errorf("last cell = %+v, %v", cells, err)
		}
	})
}

func TestCutAndPaste(t *testing.T) {
	s := NewSession()
	sh := s.file.Sheet(sheet1)
	bold := &werkbook.Style{Font: &werkbook.Font{Bold: true}}
	if err := sh.SetStyle("B1", bold); err != nil {
		t.Fatal(err)
	}
	mustSet(t, s, sheet1, "A1", "5", "B1", "=A1*2", "D1", "=B1")
	before := dump(t, s, sheet1)
	depth := s.History().Depth

	cut := mustCopy(t, s, sheet1, "A1:B1", true)
	if cut.Text != "5\t10" || !cut.Internal || cut.Cells != 2 {
		t.Errorf("cut result = %+v", cut)
	}
	if cut.History.Depth != depth+1 || cut.History.UndoLabel != "Cut" {
		t.Errorf("history after cut = %+v", cut.History)
	}
	// The source is cleared at once; the style stays.
	wantInputs(t, s, sheet1, map[string]string{"D1": "=B1"})
	if st, _ := sh.GetStyle("B1"); st != bold {
		t.Errorf("B1 style after cut = %+v, want the style to stay", st)
	}

	// A paste of a cut clipboard does not shift the references.
	res := mustPaste(t, s, sheet1, "A5", cut.Text)
	if res.Mode != "internal" {
		t.Fatalf("mode = %q", res.Mode)
	}
	wantInputs(t, s, sheet1, map[string]string{"D1": "=B1", "A5": "5", "B5": "=A1*2"})
	if st, _ := sh.GetStyle("B5"); st != bold {
		t.Errorf("B5 style = %+v, want the style of the cut cell", st)
	}

	// Two steps: the cut and the paste.
	mustUndo(t, s)
	mustUndo(t, s)
	if got := dump(t, s, sheet1); got != before {
		t.Errorf("state after two undos:\n%s\nwant:\n%s", got, before)
	}

	// A cut of empty cells makes no step.
	h := s.History()
	if got := mustCopy(t, s, sheet1, "F8:G9", true); got.History.Depth != h.Depth {
		t.Errorf("a cut of empty cells made a step: %+v", got.History)
	}
}

func TestPasteSpill(t *testing.T) {
	t.Run("anchor and followers paste only the anchor formula", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(3)*2", "C2", "old", "C3", "old")
		cp := mustCopy(t, s, sheet1, "A1:A3", false)
		if cp.Text != "2\n4\n6" {
			t.Errorf("text = %q", cp.Text)
		}
		res := mustPaste(t, s, sheet1, "C1", cp.Text)
		if res.Changed != 3 {
			t.Errorf("result = %+v", res)
		}
		// The old values in C2:C3 are cleared, so the pasted anchor spills.
		cells := cellsByRef(t, s, sheet1, 1, 3, 3, 3)
		if cells["C1"].Input != "=SEQUENCE(3)*2" || cells["C2"].Display != "4" || !cells["C2"].Spill || !cells["C3"].Spill {
			t.Errorf("pasted spill = %+v", cells)
		}
	})

	t.Run("followers with no anchor paste as values", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(3)*2")
		cp := mustCopy(t, s, sheet1, "A2:A3", false)
		mustPaste(t, s, sheet1, "C1", cp.Text)
		cells := cellsByRef(t, s, sheet1, 1, 3, 2, 3)
		if cells["C1"].Input != "4" || cells["C1"].Formula || cells["C1"].Spill || cells["C2"].Input != "6" {
			t.Errorf("pasted values = %+v", cells)
		}
	})

	t.Run("destination followers are skipped", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(3)", "C1", "x", "C2", "y")
		before := dump(t, s, sheet1)
		cp := mustCopy(t, s, sheet1, "C1:C2", false)
		res := mustPaste(t, s, sheet1, "A2", cp.Text)
		if res.Skipped != 2 || res.Changed != 0 || res.History.Depth != 3 {
			t.Errorf("result = %+v, want 2 skipped and no step", res)
		}
		if got := dump(t, s, sheet1); got != before {
			t.Errorf("the sheet changed:\n%s", got)
		}
		// A text paste has the same rule.
		res = mustPaste(t, s, sheet1, "A3:B3", "p\tq")
		if res.Mode != "text" || res.Skipped != 1 || res.Changed != 1 {
			t.Errorf("text paste result = %+v, want 1 skipped and 1 written", res)
		}
		if got := inputsOf(t, s, sheet1)["B3"]; got != "q" {
			t.Errorf("B3 = %q, want q", got)
		}
	})

	t.Run("followers are written when their anchor is overwritten too", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(3)", "C1", "x", "C2", "y")
		before := dump(t, s, sheet1)
		cp := mustCopy(t, s, sheet1, "C1:C2", false)
		res := mustPaste(t, s, sheet1, "A1", cp.Text)
		if res.Skipped != 0 || res.Changed != 2 {
			t.Errorf("result = %+v", res)
		}
		wantInputs(t, s, sheet1, map[string]string{"A1": "x", "A2": "y", "C1": "x", "C2": "y"})
		checkUndoRedo(t, s, sheet1, before)
	})
}

func TestTextPaste(t *testing.T) {
	s := NewSession()
	sh := s.file.Sheet(sheet1)
	bold := &werkbook.Style{Font: &werkbook.Font{Bold: true}}
	if err := sh.SetStyle("B2", bold); err != nil {
		t.Fatal(err)
	}
	mustSet(t, s, sheet1, "B2", "old", "C3", "old")
	before := dump(t, s, sheet1)

	text := "name\t12.5\tTRUE\r\n=B2&\"!\"\t\t'007\r\n\"multi\nline\"\t=SUM(1,2)\r\n"
	res := mustPaste(t, s, sheet1, "B2", text)
	if res.Mode != "text" || res.Range != rangeOf(t, "B2:D4") || res.Changed != 9 || res.Clipped {
		t.Errorf("result = %+v", res)
	}
	cells := cellsByRef(t, s, sheet1, 1, 1, 6, 6)
	want := map[string][2]string{ // input, type
		"B2": {"name", "string"}, "C2": {"12.5", "number"}, "D2": {"TRUE", "bool"},
		"B3": {`=B2&"!"`, "string"}, "D3": {"'007", "string"},
		"B4": {"multi\nline", "string"}, "C4": {"=SUM(1,2)", "number"},
	}
	for ref, w := range want {
		if got := cells[ref]; got.Input != w[0] || got.Type != w[1] {
			t.Errorf("%s = %q (%s), want %q (%s)", ref, got.Input, got.Type, w[0], w[1])
		}
	}
	// The empty field cleared C3; the short last row cleared D4 (it was absent).
	if _, ok := cells["C3"]; ok {
		t.Errorf("C3 is not cleared: %+v", cells["C3"])
	}
	if len(cells) != len(want) {
		t.Errorf("got %d cells, want %d: %+v", len(cells), len(want), cells)
	}
	// The style of the destination stays.
	if st, _ := sh.GetStyle("B2"); st != bold {
		t.Errorf("B2 style = %+v, want the old style", st)
	}
	checkUndoRedo(t, s, sheet1, before)

	t.Run("one value fills a range", func(t *testing.T) {
		s := NewSession()
		res := mustPaste(t, s, sheet1, "A1:B2", "7")
		if res.Changed != 4 || res.Range != rangeOf(t, "A1:B2") {
			t.Errorf("result = %+v", res)
		}
		wantInputs(t, s, sheet1, map[string]string{"A1": "7", "B1": "7", "A2": "7", "B2": "7"})
	})

	t.Run("a block repeats in a whole multiple", func(t *testing.T) {
		s := NewSession()
		mustPaste(t, s, sheet1, "A1:B4", "a\tb\nc\td")
		wantInputs(t, s, sheet1, map[string]string{
			"A1": "a", "B1": "b", "A2": "c", "B2": "d",
			"A3": "a", "B3": "b", "A4": "c", "B4": "d",
		})
	})

	t.Run("empty text is nothing to paste", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "x", "B1", "y")
		for _, text := range []string{"", "\n", "\r\n"} {
			res := mustPaste(t, s, sheet1, "A1:B1", text)
			if res.Mode != "text" || res.Changed != 0 || res.Range != rangeOf(t, "A1") || res.History.Depth != 2 {
				t.Errorf("paste of %q: result = %+v, want no change and no step", text, res)
			}
		}
		// The same with an id that says "the clipboard has no id".
		if res := mustPasteID(t, s, sheet1, "A1", "", ""); res.Changed != 0 {
			t.Errorf("result = %+v", res)
		}
		wantInputs(t, s, sheet1, map[string]string{"A1": "x", "B1": "y"})
		// A text with an empty field next to data still clears that cell.
		mustPaste(t, s, sheet1, "A1", "\tz")
		wantInputs(t, s, sheet1, map[string]string{"B1": "z"})
	})

	t.Run("a paste that changes nothing makes no step", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "x")
		res := mustPaste(t, s, sheet1, "A1", "x")
		if res.History.Depth != 1 {
			t.Errorf("history = %+v, want depth 1", res.History)
		}
	})

	t.Run("data that does not fit is cut", func(t *testing.T) {
		s := NewSession()
		res, err := s.Paste(sheet1, Range{1, werkbook.MaxColumns, 1, werkbook.MaxColumns}, "a\tb\tc\nd\te\tf", nil)
		if err != nil {
			t.Fatal(err)
		}
		if !res.Clipped || res.Changed != 2 || res.Range != (Range{1, werkbook.MaxColumns, 2, werkbook.MaxColumns}) {
			t.Errorf("result = %+v", res)
		}
	})
}

func TestPasteMergedCells(t *testing.T) {
	newSession := func(t *testing.T) *Session {
		s := NewSession()
		if err := s.file.Sheet(sheet1).MergeCell("B2", "C3"); err != nil {
			t.Fatal(err)
		}
		mustSet(t, s, sheet1, "A1", "1", "E1", "=A1", "F1", "x")
		return s
	}

	t.Run("text paste writes the top-left cell of a merge and skips the rest", func(t *testing.T) {
		s := newSession(t)
		res := mustPaste(t, s, sheet1, "A2", "a\tb\tc\td\ne\tf\tg\th")
		// Row 2: A2, B2 (top-left), C2 (covered), D2. Row 3: A3, B3 and C3 (covered), D3.
		if res.Changed != 5 || res.Skipped != 0 {
			t.Errorf("result = %+v, want 5 cells written", res)
		}
		wantInputs(t, s, sheet1, map[string]string{
			"A1": "1", "E1": "=A1", "F1": "x",
			"A2": "a", "B2": "b", "D2": "d", "A3": "e", "D3": "h",
		})
	})

	t.Run("a destination with a merge gets the values through the text path", func(t *testing.T) {
		s := newSession(t)
		cp := mustCopy(t, s, sheet1, "E1:F1", false) // "=A1" (shown 1) and "x"
		before := dump(t, s, sheet1)
		// The paste at A2:B2 touches the merge B2:C3.
		res := mustPasteID(t, s, sheet1, "A2", cp.Text, cp.ID)
		if res.Mode != "text" || res.Changed != 2 || res.Range != rangeOf(t, "A2:B2") {
			t.Errorf("result = %+v, want a text paste of 2 cells", res)
		}
		// Values, not formulas: A2 gets the shown value of the formula.
		wantInputs(t, s, sheet1, map[string]string{"A1": "1", "E1": "=A1", "F1": "x", "A2": "1", "B2": "x"})
		checkUndoRedo(t, s, sheet1, before)
		// A destination that does not touch the merge is an internal paste.
		if res := mustPasteID(t, s, sheet1, "E5", cp.Text, cp.ID); res.Mode != "internal" {
			t.Errorf("mode = %q", res.Mode)
		}
		if got := inputsOf(t, s, sheet1)["E5"]; got != "=A5" {
			t.Errorf("E5 = %q, want the shifted formula =A5", got)
		}
	})

	t.Run("a copied range with a merge is pasted as values", func(t *testing.T) {
		// The case of the review: merge C1:D2 with text in C1.
		s := NewSession()
		if err := s.file.Sheet(sheet1).MergeCell("C1", "D2"); err != nil {
			t.Fatal(err)
		}
		mustSet(t, s, sheet1, "C1", "merged", "E1", "=1+1", "E2", "5")
		cp := mustCopy(t, s, sheet1, "C1:E2", false)
		if !cp.Internal || cp.Text != "merged\t\t2\n\t\t5" {
			t.Fatalf("copy = %+v", cp)
		}
		for _, id := range []*string{nil, &cp.ID} {
			res, err := s.Paste(sheet1, rangeOf(t, "J10"), cp.Text, id)
			if err != nil {
				t.Fatalf("Paste: %v", err)
			}
			if res.Mode != "text" || res.Range != rangeOf(t, "J10:L11") {
				t.Errorf("result = %+v, want a text paste to J10:L11", res)
			}
		}
		got := inputsOf(t, s, sheet1)
		if got["J10"] != "merged" || got["L10"] != "2" || got["L11"] != "5" {
			t.Errorf("pasted cells = %v", got)
		}
		// The paste made no merge at the destination.
		if n := len(mustMeta(t, s, sheet1).Merges); n != 1 {
			t.Errorf("the sheet has %d merges, want 1", n)
		}
	})
}

// A copy of formatted constants must come back with the same types and
// values through the text path (a paste in a second tab, or a paste after
// the internal clipboard was replaced).
func TestCopyTextRoundTripOfFormattedConstants(t *testing.T) {
	s := NewSession()
	sh := s.file.Sheet(sheet1)
	styles := map[string]*werkbook.Style{
		"A1": {NumFmt: "0%"},
		"A2": {NumFmt: "#,##0.00"},
		"A3": {NumFmtID: 14},
		"A5": {NumFmt: "0.00"},
	}
	for ref, st := range styles {
		if err := sh.SetStyle(ref, st); err != nil {
			t.Fatal(err)
		}
	}
	mustSet(t, s, sheet1, "A1", "0.5", "A2", "1234.5", "A3", "45000", "A4", "'007", "A5", "0.1234567",
		"A6", "TRUE", "A7", "'=text", "A8", "'TRUE", "A9", "plain", "A10", "=A2*2")
	shown := cellsByRef(t, s, sheet1, 1, 1, 10, 1)
	if shown["A1"].Display != "50%" || shown["A2"].Display != "1,234.50" || shown["A3"].Display != "3/15/2023" {
		t.Fatalf("test setup: displays = %q %q %q", shown["A1"].Display, shown["A2"].Display, shown["A3"].Display)
	}

	cp := mustCopy(t, s, sheet1, "A1:A10", false)
	want := "0.5\n1234.5\n45000\n'007\n0.1234567\nTRUE\n'=text\n'TRUE\nplain\n2469"
	if cp.Text != want {
		t.Errorf("text = %q\nwant   %q", cp.Text, want)
	}

	// The same session, but the clipboard has no id: a text paste.
	res := mustPasteID(t, s, sheet1, "C1", cp.Text, "")
	if res.Mode != "text" {
		t.Fatalf("mode = %q, want text", res.Mode)
	}
	// A second session (a second browser tab).
	s2 := NewSession()
	if res := mustPasteID(t, s2, sheet1, "C1", cp.Text, cp.ID); res.Mode != "text" {
		t.Fatalf("mode in a second session = %q, want text", res.Mode)
	}

	for name, session := range map[string]*Session{"same session": s, "second session": s2} {
		pasted := cellsByRef(t, session, sheet1, 1, 3, 10, 3)
		for row := 1; row <= 9; row++ {
			src := shown[fmt.Sprintf("A%d", row)]
			got := pasted[fmt.Sprintf("C%d", row)]
			if got.Type != src.Type || got.Input != src.Input {
				t.Errorf("%s, row %d: pasted %q (%s), want %q (%s)", name, row, got.Input, got.Type, src.Input, src.Type)
			}
		}
		// The formula cell gives its value.
		if got := pasted["C10"]; got.Input != "2469" || got.Type != "number" || got.Formula {
			t.Errorf("%s: pasted formula cell = %+v, want the number 2469", name, got)
		}
	}
}

// The id of a copy decides if a paste uses the internal clipboard. The text
// compare is only the fallback for a caller that cannot read the id.
func TestPasteClipboardID(t *testing.T) {
	setup := func(t *testing.T) (*Session, CopyResult) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=B1*2", "B1", "2")
		cp := mustCopy(t, s, sheet1, "A1", false)
		if cp.Text != "4" {
			t.Fatalf("copy text = %q", cp.Text)
		}
		return s, cp
	}

	t.Run("equal id: internal paste", func(t *testing.T) {
		s, cp := setup(t)
		res := mustPasteID(t, s, sheet1, "D5", cp.Text, cp.ID)
		if res.Mode != "internal" || inputsOf(t, s, sheet1)["D5"] != "=E5*2" {
			t.Errorf("result = %+v, D5 = %q", res, inputsOf(t, s, sheet1)["D5"])
		}
	})
	t.Run("equal id with a text that the system changed: still internal", func(t *testing.T) {
		s, cp := setup(t)
		if res := mustPasteID(t, s, sheet1, "D5", " 4 \r\n", cp.ID); res.Mode != "internal" {
			t.Errorf("mode = %q", res.Mode)
		}
	})
	t.Run("no id on the clipboard: the same text from a different program is a text paste", func(t *testing.T) {
		// The case of the review: the user copied the text "4" in a
		// different program.
		s, _ := setup(t)
		res := mustPasteID(t, s, sheet1, "D5", "4", "")
		got := cellsByRef(t, s, sheet1, 5, 4, 5, 4)["D5"]
		if res.Mode != "text" || got.Input != "4" || got.Formula {
			t.Errorf("result = %+v, D5 = %+v; want the number 4", res, got)
		}
	})
	t.Run("different id: text paste", func(t *testing.T) {
		s, cp := setup(t)
		if res := mustPasteID(t, s, sheet1, "D5", cp.Text, "0123456789abcdef"); res.Mode != "text" {
			t.Errorf("mode = %q", res.Mode)
		}
	})
	t.Run("id of an older copy: text paste", func(t *testing.T) {
		s, first := setup(t)
		mustCopy(t, s, sheet1, "B1", false)
		if res := mustPasteID(t, s, sheet1, "D5", first.Text, first.ID); res.Mode != "text" {
			t.Errorf("mode = %q", res.Mode)
		}
	})
	t.Run("caller that cannot read the id: text compare", func(t *testing.T) {
		s, cp := setup(t)
		if res := mustPaste(t, s, sheet1, "D5", cp.Text); res.Mode != "internal" {
			t.Errorf("mode = %q, want internal", res.Mode)
		}
		if res := mustPaste(t, s, sheet1, "D6", "5"); res.Mode != "text" {
			t.Errorf("mode = %q, want text", res.Mode)
		}
	})
	t.Run("copy with no internal clipboard has no id", func(t *testing.T) {
		s := NewSession()
		// 200,001 cells with a value are above the snapshot limit.
		edits := make([]Edit, 0, maxOpCells)
		for i := 0; i < maxOpCells; i++ {
			edits = append(edits, Edit{Row: i/200 + 1, Col: i%200 + 1, Input: "1"})
		}
		if err := s.SetCells(sheet1, edits); err != nil {
			t.Fatal(err)
		}
		mustSet(t, s, sheet1, "A1001", "1")
		cp, err := s.Copy(sheet1, Range{1, 1, 1001, 200}, false)
		if err != nil {
			t.Fatal(err)
		}
		if cp.Internal || cp.ID != "" || cp.Cells != maxOpCells+1 {
			t.Errorf("copy = internal %v id %q cells %d", cp.Internal, cp.ID, cp.Cells)
		}
	})
}

func TestCutAndPasteOnADifferentSheet(t *testing.T) {
	// The case of the review.
	s := NewSession()
	if _, err := s.AddSheet("Two"); err != nil {
		t.Fatal(err)
	}
	if _, err := s.AddSheet("My Data"); err != nil {
		t.Fatal(err)
	}
	mustSet(t, s, sheet1, "A1", "10", "B1", "=A1*2", "B2", "=SUM(A1:A3)+Two!A1+$A$1", "B3", `="A1"&A1`)
	mustSet(t, s, "Two", "A1", "7")
	mustSet(t, s, "My Data", "A1", "100", "C1", "=A1+1")

	cut := mustCopy(t, s, sheet1, "B1:B3", true)
	res := mustPasteID(t, s, "Two", "B1", cut.Text, cut.ID)
	if res.Mode != "internal" {
		t.Fatalf("mode = %q", res.Mode)
	}
	cells := cellsByRef(t, s, "Two", 1, 2, 3, 2)
	// A moved formula reads the same cells as before: references with no
	// sheet name get the source sheet; qualified references and text stay.
	want := map[string][2]string{
		"B1": {"=Sheet1!A1*2", "20"},
		"B2": {"=SUM(Sheet1!A1:A3)+Two!A1+Sheet1!$A$1", "27"},
		"B3": {`="A1"&Sheet1!A1`, "A110"},
	}
	for ref, w := range want {
		if got := cells[ref]; got.Input != w[0] || got.Display != w[1] {
			t.Errorf("Two!%s = %q (%s), want %q (%s)", ref, got.Input, got.Display, w[0], w[1])
		}
	}

	// A source sheet name that needs quotes.
	cut = mustCopy(t, s, "My Data", "C1", true)
	mustPasteID(t, s, sheet1, "D4", cut.Text, cut.ID)
	if got := cellsByRef(t, s, sheet1, 4, 4, 4, 4)["D4"]; got.Input != "='My Data'!A1+1" || got.Display != "101" {
		t.Errorf("Sheet1!D4 = %q (%s), want ='My Data'!A1+1 (101)", got.Input, got.Display)
	}

	// On the source sheet a cut keeps the text as it is.
	mustSet(t, s, sheet1, "A5", "=A1+1")
	cut = mustCopy(t, s, sheet1, "A5", true)
	mustPasteID(t, s, sheet1, "F9", cut.Text, cut.ID)
	if got := inputsOf(t, s, sheet1)["F9"]; got != "=A1+1" {
		t.Errorf("Sheet1!F9 = %q, want =A1+1", got)
	}

	// A copy (not a cut) to a different sheet reads the destination sheet.
	cp := mustCopy(t, s, sheet1, "F9", false)
	mustPasteID(t, s, "Two", "F9", cp.Text, cp.ID)
	if got := cellsByRef(t, s, "Two", 9, 6, 9, 6)["F9"]; got.Input != "=A1+1" || got.Display != "8" {
		t.Errorf("Two!F9 = %q (%s), want =A1+1 (8)", got.Input, got.Display)
	}

	// The source sheet is renamed between the cut and the paste: the
	// reference gets the name that the sheet has now.
	mustSet(t, s, sheet1, "A6", "=A1")
	cut = mustCopy(t, s, sheet1, "A6", true)
	if _, err := s.RenameSheet(sheet1, "Input"); err != nil {
		t.Fatal(err)
	}
	mustPasteID(t, s, "Two", "H1", cut.Text, cut.ID)
	if got := cellsByRef(t, s, "Two", 1, 8, 1, 8)["H1"]; got.Input != "=Input!A1" || got.Display != "10" {
		t.Errorf("Two!H1 = %q (%s), want =Input!A1 (10)", got.Input, got.Display)
	}
}

func TestUndoOfCutMakesThePasteACopy(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "3", "B1", "=A1*2")
	cut := mustCopy(t, s, sheet1, "B1", true)

	// Undo the cut: B1 is back, so the clipboard is a copy of it.
	if undo := mustUndo(t, s); undo.Label != "Cut" {
		t.Fatalf("undo label = %q", undo.Label)
	}
	mustPasteID(t, s, sheet1, "D5", cut.Text, cut.ID)
	wantInputs(t, s, sheet1, map[string]string{"A1": "3", "B1": "=A1*2", "D5": "=C5*2"})

	// Undo then redo of the cut: the clipboard is a cut again, and a paste
	// does not shift.
	s = NewSession()
	mustSet(t, s, sheet1, "A1", "3", "B1", "=A1*2")
	cut = mustCopy(t, s, sheet1, "B1", true)
	mustUndo(t, s)
	if redo := mustRedo(t, s); redo.Label != "Cut" {
		t.Fatalf("redo label = %q", redo.Label)
	}
	mustPasteID(t, s, sheet1, "D5", cut.Text, cut.ID)
	wantInputs(t, s, sheet1, map[string]string{"A1": "3", "D5": "=A1*2"})

	// A newer copy is not changed by an undo or a redo of the older cut.
	s = NewSession()
	mustSet(t, s, sheet1, "A1", "3", "B1", "=A1*2")
	mustCopy(t, s, sheet1, "B1", true)
	cp := mustCopy(t, s, sheet1, "A1", false)
	mustUndo(t, s) // the cut
	if s.clip == nil || s.clip.id != cp.ID || s.clip.cut {
		t.Errorf("the newer clipboard changed after the undo: %+v", s.clip)
	}
	mustRedo(t, s)
	if s.clip == nil || s.clip.id != cp.ID || s.clip.cut {
		t.Errorf("the newer clipboard changed after the redo: %+v", s.clip)
	}

	// A cut of empty cells makes no step and no cut mark in the history.
	s = NewSession()
	mustCopy(t, s, sheet1, "A1:B2", true)
	if got := s.History(); got != (HistoryState{}) {
		t.Errorf("history after a cut of empty cells = %+v", got)
	}
}

func TestClipboardLifetime(t *testing.T) {
	setup := func(t *testing.T) (*Session, string) {
		s := NewSession()
		if _, err := s.AddSheet("Two"); err != nil {
			t.Fatal(err)
		}
		mustSet(t, s, "Two", "A1", "=B1+1")
		return s, mustCopy(t, s, "Two", "A1", false).Text
	}
	pasteMode := func(t *testing.T, s *Session, text string) string {
		t.Helper()
		return mustPaste(t, s, s.Sheets()[0].Name, "D4", text).Mode
	}

	t.Run("stays after writes, undo and a rename", func(t *testing.T) {
		s, text := setup(t)
		mustSet(t, s, sheet1, "A1", "x")
		mustUndo(t, s)
		if _, err := s.RenameSheet("Two", "Zwei"); err != nil {
			t.Fatal(err)
		}
		if got := pasteMode(t, s, text); got != "internal" {
			t.Errorf("mode = %q, want internal", got)
		}
	})
	t.Run("cleared by new", func(t *testing.T) {
		s, text := setup(t)
		s.New()
		if got := pasteMode(t, s, text); got != "text" {
			t.Errorf("mode = %q, want text", got)
		}
	})
	t.Run("cleared by open", func(t *testing.T) {
		s, text := setup(t)
		data, err := s.Save()
		if err != nil {
			t.Fatal(err)
		}
		if _, err := s.Open(data); err != nil {
			t.Fatal(err)
		}
		if got := pasteMode(t, s, text); got != "text" {
			t.Errorf("mode = %q, want text", got)
		}
	})
	t.Run("cleared by delete of the source sheet", func(t *testing.T) {
		s, text := setup(t)
		if _, err := s.DeleteSheet("Two"); err != nil {
			t.Fatal(err)
		}
		if got := pasteMode(t, s, text); got != "text" {
			t.Errorf("mode = %q, want text", got)
		}
	})
	t.Run("stays after delete of a different sheet", func(t *testing.T) {
		s, text := setup(t)
		if _, err := s.AddSheet("Three"); err != nil {
			t.Fatal(err)
		}
		if _, err := s.DeleteSheet("Three"); err != nil {
			t.Fatal(err)
		}
		if got := pasteMode(t, s, text); got != "internal" {
			t.Errorf("mode = %q, want internal", got)
		}
	})
	t.Run("replaced by the next copy", func(t *testing.T) {
		s, text := setup(t)
		mustSet(t, s, sheet1, "A1", "other")
		mustCopy(t, s, sheet1, "A1", false)
		if got := pasteMode(t, s, text); got != "text" {
			t.Errorf("mode = %q, want text (the first clipboard is gone)", got)
		}
	})
}

func TestCopyAndPasteErrors(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "1")
	h := s.History()
	if _, err := s.Copy("Nope", Range{1, 1, 1, 1}, false); err == nil {
		t.Error("Copy on an unknown sheet gave no error")
	}
	if _, err := s.Copy(sheet1, Range{0, 1, 1, 1}, true); err == nil {
		t.Error("Copy with a bad range gave no error")
	}
	if _, err := s.Paste("Nope", Range{1, 1, 1, 1}, "x", nil); err == nil {
		t.Error("Paste on an unknown sheet gave no error")
	}
	if _, err := s.Paste(sheet1, Range{1, 1, 1, werkbook.MaxColumns + 1}, "x", nil); err == nil {
		t.Error("Paste with a bad range gave no error")
	}
	// One value into a range above the cell limit.
	if _, err := s.Paste(sheet1, Range{1, 1, 1001, 200}, "x", nil); err == nil || !strings.Contains(err.Error(), "200000") {
		t.Errorf("Paste above the cell limit: error = %v", err)
	}
	if s.History() != h {
		t.Errorf("history changed: %+v", s.History())
	}
	wantInputs(t, s, sheet1, map[string]string{"A1": "1"})
}
