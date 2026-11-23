package bridge

import (
	"encoding/json"
	"strings"
	"testing"

	werkbook "github.com/jpoz/werkbook"
)

const sheet1 = "Sheet1"

// mustSet writes cells in order. Each pair is an A1 reference and input text.
func mustSet(t *testing.T, s *Session, sheet string, pairs ...string) {
	t.Helper()
	if len(pairs)%2 != 0 {
		t.Fatalf("mustSet: odd number of arguments")
	}
	for i := 0; i < len(pairs); i += 2 {
		col, row, err := werkbook.CellNameToCoordinates(pairs[i])
		if err != nil {
			t.Fatalf("bad ref %q: %v", pairs[i], err)
		}
		if err := s.SetCell(sheet, row, col, pairs[i+1]); err != nil {
			t.Fatalf("SetCell(%s, %q): %v", pairs[i], pairs[i+1], err)
		}
	}
}

// cellsByRef reads a range and indexes the result by A1 reference.
func cellsByRef(t *testing.T, s *Session, sheet string, row1, col1, row2, col2 int) map[string]Cell {
	t.Helper()
	cells, err := s.GetRange(sheet, row1, col1, row2, col2)
	if err != nil {
		t.Fatalf("GetRange: %v", err)
	}
	out := make(map[string]Cell, len(cells))
	for _, c := range cells {
		ref, err := werkbook.CoordinatesToCellName(c.Col, c.Row)
		if err != nil {
			t.Fatalf("bad cell coordinates %d,%d: %v", c.Row, c.Col, err)
		}
		if _, dup := out[ref]; dup {
			t.Fatalf("cell %s is in the result two times", ref)
		}
		out[ref] = c
	}
	return out
}

func TestNewSessionIsEmptyWorkbook(t *testing.T) {
	s := NewSession()

	sheets := s.Sheets()
	want := []SheetInfo{{Name: sheet1, MaxRow: 0, MaxCol: 0, Visible: true}}
	if len(sheets) != 1 || sheets[0] != want[0] {
		t.Fatalf("Sheets() = %+v, want %+v", sheets, want)
	}

	cells, err := s.GetRange(sheet1, 1, 1, 100, 26)
	if err != nil {
		t.Fatalf("GetRange: %v", err)
	}
	if cells == nil || len(cells) != 0 {
		t.Fatalf("GetRange on an empty sheet = %#v, want a non-nil empty slice", cells)
	}

	meta, err := s.SheetMeta(sheet1)
	if err != nil {
		t.Fatalf("SheetMeta: %v", err)
	}
	if meta.ColWidths == nil || meta.RowHeights == nil || meta.Merges == nil {
		t.Fatalf("SheetMeta has nil fields: %#v", meta)
	}
	if len(meta.ColWidths)+len(meta.RowHeights)+len(meta.Merges) != 0 {
		t.Fatalf("SheetMeta of an empty sheet = %+v, want no entries", meta)
	}

	// An empty workbook must save and open again.
	data, err := s.Save()
	if err != nil {
		t.Fatalf("Save: %v", err)
	}
	if _, err := NewSession().Open(data); err != nil {
		t.Fatalf("Open of a saved empty workbook: %v", err)
	}
}

func TestParseInput(t *testing.T) {
	tests := []struct {
		input string
		want  parsedInput
	}{
		{"", parsedInput{kind: inputClear}},
		{"=SUM(A1:A3)", parsedInput{kind: inputFormula, formula: "=SUM(A1:A3)"}},
		{"=", parsedInput{kind: inputString, str: "="}},
		{"TRUE", parsedInput{kind: inputBool, b: true}},
		{"true", parsedInput{kind: inputBool, b: true}},
		{"False", parsedInput{kind: inputBool, b: false}},
		{"12", parsedInput{kind: inputNumber, num: 12}},
		{"-3.5", parsedInput{kind: inputNumber, num: -3.5}},
		{".5", parsedInput{kind: inputNumber, num: 0.5}},
		{"1e3", parsedInput{kind: inputNumber, num: 1000}},
		{" 42 ", parsedInput{kind: inputNumber, num: 42}},
		{"'123", parsedInput{kind: inputString, str: "123"}},
		{"'=A1", parsedInput{kind: inputString, str: "=A1"}},
		{"'", parsedInput{kind: inputString, str: ""}},
		{"hello", parsedInput{kind: inputString, str: "hello"}},
		{" padded ", parsedInput{kind: inputString, str: " padded "}},
		// Text that strconv.ParseFloat accepts but a spreadsheet does not.
		{"inf", parsedInput{kind: inputString, str: "inf"}},
		{"NaN", parsedInput{kind: inputString, str: "NaN"}},
		{"0x10", parsedInput{kind: inputString, str: "0x10"}},
		{"1_000", parsedInput{kind: inputString, str: "1_000"}},
		{"1e999", parsedInput{kind: inputString, str: "1e999"}},
		{"-", parsedInput{kind: inputString, str: "-"}},
		{"1-2", parsedInput{kind: inputString, str: "1-2"}},
	}
	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			if got := parseInput(tt.input); got != tt.want {
				t.Fatalf("parseInput(%q) = %+v, want %+v", tt.input, got, tt.want)
			}
		})
	}
}

func TestSetCellInputRules(t *testing.T) {
	tests := []struct {
		name        string
		input       string
		wantPresent bool
		wantType    string
		wantDisplay string
		wantInput   string
		wantFormula bool
	}{
		{name: "number", input: "12.5", wantPresent: true, wantType: "number", wantDisplay: "12.5", wantInput: "12.5"},
		{name: "integer", input: "7", wantPresent: true, wantType: "number", wantDisplay: "7", wantInput: "7"},
		{name: "bool", input: "true", wantPresent: true, wantType: "bool", wantDisplay: "TRUE", wantInput: "TRUE"},
		{name: "string", input: "hello", wantPresent: true, wantType: "string", wantDisplay: "hello", wantInput: "hello"},
		{name: "formula", input: "=1+2", wantPresent: true, wantType: "number", wantDisplay: "3", wantInput: "=1+2", wantFormula: true},
		{name: "lone equals sign", input: "=", wantPresent: true, wantType: "string", wantDisplay: "=", wantInput: "="},
		// The editor text keeps the apostrophe, so a second commit of the
		// same text does not change the type.
		{name: "forced numeric string", input: "'123", wantPresent: true, wantType: "string", wantDisplay: "123", wantInput: "'123"},
		{name: "forced formula string", input: "'=A1", wantPresent: true, wantType: "string", wantDisplay: "=A1", wantInput: "'=A1"},
		{name: "forced bool string", input: "'TRUE", wantPresent: true, wantType: "string", wantDisplay: "TRUE", wantInput: "'TRUE"},
		{name: "clear", input: "", wantPresent: false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := NewSession()
			mustSet(t, s, sheet1, "B2", "before", "B2", tt.input)
			got, ok := cellsByRef(t, s, sheet1, 1, 1, 10, 10)["B2"]
			if ok != tt.wantPresent {
				t.Fatalf("cell present = %v, want %v (cell %+v)", ok, tt.wantPresent, got)
			}
			if !ok {
				return
			}
			if got.Row != 2 || got.Col != 2 {
				t.Errorf("coordinates = %d,%d, want 2,2", got.Row, got.Col)
			}
			if got.Type != tt.wantType || got.Display != tt.wantDisplay || got.Input != tt.wantInput || got.Formula != tt.wantFormula {
				t.Errorf("cell = %+v, want type %q display %q input %q formula %v",
					got, tt.wantType, tt.wantDisplay, tt.wantInput, tt.wantFormula)
			}
			if got.Spill {
				t.Errorf("spill = true for a cell that is not a spill follower")
			}

			// The editor text must give the same cell when it is written again.
			mustSet(t, s, sheet1, "B2", got.Input)
			again := cellsByRef(t, s, sheet1, 1, 1, 10, 10)["B2"]
			if again != got {
				t.Errorf("second commit of %q changed the cell: %+v -> %+v", got.Input, got, again)
			}
		})
	}
}

func TestFormulaRecalculatesAfterEdit(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "1", "A2", "2", "A3", "=SUM(A1:A2)", "B3", "=A3*2")

	cells := cellsByRef(t, s, sheet1, 1, 1, 5, 5)
	if got := cells["A3"]; got.Display != "3" || !got.Formula || got.Input != "=SUM(A1:A2)" || got.Type != "number" {
		t.Fatalf("A3 = %+v, want formula cell with display 3", got)
	}
	if got := cells["B3"].Display; got != "6" {
		t.Fatalf("B3 display = %q, want 6", got)
	}

	// An edit of a precedent changes the direct and the indirect dependent.
	mustSet(t, s, sheet1, "A1", "10")
	cells = cellsByRef(t, s, sheet1, 1, 1, 5, 5)
	if got := cells["A3"].Display; got != "12" {
		t.Errorf("A3 display after edit = %q, want 12", got)
	}
	if got := cells["B3"].Display; got != "24" {
		t.Errorf("B3 display after edit = %q, want 24", got)
	}

	// A cleared precedent counts as zero.
	mustSet(t, s, sheet1, "A2", "")
	cells = cellsByRef(t, s, sheet1, 1, 1, 5, 5)
	if _, ok := cells["A2"]; ok {
		t.Errorf("cleared cell A2 is in the result")
	}
	if got := cells["A3"].Display; got != "10" {
		t.Errorf("A3 display after clear = %q, want 10", got)
	}

	// A formula replaced by a value is not a formula cell any more.
	mustSet(t, s, sheet1, "A3", "5")
	cells = cellsByRef(t, s, sheet1, 1, 1, 5, 5)
	if got := cells["A3"]; got.Formula || got.Input != "5" {
		t.Errorf("A3 = %+v, want a plain number", got)
	}
	if got := cells["B3"].Display; got != "10" {
		t.Errorf("B3 display = %q, want 10", got)
	}
}

func TestErrorValues(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		want    string
	}{
		{"division by zero", "=1/0", "#DIV/0!"},
		{"unknown function", "=NOSUCHFUNCTION(1)", "#NAME?"},
		{"not available", "=NA()", "#N/A"},
		{"bad operand", `="a"+1`, "#VALUE!"},
		{"syntax error", "=SUM(", "#NAME?"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := NewSession()
			mustSet(t, s, sheet1, "A1", tt.formula, "B1", "=A1+1")
			cells := cellsByRef(t, s, sheet1, 1, 1, 3, 3)
			got := cells["A1"]
			if got.Type != "error" || got.Display != tt.want || !got.Formula || got.Input != tt.formula {
				t.Fatalf("A1 = %+v, want error %s with input %q", got, tt.want, tt.formula)
			}
			// The error goes through to a dependent.
			if dep := cells["B1"]; dep.Type != "error" || dep.Display != tt.want {
				t.Errorf("B1 = %+v, want error %s", dep, tt.want)
			}
		})
	}
}

func TestSpillCells(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "B2", "=SEQUENCE(3,2)")

	cells := cellsByRef(t, s, sheet1, 1, 1, 10, 10)
	want := map[string]string{"B2": "1", "C2": "2", "B3": "3", "C3": "4", "B4": "5", "C4": "6"}
	if len(cells) != len(want) {
		t.Fatalf("got %d cells, want %d: %+v", len(cells), len(want), cells)
	}
	for ref, display := range want {
		got, ok := cells[ref]
		if !ok {
			t.Fatalf("cell %s is missing", ref)
		}
		if got.Display != display || got.Type != "number" {
			t.Errorf("%s = %+v, want display %s", ref, got, display)
		}
		isAnchor := ref == "B2"
		if got.Formula != isAnchor || got.Spill == isAnchor {
			t.Errorf("%s: formula=%v spill=%v, want formula=%v spill=%v", ref, got.Formula, got.Spill, isAnchor, !isAnchor)
		}
		if !isAnchor && got.Input != display {
			t.Errorf("%s input = %q, want %q", ref, got.Input, display)
		}
	}

	// The sheet size includes the spill range.
	if info := s.Sheets()[0]; info.MaxRow != 4 || info.MaxCol != 3 {
		t.Errorf("sheet size = %dx%d, want 4x3", info.MaxRow, info.MaxCol)
	}

	// A range that contains only followers still returns them, in row-major order.
	part, err := s.GetRange(sheet1, 3, 3, 4, 3)
	if err != nil {
		t.Fatalf("GetRange: %v", err)
	}
	if len(part) != 2 || part[0].Display != "4" || part[1].Display != "6" || !part[0].Spill || !part[1].Spill {
		t.Errorf("partial range = %+v, want followers 4 and 6", part)
	}

	// A formula can read a follower.
	mustSet(t, s, sheet1, "E1", "=C4*10")
	if got := cellsByRef(t, s, sheet1, 1, 5, 1, 5)["E1"].Display; got != "60" {
		t.Errorf("E1 display = %q, want 60", got)
	}

	// A value in the spill range blocks the spill.
	mustSet(t, s, sheet1, "C3", "x")
	cells = cellsByRef(t, s, sheet1, 1, 1, 4, 4)
	if got := cells["B2"]; got.Type != "error" || got.Display != "#SPILL!" || !got.Formula {
		t.Errorf("blocked anchor = %+v, want #SPILL!", got)
	}
	if got := cells["C3"]; got.Spill || got.Display != "x" {
		t.Errorf("blocker = %+v, want a plain string", got)
	}
	if _, ok := cells["B4"]; ok {
		t.Errorf("follower B4 is in the result while the spill is blocked")
	}

	// When the blocker is cleared, the spill comes back.
	mustSet(t, s, sheet1, "C3", "")
	cells = cellsByRef(t, s, sheet1, 1, 1, 4, 4)
	if got := cells["B2"].Display; got != "1" {
		t.Errorf("anchor display = %q, want 1", got)
	}
	if got := cells["C3"]; !got.Spill || got.Display != "4" {
		t.Errorf("C3 = %+v, want spill follower 4", got)
	}

	// When the anchor is cleared, the followers go away.
	mustSet(t, s, sheet1, "B2", "", "E1", "")
	if cells := cellsByRef(t, s, sheet1, 1, 1, 10, 10); len(cells) != 0 {
		t.Errorf("cells after the anchor was cleared = %+v, want none", cells)
	}
}

func TestSpillCellsRowMajorOrder(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "left", "B1", "=SEQUENCE(2,2)", "A2", "below", "D2", "right")
	cells, err := s.GetRange(sheet1, 1, 1, 5, 5)
	if err != nil {
		t.Fatalf("GetRange: %v", err)
	}
	var got []string
	for _, c := range cells {
		ref, _ := werkbook.CoordinatesToCellName(c.Col, c.Row)
		got = append(got, ref)
	}
	want := "A1 B1 C1 A2 B2 C2 D2"
	if strings.Join(got, " ") != want {
		t.Fatalf("order = %v, want %s", got, want)
	}
}

func TestOutOfRangeCoordinates(t *testing.T) {
	type call func(s *Session) error
	getRange := func(r1, c1, r2, c2 int) call {
		return func(s *Session) error {
			_, err := s.GetRange(sheet1, r1, c1, r2, c2)
			return err
		}
	}
	setCell := func(row, col int) call {
		return func(s *Session) error { return s.SetCell(sheet1, row, col, "x") }
	}
	setCells := func(row, col int) call {
		return func(s *Session) error {
			return s.SetCells(sheet1, []Edit{{Row: 1, Col: 1, Input: "ok"}, {Row: row, Col: col, Input: "x"}})
		}
	}
	tests := []struct {
		name string
		call call
	}{
		{"getRange row zero", getRange(0, 1, 5, 5)},
		{"getRange column zero", getRange(1, 0, 5, 5)},
		{"getRange negative", getRange(-1, -1, 5, 5)},
		{"getRange row past the end", getRange(1, 1, werkbook.MaxRows+1, 5)},
		{"getRange column past the end", getRange(1, 1, 5, werkbook.MaxColumns+1)},
		{"setCell row zero", setCell(0, 1)},
		{"setCell column zero", setCell(1, 0)},
		{"setCell row past the end", setCell(werkbook.MaxRows+1, 1)},
		{"setCell column past the end", setCell(1, werkbook.MaxColumns+1)},
		{"setCells row zero", setCells(0, 1)},
		{"setCells column past the end", setCells(1, werkbook.MaxColumns+1)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := NewSession()
			err := tt.call(s)
			if err == nil {
				t.Fatalf("got no error")
			}
			if !strings.Contains(err.Error(), "out of range") {
				t.Errorf("error = %q, want it to contain %q", err, "out of range")
			}
			// A rejected call must not change the workbook. For SetCells this
			// includes the valid edit that came before the bad one.
			if cells := cellsByRef(t, s, sheet1, 1, 1, 10, 10); len(cells) != 0 {
				t.Errorf("workbook changed after a rejected call: %+v", cells)
			}
		})
	}

	t.Run("the last cell of the grid is valid", func(t *testing.T) {
		s := NewSession()
		if err := s.SetCell(sheet1, werkbook.MaxRows, werkbook.MaxColumns, "end"); err != nil {
			t.Fatalf("SetCell: %v", err)
		}
		cells, err := s.GetRange(sheet1, werkbook.MaxRows, werkbook.MaxColumns, werkbook.MaxRows, werkbook.MaxColumns)
		if err != nil {
			t.Fatalf("GetRange: %v", err)
		}
		if len(cells) != 1 || cells[0].Display != "end" {
			t.Fatalf("cells = %+v, want the one cell", cells)
		}
	})

	t.Run("reversed corners give the same range", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "B2", "x")
		cells, err := s.GetRange(sheet1, 3, 3, 1, 1)
		if err != nil {
			t.Fatalf("GetRange: %v", err)
		}
		if len(cells) != 1 || cells[0].Display != "x" {
			t.Fatalf("cells = %+v, want B2", cells)
		}
	})
}

func TestUnknownSheet(t *testing.T) {
	tests := []struct {
		name string
		call func(s *Session) error
	}{
		{"getRange", func(s *Session) error { _, err := s.GetRange("Nope", 1, 1, 2, 2); return err }},
		{"sheetMeta", func(s *Session) error { _, err := s.SheetMeta("Nope"); return err }},
		{"setCell", func(s *Session) error { return s.SetCell("Nope", 1, 1, "x") }},
		{"setCells", func(s *Session) error { return s.SetCells("Nope", []Edit{{Row: 1, Col: 1, Input: "x"}}) }},
		{"setCells with no edits", func(s *Session) error { return s.SetCells("Nope", nil) }},
		{"renameSheet", func(s *Session) error { _, err := s.RenameSheet("Nope", "Other"); return err }},
		{"deleteSheet", func(s *Session) error { _, err := s.DeleteSheet("Nope"); return err }},
		{"empty name", func(s *Session) error { _, err := s.GetRange("", 1, 1, 2, 2); return err }},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := NewSession()
			err := tt.call(s)
			if err == nil {
				t.Fatalf("got no error")
			}
			if !strings.Contains(err.Error(), "not found") {
				t.Errorf("error = %q, want it to contain %q", err, "not found")
			}
			if got := s.Sheets(); len(got) != 1 || got[0].Name != sheet1 {
				t.Errorf("sheets changed: %+v", got)
			}
		})
	}
}

func TestSetCellsBatch(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "D1", "=SUM(A1:A3)")

	err := s.SetCells(sheet1, []Edit{
		{Row: 1, Col: 1, Input: "1"},
		{Row: 2, Col: 1, Input: "2"},
		{Row: 3, Col: 1, Input: "3"},
		{Row: 1, Col: 2, Input: "=A1*100"},
		{Row: 2, Col: 2, Input: "text"},
		// A later edit of the same cell wins.
		{Row: 3, Col: 1, Input: "30"},
	})
	if err != nil {
		t.Fatalf("SetCells: %v", err)
	}
	cells := cellsByRef(t, s, sheet1, 1, 1, 5, 5)
	want := map[string]string{"A1": "1", "A2": "2", "A3": "30", "B1": "100", "B2": "text", "D1": "33"}
	if len(cells) != len(want) {
		t.Fatalf("got %d cells, want %d: %+v", len(cells), len(want), cells)
	}
	for ref, display := range want {
		if got := cells[ref].Display; got != display {
			t.Errorf("%s display = %q, want %q", ref, got, display)
		}
	}

	// The batch state must be closed: a single write after it recalculates.
	mustSet(t, s, sheet1, "A1", "100")
	if got := cellsByRef(t, s, sheet1, 1, 1, 5, 5)["D1"].Display; got != "132" {
		t.Errorf("D1 display after a single write = %q, want 132", got)
	}

	// A batch can clear cells.
	if err := s.SetCells(sheet1, []Edit{{Row: 1, Col: 1, Input: ""}, {Row: 2, Col: 1, Input: ""}}); err != nil {
		t.Fatalf("SetCells: %v", err)
	}
	if got := cellsByRef(t, s, sheet1, 1, 1, 5, 5)["D1"].Display; got != "30" {
		t.Errorf("D1 display after a clear batch = %q, want 30", got)
	}

	// An empty batch is valid and changes nothing.
	if err := s.SetCells(sheet1, nil); err != nil {
		t.Fatalf("SetCells(nil): %v", err)
	}
}

func TestSheetOperations(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "5")

	sheets, err := s.AddSheet("Data")
	if err != nil {
		t.Fatalf("AddSheet: %v", err)
	}
	if len(sheets) != 2 || sheets[0].Name != sheet1 || sheets[1].Name != "Data" || !sheets[1].Visible {
		t.Fatalf("sheets after add = %+v", sheets)
	}
	if sheets[0].MaxRow != 1 || sheets[0].MaxCol != 1 || sheets[1].MaxRow != 0 || sheets[1].MaxCol != 0 {
		t.Errorf("sheet sizes = %+v, want 1x1 and 0x0", sheets)
	}

	if _, err := s.AddSheet("Data"); err == nil {
		t.Errorf("AddSheet with a name in use gave no error")
	}
	if _, err := s.AddSheet(""); err == nil {
		t.Errorf("AddSheet with an empty name gave no error")
	}

	// A formula on a different sheet follows a rename.
	mustSet(t, s, "Data", "A1", "=Sheet1!A1+1")
	sheets, err = s.RenameSheet(sheet1, "Input")
	if err != nil {
		t.Fatalf("RenameSheet: %v", err)
	}
	if sheets[0].Name != "Input" {
		t.Fatalf("sheets after rename = %+v", sheets)
	}
	got := cellsByRef(t, s, "Data", 1, 1, 1, 1)["A1"]
	if got.Display != "6" || got.Input != "=Input!A1+1" {
		t.Errorf("Data!A1 after rename = %+v, want display 6 and input =Input!A1+1", got)
	}
	if _, err := s.GetRange(sheet1, 1, 1, 1, 1); err == nil {
		t.Errorf("the old sheet name is still valid after a rename")
	}

	if _, err := s.RenameSheet("Input", "Data"); err == nil {
		t.Errorf("RenameSheet to a name in use gave no error")
	}
	if _, err := s.RenameSheet("Input", ""); err == nil {
		t.Errorf("RenameSheet to an empty name gave no error")
	}

	sheets, err = s.DeleteSheet("Input")
	if err != nil {
		t.Fatalf("DeleteSheet: %v", err)
	}
	if len(sheets) != 1 || sheets[0].Name != "Data" {
		t.Fatalf("sheets after delete = %+v", sheets)
	}
	// The formula now refers to a sheet that does not exist.
	if got := cellsByRef(t, s, "Data", 1, 1, 1, 1)["A1"]; got.Type != "error" {
		t.Errorf("Data!A1 after delete of its precedent sheet = %+v, want an error value", got)
	}

	if _, err := s.DeleteSheet("Data"); err == nil {
		t.Errorf("DeleteSheet of the last sheet gave no error")
	}
	if got := s.Sheets(); len(got) != 1 {
		t.Errorf("sheets = %+v, want one sheet", got)
	}
}

func TestRoundTripOpenEditSaveOpen(t *testing.T) {
	// Build a source file with the library, with data the session must keep.
	src := werkbook.New()
	sh := src.Sheet(sheet1)
	steps := []error{
		sh.SetValue("A1", "Item"),
		sh.SetValue("B1", "Price"),
		sh.SetValue("A2", "Apple"),
		sh.SetValue("B2", 1.5),
		sh.SetValue("A3", "Pear"),
		sh.SetValue("B3", 2.25),
		sh.SetFormula("B4", "SUM(B2:B3)"),
		sh.SetValue("C2", true),
		sh.SetStyle("A1", &werkbook.Style{Font: &werkbook.Font{Bold: true}}),
		sh.SetStyle("B4", &werkbook.Style{NumFmt: "0.00"}),
		sh.SetColumnWidth("A", 20),
		sh.SetRowHeight(1, 30),
		sh.MergeCell("D1", "E2"),
	}
	for i, err := range steps {
		if err != nil {
			t.Fatalf("setup step %d: %v", i, err)
		}
	}
	if _, err := src.NewSheet("Notes"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	var buf strings.Builder
	if _, err := src.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo: %v", err)
	}

	// Open.
	s := NewSession()
	sheets, err := s.Open([]byte(buf.String()))
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	if len(sheets) != 2 || sheets[0].Name != sheet1 || sheets[1].Name != "Notes" {
		t.Fatalf("sheets = %+v", sheets)
	}
	cells := cellsByRef(t, s, sheet1, 1, 1, 10, 10)
	if got := cells["B4"]; got.Display != "3.75" || !got.Formula || got.Input != "=SUM(B2:B3)" {
		t.Fatalf("B4 after open = %+v, want 3.75 from =SUM(B2:B3)", got)
	}
	if got := cells["A1"]; got.Style == nil || !got.Style.Bold {
		t.Errorf("A1 style after open = %+v, want bold", got.Style)
	}
	if got := cells["C2"]; got.Type != "bool" || got.Display != "TRUE" {
		t.Errorf("C2 after open = %+v, want TRUE", got)
	}

	// Edit: a value, a new formula, a new string, and a sheet rename.
	mustSet(t, s, sheet1, "B2", "10", "C4", "=B4*2", "A5", "'007")
	if err := s.SetCells("Notes", []Edit{{Row: 1, Col: 1, Input: "note"}, {Row: 2, Col: 1, Input: "=Sheet1!B4"}}); err != nil {
		t.Fatalf("SetCells: %v", err)
	}
	if _, err := s.RenameSheet("Notes", "Remarks"); err != nil {
		t.Fatalf("RenameSheet: %v", err)
	}

	// Save, then open in a new session.
	data, err := s.Save()
	if err != nil {
		t.Fatalf("Save: %v", err)
	}
	if len(data) < 4 || string(data[:2]) != "PK" {
		t.Fatalf("saved data is not a ZIP file (%d bytes)", len(data))
	}
	s2 := NewSession()
	sheets, err = s2.Open(data)
	if err != nil {
		t.Fatalf("Open of the saved file: %v", err)
	}
	if len(sheets) != 2 || sheets[0].Name != sheet1 || sheets[1].Name != "Remarks" {
		t.Fatalf("sheets after round trip = %+v", sheets)
	}

	cells = cellsByRef(t, s2, sheet1, 1, 1, 10, 10)
	want := map[string]Cell{
		"A1": {Row: 1, Col: 1, Display: "Item", Input: "Item", Type: "string", Style: &CellStyle{Bold: true}},
		"B2": {Row: 2, Col: 2, Display: "10", Input: "10", Type: "number"},
		"B3": {Row: 3, Col: 2, Display: "2.25", Input: "2.25", Type: "number"},
		"B4": {Row: 4, Col: 2, Display: "12.25", Input: "=SUM(B2:B3)", Type: "number", Formula: true, Style: &CellStyle{NumFmt: "custom"}},
		"C2": {Row: 2, Col: 3, Display: "TRUE", Input: "TRUE", Type: "bool"},
		"C4": {Row: 4, Col: 3, Display: "24.5", Input: "=B4*2", Type: "number", Formula: true},
		"A5": {Row: 5, Col: 1, Display: "007", Input: "'007", Type: "string"},
	}
	for ref, w := range want {
		got, ok := cells[ref]
		if !ok {
			t.Errorf("%s is missing after the round trip", ref)
			continue
		}
		gotStyle, wantStyle := got.Style, w.Style
		got.Style, w.Style = nil, nil
		if got != w {
			t.Errorf("%s = %+v, want %+v", ref, got, w)
		}
		if (gotStyle == nil) != (wantStyle == nil) || (gotStyle != nil && *gotStyle != *wantStyle) {
			t.Errorf("%s style = %+v, want %+v", ref, gotStyle, wantStyle)
		}
	}

	remarks := cellsByRef(t, s2, "Remarks", 1, 1, 5, 5)
	if got := remarks["A1"].Display; got != "note" {
		t.Errorf("Remarks!A1 = %q, want note", got)
	}
	if got := remarks["A2"]; got.Display != "12.25" || got.Input != "=Sheet1!B4" {
		t.Errorf("Remarks!A2 = %+v, want 12.25 from =Sheet1!B4", got)
	}

	// The layout data survives the round trip.
	meta, err := s2.SheetMeta(sheet1)
	if err != nil {
		t.Fatalf("SheetMeta: %v", err)
	}
	if got := meta.ColWidths[1]; got != 145 {
		t.Errorf("width of column A = %d px, want 145", got)
	}
	if got := meta.RowHeights[1]; got != 40 {
		t.Errorf("height of row 1 = %d px, want 40", got)
	}
	if len(meta.Merges) != 1 || meta.Merges[0] != (Merge{Row1: 1, Col1: 4, Row2: 2, Col2: 5}) {
		t.Errorf("merges = %+v, want D1:E2", meta.Merges)
	}

	// The opened workbook stays editable: formulas recalculate.
	mustSet(t, s2, sheet1, "B3", "0")
	if got := cellsByRef(t, s2, sheet1, 4, 2, 4, 3); got["B4"].Display != "10.00" || got["C4"].Display != "20" {
		t.Errorf("after edit in the opened file: B4=%q C4=%q, want 10.00 and 20", got["B4"].Display, got["C4"].Display)
	}
}

func TestRoundTripSpill(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "=SEQUENCE(3)", "C1", "=SUM(A1:A3)")
	data, err := s.Save()
	if err != nil {
		t.Fatalf("Save: %v", err)
	}

	s2 := NewSession()
	if _, err := s2.Open(data); err != nil {
		t.Fatalf("Open: %v", err)
	}
	cells := cellsByRef(t, s2, sheet1, 1, 1, 5, 5)
	if got := cells["A1"]; !got.Formula || got.Spill || got.Display != "1" {
		t.Errorf("anchor = %+v, want formula cell 1", got)
	}
	for ref, display := range map[string]string{"A2": "2", "A3": "3"} {
		if got := cells[ref]; !got.Spill || got.Formula || got.Display != display {
			t.Errorf("%s = %+v, want spill follower %s", ref, got, display)
		}
	}
	if got := cells["C1"].Display; got != "6" {
		t.Errorf("C1 = %q, want 6", got)
	}
	if len(cells) != 4 {
		t.Errorf("got %d cells, want 4: %+v", len(cells), cells)
	}
}

func TestOpenRejectsBadData(t *testing.T) {
	tests := []struct {
		name string
		data []byte
	}{
		{"nil", nil},
		{"empty", []byte{}},
		{"not a zip", []byte("this is not a workbook")},
		{"truncated zip", []byte("PK\x03\x04truncated")},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := NewSession()
			mustSet(t, s, sheet1, "A1", "keep")
			if _, err := s.Open(tt.data); err == nil {
				t.Fatalf("Open gave no error")
			}
			// The session keeps the workbook it had.
			if got := cellsByRef(t, s, sheet1, 1, 1, 1, 1)["A1"].Display; got != "keep" {
				t.Errorf("A1 after a failed open = %q, want keep", got)
			}
		})
	}
}

func TestNewReplacesWorkbook(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "old")
	if _, err := s.AddSheet("Second"); err != nil {
		t.Fatalf("AddSheet: %v", err)
	}
	sheets := s.New()
	if len(sheets) != 1 || sheets[0] != (SheetInfo{Name: sheet1, Visible: true}) {
		t.Fatalf("sheets after New = %+v", sheets)
	}
	if cells := cellsByRef(t, s, sheet1, 1, 1, 5, 5); len(cells) != 0 {
		t.Errorf("cells after New = %+v, want none", cells)
	}
}

func TestStyleAndDisplay(t *testing.T) {
	tests := []struct {
		name        string
		input       string
		style       *werkbook.Style
		wantDisplay string
		wantType    string
		wantStyle   *CellStyle
	}{
		{
			name:        "no style",
			input:       "1.5",
			wantDisplay: "1.5",
			wantType:    "number",
		},
		{
			name:  "font, fill and alignment",
			input: "x",
			style: &werkbook.Style{
				Font:      &werkbook.Font{Bold: true, Italic: true, Underline: true, Color: "ff0000"},
				Fill:      &werkbook.Fill{Color: "FFFF00"},
				Alignment: &werkbook.Alignment{Horizontal: werkbook.HAlignCenter},
			},
			wantDisplay: "x",
			wantType:    "string",
			wantStyle:   &CellStyle{Bold: true, Italic: true, Underline: true, Color: "#FF0000", Fill: "#FFFF00", Align: "center"},
		},
		{
			name:        "ARGB color",
			input:       "x",
			style:       &werkbook.Style{Fill: &werkbook.Fill{Color: "FF00B050"}},
			wantDisplay: "x",
			wantType:    "string",
			wantStyle:   &CellStyle{Fill: "#00B050"},
		},
		{
			name:        "color that is not hex is dropped",
			input:       "x",
			style:       &werkbook.Style{Font: &werkbook.Font{Color: "theme1"}, Alignment: &werkbook.Alignment{Horizontal: werkbook.HAlignRight}},
			wantDisplay: "x",
			wantType:    "string",
			wantStyle:   &CellStyle{Align: "right"},
		},
		{
			name:        "style with no field of the subset is omitted",
			input:       "x",
			style:       &werkbook.Style{Font: &werkbook.Font{Name: "Arial", Size: 14}, Alignment: &werkbook.Alignment{WrapText: true}},
			wantDisplay: "x",
			wantType:    "string",
		},
		{
			name:        "custom number format",
			input:       "1234.5",
			style:       &werkbook.Style{NumFmt: "#,##0.00"},
			wantDisplay: "1,234.50",
			wantType:    "number",
			wantStyle:   &CellStyle{NumFmt: "number"},
		},
		{
			name:        "built-in percent format",
			input:       "0.5",
			style:       &werkbook.Style{NumFmtID: 9},
			wantDisplay: "50%",
			wantType:    "number",
			wantStyle:   &CellStyle{NumFmt: "percent"},
		},
		{
			name:        "built-in date format",
			input:       "45292", // 2024-01-01
			style:       &werkbook.Style{NumFmtID: 14},
			wantDisplay: "1/1/2024",
			wantType:    "number",
			wantStyle:   &CellStyle{NumFmt: "date"},
		},
		{
			name:        "custom date format",
			input:       "45292",
			style:       &werkbook.Style{NumFmt: "yyyy-mm-dd"},
			wantDisplay: "2024-01-01",
			wantType:    "number",
			wantStyle:   &CellStyle{NumFmt: "custom"},
		},
		{
			name:        "negative number in a two-section format",
			input:       "-5",
			style:       &werkbook.Style{NumFmtID: 37},
			wantDisplay: "(5)",
			wantType:    "number",
			wantStyle:   &CellStyle{NumFmt: "custom"},
		},
		{
			name:        "general format code",
			input:       "2.5",
			style:       &werkbook.Style{NumFmt: "General"},
			wantDisplay: "2.5",
			wantType:    "number",
		},
		{
			name:        "unknown built-in format falls back to general",
			input:       "2.5",
			style:       &werkbook.Style{NumFmtID: 163},
			wantDisplay: "2.5",
			wantType:    "number",
			wantStyle:   &CellStyle{NumFmt: "custom"},
		},
		{
			name:        "number format does not change a string",
			input:       "text",
			style:       &werkbook.Style{NumFmt: "0.00"},
			wantDisplay: "text",
			wantType:    "string",
			wantStyle:   &CellStyle{NumFmt: "custom"},
		},
		{
			name:        "number format does not change an error",
			input:       "=1/0",
			style:       &werkbook.Style{NumFmt: "0.00"},
			wantDisplay: "#DIV/0!",
			wantType:    "error",
			wantStyle:   &CellStyle{NumFmt: "custom"},
		},
		{
			name:        "number format applies to a formula result",
			input:       "=1/4",
			style:       &werkbook.Style{NumFmt: "0.0%"},
			wantDisplay: "25.0%",
			wantType:    "number",
			wantStyle:   &CellStyle{NumFmt: "custom"},
		},
		{
			name:        "empty cell with a fill is returned",
			input:       "",
			style:       &werkbook.Style{Fill: &werkbook.Fill{Color: "CCCCCC"}},
			wantDisplay: "",
			wantType:    "empty",
			wantStyle:   &CellStyle{Fill: "#CCCCCC"},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := NewSession()
			if tt.style != nil {
				if err := s.file.Sheet(sheet1).SetStyle("A1", tt.style); err != nil {
					t.Fatalf("SetStyle: %v", err)
				}
			}
			mustSet(t, s, sheet1, "A1", tt.input)
			got, ok := cellsByRef(t, s, sheet1, 1, 1, 1, 1)["A1"]
			if !ok {
				t.Fatalf("A1 is not in the result")
			}
			if got.Display != tt.wantDisplay || got.Type != tt.wantType {
				t.Errorf("display %q type %q, want display %q type %q", got.Display, got.Type, tt.wantDisplay, tt.wantType)
			}
			switch {
			case tt.wantStyle == nil && got.Style != nil:
				t.Errorf("style = %+v, want none", *got.Style)
			case tt.wantStyle != nil && got.Style == nil:
				t.Errorf("style is missing, want %+v", *tt.wantStyle)
			case tt.wantStyle != nil && *got.Style != *tt.wantStyle:
				t.Errorf("style = %+v, want %+v", *got.Style, *tt.wantStyle)
			}
		})
	}

	t.Run("empty cell with a style that the grid does not show is omitted", func(t *testing.T) {
		s := NewSession()
		if err := s.file.Sheet(sheet1).SetStyle("A1", &werkbook.Style{NumFmt: "0.00"}); err != nil {
			t.Fatalf("SetStyle: %v", err)
		}
		if cells := cellsByRef(t, s, sheet1, 1, 1, 1, 1); len(cells) != 0 {
			t.Errorf("cells = %+v, want none", cells)
		}
	})

	t.Run("a cleared cell keeps its style", func(t *testing.T) {
		s := NewSession()
		if err := s.file.Sheet(sheet1).SetStyle("A1", &werkbook.Style{Font: &werkbook.Font{Bold: true}}); err != nil {
			t.Fatalf("SetStyle: %v", err)
		}
		mustSet(t, s, sheet1, "A1", "x", "A1", "", "A1", "y")
		got := cellsByRef(t, s, sheet1, 1, 1, 1, 1)["A1"]
		if got.Display != "y" || got.Style == nil || !got.Style.Bold {
			t.Errorf("A1 = %+v (style %+v), want bold y", got, got.Style)
		}
	})
}

func TestSheetMeta(t *testing.T) {
	s := NewSession()
	sh := s.file.Sheet(sheet1)
	steps := []error{
		sh.SetColumnWidth("B", 10),
		sh.SetColumnWidth("XFD", 8.43),
		sh.SetRowHeight(3, 15),
		sh.SetRowHeight(7, 22.5),
		sh.MergeCell("A1", "C1"),
		sh.MergeCell("B5", "B9"),
	}
	for i, err := range steps {
		if err != nil {
			t.Fatalf("setup step %d: %v", i, err)
		}
	}

	meta, err := s.SheetMeta(sheet1)
	if err != nil {
		t.Fatalf("SheetMeta: %v", err)
	}
	wantWidths := map[int]int{2: 75, werkbook.MaxColumns: 64}
	if len(meta.ColWidths) != len(wantWidths) {
		t.Errorf("colWidths = %v, want %v", meta.ColWidths, wantWidths)
	}
	for col, px := range wantWidths {
		if meta.ColWidths[col] != px {
			t.Errorf("width of column %d = %d px, want %d", col, meta.ColWidths[col], px)
		}
	}
	wantHeights := map[int]int{3: 20, 7: 30}
	if len(meta.RowHeights) != len(wantHeights) {
		t.Errorf("rowHeights = %v, want %v", meta.RowHeights, wantHeights)
	}
	for row, px := range wantHeights {
		if meta.RowHeights[row] != px {
			t.Errorf("height of row %d = %d px, want %d", row, meta.RowHeights[row], px)
		}
	}
	wantMerges := []Merge{{Row1: 1, Col1: 1, Row2: 1, Col2: 3}, {Row1: 5, Col1: 2, Row2: 9, Col2: 2}}
	if len(meta.Merges) != len(wantMerges) {
		t.Fatalf("merges = %+v, want %+v", meta.Merges, wantMerges)
	}
	for i, m := range wantMerges {
		if meta.Merges[i] != m {
			t.Errorf("merge %d = %+v, want %+v", i, meta.Merges[i], m)
		}
	}
}

// The WASM wrapper sends results to JavaScript as JSON, so the field names
// are part of the contract in cmd/web/README.md.
func TestJSONShape(t *testing.T) {
	tests := []struct {
		name string
		v    any
		want string
	}{
		{
			name: "plain cell omits spill and style",
			v:    Cell{Row: 1, Col: 2, Display: "3", Input: "=1+2", Type: "number", Formula: true},
			want: `{"row":1,"col":2,"display":"3","input":"=1+2","type":"number","formula":true}`,
		},
		{
			name: "spill cell with a style",
			v:    Cell{Row: 2, Col: 1, Display: "x", Input: "x", Type: "string", Spill: true, Style: &CellStyle{Bold: true, Fill: "#FFFF00", Align: "right"}},
			want: `{"row":2,"col":1,"display":"x","input":"x","type":"string","formula":false,"spill":true,"style":{"bold":true,"fill":"#FFFF00","align":"right"}}`,
		},
		{
			name: "sheet info",
			v:    SheetInfo{Name: "S", MaxRow: 3, MaxCol: 4, Visible: true},
			want: `{"name":"S","maxRow":3,"maxCol":4,"visible":true}`,
		},
		{
			name: "sheet meta",
			v:    SheetMeta{ColWidths: map[int]int{2: 75}, RowHeights: map[int]int{}, Merges: []Merge{{Row1: 1, Col1: 1, Row2: 2, Col2: 3}}, Warnings: []string{}},
			want: `{"colWidths":{"2":75},"rowHeights":{},"merges":[{"row1":1,"col1":1,"row2":2,"col2":3}],"warnings":[],"freeze":{"rows":0,"cols":0}}`,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := json.Marshal(tt.v)
			if err != nil {
				t.Fatalf("Marshal: %v", err)
			}
			if string(got) != tt.want {
				t.Errorf("JSON = %s\nwant   %s", got, tt.want)
			}
		})
	}
}
