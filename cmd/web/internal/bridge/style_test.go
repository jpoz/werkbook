package bridge

import (
	"encoding/json"
	"strings"
	"testing"

	werkbook "github.com/jpoz/werkbook"
)

// patchFromJSON decodes a patch as the WASM wrapper does: strict, with null
// for "remove".
func patchFromJSON(t *testing.T, text string) StylePatch {
	t.Helper()
	patch, err := decodePatch(text)
	if err != nil {
		t.Fatalf("patch %s: %v", text, err)
	}
	return patch
}

func decodePatch(text string) (StylePatch, error) {
	var patch StylePatch
	dec := json.NewDecoder(strings.NewReader(text))
	dec.DisallowUnknownFields()
	err := dec.Decode(&patch)
	return patch, err
}

func mustStyle(t *testing.T, s *Session, rg, patch string) StyleResult {
	t.Helper()
	res, err := s.SetStyle(sheet1, rangeOf(t, rg), patchFromJSON(t, patch))
	if err != nil {
		t.Fatalf("SetStyle(%s, %s): %v", rg, patch, err)
	}
	return res
}

func styleOf(t *testing.T, s *Session, ref string) CellStyle {
	t.Helper()
	col, row, err := werkbook.CellNameToCoordinates(ref)
	if err != nil {
		t.Fatal(err)
	}
	st, err := s.file.Sheet(sheet1).GetStyle(ref)
	if err != nil {
		t.Fatal(err)
	}
	_ = col
	_ = row
	if cs := cellStyle(st); cs != nil {
		return *cs
	}
	return CellStyle{}
}

func TestSetStylePatchFields(t *testing.T) {
	tests := []struct {
		name  string
		patch string
		want  CellStyle
	}{
		{"bold", `{"bold":true}`, CellStyle{Bold: true}},
		{"italic", `{"italic":true}`, CellStyle{Italic: true}},
		{"underline", `{"underline":true}`, CellStyle{Underline: true}},
		{"font color", `{"color":"#ff0000"}`, CellStyle{Color: "#FF0000"}},
		{"fill", `{"fill":"#00B050"}`, CellStyle{Fill: "#00B050"}},
		{"align left", `{"align":"left"}`, CellStyle{Align: "left"}},
		{"align center", `{"align":"center"}`, CellStyle{Align: "center"}},
		{"align right", `{"align":"right"}`, CellStyle{Align: "right"}},
		{"number", `{"numFmt":"number"}`, CellStyle{NumFmt: "number"}},
		{"currency", `{"numFmt":"currency"}`, CellStyle{NumFmt: "currency"}},
		{"percent", `{"numFmt":"percent"}`, CellStyle{NumFmt: "percent"}},
		{"date", `{"numFmt":"date"}`, CellStyle{NumFmt: "date"}},
		{"all fields", `{"bold":true,"italic":true,"underline":true,"color":"#112233","fill":"#445566","align":"center","numFmt":"percent"}`,
			CellStyle{Bold: true, Italic: true, Underline: true, Color: "#112233", Fill: "#445566", Align: "center", NumFmt: "percent"}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := NewSession()
			mustSet(t, s, sheet1, "A1", "1", "B2", "x")
			before := dump(t, s, sheet1)
			res := mustStyle(t, s, "A1:B2", tt.patch)
			if res.Changed != 4 || res.Clamped || res.Range != rangeOf(t, "A1:B2") {
				t.Errorf("result = %+v", res)
			}
			if res.History.Depth != 3 || res.History.UndoLabel != "Format" {
				t.Errorf("history = %+v", res.History)
			}
			for _, ref := range []string{"A1", "B1", "A2", "B2"} {
				if got := styleOf(t, s, ref); got != tt.want {
					t.Errorf("%s style = %+v, want %+v", ref, got, tt.want)
				}
			}
			// Values and types do not change.
			if got := inputsOf(t, s, sheet1); got["A1"] != "1" || got["B2"] != "x" {
				t.Errorf("inputs = %v", got)
			}
			// The four cells share one style object.
			sh := s.file.Sheet(sheet1)
			a1, _ := sh.GetStyle("A1")
			b2, _ := sh.GetStyle("B2")
			if a1 == nil || a1 != b2 {
				t.Errorf("A1 and B2 have different style objects: %p %p", a1, b2)
			}
			undo := mustUndo(t, s)
			if undo.Label != "Format" || *undo.Select != rangeOf(t, "A1:B2") {
				t.Errorf("undo = %+v %v", undo, undo.Select)
			}
			if got := dump(t, s, sheet1); got != before {
				t.Errorf("state after undo:\n%s\nwant:\n%s", got, before)
			}
			mustRedo(t, s)
			if got := styleOf(t, s, "B1"); got != tt.want {
				t.Errorf("B1 style after redo = %+v", got)
			}
		})
	}
}

func TestSetStyleRemoveAndToggle(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "1")
	mustStyle(t, s, "A1:A2", `{"bold":true,"italic":true,"color":"#FF0000","fill":"#FFFF00","align":"center","numFmt":"date"}`)

	steps := []struct {
		patch string
		want  CellStyle
	}{
		{`{"bold":false}`, CellStyle{Italic: true, Color: "#FF0000", Fill: "#FFFF00", Align: "center", NumFmt: "date"}},
		{`{"color":null}`, CellStyle{Italic: true, Fill: "#FFFF00", Align: "center", NumFmt: "date"}},
		{`{"fill":null}`, CellStyle{Italic: true, Align: "center", NumFmt: "date"}},
		{`{"align":"general"}`, CellStyle{Italic: true, NumFmt: "date"}},
		{`{"numFmt":"general"}`, CellStyle{Italic: true}},
		{`{"italic":false}`, CellStyle{}},
	}
	for _, st := range steps {
		res := mustStyle(t, s, "A1:A2", st.patch)
		if res.Changed != 2 {
			t.Errorf("patch %s: changed = %d, want 2", st.patch, res.Changed)
		}
		for _, ref := range []string{"A1", "A2"} {
			if got := styleOf(t, s, ref); got != st.want {
				t.Errorf("patch %s: %s style = %+v, want %+v", st.patch, ref, got, st.want)
			}
		}
	}
	// The style of A1 is the default style again (nil), and the empty cell
	// A2 is not a stored cell any more.
	sh := s.file.Sheet(sheet1)
	if st, _ := sh.GetStyle("A1"); st != nil {
		t.Errorf("A1 style = %+v, want nil", st)
	}
	if sh.SnapshotCell(1, 2).Exists() {
		t.Errorf("the empty cell A2 with the default style is still stored")
	}
	if info := s.Sheets()[0]; info.MaxRow != 1 {
		t.Errorf("MaxRow = %d, want 1", info.MaxRow)
	}

	// A patch that changes nothing makes no step and reports 0.
	depth := s.History().Depth
	for _, patch := range []string{`{"bold":false}`, `{"color":null,"fill":null}`, `{"align":"general","numFmt":"general"}`, `{}`} {
		res := mustStyle(t, s, "A1:C3", patch)
		if res.Changed != 0 || res.History.Depth != depth {
			t.Errorf("patch %s: result = %+v, want no change", patch, res)
		}
	}
	mustStyle(t, s, "A1", `{"bold":true}`)
	if res := mustStyle(t, s, "A1", `{"bold":true}`); res.Changed != 0 || res.History.Depth != depth+1 {
		t.Errorf("same patch a second time: %+v", res)
	}
}

func TestSetStyleKeepsOtherStylePartsAndSharedPointers(t *testing.T) {
	s := NewSession()
	sh := s.file.Sheet(sheet1)
	// One style object for A1:C1, as cells of a file have that used one
	// cellXfs entry. It has parts that the patch has no field for.
	shared := &werkbook.Style{
		Font:      &werkbook.Font{Name: "Arial", Size: 14, Italic: true},
		Border:    &werkbook.Border{Bottom: werkbook.BorderSide{Style: werkbook.BorderThin, Color: "000000"}},
		Alignment: &werkbook.Alignment{Vertical: werkbook.VAlignCenter, WrapText: true},
		NumFmt:    "0.000",
	}
	if err := sh.SetRangeStyle("A1:C1", shared); err != nil {
		t.Fatal(err)
	}
	mustSet(t, s, sheet1, "A1", "1", "B1", "2", "C1", "3")
	before := *shared
	beforeFont := *shared.Font

	res := mustStyle(t, s, "A1:B1", `{"bold":true,"fill":"#FFFF00","align":"right"}`)
	if res.Changed != 2 {
		t.Errorf("changed = %d", res.Changed)
	}
	// C1 is out of the range: its style object is the old one, unchanged.
	c1, _ := sh.GetStyle("C1")
	if c1 != shared || *shared.Font != beforeFont || shared.Fill != nil || shared.NumFmt != before.NumFmt || shared.Alignment.Horizontal != werkbook.HAlignGeneral {
		t.Errorf("the shared style object changed: %+v (font %+v)", *shared, *shared.Font)
	}
	a1, _ := sh.GetStyle("A1")
	b1, _ := sh.GetStyle("B1")
	if a1 == shared || a1 != b1 {
		t.Errorf("A1 and B1 must share one new style object: %p %p (old %p)", a1, b1, shared)
	}
	// The new style has the patch and all the other parts.
	if !a1.Font.Bold || !a1.Font.Italic || a1.Font.Name != "Arial" || a1.Font.Size != 14 {
		t.Errorf("font = %+v", *a1.Font)
	}
	if a1.Fill == nil || a1.Fill.Color != "FFFF00" {
		t.Errorf("fill = %+v", a1.Fill)
	}
	if a1.Border == nil || a1.Border.Bottom.Style != werkbook.BorderThin {
		t.Errorf("border = %+v", a1.Border)
	}
	if a1.Alignment.Horizontal != werkbook.HAlignRight || a1.Alignment.Vertical != werkbook.VAlignCenter || !a1.Alignment.WrapText {
		t.Errorf("alignment = %+v", *a1.Alignment)
	}
	if a1.NumFmt != "0.000" {
		t.Errorf("numFmt = %q, want the custom format to stay", a1.NumFmt)
	}
	if got := styleOf(t, s, "A1"); got.NumFmt != "custom" {
		t.Errorf("reported numFmt = %q, want custom", got.NumFmt)
	}

	// Undo gives the old object back to A1 and B1.
	mustUndo(t, s)
	if st, _ := sh.GetStyle("A1"); st != shared {
		t.Errorf("A1 style after undo = %p, want the shared object %p", st, shared)
	}

	// A removal of all patch fields leaves the parts with no field.
	mustStyle(t, s, "A1", `{"italic":false,"numFmt":"general"}`)
	a1, _ = sh.GetStyle("A1")
	if a1 == nil || a1.Font == nil || a1.Font.Name != "Arial" || a1.Border == nil || a1.NumFmt != "" {
		t.Errorf("style after removal = %+v", a1)
	}
}

func TestSetStyleNumFmtDisplayAndRoundTrip(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "1234.5", "A2", "1234.5", "A3", "0.25", "A4", "45000", "A5", "1234.5", "A6", "text", "A7", "=1/3")
	mustStyle(t, s, "A1", `{"numFmt":"number"}`)
	mustStyle(t, s, "A2", `{"numFmt":"currency"}`)
	mustStyle(t, s, "A3", `{"numFmt":"percent"}`)
	mustStyle(t, s, "A4", `{"numFmt":"date"}`)
	mustStyle(t, s, "A5", `{"numFmt":"general","bold":true}`)
	mustStyle(t, s, "A6", `{"numFmt":"date"}`)
	mustStyle(t, s, "A7", `{"numFmt":"percent"}`)

	want := map[string][2]string{ // display, numFmt
		"A1": {"1,234.50", "number"},
		"A2": {"$1,234.50", "currency"},
		"A3": {"25%", "percent"},
		"A4": {"3/15/2023", "date"},
		"A5": {"1234.5", ""},
		"A6": {"text", "date"}, // a date format on text shows the text
		"A7": {"33%", "percent"},
	}
	check := func(stage string, session *Session) {
		t.Helper()
		cells := cellsByRef(t, session, sheet1, 1, 1, 7, 1)
		for ref, w := range want {
			got := cells[ref]
			numFmt := ""
			if got.Style != nil {
				numFmt = got.Style.NumFmt
			}
			if got.Display != w[0] || numFmt != w[1] {
				t.Errorf("%s: %s = display %q numFmt %q, want %q %q", stage, ref, got.Display, numFmt, w[0], w[1])
			}
		}
		// The values are not changed by a format.
		if got := cells["A3"].Input; got != "0.25" {
			t.Errorf("%s: A3 input = %q", stage, got)
		}
		if st := cells["A5"].Style; st == nil || !st.Bold {
			t.Errorf("%s: A5 style = %+v, want bold", stage, st)
		}
	}
	check("in the session", s)

	data, err := s.Save()
	if err != nil {
		t.Fatal(err)
	}
	s2 := NewSession()
	if _, err := s2.Open(data); err != nil {
		t.Fatal(err)
	}
	check("after save and open", s2)

	// A format from a file that is not a preset is "custom", and stays when
	// a patch has no numFmt.
	if err := s.file.Sheet(sheet1).SetStyle("C1", &werkbook.Style{NumFmt: "0.0000"}); err != nil {
		t.Fatal(err)
	}
	mustSet(t, s, sheet1, "C1", "2")
	mustStyle(t, s, "C1", `{"bold":true}`)
	if got := cellsByRef(t, s, sheet1, 1, 3, 1, 3)["C1"]; got.Display != "2.0000" || got.Style.NumFmt != "custom" || !got.Style.Bold {
		t.Errorf("C1 = %+v (style %+v)", got, got.Style)
	}

	// An empty cell with only a number format is not in a range result.
	mustStyle(t, s, "E5", `{"numFmt":"percent"}`)
	if _, ok := cellsByRef(t, s, sheet1, 5, 5, 5, 5)["E5"]; ok {
		t.Errorf("the empty cell E5 with only a number format is in the result")
	}
	if !s.file.Sheet(sheet1).SnapshotCell(5, 5).Exists() {
		t.Errorf("the styled cell E5 is not stored")
	}
	// It applies when the cell gets a value.
	mustSet(t, s, sheet1, "E5", "0.5")
	if got := cellsByRef(t, s, sheet1, 5, 5, 5, 5)["E5"].Display; got != "50%" {
		t.Errorf("E5 = %q, want 50%%", got)
	}
}

// A negative number with a date format shows number signs, as in Excel, and
// not a date before 1900.
func TestDateFormatOfANegativeNumber(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "-1234.567", "A2", "-0.5", "A3", "0.5", "A4", "0", "A5", "=A1*2", "A6", "45000", "A7", "-3", "A8", "-3")
	mustStyle(t, s, "A1:A6", `{"numFmt":"date"}`)
	if err := s.file.Sheet(sheet1).SetStyle("A7", &werkbook.Style{NumFmt: "h:mm:ss"}); err != nil {
		t.Fatal(err)
	}
	cells := cellsByRef(t, s, sheet1, 1, 1, 8, 1)
	want := map[string]string{
		"A1": "########", "A2": "########", "A5": "########", // negative values and a negative formula result
		"A3": "1/0/1900", "A4": "1/0/1900", "A6": "3/15/2023",
		"A7": "########", // a time format is a date format too
		"A8": "-3",       // no format: the number
	}
	for ref, w := range want {
		if got := cells[ref].Display; got != w {
			t.Errorf("%s display = %q, want %q", ref, got, w)
		}
	}
	// The value and the editor text do not change.
	if got := cells["A1"]; got.Input != "-1234.567" || got.Type != "number" {
		t.Errorf("A1 = %+v", got)
	}
	// The other presets show a negative number.
	for preset, w := range map[string]string{"number": "-1,234.57", "currency": "-$1,234.57", "percent": "-123457%", "general": "-1234.567"} {
		mustStyle(t, s, "A1", `{"numFmt":"`+preset+`"}`)
		if got := cellsByRef(t, s, sheet1, 1, 1, 1, 1)["A1"].Display; got != w {
			t.Errorf("preset %s: display = %q, want %q", preset, got, w)
		}
	}
	// The copy text of a formula cell is its display text.
	mustStyle(t, s, "A5", `{"numFmt":"date"}`)
	if cp := mustCopy(t, s, sheet1, "A5", false); cp.Text != "########" {
		t.Errorf("copy text = %q", cp.Text)
	}
}

func TestSetStyleClampAndSpecialCells(t *testing.T) {
	t.Run("full columns, full rows and select all are limited to the used range", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "1", "C4", "x")
		tests := []struct {
			name      string
			rg        Range
			wantRange string
			changed   int
		}{
			{"column A", Range{1, 1, werkbook.MaxRows, 1}, "A1:A4", 4},
			{"columns B:C", Range{1, 2, werkbook.MaxRows, 3}, "B1:C4", 8},
			{"row 4", Range{4, 1, 4, werkbook.MaxColumns}, "A4:C4", 3},
			{"all", Range{1, 1, werkbook.MaxRows, werkbook.MaxColumns}, "A1:C4", 12},
		}
		for _, tt := range tests {
			res, err := s.SetStyle(sheet1, tt.rg, patchFromJSON(t, `{"bold":true}`))
			if err != nil {
				t.Fatalf("%s: %v", tt.name, err)
			}
			if !res.Clamped || res.Range != rangeOf(t, tt.wantRange) || res.Changed != tt.changed {
				t.Errorf("%s: result = %+v, want range %s changed %d clamped", tt.name, res, tt.wantRange, tt.changed)
			}
			mustUndo(t, s)
		}
		// Rows and columns with no used cell: nothing changes.
		depth := s.History().Depth
		for _, rg := range []Range{{1, 9, werkbook.MaxRows, 9}, {20, 1, 22, werkbook.MaxColumns}} {
			res, err := s.SetStyle(sheet1, rg, patchFromJSON(t, `{"bold":true}`))
			if err != nil {
				t.Fatal(err)
			}
			want := Range{rg.Row1, rg.Col1, rg.Row1, rg.Col1}
			if !res.Clamped || res.Changed != 0 || res.Range != want || res.History.Depth != depth {
				t.Errorf("range %+v: result = %+v, want no change", rg, res)
			}
		}
		// An empty sheet.
		empty := NewSession()
		res, err := empty.SetStyle(sheet1, Range{1, 1, werkbook.MaxRows, werkbook.MaxColumns}, patchFromJSON(t, `{"bold":true}`))
		if err != nil || res.Changed != 0 || !res.Clamped || res.Range != (Range{1, 1, 1, 1}) {
			t.Errorf("empty sheet: %+v, %v", res, err)
		}
		// A range that is not full rows or columns is not limited.
		res = mustStyle(t, s, "J10:K11", `{"bold":true}`)
		if res.Clamped || res.Changed != 4 {
			t.Errorf("result = %+v", res)
		}
	})

	t.Run("merged range: the style goes to all its cells", func(t *testing.T) {
		s := NewSession()
		if err := s.file.Sheet(sheet1).MergeCell("B2", "C3"); err != nil {
			t.Fatal(err)
		}
		mustSet(t, s, sheet1, "B2", "merged")
		res := mustStyle(t, s, "B2:C3", `{"fill":"#FFFF00"}`)
		if res.Changed != 4 {
			t.Errorf("changed = %d, want 4", res.Changed)
		}
		for _, ref := range []string{"B2", "C2", "B3", "C3"} {
			if got := styleOf(t, s, ref); got.Fill != "#FFFF00" {
				t.Errorf("%s fill = %q", ref, got.Fill)
			}
		}
	})

	t.Run("spill followers take a style and the spill stays", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(3)")
		before := dump(t, s, sheet1)
		res := mustStyle(t, s, "A1:A3", `{"bold":true,"numFmt":"number"}`)
		if res.Changed != 3 {
			t.Errorf("changed = %d", res.Changed)
		}
		cells := cellsByRef(t, s, sheet1, 1, 1, 3, 1)
		if got := cells["A1"]; got.Input != "=SEQUENCE(3)" || got.Display != "1.00" || !got.Style.Bold {
			t.Errorf("anchor = %+v", got)
		}
		for ref, display := range map[string]string{"A2": "2.00", "A3": "3.00"} {
			got := cells[ref]
			if !got.Spill || got.Display != display || got.Style == nil || !got.Style.Bold {
				t.Errorf("%s = %+v (style %+v), want a styled spill follower %s", ref, got, got.Style, display)
			}
		}
		checkUndoRedo(t, s, sheet1, before)
	})

	t.Run("formula cells keep their formula and value", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "4", "B1", "=A1*2", "C1", "=B1+1")
		mustStyle(t, s, "B1", `{"italic":true}`)
		cells := cellsByRef(t, s, sheet1, 1, 1, 1, 3)
		if got := cells["B1"]; got.Input != "=A1*2" || got.Display != "8" || !got.Formula {
			t.Errorf("B1 = %+v", got)
		}
		if got := cells["C1"].Display; got != "9" {
			t.Errorf("C1 = %q, want 9", got)
		}
	})
}

func TestSetStyleErrors(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "1")
	before := dump(t, s, sheet1)
	history := s.History()

	// Patch texts that the decoder or the validation must reject.
	badPatches := []string{
		`{"bold":"yes"}`, `{"bold":1}`, `{"color":"red"}`, `{"color":"#FFF"}`, `{"color":"FF0000"}`, `{"color":"#GG0000"}`,
		`{"color":5}`, `{"fill":"#12345"}`, `{"fill":true}`, `{"align":"justify"}`, `{"align":5}`,
		`{"numFmt":"custom"}`, `{"numFmt":"0.00"}`, `{"numFmt":null2}`, `{"border":"thin"}`, `{"size":12}`, `[]`, `"bold"`,
	}
	for _, text := range badPatches {
		patch, err := decodePatch(text)
		if err == nil {
			_, err = s.SetStyle(sheet1, Range{1, 1, 1, 1}, patch)
		}
		if err == nil {
			t.Errorf("patch %s gave no error", text)
		}
	}
	if _, err := s.SetStyle("Nope", Range{1, 1, 1, 1}, StylePatch{Bold: boolPtr(true)}); err == nil {
		t.Error("unknown sheet gave no error")
	}
	if _, err := s.SetStyle(sheet1, Range{0, 1, 1, 1}, StylePatch{Bold: boolPtr(true)}); err == nil {
		t.Error("row 0 gave no error")
	}
	// 201 x 1000 cells that are not full rows or columns.
	if _, err := s.SetStyle(sheet1, Range{1, 1, 201, 1000}, StylePatch{Bold: boolPtr(true)}); err == nil || !strings.Contains(err.Error(), "200000") {
		t.Errorf("range above the cell limit: error = %v", err)
	}
	if got := s.History(); got != history {
		t.Errorf("history changed: %+v", got)
	}
	if got := dump(t, s, sheet1); got != before {
		t.Errorf("the workbook changed")
	}

	// null is "remove" only for the two colors; for the other fields null
	// is "no change", as an absent field is.
	patch := patchFromJSON(t, `{"bold":null,"align":null,"numFmt":null,"color":null}`)
	if patch.Bold != nil || patch.Align != nil || patch.NumFmt != nil || !patch.Color.Set || patch.Fill.Set {
		t.Errorf("patch with nulls = %+v", patch)
	}
	// Validation of fields in a Go value.
	if _, err := s.SetStyle(sheet1, Range{1, 1, 1, 1}, StylePatch{Align: strPtr("top")}); err == nil {
		t.Error("align top gave no error")
	}
}

func TestStylePatchJSONAndCellStyleJSON(t *testing.T) {
	patch := patchFromJSON(t, `{"bold":true,"italic":false,"color":"#aabbcc","fill":null,"align":"general","numFmt":"currency"}`)
	if patch.Bold == nil || !*patch.Bold || patch.Italic == nil || *patch.Italic || patch.Underline != nil {
		t.Errorf("bool fields = %+v", patch)
	}
	if patch.Color != (ColorPatch{Set: true, Value: "AABBCC"}) || patch.Fill != (ColorPatch{Set: true}) {
		t.Errorf("color fields = %+v %+v", patch.Color, patch.Fill)
	}
	if *patch.Align != "general" || *patch.NumFmt != "currency" {
		t.Errorf("align %q numFmt %q", *patch.Align, *patch.NumFmt)
	}

	got, err := json.Marshal(CellStyle{Bold: true, NumFmt: "currency"})
	if err != nil || string(got) != `{"bold":true,"numFmt":"currency"}` {
		t.Errorf("CellStyle JSON = %s, %v", got, err)
	}
	got, _ = json.Marshal(StyleResult{Changed: 2, Range: Range{1, 1, 1, 2}, Clamped: true})
	want := `{"history":{"canUndo":false,"canRedo":false,"undoLabel":"","redoLabel":"","depth":0},"changed":2,"range":{"row1":1,"col1":1,"row2":1,"col2":2},"clamped":true}`
	if string(got) != want {
		t.Errorf("StyleResult JSON = %s\nwant %s", got, want)
	}
}
