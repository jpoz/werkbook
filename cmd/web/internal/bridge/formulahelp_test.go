package bridge

import (
	"encoding/json"
	"sort"
	"strings"
	"testing"
	"unicode/utf16"

	werkbook "github.com/jpoz/werkbook"
	"github.com/jpoz/werkbook/formula"
)

// u16 returns the UTF-16 length of s, the unit of the contract offsets.
func u16(s string) int { return len(utf16.Encode([]rune(s))) }

func TestFunctions(t *testing.T) {
	s := NewSession()
	fns := s.Functions()

	names := make([]string, len(fns))
	byName := map[string]FunctionInfo{}
	for i, f := range fns {
		names[i] = f.Name
		byName[f.Name] = f
		if f.Args == nil {
			t.Errorf("%s: args is nil; the JSON must have an array", f.Name)
		}
		if f.Name != strings.ToUpper(f.Name) {
			t.Errorf("%s is not upper case", f.Name)
		}
	}
	if !sort.StringsAreSorted(names) {
		t.Errorf("the list is not sorted")
	}
	// All registered names are in the list.
	for _, name := range formula.RegisteredFunctions() {
		if _, ok := byName[name]; !ok {
			t.Errorf("registered function %s is not in the list", name)
		}
	}
	if len(fns) < 400 {
		t.Errorf("the list has %d functions, want more than 400", len(fns))
	}

	sum := byName["SUM"]
	if sum.Signature != "SUM(number1, [number2], ...)" || sum.Summary == "" || len(sum.Args) != 2 || !sum.Args[1].Repeat || !sum.Variadic {
		t.Errorf("SUM = %+v", sum)
	}
	sumif := byName["SUMIF"]
	wantArgs := []FunctionArg{{Name: "range"}, {Name: "criteria"}, {Name: "sum_range", Optional: true}}
	if len(sumif.Args) != 3 || sumif.Args[0] != wantArgs[0] || sumif.Args[1] != wantArgs[1] || sumif.Args[2] != wantArgs[2] {
		t.Errorf("SUMIF args = %+v, want %+v", sumif.Args, wantArgs)
	}
	// LET is a form of the parser, not a registered function; a user still
	// types it as a function.
	if let := byName["LET"]; let.Summary == "" || len(let.Args) == 0 {
		t.Errorf("LET = %+v, want an authored record", let)
	}
	// A function with no authored record.
	var plain *FunctionInfo
	for i := range fns {
		if fns[i].Summary == "" {
			plain = &fns[i]
			break
		}
	}
	if plain == nil {
		t.Fatalf("each function has an authored record; this test needs an update")
	}
	if plain.Signature != plain.Name+"(...)" || len(plain.Args) != 0 {
		t.Errorf("function with no record = %+v", *plain)
	}

	got, err := json.Marshal(sumif)
	if err != nil {
		t.Fatal(err)
	}
	want := `{"name":"SUMIF","signature":"SUMIF(range, criteria, [sum_range])","summary":"Adds the cells that meet one condition.",` +
		`"args":[{"name":"range","optional":false,"repeat":false},{"name":"criteria","optional":false,"repeat":false},{"name":"sum_range","optional":true,"repeat":false}],` +
		`"minArgs":2,"variadic":false}`
	if string(got) != want {
		t.Errorf("JSON = %s\nwant   %s", got, want)
	}
}

// The signatures of SUMIF and AVERAGEIF show the last range as optional,
// as Excel does, while the function registry reports three required
// arguments. This test is the proof that the engine accepts both forms.
func TestOptionalRangeOfSumifAndAverageif(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "1", "A2", "5", "A3", "9", "B1", "10", "B2", "20", "B3", "30",
		"D1", `=SUMIF(A1:A3,">2")`, "D2", `=SUMIF(A1:A3,">2",B1:B3)`,
		"D3", `=AVERAGEIF(A1:A3,">2")`, "D4", `=AVERAGEIF(A1:A3,">2",B1:B3)`)
	cells := cellsByRef(t, s, sheet1, 1, 4, 4, 4)
	for ref, want := range map[string]string{"D1": "14", "D2": "50", "D3": "7", "D4": "25"} {
		if got := cells[ref].Display; got != want {
			t.Errorf("%s = %q, want %q", ref, got, want)
		}
	}
}

func TestFormulaContextPrefixAndCall(t *testing.T) {
	s := NewSession()
	tests := []struct {
		name       string
		text       string // "|" marks the caret
		wantPrefix *TextSpan
		wantCall   *CallInfo
	}{
		{"prefix counts the equals sign", "=SU|", &TextSpan{1, 3, "SU"}, nil},
		{"prefix in a call", "=IF(A1,SU|", &TextSpan{7, 9, "SU"}, &CallInfo{"IF", 1, 1}},
		{"second argument", "=SUMIF(A1:A3,|", nil, &CallInfo{"SUMIF", 1, 1}},
		{"nested call", "=IF(SUM(A1,|", nil, &CallInfo{"SUM", 1, 4}},
		{"lower case name", "=sum(|", nil, &CallInfo{"SUM", 0, 1}},
		{"cell address prefix", "=A1|", &TextSpan{1, 3, "A1"}, nil},
		{"in a string", `="SU|`, nil, nil},
		{"only the equals sign", "=|", nil, nil},
		{"caret before the equals sign", "|=SUM(A1", nil, nil},
		{"caret after the equals sign", "=|SUM(A1", nil, nil},
		{"no equals sign", "SU|", nil, nil},
		{"text that looks like a call", "SUM(A1,|", nil, nil},
		{"apostrophe text", "'=SU|", nil, nil},
		{"empty", "|", nil, nil},
		// Offsets are UTF-16 code units: "é" is 2 bytes and 1 unit; the
		// emoji is 4 bytes and 2 units.
		{"two-byte character before the caret", `="é"&SU|`, &TextSpan{5, 7, "SU"}, nil},
		{"surrogate pair before the caret", `="😀"&SU|`, &TextSpan{6, 8, "SU"}, nil},
		{"non-ASCII string before a call", `=IF(A1="日本",SUM(B1,|`, nil, &CallInfo{"SUM", 1, 12}},
		{"surrogate pair before a call", `="😀😀"&LEFT(A1,|`, nil, &CallInfo{"LEFT", 1, 8}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			i := strings.IndexByte(tt.text, '|')
			text := tt.text[:i] + tt.text[i+1:]
			ctx := s.FormulaContext(text, u16(tt.text[:i]))
			switch {
			case (ctx.Prefix == nil) != (tt.wantPrefix == nil):
				t.Errorf("prefix = %+v, want %+v", ctx.Prefix, tt.wantPrefix)
			case ctx.Prefix != nil && *ctx.Prefix != *tt.wantPrefix:
				t.Errorf("prefix = %+v, want %+v", *ctx.Prefix, *tt.wantPrefix)
			}
			switch {
			case (ctx.Call == nil) != (tt.wantCall == nil):
				t.Errorf("call = %+v, want %+v", ctx.Call, tt.wantCall)
			case ctx.Call != nil && *ctx.Call != *tt.wantCall:
				t.Errorf("call = %+v, want %+v", *ctx.Call, *tt.wantCall)
			}
			if ctx.Refs == nil {
				t.Errorf("refs is nil; the JSON must have an array")
			}
			// The prefix span is the prefix text in UTF-16 units.
			if p := ctx.Prefix; p != nil {
				units := utf16.Encode([]rune(text))
				if got := string(utf16.Decode(units[p.Start:p.End])); got != p.Text {
					t.Errorf("text at the prefix span = %q, want %q", got, p.Text)
				}
			}
		})
	}
}

// refWant is the expected form of a RefSpan without its end sheet.
type refWant struct {
	Start, End        int
	Text, Sheet, Kind string
	Range             Range
	Color             int
}

// A 3D reference reports both sheets, so that the UI can draw the outline on
// each sheet of the span.
func TestFormulaContext3DReference(t *testing.T) {
	s := NewSession()
	ctx := s.FormulaContext("=SUM(Sheet1:Sheet3!A1)+SUM('Jan 1:Dec 31'!B2:C3)+Sheet2!D4+E5", 5)
	if len(ctx.Refs) != 4 {
		t.Fatalf("refs = %+v", ctx.Refs)
	}
	want := [][2]string{{"Sheet1", "Sheet3"}, {"Jan 1", "Dec 31"}, {"Sheet2", ""}, {"", ""}}
	for i, w := range want {
		if got := ctx.Refs[i]; got.Sheet != w[0] || got.SheetEnd != w[1] {
			t.Errorf("ref %d: sheet %q sheetEnd %q, want %q %q", i, got.Sheet, got.SheetEnd, w[0], w[1])
		}
	}
	// The same range on a different span of sheets is a different color.
	ctx = s.FormulaContext("=SUM(Sheet1:Sheet3!A1)+SUM(Sheet1:Sheet2!A1)+SUM(sheet1:SHEET3!$A$1)+Sheet1!A1", 0)
	colors := []int{ctx.Refs[0].Color, ctx.Refs[1].Color, ctx.Refs[2].Color, ctx.Refs[3].Color}
	if colors[0] != 0 || colors[1] != 1 || colors[2] != 0 || colors[3] != 2 {
		t.Errorf("colors = %v, want [0 1 0 2]", colors)
	}
	// The JSON has sheetEnd only for a 3D reference.
	encoded, err := json.Marshal(s.FormulaContext("=SUM(Jan:Dec!A1)+B1", 0).Refs)
	if err != nil {
		t.Fatal(err)
	}
	want3D := `[{"start":5,"end":15,"text":"Jan:Dec!A1","sheet":"Jan","sheetEnd":"Dec","kind":"cell","range":{"row1":1,"col1":1,"row2":1,"col2":1},"color":0},` +
		`{"start":17,"end":19,"text":"B1","sheet":"","kind":"cell","range":{"row1":1,"col1":2,"row2":1,"col2":2},"color":1}]`
	if string(encoded) != want3D {
		t.Errorf("JSON = %s\nwant   %s", encoded, want3D)
	}
}

func TestFormulaContextRefs(t *testing.T) {
	s := NewSession()
	full := werkbook.MaxRows
	wide := werkbook.MaxColumns
	tests := []struct {
		name string
		text string
		want []refWant
	}{
		{"no formula", "A1+B2", []refWant{}},
		{"no references", "=1+2", []refWant{}},
		{
			"cell and range",
			"=A1+B2:C3",
			[]refWant{
				{1, 3, "A1", "", "cell", Range{1, 1, 1, 1}, 0},
				{4, 9, "B2:C3", "", "range", Range{2, 2, 3, 3}, 1},
			},
		},
		{
			"equal references have the same color; absolute marks do not count",
			"=A1+$A$1+a1+B1+A1:A1",
			[]refWant{
				{1, 3, "A1", "", "cell", Range{1, 1, 1, 1}, 0},
				{4, 8, "$A$1", "", "cell", Range{1, 1, 1, 1}, 0},
				{9, 11, "a1", "", "cell", Range{1, 1, 1, 1}, 0},
				{12, 14, "B1", "", "cell", Range{1, 2, 1, 2}, 1},
				{15, 20, "A1:A1", "", "range", Range{1, 1, 1, 1}, 0},
			},
		},
		{
			"sheet is part of the color key, with no letter case",
			"=A1+Sheet2!A1+sheet2!A1+'My Sheet'!A1",
			[]refWant{
				{1, 3, "A1", "", "cell", Range{1, 1, 1, 1}, 0},
				{4, 13, "Sheet2!A1", "Sheet2", "cell", Range{1, 1, 1, 1}, 1},
				{14, 23, "sheet2!A1", "sheet2", "cell", Range{1, 1, 1, 1}, 1},
				{24, 37, "'My Sheet'!A1", "My Sheet", "cell", Range{1, 1, 1, 1}, 2},
			},
		},
		{
			"full columns and full rows",
			"=SUM(B:C)+SUM(2:3)",
			[]refWant{
				{5, 8, "B:C", "", "cols", Range{1, 2, full, 3}, 0},
				{14, 17, "2:3", "", "rows", Range{2, 1, 3, wide}, 1},
			},
		},
		{
			"reversed range is normalized",
			"=SUM(C3:A1)",
			[]refWant{{5, 10, "C3:A1", "", "range", Range{1, 1, 3, 3}, 0}},
		},
		{
			"incomplete formula gives the complete references",
			"=SUM(A1:",
			[]refWant{{5, 7, "A1", "", "cell", Range{1, 1, 1, 1}, 0}},
		},
		{
			"strings, names and structured references are not references",
			`="A1"&Total&Table1[B2]&C3`,
			[]refWant{{23, 25, "C3", "", "cell", Range{3, 3, 3, 3}, 0}},
		},
		{
			"offsets after text that is not ASCII are UTF-16 units",
			`="é😀"&A1+B2`,
			[]refWant{
				{7, 9, "A1", "", "cell", Range{1, 1, 1, 1}, 0},
				{10, 12, "B2", "", "cell", Range{2, 2, 2, 2}, 1},
			},
		},
		{
			"3D reference reports its first sheet",
			"=SUM(Jan:Dec!B2)",
			[]refWant{{5, 15, "Jan:Dec!B2", "Jan", "cell", Range{2, 2, 2, 2}, 0}},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctx := s.FormulaContext(tt.text, 0)
			if len(ctx.Refs) != len(tt.want) {
				t.Fatalf("refs = %+v, want %+v", ctx.Refs, tt.want)
			}
			units := utf16.Encode([]rune(tt.text))
			for i, w := range tt.want {
				got := ctx.Refs[i]
				if gotWant := (refWant{got.Start, got.End, got.Text, got.Sheet, got.Kind, got.Range, got.Color}); gotWant != w {
					t.Errorf("ref %d = %+v, want %+v", i, got, w)
				}
				if text := string(utf16.Decode(units[got.Start:got.End])); text != got.Text {
					t.Errorf("ref %d: text at the span = %q, want %q", i, text, got.Text)
				}
			}
		})
	}

	// The caret does not change the references.
	text := "=SUM(A1:B2,C3)"
	for caret := -1; caret <= u16(text)+1; caret++ {
		if got := s.FormulaContext(text, caret); len(got.Refs) != 2 {
			t.Errorf("caret %d: %d refs, want 2", caret, len(got.Refs))
		}
	}
}

func TestFormulaContextJSONShape(t *testing.T) {
	s := NewSession()
	got, err := json.Marshal(s.FormulaContext("=SUMIF(A1:A3,", 13))
	if err != nil {
		t.Fatal(err)
	}
	want := `{"prefix":null,"call":{"name":"SUMIF","argIndex":1,"nameStart":1},` +
		`"refs":[{"start":7,"end":12,"text":"A1:A3","sheet":"","kind":"range","range":{"row1":1,"col1":1,"row2":3,"col2":1},"color":0}]}`
	if string(got) != want {
		t.Errorf("JSON = %s\nwant   %s", got, want)
	}
	got, _ = json.Marshal(s.FormulaContext("text", 2))
	if string(got) != `{"prefix":null,"call":null,"refs":[]}` {
		t.Errorf("JSON of an empty context = %s", got)
	}
	got, _ = json.Marshal(s.FormulaContext("=SU", 3))
	if string(got) != `{"prefix":{"start":1,"end":3,"text":"SU"},"call":null,"refs":[]}` {
		t.Errorf("JSON with a prefix = %s", got)
	}
}

func TestCycleRef(t *testing.T) {
	s := NewSession()

	// Four states and back, with the caret after the reference.
	text := "=A1+B2"
	caret := 3
	for _, want := range []string{"=$A$1+B2", "=A$1+B2", "=$A1+B2", "=A1+B2"} {
		res := s.CycleRef(text, caret, caret)
		if !res.Changed || res.Text != want {
			t.Fatalf("CycleRef(%q, %d) = %+v, want text %q", text, caret, res, want)
		}
		if res.SelStart != res.SelEnd || res.SelEnd != strings.Index(want, "+") {
			t.Errorf("caret after %q = [%d,%d], want the end of the reference", want, res.SelStart, res.SelEnd)
		}
		text, caret = res.Text, res.SelEnd
	}

	tests := []struct {
		name               string
		text               string
		selStart, selEnd   int
		want               string
		wantStart, wantEnd int
		wantChanged        bool
	}{
		{"range", "=SUM(B2:C3)", 8, 8, "=SUM($B$2:$C$3)", 14, 14, true},
		{"columns", "=SUM(A:A)", 6, 6, "=SUM($A:$A)", 10, 10, true},
		{"rows", "=SUM(1:3)", 6, 6, "=SUM($1:$3)", 10, 10, true},
		{"rows back", "=SUM($1:$3)", 6, 6, "=SUM(1:3)", 8, 8, true},
		{"selection over two references", "=A1+B2", 1, 6, "=$A$1+$B$2", 1, 10, true},
		{"reversed selection", "=A1+B2", 6, 1, "=$A$1+$B$2", 1, 10, true},
		{"caret at the reference start", "=A1", 1, 1, "=$A$1", 5, 5, true},
		{"caret before the equals sign", "=A1", 0, 0, "=$A$1", 5, 5, true},
		{"no reference at the caret", "=1+A1", 2, 2, "=1+A1", 2, 2, false},
		{"no equals sign", "A1", 2, 2, "A1", 2, 2, false},
		{"empty", "", 0, 0, "", 0, 0, false},
		{"only the equals sign", "=", 1, 1, "=", 1, 1, false},
		{"reference in a string", `="A1"`, 4, 4, `="A1"`, 4, 4, false},
		// UTF-16 offsets: `="😀"&` is 6 units, then A1.
		{"after a surrogate pair", `="😀"&A1`, 8, 8, `="😀"&$A$1`, 10, 10, true},
		{"after a two-byte character", `="é"&A1`, 7, 7, `="é"&$A$1`, 9, 9, true},
		{"sheet qualified", "=Sheet2!A1", 10, 10, "=Sheet2!$A$1", 12, 12, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			res := s.CycleRef(tt.text, tt.selStart, tt.selEnd)
			want := CycleResult{Text: tt.want, SelStart: tt.wantStart, SelEnd: tt.wantEnd, Changed: tt.wantChanged}
			if res != want {
				t.Errorf("CycleRef(%q, %d, %d) = %+v, want %+v", tt.text, tt.selStart, tt.selEnd, res, want)
			}
		})
	}

	// The cycled text is a formula that the session accepts.
	mustSet(t, s, sheet1, "A1", "4", "B1", s.CycleRef("=A1*2", 3, 3).Text)
	if got := cellsByRef(t, s, sheet1, 1, 2, 1, 2)["B1"]; got.Input != "=$A$1*2" || got.Display != "8" {
		t.Errorf("B1 = %+v", got)
	}
}

func TestUTF16Offsets(t *testing.T) {
	s := "aé\U0001F600b" // a, é (2 bytes), emoji (4 bytes, 2 units), b
	tests := []struct{ units, bytes int }{
		{0, 0}, {1, 1}, {2, 3}, {4, 7}, {5, 8},
	}
	for _, tt := range tests {
		if got := utf16ToByte(s, tt.units); got != tt.bytes {
			t.Errorf("utf16ToByte(%d) = %d, want %d", tt.units, got, tt.bytes)
		}
		if got := byteToUTF16(s, tt.bytes); got != tt.units {
			t.Errorf("byteToUTF16(%d) = %d, want %d", tt.bytes, got, tt.units)
		}
	}
	// Out of the text, and in the middle of a surrogate pair.
	if got := utf16ToByte(s, -3); got != 0 {
		t.Errorf("negative offset = %d", got)
	}
	if got := utf16ToByte(s, 99); got != len(s) {
		t.Errorf("offset past the end = %d", got)
	}
	if got := utf16ToByte(s, 3); got != 3 {
		t.Errorf("offset in a surrogate pair = %d, want the start of the character (3)", got)
	}
	if got := byteToUTF16(s, 99); got != 5 {
		t.Errorf("byte offset past the end = %d", got)
	}
}
