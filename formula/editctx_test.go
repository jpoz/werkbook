package formula

import (
	"fmt"
	"strings"
	"testing"
)

// The caret position in a test text is the '|' character.
func splitCaret(t *testing.T, text string) (string, int) {
	t.Helper()
	i := strings.IndexByte(text, '|')
	if i < 0 {
		t.Fatalf("test text %q has no caret mark", text)
	}
	return text[:i] + text[i+1:], i
}

func describeEdit(info EditInfo) (prefix, call string) {
	prefix, call = "-", "-"
	if info.Prefix != nil {
		prefix = fmt.Sprintf("%s@%d-%d", info.Prefix.Text, info.Prefix.Start, info.Prefix.End)
	}
	if info.Call != nil {
		call = fmt.Sprintf("%s#%d@%d", info.Call.Name, info.Call.ArgIndex, info.Call.NameStart)
	}
	return prefix, call
}

func TestEditContext(t *testing.T) {
	tests := []struct {
		name       string
		text       string // with the caret mark
		wantPrefix string // "text@start-end" or "-"
		wantCall   string // "NAME#argIndex@nameStart" or "-"
	}{
		// Prefix.
		{"empty", "|", "-", "-"},
		{"prefix at the start", "SU|", "SU@0-2", "-"},
		{"one letter", "S|", "S@0-1", "-"},
		{"lower case", "su|", "su@0-2", "-"},
		{"caret at 0 before text", "|SUM", "-", "-"},
		{"after plus", "1+SU|", "SU@2-4", "-"},
		{"after minus", "-SU|", "SU@1-3", "-"},
		{"after each operator", "1*A|", "A@2-3", "-"},
		{"after divide", "1/A|", "A@2-3", "-"},
		{"after power", "1^A|", "A@2-3", "-"},
		{"after ampersand", "\"x\"&LE|", "LE@4-6", "-"},
		{"after comparison", "A1<>SU|", "SU@4-6", "-"},
		{"after less or equal", "A1<=SU|", "SU@4-6", "-"},
		{"after equals", "A1=SU|", "SU@3-5", "-"},
		{"after a space", "1 + SU|", "SU@4-6", "-"},
		{"after open parenthesis", "(SU|", "SU@1-3", "-"},
		{"after function parenthesis", "IF(SU|", "SU@3-5", "IF#0@0"},
		{"after a comma", "IF(A1,SU|", "SU@6-8", "IF#1@0"},
		{"after a comma and a space", "IF(A1, SU|", "SU@7-9", "IF#1@0"},
		{"in an array literal", "{1,SU|", "SU@3-5", "-"},
		{"name with a dot", "STDEV.S|", "STDEV.S@0-7", "-"},
		{"name with a dot at the end", "CEILING.|", "CEILING.@0-8", "-"},
		{"name with a digit", "LOG1|", "LOG1@0-4", "-"},
		{"name with an underscore", "MY_F|", "MY_F@0-4", "-"},
		{"cell address is a prefix (no function matches it)", "A1|", "A1@0-2", "-"},
		{"two-letter cell address", "AB12|", "AB12@0-4", "-"},
		{"prefix before more text", "SU|+1", "SU@0-2", "-"},
		{"prefix before a closing parenthesis", "IF(SU|)", "SU@3-5", "IF#0@0"},
		{"prefix before a comma", "IF(SU|,1)", "SU@3-5", "IF#0@0"},

		// No prefix.
		{"in a string", "\"SU|", "-", "-"},
		{"in a closed string", "\"SU|M\"", "-", "-"},
		{"at the closing quote of a string", "\"SU|\"", "-", "-"},
		{"after a string", "\"SU\"|", "-", "-"},
		{"in a string in a call", "IF(A1=\"SU|", "-", "IF#0@0"},
		{"after a string with doubled quotes", "\"a\"\"b\"&SU|", "SU@7-9", "-"},
		{"in a string with doubled quotes", "\"a\"\"SU|", "-", "-"},
		{"in the middle of a name", "SU|M", "-", "-"},
		{"in the middle of a name before a digit", "LOG|10", "-", "-"},
		{"at the end of a complete function name", "SUM|(", "-", "-"},
		{"after a number", "12|", "-", "-"},
		{"after a number with an exponent", "1E|", "-", "-"},
		{"after a number then letters", "2ab|", "-", "-"},
		{"after a decimal", "1.5|", "-", "-"},
		{"after an operator", "1+|", "-", "-"},
		{"after a parenthesis", "SUM(|", "-", "SUM#0@0"},
		{"absolute reference", "$A|", "-", "-"},
		{"absolute row", "A$|", "-", "-"},
		{"after a sheet qualifier", "Sheet1!A|", "-", "-"},
		{"sheet name before the bang", "Sheet1|!A1", "-", "-"},
		{"in a quoted sheet name", "'My SU|", "-", "-"},
		{"in a closed quoted sheet name", "'My SU|M'!A1", "-", "-"},
		{"after a quoted sheet qualifier", "'My Sheet'!A|", "-", "-"},
		{"in a structured reference", "Table1[Co|", "-", "-"},
		{"in a closed structured reference", "Table1[Co|l]", "-", "-"},
		{"name that is not ASCII", "Grö|", "-", "-"},
		{"name that starts with an underscore", "_x|", "-", "-"},
		{"name with a backslash", `a\b|`, "-", "-"},
		{"error literal", "#N/A|", "-", "-"},

		// Call and argument index.
		{"first argument", "SUM(|", "-", "SUM#0@0"},
		{"first argument with text", "SUM(A1|", "A1@4-6", "SUM#0@0"},
		{"second argument", "SUM(A1,|", "-", "SUM#1@0"},
		{"third argument", "IF(A1>0,\"y\",|", "-", "IF#2@0"},
		{"empty arguments", "IF(,,|", "-", "IF#2@0"},
		{"name in lower case is reported in upper case", "sumif(a1:a3,|", "-", "SUMIF#1@0"},
		{"call after other text", "1+SUM(A1,|", "-", "SUM#1@2"},
		{"nested call: the innermost open call", "IF(SUM(A1,|", "-", "SUM#1@3"},
		{"nested call that is closed", "IF(SUM(A1,B1),|", "-", "IF#1@0"},
		{"nested call that is closed, later argument", "IF(SUM(A1,B1),MAX(1,2,3),|", "-", "IF#2@0"},
		{"three levels", "IF(A1,SUM(B1,MAX(C1,D1,|", "-", "MAX#2@13"},
		{"commas in an array literal do not count", "SUM({1,2,3},|", "-", "SUM#1@0"},
		{"caret in an array literal", "SUM({1,2|", "-", "SUM#0@0"},
		{"array literal with rows", "SUM({1,2;3,4},A1,|", "-", "SUM#2@0"},
		{"commas in a string do not count", "IF(A1=\"a,b,c\",|", "-", "IF#1@0"},
		{"commas in plain parentheses do not count", "SUM((A1,B1),|", "-", "SUM#1@0"},
		{"caret in plain parentheses in a call", "SUM((A1,B1|", "B1@8-10", "SUM#0@0"},
		{"plain parentheses with no call", "(1+2,|", "-", "-"},
		{"commas in a structured reference do not count", "SUM(Table1[[#This Row],[Col]],|", "-", "SUM#1@0"},
		{"call is closed", "SUM(A1)|", "-", "-"},
		{"call is closed, then text", "SUM(A1)+|", "-", "-"},
		{"caret in the function name is not in the call", "SU|M(A1,B1)", "-", "-"},
		{"caret at the open parenthesis", "SUM|(A1,B1)", "-", "-"},
		{"caret in the middle of a complete call", "SUM(A1,|B1,C1)", "-", "SUM#1@0"},
		{"more closing than opening parentheses", "1)+SUM(A1,|", "-", "SUM#1@3"},
		{"function with a dot", "STDEV.S(A1,|", "-", "STDEV.S#1@0"},
		{"function with digits", "LOG10(|", "-", "LOG10#0@0"},
		// A text argument keeps the hint of its call (the review of M2).
		{"caret in a text argument keeps the call", "TEXT(A1,\"0.0|", "-", "TEXT#1@0"},
		{"caret in the first text argument", "VLOOKUP(\"ab|", "-", "VLOOKUP#0@0"},
		{"caret in a closed text argument", "IF(A1=\"x|\",1,2)", "-", "IF#0@0"},
		{"caret at the opening quote of a text argument", "IF(A1=\"|", "-", "IF#0@0"},
		{"caret after the closing quote", "IF(A1=\"x\"|", "-", "IF#0@0"},
		{"commas in the text before the caret do not count", "IF(A1,\"a,b,c|", "-", "IF#1@0"},
		{"text argument in a nested call", "IF(LEFT(\"ab|", "-", "LEFT#0@3"},
		{"text with doubled quotes", "IF(A1,\"say \"\"hi\"\", ok|", "-", "IF#1@0"},
		{"string with no call", "\"abc|", "-", "-"},
		{"caret in a quoted sheet name keeps the call", "SUM('My Sh|", "-", "SUM#0@0"},
		{"caret in a structured reference keeps the call", "SUM(A1,Table1[Co|", "-", "SUM#1@0"},

		// Text that is not ASCII before the caret: offsets are bytes.
		{"non-ASCII string before the caret", "\"é\"&SU|", "SU@5-7", "-"},
		{"non-ASCII string before a call", "IF(A1=\"日本\",SU|", "SU@15-17", "IF#1@0"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			src, caret := splitCaret(t, tt.text)
			prefix, call := describeEdit(EditContext(src, caret))
			if prefix != tt.wantPrefix {
				t.Errorf("EditContext(%q, %d) prefix = %s, want %s", src, caret, prefix, tt.wantPrefix)
			}
			if call != tt.wantCall {
				t.Errorf("EditContext(%q, %d) call = %s, want %s", src, caret, call, tt.wantCall)
			}
		})
	}
}

func TestEditContextRefsAndCaretLimits(t *testing.T) {
	// Refs are the complete references of the full text, not only the part
	// before the caret.
	info := EditContext("SUM(A1:B2,Sheet2!C3)+\"D4\"+E", 4)
	if got := describeSites("SUM(A1:B2,Sheet2!C3)+\"D4\"+E"); len(info.Refs) != 2 || len(got) != 2 {
		t.Fatalf("refs = %v", got)
	}
	// An incomplete range gives the references that are complete.
	info = EditContext("SUM(A1:", 7)
	if len(info.Refs) != 1 || info.Refs[0].Kind != RefCell || info.Refs[0].From != (RefPoint{Col: 1, Row: 1}) {
		t.Errorf("refs of SUM(A1: = %+v, want A1", info.Refs)
	}
	if info.Call == nil || info.Call.Name != "SUM" {
		t.Errorf("call = %+v", info.Call)
	}

	// A caret out of the text is moved to the nearest end.
	if p, c := describeEdit(EditContext("SUM(A1,SU", 99)); p != "SU@7-9" || c != "SUM#1@0" {
		t.Errorf("caret past the end: %s %s", p, c)
	}
	if p, c := describeEdit(EditContext("SUM(A1,SU", -5)); p != "-" || c != "-" {
		t.Errorf("negative caret: %s %s", p, c)
	}
	// It never fails on text that is not a formula.
	for _, src := range append([]string{"'", "\"", "[", "((((", "))))", "{{{", "}}}", ",,,", "$", "!"}, realFormulas...) {
		for caret := 0; caret <= len(src); caret++ {
			_ = EditContext(src, caret)
		}
	}
}

func TestCycleAbsolute(t *testing.T) {
	tests := []struct {
		name string
		text string // "[" and "]" mark the selection; "|" marks a caret
		want string // same marks for the new selection
	}{
		// Four states and back.
		{"relative to absolute", "A1|", "$A$1|"},
		{"absolute to absolute row", "$A$1|", "A$1|"},
		{"absolute row to absolute column", "A$1|", "$A1|"},
		{"absolute column to relative", "$A1|", "A1|"},

		// Caret positions.
		{"caret in the reference", "A|1+B2", "$A$1|+B2"},
		{"caret at the start of the reference", "|A1+B2", "$A$1|+B2"},
		{"caret immediately after the reference", "A1|+B2", "$A$1|+B2"},
		{"caret at the second reference", "A1+B|2", "A1+$B$2|"},
		{"caret at the start of the second reference", "A1+|B2", "A1+$B$2|"},
		{"caret in a call", "SUM(A1|,B2)", "SUM($A$1|,B2)"},
		{"lower case is written in upper case", "a1|", "$A$1|"},

		// Ranges cycle both ends, from the state of the first end.
		{"range", "SUM(A1:B2|)", "SUM($A$1:$B$2|)"},
		{"range second state", "SUM($A$1:$B$2|)", "SUM(A$1:B$2|)"},
		{"range third state", "SUM(A$1:B$2|)", "SUM($A1:$B2|)"},
		{"range fourth state", "SUM($A1:$B2|)", "SUM(A1:B2|)"},
		{"caret in the first end of a range", "SUM(A|1:B2)", "SUM($A$1:$B$2|)"},
		{"range with different ends follows the first end", "$A$1:B2|", "A$1:B$2|"},
		{"range with a mixed first end", "A$1:$B$2|", "$A1:$B2|"},

		// Column-only and row-only references.
		{"columns", "SUM(A:A|)", "SUM($A:$A|)"},
		{"columns back", "SUM($A:$A|)", "SUM(A:A|)"},
		{"column range", "A:C|", "$A:$C|"},
		{"half absolute columns follow the first end", "$A:C|", "A:C|"},
		{"rows", "SUM(1:3|)", "SUM($1:$3|)"},
		{"rows back", "SUM($1:$3|)", "SUM(1:3|)"},

		// Sheet qualifiers stay.
		{"sheet qualified", "Sheet2!A1|", "Sheet2!$A$1|"},
		{"quoted sheet", "'My Sheet'!A1:B2|", "'My Sheet'!$A$1:$B$2|"},
		{"caret in the sheet name", "She|et2!A1", "Sheet2!$A$1|"},
		{"sheet rows", "Data!1:2|", "Data!$1:$2|"},

		// Selections.
		{"selection over one reference", "[A1]+B2", "[$A$1]+B2"},
		{"selection over two references", "[A1+B2]", "[$A$1+$B$2]"},
		{"selection that touches two references in part", "A[1+B]2+C3", "[$A$1+$B$2]+C3"},
		{"selection over all", "[SUM(A1,$B$2,C$3)]", "SUM([$A$1,B$2,$C3])"},
		{"each reference cycles from its own state", "[A1+$A$1+A$1+$A1]", "[$A$1+A$1+$A1+A1]"},
		{"selection with no reference", "[1+2]+A1", "[1+2]+A1"},
		{"selection that ends at a reference start", "[1+]A1", "1+[$A$1]"},

		// Nothing to change.
		{"no reference at the caret", "A1+1|+B2", "A1+1|+B2"},
		{"caret in a number", "1|2", "1|2"},
		{"caret after a function name", "SUM|(A1)", "SUM|(A1)"},
		{"reference in a string", "\"A1|\"", "\"A1|\""},
		{"name", "Total|", "Total|"},
		{"structured reference", "Table1[A1|]", "Table1[A1|]"},
		{"empty", "|", "|"},

		// Other text stays.
		{"spaces and case of other text stay", " sum( a1| , b2 ) ", " sum( $A$1| , b2 ) "},
		{"spill operator stays", "A1|#", "$A$1|#"},
		{"text that is not ASCII before the reference", "\"é\"&A1|", "\"é\"&$A$1|"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			src, selStart, selEnd := splitSelection(t, tt.text)
			wantSrc, wantStart, wantEnd := splitSelection(t, tt.want)
			out, newStart, newEnd, changed := CycleAbsolute(src, selStart, selEnd)
			if out != wantSrc || newStart != wantStart || newEnd != wantEnd {
				t.Errorf("CycleAbsolute(%q, %d, %d) = %q [%d,%d], want %q [%d,%d]",
					src, selStart, selEnd, out, newStart, newEnd, wantSrc, wantStart, wantEnd)
			}
			if changed != (out != src) && !strings.EqualFold(out, src) {
				t.Errorf("changed = %v for %q -> %q", changed, src, out)
			}
			if wantSrc == src && changed {
				t.Errorf("changed = true but the text is the same")
			}
		})
	}
}

func splitSelection(t *testing.T, text string) (src string, start, end int) {
	t.Helper()
	if i := strings.IndexByte(text, '|'); i >= 0 {
		return text[:i] + text[i+1:], i, i
	}
	i := strings.IndexByte(text, '[')
	j := strings.LastIndexByte(text, ']')
	if i < 0 || j < i {
		t.Fatalf("test text %q has no selection marks", text)
	}
	return text[:i] + text[i+1:j] + text[j+1:], i, j - 1
}

func TestCycleAbsoluteFullCycleAndLimits(t *testing.T) {
	// Four presses give the first text again (in upper case).
	for _, src := range []string{"A1", "SUM(A1:B2)", "Sheet2!C3", "A1+B2*C3"} {
		text := src
		start, end := 0, len(src)
		seen := map[string]bool{}
		for i := 0; i < 4; i++ {
			var changed bool
			text, start, end, changed = CycleAbsolute(text, start, end)
			if !changed {
				t.Fatalf("%q: press %d changed nothing", src, i+1)
			}
			if seen[text] {
				t.Errorf("%q: state %q occurred two times", src, text)
			}
			seen[text] = true
			if _, err := Parse(text); err != nil {
				t.Errorf("%q: state %q does not parse: %v", src, text, err)
			}
		}
		if text != src {
			t.Errorf("%q after four presses = %q", src, text)
		}
	}
	// Two presses for column-only and row-only references.
	for _, src := range []string{"A:C", "SUM(2:5)"} {
		once, _, _, _ := CycleAbsolute(src, 0, len(src))
		twice, _, _, _ := CycleAbsolute(once, 0, len(once))
		if once == src || twice != src {
			t.Errorf("%q -> %q -> %q, want a cycle of two", src, once, twice)
		}
	}
	// A selection with reversed ends and with ends out of the text.
	if out, s, e, ok := CycleAbsolute("A1+B2", 99, -3); !ok || out != "$A$1+$B$2" || s != 0 || e != 9 {
		t.Errorf("reversed selection = %q [%d,%d] %v", out, s, e, ok)
	}
	for _, src := range realFormulas {
		for caret := 0; caret <= len(src); caret++ {
			_, _, _, _ = CycleAbsolute(src, caret, caret)
		}
	}
}
