package bridge

import (
	"reflect"
	"testing"
)

func TestParseTSV(t *testing.T) {
	tests := []struct {
		name string
		text string
		want [][]string
	}{
		{"empty text is one empty field", "", [][]string{{""}}},
		{"one field", "a", [][]string{{"a"}}},
		{"one row", "a\tb\tc", [][]string{{"a", "b", "c"}}},
		{"rows with LF", "a\tb\nc\td", [][]string{{"a", "b"}, {"c", "d"}}},
		{"rows with CRLF", "a\tb\r\nc\td\r\n", [][]string{{"a", "b"}, {"c", "d"}}},
		{"rows with CR", "a\rb\r", [][]string{{"a"}, {"b"}}},
		{"one line end at the end is not a row", "a\n", [][]string{{"a"}}},
		{"two line ends at the end give one empty row", "a\n\n", [][]string{{"a"}, {""}}},
		{"only a line end", "\n", [][]string{{""}}},
		{"empty fields", "\t\tx\n\t", [][]string{{"", "", "x"}, {"", ""}}},
		{"tab at the end", "a\t", [][]string{{"a", ""}}},
		{"empty row in the middle", "a\n\nb", [][]string{{"a"}, {""}, {"b"}}},
		{"ragged rows", "a\tb\tc\nd", [][]string{{"a", "b", "c"}, {"d"}}},
		{"quoted field with a tab", "\"x\ty\"\tz", [][]string{{"x\ty", "z"}}},
		{"quoted field with a line end", "\"line1\nline2\"\tz\nnext", [][]string{{"line1\nline2", "z"}, {"next"}}},
		{"quoted field with CRLF inside", "\"a\r\nb\"", [][]string{{"a\r\nb"}}},
		{"doubled quotes", `"say ""hi"""`, [][]string{{`say "hi"`}}},
		{"empty quoted field", "\"\"\tb", [][]string{{"", "b"}}},
		{"quote in the middle is literal", `5" pipe` + "\t" + `a"b"c`, [][]string{{`5" pipe`, `a"b"c`}}},
		{"text after a closing quote stays", `"ab"cd`, [][]string{{"abcd"}}},
		{"quoted field with no end", "\"abc\td\ne", [][]string{{"abc\td\ne"}}},
		{"spaces stay", " a \t b ", [][]string{{" a ", " b "}}},
		{"text that is not ASCII", "é\tß\n日本", [][]string{{"é", "ß"}, {"日本"}}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := parseTSV(tt.text); !reflect.DeepEqual(got, tt.want) {
				t.Errorf("parseTSV(%q) = %q, want %q", tt.text, got, tt.want)
			}
		})
	}
}

func TestWriteTSV(t *testing.T) {
	tests := []struct {
		name string
		rows [][]string
		want string
	}{
		{"one cell", [][]string{{"a"}}, "a"},
		{"one empty cell", [][]string{{""}}, ""},
		{"grid", [][]string{{"a", "b"}, {"c", "d"}}, "a\tb\nc\td"},
		{"empty cells", [][]string{{"", "x", ""}, {"", "", ""}}, "\tx\t\n\t\t"},
		{"tab", [][]string{{"x\ty", "z"}}, "\"x\ty\"\tz"},
		{"line end", [][]string{{"a\nb"}}, "\"a\nb\""},
		{"carriage return", [][]string{{"a\rb"}}, "\"a\rb\""},
		{"quote", [][]string{{`say "hi"`}}, `"say ""hi"""`},
		{"no rows", nil, ""},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := writeTSV(tt.rows); got != tt.want {
				t.Errorf("writeTSV(%q) = %q, want %q", tt.rows, got, tt.want)
			}
		})
	}
}

func TestTSVRoundTrip(t *testing.T) {
	grids := [][][]string{
		{{"a"}},
		{{"a", "b"}, {"c", "d"}},
		{{"", ""}, {"", "x"}},
		{{"tab\there", "line\nbreak", `quote"mid`}, {`"start`, "cr\r\nlf", ""}},
		{{" spaces ", "'apostrophe", "=1+1"}, {"TRUE", "1.5", "#N/A"}},
		{{`""`, `"`, "\t"}},
	}
	for _, grid := range grids {
		text := writeTSV(grid)
		if got := parseTSV(text); !reflect.DeepEqual(got, grid) {
			t.Errorf("round trip of %q through %q = %q", grid, text, got)
		}
		// A different line end style gives the same grid, except in quoted fields.
		if got := parseTSV(text + "\n"); !reflect.DeepEqual(got, grid) {
			t.Errorf("round trip with a line end at the end of %q = %q", grid, got)
		}
	}
}

func TestNormalizeClipboardText(t *testing.T) {
	tests := []struct{ a, b string }{
		{"a\tb\nc\td", "a\tb\r\nc\td\r\n"},
		{"a\nb", "a\rb\r"},
		{"a", "a\n"},
		{"", "\n"},
		{"\"x\ny\"", "\"x\r\ny\""},
	}
	for _, tt := range tests {
		if normalizeClipboardText(tt.a) != normalizeClipboardText(tt.b) {
			t.Errorf("%q and %q are not equal after normalization", tt.a, tt.b)
		}
	}
	// Only one line end at the end is ignored.
	if normalizeClipboardText("a\n\n") == normalizeClipboardText("a") {
		t.Errorf("an empty last row was ignored")
	}
	if normalizeClipboardText("a\tb") == normalizeClipboardText("a b") {
		t.Errorf("different texts are equal")
	}
}
