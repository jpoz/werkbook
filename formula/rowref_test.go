package formula

import (
	"strings"
	"testing"
)

// Absolute row-only references ($1:$3) are the form that Excel writes for
// print titles and that F4 gives for a row range.

func TestLexAbsoluteRowOnlyReferences(t *testing.T) {
	tests := []struct {
		src  string
		want string // "Type:Value" of each token, without EOF
	}{
		{"$1:$3", "CellRef:$1 Colon:: CellRef:$3"},
		{"1:$3", "Number:1 Colon:: CellRef:$3"},
		{"$1:3", "CellRef:$1 Colon:: Number:3"},
		{"Sheet1!$1:$3", "CellRef:Sheet1!$1 Colon:: CellRef:$3"},
		{"'My Sheet'!$1:$3", "CellRef:'My Sheet'!$1 Colon:: CellRef:$3"},
		{"SUM($1:$3)", "Func:SUM( CellRef:$1 Colon:: CellRef:$3 RParen:)"},
		{"$10:$1048576", "CellRef:$10 Colon:: CellRef:$1048576"},
		// Forms that must not change.
		{"$A$1", "CellRef:$A$1"},
		{"A$1", "CellRef:A$1"},
		{"$A1", "CellRef:$A1"},
		{"$A:$A", "CellRef:$A Colon:: CellRef:$A"},
		{"1:3", "Number:1 Colon:: Number:3"},
		{"LOG10(1)", "Func:LOG10( Number:1 RParen:)"},
		{"TRUE", "Bool:TRUE"},
		{"A1*$B$2", "CellRef:A1 Op:* CellRef:$B$2"},
	}
	for _, tt := range tests {
		tokens, err := Tokenize(tt.src)
		if err != nil {
			t.Errorf("Tokenize(%q): %v", tt.src, err)
			continue
		}
		var parts []string
		for _, tok := range tokens {
			if tok.Type != TokEOF {
				parts = append(parts, tok.Type.String()+":"+tok.Value)
			}
		}
		if got := strings.Join(parts, " "); got != tt.want {
			t.Errorf("Tokenize(%q) = %s, want %s", tt.src, got, tt.want)
		}
	}

	for _, src := range []string{"$", "$+1", "$1a", "$1$", "$$1", "1+$"} {
		if _, err := Tokenize(src); err == nil {
			t.Errorf("Tokenize(%q) gave no error", src)
		}
	}
}

func TestParseCellRefTokenAbsoluteRow(t *testing.T) {
	tests := []struct {
		raw  string
		want CellRef
	}{
		{"$3", CellRef{Row: 3, AbsRow: true}},
		{"Sheet1!$3", CellRef{Sheet: "Sheet1", Row: 3, AbsRow: true}},
		{"'My Sheet'!$12", CellRef{Sheet: "My Sheet", Row: 12, AbsRow: true}},
		{"Sheet1!3", CellRef{Sheet: "Sheet1", Row: 3}},
		// Forms that must not change.
		{"$A$1", CellRef{Col: 1, Row: 1, AbsCol: true, AbsRow: true}},
		{"A$1", CellRef{Col: 1, Row: 1, AbsRow: true}},
		{"$A1", CellRef{Col: 1, Row: 1, AbsCol: true}},
		{"$A", CellRef{Col: 1, AbsCol: true}},
		{"Sheet1!$B", CellRef{Sheet: "Sheet1", Col: 2, AbsCol: true}},
	}
	for _, tt := range tests {
		got, err := parseCellRefToken(tt.raw)
		if err != nil {
			t.Errorf("parseCellRefToken(%q): %v", tt.raw, err)
			continue
		}
		if *got != tt.want {
			t.Errorf("parseCellRefToken(%q) = %+v, want %+v", tt.raw, *got, tt.want)
		}
	}
	for _, raw := range []string{"$", "$0", "$1048577", "Sheet1!$", "Sheet1!$0", "$1x", "3"} {
		if got, err := parseCellRefToken(raw); err == nil {
			t.Errorf("parseCellRefToken(%q) = %+v, want an error", raw, *got)
		}
	}
}

func TestParseAbsoluteRowRanges(t *testing.T) {
	type end struct {
		row    int
		absRow bool
	}
	tests := []struct {
		src      string
		sheet    string
		from, to end
	}{
		{"$1:$3", "", end{1, true}, end{3, true}},
		{"1:$3", "", end{1, false}, end{3, true}},
		{"$1:3", "", end{1, true}, end{3, false}},
		{"1:3", "", end{1, false}, end{3, false}},
		{"Sheet1!$1:$3", "Sheet1", end{1, true}, end{3, true}},
		{"Sheet1!1:$3", "Sheet1", end{1, false}, end{3, true}},
		{"Sheet1!$1:3", "Sheet1", end{1, true}, end{3, false}},
		{"'My Sheet'!$1:$2", "My Sheet", end{1, true}, end{2, true}},
		{"$5:$5", "", end{5, true}, end{5, true}},
	}
	for _, tt := range tests {
		node, err := Parse(tt.src)
		if err != nil {
			t.Errorf("Parse(%q): %v", tt.src, err)
			continue
		}
		rng, ok := node.(*RangeRef)
		if !ok {
			t.Errorf("Parse(%q) = %T, want *RangeRef", tt.src, node)
			continue
		}
		// A row range covers all columns; the flags of each end are kept.
		if rng.From.Col != 1 || rng.To.Col != maxCol {
			t.Errorf("Parse(%q) columns = %d..%d, want 1..%d", tt.src, rng.From.Col, rng.To.Col, maxCol)
		}
		if rng.From.Row != tt.from.row || rng.From.AbsRow != tt.from.absRow || rng.To.Row != tt.to.row || rng.To.AbsRow != tt.to.absRow {
			t.Errorf("Parse(%q) rows = %d(abs %v)..%d(abs %v), want %+v..%+v",
				tt.src, rng.From.Row, rng.From.AbsRow, rng.To.Row, rng.To.AbsRow, tt.from, tt.to)
		}
		if rng.From.AbsCol || rng.To.AbsCol {
			t.Errorf("Parse(%q) has an absolute column flag", tt.src)
		}
		if rng.From.Sheet != tt.sheet {
			t.Errorf("Parse(%q) sheet = %q, want %q", tt.src, rng.From.Sheet, tt.sheet)
		}
	}

	// In an expression.
	for _, src := range []string{"SUM($1:$3)", "SUM($1:$3,A1)+1", "COUNT(1:$3)*2", "SUM(Sheet1!$1:$2)", "ROWS($2:$5)"} {
		if _, err := Parse(src); err != nil {
			t.Errorf("Parse(%q): %v", src, err)
		}
	}

	// A lone absolute row, and a mix with a cell or a column, stay errors.
	for _, src := range []string{"$1", "$1+1", "SUM($1)", "1+$3", "A1:$3", "$3:A1", "A:$3", "$3:B", "$1 $3"} {
		if node, err := Parse(src); err == nil {
			t.Errorf("Parse(%q) = %v, want an error", src, node)
		}
	}

	// Forms that must not change.
	for _, src := range []string{"$A$1", "A$1+$A1", "SUM($A:$A)", "SUM(A:A)", "LOG10(100)", "TRUE", "SUM(1:3)", "SUM(Sheet1!2:3)", "A1:B2"} {
		if _, err := Parse(src); err != nil {
			t.Errorf("Parse(%q): %v", src, err)
		}
	}
}

// The reference scanner, the shift and the F4 cycle already write the
// absolute row form. Each result must be a formula that the parser reads.
func TestAbsoluteRowRangesFromTheRewriters(t *testing.T) {
	cycled, _, _, changed := CycleAbsolute("SUM(1:3)", 5, 5)
	if !changed || cycled != "SUM($1:$3)" {
		t.Fatalf("CycleAbsolute = %q, %v", cycled, changed)
	}
	if _, err := Parse(cycled); err != nil {
		t.Errorf("Parse(%q): %v", cycled, err)
	}
	for _, src := range []string{"SUM($1:3)", "SUM(1:$3)", "SUM(Data!$1:$2)"} {
		shifted := ShiftRefs(src, 0, 2)
		if _, err := Parse(shifted); err != nil {
			t.Errorf("Parse(%q) (shift of %q): %v", shifted, src, err)
		}
	}
}
