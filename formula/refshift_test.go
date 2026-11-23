package formula

import (
	"strings"
	"testing"
)

func TestShiftRefs(t *testing.T) {
	tests := []struct {
		name       string
		src        string
		dCol, dRow int
		want       string
	}{
		// Relative and absolute parts.
		{"relative cell down", "A1", 0, 1, "A2"},
		{"relative cell right", "A1", 1, 0, "B1"},
		{"relative cell both", "B2", 2, 3, "D5"},
		{"relative cell up and left", "C5", -2, -4, "A1"},
		{"absolute cell stays", "$A$1", 3, 3, "$A$1"},
		{"absolute column", "$A1", 3, 3, "$A4"},
		{"absolute row", "A$1", 3, 3, "D$1"},
		{"all four forms", "A1+$A1+A$1+$A$1", 1, 1, "B2+$A2+B$1+$A$1"},
		{"column past Z", "Z1", 1, 0, "AA1"},
		{"lower case is written in upper case when it moves", "a1+$b$2", 1, 1, "B2+$b$2"},

		// Ranges.
		{"range", "SUM(A1:B2)", 1, 1, "SUM(B2:C3)"},
		{"range with an absolute start", "SUM($A$1:A1)", 0, 4, "SUM($A$1:A5)"},
		{"range with mixed parts", "$A1:B$2", 2, 2, "$A3:D$2"},
		{"fully absolute range is unchanged text", "sum( $a$1:$b$2 )", 5, 5, "sum( $a$1:$b$2 )"},

		// Column-only and row-only references shift on their own axis.
		{"columns right", "SUM(A:C)", 2, 0, "SUM(C:E)"},
		{"columns do not move with rows", "SUM(A:C)", 0, 5, "SUM(A:C)"},
		{"absolute columns", "SUM($A:$A)", 2, 0, "SUM($A:$A)"},
		{"half absolute columns", "SUM($A:B)", 2, 0, "SUM($A:D)"},
		{"rows down", "SUM(1:3)", 0, 2, "SUM(3:5)"},
		{"rows do not move with columns", "SUM(1:3)", 4, 0, "SUM(1:3)"},
		{"absolute rows", "SUM($1:$3)", 0, 2, "SUM($1:$3)"},
		{"half absolute rows", "SUM($1:3)", 0, 2, "SUM($1:5)"},

		// Sheet edges.
		{"cell above row 1", "A1", 0, -1, "#REF!"},
		{"cell left of column A", "A1", -1, 0, "#REF!"},
		{"cell below the last row", "A1048576", 0, 1, "#REF!"},
		{"cell right of the last column", "XFD1", 1, 0, "#REF!"},
		{"last cell stays valid", "XFC1048575", 1, 1, "XFD1048576"},
		{"range with one end off the sheet", "SUM(A1:B2)", 0, -1, "SUM(#REF!)"},
		{"range end off the sheet at the right", "SUM(A1:XFD1)", 1, 0, "SUM(#REF!)"},
		{"full column right of the edge", "SUM(XFD:XFD)", 1, 0, "SUM(#REF!)"},
		{"full column left of the edge", "SUM(A:B)", -1, 0, "SUM(#REF!)"},
		{"row range above row 1", "SUM(1:3)", 0, -1, "SUM(#REF!)"},
		{"row range below the last row", "SUM(5:1048576)", 0, 1, "SUM(#REF!)"},
		{"absolute part never leaves", "$A$1+A$1+$A1", -5, -5, "$A$1+#REF!+#REF!"},
		{"one of two references becomes #REF!", "A1+B2", -1, 0, "#REF!+A2"},
		{"the #REF! has no sheet qualifier", "Sheet2!A1+1", 0, -1, "#REF!+1"},
		{"spill operator goes with an invalid reference", "SUM(A1#)", 0, -1, "SUM(#REF!)"},

		// Sheet qualifiers.
		{"sheet qualified", "Sheet2!A1", 1, 1, "Sheet2!B2"},
		{"quoted sheet", "'My Sheet'!A1:B2", 1, 1, "'My Sheet'!B2:C3"},
		{"quoted sheet with a quote", "'It''s'!$A1", 1, 1, "'It''s'!$A2"},
		{"3D", "SUM(Jan:Dec!B5)", 1, 0, "SUM(Jan:Dec!C5)"},
		{"sheet rows", "SUM(Data!1:2)", 0, 1, "SUM(Data!2:3)"},

		// Cases from the review of milestone 1.
		{"reversed range left of column A", "SUM(B2:A1)", -1, 0, "SUM(#REF!)"},
		{"reversed range keeps its corner order", "SUM(B2:A1)", 1, 1, "SUM(C3:B2)"},
		{"quoted 3D range", "SUM('Sheet 1:Sheet 3'!A1:B2)", 1, 1, "SUM('Sheet 1:Sheet 3'!B2:C3)"},
		{"two qualified ends are two references", "'Sheet1'!A1:'Sheet1'!B2", 1, 1, "'Sheet1'!B2:'Sheet1'!C3"},
		{"quoted sheet name that looks like a cell", "'A1'!B2", 1, 1, "'A1'!C3"},
		{"sheet name that looks like a cell", "AB12!C3", 1, 1, "AB12!D4"},
		{"spill reference right of the last column", "XFD1#", 1, 0, "#REF!"},
		{"R1C1-like names stay", "R1C1+RC+R[1]C[1]", 1, 1, "R1C1+RC+R[1]C[1]"},
		{"range with spaces: one end off the sheet makes the full range #REF!", "SUM(A1 : B2)", -1, 0, "SUM(#REF!)"},
		{"range with spaces is written without them when it moves", "SUM(A1 : B2)", 1, 1, "SUM(B2:C3)"},
		{"range with spaces and absolute parts only is unchanged", "SUM($A$1 : $B$2)", 1, 1, "SUM($A$1 : $B$2)"},
		{"up: relative parts become #REF!, absolute parts stay", "A1+B1+A$1+$A$1", 0, -1, "#REF!+#REF!+A$1+$A$1"},
		{"left: relative parts become #REF!, absolute parts stay", "A1+A2+$A1+$A$1", -1, 0, "#REF!+#REF!+$A1+$A$1"},

		// Text that must not change.
		{"many references and other text", " IF( A1>0 , \"A1\" , B$2*Sheet2!$C3 ) ", 1, 2, " IF( B3>0 , \"A1\" , C$2*Sheet2!$C5 ) "},
		{"functions and names", "LOG10(A1)+Total+TRUE", 1, 1, "LOG10(B2)+Total+TRUE"},
		{"structured reference", "Table1[[#This Row],[A1]]+A1", 1, 1, "Table1[[#This Row],[A1]]+B2"},
		{"external workbook", "[Book.xlsx]Sheet1!A1+A1", 1, 1, "[Book.xlsx]Sheet1!A1+B2"},
		{"spill operator", "SUM(C12#)", 1, 1, "SUM(D13#)"},
		{"string that is not terminated", "A1&\"B2", 1, 1, "B2&\"B2"},
		{"error literals", "#REF!+A1+#N/A", 1, 1, "#REF!+B2+#N/A"},
		{"numbers", "A1*1E5+2.5", 1, 1, "B2*1E5+2.5"},
		{"empty", "", 1, 1, ""},
		{"no references", "1+2", 1, 1, "1+2"},
		{"zero distance", "a1+B2", 0, 0, "a1+B2"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := ShiftRefs(tt.src, tt.dCol, tt.dRow); got != tt.want {
				t.Errorf("ShiftRefs(%q, %d, %d) = %q, want %q", tt.src, tt.dCol, tt.dRow, got, tt.want)
			}
		})
	}
}

// A shift and its inverse give the first text when no reference left the sheet.
func TestShiftRefsInverse(t *testing.T) {
	for _, src := range realFormulas {
		shifted := ShiftRefs(src, 3, 7)
		back := ShiftRefs(shifted, -3, -7)
		want := ShiftRefs(src, 1, 1)
		want = ShiftRefs(want, -1, -1) // upper-case form of src
		if back != want {
			t.Errorf("shift and inverse of %q = %q, want %q", src, back, want)
		}
	}
}

// The engine must read each shifted formula in the same way as a formula
// that was typed with the shifted references.
func TestShiftRefsResultParses(t *testing.T) {
	parsed := 0
	for _, src := range realFormulas {
		if _, err := Parse(src); err != nil {
			continue // the list also has text that only the scanner accepts
		}
		parsed++
		for _, d := range [][2]int{{2, 2}, {-1, 0}, {0, -1}, {16384, 0}, {0, 1048576}} {
			shifted := ShiftRefs(src, d[0], d[1])
			// Known limit of the parser, not of the shift: #REF! as an
			// operand of an intersection ("#REF! A2:B6") or as a member of
			// a union ("(#REF!,B1:B3)") is a parse error. Excel accepts
			// both. The cell shows an error value in both programs.
			if strings.Contains(shifted, "#REF! ") || strings.Contains(shifted, "(#REF!,") {
				continue
			}
			if _, err := Parse(shifted); err != nil {
				t.Errorf("Parse(%q) (shift of %q by %v): %v", shifted, src, d, err)
			}
		}
	}
	if parsed < 30 {
		t.Errorf("only %d formulas of the list parse; the test needs 30 or more", parsed)
	}
}
