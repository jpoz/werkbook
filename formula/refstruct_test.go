package formula

import (
	"fmt"
	"strings"
	"testing"
)

func TestShiftSpan(t *testing.T) {
	type want struct {
		lo, hi int
		gone   bool
	}
	keep := func(lo, hi int) want { return want{lo, hi, false} }
	gone := want{0, 0, true}

	insert := StructuralEdit{At: 5, Count: 2}            // 2 new rows at 5 and 6
	del := StructuralEdit{At: 5, Count: 3, Delete: true} // rows 5, 6, 7 are removed
	insertCols := StructuralEdit{Cols: true, At: 5, Count: 2}
	tests := []struct {
		name   string
		e      StructuralEdit
		lo, hi int
		want   want
	}{
		// Insert, point.
		{"insert: point before", insert, 4, 4, keep(4, 4)},
		{"insert: point at the start", insert, 5, 5, keep(7, 7)},
		{"insert: point after", insert, 9, 9, keep(11, 11)},
		{"insert: first row", insert, 1, 1, keep(1, 1)},
		// Insert, span.
		{"insert: span before", insert, 1, 4, keep(1, 4)},
		{"insert: span that ends at the insert row grows", insert, 2, 5, keep(2, 7)},
		{"insert: span around the insert row grows", insert, 3, 8, keep(3, 10)},
		{"insert: span that starts at the insert row moves", insert, 5, 8, keep(7, 10)},
		{"insert: span after moves", insert, 6, 9, keep(8, 11)},
		// Insert at the end of the sheet.
		{"insert: point at the last row is pushed off", insert, refMaxRow, refMaxRow, gone},
		{"insert: point one row before the last is pushed off", insert, refMaxRow - 1, refMaxRow - 1, gone},
		{"insert: point that moves to the last row", insert, refMaxRow - 2, refMaxRow - 2, keep(refMaxRow, refMaxRow)},
		{"insert: span that grows past the end is cut", insert, 3, refMaxRow - 1, keep(3, refMaxRow)},
		{"insert: span that moves partly off is cut", insert, 10, refMaxRow, keep(12, refMaxRow)},
		{"insert: full axis is unchanged", insert, 1, refMaxRow, keep(1, refMaxRow)},

		// Delete, point.
		{"delete: point before", del, 4, 4, keep(4, 4)},
		{"delete: point at the band start", del, 5, 5, gone},
		{"delete: point in the band", del, 6, 6, gone},
		{"delete: point at the band end", del, 7, 7, gone},
		{"delete: point after", del, 8, 8, keep(5, 5)},
		{"delete: point far after", del, 100, 100, keep(97, 97)},
		// Delete, span.
		{"delete: span before", del, 1, 4, keep(1, 4)},
		{"delete: span that is the band", del, 5, 7, gone},
		{"delete: span in the band", del, 6, 7, gone},
		{"delete: span whose end is in the band shrinks", del, 2, 6, keep(2, 4)},
		{"delete: span that ends at the band end shrinks", del, 2, 7, keep(2, 4)},
		{"delete: span whose start is in the band shrinks", del, 6, 10, keep(5, 7)},
		{"delete: span that starts at the band start shrinks", del, 5, 10, keep(5, 7)},
		{"delete: span around the band shrinks", del, 2, 10, keep(2, 7)},
		{"delete: span after moves", del, 8, 12, keep(5, 9)},
		{"delete: span that touches the band from before", del, 3, 4, keep(3, 4)},
		{"delete: full axis is unchanged", del, 1, refMaxRow, keep(1, refMaxRow)},
		{"delete: span to the last row shrinks", del, 2, refMaxRow, keep(2, refMaxRow-3)},

		// Columns have their own end.
		{"columns: point at the last column is pushed off", insertCols, refMaxCol, refMaxCol, gone},
		{"columns: full axis is unchanged", insertCols, 1, refMaxCol, keep(1, refMaxCol)},
		{"columns: a span of 1..16384 rows is not the full row axis", insert, 1, refMaxCol, keep(1, refMaxCol+2)},

		// No edit.
		{"count 0", StructuralEdit{At: 5, Count: 0}, 5, 9, keep(5, 9)},
		{"negative count", StructuralEdit{At: 5, Count: -2, Delete: true}, 5, 9, keep(5, 9)},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			lo, hi, isGone := tt.e.ShiftSpan(tt.lo, tt.hi)
			if (want{lo, hi, isGone}) != tt.want {
				t.Errorf("ShiftSpan(%d, %d) = (%d, %d, %v), want %+v", tt.lo, tt.hi, lo, hi, isGone, tt.want)
			}
		})
	}
}

// An insert and the delete of the same lines give the first span again, for
// each span that the insert did not cut or push off.
func TestShiftSpanInsertThenDeleteIsIdentity(t *testing.T) {
	for at := 1; at <= 8; at++ {
		for count := 1; count <= 3; count++ {
			ins := StructuralEdit{At: at, Count: count}
			del := StructuralEdit{At: at, Count: count, Delete: true}
			for lo := 1; lo <= 10; lo++ {
				for hi := lo; hi <= 10; hi++ {
					lo2, hi2, gone := ins.ShiftSpan(lo, hi)
					if gone {
						t.Fatalf("insert %+v made %d..%d gone", ins, lo, hi)
					}
					lo3, hi3, gone := del.ShiftSpan(lo2, hi2)
					if gone || lo3 != lo || hi3 != hi {
						t.Errorf("insert then delete at %d x%d: %d..%d -> %d..%d -> %d..%d (gone %v)", at, count, lo, hi, lo2, hi2, lo3, hi3, gone)
					}
				}
			}
		}
	}
}

func TestApplyStructuralEditMatrix(t *testing.T) {
	names := []string{"Sheet1", "Sheet2", "Sheet3", "My Data"}
	insRows := StructuralEdit{Sheet: "Sheet1", At: 5, Count: 2, SheetNames: names}
	delRows := StructuralEdit{Sheet: "Sheet1", At: 5, Count: 3, Delete: true, SheetNames: names}
	insCols := StructuralEdit{Sheet: "Sheet1", Cols: true, At: 3, Count: 2, SheetNames: names}               // at C
	delCols := StructuralEdit{Sheet: "Sheet1", Cols: true, At: 3, Count: 2, Delete: true, SheetNames: names} // C:D

	tests := []struct {
		name string
		e    StructuralEdit
		src  string
		want string
	}{
		// Insert rows: point.
		{"insert rows: point before", insRows, "A4", "A4"},
		{"insert rows: point at the start", insRows, "A5", "A7"},
		{"insert rows: point after", insRows, "A9", "A11"},
		// Insert rows: range.
		{"insert rows: range before", insRows, "SUM(A1:A4)", "SUM(A1:A4)"},
		{"insert rows: range that ends at the insert row grows", insRows, "SUM(A1:A5)", "SUM(A1:A7)"},
		{"insert rows: range around the insert row grows", insRows, "SUM(A3:B8)", "SUM(A3:B10)"},
		{"insert rows: range that starts at the insert row moves", insRows, "SUM(A5:A8)", "SUM(A7:A10)"},
		{"insert rows: range after moves", insRows, "SUM(A6:A9)", "SUM(A8:A11)"},
		// Insert rows: full column and full row.
		{"insert rows: full column is unchanged", insRows, "SUM(A:C)", "SUM(A:C)"},
		{"insert rows: row range before", insRows, "SUM(1:4)", "SUM(1:4)"},
		{"insert rows: row range around grows", insRows, "SUM(3:8)", "SUM(3:10)"},
		{"insert rows: row range after moves", insRows, "SUM(5:5)", "SUM(7:7)"},
		{"insert rows: all rows are unchanged", insRows, "SUM(1:1048576)", "SUM(1:1048576)"},
		{"insert rows: a cell range over all rows is unchanged", insRows, "SUM(A1:A1048576)", "SUM(A1:A1048576)"},

		// Delete rows: point.
		{"delete rows: point before", delRows, "A4", "A4"},
		{"delete rows: point at the band start", delRows, "A5+1", "#REF!+1"},
		{"delete rows: point in the band", delRows, "A6", "#REF!"},
		{"delete rows: point at the band end", delRows, "A7", "#REF!"},
		{"delete rows: point after", delRows, "A8", "A5"},
		// Delete rows: range.
		{"delete rows: range before", delRows, "SUM(A1:A4)", "SUM(A1:A4)"},
		{"delete rows: full range in the band", delRows, "SUM(A5:B7)", "SUM(#REF!)"},
		{"delete rows: range in the band", delRows, "SUM(A6:A6)", "SUM(#REF!)"},
		{"delete rows: partial at the end", delRows, "SUM(A2:A6)", "SUM(A2:A4)"},
		{"delete rows: partial at the start", delRows, "SUM(A6:A10)", "SUM(A5:A7)"},
		{"delete rows: band in the range", delRows, "SUM(A1:A10)", "SUM(A1:A7)"},
		{"delete rows: range after", delRows, "SUM(A8:A12)", "SUM(A5:A9)"},
		{"delete rows: full column is unchanged", delRows, "SUM(A:A)", "SUM(A:A)"},
		{"delete rows: row range in the band", delRows, "SUM(5:7)", "SUM(#REF!)"},
		{"delete rows: row range partly in the band", delRows, "SUM(4:6)", "SUM(4:4)"},
		{"delete rows: row range after", delRows, "SUM(9:10)", "SUM(6:7)"},

		// Insert columns.
		{"insert cols: point before", insCols, "B1", "B1"},
		{"insert cols: point at the start", insCols, "C1", "E1"},
		{"insert cols: point after", insCols, "Z1", "AB1"},
		{"insert cols: range around grows", insCols, "SUM(A1:D9)", "SUM(A1:F9)"},
		{"insert cols: range that starts at the insert column moves", insCols, "SUM(C1:D9)", "SUM(E1:F9)"},
		{"insert cols: column range around grows", insCols, "SUM(B:C)", "SUM(B:E)"},
		{"insert cols: column range after moves", insCols, "SUM(D:D)", "SUM(F:F)"},
		{"insert cols: full row is unchanged", insCols, "SUM(1:3)", "SUM(1:3)"},
		{"insert cols: all columns are unchanged", insCols, "SUM(A:XFD)", "SUM(A:XFD)"},

		// Delete columns.
		{"delete cols: point before", delCols, "B1", "B1"},
		{"delete cols: point in the band", delCols, "C1+D1", "#REF!+#REF!"},
		{"delete cols: point after", delCols, "E1", "C1"},
		{"delete cols: full range in the band", delCols, "SUM(C1:D5)", "SUM(#REF!)"},
		{"delete cols: partial at the end", delCols, "SUM(A1:C5)", "SUM(A1:B5)"},
		{"delete cols: partial at the start", delCols, "SUM(D1:F5)", "SUM(C1:D5)"},
		{"delete cols: column range in the band", delCols, "SUM(C:D)", "SUM(#REF!)"},
		{"delete cols: column range around", delCols, "SUM(A:F)", "SUM(A:D)"},
		{"delete cols: full row is unchanged", delCols, "SUM(2:2)", "SUM(2:2)"},

		// Absolute forms are adjusted as relative forms are.
		{"absolute cell moves", insRows, "$A$5+A$5+$A5", "$A$7+A$7+$A7"},
		{"absolute range grows", insRows, "SUM($A$1:$A$9)", "SUM($A$1:$A$11)"},
		{"absolute cell deleted", delRows, "$A$6", "#REF!"},
		{"absolute row range", insRows, "SUM($5:$6)", "SUM($7:$8)"},
		{"absolute column range", delCols, "SUM($E:$F)", "SUM($C:$D)"},
		{"mixed flags stay on their corner", insCols, "SUM($B1:C$2)", "SUM($B1:E$2)"},

		// A reversed range keeps its corner order.
		{"reversed range grows", insRows, "SUM(A8:A3)", "SUM(A10:A3)"},
		{"reversed range shrinks", delRows, "SUM(A10:A6)", "SUM(A7:A5)"},
		{"reversed range with flags", insRows, "SUM($A$8:B3)", "SUM($A$10:B3)"},

		// The edge of the sheet.
		{"insert rows pushes the last row off", insRows, "A1048576", "#REF!"},
		{"insert rows: one before the last is pushed off", insRows, "A1048575+A1048574", "#REF!+A1048576"},
		{"insert rows: range to the last row is cut at the end", insRows, "SUM(A9:A1048576)", "SUM(A11:A1048576)"},
		{"insert cols pushes the last column off", insCols, "XFD1+XFC1+XFB1", "#REF!+#REF!+XFD1"},

		// Many references; only some change.
		{"two references, one becomes #REF!", delRows, "A6+A9", "#REF!+A6"},
		{"other text stays", insRows, " if( a5>0 , \"A5\" , b$9 ) ", " if( A7>0 , \"A5\" , B$11 ) "},
		{"no reference changes: text is the same", insRows, " sum( a1:a2 )+b3", " sum( a1:a2 )+b3"},
		{"spill operator moves with its reference", insRows, "SUM(A5#)", "SUM(A7#)"},
		{"spill operator goes with a deleted reference", delRows, "SUM(A6#)", "SUM(#REF!)"},
		{"range with spaces", delRows, "SUM(A6 : A10)", "SUM(A5:A7)"},

		// Opaque spans stay; references around them are adjusted.
		{"structured reference", insRows, "Table1[[#This Row],[A5]]+A5", "Table1[[#This Row],[A5]]+A7"},
		{"external workbook", insRows, "[Book.xlsx]Sheet1!A5+A5", "[Book.xlsx]Sheet1!A5+A7"},
		{"quoted external workbook", insRows, "'[Book.xlsx]Sheet1'!A5+A5", "'[Book.xlsx]Sheet1'!A5+A7"},
		{"string", insRows, "\"A5\"&A5", "\"A5\"&A7"},
		{"unterminated string", insRows, "A5&\"A5", "A7&\"A5"},
		{"error literal", delRows, "#REF!+A8+#N/A", "#REF!+A5+#N/A"},
		{"function with digits", insRows, "LOG10(A5)", "LOG10(A7)"},
		{"name", insRows, "Total5+A5", "Total5+A7"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, changed := ApplyStructuralEdit(tt.src, "Sheet1", tt.e)
			if got != tt.want {
				t.Errorf("ApplyStructuralEdit(%q) = %q, want %q", tt.src, got, tt.want)
			}
			if changed != (got != tt.src) {
				t.Errorf("changed = %v for %q -> %q", changed, tt.src, got)
			}
		})
	}
}

func TestApplyStructuralEditSheets(t *testing.T) {
	names := []string{"Sheet1", "Sheet2", "Sheet3", "My Data", "A1"}
	e := StructuralEdit{Sheet: "Sheet2", At: 5, Count: 2, SheetNames: names}
	del := StructuralEdit{Sheet: "My Data", At: 5, Count: 1, Delete: true, SheetNames: names}

	tests := []struct {
		name         string
		e            StructuralEdit
		formulaSheet string
		src          string
		want         string
	}{
		// References with no sheet name follow only on the edited sheet.
		{"unqualified, formula on the edited sheet", e, "Sheet2", "A5+B9", "A7+B11"},
		{"unqualified, formula on a different sheet", e, "Sheet1", "A5+B9", "A5+B9"},
		{"formula sheet with a different letter case", e, "sheet2", "A5", "A7"},
		// Qualified references follow from all sheets.
		{"qualified, from a different sheet", e, "Sheet1", "Sheet2!A5+A5", "Sheet2!A7+A5"},
		{"qualified, from the edited sheet", e, "Sheet2", "Sheet2!A5+A5", "Sheet2!A7+A7"},
		{"qualified with a different letter case", e, "Sheet1", "SHEET2!A5+sheet2!$A$5", "SHEET2!A7+sheet2!$A$7"},
		{"qualified to a different sheet does not change", e, "Sheet2", "Sheet1!A5+Sheet3!A5", "Sheet1!A5+Sheet3!A5"},
		{"quotes that are not necessary stay", e, "Sheet1", "'Sheet2'!A5", "'Sheet2'!A7"},
		{"qualified range and columns", e, "Sheet1", "SUM(Sheet2!A1:B9)+SUM(Sheet2!A:A)+SUM(Sheet2!5:5)", "SUM(Sheet2!A1:B11)+SUM(Sheet2!A:A)+SUM(Sheet2!7:7)"},
		// A deleted reference has no sheet name.
		{"deleted qualified reference", del, "Sheet1", "'My Data'!A5+'My Data'!A6", "#REF!+'My Data'!A5"},
		{"deleted qualified range", del, "Sheet1", "SUM('My Data'!A5:C5)*2", "SUM(#REF!)*2"},
		{"quoted sheet name", del, "Sheet1", "'my data'!B9", "'my data'!B8"},
		// A sheet name that looks like a cell.
		{"sheet named A1", StructuralEdit{Sheet: "A1", At: 1, Count: 1, SheetNames: names}, "Sheet1", "'A1'!B2+A1", "'A1'!B3+A1"},
		// 3D references follow when the edited sheet is in the span.
		{"3D with the edited sheet in the span", e, "Sheet1", "SUM(Sheet1:Sheet3!A5)", "SUM(Sheet1:Sheet3!A7)"},
		{"3D with the edited sheet at the start of the span", e, "Sheet1", "SUM(Sheet2:Sheet3!A5)", "SUM(Sheet2:Sheet3!A7)"},
		{"3D with the edited sheet at the end of the span", e, "Sheet1", "SUM(Sheet1:Sheet2!A5)", "SUM(Sheet1:Sheet2!A7)"},
		{"3D in reversed sheet order", e, "Sheet1", "SUM(Sheet3:Sheet1!A5)", "SUM(Sheet3:Sheet1!A7)"},
		{"3D with the edited sheet out of the span", e, "Sheet2", "SUM(Sheet3:'My Data'!A5)", "SUM(Sheet3:'My Data'!A5)"},
		{"quoted 3D", del, "Sheet1", "SUM('Sheet3:My Data'!A9)", "SUM('Sheet3:My Data'!A8)"},
		{"3D with a sheet that is not in the workbook", e, "Sheet1", "SUM(Sheet1:Gone!A5)", "SUM(Sheet1:Gone!A5)"},
		{"3D deleted", StructuralEdit{Sheet: "Sheet2", At: 5, Count: 1, Delete: true, SheetNames: names}, "Sheet1", "SUM(Sheet1:Sheet3!A5)", "SUM(#REF!)"},
		// A defined name value has no formula sheet.
		{"defined name value", e, "", "Sheet2!$A$5:$B$9", "Sheet2!$A$7:$B$11"},
		{"defined name value with no sheet name does not change", e, "", "$A$5", "$A$5"},
		{"print titles", e, "", "Sheet2!$5:$6", "Sheet2!$7:$8"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, _ := ApplyStructuralEdit(tt.src, tt.formulaSheet, tt.e)
			if got != tt.want {
				t.Errorf("ApplyStructuralEdit(%q on %q) = %q, want %q", tt.src, tt.formulaSheet, got, tt.want)
			}
		})
	}
}

// The reference rewrite and ShiftSpan are one rule: for each point and span,
// the text result is the span result.
func TestApplyStructuralEditAgreesWithShiftSpan(t *testing.T) {
	for _, del := range []bool{false, true} {
		for _, cols := range []bool{false, true} {
			for at := 1; at <= 7; at++ {
				for count := 1; count <= 3; count++ {
					e := StructuralEdit{Sheet: "S", Cols: cols, At: at, Count: count, Delete: del}
					for lo := 1; lo <= 9; lo++ {
						for hi := lo; hi <= 9; hi++ {
							src := fmt.Sprintf("B%d:D%d", lo, hi)
							if cols {
								src = fmt.Sprintf("%s2:%s4", ColNumberToLetters(lo), ColNumberToLetters(hi))
							}
							lo2, hi2, gone := e.ShiftSpan(lo, hi)
							want := "#REF!"
							if !gone {
								want = fmt.Sprintf("B%d:D%d", lo2, hi2)
								if cols {
									want = fmt.Sprintf("%s2:%s4", ColNumberToLetters(lo2), ColNumberToLetters(hi2))
								}
							}
							if got, _ := ApplyStructuralEdit(src, "S", e); got != want {
								t.Errorf("edit %+v: %s -> %s, want %s", e, src, got, want)
							}
						}
					}
				}
			}
		}
	}
}

// An insert and then the delete of the same lines give the first formula
// (in upper case), and each result is a formula that the parser reads.
func TestApplyStructuralEditInverseAndParse(t *testing.T) {
	names := []string{"Sheet1", "Sheet2", "Data", "My Sheet", "Price List", "Sales", "It's", "Jan", "Dec"}
	for _, src := range realFormulas {
		_, parseErr := Parse(src)
		for _, cols := range []bool{false, true} {
			ins := StructuralEdit{Sheet: "Sheet1", Cols: cols, At: 2, Count: 3, SheetNames: names}
			del := ins
			del.Delete = true
			inserted, _ := ApplyStructuralEdit(src, "Sheet1", ins)
			back, _ := ApplyStructuralEdit(inserted, "Sheet1", del)
			// A reference that moved is written in upper case; one that did
			// not move keeps its case. The compare ignores letter case.
			if !strings.EqualFold(back, src) {
				t.Errorf("cols=%v: %q -> %q -> %q, want the first text", cols, src, inserted, back)
			}
			if parseErr == nil {
				if _, err := Parse(inserted); err != nil {
					t.Errorf("Parse(%q) (insert in %q): %v", inserted, src, err)
				}
			}
		}
	}
	// Edits that change nothing.
	for _, e := range []StructuralEdit{{Sheet: "Sheet1", At: 1, Count: 0}, {Sheet: "Sheet1", At: 1, Count: -1, Delete: true}} {
		if got, changed := ApplyStructuralEdit("A1+B2", "Sheet1", e); got != "A1+B2" || changed {
			t.Errorf("edit %+v changed the text to %q", e, got)
		}
	}
	if got, changed := ApplyStructuralEdit("", "Sheet1", StructuralEdit{Sheet: "Sheet1", At: 1, Count: 1}); got != "" || changed {
		t.Errorf("empty text became %q", got)
	}
}
