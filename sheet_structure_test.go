package werkbook

import (
	"bytes"
	"fmt"
	"sort"
	"strings"
	"testing"
)

// dumpWorkbook gives a text form of the full state that a structural edit
// can change: cells (formula, value, flags, style), spill ranges, merged
// ranges, sizes, defined names and tables, for all sheets.
func dumpWorkbook(t *testing.T, f *File) string {
	t.Helper()
	return dumpWorkbookState(t, f, false)
}

// dumpWorkbookState with content true leaves out what a save and open
// changes but that is not content: the placeholder cells that a file has for
// spill followers (and with them the size of the stored grid).
func dumpWorkbookState(t *testing.T, f *File, content bool) string {
	t.Helper()
	f.Recalculate()
	var b strings.Builder
	for _, name := range f.SheetNames() {
		s := f.Sheet(name)
		if content {
			fmt.Fprintf(&b, "== %s\n", name)
		} else {
			fmt.Fprintf(&b, "== %s (max %d x %d)\n", name, s.MaxRow(), s.MaxCol())
		}
		rowNums := make([]int, 0, len(s.rows))
		for n := range s.rows {
			rowNums = append(rowNums, n)
		}
		sort.Ints(rowNums)
		for _, n := range rowNums {
			r := s.rows[n]
			if r.num != n {
				t.Errorf("%s: row at key %d has num %d", name, n, r.num)
			}
			if r.height != 0 || r.hidden {
				fmt.Fprintf(&b, "row %d height=%v hidden=%v\n", n, r.height, r.hidden)
			}
			cols := make([]int, 0, len(r.cells))
			for c := range r.cells {
				cols = append(cols, c)
			}
			sort.Ints(cols)
			for _, c := range cols {
				cell := r.cells[c]
				if cell.col != c {
					t.Errorf("%s: cell at row %d key %d has col %d", name, n, c, cell.col)
				}
				ref, _ := CoordinatesToCellName(c, n)
				v, _ := s.GetValue(ref)
				if content && cell.formula == "" && cell.value.Type == TypeEmpty && cell.style == nil {
					continue
				}
				fmt.Fprintf(&b, "%s f=%q v=%s", ref, cell.formula, valueText(v))
				if cell.isArrayFormula {
					b.WriteString(" array")
				}
				if cell.dynamicArraySpill {
					toCol, toRow, ok := s.SpillBounds(c, n)
					fmt.Fprintf(&b, " spill=%d,%d,%v", toCol, toRow, ok)
				}
				if cell.style != nil {
					fmt.Fprintf(&b, " style=%p", cell.style)
				}
				b.WriteString("\n")
			}
		}
		for _, mr := range s.merges {
			fmt.Fprintf(&b, "merge %s:%s\n", mr.Start, mr.End)
		}
		if rows, cols := s.FreezePanes(); rows != 0 || cols != 0 {
			fmt.Fprintf(&b, "freeze rows=%d cols=%d\n", rows, cols)
		}
		widths := s.ColumnWidths()
		cols := make([]int, 0, len(widths))
		for c := range widths {
			cols = append(cols, c)
		}
		sort.Ints(cols)
		for _, c := range cols {
			fmt.Fprintf(&b, "width %s=%v\n", ColumnNumberToName(c), widths[c])
		}
	}
	for _, dn := range f.DefinedNames() {
		fmt.Fprintf(&b, "name %s[%d]=%s\n", dn.Name, dn.LocalSheetID, dn.Value)
	}
	for _, td := range f.Tables() {
		fmt.Fprintf(&b, "table %s on %s ref=%s cols=%v header=%d totals=%d\n", td.DisplayName, td.SheetName, td.Ref, td.Columns, td.HeaderRowCount, td.TotalsRowCount)
	}
	for _, ti := range f.tables {
		fmt.Fprintf(&b, "tableinfo %s %d,%d-%d,%d cols=%v\n", ti.Name, ti.FirstCol, ti.FirstRow, ti.LastCol, ti.LastRow, ti.Columns)
	}
	return b.String()
}

func valueText(v Value) string {
	switch v.Type {
	case TypeNumber:
		return fmt.Sprintf("n:%v", v.Number)
	case TypeString:
		return fmt.Sprintf("s:%q", v.String)
	case TypeBool:
		return fmt.Sprintf("b:%v", v.Bool)
	case TypeError:
		return "e:" + v.String
	default:
		return "-"
	}
}

// stripStylePointers removes the style pointer text, which is different in
// a workbook that was saved and opened.
func stripStylePointers(dump string) string {
	lines := strings.Split(dump, "\n")
	for i, line := range lines {
		if j := strings.Index(line, " style=0x"); j >= 0 {
			lines[i] = line[:j] + " style"
		}
	}
	return strings.Join(lines, "\n")
}

type structOp struct {
	sheet string
	kind  string // "insRows", "delRows", "insCols", "delCols"
	at, n int
}

func (op structOp) run(f *File) (*StructuralUndo, error) {
	s := f.Sheet(op.sheet)
	switch op.kind {
	case "insRows":
		return s.InsertRows(op.at, op.n)
	case "delRows":
		return s.DeleteRows(op.at, op.n)
	case "insCols":
		return s.InsertCols(op.at, op.n)
	default:
		return s.DeleteCols(op.at, op.n)
	}
}

func mustSetCells(t *testing.T, s *Sheet, pairs ...any) {
	t.Helper()
	for i := 0; i < len(pairs); i += 2 {
		ref := pairs[i].(string)
		var err error
		if text, ok := pairs[i+1].(string); ok && strings.HasPrefix(text, "=") {
			err = s.SetFormula(ref, text)
		} else {
			err = s.SetValue(ref, pairs[i+1])
		}
		if err != nil {
			t.Fatalf("set %s: %v", ref, err)
		}
	}
}

// checkStructuralEdit runs one edit and checks the parts of the contract
// that are the same for each case:
//   - the expected formulas and values
//   - a save and open gives the same workbook
//   - Undo gives the start state exactly
//   - the same edit after the undo gives the same result as the first run
func checkStructuralEdit(t *testing.T, f *File, op structOp, wantFormulas map[string]string, wantValues map[string]any) *StructuralUndo {
	t.Helper()
	before := dumpWorkbook(t, f)
	u, err := op.run(f)
	if err != nil {
		t.Fatalf("%+v: %v", op, err)
	}
	after := dumpWorkbook(t, f)
	checkCells(t, f, wantFormulas, wantValues)

	// Save and open.
	var buf bytes.Buffer
	if _, err := f.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo after %+v: %v", op, err)
	}
	opened, err := OpenReader(&buf)
	if err != nil {
		t.Fatalf("open after %+v: %v", op, err)
	}
	if got, want := stripStylePointers(dumpWorkbookState(t, opened, true)), stripStylePointers(dumpWorkbookState(t, f, true)); got != want {
		t.Errorf("%+v: state after save and open:\n%s\nwant:\n%s", op, got, want)
	}

	// Undo.
	if u.Cost() < 0 {
		t.Errorf("Cost = %d", u.Cost())
	}
	if err := u.Undo(); err != nil {
		t.Fatalf("Undo of %+v: %v", op, err)
	}
	if got := dumpWorkbook(t, f); got != before {
		t.Errorf("%+v: state after Undo:\n%s\nwant:\n%s", op, got, before)
	}
	if err := u.Undo(); err == nil {
		t.Errorf("a second Undo gave no error")
	}

	// Redo: the same edit again.
	u2, err := op.run(f)
	if err != nil {
		t.Fatalf("%+v after undo: %v", op, err)
	}
	if got := dumpWorkbook(t, f); got != after {
		t.Errorf("%+v: state after the edit was run again:\n%s\nwant:\n%s", op, got, after)
	}
	return u2
}

// checkCells checks formulas ("Sheet!A1" -> text, "" = no formula) and
// values ("Sheet!A1" -> number, string, bool, nil for empty, or an error
// text that starts with '#').
func checkCells(t *testing.T, f *File, wantFormulas map[string]string, wantValues map[string]any) {
	t.Helper()
	split := func(key string) (*Sheet, string) {
		i := strings.LastIndex(key, "!")
		s := f.Sheet(key[:i])
		if s == nil {
			t.Fatalf("no sheet for %q", key)
		}
		return s, key[i+1:]
	}
	for key, want := range wantFormulas {
		s, ref := split(key)
		if got, _ := s.GetFormula(ref); got != want {
			t.Errorf("%s formula = %q, want %q", key, got, want)
		}
	}
	for key, want := range wantValues {
		s, ref := split(key)
		v, err := s.GetValue(ref)
		if err != nil {
			t.Fatalf("GetValue(%s): %v", key, err)
		}
		var ok bool
		switch w := want.(type) {
		case nil:
			ok = v.IsEmpty()
		case int:
			ok = v.Type == TypeNumber && v.Number == float64(w)
		case float64:
			ok = v.Type == TypeNumber && v.Number == w
		case bool:
			ok = v.Type == TypeBool && v.Bool == w
		case string:
			if strings.HasPrefix(w, "#") {
				ok = v.Type == TypeError && v.String == w
			} else {
				ok = v.Type == TypeString && v.String == w
			}
		}
		if !ok {
			t.Errorf("%s value = %s, want %v", key, valueText(v), want)
		}
	}
}

// newStructureWorkbook builds the workbook of the row and column cases:
//
//	Sheet1: A1..A6 = 1..6, B1..B6 = 10..60, formulas in D and E
//	Other:  formulas that read Sheet1
func newStructureWorkbook(t *testing.T) *File {
	t.Helper()
	f := New()
	s := f.Sheet("Sheet1")
	other, err := f.NewSheet("Other")
	if err != nil {
		t.Fatal(err)
	}
	for r := 1; r <= 6; r++ {
		mustSetCells(t, s, fmt.Sprintf("A%d", r), r, fmt.Sprintf("B%d", r), r*10)
	}
	mustSetCells(t, s,
		"D1", "=SUM(A1:A6)",
		"D2", "=A5",
		"D3", "=$A$3+A$3+$A3",
		"D4", "=SUM(A:A)",
		"D5", "=SUM(3:3)",
		"D6", "=B2*2",
		"E1", "=SUM(A2:B4)",
		"E2", "=Other!A1+A1",
	)
	mustSetCells(t, other,
		"A1", "=Sheet1!A5",
		"A2", "=SUM(Sheet1!A1:A6)",
		"A3", "=A1+1",
		"A4", "=SUM(Sheet1!A:B)",
		"A5", "='Sheet1'!$B$2",
	)
	return f
}

// --- (a) rows, cells and formulas -----------------------------------------

func TestInsertRows(t *testing.T) {
	t.Run("in the data", func(t *testing.T) {
		f := newStructureWorkbook(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "insRows", 3, 2},
			map[string]string{
				// D1 was on row 1 and stays; its range grows.
				"Sheet1!D1": "SUM(A1:A8)",
				"Sheet1!D2": "A7",
				// Row 3 moved to row 5, with its formula.
				"Sheet1!D3": "", "Sheet1!D4": "",
				"Sheet1!D5": "$A$5+A$5+$A5",
				"Sheet1!D6": "SUM(A:A)", // a full column does not change
				"Sheet1!D7": "SUM(5:5)",
				"Sheet1!D8": "B2*2",
				"Sheet1!E1": "SUM(A2:B6)",
				"Sheet1!E2": "Other!A1+A1",
				// Formulas on a different sheet follow.
				"Other!A1": "Sheet1!A7",
				"Other!A2": "SUM(Sheet1!A1:A8)",
				"Other!A3": "A1+1", // refers to its own sheet: no change
				"Other!A4": "SUM(Sheet1!A:B)",
				"Other!A5": "'Sheet1'!$B$2",
			},
			map[string]any{
				"Sheet1!A2": 2, "Sheet1!A3": nil, "Sheet1!A4": nil, "Sheet1!A5": 3, "Sheet1!A8": 6,
				"Sheet1!B5": 30,
				"Sheet1!D1": 21, "Sheet1!D2": 5, "Sheet1!D5": 9, "Sheet1!D6": 21, "Sheet1!D8": 40,
				"Other!A1": 5, "Other!A2": 21, "Other!A4": 231,
			})
	})

	t.Run("at row 1", func(t *testing.T) {
		f := newStructureWorkbook(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "insRows", 1, 1},
			map[string]string{
				"Sheet1!D1": "", "Sheet1!D2": "SUM(A2:A7)", "Sheet1!D3": "A6",
				"Sheet1!D4": "$A$4+A$4+$A4",
				"Other!A2":  "SUM(Sheet1!A2:A7)",
			},
			map[string]any{"Sheet1!A1": nil, "Sheet1!A2": 1, "Sheet1!D2": 21, "Other!A2": 21})
	})

	t.Run("at the last used row: a range that ends there grows", func(t *testing.T) {
		f := newStructureWorkbook(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "insRows", 6, 1},
			map[string]string{"Sheet1!D1": "SUM(A1:A7)", "Sheet1!D7": "B2*2", "Sheet1!D6": ""},
			map[string]any{"Sheet1!A6": nil, "Sheet1!A7": 6, "Sheet1!D1": 21})
	})

	t.Run("below the data: nothing changes but the edit is valid", func(t *testing.T) {
		f := newStructureWorkbook(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "insRows", 7, 3},
			map[string]string{"Sheet1!D1": "SUM(A1:A6)", "Sheet1!D2": "A5"},
			map[string]any{"Sheet1!D1": 21})
	})

	t.Run("on a sheet that others read, formulas on that sheet follow", func(t *testing.T) {
		f := newStructureWorkbook(t)
		checkStructuralEdit(t, f, structOp{"Other", "insRows", 1, 2},
			map[string]string{
				"Other!A3": "Sheet1!A5", "Other!A5": "A3+1",
				// Sheet1!E2 reads Other!A1, which moved to A3.
				"Sheet1!E2": "Other!A3+A1",
				"Sheet1!D2": "A5", // a reference to Sheet1 itself: no change
			},
			map[string]any{"Other!A3": 5, "Sheet1!E2": 6})
	})
}

func TestDeleteRows(t *testing.T) {
	t.Run("in the data", func(t *testing.T) {
		f := newStructureWorkbook(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "delRows", 4, 2},
			map[string]string{
				"Sheet1!D1": "SUM(A1:A4)",
				"Sheet1!D2": "#REF!", // A5 is deleted
				"Sheet1!D3": "$A$3+A$3+$A3",
				"Sheet1!D4": "B2*2", // was D6
				"Sheet1!E1": "SUM(A2:B3)",
				"Other!A1":  "#REF!",
				"Other!A2":  "SUM(Sheet1!A1:A4)",
				"Other!A3":  "A1+1",
			},
			map[string]any{
				"Sheet1!A4": 6, "Sheet1!A5": nil,
				"Sheet1!D1": 12, "Sheet1!D2": "#REF!", "Sheet1!D4": 40,
				"Other!A1": "#REF!", "Other!A2": 12, "Other!A3": "#REF!",
			})
	})

	t.Run("the first row", func(t *testing.T) {
		f := newStructureWorkbook(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "delRows", 1, 1},
			map[string]string{
				"Sheet1!D1": "A4", // was D2 = A5
				"Sheet1!D2": "$A$2+A$2+$A2",
				"Sheet1!E1": "Other!A1+#REF!",
				"Other!A2":  "SUM(Sheet1!A1:A5)",
			},
			map[string]any{"Sheet1!A1": 2, "Sheet1!D1": 5, "Other!A2": 20})
	})

	t.Run("a formula with two references where one becomes #REF!", func(t *testing.T) {
		f := New()
		s := f.Sheet("Sheet1")
		mustSetCells(t, s, "A1", 1, "A2", 2, "A3", 3, "C1", "=A2+A3", "C5", `=IF(A2>0,"A2",A3)`)
		checkStructuralEdit(t, f, structOp{"Sheet1", "delRows", 2, 1},
			map[string]string{"Sheet1!C1": "#REF!+A2", "Sheet1!C4": `IF(#REF!>0,"A2",A2)`},
			map[string]any{"Sheet1!C1": "#REF!"})
	})

	t.Run("all used rows", func(t *testing.T) {
		f := newStructureWorkbook(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "delRows", 1, 6},
			map[string]string{"Other!A1": "#REF!", "Other!A2": "SUM(#REF!)", "Other!A4": "SUM(Sheet1!A:B)"},
			map[string]any{"Sheet1!A1": nil, "Other!A4": 0})
		if s := f.Sheet("Sheet1"); s.MaxRow() != 0 || s.MaxCol() != 0 {
			t.Errorf("size after all rows were deleted = %dx%d", s.MaxRow(), s.MaxCol())
		}
	})

	t.Run("rows below the data", func(t *testing.T) {
		f := newStructureWorkbook(t)
		u := checkStructuralEdit(t, f, structOp{"Sheet1", "delRows", 50, 10},
			map[string]string{"Sheet1!D1": "SUM(A1:A6)"}, map[string]any{"Sheet1!D1": 21})
		if u.Cost() != 0 {
			t.Errorf("Cost of an edit that changes nothing = %d", u.Cost())
		}
	})

	t.Run("the last rows of the sheet", func(t *testing.T) {
		f := New()
		s := f.Sheet("Sheet1")
		mustSetCells(t, s, "A1048576", "end", "B1", "=A1048576", "B2", "=SUM(A1048570:A1048576)")
		checkStructuralEdit(t, f, structOp{"Sheet1", "delRows", MaxRows - 1, 2},
			map[string]string{"Sheet1!B1": "#REF!", "Sheet1!B2": "SUM(A1048570:A1048574)"}, nil)
	})
}

// RemoveRow is DeleteRows(row, 1) for the callers that exist.
func TestRemoveRowAdjustsFormulas(t *testing.T) {
	f := newStructureWorkbook(t)
	s := f.Sheet("Sheet1")
	if err := s.RemoveRow(2); err != nil {
		t.Fatal(err)
	}
	checkCells(t, f,
		map[string]string{
			"Sheet1!D1": "SUM(A1:A5)",
			"Sheet1!D2": "$A$2+A$2+$A2", // row 3 moved up, with its formula
			"Sheet1!E1": "SUM(A2:B3)",
			"Other!A1":  "Sheet1!A4",
			"Other!A5":  "#REF!",
		},
		map[string]any{"Sheet1!A2": 3, "Sheet1!D1": 19, "Other!A1": 5})
	for _, row := range []int{0, -1, MaxRows + 1} {
		if err := s.RemoveRow(row); err == nil {
			t.Errorf("RemoveRow(%d) gave no error", row)
		}
	}
}

// --- (b) merged ranges and row heights ---------------------------------------

func TestStructuralEditMergesAndRowHeights(t *testing.T) {
	newFile := func(t *testing.T) *File {
		f := New()
		s := f.Sheet("Sheet1")
		for _, m := range [][2]string{{"B2", "C4"}, {"E5", "F5"}, {"A8", "A9"}, {"H1", "H2"}} {
			if err := s.MergeCell(m[0], m[1]); err != nil {
				t.Fatal(err)
			}
		}
		for row, h := range map[int]float64{1: 20, 3: 30, 6: 40, 10: 50} {
			if err := s.SetRowHeight(row, h); err != nil {
				t.Fatal(err)
			}
		}
		mustSetCells(t, s, "B2", "merged", "E5", "wide", "A8", "tall")
		return f
	}
	merges := func(f *File) string {
		var parts []string
		for _, mr := range f.Sheet("Sheet1").MergeCells() {
			parts = append(parts, mr.Start+":"+mr.End)
		}
		sort.Strings(parts)
		return strings.Join(parts, " ")
	}
	heights := func(f *File) string {
		s := f.Sheet("Sheet1")
		var parts []string
		for row := 1; row <= 20; row++ {
			if h, _ := s.GetRowHeight(row); h != 0 {
				parts = append(parts, fmt.Sprintf("%d=%v", row, h))
			}
		}
		return strings.Join(parts, " ")
	}

	tests := []struct {
		name        string
		op          structOp
		wantMerges  string
		wantHeights string
	}{
		{"insert rows in a merged range: it grows", structOp{"Sheet1", "insRows", 3, 2}, "A10:A11 B2:C6 E7:F7 H1:H2", "1=20 5=30 8=40 12=50"},
		{"insert rows at the start of a merged range: it moves", structOp{"Sheet1", "insRows", 2, 1}, "A9:A10 B3:C5 E6:F6 H1:H3", "1=20 4=30 7=40 11=50"},
		{"insert rows below all", structOp{"Sheet1", "insRows", 11, 4}, "A8:A9 B2:C4 E5:F5 H1:H2", "1=20 3=30 6=40 10=50"},
		{"delete a row of a merged range: it shrinks", structOp{"Sheet1", "delRows", 3, 1}, "A7:A8 B2:C3 E4:F4 H1:H2", "1=20 5=40 9=50"},
		{"delete rows so that a range has one row but two columns", structOp{"Sheet1", "delRows", 3, 2}, "A6:A7 B2:C2 E3:F3 H1:H2", "1=20 4=40 8=50"},
		{"delete one of two rows of a one-column range: it is removed", structOp{"Sheet1", "delRows", 8, 1}, "B2:C4 E5:F5 H1:H2", "1=20 3=30 6=40 9=50"},
		{"delete all rows of a merged range: it is removed", structOp{"Sheet1", "delRows", 2, 3}, "A5:A6 E2:F2", "1=20 3=40 7=50"},
		{"insert columns in a merged range: it grows", structOp{"Sheet1", "insCols", 3, 1}, "A8:A9 B2:D4 F5:G5 I1:I2", "1=20 3=30 6=40 10=50"},
		{"delete a column of a merged range", structOp{"Sheet1", "delCols", 2, 1}, "A8:A9 B2:B4 D5:E5 G1:G2", "1=20 3=30 6=40 10=50"},
		{"delete a column so that a one-row range has one cell: it is removed", structOp{"Sheet1", "delCols", 5, 1}, "A8:A9 B2:C4 G1:G2", "1=20 3=30 6=40 10=50"},
		{"delete the column of one-column ranges", structOp{"Sheet1", "delCols", 1, 1}, "A2:B4 D5:E5 G1:G2", "1=20 3=30 6=40 10=50"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := newFile(t)
			checkStructuralEdit(t, f, tt.op, nil, nil)
			if got := merges(f); got != tt.wantMerges {
				t.Errorf("merges = %s, want %s", got, tt.wantMerges)
			}
			if got := heights(f); got != tt.wantHeights {
				t.Errorf("row heights = %s, want %s", got, tt.wantHeights)
			}
		})
	}

	t.Run("a hidden row and a row with only a height move with their number", func(t *testing.T) {
		f := New()
		s := f.Sheet("Sheet1")
		mustSetCells(t, s, "A5", "x")
		s.rows[5].hidden = true
		if err := s.SetRowHeight(9, 33); err != nil {
			t.Fatal(err)
		}
		checkStructuralEdit(t, f, structOp{"Sheet1", "insRows", 2, 3}, nil, map[string]any{"Sheet1!A8": "x"})
		if r := s.rows[8]; r == nil || !r.hidden {
			t.Errorf("the hidden flag did not move to row 8")
		}
		if h, _ := s.GetRowHeight(12); h != 33 {
			t.Errorf("height of row 12 = %v, want 33", h)
		}
	})
}

// --- (c) columns and column widths -------------------------------------------

func TestInsertAndDeleteCols(t *testing.T) {
	newFile := func(t *testing.T) *File {
		f := New()
		s := f.Sheet("Sheet1")
		other, err := f.NewSheet("Other")
		if err != nil {
			t.Fatal(err)
		}
		// A1..F1 = 1..6, formulas on row 3 and 4.
		for c := 1; c <= 6; c++ {
			ref, _ := CoordinatesToCellName(c, 1)
			mustSetCells(t, s, ref, c)
		}
		mustSetCells(t, s,
			"A3", "=SUM(A1:F1)", "B3", "=E1", "C3", "=$C$1+C$1+$C1", "D3", "=SUM(1:1)", "E3", "=SUM(C:C)", "F3", "=B1*2",
			"A4", "=SUM(B1:D2)")
		mustSetCells(t, other, "A1", "=Sheet1!E1", "A2", "=SUM(Sheet1!A1:F1)", "A3", "=SUM(Sheet1!1:1)", "A4", "=B1")
		for col, w := range map[string]float64{"A": 5, "C": 15, "F": 25} {
			if err := s.SetColumnWidth(col, w); err != nil {
				t.Fatal(err)
			}
		}
		if err := s.SetStyle("C1", &Style{Font: &Font{Bold: true}}); err != nil {
			t.Fatal(err)
		}
		return f
	}
	widths := func(f *File) string {
		w := f.Sheet("Sheet1").ColumnWidths()
		var parts []string
		for c := 1; c <= 12; c++ {
			if w[c] != 0 {
				parts = append(parts, fmt.Sprintf("%s=%v", ColumnNumberToName(c), w[c]))
			}
		}
		return strings.Join(parts, " ")
	}

	t.Run("insert columns in the data", func(t *testing.T) {
		f := newFile(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "insCols", 3, 2},
			map[string]string{
				"Sheet1!A3": "SUM(A1:H1)",
				"Sheet1!B3": "G1",
				"Sheet1!C3": "", "Sheet1!D3": "",
				"Sheet1!E3": "$E$1+E$1+$E1",
				"Sheet1!F3": "SUM(1:1)", // a full row does not change
				"Sheet1!G3": "SUM(E:E)",
				"Sheet1!H3": "B1*2",
				"Sheet1!A4": "SUM(B1:F2)",
				"Other!A1":  "Sheet1!G1",
				"Other!A2":  "SUM(Sheet1!A1:H1)",
				"Other!A3":  "SUM(Sheet1!1:1)",
				"Other!A4":  "B1",
			},
			map[string]any{
				"Sheet1!C1": nil, "Sheet1!D1": nil, "Sheet1!E1": 3, "Sheet1!H1": 6,
				"Sheet1!A3": 21, "Sheet1!B3": 5, "Sheet1!E3": 9, "Sheet1!F3": 21,
				"Other!A1": 5, "Other!A2": 21,
			})
		if got := widths(f); got != "A=5 E=15 H=25" {
			t.Errorf("widths = %s, want A=5 E=15 H=25", got)
		}
		if st, _ := f.Sheet("Sheet1").GetStyle("E1"); st == nil || !st.Font.Bold {
			t.Errorf("the style did not move with its cell")
		}
	})

	t.Run("insert a column at A", func(t *testing.T) {
		f := newFile(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "insCols", 1, 1},
			map[string]string{"Sheet1!B3": "SUM(B1:G1)", "Sheet1!C3": "F1", "Sheet1!A3": "", "Other!A1": "Sheet1!F1"},
			map[string]any{"Sheet1!A1": nil, "Sheet1!B1": 1, "Sheet1!B3": 21, "Other!A1": 5})
		if got := widths(f); got != "B=5 D=15 G=25" {
			t.Errorf("widths = %s", got)
		}
	})

	t.Run("delete columns in the data", func(t *testing.T) {
		f := newFile(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "delCols", 3, 2},
			map[string]string{
				"Sheet1!A3": "SUM(A1:D1)",
				"Sheet1!B3": "C1",         // E1 moved to C1
				"Sheet1!C3": "SUM(#REF!)", // was E3 = SUM(C:C)
				"Sheet1!D3": "B1*2",
				"Sheet1!A4": "SUM(B1:B2)",
				"Other!A1":  "Sheet1!C1",
				"Other!A2":  "SUM(Sheet1!A1:D1)",
			},
			map[string]any{"Sheet1!C1": 5, "Sheet1!D1": 6, "Sheet1!A3": 14, "Sheet1!B3": 5, "Other!A2": 14})
		if got := widths(f); got != "A=5 D=25" {
			t.Errorf("widths = %s, want A=5 D=25", got)
		}
	})

	t.Run("delete the column that a formula reads", func(t *testing.T) {
		f := newFile(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "delCols", 5, 1},
			map[string]string{"Sheet1!B3": "#REF!", "Other!A1": "#REF!", "Sheet1!A3": "SUM(A1:E1)"},
			map[string]any{"Sheet1!B3": "#REF!", "Other!A1": "#REF!", "Sheet1!A3": 16})
	})

	t.Run("delete all used columns", func(t *testing.T) {
		f := newFile(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "delCols", 1, 6},
			map[string]string{"Other!A1": "#REF!", "Other!A3": "SUM(Sheet1!1:1)"},
			map[string]any{"Other!A3": 0})
		if s := f.Sheet("Sheet1"); s.MaxRow() != 0 || len(s.ColumnWidths()) != 0 {
			t.Errorf("after all columns were deleted: MaxRow %d, widths %v", s.MaxRow(), s.ColumnWidths())
		}
	})

	t.Run("columns right of the data", func(t *testing.T) {
		f := newFile(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "insCols", 20, 3}, map[string]string{"Sheet1!A3": "SUM(A1:F1)"}, nil)
		checkStructuralEdit(t, f, structOp{"Sheet1", "delCols", 20, 3}, map[string]string{"Sheet1!A3": "SUM(A1:F1)"}, nil)
	})
}

// --- errors ---------------------------------------------------------------------

func TestStructuralEditErrorsLeaveTheWorkbookUnchanged(t *testing.T) {
	f := newStructureWorkbook(t)
	s := f.Sheet("Sheet1")
	mustSetCells(t, s, "A1048576", "last row", "XFD1", "last column")
	if err := s.MergeCell("G1048570", "H1048575"); err != nil {
		t.Fatal(err)
	}
	if err := s.MergeCell("XFA9", "XFC9"); err != nil {
		t.Fatal(err)
	}
	before := dumpWorkbook(t, f)

	tests := []struct {
		name string
		op   structOp
		want string
	}{
		{"row 0", structOp{"Sheet1", "insRows", 0, 1}, "out of range"},
		{"row past the end", structOp{"Sheet1", "delRows", MaxRows + 1, 1}, "out of range"},
		{"column 0", structOp{"Sheet1", "insCols", 0, 1}, "out of range"},
		{"column past the end", structOp{"Sheet1", "delCols", MaxColumns + 1, 1}, "out of range"},
		{"count 0", structOp{"Sheet1", "insRows", 1, 0}, "do not fit"},
		{"negative count", structOp{"Sheet1", "delCols", 1, -2}, "do not fit"},
		{"delete past the last row", structOp{"Sheet1", "delRows", MaxRows, 2}, "do not fit"},
		{"delete past the last column", structOp{"Sheet1", "delCols", MaxColumns - 1, 3}, "do not fit"},
		{"insert pushes a value off the last row", structOp{"Sheet1", "insRows", 1, 1}, "A1048576 is not empty"},
		{"insert pushes a value off the last column", structOp{"Sheet1", "insCols", 5, 2}, "XFD1 is not empty"},
		{"insert below the value still pushes a merged range", structOp{"Sheet1", "insRows", 1048570, 1}, "is not empty"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			u, err := tt.op.run(f)
			if err == nil || u != nil {
				t.Fatalf("%+v gave no error", tt.op)
			}
			if !strings.Contains(err.Error(), tt.want) {
				t.Errorf("error = %q, want it to contain %q", err, tt.want)
			}
			if got := dumpWorkbook(t, f); got != before {
				t.Errorf("the workbook changed after a failed edit")
			}
		})
	}

	t.Run("merged range that would move off the sheet", func(t *testing.T) {
		f := New()
		s := f.Sheet("Sheet1")
		if err := s.MergeCell("A1048570", "B1048575"); err != nil {
			t.Fatal(err)
		}
		if err := s.MergeCell("XFA2", "XFC3"); err != nil {
			t.Fatal(err)
		}
		before := dumpWorkbook(t, f)
		// One row fits (the range ends at the last row); two do not.
		if _, err := s.InsertRows(5, 2); err == nil || !strings.Contains(err.Error(), "merged range") {
			t.Errorf("InsertRows: error = %v, want a merged range error", err)
		}
		if _, err := s.InsertCols(1, 2); err == nil || !strings.Contains(err.Error(), "merged range") {
			t.Errorf("InsertCols: error = %v, want a merged range error", err)
		}
		// An insert in the range that would grow it past the end.
		if _, err := s.InsertRows(1048572, 2); err == nil || !strings.Contains(err.Error(), "merged range") {
			t.Errorf("InsertRows in the range: error = %v, want a merged range error", err)
		}
		if got := dumpWorkbook(t, f); got != before {
			t.Errorf("the workbook changed after a failed edit")
		}
		checkStructuralEdit(t, f, structOp{"Sheet1", "insRows", 5, 1}, nil, nil)
		checkStructuralEdit(t, f, structOp{"Sheet1", "insCols", 1, 1}, nil, nil)
	})

	t.Run("style-only cells at the end are dropped by an insert and come back with Undo", func(t *testing.T) {
		f := New()
		s := f.Sheet("Sheet1")
		st := &Style{Fill: &Fill{Color: "FF0000"}}
		for _, ref := range []string{"A1048576", "B1048575", "XFD3"} {
			if err := s.SetStyle(ref, st); err != nil {
				t.Fatal(err)
			}
		}
		if err := s.SetRowHeight(MaxRows, 44); err != nil {
			t.Fatal(err)
		}
		mustSetCells(t, s, "A1", "x")
		u := checkStructuralEdit(t, f, structOp{"Sheet1", "insRows", 1, 2}, nil, map[string]any{"Sheet1!A3": "x"})
		// The styled cell XFD3 moved to row 5; the cells at the end are gone.
		if s.MaxRow() != 5 {
			t.Errorf("MaxRow after the insert = %d, want 5", s.MaxRow())
		}
		if u.Cost() != 2 {
			t.Errorf("Cost = %d, want 2 (the two dropped cells)", u.Cost())
		}
		checkStructuralEdit(t, f, structOp{"Sheet1", "insCols", 2, 1}, nil, nil)
		if s.MaxCol() >= MaxColumns {
			t.Errorf("MaxCol after the insert = %d; the style-only cell in XFD was not dropped", s.MaxCol())
		}
	})
}

// --- (d) defined names --------------------------------------------------------

func TestStructuralEditDefinedNames(t *testing.T) {
	newFile := func(t *testing.T) *File {
		f := newStructureWorkbook(t)
		names := []DefinedName{
			{Name: "Data", Value: "Sheet1!$A$1:$B$6", LocalSheetID: -1},
			{Name: "Middle", Value: "Sheet1!$A$3:$A$4", LocalSheetID: -1},
			{Name: "OneCell", Value: "Sheet1!$B$5", LocalSheetID: -1},
			{Name: "Print_Titles", Value: "Sheet1!$1:$2", LocalSheetID: 0},
			{Name: "ColB", Value: "Sheet1!$B:$B", LocalSheetID: -1},
			{Name: "OnOther", Value: "Other!$A$1:$A$3", LocalSheetID: -1},
			{Name: "Local", Value: "$A$5", LocalSheetID: 0},      // no sheet name, scope Sheet1
			{Name: "LocalOther", Value: "$A$5", LocalSheetID: 1}, // no sheet name, scope Other
			{Name: "Rate", Value: "0.25", LocalSheetID: -1},
			{Name: "Quoted", Value: "'Sheet1'!$A$6", LocalSheetID: -1},
		}
		for _, dn := range names {
			if err := f.SetDefinedName(dn); err != nil {
				t.Fatal(err)
			}
		}
		mustSetCells(t, f.Sheet("Other"), "C1", "=SUM(Data)", "C2", "=SUM(Middle)", "C3", "=OneCell*Rate")
		return f
	}
	values := func(f *File) map[string]string {
		out := map[string]string{}
		for _, dn := range f.DefinedNames() {
			out[dn.Name] = dn.Value
		}
		return out
	}
	check := func(t *testing.T, f *File, want map[string]string) {
		t.Helper()
		got := values(f)
		for name, w := range want {
			if got[name] != w {
				t.Errorf("name %s = %q, want %q", name, got[name], w)
			}
		}
	}

	t.Run("insert rows: ranges grow and move", func(t *testing.T) {
		f := newFile(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "insRows", 4, 2}, nil,
			map[string]any{"Other!C1": 231, "Other!C2": 7, "Other!C3": 12.5})
		check(t, f, map[string]string{
			"Data": "Sheet1!$A$1:$B$8", "Middle": "Sheet1!$A$3:$A$6", "OneCell": "Sheet1!$B$7",
			"Print_Titles": "Sheet1!$1:$2", "ColB": "Sheet1!$B:$B", "OnOther": "Other!$A$1:$A$3",
			"Local": "$A$7", "LocalOther": "$A$5", "Rate": "0.25", "Quoted": "'Sheet1'!$A$8",
		})
	})

	t.Run("insert rows in the print titles", func(t *testing.T) {
		f := newFile(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "insRows", 2, 1}, nil, nil)
		check(t, f, map[string]string{"Print_Titles": "Sheet1!$1:$3"})
		f = newFile(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "insRows", 1, 1}, nil, nil)
		check(t, f, map[string]string{"Print_Titles": "Sheet1!$2:$3"})
	})

	t.Run("delete rows: a range shrinks, one becomes #REF!", func(t *testing.T) {
		f := newFile(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "delRows", 3, 2}, nil,
			map[string]any{"Other!C1": 154, "Other!C2": "#REF!", "Other!C3": 12.5})
		check(t, f, map[string]string{
			"Data": "Sheet1!$A$1:$B$4", "Middle": "#REF!", "OneCell": "Sheet1!$B$3",
			"Print_Titles": "Sheet1!$1:$2", "Local": "$A$3", "LocalOther": "$A$5", "Quoted": "'Sheet1'!$A$4",
		})
	})

	t.Run("delete the cell of a name", func(t *testing.T) {
		f := newFile(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "delRows", 5, 1}, nil, map[string]any{"Other!C3": "#REF!"})
		check(t, f, map[string]string{"OneCell": "#REF!", "Local": "#REF!", "Data": "Sheet1!$A$1:$B$5"})
	})

	t.Run("columns", func(t *testing.T) {
		f := newFile(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "insCols", 2, 1}, nil, map[string]any{"Other!C1": 231})
		check(t, f, map[string]string{
			"Data": "Sheet1!$A$1:$C$6", "OneCell": "Sheet1!$C$5", "ColB": "Sheet1!$C:$C", "Print_Titles": "Sheet1!$1:$2", "Local": "$A$5",
		})
		f = newFile(t)
		checkStructuralEdit(t, f, structOp{"Sheet1", "delCols", 2, 1}, nil, map[string]any{"Other!C1": 21})
		check(t, f, map[string]string{"Data": "Sheet1!$A$1:$A$6", "OneCell": "#REF!", "ColB": "#REF!", "Middle": "Sheet1!$A$3:$A$4"})
	})

	t.Run("an edit on a different sheet changes only its names", func(t *testing.T) {
		f := newFile(t)
		checkStructuralEdit(t, f, structOp{"Other", "insRows", 1, 2}, nil, nil)
		check(t, f, map[string]string{
			"Data": "Sheet1!$A$1:$B$6", "OnOther": "Other!$A$3:$A$5", "Local": "$A$5", "LocalOther": "$A$7",
		})
	})
}

// --- (e) tables -------------------------------------------------------------------

func TestStructuralEditTables(t *testing.T) {
	// Table "Sales" at B3:D7: header row 3, data rows 4..6, totals row 7.
	newFile := func(t *testing.T) *File {
		f := New()
		s := f.Sheet("Sheet1")
		mustSetCells(t, s,
			"B3", "Item", "C3", "Qty", "D3", "Price",
			"B4", "a", "C4", 1, "D4", 10,
			"B5", "b", "C5", 2, "D5", 20,
			"B6", "c", "C6", 3, "D6", 30,
			"B7", "Total", "C7", 6, "D7", 60,
			"F1", "=SUM(Sales[Qty])", "F2", "=SUM(Sales[Price])", "F3", "=C5",
		)
		if err := s.AddTable(Table{Name: "Sales", Ref: "B3:D7", TotalsRowCount: 1}); err != nil {
			t.Fatal(err)
		}
		other, err := f.NewSheet("Other")
		if err != nil {
			t.Fatal(err)
		}
		mustSetCells(t, other, "A1", "x", "B1", "y", "A2", 1, "B2", 2, "D1", "=SUM(Sales[Qty])")
		if err := other.AddTable(Table{Name: "Small", Ref: "A1:B2"}); err != nil {
			t.Fatal(err)
		}
		return f
	}
	table := func(t *testing.T, f *File, name string) (Table, bool) {
		t.Helper()
		for _, td := range f.Tables() {
			if td.DisplayName == name {
				return td, true
			}
		}
		return Table{}, false
	}
	type want struct {
		ref    string
		cols   string
		totals int
	}
	check := func(t *testing.T, f *File, w want) {
		t.Helper()
		td, ok := table(t, f, "Sales")
		if !ok {
			t.Fatalf("table Sales is gone")
		}
		if td.Ref != w.ref || strings.Join(td.Columns, ",") != w.cols || td.TotalsRowCount != w.totals {
			t.Errorf("table = ref %s cols %v totals %d, want %+v", td.Ref, td.Columns, td.TotalsRowCount, w)
		}
		// The table on the other sheet never changes.
		if small, ok := table(t, f, "Small"); !ok || small.Ref != "A1:B2" {
			t.Errorf("table Small = %+v, %v", small, ok)
		}
	}

	tests := []struct {
		name   string
		op     structOp
		want   want
		values map[string]any
	}{
		{"insert rows above: the table moves", structOp{"Sheet1", "insRows", 2, 2}, want{"B5:D9", "Item,Qty,Price", 1}, map[string]any{"Sheet1!F1": 6, "Sheet1!F4": 60, "Other!D1": 6}},
		{"insert rows at the header row: the table moves", structOp{"Sheet1", "insRows", 3, 1}, want{"B4:D8", "Item,Qty,Price", 1}, map[string]any{"Sheet1!F1": 6}},
		{"insert rows in the data: the table grows", structOp{"Sheet1", "insRows", 5, 2}, want{"B3:D9", "Item,Qty,Price", 1}, map[string]any{"Sheet1!F1": 6, "Sheet1!F2": 60}},
		{"insert rows below: no change", structOp{"Sheet1", "insRows", 8, 2}, want{"B3:D7", "Item,Qty,Price", 1}, map[string]any{"Sheet1!F1": 6}},
		{"delete a data row: the table shrinks", structOp{"Sheet1", "delRows", 5, 1}, want{"B3:D6", "Item,Qty,Price", 1}, map[string]any{"Sheet1!F1": 4, "Sheet1!F2": 40}},
		{"delete the totals row", structOp{"Sheet1", "delRows", 7, 1}, want{"B3:D6", "Item,Qty,Price", 0}, map[string]any{"Sheet1!F1": 6, "Sheet1!F2": 60}},
		// Row 1 had F1; F2 (the Price sum) is now F1.
		{"delete rows above: the table moves", structOp{"Sheet1", "delRows", 1, 1}, want{"B2:D6", "Item,Qty,Price", 1}, map[string]any{"Sheet1!F1": 60, "Other!D1": 6}},
		{"insert a column left: the table moves", structOp{"Sheet1", "insCols", 2, 1}, want{"C3:E7", "Item,Qty,Price", 1}, map[string]any{"Sheet1!G1": 6, "Sheet1!G2": 60}},
		{"insert a column in the table: a new column with a unique name", structOp{"Sheet1", "insCols", 3, 1}, want{"B3:E7", "Item,Column1,Qty,Price", 1}, map[string]any{"Sheet1!G1": 6, "Sheet1!C3": "Column1", "Sheet1!D3": "Qty"}},
		{"insert two columns in the table", structOp{"Sheet1", "insCols", 4, 2}, want{"B3:F7", "Item,Qty,Column1,Column2,Price", 1}, map[string]any{"Sheet1!D3": "Column1", "Sheet1!E3": "Column2", "Sheet1!D4": nil}},
		{"insert a column right of the table: no change", structOp{"Sheet1", "insCols", 5, 1}, want{"B3:D7", "Item,Qty,Price", 1}, nil},
		{"delete a column of the table", structOp{"Sheet1", "delCols", 3, 1}, want{"B3:C7", "Item,Price", 1}, map[string]any{"Sheet1!E2": 60}},
		{"delete the first column of the table", structOp{"Sheet1", "delCols", 2, 1}, want{"B3:C7", "Qty,Price", 1}, map[string]any{"Sheet1!E1": 6}},
		{"delete two columns of the table", structOp{"Sheet1", "delCols", 3, 2}, want{"B3:B7", "Item", 1}, nil},
		{"an edit on a different sheet", structOp{"Other", "insRows", 5, 1}, want{"B3:D7", "Item,Qty,Price", 1}, map[string]any{"Other!D1": 6}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := newFile(t)
			checkStructuralEdit(t, f, tt.op, nil, tt.values)
			check(t, f, tt.want)
		})
	}

	t.Run("a table that is fully in the deleted band is removed", func(t *testing.T) {
		for _, op := range []structOp{{"Sheet1", "delRows", 3, 5}, {"Sheet1", "delRows", 2, 7}, {"Sheet1", "delCols", 2, 3}, {"Sheet1", "delCols", 1, 5}} {
			f := newFile(t)
			checkStructuralEdit(t, f, op, nil, nil)
			if _, ok := table(t, f, "Sales"); ok {
				t.Errorf("%+v: the table is still in the workbook", op)
			}
			if len(f.Tables()) != 1 || len(f.tables) != 1 {
				t.Errorf("%+v: tables = %d, infos = %d, want 1 and 1", op, len(f.Tables()), len(f.tables))
			}
		}
	})

	t.Run("a unique column name when Column1 is in use", func(t *testing.T) {
		f := newFile(t)
		s := f.Sheet("Sheet1")
		if _, err := s.InsertCols(3, 1); err != nil {
			t.Fatal(err)
		}
		if _, err := s.InsertCols(3, 1); err != nil {
			t.Fatal(err)
		}
		td, _ := table(t, f, "Sales")
		if got := strings.Join(td.Columns, ","); got != "Item,Column2,Column1,Qty,Price" {
			t.Errorf("columns = %s", got)
		}
	})

	t.Run("errors", func(t *testing.T) {
		f := newFile(t)
		before := dumpWorkbook(t, f)
		bad := []struct {
			op   structOp
			want string
		}{
			{structOp{"Sheet1", "delRows", 3, 1}, `cannot delete the header row of table "Sales"`},
			{structOp{"Sheet1", "delRows", 2, 3}, `cannot delete the header row of table "Sales"`},
			{structOp{"Sheet1", "delRows", 3, 4}, `cannot delete the header row of table "Sales"`},
			{structOp{"Sheet1", "delRows", 4, 3}, `cannot delete all data rows of table "Sales"`},
			{structOp{"Sheet1", "delRows", 4, 4}, `cannot delete all data rows of table "Sales"`},
			{structOp{"Other", "delRows", 1, 1}, `cannot delete the header row of table "Small"`},
			{structOp{"Other", "delRows", 2, 1}, `cannot delete all data rows of table "Small"`},
		}
		for _, b := range bad {
			if _, err := b.op.run(f); err == nil || !strings.Contains(err.Error(), b.want) {
				t.Errorf("%+v: error = %v, want %q", b.op, err, b.want)
			}
			if got := dumpWorkbook(t, f); got != before {
				t.Fatalf("%+v changed the workbook", b.op)
			}
		}
	})

	t.Run("a table at the end of the sheet cannot move off", func(t *testing.T) {
		f := New()
		s := f.Sheet("Sheet1")
		mustSetCells(t, s, "A1048574", "h")
		if err := s.AddTable(Table{Name: "Low", Ref: "A1048574:B1048575"}); err != nil {
			t.Fatal(err)
		}
		if _, err := s.InsertRows(1, 2); err == nil || !strings.Contains(err.Error(), "table") {
			t.Errorf("error = %v, want a table error", err)
		}
		checkStructuralEdit(t, f, structOp{"Sheet1", "insRows", 1, 1}, nil, nil)
	})
}

// --- (f) spill ranges -----------------------------------------------------------

func TestStructuralEditSpill(t *testing.T) {
	t.Run("an anchor that moves still spills", func(t *testing.T) {
		f := New()
		s := f.Sheet("Sheet1")
		mustSetCells(t, s, "B2", "=SEQUENCE(3)", "D1", "=SUM(B2:B4)", "D2", "=B4")
		f.Recalculate()
		checkStructuralEdit(t, f, structOp{"Sheet1", "insRows", 1, 2},
			map[string]string{"Sheet1!B4": "SEQUENCE(3)", "Sheet1!D3": "SUM(B4:B6)", "Sheet1!D4": "B6"},
			map[string]any{"Sheet1!B4": 1, "Sheet1!B5": 2, "Sheet1!B6": 3, "Sheet1!B2": nil, "Sheet1!D3": 6, "Sheet1!D4": 3})
		f.Recalculate()
		if toCol, toRow, ok := s.SpillBounds(2, 4); !ok || toCol != 2 || toRow != 6 {
			t.Errorf("SpillBounds at the new anchor = %d,%d,%v, want 2,6,true", toCol, toRow, ok)
		}
		if _, _, ok := s.SpillBounds(2, 2); ok {
			t.Errorf("the old anchor position still has a spill state")
		}
	})

	t.Run("columns", func(t *testing.T) {
		f := New()
		s := f.Sheet("Sheet1")
		mustSetCells(t, s, "B2", "=SEQUENCE(1,3)", "A5", "=SUM(B2:D2)")
		checkStructuralEdit(t, f, structOp{"Sheet1", "insCols", 1, 1},
			map[string]string{"Sheet1!C2": "SEQUENCE(1,3)", "Sheet1!B5": "SUM(C2:E2)"},
			map[string]any{"Sheet1!C2": 1, "Sheet1!E2": 3, "Sheet1!B5": 6})
		checkStructuralEdit(t, f, structOp{"Sheet1", "delCols", 1, 2},
			map[string]string{"Sheet1!A2": "SEQUENCE(1,3)"},
			map[string]any{"Sheet1!A2": 1, "Sheet1!C2": 3})
	})

	t.Run("an insert in a spill range leaves the spill whole", func(t *testing.T) {
		f := New()
		s := f.Sheet("Sheet1")
		mustSetCells(t, s, "A1", "=SEQUENCE(4)", "C1", "=A4")
		checkStructuralEdit(t, f, structOp{"Sheet1", "insRows", 3, 1},
			map[string]string{"Sheet1!C1": "A5"},
			// The anchor stays; its result still has four rows from A1.
			map[string]any{"Sheet1!A1": 1, "Sheet1!A3": 3, "Sheet1!A4": 4, "Sheet1!A5": nil})
	})

	t.Run("a blocked spill is freed by a delete of the blocker row", func(t *testing.T) {
		f := New()
		s := f.Sheet("Sheet1")
		mustSetCells(t, s, "A1", "=SEQUENCE(1,3)", "C1", "x", "E5", "keep")
		// C1 blocks the spill to the right; move the blocker down first.
		if v, _ := s.GetValue("A1"); v.String != "#SPILL!" {
			t.Fatalf("test setup: A1 = %s, want #SPILL!", valueText(v))
		}
		checkStructuralEdit(t, f, structOp{"Sheet1", "delCols", 3, 1},
			nil, map[string]any{"Sheet1!A1": 1, "Sheet1!B1": 2, "Sheet1!C1": 3, "Sheet1!D5": "keep"})
	})

	t.Run("a deleted anchor removes its spill", func(t *testing.T) {
		f := New()
		s := f.Sheet("Sheet1")
		mustSetCells(t, s, "A2", "=SEQUENCE(3)", "C1", "=SUM(A2:A4)")
		checkStructuralEdit(t, f, structOp{"Sheet1", "delRows", 2, 1},
			map[string]string{"Sheet1!C1": "SUM(A2:A3)"},
			map[string]any{"Sheet1!A2": nil, "Sheet1!A3": nil, "Sheet1!C1": 0})
		if len(s.spill.anchors) != 0 {
			t.Errorf("spill anchors after the delete = %v", s.spill.anchors)
		}
	})

	t.Run("the imported spill range moves with its anchor", func(t *testing.T) {
		f := New()
		s := f.Sheet("Sheet1")
		mustSetCells(t, s, "B2", "=SEQUENCE(3)")
		s.setSpillFormulaRef(2, 2, "B2:B4")
		if _, err := s.InsertRows(1, 3); err != nil {
			t.Fatal(err)
		}
		if state, ok := s.spillState(2, 5); !ok || state.formulaRef != "B5:B7" {
			t.Errorf("spill state at B5 = %+v, %v; want the ref B5:B7", state, ok)
		}
		if _, err := s.InsertCols(1, 1); err != nil {
			t.Fatal(err)
		}
		if state, ok := s.spillState(3, 5); !ok || state.formulaRef != "C5:C7" {
			t.Errorf("spill state at C5 = %+v, %v; want the ref C5:C7", state, ok)
		}
	})
}

// --- (g) undo ---------------------------------------------------------------------

func TestStructuralUndoRestoresPointersAndStyles(t *testing.T) {
	f := newStructureWorkbook(t)
	s := f.Sheet("Sheet1")
	bold := &Style{Font: &Font{Bold: true}}
	if err := s.SetRangeStyle("A3:B4", bold); err != nil {
		t.Fatal(err)
	}
	cellBefore := s.rows[4].cells[1]

	u, err := s.DeleteRows(3, 2)
	if err != nil {
		t.Fatal(err)
	}
	// 4 cells of A3:B4, D3 and D4 are removed; the changed formulas are
	// D1, D2, D5 (SUM(3:3)), E1 on Sheet1 and A1, A2 on Other.
	if got := u.Cost(); got != 6+6 {
		t.Errorf("Cost = %d, want 12", got)
	}
	if err := u.Undo(); err != nil {
		t.Fatal(err)
	}
	// The same cell objects are back, so the shared style pointer is too.
	if s.rows[4].cells[1] != cellBefore {
		t.Errorf("Undo made a new cell object for A4")
	}
	if st, _ := s.GetStyle("A4"); st != bold {
		t.Errorf("style pointer after Undo = %p, want %p", st, bold)
	}

	// Undo on a sheet that is not in the workbook any more is an error.
	u, err = s.InsertRows(1, 1)
	if err != nil {
		t.Fatal(err)
	}
	if err := f.DeleteSheet("Sheet1"); err != nil {
		t.Fatal(err)
	}
	if err := u.Undo(); err == nil {
		t.Errorf("Undo for a deleted sheet gave no error")
	}
	var none *StructuralUndo
	if err := none.Undo(); err == nil {
		t.Errorf("Undo of a nil value gave no error")
	}
}

// The undo data of an edit is valid only for the state immediately after
// the edit. Undo must refuse a workbook that changed, and must not damage it.
func TestStructuralUndoFailsAfterALaterChange(t *testing.T) {
	changes := []struct {
		name   string
		change func(t *testing.T, f *File)
	}{
		{"a second structural edit", func(t *testing.T, f *File) {
			if _, err := f.Sheet("Sheet1").InsertRows(1, 1); err != nil {
				t.Fatal(err)
			}
		}},
		{"a structural edit on a different sheet", func(t *testing.T, f *File) {
			if _, err := f.Sheet("Other").DeleteCols(1, 1); err != nil {
				t.Fatal(err)
			}
		}},
		{"a value", func(t *testing.T, f *File) { mustSetCells(t, f.Sheet("Sheet1"), "H9", 5) }},
		{"a formula", func(t *testing.T, f *File) { mustSetCells(t, f.Sheet("Other"), "H9", "=1+1") }},
		{"a style", func(t *testing.T, f *File) {
			if err := f.Sheet("Sheet1").SetStyle("A1", &Style{Font: &Font{Bold: true}}); err != nil {
				t.Fatal(err)
			}
		}},
		{"a range style", func(t *testing.T, f *File) {
			if err := f.Sheet("Sheet1").SetRangeStyle("A1:B2", &Style{Font: &Font{Bold: true}}); err != nil {
				t.Fatal(err)
			}
		}},
		{"a restored cell", func(t *testing.T, f *File) {
			if err := f.Sheet("Sheet1").RestoreCell(1, 1, CellSnapshot{}); err != nil {
				t.Fatal(err)
			}
		}},
		{"a merged range", func(t *testing.T, f *File) {
			if err := f.Sheet("Sheet1").MergeCell("J1", "K2"); err != nil {
				t.Fatal(err)
			}
		}},
		{"a column width", func(t *testing.T, f *File) {
			if err := f.Sheet("Sheet1").SetColumnWidth("A", 30); err != nil {
				t.Fatal(err)
			}
		}},
		{"a row height", func(t *testing.T, f *File) {
			if err := f.Sheet("Sheet1").SetRowHeight(2, 30); err != nil {
				t.Fatal(err)
			}
		}},
		{"a freeze", func(t *testing.T, f *File) {
			if err := f.Sheet("Sheet1").SetFreezePanes(1, 0); err != nil {
				t.Fatal(err)
			}
		}},
		{"a defined name", func(t *testing.T, f *File) {
			if err := f.SetDefinedName(DefinedName{Name: "N", Value: "Sheet1!$A$1", LocalSheetID: -1}); err != nil {
				t.Fatal(err)
			}
		}},
		{"a table", func(t *testing.T, f *File) {
			mustSetCells(t, f.Sheet("Other"), "J1", "h")
			if err := f.Sheet("Other").AddTable(Table{Name: "TableNine", Ref: "J1:K2"}); err != nil {
				t.Fatal(err)
			}
		}},
		{"a new sheet", func(t *testing.T, f *File) {
			if _, err := f.NewSheet("Third"); err != nil {
				t.Fatal(err)
			}
		}},
		{"a renamed sheet", func(t *testing.T, f *File) {
			if err := f.SetSheetName("Other", "Renamed"); err != nil {
				t.Fatal(err)
			}
		}},
		{"a hidden sheet", func(t *testing.T, f *File) {
			if err := f.SetSheetVisible("Other", false); err != nil {
				t.Fatal(err)
			}
		}},
		{"a copied sheet", func(t *testing.T, f *File) {
			if _, err := f.CopySheet("Other", "Copy"); err != nil {
				t.Fatal(err)
			}
		}},
	}
	for _, tt := range changes {
		t.Run(tt.name, func(t *testing.T) {
			f := newStructureWorkbook(t)
			u, err := f.Sheet("Sheet1").DeleteRows(3, 2)
			if err != nil {
				t.Fatal(err)
			}
			tt.change(t, f)
			changed := dumpWorkbook(t, f)
			err = u.Undo()
			if err == nil || !strings.Contains(err.Error(), "changed after the edit") {
				t.Fatalf("Undo after %s: error = %v, want a refusal", tt.name, err)
			}
			if got := dumpWorkbook(t, f); got != changed {
				t.Errorf("the refused Undo changed the workbook")
			}
			// The refusal does not use up the record, and does not make it
			// valid: it still fails.
			if err := u.Undo(); err == nil || !strings.Contains(err.Error(), "changed after the edit") {
				t.Errorf("second Undo: error = %v", err)
			}
		})
	}

	t.Run("reads, a recalculation and a save are not changes", func(t *testing.T) {
		f := newStructureWorkbook(t)
		s := f.Sheet("Sheet1")
		before := dumpWorkbook(t, f)
		mark := f.ChangeMark()
		u, err := s.DeleteRows(3, 2)
		if err != nil {
			t.Fatal(err)
		}
		if f.ChangeMark() == mark {
			t.Errorf("the edit gave no new change mark")
		}
		f.Recalculate()
		_, _ = s.GetValue("D1")
		_, _ = s.GetFormula("D1")
		_ = s.MaxRow()
		_ = f.DefinedNames()
		_ = s.SnapshotCell(1, 1)
		_, _ = s.FreezePanes()
		var buf bytes.Buffer
		if _, err := f.WriteTo(&buf); err != nil {
			t.Fatal(err)
		}
		if err := u.Undo(); err != nil {
			t.Fatalf("Undo after reads: %v", err)
		}
		if got := dumpWorkbook(t, f); got != before {
			t.Errorf("state after Undo is not the start state")
		}
		// Undo gives the workbook the mark from before the edit.
		if got := f.ChangeMark(); got != mark {
			t.Errorf("change mark after Undo = %d, want %d", got, mark)
		}
	})

	t.Run("a later change that the caller reversed, with the mark restored", func(t *testing.T) {
		// This is what a history does: the later step is undone first.
		f := newStructureWorkbook(t)
		s := f.Sheet("Sheet1")
		before := dumpWorkbook(t, f)
		u, err := s.DeleteRows(3, 2)
		if err != nil {
			t.Fatal(err)
		}
		mark := f.ChangeMark()
		old := s.SnapshotCell(1, 1)
		mustSetCells(t, s, "A1", 99)
		if err := u.Undo(); err == nil {
			t.Fatalf("Undo with a later change gave no error")
		}
		if err := s.RestoreCell(1, 1, old); err != nil {
			t.Fatal(err)
		}
		// The cell is back, but the library cannot know that.
		if err := u.Undo(); err == nil {
			t.Fatalf("Undo before the mark was restored gave no error")
		}
		f.RestoreChangeMark(mark)
		if err := u.Undo(); err != nil {
			t.Fatalf("Undo after the mark was restored: %v", err)
		}
		if got := dumpWorkbook(t, f); got != before {
			t.Errorf("state after Undo is not the start state")
		}
	})

	t.Run("two edits are undone in reverse order only", func(t *testing.T) {
		f := newStructureWorkbook(t)
		s := f.Sheet("Sheet1")
		before := dumpWorkbook(t, f)
		first, err := s.DeleteRows(5, 1)
		if err != nil {
			t.Fatal(err)
		}
		second, err := s.InsertRows(1, 1)
		if err != nil {
			t.Fatal(err)
		}
		if err := first.Undo(); err == nil {
			t.Fatalf("Undo of the first edit before the second gave no error")
		}
		if err := second.Undo(); err != nil {
			t.Fatalf("Undo of the second edit: %v", err)
		}
		if err := first.Undo(); err != nil {
			t.Fatalf("Undo of the first edit: %v", err)
		}
		if got := dumpWorkbook(t, f); got != before {
			t.Errorf("state after both undos is not the start state")
		}
	})
}

// A structured reference to a table column that a column delete removed
// becomes #REF!, as a cell reference to a deleted cell does.
func TestDeleteTableColumnInvalidatesStructuredReferences(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	other, err := f.NewSheet("Other")
	if err != nil {
		t.Fatal(err)
	}
	mustSetCells(t, s,
		"B1", "h1", "C1", "h2", "D1", "Total",
		"B2", 10, "C2", 1, "D2", "=Tbl[[#This Row],[h1]]+Tbl[[#This Row],[h2]]",
		"B3", 30, "C3", 2, "D3", "=[@h1]*2",
		"B4", 50, "C4", 3, "D4", "=[@h2]+1",
		"F1", "=SUM(Tbl[h1])", "F2", "=SUM(Tbl[h2])", "F3", "=SUM(tbl[H1])+SUM(Tbl[h2])",
		"F4", `=SUM(Tbl[h2])&"Tbl[h1]"`, "F5", "=SUM(Tbl[[h1]:[h2]])", "F6", "=ROWS(Tbl[#All])",
	)
	if err := s.AddTable(Table{Name: "Tbl", Ref: "B1:D4"}); err != nil {
		t.Fatal(err)
	}
	mustSetCells(t, other, "A1", "=SUM(Tbl[h1])", "A2", "=SUM(Tbl[h2])", "A3", "=Other1[h1]")
	if v, _ := s.GetValue("F1"); v.Number != 90 {
		t.Fatalf("test setup: SUM(Tbl[h1]) = %s, want 90", valueText(v))
	}

	// Delete sheet column B, the table column h1.
	checkStructuralEdit(t, f, structOp{"Sheet1", "delCols", 2, 1},
		map[string]string{
			"Sheet1!E1": "SUM(#REF!)",
			"Sheet1!E2": "SUM(Tbl[h2])",
			"Sheet1!E3": "SUM(#REF!)+SUM(Tbl[h2])",
			// Text in a string is not a reference.
			"Sheet1!E4": `SUM(Tbl[h2])&"Tbl[h1]"`,
			"Sheet1!E5": "SUM(#REF!)",
			"Sheet1!E6": "ROWS(Tbl[#All])",
			// In the table: with and with no table name.
			"Sheet1!C2": "#REF!+Tbl[[#This Row],[h2]]",
			"Sheet1!C3": "#REF!*2",
			"Sheet1!C4": "[@h2]+1",
			// On a different sheet.
			"Other!A1": "SUM(#REF!)",
			"Other!A2": "SUM(Tbl[h2])",
			// A different table name is not changed.
			"Other!A3": "Other1[h1]",
		},
		// (The engine does not evaluate the form with no table name, [@h2];
		// the test checks its text only.)
		map[string]any{
			"Sheet1!E1": "#REF!", "Sheet1!E2": 6, "Sheet1!C2": "#REF!",
			"Other!A1": "#REF!", "Other!A2": 6,
		})

	// A delete of columns that are not in the table changes no structured
	// reference, and a row delete changes none.
	g := New()
	gs := g.Sheet("Sheet1")
	mustSetCells(t, gs, "B1", "h1", "C1", "h2", "B2", 1, "C2", 2, "B3", 3, "C3", 4, "E1", "=SUM(Tbl[h1])", "A5", "x")
	if err := gs.AddTable(Table{Name: "Tbl", Ref: "B1:C3"}); err != nil {
		t.Fatal(err)
	}
	checkStructuralEdit(t, g, structOp{"Sheet1", "delCols", 1, 1}, map[string]string{"Sheet1!D1": "SUM(Tbl[h1])"}, map[string]any{"Sheet1!D1": 4})
	checkStructuralEdit(t, g, structOp{"Sheet1", "delRows", 3, 1}, map[string]string{"Sheet1!D1": "SUM(Tbl[h1])"}, map[string]any{"Sheet1!D1": 1})
}

// A reference with a sheet name in a different letter case follows a
// structural edit. (The engine does not yet resolve such a name when it
// evaluates; the rewrite of the text is correct.)
func TestStructuralEditRewritesSheetNamesWithoutLetterCase(t *testing.T) {
	f := New()
	data := f.Sheet("Sheet1")
	if err := f.SetSheetName("Sheet1", "Data"); err != nil {
		t.Fatal(err)
	}
	other, err := f.NewSheet("Other Sheet")
	if err != nil {
		t.Fatal(err)
	}
	mustSetCells(t, data, "A1", 1, "A2", 2, "A3", 3, "A4", 4, "A5", 5)
	mustSetCells(t, other,
		"B1", "=data!a4", "B2", "=DATA!$A$4", "B3", "=SUM(dAtA!A1:A5)", "B4", "='data'!A3", "B5", "=Data!A5+data!a5",
		"B6", "=a4", "B7", "='other sheet'!a4")

	if _, err := data.DeleteRows(3, 1); err != nil {
		t.Fatal(err)
	}
	checkCells(t, f, map[string]string{
		"Other Sheet!B1": "data!A3",
		"Other Sheet!B2": "DATA!$A$3",
		"Other Sheet!B3": "SUM(dAtA!A1:A4)",
		"Other Sheet!B4": "#REF!",
		"Other Sheet!B5": "Data!A4+data!A4",
		// References to the sheet of the formula do not change.
		"Other Sheet!B6": "a4",
		"Other Sheet!B7": "'other sheet'!a4",
	}, nil)

	// An edit on the other sheet, named in lower case in its own formulas.
	if _, err := other.InsertRows(1, 2); err != nil {
		t.Fatal(err)
	}
	checkCells(t, f, map[string]string{
		"Other Sheet!B8": "A6",
		"Other Sheet!B9": "'other sheet'!A6",
		"Other Sheet!B3": "data!A3",
	}, nil)
}

// A sequence of edits and their undos in reverse order gives the start state.
func TestStructuralUndoSequence(t *testing.T) {
	f := newStructureWorkbook(t)
	s := f.Sheet("Sheet1")
	if err := s.MergeCell("A8", "B9"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetDefinedName(DefinedName{Name: "Data", Value: "Sheet1!$A$1:$B$6", LocalSheetID: -1}); err != nil {
		t.Fatal(err)
	}
	mustSetCells(t, s, "G1", "=SEQUENCE(2,2)")
	start := dumpWorkbook(t, f)

	ops := []structOp{
		{"Sheet1", "insRows", 2, 3},
		{"Sheet1", "delCols", 1, 1},
		{"Other", "insRows", 1, 1},
		{"Sheet1", "delRows", 4, 4},
		{"Sheet1", "insCols", 2, 2},
		{"Other", "delCols", 1, 1},
		{"Sheet1", "delRows", 1, 20},
	}
	var undos []*StructuralUndo
	var states []string
	for _, op := range ops {
		states = append(states, dumpWorkbook(t, f))
		u, err := op.run(f)
		if err != nil {
			t.Fatalf("%+v: %v", op, err)
		}
		undos = append(undos, u)
	}
	for i := len(undos) - 1; i >= 0; i-- {
		if err := undos[i].Undo(); err != nil {
			t.Fatalf("Undo %d: %v", i, err)
		}
		if got := dumpWorkbook(t, f); got != states[i] {
			t.Fatalf("state after Undo of step %d (%+v):\n%s\nwant:\n%s", i, ops[i], got, states[i])
		}
	}
	if got := dumpWorkbook(t, f); got != start {
		t.Errorf("the workbook is not at its start state")
	}
}
