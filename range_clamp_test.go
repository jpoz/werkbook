package werkbook

import (
	"fmt"
	"sort"
	"testing"

	"github.com/jpoz/werkbook/formula"
)

// rangeClampTestSheet builds a sparse sheet: data in rows 1..5 and 10, a row
// that has only a height (20), a row with a style-only cell (30), and a
// column far at the right.
func rangeClampTestSheet(t *testing.T) (*File, *Sheet) {
	t.Helper()
	f := New()
	s := f.Sheet("Sheet1")
	steps := []error{
		s.SetValue("A1", 1), s.SetValue("B1", 2), s.SetValue("C1", 3),
		s.SetValue("A2", 4), s.SetValue("C2", 5),
		s.SetValue("B3", 6),
		s.SetValue("A5", 7), s.SetValue("Z5", 8),
		s.SetValue("B10", 9),
		s.SetRowHeight(20, 30),
		s.SetStyle("D30", &Style{Font: &Font{Bold: true}}),
	}
	for i, err := range steps {
		if err != nil {
			t.Fatalf("setup step %d: %v", i, err)
		}
	}
	return f, s
}

// rangeClampBounds can return a stand-in for MaxRow and no MaxCol. For each
// request, the clamp must be the same as the clamp from the exact bounds.
func TestRangeClampBoundsGiveTheExactClamp(t *testing.T) {
	_, s := rangeClampTestSheet(t)
	empty := New().Sheet("Sheet1")

	rowEnds := []int{1, 2, 3, 4, 5, 6, 9, 10, 11, 19, 20, 21, 29, 30, 31, 100, MaxRows}
	rowStarts := []int{1, 2, 4, 5, 6, 10, 11, 20, 30, 31, 500}
	colRanges := [][2]int{{1, 1}, {1, 3}, {2, 2}, {4, 26}, {26, 27}, {1, 26}, {1, 4000}, {1, MaxColumns}, {2, MaxColumns}, {MaxColumns, MaxColumns}}

	for name, sheet := range map[string]*Sheet{"sparse": s, "empty": empty} {
		exactRow, exactCol := sheet.MaxRow(), sheet.MaxCol()
		for _, fromRow := range rowStarts {
			for _, toRow := range rowEnds {
				for _, cols := range colRanges {
					req := rangeMaterializationRequest{
						sheet: "Sheet1", fromCol: cols[0], fromRow: fromRow, toCol: cols[1], toRow: toRow,
					}
					wantRow, wantCol := clampMaterializedRange(req, exactRow, exactCol)
					gotMaxRow, gotMaxCol := sheet.rangeClampBounds(req)
					gotRow, gotCol := clampMaterializedRange(req, gotMaxRow, gotMaxCol)
					if gotRow != wantRow || gotCol != wantCol {
						t.Errorf("%s sheet, request %+v: clamp = (%d, %d), want (%d, %d)", name, req, gotRow, gotCol, wantRow, wantCol)
					}
				}
			}
		}
	}
}

// ForEachCell reads a small range by row and column number and a large range
// by a scan of the maps. Both paths must visit exactly the stored cells of
// the range, one time each.
func TestSheetRangeGridForEachCellVisitsTheStoredCells(t *testing.T) {
	_, s := rangeClampTestSheet(t)
	grid := sheetRangeGrid{sheet: s, maxRow: s.MaxRow(), maxCol: s.MaxCol()}

	stored := func(fromCol, fromRow, toCol, toRow int) []string {
		var out []string
		for rowNum, r := range s.rows {
			for colNum := range r.cells {
				if rowNum >= fromRow && rowNum <= toRow && colNum >= fromCol && colNum <= toCol {
					out = append(out, fmt.Sprintf("%d,%d", rowNum, colNum))
				}
			}
		}
		sort.Strings(out)
		return out
	}

	ranges := [][4]int{ // fromCol, fromRow, toCol, toRow
		{1, 1, 1, 1},                // one cell
		{2, 2, 2, 2},                // one cell that is absent
		{1, 1, 3, 3},                // small block: by number on both axes
		{1, 1, 3, 2},                // fewer rows than the sheet, as many columns as row 1 has
		{1, 1, 2, 1},                // fewer columns than row 1 has
		{1, 1, 26, 5},               // wider than each row: map scan of the cells
		{1, 1, MaxColumns, 5},       // full rows
		{1, 1, 1, MaxRows},          // full column: map scan of the rows
		{2, 1, 2, 30},               // more rows than the sheet has
		{1, 6, 26, 9},               // only absent rows
		{1, 20, 26, 20},             // a row with a height and no cells
		{4, 30, 4, 30},              // a style-only cell
		{1, 1, MaxColumns, MaxRows}, // the full sheet
		{3, 3, 1, 1},                // an inverted range visits nothing
	}
	for _, rg := range ranges {
		var got []string
		seen := map[string]int{}
		grid.ForEachCell("Sheet1", rg[0], rg[1], rg[2], rg[3], func(col, row int, _ formula.Value, _ bool) {
			key := fmt.Sprintf("%d,%d", row, col)
			seen[key]++
			got = append(got, key)
		})
		sort.Strings(got)
		want := stored(rg[0], rg[1], rg[2], rg[3])
		if fmt.Sprint(got) != fmt.Sprint(want) {
			t.Errorf("range %v: visited %v, want %v", rg, got, want)
		}
		for key, n := range seen {
			if n != 1 {
				t.Errorf("range %v: cell %s visited %d times", rg, key, n)
			}
		}
	}

	// The callback gets the value and the spill-slot state of each cell.
	values := map[string]float64{}
	occupied := map[string]bool{}
	grid.ForEachCell("Sheet1", 1, 1, 4, 30, func(col, row int, v formula.Value, occupies bool) {
		key := fmt.Sprintf("%d,%d", row, col)
		values[key] = v.Num
		occupied[key] = occupies
	})
	if values["1,3"] != 3 || values["10,2"] != 9 || !occupied["1,3"] {
		t.Errorf("values = %v occupied = %v", values, occupied)
	}
	if occ, ok := occupied["30,4"]; !ok || occ {
		t.Errorf("the style-only cell D30: visited %v, occupies %v; want visited and not occupying", ok, occ)
	}

	// A nil callback and a nil sheet do nothing.
	grid.ForEachCell("Sheet1", 1, 1, 3, 3, nil)
	sheetRangeGrid{}.ForEachCell("Sheet1", 1, 1, 3, 3, func(int, int, formula.Value, bool) {
		t.Errorf("callback called for a nil sheet")
	})
}

// The same formulas, read through the engine: trimmed ranges, full rows and
// columns, single cells, ranges that end on absent rows, and ranges past the
// used range.
func TestRangeReadsOnASparseSheet(t *testing.T) {
	f, s := rangeClampTestSheet(t)
	other, err := f.NewSheet("Out")
	if err != nil {
		t.Fatal(err)
	}
	tests := []struct {
		formula string
		want    float64
	}{
		{"SUM(Sheet1!A1:C1)", 6},
		{"SUM(Sheet1!A1:C3)", 21},
		{"SUM(Sheet1!A1:C4)", 21}, // ends on an absent row
		{"SUM(Sheet1!A1:C5)", 28},
		{"SUM(Sheet1!A1:C9)", 28}, // ends on an absent row before a later row
		{"SUM(Sheet1!A1:C10)", 37},
		{"SUM(Sheet1!A1:C20)", 37},   // ends on a row with only a height
		{"SUM(Sheet1!A1:C1000)", 37}, // ends past the used range
		{"SUM(Sheet1!A6:C9)", 0},     // only absent rows
		{"SUM(Sheet1!A40:C50)", 0},   // fully past the used range
		{"SUM(Sheet1!B3)", 6},        // one cell
		{"SUM(Sheet1!B4)", 0},        // one absent cell
		{"SUM(Sheet1!A:A)", 12},      // full column
		{"SUM(Sheet1!A:C)", 37},
		{"SUM(Sheet1!A:Z)", 45},
		{"SUM(Sheet1!5:5)", 15}, // full row
		{"SUM(Sheet1!1:5)", 36},
		{"SUM(Sheet1!1:1048576)", 45}, // all rows
		{"SUM(Sheet1!6:9)", 0},        // full rows that are absent
		{"COUNT(Sheet1!A1:Z30)", 9},
		{"COUNTA(Sheet1!A1:Z30)", 9},
		{"COUNTBLANK(Sheet1!A1:C3)", 3},
		{"ROWS(Sheet1!A1:C1000)", 1000}, // the logical size, not the trimmed size
		{"COLUMNS(Sheet1!1:5)", 16384},
		{"ROWS(Sheet1!A:A)", 1048576},
		{"SUM(Sheet1!Z1:Z1048576)", 8},
		{"SUMPRODUCT(Sheet1!A1:A5,Sheet1!A1:A5)", 66},
		{"MAX(Sheet1!A1:XFD10)", 9},
	}
	for i, tt := range tests {
		ref := fmt.Sprintf("A%d", i+1)
		if err := other.SetFormula(ref, tt.formula); err != nil {
			t.Fatalf("SetFormula(%s): %v", tt.formula, err)
		}
	}
	check := func(stage string) {
		t.Helper()
		for i, tt := range tests {
			v, err := other.GetValue(fmt.Sprintf("A%d", i+1))
			if err != nil {
				t.Fatal(err)
			}
			if v.Type != TypeNumber || v.Number != tt.want {
				t.Errorf("%s: %s = %#v, want %v", stage, tt.formula, v, tt.want)
			}
		}
	}
	check("lazy read")
	f.Recalculate()
	check("after Recalculate")

	// A new last row changes each range that reaches it.
	if err := s.SetValue("C40", 100); err != nil {
		t.Fatal(err)
	}
	setWant := func(formula string, want float64) {
		t.Helper()
		for i := range tests {
			if tests[i].formula == formula {
				tests[i].want = want
				return
			}
		}
		t.Fatalf("no test for %s", formula)
	}
	// A1:Z30 and the other ranges that end before row 40 do not change.
	after := map[string]float64{
		"SUM(Sheet1!A1:C1000)":  137,
		"SUM(Sheet1!A40:C50)":   100,
		"SUM(Sheet1!A:C)":       137,
		"SUM(Sheet1!A:Z)":       145,
		"SUM(Sheet1!1:1048576)": 145,
	}
	before := map[string]float64{}
	for formula, want := range after {
		for _, tt := range tests {
			if tt.formula == formula {
				before[formula] = tt.want
			}
		}
		setWant(formula, want)
	}
	check("after a write below the used range")

	// The removal of the last row trims the ranges again.
	if err := s.RestoreCell(3, 40, CellSnapshot{}); err != nil {
		t.Fatal(err)
	}
	for formula, want := range before {
		setWant(formula, want)
	}
	check("after the last row was removed")
}

// Ranges in, across and past spill ranges. A spill follower is not a stored
// cell, so a spill can go past the last stored row: the range clamp must not
// cut the range before the spill is read.
func TestRangeReadsWithSpills(t *testing.T) {
	f := New()
	vertical, err := f.NewSheet("V")
	if err != nil {
		t.Fatal(err)
	}
	horizontal, err := f.NewSheet("H")
	if err != nil {
		t.Fatal(err)
	}
	out := f.Sheet("Sheet1")
	// V: A1 = 100, A2 spills 1..5 to A2:A6. The only stored rows are 1 and 2.
	mustSetCells(t, vertical, "A1", 100, "A2", "=SEQUENCE(5)", "C1", 7)
	// H: A1 spills 1..4 to A1:D1; row 3 has a value.
	mustSetCells(t, horizontal, "A1", "=SEQUENCE(1,4)", "B3", 50)
	if vertical.MaxRow() != 2 {
		t.Fatalf("test setup: V has %d stored rows, want 2", vertical.MaxRow())
	}

	tests := []struct {
		formula string
		want    float64
	}{
		// In the spill.
		{"SUM(V!A2:A6)", 15},
		{"SUM(V!A3:A5)", 9},
		{"SUM(V!A4)", 3},
		// Across the start and the end of the spill.
		{"SUM(V!A1:A3)", 103},
		{"SUM(V!A5:A9)", 9},
		{"SUM(V!A1:A100)", 115},
		// Past the spill.
		{"SUM(V!A7:A20)", 0},
		// Full column and full row.
		{"SUM(V!A:A)", 115},
		{"SUM(V!A:C)", 122},
		{"SUM(V!4:4)", 3},
		{"COUNT(V!A:A)", 6},
		{"ROWS(V!A2:A6)", 5},
		{"MAX(V!A1:C6)", 100},
		{"INDEX(V!A1:A6,6)", 5},
		// Horizontal spill.
		{"SUM(H!A1:D1)", 10},
		{"SUM(H!B1:C1)", 5},
		{"SUM(H!C1:Z1)", 7},
		{"SUM(H!1:1)", 10},
		{"SUM(H!1:3)", 60},
		{"SUM(H!D:D)", 4},
		{"SUM(H!E1:H1)", 0},
		{"COUNT(H!A1:XFD1)", 4},
	}
	for i, tt := range tests {
		mustSetCells(t, out, fmt.Sprintf("A%d", i+1), "="+tt.formula)
	}
	check := func(stage string) {
		t.Helper()
		for i, tt := range tests {
			v, err := out.GetValue(fmt.Sprintf("A%d", i+1))
			if err != nil {
				t.Fatal(err)
			}
			if v.Type != TypeNumber || v.Number != tt.want {
				t.Errorf("%s: %s = %s, want %v", stage, tt.formula, valueText(v), tt.want)
			}
		}
	}
	check("lazy read")
	f.Recalculate()
	check("after Recalculate")

	// The clamp for a range that ends in the spill, past the stored rows:
	// the stand-in bound and the exact bound give the same clamp.
	for _, toRow := range []int{2, 3, 6, 7, 100} {
		req := rangeMaterializationRequest{sheet: "V", fromCol: 1, fromRow: 1, toCol: 1, toRow: toRow}
		wantRow, wantCol := clampMaterializedRange(req, vertical.MaxRow(), vertical.MaxCol())
		maxRow, maxCol := vertical.rangeClampBounds(req)
		gotRow, gotCol := clampMaterializedRange(req, maxRow, maxCol)
		if gotRow != wantRow || gotCol != wantCol {
			t.Errorf("clamp for A1:A%d = (%d, %d), want (%d, %d)", toRow, gotRow, gotCol, wantRow, wantCol)
		}
	}
}

// A trimmed full-column array result keeps its size: the stand-in for MaxRow
// must not change how far a full-column reference spills.
func TestFullColumnSpillSizeFollowsTheUsedRange(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	out, err := f.NewSheet("Out")
	if err != nil {
		t.Fatal(err)
	}
	for r := 1; r <= 4; r++ {
		if err := s.SetValue(fmt.Sprintf("A%d", r), r*10); err != nil {
			t.Fatal(err)
		}
	}
	if err := out.SetFormula("A1", "Sheet1!A1:A4*2"); err != nil {
		t.Fatal(err)
	}
	if err := out.SetFormula("C1", "SUM(Sheet1!A1:A3)+SUM(Sheet1!A2:A1000)"); err != nil {
		t.Fatal(err)
	}
	f.Recalculate()
	if toCol, toRow, ok := out.SpillBounds(1, 1); !ok || toCol != 1 || toRow != 4 {
		t.Errorf("spill bounds = %d,%d,%v, want 1,4,true", toCol, toRow, ok)
	}
	if v, _ := out.GetValue("A4"); v.Number != 80 {
		t.Errorf("Out!A4 = %#v, want 80", v)
	}
	if v, _ := out.GetValue("C1"); v.Number != 150 {
		t.Errorf("Out!C1 = %#v, want 150", v)
	}
}

// The cost of a recalculation must grow with the number of formulas, not
// with its square. The test compares counts of work, not time: each SUM over
// one row must read only the cells of its row.
func TestRecalculateRowSumsReadsOnlyTheRangeCells(t *testing.T) {
	const rows = 300
	f := newRowSumsWorkbook(t, rows)
	s := f.Sheet("Sheet1")
	f.Recalculate()
	for _, r := range []int{1, 150, rows} {
		v, err := s.GetValue(fmt.Sprintf("J%d", r))
		if err != nil {
			t.Fatal(err)
		}
		want := float64(9*r + 45)
		if v.Number != want {
			t.Errorf("J%d = %v, want %v", r, v.Number, want)
		}
	}
	// A row range in the data is read by number: no scan of the sheet.
	req := rangeMaterializationRequest{sheet: "Sheet1", fromCol: 1, fromRow: 150, toCol: 9, toRow: 150}
	if maxRow, maxCol := s.rangeClampBounds(req); maxRow != 150 || maxCol != 0 {
		t.Errorf("rangeClampBounds = (%d, %d), want the stand-in (150, 0)", maxRow, maxCol)
	}
	visited := 0
	sheetRangeGrid{sheet: s}.ForEachCell("Sheet1", 1, 150, 9, 150, func(int, int, formula.Value, bool) { visited++ })
	if visited != 9 {
		t.Errorf("ForEachCell visited %d cells, want 9", visited)
	}
}
