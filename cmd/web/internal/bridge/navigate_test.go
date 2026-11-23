package bridge

import (
	"testing"

	werkbook "github.com/jpoz/werkbook"
)

// edge runs FindEdge from an A1 reference and returns the result as an A1
// reference.
func edge(t *testing.T, s *Session, from string, dRow, dCol int) string {
	t.Helper()
	col, row, err := werkbook.CellNameToCoordinates(from)
	if err != nil {
		t.Fatal(err)
	}
	res, err := s.FindEdge(sheet1, row, col, dRow, dCol)
	if err != nil {
		t.Fatalf("FindEdge(%s, %d, %d): %v", from, dRow, dCol, err)
	}
	ref, err := werkbook.CoordinatesToCellName(res.Col, res.Row)
	if err != nil {
		t.Fatalf("FindEdge(%s) gave %+v: %v", from, res, err)
	}
	return ref
}

const (
	down, up, right, left = 0, 1, 2, 3
)

var directions = [4][2]int{{1, 0}, {-1, 0}, {0, 1}, {0, -1}}

func TestFindEdgeRules(t *testing.T) {
	// A block C3:E5, a gap, single cells, on both axes:
	//
	//	   A  B  C  D  E  F  G  H
	//	1        .  x  .
	//	2
	//	3  x     x  x  x     x
	//	4        x  x  x
	//	5        x  x  x
	//	6
	//	7           x
	//	8           x
	s := NewSession()
	mustSet(t, s, sheet1,
		"D1", "x",
		"A3", "x", "C3", "x", "D3", "x", "E3", "x", "G3", "x",
		"C4", "x", "D4", "x", "E4", "x",
		"C5", "x", "D5", "x", "E5", "x",
		"D7", "x", "D8", "x",
	)
	lastRow := "D1048576"
	tests := []struct {
		name string
		from string
		dir  int
		want string
	}{
		// Rule 1: in a run, go to its last cell.
		{"down in a run", "D3", down, "D5"},
		{"down from the middle of a run", "D4", down, "D5"},
		{"up in a run", "D5", up, "D3"},
		{"right in a run", "C3", right, "E3"},
		{"left in a run", "E3", left, "C3"},
		{"down in a run of two", "D7", down, "D8"},

		// Rule 2: from the end of a run or from an empty cell, go to the
		// next non-empty cell.
		{"down from the end of a run", "D5", down, "D7"},
		{"up from the start of a run", "D3", up, "D1"},
		{"right from the end of a run", "E3", right, "G3"},
		{"left from the start of a run", "C3", left, "A3"},
		{"down from an empty cell", "D2", down, "D3"},
		{"down from an empty cell between runs", "D6", down, "D7"},
		{"up from an empty cell", "D6", up, "D5"},
		{"right from an empty cell", "B3", right, "C3"},
		{"left from an empty cell", "F3", left, "E3"},
		{"down from a single cell to the next", "D1", down, "D3"},
		{"left from a single cell to the next", "G3", left, "E3"},

		// Rule 3: with no cell in the direction, go to the edge of the sheet.
		{"down from the last cell", "D8", down, lastRow},
		{"up from the first cell", "D1", up, "D1"},
		{"right from the last cell", "G3", right, "XFD3"},
		{"left from the first cell", "A3", left, "A3"},
		{"down in an empty column", "H2", down, "H1048576"},
		{"up in an empty column", "H9", up, "H1"},
		{"right in an empty row", "B2", right, "XFD2"},
		{"left in an empty row", "F2", left, "A2"},
		{"right in an empty row after the used range", "A20", right, "XFD20"},

		// Start out of the used range.
		{"up from far below goes to the last cell of the column", "D5000", up, "D8"},
		{"left from far right goes to the last cell of the row", "ZZ3", left, "G3"},
		{"up from the last row", lastRow, up, "D8"},
		{"left from the last column", "XFD3", left, "G3"},
		{"up from far below in an empty column", "H5000", up, "H1"},
		{"down from far below", "D5000", down, lastRow},

		// Sheet edges as the start cell.
		{"down at the last row stays", lastRow, down, lastRow},
		{"right at the last column stays", "XFD3", right, "XFD3"},
		{"up at row 1 stays", "C1", up, "C1"},
		{"left at column A stays", "A7", left, "A7"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			d := directions[tt.dir]
			if got := edge(t, s, tt.from, d[0], d[1]); got != tt.want {
				t.Errorf("FindEdge(%s, %v) = %s, want %s", tt.from, d, got, tt.want)
			}
		})
	}

	// On an empty sheet each direction goes to the sheet edge.
	empty := NewSession()
	for dir, want := range map[int]string{down: "C1048576", up: "C1", right: "XFD5", left: "A5"} {
		d := directions[dir]
		if got := edge(t, empty, "C5", d[0], d[1]); got != want {
			t.Errorf("empty sheet, direction %v: %s, want %s", d, got, want)
		}
	}
}

func TestFindEdgeWhatIsEmpty(t *testing.T) {
	s := NewSession()
	sh := s.file.Sheet(sheet1)
	// Column A: a value, a formula that gives "", an error, a bool: one run.
	mustSet(t, s, sheet1, "A1", "x", "A2", `=""`, "A3", "=1/0", "A4", "FALSE", "A5", "0")
	if got := edge(t, s, "A1", 1, 0); got != "A5" {
		t.Errorf("run with an empty-text formula, an error, FALSE and 0: %s, want A5", got)
	}

	// Column C: a cell with only a style is empty.
	if err := sh.SetStyle("C2", &werkbook.Style{Fill: &werkbook.Fill{Color: "FF0000"}}); err != nil {
		t.Fatal(err)
	}
	mustSet(t, s, sheet1, "C1", "x", "C3", "x")
	if got := edge(t, s, "C1", 1, 0); got != "C3" {
		t.Errorf("down over a style-only cell: %s, want C3 (the styled cell is empty)", got)
	}
	if got := edge(t, s, "C2", 1, 0); got != "C3" {
		t.Errorf("down from a style-only cell: %s, want C3", got)
	}

	// Column E: a spill counts as data, anchor and followers.
	mustSet(t, s, sheet1, "E2", "=SEQUENCE(4)")
	if got := edge(t, s, "E2", 1, 0); got != "E5" {
		t.Errorf("down in a spill: %s, want E5", got)
	}
	if got := edge(t, s, "E1", 1, 0); got != "E2" {
		t.Errorf("down to a spill: %s, want E2", got)
	}
	if got := edge(t, s, "E9", -1, 0); got != "E5" {
		t.Errorf("up to the last follower: %s, want E5", got)
	}
	if got := edge(t, s, "E4", -1, 0); got != "E2" {
		t.Errorf("up in a spill: %s, want E2", got)
	}
	// From far below, with the follower as the last used cell of the sheet.
	if got := edge(t, s, "E90000", -1, 0); got != "E5" {
		t.Errorf("up from far below to the last follower: %s, want E5", got)
	}
	// A horizontal spill.
	mustSet(t, s, sheet1, "G8", "=SEQUENCE(1,3)")
	if got := edge(t, s, "G8", 0, 1); got != "I8" {
		t.Errorf("right in a spill: %s, want I8", got)
	}
	if got := edge(t, s, "Z8", 0, -1); got != "I8" {
		t.Errorf("left to the last follower: %s, want I8", got)
	}

	// A cleared cell is empty again.
	mustSet(t, s, sheet1, "A3", "")
	if got := edge(t, s, "A1", 1, 0); got != "A2" {
		t.Errorf("after a clear in the run: %s, want A2", got)
	}
}

func TestFindEdgeMergedRanges(t *testing.T) {
	s := NewSession()
	sh := s.file.Sheet(sheet1)
	// B2:C3 is a merged range with a value; B6:C7 is an empty merged range.
	for _, m := range [][2]string{{"B2", "C3"}, {"B6", "C7"}} {
		if err := sh.MergeCell(m[0], m[1]); err != nil {
			t.Fatal(err)
		}
	}
	mustSet(t, s, sheet1, "B1", "x", "B2", "merged", "B4", "x", "B9", "x", "A2", "x", "D2", "x", "D3", "y", "A3", "y")

	tests := []struct {
		name       string
		from       string
		dRow, dCol int
		want       string
	}{
		// A merged range is one cell: B1, B2:C3, B4 are a run of three.
		{"down through a merged range", "B1", 1, 0, "B4"},
		{"up through a merged range", "B4", -1, 0, "B1"},
		// A result in a merged range is its top-left cell.
		{"down into a merged range from an empty start", "C1", 1, 0, "B2"},
		// The scan starts from the edge of the merged range of the start cell.
		{"down from the top-left cell of a merged range", "B2", 1, 0, "B4"},
		{"down from a covered cell in column C", "C3", 1, 0, "C1048576"},
		{"right from a merged range", "B2", 0, 1, "D2"},
		{"right from the second row of a merged range", "B3", 0, 1, "D3"},
		{"left from a merged range", "C3", 0, -1, "A3"},
		{"right through a merged range", "A2", 0, 1, "D2"},
		{"left through a merged range on its second row", "D3", 0, -1, "A3"},
		// An empty merged range is an empty cell.
		{"down over an empty merged range", "B4", 1, 0, "B9"},
		{"down from an empty merged range", "B6", 1, 0, "B9"},
		// Up column C from the empty range: the next data is the merged
		// range B2:C3, which covers C3.
		{"up from an empty merged range", "C7", -1, 0, "B2"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := edge(t, s, tt.from, tt.dRow, tt.dCol); got != tt.want {
				t.Errorf("FindEdge(%s, %d, %d) = %s, want %s", tt.from, tt.dRow, tt.dCol, got, tt.want)
			}
		})
	}
}

// From the last row or column of the sheet, with a small used range: the
// result is the last data cell of the lane, and the scan does not visit a
// million empty cells one by one.
func TestFindEdgeFromTheSheetEndWithASmallUsedRange(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "B2", "x", "B3", "y", "D2", "z")
	// A row that has only a height below the data is not data.
	if _, err := s.SetRowHeight(sheet1, 50, 50, 40); err != nil {
		t.Fatal(err)
	}
	tests := []struct {
		from       string
		dRow, dCol int
		want       string
	}{
		{"B1048576", -1, 0, "B3"},
		{"B1048575", -1, 0, "B3"},
		{"D1048576", -1, 0, "D2"},
		{"A1048576", -1, 0, "A1"}, // an empty column: the sheet edge
		{"XFD2", 0, -1, "D2"},
		{"XFD3", 0, -1, "B3"},
		{"XFD1", 0, -1, "A1"},
		{"XFD1048576", -1, 0, "XFD1"},
		{"XFD1048576", 0, -1, "A1048576"},
		{"B4", 1, 0, "B1048576"},
		{"B51", -1, 0, "B3"}, // over the row with only a height
	}
	for _, tt := range tests {
		if got := edge(t, s, tt.from, tt.dRow, tt.dCol); got != tt.want {
			t.Errorf("FindEdge(%s, %d, %d) = %s, want %s", tt.from, tt.dRow, tt.dCol, got, tt.want)
		}
	}
}

func TestFindEdgeErrors(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "x")
	h := s.History()
	bad := []struct {
		sheet                string
		row, col, dRow, dCol int
	}{
		{"Nope", 1, 1, 1, 0},
		{sheet1, 0, 1, 1, 0},
		{sheet1, 1, 0, 1, 0},
		{sheet1, werkbook.MaxRows + 1, 1, -1, 0},
		{sheet1, 1, werkbook.MaxColumns + 1, 0, -1},
		{sheet1, 1, 1, 0, 0},
		{sheet1, 1, 1, 1, 1},
		{sheet1, 1, 1, -1, 1},
		{sheet1, 1, 1, 2, 0},
		{sheet1, 1, 1, 0, -2},
	}
	for _, b := range bad {
		if res, err := s.FindEdge(b.sheet, b.row, b.col, b.dRow, b.dCol); err == nil {
			t.Errorf("FindEdge(%+v) = %+v, want an error", b, res)
		}
	}
	// A read does not change the history.
	if _, err := s.FindEdge(sheet1, 1, 1, 1, 0); err != nil {
		t.Fatal(err)
	}
	if s.History() != h {
		t.Errorf("FindEdge changed the history")
	}
}

func region(t *testing.T, s *Session, cell string) (string, bool) {
	t.Helper()
	col, row, err := werkbook.CellNameToCoordinates(cell)
	if err != nil {
		t.Fatal(err)
	}
	res, err := s.CurrentRegion(sheet1, row, col)
	if err != nil {
		t.Fatalf("CurrentRegion(%s): %v", cell, err)
	}
	a, _ := werkbook.CoordinatesToCellName(res.Range.Col1, res.Range.Row1)
	b, _ := werkbook.CoordinatesToCellName(res.Range.Col2, res.Range.Row2)
	return a + ":" + b, res.Header
}

func TestCurrentRegionShapes(t *testing.T) {
	// Two regions and a single cell:
	//
	//	   A  B  C  D  E  F  G
	//	1
	//	2     x  x        x
	//	3     x     x
	//	4        x  x
	//	5
	//	6  x
	//	7     x              (diagonal neighbor of A6)
	s := NewSession()
	mustSet(t, s, sheet1,
		"B2", "x", "C2", "x", "B3", "x", "D3", "x", "C4", "x", "D4", "x",
		"F2", "x",
		"A6", "x", "B7", "x",
	)
	tests := []struct {
		name string
		cell string
		want string
	}{
		{"from a corner cell", "B2", "B2:D4"},
		{"from the last cell", "D4", "B2:D4"},
		{"from an empty cell in the region", "C3", "B2:D4"},
		{"from an empty corner of the region", "D2", "B2:D4"},
		{"from an empty cell next to the region", "A3", "A2:D4"},
		{"from an empty cell that touches a corner", "A1", "A1:D4"},
		{"from an empty cell between two regions joins them", "E3", "B2:F4"},
		{"single cell", "F2", "F2:F2"},
		{"cells that touch at a corner are one region", "A6", "A6:B7"},
		{"from the second cell of a diagonal pair", "B7", "A6:B7"},
		{"empty cell with no neighbor", "J20", "J20:J20"},
		{"last cell of the sheet", "XFD1048576", "XFD1048576:XFD1048576"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got, _ := region(t, s, tt.cell); got != tt.want {
				t.Errorf("CurrentRegion(%s) = %s, want %s", tt.cell, got, tt.want)
			}
		})
	}

	t.Run("region at the sheet corner", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "x", "B1", "x", "A2", "x")
		if got, _ := region(t, s, "B2"); got != "A1:B2" {
			t.Errorf("region = %s, want A1:B2", got)
		}
	})

	t.Run("spill, formulas and merged ranges are data", func(t *testing.T) {
		s := NewSession()
		if err := s.file.Sheet(sheet1).MergeCell("C1", "D2"); err != nil {
			t.Fatal(err)
		}
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(3)", "B1", `=""`, "C1", "merged", "E2", "x")
		// A1:A3 is a spill; C1:D2 is a merged range; E2 touches its second row.
		if got, _ := region(t, s, "A2"); got != "A1:E3" {
			t.Errorf("region = %s, want A1:E3", got)
		}
	})

	t.Run("style-only cells do not make a region", func(t *testing.T) {
		s := NewSession()
		if err := s.file.Sheet(sheet1).SetRangeStyle("A1:C3", &werkbook.Style{Font: &werkbook.Font{Bold: true}}); err != nil {
			t.Fatal(err)
		}
		mustSet(t, s, sheet1, "B2", "x")
		if got, _ := region(t, s, "B2"); got != "B2:B2" {
			t.Errorf("region = %s, want B2:B2", got)
		}
	})

	t.Run("large region", func(t *testing.T) {
		s := NewSession()
		edits := make([]Edit, 0, 3000)
		for r := 1; r <= 1000; r++ {
			for c := 1; c <= 3; c++ {
				edits = append(edits, Edit{Row: r, Col: c, Input: "1"})
			}
		}
		if err := s.SetCells(sheet1, edits); err != nil {
			t.Fatal(err)
		}
		if got, _ := region(t, s, "B500"); got != "A1:C1000" {
			t.Errorf("region = %s, want A1:C1000", got)
		}
	})

	t.Run("errors", func(t *testing.T) {
		s := NewSession()
		if _, err := s.CurrentRegion("Nope", 1, 1); err == nil {
			t.Error("unknown sheet gave no error")
		}
		if _, err := s.CurrentRegion(sheet1, 0, 1); err == nil {
			t.Error("row 0 gave no error")
		}
		if _, err := s.CurrentRegion(sheet1, 1, werkbook.MaxColumns+1); err == nil {
			t.Error("a column past the end gave no error")
		}
	})
}

func TestCurrentRegionHeader(t *testing.T) {
	tests := []struct {
		name string
		set  []string
		want bool
	}{
		{"text over numbers", []string{"A1", "Name", "B1", "Qty", "A2", "a", "B2", "5"}, true},
		{"text over a formula that gives a number", []string{"A1", "Total", "A2", "=1+1"}, true},
		{"text over a bool", []string{"A1", "Done", "A2", "TRUE"}, true},
		{"text over an error", []string{"A1", "Rate", "A2", "=1/0"}, true},
		{"text over text only", []string{"A1", "Name", "B1", "City", "A2", "a", "B2", "b"}, false},
		{"numbers in the first row", []string{"A1", "1", "B1", "2", "A2", "3", "B2", "4"}, false},
		{"one number in the first row", []string{"A1", "Name", "B1", "2024", "A2", "a", "B2", "5"}, false},
		{"empty cell in the first row is permitted", []string{"A1", "Name", "C1", "Qty", "A2", "a", "B2", "x", "C2", "5"}, true},
		{"first row with no text", []string{"B1", "1", "A2", "a", "B2", "5"}, false},
		{"one row", []string{"A1", "Name", "B1", "Qty"}, false},
		{"one cell", []string{"A1", "Name"}, false},
		{"formula that gives text is text", []string{"A1", `="Na"&"me"`, "A2", "5"}, true},
		{"text-like number in the second row is text", []string{"A1", "Name", "A2", "'5"}, false},
		{"second row empty under the header, number in the third row", []string{"A1", "Name", "B2", "x", "A3", "5"}, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := NewSession()
			mustSet(t, s, sheet1, tt.set...)
			if _, got := region(t, s, "A1"); got != tt.want {
				t.Errorf("header = %v, want %v", got, tt.want)
			}
		})
	}
}
