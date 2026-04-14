package werkbook

import (
	"testing"

	"github.com/jpoz/werkbook/formula"
)

type fakeRangeGrid struct {
	maxRow int
	maxCol int
	cells  []rangeMaterializationCell
}

func (g *fakeRangeGrid) MaxRow(string) int {
	return g.maxRow
}

func (g *fakeRangeGrid) MaxCol(string) int {
	return g.maxCol
}

func (g *fakeRangeGrid) ForEachCell(
	_ string,
	fromCol, fromRow, toCol, toRow int,
	fn func(col, row int, value formula.Value, occupies bool),
) {
	for _, cell := range g.cells {
		if cell.col < fromCol || cell.col > toCol {
			continue
		}
		if cell.row < fromRow || cell.row > toRow {
			continue
		}
		fn(cell.col, cell.row, cell.value, cell.occupies)
	}
}

type fakeRangeSpills struct {
	anchors []rangeSpillAnchor
}

func (s *fakeRangeSpills) Anchors(string) []rangeSpillAnchor {
	return s.anchors
}

func gridCell(ref string, value formula.Value, occupies bool) rangeMaterializationCell {
	col, row, err := CellNameToCoordinates(ref)
	if err != nil {
		panic(err)
	}
	return rangeMaterializationCell{
		col:      col,
		row:      row,
		value:    value,
		occupies: occupies,
	}
}

func spillAnchor(ref, toRef string, raw formula.Value) rangeSpillAnchor {
	col, row, err := CellNameToCoordinates(ref)
	if err != nil {
		panic(err)
	}
	toCol, toRow, err := CellNameToCoordinates(toRef)
	if err != nil {
		panic(err)
	}
	return rangeSpillAnchor{
		col:   col,
		row:   row,
		toCol: toCol,
		toRow: toRow,
		raw:   raw,
	}
}

func assertCellValue(t *testing.T, cells [][]formula.Value, ref string, want formula.Value) {
	t.Helper()
	col, row, err := CellNameToCoordinates(ref)
	if err != nil {
		t.Fatalf("CellNameToCoordinates(%s): %v", ref, err)
	}
	if row <= 0 || row > len(cells) {
		t.Fatalf("%s row %d out of range for %dx%d matrix", ref, row, len(cells), matrixWidth(cells))
	}
	if col <= 0 || col > len(cells[row-1]) {
		t.Fatalf("%s col %d out of range for %dx%d matrix", ref, col, len(cells), matrixWidth(cells))
	}
	got := cells[row-1][col-1]
	if got.Type != want.Type || got.Num != want.Num || got.Str != want.Str || got.Bool != want.Bool || got.Err != want.Err || got.NoSpill != want.NoSpill || got.RangeOverflow != want.RangeOverflow {
		t.Fatalf("%s = %#v, want %#v", ref, got, want)
	}
}

func matrixWidth(cells [][]formula.Value) int {
	if len(cells) == 0 {
		return 0
	}
	return len(cells[0])
}

func TestMaterializeRangeMissingSheetReturnsREFMatrix(t *testing.T) {
	res := materializeRange(rangeMaterializationRequest{
		sheet:   "Missing",
		fromCol: 1,
		fromRow: 1,
		toCol:   2,
		toRow:   2,
	}, nil, nil)

	if res.overflow {
		t.Fatalf("overflow = true, want false")
	}
	if len(res.cells) != 2 || len(res.cells[0]) != 2 {
		t.Fatalf("shape = %dx%d, want 2x2", len(res.cells), matrixWidth(res.cells))
	}
	want := formula.ErrorVal(formula.ErrValREF)
	assertCellValue(t, res.cells, "A1", want)
	assertCellValue(t, res.cells, "B1", want)
	assertCellValue(t, res.cells, "A2", want)
	assertCellValue(t, res.cells, "B2", want)
}

func TestMaterializeRangeMissingSheetOverflowSentinel(t *testing.T) {
	res := materializeRange(rangeMaterializationRequest{
		sheet:   "Missing",
		fromCol: 1,
		fromRow: 1,
		toCol:   2,
		toRow:   MaxRows,
	}, nil, nil)

	if !res.overflow {
		t.Fatalf("overflow = false, want true")
	}
	if len(res.cells) != 1 || len(res.cells[0]) != 1 {
		t.Fatalf("overflow shape = %dx%d, want 1x1", len(res.cells), matrixWidth(res.cells))
	}
	cell := res.cells[0][0]
	if cell.Type != formula.ValueError || cell.Err != formula.ErrValREF || !cell.RangeOverflow {
		t.Fatalf("overflow cell = %#v, want #REF! overflow", cell)
	}
}

func TestMaterializeRangeClampsFullRowAndColumn(t *testing.T) {
	grid := &fakeRangeGrid{
		maxRow: 3,
		maxCol: 4,
		cells: []rangeMaterializationCell{
			gridCell("A1", formula.NumberVal(10), true),
			gridCell("C1", formula.NumberVal(30), true),
			gridCell("D3", formula.NumberVal(40), true),
		},
	}

	tests := []struct {
		name string
		req  rangeMaterializationRequest
		want string
	}{
		{
			name: "full column clamp",
			req: rangeMaterializationRequest{
				sheet:   "Sheet1",
				fromCol: 1,
				fromRow: 1,
				toCol:   MaxColumns,
				toRow:   2,
			},
			want: "A1:D2",
		},
		{
			name: "full row clamp",
			req: rangeMaterializationRequest{
				sheet:   "Sheet1",
				fromCol: 1,
				fromRow: 1,
				toCol:   MaxColumns,
				toRow:   1,
			},
			want: "A1:D1",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			res := materializeRange(tt.req, grid, nil)
			if res.overflow {
				t.Fatalf("overflow = true, want false")
			}
			wantFrom, wantTo, err := rangeBoundRefs(tt.want)
			if err != nil {
				t.Fatalf("rangeBoundRefs(%s): %v", tt.want, err)
			}
			if res.bounds.FromCol != wantFrom.col || res.bounds.FromRow != wantFrom.row || res.bounds.ToCol != wantTo.col || res.bounds.ToRow != wantTo.row {
				t.Fatalf("bounds = %+v, want %s", res.bounds, tt.want)
			}
		})
	}
}

func TestMaterializeRangeSparseCellsAndTrailingBlanks(t *testing.T) {
	grid := &fakeRangeGrid{
		maxRow: 3,
		maxCol: 3,
		cells: []rangeMaterializationCell{
			gridCell("A1", formula.NumberVal(1), true),
			gridCell("C3", formula.NumberVal(9), true),
		},
	}

	res := materializeRange(rangeMaterializationRequest{
		sheet:   "Sheet1",
		fromCol: 1,
		fromRow: 1,
		toCol:   3,
		toRow:   3,
	}, grid, nil)

	if res.overflow {
		t.Fatalf("overflow = true, want false")
	}
	if len(res.cells) != 3 || len(res.cells[0]) != 3 {
		t.Fatalf("shape = %dx%d, want 3x3", len(res.cells), matrixWidth(res.cells))
	}
	assertCellValue(t, res.cells, "A1", formula.NumberVal(1))
	assertCellValue(t, res.cells, "B1", formula.EmptyVal())
	assertCellValue(t, res.cells, "C1", formula.EmptyVal())
	assertCellValue(t, res.cells, "B2", formula.EmptyVal())
	assertCellValue(t, res.cells, "C3", formula.NumberVal(9))
}

func TestMaterializeRangePlaceholderOccupancyBlocksOnlyWhenOccupied(t *testing.T) {
	tests := []struct {
		name     string
		occupied bool
		wantB1   formula.Value
	}{
		{
			name:     "placeholder does not block",
			occupied: false,
			wantB1:   formula.NumberVal(20),
		},
		{
			name:     "occupied cell blocks spill",
			occupied: true,
			wantB1:   formula.StringVal("block"),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			placeholderValue := formula.EmptyVal()
			wantB1 := tt.wantB1
			if tt.occupied {
				placeholderValue = tt.wantB1
			}
			grid := &fakeRangeGrid{
				maxRow: 1,
				maxCol: 2,
				cells: []rangeMaterializationCell{
					gridCell("A1", formula.NumberVal(10), true),
					gridCell("B1", placeholderValue, tt.occupied),
				},
			}
			res := materializeRange(rangeMaterializationRequest{
				sheet:   "Sheet1",
				fromCol: 1,
				fromRow: 1,
				toCol:   2,
				toRow:   1,
			}, grid, &fakeRangeSpills{
				anchors: []rangeSpillAnchor{
					spillAnchor("A1", "B1", formula.Value{
						Type:  formula.ValueArray,
						Array: [][]formula.Value{{formula.NumberVal(10), formula.NumberVal(20)}},
					}),
				},
			})

			if res.overflow {
				t.Fatalf("overflow = true, want false")
			}
			assertCellValue(t, res.cells, "A1", formula.NumberVal(10))
			assertCellValue(t, res.cells, "B1", wantB1)
		})
	}
}

func TestMaterializeRangeSpillGrowthAndMetadata(t *testing.T) {
	grid := &fakeRangeGrid{
		maxRow: 2,
		maxCol: 1,
		cells: []rangeMaterializationCell{
			gridCell("A1", formula.NumberVal(1), true),
		},
	}
	res := materializeRange(rangeMaterializationRequest{
		sheet:   "Sheet1",
		fromCol: 1,
		fromRow: 1,
		toCol:   1,
		toRow:   5,
	}, grid, &fakeRangeSpills{
		anchors: []rangeSpillAnchor{
			spillAnchor("A1", "A4", formula.Value{
				Type:  formula.ValueArray,
				Array: [][]formula.Value{{formula.NumberVal(1)}, {formula.NumberVal(2)}, {formula.NumberVal(3)}, {formula.NumberVal(4)}},
			}),
			spillAnchor("A4", "A5", formula.Value{
				Type:  formula.ValueArray,
				Array: [][]formula.Value{{formula.NumberVal(100)}, {formula.NumberVal(200)}},
			}),
		},
	})

	if res.overflow {
		t.Fatalf("overflow = true, want false")
	}
	if res.bounds.ToRow != 5 || res.bounds.ToCol != 1 {
		t.Fatalf("bounds = %+v, want A1:A5", res.bounds)
	}
	if len(res.discoveredDeps) != 2 {
		t.Fatalf("discoveredDeps len = %d, want 2", len(res.discoveredDeps))
	}
	assertCellValue(t, res.cells, "A1", formula.NumberVal(1))
	assertCellValue(t, res.cells, "A2", formula.NumberVal(2))
	assertCellValue(t, res.cells, "A3", formula.NumberVal(3))
	assertCellValue(t, res.cells, "A4", formula.NumberVal(100))
	assertCellValue(t, res.cells, "A5", formula.NumberVal(200))
}

func TestMaterializeRangeBlockedSpillDoesNotGrow(t *testing.T) {
	grid := &fakeRangeGrid{
		maxRow: 2,
		maxCol: 2,
		cells: []rangeMaterializationCell{
			gridCell("A1", formula.ErrorVal(formula.ErrValSPILL), true),
			gridCell("B1", formula.StringVal("blocker"), true),
		},
	}
	res := materializeRange(rangeMaterializationRequest{
		sheet:   "Sheet1",
		fromCol: 1,
		fromRow: 1,
		toCol:   2,
		toRow:   2,
	}, grid, &fakeRangeSpills{
		anchors: []rangeSpillAnchor{
			spillAnchor("A1", "A1", formula.Value{
				Type:  formula.ValueArray,
				Array: [][]formula.Value{{formula.NumberVal(1), formula.NumberVal(2)}},
			}),
		},
	})

	if res.overflow {
		t.Fatalf("overflow = true, want false")
	}
	if len(res.discoveredDeps) != 0 {
		t.Fatalf("discoveredDeps = %#v, want none", res.discoveredDeps)
	}
	assertCellValue(t, res.cells, "A1", formula.ErrorVal(formula.ErrValSPILL))
	assertCellValue(t, res.cells, "B1", formula.StringVal("blocker"))
}

func TestMaterializeRangeReversedBoundsReturnEmptyResult(t *testing.T) {
	res := materializeRange(rangeMaterializationRequest{
		sheet:   "Sheet1",
		fromCol: 3,
		fromRow: 3,
		toCol:   2,
		toRow:   2,
	}, &fakeRangeGrid{maxRow: 3, maxCol: 3}, nil)

	if res.overflow {
		t.Fatalf("overflow = true, want false")
	}
	if res.cells != nil {
		t.Fatalf("cells = %#v, want nil", res.cells)
	}
}

func rangeBoundRefs(ref string) (from, to cellCoord, err error) {
	col1, row1, col2, row2, err := RangeToCoordinates(ref)
	if err != nil {
		return cellCoord{}, cellCoord{}, err
	}
	return cellCoord{col: col1, row: row1}, cellCoord{col: col2, row: row2}, nil
}

type cellCoord struct {
	col int
	row int
}
