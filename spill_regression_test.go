package werkbook_test

import (
	"bytes"
	"os"
	"path/filepath"
	"strconv"
	"testing"

	"github.com/jpoz/werkbook"
	"github.com/jpoz/werkbook/ooxml"
)

type spillWant struct {
	cell string
	typ  werkbook.ValueType
	num  float64
	str  string
}

func numWant(cell string, want float64) spillWant {
	return spillWant{cell: cell, typ: werkbook.TypeNumber, num: want}
}

func strWant(cell, want string) spillWant {
	return spillWant{cell: cell, typ: werkbook.TypeString, str: want}
}

func mustSetSheetName(t *testing.T, f *werkbook.File, oldName, newName string) {
	t.Helper()
	if err := f.SetSheetName(oldName, newName); err != nil {
		t.Fatalf("SetSheetName(%q,%q): %v", oldName, newName, err)
	}
}

func mustNewSheet(t *testing.T, f *werkbook.File, name string) *werkbook.Sheet {
	t.Helper()
	s, err := f.NewSheet(name)
	if err != nil {
		t.Fatalf("NewSheet(%q): %v", name, err)
	}
	return s
}

func mustSetValue(t *testing.T, s *werkbook.Sheet, cell string, value any) {
	t.Helper()
	if err := s.SetValue(cell, value); err != nil {
		t.Fatalf("SetValue(%s): %v", cell, err)
	}
}

func mustSetFormula(t *testing.T, s *werkbook.Sheet, cell, formula string) {
	t.Helper()
	if err := s.SetFormula(cell, formula); err != nil {
		t.Fatalf("SetFormula(%s): %v", cell, err)
	}
}

func assertSheetWants(t *testing.T, s *werkbook.Sheet, wants ...spillWant) {
	t.Helper()
	for _, want := range wants {
		val, err := s.GetValue(want.cell)
		if err != nil {
			t.Fatalf("GetValue(%s): %v", want.cell, err)
		}
		if val.Type != want.typ {
			t.Fatalf("%s type = %v, want %v", want.cell, val.Type, want.typ)
		}
		switch want.typ {
		case werkbook.TypeNumber:
			if val.Number != want.num {
				t.Fatalf("%s = %g, want %g", want.cell, val.Number, want.num)
			}
		case werkbook.TypeString:
			if val.String != want.str {
				t.Fatalf("%s = %q, want %q", want.cell, val.String, want.str)
			}
		default:
			t.Fatalf("unsupported want type %v for %s", want.typ, want.cell)
		}
	}
}

func assertEmptyCell(t *testing.T, s *werkbook.Sheet, cell string) {
	t.Helper()
	val, err := s.GetValue(cell)
	if err != nil {
		t.Fatalf("GetValue(%s): %v", cell, err)
	}
	if val.Type != werkbook.TypeEmpty {
		t.Fatalf("%s = %#v, want empty cell", cell, val)
	}
}

func newSpillHarness(t *testing.T) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet, *werkbook.Sheet) {
	t.Helper()

	f := werkbook.New()
	mustSetSheetName(t, f, "Sheet1", "Data")
	data := f.Sheet("Data")
	spill := mustNewSheet(t, f, "Spill")
	calc := mustNewSheet(t, f, "Calc")
	return f, data, spill, calc
}

func buildVerticalFilterHarness(t *testing.T) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet, *werkbook.Sheet) {
	t.Helper()

	f, data, spill, calc := newSpillHarness(t)
	rows := []struct {
		include bool
		amount  float64
	}{
		{true, 10},
		{false, 20},
		{true, 30},
		{false, 40},
		{true, 50},
	}
	for i, row := range rows {
		cellRow := i + 2
		mustSetValue(t, data, "A"+strconv.Itoa(cellRow), row.include)
		mustSetValue(t, data, "B"+strconv.Itoa(cellRow), row.amount)
	}
	mustSetFormula(t, spill, "B2", `FILTER(Data!B2:B6,Data!A2:A6)`)
	return f, data, spill, calc
}

func buildHorizontalHStackHarness(t *testing.T) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet, *werkbook.Sheet) {
	t.Helper()

	f, data, spill, calc := newSpillHarness(t)
	_ = data
	mustSetFormula(t, spill, "B2", `HSTACK(10,20,30)`)
	return f, data, spill, calc
}

func buildRectangularSequenceHarness(t *testing.T) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet, *werkbook.Sheet) {
	t.Helper()

	f, data, spill, calc := newSpillHarness(t)
	_ = data
	mustSetFormula(t, spill, "B2", `SEQUENCE(2,2,10,5)`)
	return f, data, spill, calc
}

func buildResizableFilterHarness(t *testing.T) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet, *werkbook.Sheet) {
	t.Helper()

	f, data, spill, calc := newSpillHarness(t)
	rows := []struct {
		include bool
		amount  float64
	}{
		{true, 10},
		{false, 20},
		{true, 30},
		{false, 40},
	}
	for i, row := range rows {
		cellRow := i + 2
		mustSetValue(t, data, "A"+strconv.Itoa(cellRow), row.include)
		mustSetValue(t, data, "B"+strconv.Itoa(cellRow), row.amount)
	}
	mustSetFormula(t, spill, "B2", `FILTER(Data!B2:B5,Data!A2:A5)`)
	return f, data, spill, calc
}

func saveAndReopen(t *testing.T, f *werkbook.File) *werkbook.File {
	t.Helper()

	path := filepath.Join(t.TempDir(), "spill-roundtrip.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	return f2
}

func writeFixtureWorkbook(t *testing.T, fixture *ooxml.WorkbookData, filename string) string {
	t.Helper()

	var buf bytes.Buffer
	if err := ooxml.WriteWorkbook(&buf, fixture); err != nil {
		t.Fatalf("WriteWorkbook: %v", err)
	}
	path := filepath.Join(t.TempDir(), filename)
	if err := os.WriteFile(path, buf.Bytes(), 0o600); err != nil {
		t.Fatalf("WriteFile(%s): %v", path, err)
	}
	return path
}

func TestSpillRepresentativeShapes(t *testing.T) {
	tests := []struct {
		name  string
		build func(*testing.T) (*werkbook.File, *werkbook.Sheet)
		wants []spillWant
	}{
		{
			name: "vertical filter",
			build: func(t *testing.T) (*werkbook.File, *werkbook.Sheet) {
				f, _, spill, _ := buildVerticalFilterHarness(t)
				return f, spill
			},
			wants: []spillWant{
				numWant("B2", 10),
				numWant("B3", 30),
				numWant("B4", 50),
			},
		},
		{
			name: "horizontal hstack",
			build: func(t *testing.T) (*werkbook.File, *werkbook.Sheet) {
				f, _, spill, _ := buildHorizontalHStackHarness(t)
				return f, spill
			},
			wants: []spillWant{
				numWant("B2", 10),
				numWant("C2", 20),
				numWant("D2", 30),
			},
		},
		{
			name: "rectangular sequence",
			build: func(t *testing.T) (*werkbook.File, *werkbook.Sheet) {
				f, _, spill, _ := buildRectangularSequenceHarness(t)
				return f, spill
			},
			wants: []spillWant{
				numWant("B2", 10),
				numWant("C2", 15),
				numWant("B3", 20),
				numWant("C3", 25),
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f, spill := tt.build(t)
			f.Recalculate()
			assertSheetWants(t, spill, tt.wants...)
		})
	}
}

func TestSpillConsumerMatrix(t *testing.T) {
	tests := []struct {
		name     string
		build    func(*testing.T) (*werkbook.File, *werkbook.Sheet)
		formulas map[string]string
		wants    []spillWant
	}{
		{
			name: "vertical spill consumers",
			build: func(t *testing.T) (*werkbook.File, *werkbook.Sheet) {
				f, _, _, calc := buildVerticalFilterHarness(t)
				return f, calc
			},
			formulas: map[string]string{
				"A1": `Spill!B3`,
				"A2": `SUM(Spill!B2:B3)`,
				"A3": `COUNTIF(Spill!B2:B10,">20")`,
				"A4": `MATCH(50,Spill!B2:B10,0)`,
			},
			wants: []spillWant{
				numWant("A1", 30),
				numWant("A2", 40),
				numWant("A3", 2),
				numWant("A4", 3),
			},
		},
		{
			name: "horizontal spill consumers",
			build: func(t *testing.T) (*werkbook.File, *werkbook.Sheet) {
				f, _, _, calc := buildHorizontalHStackHarness(t)
				return f, calc
			},
			formulas: map[string]string{
				"A1": `Spill!D2`,
				"A2": `SUM(Spill!B2:D2)`,
				"A3": `MATCH(30,Spill!B2:D2,0)`,
			},
			wants: []spillWant{
				numWant("A1", 30),
				numWant("A2", 60),
				numWant("A3", 3),
			},
		},
		{
			name: "rectangular spill consumers",
			build: func(t *testing.T) (*werkbook.File, *werkbook.Sheet) {
				f, _, _, calc := buildRectangularSequenceHarness(t)
				return f, calc
			},
			formulas: map[string]string{
				"A1": `Spill!C3`,
				"A2": `SUM(Spill!B2:C3)`,
				"A3": `INDEX(Spill!B2:C3,2,1)`,
				"A4": `SUM(Spill!B3:C3)`,
			},
			wants: []spillWant{
				numWant("A1", 25),
				numWant("A2", 70),
				numWant("A3", 20),
				numWant("A4", 45),
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f, calc := tt.build(t)
			for cell, formula := range tt.formulas {
				mustSetFormula(t, calc, cell, formula)
			}
			f.Recalculate()
			assertSheetWants(t, calc, tt.wants...)
		})
	}
}

func TestSpillRecalculationTracksGrowthAndShrink(t *testing.T) {
	f, data, spill, calc := buildResizableFilterHarness(t)
	mustSetFormula(t, calc, "A1", `SUM(Spill!B2:B10)`)
	mustSetFormula(t, calc, "A2", `COUNTA(Spill!B2:B10)`)

	f.Recalculate()
	assertSheetWants(t, spill, numWant("B2", 10), numWant("B3", 30))
	assertEmptyCell(t, spill, "B4")
	assertSheetWants(t, calc, numWant("A1", 40), numWant("A2", 2))

	mustSetValue(t, data, "A5", true)
	f.Recalculate()
	assertSheetWants(t, spill, numWant("B2", 10), numWant("B3", 30), numWant("B4", 40))
	assertSheetWants(t, calc, numWant("A1", 80), numWant("A2", 3))

	mustSetValue(t, data, "A4", false)
	mustSetValue(t, data, "A5", false)
	f.Recalculate()
	assertSheetWants(t, spill, numWant("B2", 10))
	assertEmptyCell(t, spill, "B3")
	assertEmptyCell(t, spill, "B4")
	assertSheetWants(t, calc, numWant("A1", 10), numWant("A2", 1))
}

func TestImportedDynamicArrayMetadataMatrixWithoutPlaceholders(t *testing.T) {
	const (
		dataSheet    = "Data"
		spillSheet   = "Spill"
		calcSheet    = "Calc"
		spillAnchor  = "B2"
		spillRef     = "B2:B4"
		userFormula  = `FILTER(Data!B2:B6,Data!A2:A6)`
		ooxmlFormula = `_xlfn._xlws.FILTER(Data!B2:B6,Data!A2:A6)`
	)

	tests := []struct {
		name            string
		includeMetadata bool
		wantSpill       []spillWant
		emptyCells      []string
		wantCalc        []spillWant
	}{
		{
			name:            "metadata present spills without placeholders",
			includeMetadata: true,
			wantSpill: []spillWant{
				numWant("B2", 10),
				numWant("B3", 30),
				numWant("B4", 50),
			},
			wantCalc: []spillWant{
				numWant("A1", 90),
				numWant("B1", 3),
			},
		},
		{
			name:            "metadata absent stays scalar without placeholders",
			includeMetadata: false,
			wantSpill: []spillWant{
				numWant("B2", 10),
			},
			emptyCells: []string{"B3", "B4"},
			wantCalc: []spillWant{
				numWant("A1", 10),
				numWant("B1", 1),
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			anchorCell := ooxml.CellData{
				Ref:     spillAnchor,
				Value:   "10",
				Formula: ooxmlFormula,
			}
			if tt.includeMetadata {
				anchorCell.FormulaType = "array"
				anchorCell.FormulaRef = spillRef
				anchorCell.IsDynamicArray = true
			}

			fixture := &ooxml.WorkbookData{
				Styles: []ooxml.StyleData{{}},
				Sheets: []ooxml.SheetData{
					{
						Name: dataSheet,
						Rows: []ooxml.RowData{
							{Num: 1, Cells: []ooxml.CellData{
								{Ref: "A1", Type: "s", Value: "Include"},
								{Ref: "B1", Type: "s", Value: "Amount"},
							}},
							{Num: 2, Cells: []ooxml.CellData{
								{Ref: "A2", Type: "b", Value: "1"},
								{Ref: "B2", Value: "10"},
							}},
							{Num: 3, Cells: []ooxml.CellData{
								{Ref: "A3", Type: "b", Value: "0"},
								{Ref: "B3", Value: "20"},
							}},
							{Num: 4, Cells: []ooxml.CellData{
								{Ref: "A4", Type: "b", Value: "1"},
								{Ref: "B4", Value: "30"},
							}},
							{Num: 5, Cells: []ooxml.CellData{
								{Ref: "A5", Type: "b", Value: "0"},
								{Ref: "B5", Value: "40"},
							}},
							{Num: 6, Cells: []ooxml.CellData{
								{Ref: "A6", Type: "b", Value: "1"},
								{Ref: "B6", Value: "50"},
							}},
						},
					},
					{
						Name: spillSheet,
						Rows: []ooxml.RowData{
							{Num: 1, Cells: []ooxml.CellData{{Ref: "B1", Type: "s", Value: "Filtered"}}},
							{Num: 2, Cells: []ooxml.CellData{anchorCell}},
						},
					},
					{
						Name: calcSheet,
						Rows: []ooxml.RowData{{Num: 1, Cells: []ooxml.CellData{
							{Ref: "A1", Value: "10", Formula: `SUM(Spill!B:B)`},
							{Ref: "B1", Value: "1", Formula: `COUNT(Spill!B:B)`},
						}}},
					},
				},
			}

			path := writeFixtureWorkbook(t, fixture, "spill-import-matrix.xlsx")
			f, err := werkbook.Open(path)
			if err != nil {
				t.Fatalf("Open: %v", err)
			}

			gotFormula, err := f.Sheet(spillSheet).GetFormula(spillAnchor)
			if err != nil {
				t.Fatalf("GetFormula(%s): %v", spillAnchor, err)
			}
			if gotFormula != userFormula {
				t.Fatalf("formula round-trip = %q, want %q", gotFormula, userFormula)
			}

			f.Recalculate()

			assertSheetWants(t, f.Sheet(spillSheet), tt.wantSpill...)
			for _, cell := range tt.emptyCells {
				assertEmptyCell(t, f.Sheet(spillSheet), cell)
			}
			assertSheetWants(t, f.Sheet(calcSheet), tt.wantCalc...)
		})
	}
}

func TestDynamicArrayRepresentativeRoundTrip(t *testing.T) {
	tests := []struct {
		name        string
		build       func(*testing.T) (*werkbook.File, *werkbook.Sheet)
		formulaCell string
		wantFormula string
		wantSpill   []spillWant
	}{
		{
			name: "filter round-trip",
			build: func(t *testing.T) (*werkbook.File, *werkbook.Sheet) {
				f, _, spill, _ := buildVerticalFilterHarness(t)
				return f, spill
			},
			formulaCell: "B2",
			wantFormula: `FILTER(Data!B2:B6,Data!A2:A6)`,
			wantSpill: []spillWant{
				numWant("B2", 10),
				numWant("B3", 30),
				numWant("B4", 50),
			},
		},
		{
			name: "hstack round-trip",
			build: func(t *testing.T) (*werkbook.File, *werkbook.Sheet) {
				f, _, spill, _ := buildHorizontalHStackHarness(t)
				return f, spill
			},
			formulaCell: "B2",
			wantFormula: `HSTACK(10,20,30)`,
			wantSpill: []spillWant{
				numWant("B2", 10),
				numWant("C2", 20),
				numWant("D2", 30),
			},
		},
		{
			name: "sequence round-trip",
			build: func(t *testing.T) (*werkbook.File, *werkbook.Sheet) {
				f, _, spill, _ := buildRectangularSequenceHarness(t)
				return f, spill
			},
			formulaCell: "B2",
			wantFormula: `SEQUENCE(2,2,10,5)`,
			wantSpill: []spillWant{
				numWant("B2", 10),
				numWant("C2", 15),
				numWant("B3", 20),
				numWant("C3", 25),
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f, spill := tt.build(t)
			f.Recalculate()

			f2 := saveAndReopen(t, f)
			f2.Recalculate()

			spill2 := f2.Sheet(spill.Name())
			gotFormula, err := spill2.GetFormula(tt.formulaCell)
			if err != nil {
				t.Fatalf("GetFormula(%s): %v", tt.formulaCell, err)
			}
			if gotFormula != tt.wantFormula {
				t.Fatalf("formula round-trip = %q, want %q", gotFormula, tt.wantFormula)
			}
			assertSheetWants(t, spill2, tt.wantSpill...)
		})
	}
}

// TestSpillRangeAggregation locks down that aggregator functions see spilled
// values when reading bounded, full-column, full-row, cross-sheet, and
// defined-name references that cover a dynamic-array spill.
//
// Why this matrix exists: range aggregation over spill ranges has broken
// multiple times (PRs #39, #40, #50). Each combination below is a seam
// that any future refactor to range materialization could silently regress.
//
// Shapes under test:
//   - vertical:    FILTER(Data!B2:B6, Data!A2:A6)     spills B2:B4 on Spill
//   - horizontal:  HSTACK(10, 20, 30)                 spills B2:D2 on Spill
//   - rectangular: SEQUENCE(2, 2, 10, 5)              spills B2:C3 on Spill
func TestSpillRangeAggregation(t *testing.T) {
	type harness struct {
		name    string
		build   func(*testing.T) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet)
		shape   string // "vertical" | "horizontal" | "rectangular"
		spillTL string // top-left of spilled range, e.g. "B2"
	}

	harnesses := []harness{
		{
			name: "vertical_filter",
			build: func(t *testing.T) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet) {
				f, _, spill, calc := buildVerticalFilterHarness(t)
				return f, spill, calc
			},
			shape:   "vertical",
			spillTL: "B2",
		},
		{
			name: "horizontal_hstack",
			build: func(t *testing.T) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet) {
				f, _, spill, calc := buildHorizontalHStackHarness(t)
				return f, spill, calc
			},
			shape:   "horizontal",
			spillTL: "B2",
		},
		{
			name: "rectangular_sequence",
			build: func(t *testing.T) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet) {
				f, _, spill, calc := buildRectangularSequenceHarness(t)
				return f, spill, calc
			},
			shape:   "rectangular",
			spillTL: "B2",
		},
	}

	// Expected totals over each shape's full spill.
	//   vertical:    [10, 30, 50]                sum=90  count=3  min=10 max=50
	//   horizontal:  [10, 20, 30]                sum=60  count=3  min=10 max=30
	//   rectangular: [[10,15],[20,25]]           sum=70  count=4  min=10 max=25
	type expect struct {
		sum, count, min, max float64
	}
	wants := map[string]expect{
		"vertical":    {sum: 90, count: 3, min: 10, max: 50},
		"horizontal":  {sum: 60, count: 3, min: 10, max: 30},
		"rectangular": {sum: 70, count: 4, min: 10, max: 25},
	}

	// rangeFor returns reference strings for different range styles over the
	// spill on sheet "Spill". "bounded" is a generous fixed rectangle that
	// fully covers the spill; "fullcol" uses Spill!B:D which covers every
	// shape's columns.
	//
	// Note: full-row references (Spill!2:3) are intentionally excluded; see
	// TestSpillFullRowRangeReference for that gap.
	rangeFor := func(style, shape string) string {
		switch style {
		case "bounded":
			switch shape {
			case "vertical":
				return "Spill!B2:B10"
			case "horizontal":
				return "Spill!B2:Z2"
			case "rectangular":
				return "Spill!B2:D10"
			}
		case "fullcol":
			return "Spill!B:D"
		case "crosssheet_bounded":
			// same as bounded but emphasizes the cross-sheet path
			switch shape {
			case "vertical":
				return "Spill!B2:B20"
			case "horizontal":
				return "Spill!B2:H2"
			case "rectangular":
				return "Spill!B2:E10"
			}
		}
		t.Fatalf("unknown style/shape combo: %s/%s", style, shape)
		return ""
	}

	styles := []string{"bounded", "fullcol", "crosssheet_bounded"}

	for _, h := range harnesses {
		for _, style := range styles {
			t.Run(h.name+"/"+style, func(t *testing.T) {
				f, _, calc := h.build(t)

				ref := rangeFor(style, h.shape)
				mustSetFormula(t, calc, "A1", "SUM("+ref+")")
				mustSetFormula(t, calc, "A2", "COUNT("+ref+")")
				mustSetFormula(t, calc, "A3", "COUNTA("+ref+")")
				mustSetFormula(t, calc, "A4", "MIN("+ref+")")
				mustSetFormula(t, calc, "A5", "MAX("+ref+")")
				mustSetFormula(t, calc, "A6", "AVERAGE("+ref+")")
				mustSetFormula(t, calc, "A7", `COUNTIF(`+ref+`,">0")`)
				mustSetFormula(t, calc, "A8", `SUMIF(`+ref+`,">0")`)

				f.Recalculate()

				w := wants[h.shape]
				assertSheetWants(t, calc,
					numWant("A1", w.sum),
					numWant("A2", w.count),
					numWant("A3", w.count),
					numWant("A4", w.min),
					numWant("A5", w.max),
					numWant("A6", w.sum/w.count),
					numWant("A7", w.count),
					numWant("A8", w.sum),
				)
			})
		}
	}
}

// TestSpillRangeAggregationWithDefinedName locks down that aggregator
// functions see spilled values through a workbook-scoped defined name.
// This is the exact path that regressed in PR #50 (jpoz/fix-spill-reads),
// where full-column references inside defined names did not expand to
// include published spill bounds.
func TestSpillRangeAggregationWithDefinedName(t *testing.T) {
	f, _, _, calc := buildVerticalFilterHarness(t)

	f.AddDefinedName(werkbook.DefinedName{
		Name:         "SpillCol",
		Value:        "Spill!$B:$B",
		LocalSheetID: -1,
	})
	f.AddDefinedName(werkbook.DefinedName{
		Name:         "SpillBounded",
		Value:        "Spill!$B$2:$B$20",
		LocalSheetID: -1,
	})

	mustSetFormula(t, calc, "A1", `SUM(SpillCol)`)
	mustSetFormula(t, calc, "A2", `COUNT(SpillCol)`)
	mustSetFormula(t, calc, "A3", `SUM(SpillBounded)`)
	mustSetFormula(t, calc, "A4", `COUNTIF(SpillBounded,">20")`)

	f.Recalculate()

	assertSheetWants(t, calc,
		numWant("A1", 90),
		numWant("A2", 3),
		numWant("A3", 90),
		numWant("A4", 2),
	)
}

func TestSpillFromRangeArithmetic(t *testing.T) {
	f, data, spill, _ := newSpillHarness(t)
	for i, v := range []float64{10, 20, 30} {
		mustSetValue(t, data, "B"+strconv.Itoa(i+2), v)
	}
	mustSetFormula(t, spill, "B2", `Data!B2:B4*2`)

	f.Recalculate()
	assertSheetWants(t, spill,
		numWant("B2", 20),
		numWant("B3", 40),
		numWant("B4", 60),
	)

	f2 := saveAndReopen(t, f)
	spill2 := f2.Sheet("Spill")
	if spill2 == nil {
		t.Fatal("reopened workbook missing Spill sheet")
	}
	assertSheetWants(t, spill2,
		numWant("B2", 20),
		numWant("B3", 40),
		numWant("B4", 60),
	)
}

func TestSpillFromReciprocalRangeArithmetic(t *testing.T) {
	f, data, spill, _ := newSpillHarness(t)
	for i, v := range []float64{2, 4, 5} {
		mustSetValue(t, data, "B"+strconv.Itoa(i+2), v)
	}
	mustSetFormula(t, spill, "B2", `1/Data!B2:B4`)

	f.Recalculate()
	assertSheetWants(t, spill,
		numWant("B2", 0.5),
		numWant("B3", 0.25),
		numWant("B4", 0.2),
	)
}

func TestSpillFromIFBroadcast(t *testing.T) {
	f, data, spill, _ := newSpillHarness(t)
	for i, v := range []float64{10, 20, 30} {
		mustSetValue(t, data, "B"+strconv.Itoa(i+2), v)
	}
	mustSetFormula(t, spill, "B2", `IF(Data!B2:B4>15,Data!B2:B4,0)`)

	f.Recalculate()
	assertSheetWants(t, spill,
		numWant("B2", 0),
		numWant("B3", 20),
		numWant("B4", 30),
	)
}

func TestSpillProbeKeepsScalarReductionScalar(t *testing.T) {
	f, data, spill, _ := newSpillHarness(t)
	for i, v := range []float64{10, 20, 30} {
		mustSetValue(t, data, "B"+strconv.Itoa(i+2), v)
	}
	mustSetFormula(t, spill, "B2", `SUM(Data!B2:B4)`)

	f.Recalculate()
	assertSheetWants(t, spill, numWant("B2", 60))
	assertEmptyCell(t, spill, "B3")
}

func TestNonLexicalSpillSerializesDynamicArrayMetadata(t *testing.T) {
	f, data, spill, _ := newSpillHarness(t)
	for i, v := range []float64{10, 20, 30} {
		mustSetValue(t, data, "B"+strconv.Itoa(i+2), v)
	}
	mustSetFormula(t, spill, "B2", `Data!B2:B4*2`)
	f.Recalculate()

	path := filepath.Join(t.TempDir(), "range-arithmetic-spill.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	r, err := os.Open(path)
	if err != nil {
		t.Fatalf("Open saved xlsx: %v", err)
	}
	defer r.Close()
	info, err := r.Stat()
	if err != nil {
		t.Fatalf("Stat saved xlsx: %v", err)
	}
	dataFile, err := ooxml.ReadWorkbook(r, info.Size())
	if err != nil {
		t.Fatalf("ReadWorkbook: %v", err)
	}

	var found bool
	for _, sd := range dataFile.Sheets {
		if sd.Name != "Spill" {
			continue
		}
		for _, rd := range sd.Rows {
			for _, cd := range rd.Cells {
				if cd.Ref != "B2" {
					continue
				}
				found = true
				if !cd.IsDynamicArray {
					t.Fatalf("B2 IsDynamicArray = false, want true")
				}
				if cd.FormulaType != "array" || cd.FormulaRef != "B2:B4" {
					t.Fatalf("B2 formula metadata = type %q ref %q, want array B2:B4",
						cd.FormulaType, cd.FormulaRef)
				}
			}
		}
	}
	if !found {
		t.Fatal("saved workbook missing Spill!B2")
	}
}

// TestSpillArrayContextInheritance locks down that scalar functions
// registered in formula.inheritedArrayArgFuncs actually lift to array
// context when called inside an array-forcing outer (FILTER, SUMPRODUCT,
// HSTACK, etc.). Bugs in this registry are the single most common spill
// failure mode:
//
//   - PR #43: IFERROR/IFNA didn't inherit when wrapping errors inside SUMPRODUCT
//   - PR #44: IF was missing from the list; SUMPRODUCT(mask*IF(...)) broke
//   - PR #46: DATEVALUE was missing; FILTER on DATEVALUE(text_col) broke
//
// Each subtest below covers one function from inheritedArrayArgFuncs. Add a
// row here every time a new function is added to that registry.
func TestSpillArrayContextInheritance(t *testing.T) {
	type rowData struct {
		label  string
		amount float64
		date   string // YYYY-MM-DD text
	}
	// Helper: build a fresh workbook seeded with the same small dataset on
	// sheet Data. B = amount, C = date-text, A = include flag.
	build := func(t *testing.T) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet) {
		t.Helper()
		rows := []rowData{
			{"alpha", 10, "2026-01-01"},
			{"beta", -20, "2026-02-01"},
			{"gamma", 30, "2026-03-01"},
			{"delta", -40, "2026-04-01"},
		}
		f, data, spill, calc := newSpillHarness(t)
		for i, r := range rows {
			n := i + 2
			mustSetValue(t, data, "A"+strconv.Itoa(n), r.label)
			mustSetValue(t, data, "B"+strconv.Itoa(n), r.amount)
			mustSetValue(t, data, "C"+strconv.Itoa(n), r.date)
		}
		_ = spill
		return f, data, calc
	}

	t.Run("IF_inside_SUMPRODUCT", func(t *testing.T) {
		f, _, calc := build(t)
		// Count positive amounts: legacy IF must evaluate element-wise inside
		// the outer SUMPRODUCT array context (PR #44).
		mustSetFormula(t, calc, "A1", `SUMPRODUCT(IF(Data!B2:B5>0,1,0))`)
		f.Recalculate()
		assertSheetWants(t, calc, numWant("A1", 2)) // alpha + gamma
	})

	t.Run("IFERROR_inside_SUMPRODUCT", func(t *testing.T) {
		f, _, calc := build(t)
		// IFERROR inside a legacy array-forcing wrapper (SUMPRODUCT) follows
		// Excel's scalar implicit-intersection rule: the range ref collapses
		// to the formula-cell row before IFERROR runs. Formula is at A1 on
		// calc, so row 1 is outside Data!B2:B5 → implicit intersection yields
		// #VALUE!, 1/#VALUE! is #VALUE!, IFERROR catches it and returns 0.
		// Verified against Excel in
		// testdata/error_propagation/11_iferror_in_sumproduct.xlsx (B1).
		mustSetFormula(t, calc, "A1", `SUMPRODUCT(IFERROR(1/Data!B2:B5,0))`)
		f.Recalculate()
		val, err := calc.GetValue("A1")
		if err != nil {
			t.Fatalf("GetValue: %v", err)
		}
		if val.Type != werkbook.TypeNumber || val.Number != 0 {
			t.Fatalf("IFERROR inside SUMPRODUCT = %#v, want 0", val)
		}
	})

	t.Run("IFNA_inside_SUMPRODUCT", func(t *testing.T) {
		f, _, calc := build(t)
		// IFNA inside SUMPRODUCT follows the same scalar rule: D2:D5 is
		// implicit-intersected at row 1 (outside [2,5]) → #VALUE!. IFNA
		// only catches #N/A, not #VALUE!, so the error propagates through
		// SUMPRODUCT. Verified against Excel in
		// testdata/error_propagation/11_iferror_in_sumproduct.xlsx (A1).
		mustSetValue(t, calc, "D2", 10.0)
		mustSetFormula(t, calc, "D3", `NA()`)
		mustSetValue(t, calc, "D4", 20.0)
		mustSetFormula(t, calc, "D5", `NA()`)
		mustSetFormula(t, calc, "A1", `SUMPRODUCT(IFNA(D2:D5,0))`)
		// Control: the raw sum still propagates #N/A.
		mustSetFormula(t, calc, "A2", `SUMPRODUCT(D2:D5)`)
		f.Recalculate()
		val, err := calc.GetValue("A1")
		if err != nil {
			t.Fatalf("GetValue: %v", err)
		}
		if val.Type != werkbook.TypeError || val.String != "#VALUE!" {
			t.Fatalf("A1 = %#v, want #VALUE!", val)
		}
		ctrl, _ := calc.GetValue("A2")
		if ctrl.Type != werkbook.TypeError || ctrl.String != "#N/A" {
			t.Fatalf("control A2 should be #N/A error, got %#v", ctrl)
		}
	})

	t.Run("ABS_inside_SUMPRODUCT", func(t *testing.T) {
		f, _, calc := build(t)
		// Without array inheritance, ABS would implicit-intersect to a
		// single scalar and SUMPRODUCT would see |B2|.
		mustSetFormula(t, calc, "A1", `SUMPRODUCT(ABS(Data!B2:B5))`)
		f.Recalculate()
		assertSheetWants(t, calc, numWant("A1", 100)) // 10+20+30+40
	})

	t.Run("ISNUMBER_inside_SUMPRODUCT", func(t *testing.T) {
		f, _, calc := build(t)
		// Mix of numbers and labels: count numeric entries across A:C cols.
		// Uses array context so ISNUMBER sees the whole range.
		mustSetFormula(t, calc, "A1", `SUMPRODUCT(--ISNUMBER(Data!B2:B5))`)
		f.Recalculate()
		assertSheetWants(t, calc, numWant("A1", 4))
	})

	t.Run("NOT_inside_SUMPRODUCT", func(t *testing.T) {
		f, _, calc := build(t)
		// Count non-negative values with NOT.
		mustSetFormula(t, calc, "A1", `SUMPRODUCT(--NOT(Data!B2:B5<0))`)
		f.Recalculate()
		assertSheetWants(t, calc, numWant("A1", 2))
	})

	t.Run("N_inside_SUMPRODUCT", func(t *testing.T) {
		f, _, calc := build(t)
		mustSetFormula(t, calc, "A1", `SUMPRODUCT(N(Data!B2:B5>0))`)
		f.Recalculate()
		assertSheetWants(t, calc, numWant("A1", 2))
	})

	t.Run("DATEVALUE_inside_FILTER", func(t *testing.T) {
		f, _, calc := build(t)
		// PR #46: FILTER with DATEVALUE converting a text-date column must
		// lift DATEVALUE to array context so the predicate sees a range.
		// Keep only rows whose date is after 2026-02-15.
		mustSetFormula(t, calc, "A1",
			`FILTER(Data!B2:B5, DATEVALUE(Data!C2:C5)>DATEVALUE("2026-02-15"))`)
		f.Recalculate()
		// Matches rows with date 2026-03-01 and 2026-04-01 → [30, -40]
		assertSheetWants(t, calc,
			numWant("A1", 30),
			numWant("A2", -40),
		)
	})

	t.Run("DATEVALUE_inside_SUMPRODUCT", func(t *testing.T) {
		f, _, calc := build(t)
		mustSetFormula(t, calc, "A1",
			`SUMPRODUCT((DATEVALUE(Data!C2:C5)>DATEVALUE("2026-02-15"))*Data!B2:B5)`)
		f.Recalculate()
		assertSheetWants(t, calc, numWant("A1", -10)) // 30 + -40
	})
}

func absDiff(a, b float64) float64 {
	if a > b {
		return a - b
	}
	return b - a
}

// TestSpillFullColumnSiblingFilters locks down the exact shape that
// regressed in PR #50 (jpoz/fix-spill-reads): a defined name that points
// at a full-column block (Out!$A:$H) where each column is a sibling
// FILTER anchor. The defined-name resolver must honor every sibling's
// published spill bounds when materializing the rectangular block.
//
// A full-size version lives in full_column_filter_bench_test.go as a
// benchmark; this is the small-scale correctness test.
func TestSpillFullColumnSiblingFilters(t *testing.T) {
	f := werkbook.New(werkbook.FirstSheet("Source"))
	source := f.Sheet("Source")

	headers := []string{"Account", "Borrower", "City", "State", "Balance", "DaysLate", "Bucket", "Delinquent"}
	for i, h := range headers {
		ref, _ := werkbook.CoordinatesToCellName(i+1, 1)
		mustSetValue(t, source, ref, h)
	}

	// 6 rows, 3 delinquent (rows 2, 4, 6 → delinquent=1).
	type record struct {
		id         int
		borrower   string
		city       string
		state      string
		balance    float64
		daysLate   float64
		bucket     string
		delinquent float64
	}
	rows := []record{
		{1, "Alice", "NY", "NY", 100, 5, "B1", 0},
		{2, "Bob", "LA", "CA", 200, 15, "B2", 1},
		{3, "Carol", "SF", "CA", 300, 0, "B0", 0},
		{4, "Dan", "DC", "DC", 400, 35, "B3", 1},
		{5, "Eve", "SEA", "WA", 500, 3, "B1", 0},
		{6, "Frank", "BOS", "MA", 600, 90, "B3", 1},
	}
	for i, r := range rows {
		rowN := i + 2
		refA, _ := werkbook.CoordinatesToCellName(1, rowN)
		mustSetValue(t, source, refA, float64(r.id))
		refB, _ := werkbook.CoordinatesToCellName(2, rowN)
		mustSetValue(t, source, refB, r.borrower)
		refC, _ := werkbook.CoordinatesToCellName(3, rowN)
		mustSetValue(t, source, refC, r.city)
		refD, _ := werkbook.CoordinatesToCellName(4, rowN)
		mustSetValue(t, source, refD, r.state)
		refE, _ := werkbook.CoordinatesToCellName(5, rowN)
		mustSetValue(t, source, refE, r.balance)
		refF, _ := werkbook.CoordinatesToCellName(6, rowN)
		mustSetValue(t, source, refF, r.daysLate)
		refG, _ := werkbook.CoordinatesToCellName(7, rowN)
		mustSetValue(t, source, refG, r.bucket)
		refH, _ := werkbook.CoordinatesToCellName(8, rowN)
		mustSetValue(t, source, refH, r.delinquent)
	}

	out := mustNewSheet(t, f, "Out")
	for i, h := range headers {
		ref, _ := werkbook.CoordinatesToCellName(i+1, 1)
		mustSetValue(t, out, ref, h)
	}
	// 8 sibling FILTER anchors, all consuming the same Delinquent criteria.
	for col := 1; col <= 8; col++ {
		colName := werkbook.ColumnNumberToName(col)
		formula := `FILTER(Source!` + colName + `2:` + colName + `7,Source!H2:H7<>0,"No rows")`
		anchor, _ := werkbook.CoordinatesToCellName(col, 2)
		mustSetFormula(t, out, anchor, formula)
	}

	// Workbook-scoped defined name over the full column block.
	if err := f.SetDefinedName(werkbook.DefinedName{
		Name:         "Delinquents",
		Value:        `Out!$A:$H`,
		LocalSheetID: -1,
	}); err != nil {
		t.Fatalf("SetDefinedName: %v", err)
	}

	f.Recalculate()

	vals, err := f.ResolveDefinedName("Delinquents", -1)
	if err != nil {
		t.Fatalf("ResolveDefinedName: %v", err)
	}
	// Expected rows: 1 header + 3 delinquent rows = 4 rows
	if got := len(vals); got != 4 {
		t.Fatalf("rows = %d, want 4 (header + 3 spilled)", got)
	}
	if got := len(vals[0]); got != 8 {
		t.Fatalf("cols = %d, want 8", got)
	}

	// Spot-check a few cells. Header row:
	if vals[0][1].String != "Borrower" {
		t.Fatalf("header[1] = %q, want Borrower", vals[0][1].String)
	}
	// First data row is id=2 (Bob).
	if vals[1][0].Type != werkbook.TypeNumber || vals[1][0].Number != 2 {
		t.Fatalf("row1 col0 = %#v, want 2", vals[1][0])
	}
	if vals[1][1].String != "Bob" {
		t.Fatalf("row1 col1 = %q, want Bob", vals[1][1].String)
	}
	// Third data row is id=6 (Frank, balance 600).
	if vals[3][0].Type != werkbook.TypeNumber || vals[3][0].Number != 6 {
		t.Fatalf("row3 col0 = %#v, want 6", vals[3][0])
	}
	if vals[3][4].Type != werkbook.TypeNumber || vals[3][4].Number != 600 {
		t.Fatalf("row3 col4 = %#v, want 600", vals[3][4])
	}
}

// TestSpillCascade locks down cascading dynamic-array dependencies:
// anchor A feeds anchor B which feeds anchor C. When A's upstream data
// changes, every downstream anchor must recompute with fresh bounds.
// This exercises the dependency graph + spill overlay rebuild path.
func TestSpillCascade(t *testing.T) {
	f := werkbook.New()
	mustSetSheetName(t, f, "Sheet1", "Data")
	data := f.Sheet("Data")
	mid := mustNewSheet(t, f, "Mid")
	out := mustNewSheet(t, f, "Out")

	// Seed data with 5 rows, flag column A and amount column B.
	seeds := []struct {
		keep bool
		amt  float64
	}{
		{true, 10},
		{false, 20},
		{true, 30},
		{true, 40},
		{false, 50},
	}
	for i, s := range seeds {
		n := i + 2
		mustSetValue(t, data, "A"+strconv.Itoa(n), s.keep)
		mustSetValue(t, data, "B"+strconv.Itoa(n), s.amt)
	}

	// Anchor A: FILTER returns kept amounts → [10, 30, 40]
	mustSetFormula(t, mid, "B2", `FILTER(Data!B2:B6,Data!A2:A6)`)
	// Anchor B: ANCHORARRAY is the programmatic form of Excel's spill
	// reference operator (Mid!B2#). It reads anchor A's dynamic bounds,
	// making the downstream anchor re-evaluate whenever A's shape changes.
	// Double every element.
	mustSetFormula(t, out, "B2", `ANCHORARRAY(Mid!B2)*2`)
	// Scalar consumer of the final cascade. Uses a generous bounded range
	// so the consumer picks up whatever shape B2 spills to.
	mustSetFormula(t, out, "A1", `SUM(B2:B20)`)

	f.Recalculate()
	// (10+30+40) * 2 = 160
	assertSheetWants(t, out, numWant("A1", 160))

	// Mutate upstream: flip row 3 to keep, row 5 to drop.
	// A2=T(10), A3=T(20), A4=T(30), A5=F, A6=F → kept = [10,20,30]
	mustSetValue(t, data, "A3", true)
	mustSetValue(t, data, "A5", false)
	f.Recalculate()
	// (10+20+30) * 2 = 120
	assertSheetWants(t, out, numWant("A1", 120))

	// Shrink upstream: disable all rows → kept empty, cascade should
	// produce an empty or zero sum and not leak stale data from prior state.
	for _, r := range []string{"A2", "A3", "A4", "A5", "A6"} {
		mustSetValue(t, data, r, false)
	}
	f.Recalculate()
	val, err := out.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue A1: %v", err)
	}
	if val.Type == werkbook.TypeNumber && val.Number == 0 {
		return // accepted: empty spill → zero sum
	}
	if val.Type == werkbook.TypeError {
		// Excel can produce #CALC! when FILTER has no matches. That
		// propagates through the cascade as an error, which is also fine.
		return
	}
	t.Fatalf("after upstream drain, expected zero or error, got %#v", val)
}

// TestSpillBlockedTransitions locks down that #SPILL! state clears when
// the blocker is removed and re-appears when re-introduced, with the
// overlay generation tracking correctly across transitions.
func TestSpillBlockedTransitions(t *testing.T) {
	f := werkbook.New()
	mustSetSheetName(t, f, "Sheet1", "Data")
	data := f.Sheet("Data")
	spill := mustNewSheet(t, f, "Spill")

	// Seed source data: [10, 30, 50] (3 rows kept)
	mustSetValue(t, data, "A2", true)
	mustSetValue(t, data, "B2", 10.0)
	mustSetValue(t, data, "A3", true)
	mustSetValue(t, data, "B3", 30.0)
	mustSetValue(t, data, "A4", true)
	mustSetValue(t, data, "B4", 50.0)

	mustSetFormula(t, spill, "B2", `FILTER(Data!B2:B4,Data!A2:A4)`)

	// Phase 1: unobstructed. B2:B4 should hold the filtered values.
	f.Recalculate()
	assertSheetWants(t, spill,
		numWant("B2", 10),
		numWant("B3", 30),
		numWant("B4", 50),
	)

	// Phase 2: place a blocker inside the spill range. Anchor B2 must
	// become #SPILL!, and downstream cells should clear.
	mustSetValue(t, spill, "B3", "blocker")
	f.Recalculate()
	anchor, err := spill.GetValue("B2")
	if err != nil {
		t.Fatalf("GetValue B2: %v", err)
	}
	if anchor.Type != werkbook.TypeError || anchor.String != "#SPILL!" {
		t.Fatalf("expected #SPILL! at B2 after blocker, got %#v", anchor)
	}

	// Phase 3: remove the blocker. The anchor must reset to spilled state
	// and the downstream cells must be repopulated.
	mustSetValue(t, spill, "B3", nil)
	f.Recalculate()
	assertSheetWants(t, spill,
		numWant("B2", 10),
		numWant("B3", 30),
		numWant("B4", 50),
	)

	// Phase 4: reintroduce the blocker elsewhere in the range. Confirms
	// that overlay invalidation works on the second blocking transition.
	mustSetValue(t, spill, "B4", "other blocker")
	f.Recalculate()
	anchor, _ = spill.GetValue("B2")
	if anchor.Type != werkbook.TypeError || anchor.String != "#SPILL!" {
		t.Fatalf("expected #SPILL! at B2 after second blocker, got %#v", anchor)
	}
}

// TestSpillFullRowRangeReference locks down sheet-qualified full-row
// range references (Spill!2:3), the mirror of the long-supported full-column
// form (Spill!B:D). Regression coverage for parser + cellref handling of
// row-only refs when a sheet qualifier is present.
func TestSpillFullRowRangeReference(t *testing.T) {
	f, _, _, calc := buildHorizontalHStackHarness(t)
	mustSetFormula(t, calc, "A1", `SUM(Spill!2:3)`)
	f.Recalculate()

	val, err := calc.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue: %v", err)
	}
	if val.Type != werkbook.TypeNumber || val.Number != 60 {
		t.Fatalf("SUM(Spill!2:3) = %#v, want number 60", val)
	}
}

// TestSpillRangeAggregationAfterGrowth locks down that aggregators re-read
// a spill after the underlying data grows/shrinks. This is a generalization
// of TestSpillRecalculationTracksGrowthAndShrink that adds SUMIF/COUNTIF/MAX
// into the consumer set, since those take separate range-materialization
// paths than plain SUM.
func TestSpillRangeAggregationAfterGrowth(t *testing.T) {
	f, data, _, calc := buildResizableFilterHarness(t)

	mustSetFormula(t, calc, "A1", `SUM(Spill!B2:B20)`)
	mustSetFormula(t, calc, "A2", `COUNT(Spill!B2:B20)`)
	mustSetFormula(t, calc, "A3", `MAX(Spill!B2:B20)`)
	mustSetFormula(t, calc, "A4", `COUNTIF(Spill!B2:B20,">25")`)
	mustSetFormula(t, calc, "A5", `SUMIF(Spill!B2:B20,">25")`)

	f.Recalculate()
	// Initial: FILTER result = [10, 30]
	assertSheetWants(t, calc,
		numWant("A1", 40),
		numWant("A2", 2),
		numWant("A3", 30),
		numWant("A4", 1),
		numWant("A5", 30),
	)

	// Grow to [10, 30, 40]
	mustSetValue(t, data, "A5", true)
	f.Recalculate()
	assertSheetWants(t, calc,
		numWant("A1", 80),
		numWant("A2", 3),
		numWant("A3", 40),
		numWant("A4", 2),
		numWant("A5", 70),
	)

	// Shrink to [10]
	mustSetValue(t, data, "A4", false)
	mustSetValue(t, data, "A5", false)
	f.Recalculate()
	assertSheetWants(t, calc,
		numWant("A1", 10),
		numWant("A2", 1),
		numWant("A3", 10),
		numWant("A4", 0),
		numWant("A5", 0),
	)
}
