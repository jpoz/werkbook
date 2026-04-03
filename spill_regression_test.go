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
