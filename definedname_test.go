package werkbook_test

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/jpoz/werkbook"
	"github.com/jpoz/werkbook/ooxml"
)

func TestDefinedNameFormulaEval(t *testing.T) {
	// Build a WorkbookData with a defined name and a formula that uses it.
	data := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{
			{
				Name: "Sheet1",
				Rows: []ooxml.RowData{
					{
						Num: 1,
						Cells: []ooxml.CellData{
							{Ref: "A1", Value: "100"},         // plain value
							{Ref: "B1", Formula: "A1+MyName"}, // uses defined name
						},
					},
					{
						Num: 10,
						Cells: []ooxml.CellData{
							{Ref: "A10", Value: "42"}, // target of the defined name
						},
					},
				},
			},
		},
		DefinedNames: []ooxml.DefinedName{
			{Name: "MyName", Value: "Sheet1!$A$10", LocalSheetID: -1},
		},
	}

	// Write to a temp file and re-open to exercise the full path.
	dir := t.TempDir()
	path := filepath.Join(dir, "defname.xlsx")
	out, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := ooxml.WriteWorkbook(out, data); err != nil {
		out.Close()
		t.Fatal(err)
	}
	out.Close()

	f, err := werkbook.Open(path)
	if err != nil {
		t.Fatal(err)
	}

	s := f.Sheet("Sheet1")
	if s == nil {
		t.Fatal("Sheet1 not found")
	}

	// B1 = A1 + MyName = 100 + 42 = 142
	v, err := s.GetValue("B1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != werkbook.TypeNumber || v.Number != 142 {
		t.Errorf("B1 = %v, want 142", v)
	}
}

func TestDefinedNameRoundTrip(t *testing.T) {
	// Create a workbook with defined names, save, re-open, check names are preserved.
	data := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{
			{
				Name: "Sheet1",
				Rows: []ooxml.RowData{
					{Num: 1, Cells: []ooxml.CellData{{Ref: "A1", Value: "1"}}},
				},
			},
		},
		DefinedNames: []ooxml.DefinedName{
			{Name: "TestName", Value: "Sheet1!$A$1", LocalSheetID: -1},
			{Name: "LocalName", Value: "Sheet1!$B$2", LocalSheetID: 0},
		},
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "roundtrip.xlsx")
	out, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := ooxml.WriteWorkbook(out, data); err != nil {
		out.Close()
		t.Fatal(err)
	}
	out.Close()

	// Re-read and verify defined names survive.
	osf, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer osf.Close()
	info, _ := osf.Stat()
	data2, err := ooxml.ReadWorkbook(osf, info.Size())
	if err != nil {
		t.Fatal(err)
	}

	if len(data2.DefinedNames) != 2 {
		t.Fatalf("got %d defined names, want 2", len(data2.DefinedNames))
	}
	if data2.DefinedNames[0].Name != "TestName" || data2.DefinedNames[0].Value != "Sheet1!$A$1" {
		t.Errorf("defined name 0 = %+v", data2.DefinedNames[0])
	}
	if data2.DefinedNames[1].Name != "LocalName" || data2.DefinedNames[1].LocalSheetID != 0 {
		t.Errorf("defined name 1 = %+v", data2.DefinedNames[1])
	}
}

func TestDefinedNamesAccessor(t *testing.T) {
	data := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{
			{Name: "Sheet1"},
			{Name: "Sheet2"},
		},
		DefinedNames: []ooxml.DefinedName{
			{Name: "GlobalName", Value: "Sheet1!$A$1", LocalSheetID: -1},
			{Name: "LocalName", Value: "Sheet2!$B$2", LocalSheetID: 1},
		},
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "defined-names.xlsx")
	out, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := ooxml.WriteWorkbook(out, data); err != nil {
		out.Close()
		t.Fatal(err)
	}
	if err := out.Close(); err != nil {
		t.Fatal(err)
	}

	f, err := werkbook.Open(path)
	if err != nil {
		t.Fatal(err)
	}

	got := f.DefinedNames()
	if len(got) != 2 {
		t.Fatalf("got %d defined names, want 2", len(got))
	}
	if got[0].Name != "GlobalName" || got[0].Value != "Sheet1!$A$1" || got[0].LocalSheetID != -1 {
		t.Fatalf("defined name 0 = %+v", got[0])
	}
	if got[1].Name != "LocalName" || got[1].Value != "Sheet2!$B$2" || got[1].LocalSheetID != 1 {
		t.Fatalf("defined name 1 = %+v", got[1])
	}

	got[0].Name = "mutated"
	again := f.DefinedNames()
	if again[0].Name != "GlobalName" {
		t.Fatalf("DefinedNames returned aliasing slice, got %+v", again[0])
	}
}

func TestResolveDefinedNameSingleCell(t *testing.T) {
	data := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{
			{
				Name: "Sheet1",
				Rows: []ooxml.RowData{
					{Num: 1, Cells: []ooxml.CellData{{Ref: "A1", Value: "42"}}},
				},
			},
		},
		DefinedNames: []ooxml.DefinedName{
			{Name: "MyCell", Value: "Sheet1!$A$1", LocalSheetID: -1},
		},
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "resolve_single.xlsx")
	out, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := ooxml.WriteWorkbook(out, data); err != nil {
		out.Close()
		t.Fatal(err)
	}
	out.Close()

	f, err := werkbook.Open(path)
	if err != nil {
		t.Fatal(err)
	}

	vals, err := f.ResolveDefinedName("MyCell", -1)
	if err != nil {
		t.Fatal(err)
	}
	if len(vals) != 1 || len(vals[0]) != 1 {
		t.Fatalf("expected 1x1 grid, got %dx%d", len(vals), len(vals[0]))
	}
	if vals[0][0].Type != werkbook.TypeNumber || vals[0][0].Number != 42 {
		t.Errorf("got %v, want 42", vals[0][0])
	}
}

func TestResolveDefinedNameRange(t *testing.T) {
	data := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{
			{
				Name: "Sheet1",
				Rows: []ooxml.RowData{
					{Num: 1, Cells: []ooxml.CellData{
						{Ref: "A1", Value: "1"},
						{Ref: "B1", Value: "2"},
					}},
					{Num: 2, Cells: []ooxml.CellData{
						{Ref: "A2", Value: "3"},
						{Ref: "B2", Value: "4"},
					}},
				},
			},
		},
		DefinedNames: []ooxml.DefinedName{
			{Name: "MyRange", Value: "Sheet1!$A$1:$B$2", LocalSheetID: -1},
		},
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "resolve_range.xlsx")
	out, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := ooxml.WriteWorkbook(out, data); err != nil {
		out.Close()
		t.Fatal(err)
	}
	out.Close()

	f, err := werkbook.Open(path)
	if err != nil {
		t.Fatal(err)
	}

	vals, err := f.ResolveDefinedName("MyRange", -1)
	if err != nil {
		t.Fatal(err)
	}
	if len(vals) != 2 || len(vals[0]) != 2 {
		t.Fatalf("expected 2x2 grid, got %dx%d", len(vals), len(vals[0]))
	}
	want := [][]float64{{1, 2}, {3, 4}}
	for r, row := range vals {
		for c, v := range row {
			if v.Type != werkbook.TypeNumber || v.Number != want[r][c] {
				t.Errorf("[%d][%d] = %v, want %v", r, c, v, want[r][c])
			}
		}
	}
}

func TestResolveDefinedNameCaseInsensitive(t *testing.T) {
	data := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{
			{
				Name: "Sheet1",
				Rows: []ooxml.RowData{
					{Num: 1, Cells: []ooxml.CellData{{Ref: "A1", Value: "99"}}},
				},
			},
		},
		DefinedNames: []ooxml.DefinedName{
			{Name: "Revenue", Value: "Sheet1!$A$1", LocalSheetID: -1},
		},
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "resolve_case.xlsx")
	out, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := ooxml.WriteWorkbook(out, data); err != nil {
		out.Close()
		t.Fatal(err)
	}
	out.Close()

	f, err := werkbook.Open(path)
	if err != nil {
		t.Fatal(err)
	}

	vals, err := f.ResolveDefinedName("revenue", -1)
	if err != nil {
		t.Fatal(err)
	}
	if vals[0][0].Type != werkbook.TypeNumber || vals[0][0].Number != 99 {
		t.Errorf("got %v, want 99", vals[0][0])
	}
}

func TestResolveDefinedNameSheetScoped(t *testing.T) {
	data := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{
			{
				Name: "Sheet1",
				Rows: []ooxml.RowData{
					{Num: 1, Cells: []ooxml.CellData{{Ref: "A1", Value: "10"}}},
				},
			},
			{
				Name: "Sheet2",
				Rows: []ooxml.RowData{
					{Num: 1, Cells: []ooxml.CellData{{Ref: "A1", Value: "20"}}},
				},
			},
		},
		DefinedNames: []ooxml.DefinedName{
			{Name: "Rate", Value: "Sheet1!$A$1", LocalSheetID: -1}, // global
			{Name: "Rate", Value: "Sheet2!$A$1", LocalSheetID: 1},  // local to Sheet2
		},
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "resolve_scope.xlsx")
	out, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := ooxml.WriteWorkbook(out, data); err != nil {
		out.Close()
		t.Fatal(err)
	}
	out.Close()

	f, err := werkbook.Open(path)
	if err != nil {
		t.Fatal(err)
	}

	// From Sheet2 context (index 1), should get the local name → 20.
	vals, err := f.ResolveDefinedName("Rate", 1)
	if err != nil {
		t.Fatal(err)
	}
	if vals[0][0].Number != 20 {
		t.Errorf("sheet-scoped: got %v, want 20", vals[0][0])
	}

	// From Sheet1 context (index 0), should get the global name → 10.
	vals, err = f.ResolveDefinedName("Rate", 0)
	if err != nil {
		t.Fatal(err)
	}
	if vals[0][0].Number != 10 {
		t.Errorf("global: got %v, want 10", vals[0][0])
	}
}

func TestResolveDefinedNameNotFound(t *testing.T) {
	data := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{
			{Name: "Sheet1", Rows: []ooxml.RowData{
				{Num: 1, Cells: []ooxml.CellData{{Ref: "A1", Value: "1"}}},
			}},
		},
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "resolve_notfound.xlsx")
	out, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := ooxml.WriteWorkbook(out, data); err != nil {
		out.Close()
		t.Fatal(err)
	}
	out.Close()

	f, err := werkbook.Open(path)
	if err != nil {
		t.Fatal(err)
	}

	_, err = f.ResolveDefinedName("DoesNotExist", -1)
	if err == nil {
		t.Fatal("expected error for missing defined name")
	}
}

func TestSetDefinedNameRebuildsFormulaState(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	if err := s.SetValue("A1", 5); err != nil {
		t.Fatal(err)
	}
	if err := s.SetValue("A2", 7); err != nil {
		t.Fatal(err)
	}
	if err := s.SetFormula("B1", "Target+1"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetDefinedName(werkbook.DefinedName{
		Name:         "Target",
		Value:        "Sheet1!$A$1",
		LocalSheetID: -1,
	}); err != nil {
		t.Fatal(err)
	}

	v, err := s.GetValue("B1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != werkbook.TypeNumber || v.Number != 6 {
		t.Fatalf("B1 = %#v, want 6", v)
	}

	if err := f.SetDefinedName(werkbook.DefinedName{
		Name:         "Target",
		Value:        "Sheet1!$A$2",
		LocalSheetID: -1,
	}); err != nil {
		t.Fatal(err)
	}

	v, err = s.GetValue("B1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != werkbook.TypeNumber || v.Number != 8 {
		t.Fatalf("B1 after rename = %#v, want 8", v)
	}
}

func TestLocalSheetScopedDefinedNamesFixtures(t *testing.T) {
	// Build a workbook with locally-scoped defined names programmatically
	// instead of relying on external fixture files.
	//
	// Sheets: "North Ops" (idx 0), "South & West" (idx 1), "Summary" (idx 2)
	//
	// Each data sheet has raw data in B/C (owner, tag) and F/G (bool, adjusted score).
	// Defined names are sheet-local: EscalateFloor, FocusPattern, DeskLabel, LiteralMask.
	// Global names: TargetOwner, GlobalBoost.
	// Summary sheet formulas reference these names to produce expected results.

	northRows := []ooxml.RowData{
		{Num: 1, Cells: []ooxml.CellData{
			{Ref: "H1", Value: "180"},
			{Ref: "H2", Type: "s", Value: "X*"},
			{Ref: "H3", Type: "s", Value: "North Desk"},
			{Ref: "H4", Type: "s", Value: "X-1"},
		}},
		{Num: 2, Cells: []ooxml.CellData{
			{Ref: "B2", Type: "s", Value: "Alpha"},
			{Ref: "C2", Type: "s", Value: "X-1"},
			{Ref: "F2", Type: "b", Value: "1"},
			{Ref: "G2", Value: "235"},
		}},
		{Num: 3, Cells: []ooxml.CellData{
			{Ref: "B3", Type: "s", Value: "Beta"},
			{Ref: "C3", Type: "s", Value: "Y-1"},
			{Ref: "F3", Type: "b", Value: "0"},
			{Ref: "G3", Value: "160"},
		}},
		{Num: 4, Cells: []ooxml.CellData{
			{Ref: "B4", Type: "s", Value: "Alpha"},
			{Ref: "C4", Type: "s", Value: "X-2"},
			{Ref: "F4", Type: "b", Value: "1"},
			{Ref: "G4", Value: "215"},
		}},
		{Num: 5, Cells: []ooxml.CellData{
			{Ref: "B5", Type: "s", Value: "Alpha"},
			{Ref: "C5", Type: "s", Value: "Y-2"},
			{Ref: "F5", Type: "b", Value: "0"},
			{Ref: "G5", Value: "240"},
		}},
		{Num: 6, Cells: []ooxml.CellData{
			{Ref: "B6", Type: "s", Value: "Alpha"},
			{Ref: "C6", Type: "s", Value: "X-3"},
			{Ref: "F6", Type: "b", Value: "0"},
			{Ref: "G6", Value: "120"},
		}},
	}

	southRows := []ooxml.RowData{
		{Num: 1, Cells: []ooxml.CellData{
			{Ref: "H1", Value: "250"},
			{Ref: "H2", Type: "s", Value: "P*"},
			{Ref: "H3", Type: "s", Value: "Field Desk"},
			{Ref: "H4", Type: "s", Value: "P-3"},
		}},
		{Num: 2, Cells: []ooxml.CellData{
			{Ref: "B2", Type: "s", Value: "Alpha"},
			{Ref: "C2", Type: "s", Value: "P-1"},
			{Ref: "F2", Type: "b", Value: "1"},
			{Ref: "G2", Value: "285"},
		}},
		{Num: 3, Cells: []ooxml.CellData{
			{Ref: "B3", Type: "s", Value: "Beta"},
			{Ref: "C3", Type: "s", Value: "P-2"},
			{Ref: "F3", Type: "b", Value: "0"},
			{Ref: "G3", Value: "180"},
		}},
		{Num: 4, Cells: []ooxml.CellData{
			{Ref: "B4", Type: "s", Value: "Alpha"},
			{Ref: "C4", Type: "s", Value: "Q-1"},
			{Ref: "F4", Type: "b", Value: "0"},
			{Ref: "G4", Value: "310"},
		}},
		{Num: 5, Cells: []ooxml.CellData{
			{Ref: "B5", Type: "s", Value: "Alpha"},
			{Ref: "C5", Type: "s", Value: "P-3"},
			{Ref: "F5", Type: "b", Value: "1"},
			{Ref: "G5", Value: "300"},
		}},
		{Num: 6, Cells: []ooxml.CellData{
			{Ref: "B6", Type: "s", Value: "Gamma"},
			{Ref: "C6", Type: "s", Value: "Q-2"},
			{Ref: "F6", Type: "b", Value: "0"},
			{Ref: "G6", Value: "140"},
		}},
	}

	summaryRows := []ooxml.RowData{
		{Num: 1, Cells: []ooxml.CellData{
			{Ref: "A1", Type: "s", Value: "Alpha"},
			{Ref: "A2", Value: "35"},
		}},
		{Num: 2, Cells: []ooxml.CellData{
			{Ref: "B2", Formula: "'North Ops'!EscalateFloor"},
		}},
		{Num: 3, Cells: []ooxml.CellData{
			{Ref: "B3", Formula: "'South & West'!EscalateFloor"},
		}},
		{Num: 4, Cells: []ooxml.CellData{
			{Ref: "B4", Formula: "COUNTIFS('North Ops'!B2:B6,TargetOwner,'North Ops'!C2:C6,'North Ops'!FocusPattern)"},
		}},
		{Num: 5, Cells: []ooxml.CellData{
			{Ref: "B5", Formula: "COUNTIFS('South & West'!B2:B6,TargetOwner,'South & West'!C2:C6,'South & West'!FocusPattern)"},
		}},
		{Num: 6, Cells: []ooxml.CellData{
			{Ref: "B6", Formula: "SUMIF('North Ops'!F2:F6,TRUE,'North Ops'!G2:G6)"},
		}},
		{Num: 7, Cells: []ooxml.CellData{
			{Ref: "B7", Formula: "SUMIF('South & West'!F2:F6,TRUE,'South & West'!G2:G6)"},
		}},
		{Num: 8, Cells: []ooxml.CellData{
			{Ref: "B8", Formula: "'North Ops'!DeskLabel"},
		}},
		{Num: 9, Cells: []ooxml.CellData{
			{Ref: "B9", Formula: "'South & West'!DeskLabel"},
		}},
		{Num: 10, Cells: []ooxml.CellData{
			{Ref: "B10", Formula: "IF(B6>B7,B8,B9)"},
		}},
		{Num: 11, Cells: []ooxml.CellData{
			{Ref: "B11", Formula: "COUNTIF('North Ops'!C2:C6,'North Ops'!LiteralMask)"},
		}},
		{Num: 12, Cells: []ooxml.CellData{
			{Ref: "B12", Formula: "COUNTIF('South & West'!C2:C6,'South & West'!LiteralMask)"},
		}},
		{Num: 13, Cells: []ooxml.CellData{
			{Ref: "B13", Formula: "COUNTIF('North Ops'!F2:F6,TRUE)+COUNTIF('South & West'!F2:F6,TRUE)"},
		}},
		{Num: 14, Cells: []ooxml.CellData{
			{Ref: "B14", Formula: "SUM(B6:B7)"},
		}},
		{Num: 15, Cells: []ooxml.CellData{
			{Ref: "B15", Formula: "B3-B2"},
		}},
		{Num: 16, Cells: []ooxml.CellData{
			{Ref: "B16", Formula: "IF(B14>1000,\"OVER\",\"UNDER\")"},
		}},
		{Num: 17, Cells: []ooxml.CellData{
			{Ref: "B17", Formula: "IF(B4>B5,B8,B9)"},
		}},
	}

	data := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{
			{Name: "North Ops", Rows: northRows},
			{Name: "South & West", Rows: southRows},
			{Name: "Summary", Rows: summaryRows},
		},
		DefinedNames: []ooxml.DefinedName{
			// Global names
			{Name: "TargetOwner", Value: "Summary!$A$1", LocalSheetID: -1},
			{Name: "GlobalBoost", Value: "Summary!$A$2", LocalSheetID: -1},
			// North Ops local names (sheet index 0)
			{Name: "EscalateFloor", Value: "'North Ops'!$H$1", LocalSheetID: 0},
			{Name: "FocusPattern", Value: "'North Ops'!$H$2", LocalSheetID: 0},
			{Name: "DeskLabel", Value: "'North Ops'!$H$3", LocalSheetID: 0},
			{Name: "LiteralMask", Value: "'North Ops'!$H$4", LocalSheetID: 0},
			// South & West local names (sheet index 1)
			{Name: "EscalateFloor", Value: "'South & West'!$H$1", LocalSheetID: 1},
			{Name: "FocusPattern", Value: "'South & West'!$H$2", LocalSheetID: 1},
			{Name: "DeskLabel", Value: "'South & West'!$H$3", LocalSheetID: 1},
			{Name: "LiteralMask", Value: "'South & West'!$H$4", LocalSheetID: 1},
		},
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "local_name_scopes.xlsx")
	out, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := ooxml.WriteWorkbook(out, data); err != nil {
		out.Close()
		t.Fatal(err)
	}
	out.Close()

	f, err := werkbook.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	f.Recalculate()

	summary := f.Sheet("Summary")
	if summary == nil {
		t.Fatal("Summary sheet not found")
	}

	expected := []struct {
		cell string
		want any
	}{
		{cell: "B2", want: float64(180)},
		{cell: "B3", want: float64(250)},
		{cell: "B4", want: float64(3)},
		{cell: "B5", want: float64(2)},
		{cell: "B6", want: float64(450)},
		{cell: "B7", want: float64(585)},
		{cell: "B8", want: "North Desk"},
		{cell: "B9", want: "Field Desk"},
		{cell: "B10", want: "Field Desk"},
		{cell: "B11", want: float64(1)},
		{cell: "B12", want: float64(1)},
		{cell: "B13", want: float64(4)},
		{cell: "B14", want: float64(1035)},
		{cell: "B15", want: float64(70)},
		{cell: "B16", want: "OVER"},
		{cell: "B17", want: "North Desk"},
	}

	for _, tc := range expected {
		v, err := summary.GetValue(tc.cell)
		if err != nil {
			t.Fatalf("%s: %v", tc.cell, err)
		}
		switch want := tc.want.(type) {
		case float64:
			if v.Type != werkbook.TypeNumber || v.Number != want {
				t.Fatalf("%s = %#v, want %v", tc.cell, v, want)
			}
		case string:
			if v.Type != werkbook.TypeString || v.String != want {
				t.Fatalf("%s = %#v, want %q", tc.cell, v, want)
			}
		default:
			t.Fatalf("unsupported expected type %T", tc.want)
		}
	}
}
