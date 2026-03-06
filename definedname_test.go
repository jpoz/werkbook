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
