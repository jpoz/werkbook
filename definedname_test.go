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
							{Ref: "A1", Value: "100"},          // plain value
							{Ref: "B1", Formula: "A1+MyName"},  // uses defined name
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
