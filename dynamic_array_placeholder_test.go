package werkbook_test

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/jpoz/werkbook"
	"github.com/jpoz/werkbook/ooxml"
)

func TestImportedDynamicArrayPlaceholdersDoNotBlockSpill(t *testing.T) {
	const (
		dataSheet    = "Data"
		spillSheet   = "Spill"
		calcSheet    = "Calc"
		spillAnchor  = "B2"
		spillRef     = "B2:B4"
		userFormula  = `FILTER(Data!B2:B6,Data!A2:A6)`
		ooxmlFormula = `_xlfn._xlws.FILTER(Data!B2:B6,Data!A2:A6)`
	)

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
					{Num: 2, Cells: []ooxml.CellData{{
						Ref:            spillAnchor,
						Value:          "10",
						Formula:        ooxmlFormula,
						FormulaType:    "array",
						FormulaRef:     spillRef,
						IsDynamicArray: true,
					}}},
					{Num: 3, Cells: []ooxml.CellData{{Ref: "B3"}}},
					{Num: 4, Cells: []ooxml.CellData{{Ref: "B4"}}},
				},
			},
			{
				Name: calcSheet,
				Rows: []ooxml.RowData{{Num: 1, Cells: []ooxml.CellData{
					{Ref: "A1", Value: "10", Formula: `SUM(Spill!B:B)`},
					{Ref: "B1", Value: "1", Formula: `COUNT(Spill!B:B)`},
					{Ref: "C1", Value: "0", Formula: `Spill!B3`},
				}}},
			},
		},
	}

	dir := t.TempDir()
	srcPath := filepath.Join(dir, "dynamic-array-source.xlsx")
	var buf bytes.Buffer
	if err := ooxml.WriteWorkbook(&buf, fixture); err != nil {
		t.Fatalf("WriteWorkbook fixture: %v", err)
	}
	if err := os.WriteFile(srcPath, buf.Bytes(), 0o600); err != nil {
		t.Fatalf("WriteFile fixture: %v", err)
	}

	srcSpillXML := string(readSheetXML(t, srcPath, "xl/worksheets/sheet2.xml"))
	for _, want := range []string{`r="B3"`, `r="B4"`} {
		if !strings.Contains(srcSpillXML, want) {
			t.Fatalf("fixture workbook missing placeholder spill cell %q\nxml: %s", want, srcSpillXML)
		}
	}

	f, err := werkbook.Open(srcPath)
	if err != nil {
		t.Fatalf("Open fixture: %v", err)
	}

	got, err := f.Sheet(spillSheet).GetFormula(spillAnchor)
	if err != nil {
		t.Fatalf("GetFormula(%s): %v", spillAnchor, err)
	}
	if got != userFormula {
		t.Fatalf("formula round-trip = %q, want %q", got, userFormula)
	}

	f.Recalculate()

	assertNumber := func(sheetName, cell string, want float64) {
		t.Helper()
		val, err := f.Sheet(sheetName).GetValue(cell)
		if err != nil {
			t.Fatalf("GetValue(%s!%s): %v", sheetName, cell, err)
		}
		if val.Type != werkbook.TypeNumber || val.Number != want {
			t.Fatalf("%s!%s = %#v, want %g", sheetName, cell, val, want)
		}
	}

	assertNumber(spillSheet, "B2", 10)
	assertNumber(spillSheet, "B3", 30)
	assertNumber(spillSheet, "B4", 50)
	assertNumber(calcSheet, "A1", 90)
	assertNumber(calcSheet, "B1", 3)
	assertNumber(calcSheet, "C1", 30)

	dstPath := filepath.Join(dir, "dynamic-array-roundtrip.xlsx")
	if err := f.SaveAs(dstPath); err != nil {
		t.Fatalf("SaveAs round-trip workbook: %v", err)
	}

	r, err := os.Open(dstPath)
	if err != nil {
		t.Fatalf("Open round-trip xlsx: %v", err)
	}
	defer r.Close()
	info, err := r.Stat()
	if err != nil {
		t.Fatalf("Stat round-trip xlsx: %v", err)
	}
	data, err := ooxml.ReadWorkbook(r, info.Size())
	if err != nil {
		t.Fatalf("ReadWorkbook round-trip: %v", err)
	}

	cachedValue := func(sheetName, cell string) string {
		t.Helper()
		for _, sd := range data.Sheets {
			if sd.Name != sheetName {
				continue
			}
			for _, rd := range sd.Rows {
				for _, cd := range rd.Cells {
					if cd.Ref == cell {
						return cd.Value
					}
				}
			}
		}
		t.Fatalf("cell %s!%s not found in saved workbook", sheetName, cell)
		return ""
	}

	if got := cachedValue(calcSheet, "A1"); got != "90" {
		t.Fatalf("saved cache %s!A1 = %q, want 90", calcSheet, got)
	}
	if got := cachedValue(calcSheet, "B1"); got != "3" {
		t.Fatalf("saved cache %s!B1 = %q, want 3", calcSheet, got)
	}
	if got := cachedValue(calcSheet, "C1"); got != "30" {
		t.Fatalf("saved cache %s!C1 = %q, want 30", calcSheet, got)
	}
}
