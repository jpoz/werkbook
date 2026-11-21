package ooxml

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"strconv"
	"testing"
)

// writeParts serializes data with WriteWorkbook and returns every file in the
// resulting zip keyed by its archive path (e.g. "xl/sharedStrings.xml"). Use
// this when a test needs to inspect the bytes that werkbook actually emits,
// rather than round-tripping through the reader.
func writeParts(t *testing.T, data *WorkbookData) map[string][]byte {
	t.Helper()
	var buf bytes.Buffer
	if err := WriteWorkbook(&buf, data); err != nil {
		t.Fatalf("WriteWorkbook: %v", err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}
	parts := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open %s: %v", f.Name, err)
		}
		b, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatalf("read %s: %v", f.Name, err)
		}
		parts[f.Name] = b
	}
	return parts
}

// requirePart fails the test if name is missing from parts and returns its
// contents otherwise.
func requirePart(t *testing.T, parts map[string][]byte, name string) []byte {
	t.Helper()
	b, ok := parts[name]
	if !ok {
		t.Fatalf("missing part %q; have: %v", name, partNames(parts))
	}
	return b
}

// unmarshalPart decodes an emitted part into v. Fails the test on error.
func unmarshalPart(t *testing.T, parts map[string][]byte, name string, v any) {
	t.Helper()
	b := requirePart(t, parts, name)
	if err := xml.Unmarshal(b, v); err != nil {
		t.Fatalf("unmarshal %s: %v", name, err)
	}
}

func partNames(parts map[string][]byte) []string {
	names := make([]string, 0, len(parts))
	for k := range parts {
		names = append(names, k)
	}
	return names
}

// ---------------------------------------------------------------------------
// Shared string table deduplication
// ---------------------------------------------------------------------------

// TestWrite_SharedStringTableDedupes confirms the writer collapses identical
// string values into a single shared-string entry, with count/uniqueCount
// reflecting total occurrences and distinct values respectively.
func TestWrite_SharedStringTableDedupes(t *testing.T) {
	const occurrences = 100
	rows := make([]RowData, 0, occurrences)
	for i := range occurrences {
		rows = append(rows, RowData{
			Num: i + 1,
			Cells: []CellData{
				{Ref: cellRef(1, i+1), Type: "s", Value: "hello"},
			},
		})
	}
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{Name: "Sheet1", Rows: rows}},
	}

	parts := writeParts(t, data)

	var sst xlsxSST
	unmarshalPart(t, parts, "xl/sharedStrings.xml", &sst)

	if got := len(sst.SI); got != 1 {
		t.Errorf("len(<si>) = %d, want 1", got)
	}
	if sst.UniqueCount != 1 {
		t.Errorf("uniqueCount = %d, want 1", sst.UniqueCount)
	}
	if sst.Count != occurrences {
		t.Errorf("count = %d, want %d", sst.Count, occurrences)
	}
	if sst.SI[0].T == nil || *sst.SI[0].T != "hello" {
		t.Errorf("si[0].t = %v, want \"hello\"", sst.SI[0].T)
	}

	// Every cell should reference index 0 in the SST.
	var ws xlsxWorksheet
	unmarshalPart(t, parts, "xl/worksheets/sheet1.xml", &ws)
	for _, row := range ws.SheetData.Rows {
		for _, c := range row.Cells {
			if c.T != "s" || c.V != "0" {
				t.Errorf("cell %s: T=%q V=%q, want T=\"s\" V=\"0\"", c.R, c.T, c.V)
			}
		}
	}
}

// ---------------------------------------------------------------------------
// Shared formula emission
// ---------------------------------------------------------------------------

// TestWrite_SharedFormulasEmittedAsStandalone pins the current behavior: when
// a caller provides shared-formula CellData (a master with FormulaType="shared"
// plus children sharing its SharedIndex), the writer emits each cell as an
// independent <f> element with no t="shared" or si attribute. werkbook
// expands shared formulas on read and does not re-deduplicate them on write,
// so identical formulas are re-materialized per cell.
//
// If werkbook ever adds a writer-side shared-formula optimization, this test
// should be updated to assert t="shared" + si on the master and bare si= on
// children.
func TestWrite_SharedFormulasEmittedAsStandalone(t *testing.T) {
	// Simulate what the reader produces after expandSharedFormulas: three
	// cells, each with a fully-materialized formula, none carrying shared
	// markers. The writer should round-trip these as standalone formulas.
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{
				{Num: 1, Cells: []CellData{
					{Ref: "A1", Value: "10"},
					{Ref: "B1", Formula: "A1+1", SharedIndex: -1},
				}},
				{Num: 2, Cells: []CellData{
					{Ref: "A2", Value: "20"},
					{Ref: "B2", Formula: "A2+1", SharedIndex: -1},
				}},
				{Num: 3, Cells: []CellData{
					{Ref: "A3", Value: "30"},
					{Ref: "B3", Formula: "A3+1", SharedIndex: -1},
				}},
			},
		}},
	}

	parts := writeParts(t, data)

	var ws xlsxWorksheet
	unmarshalPart(t, parts, "xl/worksheets/sheet1.xml", &ws)

	var formulaCells []xlsxC
	for _, row := range ws.SheetData.Rows {
		for _, c := range row.Cells {
			if c.FE != nil {
				formulaCells = append(formulaCells, c)
			}
		}
	}
	if got := len(formulaCells); got != 3 {
		t.Fatalf("formula cells = %d, want 3", got)
	}
	for _, c := range formulaCells {
		if c.FE.T != "" {
			t.Errorf("cell %s: f@t = %q, want empty (no shared marker)", c.R, c.FE.T)
		}
		if c.FE.Ref != "" {
			t.Errorf("cell %s: f@ref = %q, want empty", c.R, c.FE.Ref)
		}
		if c.FE.Si != 0 {
			t.Errorf("cell %s: f@si = %d, want 0", c.R, c.FE.Si)
		}
		if c.FE.Text == "" {
			t.Errorf("cell %s: formula text is empty", c.R)
		}
	}

	// Sanity: no <f t="shared"> substring anywhere in the sheet.
	sheetXML := requirePart(t, parts, "xl/worksheets/sheet1.xml")
	if bytes.Contains(sheetXML, []byte(`t="shared"`)) {
		t.Errorf("sheet XML unexpectedly contains t=\"shared\":\n%s", sheetXML)
	}
}

// cellRef builds an A1-style reference from 1-based column and row numbers.
// Only supports the first 26 columns, which is all these tests need.
func cellRef(col, row int) string {
	if col < 1 || col > 26 {
		panic("cellRef: col out of range")
	}
	return string(rune('A'+col-1)) + strconv.Itoa(row)
}
