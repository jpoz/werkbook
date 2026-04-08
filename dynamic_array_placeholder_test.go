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

func TestImportedPlainDynamicArrayFormulaDoesNotSpillWithoutMetadata(t *testing.T) {
	const (
		dataSheet    = "Data"
		spillSheet   = "Spill"
		calcSheet    = "Calc"
		spillAnchor  = "B2"
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
						Ref:     spillAnchor,
						Value:   "10",
						Formula: ooxmlFormula,
					}}},
					{Num: 3, Cells: []ooxml.CellData{{Ref: "B3"}}},
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
	srcPath := filepath.Join(dir, "plain-dynamic-array-source.xlsx")
	var buf bytes.Buffer
	if err := ooxml.WriteWorkbook(&buf, fixture); err != nil {
		t.Fatalf("WriteWorkbook fixture: %v", err)
	}
	if err := os.WriteFile(srcPath, buf.Bytes(), 0o600); err != nil {
		t.Fatalf("WriteFile fixture: %v", err)
	}

	srcSpillXML := string(readSheetXML(t, srcPath, "xl/worksheets/sheet2.xml"))
	for _, unexpected := range []string{`cm="1"`, `<f t="array"`, `aca="1"`} {
		if strings.Contains(srcSpillXML, unexpected) {
			t.Fatalf("fixture workbook unexpectedly wrote dynamic-array metadata %q\nxml: %s", unexpected, srcSpillXML)
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
	val, err := f.Sheet(spillSheet).GetValue("B3")
	if err != nil {
		t.Fatalf("GetValue(%s!B3): %v", spillSheet, err)
	}
	if val.Type != werkbook.TypeEmpty {
		t.Fatalf("%s!B3 = %#v, want empty", spillSheet, val)
	}
	assertNumber(calcSheet, "A1", 10)
	assertNumber(calcSheet, "B1", 1)
	assertNumber(calcSheet, "C1", 0)
}

// TestSpillShrinkClearsStaleData verifies that when a FILTER result shrinks
// (e.g., from 3 rows to 1 row after data changes), the old spill cells
// become empty and aggregate functions like SUM/COUNT reflect the smaller result.
func TestSpillShrinkClearsStaleData(t *testing.T) {
	const (
		dataSheet  = "Data"
		spillSheet = "Spill"
		calcSheet  = "Calc"
	)

	// Data: 5 rows, 3 included initially.
	f := werkbook.New(werkbook.FirstSheet(dataSheet))
	ds := f.Sheet(dataSheet)
	for _, c := range []struct {
		cell string
		val  any
	}{
		{"A2", true}, {"B2", 10.0},
		{"A3", false}, {"B3", 20.0},
		{"A4", true}, {"B4", 30.0},
		{"A5", false}, {"B5", 40.0},
		{"A6", true}, {"B6", 50.0},
	} {
		if err := ds.SetValue(c.cell, c.val); err != nil {
			t.Fatalf("SetValue %s: %v", c.cell, err)
		}
	}

	ss, err := f.NewSheet(spillSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := ss.SetFormula("B2", "FILTER(Data!B2:B6,Data!A2:A6)"); err != nil {
		t.Fatal(err)
	}

	cs, err := f.NewSheet(calcSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := cs.SetFormula("A1", "SUM(Spill!B:B)"); err != nil {
		t.Fatal(err)
	}
	if err := cs.SetFormula("B1", "COUNT(Spill!B:B)"); err != nil {
		t.Fatal(err)
	}

	// Phase 1: FILTER returns 3 rows (10, 30, 50).
	f.Recalculate()

	assertNum := func(sheet, cell string, want float64) {
		t.Helper()
		v, err := f.Sheet(sheet).GetValue(cell)
		if err != nil {
			t.Fatalf("GetValue(%s!%s): %v", sheet, cell, err)
		}
		if v.Type != werkbook.TypeNumber || v.Number != want {
			t.Fatalf("%s!%s = %#v, want %g", sheet, cell, v, want)
		}
	}
	assertEmpty := func(sheet, cell string) {
		t.Helper()
		v, err := f.Sheet(sheet).GetValue(cell)
		if err != nil {
			t.Fatalf("GetValue(%s!%s): %v", sheet, cell, err)
		}
		if v.Type != werkbook.TypeEmpty {
			t.Fatalf("%s!%s = %#v, want empty", sheet, cell, v)
		}
	}

	assertNum(spillSheet, "B2", 10)
	assertNum(spillSheet, "B3", 30)
	assertNum(spillSheet, "B4", 50)
	assertNum(calcSheet, "A1", 90)
	assertNum(calcSheet, "B1", 3)

	// Phase 2: Change data so only 1 row matches.
	if err := ds.SetValue("A4", false); err != nil {
		t.Fatal(err)
	}
	if err := ds.SetValue("A6", false); err != nil {
		t.Fatal(err)
	}

	f.Recalculate()

	assertNum(spillSheet, "B2", 10)
	assertEmpty(spillSheet, "B3")
	assertEmpty(spillSheet, "B4")
	assertNum(calcSheet, "A1", 10)
	assertNum(calcSheet, "B1", 1)
}

// TestCachedSpillValuesOverwrittenOnRecalc verifies that non-empty cached spill
// values from OOXML import are replaced by live FILTER results after recalc.
func TestCachedSpillValuesOverwrittenOnRecalc(t *testing.T) {
	const (
		dataSheet  = "Data"
		spillSheet = "Spill"
		calcSheet  = "Calc"
	)

	ooxmlFormula := `_xlfn._xlws.FILTER(Data!B2:B6,Data!A2:A6)`

	fixture := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{
			{
				Name: dataSheet,
				Rows: []ooxml.RowData{
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
					// Anchor with dynamic array metadata. Cached value is stale (999).
					{Num: 2, Cells: []ooxml.CellData{{
						Ref:            "B2",
						Value:          "999",
						Formula:        ooxmlFormula,
						FormulaType:    "array",
						FormulaRef:     "B2:B4",
						IsDynamicArray: true,
					}}},
					// Cached spill cells with stale non-empty values.
					{Num: 3, Cells: []ooxml.CellData{{Ref: "B3", Value: "888"}}},
					{Num: 4, Cells: []ooxml.CellData{{Ref: "B4", Value: "777"}}},
				},
			},
			{
				Name: calcSheet,
				Rows: []ooxml.RowData{{Num: 1, Cells: []ooxml.CellData{
					{Ref: "A1", Value: "0", Formula: `SUM(Spill!B:B)`},
					{Ref: "B1", Value: "0", Formula: `COUNT(Spill!B:B)`},
				}}},
			},
		},
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "cached-spill.xlsx")
	var buf bytes.Buffer
	if err := ooxml.WriteWorkbook(&buf, fixture); err != nil {
		t.Fatalf("WriteWorkbook: %v", err)
	}
	if err := os.WriteFile(path, buf.Bytes(), 0o600); err != nil {
		t.Fatal(err)
	}

	f, err := werkbook.Open(path)
	if err != nil {
		t.Fatal(err)
	}

	f.Recalculate()

	assertNum := func(sheet, cell string, want float64) {
		t.Helper()
		v, err := f.Sheet(sheet).GetValue(cell)
		if err != nil {
			t.Fatalf("GetValue(%s!%s): %v", sheet, cell, err)
		}
		if v.Type != werkbook.TypeNumber || v.Number != want {
			t.Fatalf("%s!%s = %#v, want %g", sheet, cell, v, want)
		}
	}

	// FILTER(Data!B2:B6, Data!A2:A6) → {10, 30, 50}
	assertNum(spillSheet, "B2", 10)
	assertNum(spillSheet, "B3", 30)
	assertNum(spillSheet, "B4", 50)
	// SUM and COUNT should reflect live values, not cached (999+888+777).
	assertNum(calcSheet, "A1", 90)
	assertNum(calcSheet, "B1", 3)
}

// TestSpillBlockedByOccupiedCell verifies that a user-entered value in the
// spill range blocks the spill at that position.
func TestSpillBlockedByOccupiedCell(t *testing.T) {
	const (
		dataSheet  = "Data"
		spillSheet = "Spill"
	)

	f := werkbook.New(werkbook.FirstSheet(dataSheet))
	ds := f.Sheet(dataSheet)
	for _, c := range []struct {
		cell string
		val  any
	}{
		{"A2", true}, {"B2", 10.0},
		{"A3", false}, {"B3", 20.0},
		{"A4", true}, {"B4", 30.0},
		{"A5", false}, {"B5", 40.0},
		{"A6", true}, {"B6", 50.0},
	} {
		if err := ds.SetValue(c.cell, c.val); err != nil {
			t.Fatal(err)
		}
	}

	ss, err := f.NewSheet(spillSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := ss.SetFormula("B2", "FILTER(Data!B2:B6,Data!A2:A6)"); err != nil {
		t.Fatal(err)
	}
	// Put a user value in B4, which is in the spill range.
	if err := ss.SetValue("B4", 12345.0); err != nil {
		t.Fatal(err)
	}
	f.Recalculate()

	// B4 keeps the user value — it blocks the spill.
	v, err := ss.GetValue("B4")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != werkbook.TypeNumber || v.Number != 12345 {
		t.Fatalf("B4 = %#v, want 12345 (user value should block spill)", v)
	}

	// B2 (the anchor) should show #SPILL! because B4 blocks the committed range.
	v2, _ := ss.GetValue("B2")
	if v2.Type != werkbook.TypeError || v2.String != "#SPILL!" {
		t.Fatalf("B2 = %#v, want #SPILL! error", v2)
	}

	// B3 should be empty (spill didn't happen).
	v3, _ := ss.GetValue("B3")
	if v3.Type != werkbook.TypeEmpty {
		t.Fatalf("B3 = %#v, want empty (spill blocked)", v3)
	}

	// Remove the blocking value — spill should work now.
	if err := ss.SetValue("B4", nil); err != nil {
		t.Fatal(err)
	}
	f.Recalculate()

	v2after, _ := ss.GetValue("B2")
	if v2after.Type != werkbook.TypeNumber || v2after.Number != 10 {
		t.Fatalf("B2 after unblock = %#v, want 10", v2after)
	}
	v3after, _ := ss.GetValue("B3")
	if v3after.Type != werkbook.TypeNumber || v3after.Number != 30 {
		t.Fatalf("B3 after unblock = %#v, want 30", v3after)
	}
	v4after, _ := ss.GetValue("B4")
	if v4after.Type != werkbook.TypeNumber || v4after.Number != 50 {
		t.Fatalf("B4 after unblock = %#v, want 50", v4after)
	}
}

func TestSpillBlockedByOccupiedCellWithoutRecalculate(t *testing.T) {
	const (
		dataSheet  = "Data"
		spillSheet = "Spill"
	)

	f := werkbook.New(werkbook.FirstSheet(dataSheet))
	ds := f.Sheet(dataSheet)
	for _, c := range []struct {
		cell string
		val  any
	}{
		{"A2", true}, {"B2", 10.0},
		{"A3", false}, {"B3", 20.0},
		{"A4", true}, {"B4", 30.0},
		{"A5", false}, {"B5", 40.0},
		{"A6", true}, {"B6", 50.0},
	} {
		if err := ds.SetValue(c.cell, c.val); err != nil {
			t.Fatal(err)
		}
	}

	ss, err := f.NewSheet(spillSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := ss.SetFormula("B2", "FILTER(Data!B2:B6,Data!A2:A6)"); err != nil {
		t.Fatal(err)
	}

	v, err := ss.GetValue("B3")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != werkbook.TypeNumber || v.Number != 30 {
		t.Fatalf("B3 before blocker = %#v, want 30", v)
	}

	if err := ss.SetValue("B4", 12345.0); err != nil {
		t.Fatal(err)
	}

	v2, err := ss.GetValue("B2")
	if err != nil {
		t.Fatal(err)
	}
	if v2.Type != werkbook.TypeError || v2.String != "#SPILL!" {
		t.Fatalf("B2 after blocker = %#v, want #SPILL!", v2)
	}
	v3, err := ss.GetValue("B3")
	if err != nil {
		t.Fatal(err)
	}
	if v3.Type != werkbook.TypeEmpty {
		t.Fatalf("B3 after blocker = %#v, want empty", v3)
	}

	if err := ss.SetValue("B4", nil); err != nil {
		t.Fatal(err)
	}

	v4, err := ss.GetValue("B3")
	if err != nil {
		t.Fatal(err)
	}
	if v4.Type != werkbook.TypeNumber || v4.Number != 30 {
		t.Fatalf("B3 after unblock = %#v, want 30", v4)
	}
	v5, err := ss.GetValue("B4")
	if err != nil {
		t.Fatal(err)
	}
	if v5.Type != werkbook.TypeNumber || v5.Number != 50 {
		t.Fatalf("B4 after unblock = %#v, want 50", v5)
	}
}

func TestSpillBlockedByFormulaCell(t *testing.T) {
	const (
		dataSheet  = "Data"
		spillSheet = "Spill"
	)

	f := werkbook.New(werkbook.FirstSheet(dataSheet))
	ds := f.Sheet(dataSheet)
	for _, c := range []struct {
		cell string
		val  any
	}{
		{"A2", true}, {"B2", 10.0},
		{"A3", false}, {"B3", 20.0},
		{"A4", true}, {"B4", 30.0},
		{"A5", false}, {"B5", 40.0},
		{"A6", true}, {"B6", 50.0},
	} {
		if err := ds.SetValue(c.cell, c.val); err != nil {
			t.Fatal(err)
		}
	}

	ss, err := f.NewSheet(spillSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := ss.SetFormula("B2", "FILTER(Data!B2:B6,Data!A2:A6)"); err != nil {
		t.Fatal(err)
	}
	if err := ss.SetFormula("B4", "12345"); err != nil {
		t.Fatal(err)
	}

	v, err := ss.GetValue("B2")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != werkbook.TypeError || v.String != "#SPILL!" {
		t.Fatalf("B2 = %#v, want #SPILL!", v)
	}
	v2, err := ss.GetValue("B3")
	if err != nil {
		t.Fatal(err)
	}
	if v2.Type != werkbook.TypeEmpty {
		t.Fatalf("B3 = %#v, want empty while spill blocked", v2)
	}
	v3, err := ss.GetValue("B4")
	if err != nil {
		t.Fatal(err)
	}
	if v3.Type != werkbook.TypeNumber || v3.Number != 12345 {
		t.Fatalf("B4 = %#v, want 12345 formula blocker", v3)
	}

	if err := ss.SetValue("B4", nil); err != nil {
		t.Fatal(err)
	}

	v4, err := ss.GetValue("B3")
	if err != nil {
		t.Fatal(err)
	}
	if v4.Type != werkbook.TypeNumber || v4.Number != 30 {
		t.Fatalf("B3 after clearing formula blocker = %#v, want 30", v4)
	}
}

// TestSpillConflict2DInteriorBlocker verifies #SPILL! when a cell in the
// interior of a 2D spill rectangle (SEQUENCE(3,3)) is occupied.
func TestSpillConflict2DInteriorBlocker(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// SEQUENCE(3,3) anchored at A1 should fill A1:C3 with 1..9.
	if err := s.SetFormula("A1", "SEQUENCE(3,3)"); err != nil {
		t.Fatal(err)
	}
	// Place a blocker at B2, which is in the interior of the 3×3 rectangle.
	if err := s.SetValue("B2", "blocker"); err != nil {
		t.Fatal(err)
	}
	f.Recalculate()

	// A1 (anchor) should show #SPILL!.
	v, err := s.GetValue("A1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != werkbook.TypeError || v.String != "#SPILL!" {
		t.Fatalf("A1 = %#v, want #SPILL! error", v)
	}

	// Other cells in the rectangle (except the blocker) should be empty.
	for _, cell := range []string{"A2", "A3", "C1", "C2", "C3"} {
		cv, _ := s.GetValue(cell)
		if cv.Type != werkbook.TypeEmpty {
			t.Fatalf("%s = %#v, want empty (spill blocked)", cell, cv)
		}
	}

	// B2 retains the user value.
	bv, _ := s.GetValue("B2")
	if bv.Type != werkbook.TypeString || bv.String != "blocker" {
		t.Fatalf("B2 = %#v, want \"blocker\"", bv)
	}

	// Remove blocker — spill should succeed.
	if err := s.SetValue("B2", nil); err != nil {
		t.Fatal(err)
	}
	f.Recalculate()

	// Verify the full 3×3 grid: SEQUENCE(3,3) produces 1..9 row-major.
	expected := []struct {
		cell string
		num  float64
	}{
		{"A1", 1}, {"B1", 2}, {"C1", 3},
		{"A2", 4}, {"B2", 5}, {"C2", 6},
		{"A3", 7}, {"B3", 8}, {"C3", 9},
	}
	for _, e := range expected {
		cv, _ := s.GetValue(e.cell)
		if cv.Type != werkbook.TypeNumber || cv.Number != e.num {
			t.Fatalf("%s = %#v, want %g", e.cell, cv, e.num)
		}
	}
}

// TestCrossSheetSUMOnSpillRange verifies that SUM/COUNT on another sheet's
// spill range work correctly through data changes.
func TestCrossSheetSUMOnSpillRange(t *testing.T) {
	const (
		dataSheet  = "Data"
		spillSheet = "Spill"
		calcSheet  = "Calc"
	)

	f := werkbook.New(werkbook.FirstSheet(dataSheet))
	ds := f.Sheet(dataSheet)
	for _, c := range []struct {
		cell string
		val  any
	}{
		{"A2", true}, {"B2", 100.0},
		{"A3", true}, {"B3", 200.0},
		{"A4", false}, {"B4", 300.0},
	} {
		if err := ds.SetValue(c.cell, c.val); err != nil {
			t.Fatal(err)
		}
	}

	ss, err := f.NewSheet(spillSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := ss.SetFormula("A2", "FILTER(Data!B2:B4,Data!A2:A4)"); err != nil {
		t.Fatal(err)
	}

	cs, err := f.NewSheet(calcSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := cs.SetFormula("A1", "SUM(Spill!A:A)"); err != nil {
		t.Fatal(err)
	}
	if err := cs.SetFormula("B1", "COUNT(Spill!A:A)"); err != nil {
		t.Fatal(err)
	}

	f.Recalculate()

	assertNum := func(sheet, cell string, want float64) {
		t.Helper()
		v, err := f.Sheet(sheet).GetValue(cell)
		if err != nil {
			t.Fatalf("GetValue(%s!%s): %v", sheet, cell, err)
		}
		if v.Type != werkbook.TypeNumber || v.Number != want {
			t.Fatalf("%s!%s = %#v, want %g", sheet, cell, v, want)
		}
	}

	// FILTER returns {100, 200}, SUM = 300, COUNT = 2
	assertNum(calcSheet, "A1", 300)
	assertNum(calcSheet, "B1", 2)

	// Include the third row too.
	if err := ds.SetValue("A4", true); err != nil {
		t.Fatal(err)
	}
	f.Recalculate()

	// Now FILTER returns {100, 200, 300}, SUM = 600, COUNT = 3
	assertNum(calcSheet, "A1", 600)
	assertNum(calcSheet, "B1", 3)

	// Shrink back to 1 row.
	if err := ds.SetValue("A3", false); err != nil {
		t.Fatal(err)
	}
	if err := ds.SetValue("A4", false); err != nil {
		t.Fatal(err)
	}
	f.Recalculate()

	// FILTER returns {100}, SUM = 100, COUNT = 1
	assertNum(calcSheet, "A1", 100)
	assertNum(calcSheet, "B1", 1)
}

func TestCrossSheetSUMOnSpillRangeWithoutRecalculate(t *testing.T) {
	const (
		dataSheet  = "Data"
		spillSheet = "Spill"
		calcSheet  = "Calc"
	)

	f := werkbook.New(werkbook.FirstSheet(dataSheet))
	ds := f.Sheet(dataSheet)
	for _, c := range []struct {
		cell string
		val  any
	}{
		{"A2", true}, {"B2", 100.0},
		{"A3", true}, {"B3", 200.0},
		{"A4", false}, {"B4", 300.0},
	} {
		if err := ds.SetValue(c.cell, c.val); err != nil {
			t.Fatal(err)
		}
	}

	ss, err := f.NewSheet(spillSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := ss.SetFormula("A2", "FILTER(Data!B2:B4,Data!A2:A4)"); err != nil {
		t.Fatal(err)
	}

	cs, err := f.NewSheet(calcSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := cs.SetFormula("A1", "SUM(Spill!A:A)"); err != nil {
		t.Fatal(err)
	}
	if err := cs.SetFormula("B1", "COUNT(Spill!A:A)"); err != nil {
		t.Fatal(err)
	}

	assertNum := func(sheet, cell string, want float64) {
		t.Helper()
		v, err := f.Sheet(sheet).GetValue(cell)
		if err != nil {
			t.Fatalf("GetValue(%s!%s): %v", sheet, cell, err)
		}
		if v.Type != werkbook.TypeNumber || v.Number != want {
			t.Fatalf("%s!%s = %#v, want %g", sheet, cell, v, want)
		}
	}

	assertNum(calcSheet, "A1", 300)
	assertNum(calcSheet, "B1", 2)

	if err := ds.SetValue("A4", true); err != nil {
		t.Fatal(err)
	}

	assertNum(calcSheet, "A1", 600)
	assertNum(calcSheet, "B1", 3)

	if err := ds.SetValue("A3", false); err != nil {
		t.Fatal(err)
	}
	if err := ds.SetValue("A4", false); err != nil {
		t.Fatal(err)
	}

	assertNum(calcSheet, "A1", 100)
	assertNum(calcSheet, "B1", 1)
}

// TestSetValueOnSpillAnchorClearsSpill verifies that overwriting a spill
// anchor cell with a plain value removes the spill entirely.
func TestSetValueOnSpillAnchorClearsSpill(t *testing.T) {
	const (
		dataSheet  = "Data"
		spillSheet = "Spill"
	)

	f := werkbook.New(werkbook.FirstSheet(dataSheet))
	ds := f.Sheet(dataSheet)
	for _, c := range []struct {
		cell string
		val  any
	}{
		{"A2", true}, {"B2", 10.0},
		{"A3", true}, {"B3", 20.0},
	} {
		if err := ds.SetValue(c.cell, c.val); err != nil {
			t.Fatal(err)
		}
	}

	ss, err := f.NewSheet(spillSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := ss.SetFormula("B2", "FILTER(Data!B2:B3,Data!A2:A3)"); err != nil {
		t.Fatal(err)
	}

	f.Recalculate()

	// Verify spill works.
	v, _ := ss.GetValue("B3")
	if v.Type != werkbook.TypeNumber || v.Number != 20 {
		t.Fatalf("B3 = %#v, want 20 (spill)", v)
	}

	// Overwrite anchor with a plain value.
	if err := ss.SetValue("B2", 42.0); err != nil {
		t.Fatal(err)
	}
	f.Recalculate()

	// B3 should now be empty (no more spill anchor).
	v3, _ := ss.GetValue("B3")
	if v3.Type != werkbook.TypeEmpty {
		t.Fatalf("B3 = %#v, want empty after anchor overwritten", v3)
	}
}

func TestSetFormulaOnSpillAnchorClearsSpillWithoutRecalculate(t *testing.T) {
	const (
		dataSheet  = "Data"
		spillSheet = "Spill"
	)

	f := werkbook.New(werkbook.FirstSheet(dataSheet))
	ds := f.Sheet(dataSheet)
	for _, c := range []struct {
		cell string
		val  any
	}{
		{"A2", true}, {"B2", 10.0},
		{"A3", true}, {"B3", 20.0},
	} {
		if err := ds.SetValue(c.cell, c.val); err != nil {
			t.Fatal(err)
		}
	}

	ss, err := f.NewSheet(spillSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := ss.SetFormula("B2", "FILTER(Data!B2:B3,Data!A2:A3)"); err != nil {
		t.Fatal(err)
	}

	v, err := ss.GetValue("B3")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != werkbook.TypeNumber || v.Number != 20 {
		t.Fatalf("B3 = %#v, want 20 before anchor replacement", v)
	}

	if err := ss.SetFormula("B2", "42"); err != nil {
		t.Fatal(err)
	}

	v2, err := ss.GetValue("B2")
	if err != nil {
		t.Fatal(err)
	}
	if v2.Type != werkbook.TypeNumber || v2.Number != 42 {
		t.Fatalf("B2 = %#v, want 42 after anchor replacement", v2)
	}
	v3, err := ss.GetValue("B3")
	if err != nil {
		t.Fatal(err)
	}
	if v3.Type != werkbook.TypeEmpty {
		t.Fatalf("B3 = %#v, want empty after replacing spill anchor with scalar formula", v3)
	}
}

// TestMultipleSpillAnchorsOnSameSheet verifies that two independent dynamic
// array formulas on the same sheet spill correctly without interfering.
func TestMultipleSpillAnchorsOnSameSheet(t *testing.T) {
	const dataSheet = "Data"
	const spillSheet = "Spill"

	f := werkbook.New(werkbook.FirstSheet(dataSheet))
	ds := f.Sheet(dataSheet)
	for _, c := range []struct {
		cell string
		val  any
	}{
		{"A2", true}, {"B2", 1.0}, {"C2", 100.0},
		{"A3", false}, {"B3", 2.0}, {"C3", 200.0},
		{"A4", true}, {"B4", 3.0}, {"C4", 300.0},
	} {
		if err := ds.SetValue(c.cell, c.val); err != nil {
			t.Fatal(err)
		}
	}

	ss, err := f.NewSheet(spillSheet)
	if err != nil {
		t.Fatal(err)
	}
	// Two independent FILTER formulas in different columns.
	if err := ss.SetFormula("A2", "FILTER(Data!B2:B4,Data!A2:A4)"); err != nil {
		t.Fatal(err)
	}
	if err := ss.SetFormula("C2", "FILTER(Data!C2:C4,Data!A2:A4)"); err != nil {
		t.Fatal(err)
	}

	f.Recalculate()

	assertNum := func(cell string, want float64) {
		t.Helper()
		v, err := ss.GetValue(cell)
		if err != nil {
			t.Fatalf("GetValue(%s): %v", cell, err)
		}
		if v.Type != werkbook.TypeNumber || v.Number != want {
			t.Fatalf("%s = %#v, want %g", cell, v, want)
		}
	}

	// Column A: FILTER → {1, 3}
	assertNum("A2", 1)
	assertNum("A3", 3)
	// Column C: FILTER → {100, 300}
	assertNum("C2", 100)
	assertNum("C3", 300)
}

// TestSpillWithBoundedRangeReference verifies that a bounded range reference
// like SUM(A2:A10) picks up spill values within that range.
func TestSpillWithBoundedRangeReference(t *testing.T) {
	const dataSheet = "Data"
	const spillSheet = "Spill"

	f := werkbook.New(werkbook.FirstSheet(dataSheet))
	ds := f.Sheet(dataSheet)
	for _, c := range []struct {
		cell string
		val  any
	}{
		{"A2", true}, {"B2", 5.0},
		{"A3", true}, {"B3", 15.0},
		{"A4", true}, {"B4", 25.0},
	} {
		if err := ds.SetValue(c.cell, c.val); err != nil {
			t.Fatal(err)
		}
	}

	ss, err := f.NewSheet(spillSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := ss.SetFormula("A2", "FILTER(Data!B2:B4,Data!A2:A4)"); err != nil {
		t.Fatal(err)
	}
	// Bounded range that overlaps the spill.
	if err := ss.SetFormula("C1", "SUM(A2:A10)"); err != nil {
		t.Fatal(err)
	}
	// Direct reference to a spill cell.
	if err := ss.SetFormula("D1", "A3"); err != nil {
		t.Fatal(err)
	}

	f.Recalculate()

	v, _ := ss.GetValue("C1")
	if v.Type != werkbook.TypeNumber || v.Number != 45 {
		t.Fatalf("SUM(A2:A10) = %#v, want 45", v)
	}

	v2, _ := ss.GetValue("D1")
	if v2.Type != werkbook.TypeNumber || v2.Number != 15 {
		t.Fatalf("=A3 = %#v, want 15 (spill value)", v2)
	}
}

// TestNestedDynamicArrayFunctions verifies that SORT(UNIQUE(FILTER(...)))
// spills the final composed result.
func TestNestedDynamicArrayFunctions(t *testing.T) {
	const dataSheet = "Data"
	const spillSheet = "Spill"

	f := werkbook.New(werkbook.FirstSheet(dataSheet))
	ds := f.Sheet(dataSheet)
	// Data with duplicates: include {30, 10, 30, 10, 20} → unique sorted = {10, 20, 30}
	for _, c := range []struct {
		cell string
		val  any
	}{
		{"A2", true}, {"B2", 30.0},
		{"A3", true}, {"B3", 10.0},
		{"A4", true}, {"B4", 30.0},
		{"A5", true}, {"B5", 10.0},
		{"A6", true}, {"B6", 20.0},
	} {
		if err := ds.SetValue(c.cell, c.val); err != nil {
			t.Fatal(err)
		}
	}

	ss, err := f.NewSheet(spillSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := ss.SetFormula("A2", "SORT(UNIQUE(FILTER(Data!B2:B6,Data!A2:A6)))"); err != nil {
		t.Fatal(err)
	}

	f.Recalculate()

	assertNum := func(cell string, want float64) {
		t.Helper()
		v, err := ss.GetValue(cell)
		if err != nil {
			t.Fatalf("GetValue(%s): %v", cell, err)
		}
		if v.Type != werkbook.TypeNumber || v.Number != want {
			t.Fatalf("%s = %#v, want %g", cell, v, want)
		}
	}
	assertNum("A2", 10)
	assertNum("A3", 20)
	assertNum("A4", 30)
}

// TestSheetCopyWithSpill verifies that cloning a sheet with a dynamic array
// formula produces an independent spill on the cloned sheet.
func TestSheetCopyWithSpill(t *testing.T) {
	const dataSheet = "Data"
	const spillSheet = "Spill"

	f := werkbook.New(werkbook.FirstSheet(dataSheet))
	ds := f.Sheet(dataSheet)
	for _, c := range []struct {
		cell string
		val  any
	}{
		{"A2", true}, {"B2", 7.0},
		{"A3", true}, {"B3", 14.0},
	} {
		if err := ds.SetValue(c.cell, c.val); err != nil {
			t.Fatal(err)
		}
	}

	ss, err := f.NewSheet(spillSheet)
	if err != nil {
		t.Fatal(err)
	}
	if err := ss.SetFormula("A2", "FILTER(Data!B2:B3,Data!A2:A3)"); err != nil {
		t.Fatal(err)
	}

	f.Recalculate()

	clone, err := f.CloneSheetFrom(ss, "SpillCopy")
	if err != nil {
		t.Fatal(err)
	}

	f.Recalculate()

	// Both sheets should have correct spill values.
	for _, s := range []*werkbook.Sheet{ss, clone} {
		v2, _ := s.GetValue("A2")
		if v2.Type != werkbook.TypeNumber || v2.Number != 7 {
			t.Fatalf("%s!A2 = %#v, want 7", s.Name(), v2)
		}
		v3, _ := s.GetValue("A3")
		if v3.Type != werkbook.TypeNumber || v3.Number != 14 {
			t.Fatalf("%s!A3 = %#v, want 14", s.Name(), v3)
		}
	}
}
