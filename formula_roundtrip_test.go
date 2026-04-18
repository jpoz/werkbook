package werkbook_test

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/jpoz/werkbook"
	"github.com/jpoz/werkbook/ooxml"
)

func TestSetFormulaGetFormula(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	if err := s.SetFormula("A1", "SUM(B1:B5)"); err != nil {
		t.Fatalf("SetFormula: %v", err)
	}

	got, err := s.GetFormula("A1")
	if err != nil {
		t.Fatalf("GetFormula: %v", err)
	}
	if got != "SUM(B1:B5)" {
		t.Errorf("GetFormula = %q, want %q", got, "SUM(B1:B5)")
	}

	// Cell with no formula returns "".
	got, err = s.GetFormula("B1")
	if err != nil {
		t.Fatalf("GetFormula B1: %v", err)
	}
	if got != "" {
		t.Errorf("GetFormula B1 = %q, want empty", got)
	}
}

func TestSetFormulaStripsLeadingEquals(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	if err := s.SetFormula("A1", "=SUM(1,2,3)"); err != nil {
		t.Fatalf("SetFormula: %v", err)
	}

	got, err := s.GetFormula("A1")
	if err != nil {
		t.Fatalf("GetFormula: %v", err)
	}
	if got != "SUM(1,2,3)" {
		t.Errorf("GetFormula = %q, want %q (leading '=' should be stripped)", got, "SUM(1,2,3)")
	}

	v, err := s.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue: %v", err)
	}
	if v.Type != werkbook.TypeNumber || v.Number != 6 {
		t.Errorf("GetValue = %+v, want TypeNumber 6 (formula should evaluate normally)", v)
	}
}

func TestFormulaRoundTrip(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetValue("A1", 10)
	s.SetFormula("B1", "A1*2")

	dir := t.TempDir()
	path := filepath.Join(dir, "formulas.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	s2 := f2.Sheet("Sheet1")

	got, err := s2.GetFormula("B1")
	if err != nil {
		t.Fatalf("GetFormula: %v", err)
	}
	if got != "A1*2" {
		t.Errorf("formula round-trip = %q, want %q", got, "A1*2")
	}
}

func TestFormulaWithValue(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetValue("A1", 10)
	s.SetValue("B1", 20) // cached value
	s.SetFormula("B1", "A1*2")

	dir := t.TempDir()
	path := filepath.Join(dir, "fv.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	s2 := f2.Sheet("Sheet1")

	formula, _ := s2.GetFormula("B1")
	if formula != "A1*2" {
		t.Errorf("formula = %q, want %q", formula, "A1*2")
	}

	val, _ := s2.GetValue("B1")
	if val.Type != werkbook.TypeNumber || val.Number != 20 {
		t.Errorf("cached value = %#v, want 20", val)
	}
}

func TestDynamicArraySpillCellsRoundTrip(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetFormula("B2", "SEQUENCE(2,3,10,5)")
	f.Recalculate()

	dir := t.TempDir()
	path := filepath.Join(dir, "dynamic-spill.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	s2 := f2.Sheet("Sheet1")

	tests := map[string]float64{
		"B2": 10,
		"C2": 15,
		"D2": 20,
		"B3": 25,
		"C3": 30,
		"D3": 35,
	}
	for cell, want := range tests {
		val, err := s2.GetValue(cell)
		if err != nil {
			t.Fatalf("GetValue(%s): %v", cell, err)
		}
		if val.Type != werkbook.TypeNumber || val.Number != want {
			t.Fatalf("%s = %#v, want %g", cell, val, want)
		}
	}
}

func TestFormulaCellInXML(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetFormula("A1", "1+1")

	dir := t.TempDir()
	path := filepath.Join(dir, "xmlcheck.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	sheetXML := readSheetXML(t, path, "xl/worksheets/sheet1.xml")

	type xmlC struct {
		R string `xml:"r,attr"`
		F string `xml:"f"`
		V string `xml:"v"`
	}
	type xmlRow struct {
		Cells []xmlC `xml:"c"`
	}
	type xmlSheet struct {
		Rows []xmlRow `xml:"sheetData>row"`
	}
	var ws xmlSheet
	if err := xml.Unmarshal(sheetXML, &ws); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	found := false
	for _, row := range ws.Rows {
		for _, c := range row.Cells {
			if c.R == "A1" && c.F == "1+1" {
				found = true
			}
		}
	}
	if !found {
		t.Errorf("expected <f>1+1</f> in A1, XML: %s", sheetXML)
	}
}

func TestDynamicArrayFormulaPrefixesInXML(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetFormula("A1", `SORT(UNIQUE(FILTER(B1:B10,B1:B10<>"")))`)

	dir := t.TempDir()
	path := filepath.Join(dir, "dynamic-array-prefixes.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	sheetXML := string(readSheetXML(t, path, "xl/worksheets/sheet1.xml"))
	want := `_xlfn._xlws.SORT(_xlfn.UNIQUE(_xlfn._xlws.FILTER(B1:B10,B1:B10&lt;&gt;&#34;&#34;)))`
	if !strings.Contains(sheetXML, want) {
		t.Fatalf("dynamic array formula XML missing expected prefixes\nwant: %s\nxml: %s", want, sheetXML)
	}
}

func TestOfficeEraFormulaPrefixesInXML(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	formulas := map[string]string{
		"A1":  "ACOT(0)",
		"A2":  "ACOTH(2)",
		"A3":  "BITAND(1,5)",
		"A4":  "BITLSHIFT(1,1)",
		"A5":  "BITOR(1,5)",
		"A6":  "BITRSHIFT(8,1)",
		"A7":  "BITXOR(1,5)",
		"A8":  "ERF.PRECISE(0.5)",
		"A9":  "ERFC.PRECISE(0.5)",
		"A10": "PDURATION(0.025,2000,2200)",
		"A11": "RRI(96,10000,11000)",
	}
	for cell, formula := range formulas {
		s.SetFormula(cell, formula)
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "office-era-prefixes.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	sheetXML := string(readSheetXML(t, path, "xl/worksheets/sheet1.xml"))
	for _, formula := range formulas {
		want := "<f>_xlfn." + formula + "</f>"
		if !strings.Contains(sheetXML, want) {
			t.Fatalf("formula XML missing expected prefix\nwant: %s\nxml: %s", want, sheetXML)
		}
	}
}

// These tests lock down the exact OOXML fragments that previously triggered
// Excel's repair dialog. They inspect the generated package directly so the
// regression suite stays pure Go and does not need to automate Excel.
func TestFutureFunctionWorkbookXMLIncludesCalcFeatures(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetFormula("A1", "ACOT(1)")

	dir := t.TempDir()
	path := filepath.Join(dir, "future-features.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	workbookXML := string(readSheetXML(t, path, "xl/workbook.xml"))
	if !strings.Contains(workbookXML, `calcId="181029"`) {
		t.Fatalf("expected future-function workbook to emit calcPr calcId, XML: %s", workbookXML)
	}
	if !strings.Contains(workbookXML, `xcalcf:calcFeatures`) {
		t.Fatalf("expected future-function workbook to emit calcFeatures metadata, XML: %s", workbookXML)
	}
	if !strings.Contains(workbookXML, `microsoft.com:LET_WF`) {
		t.Fatalf("expected future-function workbook to emit Excel calc feature bundle, XML: %s", workbookXML)
	}
}

func TestFutureFunctionWorkbookXMLChildOrder(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetValue("A1", 1)
	s.SetFormula("B1", "ACOT(A1)")
	if err := f.SetDefinedName(werkbook.DefinedName{
		Name:         "InputCell",
		Value:        "Sheet1!$A$1",
		LocalSheetID: -1,
	}); err != nil {
		t.Fatalf("SetDefinedName: %v", err)
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "future-order.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	workbookXML := string(readSheetXML(t, path, "xl/workbook.xml"))
	got := directChildElementNames(t, []byte(workbookXML))
	want := []string{"sheets", "definedNames", "calcPr", "extLst"}
	if len(got) < len(want) {
		t.Fatalf("workbook child elements = %v, want at least %v", got, want)
	}
	for i, name := range want {
		if got[i] != name {
			t.Fatalf("workbook child elements = %v, want prefix %v", got, want)
		}
	}
}

func TestLegacyWorkbookOmitsCalcFeaturesMetadata(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetFormula("A1", "SUM(B1:B5)")

	dir := t.TempDir()
	path := filepath.Join(dir, "legacy-features.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	workbookXML := string(readSheetXML(t, path, "xl/workbook.xml"))
	if strings.Contains(workbookXML, `xcalcf:calcFeatures`) {
		t.Fatalf("legacy workbook should not emit calcFeatures metadata, XML: %s", workbookXML)
	}
}

func TestLETFormulaSerializationUsesXlpmPrefixes(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetFormula("A1", "LET(x,5,x+1)")

	dir := t.TempDir()
	path := filepath.Join(dir, "let-prefixes.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	sheetXML := string(readSheetXML(t, path, "xl/worksheets/sheet1.xml"))
	want := `<f>_xlfn.LET(_xlpm.x,5,_xlpm.x+1)</f>`
	if !strings.Contains(sheetXML, want) {
		t.Fatalf("LET formula XML missing expected _xlpm prefixes\nwant: %s\nxml: %s", want, sheetXML)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	s2 := f2.Sheet("Sheet1")

	got, err := s2.GetFormula("A1")
	if err != nil {
		t.Fatalf("GetFormula: %v", err)
	}
	if got != "LET(x,5,x+1)" {
		t.Fatalf("formula round-trip = %q, want %q", got, "LET(x,5,x+1)")
	}

	val, err := s2.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue: %v", err)
	}
	if val.Type != werkbook.TypeNumber || val.Number != 6 {
		t.Fatalf("LET cached value = %#v, want 6", val)
	}
}

func TestLambdaFamilyFormulasUseXlpmPrefixesInXML(t *testing.T) {
	tests := []struct {
		name       string
		formula    string
		wantXML    string
		setupCells map[string]any
	}{
		{
			name:    "lambda",
			formula: "LAMBDA(x,x+1)",
			wantXML: `_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)`,
		},
		{
			name:    "map lambda",
			formula: "MAP(A1:A3,LAMBDA(x,x+1))",
			wantXML: `_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1))`,
			setupCells: map[string]any{
				"A1": 1,
				"A2": 2,
				"A3": 3,
			},
		},
		{
			name:    "byrow lambda",
			formula: "BYROW(A1:B2,LAMBDA(r,SUM(r)))",
			wantXML: `_xlfn.BYROW(A1:B2,_xlfn.LAMBDA(_xlpm.r,SUM(_xlpm.r)))`,
			setupCells: map[string]any{
				"A1": 1,
				"B1": 2,
				"A2": 3,
				"B2": 4,
			},
		},
		{
			name:    "bycol lambda",
			formula: "BYCOL(A1:B2,LAMBDA(c,SUM(c)))",
			wantXML: `_xlfn.BYCOL(A1:B2,_xlfn.LAMBDA(_xlpm.c,SUM(_xlpm.c)))`,
			setupCells: map[string]any{
				"A1": 1,
				"B1": 2,
				"A2": 3,
				"B2": 4,
			},
		},
		{
			name:    "makearray lambda",
			formula: "MAKEARRAY(2,2,LAMBDA(r,c,r+c))",
			wantXML: `_xlfn.MAKEARRAY(2,2,_xlfn.LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r+_xlpm.c))`,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			assertFormulaXMLRoundTrip(t, tt.formula, tt.wantXML, tt.setupCells)
		})
	}
}

func TestOfficeEraFormulaPrefixesRestoredOnResave(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	formulas := map[string]string{
		"A1":  "ACOT(0)",
		"A2":  "ACOTH(2)",
		"A3":  "BITAND(1,5)",
		"A4":  "BITLSHIFT(1,1)",
		"A5":  "BITOR(1,5)",
		"A6":  "BITRSHIFT(8,1)",
		"A7":  "BITXOR(1,5)",
		"A8":  "ERF.PRECISE(0.5)",
		"A9":  "ERFC.PRECISE(0.5)",
		"A10": "PDURATION(0.025,2000,2200)",
		"A11": "RRI(96,10000,11000)",
	}
	for cell, formula := range formulas {
		s.SetFormula(cell, formula)
	}

	dir := t.TempDir()
	prefixedPath := filepath.Join(dir, "prefixed.xlsx")
	if err := f.SaveAs(prefixedPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	legacyPath := filepath.Join(dir, "legacy.xlsx")
	rewriteZipEntry(t, prefixedPath, legacyPath, "xl/worksheets/sheet1.xml", func(data []byte) []byte {
		xml := string(data)
		xml = strings.ReplaceAll(xml, "_xlfn._xlws.", "")
		xml = strings.ReplaceAll(xml, "_xlfn.", "")
		return []byte(xml)
	})

	legacyXML := string(readSheetXML(t, legacyPath, "xl/worksheets/sheet1.xml"))
	if strings.Contains(legacyXML, "_xlfn.") {
		t.Fatalf("legacy test fixture still contains _xlfn prefix: %s", legacyXML)
	}

	f2, err := werkbook.Open(legacyPath)
	if err != nil {
		t.Fatalf("Open legacy workbook: %v", err)
	}
	s2 := f2.Sheet("Sheet1")
	for cell, want := range formulas {
		got, err := s2.GetFormula(cell)
		if err != nil {
			t.Fatalf("GetFormula(%s): %v", cell, err)
		}
		if got != want {
			t.Fatalf("GetFormula(%s) = %q, want %q", cell, got, want)
		}
	}

	resavedPath := filepath.Join(dir, "resaved.xlsx")
	if err := f2.SaveAs(resavedPath); err != nil {
		t.Fatalf("SaveAs resaved workbook: %v", err)
	}

	resavedXML := string(readSheetXML(t, resavedPath, "xl/worksheets/sheet1.xml"))
	for _, formula := range formulas {
		want := "<f>_xlfn." + formula + "</f>"
		if !strings.Contains(resavedXML, want) {
			t.Fatalf("resaved formula XML missing expected prefix\nwant: %s\nxml: %s", want, resavedXML)
		}
	}
}

func TestDynamicArrayFormulaSerializationIncludesMetadataXML(t *testing.T) {
	f := werkbook.New(werkbook.FirstSheet("Out - Ledger Summary"))
	s := f.Sheet("Out - Ledger Summary")
	if _, err := f.NewSheet("treasury-ledger"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	s.SetFormula("A2", `SORT(UNIQUE(FILTER('treasury-ledger'!E2:E1000,'treasury-ledger'!E2:E1000<>"")))`)
	s.SetFormula("B2", `COUNTIFS('treasury-ledger'!E:E,ANCHORARRAY(A2))`)
	s.SetFormula("C2", `SUMIFS('treasury-ledger'!F:F,'treasury-ledger'!E:E,ANCHORARRAY(A2))/100`)
	s.SetFormula("D2", `SINGLE(IF(SUMIFS('treasury-ledger'!F:F,'treasury-ledger'!E:E,ANCHORARRAY(A2))>0,"Inflow","Outflow"))`)

	val, _ := s.GetValue("A2")
	if val.Type != werkbook.TypeError || val.String != "#CALC!" {
		t.Fatalf("A2 cached value = %#v, want #CALC!", val)
	}
	val, _ = s.GetValue("D2")
	if val.Type != werkbook.TypeError || val.String != "#NAME?" {
		t.Fatalf("D2 cached value = %#v, want #NAME?", val)
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "dynamic-array-metadata.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	sheetXML := string(readSheetXML(t, path, "xl/worksheets/sheet1.xml"))
	if !strings.Contains(sheetXML, `r="A2"`) || !strings.Contains(sheetXML, `cm="1"`) || !strings.Contains(sheetXML, `ref="A2" ca="1">_xlfn._xlws.SORT(_xlfn.UNIQUE(_xlfn._xlws.FILTER(`) {
		t.Fatalf("expected A2 to be written with dynamic-array metadata, XML: %s", sheetXML)
	}
	if !strings.Contains(sheetXML, `r="D2" t="e" cm="1"`) || !strings.Contains(sheetXML, `ref="D2" ca="1">_xlfn.SINGLE(`) {
		t.Fatalf("expected D2 SINGLE formula to carry dynamic-array metadata, XML: %s", sheetXML)
	}
	if !strings.Contains(sheetXML, `cm="1"`) || !strings.Contains(sheetXML, `<f t="array"`) {
		t.Fatalf("expected dynamic-array metadata in sheet XML: %s", sheetXML)
	}

	workbookRels := string(readSheetXML(t, path, "xl/_rels/workbook.xml.rels"))
	if !strings.Contains(workbookRels, `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"`) {
		t.Fatalf("expected workbook relationships to include sheet metadata, XML: %s", workbookRels)
	}

	contentTypes := string(readSheetXML(t, path, "[Content_Types].xml"))
	if !strings.Contains(contentTypes, `PartName="/xl/metadata.xml"`) {
		t.Fatalf("expected content types to include metadata.xml, XML: %s", contentTypes)
	}
	if !zipHasEntry(t, path, "xl/metadata.xml") {
		t.Fatal("metadata.xml should be written for dynamic array formulas")
	}

	r, err := os.Open(path)
	if err != nil {
		t.Fatalf("Open temp xlsx: %v", err)
	}
	defer r.Close()
	info, err := r.Stat()
	if err != nil {
		t.Fatalf("Stat temp xlsx: %v", err)
	}
	data, err := ooxml.ReadWorkbook(r, info.Size())
	if err != nil {
		t.Fatalf("ReadWorkbook: %v", err)
	}
	foundA2 := false
	for _, sd := range data.Sheets {
		if sd.Name != "Out - Ledger Summary" {
			continue
		}
		for _, rd := range sd.Rows {
			for _, cd := range rd.Cells {
				if cd.Ref != "A2" {
					continue
				}
				foundA2 = true
				if !cd.IsDynamicArray {
					t.Fatalf("expected ooxml reader to mark A2 as dynamic array: %#v", cd)
				}
			}
		}
	}
	if !foundA2 {
		t.Fatal("A2 not found in parsed workbook data")
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	s2 := f2.Sheet("Out - Ledger Summary")
	val, _ = s2.GetValue("A2")
	if val.Type != werkbook.TypeString || val.String != "#CALC!" {
		t.Fatalf("A2 round-trip value = %#v, want string #CALC!", val)
	}
	got, err := s2.GetFormula("D2")
	if err != nil {
		t.Fatalf("GetFormula D2: %v", err)
	}
	wantFormula := `SINGLE(IF(SUMIFS('treasury-ledger'!F:F,'treasury-ledger'!E:E,ANCHORARRAY(A2))>0,"Inflow","Outflow"))`
	if got != wantFormula {
		t.Fatalf("formula round-trip = %q, want %q", got, wantFormula)
	}
}

func TestOpenSavePreservesImportedDynamicArrayMetadata(t *testing.T) {
	const (
		sheetName    = "Out - Fee Accruals"
		sourceSheet  = "venture-fee-accrual-line-items"
		anchorCell   = "A2"
		spillRef     = "A2:A20"
		userFormula  = `FILTER('venture-fee-accrual-line-items'!E2:E41,DATEVALUE('venture-fee-accrual-line-items'!E2:E41)>=TODAY(),"No upcoming fees")`
		ooxmlFormula = `_xlfn._xlws.FILTER('venture-fee-accrual-line-items'!E2:E41,DATEVALUE('venture-fee-accrual-line-items'!E2:E41)>=TODAY(),"No upcoming fees")`
	)

	fixture := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{
			{
				Name: sheetName,
				Rows: []ooxml.RowData{{Num: 2, Cells: []ooxml.CellData{{
					Ref:            anchorCell,
					Type:           "str",
					Value:          "2026-04-01",
					Formula:        ooxmlFormula,
					FormulaType:    "array",
					FormulaRef:     spillRef,
					IsDynamicArray: true,
				}}}},
			},
			{
				Name: sourceSheet,
				Rows: []ooxml.RowData{
					{Num: 1, Cells: []ooxml.CellData{{Ref: "E1", Type: "s", Value: "fee_date"}}},
					{Num: 2, Cells: []ooxml.CellData{{Ref: "E2", Type: "s", Value: "2026-04-01"}}},
				},
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

	srcSheetXML := string(readSheetXML(t, srcPath, "xl/worksheets/sheet1.xml"))
	for _, want := range []string{
		`cm="1"`,
		`<f t="array"`,
		`aca="1"`,
		`ref="A2:A20"`,
		`ca="1">_xlfn._xlws.FILTER(`,
	} {
		if !strings.Contains(srcSheetXML, want) {
			t.Fatalf("fixture workbook missing dynamic-array metadata %q\nxml: %s", want, srcSheetXML)
		}
	}
	if !zipHasEntry(t, srcPath, "xl/metadata.xml") {
		t.Fatal("fixture workbook missing xl/metadata.xml")
	}

	f, err := werkbook.Open(srcPath)
	if err != nil {
		t.Fatalf("Open fixture: %v", err)
	}
	got, err := f.Sheet(sheetName).GetFormula(anchorCell)
	if err != nil {
		t.Fatalf("GetFormula(%s): %v", anchorCell, err)
	}
	if got != userFormula {
		t.Fatalf("formula round-trip = %q, want %q", got, userFormula)
	}

	dstPath := filepath.Join(dir, "dynamic-array-roundtrip.xlsx")
	if err := f.SaveAs(dstPath); err != nil {
		t.Fatalf("SaveAs round-trip workbook: %v", err)
	}

	dstSheetXML := string(readSheetXML(t, dstPath, "xl/worksheets/sheet1.xml"))
	for _, want := range []string{
		`cm="1"`,
		`<f t="array"`,
		`aca="1"`,
		`ref="A2:A20"`,
		`ca="1">_xlfn._xlws.FILTER(`,
	} {
		if !strings.Contains(dstSheetXML, want) {
			t.Fatalf("round-trip workbook missing dynamic-array metadata %q\nxml: %s", want, dstSheetXML)
		}
	}

	workbookRels := string(readSheetXML(t, dstPath, "xl/_rels/workbook.xml.rels"))
	if !strings.Contains(workbookRels, `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"`) {
		t.Fatalf("expected workbook relationships to preserve sheet metadata, XML: %s", workbookRels)
	}

	contentTypes := string(readSheetXML(t, dstPath, "[Content_Types].xml"))
	if !strings.Contains(contentTypes, `PartName="/xl/metadata.xml"`) {
		t.Fatalf("expected content types to preserve metadata.xml, XML: %s", contentTypes)
	}
	if !zipHasEntry(t, dstPath, "xl/metadata.xml") {
		t.Fatal("round-trip workbook missing xl/metadata.xml")
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

	foundAnchor := false
	for _, sd := range data.Sheets {
		if sd.Name != sheetName {
			continue
		}
		for _, rd := range sd.Rows {
			for _, cd := range rd.Cells {
				if cd.Ref != anchorCell {
					continue
				}
				foundAnchor = true
				if !cd.IsDynamicArray {
					t.Fatalf("expected %s to remain a dynamic array anchor: %#v", anchorCell, cd)
				}
				if cd.FormulaRef != spillRef {
					t.Fatalf("FormulaRef = %q, want %q", cd.FormulaRef, spillRef)
				}
			}
		}
	}
	if !foundAnchor {
		t.Fatalf("%s not found in parsed round-trip workbook data", anchorCell)
	}
}

func TestOpenSaveRecomputesImportedDynamicArraySpillRefAfterRecalc(t *testing.T) {
	const (
		dataSheet    = "Data"
		spillSheet   = "Spill"
		spillAnchor  = "B2"
		originalRef  = "B2:B4"
		updatedRef   = "B2:B3"
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
						{Ref: "A3", Type: "b", Value: "1"},
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
						{Ref: "A6", Type: "b", Value: "0"},
						{Ref: "B6", Value: "50"},
					}},
				},
			},
			{
				Name: spillSheet,
				Rows: []ooxml.RowData{
					{Num: 2, Cells: []ooxml.CellData{{
						Ref:            spillAnchor,
						Type:           "str",
						Value:          "10",
						Formula:        ooxmlFormula,
						FormulaType:    "array",
						FormulaRef:     originalRef,
						IsDynamicArray: true,
					}}},
					{Num: 3, Cells: []ooxml.CellData{{Ref: "B3"}}},
					{Num: 4, Cells: []ooxml.CellData{{Ref: "B4"}}},
				},
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

	f, err := werkbook.Open(srcPath)
	if err != nil {
		t.Fatalf("Open fixture: %v", err)
	}
	if got, err := f.Sheet(spillSheet).GetFormula(spillAnchor); err != nil {
		t.Fatalf("GetFormula(%s): %v", spillAnchor, err)
	} else if got != userFormula {
		t.Fatalf("formula round-trip = %q, want %q", got, userFormula)
	}

	if err := f.Sheet(dataSheet).SetValue("A4", false); err != nil {
		t.Fatalf("SetValue(A4): %v", err)
	}

	if val, err := f.Sheet(spillSheet).GetValue(spillAnchor); err != nil {
		t.Fatalf("GetValue(%s): %v", spillAnchor, err)
	} else if val.Type != werkbook.TypeNumber || val.Number != 10 {
		t.Fatalf("%s value = %#v, want 10", spillAnchor, val)
	}
	if val, err := f.Sheet(spillSheet).GetValue("B3"); err != nil {
		t.Fatalf("GetValue(B3): %v", err)
	} else if val.Type != werkbook.TypeNumber || val.Number != 20 {
		t.Fatalf("B3 value = %#v, want 20 after spill shrink", val)
	}
	if val, err := f.Sheet(spillSheet).GetValue("B4"); err != nil {
		t.Fatalf("GetValue(B4): %v", err)
	} else if val.Type != werkbook.TypeEmpty {
		t.Fatalf("B4 value = %#v, want empty after spill shrink", val)
	}

	dstPath := filepath.Join(dir, "dynamic-array-roundtrip.xlsx")
	if err := f.SaveAs(dstPath); err != nil {
		t.Fatalf("SaveAs round-trip workbook: %v", err)
	}

	dstSheetXML := string(readSheetXML(t, dstPath, "xl/worksheets/sheet2.xml"))
	if strings.Contains(dstSheetXML, `ref="`+originalRef+`"`) {
		t.Fatalf("saved worksheet still contains stale spill ref %q\nxml: %s", originalRef, dstSheetXML)
	}
	if !strings.Contains(dstSheetXML, `ref="`+updatedRef+`"`) {
		t.Fatalf("saved worksheet missing updated spill ref %q\nxml: %s", updatedRef, dstSheetXML)
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

	foundAnchor := false
	for _, sd := range data.Sheets {
		if sd.Name != spillSheet {
			continue
		}
		for _, rd := range sd.Rows {
			for _, cd := range rd.Cells {
				if cd.Ref != spillAnchor {
					continue
				}
				foundAnchor = true
				if !cd.IsDynamicArray {
					t.Fatalf("expected %s to remain a dynamic array anchor: %#v", spillAnchor, cd)
				}
				if cd.FormulaRef != updatedRef {
					t.Fatalf("FormulaRef = %q, want %q", cd.FormulaRef, updatedRef)
				}
			}
		}
	}
	if !foundAnchor {
		t.Fatalf("%s not found in parsed round-trip workbook data", spillAnchor)
	}
}

func TestFormulaAndValueCoexist(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set up inputs so the formula evaluates to 42.
	s.SetValue("B1", 21)
	s.SetValue("C1", 21)
	s.SetFormula("A1", "B1+C1")

	// Evaluate so the cached value is saved to the file.
	val, _ := s.GetValue("A1")
	if val.Type != werkbook.TypeNumber || val.Number != 42 {
		t.Fatalf("pre-save value = %#v, want 42", val)
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "coexist.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	s2 := f2.Sheet("Sheet1")

	val, _ = s2.GetValue("A1")
	if val.Type != werkbook.TypeNumber || val.Number != 42 {
		t.Errorf("value = %#v, want 42", val)
	}
	formula, _ := s2.GetFormula("A1")
	if formula != "B1+C1" {
		t.Errorf("formula = %q, want %q", formula, "B1+C1")
	}
}

func TestEmptyFormulaIgnored(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetValue("A1", "hello")

	dir := t.TempDir()
	path := filepath.Join(dir, "nof.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	data := readSheetXML(t, path, "xl/worksheets/sheet1.xml")
	content := string(data)
	if strings.Contains(content, "<f>") || strings.Contains(content, "</f>") {
		t.Errorf("unexpected <f> element in XML for cell without formula: %s", content)
	}
}

// assertFormulaXMLRoundTrip verifies that the serialized worksheet XML
// contains the expected formula fragment and then reopens the workbook to
// ensure StripXlfnPrefixes restores the user-facing formula text. This keeps
// the regression anchored to raw OOXML without calling external spreadsheet
// apps.
func assertFormulaXMLRoundTrip(t *testing.T, userFormula, wantXML string, setupCells map[string]any) {
	t.Helper()

	f := werkbook.New()
	s := f.Sheet("Sheet1")
	for cell, value := range setupCells {
		if err := s.SetValue(cell, value); err != nil {
			t.Fatalf("SetValue(%s): %v", cell, err)
		}
	}
	if err := s.SetFormula("H1", userFormula); err != nil {
		t.Fatalf("SetFormula(H1=%q): %v", userFormula, err)
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "formula-roundtrip.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	sheetXML := string(readSheetXML(t, path, "xl/worksheets/sheet1.xml"))
	if !strings.Contains(sheetXML, wantXML) {
		t.Fatalf("formula XML missing expected fragment\nwant: %s\nxml: %s", wantXML, sheetXML)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	got, err := f2.Sheet("Sheet1").GetFormula("H1")
	if err != nil {
		t.Fatalf("GetFormula: %v", err)
	}
	if got != userFormula {
		t.Fatalf("formula round-trip = %q, want %q", got, userFormula)
	}
}

func directChildElementNames(t *testing.T, doc []byte) []string {
	t.Helper()

	dec := xml.NewDecoder(strings.NewReader(string(doc)))
	var (
		depth int
		names []string
	)
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return names
		}
		if err != nil {
			t.Fatalf("decode xml: %v", err)
		}
		switch tok := tok.(type) {
		case xml.StartElement:
			if depth == 1 {
				names = append(names, tok.Name.Local)
			}
			depth++
		case xml.EndElement:
			depth--
		}
	}
}

// readSheetXML extracts a file from the XLSX zip archive.
func readSheetXML(t *testing.T, xlsxPath, entryName string) []byte {
	t.Helper()
	zr, err := zip.OpenReader(xlsxPath)
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}
	defer zr.Close()

	for _, zf := range zr.File {
		if zf.Name == entryName {
			rc, err := zf.Open()
			if err != nil {
				t.Fatalf("open %s: %v", entryName, err)
			}
			defer rc.Close()
			data, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("read %s: %v", entryName, err)
			}
			return data
		}
	}
	t.Fatalf("%s not found in zip", entryName)
	return nil
}

func zipHasEntry(t *testing.T, xlsxPath, entryName string) bool {
	t.Helper()
	zr, err := zip.OpenReader(xlsxPath)
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}
	defer zr.Close()

	for _, zf := range zr.File {
		if zf.Name == entryName {
			return true
		}
	}
	return false
}

func rewriteZipEntry(t *testing.T, srcPath, dstPath, entryName string, rewrite func([]byte) []byte) {
	t.Helper()

	zr, err := zip.OpenReader(srcPath)
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}
	defer zr.Close()

	out, err := os.Create(dstPath)
	if err != nil {
		t.Fatalf("create output zip: %v", err)
	}
	defer out.Close()

	zw := zip.NewWriter(out)
	defer zw.Close()

	for _, zf := range zr.File {
		rc, err := zf.Open()
		if err != nil {
			t.Fatalf("open %s: %v", zf.Name, err)
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatalf("read %s: %v", zf.Name, err)
		}
		if zf.Name == entryName {
			data = rewrite(data)
		}

		hdr := zf.FileHeader
		w, err := zw.CreateHeader(&hdr)
		if err != nil {
			t.Fatalf("create %s: %v", zf.Name, err)
		}
		if _, err := w.Write(data); err != nil {
			t.Fatalf("write %s: %v", zf.Name, err)
		}
	}

	if err := zw.Close(); err != nil {
		t.Fatalf("close zip writer: %v", err)
	}
}
