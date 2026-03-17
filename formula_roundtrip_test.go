package werkbook_test

import (
	"archive/zip"
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
	want := `<f>_xlfn._xlws.SORT(_xlfn.UNIQUE(_xlfn._xlws.FILTER(B1:B10,B1:B10&lt;&gt;&#34;&#34;)))</f>`
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
			wantXML: `<f>_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)</f>`,
		},
		{
			name:    "map lambda",
			formula: "MAP(A1:A3,LAMBDA(x,x+1))",
			wantXML: `<f>_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1))</f>`,
			setupCells: map[string]any{
				"A1": 1,
				"A2": 2,
				"A3": 3,
			},
		},
		{
			name:    "byrow lambda",
			formula: "BYROW(A1:B2,LAMBDA(r,SUM(r)))",
			wantXML: `<f>_xlfn.BYROW(A1:B2,_xlfn.LAMBDA(_xlpm.r,SUM(_xlpm.r)))</f>`,
			setupCells: map[string]any{
				"A1": 1,
				"B1": 2,
				"A2": 3,
				"B2": 4,
			},
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

func TestDynamicArrayFormulaSerializationAvoidsMetadataXML(t *testing.T) {
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
	if !strings.Contains(sheetXML, `<c r="A2" t="str"><f>_xlfn._xlws.SORT(_xlfn.UNIQUE(_xlfn._xlws.FILTER(`) {
		t.Fatalf("expected A2 to be written as a plain formula string cell, XML: %s", sheetXML)
	}
	if strings.Contains(sheetXML, `cm="1"`) {
		t.Fatalf("expected dynamic-array cells to omit cm metadata, XML: %s", sheetXML)
	}
	if strings.Contains(sheetXML, `<f t="array"`) {
		t.Fatalf("expected dynamic-array cells to omit legacy array markup, XML: %s", sheetXML)
	}
	if !strings.Contains(sheetXML, `<c r="D2" t="e"><f>_xlfn.SINGLE(`) {
		t.Fatalf("expected D2 SINGLE formula to use _xlfn prefix and normal error typing, XML: %s", sheetXML)
	}

	workbookRels := string(readSheetXML(t, path, "xl/_rels/workbook.xml.rels"))
	if strings.Contains(workbookRels, `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"`) {
		t.Fatalf("expected workbook relationships to omit sheet metadata, XML: %s", workbookRels)
	}

	contentTypes := string(readSheetXML(t, path, "[Content_Types].xml"))
	if strings.Contains(contentTypes, `PartName="/xl/metadata.xml"`) {
		t.Fatalf("expected content types to omit metadata.xml, XML: %s", contentTypes)
	}
	if zipHasEntry(t, path, "xl/metadata.xml") {
		t.Fatal("metadata.xml should not be written for dynamic array formulas")
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

// assertFormulaXMLRoundTrip verifies the exact serialized <f> fragment and
// then reopens the workbook to ensure StripXlfnPrefixes restores the user-
// facing formula text. This keeps the regression anchored to raw OOXML
// without calling external spreadsheet apps.
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
