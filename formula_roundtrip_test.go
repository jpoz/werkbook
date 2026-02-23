package werkbook_test

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"path/filepath"
	"strings"
	"testing"

	"github.com/jpoz/werkbook"
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
