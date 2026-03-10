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

// TestNumericValuesNoScientificNotation verifies that large and small numbers
// are written in plain decimal notation (not scientific notation like "1e+06")
// in the XLSX XML, which OOXML-compliant readers require.
func TestNumericValuesNoScientificNotation(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	values := map[string]float64{
		"A1": 1000000,
		"A2": -2500000,
		"A3": 100000000,
		"A4": 28614898,
		"A5": 0.001,
		"A6": 123456789.123,
		"A7": -1999830,
		"A8": 0,
		"A9": 1e15,
	}
	for ref, val := range values {
		if err := s.SetValue(ref, val); err != nil {
			t.Fatalf("SetValue(%q, %v): %v", ref, val, err)
		}
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "numeric.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	// Read the raw sheet XML from the zip.
	zr, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}
	defer zr.Close()

	var sheetXML string
	for _, zf := range zr.File {
		if zf.Name == "xl/worksheets/sheet1.xml" {
			rc, err := zf.Open()
			if err != nil {
				t.Fatalf("open sheet1: %v", err)
			}
			b, _ := io.ReadAll(rc)
			rc.Close()
			sheetXML = string(b)
		}
	}
	if sheetXML == "" {
		t.Fatal("sheet1.xml not found in zip")
	}

	// Parse cells from the XML.
	type xmlC struct {
		R string `xml:"r,attr"`
		T string `xml:"t,attr"`
		V string `xml:"v"`
	}
	type xmlRow struct {
		C []xmlC `xml:"c"`
	}
	type xmlSheet struct {
		SheetData struct {
			Row []xmlRow `xml:"row"`
		} `xml:"sheetData"`
	}
	var ws xmlSheet
	if err := xml.Unmarshal([]byte(sheetXML), &ws); err != nil {
		t.Fatalf("parse sheet XML: %v", err)
	}

	cells := make(map[string]string)
	for _, r := range ws.SheetData.Row {
		for _, c := range r.C {
			cells[c.R] = c.V
		}
	}

	// Verify no cell value contains scientific notation characters.
	for ref, v := range cells {
		if strings.ContainsAny(v, "eE") {
			t.Errorf("cell %s value %q uses scientific notation", ref, v)
		}
	}

	// Verify specific expected values.
	expected := map[string]string{
		"A1": "1000000",
		"A2": "-2500000",
		"A3": "100000000",
		"A4": "28614898",
		"A5": "0.001",
		"A7": "-1999830",
		"A8": "0",
		"A9": "1000000000000000",
	}
	for ref, want := range expected {
		got, ok := cells[ref]
		if !ok {
			t.Errorf("cell %s not found in XML", ref)
			continue
		}
		if got != want {
			t.Errorf("cell %s = %q, want %q", ref, got, want)
		}
	}
}

// TestLargeNumberRoundTrip verifies that large numbers survive a write+read
// cycle without loss of precision.
func TestLargeNumberRoundTrip(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	values := map[string]float64{
		"A1": 1000000,
		"A2": -2500000,
		"A3": 100000000,
		"A4": 28614898,
		"A5": 0.001,
		"A6": 123456789.123,
		"A7": -1999830,
		"A8": 1e15,
	}
	for ref, val := range values {
		if err := s.SetValue(ref, val); err != nil {
			t.Fatalf("SetValue(%q, %v): %v", ref, val, err)
		}
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "largenum.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	s2 := f2.Sheet("Sheet1")

	for ref, want := range values {
		v, err := s2.GetValue(ref)
		if err != nil {
			t.Errorf("GetValue(%q): %v", ref, err)
			continue
		}
		got := v.Raw()
		if got != want {
			t.Errorf("GetValue(%q) = %v, want %v", ref, got, want)
		}
	}
}
