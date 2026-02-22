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

func TestSetValueAndSave(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	if err := s.SetValue("A1", "hello"); err != nil {
		t.Fatalf("SetValue string: %v", err)
	}
	if err := s.SetValue("B1", 42); err != nil {
		t.Fatalf("SetValue int: %v", err)
	}
	if err := s.SetValue("C1", 3.14); err != nil {
		t.Fatalf("SetValue float: %v", err)
	}
	if err := s.SetValue("A2", true); err != nil {
		t.Fatalf("SetValue bool: %v", err)
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "test.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	// Open the ZIP and verify the sheet XML
	zr, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}
	defer zr.Close()

	// Verify sharedStrings.xml exists
	var hasSST bool
	var sheetXML string
	for _, zf := range zr.File {
		if zf.Name == "xl/sharedStrings.xml" {
			hasSST = true
		}
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

	if !hasSST {
		t.Error("expected xl/sharedStrings.xml in ZIP")
	}

	// Verify the sheet XML contains our cell references
	for _, ref := range []string{"A1", "B1", "C1", "A2"} {
		if !strings.Contains(sheetXML, ref) {
			t.Errorf("sheet XML missing cell ref %q", ref)
		}
	}

	// Parse and verify cell types
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

	cells := make(map[string]xmlC)
	for _, r := range ws.SheetData.Row {
		for _, c := range r.C {
			cells[c.R] = c
		}
	}

	// A1 is a string (SST index 0)
	if c, ok := cells["A1"]; !ok || c.T != "s" || c.V != "0" {
		t.Errorf("A1: got %+v, want t=s v=0", cells["A1"])
	}
	// B1 is a number
	if c, ok := cells["B1"]; !ok || c.T != "" || c.V != "42" {
		t.Errorf("B1: got %+v, want t= v=42", cells["B1"])
	}
	// C1 is a number
	if c, ok := cells["C1"]; !ok || c.T != "" || c.V != "3.14" {
		t.Errorf("C1: got %+v, want t= v=3.14", cells["C1"])
	}
	// A2 is a bool
	if c, ok := cells["A2"]; !ok || c.T != "b" || c.V != "1" {
		t.Errorf("A2: got %+v, want t=b v=1", cells["A2"])
	}
}

func TestGetValue(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", "hello")
	s.SetValue("B1", 42)
	s.SetValue("C1", true)

	tests := []struct {
		cell string
		want any
	}{
		{"A1", "hello"},
		{"B1", float64(42)},
		{"C1", true},
		{"D1", nil}, // empty
	}
	for _, tt := range tests {
		v, err := s.GetValue(tt.cell)
		if err != nil {
			t.Errorf("GetValue(%q): %v", tt.cell, err)
			continue
		}
		got := v.Raw()
		if got != tt.want {
			t.Errorf("GetValue(%q).Raw() = %v (%T), want %v (%T)", tt.cell, got, got, tt.want, tt.want)
		}
	}
}
