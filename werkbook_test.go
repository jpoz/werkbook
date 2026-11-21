package werkbook_test

import (
	"archive/zip"
	"bytes"
	"os"
	"path/filepath"
	"slices"
	"testing"

	"github.com/jpoz/werkbook"
)

func TestNewSaveAsProducesValidZIP(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "test.xlsx")

	f := werkbook.New()
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	// Verify it's a valid ZIP
	zr, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}
	defer zr.Close()

	expected := []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"xl/workbook.xml",
		"xl/_rels/workbook.xml.rels",
		"xl/styles.xml",
		"xl/worksheets/sheet1.xml",
	}

	var got []string
	for _, f := range zr.File {
		got = append(got, f.Name)
	}

	for _, want := range expected {
		if !slices.Contains(got, want) {
			t.Errorf("missing ZIP entry %q; got %v", want, got)
		}
	}
}

func TestNewHasDefaultSheet(t *testing.T) {
	f := werkbook.New()
	names := f.SheetNames()
	if len(names) != 1 || names[0] != "Sheet1" {
		t.Errorf("expected [Sheet1], got %v", names)
	}
	if f.Sheet("Sheet1") == nil {
		t.Error("Sheet(\"Sheet1\") returned nil")
	}
	if f.Sheet("nope") != nil {
		t.Error("Sheet(\"nope\") should return nil")
	}
}

func TestSaveAsCreatesFile(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "out.xlsx")

	if err := werkbook.New().SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	info, err := os.Stat(path)
	if err != nil {
		t.Fatalf("stat: %v", err)
	}
	if info.Size() == 0 {
		t.Error("file is empty")
	}
}

func TestSetDate1904(t *testing.T) {
	f := werkbook.New()

	// Default should be false.
	if f.Date1904() {
		t.Error("Date1904 should default to false")
	}

	// Enable.
	f.SetDate1904(true)
	if !f.Date1904() {
		t.Error("Date1904 should be true after SetDate1904(true)")
	}

	// Setting to same value should be a no-op (early return path).
	f.SetDate1904(true)
	if !f.Date1904() {
		t.Error("Date1904 should still be true")
	}

	// Disable.
	f.SetDate1904(false)
	if f.Date1904() {
		t.Error("Date1904 should be false after SetDate1904(false)")
	}
}

func TestCellFormula(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	if err := s.SetFormula("A1", "1+2"); err != nil {
		t.Fatal(err)
	}
	if err := s.SetValue("B1", "plain"); err != nil {
		t.Fatal(err)
	}

	for row := range s.Rows() {
		for _, cell := range row.Cells() {
			switch cell.Col() {
			case 1: // A1
				if cell.Formula() != "1+2" {
					t.Errorf("A1.Formula() = %q, want '1+2'", cell.Formula())
				}
			case 2: // B1
				if cell.Formula() != "" {
					t.Errorf("B1.Formula() = %q, want empty", cell.Formula())
				}
			}
		}
	}
}

func TestCalcProperties(t *testing.T) {
	f := werkbook.New()

	// Default should be zero value.
	cp := f.CalcProperties()
	if cp.Mode != "" || cp.ID != 0 {
		t.Error("default CalcProperties should be zero value")
	}

	// Set and get.
	f.SetCalcProperties(werkbook.CalcProperties{
		Mode:           "auto",
		ID:             12345,
		FullCalcOnLoad: true,
	})
	cp = f.CalcProperties()
	if cp.Mode != "auto" {
		t.Errorf("Mode = %q, want auto", cp.Mode)
	}
	if cp.ID != 12345 {
		t.Errorf("ID = %d, want 12345", cp.ID)
	}
	if !cp.FullCalcOnLoad {
		t.Error("FullCalcOnLoad should be true")
	}
}

func TestWriteToAndOpenReader(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	if err := s.SetValue("A1", "hello"); err != nil {
		t.Fatal(err)
	}
	if err := s.SetFormula("B1", `A1&" world"`); err != nil {
		t.Fatal(err)
	}

	var buf bytes.Buffer
	n, err := f.WriteTo(&buf)
	if err != nil {
		t.Fatalf("WriteTo: %v", err)
	}
	if n != int64(buf.Len()) {
		t.Fatalf("WriteTo wrote %d bytes, buffer has %d", n, buf.Len())
	}

	f2, err := werkbook.OpenReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}

	v, err := f2.Sheet("Sheet1").GetValue("B1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != werkbook.TypeString || v.String != "hello world" {
		t.Fatalf("B1 = %#v, want \"hello world\"", v)
	}
}
