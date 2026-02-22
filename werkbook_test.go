package werkbook_test

import (
	"archive/zip"
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
