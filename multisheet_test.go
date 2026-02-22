package werkbook_test

import (
	"path/filepath"
	"slices"
	"testing"

	"github.com/jpoz/werkbook"
)

func TestMultiSheetRoundTrip(t *testing.T) {
	f := werkbook.New()

	// Sheet1 already exists
	f.Sheet("Sheet1").SetValue("A1", "first")

	s2, err := f.NewSheet("Data")
	if err != nil {
		t.Fatalf("NewSheet(Data): %v", err)
	}
	s2.SetValue("A1", "second")
	s2.SetValue("B1", 200)

	s3, err := f.NewSheet("Summary")
	if err != nil {
		t.Fatalf("NewSheet(Summary): %v", err)
	}
	s3.SetValue("A1", "third")
	s3.SetValue("C3", true)

	dir := t.TempDir()
	path := filepath.Join(dir, "multi.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	// Re-open
	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	names := f2.SheetNames()
	if !slices.Equal(names, []string{"Sheet1", "Data", "Summary"}) {
		t.Fatalf("expected [Sheet1, Data, Summary], got %v", names)
	}

	// Verify data
	v, _ := f2.Sheet("Sheet1").GetValue("A1")
	if v.Raw() != "first" {
		t.Errorf("Sheet1!A1 = %v, want first", v.Raw())
	}

	v, _ = f2.Sheet("Data").GetValue("A1")
	if v.Raw() != "second" {
		t.Errorf("Data!A1 = %v, want second", v.Raw())
	}
	v, _ = f2.Sheet("Data").GetValue("B1")
	if v.Raw() != float64(200) {
		t.Errorf("Data!B1 = %v, want 200", v.Raw())
	}

	v, _ = f2.Sheet("Summary").GetValue("A1")
	if v.Raw() != "third" {
		t.Errorf("Summary!A1 = %v, want third", v.Raw())
	}
	v, _ = f2.Sheet("Summary").GetValue("C3")
	if v.Raw() != true {
		t.Errorf("Summary!C3 = %v, want true", v.Raw())
	}
}

func TestDeleteSheet(t *testing.T) {
	f := werkbook.New()
	f.NewSheet("Sheet2")
	f.NewSheet("Sheet3")

	if err := f.DeleteSheet("Sheet2"); err != nil {
		t.Fatalf("DeleteSheet: %v", err)
	}

	names := f.SheetNames()
	if !slices.Equal(names, []string{"Sheet1", "Sheet3"}) {
		t.Errorf("expected [Sheet1, Sheet3], got %v", names)
	}

	// Can't delete nonexistent sheet
	if err := f.DeleteSheet("nope"); err == nil {
		t.Error("expected error deleting nonexistent sheet")
	}

	// Can't delete last sheet
	f.DeleteSheet("Sheet3")
	if err := f.DeleteSheet("Sheet1"); err == nil {
		t.Error("expected error deleting last sheet")
	}
}

func TestNewSheetDuplicate(t *testing.T) {
	f := werkbook.New()
	_, err := f.NewSheet("Sheet1")
	if err == nil {
		t.Error("expected error creating duplicate sheet")
	}
}
