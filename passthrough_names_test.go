package werkbook_test

import (
	"path/filepath"
	"sort"
	"testing"

	"github.com/jpoz/werkbook"
)

func TestPassthroughElementNames(t *testing.T) {
	// A new sheet has no passthrough elements.
	if got := werkbook.New().Sheet("Sheet1").PassthroughElementNames(); len(got) != 0 {
		t.Errorf("new sheet: names = %v, want none", got)
	}

	path := filepath.Join(t.TempDir(), "template.xlsx")
	writeOpaquePassthroughFixture(t, path)
	f, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	s := f.Sheet("Sheet1")
	names := s.PassthroughElementNames()

	if !sort.StringsAreSorted(names) {
		t.Errorf("names are not sorted: %v", names)
	}
	seen := map[string]bool{}
	for _, name := range names {
		if name == "" || seen[name] {
			t.Errorf("empty or duplicate name %q in %v", name, names)
		}
		seen[name] = true
	}
	// The elements of the fixture that hold cell addresses or objects.
	for _, want := range []string{"conditionalFormatting", "dataValidations", "hyperlinks", "drawing", "legacyDrawing"} {
		if !seen[want] {
			t.Errorf("names = %v, want %q in the list", names, want)
		}
	}
	// Parts that the library models are not passthrough.
	for _, modeled := range []string{"sheetData", "mergeCells", "cols"} {
		if seen[modeled] {
			t.Errorf("names = %v; %q is modeled and must not be in the list", names, modeled)
		}
	}

	// The result is a copy.
	if len(names) > 0 {
		names[0] = "changed"
		if again := s.PassthroughElementNames(); again[0] == "changed" {
			t.Errorf("a change of the result changed the sheet")
		}
	}

	// A structural edit keeps the elements (they are not adjusted).
	before := s.PassthroughElementNames()
	if _, err := s.InsertRows(1, 1); err != nil {
		t.Fatalf("InsertRows: %v", err)
	}
	after := s.PassthroughElementNames()
	if len(after) != len(before) {
		t.Errorf("names after InsertRows = %v, want %v", after, before)
	}
}
