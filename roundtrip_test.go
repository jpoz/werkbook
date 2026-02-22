package werkbook_test

import (
	"path/filepath"
	"testing"

	"github.com/jpoz/werkbook"
)

func TestRoundTrip(t *testing.T) {
	// Write a file with various cell types.
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	cells := map[string]any{
		"A1": "hello",
		"B1": 42,
		"C1": 3.14,
		"A2": true,
		"B2": false,
		"A3": "world",
		"D5": 0,
	}
	for ref, val := range cells {
		if err := s.SetValue(ref, val); err != nil {
			t.Fatalf("SetValue(%q, %v): %v", ref, val, err)
		}
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "roundtrip.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	// Read it back.
	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	names := f2.SheetNames()
	if len(names) != 1 || names[0] != "Sheet1" {
		t.Fatalf("expected [Sheet1], got %v", names)
	}

	s2 := f2.Sheet("Sheet1")

	// Check each value.
	expected := map[string]any{
		"A1": "hello",
		"B1": float64(42),
		"C1": 3.14,
		"A2": true,
		"B2": false,
		"A3": "world",
		"D5": float64(0),
	}
	for ref, want := range expected {
		v, err := s2.GetValue(ref)
		if err != nil {
			t.Errorf("GetValue(%q): %v", ref, err)
			continue
		}
		got := v.Raw()
		if got != want {
			t.Errorf("GetValue(%q) = %v (%T), want %v (%T)", ref, got, got, want, want)
		}
	}

	// Empty cells should return TypeEmpty.
	v, err := s2.GetValue("Z99")
	if err != nil {
		t.Fatalf("GetValue(Z99): %v", err)
	}
	if !v.IsEmpty() {
		t.Errorf("expected empty value for Z99, got %#v", v)
	}
}

func TestRoundTripDuplicateStrings(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Same string in multiple cells should use the same SST index.
	s.SetValue("A1", "dup")
	s.SetValue("A2", "dup")
	s.SetValue("A3", "other")

	dir := t.TempDir()
	path := filepath.Join(dir, "dup.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	s2 := f2.Sheet("Sheet1")

	for _, ref := range []string{"A1", "A2"} {
		v, _ := s2.GetValue(ref)
		if v.Raw() != "dup" {
			t.Errorf("GetValue(%q) = %v, want dup", ref, v.Raw())
		}
	}
	v, _ := s2.GetValue("A3")
	if v.Raw() != "other" {
		t.Errorf("GetValue(A3) = %v, want other", v.Raw())
	}
}
