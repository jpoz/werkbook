package werkbook

import (
	"testing"

	"github.com/jpoz/werkbook/formula"
)

func TestTableStructuredReferences(t *testing.T) {
	// Create a workbook with a table and formulas that use structured references.
	f := New(FirstSheet("Data"))
	s := f.Sheet("Data")

	// Set up a table: A1:C5 with header row
	// Headers: Name, Price, Qty
	s.SetValue("A1", "Name")
	s.SetValue("B1", "Price")
	s.SetValue("C1", "Qty")
	// Data rows
	s.SetValue("A2", "Apple")
	s.SetValue("B2", 1.5)
	s.SetValue("C2", 10)
	s.SetValue("A3", "Banana")
	s.SetValue("B3", 0.75)
	s.SetValue("C3", 20)
	s.SetValue("A4", "Cherry")
	s.SetValue("B4", 3.0)
	s.SetValue("C4", 5)

	// Register a table definition.
	f.tables = []formula.TableInfo{
		{
			Name:       "Products",
			SheetName:  "Data",
			Columns:    []string{"Name", "Price", "Qty"},
			FirstCol:   1,
			FirstRow:   1,
			LastCol:    3,
			LastRow:    4,
			HeaderRows: 1,
			TotalRows:  0,
		},
	}

	// Set formula using table structured reference.
	s.SetFormula("D1", "SUM(Products[Price])")

	v, err := s.GetValue("D1")
	if err != nil {
		t.Fatalf("GetValue error: %v", err)
	}
	if v.Type != TypeNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	// 1.5 + 0.75 + 3.0 = 5.25
	if v.Number != 5.25 {
		t.Errorf("SUM(Products[Price]) = %v, want 5.25", v.Number)
	}

	// Test [#This Row] reference.
	s.SetFormula("D2", "Products[[#This Row],[Price]]*Products[[#This Row],[Qty]]")
	v, err = s.GetValue("D2")
	if err != nil {
		t.Fatalf("GetValue error: %v", err)
	}
	if v.Type != TypeNumber {
		t.Fatalf("expected number, got %v", v.Type)
	}
	// 1.5 * 10 = 15
	if v.Number != 15 {
		t.Errorf("Price*Qty for row 2 = %v, want 15", v.Number)
	}
}
