package werkbook_test

import (
	"testing"

	"github.com/jpoz/werkbook"
)

// TestFormulaCacheInvalidation verifies that changing a cell value causes
// dependent formulas to re-evaluate, not return stale cached results.
func TestFormulaCacheInvalidation(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 10)
	s.SetFormula("B1", "A1*2")

	// First read: evaluates and caches.
	val, err := s.GetValue("B1")
	if err != nil {
		t.Fatalf("GetValue B1: %v", err)
	}
	if val.Number != 20 {
		t.Fatalf("B1 = %g, want 20", val.Number)
	}

	// Change the dependency.
	s.SetValue("A1", 50)

	// Must re-evaluate, not return the stale 20.
	val, err = s.GetValue("B1")
	if err != nil {
		t.Fatalf("GetValue B1 after mutation: %v", err)
	}
	if val.Number != 100 {
		t.Errorf("B1 after A1=50: got %g, want 100", val.Number)
	}
}

// TestFormulaCacheInvalidationChain verifies transitive re-evaluation:
// C1 depends on B1, B1 depends on A1. Changing A1 should propagate through.
func TestFormulaCacheInvalidationChain(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 5)
	s.SetFormula("B1", "A1+10")
	s.SetFormula("C1", "B1*3")

	val, _ := s.GetValue("C1")
	if val.Number != 45 {
		t.Fatalf("C1 = %g, want 45", val.Number)
	}

	s.SetValue("A1", 10)

	val, _ = s.GetValue("C1")
	if val.Number != 60 {
		t.Errorf("C1 after A1=10: got %g, want 60", val.Number)
	}
}

// TestFormulaReturningZero verifies that a formula evaluating to 0 is
// properly cached and still re-evaluates when dependencies change.
func TestFormulaReturningZero(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 5)
	s.SetValue("A2", 5)
	s.SetFormula("B1", "A1-A2") // = 0

	val, _ := s.GetValue("B1")
	if val.Type != werkbook.TypeNumber || val.Number != 0 {
		t.Fatalf("B1 = %#v, want 0", val)
	}

	s.SetValue("A1", 10)

	val, _ = s.GetValue("B1")
	if val.Number != 5 {
		t.Errorf("B1 after A1=10: got %g, want 5", val.Number)
	}
}

// TestFormulaReturningFalse verifies that a formula evaluating to FALSE is
// properly cached and still re-evaluates when dependencies change.
func TestFormulaReturningFalse(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 1)
	s.SetFormula("B1", "A1>10") // = FALSE

	val, _ := s.GetValue("B1")
	if val.Type != werkbook.TypeBool || val.Bool != false {
		t.Fatalf("B1 = %#v, want FALSE", val)
	}

	s.SetValue("A1", 20)

	val, _ = s.GetValue("B1")
	if val.Type != werkbook.TypeBool || val.Bool != true {
		t.Errorf("B1 after A1=20: got %#v, want TRUE", val)
	}
}

// TestFormulaReturningEmptyString verifies that a formula evaluating to ""
// is properly cached and still re-evaluates when dependencies change.
func TestFormulaReturningEmptyString(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 1)
	s.SetFormula("B1", `IF(A1>10,"big","")`) // = ""

	val, _ := s.GetValue("B1")
	if val.Type != werkbook.TypeString || val.String != "" {
		t.Fatalf("B1 = %#v, want empty string", val)
	}

	s.SetValue("A1", 20)

	val, _ = s.GetValue("B1")
	if val.Type != werkbook.TypeString || val.String != "big" {
		t.Errorf("B1 after A1=20: got %#v, want \"big\"", val)
	}
}

// TestSetValueClearsFormula verifies that SetValue on a formula cell
// removes the formula.
func TestSetValueClearsFormula(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetFormula("A1", "1+1")
	val, _ := s.GetValue("A1")
	if val.Number != 2 {
		t.Fatalf("A1 = %g, want 2", val.Number)
	}

	// SetValue should replace the formula.
	s.SetValue("A1", 99)

	val, _ = s.GetValue("A1")
	if val.Number != 99 {
		t.Errorf("A1 after SetValue: got %g, want 99", val.Number)
	}
	formula, _ := s.GetFormula("A1")
	if formula != "" {
		t.Errorf("formula after SetValue: got %q, want empty", formula)
	}
}

// TestCrossSheetFormula verifies that a formula can reference a cell on
// another sheet.
func TestCrossSheetFormula(t *testing.T) {
	f := werkbook.New()
	s1 := f.Sheet("Sheet1")

	s2, _ := f.NewSheet("Sheet2")
	s2.SetValue("A1", 100)

	s1.SetFormula("A1", "Sheet2!A1*2")

	val, err := s1.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue: %v", err)
	}
	if val.Type != werkbook.TypeNumber || val.Number != 200 {
		t.Errorf("Sheet1!A1 = %#v, want 200", val)
	}
}

// TestCrossSheetRange verifies that a formula can reference a range on
// another sheet.
func TestCrossSheetRange(t *testing.T) {
	f := werkbook.New()
	s1 := f.Sheet("Sheet1")

	s2, _ := f.NewSheet("Data")
	s2.SetValue("A1", 10)
	s2.SetValue("A2", 20)
	s2.SetValue("A3", 30)

	s1.SetFormula("A1", "SUM(Data!A1:A3)")

	val, err := s1.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue: %v", err)
	}
	if val.Type != werkbook.TypeNumber || val.Number != 60 {
		t.Errorf("SUM(Data!A1:A3) = %#v, want 60", val)
	}
}

// TestCrossSheetInvalidation verifies that changing a value on one sheet
// causes formulas on another sheet that reference it to re-evaluate.
func TestCrossSheetInvalidation(t *testing.T) {
	f := werkbook.New()
	s1 := f.Sheet("Sheet1")

	s2, _ := f.NewSheet("Sheet2")
	s2.SetValue("A1", 10)

	s1.SetFormula("A1", "Sheet2!A1+5")

	val, _ := s1.GetValue("A1")
	if val.Number != 15 {
		t.Fatalf("Sheet1!A1 = %g, want 15", val.Number)
	}

	// Change the dependency on Sheet2.
	s2.SetValue("A1", 100)

	val, _ = s1.GetValue("A1")
	if val.Number != 105 {
		t.Errorf("Sheet1!A1 after Sheet2!A1=100: got %g, want 105", val.Number)
	}
}

// TestCrossSheetFormulaChain verifies formula chains across sheets:
// Sheet3!A1 depends on Sheet2!A1, Sheet2!A1 depends on Sheet1!A1.
func TestCrossSheetFormulaChain(t *testing.T) {
	f := werkbook.New()
	s1 := f.Sheet("Sheet1")
	s1.SetValue("A1", 5)

	s2, _ := f.NewSheet("Sheet2")
	s2.SetFormula("A1", "Sheet1!A1*10")

	s3, _ := f.NewSheet("Sheet3")
	s3.SetFormula("A1", "Sheet2!A1+1")

	val, _ := s3.GetValue("A1")
	if val.Number != 51 {
		t.Fatalf("Sheet3!A1 = %g, want 51", val.Number)
	}

	s1.SetValue("A1", 8)

	val, _ = s3.GetValue("A1")
	if val.Number != 81 {
		t.Errorf("Sheet3!A1 after Sheet1!A1=8: got %g, want 81", val.Number)
	}
}

// TestCrossSheetCircularRef verifies that circular references spanning
// multiple sheets are detected.
func TestCrossSheetCircularRef(t *testing.T) {
	f := werkbook.New()
	s1 := f.Sheet("Sheet1")

	s2, _ := f.NewSheet("Sheet2")

	s1.SetFormula("A1", "Sheet2!A1+1")
	s2.SetFormula("A1", "Sheet1!A1+1")

	val, _ := s1.GetValue("A1")
	if val.Type != werkbook.TypeError {
		t.Errorf("expected error for circular ref, got %#v", val)
	}
}

// TestBadSheetRefReturnsError verifies that referencing a non-existent
// sheet in a formula produces a #REF! error.
func TestBadSheetRefReturnsError(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetFormula("A1", "NoSuchSheet!A1")

	val, _ := s.GetValue("A1")
	if val.Type != werkbook.TypeError {
		t.Errorf("expected error for bad sheet ref, got %#v", val)
	}
}
