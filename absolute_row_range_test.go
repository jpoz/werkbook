package werkbook

import (
	"bytes"
	"testing"
)

// Row ranges with absolute marks ($1:$3) must evaluate as the same range
// without marks does, and must survive a save.
func TestAbsoluteRowRangeEvaluation(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	other, err := f.NewSheet("Data")
	if err != nil {
		t.Fatal(err)
	}
	steps := []error{
		s.SetValue("A1", 1), s.SetValue("B2", 2), s.SetValue("C3", 4), s.SetValue("A4", 100),
		other.SetValue("A1", 10), other.SetValue("Z2", 20), other.SetValue("A3", 1000),
	}
	for i, err := range steps {
		if err != nil {
			t.Fatalf("setup step %d: %v", i, err)
		}
	}

	tests := []struct {
		formula string
		want    float64
	}{
		{"SUM($1:$3)", 7},
		{"SUM(1:$3)", 7},
		{"SUM($1:3)", 7},
		{"SUM(1:3)", 7},
		{"SUM($2:$2)", 2},
		{"SUM($1:$3)+SUM($4:$4)", 107},
		{"SUM(Data!$1:$2)", 30},
		{"SUM('Data'!1:$2)", 30},
		{"COUNT($1:$4)", 4},
		{"ROWS($2:$5)", 4},
		{"COLUMNS($2:$5)", 16384},
	}
	for i, tt := range tests {
		// The formulas are in column F on rows 10 and below, out of the
		// ranges that they read.
		ref, _ := CoordinatesToCellName(6, 10+i)
		if err := s.SetFormula(ref, tt.formula); err != nil {
			t.Fatalf("SetFormula(%s): %v", tt.formula, err)
		}
		v, err := s.GetValue(ref)
		if err != nil {
			t.Fatal(err)
		}
		if v.Type != TypeNumber || v.Number != tt.want {
			t.Errorf("%s = %#v, want %v", tt.formula, v, tt.want)
		}
	}

	// A precedent in the row range makes the formula recalculate.
	if err := s.SetValue("XFD2", 50); err != nil {
		t.Fatal(err)
	}
	if v, _ := s.GetValue("F10"); v.Number != 57 {
		t.Errorf("SUM($1:$3) after a write to XFD2 = %v, want 57", v.Number)
	}

	// A lone absolute row is an error value, not a number. The cell is out
	// of the row ranges above.
	if err := s.SetFormula("H30", "$3"); err != nil {
		t.Fatal(err)
	}
	if v, _ := s.GetValue("H30"); v.Type != TypeError {
		t.Errorf("=$3 gives %#v, want an error value", v)
	}

	// Save and open: the formula text and the result stay.
	var buf bytes.Buffer
	if _, err := f.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	f2, err := OpenReader(&buf)
	if err != nil {
		t.Fatal(err)
	}
	s2 := f2.Sheet("Sheet1")
	if got, _ := s2.GetFormula("F10"); got != "SUM($1:$3)" {
		t.Errorf("formula after save and open = %q", got)
	}
	f2.Recalculate()
	if v, _ := s2.GetValue("F10"); v.Number != 57 {
		t.Errorf("SUM($1:$3) after save and open = %#v, want 57", v)
	}
}

// Excel writes print titles as a defined name with an absolute row range.
func TestDefinedNameWithAbsoluteRowRange(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	for _, err := range []error{s.SetValue("A1", 3), s.SetValue("C2", 4), s.SetValue("A3", 100)} {
		if err != nil {
			t.Fatal(err)
		}
	}
	if err := f.SetDefinedName(DefinedName{Name: "TopRows", Value: "Sheet1!$1:$2", LocalSheetID: -1}); err != nil {
		t.Fatal(err)
	}
	if err := s.SetFormula("E5", "SUM(TopRows)"); err != nil {
		t.Fatal(err)
	}
	if v, _ := s.GetValue("E5"); v.Type != TypeNumber || v.Number != 7 {
		t.Errorf("SUM(TopRows) = %#v, want 7", v)
	}
}
