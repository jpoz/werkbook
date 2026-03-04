package formula

import "testing"

func TestExpandDefinedNames(t *testing.T) {
	names := []DefinedNameInfo{
		{Name: "OneRange", Value: "Sheet1!$A$10", LocalSheetID: -1},
		{Name: "valuation_date", Value: "Sheet1!$A$1", LocalSheetID: -1},
		{Name: "MyConst", Value: "42", LocalSheetID: -1},
	}

	tests := []struct {
		formula  string
		sheetIdx int
		want     string
	}{
		// Simple name replacement.
		{"B1+OneRange", 0, "B1+Sheet1!$A$10"},
		// Case-insensitive match.
		{"B1+onerange", 0, "B1+Sheet1!$A$10"},
		// Multiple names in one formula.
		{"OneRange+valuation_date", 0, "Sheet1!$A$10+Sheet1!$A$1"},
		// Name inside string literal should NOT be replaced.
		{`"OneRange"`, 0, `"OneRange"`},
		// Name followed by '(' is a function call, not replaced.
		{"OneRange(1,2)", 0, "OneRange(1,2)"},
		// Name followed by '!' is a sheet prefix, not replaced.
		{"OneRange!A1", 0, "OneRange!A1"},
		// Name preceded by '!' (sheet-qualified ref), not replaced.
		{"Sheet1!OneRange", 0, "Sheet1!OneRange"},
		// Constant value.
		{"A1+MyConst", 0, "A1+42"},
		// No names to expand.
		{"A1+B1", 0, "A1+B1"},
		// Name at end of formula.
		{"SUM(A1:A5)+OneRange", 0, "SUM(A1:A5)+Sheet1!$A$10"},
	}

	for _, tt := range tests {
		got := ExpandDefinedNames(tt.formula, names, tt.sheetIdx)
		if got != tt.want {
			t.Errorf("ExpandDefinedNames(%q) = %q, want %q", tt.formula, got, tt.want)
		}
	}
}

func TestExpandDefinedNamesLocalScope(t *testing.T) {
	names := []DefinedNameInfo{
		{Name: "Rate", Value: "0.05", LocalSheetID: 0},       // local to sheet 0
		{Name: "Rate", Value: "0.10", LocalSheetID: 1},       // local to sheet 1
		{Name: "GlobalName", Value: "Sheet1!$B$2", LocalSheetID: -1}, // global
	}

	// On sheet 0, should get the sheet-0-local value.
	got := ExpandDefinedNames("A1*Rate", names, 0)
	if want := "A1*0.05"; got != want {
		t.Errorf("sheet 0: got %q, want %q", got, want)
	}

	// On sheet 1, should get the sheet-1-local value.
	got = ExpandDefinedNames("A1*Rate", names, 1)
	if want := "A1*0.10"; got != want {
		t.Errorf("sheet 1: got %q, want %q", got, want)
	}

	// On sheet 2, no local Rate is visible, so Rate is not expanded.
	got = ExpandDefinedNames("A1*Rate", names, 2)
	if want := "A1*Rate"; got != want {
		t.Errorf("sheet 2: got %q, want %q", got, want)
	}

	// Global name should be visible on any sheet.
	got = ExpandDefinedNames("GlobalName+1", names, 2)
	if want := "Sheet1!$B$2+1"; got != want {
		t.Errorf("global on sheet 2: got %q, want %q", got, want)
	}
}

func TestExpandDefinedNamesEmpty(t *testing.T) {
	// No names at all should return formula unchanged.
	got := ExpandDefinedNames("A1+B1", nil, 0)
	if want := "A1+B1"; got != want {
		t.Errorf("nil names: got %q, want %q", got, want)
	}
}
