package fuzz

import (
	"strings"
	"testing"
)

func TestExtractFunctionsFromFormula(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		want    []string
	}{
		{"single", "SUM(A1:A3)", []string{"SUM"}},
		{"nested", "IF(SUM(A1:A3)>10,AVERAGE(A1:A3),0)", []string{"AVERAGE", "IF", "SUM"}},
		{"dot notation", "CEILING.MATH(3.7)", []string{"CEILING.MATH"}},
		{"duplicates", "SUM(A1)+SUM(A2)", []string{"SUM"}},
		{"case insensitive", "sum(A1:A3)", []string{"SUM"}},
		{"no functions", "A1+B1", nil},
		{"complex nesting", "IF(AND(A1>0,OR(B1<5,C1=0)),ROUND(SUM(D1:D5),2),0)", []string{"AND", "IF", "OR", "ROUND", "SUM"}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := ExtractFunctionsFromFormula(tt.formula)
			if len(got) != len(tt.want) {
				t.Errorf("ExtractFunctionsFromFormula(%q) = %v, want %v", tt.formula, got, tt.want)
				return
			}
			for i := range got {
				if got[i] != tt.want[i] {
					t.Errorf("ExtractFunctionsFromFormula(%q)[%d] = %q, want %q", tt.formula, i, got[i], tt.want[i])
				}
			}
		})
	}
}

func TestExtractFunctions(t *testing.T) {
	spec := &TestSpec{
		Sheets: []SheetSpec{
			{
				Name: "Sheet1",
				Cells: []CellSpec{
					{Ref: "A1", Value: 10},
					{Ref: "B1", Formula: "SUM(A1:A3)"},
					{Ref: "C1", Formula: "IF(A1>0,AVERAGE(A1:A3),0)"},
				},
			},
			{
				Name: "Sheet2",
				Cells: []CellSpec{
					{Ref: "A1", Formula: "ROUND(Sheet1!B1,2)"},
				},
			},
		},
	}

	got := ExtractFunctions(spec)
	want := []string{"AVERAGE", "IF", "ROUND", "SUM"}
	if len(got) != len(want) {
		t.Errorf("ExtractFunctions = %v, want %v", got, want)
		return
	}
	for i := range got {
		if got[i] != want[i] {
			t.Errorf("ExtractFunctions[%d] = %q, want %q", i, got[i], want[i])
		}
	}
}

func TestFunctionCoverageRecord(t *testing.T) {
	fc := NewFunctionCoverage()
	fc.Record([]string{"SUM", "IF"}, true)
	fc.Record([]string{"SUM", "TRUNC"}, false)

	if fc.Tested["SUM"].Passed != 1 || fc.Tested["SUM"].Failed != 1 {
		t.Errorf("SUM: got passed=%d failed=%d, want 1/1", fc.Tested["SUM"].Passed, fc.Tested["SUM"].Failed)
	}
	if fc.Tested["IF"].Passed != 1 || fc.Tested["IF"].Failed != 0 {
		t.Errorf("IF: got passed=%d failed=%d, want 1/0", fc.Tested["IF"].Passed, fc.Tested["IF"].Failed)
	}
	if fc.Tested["TRUNC"].Passed != 0 || fc.Tested["TRUNC"].Failed != 1 {
		t.Errorf("TRUNC: got passed=%d failed=%d, want 0/1", fc.Tested["TRUNC"].Passed, fc.Tested["TRUNC"].Failed)
	}
}

func TestFunctionCoverageFailingFunctions(t *testing.T) {
	fc := NewFunctionCoverage()
	fc.Record([]string{"SUM"}, true)
	fc.Record([]string{"TRUNC"}, false)
	fc.Record([]string{"IFS"}, false)
	fc.Record([]string{"IFS"}, true) // mixed — should NOT be in failing list

	failing := fc.FailingFunctions()
	if len(failing) != 1 || failing[0] != "TRUNC" {
		t.Errorf("FailingFunctions = %v, want [TRUNC]", failing)
	}
}

func TestFunctionCoverageSummary(t *testing.T) {
	fc := NewFunctionCoverage()
	fc.Record([]string{"SUM", "IF"}, true)
	fc.Record([]string{"TRUNC"}, false)

	summary := fc.Summary()
	if !strings.Contains(summary, "Passing") {
		t.Error("Summary should contain 'Passing'")
	}
	if !strings.Contains(summary, "Failing") {
		t.Error("Summary should contain 'Failing'")
	}
	if !strings.Contains(summary, "Untested") {
		t.Error("Summary should contain 'Untested'")
	}
}

func TestFunctionCoverageMarkBroken(t *testing.T) {
	fc := NewFunctionCoverage()
	fc.MarkBroken([]string{"TRUNC", "IFS"})
	fc.MarkBroken([]string{"TRUNC"})

	broken := fc.BrokenList()
	if len(broken) != 2 {
		t.Fatalf("BrokenList = %v, want 2 items", broken)
	}

	if fc.Broken["TRUNC"] != 2 {
		t.Errorf("TRUNC broken count = %d, want 2", fc.Broken["TRUNC"])
	}
}

func TestFunctionCoverageMarkFixed(t *testing.T) {
	fc := NewFunctionCoverage()
	fc.MarkBroken([]string{"TRUNC", "IFS"})
	fc.MarkFixed([]string{"TRUNC"})

	broken := fc.BrokenList()
	if len(broken) != 1 || broken[0] != "IFS" {
		t.Errorf("BrokenList after fix = %v, want [IFS]", broken)
	}
}

func TestFunctionCoverageIsHardBroken(t *testing.T) {
	fc := NewFunctionCoverage()
	fc.MarkBroken([]string{"TRUNC"})
	fc.MarkBroken([]string{"TRUNC"})
	if fc.IsHardBroken("TRUNC") {
		t.Error("TRUNC should not be hard broken after 2 failures")
	}
	fc.MarkBroken([]string{"TRUNC"})
	if !fc.IsHardBroken("TRUNC") {
		t.Error("TRUNC should be hard broken after 3 failures")
	}
}

func TestExtractBrokenFunctionsFromMismatches_NameError(t *testing.T) {
	spec := &TestSpec{
		Sheets: []SheetSpec{{
			Name: "Sheet1",
			Cells: []CellSpec{
				{Ref: "C1", Formula: "TRUNC(3.7)"},
			},
		}},
	}
	mismatches := []Mismatch{{
		Ref:      "Sheet1!C1",
		Werkbook: "#NAME?",
		Oracle:   "3",
	}}

	broken := ExtractBrokenFunctionsFromMismatches(spec, mismatches)
	if len(broken) != 1 || broken[0] != "TRUNC" {
		t.Errorf("got %v, want [TRUNC]", broken)
	}
}

func TestExtractBrokenFunctionsFromMismatches_ValueMismatch(t *testing.T) {
	spec := &TestSpec{
		Sheets: []SheetSpec{{
			Name: "Sheet1",
			Cells: []CellSpec{
				{Ref: "C1", Formula: "IF(A1>0,TRUNC(A1),0)"},
			},
		}},
	}
	mismatches := []Mismatch{{
		Ref:      "Sheet1!C1",
		Werkbook: "3.7",
		Oracle:   "3",
	}}

	broken := ExtractBrokenFunctionsFromMismatches(spec, mismatches)
	if len(broken) != 2 {
		t.Errorf("got %v, want [IF TRUNC]", broken)
	}
}

func TestExtractBrokenFunctionsFromMismatches_BareRef(t *testing.T) {
	spec := &TestSpec{
		Sheets: []SheetSpec{{
			Name: "Sheet1",
			Cells: []CellSpec{
				{Ref: "C1", Formula: "TRUNC(3.7)"},
			},
		}},
	}
	mismatches := []Mismatch{{
		Ref:      "C1",
		Werkbook: "#NAME?",
		Oracle:   "3",
	}}

	broken := ExtractBrokenFunctionsFromMismatches(spec, mismatches)
	if len(broken) != 1 || broken[0] != "TRUNC" {
		t.Errorf("got %v, want [TRUNC]", broken)
	}
}
