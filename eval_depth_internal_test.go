package werkbook

import (
	"fmt"
	"testing"
)

// After a deep evaluation, no evaluation state must stay in the workbook.
func TestEvalDeferred_LeavesNoState(t *testing.T) {
	const n = 5_000
	f := New()
	s := f.Sheet("Sheet1")
	f.BeginBatchWrite()
	if err := s.SetValue("A1", 1); err != nil {
		t.Fatal(err)
	}
	for row := 2; row <= n; row++ {
		if err := s.SetFormula(fmt.Sprintf("A%d", row), fmt.Sprintf("A%d+1", row-1)); err != nil {
			t.Fatal(err)
		}
	}
	// A cycle below the chain leaves waiting entries on the explicit stack
	// when the evaluation completes with an error value.
	if err := s.SetFormula("A1", fmt.Sprintf("A%d+1", n)); err != nil {
		t.Fatal(err)
	}
	f.EndBatchWrite()

	if _, err := s.GetValue(fmt.Sprintf("A%d", n)); err != nil {
		t.Fatal(err)
	}
	if f.probe != nil {
		t.Fatal("probe is not cleared")
	}
	if len(f.evaluating) != 0 {
		t.Fatalf("%d cells stay in evaluation", len(f.evaluating))
	}
	gen := f.calcGen
	for row := 1; row <= n; row++ {
		if s.rows[row].cells[1].needsEval(gen) {
			t.Fatalf("A%d is not resolved", row)
		}
	}
}
