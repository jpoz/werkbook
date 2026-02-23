package formula

import (
	"sort"
	"testing"
)

func sortQC(cells []QualifiedCell) {
	sort.Slice(cells, func(i, j int) bool {
		a, b := cells[i], cells[j]
		if a.Sheet != b.Sheet {
			return a.Sheet < b.Sheet
		}
		if a.Col != b.Col {
			return a.Col < b.Col
		}
		return a.Row < b.Row
	})
}

func qcEqual(a, b []QualifiedCell) bool {
	if len(a) != len(b) {
		return false
	}
	sortQC(a)
	sortQC(b)
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

func TestRegisterAndDirectDependents(t *testing.T) {
	g := NewDepGraph()

	// B1 = A1 + A2
	b1 := QualifiedCell{Sheet: "Sheet1", Col: 2, Row: 1}
	g.Register(b1, "Sheet1", []CellAddr{
		{Col: 1, Row: 1},
		{Col: 1, Row: 2},
	}, nil)

	deps := g.DirectDependents(QualifiedCell{Sheet: "Sheet1", Col: 1, Row: 1})
	if !qcEqual(deps, []QualifiedCell{b1}) {
		t.Errorf("expected [B1], got %v", deps)
	}

	deps = g.DirectDependents(QualifiedCell{Sheet: "Sheet1", Col: 1, Row: 2})
	if !qcEqual(deps, []QualifiedCell{b1}) {
		t.Errorf("expected [B1], got %v", deps)
	}

	// No dependents for unrelated cell.
	deps = g.DirectDependents(QualifiedCell{Sheet: "Sheet1", Col: 3, Row: 1})
	if len(deps) != 0 {
		t.Errorf("expected no dependents, got %v", deps)
	}
}

func TestUnregister(t *testing.T) {
	g := NewDepGraph()

	b1 := QualifiedCell{Sheet: "Sheet1", Col: 2, Row: 1}
	g.Register(b1, "Sheet1", []CellAddr{
		{Col: 1, Row: 1},
	}, nil)

	g.Unregister(b1)

	deps := g.DirectDependents(QualifiedCell{Sheet: "Sheet1", Col: 1, Row: 1})
	if len(deps) != 0 {
		t.Errorf("expected no dependents after unregister, got %v", deps)
	}
}

func TestTransitiveDependents(t *testing.T) {
	g := NewDepGraph()

	// A1 → B1 → C1
	b1 := QualifiedCell{Sheet: "S", Col: 2, Row: 1}
	c1 := QualifiedCell{Sheet: "S", Col: 3, Row: 1}

	g.Register(b1, "S", []CellAddr{{Col: 1, Row: 1}}, nil)
	g.Register(c1, "S", []CellAddr{{Col: 2, Row: 1}}, nil)

	deps := g.Dependents(QualifiedCell{Sheet: "S", Col: 1, Row: 1})
	if !qcEqual(deps, []QualifiedCell{b1, c1}) {
		t.Errorf("expected [B1, C1], got %v", deps)
	}
}

func TestRangeContainment(t *testing.T) {
	g := NewDepGraph()

	// D1 = SUM(A1:C3)
	d1 := QualifiedCell{Sheet: "S", Col: 4, Row: 1}
	g.Register(d1, "S", nil, []RangeAddr{
		{FromCol: 1, FromRow: 1, ToCol: 3, ToRow: 3},
	})

	// B2 is inside the range.
	deps := g.DirectDependents(QualifiedCell{Sheet: "S", Col: 2, Row: 2})
	if !qcEqual(deps, []QualifiedCell{d1}) {
		t.Errorf("expected [D1], got %v", deps)
	}

	// D4 is outside the range.
	deps = g.DirectDependents(QualifiedCell{Sheet: "S", Col: 4, Row: 4})
	if len(deps) != 0 {
		t.Errorf("expected no dependents, got %v", deps)
	}
}

func TestCrossSheetDeps(t *testing.T) {
	g := NewDepGraph()

	// Sheet2!A1 = Sheet1!A1 * 2
	s2a1 := QualifiedCell{Sheet: "Sheet2", Col: 1, Row: 1}
	g.Register(s2a1, "Sheet2", []CellAddr{
		{Sheet: "Sheet1", Col: 1, Row: 1},
	}, nil)

	deps := g.DirectDependents(QualifiedCell{Sheet: "Sheet1", Col: 1, Row: 1})
	if !qcEqual(deps, []QualifiedCell{s2a1}) {
		t.Errorf("expected [Sheet2!A1], got %v", deps)
	}

	// Different sheet should have no dependents.
	deps = g.DirectDependents(QualifiedCell{Sheet: "Sheet2", Col: 1, Row: 1})
	if len(deps) != 0 {
		t.Errorf("expected no dependents, got %v", deps)
	}
}

func TestReRegistrationReplacesEdges(t *testing.T) {
	g := NewDepGraph()

	b1 := QualifiedCell{Sheet: "S", Col: 2, Row: 1}

	// B1 = A1
	g.Register(b1, "S", []CellAddr{{Col: 1, Row: 1}}, nil)

	// Re-register: B1 = A2 (no longer depends on A1)
	g.Register(b1, "S", []CellAddr{{Col: 1, Row: 2}}, nil)

	deps := g.DirectDependents(QualifiedCell{Sheet: "S", Col: 1, Row: 1})
	if len(deps) != 0 {
		t.Errorf("expected no dependents for A1 after re-registration, got %v", deps)
	}

	deps = g.DirectDependents(QualifiedCell{Sheet: "S", Col: 1, Row: 2})
	if !qcEqual(deps, []QualifiedCell{b1}) {
		t.Errorf("expected [B1] for A2, got %v", deps)
	}
}

func TestUnqualifiedRefResolution(t *testing.T) {
	g := NewDepGraph()

	// Formula on Sheet2, referencing A1 without sheet qualifier → resolves to Sheet2!A1
	b1 := QualifiedCell{Sheet: "Sheet2", Col: 2, Row: 1}
	g.Register(b1, "Sheet2", []CellAddr{
		{Col: 1, Row: 1}, // unqualified
	}, nil)

	// Should be dependent on Sheet2!A1, not Sheet1!A1.
	deps := g.DirectDependents(QualifiedCell{Sheet: "Sheet2", Col: 1, Row: 1})
	if !qcEqual(deps, []QualifiedCell{b1}) {
		t.Errorf("expected [Sheet2!B1], got %v", deps)
	}

	deps = g.DirectDependents(QualifiedCell{Sheet: "Sheet1", Col: 1, Row: 1})
	if len(deps) != 0 {
		t.Errorf("expected no dependents on Sheet1!A1, got %v", deps)
	}
}

func TestUnregisterClearsRangeSubs(t *testing.T) {
	g := NewDepGraph()

	d1 := QualifiedCell{Sheet: "S", Col: 4, Row: 1}
	g.Register(d1, "S", nil, []RangeAddr{
		{FromCol: 1, FromRow: 1, ToCol: 3, ToRow: 3},
	})

	g.Unregister(d1)

	deps := g.DirectDependents(QualifiedCell{Sheet: "S", Col: 2, Row: 2})
	if len(deps) != 0 {
		t.Errorf("expected no dependents after unregistering range sub, got %v", deps)
	}
}

func TestFanOut(t *testing.T) {
	g := NewDepGraph()

	// A1 is read by B1, C1, D1
	b1 := QualifiedCell{Sheet: "S", Col: 2, Row: 1}
	c1 := QualifiedCell{Sheet: "S", Col: 3, Row: 1}
	d1 := QualifiedCell{Sheet: "S", Col: 4, Row: 1}

	g.Register(b1, "S", []CellAddr{{Col: 1, Row: 1}}, nil)
	g.Register(c1, "S", []CellAddr{{Col: 1, Row: 1}}, nil)
	g.Register(d1, "S", []CellAddr{{Col: 1, Row: 1}}, nil)

	deps := g.DirectDependents(QualifiedCell{Sheet: "S", Col: 1, Row: 1})
	expected := []QualifiedCell{b1, c1, d1}
	if !qcEqual(deps, expected) {
		t.Errorf("expected %v, got %v", expected, deps)
	}
}

func TestDependentsNoCycles(t *testing.T) {
	g := NewDepGraph()

	// Diamond: A1 → B1, A1 → C1, B1 → D1, C1 → D1
	b1 := QualifiedCell{Sheet: "S", Col: 2, Row: 1}
	c1 := QualifiedCell{Sheet: "S", Col: 3, Row: 1}
	d1 := QualifiedCell{Sheet: "S", Col: 4, Row: 1}

	g.Register(b1, "S", []CellAddr{{Col: 1, Row: 1}}, nil)
	g.Register(c1, "S", []CellAddr{{Col: 1, Row: 1}}, nil)
	g.Register(d1, "S", []CellAddr{{Col: 2, Row: 1}, {Col: 3, Row: 1}}, nil)

	deps := g.Dependents(QualifiedCell{Sheet: "S", Col: 1, Row: 1})
	// All three should appear exactly once.
	if len(deps) != 3 {
		t.Fatalf("expected 3 dependents, got %d: %v", len(deps), deps)
	}
	expected := []QualifiedCell{b1, c1, d1}
	if !qcEqual(deps, expected) {
		t.Errorf("expected %v, got %v", expected, deps)
	}
}

func TestRangeUnqualifiedResolution(t *testing.T) {
	g := NewDepGraph()

	// Formula on Sheet1, referencing SUM(A1:A3) without sheet qualifier.
	b1 := QualifiedCell{Sheet: "Sheet1", Col: 2, Row: 1}
	g.Register(b1, "Sheet1", nil, []RangeAddr{
		{FromCol: 1, FromRow: 1, ToCol: 1, ToRow: 3}, // unqualified
	})

	deps := g.DirectDependents(QualifiedCell{Sheet: "Sheet1", Col: 1, Row: 2})
	if !qcEqual(deps, []QualifiedCell{b1}) {
		t.Errorf("expected [B1], got %v", deps)
	}

	// Different sheet should not match.
	deps = g.DirectDependents(QualifiedCell{Sheet: "Sheet2", Col: 1, Row: 2})
	if len(deps) != 0 {
		t.Errorf("expected no dependents on Sheet2, got %v", deps)
	}
}
