package werkbook_test

import (
	"bytes"
	"fmt"
	"testing"

	"github.com/jpoz/werkbook"
)

// A chain of this length caused a fatal stack overflow before the engine
// resolved deep precedent chains with an explicit stack.
const deepChainLen = 200_000

func buildChain(t *testing.T, n int, formulaFor func(row int) string) (*werkbook.File, *werkbook.Sheet) {
	t.Helper()
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	f.BeginBatchWrite()
	defer f.EndBatchWrite()
	if err := s.SetValue("A1", 1); err != nil {
		t.Fatal(err)
	}
	for row := 2; row <= n; row++ {
		if err := s.SetFormula(fmt.Sprintf("A%d", row), formulaFor(row)); err != nil {
			t.Fatal(err)
		}
	}
	return f, s
}

func requireNumber(t *testing.T, s *werkbook.Sheet, cell string, want float64) {
	t.Helper()
	v, err := s.GetValue(cell)
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != werkbook.TypeNumber || v.Number != want {
		t.Fatalf("%s = %#v, want number %v", cell, v, want)
	}
}

func TestDeepChain_CellRefs(t *testing.T) {
	_, s := buildChain(t, deepChainLen, func(row int) string {
		return fmt.Sprintf("A%d+1", row-1)
	})
	requireNumber(t, s, fmt.Sprintf("A%d", deepChainLen), deepChainLen)
	// A cell in the middle of the chain is correct too.
	requireNumber(t, s, "A1000", 1000)
}

// The chain is short because each range read scans the sheet for its bounds,
// which makes a long range chain slow. The chain is still deeper than the
// recursion limit.
func TestDeepChain_Ranges(t *testing.T) {
	const n = 2_000
	_, s := buildChain(t, n, func(row int) string {
		return fmt.Sprintf("SUM(A%d:A%d)+1", row-1, row-1)
	})
	requireNumber(t, s, fmt.Sprintf("A%d", n), n)
}

// References that a formula computes at run time are not known before the
// evaluation. They must not use the Go stack either.
func TestDeepChain_Indirect(t *testing.T) {
	const n = 60_000
	_, s := buildChain(t, n, func(int) string {
		return `INDIRECT("A"&(ROW()-1))+1`
	})
	requireNumber(t, s, fmt.Sprintf("A%d", n), n)
}

// Each cell reads all the cells above it in one range. The cells that an
// attempt finds stale are on the explicit stack already.
func TestDeepChain_CumulativeRanges(t *testing.T) {
	const n = 3_000
	_, s := buildChain(t, n, func(row int) string {
		return fmt.Sprintf("COUNT(A$1:A%d)+1", row-1)
	})
	requireNumber(t, s, fmt.Sprintf("A%d", n), n)
}

// Each cell reads all the cells below it, which is the worst order for the
// explicit stack: each attempt replaces the entries of the cells below.
func TestDeepChain_ReverseCumulativeRanges(t *testing.T) {
	const n = 3_000
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	f.BeginBatchWrite()
	if err := s.SetValue(fmt.Sprintf("A%d", n), 1); err != nil {
		t.Fatal(err)
	}
	for row := 1; row < n; row++ {
		if err := s.SetFormula(fmt.Sprintf("A%d", row), fmt.Sprintf("COUNT(A%d:A$%d)+1", row+1, n)); err != nil {
			t.Fatal(err)
		}
	}
	f.EndBatchWrite()
	requireNumber(t, s, "A1", n)
	requireNumber(t, s, "A1500", n-1499)
}

// deepChainLengths are the lengths of the chain that deepShallowCase puts
// above the probe cell. The lengths near the recursion limit of 256 put the
// probe cell and its precedents at each position relative to the point where
// the engine stops the recursion.
var deepChainLengths = []int{253, 254, 255, 256, 257, 258, 259, 2_000}

// deepShallowCase builds the same cells in two workbooks. In the first one, a
// chain in column Z reads the probe cell, so the cells are evaluated at or
// below the recursion limit. In the second one, the probe cell is read
// directly. All the given cells must have the same value in the two
// workbooks.
func deepShallowCase(t *testing.T, cells map[string]any, probe string, check []string) {
	t.Helper()
	for _, n := range deepChainLengths {
		deepShallowCaseN(t, n, cells, probe, check)
	}
}

func deepShallowCaseN(t *testing.T, n int, cells map[string]any, probe string, check []string) {
	t.Helper()
	build := func(deep bool) *werkbook.Sheet {
		f := werkbook.New()
		s := f.Sheet("Sheet1")
		f.BeginBatchWrite()
		defer f.EndBatchWrite()
		for ref, v := range cells {
			var err error
			if formulaText, ok := v.(string); ok && len(formulaText) > 0 && formulaText[0] == '=' {
				err = s.SetFormula(ref, formulaText)
			} else {
				err = s.SetValue(ref, v)
			}
			if err != nil {
				t.Fatal(err)
			}
		}
		if deep {
			if err := s.SetFormula("Z1", probe+"*1"); err != nil {
				t.Fatal(err)
			}
			for row := 2; row <= n; row++ {
				if err := s.SetFormula(fmt.Sprintf("Z%d", row), fmt.Sprintf("Z%d", row-1)); err != nil {
					t.Fatal(err)
				}
			}
		}
		return s
	}

	deep := build(true)
	if _, err := deep.GetValue(fmt.Sprintf("Z%d", n)); err != nil {
		t.Fatal(err)
	}
	shallow := build(false)
	if _, err := shallow.GetValue(probe); err != nil {
		t.Fatal(err)
	}
	for _, ref := range check {
		got, err := deep.GetValue(ref)
		if err != nil {
			t.Fatal(err)
		}
		want, err := shallow.GetValue(ref)
		if err != nil {
			t.Fatal(err)
		}
		if got != want {
			t.Errorf("chain %d, %s: deep evaluation gives %#v, shallow evaluation gives %#v", n, ref, got, want)
		}
	}
}

// A reference that is used only for its address must not cause an evaluation
// of the cell. Before, the deep path evaluated B1 while A1 was in evaluation,
// and B1 cached a false circular-reference error.
func TestDeepChain_AddressOnlyReference(t *testing.T) {
	cells := map[string]any{
		"A1": "=COLUMN(B1)",
		"B1": "=A1+1",
	}
	deepShallowCase(t, cells, "A1", []string{"A1", "B1"})

	// Make sure that the values are the correct ones, not only equal.
	for _, n := range deepChainLengths {
		f := werkbook.New()
		s := f.Sheet("Sheet1")
		f.BeginBatchWrite()
		for ref, v := range cells {
			if err := s.SetFormula(ref, v.(string)); err != nil {
				t.Fatal(err)
			}
		}
		if err := s.SetFormula("Z1", "A1*1"); err != nil {
			t.Fatal(err)
		}
		for row := 2; row <= n; row++ {
			if err := s.SetFormula(fmt.Sprintf("Z%d", row), fmt.Sprintf("Z%d", row-1)); err != nil {
				t.Fatal(err)
			}
		}
		f.EndBatchWrite()
		requireNumber(t, s, fmt.Sprintf("Z%d", n), 2)
		requireNumber(t, s, "A1", 2)
		requireNumber(t, s, "B1", 3)
	}
}

func TestDeepChain_SameResultsAsShallow(t *testing.T) {
	tests := []struct {
		name  string
		cells map[string]any
		probe string
		check []string
	}{
		{
			name:  "row and rows use only the address",
			cells: map[string]any{"A1": "=ROW(B5)+ROWS(B1:B3)", "B1": "=A1+1", "B5": "=A1+2"},
			probe: "A1",
			check: []string{"A1", "B1", "B5"},
		},
		{
			name:  "branch that is not taken",
			cells: map[string]any{"A1": "=IF(C1>0,10,B1)", "B1": "=A1+1", "C1": "=1+1"},
			probe: "A1",
			check: []string{"A1", "B1", "C1"},
		},
		{
			name:  "direct circular reference",
			cells: map[string]any{"A1": "=B1+1", "B1": "=A1+1"},
			probe: "A1",
			check: []string{"A1", "B1"},
		},
		{
			name:  "self reference",
			cells: map[string]any{"A1": "=A1+1"},
			probe: "A1",
			check: []string{"A1"},
		},
		{
			name:  "many stale precedents in separate reads",
			cells: map[string]any{"A1": "=B1+B2+B3+SUM(B1:B3)", "B1": "=1+1", "B2": "=B1*2", "B3": "=B2+B1"},
			probe: "A1",
			check: []string{"A1", "B1", "B2", "B3"},
		},
		{
			name:  "spill follower",
			cells: map[string]any{"A1": "=E3*2", "E1": "=SEQUENCE(3)"},
			probe: "A1",
			check: []string{"A1", "E1", "E2", "E3"},
		},
		{
			name:  "range over a spill",
			cells: map[string]any{"A1": "=SUM(E1:E5)", "E1": "=SEQUENCE(4)*B1", "B1": "=1+1"},
			probe: "A1",
			check: []string{"A1", "E1", "E4", "B1"},
		},
		{
			name:  "spill anchor operator",
			cells: map[string]any{"A1": "=SUM(E1#)", "E1": "=SEQUENCE(4)*B1", "B1": "=1+2"},
			probe: "A1",
			check: []string{"A1", "E1", "E4"},
		},
		{
			name:  "blocked spill",
			cells: map[string]any{"A1": "=E1", "E1": "=SEQUENCE(3)", "E2": "blocker"},
			probe: "A1",
			check: []string{"A1", "E1", "E2", "E3"},
		},
		{
			name:  "formula with an error",
			cells: map[string]any{"A1": "=B1+1", "B1": "=NOSUCHFUNCTION(1)"},
			probe: "A1",
			check: []string{"A1", "B1"},
		},
		{
			name:  "full column range",
			cells: map[string]any{"A1": "=SUM(B:B)", "B1": "=1+1", "B2": "=B1*3", "B7": 5},
			probe: "A1",
			check: []string{"A1", "B1", "B2"},
		},
		{
			name:  "offset reference",
			cells: map[string]any{"A1": "=OFFSET(B1,1,0)+1", "B1": "=A1+1", "B2": "=2*3"},
			probe: "A1",
			check: []string{"A1", "B1", "B2"},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			deepShallowCase(t, tt.cells, tt.probe, tt.check)
		})
	}
}

func TestDeepChain_Recalculate(t *testing.T) {
	f, s := buildChain(t, deepChainLen, func(row int) string {
		return fmt.Sprintf("A%d+1", row-1)
	})
	f.Recalculate()
	requireNumber(t, s, fmt.Sprintf("A%d", deepChainLen), deepChainLen)

	// A change at the start of the chain goes through to the end.
	if err := s.SetValue("A1", 11); err != nil {
		t.Fatal(err)
	}
	f.Recalculate()
	requireNumber(t, s, fmt.Sprintf("A%d", deepChainLen), deepChainLen+10)
}

func TestDeepChain_CrossSheet(t *testing.T) {
	const n = 50_000
	f := werkbook.New()
	a := f.Sheet("Sheet1")
	b, err := f.NewSheet("Other")
	if err != nil {
		t.Fatal(err)
	}
	f.BeginBatchWrite()
	if err := a.SetValue("A1", 1); err != nil {
		t.Fatal(err)
	}
	for row := 2; row <= n; row++ {
		// The chain alternates between the two sheets.
		if err := b.SetFormula(fmt.Sprintf("A%d", row), fmt.Sprintf("Sheet1!A%d+1", row-1)); err != nil {
			t.Fatal(err)
		}
		if err := a.SetFormula(fmt.Sprintf("A%d", row), fmt.Sprintf("Other!A%d+1", row)); err != nil {
			t.Fatal(err)
		}
	}
	f.EndBatchWrite()
	requireNumber(t, a, fmt.Sprintf("A%d", n), float64(2*n-1))
}

func TestDeepChain_SaveAndOpen(t *testing.T) {
	const n = 50_000
	f, _ := buildChain(t, n, func(row int) string {
		return fmt.Sprintf("A%d+1", row-1)
	})
	var buf bytes.Buffer
	if _, err := f.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	g, err := werkbook.OpenReader(&buf)
	if err != nil {
		t.Fatal(err)
	}
	g.Recalculate()
	requireNumber(t, g.Sheet("Sheet1"), fmt.Sprintf("A%d", n), n)
}

// A reference cycle that is deeper than the recursion limit must give an
// error value. It must not hang or overflow the stack.
func TestDeepChain_Cycle(t *testing.T) {
	const n = 20_000
	f, s := buildChain(t, n, func(row int) string {
		return fmt.Sprintf("A%d+1", row-1)
	})
	f.BeginBatchWrite()
	if err := s.SetFormula("A1", fmt.Sprintf("A%d+1", n)); err != nil {
		t.Fatal(err)
	}
	f.EndBatchWrite()
	v, err := s.GetValue(fmt.Sprintf("A%d", n))
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != werkbook.TypeError {
		t.Fatalf("A%d = %#v, want an error value for a circular reference", n, v)
	}
}

// Wide dependency graphs go through the same path: many cells read one deep
// chain, and each cell in the chain is evaluated one time.
func TestDeepChain_SharedPrecedents(t *testing.T) {
	const n = 30_000
	f, s := buildChain(t, n, func(row int) string {
		return fmt.Sprintf("A%d+1", row-1)
	})
	f.BeginBatchWrite()
	for row := 1; row <= n; row++ {
		if err := s.SetFormula(fmt.Sprintf("B%d", row), fmt.Sprintf("A%d*2", row)); err != nil {
			t.Fatal(err)
		}
	}
	if err := s.SetFormula("C1", fmt.Sprintf("SUM(B1:B%d)", n)); err != nil {
		t.Fatal(err)
	}
	f.EndBatchWrite()
	requireNumber(t, s, "C1", float64(n)*float64(n+1))
}
