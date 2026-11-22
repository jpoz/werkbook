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
// recursion limit. TestResolvePrecedentsIteratively_Ranges checks that range
// precedents are resolved without recursion.
func TestDeepChain_Ranges(t *testing.T) {
	const n = 2_000
	_, s := buildChain(t, n, func(row int) string {
		return fmt.Sprintf("SUM(A%d:A%d)+1", row-1, row-1)
	})
	requireNumber(t, s, fmt.Sprintf("A%d", n), n)
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
