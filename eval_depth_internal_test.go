package werkbook

import (
	"fmt"
	"sort"
	"testing"

	"github.com/jpoz/werkbook/formula"
)

// resolveLastCellPrecedents builds a chain in column A with the given
// formula for each row, runs resolvePrecedentsIteratively on the last cell,
// and returns the sheet. It does not evaluate the last cell.
func resolveLastCellPrecedents(t *testing.T, n int, formulaFor func(row int) string) *Sheet {
	t.Helper()
	f := New()
	s := f.Sheet("Sheet1")
	f.BeginBatchWrite()
	if err := s.SetValue("A1", 1); err != nil {
		t.Fatal(err)
	}
	for row := 2; row <= n; row++ {
		if err := s.SetFormula(fmt.Sprintf("A%d", row), formulaFor(row)); err != nil {
			t.Fatal(err)
		}
	}
	f.EndBatchWrite()

	last := s.rows[n].cells[1]
	// The caller of resolvePrecedentsIteratively has the root in evaluation.
	key := cellKey{sheet: s.name, col: 1, row: n}
	f.evaluating = map[cellKey]bool{key: true}
	f.resolvePrecedentsIteratively(s, last, 1, n)
	delete(f.evaluating, key)
	return s
}

func requireChainResolved(t *testing.T, s *Sheet, n int) {
	t.Helper()
	gen := s.file.calcGen
	for row := 2; row < n; row++ {
		c := s.rows[row].cells[1]
		if c.needsEval(gen) {
			t.Fatalf("A%d is not resolved", row)
		}
		if c.value.Type != TypeNumber || c.value.Number != float64(row) {
			t.Fatalf("A%d = %#v, want %d", row, c.value, row)
		}
	}
	if !s.rows[n].cells[1].needsEval(gen) {
		t.Fatalf("A%d (the root) must stay not evaluated", n)
	}
}

func TestResolvePrecedentsIteratively_CellRefs(t *testing.T) {
	const n = 3_000
	s := resolveLastCellPrecedents(t, n, func(row int) string {
		return fmt.Sprintf("A%d+1", row-1)
	})
	requireChainResolved(t, s, n)
}

func TestResolvePrecedentsIteratively_Ranges(t *testing.T) {
	const n = 3_000
	s := resolveLastCellPrecedents(t, n, func(row int) string {
		return fmt.Sprintf("SUM(A%d:A%d)+1", row-1, row-1)
	})
	requireChainResolved(t, s, n)
}

func TestForEachCellInRange(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	for _, ref := range []string{"A1", "B2", "C3", "D4", "B9", "E2"} {
		if err := s.SetValue(ref, 1); err != nil {
			t.Fatal(err)
		}
	}

	tests := []struct {
		name string
		rng  formula.RangeAddr
		want []string
	}{
		{"small range uses direct lookup", formula.RangeAddr{FromCol: 2, FromRow: 2, ToCol: 3, ToRow: 3}, []string{"B2", "C3"}},
		{"full column scans stored rows", formula.RangeAddr{FromCol: 2, FromRow: 1, ToCol: 2, ToRow: MaxRows}, []string{"B2", "B9"}},
		{"full row scans stored cells", formula.RangeAddr{FromCol: 1, FromRow: 2, ToCol: MaxColumns, ToRow: 2}, []string{"B2", "E2"}},
		{"single cell", formula.RangeAddr{FromCol: 4, FromRow: 4, ToCol: 4, ToRow: 4}, []string{"D4"}},
		{"empty area", formula.RangeAddr{FromCol: 6, FromRow: 6, ToCol: 8, ToRow: 8}, nil},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var got []string
			s.forEachCellInRange(tt.rng, func(_ *Cell, col, row int) {
				ref, err := CoordinatesToCellName(col, row)
				if err != nil {
					t.Fatal(err)
				}
				got = append(got, ref)
			})
			sort.Strings(got)
			if fmt.Sprint(got) != fmt.Sprint(tt.want) {
				t.Fatalf("got %v, want %v", got, tt.want)
			}
		})
	}
}
