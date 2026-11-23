package werkbook

import (
	"fmt"
	"testing"
)

// newRowSumsWorkbook builds the shape of a usual data sheet: rows of nine
// numbers with one SUM over the row in column J.
func newRowSumsWorkbook(tb testing.TB, rows int) *File {
	tb.Helper()
	f := New()
	s := f.Sheet("Sheet1")
	f.BeginBatchWrite()
	for r := 1; r <= rows; r++ {
		for c := 1; c <= 9; c++ {
			ref, _ := CoordinatesToCellName(c, r)
			if err := s.SetValue(ref, float64(r+c)); err != nil {
				tb.Fatal(err)
			}
		}
		if err := s.SetFormula(fmt.Sprintf("J%d", r), fmt.Sprintf("SUM(A%d:I%d)", r, r)); err != nil {
			tb.Fatal(err)
		}
	}
	f.EndBatchWrite()
	return f
}

// BenchmarkRecalculateRowSums measures a full recalculation. Its cost must
// grow linearly with the number of rows: each formula reads nine cells.
func BenchmarkRecalculateRowSums(b *testing.B) {
	for _, rows := range []int{2000, 8000} {
		b.Run(fmt.Sprintf("rows=%d", rows), func(b *testing.B) {
			f := newRowSumsWorkbook(b, rows)
			f.Recalculate()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				f.Recalculate()
			}
		})
	}
}

// BenchmarkEditThenRecalculateRowSums measures what an editor does for each
// committed cell: one write, then a recalculation.
func BenchmarkEditThenRecalculateRowSums(b *testing.B) {
	for _, rows := range []int{2000, 8000} {
		b.Run(fmt.Sprintf("rows=%d", rows), func(b *testing.B) {
			f := newRowSumsWorkbook(b, rows)
			s := f.Sheet("Sheet1")
			f.Recalculate()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				if err := s.SetValue("A1", float64(i)); err != nil {
					b.Fatal(err)
				}
				f.Recalculate()
			}
		})
	}
}
