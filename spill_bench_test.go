package werkbook_test

import (
	"fmt"
	"testing"

	"github.com/jpoz/werkbook"
)

var spillBenchmarkSink float64

func BenchmarkSpillPointLookup(b *testing.B) {
	for _, rows := range []int{100, 1000, 5000} {
		b.Run(fmt.Sprintf("rows=%d", rows), func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				b.StopTimer()
				f := buildSpillPointLookupWorkbook(b, rows)
				b.StartTimer()
				f.Recalculate()
				v, err := f.Sheet("Calc").GetValue("A1")
				if err != nil {
					b.Fatal(err)
				}
				spillBenchmarkSink = v.Number
			}
		})
	}
}

func BenchmarkSpillRangeAggregate(b *testing.B) {
	for _, rows := range []int{100, 1000, 5000} {
		b.Run(fmt.Sprintf("rows=%d", rows), func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				b.StopTimer()
				f := buildSpillRangeAggregateWorkbook(b, rows)
				b.StartTimer()
				f.Recalculate()
				v, err := f.Sheet("Calc").GetValue("A1")
				if err != nil {
					b.Fatal(err)
				}
				spillBenchmarkSink = v.Number
			}
		})
	}
}

func BenchmarkSpillMatchFullColumn(b *testing.B) {
	for _, rows := range []int{100, 1000, 5000} {
		b.Run(fmt.Sprintf("rows=%d", rows), func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				b.StopTimer()
				f := buildSpillMatchFullColumnWorkbook(b, rows)
				b.StartTimer()
				f.Recalculate()
				v, err := f.Sheet("Calc").GetValue("A1")
				if err != nil {
					b.Fatal(err)
				}
				spillBenchmarkSink = v.Number
			}
		})
	}
}

func BenchmarkSpillManyAnchors(b *testing.B) {
	for _, anchors := range []int{10, 50, 100} {
		b.Run(fmt.Sprintf("anchors=%d", anchors), func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				b.StopTimer()
				f := buildManySpillAnchorsWorkbook(b, anchors, 200)
				b.StartTimer()
				f.Recalculate()
				v, err := f.Sheet("Calc").GetValue("A1")
				if err != nil {
					b.Fatal(err)
				}
				spillBenchmarkSink = v.Number
			}
		})
	}
}

func BenchmarkSpillLazyAfterEdit(b *testing.B) {
	for _, rows := range []int{100, 1000, 5000} {
		b.Run(fmt.Sprintf("rows=%d", rows), func(b *testing.B) {
			_, data, calc := buildLazySpillEditWorkbook(b, rows)
			toggleCell := benchCellRef(b, 1, rows+1)
			include := false

			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				if err := data.SetValue(toggleCell, include); err != nil {
					b.Fatal(err)
				}
				sum, err := calc.GetValue("A1")
				if err != nil {
					b.Fatal(err)
				}
				count, err := calc.GetValue("B1")
				if err != nil {
					b.Fatal(err)
				}
				spillBenchmarkSink = sum.Number + count.Number
				include = !include
			}
		})
	}
}

func buildSpillPointLookupWorkbook(tb testing.TB, rows int) *werkbook.File {
	tb.Helper()

	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	for row := 2; row <= rows+1; row++ {
		benchMustSetValue(tb, data, benchCellRef(tb, 1, row), true)
		benchMustSetValue(tb, data, benchCellRef(tb, 2, row), float64(row-1))
	}

	spill := benchMustNewSheet(tb, f, "Spill")
	benchMustSetValue(tb, spill, "B1", "Filtered")
	benchMustSetFormula(tb, spill, "B2", fmt.Sprintf("FILTER(Data!B2:B%d,Data!A2:A%d)", rows+1, rows+1))

	calc := benchMustNewSheet(tb, f, "Calc")
	for row := 1; row <= rows; row++ {
		benchMustSetFormula(tb, calc, benchCellRef(tb, 1, row), fmt.Sprintf("Spill!B%d*2", row+1))
	}

	return f
}

func buildSpillRangeAggregateWorkbook(tb testing.TB, rows int) *werkbook.File {
	tb.Helper()

	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	for row := 2; row <= rows+1; row++ {
		benchMustSetValue(tb, data, benchCellRef(tb, 1, row), true)
		benchMustSetValue(tb, data, benchCellRef(tb, 2, row), float64(row-1))
	}

	spill := benchMustNewSheet(tb, f, "Spill")
	benchMustSetValue(tb, spill, "B1", "Filtered")
	benchMustSetFormula(tb, spill, "B2", fmt.Sprintf("FILTER(Data!B2:B%d,Data!A2:A%d)", rows+1, rows+1))

	calc := benchMustNewSheet(tb, f, "Calc")
	formulas := []string{
		`SUM(Spill!B:B)`,
		`COUNT(Spill!B:B)`,
		`AVERAGE(Spill!B:B)`,
		`MIN(Spill!B:B)`,
		`MAX(Spill!B:B)`,
		`SUMIF(Spill!B:B,">0")`,
		`COUNTIF(Spill!B:B,">0")`,
		fmt.Sprintf(`MATCH(%d,Spill!B:B,0)`, rows),
	}
	for i, formula := range formulas {
		benchMustSetFormula(tb, calc, benchCellRef(tb, 1, i+1), formula)
	}

	return f
}

func buildSpillMatchFullColumnWorkbook(tb testing.TB, rows int) *werkbook.File {
	tb.Helper()

	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	for row := 2; row <= rows+1; row++ {
		benchMustSetValue(tb, data, benchCellRef(tb, 1, row), true)
		benchMustSetValue(tb, data, benchCellRef(tb, 2, row), float64(row-1))
	}

	spill := benchMustNewSheet(tb, f, "Spill")
	benchMustSetValue(tb, spill, "B1", "Filtered")
	benchMustSetFormula(tb, spill, "B2", fmt.Sprintf("FILTER(Data!B2:B%d,Data!A2:A%d)", rows+1, rows+1))

	calc := benchMustNewSheet(tb, f, "Calc")
	benchMustSetFormula(tb, calc, "A1", fmt.Sprintf("MATCH(%d,Spill!B:B,0)", rows))

	return f
}

func buildManySpillAnchorsWorkbook(tb testing.TB, anchors, rows int) *werkbook.File {
	tb.Helper()

	f := werkbook.New(werkbook.FirstSheet("Spill"))
	spill := f.Sheet("Spill")
	calc := benchMustNewSheet(tb, f, "Calc")
	targetRow := rows
	if targetRow < 2 {
		targetRow = 2
	}

	for i := 0; i < anchors; i++ {
		col := 1 + i*2
		benchMustSetFormula(tb, spill, benchCellRef(tb, col, 1), fmt.Sprintf("SEQUENCE(%d,1,%d,1)", rows, i*1000+1))
		benchMustSetFormula(tb, calc, benchCellRef(tb, 1, i+1), fmt.Sprintf("Spill!%s", benchCellRef(tb, col, targetRow)))
	}

	return f
}

func buildLazySpillEditWorkbook(tb testing.TB, rows int) (*werkbook.File, *werkbook.Sheet, *werkbook.Sheet) {
	tb.Helper()

	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	for row := 2; row <= rows+1; row++ {
		benchMustSetValue(tb, data, benchCellRef(tb, 1, row), true)
		benchMustSetValue(tb, data, benchCellRef(tb, 2, row), float64(row-1))
	}

	spill := benchMustNewSheet(tb, f, "Spill")
	benchMustSetFormula(tb, spill, "B2", fmt.Sprintf("FILTER(Data!B2:B%d,Data!A2:A%d)", rows+1, rows+1))

	calc := benchMustNewSheet(tb, f, "Calc")
	benchMustSetFormula(tb, calc, "A1", "SUM(Spill!B:B)")
	benchMustSetFormula(tb, calc, "B1", "COUNT(Spill!B:B)")

	f.Recalculate()
	return f, data, calc
}

func benchMustNewSheet(tb testing.TB, f *werkbook.File, name string) *werkbook.Sheet {
	tb.Helper()
	s, err := f.NewSheet(name)
	if err != nil {
		tb.Fatal(err)
	}
	return s
}

func benchMustSetValue(tb testing.TB, s *werkbook.Sheet, cell string, value any) {
	tb.Helper()
	if err := s.SetValue(cell, value); err != nil {
		tb.Fatalf("SetValue(%s): %v", cell, err)
	}
}

func benchMustSetFormula(tb testing.TB, s *werkbook.Sheet, cell string, formula string) {
	tb.Helper()
	if err := s.SetFormula(cell, formula); err != nil {
		tb.Fatalf("SetFormula(%s): %v", cell, err)
	}
}

func benchCellRef(tb testing.TB, col, row int) string {
	tb.Helper()
	ref, err := werkbook.CoordinatesToCellName(col, row)
	if err != nil {
		tb.Fatal(err)
	}
	return ref
}
