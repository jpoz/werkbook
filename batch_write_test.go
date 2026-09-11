package werkbook_test

import (
	"fmt"
	"testing"

	"github.com/jpoz/werkbook"
)

// fanOutWorkbook builds a Data sheet of `rows` rows and a Calc sheet where each of
// `formulas` cells aggregates the whole Data column with a full-column ref, so every
// Data write has every Calc cell as a dependent.
func fanOutWorkbook(t testing.TB, rows, formulas int) *werkbook.File {
	f := werkbook.New()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	calc, err := f.NewSheet("Calc")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	for i := 1; i <= formulas; i++ {
		if err := calc.SetFormula(fmt.Sprintf("A%d", i), "SUM(Data!$A:$A)"); err != nil {
			t.Fatalf("SetFormula: %v", err)
		}
		if err := calc.SetFormula(fmt.Sprintf("B%d", i), fmt.Sprintf("A%d*2", i)); err != nil {
			t.Fatalf("SetFormula: %v", err)
		}
	}
	return f
}

// num reads a cell and returns its numeric value.
func num(t testing.TB, s *werkbook.Sheet, cell string) float64 {
	v, err := s.GetValue(cell)
	if err != nil {
		t.Fatalf("GetValue %s: %v", cell, err)
	}
	return v.Number
}

func writeRows(t testing.TB, f *werkbook.File, rows int) {
	data := f.Sheet("Data")
	for r := 1; r <= rows; r++ {
		if err := data.SetValue(fmt.Sprintf("A%d", r), 1); err != nil {
			t.Fatalf("SetValue: %v", err)
		}
	}
}

func TestBatchWriteValuesVisibleAfterRecalculate(t *testing.T) {
	f := fanOutWorkbook(t, 0, 5)
	f.Recalculate()
	if got := num(t, f.Sheet("Calc"), "A1"); got != 0 {
		t.Fatalf("before writes: A1 = %v, want 0", got)
	}

	f.BeginBatchWrite()
	writeRows(t, f, 10)
	f.EndBatchWrite()
	f.Recalculate()

	calc := f.Sheet("Calc")
	if got := num(t, calc, "A5"); got != 10 {
		t.Fatalf("after batch: A5 = %v, want 10", got)
	}
	if got := num(t, calc, "B5"); got != 20 {
		t.Fatalf("after batch: B5 = %v, want 20 (second-hop dependent)", got)
	}
}

func TestBatchWriteValuesVisibleToLazyGetValue(t *testing.T) {
	f := fanOutWorkbook(t, 0, 3)
	f.Recalculate()
	calc := f.Sheet("Calc")
	_ = num(t, calc, "B3") // populate caches

	f.BeginBatchWrite()
	writeRows(t, f, 4)
	f.EndBatchWrite()

	// No Recalculate: lazy reads must still see the new data through calcGen.
	if got := num(t, calc, "A1"); got != 4 {
		t.Fatalf("lazy A1 = %v, want 4", got)
	}
	if got := num(t, calc, "B3"); got != 8 {
		t.Fatalf("lazy B3 = %v, want 8", got)
	}
}

func TestBatchWriteNestsAndRestoresInvalidation(t *testing.T) {
	f := fanOutWorkbook(t, 0, 2)
	f.BeginBatchWrite()
	f.BeginBatchWrite()
	f.EndBatchWrite()
	writeRows(t, f, 2)
	f.EndBatchWrite()
	f.EndBatchWrite() // unbalanced extra call must be a no-op

	// Outside the batch, a write must invalidate as before.
	if err := f.Sheet("Data").SetValue("A3", 5); err != nil {
		t.Fatalf("SetValue: %v", err)
	}
	if got := num(t, f.Sheet("Calc"), "B2"); got != 14 {
		t.Fatalf("B2 = %v, want 14", got)
	}
}

func TestBatchWriteFormulaOverwriteInsideBatch(t *testing.T) {
	f := fanOutWorkbook(t, 0, 1)
	writeRows(t, f, 3)
	f.Recalculate()
	calc := f.Sheet("Calc")
	if got := num(t, calc, "B1"); got != 6 {
		t.Fatalf("B1 = %v, want 6", got)
	}
	f.BeginBatchWrite()
	if err := calc.SetFormula("A1", "SUM(Data!$A:$A)+100"); err != nil {
		t.Fatalf("SetFormula: %v", err)
	}
	f.EndBatchWrite()
	if got := num(t, calc, "B1"); got != 206 {
		t.Fatalf("B1 after formula change in batch = %v, want 206", got)
	}
}

func BenchmarkBulkWriteUnbatched(b *testing.B) {
	for i := 0; i < b.N; i++ {
		f := fanOutWorkbook(b, 0, 500)
		writeRows(b, f, 200)
	}
}

func BenchmarkBulkWriteBatched(b *testing.B) {
	for i := 0; i < b.N; i++ {
		f := fanOutWorkbook(b, 0, 500)
		f.BeginBatchWrite()
		writeRows(b, f, 200)
		f.EndBatchWrite()
	}
}
