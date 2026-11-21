package werkbook_test

import (
	"strconv"
	"testing"

	"github.com/jpoz/werkbook"
)

func TestBridgeWorkbookCompat(t *testing.T) {
	t.Run("rows_empty_filter_maps_calc_to_value", func(t *testing.T) {
		f := werkbook.New()
		s := f.Sheet("Sheet1")

		for i, region := range []string{"north", "north", "south", "south", "east"} {
			mustSetValue(t, s, "A"+strconv.Itoa(i+1), region)
		}
		mustSetFormula(t, s, "C1", `ROWS(FILTER(A1:A5,A1:A5="west"))`)

		got, err := s.GetValue("C1")
		if err != nil {
			t.Fatalf("GetValue(C1): %v", err)
		}
		if got.Type != werkbook.TypeError || got.String != "#VALUE!" {
			t.Fatalf("C1 = %#v, want #VALUE!", got)
		}
	})

	t.Run("index_selector_spills", func(t *testing.T) {
		f := werkbook.New()
		s := f.Sheet("Sheet1")

		for i, v := range []float64{10, 20, 30, 40, 50} {
			mustSetValue(t, s, "A"+strconv.Itoa(i+1), v)
		}
		mustSetFormula(t, s, "B1", `INDEX(A1:A5,SEQUENCE(5))`)

		assertSheetWants(t, s,
			numWant("B1", 10),
			numWant("B2", 20),
			numWant("B3", 30),
			numWant("B4", 40),
			numWant("B5", 50),
		)
	})

	t.Run("criteria_functions_broadcast_unique", func(t *testing.T) {
		f := werkbook.New()
		s := f.Sheet("Sheet1")

		keys := []string{"A", "B", "A", "C", "B", "A"}
		vals := []float64{10, 20, 15, 30, 25, 5}
		for i := range keys {
			row := i + 1
			mustSetValue(t, s, "A"+strconv.Itoa(row), keys[i])
			mustSetValue(t, s, "B"+strconv.Itoa(row), vals[i])
		}
		mustSetFormula(t, s, "D1", `COUNTIF(A1:A6,UNIQUE(A1:A6))`)
		mustSetFormula(t, s, "E1", `SUMIF(A1:A6,UNIQUE(A1:A6),B1:B6)`)

		assertSheetWants(t, s,
			numWant("D1", 3),
			numWant("D2", 2),
			numWant("D3", 1),
			numWant("E1", 30),
			numWant("E2", 45),
			numWant("E3", 30),
		)
	})

	t.Run("single_and_at_scalarize_mid_chain", func(t *testing.T) {
		f := werkbook.New()
		s := f.Sheet("Sheet1")

		keys := []string{"a", "b", "c", "d"}
		vals := []float64{10, 20, 30, 40}
		for i := range keys {
			row := i + 1
			mustSetValue(t, s, "A"+strconv.Itoa(row), keys[i])
			mustSetValue(t, s, "B"+strconv.Itoa(row), vals[i])
		}
		mustSetFormula(t, s, "D1", `IF(SINGLE(FILTER(A1:A4,B1:B4>15))="b","yes","no")`)
		mustSetFormula(t, s, "D2", `IF(@FILTER(A1:A4,B1:B4>15)="b","yes","no")`)

		assertSheetWants(t, s,
			strWant("D1", "yes"),
			strWant("D2", "yes"),
		)
	})
}
