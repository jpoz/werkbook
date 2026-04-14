//go:build fixtures

// Package-external fixture generator for testdata/spill_fixtures/*.xlsx.
//
// This file is compiled only when the "fixtures" build tag is set, so it
// never runs during normal `go test`. To (re)generate fixtures, run:
//
//	go test -tags=fixtures -run TestGenerateSpillFixtures ./...
//
// After running this, the new .xlsx files land in ../testdata/spill_fixtures/
// but their formula cells have no cached values yet. The next step is to
// resave them through real Excel so formula results get cached:
//
//	../testdata/resave.sh
//
// Once resaved, the cached-value parity test (TestSpillFixturesCachedValueParity
// in spill_fixtures_test.go) will compare werkbook's computation to the
// values Excel wrote.
//
// Each fixture targets a specific spill edge case. When adding a new
// fixture, pick the next available number and document *what seam it
// tests* in a comment on the builder function.
package werkbook_test

import (
	"os"
	"path/filepath"
	"strconv"
	"testing"

	"github.com/jpoz/werkbook"
)

const spillFixturesGenDir = "../testdata/spill_fixtures"

type spillFixtureBuilder struct {
	filename string
	build    func(t *testing.T) *werkbook.File
	purpose  string
}

// spillFixtureBuilders lists every fixture this generator produces.
// Each entry documents the exact seam being locked down.
var spillFixtureBuilders = []spillFixtureBuilder{
	{
		filename: "11_filter_full_column_criteria.xlsx",
		purpose:  "FILTER with full-column criteria ranges (PR #50 regression path)",
		build:    buildFixture11,
	},
	{
		filename: "12_sum_over_bounded_and_spillref.xlsx",
		purpose:  "SUM over both bounded and spill-ref (#) forms of the same anchor",
		build:    buildFixture12,
	},
	{
		filename: "13_spill_cascade_chain.xlsx",
		purpose:  "ANCHORARRAY cascade: A spills -> B reads A's spill -> C reads B",
		build:    buildFixture13,
	},
	{
		filename: "14_spill_blocker_then_cleared.xlsx",
		purpose:  "#SPILL! error where the blocker sits inside the would-be rect",
		build:    buildFixture14,
	},
	{
		filename: "15_sequence_medium_1000.xlsx",
		purpose:  "SEQUENCE(1000) -- range growth at a scale that catches perf bugs",
		build:    buildFixture15,
	},
	{
		filename: "16_sort_unique_filter.xlsx",
		purpose:  "SORT(UNIQUE(FILTER(...))) nested dynamic-array chain",
		build:    buildFixture16,
	},
	{
		filename: "17_let_with_spill.xlsx",
		purpose:  "LET binding a dynamic array and returning it",
		build:    buildFixture17,
	},
	{
		filename: "18_if_array_lifted.xlsx",
		purpose:  "IF(range>threshold, a, b) as a dynamic array",
		build:    buildFixture18,
	},
	{
		filename: "19_datevalue_in_filter.xlsx",
		purpose:  "PR #46 exact repro: DATEVALUE inside FILTER predicate",
		build:    buildFixture19,
	},
	{
		filename: "20_countifs_over_spill.xlsx",
		purpose:  "COUNTIFS with multiple criteria ranges over spill outputs",
		build:    buildFixture20,
	},
	{
		filename: "21_textjoin_over_spill.xlsx",
		purpose:  "TEXTJOIN aggregation over a FILTER spill",
		build:    buildFixture21,
	},
	{
		filename: "22_tocol_torow_reshape.xlsx",
		purpose:  "TOCOL / TOROW shape transforms over a source range",
		build:    buildFixture22,
	},
	{
		filename: "23_take_drop_slicing.xlsx",
		purpose:  "TAKE and DROP slicing a larger spilled array",
		build:    buildFixture23,
	},
	{
		filename: "24_xlookup_array_return.xlsx",
		purpose:  "XLOOKUP returning an array that spills",
		build:    buildFixture24,
	},
	{
		filename: "25_cross_sheet_defined_name_spill.xlsx",
		purpose:  "Workbook-scoped defined name referencing another sheet's spill",
		build:    buildFixture25,
	},
	{
		filename: "26_cascade_growth_then_shrink.xlsx",
		purpose:  "Cascaded spill where upstream data grows then shrinks",
		build:    buildFixture26,
	},
	{
		filename: "27_spill_anchor_at_col_row_edge.xlsx",
		purpose:  "Anchor placed at A1 (edge of the sheet) with vertical spill",
		build:    buildFixture27,
	},
	{
		filename: "28_multiple_parallel_spills.xlsx",
		purpose:  "Eight side-by-side FILTER anchors -- catches overlay rebuild issues",
		build:    buildFixture28,
	},
	{
		filename: "29_hstack_vstack_mix.xlsx",
		purpose:  "VSTACK and HSTACK composing multiple literal and range arguments",
		build:    buildFixture29,
	},
	{
		filename: "30_sumproduct_over_spill_mask.xlsx",
		purpose:  "SUMPRODUCT reading a spill and multiplying by a criteria mask",
		build:    buildFixture30,
	},
	{
		filename: "31_sumproduct_nested_if_elementwise.xlsx",
		purpose:  "SUMPRODUCT nested IF with the formula row outside the referenced rows (#VALUE! via implicit intersection)",
		build:    buildFixture31,
	},
}

// TestGenerateSpillFixtures writes each entry in spillFixtureBuilders to
// testdata/spill_fixtures/. It overwrites existing files. After writing,
// run ../testdata/resave.sh to cache formula values via real Excel.
func TestGenerateSpillFixtures(t *testing.T) {
	if err := os.MkdirAll(spillFixturesGenDir, 0o755); err != nil {
		t.Fatalf("MkdirAll: %v", err)
	}
	for _, fb := range spillFixtureBuilders {
		t.Run(fb.filename, func(t *testing.T) {
			f := fb.build(t)
			// Tell Excel to recalculate all formulas on open so it
			// caches its own computed values, not werkbook's.
			f.SetCalcProperties(werkbook.CalcProperties{FullCalcOnLoad: true})
			f.Recalculate()
			path := filepath.Join(spillFixturesGenDir, fb.filename)
			if err := f.SaveAs(path); err != nil {
				t.Fatalf("SaveAs(%s): %v", path, err)
			}
			t.Logf("wrote %s (%s)", path, fb.purpose)
		})
	}
	t.Log("NEXT: run ../testdata/resave.sh to cache computed values via Excel")
}

// ---------------------------------------------------------------------------
// Fixture builders
// ---------------------------------------------------------------------------
//
// Each function returns a fresh *werkbook.File configured to exercise one
// specific spill seam. Keep them small: 5-20 data rows and 1-5 formulas is
// plenty to lock down behavior without making resave.sh slow.

func seedBasicDataSheet(t *testing.T, rows int) *werkbook.File {
	t.Helper()
	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	mustSetValue(t, data, "A1", "Keep")
	mustSetValue(t, data, "B1", "Amount")
	mustSetValue(t, data, "C1", "Label")
	for i := 0; i < rows; i++ {
		n := i + 2
		mustSetValue(t, data, "A"+strconv.Itoa(n), i%2 == 0)
		mustSetValue(t, data, "B"+strconv.Itoa(n), float64((i+1)*10))
		mustSetValue(t, data, "C"+strconv.Itoa(n), "row-"+strconv.Itoa(i+1))
	}
	return f
}

func buildFixture11(t *testing.T) *werkbook.File {
	f := seedBasicDataSheet(t, 6)
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, out, "A1", `FILTER(Data!B:B,Data!A:A=TRUE,"none")`)
	mustSetFormula(t, out, "D1", `SUM(A1:A20)`)
	mustSetFormula(t, out, "D2", `COUNT(A1:A20)`)
	return f
}

func buildFixture12(t *testing.T) *werkbook.File {
	f := seedBasicDataSheet(t, 5)
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, out, "B2", `FILTER(Data!B2:B6,Data!A2:A6=TRUE)`)
	// Bounded read
	mustSetFormula(t, out, "D1", `SUM(B2:B10)`)
	// Spill-ref read via ANCHORARRAY
	mustSetFormula(t, out, "D2", `SUM(ANCHORARRAY(B2))`)
	return f
}

func buildFixture13(t *testing.T) *werkbook.File {
	f := seedBasicDataSheet(t, 5)
	mid, _ := f.NewSheet("Mid")
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, mid, "B2", `FILTER(Data!B2:B6,Data!A2:A6=TRUE)`)
	mustSetFormula(t, out, "B2", `ANCHORARRAY(Mid!B2)*2`)
	mustSetFormula(t, out, "D1", `SUM(B2:B20)`)
	return f
}

func buildFixture14(t *testing.T) *werkbook.File {
	f := seedBasicDataSheet(t, 5)
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, out, "B2", `FILTER(Data!B2:B6,Data!A2:A6=TRUE)`)
	// Blocker sits inside the spill rect.
	mustSetValue(t, out, "B3", "BLOCKER")
	mustSetFormula(t, out, "D1", `B2`)
	return f
}

func buildFixture15(t *testing.T) *werkbook.File {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	mustSetFormula(t, s, "A1", `SEQUENCE(1000)`)
	mustSetFormula(t, s, "C1", `SUM(A1:A1100)`)
	mustSetFormula(t, s, "C2", `COUNT(A1:A1100)`)
	return f
}

func buildFixture16(t *testing.T) *werkbook.File {
	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	values := []float64{30, 10, 20, 10, 40, 30, 50, 20}
	for i, v := range values {
		mustSetValue(t, data, "A"+strconv.Itoa(i+1), v)
	}
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, out, "A1", `SORT(UNIQUE(FILTER(Data!A1:A8,Data!A1:A8>15)))`)
	return f
}

func buildFixture17(t *testing.T) *werkbook.File {
	f := seedBasicDataSheet(t, 5)
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, out, "A1", `LET(src,FILTER(Data!B2:B6,Data!A2:A6=TRUE),src*2)`)
	return f
}

func buildFixture18(t *testing.T) *werkbook.File {
	f := seedBasicDataSheet(t, 6)
	out, _ := f.NewSheet("Out")
	// IF lifted to array: returns array of {"big","small",...}
	mustSetFormula(t, out, "A1", `IF(Data!B2:B7>30,"big","small")`)
	return f
}

func buildFixture19(t *testing.T) *werkbook.File {
	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	mustSetValue(t, data, "A1", "Date")
	mustSetValue(t, data, "B1", "Amount")
	seeds := []struct {
		date string
		amt  float64
	}{
		{"2026-01-15", 100},
		{"2026-02-20", 200},
		{"2026-03-05", 300},
		{"2026-04-10", 400},
		{"2026-05-25", 500},
	}
	for i, s := range seeds {
		n := i + 2
		mustSetValue(t, data, "A"+strconv.Itoa(n), s.date)
		mustSetValue(t, data, "B"+strconv.Itoa(n), s.amt)
	}
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, out, "A1",
		`FILTER(Data!B2:B6,DATEVALUE(Data!A2:A6)>DATEVALUE("2026-03-01"),"none")`)
	return f
}

func buildFixture20(t *testing.T) *werkbook.File {
	f := seedBasicDataSheet(t, 8)
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, out, "A1", `FILTER(Data!B2:B9,Data!A2:A9=TRUE)`)
	mustSetFormula(t, out, "B1", `COUNTIFS(A1:A20,">20")`)
	mustSetFormula(t, out, "B2", `COUNTIFS(A1:A20,">=10",A1:A20,"<=50")`)
	return f
}

func buildFixture21(t *testing.T) *werkbook.File {
	f := seedBasicDataSheet(t, 5)
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, out, "A1", `FILTER(Data!C2:C6,Data!A2:A6=TRUE)`)
	mustSetFormula(t, out, "B1", `TEXTJOIN(",",TRUE,A1:A20)`)
	return f
}

func buildFixture22(t *testing.T) *werkbook.File {
	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	for r := 1; r <= 3; r++ {
		for c := 1; c <= 3; c++ {
			ref, _ := werkbook.CoordinatesToCellName(c, r)
			mustSetValue(t, data, ref, float64(r*10+c))
		}
	}
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, out, "A1", `TOCOL(Data!A1:C3)`)
	mustSetFormula(t, out, "D1", `TOROW(Data!A1:C3)`)
	return f
}

func buildFixture23(t *testing.T) *werkbook.File {
	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	for i := 1; i <= 10; i++ {
		mustSetValue(t, data, "A"+strconv.Itoa(i), float64(i*5))
	}
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, out, "A1", `TAKE(Data!A1:A10,3)`)
	mustSetFormula(t, out, "C1", `DROP(Data!A1:A10,3)`)
	return f
}

func buildFixture24(t *testing.T) *werkbook.File {
	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	mustSetValue(t, data, "A1", "Key")
	mustSetValue(t, data, "B1", "V1")
	mustSetValue(t, data, "C1", "V2")
	keys := []string{"alpha", "beta", "gamma"}
	for i, k := range keys {
		n := i + 2
		mustSetValue(t, data, "A"+strconv.Itoa(n), k)
		mustSetValue(t, data, "B"+strconv.Itoa(n), float64((i+1)*10))
		mustSetValue(t, data, "C"+strconv.Itoa(n), float64((i+1)*100))
	}
	out, _ := f.NewSheet("Out")
	// XLOOKUP returning a 2-wide array that spills.
	mustSetFormula(t, out, "A1", `XLOOKUP("beta",Data!A2:A4,Data!B2:C4)`)
	return f
}

func buildFixture25(t *testing.T) *werkbook.File {
	f := seedBasicDataSheet(t, 5)
	mid, _ := f.NewSheet("Mid")
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, mid, "B2", `FILTER(Data!B2:B6,Data!A2:A6=TRUE)`)
	// Defined name covering the full spill column on Mid.
	if err := f.SetDefinedName(werkbook.DefinedName{
		Name:         "MidSpill",
		Value:        `Mid!$B:$B`,
		LocalSheetID: -1,
	}); err != nil {
		t.Fatalf("SetDefinedName: %v", err)
	}
	mustSetFormula(t, out, "A1", `SUM(MidSpill)`)
	mustSetFormula(t, out, "A2", `COUNT(MidSpill)`)
	return f
}

func buildFixture26(t *testing.T) *werkbook.File {
	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	// Eight rows, first four keep=TRUE, last four keep=FALSE.
	mustSetValue(t, data, "A1", "Keep")
	mustSetValue(t, data, "B1", "Amount")
	for i := 0; i < 8; i++ {
		n := i + 2
		mustSetValue(t, data, "A"+strconv.Itoa(n), i < 4)
		mustSetValue(t, data, "B"+strconv.Itoa(n), float64((i+1)*10))
	}
	mid, _ := f.NewSheet("Mid")
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, mid, "B2", `FILTER(Data!B2:B9,Data!A2:A9=TRUE)`)
	mustSetFormula(t, out, "B2", `ANCHORARRAY(Mid!B2)+100`)
	mustSetFormula(t, out, "A1", `SUM(B2:B20)`)
	return f
}

func buildFixture27(t *testing.T) *werkbook.File {
	f := werkbook.New()
	s := f.Sheet("Sheet1")
	// Anchor at A1: watches that edge position doesn't trip the overlay.
	mustSetFormula(t, s, "A1", `SEQUENCE(5)`)
	mustSetFormula(t, s, "C1", `SUM(A1:A20)`)
	return f
}

func buildFixture28(t *testing.T) *werkbook.File {
	f := werkbook.New(werkbook.FirstSheet("Source"))
	source := f.Sheet("Source")
	headers := []string{"ID", "Name", "City", "State", "Amount", "Days", "Bucket", "Flag"}
	for i, h := range headers {
		ref, _ := werkbook.CoordinatesToCellName(i+1, 1)
		mustSetValue(t, source, ref, h)
	}
	for r := 2; r <= 7; r++ {
		mustSetValue(t, source, "A"+strconv.Itoa(r), float64(r-1))
		mustSetValue(t, source, "B"+strconv.Itoa(r), "name-"+strconv.Itoa(r-1))
		mustSetValue(t, source, "C"+strconv.Itoa(r), "city-"+strconv.Itoa((r-1)%3))
		mustSetValue(t, source, "D"+strconv.Itoa(r), "state-"+strconv.Itoa((r-1)%2))
		mustSetValue(t, source, "E"+strconv.Itoa(r), float64((r-1)*100))
		mustSetValue(t, source, "F"+strconv.Itoa(r), float64((r-1)*3))
		mustSetValue(t, source, "G"+strconv.Itoa(r), "bucket-"+strconv.Itoa((r-1)%4))
		mustSetValue(t, source, "H"+strconv.Itoa(r), float64((r-1)%2))
	}
	out, _ := f.NewSheet("Out")
	for col := 1; col <= 8; col++ {
		colName := werkbook.ColumnNumberToName(col)
		anchor, _ := werkbook.CoordinatesToCellName(col, 2)
		formula := `FILTER(Source!` + colName + `2:` + colName + `7,Source!H2:H7<>0,"none")`
		mustSetFormula(t, out, anchor, formula)
	}
	return f
}

func buildFixture29(t *testing.T) *werkbook.File {
	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	for i := 1; i <= 3; i++ {
		mustSetValue(t, data, "A"+strconv.Itoa(i), float64(i))
		mustSetValue(t, data, "B"+strconv.Itoa(i), float64(i*10))
	}
	out, _ := f.NewSheet("Out")
	mustSetFormula(t, out, "A1", `VSTACK(Data!A1:A3,Data!B1:B3,100)`)
	mustSetFormula(t, out, "C1", `HSTACK(Data!A1:A3,Data!B1:B3)`)
	return f
}

func buildFixture30(t *testing.T) *werkbook.File {
	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	for i := 1; i <= 6; i++ {
		n := i + 1
		mustSetValue(t, data, "A"+strconv.Itoa(n), float64(i*10))
		mustSetValue(t, data, "B"+strconv.Itoa(n), i%2 == 0)
	}
	out, _ := f.NewSheet("Out")
	// Spill the filtered values and their doubled form, then aggregate.
	mustSetFormula(t, out, "A2", `FILTER(Data!A2:A7,Data!B2:B7=TRUE)`)
	mustSetFormula(t, out, "C1", `SUMPRODUCT(ANCHORARRAY(A2)*2)`)
	return f
}

// buildFixture31 locks down the workbook-layout case where SUMPRODUCT's nested
// IF is evaluated from row 1 while the referenced rows start at row 2. Excel
// applies implicit intersection instead of element-wise IF evaluation, so the
// formula yields #VALUE!.
func buildFixture31(t *testing.T) *werkbook.File {
	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	// A=status, B=amount, C=cost, D=rate
	mustSetValue(t, data, "A1", "Status")
	mustSetValue(t, data, "B1", "Amount")
	mustSetValue(t, data, "C1", "Cost")
	mustSetValue(t, data, "D1", "Rate")

	mustSetValue(t, data, "A2", "completed")
	mustSetValue(t, data, "B2", 5.0)
	mustSetValue(t, data, "C2", 0.0)
	mustSetValue(t, data, "D2", 1.0)

	mustSetValue(t, data, "A3", "completed")
	mustSetValue(t, data, "B3", 5.0)
	mustSetValue(t, data, "C3", 5.0)

	mustSetValue(t, data, "A4", "completed")
	mustSetValue(t, data, "B4", 10.0)
	mustSetValue(t, data, "C4", 0.0)

	out, _ := f.NewSheet("Out")
	// The formula is intentionally placed on row 1 while its array inputs begin
	// on row 2 so cached-value parity exercises the #VALUE! case.
	mustSetFormula(t, out, "A1",
		`SUMPRODUCT((Data!A2:A4="completed")*IF(Data!B2:B4-Data!C2:C4*Data!D2>0,Data!B2:B4-Data!C2:C4*Data!D2,0))`)
	return f
}
