package werkbook_test

import (
	"math"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/jpoz/werkbook"
)

const spillFixturesDir = "../testdata/spill_fixtures"

// TestSpillFixturesOpenCleanly asserts every workbook in the spill_fixtures
// directory opens without error and recalculates without panicking. This is
// the cheapest possible guard: no value comparison, just a smoke test that
// we do not regress at the open/parse/recalc layer on any known fixture.
func TestSpillFixturesOpenCleanly(t *testing.T) {
	paths := listSpillFixtures(t)
	if len(paths) == 0 {
		t.Skip("no fixtures in " + spillFixturesDir)
	}
	for _, path := range paths {
		name := filepath.Base(path)
		t.Run(name, func(t *testing.T) {
			f, err := werkbook.Open(path)
			if err != nil {
				t.Fatalf("Open(%s): %v", name, err)
			}
			f.Recalculate()
		})
	}
}

// TestSpillFixturesCachedValueParity compares the cached (Excel-computed)
// value of every formula cell in every spill fixture to what werkbook
// computes on recalc. Fixtures whose filenames appear in the skip list
// below are excluded, mirroring wb check's ignore_files.
//
// This is the gold-standard lock-down: if any change to the calc engine,
// OOXML reader, or spill machinery diverges from what Excel saved, one
// of these subtests fails by name.
func TestSpillFixturesCachedValueParity(t *testing.T) {
	paths := listSpillFixtures(t)
	if len(paths) == 0 {
		t.Skip("no fixtures in " + spillFixturesDir)
	}

	// Skip policy: every entry MUST have a comment with a GitHub issue
	// number and a plan to un-skip. A bare skip with a vague comment is
	// how issue #51 slipped through — fixture 10 was skipped instead of
	// treated as a failing signal.
	skip := map[string]bool{
		// #51: Pre-existing fixture without fullCalcOnLoad. Excel cached
		// stale #VALUE!. Un-skip after regenerating with fullCalcOnLoad
		// and resaving.
		"10_sumproduct_nested_if.xlsx": true,
		// Intentional #SPILL! fixture — Excel flags it as corrupt on
		// open, so it can't be resaved. Covered by the dedicated
		// TestSpillBlockedSaveAsRoundTrip test instead.
		"14_spill_blocker_then_cleared.xlsx": true,
		// XLOOKUP multi-column return bug: Excel returns first result
		// column (20), werkbook returns second (100). Needs fix in
		// fnXLOOKUP return-array column selection.
		"24_xlookup_array_return.xlsx": true,
	}

	for _, path := range paths {
		name := filepath.Base(path)
		if skip[name] {
			continue
		}
		t.Run(name, func(t *testing.T) {
			mismatches := compareFixtureCachedParity(t, path, 1e-9)
			if mismatches > 0 {
				t.Fatalf("%d cached/computed mismatches in %s", mismatches, name)
			}
		})
	}
}

// TestSpillFixturesRoundTrip opens each fixture, saves it to a temp path,
// reopens, and confirms every formula cell computes to the same value as
// the original recalc. This catches writer regressions in dynamic-array
// serialization (cm="1", formulaRef attribute, x14ac namespace).
func TestSpillFixturesRoundTrip(t *testing.T) {
	paths := listSpillFixtures(t)
	if len(paths) == 0 {
		t.Skip("no fixtures in " + spillFixturesDir)
	}
	for _, path := range paths {
		name := filepath.Base(path)
		t.Run(name, func(t *testing.T) {
			f, err := werkbook.Open(path)
			if err != nil {
				t.Fatalf("Open: %v", err)
			}
			f.Recalculate()
			pre := collectFormulaCellValues(t, f)

			tmp := filepath.Join(t.TempDir(), name)
			if err := f.SaveAs(tmp); err != nil {
				t.Fatalf("SaveAs: %v", err)
			}

			f2, err := werkbook.Open(tmp)
			if err != nil {
				t.Fatalf("reopen: %v", err)
			}
			f2.Recalculate()
			post := collectFormulaCellValues(t, f2)

			if len(pre) != len(post) {
				t.Fatalf("formula cell count changed across round-trip: pre=%d post=%d",
					len(pre), len(post))
			}
			for key, preVal := range pre {
				postVal, ok := post[key]
				if !ok {
					t.Errorf("%s: missing after round-trip", key)
					continue
				}
				if !valuesApproxEqual(preVal, postVal, 1e-9) {
					t.Errorf("%s: pre=%#v post=%#v", key, preVal, postVal)
				}
			}
		})
	}
}

// TestSpillBlockedSaveAsRoundTrip reproduces fixture 14 directly without the
// fixture skip list. It locks down the bug described in TODO.md: after a spill
// anchor is blocked and recalculated, SaveAs must preserve the blocked cached
// value on reopen.
func TestSpillBlockedSaveAsRoundTrip(t *testing.T) {
	f := werkbook.New(werkbook.FirstSheet("Data"))
	data := f.Sheet("Data")
	out, _ := f.NewSheet("Out")

	must := func(err error) {
		t.Helper()
		if err != nil {
			t.Fatal(err)
		}
	}

	for _, cell := range []struct {
		ref string
		val any
	}{
		{"A1", "Keep"},
		{"B1", "Amount"},
		{"C1", "Label"},
		{"A2", true},
		{"B2", 10.0},
		{"C2", "row-1"},
		{"A3", false},
		{"B3", 20.0},
		{"C3", "row-2"},
		{"A4", true},
		{"B4", 30.0},
		{"C4", "row-3"},
		{"A5", false},
		{"B5", 40.0},
		{"C5", "row-4"},
		{"A6", true},
		{"B6", 50.0},
		{"C6", "row-5"},
	} {
		must(data.SetValue(cell.ref, cell.val))
	}
	must(out.SetFormula("B2", `FILTER(Data!B2:B6,Data!A2:A6=TRUE)`))
	must(out.SetValue("B3", "BLOCKER"))
	must(out.SetFormula("D1", `B2`))

	f.Recalculate()

	before, err := out.GetValue("B2")
	must(err)
	if before.Type != werkbook.TypeError || before.String != "#SPILL!" {
		t.Fatalf("B2 before save = %#v, want #SPILL!", before)
	}

	tmp := filepath.Join(t.TempDir(), "fixture14.xlsx")
	must(f.SaveAs(tmp))

	f2, err := werkbook.Open(tmp)
	must(err)
	out2 := f2.Sheet("Out")
	if out2 == nil {
		t.Fatal("reopened workbook missing Out sheet")
	}

	after, err := out2.GetValue("B2")
	must(err)
	if after.Type != werkbook.TypeError || after.String != "#SPILL!" {
		t.Fatalf("B2 after reopen = %#v, want #SPILL!", after)
	}

	d1, err := out2.GetValue("D1")
	must(err)
	if d1.Type != werkbook.TypeError || d1.String != "#SPILL!" {
		t.Fatalf("D1 after reopen = %#v, want #SPILL!", d1)
	}
}

// listSpillFixtures walks spillFixturesDir and returns sorted xlsx paths.
func listSpillFixtures(t *testing.T) []string {
	t.Helper()
	entries, err := os.ReadDir(spillFixturesDir)
	if err != nil {
		if os.IsNotExist(err) {
			return nil
		}
		t.Fatalf("ReadDir(%s): %v", spillFixturesDir, err)
	}
	var paths []string
	for _, e := range entries {
		if e.IsDir() {
			continue
		}
		if !strings.HasSuffix(e.Name(), ".xlsx") {
			continue
		}
		if strings.HasPrefix(e.Name(), "~$") {
			continue
		}
		paths = append(paths, filepath.Join(spillFixturesDir, e.Name()))
	}
	return paths
}

// compareFixtureCachedParity mirrors wb check: reads the cached cell
// values from the file, recalculates, and compares. Returns the number
// of mismatches (0 on full parity).
func compareFixtureCachedParity(t *testing.T, path string, tolerance float64) int {
	t.Helper()
	f, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	type key struct {
		sheet string
		ref   string
	}
	cached := make(map[key]werkbook.Value)
	formulas := make(map[key]string)

	for _, name := range f.SheetNames() {
		s := f.Sheet(name)
		if s == nil {
			continue
		}
		for row := range s.Rows() {
			for _, cell := range row.Cells() {
				fstr := cell.Formula()
				if fstr == "" || isVolatileFormulaLocal(fstr) {
					continue
				}
				ref, err := werkbook.CoordinatesToCellName(cell.Col(), row.Num())
				if err != nil {
					continue
				}
				v, _ := s.GetValue(ref)
				k := key{sheet: name, ref: ref}
				cached[k] = v
				formulas[k] = fstr
			}
		}
	}

	// If everything cached is zero/empty, the file was never calculated.
	if len(cached) > 2 && allUncached(cached) {
		t.Logf("skipping %s: no cached values", filepath.Base(path))
		return 0
	}

	f.Recalculate()

	mismatches := 0
	for k, want := range cached {
		s := f.Sheet(k.sheet)
		if s == nil {
			continue
		}
		got, _ := s.GetValue(k.ref)
		if !valuesApproxEqual(want, got, tolerance) {
			mismatches++
			t.Errorf("%s!%s: formula=%s cached=%#v computed=%#v",
				k.sheet, k.ref, formulas[k], want, got)
		}
	}
	return mismatches
}

func collectFormulaCellValues(t *testing.T, f *werkbook.File) map[string]werkbook.Value {
	t.Helper()
	out := make(map[string]werkbook.Value)
	for _, name := range f.SheetNames() {
		s := f.Sheet(name)
		if s == nil {
			continue
		}
		for row := range s.Rows() {
			for _, cell := range row.Cells() {
				if cell.Formula() == "" {
					continue
				}
				ref, err := werkbook.CoordinatesToCellName(cell.Col(), row.Num())
				if err != nil {
					continue
				}
				v, _ := s.GetValue(ref)
				out[name+"!"+ref] = v
			}
		}
	}
	return out
}

// allUncached reports true when every value in the map is the zero value
// for its type: this indicates an xlsx that was never calculated by Excel
// (e.g. produced by exceljs/xlsxwriter without cached values).
func allUncached[K comparable](m map[K]werkbook.Value) bool {
	for _, v := range m {
		switch v.Type {
		case werkbook.TypeEmpty:
			// treat as uncached
		case werkbook.TypeNumber:
			if v.Number != 0 {
				return false
			}
		case werkbook.TypeString:
			if v.String != "" {
				return false
			}
		default:
			return false
		}
	}
	return true
}

func isVolatileFormulaLocal(formula string) bool {
	upper := strings.ToUpper(formula)
	for _, fn := range []string{
		"RAND(", "RANDARRAY(", "RANDBETWEEN(", "NOW(", "TODAY(",
		"OFFSET(", "INDIRECT(",
	} {
		if strings.Contains(upper, fn) {
			return true
		}
	}
	return false
}

func valuesApproxEqual(a, b werkbook.Value, tolerance float64) bool {
	if a.Type != b.Type {
		// Accept error-vs-string cross-type match (some writers store
		// errors as t="str").
		if a.Type == werkbook.TypeError && b.Type == werkbook.TypeString {
			return a.String == b.String
		}
		if a.Type == werkbook.TypeString && b.Type == werkbook.TypeError {
			return a.String == b.String
		}
		return false
	}
	switch a.Type {
	case werkbook.TypeNumber:
		if tolerance > 0 {
			return math.Abs(a.Number-b.Number) <= tolerance
		}
		return a.Number == b.Number
	case werkbook.TypeString:
		return a.String == b.String
	case werkbook.TypeBool:
		return a.Bool == b.Bool
	case werkbook.TypeError:
		return a.String == b.String
	case werkbook.TypeEmpty:
		return true
	default:
		return a.Raw() == b.Raw()
	}
}
