package werkbook

import (
	"testing"

	"github.com/jpoz/werkbook/formula"
	"github.com/jpoz/werkbook/ooxml"
)

// Package-internal unit tests for the spill state machine in spill.go.
// These exercise pure helper functions and Cell-level flag transitions
// that are hard to observe through the public API.

func numArray(rows [][]float64) formula.Value {
	out := make([][]formula.Value, len(rows))
	for i, row := range rows {
		rr := make([]formula.Value, len(row))
		for j, v := range row {
			rr[j] = formula.NumberVal(v)
		}
		out[i] = rr
	}
	return formula.Value{Type: formula.ValueArray, Array: out}
}

func TestSpillArrayRect(t *testing.T) {
	tests := []struct {
		name                 string
		raw                  formula.Value
		anchorCol, anchorRow int
		wantToCol, wantToRow int
		wantOK               bool
	}{
		{
			name:      "scalar not an array",
			raw:       formula.NumberVal(42),
			anchorCol: 5, anchorRow: 3,
			wantOK: false,
		},
		{
			name:      "NoSpill flag blocks spill",
			raw:       formula.Value{Type: formula.ValueArray, NoSpill: true, Array: [][]formula.Value{{formula.NumberVal(1)}}},
			anchorCol: 2, anchorRow: 2,
			wantOK: false,
		},
		{
			name:      "empty array",
			raw:       numArray(nil),
			anchorCol: 1, anchorRow: 1,
			wantOK: false,
		},
		{
			name:      "array with empty row",
			raw:       formula.Value{Type: formula.ValueArray, Array: [][]formula.Value{{}}},
			anchorCol: 1, anchorRow: 1,
			wantOK: false,
		},
		{
			name:      "1x1 array spills to anchor itself",
			raw:       numArray([][]float64{{10}}),
			anchorCol: 4, anchorRow: 7,
			wantToCol: 4, wantToRow: 7,
			wantOK: true,
		},
		{
			name:      "vertical 3x1 spill",
			raw:       numArray([][]float64{{10}, {20}, {30}}),
			anchorCol: 2, anchorRow: 5,
			wantToCol: 2, wantToRow: 7,
			wantOK: true,
		},
		{
			name:      "horizontal 1x3 spill",
			raw:       numArray([][]float64{{10, 20, 30}}),
			anchorCol: 3, anchorRow: 2,
			wantToCol: 5, wantToRow: 2,
			wantOK: true,
		},
		{
			name:      "rectangular 2x3 spill",
			raw:       numArray([][]float64{{1, 2, 3}, {4, 5, 6}}),
			anchorCol: 1, anchorRow: 1,
			wantToCol: 3, wantToRow: 2,
			wantOK: true,
		},
		{
			name: "jagged array uses widest row for col bound",
			raw: formula.Value{Type: formula.ValueArray, Array: [][]formula.Value{
				{formula.NumberVal(1), formula.NumberVal(2)},
				{formula.NumberVal(3), formula.NumberVal(4), formula.NumberVal(5)},
			}},
			anchorCol: 10, anchorRow: 10,
			wantToCol: 12, wantToRow: 11,
			wantOK: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			toCol, toRow, ok := spillArrayRect(tt.raw, tt.anchorCol, tt.anchorRow)
			if ok != tt.wantOK {
				t.Fatalf("ok = %v, want %v", ok, tt.wantOK)
			}
			if !ok {
				return
			}
			if toCol != tt.wantToCol || toRow != tt.wantToRow {
				t.Fatalf("rect = (%d,%d), want (%d,%d)", toCol, toRow, tt.wantToCol, tt.wantToRow)
			}
		})
	}
}

func TestIsDynamicArrayAnchor(t *testing.T) {
	tests := []struct {
		name string
		cell *Cell
		want bool
	}{
		{name: "nil cell", cell: nil, want: false},
		{name: "no formula", cell: &Cell{}, want: false},
		{
			name: "CSE array formula is not a dynamic anchor",
			cell: &Cell{formula: "FILTER(A:A,B:B)", isArrayFormula: true, dynamicArraySpill: true},
			want: false,
		},
		{
			name: "formula without dynamicArraySpill is not an anchor",
			cell: &Cell{formula: "FILTER(A:A,B:B)", dynamicArraySpill: false},
			want: false,
		},
		{
			name: "formula with dynamicArraySpill is an anchor",
			cell: &Cell{formula: "FILTER(A:A,B:B)", dynamicArraySpill: true},
			want: true,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := isDynamicArrayAnchor(tt.cell); got != tt.want {
				t.Fatalf("isDynamicArrayAnchor = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestPublishSpillState_IdempotentWithinGen(t *testing.T) {
	f := New()
	s, err := f.NewSheet("Spill")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	f.calcGen = 5

	c := &Cell{formula: "SEQUENCE(3)", dynamicArraySpill: true}
	raw := numArray([][]float64{{1}, {2}, {3}})

	// First publish should update fields and invalidate the overlay.
	s.ensureSpillOverlay() // build with current gen
	s.publishSpillState(c, 2, 2, raw, false)
	state, ok := s.spillState(2, 2)
	if !ok {
		t.Fatalf("missing spill state for B2")
	}
	if state.gen != 5 {
		t.Fatalf("spill gen = %d, want 5", state.gen)
	}
	if state.publishedToCol != 2 || state.publishedToRow != 4 {
		t.Fatalf("publishedTo = (%d,%d), want (2,4)", state.publishedToCol, state.publishedToRow)
	}

	// Second call with identical rect must be a no-op: overlay gen stays.
	s.spill.gen = f.calcGen
	s.publishSpillState(c, 2, 2, raw, false)
	if s.spill.gen != f.calcGen {
		t.Fatalf("overlay gen was invalidated on idempotent publish")
	}

	// A call with a different rect must invalidate the overlay.
	raw2 := numArray([][]float64{{1}, {2}, {3}, {4}})
	s.publishSpillState(c, 2, 2, raw2, false)
	if s.spill.gen == f.calcGen {
		t.Fatalf("overlay gen should be invalidated on rect change")
	}
	state, _ = s.spillState(2, 2)
	if state.publishedToRow != 5 {
		t.Fatalf("publishedToRow = %d after grow, want 5", state.publishedToRow)
	}
}

func TestPublishSpillState_BlockedKeepsAttemptedButNotPublished(t *testing.T) {
	f := New()
	s, err := f.NewSheet("Spill")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	f.calcGen = 7

	c := &Cell{formula: "FILTER(A:A,B:B)", dynamicArraySpill: true}
	raw := numArray([][]float64{{1}, {2}, {3}})

	// blocked=true: attempted rect is still tracked, but published collapses.
	s.publishSpillState(c, 3, 3, raw, true)
	state, ok := s.spillState(3, 3)
	if !ok {
		t.Fatalf("missing spill state for C3")
	}
	if state.attemptedToCol != 3 || state.attemptedToRow != 5 {
		t.Fatalf("attemptedTo = (%d,%d), want (3,5)",
			state.attemptedToCol, state.attemptedToRow)
	}
	if state.publishedToCol != 3 || state.publishedToRow != 3 {
		t.Fatalf("publishedTo = (%d,%d), want (3,3) while blocked",
			state.publishedToCol, state.publishedToRow)
	}
	if !state.blocked {
		t.Fatalf("blocked = false, want true")
	}
}

func TestClearSpillState_IdempotentWhenAlreadyClear(t *testing.T) {
	f := New()
	s, err := f.NewSheet("Spill")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	// First clear on an already-clear cell must not invalidate the overlay.
	s.spill.gen = 42
	s.clearSpillState(2, 2)
	if s.spill.gen != 42 {
		t.Fatalf("overlay gen touched by no-op clear")
	}

	// Set some state, then clear, and confirm the overlay IS invalidated.
	state := s.ensureSpillState(2, 2)
	state.gen = 3
	state.publishedToCol = 7
	state.publishedToRow = 9
	state.attemptedToCol = 7
	state.attemptedToRow = 9
	s.clearSpillState(2, 2)
	if s.spill.gen == 42 {
		t.Fatalf("overlay gen should be invalidated after real clear")
	}
	if _, ok := s.spillState(2, 2); ok {
		t.Fatalf("spill state still present after clear")
	}
}

func TestSpillStateHasPublishedSpill(t *testing.T) {
	tests := []struct {
		name  string
		state *spillAnchorState
		gen   uint64
		want  bool
	}{
		{name: "nil state", state: nil, gen: 1, want: false},
		{
			name: "stale gen",
			state: &spillAnchorState{
				gen: 1, publishedToCol: 10, publishedToRow: 10,
			},
			gen:  2,
			want: false,
		},
		{
			name: "anchor-only (published == anchor)",
			state: &spillAnchorState{
				gen: 2, publishedToCol: 5, publishedToRow: 5,
			},
			gen:  2,
			want: false, // 5 > 5 is false on both axes
		},
		{
			name: "published extends column",
			state: &spillAnchorState{
				gen: 2, publishedToCol: 6, publishedToRow: 5,
			},
			gen:  2,
			want: true,
		},
		{
			name: "published extends row",
			state: &spillAnchorState{
				gen: 2, publishedToCol: 5, publishedToRow: 7,
			},
			gen:  2,
			want: true,
		},
		{
			name: "blocked",
			state: &spillAnchorState{
				gen: 2, publishedToCol: 6, publishedToRow: 7, blocked: true,
			},
			gen:  2,
			want: false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := spillStateHasPublishedSpill(tt.state, 5, 5, tt.gen)
			if got != tt.want {
				t.Fatalf("spillStateHasPublishedSpill = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestSpillStateFormulaRef(t *testing.T) {
	// Anchor at B2 (col 2, row 2).
	state := &spillAnchorState{
		gen:            3,
		publishedToCol: 4, // D
		publishedToRow: 5,
	}
	ref, ok := spillStateFormulaRef(state, 2, 2)
	if !ok {
		t.Fatalf("spillStateFormulaRef returned ok=false")
	}
	if ref != "B2:D5" {
		t.Fatalf("ref = %q, want B2:D5", ref)
	}

	// Anchor whose published bounds equal the anchor should report no ref.
	scalar := &spillAnchorState{
		gen:            3,
		publishedToCol: 2,
		publishedToRow: 2,
	}
	if _, ok := spillStateFormulaRef(scalar, 2, 2); ok {
		t.Fatalf("expected no ref for non-spilled anchor")
	}
}

func TestImportedSpillFormulaRefLivesInOverlay(t *testing.T) {
	const formulaRef = "B2:B4"
	f, err := fileFromData(&ooxml.WorkbookData{
		Sheets: []ooxml.SheetData{{
			Name: "Sheet1",
			Rows: []ooxml.RowData{{
				Num: 2,
				Cells: []ooxml.CellData{{
					Ref:            "B2",
					Value:          "1",
					Formula:        "SEQUENCE(3)",
					FormulaType:    "array",
					FormulaRef:     formulaRef,
					IsDynamicArray: true,
				}},
			}},
		}},
	}, openConfig{})
	if err != nil {
		t.Fatalf("fileFromData: %v", err)
	}
	s := f.Sheet("Sheet1")
	state, ok := s.spillState(2, 2)
	if !ok {
		t.Fatalf("missing overlay state for imported dynamic-array anchor")
	}
	if state.formulaRef != formulaRef {
		t.Fatalf("formulaRef = %q, want %q", state.formulaRef, formulaRef)
	}
}

func TestBuildWorkbookDataDoesNotClobberBlockedSpillAnchor(t *testing.T) {
	f := New(FirstSheet("Data"))
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
	r := out.rows[2]
	b2 := r.cells[2]
	if b2.value.Type != TypeError || b2.value.String != "#SPILL!" {
		t.Fatalf("B2 before buildWorkbookData = %#v, want #SPILL!", b2.value)
	}
	if b2.rawValue.Type != formula.ValueArray {
		t.Fatalf("B2 rawValue before buildWorkbookData = %#v, want ValueArray", b2.rawValue)
	}

	wbdata := f.buildWorkbookData()
	var foundB2, foundD1 bool
	for _, sd := range wbdata.Sheets {
		if sd.Name != "Out" {
			continue
		}
		for _, rd := range sd.Rows {
			for _, cd := range rd.Cells {
				switch cd.Ref {
				case "B2":
					foundB2 = true
					if cd.Type != "" || cd.Value != "" {
						t.Fatalf("WorkbookData B2 cached value = %#v, want empty while blocked", cd)
					}
					if !cd.IsDynamicArray {
						t.Fatalf("WorkbookData B2 IsDynamicArray = false, want true")
					}
					if cd.FormulaRef != "" {
						t.Fatalf("WorkbookData B2 FormulaRef = %q, want empty while blocked", cd.FormulaRef)
					}
				case "D1":
					foundD1 = true
					if cd.Type != "e" || cd.Value != "#SPILL!" {
						t.Fatalf("WorkbookData D1 = %#v, want error #SPILL!", cd)
					}
				}
			}
		}
	}
	if !foundB2 || !foundD1 {
		t.Fatalf("WorkbookData missing B2 or D1: B2=%v D1=%v", foundB2, foundD1)
	}
}
