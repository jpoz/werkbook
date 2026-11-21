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

	raw := numArray([][]float64{{1}, {2}, {3}})
	plan := newSpillPlan(raw, 2, 2)

	// First publish should update fields and invalidate the overlay.
	s.ensureSpillOverlay() // build with current gen
	s.publishSpillState(2, 2, plan)
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
	s.publishSpillState(2, 2, plan)
	if s.spill.gen != f.calcGen {
		t.Fatalf("overlay gen was invalidated on idempotent publish")
	}

	// A call with a different rect must invalidate the overlay.
	raw2 := numArray([][]float64{{1}, {2}, {3}, {4}})
	s.publishSpillState(2, 2, newSpillPlan(raw2, 2, 2))
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

	raw := numArray([][]float64{{1}, {2}, {3}})
	plan := newSpillPlan(raw, 3, 3)
	plan.Blocked = true
	plan.PublishedToCol = 3
	plan.PublishedToRow = 3

	// blocked=true: attempted rect is still tracked, but published collapses.
	s.publishSpillState(3, 3, plan)
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

func TestCellEvalOutcomeSeparatesBlockedDisplayFromRawSpill(t *testing.T) {
	f := New(FirstSheet("Spill"))
	s := f.Sheet("Spill")

	if err := s.SetFormula("B2", `SEQUENCE(3)`); err != nil {
		t.Fatalf("SetFormula(B2): %v", err)
	}
	if err := s.SetValue("B3", "blocker"); err != nil {
		t.Fatalf("SetValue(B3): %v", err)
	}

	cell := s.rows[2].cells[2]
	outcome, err := s.evalCellOutcome(cell, 2, 2)
	if err != nil {
		t.Fatalf("evalCellOutcome: %v", err)
	}
	if outcome.Display.Type != formula.ValueError || outcome.Display.Err != formula.ErrValSPILL {
		t.Fatalf("Display = %#v, want #SPILL!", outcome.Display)
	}
	if outcome.Spill == nil || !outcome.Spill.Blocked {
		t.Fatalf("Spill = %#v, want blocked spill plan", outcome.Spill)
	}
	raw := outcome.RawValue()
	if raw.Type != formula.ValueArray {
		t.Fatalf("RawValue.Type = %v, want ValueArray", raw.Type)
	}
	if len(raw.Array) != 3 || len(raw.Array[0]) != 1 {
		t.Fatalf("RawValue.Array shape = %v, want 3x1", raw.Array)
	}
	if raw.Array[1][0].Type != formula.ValueNumber || raw.Array[1][0].Num != 2 {
		t.Fatalf("RawValue second row = %#v, want number 2", raw.Array[1][0])
	}
}

func TestSpillOverlayIndexLookup(t *testing.T) {
	f := New(FirstSheet("Spill"))
	s := f.Sheet("Spill")

	if err := s.SetFormula("B2", `SEQUENCE(2,2,10,1)`); err != nil {
		t.Fatalf("SetFormula(B2): %v", err)
	}
	f.Recalculate()

	overlay := s.ensureSpillOverlay()
	if got := overlay.index.lookup(2, 2); got != nil {
		t.Fatalf("lookup(B2) = %#v, want nil for anchor cell", got)
	}
	for _, cell := range []struct {
		col, row int
		want     float64
	}{
		{3, 2, 11},
		{2, 3, 12},
		{3, 3, 13},
	} {
		span := overlay.index.lookup(cell.col, cell.row)
		if span == nil {
			t.Fatalf("lookup(%d,%d) = nil, want spill span", cell.col, cell.row)
		}
		got, ok := s.spillFormulaValueAt(cell.col, cell.row)
		if !ok {
			t.Fatalf("spillFormulaValueAt(%d,%d) = missing", cell.col, cell.row)
		}
		if got.Type != formula.ValueNumber || got.Num != cell.want {
			t.Fatalf("spillFormulaValueAt(%d,%d) = %#v, want %g", cell.col, cell.row, got, cell.want)
		}
	}
}

func TestSpillOverlayIndexLookupSkipsJaggedHole(t *testing.T) {
	f := New(FirstSheet("Spill"))
	s := f.Sheet("Spill")
	f.calcGen = 1

	raw1 := formula.Value{Type: formula.ValueArray, Array: [][]formula.Value{
		{formula.NumberVal(1), formula.NumberVal(2), formula.NumberVal(3)},
		{formula.NumberVal(4)},
	}}
	raw2 := formula.Value{Type: formula.ValueArray, Array: [][]formula.Value{{
		formula.NumberVal(10),
		formula.NumberVal(11),
	}}}

	if s.rows[2] == nil {
		s.rows[2] = &Row{num: 2, cells: make(map[int]*Cell)}
	}
	s.rows[2].cells[2] = &Cell{
		col:               2,
		formula:           "JAGGED()",
		dynamicArraySpill: true,
		rawValue:          raw1,
		rawCachedGen:      f.calcGen,
	}
	if s.rows[3] == nil {
		s.rows[3] = &Row{num: 3, cells: make(map[int]*Cell)}
	}
	s.rows[3].cells[3] = &Cell{
		col:               3,
		formula:           "ROWPAIR()",
		dynamicArraySpill: true,
		rawValue:          raw2,
		rawCachedGen:      f.calcGen,
	}

	s.publishSpillState(2, 2, newSpillPlan(raw1, 2, 2))
	s.publishSpillState(3, 3, newSpillPlan(raw2, 3, 3))

	overlay := s.ensureSpillOverlay()
	if got := overlay.index.lookup(4, 3); got == nil {
		t.Fatalf("lookup(D3) = nil, want spill span from C3")
	}
	got, ok := s.spillFormulaValueAt(4, 3)
	if !ok {
		t.Fatalf("spillFormulaValueAt(D3) = missing")
	}
	if got.Type != formula.ValueNumber || got.Num != 11 {
		t.Fatalf("spillFormulaValueAt(D3) = %#v, want 11", got)
	}
}

func TestRefreshSpillAnchorsForPointSkipsAnchorsAfterPoint(t *testing.T) {
	f := New(FirstSheet("Spill"))
	s := f.Sheet("Spill")

	if err := s.SetFormula("A1", `SEQUENCE(3,1)`); err != nil {
		t.Fatalf("SetFormula(A1): %v", err)
	}
	if err := s.SetFormula("C1", `SEQUENCE(3,1)`); err != nil {
		t.Fatalf("SetFormula(C1): %v", err)
	}
	if err := s.SetFormula("A3", `SEQUENCE(3,1)`); err != nil {
		t.Fatalf("SetFormula(A3): %v", err)
	}
	gen := f.calcGen

	if !s.refreshSpillAnchorsForPoint(2, 2) {
		t.Fatalf("refreshSpillAnchorsForPoint(B2) = false, want true")
	}

	a1 := s.rows[1].cells[1]
	c1 := s.rows[1].cells[3]
	a3 := s.rows[3].cells[1]
	if a1.rawCachedGen != gen {
		t.Fatalf("A1 rawCachedGen = %d, want %d", a1.rawCachedGen, gen)
	}
	if c1.rawCachedGen == gen {
		t.Fatalf("C1 was refreshed for B2, want skipped because it is after the point column")
	}
	if a3.rawCachedGen == gen {
		t.Fatalf("A3 was refreshed for B2, want skipped because it is after the point row")
	}
}

func TestRefreshSpillAnchorsForRangeSkipsAnchorsAfterRange(t *testing.T) {
	f := New(FirstSheet("Spill"))
	s := f.Sheet("Spill")

	if err := s.SetFormula("A1", `SEQUENCE(4,2)`); err != nil {
		t.Fatalf("SetFormula(A1): %v", err)
	}
	if err := s.SetFormula("D1", `SEQUENCE(4,1)`); err != nil {
		t.Fatalf("SetFormula(D1): %v", err)
	}
	if err := s.SetFormula("A5", `SEQUENCE(4,2)`); err != nil {
		t.Fatalf("SetFormula(A5): %v", err)
	}
	gen := f.calcGen

	req := rangeMaterializationRequest{
		sheet:   "Spill",
		fromCol: 2,
		fromRow: 2,
		toCol:   3,
		toRow:   4,
	}
	if !s.refreshSpillAnchorsForRange(req, req.toCol) {
		t.Fatalf("refreshSpillAnchorsForRange(B2:C4) = false, want true")
	}

	a1 := s.rows[1].cells[1]
	d1 := s.rows[1].cells[4]
	a5 := s.rows[5].cells[1]
	if a1.rawCachedGen != gen {
		t.Fatalf("A1 rawCachedGen = %d, want %d", a1.rawCachedGen, gen)
	}
	if d1.rawCachedGen == gen {
		t.Fatalf("D1 was refreshed for B2:C4, want skipped because it is after the range column")
	}
	if a5.rawCachedGen == gen {
		t.Fatalf("A5 was refreshed for B2:C4, want skipped because it is after the range row")
	}
}

func TestSpillBoundsTrackRecalculation(t *testing.T) {
	f := New(FirstSheet("Data"))
	data := f.Sheet("Data")
	spill, err := f.NewSheet("Spill")
	if err != nil {
		t.Fatalf("NewSheet(Spill): %v", err)
	}

	for _, cell := range []struct {
		ref string
		val any
	}{
		{"A2", true},
		{"B2", 10.0},
		{"A3", true},
		{"B3", 20.0},
		{"A4", false},
		{"B4", 30.0},
	} {
		if err := data.SetValue(cell.ref, cell.val); err != nil {
			t.Fatalf("SetValue(%s): %v", cell.ref, err)
		}
	}
	if err := spill.SetFormula("B2", `FILTER(Data!B2:B4,Data!A2:A4)`); err != nil {
		t.Fatalf("SetFormula(B2): %v", err)
	}

	f.Recalculate()
	toCol, toRow, ok := spill.SpillBounds(2, 2)
	if !ok || toCol != 2 || toRow != 3 {
		t.Fatalf("SpillBounds after initial recalc = (%d,%d,%v), want (2,3,true)", toCol, toRow, ok)
	}

	if err := data.SetValue("A4", true); err != nil {
		t.Fatalf("SetValue(A4): %v", err)
	}
	f.Recalculate()
	toCol, toRow, ok = spill.SpillBounds(2, 2)
	if !ok || toCol != 2 || toRow != 4 {
		t.Fatalf("SpillBounds after growth = (%d,%d,%v), want (2,4,true)", toCol, toRow, ok)
	}

	if err := data.SetValue("A3", false); err != nil {
		t.Fatalf("SetValue(A3): %v", err)
	}
	if err := data.SetValue("A4", false); err != nil {
		t.Fatalf("SetValue(A4 second): %v", err)
	}
	f.Recalculate()
	if _, _, ok := spill.SpillBounds(2, 2); ok {
		t.Fatalf("SpillBounds after shrink-to-anchor reported a spill, want none")
	}
}

func TestNonDynamicRangeRefDisplayUsesImplicitIntersection(t *testing.T) {
	f := New(FirstSheet("data"))
	data := f.Sheet("data")
	results, err := f.NewSheet("results")
	if err != nil {
		t.Fatalf("NewSheet(results): %v", err)
	}
	for i, row := range []struct {
		a float64
		b float64
	}{
		{1, 10},
		{2, 20},
		{3, 30},
		{4, 40},
		{5, 50},
	} {
		refA, _ := CoordinatesToCellName(1, i+2)
		refB, _ := CoordinatesToCellName(2, i+2)
		if err := data.SetValue(refA, row.a); err != nil {
			t.Fatalf("SetValue(%s): %v", refA, err)
		}
		if err := data.SetValue(refB, row.b); err != nil {
			t.Fatalf("SetValue(%s): %v", refB, err)
		}
	}
	if err := results.SetFormula("B2", `OFFSET(data!A2:B6, 0, 0)`); err != nil {
		t.Fatalf("SetFormula(B2): %v", err)
	}
	results.rows[2].cells[2].dynamicArraySpill = false

	f.Recalculate()

	got, err := results.GetValue("B2")
	if err != nil {
		t.Fatalf("GetValue(B2): %v", err)
	}
	if got.Type != TypeNumber || got.Number != 10 {
		t.Fatalf("B2 = %#v, want 10 via implicit intersection", got)
	}
}
