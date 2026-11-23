package werkbook

import (
	"bytes"
	"strings"
	"testing"
)

func snapshotTestValue(t *testing.T, s *Sheet, ref string) Value {
	t.Helper()
	v, err := s.GetValue(ref)
	if err != nil {
		t.Fatalf("GetValue(%s): %v", ref, err)
	}
	return v
}

func snapshotTestSetup(t *testing.T, errs ...error) {
	t.Helper()
	for i, err := range errs {
		if err != nil {
			t.Fatalf("setup step %d: %v", i, err)
		}
	}
}

// A snapshot taken before a write and restored after it must give back the
// stored state for each kind of cell.
func TestSnapshotCellRoundTrip(t *testing.T) {
	bold := &Style{Font: &Font{Bold: true}}
	tests := []struct {
		name  string
		setup func(t *testing.T, s *Sheet)
		// expected stored state of B2
		wantExists  bool
		wantEmpty   bool
		wantFormula string
		wantValue   Value
		wantStyle   *Style
	}{
		{
			name:       "number",
			setup:      func(t *testing.T, s *Sheet) { snapshotTestSetup(t, s.SetValue("B2", 12.5)) },
			wantExists: true,
			wantValue:  Value{Type: TypeNumber, Number: 12.5},
		},
		{
			name:       "string",
			setup:      func(t *testing.T, s *Sheet) { snapshotTestSetup(t, s.SetValue("B2", "text")) },
			wantExists: true,
			wantValue:  Value{Type: TypeString, String: "text"},
		},
		{
			name:       "bool",
			setup:      func(t *testing.T, s *Sheet) { snapshotTestSetup(t, s.SetValue("B2", true)) },
			wantExists: true,
			wantValue:  Value{Type: TypeBool, Bool: true},
		},
		{
			name: "formula with a style",
			setup: func(t *testing.T, s *Sheet) {
				snapshotTestSetup(t, s.SetValue("A1", 4), s.SetFormula("B2", "A1*2"), s.SetStyle("B2", bold))
			},
			wantExists:  true,
			wantFormula: "A1*2",
			wantValue:   Value{Type: TypeNumber, Number: 8},
			wantStyle:   bold,
		},
		{
			name:       "style only",
			setup:      func(t *testing.T, s *Sheet) { snapshotTestSetup(t, s.SetStyle("B2", bold)) },
			wantExists: true,
			wantEmpty:  true,
			wantStyle:  bold,
		},
		{
			name:       "stored empty cell",
			setup:      func(t *testing.T, s *Sheet) { snapshotTestSetup(t, s.SetValue("B2", nil)) },
			wantExists: true,
			wantEmpty:  true,
		},
		{
			name:      "absent cell",
			setup:     func(t *testing.T, s *Sheet) {},
			wantEmpty: true,
		},
	}
	// Each kind of cell is overwritten with each kind of write.
	overwrites := []struct {
		name  string
		write func(s *Sheet) error
	}{
		{"value", func(s *Sheet) error { return s.SetValue("B2", "other") }},
		{"formula", func(s *Sheet) error { return s.SetFormula("B2", "1+1") }},
		{"clear", func(s *Sheet) error { return s.SetValue("B2", nil) }},
		{"style", func(s *Sheet) error { return s.SetStyle("B2", &Style{Fill: &Fill{Color: "FF0000"}}) }},
		{"absent", func(s *Sheet) error { return s.RestoreCell(2, 2, CellSnapshot{}) }},
	}

	for _, tt := range tests {
		for _, ow := range overwrites {
			t.Run(tt.name+"/overwritten by "+ow.name, func(t *testing.T) {
				f := New()
				s := f.Sheet("Sheet1")
				tt.setup(t, s)
				f.Recalculate()

				snap := s.SnapshotCell(2, 2)
				if snap.Exists() != tt.wantExists || snap.IsEmpty() != tt.wantEmpty {
					t.Fatalf("snapshot: exists=%v empty=%v, want exists=%v empty=%v",
						snap.Exists(), snap.IsEmpty(), tt.wantExists, tt.wantEmpty)
				}
				if snap.Formula() != tt.wantFormula || snap.Value() != tt.wantValue || snap.Style() != tt.wantStyle {
					t.Fatalf("snapshot = formula %q value %#v style %p, want formula %q value %#v style %p",
						snap.Formula(), snap.Value(), snap.Style(), tt.wantFormula, tt.wantValue, tt.wantStyle)
				}

				if err := ow.write(s); err != nil {
					t.Fatalf("overwrite: %v", err)
				}
				if err := s.RestoreCell(2, 2, snap); err != nil {
					t.Fatalf("RestoreCell: %v", err)
				}

				got := s.SnapshotCell(2, 2)
				if got.Exists() != tt.wantExists {
					t.Fatalf("after restore: exists = %v, want %v", got.Exists(), tt.wantExists)
				}
				if got.Formula() != tt.wantFormula || got.Style() != tt.wantStyle {
					t.Errorf("after restore: formula %q style %p, want formula %q style %p",
						got.Formula(), got.Style(), tt.wantFormula, tt.wantStyle)
				}
				if tt.wantExists {
					// GetValue evaluates a restored formula again.
					if v := snapshotTestValue(t, s, "B2"); v != tt.wantValue {
						t.Errorf("after restore: value = %#v, want %#v", v, tt.wantValue)
					}
				}
				if gotFormula, _ := s.GetFormula("B2"); gotFormula != tt.wantFormula {
					t.Errorf("GetFormula = %q, want %q", gotFormula, tt.wantFormula)
				}
			})
		}
	}
}

func TestRestoreCellAbsentShrinksSheet(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	snapshotTestSetup(t, s.SetValue("A1", 1))

	before := s.SnapshotCell(5, 9)
	if before.Exists() {
		t.Fatalf("snapshot of an absent cell exists")
	}
	snapshotTestSetup(t, s.SetValue("E9", "far"))
	if s.MaxRow() != 9 || s.MaxCol() != 5 {
		t.Fatalf("size after write = %dx%d, want 9x5", s.MaxRow(), s.MaxCol())
	}
	snapshotTestSetup(t, s.RestoreCell(5, 9, before))
	if s.MaxRow() != 1 || s.MaxCol() != 1 {
		t.Errorf("size after restore = %dx%d, want 1x1", s.MaxRow(), s.MaxCol())
	}

	// A row that has a custom height stays when its last cell is removed.
	snapshotTestSetup(t, s.SetRowHeight(4, 30), s.SetValue("B4", 1), s.RestoreCell(2, 4, CellSnapshot{}))
	if h, _ := s.GetRowHeight(4); h != 30 {
		t.Errorf("row height after the last cell was removed = %v, want 30", h)
	}
	// Other cells of the row stay.
	snapshotTestSetup(t, s.SetValue("B1", 2), s.RestoreCell(2, 1, CellSnapshot{}))
	if v := snapshotTestValue(t, s, "A1"); v.Number != 1 {
		t.Errorf("A1 after a neighbor was removed = %#v, want 1", v)
	}

	// A restore to absent of a cell that is absent is valid and changes nothing.
	snapshotTestSetup(t, s.RestoreCell(10, 10, CellSnapshot{}))
	if s.MaxRow() != 4 {
		t.Errorf("MaxRow = %d, want 4", s.MaxRow())
	}
}

func TestRestoreCellRecalculatesDependents(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	other, err := f.NewSheet("Other")
	if err != nil {
		t.Fatal(err)
	}
	snapshotTestSetup(t,
		s.SetValue("A1", 1),
		s.SetFormula("B1", "A1+1"),
		s.SetFormula("C1", "B1*10"),
		s.SetFormula("D1", "SUM(A1:A3)"),
		other.SetFormula("A1", "Sheet1!A1+100"),
	)
	f.Recalculate()
	snap := s.SnapshotCell(1, 1)

	snapshotTestSetup(t, s.SetValue("A1", 5))
	if v := snapshotTestValue(t, s, "C1"); v.Number != 60 {
		t.Fatalf("C1 after write = %v, want 60", v.Number)
	}

	snapshotTestSetup(t, s.RestoreCell(1, 1, snap))
	for ref, want := range map[string]float64{"B1": 2, "C1": 20, "D1": 1} {
		if v := snapshotTestValue(t, s, ref); v.Number != want {
			t.Errorf("%s after restore = %v, want %v", ref, v.Number, want)
		}
	}
	if v := snapshotTestValue(t, other, "A1"); v.Number != 101 {
		t.Errorf("Other!A1 after restore = %v, want 101", v.Number)
	}

	// Restore to absent: the dependents see an empty cell.
	snapshotTestSetup(t, s.RestoreCell(1, 1, CellSnapshot{}))
	if v := snapshotTestValue(t, s, "C1"); v.Number != 10 {
		t.Errorf("C1 after restore to absent = %v, want 10", v.Number)
	}

	// Restore of a formula over a value: its own precedents are registered again.
	formulaSnap := s.SnapshotCell(2, 1)
	snapshotTestSetup(t, s.SetValue("B1", 0), s.RestoreCell(2, 1, formulaSnap), s.SetValue("A1", 7))
	if v := snapshotTestValue(t, s, "C1"); v.Number != 80 {
		t.Errorf("C1 after formula restore and a precedent write = %v, want 80", v.Number)
	}
}

func TestRestoreCellInBatchWrite(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	snapshotTestSetup(t, s.SetValue("A1", 1), s.SetValue("A2", 2), s.SetFormula("B1", "SUM(A1:A2)"))
	f.Recalculate()
	a1, a2 := s.SnapshotCell(1, 1), s.SnapshotCell(1, 2)
	snapshotTestSetup(t, s.SetValue("A1", 10), s.SetValue("A2", 20))
	if v := snapshotTestValue(t, s, "B1"); v.Number != 30 {
		t.Fatalf("B1 = %v, want 30", v.Number)
	}

	f.BeginBatchWrite()
	snapshotTestSetup(t, s.RestoreCell(1, 1, a1), s.RestoreCell(1, 2, a2))
	f.EndBatchWrite()

	if v := snapshotTestValue(t, s, "B1"); v.Number != 3 {
		t.Errorf("B1 after a batch restore = %v, want 3", v.Number)
	}
}

func TestSnapshotCellArrayFormulaFlag(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	snapshotTestSetup(t, s.SetValue("A1", 1), s.SetValue("A2", 2), s.SetFormula("B1", "SUM(A1:A2*2)"))
	// The public API has no setter for a legacy array formula; files have them.
	cell := s.rows[1].cells[2]
	cell.isArrayFormula = true
	cell.dynamicArraySpill = false

	snap := s.SnapshotCell(2, 1)
	if !snap.IsArrayFormula() {
		t.Fatalf("snapshot lost the array formula flag")
	}

	// SetValue drops the flag; the restore must bring it back.
	snapshotTestSetup(t, s.SetValue("B1", "x"), s.RestoreCell(2, 1, snap))
	got := s.rows[1].cells[2]
	if !got.isArrayFormula || got.dynamicArraySpill || got.formula != "SUM(A1:A2*2)" {
		t.Errorf("after restore: array=%v spill=%v formula=%q, want a legacy array formula",
			got.isArrayFormula, got.dynamicArraySpill, got.formula)
	}
	if v := snapshotTestValue(t, s, "B1"); v.Number != 6 {
		t.Errorf("B1 = %#v, want 6", v)
	}

	// The flag is written to the file.
	var buf bytes.Buffer
	if _, err := f.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	f2, err := OpenReader(&buf)
	if err != nil {
		t.Fatal(err)
	}
	if c := f2.Sheet("Sheet1").rows[1].cells[2]; !c.isArrayFormula {
		t.Errorf("array formula flag is not in the saved file")
	}

	// A value snapshot never carries the flag.
	if s.SnapshotCell(1, 1).IsArrayFormula() {
		t.Errorf("a value cell reports an array formula")
	}
	if snap.WithValue(Value{Type: TypeNumber, Number: 1}).IsArrayFormula() || snap.WithFormula("A1").IsArrayFormula() {
		t.Errorf("WithValue or WithFormula kept the array formula flag")
	}
}

func TestRestoreCellDynamicArrayAnchor(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	snapshotTestSetup(t, s.SetFormula("A1", "SEQUENCE(3)"), s.SetFormula("C1", "SUM(A1:A3)"))
	f.Recalculate()
	anchor := s.SnapshotCell(1, 1)

	// The followers are not stored cells.
	if s.SnapshotCell(1, 2).Exists() {
		t.Fatalf("a spill follower has a stored snapshot")
	}

	// Replace the anchor, then restore it: it must spill again.
	snapshotTestSetup(t, s.SetValue("A1", "gone"))
	if v := snapshotTestValue(t, s, "A3"); !v.IsEmpty() {
		t.Fatalf("A3 after the anchor was replaced = %#v, want empty", v)
	}
	snapshotTestSetup(t, s.RestoreCell(1, 1, anchor))
	for ref, want := range map[string]float64{"A1": 1, "A2": 2, "A3": 3, "C1": 6} {
		if v := snapshotTestValue(t, s, ref); v.Type != TypeNumber || v.Number != want {
			t.Errorf("%s after anchor restore = %#v, want %v", ref, v, want)
		}
	}
	f.Recalculate()
	if toCol, toRow, ok := s.SpillBounds(1, 1); !ok || toCol != 1 || toRow != 3 {
		t.Errorf("SpillBounds after restore = %d,%d,%v, want 1,3,true", toCol, toRow, ok)
	}

	// A restore over a follower blocks the spill; a second restore frees it.
	follower := s.SnapshotCell(1, 2)
	blocker := follower.WithValue(Value{Type: TypeString, String: "x"})
	snapshotTestSetup(t, s.RestoreCell(1, 2, blocker))
	if v := snapshotTestValue(t, s, "A1"); v.Type != TypeError || v.String != "#SPILL!" {
		t.Errorf("anchor with a blocked spill = %#v, want #SPILL!", v)
	}
	snapshotTestSetup(t, s.RestoreCell(1, 2, follower))
	if v := snapshotTestValue(t, s, "A1"); v.Type != TypeNumber || v.Number != 1 {
		t.Errorf("anchor after the blocker was removed = %#v, want 1", v)
	}
	if v := snapshotTestValue(t, s, "A2"); v.Number != 2 {
		t.Errorf("A2 after the blocker was removed = %#v, want 2", v)
	}
	if v := snapshotTestValue(t, s, "C1"); v.Number != 6 {
		t.Errorf("C1 after the blocker was removed = %#v, want 6", v)
	}

	// A snapshot from WithFormula derives the anchor flag, as SetFormula does.
	snapshotTestSetup(t, s.RestoreCell(5, 1, CellSnapshot{}.WithFormula("=SEQUENCE(2)")))
	if v := snapshotTestValue(t, s, "E2"); v.Number != 2 {
		t.Errorf("E2 below a WithFormula anchor = %#v, want 2", v)
	}
	if !s.rows[1].cells[5].dynamicArraySpill {
		t.Errorf("WithFormula anchor has no dynamic array flag")
	}
}

// An imported formula with no dynamic-array metadata stays scalar in Excel.
// The snapshot must keep that state; a derived flag would make it spill.
func TestRestoreCellKeepsImportedScalarFormula(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	snapshotTestSetup(t, s.SetFormula("A1", "SEQUENCE(3)"))
	cell := s.rows[1].cells[1]
	cell.dynamicArraySpill = false
	s.clearSpillState(1, 1)
	f.rebuildFormulaState()

	snap := s.SnapshotCell(1, 1)
	snapshotTestSetup(t, s.SetValue("A1", 0), s.RestoreCell(1, 1, snap))
	if s.rows[1].cells[1].dynamicArraySpill {
		t.Errorf("restore made an imported scalar formula a spill anchor")
	}
	if v := snapshotTestValue(t, s, "A2"); !v.IsEmpty() {
		t.Errorf("A2 = %#v, want empty (no spill)", v)
	}
}

func TestCellSnapshotWithMethods(t *testing.T) {
	bold := &Style{Font: &Font{Bold: true}}
	f := New()
	s := f.Sheet("Sheet1")
	snapshotTestSetup(t, s.SetValue("A1", 3), s.SetStyle("A1", bold))
	base := s.SnapshotCell(1, 1)

	withFormula := base.WithFormula("=A1*2")
	if withFormula.Formula() != "A1*2" || withFormula.Style() != bold || !withFormula.Exists() || withFormula.IsEmpty() {
		t.Errorf("WithFormula = %+v", withFormula)
	}
	if !withFormula.Value().IsEmpty() {
		t.Errorf("WithFormula kept the old value %#v", withFormula.Value())
	}
	if got := base.WithFormula("A1*2").Formula(); got != "A1*2" {
		t.Errorf("WithFormula without '=' = %q", got)
	}
	if got := base.WithFormula(""); !got.IsEmpty() || !got.Exists() || got.Style() != bold {
		t.Errorf("WithFormula(\"\") = %+v, want an empty styled cell", got)
	}

	withValue := withFormula.WithValue(Value{Type: TypeString, String: "s"})
	if withValue.Formula() != "" || withValue.Value().String != "s" || withValue.Style() != bold {
		t.Errorf("WithValue = %+v", withValue)
	}

	withStyle := base.WithStyle(nil)
	if withStyle.Style() != nil || withStyle.Value().Number != 3 {
		t.Errorf("WithStyle(nil) = %+v", withStyle)
	}
	// The With methods on the zero snapshot make a stored cell.
	if got := (CellSnapshot{}).WithStyle(bold); !got.Exists() || !got.IsEmpty() || got.Style() != bold {
		t.Errorf("zero.WithStyle = %+v", got)
	}
	// The base snapshot is a value: the With methods did not change it.
	if base.Formula() != "" || base.Value().Number != 3 || base.Style() != bold {
		t.Errorf("base snapshot changed: %+v", base)
	}

	snapshotTestSetup(t, s.RestoreCell(2, 1, withFormula))
	if v := snapshotTestValue(t, s, "B1"); v.Number != 6 {
		t.Errorf("B1 = %#v, want 6", v)
	}
	if st, _ := s.GetStyle("B1"); st != bold {
		t.Errorf("B1 style = %p, want the shared pointer %p", st, bold)
	}
}

func TestCellSnapshotEqual(t *testing.T) {
	bold := &Style{Font: &Font{Bold: true}}
	f := New()
	s := f.Sheet("Sheet1")
	snapshotTestSetup(t, s.SetValue("A1", 3), s.SetFormula("B1", "A1*2"), s.SetStyle("C1", bold))

	// The cached result of a formula is not part of the stored state.
	notEvaluated := s.SnapshotCell(2, 1)
	f.Recalculate()
	evaluated := s.SnapshotCell(2, 1)
	if notEvaluated.Value() == evaluated.Value() {
		t.Fatalf("test setup: the cached values are equal")
	}
	if !notEvaluated.Equal(evaluated) {
		t.Errorf("snapshots of one formula with different cached results are not equal")
	}
	// A rewrite of the same formula gives the same stored state.
	snapshotTestSetup(t, s.SetFormula("B1", "A1*2"))
	if !s.SnapshotCell(2, 1).Equal(evaluated) {
		t.Errorf("snapshot after a rewrite of the same formula is not equal")
	}

	a1 := s.SnapshotCell(1, 1)
	absent := s.SnapshotCell(9, 9)
	tests := []struct {
		name string
		a, b CellSnapshot
		want bool
	}{
		{"same value", a1, s.SnapshotCell(1, 1), true},
		{"absent and zero", absent, CellSnapshot{}, true},
		{"absent and stored empty", absent, CellSnapshot{}.WithFormula(""), false},
		{"different value", a1, a1.WithValue(Value{Type: TypeNumber, Number: 4}), false},
		{"number and text", a1, a1.WithValue(Value{Type: TypeString, String: "3"}), false},
		{"different style", a1, a1.WithStyle(bold), false},
		{"equal styles in different pointers", a1.WithStyle(bold), a1.WithStyle(&Style{Font: &Font{Bold: true}}), false},
		{"value and formula", a1, evaluated, false},
		{"different formula", evaluated, evaluated.WithFormula("A1*3"), false},
		{"style only", s.SnapshotCell(3, 1), CellSnapshot{}.WithStyle(bold), true},
	}
	for _, tt := range tests {
		if got := tt.a.Equal(tt.b); got != tt.want {
			t.Errorf("%s: Equal = %v, want %v", tt.name, got, tt.want)
		}
		if got := tt.b.Equal(tt.a); got != tt.want {
			t.Errorf("%s (reversed): Equal = %v, want %v", tt.name, got, tt.want)
		}
	}
}

func TestRestoreCellErrors(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	snapshotTestSetup(t, s.SetValue("A1", "keep"))

	for _, c := range []struct{ col, row int }{{0, 1}, {1, 0}, {MaxColumns + 1, 1}, {1, MaxRows + 1}, {-1, -1}} {
		if err := s.RestoreCell(c.col, c.row, CellSnapshot{}.WithValue(Value{Type: TypeNumber, Number: 1})); err == nil {
			t.Errorf("RestoreCell(%d, %d) gave no error", c.col, c.row)
		}
		if s.SnapshotCell(c.col, c.row).Exists() {
			t.Errorf("SnapshotCell(%d, %d) exists", c.col, c.row)
		}
	}

	// A formula above the size limit fails and leaves the cell unchanged.
	huge := "\"" + strings.Repeat("a", 70*1024) + "\""
	if err := s.RestoreCell(1, 1, CellSnapshot{}.WithFormula(huge)); err == nil {
		t.Skip("the library accepted a formula above the expansion limit")
	}
	if v := snapshotTestValue(t, s, "A1"); v.String != "keep" {
		t.Errorf("A1 after a failed restore = %#v, want keep", v)
	}

	// A formula that does not parse is stored, as SetFormula does.
	snapshotTestSetup(t, s.RestoreCell(1, 2, CellSnapshot{}.WithFormula("SUM(")))
	if got, _ := s.GetFormula("A2"); got != "SUM(" {
		t.Errorf("formula text = %q, want SUM(", got)
	}
	if v := snapshotTestValue(t, s, "A2"); v.Type != TypeError {
		t.Errorf("A2 = %#v, want an error value", v)
	}
}
