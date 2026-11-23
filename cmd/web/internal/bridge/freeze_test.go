package bridge

import (
	"encoding/json"
	"strings"
	"testing"

	werkbook "github.com/jpoz/werkbook"
)

func freezeOf(t *testing.T, s *Session, sheet string) Freeze {
	t.Helper()
	meta, err := s.SheetMeta(sheet)
	if err != nil {
		t.Fatal(err)
	}
	return meta.Freeze
}

func TestSetFreeze(t *testing.T) {
	s := NewSession()
	if _, err := s.AddSheet("Two"); err != nil {
		t.Fatal(err)
	}
	mustSet(t, s, sheet1, "A1", "=RAND()")
	rand := cellsByRef(t, s, sheet1, 1, 1, 1, 1)["A1"].Display
	if got := freezeOf(t, s, sheet1); got != (Freeze{}) {
		t.Fatalf("freeze of a new sheet = %+v", got)
	}
	depth := s.History().Depth

	res, err := s.SetFreeze(sheet1, 2, 1)
	if err != nil {
		t.Fatal(err)
	}
	if res.History.Depth != depth+1 || res.History.UndoLabel != "Freeze panes" {
		t.Errorf("history = %+v", res.History)
	}
	if got := freezeOf(t, s, sheet1); got != (Freeze{Rows: 2, Cols: 1}) {
		t.Errorf("freeze = %+v", got)
	}
	// The other sheet has its own freeze.
	if got := freezeOf(t, s, "Two"); got != (Freeze{}) {
		t.Errorf("freeze of the other sheet = %+v", got)
	}

	// The same freeze again makes no step.
	if res, err := s.SetFreeze(sheet1, 2, 1); err != nil || res.History.Depth != depth+1 {
		t.Errorf("same freeze: %+v, %v", res, err)
	}
	// A change, then a removal.
	if _, err := s.SetFreeze(sheet1, 0, 3); err != nil {
		t.Fatal(err)
	}
	if _, err := s.SetFreeze(sheet1, 0, 0); err != nil {
		t.Fatal(err)
	}
	if got := freezeOf(t, s, sheet1); got != (Freeze{}) {
		t.Errorf("freeze after removal = %+v", got)
	}

	// Undo and redo through the three steps.
	undo := mustUndo(t, s)
	if undo.Label != "Freeze panes" || undo.Sheet != sheet1 || undo.Select != nil {
		t.Errorf("undo = %+v", undo)
	}
	if got := freezeOf(t, s, sheet1); got != (Freeze{Cols: 3}) {
		t.Errorf("freeze after one undo = %+v", got)
	}
	mustUndo(t, s)
	if got := freezeOf(t, s, sheet1); got != (Freeze{Rows: 2, Cols: 1}) {
		t.Errorf("freeze after two undos = %+v", got)
	}
	mustUndo(t, s)
	if got := freezeOf(t, s, sheet1); got != (Freeze{}) {
		t.Errorf("freeze after three undos = %+v", got)
	}
	mustRedo(t, s)
	if got := freezeOf(t, s, sheet1); got != (Freeze{Rows: 2, Cols: 1}) {
		t.Errorf("freeze after redo = %+v", got)
	}

	// A freeze is a view setting: no recalculation.
	if got := cellsByRef(t, s, sheet1, 1, 1, 1, 1)["A1"].Display; got != rand {
		t.Errorf("the value of =RAND() changed across freeze calls: %s -> %s", rand, got)
	}

	// Save and open.
	data, err := s.Save()
	if err != nil {
		t.Fatal(err)
	}
	s2 := NewSession()
	if _, err := s2.Open(data); err != nil {
		t.Fatal(err)
	}
	if got := freezeOf(t, s2, sheet1); got != (Freeze{Rows: 2, Cols: 1}) {
		t.Errorf("freeze after save and open = %+v", got)
	}
	if got := freezeOf(t, s2, "Two"); got != (Freeze{}) {
		t.Errorf("freeze of the other sheet after save and open = %+v", got)
	}

	encoded, err := json.Marshal(mustMeta(t, s2, sheet1))
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(encoded), `"freeze":{"rows":2,"cols":1}`) {
		t.Errorf("sheetMeta JSON = %s", encoded)
	}
}

func TestSetFreezeErrors(t *testing.T) {
	s := NewSession()
	if _, err := s.SetFreeze(sheet1, 1, 1); err != nil {
		t.Fatal(err)
	}
	h := s.History()
	bad := []struct {
		sheet      string
		rows, cols int
	}{
		{"Nope", 1, 1}, {sheet1, -1, 0}, {sheet1, 0, -1},
		{sheet1, werkbook.MaxRows, 0}, {sheet1, 0, werkbook.MaxColumns},
	}
	for _, b := range bad {
		if _, err := s.SetFreeze(b.sheet, b.rows, b.cols); err == nil {
			t.Errorf("SetFreeze(%q, %d, %d) gave no error", b.sheet, b.rows, b.cols)
		}
	}
	if s.History() != h || freezeOf(t, s, sheet1) != (Freeze{Rows: 1, Cols: 1}) {
		t.Errorf("a failed call changed the freeze or the history")
	}
	// The largest freeze is valid.
	if _, err := s.SetFreeze(sheet1, werkbook.MaxRows-1, werkbook.MaxColumns-1); err != nil {
		t.Errorf("largest freeze: %v", err)
	}
}

// The freeze follows insert and delete of rows and columns, and their undo.
func TestFreezeFollowsStructuralEdits(t *testing.T) {
	s := newStructureSession(t)
	if _, err := s.SetFreeze(sheet1, 3, 1); err != nil {
		t.Fatal(err)
	}
	steps := []struct {
		name string
		run  func() error
		want Freeze
	}{
		{"insert rows in the frozen rows", func() error { _, err := s.InsertRows(sheet1, 2, 2); return err }, Freeze{5, 1}},
		{"insert rows below the freeze", func() error { _, err := s.InsertRows(sheet1, 20, 2); return err }, Freeze{5, 1}},
		{"delete a frozen row", func() error { _, err := s.DeleteRows(sheet1, 1, 1); return err }, Freeze{4, 1}},
		{"insert a column at A", func() error { _, err := s.InsertCols(sheet1, 1, 1); return err }, Freeze{4, 2}},
		{"delete the frozen columns", func() error { _, err := s.DeleteCols(sheet1, 1, 2); return err }, Freeze{4, 0}},
		{"an edit on a different sheet", func() error { _, err := s.InsertRows("Two", 1, 3); return err }, Freeze{4, 0}},
	}
	states := []Freeze{freezeOf(t, s, sheet1)}
	for _, st := range steps {
		if err := st.run(); err != nil {
			t.Fatalf("%s: %v", st.name, err)
		}
		if got := freezeOf(t, s, sheet1); got != st.want {
			t.Errorf("%s: freeze = %+v, want %+v", st.name, got, st.want)
		}
		states = append(states, st.want)
	}
	for i := len(steps) - 1; i >= 0; i-- {
		mustUndo(t, s)
		if got := freezeOf(t, s, sheet1); got != states[i] {
			t.Errorf("after undo of %q: freeze = %+v, want %+v", steps[i].name, got, states[i])
		}
	}
	for i := range steps {
		mustRedo(t, s)
		if got := freezeOf(t, s, sheet1); got != states[i+1] {
			t.Errorf("after redo of %q: freeze = %+v, want %+v", steps[i].name, got, states[i+1])
		}
	}
}
