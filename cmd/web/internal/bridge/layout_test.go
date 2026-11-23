package bridge

import (
	"reflect"
	"strings"
	"testing"

	werkbook "github.com/jpoz/werkbook"
)

func mustMeta(t *testing.T, s *Session, sheet string) SheetMeta {
	t.Helper()
	meta, err := s.SheetMeta(sheet)
	if err != nil {
		t.Fatalf("SheetMeta: %v", err)
	}
	return meta
}

func TestSetColWidthAndRowHeight(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "x")

	res, err := s.SetColWidth(sheet1, 2, 4, 120)
	if err != nil {
		t.Fatal(err)
	}
	if res.History.Depth != 2 || res.History.UndoLabel != "Column width" {
		t.Errorf("history = %+v", res.History)
	}
	if got := mustMeta(t, s, sheet1).ColWidths; !reflect.DeepEqual(got, map[int]int{2: 120, 3: 120, 4: 120}) {
		t.Errorf("colWidths = %v", got)
	}

	// Reversed numbers; a second call changes only the columns that differ.
	if _, err := s.SetColWidth(sheet1, 5, 4, 80); err != nil {
		t.Fatal(err)
	}
	if got := mustMeta(t, s, sheet1).ColWidths; !reflect.DeepEqual(got, map[int]int{2: 120, 3: 120, 4: 80, 5: 80}) {
		t.Errorf("colWidths = %v", got)
	}

	res, err = s.SetRowHeight(sheet1, 3, 5, 40)
	if err != nil {
		t.Fatal(err)
	}
	if res.History.Depth != 4 || res.History.UndoLabel != "Row height" {
		t.Errorf("history = %+v", res.History)
	}
	if got := mustMeta(t, s, sheet1).RowHeights; !reflect.DeepEqual(got, map[int]int{3: 40, 4: 40, 5: 40}) {
		t.Errorf("rowHeights = %v", got)
	}
	// Rows that have only a height are not in the used range.
	if info := s.Sheets()[0]; info.MaxRow != 1 || info.MaxCol != 1 {
		t.Errorf("used range = %dx%d, want 1x1", info.MaxRow, info.MaxCol)
	}

	// Undo the row heights: the heights and the empty rows go away.
	undo := mustUndo(t, s)
	if undo.Label != "Row height" || undo.Sheet != sheet1 || *undo.Select != (Range{3, 1, 5, werkbook.MaxColumns}) {
		t.Errorf("undo = %+v %v", undo, undo.Select)
	}
	if got := mustMeta(t, s, sheet1).RowHeights; len(got) != 0 {
		t.Errorf("rowHeights after undo = %v", got)
	}
	if info := undo.Sheets[0]; info.MaxRow != 1 {
		t.Errorf("MaxRow after undo = %d, want 1", info.MaxRow)
	}
	// Undo the second width call, then the first.
	undo = mustUndo(t, s)
	if undo.Label != "Column width" || *undo.Select != (Range{1, 4, werkbook.MaxRows, 5}) {
		t.Errorf("undo = %+v %v", undo, undo.Select)
	}
	if got := mustMeta(t, s, sheet1).ColWidths; !reflect.DeepEqual(got, map[int]int{2: 120, 3: 120, 4: 120}) {
		t.Errorf("colWidths after undo = %v", got)
	}
	mustUndo(t, s)
	if got := mustMeta(t, s, sheet1).ColWidths; len(got) != 0 {
		t.Errorf("colWidths after two undos = %v", got)
	}
	// Redo all three.
	mustRedo(t, s)
	mustRedo(t, s)
	mustRedo(t, s)
	meta := mustMeta(t, s, sheet1)
	if !reflect.DeepEqual(meta.ColWidths, map[int]int{2: 120, 3: 120, 4: 80, 5: 80}) || !reflect.DeepEqual(meta.RowHeights, map[int]int{3: 40, 4: 40, 5: 40}) {
		t.Errorf("meta after redo = %+v", meta)
	}
	if got := cellsByRef(t, s, sheet1, 1, 1, 1, 1)["A1"].Display; got != "x" {
		t.Errorf("A1 = %q; a size change must not change cells", got)
	}
}

func TestSizeRemoveAndLimits(t *testing.T) {
	s := NewSession()
	if _, err := s.SetColWidth(sheet1, 1, 3, 100); err != nil {
		t.Fatal(err)
	}
	if _, err := s.SetRowHeight(sheet1, 1, 3, 30); err != nil {
		t.Fatal(err)
	}

	// px <= 0 removes the custom size.
	if _, err := s.SetColWidth(sheet1, 2, 2, 0); err != nil {
		t.Fatal(err)
	}
	if _, err := s.SetRowHeight(sheet1, 2, 3, -1); err != nil {
		t.Fatal(err)
	}
	meta := mustMeta(t, s, sheet1)
	if !reflect.DeepEqual(meta.ColWidths, map[int]int{1: 100, 3: 100}) || !reflect.DeepEqual(meta.RowHeights, map[int]int{1: 30}) {
		t.Errorf("meta after removal = %+v", meta)
	}
	if info := s.Sheets()[0]; info.MaxRow != 0 {
		t.Errorf("MaxRow of a sheet with sizes and no cells = %d, want 0", info.MaxRow)
	}
	if got := s.file.Sheet(sheet1).MaxRow(); got != 1 {
		t.Errorf("stored rows after the heights were removed: last = %d, want 1", got)
	}
	mustUndo(t, s)
	mustUndo(t, s)
	meta = mustMeta(t, s, sheet1)
	if !reflect.DeepEqual(meta.ColWidths, map[int]int{1: 100, 2: 100, 3: 100}) || !reflect.DeepEqual(meta.RowHeights, map[int]int{1: 30, 2: 30, 3: 30}) {
		t.Errorf("meta after undo of the removals = %+v", meta)
	}

	// Limits.
	limits := []struct {
		cols     bool
		px, want int
	}{
		{true, 1, 12}, {true, 11, 12}, {true, 12, 12}, {true, 2000, 2000}, {true, 2001, 2000}, {true, 1 << 30, 2000},
		{false, 1, 8}, {false, 7, 8}, {false, 8, 8}, {false, 1000, 1000}, {false, 1001, 1000}, {false, 1 << 30, 1000},
	}
	for _, l := range limits {
		var got int
		if l.cols {
			if _, err := s.SetColWidth(sheet1, 9, 9, l.px); err != nil {
				t.Fatal(err)
			}
			got = mustMeta(t, s, sheet1).ColWidths[9]
		} else {
			if _, err := s.SetRowHeight(sheet1, 9, 9, l.px); err != nil {
				t.Fatal(err)
			}
			got = mustMeta(t, s, sheet1).RowHeights[9]
		}
		if got != l.want {
			t.Errorf("cols=%v px %d gives %d, want %d", l.cols, l.px, got, l.want)
		}
	}
}

// Each pixel number that the contract can set must come back the same from
// SheetMeta, also after a save.
func TestSizePixelRoundTrip(t *testing.T) {
	for px := minColWidthPx; px <= maxColWidthPx; px++ {
		if got := colWidthToPx(pxToColWidth(px)); got != px {
			t.Fatalf("column width %d px comes back as %d", px, got)
		}
	}
	for px := minRowHeightPx; px <= maxRowHeightPx; px++ {
		if got := rowHeightToPx(pxToRowHeight(px)); got != px {
			t.Fatalf("row height %d px comes back as %d", px, got)
		}
	}

	s := NewSession()
	wantCols, wantRows := map[int]int{}, map[int]int{}
	for i, px := range []int{12, 13, 64, 99, 100, 123, 257, 1999, 2000} {
		if _, err := s.SetColWidth(sheet1, i+1, i+1, px); err != nil {
			t.Fatal(err)
		}
		wantCols[i+1] = px
	}
	for i, px := range []int{8, 9, 20, 21, 33, 41, 100, 999, 1000} {
		if _, err := s.SetRowHeight(sheet1, i+1, i+1, px); err != nil {
			t.Fatal(err)
		}
		wantRows[i+1] = px
	}
	data, err := s.Save()
	if err != nil {
		t.Fatal(err)
	}
	s2 := NewSession()
	if _, err := s2.Open(data); err != nil {
		t.Fatal(err)
	}
	meta := mustMeta(t, s2, sheet1)
	if !reflect.DeepEqual(meta.ColWidths, wantCols) {
		t.Errorf("colWidths after save and open = %v, want %v", meta.ColWidths, wantCols)
	}
	if !reflect.DeepEqual(meta.RowHeights, wantRows) {
		t.Errorf("rowHeights after save and open = %v, want %v", meta.RowHeights, wantRows)
	}

	// A removed size is not in the file.
	if _, err := s.SetColWidth(sheet1, 1, 9, 0); err != nil {
		t.Fatal(err)
	}
	if _, err := s.SetRowHeight(sheet1, 1, 9, 0); err != nil {
		t.Fatal(err)
	}
	data, err = s.Save()
	if err != nil {
		t.Fatal(err)
	}
	if _, err := s2.Open(data); err != nil {
		t.Fatal(err)
	}
	if meta := mustMeta(t, s2, sheet1); len(meta.ColWidths) != 0 || len(meta.RowHeights) != 0 {
		t.Errorf("meta after removal, save and open = %+v", meta)
	}
}

func TestSizeNoStepAndErrors(t *testing.T) {
	s := NewSession()
	if _, err := s.SetColWidth(sheet1, 2, 3, 90); err != nil {
		t.Fatal(err)
	}
	if _, err := s.SetRowHeight(sheet1, 2, 3, 25); err != nil {
		t.Fatal(err)
	}
	mustUndo(t, s) // the row height is on the redo stack
	h := s.History()

	// A call that changes no size makes no step and keeps the redo stack.
	noop := []func() (LayoutResult, error){
		func() (LayoutResult, error) { return s.SetColWidth(sheet1, 2, 3, 90) },
		func() (LayoutResult, error) { return s.SetColWidth(sheet1, 3, 3, 90) },
		func() (LayoutResult, error) { return s.SetColWidth(sheet1, 7, 9, 0) },
		func() (LayoutResult, error) { return s.SetRowHeight(sheet1, 5, 9, 0) },
	}
	for i, call := range noop {
		res, err := call()
		if err != nil {
			t.Fatalf("call %d: %v", i, err)
		}
		if res.History != h || s.History() != h {
			t.Errorf("call %d changed the history: %+v, want %+v", i, res.History, h)
		}
	}
	if info := s.Sheets()[0]; info.MaxRow != 0 {
		t.Errorf("MaxRow = %d; a removal on rows with no height must make no rows", info.MaxRow)
	}

	bad := []struct {
		name string
		call func() (LayoutResult, error)
		want string
	}{
		{"unknown sheet, columns", func() (LayoutResult, error) { return s.SetColWidth("Nope", 1, 1, 50) }, "not found"},
		{"unknown sheet, rows", func() (LayoutResult, error) { return s.SetRowHeight("Nope", 1, 1, 50) }, "not found"},
		{"column 0", func() (LayoutResult, error) { return s.SetColWidth(sheet1, 0, 2, 50) }, "out of range"},
		{"column past the end", func() (LayoutResult, error) { return s.SetColWidth(sheet1, 1, werkbook.MaxColumns+1, 50) }, "out of range"},
		{"row 0", func() (LayoutResult, error) { return s.SetRowHeight(sheet1, 2, 0, 50) }, "out of range"},
		{"row past the end", func() (LayoutResult, error) { return s.SetRowHeight(sheet1, 1, werkbook.MaxRows+1, 50) }, "out of range"},
	}
	for _, b := range bad {
		_, err := b.call()
		if err == nil || !strings.Contains(err.Error(), b.want) {
			t.Errorf("%s: error = %v, want it to contain %q", b.name, err, b.want)
		}
	}
	if s.History() != h {
		t.Errorf("failed calls changed the history: %+v", s.History())
	}
	if got := mustMeta(t, s, sheet1); !reflect.DeepEqual(got.ColWidths, map[int]int{2: 90, 3: 90}) || len(got.RowHeights) != 0 {
		t.Errorf("failed calls changed the sizes: %+v", got)
	}

	// All columns in one call, and the row limit exactly.
	if _, err := s.SetColWidth(sheet1, 1, werkbook.MaxColumns, 50); err != nil {
		t.Fatalf("all columns: %v", err)
	}
	if got := len(mustMeta(t, s, sheet1).ColWidths); got != werkbook.MaxColumns {
		t.Errorf("%d columns have a width, want %d", got, werkbook.MaxColumns)
	}
	mustUndo(t, s)
	if got := mustMeta(t, s, sheet1).ColWidths; !reflect.DeepEqual(got, map[int]int{2: 90, 3: 90}) {
		t.Errorf("colWidths after undo of all columns = %d entries", len(got))
	}
	if _, err := s.SetRowHeight(sheet1, 1, maxLayoutRows, 30); err != nil {
		t.Fatalf("row limit: %v", err)
	}
	mustUndo(t, s)
	if info := s.Sheets()[0]; info.MaxRow != 0 {
		t.Errorf("MaxRow after undo of %d row heights = %d, want 0", maxLayoutRows, info.MaxRow)
	}
}

// A selection of all rows (select all, then a drag of one row handle) must
// work, as a selection of all columns does.
func TestSetRowHeightForAllRows(t *testing.T) {
	allRows := Range{1, 1, werkbook.MaxRows, werkbook.MaxColumns}

	t.Run("a height for all rows is limited to the used range", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "x", "C7", "y")
		res, err := s.SetRowHeight(sheet1, 1, werkbook.MaxRows, 30)
		if err != nil {
			t.Fatalf("SetRowHeight for all rows: %v", err)
		}
		if !res.Clamped || res.Range != (Range{1, 1, 7, werkbook.MaxColumns}) {
			t.Errorf("result = %+v, want clamped to rows 1..7", res)
		}
		heights := mustMeta(t, s, sheet1).RowHeights
		if len(heights) != 7 || heights[1] != 30 || heights[7] != 30 {
			t.Errorf("rowHeights = %v, want rows 1..7 at 30", heights)
		}
		undo := mustUndo(t, s)
		if *undo.Select != (Range{1, 1, 7, werkbook.MaxColumns}) {
			t.Errorf("undo select = %v", undo.Select)
		}
		if got := mustMeta(t, s, sheet1).RowHeights; len(got) != 0 {
			t.Errorf("rowHeights after undo = %v", got)
		}
	})

	t.Run("the used range has spill rows", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "=SEQUENCE(5)")
		res, err := s.SetRowHeight(sheet1, 1, werkbook.MaxRows, 30)
		if err != nil || res.Range.Row2 != 5 {
			t.Errorf("result = %+v, %v; want rows 1..5", res, err)
		}
	})

	t.Run("an empty sheet gets the height on the first row of the range", func(t *testing.T) {
		s := NewSession()
		res, err := s.SetRowHeight(sheet1, 1, werkbook.MaxRows, 30)
		if err != nil || !res.Clamped || res.Range != (Range{1, 1, 1, werkbook.MaxColumns}) {
			t.Errorf("result = %+v, %v", res, err)
		}
		// A range that starts below the used range.
		res, err = s.SetRowHeight(sheet1, 500, werkbook.MaxRows, 30)
		if err != nil || res.Range != (Range{500, 1, 500, werkbook.MaxColumns}) {
			t.Errorf("result = %+v, %v", res, err)
		}
	})

	t.Run("a removal for all rows has no limit and touches only rows with a height", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "x")
		for _, row := range []int{2, 900_000, werkbook.MaxRows} {
			if _, err := s.SetRowHeight(sheet1, row, row, 40); err != nil {
				t.Fatal(err)
			}
		}
		res, err := s.SetRowHeight(sheet1, 1, werkbook.MaxRows, 0)
		if err != nil {
			t.Fatalf("removal for all rows: %v", err)
		}
		if res.Clamped || res.Range != allRows || res.History.UndoLabel != "Row height" {
			t.Errorf("result = %+v", res)
		}
		if got := mustMeta(t, s, sheet1).RowHeights; len(got) != 0 {
			t.Errorf("rowHeights after the removal = %v", got)
		}
		// One step; its undo restores the three heights.
		mustUndo(t, s)
		want := map[int]int{2: 40, 900_000: 40, werkbook.MaxRows: 40}
		if got := mustMeta(t, s, sheet1).RowHeights; !reflect.DeepEqual(got, want) {
			t.Errorf("rowHeights after undo = %v, want %v", got, want)
		}
		// A removal where no row has a height makes no step.
		mustRedo(t, s)
		depth := s.History().Depth
		if res, err := s.SetRowHeight(sheet1, 1, werkbook.MaxRows, -1); err != nil || res.History.Depth != depth {
			t.Errorf("second removal: %+v, %v", res, err)
		}
	})

	t.Run("a used range above the limit is still an error", func(t *testing.T) {
		s := NewSession()
		mustSet(t, s, sheet1, "A1", "x", "A300000", "far")
		h := s.History()
		if _, err := s.SetRowHeight(sheet1, 1, werkbook.MaxRows, 30); err == nil || !strings.Contains(err.Error(), "more than 200000 rows") {
			t.Errorf("error = %v, want the row limit", err)
		}
		if s.History() != h || len(mustMeta(t, s, sheet1).RowHeights) != 0 {
			t.Errorf("a failed call changed the sheet or the history")
		}
		// A range in the limit is not clamped, also past the used range.
		res, err := s.SetRowHeight(sheet1, 299_990, 300_010, 30)
		if err != nil || res.Clamped || len(mustMeta(t, s, sheet1).RowHeights) != 21 {
			t.Errorf("result = %+v, %v", res, err)
		}
	})

	t.Run("all columns", func(t *testing.T) {
		s := NewSession()
		res, err := s.SetColWidth(sheet1, 1, werkbook.MaxColumns, 80)
		if err != nil || res.Clamped || res.Range != allRows {
			t.Errorf("result = %+v, %v", res, err)
		}
	})
}

// A size change cannot change a formula result, so it must not start a
// recalculation: a volatile formula keeps its value.
func TestResizeDoesNotRecalculate(t *testing.T) {
	s := NewSession()
	mustSet(t, s, sheet1, "A1", "=RAND()", "B1", "=A1*2")
	read := func() string { return cellsByRef(t, s, sheet1, 1, 1, 1, 2)["A1"].Display }
	first := read()
	if read() != first {
		t.Fatalf("test setup: two reads with no write give different values")
	}

	steps := []struct {
		name string
		run  func() error
	}{
		{"setColWidth", func() error { _, err := s.SetColWidth(sheet1, 2, 4, 150); return err }},
		{"setRowHeight", func() error { _, err := s.SetRowHeight(sheet1, 1, 3, 40); return err }},
		{"undo of a resize", func() error { _, err := s.Undo(); return err }},
		{"redo of a resize", func() error { _, err := s.Redo(); return err }},
		{"removal of a width", func() error { _, err := s.SetColWidth(sheet1, 2, 4, 0); return err }},
		{"sheetMeta and sheets", func() error { _, err := s.SheetMeta(sheet1); s.Sheets(); return err }},
	}
	for _, st := range steps {
		if err := st.run(); err != nil {
			t.Fatalf("%s: %v", st.name, err)
		}
		if got := read(); got != first {
			t.Errorf("after %s the value of =RAND() changed: %s -> %s", st.name, first, got)
		}
	}

	// A cell write still recalculates.
	mustSet(t, s, sheet1, "D1", "1")
	if got := read(); got == first {
		t.Errorf("a cell write did not recalculate the volatile formula")
	}
	// An undo of a cell write recalculates too.
	second := read()
	mustUndo(t, s)
	if got := read(); got == second {
		t.Errorf("an undo of a cell write did not recalculate")
	}
}

func TestSizesFromAFileStay(t *testing.T) {
	// A file from a different program has widths that are not whole pixels.
	src := werkbook.New()
	sh := src.Sheet(sheet1)
	for _, err := range []error{sh.SetColumnWidth("A", 8.43), sh.SetColumnWidth("C", 30.71), sh.SetRowHeight(2, 15.75)} {
		if err != nil {
			t.Fatal(err)
		}
	}
	var buf strings.Builder
	if _, err := src.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	s := NewSession()
	if _, err := s.Open([]byte(buf.String())); err != nil {
		t.Fatal(err)
	}
	before := mustMeta(t, s, sheet1)
	if !reflect.DeepEqual(before.ColWidths, map[int]int{1: 64, 3: 220}) || !reflect.DeepEqual(before.RowHeights, map[int]int{2: 21}) {
		t.Fatalf("meta of the file = %+v", before)
	}
	// A change and its undo give the exact old width, not a rounded one.
	if _, err := s.SetColWidth(sheet1, 1, 3, 100); err != nil {
		t.Fatal(err)
	}
	mustUndo(t, s)
	if w, _ := s.file.Sheet(sheet1).GetColumnWidth("A"); w != 8.43 {
		t.Errorf("width of A after undo = %v, want 8.43", w)
	}
	if w := s.file.Sheet(sheet1).ColumnWidths(); len(w) != 2 {
		t.Errorf("widths after undo = %v, want A and C only", w)
	}
}
