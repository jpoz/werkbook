package werkbook

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"strings"
	"testing"

	"github.com/jpoz/werkbook/ooxml"
)

func freezeTestSave(t *testing.T, f *File) []byte {
	t.Helper()
	var buf bytes.Buffer
	if _, err := f.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo: %v", err)
	}
	return buf.Bytes()
}

func freezeTestSheetXML(t *testing.T, data []byte, name string) string {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	for _, file := range zr.File {
		if file.Name != name {
			continue
		}
		rc, err := file.Open()
		if err != nil {
			t.Fatal(err)
		}
		defer rc.Close()
		text, err := io.ReadAll(rc)
		if err != nil {
			t.Fatal(err)
		}
		return string(text)
	}
	t.Fatalf("%s is not in the file", name)
	return ""
}

// sheetViewsOf returns the <sheetViews> element text of a sheet XML, or "".
func sheetViewsOf(sheetXML string) string {
	start := strings.Index(sheetXML, "<sheetViews")
	if start < 0 {
		return ""
	}
	end := strings.Index(sheetXML, "</sheetViews>")
	if end < 0 {
		return ""
	}
	return sheetXML[start : end+len("</sheetViews>")]
}

// fileWithSheetViews makes an XLSX whose first sheet has the given
// <sheetViews> element, as a file from a different program has.
func fileWithSheetViews(t *testing.T, sheetViews string) []byte {
	t.Helper()
	data := &ooxml.WorkbookData{
		Styles: []ooxml.StyleData{{}},
		Sheets: []ooxml.SheetData{{
			Name: "Sheet1",
			Rows: []ooxml.RowData{{Num: 1, Cells: []ooxml.CellData{{Ref: "A1", Value: "1"}}}},
			ExtraElements: []ooxml.RawElement{
				{Name: "sheetViews", XML: sheetViews, OrderKey: 4, Seq: 0},
				{Name: "pageMargins", XML: `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`, OrderKey: 40, Seq: 1},
			},
		}},
	}
	var buf bytes.Buffer
	if err := ooxml.WriteWorkbook(&buf, data); err != nil {
		t.Fatalf("WriteWorkbook: %v", err)
	}
	return buf.Bytes()
}

func TestSetFreezePanesSaveAndOpen(t *testing.T) {
	tests := []struct {
		name       string
		rows, cols int
		wantPane   string // "" = no pane element
	}{
		{"rows only", 2, 0, `<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>`},
		{"columns only", 0, 3, `<pane xSplit="3" topLeftCell="D1" activePane="topRight" state="frozen"/>`},
		{"rows and columns", 1, 1, `<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>`},
		{"many", 40, 28, `<pane xSplit="28" ySplit="40" topLeftCell="AC41" activePane="bottomRight" state="frozen"/>`},
		{"the largest freeze", MaxRows - 1, MaxColumns - 1, `<pane xSplit="16383" ySplit="1048575" topLeftCell="XFD1048576" activePane="bottomRight" state="frozen"/>`},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := New()
			s := f.Sheet("Sheet1")
			if err := s.SetValue("A1", 1); err != nil {
				t.Fatal(err)
			}
			if rows, cols := s.FreezePanes(); rows != 0 || cols != 0 {
				t.Fatalf("a new sheet is frozen: %d, %d", rows, cols)
			}
			if err := s.SetFreezePanes(tt.rows, tt.cols); err != nil {
				t.Fatal(err)
			}
			if rows, cols := s.FreezePanes(); rows != tt.rows || cols != tt.cols {
				t.Errorf("FreezePanes = %d, %d", rows, cols)
			}
			data := freezeTestSave(t, f)
			sheetXML := freezeTestSheetXML(t, data, "xl/worksheets/sheet1.xml")
			want := `<sheetViews><sheetView workbookViewId="0">` + tt.wantPane + `</sheetView></sheetViews>`
			if got := sheetViewsOf(sheetXML); got != want {
				t.Errorf("sheetViews = %s\nwant         %s", got, want)
			}
			// The element is before sheetData, as the schema requires, and
			// the sheet XML is well formed.
			if i, j := strings.Index(sheetXML, "<sheetViews"), strings.Index(sheetXML, "<sheetData"); i < 0 || j < 0 || i > j {
				t.Errorf("sheetViews is not before sheetData: %s", sheetXML)
			}
			if err := xml.Unmarshal([]byte(sheetXML), new(struct{})); err != nil {
				t.Errorf("sheet XML is not well formed: %v", err)
			}

			opened, err := OpenReader(bytes.NewReader(data))
			if err != nil {
				t.Fatal(err)
			}
			if rows, cols := opened.Sheet("Sheet1").FreezePanes(); rows != tt.rows || cols != tt.cols {
				t.Errorf("after save and open: FreezePanes = %d, %d, want %d, %d", rows, cols, tt.rows, tt.cols)
			}
			// A second save with no change gives the same element.
			again := freezeTestSheetXML(t, freezeTestSave(t, opened), "xl/worksheets/sheet1.xml")
			if got := sheetViewsOf(again); got != want {
				t.Errorf("sheetViews after a second save = %s", got)
			}

			// Remove the freeze: the view stays, the pane goes.
			if err := opened.Sheet("Sheet1").SetFreezePanes(0, 0); err != nil {
				t.Fatal(err)
			}
			removed := freezeTestSheetXML(t, freezeTestSave(t, opened), "xl/worksheets/sheet1.xml")
			if got := sheetViewsOf(removed); got != "" && strings.Contains(got, "<pane") {
				t.Errorf("sheetViews after the freeze was removed = %s", got)
			}
			reopened, err := OpenReader(bytes.NewReader(freezeTestSave(t, opened)))
			if err != nil {
				t.Fatal(err)
			}
			if rows, cols := reopened.Sheet("Sheet1").FreezePanes(); rows != 0 || cols != 0 {
				t.Errorf("after removal, save and open: FreezePanes = %d, %d", rows, cols)
			}
		})
	}

	t.Run("a sheet that was never frozen writes no sheetViews", func(t *testing.T) {
		f := New()
		sheetXML := freezeTestSheetXML(t, freezeTestSave(t, f), "xl/worksheets/sheet1.xml")
		if strings.Contains(sheetXML, "sheetViews") {
			t.Errorf("sheet XML = %s", sheetXML)
		}
	})

	t.Run("errors", func(t *testing.T) {
		s := New().Sheet("Sheet1")
		if err := s.SetFreezePanes(3, 2); err != nil {
			t.Fatal(err)
		}
		for _, c := range [][2]int{{-1, 0}, {0, -1}, {MaxRows, 0}, {0, MaxColumns}, {MaxRows + 5, MaxColumns + 5}} {
			if err := s.SetFreezePanes(c[0], c[1]); err == nil {
				t.Errorf("SetFreezePanes(%d, %d) gave no error", c[0], c[1])
			}
		}
		if rows, cols := s.FreezePanes(); rows != 3 || cols != 2 {
			t.Errorf("a failed call changed the freeze: %d, %d", rows, cols)
		}
	})
}

func TestFreezePanesFromAFile(t *testing.T) {
	// A view as Excel writes it: attributes, a frozen pane, and selections.
	excelViews := `<sheetViews><sheetView showGridLines="0" tabSelected="1" zoomScale="85" workbookViewId="0">` +
		`<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>` +
		`<selection pane="topRight" activeCell="B1" sqref="B1"/><selection pane="bottomLeft" activeCell="A3" sqref="A3"/>` +
		`<selection pane="bottomRight" activeCell="C5" sqref="C5"/></sheetView></sheetViews>`

	t.Run("the freeze is read", func(t *testing.T) {
		f, err := OpenReader(bytes.NewReader(fileWithSheetViews(t, excelViews)))
		if err != nil {
			t.Fatal(err)
		}
		if rows, cols := f.Sheet("Sheet1").FreezePanes(); rows != 2 || cols != 1 {
			t.Errorf("FreezePanes = %d, %d, want 2, 1", rows, cols)
		}
	})

	t.Run("no change of the freeze gives the same sheetViews bytes", func(t *testing.T) {
		f, err := OpenReader(bytes.NewReader(fileWithSheetViews(t, excelViews)))
		if err != nil {
			t.Fatal(err)
		}
		s := f.Sheet("Sheet1")
		if err := s.SetValue("B2", "edit"); err != nil {
			t.Fatal(err)
		}
		// A call with the freeze that the file has is no change.
		if err := s.SetFreezePanes(2, 1); err != nil {
			t.Fatal(err)
		}
		sheetXML := freezeTestSheetXML(t, freezeTestSave(t, f), "xl/worksheets/sheet1.xml")
		if got := sheetViewsOf(sheetXML); got != excelViews {
			t.Errorf("sheetViews = %s\nwant the bytes of the file: %s", got, excelViews)
		}
		if !strings.Contains(sheetXML, "<pageMargins") {
			t.Errorf("the other preserved element is gone: %s", sheetXML)
		}
	})

	t.Run("a new freeze keeps the other attributes of the view", func(t *testing.T) {
		f, err := OpenReader(bytes.NewReader(fileWithSheetViews(t, excelViews)))
		if err != nil {
			t.Fatal(err)
		}
		if err := f.Sheet("Sheet1").SetFreezePanes(5, 0); err != nil {
			t.Fatal(err)
		}
		data := freezeTestSave(t, f)
		sheetXML := freezeTestSheetXML(t, data, "xl/worksheets/sheet1.xml")
		// The selection of the pane that is active now (bottomLeft) stays;
		// the selections of the panes that are gone do not.
		want := `<sheetViews><sheetView showGridLines="0" tabSelected="1" zoomScale="85" workbookViewId="0">` +
			`<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/>` +
			`<selection pane="bottomLeft" activeCell="A3" sqref="A3"/></sheetView></sheetViews>`
		if got := sheetViewsOf(sheetXML); got != want {
			t.Errorf("sheetViews = %s\nwant         %s", got, want)
		}
		// Exactly one sheetViews element, in its position.
		if strings.Count(sheetXML, "<sheetViews") != 1 {
			t.Errorf("sheet XML has more than one sheetViews: %s", sheetXML)
		}
		if i, j := strings.Index(sheetXML, "<sheetViews"), strings.Index(sheetXML, "<sheetData"); i > j {
			t.Errorf("sheetViews is after sheetData")
		}
		opened, err := OpenReader(bytes.NewReader(data))
		if err != nil {
			t.Fatal(err)
		}
		if rows, cols := opened.Sheet("Sheet1").FreezePanes(); rows != 5 || cols != 0 {
			t.Errorf("after save and open: %d, %d", rows, cols)
		}
	})

	t.Run("removal keeps the view attributes", func(t *testing.T) {
		f, err := OpenReader(bytes.NewReader(fileWithSheetViews(t, excelViews)))
		if err != nil {
			t.Fatal(err)
		}
		if err := f.Sheet("Sheet1").SetFreezePanes(0, 0); err != nil {
			t.Fatal(err)
		}
		sheetXML := freezeTestSheetXML(t, freezeTestSave(t, f), "xl/worksheets/sheet1.xml")
		if !strings.Contains(sheetXML, `<sheetView showGridLines="0" tabSelected="1" zoomScale="85" workbookViewId="0"/>`) || strings.Contains(sheetXML, "<pane") {
			t.Errorf("sheet XML = %s", sheetXML)
		}
	})

	t.Run("a split that is not frozen is not a freeze and is kept", func(t *testing.T) {
		split := `<sheetViews><sheetView workbookViewId="0"><pane xSplit="2000" ySplit="1500" topLeftCell="C4" activePane="bottomRight"/></sheetView></sheetViews>`
		f, err := OpenReader(bytes.NewReader(fileWithSheetViews(t, split)))
		if err != nil {
			t.Fatal(err)
		}
		if rows, cols := f.Sheet("Sheet1").FreezePanes(); rows != 0 || cols != 0 {
			t.Errorf("FreezePanes of a split view = %d, %d, want 0, 0", rows, cols)
		}
		sheetXML := freezeTestSheetXML(t, freezeTestSave(t, f), "xl/worksheets/sheet1.xml")
		if got := sheetViewsOf(sheetXML); got != split {
			t.Errorf("sheetViews = %s, want the split of the file", got)
		}
	})

	t.Run("other forms of the pane", func(t *testing.T) {
		tests := []struct {
			views      string
			rows, cols int
		}{
			{`<sheetViews><sheetView workbookViewId="0"><pane ySplit="1.0" topLeftCell="A2" state="frozenSplit"/></sheetView></sheetViews>`, 1, 0},
			{`<sheetViews><sheetView workbookViewId="0"><pane xSplit="4" state="frozen"/></sheetView></sheetViews>`, 0, 4},
			{`<sheetViews><sheetView workbookViewId="0"/></sheetViews>`, 0, 0},
			{`<sheetViews><sheetView workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>`, 0, 0},
			{`<sheetViews><sheetView workbookViewId="0"><pane ySplit="x" state="frozen"/></sheetView></sheetViews>`, 0, 0},
			// Only the first view counts.
			{`<sheetViews><sheetView workbookViewId="0"/><sheetView workbookViewId="1"><pane ySplit="3" state="frozen"/></sheetView></sheetViews>`, 0, 0},
		}
		for _, tt := range tests {
			f, err := OpenReader(bytes.NewReader(fileWithSheetViews(t, tt.views)))
			if err != nil {
				t.Fatalf("%s: %v", tt.views, err)
			}
			if rows, cols := f.Sheet("Sheet1").FreezePanes(); rows != tt.rows || cols != tt.cols {
				t.Errorf("%s: FreezePanes = %d, %d, want %d, %d", tt.views, rows, cols, tt.rows, tt.cols)
			}
			// With no change the element is byte-equal.
			sheetXML := freezeTestSheetXML(t, freezeTestSave(t, f), "xl/worksheets/sheet1.xml")
			if got := sheetViewsOf(sheetXML); got != tt.views && !strings.HasSuffix(tt.views, "/></sheetViews>") {
				t.Errorf("sheetViews = %s, want %s", got, tt.views)
			}
		}
	})
}

// A change of the freeze changes only the pane of the first view. The other
// views, the selection of the active pane and the other children stay.
func TestFreezeChangeKeepsViewsAndSelection(t *testing.T) {
	sheetViews := func(t *testing.T, views string, change func(s *Sheet) error) string {
		t.Helper()
		f, err := OpenReader(bytes.NewReader(fileWithSheetViews(t, views)))
		if err != nil {
			t.Fatal(err)
		}
		if err := change(f.Sheet("Sheet1")); err != nil {
			t.Fatal(err)
		}
		data := freezeTestSave(t, f)
		sheetXML := freezeTestSheetXML(t, data, "xl/worksheets/sheet1.xml")
		if err := xml.Unmarshal([]byte(sheetXML), new(struct{})); err != nil {
			t.Fatalf("sheet XML is not well formed: %v\n%s", err, sheetXML)
		}
		// The file opens again and has the freeze.
		if _, err := OpenReader(bytes.NewReader(data)); err != nil {
			t.Fatalf("open of the saved file: %v", err)
		}
		return sheetViewsOf(sheetXML)
	}
	freeze := func(rows, cols int) func(*Sheet) error {
		return func(s *Sheet) error { return s.SetFreezePanes(rows, cols) }
	}

	second := `<sheetView workbookViewId="1" zoomScale="50"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A9" sqref="A9"/></sheetView>`
	tests := []struct {
		name   string
		views  string
		change func(*Sheet) error
		want   string
	}{
		{
			name:   "two views: the second view stays as it is",
			views:  `<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="C3" sqref="C3"/></sheetView>` + second + `</sheetViews>`,
			change: freeze(4, 0),
			want: `<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView>` +
				second + `</sheetViews>`,
		},
		{
			name:   "the selection of the new active pane stays",
			views:  `<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="B1" sqref="B1"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/><selection pane="bottomRight" activeCell="D7" sqref="D7:E8"/></sheetView></sheetViews>`,
			change: freeze(3, 2),
			want:   `<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="D7" sqref="D7:E8"/></sheetView></sheetViews>`,
		},
		{
			name:   "removal: the selection with no pane stays, the pane selections go",
			views:  `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection activeCell="B1" sqref="B1"/><selection pane="bottomLeft" activeCell="A5" sqref="A5"/></sheetView></sheetViews>`,
			change: freeze(0, 0),
			want:   `<sheetViews><sheetView workbookViewId="0"><selection activeCell="B1" sqref="B1"/></sheetView></sheetViews>`,
		},
		{
			name:   "a first freeze: the selection of the single pane is for a pane that is not active now",
			views:  `<sheetViews><sheetView showGridLines="0" workbookViewId="0"><selection activeCell="C3" sqref="C3"/></sheetView></sheetViews>`,
			change: freeze(1, 0),
			want:   `<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`,
		},
		{
			name:   "other children of the view stay after the selections",
			views:  `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A5" sqref="A5"/><extLst><ext uri="{x}"><y/></ext></extLst></sheetView></sheetViews>`,
			change: freeze(2, 0),
			want:   `<sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A5" sqref="A5"/><extLst><ext uri="{x}"><y/></ext></extLst></sheetView></sheetViews>`,
		},
		{
			name:  "a row insert in the frozen rows keeps the selection and the second view",
			views: `<sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B9" sqref="B9"/></sheetView>` + second + `</sheetViews>`,
			change: func(s *Sheet) error {
				_, err := s.InsertRows(1, 1)
				return err
			},
			want: `<sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B9" sqref="B9"/></sheetView>` + second + `</sheetViews>`,
		},
		{
			name:  "a row insert below the freeze changes no byte",
			views: `<sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/><selection pane="topLeft" activeCell="A1" sqref="A1"/><selection pane="bottomLeft" activeCell="B9" sqref="B9"/></sheetView>` + second + `</sheetViews>`,
			change: func(s *Sheet) error {
				_, err := s.InsertRows(10, 3)
				return err
			},
			want: `<sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/><selection pane="topLeft" activeCell="A1" sqref="A1"/><selection pane="bottomLeft" activeCell="B9" sqref="B9"/></sheetView>` + second + `</sheetViews>`,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := sheetViews(t, tt.views, tt.change); got != tt.want {
				t.Errorf("sheetViews = %s\nwant         %s", got, tt.want)
			}
		})
	}
}

func TestCopySheetCopiesTheFreeze(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	if err := s.SetFreezePanes(2, 1); err != nil {
		t.Fatal(err)
	}
	cp, err := f.CopySheet("Sheet1", "Copy")
	if err != nil {
		t.Fatal(err)
	}
	if rows, cols := cp.FreezePanes(); rows != 2 || cols != 1 {
		t.Errorf("FreezePanes of the copy = %d, %d", rows, cols)
	}
	// The two sheets are independent.
	if err := cp.SetFreezePanes(0, 0); err != nil {
		t.Fatal(err)
	}
	if rows, cols := s.FreezePanes(); rows != 2 || cols != 1 {
		t.Errorf("the source freeze changed: %d, %d", rows, cols)
	}
	opened, err := OpenReader(bytes.NewReader(freezeTestSave(t, f)))
	if err != nil {
		t.Fatal(err)
	}
	if rows, cols := opened.Sheet("Sheet1").FreezePanes(); rows != 2 || cols != 1 {
		t.Errorf("Sheet1 after save and open: %d, %d", rows, cols)
	}
	if rows, cols := opened.Sheet("Copy").FreezePanes(); rows != 0 || cols != 0 {
		t.Errorf("Copy after save and open: %d, %d", rows, cols)
	}
}

// The frozen rows and columns are the span 1..n: an insert in the span makes
// it larger and a delete makes it smaller.
func TestStructuralEditFreezePanes(t *testing.T) {
	tests := []struct {
		name               string
		op                 structOp
		wantRows, wantCols int // from a freeze of 3 rows and 2 columns
	}{
		{"insert rows in the frozen rows", structOp{"Sheet1", "insRows", 2, 2}, 5, 2},
		{"insert rows at row 1", structOp{"Sheet1", "insRows", 1, 1}, 4, 2},
		{"insert rows at the last frozen row", structOp{"Sheet1", "insRows", 3, 1}, 4, 2},
		{"insert rows at the first row that scrolls", structOp{"Sheet1", "insRows", 4, 5}, 3, 2},
		{"insert rows far below", structOp{"Sheet1", "insRows", 100, 5}, 3, 2},
		{"delete a frozen row", structOp{"Sheet1", "delRows", 2, 1}, 2, 2},
		{"delete rows across the freeze line", structOp{"Sheet1", "delRows", 3, 4}, 2, 2},
		{"delete all frozen rows", structOp{"Sheet1", "delRows", 1, 3}, 0, 2},
		{"delete all frozen rows and more", structOp{"Sheet1", "delRows", 1, 10}, 0, 2},
		{"delete rows below the freeze", structOp{"Sheet1", "delRows", 4, 3}, 3, 2},
		{"insert columns in the frozen columns", structOp{"Sheet1", "insCols", 2, 3}, 3, 5},
		{"insert a column at A", structOp{"Sheet1", "insCols", 1, 1}, 3, 3},
		{"insert a column at the first that scrolls", structOp{"Sheet1", "insCols", 3, 1}, 3, 2},
		{"delete a frozen column", structOp{"Sheet1", "delCols", 1, 1}, 3, 1},
		{"delete all frozen columns", structOp{"Sheet1", "delCols", 1, 2}, 3, 0},
		{"delete columns right of the freeze", structOp{"Sheet1", "delCols", 3, 2}, 3, 2},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := newStructureWorkbook(t)
			s := f.Sheet("Sheet1")
			if err := s.SetFreezePanes(3, 2); err != nil {
				t.Fatal(err)
			}
			// The helper checks save and open, Undo, and a second run.
			checkStructuralEdit(t, f, tt.op, nil, nil)
			if rows, cols := s.FreezePanes(); rows != tt.wantRows || cols != tt.wantCols {
				t.Errorf("FreezePanes = %d, %d, want %d, %d", rows, cols, tt.wantRows, tt.wantCols)
			}
		})
	}

	t.Run("an edit on a different sheet does not change the freeze", func(t *testing.T) {
		f := newStructureWorkbook(t)
		s := f.Sheet("Sheet1")
		if err := s.SetFreezePanes(3, 2); err != nil {
			t.Fatal(err)
		}
		if _, err := f.Sheet("Other").InsertRows(1, 5); err != nil {
			t.Fatal(err)
		}
		if rows, cols := s.FreezePanes(); rows != 3 || cols != 2 {
			t.Errorf("FreezePanes = %d, %d", rows, cols)
		}
		if rows, cols := f.Sheet("Other").FreezePanes(); rows != 0 || cols != 0 {
			t.Errorf("the other sheet got a freeze: %d, %d", rows, cols)
		}
	})
}
