package ooxml

import (
	"bytes"
	"testing"
)

// writeAndRead is a helper that writes WorkbookData to an in-memory XLSX
// and reads it back, returning the parsed result.
func writeAndRead(t *testing.T, data *WorkbookData) *WorkbookData {
	t.Helper()
	var buf bytes.Buffer
	if err := WriteWorkbook(&buf, data); err != nil {
		t.Fatalf("WriteWorkbook: %v", err)
	}
	r := bytes.NewReader(buf.Bytes())
	got, err := ReadWorkbook(r, int64(buf.Len()))
	if err != nil {
		t.Fatalf("ReadWorkbook: %v", err)
	}
	return got
}

// ---------------------------------------------------------------------------
// Basic round-trip: single sheet with string, number, bool, formula cells
// ---------------------------------------------------------------------------

func TestRoundTrip_BasicCells(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}}, // default style at index 0
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{
				{Num: 1, Cells: []CellData{
					{Ref: "A1", Type: "s", Value: "hello"},
					{Ref: "B1", Value: "42"},
					{Ref: "C1", Type: "b", Value: "1"},
				}},
				{Num: 2, Cells: []CellData{
					{Ref: "A2", Type: "s", Value: "world"},
					{Ref: "B2", Formula: "B1+1"},
				}},
			},
		}},
	}

	got := writeAndRead(t, data)

	if len(got.Sheets) != 1 {
		t.Fatalf("Sheets = %d, want 1", len(got.Sheets))
	}
	s := got.Sheets[0]
	if s.Name != "Sheet1" {
		t.Errorf("Name = %q, want Sheet1", s.Name)
	}
	if len(s.Rows) != 2 {
		t.Fatalf("Rows = %d, want 2", len(s.Rows))
	}

	// Row 1
	r1 := s.Rows[0]
	if len(r1.Cells) != 3 {
		t.Fatalf("Row1 cells = %d, want 3", len(r1.Cells))
	}
	assertCell(t, r1.Cells[0], "A1", "s", "hello", "")
	assertCell(t, r1.Cells[1], "B1", "", "42", "")
	assertCell(t, r1.Cells[2], "C1", "b", "1", "")

	// Row 2
	r2 := s.Rows[1]
	if len(r2.Cells) != 2 {
		t.Fatalf("Row2 cells = %d, want 2", len(r2.Cells))
	}
	assertCell(t, r2.Cells[0], "A2", "s", "world", "")
	if r2.Cells[1].Formula != "B1+1" {
		t.Errorf("B2 formula = %q, want B1+1", r2.Cells[1].Formula)
	}
}

// ---------------------------------------------------------------------------
// Multiple sheets
// ---------------------------------------------------------------------------

func TestRoundTrip_MultipleSheets(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{
			{Name: "First", Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}}},
			{Name: "Second", Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Type: "s", Value: "two"}}}}},
			{Name: "Third", Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "3"}}}}},
		},
	}
	got := writeAndRead(t, data)
	if len(got.Sheets) != 3 {
		t.Fatalf("Sheets = %d, want 3", len(got.Sheets))
	}
	names := []string{got.Sheets[0].Name, got.Sheets[1].Name, got.Sheets[2].Name}
	want := []string{"First", "Second", "Third"}
	for i := range names {
		if names[i] != want[i] {
			t.Errorf("Sheet[%d].Name = %q, want %q", i, names[i], want[i])
		}
	}
}

// ---------------------------------------------------------------------------
// Hidden sheet state
// ---------------------------------------------------------------------------

func TestRoundTrip_SheetState(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{
			{Name: "Visible", Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}}},
			{Name: "Hidden", State: "hidden", Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "2"}}}}},
			{Name: "VeryHidden", State: "veryHidden", Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "3"}}}}},
		},
	}
	got := writeAndRead(t, data)
	if got.Sheets[0].State != "" {
		t.Errorf("Visible.State = %q, want empty", got.Sheets[0].State)
	}
	if got.Sheets[1].State != "hidden" {
		t.Errorf("Hidden.State = %q, want hidden", got.Sheets[1].State)
	}
	if got.Sheets[2].State != "veryHidden" {
		t.Errorf("VeryHidden.State = %q, want veryHidden", got.Sheets[2].State)
	}
}

// ---------------------------------------------------------------------------
// Date1904
// ---------------------------------------------------------------------------

func TestRoundTrip_Date1904(t *testing.T) {
	data := &WorkbookData{
		Date1904: true,
		Styles:   []StyleData{{}},
		Sheets:   []SheetData{{Name: "S1", Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}}}},
	}
	got := writeAndRead(t, data)
	if !got.Date1904 {
		t.Error("Date1904 lost during round-trip")
	}
}

func TestRoundTrip_Date1904_False(t *testing.T) {
	data := &WorkbookData{
		Date1904: false,
		Styles:   []StyleData{{}},
		Sheets:   []SheetData{{Name: "S1", Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}}}},
	}
	got := writeAndRead(t, data)
	if got.Date1904 {
		t.Error("Date1904 should be false")
	}
}

// ---------------------------------------------------------------------------
// CalcProperties
// ---------------------------------------------------------------------------

func TestRoundTrip_CalcProperties(t *testing.T) {
	data := &WorkbookData{
		CalcProps: CalcPropertiesData{
			Mode:           "auto",
			ID:             191029,
			FullCalcOnLoad: true,
			ForceFullCalc:  true,
			Completed:      true,
		},
		Styles: []StyleData{{}},
		Sheets: []SheetData{{Name: "S1", Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}}}},
	}
	got := writeAndRead(t, data)
	if got.CalcProps.Mode != "auto" {
		t.Errorf("CalcProps.Mode = %q, want auto", got.CalcProps.Mode)
	}
	if got.CalcProps.ID != 191029 {
		t.Errorf("CalcProps.ID = %d, want 191029", got.CalcProps.ID)
	}
	if !got.CalcProps.FullCalcOnLoad {
		t.Error("CalcProps.FullCalcOnLoad lost")
	}
	if !got.CalcProps.ForceFullCalc {
		t.Error("CalcProps.ForceFullCalc lost")
	}
	if !got.CalcProps.Completed {
		t.Error("CalcProps.Completed lost")
	}
}

// ---------------------------------------------------------------------------
// Column widths
// ---------------------------------------------------------------------------

func TestRoundTrip_ColumnWidths(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "S1",
			ColWidths: []ColWidthData{
				{Min: 1, Max: 1, Width: 20.5},
				{Min: 3, Max: 5, Width: 15.0},
			},
			Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}},
		}},
	}
	got := writeAndRead(t, data)
	cw := got.Sheets[0].ColWidths
	if len(cw) != 2 {
		t.Fatalf("ColWidths = %d, want 2", len(cw))
	}
	if cw[0].Min != 1 || cw[0].Max != 1 || cw[0].Width != 20.5 {
		t.Errorf("ColWidth[0] = %+v", cw[0])
	}
	if cw[1].Min != 3 || cw[1].Max != 5 || cw[1].Width != 15.0 {
		t.Errorf("ColWidth[1] = %+v", cw[1])
	}
}

// ---------------------------------------------------------------------------
// Merge cells
// ---------------------------------------------------------------------------

func TestRoundTrip_MergeCells(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "S1",
			MergeCells: []MergeCellData{
				{StartAxis: "A1", EndAxis: "C1"},
				{StartAxis: "B3", EndAxis: "D5"},
			},
			Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "merged"}}}},
		}},
	}
	got := writeAndRead(t, data)
	mc := got.Sheets[0].MergeCells
	if len(mc) != 2 {
		t.Fatalf("MergeCells = %d, want 2", len(mc))
	}
	if mc[0].StartAxis != "A1" || mc[0].EndAxis != "C1" {
		t.Errorf("MergeCell[0] = %+v", mc[0])
	}
	if mc[1].StartAxis != "B3" || mc[1].EndAxis != "D5" {
		t.Errorf("MergeCell[1] = %+v", mc[1])
	}
}

// ---------------------------------------------------------------------------
// Row height and hidden rows
// ---------------------------------------------------------------------------

func TestRoundTrip_RowHeightAndHidden(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "S1",
			Rows: []RowData{
				{Num: 1, Height: 30.0, Cells: []CellData{{Ref: "A1", Value: "tall"}}},
				{Num: 2, Hidden: true, Cells: []CellData{{Ref: "A2", Value: "hidden"}}},
				{Num: 3, Cells: []CellData{{Ref: "A3", Value: "normal"}}},
			},
		}},
	}
	got := writeAndRead(t, data)
	rows := got.Sheets[0].Rows
	if len(rows) != 3 {
		t.Fatalf("Rows = %d, want 3", len(rows))
	}
	if rows[0].Height != 30.0 {
		t.Errorf("Row1.Height = %f, want 30", rows[0].Height)
	}
	if !rows[1].Hidden {
		t.Error("Row2 should be hidden")
	}
	if rows[2].Hidden {
		t.Error("Row3 should not be hidden")
	}
}

// ---------------------------------------------------------------------------
// DefinedNames
// ---------------------------------------------------------------------------

func TestRoundTrip_DefinedNames(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}},
		}},
		DefinedNames: []DefinedName{
			{Name: "GlobalName", Value: "Sheet1!$A$1:$A$10", LocalSheetID: -1},
			{Name: "LocalName", Value: "Sheet1!$B$1", LocalSheetID: 0},
		},
	}
	got := writeAndRead(t, data)
	if len(got.DefinedNames) != 2 {
		t.Fatalf("DefinedNames = %d, want 2", len(got.DefinedNames))
	}
	dn0 := got.DefinedNames[0]
	if dn0.Name != "GlobalName" || dn0.Value != "Sheet1!$A$1:$A$10" || dn0.LocalSheetID != -1 {
		t.Errorf("DefinedName[0] = %+v", dn0)
	}
	dn1 := got.DefinedNames[1]
	if dn1.Name != "LocalName" || dn1.Value != "Sheet1!$B$1" || dn1.LocalSheetID != 0 {
		t.Errorf("DefinedName[1] = %+v", dn1)
	}
}

// ---------------------------------------------------------------------------
// Styles round-trip
// ---------------------------------------------------------------------------

func TestRoundTrip_Styles(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{
			{}, // index 0 = default
			{
				FontName:          "Arial",
				FontSize:          14,
				FontBold:          true,
				FontItalic:        true,
				FontColor:         "FFFF0000",
				FillColor:         "FF00FF00",
				BorderBottomStyle: "thin",
				BorderBottomColor: "FF000000",
				HAlign:            "center",
				VAlign:            "top",
				WrapText:          true,
				NumFmt:            "#,##0.00",
			},
		},
		Sheets: []SheetData{{
			Name: "S1",
			Rows: []RowData{{Num: 1, Cells: []CellData{
				{Ref: "A1", Value: "1"},
				{Ref: "B1", Value: "2", StyleIdx: 1},
			}}},
		}},
	}
	got := writeAndRead(t, data)

	// The styled cell (B1) should have StyleIdx > 0.
	b1 := got.Sheets[0].Rows[0].Cells[1]
	if b1.StyleIdx == 0 {
		t.Error("B1 StyleIdx should be non-zero")
	}

	// Verify the style at that index has our properties.
	if b1.StyleIdx >= len(got.Styles) {
		t.Fatalf("StyleIdx %d out of range (len=%d)", b1.StyleIdx, len(got.Styles))
	}
	sd := got.Styles[b1.StyleIdx]
	if sd.FontName != "Arial" {
		t.Errorf("FontName = %q, want Arial", sd.FontName)
	}
	if sd.FontSize != 14 {
		t.Errorf("FontSize = %f, want 14", sd.FontSize)
	}
	if !sd.FontBold {
		t.Error("FontBold should be true")
	}
	if !sd.FontItalic {
		t.Error("FontItalic should be true")
	}
	if sd.FontColor != "FFFF0000" {
		t.Errorf("FontColor = %q, want FFFF0000", sd.FontColor)
	}
	if sd.FillColor != "FF00FF00" {
		t.Errorf("FillColor = %q, want FF00FF00", sd.FillColor)
	}
	if sd.BorderBottomStyle != "thin" {
		t.Errorf("BorderBottomStyle = %q, want thin", sd.BorderBottomStyle)
	}
	if sd.HAlign != "center" {
		t.Errorf("HAlign = %q, want center", sd.HAlign)
	}
	if sd.VAlign != "top" {
		t.Errorf("VAlign = %q, want top", sd.VAlign)
	}
	if !sd.WrapText {
		t.Error("WrapText should be true")
	}
	if sd.NumFmt != "#,##0.00" {
		t.Errorf("NumFmt = %q, want #,##0.00", sd.NumFmt)
	}
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

func TestRoundTrip_Tables(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "Data",
			Rows: []RowData{
				{Num: 1, Cells: []CellData{
					{Ref: "A1", Type: "s", Value: "Name"},
					{Ref: "B1", Type: "s", Value: "Score"},
				}},
				{Num: 2, Cells: []CellData{
					{Ref: "A2", Type: "s", Value: "Alice"},
					{Ref: "B2", Value: "95"},
				}},
			},
		}},
		Tables: []TableDef{{
			Name:           "SalesTable",
			DisplayName:    "SalesTable",
			Ref:            "A1:B2",
			SheetIndex:     0,
			Columns:        []string{"Name", "Score"},
			HeaderRowCount: 1,
			HasAutoFilter:  true,
			Style: &TableStyleData{
				Name:           "TableStyleMedium2",
				ShowRowStripes: true,
			},
		}},
	}
	got := writeAndRead(t, data)
	if len(got.Tables) != 1 {
		t.Fatalf("Tables = %d, want 1", len(got.Tables))
	}
	td := got.Tables[0]
	if td.Name != "SalesTable" {
		t.Errorf("Name = %q", td.Name)
	}
	if td.DisplayName != "SalesTable" {
		t.Errorf("DisplayName = %q", td.DisplayName)
	}
	if td.Ref != "A1:B2" {
		t.Errorf("Ref = %q", td.Ref)
	}
	if td.SheetIndex != 0 {
		t.Errorf("SheetIndex = %d", td.SheetIndex)
	}
	if len(td.Columns) != 2 || td.Columns[0] != "Name" || td.Columns[1] != "Score" {
		t.Errorf("Columns = %v", td.Columns)
	}
	if td.HeaderRowCount != 1 {
		t.Errorf("HeaderRowCount = %d", td.HeaderRowCount)
	}
	if !td.HasAutoFilter {
		t.Error("HasAutoFilter should be true")
	}
	if td.Style == nil {
		t.Fatal("Style is nil")
	}
	if td.Style.Name != "TableStyleMedium2" {
		t.Errorf("Style.Name = %q", td.Style.Name)
	}
	if !td.Style.ShowRowStripes {
		t.Error("ShowRowStripes should be true")
	}
}

// ---------------------------------------------------------------------------
// Table with escaped column names (newlines, tabs)
// ---------------------------------------------------------------------------

func TestRoundTrip_TableEscapedColumns(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "S1",
			Rows: []RowData{
				{Num: 1, Cells: []CellData{
					{Ref: "A1", Type: "s", Value: "Col\nA"},
					{Ref: "B1", Type: "s", Value: "Col\tB"},
				}},
				{Num: 2, Cells: []CellData{
					{Ref: "A2", Value: "1"},
					{Ref: "B2", Value: "2"},
				}},
			},
		}},
		Tables: []TableDef{{
			Name:           "T1",
			DisplayName:    "T1",
			Ref:            "A1:B2",
			SheetIndex:     0,
			Columns:        []string{"Col\nA", "Col\tB"},
			HeaderRowCount: 1,
		}},
	}
	got := writeAndRead(t, data)
	if len(got.Tables) != 1 {
		t.Fatalf("Tables = %d, want 1", len(got.Tables))
	}
	cols := got.Tables[0].Columns
	if len(cols) != 2 {
		t.Fatalf("Columns = %d, want 2", len(cols))
	}
	if cols[0] != "Col\nA" {
		t.Errorf("Column[0] = %q, want 'Col\\nA'", cols[0])
	}
	if cols[1] != "Col\tB" {
		t.Errorf("Column[1] = %q, want 'Col\\tB'", cols[1])
	}
}

// ---------------------------------------------------------------------------
// Table with totals row and no auto-filter
// ---------------------------------------------------------------------------

func TestRoundTrip_TableTotals(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "S1",
			Rows: []RowData{
				{Num: 1, Cells: []CellData{{Ref: "A1", Type: "s", Value: "X"}}},
				{Num: 2, Cells: []CellData{{Ref: "A2", Value: "1"}}},
				{Num: 3, Cells: []CellData{{Ref: "A3", Value: "1"}}},
			},
		}},
		Tables: []TableDef{{
			Name:           "T1",
			DisplayName:    "T1",
			Ref:            "A1:A3",
			SheetIndex:     0,
			Columns:        []string{"X"},
			HeaderRowCount: 1,
			TotalsRowCount: 1,
			HasAutoFilter:  false,
		}},
	}
	got := writeAndRead(t, data)
	td := got.Tables[0]
	if td.TotalsRowCount != 1 {
		t.Errorf("TotalsRowCount = %d, want 1", td.TotalsRowCount)
	}
	if td.HasAutoFilter {
		t.Error("HasAutoFilter should be false")
	}
}

// ---------------------------------------------------------------------------
// Table with HeaderRowCount=0 (no header)
// ---------------------------------------------------------------------------

func TestRoundTrip_TableNoHeader(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "S1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}},
		}},
		Tables: []TableDef{{
			Name:           "T1",
			DisplayName:    "T1",
			Ref:            "A1:A1",
			SheetIndex:     0,
			Columns:        []string{"X"},
			HeaderRowCount: 0,
			HasAutoFilter:  true,
		}},
	}
	got := writeAndRead(t, data)
	td := got.Tables[0]
	if td.HeaderRowCount != 0 {
		t.Errorf("HeaderRowCount = %d, want 0", td.HeaderRowCount)
	}
	// HasAutoFilter should be false when HeaderRowCount is 0 (writer skips autoFilter).
	if td.HasAutoFilter {
		t.Error("HasAutoFilter should be false when HeaderRowCount=0")
	}
}

// ---------------------------------------------------------------------------
// Dynamic array formulas
// ---------------------------------------------------------------------------

func TestRoundTrip_DynamicArrayFormula(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "S1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{
				Ref:            "A1",
				Formula:        "SORT(B1:B10)",
				FormulaType:    "array",
				FormulaRef:     "A1:A10",
				IsDynamicArray: true,
			}}}},
		}},
	}
	got := writeAndRead(t, data)
	c := got.Sheets[0].Rows[0].Cells[0]
	if c.Formula != "SORT(B1:B10)" {
		t.Errorf("Formula = %q", c.Formula)
	}
	if !c.IsDynamicArray {
		t.Error("IsDynamicArray should be true")
	}
}

// ---------------------------------------------------------------------------
// Shared string deduplication across cells
// ---------------------------------------------------------------------------

func TestRoundTrip_SharedStringDedup(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "S1",
			Rows: []RowData{
				{Num: 1, Cells: []CellData{
					{Ref: "A1", Type: "s", Value: "same"},
					{Ref: "B1", Type: "s", Value: "same"},
					{Ref: "C1", Type: "s", Value: "different"},
				}},
			},
		}},
	}
	got := writeAndRead(t, data)
	cells := got.Sheets[0].Rows[0].Cells
	if cells[0].Value != "same" || cells[1].Value != "same" || cells[2].Value != "different" {
		t.Errorf("Values = %q, %q, %q", cells[0].Value, cells[1].Value, cells[2].Value)
	}
}

// ---------------------------------------------------------------------------
// Empty workbook error
// ---------------------------------------------------------------------------

func TestWriteWorkbook_NoSheets(t *testing.T) {
	var buf bytes.Buffer
	err := WriteWorkbook(&buf, &WorkbookData{})
	if err == nil {
		t.Error("expected error for workbook with no sheets")
	}
}

// ---------------------------------------------------------------------------
// Encrypted file detection
// ---------------------------------------------------------------------------

func TestReadWorkbook_EncryptedFile(t *testing.T) {
	// Create a buffer starting with CFB magic bytes.
	data := make([]byte, 512)
	copy(data, cfbMagic)
	r := bytes.NewReader(data)
	_, err := ReadWorkbook(r, int64(len(data)))
	if err == nil {
		t.Fatal("expected error for encrypted file")
	}
	if err != ErrEncryptedFile {
		t.Errorf("expected ErrEncryptedFile, got: %v", err)
	}
}

// ---------------------------------------------------------------------------
// Invalid zip
// ---------------------------------------------------------------------------

func TestReadWorkbook_InvalidZip(t *testing.T) {
	data := []byte("this is not a zip file at all")
	r := bytes.NewReader(data)
	_, err := ReadWorkbook(r, int64(len(data)))
	if err == nil {
		t.Error("expected error for invalid zip")
	}
}

// ---------------------------------------------------------------------------
// Shared strings with whitespace preservation
// ---------------------------------------------------------------------------

func TestRoundTrip_WhitespaceStrings(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "S1",
			Rows: []RowData{{Num: 1, Cells: []CellData{
				{Ref: "A1", Type: "s", Value: " leading"},
				{Ref: "B1", Type: "s", Value: "trailing "},
				{Ref: "C1", Type: "s", Value: "\ttab"},
			}}},
		}},
	}
	got := writeAndRead(t, data)
	cells := got.Sheets[0].Rows[0].Cells
	if cells[0].Value != " leading" {
		t.Errorf("A1 = %q, want ' leading'", cells[0].Value)
	}
	if cells[1].Value != "trailing " {
		t.Errorf("B1 = %q, want 'trailing '", cells[1].Value)
	}
	if cells[2].Value != "\ttab" {
		t.Errorf("C1 = %q, want '\\ttab'", cells[2].Value)
	}
}

// ---------------------------------------------------------------------------
// Array formula (CSE)
// ---------------------------------------------------------------------------

func TestRoundTrip_ArrayFormula(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "S1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{
				Ref:            "A1",
				Formula:        "SUM(B1:B10*C1:C10)",
				FormulaType:    "array",
				FormulaRef:     "A1",
				IsArrayFormula: true,
				Value:          "100",
			}}}},
		}},
	}
	got := writeAndRead(t, data)
	c := got.Sheets[0].Rows[0].Cells[0]
	if c.Formula != "SUM(B1:B10*C1:C10)" {
		t.Errorf("Formula = %q", c.Formula)
	}
	if c.FormulaType != "array" {
		t.Errorf("FormulaType = %q", c.FormulaType)
	}
	if c.FormulaRef != "A1" {
		t.Errorf("FormulaRef = %q", c.FormulaRef)
	}
}

// ---------------------------------------------------------------------------
// Inline string
// ---------------------------------------------------------------------------

func TestRoundTrip_InlineString(t *testing.T) {
	// Inline strings are written as type "inlineStr" with <is><t>.
	// Note: our writer doesn't generate inline strings (it uses SST),
	// but if the data says inlineStr, we should verify it doesn't corrupt.
	// Actually, the writer treats type "s" specially but not "inlineStr".
	// Let's test that "str" type (cached formula result) round-trips.
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "S1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{
				Ref:     "A1",
				Type:    "str",
				Value:   "computed text",
				Formula: `"hello"&" world"`,
			}}}},
		}},
	}
	got := writeAndRead(t, data)
	c := got.Sheets[0].Rows[0].Cells[0]
	if c.Type != "str" {
		t.Errorf("Type = %q, want str", c.Type)
	}
	if c.Value != "computed text" {
		t.Errorf("Value = %q", c.Value)
	}
}

// ---------------------------------------------------------------------------
// Multiple tables on same sheet
// ---------------------------------------------------------------------------

func TestRoundTrip_MultipleTablesOnSheet(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "Data",
			Rows: []RowData{
				{Num: 1, Cells: []CellData{
					{Ref: "A1", Type: "s", Value: "H1"},
					{Ref: "D1", Type: "s", Value: "H2"},
				}},
				{Num: 2, Cells: []CellData{
					{Ref: "A2", Value: "1"},
					{Ref: "D2", Value: "2"},
				}},
			},
		}},
		Tables: []TableDef{
			{
				Name: "T1", DisplayName: "T1",
				Ref: "A1:A2", SheetIndex: 0,
				Columns: []string{"H1"}, HeaderRowCount: 1,
			},
			{
				Name: "T2", DisplayName: "T2",
				Ref: "D1:D2", SheetIndex: 0,
				Columns: []string{"H2"}, HeaderRowCount: 1,
			},
		},
	}
	got := writeAndRead(t, data)
	if len(got.Tables) != 2 {
		t.Fatalf("Tables = %d, want 2", len(got.Tables))
	}
	if got.Tables[0].Name != "T1" || got.Tables[1].Name != "T2" {
		t.Errorf("Table names = %q, %q", got.Tables[0].Name, got.Tables[1].Name)
	}
}

// ---------------------------------------------------------------------------
// External workbook references in defined names are stripped
// ---------------------------------------------------------------------------

func TestRoundTrip_ExternalRefDefinedNamesStripped(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{
			{Name: "Sheet1", Rows: []RowData{
				{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}},
			}},
		},
		DefinedNames: []DefinedName{
			{Name: "Good", Value: "Sheet1!$A$1", LocalSheetID: -1},
			{Name: "ExtRef", Value: "'[1]Other Sheet'!$G$5", LocalSheetID: 0},
			{Name: "AlsoExt", Value: "[2]Sheet2!$A$1", LocalSheetID: -1},
			{Name: "LocalOK", Value: "Sheet1!$B$2", LocalSheetID: 0},
		},
	}
	got := writeAndRead(t, data)

	// Only the two non-external defined names should survive.
	if len(got.DefinedNames) != 2 {
		t.Fatalf("got %d defined names, want 2: %+v", len(got.DefinedNames), got.DefinedNames)
	}
	if got.DefinedNames[0].Name != "Good" {
		t.Errorf("DefinedNames[0].Name = %q, want Good", got.DefinedNames[0].Name)
	}
	if got.DefinedNames[1].Name != "LocalOK" {
		t.Errorf("DefinedNames[1].Name = %q, want LocalOK", got.DefinedNames[1].Name)
	}
}

// ---------------------------------------------------------------------------
// All defined names are external refs → no DefinedNames element at all
// ---------------------------------------------------------------------------

func TestRoundTrip_AllExternalRefsStripped(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{
			{Name: "Sheet1", Rows: []RowData{
				{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}},
			}},
		},
		DefinedNames: []DefinedName{
			{Name: "Ext1", Value: "'[1]Sheet'!$A$1", LocalSheetID: 0},
			{Name: "Ext2", Value: "[2]Sheet2!$B$2", LocalSheetID: -1},
		},
	}
	got := writeAndRead(t, data)

	if len(got.DefinedNames) != 0 {
		t.Fatalf("got %d defined names, want 0: %+v", len(got.DefinedNames), got.DefinedNames)
	}
}

// ---------------------------------------------------------------------------
// helper
// ---------------------------------------------------------------------------

func assertCell(t *testing.T, c CellData, wantRef, wantType, wantValue, wantFormula string) {
	t.Helper()
	if c.Ref != wantRef {
		t.Errorf("Ref = %q, want %q", c.Ref, wantRef)
	}
	if c.Type != wantType {
		t.Errorf("[%s] Type = %q, want %q", wantRef, c.Type, wantType)
	}
	if c.Value != wantValue {
		t.Errorf("[%s] Value = %q, want %q", wantRef, c.Value, wantValue)
	}
	if c.Formula != wantFormula {
		t.Errorf("[%s] Formula = %q, want %q", wantRef, c.Formula, wantFormula)
	}
}
