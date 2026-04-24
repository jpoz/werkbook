package werkbook

import (
	"bytes"
	"errors"
	"slices"
	"testing"

	"github.com/jpoz/werkbook/ooxml"
)

func TestCopySheetPreservesCellsAndMetadata(t *testing.T) {
	f := New(FirstSheet("Source"))
	src := f.Sheet("Source")

	style := &Style{
		Font:      &Font{Name: "Arial", Size: 14, Bold: true},
		Fill:      &Fill{Color: "FFFF00"},
		Alignment: &Alignment{Horizontal: HAlignCenter, WrapText: true},
		NumFmt:    "0.00",
	}

	if err := src.SetValue("A1", "title"); err != nil {
		t.Fatal(err)
	}
	if err := src.SetStyle("A1", style); err != nil {
		t.Fatal(err)
	}
	if err := src.SetValue("B2", 12.5); err != nil {
		t.Fatal(err)
	}
	if err := src.SetValue("C3", true); err != nil {
		t.Fatal(err)
	}
	if err := src.SetValue("D4", Value{Type: TypeError, String: "#DIV/0!"}); err != nil {
		t.Fatal(err)
	}
	if err := src.SetFormula("E2", "B2*2"); err != nil {
		t.Fatal(err)
	}
	if err := src.SetColumnWidth("A", 18.5); err != nil {
		t.Fatal(err)
	}
	if err := src.SetColumnWidth("E", 22); err != nil {
		t.Fatal(err)
	}
	if err := src.SetRowHeight(1, 24); err != nil {
		t.Fatal(err)
	}
	if err := src.SetRowHeight(4, 40); err != nil {
		t.Fatal(err)
	}
	if err := src.MergeCell("A1", "C1"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetSheetVisible("Source", false); err != nil {
		t.Fatal(err)
	}

	copied, err := f.CopySheet("Source", "Copied")
	if err != nil {
		t.Fatalf("CopySheet: %v", err)
	}

	if !slices.Equal(f.SheetNames(), []string{"Source", "Copied"}) {
		t.Fatalf("sheet names = %v, want [Source Copied]", f.SheetNames())
	}
	if copied.Visible() {
		t.Fatal("copied sheet should preserve hidden state")
	}

	w, err := copied.GetColumnWidth("A")
	if err != nil {
		t.Fatal(err)
	}
	if w != 18.5 {
		t.Fatalf("A width = %g, want 18.5", w)
	}
	w, err = copied.GetColumnWidth("E")
	if err != nil {
		t.Fatal(err)
	}
	if w != 22 {
		t.Fatalf("E width = %g, want 22", w)
	}

	h, err := copied.GetRowHeight(1)
	if err != nil {
		t.Fatal(err)
	}
	if h != 24 {
		t.Fatalf("row 1 height = %g, want 24", h)
	}
	h, err = copied.GetRowHeight(4)
	if err != nil {
		t.Fatal(err)
	}
	if h != 40 {
		t.Fatalf("row 4 height = %g, want 40", h)
	}

	merges := copied.MergeCells()
	if len(merges) != 1 || merges[0] != (MergeRange{Start: "A1", End: "C1"}) {
		t.Fatalf("merges = %+v, want [{A1 C1}]", merges)
	}

	v, err := copied.GetValue("A1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeString || v.String != "title" {
		t.Fatalf("A1 = %#v, want string title", v)
	}

	v, err = copied.GetValue("B2")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeNumber || v.Number != 12.5 {
		t.Fatalf("B2 = %#v, want 12.5", v)
	}

	v, err = copied.GetValue("C3")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeBool || !v.Bool {
		t.Fatalf("C3 = %#v, want TRUE", v)
	}

	v, err = copied.GetValue("D4")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeError || v.String != "#DIV/0!" {
		t.Fatalf("D4 = %#v, want #DIV/0!", v)
	}

	formulaText, err := copied.GetFormula("E2")
	if err != nil {
		t.Fatal(err)
	}
	if formulaText != "B2*2" {
		t.Fatalf("E2 formula = %q, want %q", formulaText, "B2*2")
	}

	v, err = copied.GetValue("E2")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeNumber || v.Number != 25 {
		t.Fatalf("E2 = %#v, want 25", v)
	}

	srcStyle, err := src.GetStyle("A1")
	if err != nil {
		t.Fatal(err)
	}
	copiedStyle, err := copied.GetStyle("A1")
	if err != nil {
		t.Fatal(err)
	}
	if copiedStyle == nil || copiedStyle.Font == nil {
		t.Fatal("copied A1 style is missing")
	}
	if copiedStyle == srcStyle {
		t.Fatal("copied style should not alias source style")
	}
	if copiedStyle.Font == srcStyle.Font {
		t.Fatal("copied font should not alias source font")
	}
	if copiedStyle.Font.Name != "Arial" || copiedStyle.Font.Size != 14 || !copiedStyle.Font.Bold {
		t.Fatalf("copied style font = %+v, want Arial/14/bold", copiedStyle.Font)
	}

	copiedStyle.Font.Bold = false
	if !srcStyle.Font.Bold {
		t.Fatal("mutating copied style should not affect source style")
	}

	if err := copied.SetValue("B2", 7); err != nil {
		t.Fatal(err)
	}
	v, err = copied.GetValue("E2")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeNumber || v.Number != 14 {
		t.Fatalf("copied E2 after edit = %#v, want 14", v)
	}

	v, err = src.GetValue("E2")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeNumber || v.Number != 25 {
		t.Fatalf("source E2 after copied-sheet edit = %#v, want 25", v)
	}
}

func TestCloneSheetFromAnotherWorkbook(t *testing.T) {
	srcFile := New(FirstSheet("Source"))
	src := srcFile.Sheet("Source")

	if err := src.SetValue("A1", 3); err != nil {
		t.Fatal(err)
	}
	if err := src.SetFormula("B1", "A1+1"); err != nil {
		t.Fatal(err)
	}
	if err := src.SetStyle("B1", &Style{Fill: &Fill{Color: "00FF00"}}); err != nil {
		t.Fatal(err)
	}
	if err := src.SetColumnWidth("B", 30); err != nil {
		t.Fatal(err)
	}
	if err := src.SetRowHeight(1, 33); err != nil {
		t.Fatal(err)
	}
	if err := srcFile.SetSheetVisible("Source", false); err != nil {
		t.Fatal(err)
	}

	dstFile := New(FirstSheet("Existing"))
	cloned, err := dstFile.CloneSheetFrom(src, "Imported")
	if err != nil {
		t.Fatalf("CloneSheetFrom: %v", err)
	}

	if cloned == nil || cloned.Name() != "Imported" {
		t.Fatalf("cloned sheet = %#v, want Imported", cloned)
	}
	if cloned.Visible() {
		t.Fatal("cloned sheet should preserve hidden state")
	}

	formulaText, err := cloned.GetFormula("B1")
	if err != nil {
		t.Fatal(err)
	}
	if formulaText != "A1+1" {
		t.Fatalf("B1 formula = %q, want %q", formulaText, "A1+1")
	}

	v, err := cloned.GetValue("B1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeNumber || v.Number != 4 {
		t.Fatalf("B1 = %#v, want 4", v)
	}

	style, err := cloned.GetStyle("B1")
	if err != nil {
		t.Fatal(err)
	}
	if style == nil || style.Fill == nil || style.Fill.Color != "00FF00" {
		t.Fatalf("B1 style = %+v, want fill 00FF00", style)
	}

	w, err := cloned.GetColumnWidth("B")
	if err != nil {
		t.Fatal(err)
	}
	if w != 30 {
		t.Fatalf("B width = %g, want 30", w)
	}

	h, err := cloned.GetRowHeight(1)
	if err != nil {
		t.Fatal(err)
	}
	if h != 33 {
		t.Fatalf("row 1 height = %g, want 33", h)
	}

	if err := cloned.SetValue("A1", 9); err != nil {
		t.Fatal(err)
	}
	v, err = cloned.GetValue("B1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeNumber || v.Number != 10 {
		t.Fatalf("B1 after edit = %#v, want 10", v)
	}

	v, err = src.GetValue("B1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeNumber || v.Number != 4 {
		t.Fatalf("source B1 after destination edit = %#v, want 4", v)
	}
}

func TestCopySheetPreservesPassthrough(t *testing.T) {
	f := New(FirstSheet("Source"))
	src := f.Sheet("Source")

	// Simulate a sheet with passthrough metadata (as if read from a real xlsx).
	src.rootAttrs = []ooxml.RawAttr{
		{Name: "xmlns", Value: ooxml.NSSpreadsheetML},
		{Name: "xmlns:r", Value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"},
		{Name: "xmlns:mc", Value: "http://schemas.openxmlformats.org/markup-compatibility/2006"},
		{Name: "mc:Ignorable", Value: "x14ac"},
	}
	src.extraElements = []ooxml.RawElement{
		{Name: "conditionalFormatting", XML: `<conditionalFormatting sqref="B2"/>`, OrderKey: 32, Seq: 0},
		{Name: "drawing", XML: `<drawing r:id="rId2"/>`, OrderKey: 58, Seq: 1},
	}
	src.extraRels = []ooxml.OpaqueRel{
		{ID: "rId2", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", Target: "../drawings/drawing1.xml"},
		{ID: "rId3", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", Target: "https://example.com/", TargetMode: "External"},
	}

	// Add the opaque entry the drawing rel points to.
	f.opaqueEntries = append(f.opaqueEntries, ooxml.OpaqueEntry{
		Path:        "xl/drawings/drawing1.xml",
		ContentType: "application/vnd.openxmlformats-officedocument.drawing+xml",
		Data:        []byte("<drawing/>"),
	})

	copied, err := f.CopySheet("Source", "Copy")
	if err != nil {
		t.Fatalf("CopySheet: %v", err)
	}

	// Verify rootAttrs copied.
	if len(copied.rootAttrs) != len(src.rootAttrs) {
		t.Fatalf("rootAttrs len = %d, want %d", len(copied.rootAttrs), len(src.rootAttrs))
	}
	for i, want := range src.rootAttrs {
		if copied.rootAttrs[i] != want {
			t.Fatalf("rootAttrs[%d] = %+v, want %+v", i, copied.rootAttrs[i], want)
		}
	}

	// Verify extraElements copied.
	if len(copied.extraElements) != len(src.extraElements) {
		t.Fatalf("extraElements len = %d, want %d", len(copied.extraElements), len(src.extraElements))
	}
	for i, want := range src.extraElements {
		if copied.extraElements[i] != want {
			t.Fatalf("extraElements[%d] = %+v, want %+v", i, copied.extraElements[i], want)
		}
	}

	// Verify extraRels copied.
	if len(copied.extraRels) != len(src.extraRels) {
		t.Fatalf("extraRels len = %d, want %d", len(copied.extraRels), len(src.extraRels))
	}
	for i, want := range src.extraRels {
		if copied.extraRels[i] != want {
			t.Fatalf("extraRels[%d] = %+v, want %+v", i, copied.extraRels[i], want)
		}
	}

	// Same-workbook clone should not duplicate opaque entries.
	drawingCount := 0
	for _, e := range f.opaqueEntries {
		if e.Path == "xl/drawings/drawing1.xml" {
			drawingCount++
		}
	}
	if drawingCount != 1 {
		t.Fatalf("expected 1 drawing entry, got %d (same-workbook clone should share)", drawingCount)
	}

	// Verify independence: mutating copied fields doesn't affect source.
	copied.rootAttrs[0] = ooxml.RawAttr{Name: "mutated", Value: "yes"}
	if src.rootAttrs[0].Name == "mutated" {
		t.Fatal("mutating copied rootAttrs should not affect source")
	}
}

func TestCloneSheetFromCrossWorkbookOpaqueEntries(t *testing.T) {
	srcFile := New(FirstSheet("Source"))
	src := srcFile.Sheet("Source")
	if err := src.SetValue("A1", "hello"); err != nil {
		t.Fatal(err)
	}

	src.rootAttrs = []ooxml.RawAttr{
		{Name: "xmlns", Value: ooxml.NSSpreadsheetML},
		{Name: "xmlns:mc", Value: "http://schemas.openxmlformats.org/markup-compatibility/2006"},
	}
	src.extraElements = []ooxml.RawElement{
		{Name: "drawing", XML: `<drawing r:id="rId1"/>`, OrderKey: 58, Seq: 0},
	}
	src.extraRels = []ooxml.OpaqueRel{
		{ID: "rId1", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", Target: "../drawings/drawing1.xml"},
		{ID: "rId2", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", Target: "https://example.com/", TargetMode: "External"},
	}
	srcFile.opaqueEntries = []ooxml.OpaqueEntry{
		{Path: "xl/drawings/drawing1.xml", ContentType: "application/vnd.openxmlformats-officedocument.drawing+xml", Data: []byte("<srcDrawing/>")},
		{Path: "xl/drawings/_rels/drawing1.xml.rels", ContentType: "", Data: []byte("<rels/>")},
	}
	srcFile.opaqueDefaults = []ooxml.OpaqueDefault{
		{Extension: "png", ContentType: "image/png"},
	}

	dstFile := New(FirstSheet("Existing"))

	cloned, err := dstFile.CloneSheetFrom(src, "Imported")
	if err != nil {
		t.Fatalf("CloneSheetFrom: %v", err)
	}

	// Verify passthrough fields were copied.
	if len(cloned.rootAttrs) != 2 {
		t.Fatalf("rootAttrs len = %d, want 2", len(cloned.rootAttrs))
	}
	if len(cloned.extraElements) != 1 {
		t.Fatalf("extraElements len = %d, want 1", len(cloned.extraElements))
	}
	if len(cloned.extraRels) != 2 {
		t.Fatalf("extraRels len = %d, want 2", len(cloned.extraRels))
	}

	// Verify the drawing opaque entry was imported into dstFile.
	found := false
	for _, e := range dstFile.opaqueEntries {
		if e.Path == "xl/drawings/drawing1.xml" {
			found = true
			if !bytes.Equal(e.Data, []byte("<srcDrawing/>")) {
				t.Fatalf("imported drawing data = %q, want <srcDrawing/>", e.Data)
			}
		}
	}
	if !found {
		t.Fatal("drawing entry was not imported into destination workbook")
	}

	// Verify the external rel was preserved as-is (no opaque entry for it).
	if cloned.extraRels[1].TargetMode != "External" {
		t.Fatalf("expected external rel, got %+v", cloned.extraRels[1])
	}

	// Verify opaqueDefaults were imported.
	foundPNG := false
	for _, d := range dstFile.opaqueDefaults {
		if d.Extension == "png" {
			foundPNG = true
		}
	}
	if !foundPNG {
		t.Fatal("png opaqueDefault was not imported")
	}
}

func TestCloneSheetFromCrossWorkbookPathCollision(t *testing.T) {
	srcFile := New(FirstSheet("Source"))
	src := srcFile.Sheet("Source")
	src.extraRels = []ooxml.OpaqueRel{
		{ID: "rId1", Type: "drawing", Target: "../drawings/drawing1.xml"},
	}
	srcFile.opaqueEntries = []ooxml.OpaqueEntry{
		{Path: "xl/drawings/drawing1.xml", ContentType: "drawing+xml", Data: []byte("src-drawing")},
	}

	dstFile := New(FirstSheet("Existing"))
	// Pre-populate destination with a different entry at the same path.
	dstFile.opaqueEntries = []ooxml.OpaqueEntry{
		{Path: "xl/drawings/drawing1.xml", ContentType: "drawing+xml", Data: []byte("existing-drawing")},
	}

	cloned, err := dstFile.CloneSheetFrom(src, "Imported")
	if err != nil {
		t.Fatalf("CloneSheetFrom: %v", err)
	}

	// The cloned rel should have been remapped to a unique path.
	if cloned.extraRels[0].Target == "../drawings/drawing1.xml" {
		t.Fatal("expected target to be remapped due to collision")
	}
	if cloned.extraRels[0].Target != "../drawings/drawing1_2.xml" {
		t.Fatalf("remapped target = %q, want ../drawings/drawing1_2.xml", cloned.extraRels[0].Target)
	}

	// Both entries should exist in destination.
	if len(dstFile.opaqueEntries) != 2 {
		t.Fatalf("expected 2 opaque entries, got %d", len(dstFile.opaqueEntries))
	}
	original := dstFile.opaqueEntries[0]
	remapped := dstFile.opaqueEntries[1]
	if !bytes.Equal(original.Data, []byte("existing-drawing")) {
		t.Fatalf("original entry data = %q, want existing-drawing", original.Data)
	}
	if remapped.Path != "xl/drawings/drawing1_2.xml" {
		t.Fatalf("remapped entry path = %q, want xl/drawings/drawing1_2.xml", remapped.Path)
	}
	if !bytes.Equal(remapped.Data, []byte("src-drawing")) {
		t.Fatalf("remapped entry data = %q, want src-drawing", remapped.Data)
	}
}

func TestCopySheetErrors(t *testing.T) {
	f := New()

	if _, err := f.CopySheet("Missing", "Copy"); !errors.Is(err, ErrSheetNotFound) {
		t.Fatalf("CopySheet missing source error = %v, want ErrSheetNotFound", err)
	}

	if _, err := f.CopySheet("Sheet1", "Sheet1"); err == nil {
		t.Fatal("CopySheet should reject duplicate destination names")
	}

	if _, err := f.CloneSheetFrom(nil, "Copy"); err == nil {
		t.Fatal("CloneSheetFrom should reject nil source sheets")
	}
}
