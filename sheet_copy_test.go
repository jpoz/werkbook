package werkbook

import (
	"archive/zip"
	"bytes"
	"errors"
	"io"
	"path"
	"slices"
	"strings"
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

func TestResolveRelTarget(t *testing.T) {
	tests := []struct {
		target string
		want   string
	}{
		{"../drawings/drawing1.xml", "xl/drawings/drawing1.xml"},
		{"../../customXml/item1.xml", "customXml/item1.xml"},
		{"comments1.xml", "xl/worksheets/comments1.xml"},
		{"./comments1.xml", "xl/worksheets/comments1.xml"},
		// Targets rooted at the package.
		{"/xl/drawings/drawing1.xml", "xl/drawings/drawing1.xml"},
		{"/xl/drawings/../media/image1.png", "xl/media/image1.png"},
		// A target must not climb above the package root.
		{"../../../../etc/passwd", "etc/passwd"},
		{"/../xl/media/image1.png", "xl/media/image1.png"},
	}
	for _, tt := range tests {
		if got := resolveRelTarget(tt.target); got != tt.want {
			t.Errorf("resolveRelTarget(%q) = %q, want %q", tt.target, got, tt.want)
		}
	}
}

const (
	testRelTypeDrawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
	testRelTypeChart   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
	testRelTypeImage   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	testRelTypeStyle   = "http://schemas.microsoft.com/office/2011/relationships/chartStyle"

	testDrawingXML = `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>`
)

func testRelsXML(t *testing.T, rels ...ooxml.OpaqueRel) []byte {
	t.Helper()
	data, err := ooxml.MarshalRelationships(rels)
	if err != nil {
		t.Fatal(err)
	}
	return data
}

// drawingWorkbook returns a workbook whose first sheet has a drawing with a
// chart (which has a style part) and an image. tag makes the content of the
// chart, the style, and the image different between workbooks.
func drawingWorkbook(t *testing.T, sheetName, tag string) *File {
	t.Helper()
	f := New(FirstSheet(sheetName))
	s := f.Sheet(sheetName)
	if err := s.SetValue("A1", tag); err != nil {
		t.Fatal(err)
	}
	s.rootAttrs = []ooxml.RawAttr{
		{Name: "xmlns", Value: ooxml.NSSpreadsheetML},
		{Name: "xmlns:r", Value: ooxml.NSOfficeDocument},
	}
	s.extraElements = []ooxml.RawElement{
		{Name: "drawing", XML: `<drawing r:id="rId1"/>`, OrderKey: 58},
	}
	s.extraRels = []ooxml.OpaqueRel{
		{ID: "rId1", Type: testRelTypeDrawing, Target: "../drawings/drawing1.xml"},
	}
	f.opaqueEntries = []ooxml.OpaqueEntry{
		{
			Path:        "xl/drawings/drawing1.xml",
			ContentType: "application/vnd.openxmlformats-officedocument.drawing+xml",
			Data:        []byte(testDrawingXML),
		},
		{
			Path: "xl/drawings/_rels/drawing1.xml.rels",
			Data: testRelsXML(t,
				ooxml.OpaqueRel{ID: "rId1", Type: testRelTypeChart, Target: "../charts/chart1.xml"},
				ooxml.OpaqueRel{ID: "rId2", Type: testRelTypeImage, Target: "/xl/media/image1.png"},
				ooxml.OpaqueRel{ID: "rId3", Type: "hyperlink", Target: "https://example.com/", TargetMode: "External"},
			),
		},
		{
			Path:        "xl/charts/chart1.xml",
			ContentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
			Data:        []byte("<chart>" + tag + "</chart>"),
		},
		{
			Path: "xl/charts/_rels/chart1.xml.rels",
			Data: testRelsXML(t,
				ooxml.OpaqueRel{ID: "rId1", Type: testRelTypeStyle, Target: "style1.xml"},
			),
		},
		{
			Path:        "xl/charts/style1.xml",
			ContentType: "application/vnd.ms-office.chartstyle+xml",
			Data:        []byte("<style>" + tag + "</style>"),
		},
		{Path: "xl/media/image1.png", Data: []byte("png-" + tag)},
	}
	f.opaqueDefaults = []ooxml.OpaqueDefault{{Extension: "png", ContentType: "image/png"}}
	return f
}

// savedParts saves the workbook and returns the content of each zip entry.
func savedParts(t *testing.T, f *File) map[string][]byte {
	t.Helper()
	var buf bytes.Buffer
	if _, err := f.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo: %v", err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	parts := make(map[string][]byte, len(zr.File))
	for _, zf := range zr.File {
		rc, err := zf.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		if _, dup := parts[zf.Name]; dup {
			t.Fatalf("zip has two entries with the name %s", zf.Name)
		}
		parts[zf.Name] = data
	}
	return parts
}

// relTargetsOf returns the package paths of the internal relationship targets
// of a part, by relationship ID.
func relTargetsOf(t *testing.T, parts map[string][]byte, partPath string) map[string]string {
	t.Helper()
	data, ok := parts[partRelsPath(partPath)]
	if !ok {
		t.Fatalf("%s has no relationships file", partPath)
	}
	rels, err := ooxml.ParseRelationships(data)
	if err != nil {
		t.Fatal(err)
	}
	targets := make(map[string]string)
	for _, rel := range rels {
		if rel.TargetMode == "External" {
			continue
		}
		targets[rel.ID] = resolvePartRelTarget(partPath, rel.Target)
	}
	return targets
}

// requireRelationshipsResolve fails if a relationships file in the saved
// package refers to a part that is not in the package.
func requireRelationshipsResolve(t *testing.T, parts map[string][]byte) {
	t.Helper()
	for name, data := range parts {
		dir, base := path.Split(name)
		if !strings.HasSuffix(dir, "_rels/") || !strings.HasSuffix(base, ".rels") {
			continue
		}
		owner := strings.TrimSuffix(dir, "_rels/") + strings.TrimSuffix(base, ".rels")
		rels, err := ooxml.ParseRelationships(data)
		if err != nil {
			t.Fatalf("%s: %v", name, err)
		}
		for _, rel := range rels {
			if rel.TargetMode == "External" {
				continue
			}
			target := resolvePartRelTarget(owner, rel.Target)
			if _, ok := parts[target]; !ok {
				t.Errorf("%s: relationship %s refers to %s, which is not in the package", name, rel.ID, target)
			}
		}
	}
}

// requireSheetDrawingGraph follows the drawing of the given sheet part in the
// saved package and checks the content of its chart, chart style, and image.
func requireSheetDrawingGraph(t *testing.T, parts map[string][]byte, sheetPart, tag string) {
	t.Helper()
	drawing := relTargetsOf(t, parts, sheetPart)["rId1"]
	if string(parts[drawing]) != testDrawingXML {
		t.Fatalf("%s: drawing %s = %q", sheetPart, drawing, parts[drawing])
	}
	drawingTargets := relTargetsOf(t, parts, drawing)
	chart := drawingTargets["rId1"]
	if got, want := string(parts[chart]), "<chart>"+tag+"</chart>"; got != want {
		t.Errorf("%s: chart %s = %q, want %q", sheetPart, chart, got, want)
	}
	image := drawingTargets["rId2"]
	if got, want := string(parts[image]), "png-"+tag; got != want {
		t.Errorf("%s: image %s = %q, want %q", sheetPart, image, got, want)
	}
	style := relTargetsOf(t, parts, chart)["rId1"]
	if got, want := string(parts[style]), "<style>"+tag+"</style>"; got != want {
		t.Errorf("%s: chart style %s = %q, want %q", sheetPart, style, got, want)
	}
}

func TestCloneSheetFromCopiesDrawingRelationshipGraph(t *testing.T) {
	src := drawingWorkbook(t, "Source", "src")
	dst := New(FirstSheet("Existing"))

	if _, err := dst.CloneSheetFrom(src.Sheet("Source"), "Imported"); err != nil {
		t.Fatalf("CloneSheetFrom: %v", err)
	}

	parts := savedParts(t, dst)
	requireRelationshipsResolve(t, parts)
	requireSheetDrawingGraph(t, parts, "xl/worksheets/sheet2.xml", "src")
	// With no collision, the parts keep their paths.
	for _, name := range []string{
		"xl/drawings/drawing1.xml",
		"xl/charts/chart1.xml",
		"xl/charts/style1.xml",
		"xl/media/image1.png",
	} {
		if _, ok := parts[name]; !ok {
			t.Errorf("%s is not in the saved package", name)
		}
	}
}

// The two workbooks have charts and images at the same paths with different
// content. The parts of the two sheets must stay apart.
func TestCloneSheetFromRemapsCollidingDrawingGraph(t *testing.T) {
	src := drawingWorkbook(t, "Source", "src")
	dst := drawingWorkbook(t, "Existing", "dst")

	if _, err := dst.CloneSheetFrom(src.Sheet("Source"), "Imported"); err != nil {
		t.Fatalf("CloneSheetFrom: %v", err)
	}

	parts := savedParts(t, dst)
	requireRelationshipsResolve(t, parts)
	requireSheetDrawingGraph(t, parts, "xl/worksheets/sheet1.xml", "dst")
	requireSheetDrawingGraph(t, parts, "xl/worksheets/sheet2.xml", "src")

	// The drawing XML is the same in the two workbooks, but its assets are
	// not. The imported sheet must not use the drawing of the destination.
	existing := relTargetsOf(t, parts, "xl/worksheets/sheet1.xml")["rId1"]
	imported := relTargetsOf(t, parts, "xl/worksheets/sheet2.xml")["rId1"]
	if existing == imported {
		t.Fatalf("the two sheets use the same drawing part %s", existing)
	}

	// The saved file opens, and a second save keeps the graph.
	var buf bytes.Buffer
	if _, err := dst.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	reopened, err := OpenReader(&buf)
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	parts = savedParts(t, reopened)
	requireRelationshipsResolve(t, parts)
	requireSheetDrawingGraph(t, parts, "xl/worksheets/sheet1.xml", "dst")
	requireSheetDrawingGraph(t, parts, "xl/worksheets/sheet2.xml", "src")
}

// A graph that is the same in the two workbooks is not copied again.
func TestCloneSheetFromReusesIdenticalDrawingGraph(t *testing.T) {
	src := drawingWorkbook(t, "Source", "same")
	dst := drawingWorkbook(t, "Existing", "same")
	before := len(dst.opaqueEntries)

	cloned, err := dst.CloneSheetFrom(src.Sheet("Source"), "Imported")
	if err != nil {
		t.Fatalf("CloneSheetFrom: %v", err)
	}
	if len(dst.opaqueEntries) != before {
		t.Fatalf("opaque entries = %d, want %d", len(dst.opaqueEntries), before)
	}
	if got := cloned.extraRels[0].Target; got != "../drawings/drawing1.xml" {
		t.Fatalf("drawing target = %q", got)
	}
	parts := savedParts(t, dst)
	requireRelationshipsResolve(t, parts)
}

// Only one asset is different. The parts that are the same are used again,
// and the parts on the path to the different asset get new names.
func TestCloneSheetFromRemapsOnlyChangedParts(t *testing.T) {
	src := drawingWorkbook(t, "Source", "same")
	dst := drawingWorkbook(t, "Existing", "same")
	for i := range src.opaqueEntries {
		if src.opaqueEntries[i].Path == "xl/charts/style1.xml" {
			src.opaqueEntries[i].Data = []byte("<style>other</style>")
		}
	}

	if _, err := dst.CloneSheetFrom(src.Sheet("Source"), "Imported"); err != nil {
		t.Fatalf("CloneSheetFrom: %v", err)
	}

	parts := savedParts(t, dst)
	requireRelationshipsResolve(t, parts)
	requireSheetDrawingGraph(t, parts, "xl/worksheets/sheet1.xml", "same")

	drawing := relTargetsOf(t, parts, "xl/worksheets/sheet2.xml")["rId1"]
	drawingTargets := relTargetsOf(t, parts, drawing)
	chart := drawingTargets["rId1"]
	style := relTargetsOf(t, parts, chart)["rId1"]
	if got := string(parts[style]); got != "<style>other</style>" {
		t.Fatalf("imported chart style %s = %q", style, got)
	}
	if drawing == "xl/drawings/drawing1.xml" || chart == "xl/charts/chart1.xml" || style == "xl/charts/style1.xml" {
		t.Fatalf("the imported graph uses parts of the destination: %s, %s, %s", drawing, chart, style)
	}
	// The image is the same, so there is one copy.
	if image := drawingTargets["rId2"]; image != "xl/media/image1.png" {
		t.Fatalf("image = %s, want the image of the destination", image)
	}
	if _, ok := parts["xl/media/image1_2.png"]; ok {
		t.Fatal("the image was copied, but it is the same in the two workbooks")
	}
}
