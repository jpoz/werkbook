package ooxml

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"strings"
	"testing"
)

// ---------------------------------------------------------------------------
// Opaque zip entries survive round-trip
// ---------------------------------------------------------------------------

func TestRoundTrip_OpaqueEntries(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}},
		}},
		OpaqueEntries: []OpaqueEntry{
			{
				Path:        "xl/charts/chart1.xml",
				ContentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
				Data:        []byte(`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`),
			},
			{
				Path:        "xl/media/image1.png",
				ContentType: "image/png",
				Data:        []byte("fakepngdata"),
			},
		},
		OpaqueDefaults: []OpaqueDefault{
			{Extension: "png", ContentType: "image/png"},
		},
	}

	got := writeAndRead(t, data)

	if len(got.OpaqueEntries) < 2 {
		t.Fatalf("OpaqueEntries = %d, want >= 2", len(got.OpaqueEntries))
	}
	foundChart := false
	foundMedia := false
	for _, entry := range got.OpaqueEntries {
		switch entry.Path {
		case "xl/charts/chart1.xml":
			foundChart = true
			if entry.ContentType == "" {
				t.Error("chart entry lost its content type")
			}
		case "xl/media/image1.png":
			foundMedia = true
			if string(entry.Data) != "fakepngdata" {
				t.Errorf("image data = %q, want fakepngdata", string(entry.Data))
			}
		}
	}
	if !foundChart {
		t.Error("chart entry did not survive round-trip")
	}
	if !foundMedia {
		t.Error("media entry did not survive round-trip")
	}

	// Verify content types were written
	var buf bytes.Buffer
	if err := WriteWorkbook(&buf, data); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	ctData := readZipEntry(t, zr, "[Content_Types].xml")
	if !strings.Contains(ctData, "drawingml.chart+xml") {
		t.Error("[Content_Types].xml missing chart override")
	}
	if !strings.Contains(ctData, "image/png") {
		t.Error("[Content_Types].xml missing png default")
	}
}

// ---------------------------------------------------------------------------
// Opaque defaults for media extensions
// ---------------------------------------------------------------------------

func TestRoundTrip_OpaqueDefaults(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}},
		}},
		OpaqueDefaults: []OpaqueDefault{
			{Extension: "png", ContentType: "image/png"},
			{Extension: "jpeg", ContentType: "image/jpeg"},
		},
	}
	got := writeAndRead(t, data)

	pngFound := false
	jpegFound := false
	for _, d := range got.OpaqueDefaults {
		switch strings.ToLower(d.Extension) {
		case "png":
			pngFound = true
		case "jpeg":
			jpegFound = true
		}
	}
	if !pngFound {
		t.Error("png default did not survive round-trip")
	}
	if !jpegFound {
		t.Error("jpeg default did not survive round-trip")
	}
}

// ---------------------------------------------------------------------------
// Extra sheet rels (e.g. drawing relationship) survive round-trip
// ---------------------------------------------------------------------------

func TestRoundTrip_ExtraSheetRels(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}},
			ExtraRels: []OpaqueRel{
				{
					ID:     "rId1",
					Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
					Target: "../drawings/drawing1.xml",
				},
			},
		}},
	}
	got := writeAndRead(t, data)

	if len(got.Sheets[0].ExtraRels) != 1 {
		t.Fatalf("ExtraRels = %d, want 1", len(got.Sheets[0].ExtraRels))
	}
	rel := got.Sheets[0].ExtraRels[0]
	if rel.Type != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" {
		t.Errorf("rel.Type = %q", rel.Type)
	}
	if rel.Target != "../drawings/drawing1.xml" {
		t.Errorf("rel.Target = %q", rel.Target)
	}
}

// ---------------------------------------------------------------------------
// Tables + extra rels on same sheet: rIds don't collide
// ---------------------------------------------------------------------------

func TestRoundTrip_TablesAndExtraRels(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{
				{Num: 1, Cells: []CellData{
					{Ref: "A1", Type: "s", Value: "Col"},
				}},
				{Num: 2, Cells: []CellData{
					{Ref: "A2", Value: "1"},
				}},
			},
			ExtraRels: []OpaqueRel{
				{
					ID:     "rId1",
					Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
					Target: "../drawings/drawing1.xml",
				},
			},
		}},
		Tables: []TableDef{{
			Name:           "T1",
			DisplayName:    "T1",
			Ref:            "A1:A2",
			SheetIndex:     0,
			Columns:        []string{"Col"},
			HeaderRowCount: 1,
		}},
	}

	// Write and inspect the raw rels XML to verify no rId collision.
	var buf bytes.Buffer
	if err := WriteWorkbook(&buf, data); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}

	relsXML := readZipEntry(t, zr, "xl/worksheets/_rels/sheet1.xml.rels")

	// Parse the rels to check for rId uniqueness.
	var rels xlsxRelationships
	if err := xml.Unmarshal([]byte(relsXML), &rels); err != nil {
		t.Fatal(err)
	}
	seen := make(map[string]bool)
	for _, r := range rels.Relationships {
		if seen[r.ID] {
			t.Errorf("duplicate rId %q in sheet rels", r.ID)
		}
		seen[r.ID] = true
	}
	if len(rels.Relationships) < 2 {
		t.Errorf("expected at least 2 relationships (drawing + table), got %d", len(rels.Relationships))
	}

	// Verify the sheet XML tablePart rId matches the rels entry for the table.
	sheetXML := readZipEntry(t, zr, "xl/worksheets/sheet1.xml")
	tableRelID := ""
	for _, r := range rels.Relationships {
		if strings.Contains(r.Target, "table") {
			tableRelID = r.ID
			break
		}
	}
	if tableRelID == "" {
		t.Fatal("no table relationship found in rels")
	}
	if !strings.Contains(sheetXML, `r:id="`+tableRelID+`"`) {
		t.Errorf("sheet XML tablePart r:id does not match rels entry %q", tableRelID)
	}

	// Round-trip: verify both the table and the drawing rel survive.
	got := writeAndRead(t, data)
	if len(got.Tables) != 1 {
		t.Errorf("Tables = %d, want 1", len(got.Tables))
	}
	if len(got.Sheets[0].ExtraRels) != 1 {
		t.Errorf("ExtraRels = %d, want 1", len(got.Sheets[0].ExtraRels))
	}
}

// ---------------------------------------------------------------------------
// Extra sheet XML elements survive round-trip
// ---------------------------------------------------------------------------

func TestRoundTrip_ExtraSheetElements(t *testing.T) {
	drawingXML := `<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>`
	cfXML := `<conditionalFormatting sqref="A1:A10"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>`

	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}},
			ExtraElements: []RawElement{
				{
					Name:     "conditionalFormatting",
					XML:      cfXML,
					OrderKey: worksheetElementOrder["conditionalFormatting"],
					Seq:      0,
				},
				{
					Name:     "drawing",
					XML:      drawingXML,
					OrderKey: worksheetElementOrder["drawing"],
					Seq:      1,
				},
			},
		}},
	}

	var buf bytes.Buffer
	if err := WriteWorkbook(&buf, data); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}

	sheetXML := readZipEntry(t, zr, "xl/worksheets/sheet1.xml")
	if !strings.Contains(sheetXML, "conditionalFormatting") {
		t.Error("sheet XML missing conditionalFormatting element")
	}
	if !strings.Contains(sheetXML, `r:id="rId1"`) {
		t.Error("sheet XML missing drawing r:id")
	}

	// Verify ordering: conditionalFormatting before drawing, both after sheetData
	sdIdx := strings.Index(sheetXML, "sheetData")
	cfIdx := strings.Index(sheetXML, "conditionalFormatting")
	drIdx := strings.Index(sheetXML, "drawing")
	if sdIdx >= cfIdx {
		t.Error("sheetData should come before conditionalFormatting")
	}
	if cfIdx >= drIdx {
		t.Error("conditionalFormatting should come before drawing")
	}
}

// ---------------------------------------------------------------------------
// Workbook-level extra elements survive round-trip
// ---------------------------------------------------------------------------

func TestRoundTrip_WorkbookExtraElements(t *testing.T) {
	bookViewsXML := `<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"/></bookViews>`

	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}},
		}},
		ExtraElements: []RawElement{
			{
				Name:     "bookViews",
				XML:      bookViewsXML,
				OrderKey: workbookElementOrder["bookViews"],
				Seq:      0,
			},
		},
	}

	var buf bytes.Buffer
	if err := WriteWorkbook(&buf, data); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	wbXML := readZipEntry(t, zr, "xl/workbook.xml")
	if !strings.Contains(wbXML, "bookViews") {
		t.Error("workbook XML missing bookViews element")
	}
	if !strings.Contains(wbXML, "windowWidth") {
		t.Error("bookViews content was not preserved")
	}

	// Verify ordering: bookViews before sheets
	bvIdx := strings.Index(wbXML, "bookViews")
	shIdx := strings.Index(wbXML, "<sheets>")
	if bvIdx >= shIdx {
		t.Error("bookViews should come before sheets")
	}
}

// ---------------------------------------------------------------------------
// Extra root and workbook rels survive round-trip
// ---------------------------------------------------------------------------

func TestRoundTrip_ExtraWorkbookRels(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}},
		}},
		ExtraRels: []OpaqueRel{
			{
				ID:     "rId99",
				Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
				Target: "customXml/item1.xml",
			},
		},
		ExtraRootRels: []OpaqueRel{
			{
				ID:     "rId50",
				Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
				Target: "docProps/custom.xml",
			},
		},
	}

	got := writeAndRead(t, data)

	if len(got.ExtraRels) != 1 {
		t.Fatalf("ExtraRels = %d, want 1", len(got.ExtraRels))
	}
	if got.ExtraRels[0].ID != "rId99" {
		t.Errorf("ExtraRels[0].ID = %q, want rId99", got.ExtraRels[0].ID)
	}
	if len(got.ExtraRootRels) != 1 {
		t.Fatalf("ExtraRootRels = %d, want 1", len(got.ExtraRootRels))
	}
	if got.ExtraRootRels[0].ID != "rId50" {
		t.Errorf("ExtraRootRels[0].ID = %q, want rId50", got.ExtraRootRels[0].ID)
	}
}

// ---------------------------------------------------------------------------
// TargetMode is preserved on rels (needed for external hyperlinks)
// ---------------------------------------------------------------------------

func TestRoundTrip_RelTargetMode(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1"}}}},
			ExtraRels: []OpaqueRel{
				{
					ID:         "rId1",
					Type:       "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
					Target:     "https://example.com",
					TargetMode: "External",
				},
			},
		}},
	}
	got := writeAndRead(t, data)
	if len(got.Sheets[0].ExtraRels) != 1 {
		t.Fatal("ExtraRels not preserved")
	}
	if got.Sheets[0].ExtraRels[0].TargetMode != "External" {
		t.Errorf("TargetMode = %q, want External", got.Sheets[0].ExtraRels[0].TargetMode)
	}
}

// ---------------------------------------------------------------------------
// Double-write prevention: metadata.xml not duplicated
// ---------------------------------------------------------------------------

func TestRoundTrip_MetadataNotDuplicated(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{
				Ref:            "A1",
				Formula:        "SORT(B1:B10)",
				FormulaType:    "array",
				FormulaRef:     "A1:A10",
				IsDynamicArray: true,
			}}}},
		}},
		// Simulate xl/metadata.xml captured as opaque (it would be in a
		// real file opened from disk).
		OpaqueEntries: []OpaqueEntry{
			{
				Path:        "xl/metadata.xml",
				ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
				Data:        []byte(dynamicArrayMetadataXML),
			},
		},
	}

	var buf bytes.Buffer
	if err := WriteWorkbook(&buf, data); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}

	// Count how many times xl/metadata.xml appears.
	count := 0
	for _, f := range zr.File {
		if f.Name == "xl/metadata.xml" {
			count++
		}
	}
	if count != 1 {
		t.Errorf("xl/metadata.xml appears %d times, want 1", count)
	}
}

// ---------------------------------------------------------------------------
// nextRelationshipIDs
// ---------------------------------------------------------------------------

func TestNextRelationshipIDs_Basic(t *testing.T) {
	ids := nextRelationshipIDs(nil, 3)
	if len(ids) != 3 {
		t.Fatalf("len = %d, want 3", len(ids))
	}
	if ids[0] != "rId1" || ids[1] != "rId2" || ids[2] != "rId3" {
		t.Errorf("ids = %v", ids)
	}
}

func TestNextRelationshipIDs_SkipsExisting(t *testing.T) {
	existing := []OpaqueRel{
		{ID: "rId1"},
		{ID: "rId3"},
	}
	ids := nextRelationshipIDs(existing, 2)
	if len(ids) != 2 {
		t.Fatalf("len = %d, want 2", len(ids))
	}
	// Should skip rId1, rId3, rId4 (starts after max=3)
	if ids[0] != "rId4" || ids[1] != "rId5" {
		t.Errorf("ids = %v, want [rId4, rId5]", ids)
	}
}

func TestNextRelationshipIDs_AlreadyUsed(t *testing.T) {
	existing := []OpaqueRel{{ID: "rId1"}}
	ids := nextRelationshipIDs(existing, 2, "rId2", "rId3")
	// max is 3 (from alreadyUsed rId3), so starts at rId4
	if ids[0] != "rId4" || ids[1] != "rId5" {
		t.Errorf("ids = %v, want [rId4, rId5]", ids)
	}
}

func TestNextRelationshipIDs_Zero(t *testing.T) {
	ids := nextRelationshipIDs(nil, 0)
	if ids != nil {
		t.Errorf("expected nil, got %v", ids)
	}
}

// ---------------------------------------------------------------------------
// captureRootAttrsAndExtras
// ---------------------------------------------------------------------------

func TestCaptureRootAttrsAndExtras_Worksheet(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
<drawing r:id="rId1"/>
</worksheet>`

	attrs, extras, err := captureRootAttrsAndExtras(
		[]byte(xml),
		func(name string) bool {
			return name == "sheetData"
		},
		worksheetOrderKey,
	)
	if err != nil {
		t.Fatal(err)
	}

	// Should capture xmlns, xmlns:r, xmlns:mc attrs.
	if len(attrs) < 2 {
		t.Errorf("attrs = %d, want >= 2", len(attrs))
	}

	// sheetViews and drawing should be extras (sheetData is known).
	if len(extras) != 2 {
		t.Fatalf("extras = %d, want 2", len(extras))
	}
	names := map[string]bool{}
	for _, e := range extras {
		names[e.Name] = true
	}
	if !names["sheetViews"] {
		t.Error("missing sheetViews in extras")
	}
	if !names["drawing"] {
		t.Error("missing drawing in extras")
	}

	// drawing should have correct order key.
	for _, e := range extras {
		if e.Name == "drawing" && e.OrderKey != worksheetElementOrder["drawing"] {
			t.Errorf("drawing OrderKey = %d, want %d", e.OrderKey, worksheetElementOrder["drawing"])
		}
	}
}

// ---------------------------------------------------------------------------
// isRecognizedRootRelType
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// workbookPr is preserved and reconciled with Date1904
// ---------------------------------------------------------------------------

func TestWorkbookPrFragment(t *testing.T) {
	raw := func(xml string) []RawElement {
		return []RawElement{{Name: "workbookPr", XML: xml, OrderKey: workbookElementOrder["workbookPr"]}}
	}
	tests := []struct {
		name     string
		extras   []RawElement
		date1904 bool
		want     string
	}{
		{"no element, 1900 system", nil, false, ""},
		{"no element, 1904 system", nil, true, `<workbookPr date1904="1"/>`},
		{"preserved verbatim", raw(`<workbookPr defaultThemeVersion="166925" filterPrivacy="1"/>`), false, `<workbookPr defaultThemeVersion="166925" filterPrivacy="1"/>`},
		{"preserved verbatim with date1904", raw(`<workbookPr date1904="true" codeName="ThisWorkbook"/>`), true, `<workbookPr date1904="true" codeName="ThisWorkbook"/>`},
		{"explicit date1904=0 agrees", raw(`<workbookPr date1904="0"/>`), false, `<workbookPr date1904="0"/>`},
		{"date1904 switched on", raw(`<workbookPr defaultThemeVersion="166925"/>`), true, `<workbookPr date1904="1" defaultThemeVersion="166925"/>`},
		{"date1904 switched off", raw(`<workbookPr date1904="1" defaultThemeVersion="166925"/>`), false, `<workbookPr defaultThemeVersion="166925"/>`},
		{"date1904 switched off, no other attrs", raw(`<workbookPr date1904="1"/>`), false, ""},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := workbookPrFragment(tt.extras, tt.date1904)
			if err != nil {
				t.Fatal(err)
			}
			if string(got) != tt.want {
				t.Errorf("workbookPrFragment = %q, want %q", got, tt.want)
			}
		})
	}
}

func TestRoundTrip_WorkbookPrPreserved(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{Name: "Sheet1"}},
		ExtraElements: []RawElement{{
			Name:     "workbookPr",
			XML:      `<workbookPr defaultThemeVersion="166925"/>`,
			OrderKey: workbookElementOrder["workbookPr"],
		}},
	}

	got := writeAndRead(t, data)
	if got.Date1904 {
		t.Error("Date1904 = true, want false")
	}
	if !hasExtraElementName(got.ExtraElements, "workbookPr") {
		t.Fatal("workbookPr did not survive round-trip")
	}

	// Switching the date system keeps the other attributes and writes
	// exactly one workbookPr element.
	got.Date1904 = true
	var buf bytes.Buffer
	if err := WriteWorkbook(&buf, got); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	wbXML := readZipEntry(t, zr, "xl/workbook.xml")
	if n := strings.Count(wbXML, "<workbookPr"); n != 1 {
		t.Fatalf("workbookPr count = %d, want 1: %s", n, wbXML)
	}
	if !strings.Contains(wbXML, `<workbookPr date1904="1" defaultThemeVersion="166925"/><sheets>`) {
		t.Fatalf("workbookPr not reconciled with Date1904: %s", wbXML)
	}
	again, err := ReadWorkbook(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	if !again.Date1904 {
		t.Error("Date1904 = false after switch, want true")
	}
}

// ---------------------------------------------------------------------------
// Unparsed styleSheet elements survive round-trip
// ---------------------------------------------------------------------------

func TestRoundTrip_StyleExtraElements(t *testing.T) {
	const dxfs = `<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>`
	const extLst = `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst>`
	data := &WorkbookData{
		// A styled cell forces numFmts and a second cellXf, so the extras
		// must sort after every rebuilt element.
		Styles: []StyleData{{}, {NumFmt: "0.000", FontBold: true}},
		Sheets: []SheetData{{
			Name: "Sheet1",
			Rows: []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Value: "1", StyleIdx: 1}}}},
		}},
		StyleRootAttrs: []RawAttr{
			{Name: "xmlns", Value: NSSpreadsheetML},
			{Name: "xmlns:mc", Value: "http://schemas.openxmlformats.org/markup-compatibility/2006"},
			{Name: "xmlns:x14ac", Value: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"},
			{Name: "mc:Ignorable", Value: "x14ac"},
		},
		StyleExtraElements: []RawElement{
			{Name: "dxfs", XML: dxfs, OrderKey: styleSheetElementOrder["dxfs"], Seq: 0},
			{Name: "extLst", XML: extLst, OrderKey: styleSheetElementOrder["extLst"], Seq: 1},
		},
	}

	var buf bytes.Buffer
	if err := WriteWorkbook(&buf, data); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	stylesXML := readZipEntry(t, zr, "xl/styles.xml")
	if !strings.Contains(stylesXML, `mc:Ignorable="x14ac"`) {
		t.Errorf("styleSheet root attrs not preserved: %s", stylesXML)
	}
	if !strings.HasSuffix(stylesXML, "</cellStyles>"+dxfs+extLst+"</styleSheet>") {
		t.Errorf("style extras not written verbatim after cellStyles: %s", stylesXML)
	}
	if strings.Index(stylesXML, "<numFmts") > strings.Index(stylesXML, "<fonts") {
		t.Errorf("numFmts must precede fonts: %s", stylesXML)
	}

	got, err := ReadWorkbook(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	if len(got.StyleExtraElements) != 2 ||
		got.StyleExtraElements[0].XML != dxfs ||
		got.StyleExtraElements[1].XML != extLst {
		t.Errorf("StyleExtraElements = %+v, want dxfs and extLst", got.StyleExtraElements)
	}
	if rawAttrValue(got.StyleRootAttrs, "mc:Ignorable") != "x14ac" {
		t.Errorf("StyleRootAttrs = %+v, want mc:Ignorable preserved", got.StyleRootAttrs)
	}
	// The parsed cell style is still rebuilt from the model.
	if len(got.Styles) < 2 || got.Styles[1].NumFmt != "0.000" || !got.Styles[1].FontBold {
		t.Errorf("Styles = %+v, want the styled cell format preserved", got.Styles)
	}
	for _, entry := range got.OpaqueEntries {
		if entry.Path == "xl/styles.xml" {
			t.Error("xl/styles.xml must not also be kept as an opaque entry")
		}
	}
}

func TestWriteStyles_NoPassthroughMatchesModelOnly(t *testing.T) {
	data := &WorkbookData{
		Styles: []StyleData{{}},
		Sheets: []SheetData{{Name: "Sheet1"}},
	}
	var buf bytes.Buffer
	if err := WriteWorkbook(&buf, data); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	got := readZipEntry(t, zr, "xl/styles.xml")

	var want bytes.Buffer
	enc := xml.NewEncoder(&want)
	if err := enc.Encode(NewStyleSheetBuilder().Build()); err != nil {
		t.Fatal(err)
	}
	if wantXML := xmlHeader + string(selfCloseEmptyElements(want.Bytes())); got != wantXML {
		t.Errorf("styles.xml changed for a workbook with no style passthrough:\n got: %s\nwant: %s", got, wantXML)
	}
}

func TestIsRecognizedRootRelType(t *testing.T) {
	tests := []struct {
		relType string
		want    bool
	}{
		{RelTypeWorkbook, true},
		{RelTypeCoreProps, true},
		{RelTypeExtendedApp, true},
		// Strict workbook
		{NSOfficeDocumentStrict + "/officeDocument", true},
		// Unknown
		{"http://example.com/custom", false},
	}
	for _, tt := range tests {
		if got := isRecognizedRootRelType(tt.relType); got != tt.want {
			t.Errorf("isRecognizedRootRelType(%q) = %v, want %v", tt.relType, got, tt.want)
		}
	}
}

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

func readZipEntry(t *testing.T, zr *zip.Reader, name string) string {
	t.Helper()
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				t.Fatalf("open %s: %v", name, err)
			}
			var buf bytes.Buffer
			buf.ReadFrom(rc)
			rc.Close()
			return buf.String()
		}
	}
	t.Fatalf("zip entry %q not found", name)
	return ""
}
