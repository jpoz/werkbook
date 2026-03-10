package ooxml

import (
	"archive/zip"
	"encoding/xml"
	"strings"
	"testing"
)

func TestNewStyleSheetBuilder_Defaults(t *testing.T) {
	ssb := NewStyleSheetBuilder()
	ss := ssb.Build()

	if ss.Fonts.Count != 1 {
		t.Errorf("Fonts.Count = %d, want 1", ss.Fonts.Count)
	}
	if ss.Fills.Count != 2 {
		t.Errorf("Fills.Count = %d, want 2", ss.Fills.Count)
	}
	if ss.Borders.Count != 1 {
		t.Errorf("Borders.Count = %d, want 1", ss.Borders.Count)
	}
	if ss.CellXfs.Count != 1 {
		t.Errorf("CellXfs.Count = %d, want 1", ss.CellXfs.Count)
	}
	if ss.NumFmts != nil {
		t.Error("NumFmts should be nil for default builder")
	}
}

func TestStyleSheetBuilder_AddStyle_Font(t *testing.T) {
	ssb := NewStyleSheetBuilder()
	idx := ssb.AddStyle(StyleData{FontName: "Arial", FontSize: 14, FontBold: true})
	if idx == 0 {
		t.Error("AddStyle returned 0, want non-zero for styled cell")
	}
	ss := ssb.Build()
	if ss.Fonts.Count != 2 {
		t.Errorf("Fonts.Count = %d, want 2", ss.Fonts.Count)
	}
	if ss.CellXfs.Count != 2 {
		t.Errorf("CellXfs.Count = %d, want 2", ss.CellXfs.Count)
	}
}

func TestStyleSheetBuilder_AddStyle_Fill(t *testing.T) {
	ssb := NewStyleSheetBuilder()
	idx := ssb.AddStyle(StyleData{FillColor: "FFFF0000"})
	if idx == 0 {
		t.Error("AddStyle returned 0 for fill style")
	}
	ss := ssb.Build()
	// 2 default fills + 1 custom
	if ss.Fills.Count != 3 {
		t.Errorf("Fills.Count = %d, want 3", ss.Fills.Count)
	}
}

func TestStyleSheetBuilder_AddStyle_Border(t *testing.T) {
	ssb := NewStyleSheetBuilder()
	idx := ssb.AddStyle(StyleData{
		BorderLeftStyle: "thin", BorderLeftColor: "FF000000",
	})
	if idx == 0 {
		t.Error("AddStyle returned 0 for border style")
	}
	ss := ssb.Build()
	if ss.Borders.Count != 2 {
		t.Errorf("Borders.Count = %d, want 2", ss.Borders.Count)
	}
}

func TestStyleSheetBuilder_AddStyle_Alignment(t *testing.T) {
	ssb := NewStyleSheetBuilder()
	idx := ssb.AddStyle(StyleData{HAlign: "center", VAlign: "top", WrapText: true})
	if idx == 0 {
		t.Error("AddStyle returned 0 for alignment style")
	}
	ss := ssb.Build()
	xf := ss.CellXfs.Xf[idx]
	if xf.ApplyAlignment == 0 {
		t.Error("ApplyAlignment = 0, want 1")
	}
	if xf.Alignment == nil {
		t.Fatal("Alignment is nil")
	}
	if xf.Alignment.Horizontal != "center" {
		t.Errorf("Horizontal = %q, want center", xf.Alignment.Horizontal)
	}
	if xf.Alignment.Vertical != "top" {
		t.Errorf("Vertical = %q, want top", xf.Alignment.Vertical)
	}
	if xf.Alignment.WrapText == 0 {
		t.Error("WrapText = 0, want 1")
	}
}

func TestStyleSheetBuilder_AddStyle_NumFmt(t *testing.T) {
	ssb := NewStyleSheetBuilder()
	idx := ssb.AddStyle(StyleData{NumFmt: "#,##0.00"})
	if idx == 0 {
		t.Error("AddStyle returned 0 for numfmt style")
	}
	ss := ssb.Build()
	if ss.NumFmts == nil {
		t.Fatal("NumFmts is nil")
	}
	if ss.NumFmts.Count != 1 {
		t.Errorf("NumFmts.Count = %d, want 1", ss.NumFmts.Count)
	}
	if ss.NumFmts.NumFmt[0].NumFmtID != 164 {
		t.Errorf("NumFmtID = %d, want 164", ss.NumFmts.NumFmt[0].NumFmtID)
	}
	if ss.NumFmts.NumFmt[0].FormatCode != "#,##0.00" {
		t.Errorf("FormatCode = %q, want #,##0.00", ss.NumFmts.NumFmt[0].FormatCode)
	}
}

func TestStyleSheetBuilder_AddStyle_BuiltinNumFmtID(t *testing.T) {
	ssb := NewStyleSheetBuilder()
	idx := ssb.AddStyle(StyleData{NumFmtID: 14})
	if idx == 0 {
		t.Error("AddStyle returned 0 for builtin numfmt")
	}
	ss := ssb.Build()
	if ss.NumFmts != nil {
		t.Error("NumFmts should be nil for built-in numfmt")
	}
	xf := ss.CellXfs.Xf[idx]
	if xf.NumFmtID != 14 {
		t.Errorf("NumFmtID = %d, want 14", xf.NumFmtID)
	}
}

func TestStyleSheetBuilder_Dedup(t *testing.T) {
	ssb := NewStyleSheetBuilder()
	sd := StyleData{FontName: "Arial", FontSize: 12, FontBold: true}
	idx1 := ssb.AddStyle(sd)
	idx2 := ssb.AddStyle(sd)
	if idx1 != idx2 {
		t.Errorf("dedup failed: idx1=%d, idx2=%d", idx1, idx2)
	}
	ss := ssb.Build()
	if ss.Fonts.Count != 2 {
		t.Errorf("Fonts.Count = %d, want 2 (default + 1 unique)", ss.Fonts.Count)
	}
	if ss.CellXfs.Count != 2 {
		t.Errorf("CellXfs.Count = %d, want 2 (default + 1 unique)", ss.CellXfs.Count)
	}
}

func TestStyleSheetBuilder_MultipleCustomNumFmts(t *testing.T) {
	ssb := NewStyleSheetBuilder()
	ssb.AddStyle(StyleData{NumFmt: "#,##0"})
	ssb.AddStyle(StyleData{NumFmt: "0.00%"})
	// Same as first, should dedup.
	ssb.AddStyle(StyleData{NumFmt: "#,##0"})

	ss := ssb.Build()
	if ss.NumFmts.Count != 2 {
		t.Errorf("NumFmts.Count = %d, want 2", ss.NumFmts.Count)
	}
}

func TestStyleSheetBuilder_XMLRoundTrip(t *testing.T) {
	ssb := NewStyleSheetBuilder()
	ssb.AddStyle(StyleData{
		FontName:          "Helvetica",
		FontSize:          16,
		FontBold:          true,
		FontColor:         "FFFF0000",
		FillColor:         "FF00FF00",
		BorderBottomStyle: "thin",
		BorderBottomColor: "FF000000",
		HAlign:            "center",
		WrapText:          true,
		NumFmt:            "#,##0.00",
	})

	ss := ssb.Build()

	// Marshal to XML.
	data, err := xml.MarshalIndent(ss, "", "  ")
	if err != nil {
		t.Fatal(err)
	}
	xmlStr := string(data)

	// Verify key elements are present.
	checks := []string{
		`<font>`,
		`Helvetica`,
		`16`,
		`FFFF0000`,
		`solid`,
		`FF00FF00`,
		`thin`,
		`center`,
		`wrapText`,
		`#,##0.00`,
	}
	for _, c := range checks {
		if !strings.Contains(xmlStr, c) {
			t.Errorf("XML output missing %q", c)
		}
	}

	// Unmarshal back.
	var ss2 xlsxStyleSheet
	if err := xml.Unmarshal(data, &ss2); err != nil {
		t.Fatal(err)
	}
	if ss2.Fonts.Count != ss.Fonts.Count {
		t.Errorf("round-trip Fonts.Count = %d, want %d", ss2.Fonts.Count, ss.Fonts.Count)
	}
	if ss2.CellXfs.Count != ss.CellXfs.Count {
		t.Errorf("round-trip CellXfs.Count = %d, want %d", ss2.CellXfs.Count, ss.CellXfs.Count)
	}
}

func TestReadStyles_Missing(t *testing.T) {
	// Empty files map should return default.
	files := make(map[string]*zip.File)
	styles := readStyles(files)
	if len(styles) != 1 {
		t.Errorf("len = %d, want 1", len(styles))
	}
	if styles[0] != (StyleData{}) {
		t.Errorf("expected zero StyleData, got %+v", styles[0])
	}
}
