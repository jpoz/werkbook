package werkbook

import (
	"testing"

	"github.com/jpoz/werkbook/ooxml"
)

func TestStyleToStyleData_Nil(t *testing.T) {
	sd := styleToStyleData(nil)
	if sd != (ooxml.StyleData{}) {
		t.Fatalf("expected zero StyleData for nil style, got %+v", sd)
	}
}

func TestStyleToStyleData_Font(t *testing.T) {
	s := &Style{
		Font: &Font{
			Name:      "Arial",
			Size:      14,
			Bold:      true,
			Italic:    true,
			Underline: true,
			Color:     "FF0000",
		},
	}
	sd := styleToStyleData(s)
	if sd.FontName != "Arial" {
		t.Errorf("FontName = %q, want %q", sd.FontName, "Arial")
	}
	if sd.FontSize != 14 {
		t.Errorf("FontSize = %g, want %g", sd.FontSize, 14.0)
	}
	if !sd.FontBold {
		t.Error("FontBold = false, want true")
	}
	if !sd.FontItalic {
		t.Error("FontItalic = false, want true")
	}
	if !sd.FontUL {
		t.Error("FontUL = false, want true")
	}
	if sd.FontColor != "FFFF0000" {
		t.Errorf("FontColor = %q, want %q", sd.FontColor, "FFFF0000")
	}
}

func TestStyleToStyleData_Fill(t *testing.T) {
	s := &Style{Fill: &Fill{Color: "00FF00"}}
	sd := styleToStyleData(s)
	if sd.FillColor != "FF00FF00" {
		t.Errorf("FillColor = %q, want %q", sd.FillColor, "FF00FF00")
	}
}

func TestStyleToStyleData_Border(t *testing.T) {
	s := &Style{
		Border: &Border{
			Left:   BorderSide{Style: BorderThin, Color: "000000"},
			Right:  BorderSide{Style: BorderMedium, Color: "111111"},
			Top:    BorderSide{Style: BorderThick, Color: "222222"},
			Bottom: BorderSide{Style: BorderDashed, Color: "333333"},
		},
	}
	sd := styleToStyleData(s)
	if sd.BorderLeftStyle != "thin" {
		t.Errorf("BorderLeftStyle = %q, want %q", sd.BorderLeftStyle, "thin")
	}
	if sd.BorderLeftColor != "FF000000" {
		t.Errorf("BorderLeftColor = %q, want %q", sd.BorderLeftColor, "FF000000")
	}
	if sd.BorderRightStyle != "medium" {
		t.Errorf("BorderRightStyle = %q, want %q", sd.BorderRightStyle, "medium")
	}
	if sd.BorderTopStyle != "thick" {
		t.Errorf("BorderTopStyle = %q, want %q", sd.BorderTopStyle, "thick")
	}
	if sd.BorderBottomStyle != "dashed" {
		t.Errorf("BorderBottomStyle = %q, want %q", sd.BorderBottomStyle, "dashed")
	}
}

func TestStyleToStyleData_Alignment(t *testing.T) {
	s := &Style{
		Alignment: &Alignment{
			Horizontal: HAlignCenter,
			Vertical:   VAlignTop,
			WrapText:   true,
		},
	}
	sd := styleToStyleData(s)
	if sd.HAlign != "center" {
		t.Errorf("HAlign = %q, want %q", sd.HAlign, "center")
	}
	if sd.VAlign != "top" {
		t.Errorf("VAlign = %q, want %q", sd.VAlign, "top")
	}
	if !sd.WrapText {
		t.Error("WrapText = false, want true")
	}
}

func TestStyleToStyleData_NumFmt(t *testing.T) {
	s := &Style{NumFmt: "#,##0.00"}
	sd := styleToStyleData(s)
	if sd.NumFmt != "#,##0.00" {
		t.Errorf("NumFmt = %q, want %q", sd.NumFmt, "#,##0.00")
	}

	s2 := &Style{NumFmtID: 14}
	sd2 := styleToStyleData(s2)
	if sd2.NumFmtID != 14 {
		t.Errorf("NumFmtID = %d, want %d", sd2.NumFmtID, 14)
	}
}

func TestStyleDataToStyle_ZeroValue(t *testing.T) {
	s := styleDataToStyle(ooxml.StyleData{})
	if s != nil {
		t.Fatalf("expected nil for zero StyleData, got %+v", s)
	}
}

func TestStyleDataToStyle_RoundTrip(t *testing.T) {
	original := &Style{
		Font: &Font{
			Name:      "Times New Roman",
			Size:      12,
			Bold:      true,
			Italic:    false,
			Underline: true,
			Color:     "0000FF",
		},
		Fill: &Fill{Color: "FFFF00"},
		Border: &Border{
			Left:   BorderSide{Style: BorderThin, Color: "000000"},
			Bottom: BorderSide{Style: BorderDouble, Color: "FF0000"},
		},
		Alignment: &Alignment{
			Horizontal: HAlignRight,
			Vertical:   VAlignCenter,
			WrapText:   true,
		},
		NumFmt: "0.00%",
	}
	sd := styleToStyleData(original)
	result := styleDataToStyle(sd)

	if result.Font.Name != original.Font.Name {
		t.Errorf("Font.Name = %q, want %q", result.Font.Name, original.Font.Name)
	}
	if result.Font.Size != original.Font.Size {
		t.Errorf("Font.Size = %g, want %g", result.Font.Size, original.Font.Size)
	}
	if result.Font.Bold != original.Font.Bold {
		t.Errorf("Font.Bold = %v, want %v", result.Font.Bold, original.Font.Bold)
	}
	if result.Font.Underline != original.Font.Underline {
		t.Errorf("Font.Underline = %v, want %v", result.Font.Underline, original.Font.Underline)
	}
	if result.Font.Color != original.Font.Color {
		t.Errorf("Font.Color = %q, want %q", result.Font.Color, original.Font.Color)
	}
	if result.Fill.Color != original.Fill.Color {
		t.Errorf("Fill.Color = %q, want %q", result.Fill.Color, original.Fill.Color)
	}
	if result.Border.Left.Style != original.Border.Left.Style {
		t.Errorf("Border.Left.Style = %v, want %v", result.Border.Left.Style, original.Border.Left.Style)
	}
	if result.Border.Bottom.Style != original.Border.Bottom.Style {
		t.Errorf("Border.Bottom.Style = %v, want %v", result.Border.Bottom.Style, original.Border.Bottom.Style)
	}
	if result.Alignment.Horizontal != original.Alignment.Horizontal {
		t.Errorf("Alignment.Horizontal = %v, want %v", result.Alignment.Horizontal, original.Alignment.Horizontal)
	}
	if result.Alignment.WrapText != original.Alignment.WrapText {
		t.Errorf("Alignment.WrapText = %v, want %v", result.Alignment.WrapText, original.Alignment.WrapText)
	}
	if result.NumFmt != original.NumFmt {
		t.Errorf("NumFmt = %q, want %q", result.NumFmt, original.NumFmt)
	}
}

func TestSetGetStyle(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	style := &Style{
		Font: &Font{Bold: true, Size: 16, Name: "Arial"},
		Fill: &Fill{Color: "FFFF00"},
	}
	if err := s.SetStyle("A1", style); err != nil {
		t.Fatal(err)
	}

	got, err := s.GetStyle("A1")
	if err != nil {
		t.Fatal(err)
	}
	if got != style {
		t.Error("GetStyle did not return the same pointer")
	}

	// Nonexistent cell returns nil.
	got2, err := s.GetStyle("Z99")
	if err != nil {
		t.Fatal(err)
	}
	if got2 != nil {
		t.Errorf("expected nil for nonexistent cell, got %+v", got2)
	}
}

func TestSetStyle_InvalidRef(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	err := s.SetStyle("!!!", &Style{})
	if err == nil {
		t.Error("expected error for invalid cell ref")
	}
}

func TestGetStyle_InvalidRef(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_, err := s.GetStyle("!!!")
	if err == nil {
		t.Error("expected error for invalid cell ref")
	}
}

func TestCellStyleAccessor(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	// Cell with no style.
	if err := s.SetValue("A1", "hello"); err != nil {
		t.Fatal(err)
	}
	for row := range s.Rows() {
		for _, c := range row.Cells() {
			if c.Style() != nil {
				t.Error("expected nil style for unstyled cell")
			}
		}
	}

	// Cell with style.
	style := &Style{Font: &Font{Bold: true}}
	if err := s.SetStyle("A1", style); err != nil {
		t.Fatal(err)
	}
	for row := range s.Rows() {
		for _, c := range row.Cells() {
			if c.Style() != style {
				t.Error("Style() did not return expected pointer")
			}
		}
	}
}

func TestSetRangeStyle(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")

	style := &Style{Font: &Font{Bold: true}}
	if err := s.SetRangeStyle("A1:C3", style); err != nil {
		t.Fatal(err)
	}

	// Verify all cells in the range have the style.
	for _, ref := range []string{"A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"} {
		got, err := s.GetStyle(ref)
		if err != nil {
			t.Fatalf("GetStyle(%s): %v", ref, err)
		}
		if got != style {
			t.Errorf("GetStyle(%s) did not return expected style", ref)
		}
	}
}

func TestSetRangeStyle_SingleCell(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	style := &Style{Fill: &Fill{Color: "FF0000"}}
	if err := s.SetRangeStyle("B2", style); err != nil {
		t.Fatal(err)
	}
	got, _ := s.GetStyle("B2")
	if got != style {
		t.Error("single-cell range did not apply style")
	}
}

func TestSetRangeStyle_Reversed(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	style := &Style{Font: &Font{Italic: true}}
	if err := s.SetRangeStyle("C3:A1", style); err != nil {
		t.Fatal(err)
	}
	got, _ := s.GetStyle("A1")
	if got != style {
		t.Error("reversed range did not apply style to A1")
	}
	got2, _ := s.GetStyle("C3")
	if got2 != style {
		t.Error("reversed range did not apply style to C3")
	}
}

func TestSetRangeStyle_Invalid(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	err := s.SetRangeStyle("!!!:B2", &Style{})
	if err == nil {
		t.Error("expected error for invalid range")
	}
}

func TestSetRangeStyle_CreatesEmptyCells(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	style := &Style{Font: &Font{Bold: true}}
	_ = s.SetRangeStyle("A1:B2", style)

	// All cells should exist even though no values were set.
	for _, ref := range []string{"A1", "A2", "B1", "B2"} {
		got, _ := s.GetStyle(ref)
		if got == nil {
			t.Errorf("cell %s should have been created with style", ref)
		}
	}
}

func TestRGBConversions(t *testing.T) {
	if got := rgbToARGB("FF0000"); got != "FFFF0000" {
		t.Errorf("rgbToARGB = %q, want FFFF0000", got)
	}
	if got := rgbToARGB(""); got != "" {
		t.Errorf("rgbToARGB empty = %q, want empty", got)
	}
	if got := argbToRGB("FFFF0000"); got != "FF0000" {
		t.Errorf("argbToRGB = %q, want FF0000", got)
	}
	if got := argbToRGB("FF00"); got != "FF00" {
		t.Errorf("argbToRGB short = %q, want FF00 (passthrough)", got)
	}
}
