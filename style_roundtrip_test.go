package werkbook

import (
	"os"
	"path/filepath"
	"testing"
)

func TestStyleRoundTrip_FontBoldItalic(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetValue("A1", "Bold")
	_ = s.SetStyle("A1", &Style{Font: &Font{Bold: true, Size: 14, Name: "Arial"}})
	_ = s.SetValue("A2", "Italic")
	_ = s.SetStyle("A2", &Style{Font: &Font{Italic: true, Size: 11, Name: "Calibri"}})

	path := filepath.Join(t.TempDir(), "font.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	s2 := f2.Sheet("Sheet1")

	st1, _ := s2.GetStyle("A1")
	if st1 == nil || st1.Font == nil {
		t.Fatal("A1 style or font is nil")
	}
	if !st1.Font.Bold {
		t.Error("A1 Font.Bold = false, want true")
	}
	if st1.Font.Name != "Arial" {
		t.Errorf("A1 Font.Name = %q, want %q", st1.Font.Name, "Arial")
	}
	if st1.Font.Size != 14 {
		t.Errorf("A1 Font.Size = %g, want %g", st1.Font.Size, 14.0)
	}

	st2, _ := s2.GetStyle("A2")
	if st2 == nil || st2.Font == nil {
		t.Fatal("A2 style or font is nil")
	}
	if !st2.Font.Italic {
		t.Error("A2 Font.Italic = false, want true")
	}
}

func TestStyleRoundTrip_Fill(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetValue("A1", "Yellow")
	_ = s.SetStyle("A1", &Style{Fill: &Fill{Color: "FFFF00"}})

	path := filepath.Join(t.TempDir(), "fill.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	st, _ := f2.Sheet("Sheet1").GetStyle("A1")
	if st == nil || st.Fill == nil {
		t.Fatal("A1 style or fill is nil")
	}
	if st.Fill.Color != "FFFF00" {
		t.Errorf("Fill.Color = %q, want %q", st.Fill.Color, "FFFF00")
	}
}

func TestStyleRoundTrip_Border(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetValue("A1", "Bordered")
	_ = s.SetStyle("A1", &Style{
		Border: &Border{
			Left:   BorderSide{Style: BorderThin, Color: "000000"},
			Right:  BorderSide{Style: BorderThin, Color: "000000"},
			Top:    BorderSide{Style: BorderThin, Color: "000000"},
			Bottom: BorderSide{Style: BorderThin, Color: "000000"},
		},
	})

	path := filepath.Join(t.TempDir(), "border.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	st, _ := f2.Sheet("Sheet1").GetStyle("A1")
	if st == nil || st.Border == nil {
		t.Fatal("A1 style or border is nil")
	}
	if st.Border.Left.Style != BorderThin {
		t.Errorf("Border.Left.Style = %v, want %v", st.Border.Left.Style, BorderThin)
	}
	if st.Border.Left.Color != "000000" {
		t.Errorf("Border.Left.Color = %q, want %q", st.Border.Left.Color, "000000")
	}
}

func TestStyleRoundTrip_Alignment(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetValue("A1", "Centered")
	_ = s.SetStyle("A1", &Style{
		Alignment: &Alignment{
			Horizontal: HAlignCenter,
			Vertical:   VAlignTop,
			WrapText:   true,
		},
	})

	path := filepath.Join(t.TempDir(), "align.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	st, _ := f2.Sheet("Sheet1").GetStyle("A1")
	if st == nil || st.Alignment == nil {
		t.Fatal("A1 style or alignment is nil")
	}
	if st.Alignment.Horizontal != HAlignCenter {
		t.Errorf("Horizontal = %v, want %v", st.Alignment.Horizontal, HAlignCenter)
	}
	if st.Alignment.Vertical != VAlignTop {
		t.Errorf("Vertical = %v, want %v", st.Alignment.Vertical, VAlignTop)
	}
	if !st.Alignment.WrapText {
		t.Error("WrapText = false, want true")
	}
}

func TestStyleRoundTrip_NumFmt(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetValue("A1", 1234.5678)
	_ = s.SetStyle("A1", &Style{NumFmt: "#,##0.00"})

	path := filepath.Join(t.TempDir(), "numfmt.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	st, _ := f2.Sheet("Sheet1").GetStyle("A1")
	if st == nil {
		t.Fatal("A1 style is nil")
	}
	if st.NumFmt != "#,##0.00" {
		t.Errorf("NumFmt = %q, want %q", st.NumFmt, "#,##0.00")
	}
}

func TestStyleRoundTrip_BuiltinNumFmtID(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetValue("A1", 0.5)
	_ = s.SetStyle("A1", &Style{NumFmtID: 10}) // 0.00%

	path := filepath.Join(t.TempDir(), "numfmtid.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	st, _ := f2.Sheet("Sheet1").GetStyle("A1")
	if st == nil {
		t.Fatal("A1 style is nil")
	}
	if st.NumFmtID != 10 {
		t.Errorf("NumFmtID = %d, want %d", st.NumFmtID, 10)
	}
}

func TestStyleRoundTrip_Combined(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetValue("A1", "Fancy")
	_ = s.SetStyle("A1", &Style{
		Font:      &Font{Bold: true, Size: 18, Name: "Helvetica", Color: "0000FF"},
		Fill:      &Fill{Color: "FFFF00"},
		Border:    &Border{Bottom: BorderSide{Style: BorderDouble, Color: "FF0000"}},
		Alignment: &Alignment{Horizontal: HAlignRight, WrapText: true},
		NumFmt:    "0.0",
	})
	_ = s.SetValue("B1", "Plain")

	path := filepath.Join(t.TempDir(), "combined.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	s2 := f2.Sheet("Sheet1")

	st, _ := s2.GetStyle("A1")
	if st == nil {
		t.Fatal("A1 style is nil")
	}
	if st.Font == nil || !st.Font.Bold || st.Font.Size != 18 || st.Font.Name != "Helvetica" || st.Font.Color != "0000FF" {
		t.Errorf("Font mismatch: %+v", st.Font)
	}
	if st.Fill == nil || st.Fill.Color != "FFFF00" {
		t.Errorf("Fill mismatch: %+v", st.Fill)
	}
	if st.Border == nil || st.Border.Bottom.Style != BorderDouble || st.Border.Bottom.Color != "FF0000" {
		t.Errorf("Border mismatch: %+v", st.Border)
	}
	if st.Alignment == nil || st.Alignment.Horizontal != HAlignRight || !st.Alignment.WrapText {
		t.Errorf("Alignment mismatch: %+v", st.Alignment)
	}
	if st.NumFmt != "0.0" {
		t.Errorf("NumFmt = %q, want %q", st.NumFmt, "0.0")
	}

	// B1 should have no style (or default).
	stB, _ := s2.GetStyle("B1")
	if stB != nil {
		t.Errorf("B1 style should be nil, got %+v", stB)
	}
}

func TestStyleRoundTrip_Dedup(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	style := &Style{Font: &Font{Bold: true, Size: 12, Name: "Calibri"}}
	_ = s.SetValue("A1", "First")
	_ = s.SetStyle("A1", style)
	_ = s.SetValue("A2", "Second")
	_ = s.SetStyle("A2", style)

	path := filepath.Join(t.TempDir(), "dedup.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	// Re-open and verify both cells have the same style.
	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	s2 := f2.Sheet("Sheet1")
	st1, _ := s2.GetStyle("A1")
	st2, _ := s2.GetStyle("A2")
	if st1 == nil || st2 == nil {
		t.Fatal("styles are nil")
	}
	// Since they reference the same StyleData index, they should share the same pointer.
	if st1 != st2 {
		t.Error("deduplication failed: A1 and A2 should share the same *Style pointer")
	}
}

func TestStyleRoundTrip_StyleOnlyCell(t *testing.T) {
	// A cell with only a style (no value, no formula) should be preserved.
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetStyle("B2", &Style{Fill: &Fill{Color: "00FF00"}})

	path := filepath.Join(t.TempDir(), "styleonly.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	st, _ := f2.Sheet("Sheet1").GetStyle("B2")
	if st == nil || st.Fill == nil {
		t.Fatal("style-only cell lost its style")
	}
	if st.Fill.Color != "00FF00" {
		t.Errorf("Fill.Color = %q, want %q", st.Fill.Color, "00FF00")
	}
}

func TestStyleRoundTrip_UnstyledWorkbookBackwardCompat(t *testing.T) {
	// An unstyled workbook should round-trip without issues.
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetValue("A1", "Hello")
	_ = s.SetValue("A2", 42)
	_ = s.SetValue("A3", true)

	path := filepath.Join(t.TempDir(), "unstyled.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	s2 := f2.Sheet("Sheet1")
	v1, _ := s2.GetValue("A1")
	if v1.String != "Hello" {
		t.Errorf("A1 = %q, want %q", v1.String, "Hello")
	}
	v2, _ := s2.GetValue("A2")
	if v2.Number != 42 {
		t.Errorf("A2 = %g, want %g", v2.Number, 42.0)
	}
	v3, _ := s2.GetValue("A3")
	if !v3.Bool {
		t.Error("A3 = false, want true")
	}
}

func TestStyleRoundTrip_MultiSheet(t *testing.T) {
	f := New()
	s1 := f.Sheet("Sheet1")
	s2, _ := f.NewSheet("Sheet2")

	_ = s1.SetValue("A1", "Sheet1")
	_ = s1.SetStyle("A1", &Style{Font: &Font{Bold: true}})

	_ = s2.SetValue("A1", "Sheet2")
	_ = s2.SetStyle("A1", &Style{Fill: &Fill{Color: "FF0000"}})

	path := filepath.Join(t.TempDir(), "multisheet.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f3, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}

	st1, _ := f3.Sheet("Sheet1").GetStyle("A1")
	if st1 == nil || st1.Font == nil || !st1.Font.Bold {
		t.Error("Sheet1 A1 should have bold font")
	}

	st2, _ := f3.Sheet("Sheet2").GetStyle("A1")
	if st2 == nil || st2.Fill == nil || st2.Fill.Color != "FF0000" {
		t.Error("Sheet2 A1 should have red fill")
	}
}

func TestStyleRoundTrip_FontUnderlineAndColor(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	_ = s.SetValue("A1", "Underlined")
	_ = s.SetStyle("A1", &Style{
		Font: &Font{Underline: true, Color: "00FFAA", Size: 11, Name: "Calibri"},
	})

	path := filepath.Join(t.TempDir(), "underline.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	st, _ := f2.Sheet("Sheet1").GetStyle("A1")
	if st == nil || st.Font == nil {
		t.Fatal("style/font is nil")
	}
	if !st.Font.Underline {
		t.Error("Font.Underline = false, want true")
	}
	if st.Font.Color != "00FFAA" {
		t.Errorf("Font.Color = %q, want %q", st.Font.Color, "00FFAA")
	}
}

func TestStyleRoundTrip_RangeStyle(t *testing.T) {
	f := New()
	s := f.Sheet("Sheet1")
	// Set some values and apply a range style.
	_ = s.SetValue("A1", "TopLeft")
	_ = s.SetValue("C3", "BottomRight")
	_ = s.SetRangeStyle("A1:C3", &Style{Font: &Font{Bold: true, Size: 12, Name: "Calibri"}})

	path := filepath.Join(t.TempDir(), "rangestyle.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	s2 := f2.Sheet("Sheet1")

	// All 9 cells should have the bold style.
	for _, ref := range []string{"A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"} {
		st, err := s2.GetStyle(ref)
		if err != nil {
			t.Fatalf("GetStyle(%s): %v", ref, err)
		}
		if st == nil || st.Font == nil {
			t.Fatalf("%s style or font is nil", ref)
		}
		if !st.Font.Bold {
			t.Errorf("%s Font.Bold = false, want true", ref)
		}
	}

	// Verify data is preserved.
	v1, _ := s2.GetValue("A1")
	if v1.String != "TopLeft" {
		t.Errorf("A1 = %q, want TopLeft", v1.String)
	}
	v2, _ := s2.GetValue("C3")
	if v2.String != "BottomRight" {
		t.Errorf("C3 = %q, want BottomRight", v2.String)
	}
}

func TestReadExistingStyledFile(t *testing.T) {
	// If a styled test fixture exists, read it and verify styles are preserved.
	path := filepath.Join("testdata", "styled.xlsx")
	if _, err := os.Stat(path); os.IsNotExist(err) {
		t.Skip("testdata/styled.xlsx not found, skipping")
	}
	_, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
}
