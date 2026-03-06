package werkbook

import (
	"errors"
	"slices"
	"testing"
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
