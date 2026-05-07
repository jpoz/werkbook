package werkbook

import (
	"path/filepath"
	"testing"
)

// --- Unit tests for rewriteSheetRefsInFormula ---

func TestRewriteSheetRefsInFormula_QuotedRef(t *testing.T) {
	got := rewriteSheetRefsInFormula("'Out - X'!A1+'Out - X'!B2", "Out - X", "X")
	want := "X!A1+X!B2"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_QuotedToQuoted(t *testing.T) {
	got := rewriteSheetRefsInFormula("'Out - X'!A1", "Out - X", "New Sheet")
	want := "'New Sheet'!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_UnquotedRef(t *testing.T) {
	got := rewriteSheetRefsInFormula("Sheet1!A1+Sheet1!B2", "Sheet1", "Data")
	want := "Data!A1+Data!B2"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_UnquotedToQuoted(t *testing.T) {
	got := rewriteSheetRefsInFormula("Sheet1!A1", "Sheet1", "My Sheet")
	want := "'My Sheet'!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_ApostropheInName(t *testing.T) {
	got := rewriteSheetRefsInFormula("'Fund''s Data'!A1", "Fund's Data", "Renamed")
	want := "Renamed!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_ApostropheInNewName(t *testing.T) {
	got := rewriteSheetRefsInFormula("Sheet1!A1", "Sheet1", "Fund's Data")
	want := "'Fund''s Data'!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_StringLiteralPreserved(t *testing.T) {
	src := `="'Out - X' report"&'Out - X'!A1`
	got := rewriteSheetRefsInFormula(src, "Out - X", "X")
	want := `="'Out - X' report"&X!A1`
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_NoMatchUntouched(t *testing.T) {
	src := "'Other Sheet'!A1+Sheet2!B2"
	got := rewriteSheetRefsInFormula(src, "Out - X", "X")
	if got != src {
		t.Fatalf("got %q, want %q (unchanged)", got, src)
	}
}

func TestRewriteSheetRefsInFormula_EmptyFormula(t *testing.T) {
	got := rewriteSheetRefsInFormula("", "Sheet1", "Data")
	if got != "" {
		t.Fatalf("got %q, want empty", got)
	}
}

func TestRewriteSheetRefsInFormula_SameNameNoop(t *testing.T) {
	src := "Sheet1!A1"
	got := rewriteSheetRefsInFormula(src, "Sheet1", "Sheet1")
	if got != src {
		t.Fatalf("got %q, want %q", got, src)
	}
}

func TestRewriteSheetRefsInFormula_CaseInsensitiveUnquoted(t *testing.T) {
	got := rewriteSheetRefsInFormula("sheet1!A1", "Sheet1", "Data")
	want := "Data!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_MultipleSheetRefs(t *testing.T) {
	src := "SUM('Out - X'!A1:A10)+'Out - X'!B1+Sheet2!C1"
	got := rewriteSheetRefsInFormula(src, "Out - X", "Results")
	want := "SUM(Results!A1:A10)+Results!B1+Sheet2!C1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_AbsoluteRefs(t *testing.T) {
	got := rewriteSheetRefsInFormula("'Out - X'!$A$1:$C$100", "Out - X", "X")
	want := "X!$A$1:$C$100"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_NoFalsePositiveInIdentifier(t *testing.T) {
	src := "Sheet1Summary+Sheet1!A1"
	got := rewriteSheetRefsInFormula(src, "Sheet1", "Data")
	want := "Sheet1Summary+Data!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_QuotedRefNotFollowedByBang(t *testing.T) {
	src := "'Sheet1'+A1"
	got := rewriteSheetRefsInFormula(src, "Sheet1", "Data")
	if got != src {
		t.Fatalf("got %q, want %q (unchanged)", got, src)
	}
}

func TestRewriteSheetRefsInFormula_CaseInsensitiveQuoted(t *testing.T) {
	got := rewriteSheetRefsInFormula("'SHEET ONE'!A1", "Sheet One", "Data")
	want := "Data!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_Unquoted3D(t *testing.T) {
	got := rewriteSheetRefsInFormula("Sheet1:Sheet3!A1", "Sheet1", "Data")
	want := "Data:Sheet3!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_Unquoted3DBothEndpoints(t *testing.T) {
	got := rewriteSheetRefsInFormula("Sheet1:Sheet1!A1", "Sheet1", "Data")
	want := "Data:Data!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_Unquoted3DSecondOnly(t *testing.T) {
	got := rewriteSheetRefsInFormula("Other:Sheet1!A1", "Sheet1", "Data")
	want := "Other:Data!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_Unquoted3DSecondOnlyNeedsQuoting(t *testing.T) {
	got := rewriteSheetRefsInFormula("Other:Sheet1!A1", "Sheet1", "My Sheet")
	want := "'Other:My Sheet'!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_Unquoted3DNewNameNeedsQuoting(t *testing.T) {
	got := rewriteSheetRefsInFormula("Sheet1:Sheet3!A1", "Sheet1", "My Sheet")
	want := "'My Sheet:Sheet3'!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_Quoted3D(t *testing.T) {
	got := rewriteSheetRefsInFormula("'Out - X:Out - Y'!A1", "Out - X", "X")
	want := "'X:Out - Y'!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_Quoted3DBothEndpoints(t *testing.T) {
	got := rewriteSheetRefsInFormula("'Out - X:Out - X'!A1", "Out - X", "X")
	want := "X:X!A1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_Unquoted3DNoMatch(t *testing.T) {
	src := "Other:Another!A1"
	got := rewriteSheetRefsInFormula(src, "Sheet1", "Data")
	if got != src {
		t.Fatalf("got %q, want %q (unchanged)", got, src)
	}
}

func TestRewriteSheetRefsInFormula_RangeWithRepeatedSheetRef(t *testing.T) {
	got := rewriteSheetRefsInFormula("Sheet1!A1:Sheet1!B1", "Sheet1", "My Sheet")
	want := "'My Sheet'!A1:'My Sheet'!B1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_RangeWithRepeatedSheetRefNoQuoting(t *testing.T) {
	got := rewriteSheetRefsInFormula("Sheet1!A1:Sheet1!B1", "Sheet1", "Data")
	want := "Data!A1:Data!B1"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func TestRewriteSheetRefsInFormula_SumRangeAcrossSheetRef(t *testing.T) {
	got := rewriteSheetRefsInFormula("SUM(Sheet1!A1:Sheet1!A10)", "Sheet1", "My Sheet")
	want := "SUM('My Sheet'!A1:'My Sheet'!A10)"
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// --- Integration tests for SetSheetName formula rewriting ---

func TestSetSheetName_RewritesCrossSheetFormula(t *testing.T) {
	f := New()
	s1 := f.Sheet("Sheet1")
	s1.SetValue("A1", 10)

	s2, _ := f.NewSheet("Sheet2")
	s2.SetFormula("A1", "Sheet1!A1*2")

	if err := f.SetSheetName("Sheet1", "Data"); err != nil {
		t.Fatal(err)
	}

	got, _ := f.Sheet("Sheet2").GetFormula("A1")
	if got != "Data!A1*2" {
		t.Fatalf("formula = %q, want %q", got, "Data!A1*2")
	}

	v, err := f.Sheet("Sheet2").GetValue("A1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeNumber || v.Number != 20 {
		t.Fatalf("Sheet2!A1 = %v, want 20", v)
	}
}

func TestSetSheetName_RewritesQuotedCrossSheetFormula(t *testing.T) {
	f := New(FirstSheet("Out - X"))
	s1 := f.Sheet("Out - X")
	s1.SetValue("A1", 5)

	s2, _ := f.NewSheet("Sheet2")
	s2.SetFormula("A1", "'Out - X'!A1+10")

	if err := f.SetSheetName("Out - X", "X"); err != nil {
		t.Fatal(err)
	}

	got, _ := f.Sheet("Sheet2").GetFormula("A1")
	if got != "X!A1+10" {
		t.Fatalf("formula = %q, want %q", got, "X!A1+10")
	}

	v, err := f.Sheet("Sheet2").GetValue("A1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeNumber || v.Number != 15 {
		t.Fatalf("Sheet2!A1 = %v, want 15", v)
	}
}

func TestSetSheetName_RewritesDefinedNameValue(t *testing.T) {
	f := New(FirstSheet("Out - X"))
	f.Sheet("Out - X").SetValue("A1", 42)

	f.AddDefinedName(DefinedName{
		Name:         "MyRange",
		Value:        "'Out - X'!$A$1:$C$100",
		LocalSheetID: -1,
	})

	if err := f.SetSheetName("Out - X", "X"); err != nil {
		t.Fatal(err)
	}

	names := f.DefinedNames()
	if len(names) != 1 {
		t.Fatalf("expected 1 defined name, got %d", len(names))
	}
	want := "X!$A$1:$C$100"
	if names[0].Value != want {
		t.Fatalf("defined name Value = %q, want %q", names[0].Value, want)
	}
}

func TestSetSheetName_StringLiteralsNotCorrupted(t *testing.T) {
	f := New(FirstSheet("Out - X"))
	f.Sheet("Out - X").SetValue("A1", 1)

	s2, _ := f.NewSheet("Sheet2")
	s2.SetFormula("A1", `"'Out - X' report"&'Out - X'!A1`)

	if err := f.SetSheetName("Out - X", "X"); err != nil {
		t.Fatal(err)
	}

	got, _ := f.Sheet("Sheet2").GetFormula("A1")
	want := `"'Out - X' report"&X!A1`
	if got != want {
		t.Fatalf("formula = %q, want %q", got, want)
	}
}

func TestSetSheetName_ApostropheInSheetName(t *testing.T) {
	f := New(FirstSheet("Fund's Data"))
	s1 := f.Sheet("Fund's Data")
	s1.SetValue("A1", 100)

	s2, _ := f.NewSheet("Summary")
	s2.SetFormula("A1", "'Fund''s Data'!A1+1")

	if err := f.SetSheetName("Fund's Data", "Data"); err != nil {
		t.Fatal(err)
	}

	got, _ := f.Sheet("Summary").GetFormula("A1")
	if got != "Data!A1+1" {
		t.Fatalf("formula = %q, want %q", got, "Data!A1+1")
	}

	v, err := f.Sheet("Summary").GetValue("A1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeNumber || v.Number != 101 {
		t.Fatalf("Summary!A1 = %v, want 101", v)
	}
}

func TestSetSheetName_RoundTrip(t *testing.T) {
	f := New(FirstSheet("Out - X"))
	s1 := f.Sheet("Out - X")
	s1.SetValue("A1", 7)

	s2, _ := f.NewSheet("Calc")
	s2.SetFormula("A1", "'Out - X'!A1*3")

	f.AddDefinedName(DefinedName{
		Name:         "Result",
		Value:        "'Out - X'!$A$1",
		LocalSheetID: -1,
	})

	if err := f.SetSheetName("Out - X", "X"); err != nil {
		t.Fatal(err)
	}

	path := filepath.Join(t.TempDir(), "renamed.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}

	got, _ := f2.Sheet("Calc").GetFormula("A1")
	if got != "X!A1*3" {
		t.Fatalf("round-trip formula = %q, want %q", got, "X!A1*3")
	}

	v, err := f2.Sheet("Calc").GetValue("A1")
	if err != nil {
		t.Fatal(err)
	}
	if v.Type != TypeNumber || v.Number != 21 {
		t.Fatalf("Calc!A1 = %v, want 21", v)
	}

	names := f2.DefinedNames()
	if len(names) != 1 {
		t.Fatalf("expected 1 defined name, got %d", len(names))
	}
	if names[0].Value != "X!$A$1" {
		t.Fatalf("defined name Value = %q, want %q", names[0].Value, "X!$A$1")
	}
}

func TestSetSheetName_UnquotedRef(t *testing.T) {
	f := New()
	s1 := f.Sheet("Sheet1")
	s1.SetValue("A1", 3)

	s2, _ := f.NewSheet("Sheet2")
	s2.SetFormula("A1", "Sheet1!A1+Sheet1!B1")

	if err := f.SetSheetName("Sheet1", "Source"); err != nil {
		t.Fatal(err)
	}

	got, _ := f.Sheet("Sheet2").GetFormula("A1")
	if got != "Source!A1+Source!B1" {
		t.Fatalf("formula = %q, want %q", got, "Source!A1+Source!B1")
	}
}
