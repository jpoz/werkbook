package formula

import "testing"

func TestInvalidateTableColumnRefs(t *testing.T) {
	tests := []struct {
		name    string
		src     string
		columns []string
		inTable bool
		want    string
	}{
		{"one column", "SUM(Tbl[Qty])", []string{"Qty"}, false, "SUM(#REF!)"},
		{"a column that stays", "SUM(Tbl[Price])", []string{"Qty"}, false, "SUM(Tbl[Price])"},
		{"no letter case for the table and the column", "SUM(tbl[QTY])", []string{"Qty"}, false, "SUM(#REF!)"},
		{"this row", "Tbl[[#This Row],[Qty]]*2", []string{"Qty"}, false, "#REF!*2"},
		{"this row, short form", "Tbl[@Qty]*2", []string{"Qty"}, false, "#REF!*2"},
		{"this row, short form with brackets", "Tbl[@[Qty]]*2", []string{"Qty"}, false, "#REF!*2"},
		{"a column range with one removed end", "SUM(Tbl[[Qty]:[Price]])", []string{"Price"}, false, "SUM(#REF!)"},
		{"a column range that stays", "SUM(Tbl[[Qty]:[Price]])", []string{"Tax"}, false, "SUM(Tbl[[Qty]:[Price]])"},
		{"special items are not columns", "ROWS(Tbl[#All])+ROWS(Tbl[[#Headers],[Price]])", []string{"All", "Headers", "#All"}, false, "ROWS(Tbl[#All])+ROWS(Tbl[[#Headers],[Price]])"},
		{"two removed columns", "Tbl[A]+Tbl[B]+Tbl[C]", []string{"A", "C"}, false, "#REF!+Tbl[B]+#REF!"},
		{"a different table", "SUM(Other[Qty])+SUM(Tbl2[Qty])+SUM(MyTbl[Qty])", []string{"Qty"}, false, "SUM(Other[Qty])+SUM(Tbl2[Qty])+SUM(MyTbl[Qty])"},
		{"no table name, formula in the table", "[@Qty]*[@Price]", []string{"Qty"}, true, "#REF!*[@Price]"},
		{"no table name, formula in the table, plain form", "SUM([Qty])", []string{"Qty"}, true, "SUM(#REF!)"},
		{"no table name, formula out of the table", "[@Qty]*2", []string{"Qty"}, false, "[@Qty]*2"},
		{"text in a string", `"Tbl[Qty]"&Tbl[Qty]`, []string{"Qty"}, false, `"Tbl[Qty]"&#REF!`},
		{"column name with a space", "SUM(Tbl[Unit Price])", []string{"Unit Price"}, false, "SUM(#REF!)"},
		{"column name with an escaped bracket", "SUM(Tbl[a'[b'] c])", []string{"a[b] c"}, false, "SUM(#REF!)"},
		{"other text stays", " sum( Tbl[Qty] ) + a1 ", []string{"Qty"}, false, " sum( #REF! ) + a1 "},
		{"external workbook prefix", "[Book.xlsx]Sheet1!A1+Tbl[Qty]", []string{"Qty", "Book.xlsx"}, false, "[Book.xlsx]Sheet1!A1+#REF!"},
		{"quoted sheet name with a bracket", "'[Qty]Sheet'!A1+Tbl[Qty]", []string{"Qty"}, true, "'[Qty]Sheet'!A1+#REF!"},
		{"a group with no end", "SUM(Tbl[Qty", []string{"Qty"}, false, "SUM(Tbl[Qty"},
		{"no columns", "SUM(Tbl[Qty])", nil, false, "SUM(Tbl[Qty])"},
		{"no brackets", "SUM(A1:B2)", []string{"Qty"}, true, "SUM(A1:B2)"},
		{"empty", "", []string{"Qty"}, true, ""},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, changed := InvalidateTableColumnRefs(tt.src, "Tbl", tt.columns, tt.inTable)
			if got != tt.want {
				t.Errorf("InvalidateTableColumnRefs(%q) = %q, want %q", tt.src, got, tt.want)
			}
			if changed != (got != tt.src) {
				t.Errorf("changed = %v for %q -> %q", changed, tt.src, got)
			}
		})
	}
	// It never fails on text that is not a formula.
	for _, src := range append([]string{"[", "]", "[[", "]]", "'[", "\"[", "Tbl[[[", "Tbl[]", "[]"}, realFormulas...) {
		_, _ = InvalidateTableColumnRefs(src, "Tbl", []string{"Qty", "Amount", "Col"}, true)
	}
}
