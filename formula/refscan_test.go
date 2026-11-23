package formula

import (
	"fmt"
	"strings"
	"testing"
)

var refKindNames = map[RefKind]string{RefCell: "cell", RefRange: "range", RefCols: "cols", RefRows: "rows"}

// describeSite gives "text|sheet|sheetEnd|kind|body" for one site. text is
// the source span; body is the canonical form without the qualifier, which
// shows the parsed corners and absolute flags.
func describeSite(src string, s RefSite) string {
	body := s
	body.Sheet, body.SheetEnd, body.qual = "", "", ""
	return fmt.Sprintf("%s|%s|%s|%s|%s", src[s.Start:s.End], s.Sheet, s.SheetEnd, refKindNames[s.Kind], FormatRef(body))
}

func describeSites(src string) []string {
	var out []string
	for _, s := range ScanRefs(src) {
		out = append(out, describeSite(src, s))
	}
	return out
}

func TestScanRefs(t *testing.T) {
	tests := []struct {
		name string
		src  string
		want []string
	}{
		// Site forms.
		{"cell", "A1", []string{"A1|||cell|A1"}},
		{"absolute cell", "$A$1", []string{"$A$1|||cell|$A$1"}},
		{"mixed absolute", "A$1+$B2", []string{"A$1|||cell|A$1", "$B2|||cell|$B2"}},
		{"range", "SUM(A1:B2)", []string{"A1:B2|||range|A1:B2"}},
		{"range with absolute parts", "$A$1:B$2", []string{"$A$1:B$2|||range|$A$1:B$2"}},
		{"reversed range keeps its corner order", "B2:A1", []string{"B2:A1|||range|B2:A1"}},
		{"columns", "SUM(A:C)", []string{"A:C|||cols|A:C"}},
		{"absolute columns", "$A:$A", []string{"$A:$A|||cols|$A:$A"}},
		{"rows", "SUM(1:3)", []string{"1:3|||rows|1:3"}},
		{"absolute rows", "SUM($1:$3)", []string{"$1:$3|||rows|$1:$3"}},
		{"mixed absolute rows", "1:$3+$4:5", []string{"1:$3|||rows|1:$3", "$4:5|||rows|$4:5"}},
		{"last cell", "XFD1048576", []string{"XFD1048576|||cell|XFD1048576"}},
		{"lower case", "a1+b2:c3", []string{"a1|||cell|A1", "b2:c3|||range|B2:C3"}},

		// Sheet qualifiers.
		{"sheet cell", "Sheet1!A1", []string{"Sheet1!A1|Sheet1||cell|A1"}},
		{"sheet range: the qualifier covers the full range", "Sheet1!A1:B2", []string{"Sheet1!A1:B2|Sheet1||range|A1:B2"}},
		{"sheet columns", "Sheet1!$A:$B", []string{"Sheet1!$A:$B|Sheet1||cols|$A:$B"}},
		{"sheet rows", "Sheet1!1:3", []string{"Sheet1!1:3|Sheet1||rows|1:3"}},
		{"sheet absolute rows", "Sheet1!$1:$2", []string{"Sheet1!$1:$2|Sheet1||rows|$1:$2"}},
		{"quoted sheet", "'My Sheet'!A1", []string{"'My Sheet'!A1|My Sheet||cell|A1"}},
		{"quoted sheet range", "SUM('My Sheet'!A1:B2)", []string{"'My Sheet'!A1:B2|My Sheet||range|A1:B2"}},
		{"quoted sheet with a quote", "'It''s'!$B$2", []string{"'It''s'!$B$2|It's||cell|$B$2"}},
		{"quoted sheet rows", "'My Sheet'!$1:$2", []string{"'My Sheet'!$1:$2|My Sheet||rows|$1:$2"}},
		{"quotes that are not necessary stay in the span", "'Sheet1'!A1", []string{"'Sheet1'!A1|Sheet1||cell|A1"}},
		{"sheet name with underscore and dot", "My_Sheet.2!C3", []string{"My_Sheet.2!C3|My_Sheet.2||cell|C3"}},
		{"sheet name that is not ASCII", "Größe!A1+Übersicht!B2", []string{"Größe!A1|Größe||cell|A1", "Übersicht!B2|Übersicht||cell|B2"}},
		{"3D", "SUM(Sheet1:Sheet3!A1)", []string{"Sheet1:Sheet3!A1|Sheet1|Sheet3|cell|A1"}},
		{"3D range", "SUM(Sheet1:Sheet3!A1:B2)", []string{"Sheet1:Sheet3!A1:B2|Sheet1|Sheet3|range|A1:B2"}},
		{"quoted 3D", "SUM('Jan 1:Dec 31'!C5)", []string{"'Jan 1:Dec 31'!C5|Jan 1|Dec 31|cell|C5"}},
		{"sheet names that look like columns make a 3D reference", "A:B!C1", []string{"A:B!C1|A|B|cell|C1"}},
		{"cell-like first word makes the colon a range operator", "S1:S3", []string{"S1:S3|||range|S1:S3"}},

		// Several references.
		{"many", "A1+B2*Sheet2!C3-SUM(D:D)", []string{"A1|||cell|A1", "B2|||cell|B2", "Sheet2!C3|Sheet2||cell|C3", "D:D|||cols|D:D"}},
		{"spaces", " A1 +  B2 ", []string{"A1|||cell|A1", "B2|||cell|B2"}},
		{"intersection with a space", "A1:B5 B2:C6", []string{"A1:B5|||range|A1:B5", "B2:C6|||range|B2:C6"}},
		{"union in parentheses", "SUM((A1:A3,C1:C3))", []string{"A1:A3|||range|A1:A3", "C1:C3|||range|C1:C3"}},
		{"range to a function", "A1:INDEX(B:B,3)", []string{"A1|||cell|A1", "B:B|||cols|B:B"}},
		{"three-part range", "A1:B2:C3", []string{"A1:B2|||range|A1:B2", "C3|||cell|C3"}},
		{"range with spaces around the colon is one range", "A1 : B2", []string{"A1 : B2|||range|A1:B2"}},
		{"range with a space before the colon", "SUM(A1 :B2)", []string{"A1 :B2|||range|A1:B2"}},
		{"range with spaces after the colon", "SUM($A$1:  B2)", []string{"$A$1:  B2|||range|$A$1:B2"}},
		{"sheet range with spaces", "Sheet1!A1 : B2", []string{"Sheet1!A1 : B2|Sheet1||range|A1:B2"}},
		{"spaces then a function is one cell", "A1 : INDEX(B:B,3)", []string{"A1|||cell|A1", "B:B|||cols|B:B"}},
		{"intersection is not a range", "A1 B2", []string{"A1|||cell|A1", "B2|||cell|B2"}},
		{"columns with spaces are not a range", "A : C", nil},
		{"unary and percent", "-A1%+B1^2", []string{"A1|||cell|A1", "B1|||cell|B1"}},
		{"implicit intersection", "@A1:A5", []string{"A1:A5|||range|A1:A5"}},
		{"leading equals sign", "=A1", []string{"A1|||cell|A1"}},

		// Not sites.
		{"function with digits", "LOG10(100)", nil},
		{"function names", "SUM(MAX(1,2),PI())", nil},
		{"function with a dot", "STDEV.S(A1:A3)", []string{"A1:A3|||range|A1:A3"}},
		{"booleans", "IF(TRUE,FALSE,true)", nil},
		{"names", "Total+running_sum+Tax.Rate", nil},
		{"name with a backslash and an underscore", `\name+_x1`, nil},
		{"name that is longer than a column", "ABCD1+Total1", nil},
		{"column out of range", "XFE1+ZZZ9", nil},
		{"row out of range", "A1048577+A0", nil},
		{"single column is a name", "A+B", nil},
		{"numbers", "1+2.5+.5+1E5+1.5e-3+3E+2", nil},
		{"number with an exponent is not a cell", "1E5+2e3", nil},
		{"number then name", "2abc1", nil},
		{"string", `"A1"&"B2:C3"`, nil},
		{"string with doubled quotes", `"say ""A1"" now"&B1`, []string{"B1|||cell|B1"}},
		{"unterminated string", `SUM(A1,"abc B2`, []string{"A1|||cell|A1"}},
		{"array literal", "{1,2;3,4}", nil},
		{"array literal next to a reference", "SUM({1,2;3,4}*A1:B2)", []string{"A1:B2|||range|A1:B2"}},
		{"error literals", "#REF!+#N/A+#DIV/0!+#NAME?+#VALUE!", nil},
		{"error then reference", "IFERROR(#N/A,A1)", []string{"A1|||cell|A1"}},
		{"deleted sheet form", "#REF!A1+B2", []string{"B2|||cell|B2"}},
		{"sheet with an error", "Sheet1!#REF!+A1", []string{"A1|||cell|A1"}},
		{"sheet with a name", "Sheet1!Total+A1", []string{"A1|||cell|A1"}},
		{"empty", "", nil},
		{"only spaces", "   ", nil},
		{"lone dollar", "$+A1", []string{"A1|||cell|A1"}},
		{"lone quote", "'", nil},
		{"quoted name with no bang", "'abc'+A1", []string{"A1|||cell|A1"}},
		{"empty quoted sheet", "''!A1", nil},

		// Opaque spans.
		{"structured reference", "SUM(Table1[Col])+A1", []string{"A1|||cell|A1"}},
		{"structured reference with a cell-like table name", "SUM(Tbl1[Col1])", nil},
		{"structured reference this row", "Table1[[#This Row],[Col A1]]*B2", []string{"B2|||cell|B2"}},
		{"structured reference with an escaped bracket", "Table1[Col'] A1]+B2", []string{"B2|||cell|B2"}},
		{"unterminated bracket", "Table1[Col+A1", nil},
		{"external workbook", "[Book.xlsx]Sheet1!A1+B2", []string{"B2|||cell|B2"}},
		{"external workbook range", "SUM([1]Sheet1!$A$1:$B$2)+C3", []string{"C3|||cell|C3"}},
		{"quoted external workbook", "'[Book 1.xlsx]My Sheet'!A1:B2+C3", []string{"C3|||cell|C3"}},
		{"spill operator", "C12#", []string{"C12|||cell|C12"}},
		{"spill operator in a call", "SUM(Sheet1!C12#)+D1", []string{"Sheet1!C12|Sheet1||cell|C12", "D1|||cell|D1"}},
		{"characters that the lexer rejects", "A1 ? B2 ~ `C3`", []string{"A1|||cell|A1", "B2|||cell|B2", "C3|||cell|C3"}},
		{"text that is not ASCII", "A1&\"é\"&ß&B2", []string{"A1|||cell|A1", "B2|||cell|B2"}},
		{"name that is not ASCII with a cell-like tail", "Größe1+A1", []string{"A1|||cell|A1"}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := describeSites(tt.src)
			if strings.Join(got, "\n") != strings.Join(tt.want, "\n") {
				t.Errorf("ScanRefs(%q):\n got  %q\n want %q", tt.src, got, tt.want)
			}
		})
	}
}

func TestScanRefsSpans(t *testing.T) {
	src := "  SUM( 'My Sheet'!$A$1:B2 , c3 )+Sheet2!D:E"
	sites := ScanRefs(src)
	if len(sites) != 3 {
		t.Fatalf("got %d sites, want 3", len(sites))
	}
	want := []struct {
		start, end int
		from, to   RefPoint
	}{
		{7, 25, RefPoint{Col: 1, Row: 1, AbsCol: true, AbsRow: true}, RefPoint{Col: 2, Row: 2}},
		{28, 30, RefPoint{Col: 3, Row: 3}, RefPoint{Col: 3, Row: 3}},
		{33, 43, RefPoint{Col: 4}, RefPoint{Col: 5}},
	}
	for i, w := range want {
		s := sites[i]
		if s.Start != w.start || s.End != w.end {
			t.Errorf("site %d span = [%d,%d) %q, want [%d,%d) %q", i, s.Start, s.End, src[s.Start:s.End], w.start, w.end, src[w.start:w.end])
		}
		if s.From != w.from || s.To != w.to {
			t.Errorf("site %d corners = %+v %+v, want %+v %+v", i, s.From, s.To, w.from, w.to)
		}
	}
	// Sites are in text order and do not overlap.
	last := 0
	for i, s := range sites {
		if s.Start < last || s.End <= s.Start {
			t.Errorf("site %d [%d,%d) overlaps or is empty", i, s.Start, s.End)
		}
		last = s.End
	}
}

// realFormulas is a list of formulas of the kinds that workbooks have. A
// rewrite that keeps all sites must return each one byte for byte.
var realFormulas = []string{
	"SUM(A1:A10)",
	"sum( a1:a10 )",
	"  A1 +   B1  ",
	"IF(A1>0,\"yes\",\"no\")",
	"IF(AND(A1>=1,B1<>\"\"),C1*1.5,0)",
	"VLOOKUP($A2,Data!$A$1:$D$100,3,FALSE)",
	"INDEX('Price List'!$B:$B,MATCH(A2,'Price List'!$A:$A,0))",
	"XLOOKUP(A2,Sheet2!A:A,Sheet2!C:C,\"not found\")",
	"SUMIFS(Sales!$C:$C,Sales!$A:$A,$A2,Sales!$B:$B,\">=\"&B$1)",
	"COUNTIF(A:A,\"*text*\")",
	"TEXT(TODAY(),\"yyyy-mm-dd\")",
	"A1&\" \"&B1",
	"\"He said \"\"hi\"\" to \"&A1",
	"LET(x,A1*2,y,B1+1,x+y)",
	"LAMBDA(a,b,a+b)(A1,B1)",
	"SUM(Table1[Amount])",
	"Table1[[#This Row],[Qty]]*Table1[[#This Row],[Price]]",
	"SUM(Table1[[#Totals],[Amount]])",
	"[Budget.xlsx]Sheet1!$A$1+1",
	"'[Budget 2024.xlsx]Jan'!B2",
	"SUM(Sheet1:Sheet12!B5)",
	"SUM('Jan 2024:Dec 2024'!B5)",
	"SUM(1:1)",
	"SUM($1:$3)",
	"SUM(Sheet1!$1:$2)",
	"SUM(A:A)/COUNT(A:A)",
	"C12#",
	"SORT(FILTER(A2:C100,B2:B100>0),2,-1)",
	"SUM(A2#)",
	"{1,2;3,4}",
	"SUM({1,2,3}*A1:C1)",
	"-A1^2%",
	"(A1+B1)*(C1-D1)/E1",
	"A1:B5 B2:C6",
	"SUM((A1:A3,C1:C3))",
	"A1:INDEX(B:B,3)",
	"OFFSET($A$1,ROW()-1,0,1,3)",
	"INDIRECT(\"A\"&ROW())",
	"IFERROR(1/0,#N/A)",
	"#REF!+1",
	"Sheet1!#REF!",
	"#REF!A1",
	"@A1:A10",
	"1E5+2.5e-3+.5",
	"LOG10(A1)+ATAN2(B1,C1)+DEC2BIN(D1)",
	"Total*Tax_Rate+my.name",
	"ROUND(AVERAGE(B2:B13),2)&\" avg\"",
	"IF(ISBLANK('My Sheet'!A1),\"\",'My Sheet'!A1)",
	"'It''s'!A1+'It''s'!$B$2",
	"TRUE+FALSE+true",
	"Größe!A1+Ünïcödé",
	"SUM(\"abc",
	"A1 ? B2",
	"",
}

func TestRewriteRefsKeepIsByteExact(t *testing.T) {
	if len(realFormulas) < 40 {
		t.Fatalf("the formula list has %d entries, want 40 or more", len(realFormulas))
	}
	for _, src := range realFormulas {
		calls := 0
		got := RewriteRefs(src, func(s RefSite) (RefSite, RefAction) {
			calls++
			return s, RefKeep
		})
		if got != src {
			t.Errorf("RewriteRefs(keep) changed the text:\n got  %q\n want %q", got, src)
		}
		if calls != len(ScanRefs(src)) {
			t.Errorf("%q: callback ran %d times for %d sites", src, calls, len(ScanRefs(src)))
		}
		// A replace with the unchanged site changes only letter case of
		// column letters: the text is equal with no letter-case compare.
		got = RewriteRefs(src, func(s RefSite) (RefSite, RefAction) { return s, RefReplace })
		if !strings.EqualFold(got, src) {
			t.Errorf("RewriteRefs(replace same) changed the text:\n got  %q\n want %q", got, src)
		}
		// A shift of zero is the identity.
		if got := ShiftRefs(src, 0, 0); got != src {
			t.Errorf("ShiftRefs(0,0) changed %q to %q", src, got)
		}
	}
}

func TestRewriteRefsReplaceAndInvalid(t *testing.T) {
	tests := []struct {
		name string
		src  string
		fn   func(RefSite) (RefSite, RefAction)
		want string
	}{
		{
			name: "replace one of two; other text stays",
			src:  "  sum( a1 ,  B2 )&\"A1\"",
			fn: func(s RefSite) (RefSite, RefAction) {
				if s.From.Col == 1 {
					s.From.Row, s.To.Row = 7, 7
					return s, RefReplace
				}
				return s, RefKeep
			},
			want: "  sum( A7 ,  B2 )&\"A1\"",
		},
		{
			name: "invalid writes #REF! with no qualifier",
			src:  "Sheet2!A5+'My Sheet'!B1:B3+C1",
			fn: func(s RefSite) (RefSite, RefAction) {
				if s.Sheet != "" {
					return s, RefInvalid
				}
				return s, RefKeep
			},
			want: "#REF!+#REF!+C1",
		},
		{
			name: "invalid takes the spill operator with it",
			src:  "SUM(C12#)+D1#",
			fn: func(s RefSite) (RefSite, RefAction) {
				if s.From.Col == 3 {
					return s, RefInvalid
				}
				return s, RefKeep
			},
			want: "SUM(#REF!)+D1#",
		},
		{
			name: "the quote style of the qualifier stays",
			src:  "'Sheet1'!A1+Sheet1!A1+'My Sheet'!A1+'It''s'!A1",
			fn: func(s RefSite) (RefSite, RefAction) {
				s.From.Col, s.To.Col = 2, 2
				return s, RefReplace
			},
			want: "'Sheet1'!B1+Sheet1!B1+'My Sheet'!B1+'It''s'!B1",
		},
		{
			name: "a changed sheet name is quoted when necessary",
			src:  "Sheet1!A1+Sheet1!B2+Jan:Dec!C3",
			fn: func(s RefSite) (RefSite, RefAction) {
				switch s.From.Col {
				case 1:
					s.Sheet = "My Sheet"
				case 2:
					s.Sheet = "Other"
				case 3:
					s.SheetEnd = "It's"
				}
				return s, RefReplace
			},
			want: "'My Sheet'!A1+Other!B2+'Jan:It''s'!C3",
		},
		{
			name: "kind change",
			src:  "SUM(A1:B2)",
			fn: func(s RefSite) (RefSite, RefAction) {
				s.Kind = RefCell
				return s, RefReplace
			},
			want: "SUM(A1)",
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := RewriteRefs(tt.src, tt.fn); got != tt.want {
				t.Errorf("RewriteRefs(%q) = %q, want %q", tt.src, got, tt.want)
			}
		})
	}
}

func TestFormatRef(t *testing.T) {
	tests := []struct {
		site RefSite
		want string
	}{
		{RefSite{Kind: RefCell, From: RefPoint{Col: 1, Row: 1}}, "A1"},
		{RefSite{Kind: RefCell, From: RefPoint{Col: 28, Row: 10, AbsCol: true}}, "$AB10"},
		{RefSite{Kind: RefRange, From: RefPoint{Col: 1, Row: 1, AbsRow: true}, To: RefPoint{Col: 3, Row: 9, AbsCol: true, AbsRow: true}}, "A$1:$C$9"},
		{RefSite{Kind: RefCols, From: RefPoint{Col: 1}, To: RefPoint{Col: 16384, AbsCol: true}}, "A:$XFD"},
		{RefSite{Kind: RefRows, From: RefPoint{Row: 1, AbsRow: true}, To: RefPoint{Row: 1048576}}, "$1:1048576"},
		{RefSite{Sheet: "Data", Kind: RefCell, From: RefPoint{Col: 2, Row: 2}}, "Data!B2"},
		{RefSite{Sheet: "My Data", Kind: RefCols, From: RefPoint{Col: 2}, To: RefPoint{Col: 2}}, "'My Data'!B:B"},
		{RefSite{Sheet: "Jan", SheetEnd: "Dec", Kind: RefCell, From: RefPoint{Col: 2, Row: 2}}, "Jan:Dec!B2"},
	}
	for _, tt := range tests {
		if got := FormatRef(tt.site); got != tt.want {
			t.Errorf("FormatRef(%+v) = %q, want %q", tt.site, got, tt.want)
		}
		// The written form scans back to the same site.
		sites := ScanRefs(tt.want)
		if len(sites) != 1 {
			t.Errorf("ScanRefs(%q) gave %d sites, want 1", tt.want, len(sites))
			continue
		}
		got := sites[0]
		if got.Kind != tt.site.Kind || got.From != tt.site.From || got.Sheet != tt.site.Sheet || got.SheetEnd != tt.site.SheetEnd {
			t.Errorf("ScanRefs(%q) = %+v, want %+v", tt.want, got, tt.site)
		}
		if tt.site.Kind != RefCell && got.To != tt.site.To {
			t.Errorf("ScanRefs(%q).To = %+v, want %+v", tt.want, got.To, tt.site.To)
		}
	}
}

// The scanner must agree with the lexer on the formulas that the lexer
// accepts: each cell-like TokCellRef token is inside a site.
func TestScanRefsAgreesWithLexer(t *testing.T) {
	for _, src := range realFormulas {
		tokens, err := Tokenize(src)
		if err != nil {
			continue // the scanner is tolerant; the lexer is not
		}
		sites := ScanRefs(src)
		for _, tok := range tokens {
			if tok.Type != TokCellRef {
				continue
			}
			body := tok.Value
			if i := strings.LastIndexByte(body, '!'); i >= 0 {
				body = body[i+1:]
			}
			if !LooksLikeCellRef(body) {
				continue // a name or a column-only part
			}
			// Two forms are opaque by design: a reference to a different
			// workbook, and the text after a deleted-sheet #REF!.
			if strings.Contains(tok.Value, "[") || strings.HasSuffix(src[:tok.Pos], "#REF!") {
				continue
			}
			inSite := false
			for _, s := range sites {
				if tok.Pos >= s.Start && tok.End <= s.End {
					inSite = true
				}
			}
			if !inSite {
				t.Errorf("%q: lexer token %v is in no site %q", src, tok, describeSites(src))
			}
		}
	}
}

func FuzzRewriteRefsKeep(f *testing.F) {
	for _, src := range realFormulas {
		f.Add(src)
	}
	f.Add("'")
	f.Add("[")
	f.Add("$$A$$1:")
	f.Add("A1:B2:C3!D4#'x'!")
	f.Fuzz(func(t *testing.T, src string) {
		sites := ScanRefs(src)
		last := 0
		for _, s := range sites {
			if s.Start < last || s.End <= s.Start || s.End > len(src) {
				t.Fatalf("%q: bad span [%d,%d) after %d", src, s.Start, s.End, last)
			}
			last = s.End
		}
		if got := RewriteRefs(src, func(s RefSite) (RefSite, RefAction) { return s, RefKeep }); got != src {
			t.Fatalf("RewriteRefs(keep) changed %q to %q", src, got)
		}
		// A replace with the same site scans to the same number of sites.
		out := RewriteRefs(src, func(s RefSite) (RefSite, RefAction) { return s, RefReplace })
		if len(ScanRefs(out)) != len(sites) {
			t.Fatalf("%q: %d sites; after a same-site rewrite %q has %d", src, len(sites), out, len(ScanRefs(out)))
		}
		_ = ShiftRefs(src, 3, -2)
	})
}

func TestTokenEnd(t *testing.T) {
	tests := []struct {
		src  string
		want []string // source text of each token span, without the EOF token
	}{
		{"SUM(A1:B2)", []string{"SUM(", "A1", ":", "B2", ")"}},
		{" 1.5 +  \"a\"\"b\" ", []string{"1.5", "+", `"a""b"`}},
		{"'My Sheet'!A1<=Sheet2!$B$2", []string{"'My Sheet'!A1", "<=", "Sheet2!$B$2"}},
		{"{1,2;3}", []string{"{", "1", ",", "2", ";", "3", "}"}},
		{"#DIV/0!&TRUE", []string{"#DIV/0!", "&", "TRUE"}},
		{"A1:B5 B2:C6", []string{"A1", ":", "B5", " ", "B2", ":", "C6"}},
		{"A1   B2", []string{"A1", "   ", "B2"}},
		{"@A1%", []string{"@", "A1", "%"}},
	}
	for _, tt := range tests {
		tokens, err := Tokenize(tt.src)
		if err != nil {
			t.Errorf("Tokenize(%q): %v", tt.src, err)
			continue
		}
		var got []string
		for _, tok := range tokens {
			if tok.Type == TokEOF {
				if tok.Pos != len(tt.src) || tok.End != len(tt.src) {
					t.Errorf("%q: EOF token span = [%d,%d), want [%d,%d)", tt.src, tok.Pos, tok.End, len(tt.src), len(tt.src))
				}
				continue
			}
			if tok.End < tok.Pos || tok.End > len(tt.src) {
				t.Errorf("%q: token %v has End %d", tt.src, tok, tok.End)
				continue
			}
			got = append(got, tt.src[tok.Pos:tok.End])
		}
		if strings.Join(got, "\x00") != strings.Join(tt.want, "\x00") {
			t.Errorf("Tokenize(%q) spans = %q, want %q", tt.src, got, tt.want)
		}
	}
}
