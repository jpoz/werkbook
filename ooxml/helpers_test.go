package ooxml

import (
	"encoding/xml"
	"strings"
	"testing"
)

// ---------------------------------------------------------------------------
// parseOOXMLBoolString
// ---------------------------------------------------------------------------

func TestParseOOXMLBoolString(t *testing.T) {
	tests := []struct {
		in   string
		want bool
	}{
		{"1", true},
		{"true", true},
		{"on", true},
		{"TRUE", true},
		{"True", true},
		{"ON", true},
		{"0", false},
		{"false", false},
		{"off", false},
		{"", false},
		{"yes", false},   // not recognized → false
		{"maybe", false}, // not recognized → false
	}
	for _, tt := range tests {
		t.Run(tt.in, func(t *testing.T) {
			if got := parseOOXMLBoolString(tt.in); got != tt.want {
				t.Errorf("parseOOXMLBoolString(%q) = %v, want %v", tt.in, got, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// decodeOOXMLEscapes
// ---------------------------------------------------------------------------

func TestDecodeOOXMLEscapes(t *testing.T) {
	tests := []struct {
		name string
		in   string
		want string
	}{
		{"no escapes", "hello", "hello"},
		{"tab escape", "_x0009_", "\t"},
		{"newline escape", "_x000A_", "\n"},
		{"carriage return", "_x000D_", "\r"},
		{"underscore literal", "_x005F_", "_"},
		{"mid-string escape", "col_x000A_name", "col\nname"},
		{"multiple escapes", "_x0009_a_x000A_b", "\ta\nb"},
		{"incomplete escape", "_xZZZZ_", "_xZZZZ_"},
		{"too short", "_x00_", "_x00_"},
		{"no trailing underscore", "_x0009X", "_x0009X"},
		{"empty string", "", ""},
		{"space char", "_x0020_", " "},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := decodeOOXMLEscapes(tt.in); got != tt.want {
				t.Errorf("decodeOOXMLEscapes(%q) = %q, want %q", tt.in, got, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// encodeOOXMLEscapes
// ---------------------------------------------------------------------------

func TestEncodeOOXMLEscapes(t *testing.T) {
	tests := []struct {
		name string
		in   string
		want string
	}{
		{"plain text", "hello", "hello"},
		{"tab", "\t", "_x0009_"},
		{"newline", "\n", "_x000A_"},
		{"carriage return", "\r", "_x000D_"},
		{"control char", "\x01", "_x0001_"},
		{"existing escape pattern gets escaped", "_x0009_", "_x005F__x0009_"},
		{"normal underscore no hex", "a_b", "a_b"},
		{"printable chars unchanged", "abc 123 !@#", "abc 123 !@#"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := encodeOOXMLEscapes(tt.in); got != tt.want {
				t.Errorf("encodeOOXMLEscapes(%q) = %q, want %q", tt.in, got, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// encode/decode round-trip
// ---------------------------------------------------------------------------

func TestOOXMLEscapeRoundTrip(t *testing.T) {
	values := []string{
		"simple",
		"has\ttab",
		"has\nnewline",
		"\x01\x02\x03 control chars",
		"no special underscores_here",
	}
	for _, v := range values {
		encoded := encodeOOXMLEscapes(v)
		decoded := decodeOOXMLEscapes(encoded)
		if decoded != v {
			t.Errorf("round-trip failed for %q: encoded=%q decoded=%q", v, encoded, decoded)
		}
	}
}

// ---------------------------------------------------------------------------
// looksLikeOOXMLEscape
// ---------------------------------------------------------------------------

func TestLooksLikeOOXMLEscape(t *testing.T) {
	tests := []struct {
		s    string
		i    int
		want bool
	}{
		{"_x0009_", 0, true},
		{"_xABCD_", 0, true},
		{"_xabcd_", 0, true},
		{"_xGGGG_", 0, false}, // not hex
		{"_x00_", 0, false},   // too short
		{"hello", 0, false},
		{"ab_x0009_cd", 2, true},
	}
	for _, tt := range tests {
		if got := looksLikeOOXMLEscape(tt.s, tt.i); got != tt.want {
			t.Errorf("looksLikeOOXMLEscape(%q, %d) = %v, want %v", tt.s, tt.i, got, tt.want)
		}
	}
}

// ---------------------------------------------------------------------------
// resolveRelativePath
// ---------------------------------------------------------------------------

func TestResolveRelativePath(t *testing.T) {
	tests := []struct {
		base, target, want string
	}{
		{"xl/worksheets", "../tables/table1.xml", "xl/tables/table1.xml"},
		{"xl/worksheets", "table1.xml", "xl/worksheets/table1.xml"},
		{"xl", "../other/file.xml", "other/file.xml"},
		{"", "tables/table1.xml", "tables/table1.xml"},
		{"a/b/c", "../../d/e.xml", "a/d/e.xml"},
	}
	for _, tt := range tests {
		t.Run(tt.base+"/"+tt.target, func(t *testing.T) {
			if got := resolveRelativePath(tt.base, tt.target); got != tt.want {
				t.Errorf("resolveRelativePath(%q, %q) = %q, want %q", tt.base, tt.target, got, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// boolString
// ---------------------------------------------------------------------------

func TestBoolString(t *testing.T) {
	if got := boolString(true); got != "1" {
		t.Errorf("boolString(true) = %q, want \"1\"", got)
	}
	if got := boolString(false); got != "" {
		t.Errorf("boolString(false) = %q, want \"\"", got)
	}
}

// ---------------------------------------------------------------------------
// hasCalcProps
// ---------------------------------------------------------------------------

func TestHasCalcProps(t *testing.T) {
	if hasCalcProps(CalcPropertiesData{}) {
		t.Error("empty CalcPropertiesData should return false")
	}
	if !hasCalcProps(CalcPropertiesData{Mode: "auto"}) {
		t.Error("CalcPropertiesData{Mode: auto} should return true")
	}
	if !hasCalcProps(CalcPropertiesData{ID: 1}) {
		t.Error("CalcPropertiesData{ID: 1} should return true")
	}
	if !hasCalcProps(CalcPropertiesData{FullCalcOnLoad: true}) {
		t.Error("CalcPropertiesData{FullCalcOnLoad: true} should return true")
	}
	if !hasCalcProps(CalcPropertiesData{ForceFullCalc: true}) {
		t.Error("CalcPropertiesData{ForceFullCalc: true} should return true")
	}
	if !hasCalcProps(CalcPropertiesData{Completed: true}) {
		t.Error("CalcPropertiesData{Completed: true} should return true")
	}
}

// ---------------------------------------------------------------------------
// boolOOXML
// ---------------------------------------------------------------------------

func TestBoolOOXML(t *testing.T) {
	if boolOOXML(true) != 1 {
		t.Error("boolOOXML(true) should be 1")
	}
	if boolOOXML(false) != 0 {
		t.Error("boolOOXML(false) should be 0")
	}
}

// ---------------------------------------------------------------------------
// ooxmlBool UnmarshalXMLAttr
// ---------------------------------------------------------------------------

func TestOoxmlBoolUnmarshalXMLAttr(t *testing.T) {
	tests := []struct {
		val  string
		want ooxmlBool
		err  bool
	}{
		{"1", 1, false},
		{"true", 1, false},
		{"on", 1, false},
		{"0", 0, false},
		{"false", 0, false},
		{"off", 0, false},
		{"", 0, false},
		{"2", 1, false},  // non-zero int → true
		{"-1", 1, false}, // non-zero int → true
		{"abc", 0, true}, // invalid → error
	}
	for _, tt := range tests {
		t.Run(tt.val, func(t *testing.T) {
			var b ooxmlBool
			err := b.UnmarshalXMLAttr(xml.Attr{Value: tt.val})
			if (err != nil) != tt.err {
				t.Errorf("err = %v, wantErr = %v", err, tt.err)
			}
			if err == nil && b != tt.want {
				t.Errorf("ooxmlBool = %d, want %d", b, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// isExternalRef
// ---------------------------------------------------------------------------

func TestIsExternalRef(t *testing.T) {
	tests := []struct {
		value string
		want  bool
	}{
		{"[1]Sheet!$A$1", true},
		{"'[1]Other Sheet'!$G$5", true},
		{"[2]Sheet2!$B$2", true},
		{"'[99]Sheet'!$A$1", true},
		{"Sheet1!$A$1", false},
		{"Sheet1!$A$1", false},
		{"", false},
		{"SUM(A1:A10)", false},
		// Brackets in the middle should not be flagged.
		{"Sheet1!$A$1+[ignored]", false},
		{"INDEX(A1:A10,[1])", false},
	}
	for _, tt := range tests {
		t.Run(tt.value, func(t *testing.T) {
			if got := isExternalRef(tt.value); got != tt.want {
				t.Errorf("isExternalRef(%q) = %v, want %v", tt.value, got, tt.want)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// sharedIndex
// ---------------------------------------------------------------------------

func TestSharedIndex(t *testing.T) {
	if got := sharedIndex(nil); got != -1 {
		t.Errorf("sharedIndex(nil) = %d, want -1", got)
	}
	if got := sharedIndex(&xlsxF{T: "array"}); got != -1 {
		t.Errorf("sharedIndex(array) = %d, want -1", got)
	}
	if got := sharedIndex(&xlsxF{T: ""}); got != -1 {
		t.Errorf("sharedIndex(empty type) = %d, want -1", got)
	}
	if got := sharedIndex(&xlsxF{T: "shared", Si: 0}); got != 0 {
		t.Errorf("sharedIndex(shared, si=0) = %d, want 0", got)
	}
	if got := sharedIndex(&xlsxF{T: "shared", Si: 5}); got != 5 {
		t.Errorf("sharedIndex(shared, si=5) = %d, want 5", got)
	}
}

// ---------------------------------------------------------------------------
// SharedStringTable
// ---------------------------------------------------------------------------

func TestSharedStringTable(t *testing.T) {
	sst := NewSharedStringTable()

	if sst.Len() != 0 {
		t.Fatalf("new SST Len() = %d, want 0", sst.Len())
	}

	idx0 := sst.Add("hello")
	idx1 := sst.Add("world")
	idx2 := sst.Add("hello") // duplicate

	if idx0 != 0 {
		t.Errorf("first Add = %d, want 0", idx0)
	}
	if idx1 != 1 {
		t.Errorf("second Add = %d, want 1", idx1)
	}
	if idx2 != 0 {
		t.Errorf("duplicate Add = %d, want 0 (dedup)", idx2)
	}
	if sst.Len() != 2 {
		t.Errorf("Len = %d, want 2", sst.Len())
	}
	if sst.Get(0) != "hello" {
		t.Errorf("Get(0) = %q, want hello", sst.Get(0))
	}
	if sst.Get(1) != "world" {
		t.Errorf("Get(1) = %q, want world", sst.Get(1))
	}

	strs := sst.Strings()
	if len(strs) != 2 || strs[0] != "hello" || strs[1] != "world" {
		t.Errorf("Strings() = %v, want [hello world]", strs)
	}
}

func TestSharedStringTableToXML(t *testing.T) {
	sst := NewSharedStringTable()
	sst.Add("alpha")
	sst.Add("beta")
	sst.Add("alpha") // duplicate: count=3, uniqueCount=2

	x := sst.ToXML()
	if x.Count != 3 {
		t.Errorf("Count = %d, want 3", x.Count)
	}
	if x.UniqueCount != 2 {
		t.Errorf("UniqueCount = %d, want 2", x.UniqueCount)
	}
	if len(x.SI) != 2 {
		t.Fatalf("len(SI) = %d, want 2", len(x.SI))
	}
	if *x.SI[0].T != "alpha" {
		t.Errorf("SI[0].T = %q, want alpha", *x.SI[0].T)
	}
	if *x.SI[1].T != "beta" {
		t.Errorf("SI[1].T = %q, want beta", *x.SI[1].T)
	}
}

// ---------------------------------------------------------------------------
// xlsxSI MarshalXML – whitespace preservation
// ---------------------------------------------------------------------------

func TestSIMarshalXML_WhitespacePreserve(t *testing.T) {
	tests := []struct {
		name         string
		value        string
		wantPreserve bool
	}{
		{"no whitespace", "hello", false},
		{"leading space", " hello", true},
		{"trailing space", "hello ", true},
		{"leading tab", "\thello", true},
		{"trailing tab", "hello\t", true},
		{"no special", "hello world", false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := tt.value
			si := xlsxSI{T: &s}
			data, err := xml.Marshal(si)
			if err != nil {
				t.Fatal(err)
			}
			xmlStr := string(data)
			hasPreserve := strings.Contains(xmlStr, `space="preserve"`)
			if hasPreserve != tt.wantPreserve {
				t.Errorf("xml:space=preserve present=%v, want %v; xml=%s", hasPreserve, tt.wantPreserve, xmlStr)
			}
		})
	}
}

// ---------------------------------------------------------------------------
// xlsxSheet MarshalXML / UnmarshalXML
// ---------------------------------------------------------------------------

func TestXlsxSheetMarshalRoundTrip(t *testing.T) {
	original := xlsxSheet{
		Name:    "Data",
		SheetID: 3,
		RID:     "rId3",
		State:   "hidden",
	}

	data, err := xml.Marshal(original)
	if err != nil {
		t.Fatal(err)
	}
	xmlStr := string(data)

	// Verify marshaled attributes.
	if !strings.Contains(xmlStr, `name="Data"`) {
		t.Errorf("missing name attr in %s", xmlStr)
	}
	if !strings.Contains(xmlStr, `sheetId="3"`) {
		t.Errorf("missing sheetId attr in %s", xmlStr)
	}
	if !strings.Contains(xmlStr, `r:id="rId3"`) {
		t.Errorf("missing r:id attr in %s", xmlStr)
	}
	if !strings.Contains(xmlStr, `state="hidden"`) {
		t.Errorf("missing state attr in %s", xmlStr)
	}
}

func TestXlsxSheetMarshalOmitsEmptyState(t *testing.T) {
	s := xlsxSheet{Name: "Sheet1", SheetID: 1, RID: "rId1"}
	data, err := xml.Marshal(s)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(data), "state") {
		t.Errorf("empty state should be omitted: %s", string(data))
	}
}

func TestXlsxSheetUnmarshalTransitional(t *testing.T) {
	xmlStr := `<sheet name="Sheet1" sheetId="1" xmlns:r="` + NSOfficeDocument + `" r:id="rId1" state="hidden"/>`
	var s xlsxSheet
	if err := xml.Unmarshal([]byte(xmlStr), &s); err != nil {
		t.Fatal(err)
	}
	if s.Name != "Sheet1" {
		t.Errorf("Name = %q, want Sheet1", s.Name)
	}
	if s.SheetID != 1 {
		t.Errorf("SheetID = %d, want 1", s.SheetID)
	}
	if s.RID != "rId1" {
		t.Errorf("RID = %q, want rId1", s.RID)
	}
	if s.State != "hidden" {
		t.Errorf("State = %q, want hidden", s.State)
	}
}

func TestXlsxSheetUnmarshalStrict(t *testing.T) {
	xmlStr := `<sheet name="Data" sheetId="2" xmlns:r="` + NSOfficeDocumentStrict + `" r:id="rId2"/>`
	var s xlsxSheet
	if err := xml.Unmarshal([]byte(xmlStr), &s); err != nil {
		t.Fatal(err)
	}
	if s.RID != "rId2" {
		t.Errorf("RID = %q, want rId2 (strict namespace)", s.RID)
	}
}

// ---------------------------------------------------------------------------
// xlsxTablePart MarshalXML
// ---------------------------------------------------------------------------

func TestXlsxTablePartMarshalXML(t *testing.T) {
	tp := xlsxTablePart{RID: "rId5"}
	data, err := xml.Marshal(tp)
	if err != nil {
		t.Fatal(err)
	}
	xmlStr := string(data)
	if !strings.Contains(xmlStr, `r:id="rId5"`) {
		t.Errorf("missing r:id attr in %s", xmlStr)
	}
	if !strings.Contains(xmlStr, "tablePart") {
		t.Errorf("missing tablePart element in %s", xmlStr)
	}
}

// ---------------------------------------------------------------------------
// buildTableWritePlan
// ---------------------------------------------------------------------------

func TestBuildTableWritePlan(t *testing.T) {
	tables := []TableDef{
		{Name: "T1", SheetIndex: 0},
		{Name: "T2", SheetIndex: 1},
		{Name: "T3", SheetIndex: 0},
	}
	plan := buildTableWritePlan(tables)

	if len(plan[0]) != 2 {
		t.Errorf("sheet 0 tables = %d, want 2", len(plan[0]))
	}
	if len(plan[1]) != 1 {
		t.Errorf("sheet 1 tables = %d, want 1", len(plan[1]))
	}
	// PartNum is 1-based index into the global tables slice.
	if plan[0][0].PartNum != 1 {
		t.Errorf("plan[0][0].PartNum = %d, want 1", plan[0][0].PartNum)
	}
	if plan[0][1].PartNum != 3 {
		t.Errorf("plan[0][1].PartNum = %d, want 3", plan[0][1].PartNum)
	}
	if plan[1][0].PartNum != 2 {
		t.Errorf("plan[1][0].PartNum = %d, want 2", plan[1][0].PartNum)
	}
}

// ---------------------------------------------------------------------------
// siToString
// ---------------------------------------------------------------------------

func TestSiToString(t *testing.T) {
	plain := "hello"
	si := xlsxSI{T: &plain}
	if got := siToString(si); got != "hello" {
		t.Errorf("siToString plain = %q, want hello", got)
	}

	// Rich text.
	richSI := xlsxSI{R: []xlsxR{{T: "Hello "}, {T: "World"}}}
	if got := siToString(richSI); got != "Hello World" {
		t.Errorf("siToString rich = %q, want 'Hello World'", got)
	}

	// Empty.
	emptySI := xlsxSI{}
	if got := siToString(emptySI); got != "" {
		t.Errorf("siToString empty = %q, want empty", got)
	}
}

// ---------------------------------------------------------------------------
// formulaType / formulaRef
// ---------------------------------------------------------------------------

func TestFormulaTypeAndRef(t *testing.T) {
	if got := formulaType(nil); got != "" {
		t.Errorf("formulaType(nil) = %q, want empty", got)
	}
	if got := formulaRef(nil); got != "" {
		t.Errorf("formulaRef(nil) = %q, want empty", got)
	}

	fe := &xlsxF{T: "array", Ref: "A1:B2"}
	if got := formulaType(fe); got != "array" {
		t.Errorf("formulaType = %q, want array", got)
	}
	if got := formulaRef(fe); got != "A1:B2" {
		t.Errorf("formulaRef = %q, want A1:B2", got)
	}
}

// ---------------------------------------------------------------------------
// parseCellData – additional cases
// ---------------------------------------------------------------------------

func TestParseCellData_InlineStr(t *testing.T) {
	c := xlsxC{R: "A1", T: "inlineStr", IS: &xlsxIS{T: "inline text"}}
	cd := parseCellData(c, nil)
	if cd.Type != "inlineStr" || cd.Value != "inline text" {
		t.Errorf("inlineStr: Type=%q Value=%q", cd.Type, cd.Value)
	}
}

func TestParseCellData_StrType(t *testing.T) {
	c := xlsxC{R: "A1", T: "str", V: "formula result"}
	cd := parseCellData(c, nil)
	if cd.Type != "str" || cd.Value != "formula result" {
		t.Errorf("str: Type=%q Value=%q", cd.Type, cd.Value)
	}
}

func TestParseCellData_ErrorType(t *testing.T) {
	c := xlsxC{R: "A1", T: "e", V: "#DIV/0!"}
	cd := parseCellData(c, nil)
	if cd.Type != "e" || cd.Value != "#DIV/0!" {
		t.Errorf("error: Type=%q Value=%q", cd.Type, cd.Value)
	}
}

func TestParseCellData_DateType(t *testing.T) {
	c := xlsxC{R: "A1", T: "d", V: "2024-01-01"}
	cd := parseCellData(c, nil)
	if cd.Type != "d" || cd.Value != "2024-01-01" {
		t.Errorf("date: Type=%q Value=%q", cd.Type, cd.Value)
	}
}

func TestParseCellData_WithFormula(t *testing.T) {
	c := xlsxC{
		R:  "A1",
		FE: &xlsxF{Text: "SUM(B1:B10)", T: "array", Ref: "A1:A5"},
		V:  "55",
	}
	cd := parseCellData(c, nil)
	if cd.Formula != "SUM(B1:B10)" {
		t.Errorf("Formula = %q, want SUM(B1:B10)", cd.Formula)
	}
	if cd.FormulaType != "array" {
		t.Errorf("FormulaType = %q, want array", cd.FormulaType)
	}
	if cd.FormulaRef != "A1:A5" {
		t.Errorf("FormulaRef = %q, want A1:A5", cd.FormulaRef)
	}
	if !cd.IsArrayFormula {
		t.Error("expected IsArrayFormula=true")
	}
}

func TestParseCellData_DynamicArray(t *testing.T) {
	c := xlsxC{
		R:  "A1",
		CM: 1,
		FE: &xlsxF{Text: "SORT(B1:B10)", T: "array", Ref: "A1:A10"},
	}
	cd := parseCellData(c, nil)
	if !cd.IsDynamicArray {
		t.Error("expected IsDynamicArray=true for CM=1 + array formula")
	}
	if cd.IsArrayFormula {
		t.Error("expected IsArrayFormula=false for dynamic array (CM!=0)")
	}
}

func TestParseCellData_DynamicArrayPlainFormula(t *testing.T) {
	c := xlsxC{
		R:  "A1",
		FE: &xlsxF{Text: "SORT(B1:B10)"},
	}
	cd := parseCellData(c, nil)
	if !cd.IsDynamicArray {
		t.Error("expected IsDynamicArray=true for plain dynamic-array formulas")
	}
	if cd.IsArrayFormula {
		t.Error("expected IsArrayFormula=false for plain dynamic-array formulas")
	}
}

func TestParseCellData_StyleIdx(t *testing.T) {
	c := xlsxC{R: "A1", S: 5, V: "42"}
	cd := parseCellData(c, nil)
	if cd.StyleIdx != 5 {
		t.Errorf("StyleIdx = %d, want 5", cd.StyleIdx)
	}
}

func TestParseCellData_InvalidSSTIndex(t *testing.T) {
	sst := []string{"only one"}
	c := xlsxC{R: "A1", T: "s", V: "99"} // out of range
	cd := parseCellData(c, sst)
	// Should fall through without setting value.
	if cd.Value != "" {
		t.Errorf("invalid SST index: Value = %q, want empty", cd.Value)
	}
}
