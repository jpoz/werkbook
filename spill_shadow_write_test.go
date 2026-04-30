package werkbook_test

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"path"
	"strings"
	"testing"

	"github.com/jpoz/werkbook"
)

// TestSpill_WritesCachedFollowerValues verifies that dynamic-array spill
// follower cells are persisted with their cached values on save. Excel does
// not need follower values (it re-evaluates the anchor's formula on load),
// but Apple Numbers and other consumers cannot evaluate `_xlfn._xlws.FILTER`
// and rely on the cached follower values to render each spilled row.
func TestSpill_WritesCachedFollowerValues(t *testing.T) {
	f, data, spill, _ := buildVerticalFilterHarness(t)

	f.Recalculate()

	var buf bytes.Buffer
	if _, err := f.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo: %v", err)
	}

	_ = data
	_ = spill

	cells := readSheetCells(t, buf.Bytes(), "Spill")

	// Anchor: B2 — first matching value (10).
	assertCellValue(t, cells, "B2", "10")
	// Followers: B3, B4 — additional matched values (30, 50).
	assertCellValue(t, cells, "B3", "30")
	assertCellValue(t, cells, "B4", "50")

	// Followers must carry an empty <f ca="1"/> element so consumers like
	// Apple Numbers recognise them as cached spill output rather than
	// orphaned literal cells.
	assertSpillFollowerFormulaTag(t, cells, "B3")
	assertSpillFollowerFormulaTag(t, cells, "B4")
}

func assertSpillFollowerFormulaTag(t *testing.T, cells map[string]sheetCell, ref string) {
	t.Helper()
	c, ok := cells[ref]
	if !ok {
		t.Fatalf("cell %s missing", ref)
	}
	if !c.HasF {
		t.Fatalf("cell %s missing <f> element; expected empty <f ca=\"1\"/> for spill follower", ref)
	}
	if c.FCa != "1" {
		t.Fatalf("cell %s <f ca=> = %q, want %q", ref, c.FCa, "1")
	}
	if c.FText != "" {
		t.Fatalf("cell %s <f> text = %q, want empty for spill follower", ref, c.FText)
	}
}

// TestSpill_StringFollowerValuesPersisted verifies that string-typed spill
// followers are written with their cached value (resolved via the shared
// string table when applicable).
func TestSpill_StringFollowerValuesPersisted(t *testing.T) {
	f := werkbook.New()
	mustSetSheetName(t, f, "Sheet1", "Data")
	data := f.Sheet("Data")
	spill, err := f.NewSheet("Spill")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	mustSetValue(t, data, "A2", true)
	mustSetValue(t, data, "B2", "alpha")
	mustSetValue(t, data, "A3", false)
	mustSetValue(t, data, "B3", "beta")
	mustSetValue(t, data, "A4", true)
	mustSetValue(t, data, "B4", "gamma")

	mustSetFormula(t, spill, "C2", `FILTER(Data!B2:B4,Data!A2:A4)`)

	f.Recalculate()

	var buf bytes.Buffer
	if _, err := f.WriteTo(&buf); err != nil {
		t.Fatalf("WriteTo: %v", err)
	}

	assertResolvedCellValue(t, buf.Bytes(), "Spill", "C2", "alpha")
	assertResolvedCellValue(t, buf.Bytes(), "Spill", "C3", "gamma")
}

type sheetCell struct {
	Ref     string
	Type    string
	Value   string
	HasF    bool
	FCa     string
	FText   string
}

func readSheetCells(t *testing.T, data []byte, sheetName string) map[string]sheetCell {
	t.Helper()

	type sheetEntry struct {
		Name  string     `xml:"name,attr"`
		Attrs []xml.Attr `xml:",any,attr"`
	}
	type workbookXML struct {
		Sheets []sheetEntry `xml:"sheets>sheet"`
	}
	type relationship struct {
		ID     string `xml:"Id,attr"`
		Target string `xml:"Target,attr"`
	}
	type relationships struct {
		Relationships []relationship `xml:"Relationship"`
	}
	type fXML struct {
		Ca   string `xml:"ca,attr"`
		Text string `xml:",chardata"`
	}
	type cellXML struct {
		Ref   string `xml:"r,attr"`
		Type  string `xml:"t,attr"`
		Value string `xml:"v"`
		F     *fXML  `xml:"f"`
	}
	type rowXML struct {
		Cells []cellXML `xml:"c"`
	}
	type worksheetXML struct {
		Rows []rowXML `xml:"sheetData>row"`
	}

	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("zip open: %v", err)
	}

	read := func(name string) []byte {
		for _, file := range zr.File {
			if file.Name != name {
				continue
			}
			rc, err := file.Open()
			if err != nil {
				t.Fatalf("open %s: %v", name, err)
			}
			defer rc.Close()
			var buf bytes.Buffer
			if _, err := buf.ReadFrom(rc); err != nil {
				t.Fatalf("read %s: %v", name, err)
			}
			return buf.Bytes()
		}
		t.Fatalf("zip entry %q not found", name)
		return nil
	}

	var wb workbookXML
	if err := xml.Unmarshal(read("xl/workbook.xml"), &wb); err != nil {
		t.Fatalf("workbook.xml: %v", err)
	}
	var rels relationships
	if err := xml.Unmarshal(read("xl/_rels/workbook.xml.rels"), &rels); err != nil {
		t.Fatalf("workbook.xml.rels: %v", err)
	}

	target := ""
	for _, sheet := range wb.Sheets {
		if sheet.Name != sheetName {
			continue
		}
		var rid string
		for _, attr := range sheet.Attrs {
			if attr.Name.Local == "id" {
				rid = attr.Value
			}
		}
		for _, rel := range rels.Relationships {
			if rel.ID == rid {
				target = rel.Target
				break
			}
		}
		break
	}
	if target == "" {
		t.Fatalf("sheet %q not found", sheetName)
	}

	target = strings.TrimPrefix(target, "/")
	if !strings.HasPrefix(target, "xl/") {
		target = path.Join("xl", target)
	}

	var sheet worksheetXML
	if err := xml.Unmarshal(read(target), &sheet); err != nil {
		t.Fatalf("worksheet: %v", err)
	}

	out := make(map[string]sheetCell)
	for _, row := range sheet.Rows {
		for _, c := range row.Cells {
			sc := sheetCell{Ref: c.Ref, Type: c.Type, Value: c.Value}
			if c.F != nil {
				sc.HasF = true
				sc.FCa = c.F.Ca
				sc.FText = c.F.Text
			}
			out[c.Ref] = sc
		}
	}
	return out
}

func assertCellValue(t *testing.T, cells map[string]sheetCell, ref, want string) {
	t.Helper()
	c, ok := cells[ref]
	if !ok {
		t.Fatalf("cell %s missing in serialized worksheet", ref)
	}
	if c.Value != want {
		t.Fatalf("cell %s value = %q, want %q", ref, c.Value, want)
	}
}

// assertResolvedCellValue resolves the cell's value through the shared
// string table when applicable, then asserts the resulting display value.
func assertResolvedCellValue(t *testing.T, data []byte, sheetName, ref, want string) {
	t.Helper()

	cells := readSheetCells(t, data, sheetName)
	c, ok := cells[ref]
	if !ok {
		t.Fatalf("cell %s missing in serialized worksheet", ref)
	}

	got := c.Value
	if c.Type == "s" {
		got = lookupSharedString(t, data, c.Value)
	}
	if got != want {
		t.Fatalf("cell %s value = %q, want %q", ref, got, want)
	}
}

func lookupSharedString(t *testing.T, data []byte, idx string) string {
	t.Helper()

	type sstSI struct {
		T string `xml:"t"`
	}
	type sstXML struct {
		SI []sstSI `xml:"si"`
	}

	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("zip open: %v", err)
	}
	for _, file := range zr.File {
		if file.Name != "xl/sharedStrings.xml" {
			continue
		}
		rc, err := file.Open()
		if err != nil {
			t.Fatalf("open sst: %v", err)
		}
		defer rc.Close()
		var buf bytes.Buffer
		if _, err := buf.ReadFrom(rc); err != nil {
			t.Fatalf("read sst: %v", err)
		}
		var sst sstXML
		if err := xml.Unmarshal(buf.Bytes(), &sst); err != nil {
			t.Fatalf("parse sst: %v", err)
		}
		var n int
		for i, b := range idx {
			if b < '0' || b > '9' {
				t.Fatalf("invalid sst index %q", idx)
			}
			n = n*10 + int(b-'0')
			_ = i
		}
		if n < 0 || n >= len(sst.SI) {
			t.Fatalf("sst index %d out of range (len=%d)", n, len(sst.SI))
		}
		return sst.SI[n].T
	}
	t.Fatalf("sharedStrings.xml not found")
	return ""
}
