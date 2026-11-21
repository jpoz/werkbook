package ooxml

import (
	"archive/zip"
	"bytes"
	"io"
	"strings"
	"testing"
)

func TestSelfCloseEmptyElements(t *testing.T) {
	cases := []struct{ in, want string }{
		{`<mergeCell ref="A1:K1"></mergeCell>`, `<mergeCell ref="A1:K1"/>`},
		{`<c r="B3" s="2"></c>`, `<c r="B3" s="2"/>`},
		{`<t>text</t>`, `<t>text</t>`},
		{`<v>1&gt;2</v>`, `<v>1&gt;2</v>`},
		{`<t></t>`, `<t/>`},
		{`<row r="1"><c r="A1"></c></row>`, `<row r="1"><c r="A1"/></row>`},
	}
	for _, tc := range cases {
		if got := string(selfCloseEmptyElements([]byte(tc.in))); got != tc.want {
			t.Errorf("selfCloseEmptyElements(%q) = %q, want %q", tc.in, got, tc.want)
		}
	}
}

func TestWriteWorkbookSelfClosingMergeCells(t *testing.T) {
	data := &WorkbookData{
		Sheets: []SheetData{{
			Name:       "Sheet1",
			Rows:       []RowData{{Num: 1, Cells: []CellData{{Ref: "A1", Type: "s", Value: "hello"}}}},
			MergeCells: []MergeCellData{{StartAxis: "A1", EndAxis: "K1"}},
		}},
	}
	var buf bytes.Buffer
	if err := WriteWorkbook(&buf, data); err != nil {
		t.Fatalf("WriteWorkbook: %v", err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}
	for _, f := range zr.File {
		if f.Name != "xl/worksheets/sheet1.xml" {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open sheet: %v", err)
		}
		b, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatalf("read sheet: %v", err)
		}
		xml := string(b)
		if !strings.Contains(xml, `<mergeCell ref="A1:K1"/>`) {
			t.Errorf("expected self-closing mergeCell, got: %s", xml)
		}
		if strings.Contains(xml, "</mergeCell>") {
			t.Errorf("found non-self-closing mergeCell: %s", xml)
		}
		return
	}
	t.Fatal("sheet1.xml not found")
}
