package werkbook_test

import (
	"archive/zip"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/jpoz/werkbook"
)

func TestRecalculateUsesExternalLinkCache(t *testing.T) {
	f := werkbook.New()
	s := f.Sheet("Sheet1")

	if err := s.SetValue("A1", "beta"); err != nil {
		t.Fatalf("SetValue A1: %v", err)
	}
	if err := s.SetFormula("B1", `VLOOKUP(A1,'[1]Lookup'!$A$1:$B$2,2,FALSE)`); err != nil {
		t.Fatalf("SetFormula B1: %v", err)
	}
	if err := s.SetFormula("C1", "'[1]Lookup'!B1"); err != nil {
		t.Fatalf("SetFormula C1: %v", err)
	}

	dir := t.TempDir()
	basePath := filepath.Join(dir, "base.xlsx")
	if err := f.SaveAs(basePath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	patchedPath := filepath.Join(dir, "external-cache.xlsx")
	addExternalLinkCache(t, basePath, patchedPath)

	f2, err := werkbook.Open(patchedPath)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	f2.Recalculate()

	s2 := f2.Sheet("Sheet1")
	got, err := s2.GetValue("B1")
	if err != nil {
		t.Fatalf("GetValue B1: %v", err)
	}
	if got.Type != werkbook.TypeNumber || got.Number != 22 {
		t.Fatalf("B1 = %#v, want 22", got)
	}

	got, err = s2.GetValue("C1")
	if err != nil {
		t.Fatalf("GetValue C1: %v", err)
	}
	if got.Type != werkbook.TypeNumber || got.Number != 11 {
		t.Fatalf("C1 = %#v, want 11", got)
	}
}

func TestSheetQualifiedRefErrorLiteralEvaluatesToRef(t *testing.T) {
	f := werkbook.New(werkbook.FirstSheet("Main"))
	if _, err := f.NewSheet("Lookup Sheet"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	s := f.Sheet("Main")
	if err := s.SetFormula("A1", "'Lookup Sheet'!#REF!"); err != nil {
		t.Fatalf("SetFormula: %v", err)
	}

	got, err := s.GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue: %v", err)
	}
	if got.Type != werkbook.TypeError || got.String != "#REF!" {
		t.Fatalf("A1 = %#v, want #REF!", got)
	}
}

func addExternalLinkCache(t *testing.T, srcPath, dstPath string) {
	t.Helper()

	zr, err := zip.OpenReader(srcPath)
	if err != nil {
		t.Fatalf("open source xlsx: %v", err)
	}
	defer zr.Close()

	out, err := os.Create(dstPath)
	if err != nil {
		t.Fatalf("create destination xlsx: %v", err)
	}
	defer out.Close()

	zw := zip.NewWriter(out)
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open zip entry %s: %v", f.Name, err)
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatalf("read zip entry %s: %v", f.Name, err)
		}

		switch f.Name {
		case "xl/workbook.xml":
			xml := string(data)
			xml = strings.Replace(xml, "</workbook>", `<externalReferences><externalReference r:id="rId99"/></externalReferences></workbook>`, 1)
			data = []byte(xml)
		case "xl/_rels/workbook.xml.rels":
			xml := string(data)
			xml = strings.Replace(xml, "</Relationships>", `<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>`, 1)
			data = []byte(xml)
		}

		w, err := zw.Create(f.Name)
		if err != nil {
			t.Fatalf("create zip entry %s: %v", f.Name, err)
		}
		if _, err := w.Write(data); err != nil {
			t.Fatalf("write zip entry %s: %v", f.Name, err)
		}
	}

	w, err := zw.Create("xl/externalLinks/externalLink1.xml")
	if err != nil {
		t.Fatalf("create externalLink1.xml: %v", err)
	}
	if _, err := io.WriteString(w, externalLinkFixtureXML); err != nil {
		t.Fatalf("write externalLink1.xml: %v", err)
	}

	if err := zw.Close(); err != nil {
		t.Fatalf("close destination zip: %v", err)
	}
}

const externalLinkFixtureXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <externalBook>
    <sheetNames>
      <sheetName val="Lookup"/>
    </sheetNames>
    <sheetDataSet>
      <sheetData sheetId="0">
        <row r="1">
          <cell r="A1" t="str"><v>alpha</v></cell>
          <cell r="B1"><v>11</v></cell>
        </row>
        <row r="2">
          <cell r="A2" t="str"><v>beta</v></cell>
          <cell r="B2"><v>22</v></cell>
        </row>
      </sheetData>
    </sheetDataSet>
  </externalBook>
</externalLink>
`
