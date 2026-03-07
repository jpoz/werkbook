package werkbook_test

import (
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/jpoz/werkbook"
)

func TestDate1904RoundTrip(t *testing.T) {
	f := werkbook.New(werkbook.WithDate1904(true))
	s := f.Sheet("Sheet1")
	if err := s.SetValue("A1", time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)); err != nil {
		t.Fatalf("SetValue: %v", err)
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "date1904.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	workbookXML := string(readSheetXML(t, path, "xl/workbook.xml"))
	if !strings.Contains(workbookXML, `date1904="1"`) {
		t.Fatalf("expected workbookPr date1904 in workbook.xml: %s", workbookXML)
	}
	sheetXML := string(readSheetXML(t, path, "xl/worksheets/sheet1.xml"))
	if !strings.Contains(sheetXML, `<v>0</v>`) {
		t.Fatalf("expected 1904-01-01 to serialize as serial 0, XML: %s", sheetXML)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	if !f2.Date1904() {
		t.Fatal("expected reopened workbook to preserve date1904")
	}
	v, err := f2.Sheet("Sheet1").GetValue("A1")
	if err != nil {
		t.Fatalf("GetValue: %v", err)
	}
	if v.Type != werkbook.TypeNumber || v.Number != 0 {
		t.Fatalf("A1 = %#v, want serial 0", v)
	}
}

func TestCalcPropertiesRoundTrip(t *testing.T) {
	f := werkbook.New()
	f.SetCalcProperties(werkbook.CalcProperties{
		Mode:           "manual",
		ID:             187000,
		FullCalcOnLoad: true,
		ForceFullCalc:  true,
		Completed:      true,
	})

	dir := t.TempDir()
	path := filepath.Join(dir, "calc-props.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	workbookXML := string(readSheetXML(t, path, "xl/workbook.xml"))
	for _, want := range []string{
		`calcMode="manual"`,
		`calcId="187000"`,
		`fullCalcOnLoad="1"`,
		`forceFullCalc="1"`,
		`calcCompleted="1"`,
	} {
		if !strings.Contains(workbookXML, want) {
			t.Fatalf("expected workbook.xml to contain %s: %s", want, workbookXML)
		}
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	got := f2.CalcProperties()
	if got != (werkbook.CalcProperties{
		Mode:           "manual",
		ID:             187000,
		FullCalcOnLoad: true,
		ForceFullCalc:  true,
		Completed:      true,
	}) {
		t.Fatalf("CalcProperties = %#v", got)
	}
}

func TestTableAuthoringRoundTrip(t *testing.T) {
	f := werkbook.New(werkbook.FirstSheet("Data"))
	s := f.Sheet("Data")

	for ref, val := range map[string]any{
		"A1": "Name",
		"B1": "Price",
		"C1": "Qty",
		"A2": "Apple",
		"B2": 1.5,
		"C2": 10,
		"A3": "Banana",
		"B3": 0.75,
		"C3": 20,
		"A4": "Cherry",
		"B4": 3.0,
		"C4": 5,
	} {
		if err := s.SetValue(ref, val); err != nil {
			t.Fatalf("SetValue(%s): %v", ref, err)
		}
	}
	if err := s.AddTable(werkbook.Table{
		DisplayName: "Products",
		Ref:         "A1:C4",
		Style: &werkbook.TableStyle{
			Name:           "TableStyleMedium2",
			ShowRowStripes: true,
		},
	}); err != nil {
		t.Fatalf("AddTable: %v", err)
	}
	if err := s.SetFormula("D1", "SUM(Products[Price])"); err != nil {
		t.Fatalf("SetFormula: %v", err)
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "table.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	tableXML := string(readSheetXML(t, path, "xl/tables/table1.xml"))
	for _, want := range []string{
		`displayName="Products"`,
		`ref="A1:C4"`,
		`tableStyleInfo`,
		`name="Price"`,
	} {
		if !strings.Contains(tableXML, want) {
			t.Fatalf("expected table XML to contain %s: %s", want, tableXML)
		}
	}
	sheetXML := string(readSheetXML(t, path, "xl/worksheets/sheet1.xml"))
	if !strings.Contains(sheetXML, `<tableParts count="1">`) || !strings.Contains(sheetXML, `<tablePart r:id="rId1"></tablePart>`) {
		t.Fatalf("expected worksheet to reference table part: %s", sheetXML)
	}
	sheetRels := string(readSheetXML(t, path, "xl/worksheets/_rels/sheet1.xml.rels"))
	if !strings.Contains(sheetRels, `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"`) {
		t.Fatalf("expected sheet relationships to include table rel: %s", sheetRels)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	tables := f2.Tables()
	if len(tables) != 1 {
		t.Fatalf("expected 1 table, got %d", len(tables))
	}
	if tables[0].DisplayName != "Products" || tables[0].Ref != "A1:C4" {
		t.Fatalf("unexpected table metadata: %#v", tables[0])
	}
	v, err := f2.Sheet("Data").GetValue("D1")
	if err != nil {
		t.Fatalf("GetValue(D1): %v", err)
	}
	if v.Type != werkbook.TypeNumber || v.Number != 5.25 {
		t.Fatalf("D1 = %#v, want 5.25", v)
	}
}
