package werkbook_test

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/jpoz/werkbook"
)

const testXMLHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"

func TestOpenSavePreservesOpaqueOOXMLParts(t *testing.T) {
	dir := t.TempDir()
	srcPath := filepath.Join(dir, "template.xlsx")
	dstPath := filepath.Join(dir, "roundtrip.xlsx")
	writeOpaquePassthroughFixture(t, srcPath)

	f, err := werkbook.Open(srcPath)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	if err := f.Sheet("Sheet1").SetValue("B2", 42); err != nil {
		t.Fatalf("SetValue: %v", err)
	}
	if err := f.SaveAs(dstPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	workbookXML := string(readSheetXML(t, dstPath, "xl/workbook.xml"))
	for _, want := range []string{
		`<bookViews><workbookView activeTab="0"/></bookViews>`,
		`xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`,
		`mc:Ignorable="x15"`,
		`xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"`,
		`<sheet name="Sheet1" sheetId="1" r:id="rId4"></sheet>`,
	} {
		if !strings.Contains(workbookXML, want) {
			t.Fatalf("expected workbook.xml to contain %s: %s", want, workbookXML)
		}
	}

	workbookRels := string(readSheetXML(t, dstPath, "xl/_rels/workbook.xml.rels"))
	for _, want := range []string{
		`Id="rId3"`,
		`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"`,
		`Target="theme/theme1.xml"`,
		`Id="rId4"`,
		`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"`,
		`Target="worksheets/sheet1.xml"`,
		`Id="rId5"`,
		`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"`,
		`Target="styles.xml"`,
	} {
		if !strings.Contains(workbookRels, want) {
			t.Fatalf("expected workbook rels to contain %s: %s", want, workbookRels)
		}
	}

	sheetXML := string(readSheetXML(t, dstPath, "xl/worksheets/sheet1.xml"))
	for _, want := range []string{
		`xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`,
		`mc:Ignorable="x14ac"`,
		`xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"`,
		`<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>`,
		`<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>`,
		`<conditionalFormatting sqref="B2"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>`,
		`<dataValidations count="1"><dataValidation type="whole" sqref="B2"><formula1>1</formula1></dataValidation></dataValidations>`,
		`<hyperlinks><hyperlink ref="B1" r:id="rId4" tooltip="Example"/></hyperlinks>`,
		`<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`,
		`<pageSetup paperSize="9" orientation="portrait"/>`,
		`<headerFooter><oddFooter>&amp;CConfidential</oddFooter></headerFooter>`,
		`<drawing r:id="rId2"/>`,
		`<legacyDrawing r:id="rId3"/>`,
		`<tablePart r:id="rId5"></tablePart>`,
	} {
		if !strings.Contains(sheetXML, want) {
			t.Fatalf("expected sheet XML to contain %s: %s", want, sheetXML)
		}
	}

	sheetRels := string(readSheetXML(t, dstPath, "xl/worksheets/_rels/sheet1.xml.rels"))
	for _, want := range []string{
		`Id="rId2"`,
		`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"`,
		`Target="../drawings/drawing1.xml"`,
		`Id="rId3"`,
		`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"`,
		`Target="../drawings/vmlDrawing1.vml"`,
		`Id="rId4"`,
		`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"`,
		`Target="https://example.com/"`,
		`TargetMode="External"`,
		`Id="rId5"`,
		`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"`,
		`Target="../tables/table1.xml"`,
	} {
		if !strings.Contains(sheetRels, want) {
			t.Fatalf("expected sheet rels to contain %s: %s", want, sheetRels)
		}
	}

	contentTypes := string(readSheetXML(t, dstPath, "[Content_Types].xml"))
	for _, want := range []string{
		`Extension="png"`,
		`ContentType="image/png"`,
		`Extension="vml"`,
		`ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"`,
		`PartName="/xl/drawings/drawing1.xml"`,
		`application/vnd.openxmlformats-officedocument.drawing+xml`,
		`PartName="/xl/charts/chart1.xml"`,
		`application/vnd.openxmlformats-officedocument.drawingml.chart+xml`,
		`PartName="/xl/theme/theme1.xml"`,
		`application/vnd.openxmlformats-officedocument.theme+xml`,
	} {
		if !strings.Contains(contentTypes, want) {
			t.Fatalf("expected content types to contain %s: %s", want, contentTypes)
		}
	}

	for _, entry := range []string{
		"xl/theme/theme1.xml",
		"xl/drawings/drawing1.xml",
		"xl/drawings/_rels/drawing1.xml.rels",
		"xl/charts/chart1.xml",
		"xl/media/image1.png",
		"xl/drawings/vmlDrawing1.vml",
	} {
		if !zipHasEntry(t, dstPath, entry) {
			t.Fatalf("expected %s to survive round-trip", entry)
		}
		if got, want := readSheetXML(t, dstPath, entry), readSheetXML(t, srcPath, entry); !bytes.Equal(got, want) {
			t.Fatalf("%s changed across round-trip\nwant: %q\ngot:  %q", entry, want, got)
		}
	}
}

func writeOpaquePassthroughFixture(t *testing.T, path string) {
	t.Helper()

	out, err := os.Create(path)
	if err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	defer out.Close()

	zw := zip.NewWriter(out)
	writeEntry := func(name, body string) {
		t.Helper()
		w, err := zw.Create(name)
		if err != nil {
			t.Fatalf("create %s: %v", name, err)
		}
		if _, err := w.Write([]byte(body)); err != nil {
			t.Fatalf("write %s: %v", name, err)
		}
	}
	writeBinary := func(name string, body []byte) {
		t.Helper()
		w, err := zw.Create(name)
		if err != nil {
			t.Fatalf("create %s: %v", name, err)
		}
		if _, err := w.Write(body); err != nil {
			t.Fatalf("write %s: %v", name, err)
		}
	}

	writeEntry("[Content_Types].xml", testXMLHeader+`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`)

	writeEntry("_rels/.rels", testXMLHeader+`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`)

	writeEntry("xl/workbook.xml", testXMLHeader+`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" mc:Ignorable="x15">
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`)

	writeEntry("xl/_rels/workbook.xml.rels", testXMLHeader+`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`)

	writeEntry("xl/styles.xml", testXMLHeader+`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`)

	writeEntry("xl/worksheets/sheet1.xml", testXMLHeader+`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
  <sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Header</t></is></c>
      <c r="B1"><v>1</v></c>
    </row>
  </sheetData>
  <conditionalFormatting sqref="B2"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="whole" sqref="B2"><formula1>1</formula1></dataValidation></dataValidations>
  <hyperlinks><hyperlink ref="B1" r:id="rId4" tooltip="Example"/></hyperlinks>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup paperSize="9" orientation="portrait"/>
  <headerFooter><oddFooter>&amp;CConfidential</oddFooter></headerFooter>
  <drawing r:id="rId2"/>
  <legacyDrawing r:id="rId3"/>
  <tableParts count="1"><tablePart r:id="rId1"/></tableParts>
</worksheet>`)

	writeEntry("xl/worksheets/_rels/sheet1.xml.rels", testXMLHeader+`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>
</Relationships>`)

	writeEntry("xl/tables/table1.xml", testXMLHeader+`<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B2" totalsRowShown="0">
  <autoFilter ref="A1:B2"/>
  <tableColumns count="2">
    <tableColumn id="1" name="Header"/>
    <tableColumn id="2" name="Value"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>
</table>`)

	writeEntry("xl/theme/theme1.xml", testXMLHeader+`<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"/>`)

	writeEntry("xl/drawings/drawing1.xml", testXMLHeader+`<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`)

	writeEntry("xl/drawings/_rels/drawing1.xml.rels", testXMLHeader+`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`)

	writeEntry("xl/charts/chart1.xml", testXMLHeader+`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>`)
	writeEntry("xl/drawings/vmlDrawing1.vml", `<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="shape1"/></xml>`)

	png, err := base64.StdEncoding.DecodeString("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a/V0AAAAASUVORK5CYII=")
	if err != nil {
		t.Fatalf("decode png: %v", err)
	}
	writeBinary("xl/media/image1.png", png)

	if err := zw.Close(); err != nil {
		t.Fatalf("close fixture zip: %v", err)
	}
}
