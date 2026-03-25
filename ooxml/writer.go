package ooxml

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strings"
)

const xmlHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"

const (
	// Excel 16.x writes this calcId on future-function workbooks such as
	// <f>_xlfn.ACOT(1)</f> and <f>_xlfn.LET(_xlpm.x,5,_xlpm.x+1)</f>.
	defaultFutureFunctionsCalcID = 181029

	// futureFunctionsWorkbookExtXML is copied from an Excel-authored workbook.
	// Omitting this bundle causes Excel to open the file with a repair prompt
	// even when the individual sheet formulas are serialized correctly.
	futureFunctionsWorkbookExtXML = `<ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:workbookPr chartTrackingRefBase="1"/></ext><ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures"><xcalcf:calcFeatures><xcalcf:feature name="microsoft.com:RD"/><xcalcf:feature name="microsoft.com:Single"/><xcalcf:feature name="microsoft.com:FV"/><xcalcf:feature name="microsoft.com:CNMTM"/><xcalcf:feature name="microsoft.com:LET_WF"/><xcalcf:feature name="microsoft.com:LAMBDA_WF"/><xcalcf:feature name="microsoft.com:ARRAYTEXT_WF"/></xcalcf:calcFeatures></ext>`

	dynamicArrayMetadataXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`
)

type tableWriteInfo struct {
	PartNum int
	Def     TableDef
}

// WriteWorkbook writes a complete XLSX file to w from the given WorkbookData.
func WriteWorkbook(w io.Writer, data *WorkbookData) error {
	zw := zip.NewWriter(w)

	sheetCount := len(data.Sheets)
	if sheetCount == 0 {
		return fmt.Errorf("workbook must have at least one sheet")
	}
	tablePlan := buildTableWritePlan(data.Tables)
	hasCoreProps := len(data.CorePropsRaw) > 0 || hasCorePropertiesData(data.CoreProps)
	hasDynamicArrayMetadata := workbookNeedsDynamicArrayMetadata(data)

	// Build shared string table from all string cells.
	sst := NewSharedStringTable()
	for i := range data.Sheets {
		for j := range data.Sheets[i].Rows {
			for k := range data.Sheets[i].Rows[j].Cells {
				c := &data.Sheets[i].Rows[j].Cells[k]
				if c.Type == "s" && c.Formula == "" {
					idx := sst.Add(c.Value)
					c.Value = fmt.Sprintf("%d", idx)
				}
			}
		}
	}

	// Build style sheet from WorkbookData.Styles.
	ssb := NewStyleSheetBuilder()
	// styleIndexMap maps intermediate StyleData index -> cellXfs index.
	var styleIndexMap []int
	if len(data.Styles) > 0 {
		styleIndexMap = make([]int, len(data.Styles))
		// Index 0 is the default; map it to cellXfs 0.
		styleIndexMap[0] = 0
		for i := 1; i < len(data.Styles); i++ {
			styleIndexMap[i] = ssb.AddStyle(data.Styles[i])
		}
	}

	// [Content_Types].xml
	if err := writeContentTypes(zw, sheetCount, len(data.Tables), sst.Len() > 0, hasCoreProps, hasDynamicArrayMetadata); err != nil {
		return err
	}

	// _rels/.rels
	if err := writeRootRels(zw, hasCoreProps); err != nil {
		return err
	}

	// docProps/app.xml
	if err := writeAppProperties(zw, data); err != nil {
		return err
	}

	// docProps/core.xml
	if hasCoreProps {
		if err := writeCoreProperties(zw, data); err != nil {
			return err
		}
	}

	// xl/workbook.xml
	if err := writeWorkbookXML(zw, data); err != nil {
		return err
	}

	// xl/_rels/workbook.xml.rels
	if err := writeWorkbookRels(zw, sheetCount, sst.Len() > 0, hasDynamicArrayMetadata); err != nil {
		return err
	}

	// xl/styles.xml
	if err := writeXML(zw, "xl/styles.xml", ssb.Build()); err != nil {
		return err
	}

	// xl/worksheets/sheet{N}.xml
	for i, sd := range data.Sheets {
		sheetTables := tablePlan[i]
		if err := writeSheet(zw, i+1, &sd, styleIndexMap, sheetTables); err != nil {
			return err
		}
		if len(sheetTables) > 0 {
			if err := writeSheetRels(zw, i+1, sheetTables); err != nil {
				return err
			}
		}
	}

	if err := writeTables(zw, tablePlan); err != nil {
		return err
	}

	if hasDynamicArrayMetadata {
		if err := writeDynamicArrayMetadata(zw); err != nil {
			return err
		}
	}

	// xl/sharedStrings.xml (only if there are strings)
	if sst.Len() > 0 {
		if err := writeSST(zw, sst); err != nil {
			return err
		}
	}

	return zw.Close()
}

func writeContentTypes(zw *zip.Writer, sheetCount, tableCount int, hasSST, hasCoreProps, hasDynamicArrayMetadata bool) error {
	ct := xlsxTypes{
		Xmlns: contentTypesNS,
		Defaults: []xlsxDefault{
			{Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"},
			{Extension: "xml", ContentType: "application/xml"},
		},
		Overrides: []xlsxOverride{
			{PartName: "/docProps/app.xml", ContentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml"},
			{PartName: "/xl/workbook.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"},
			{PartName: "/xl/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"},
		},
	}
	if hasCoreProps {
		ct.Overrides = append(ct.Overrides, xlsxOverride{
			PartName:    "/docProps/core.xml",
			ContentType: "application/vnd.openxmlformats-package.core-properties+xml",
		})
	}
	for i := range sheetCount {
		ct.Overrides = append(ct.Overrides, xlsxOverride{
			PartName:    fmt.Sprintf("/xl/worksheets/sheet%d.xml", i+1),
			ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
		})
	}
	if hasSST {
		ct.Overrides = append(ct.Overrides, xlsxOverride{
			PartName:    "/xl/sharedStrings.xml",
			ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
		})
	}
	for i := range tableCount {
		ct.Overrides = append(ct.Overrides, xlsxOverride{
			PartName:    fmt.Sprintf("/xl/tables/table%d.xml", i+1),
			ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
		})
	}
	if hasDynamicArrayMetadata {
		ct.Overrides = append(ct.Overrides, xlsxOverride{
			PartName:    "/xl/metadata.xml",
			ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
		})
	}
	return writeXML(zw, "[Content_Types].xml", ct)
}

func writeRootRels(zw *zip.Writer, hasCoreProps bool) error {
	rels := xlsxRelationships{
		Xmlns: NSRelationships,
		Relationships: []xlsxRelationship{
			{ID: "rId1", Type: RelTypeWorkbook, Target: "xl/workbook.xml"},
		},
	}
	nextID := 2
	if hasCoreProps {
		rels.Relationships = append(rels.Relationships, xlsxRelationship{
			ID:     fmt.Sprintf("rId%d", nextID),
			Type:   RelTypeCoreProps,
			Target: "docProps/core.xml",
		})
		nextID++
	}
	rels.Relationships = append(rels.Relationships, xlsxRelationship{
		ID:     fmt.Sprintf("rId%d", nextID),
		Type:   RelTypeExtendedApp,
		Target: "docProps/app.xml",
	})
	return writeXML(zw, "_rels/.rels", rels)
}

func writeWorkbookXML(zw *zip.Writer, data *WorkbookData) error {
	wb := xlsxWorkbook{
		Xmlns:  NSSpreadsheetML,
		XmlnsR: NSOfficeDocument,
	}

	// Future-function formulas are serialized with _xlfn.* prefixes in sheet XML.
	// Excel expects the workbook part to advertise the matching calc metadata,
	// otherwise it offers to repair the file on open.
	calcProps := data.CalcProps
	if workbookNeedsFutureFunctionsMetadata(data) && calcProps.ID == 0 {
		calcProps.ID = defaultFutureFunctionsCalcID
	}
	if data.Date1904 {
		wb.WorkbookPr = &xlsxWorkbookPr{Date1904: "1"}
	}
	if hasCalcProps(calcProps) {
		wb.CalcPr = &xlsxCalcPr{
			CalcMode:       calcProps.Mode,
			CalcID:         calcProps.ID,
			FullCalcOnLoad: boolString(calcProps.FullCalcOnLoad),
			ForceFullCalc:  boolString(calcProps.ForceFullCalc),
			CalcCompleted:  boolString(calcProps.Completed),
		}
	}
	for i, sd := range data.Sheets {
		sheet := xlsxSheet{
			Name:    sd.Name,
			SheetID: i + 1,
			RID:     fmt.Sprintf("rId%d", i+1),
		}
		if sd.State != "" {
			sheet.State = sd.State
		}
		wb.Sheets.Sheet = append(wb.Sheets.Sheet, sheet)
	}
	if len(data.DefinedNames) > 0 {
		wb.DefinedNames = &xlsxDefinedNames{}
		for _, dn := range data.DefinedNames {
			// Skip defined names that reference external workbooks (e.g.
			// '[1]Sheet'!$A$1). werkbook does not preserve external link
			// metadata, so writing these produces dangling references that
			// Excel flags as corrupt.
			if isExternalRef(dn.Value) {
				continue
			}
			xdn := xlsxDefinedName{
				Name:  dn.Name,
				Value: dn.Value,
			}
			if dn.LocalSheetID >= 0 {
				id := dn.LocalSheetID
				xdn.LocalSheetID = &id
			}
			wb.DefinedNames.DefinedName = append(wb.DefinedNames.DefinedName, xdn)
		}
		if len(wb.DefinedNames.DefinedName) == 0 {
			wb.DefinedNames = nil
		}
	}
	if workbookNeedsFutureFunctionsMetadata(data) {
		wb.ExtLst = &xlsxExtLst{InnerXML: futureFunctionsWorkbookExtXML}
	}
	return writeXML(zw, "xl/workbook.xml", wb)
}

// isExternalRef returns true if a defined name value references an external
// workbook. External refs start with [n] or '[n] (e.g. "[1]Sheet!$A$1" or
// "'[1]Sheet'!$A$1").
func isExternalRef(value string) bool {
	return strings.HasPrefix(value, "[") || strings.HasPrefix(value, "'[")
}

func writeWorkbookRels(zw *zip.Writer, sheetCount int, hasSST, hasDynamicArrayMetadata bool) error {
	rels := xlsxRelationships{
		Xmlns: NSRelationships,
	}
	for i := range sheetCount {
		rels.Relationships = append(rels.Relationships, xlsxRelationship{
			ID:     fmt.Sprintf("rId%d", i+1),
			Type:   RelTypeWorksheet,
			Target: fmt.Sprintf("worksheets/sheet%d.xml", i+1),
		})
	}
	nextID := sheetCount + 1
	rels.Relationships = append(rels.Relationships, xlsxRelationship{
		ID:     fmt.Sprintf("rId%d", nextID),
		Type:   RelTypeStyles,
		Target: "styles.xml",
	})
	if hasSST {
		nextID++
		rels.Relationships = append(rels.Relationships, xlsxRelationship{
			ID:     fmt.Sprintf("rId%d", nextID),
			Type:   RelTypeSharedStr,
			Target: "sharedStrings.xml",
		})
	}
	if hasDynamicArrayMetadata {
		nextID++
		rels.Relationships = append(rels.Relationships, xlsxRelationship{
			ID:     fmt.Sprintf("rId%d", nextID),
			Type:   RelTypeSheetMetadata,
			Target: "metadata.xml",
		})
	}
	return writeXML(zw, "xl/_rels/workbook.xml.rels", rels)
}

func writeSheet(zw *zip.Writer, num int, sd *SheetData, styleIndexMap []int, tables []tableWriteInfo) error {
	ws := xlsxWorksheet{
		Xmlns: NSSpreadsheetML,
	}
	if len(tables) > 0 {
		ws.XmlnsR = NSOfficeDocument
		ws.TableParts = &xlsxTableParts{Count: len(tables)}
		for i := range tables {
			ws.TableParts.TablePart = append(ws.TableParts.TablePart, xlsxTablePart{
				RID: fmt.Sprintf("rId%d", i+1),
			})
		}
	}

	// Populate column widths.
	if len(sd.ColWidths) > 0 {
		ws.Cols = &xlsxCols{}
		for _, cw := range sd.ColWidths {
			ws.Cols.Col = append(ws.Cols.Col, xlsxCol{
				Min:         cw.Min,
				Max:         cw.Max,
				Width:       cw.Width,
				CustomWidth: 1,
			})
		}
	}
	if len(sd.MergeCells) > 0 {
		ws.MergeCells = &xlsxMergeCells{Count: len(sd.MergeCells)}
		for _, mc := range sd.MergeCells {
			ref := mc.StartAxis
			if mc.EndAxis != "" && mc.EndAxis != mc.StartAxis {
				ref += ":" + mc.EndAxis
			}
			ws.MergeCells.MergeCell = append(ws.MergeCells.MergeCell, xlsxMergeCell{Ref: ref})
		}
	}

	for _, rd := range sd.Rows {
		var hidden ooxmlBool
		if rd.Hidden {
			hidden = 1
		}
		row := xlsxRow{R: rd.Num, Hidden: hidden}
		if rd.Height != 0 {
			row.Ht = rd.Height
			row.CustomHeight = 1
		}
		for _, cd := range rd.Cells {
			c := xlsxC{
				R: cd.Ref,
				T: cd.Type,
				V: cd.Value,
			}
			if cd.Formula != "" {
				c.FE = &xlsxF{Text: cd.Formula}
				if cd.IsDynamicArray && cd.FormulaRef != "" {
					c.CM = 1
					c.FE.T = cd.FormulaType
					if c.FE.T == "" {
						c.FE.T = "array"
					}
					c.FE.Ref = cd.FormulaRef
					c.FE.Aca = 1
					c.FE.Ca = 1
				} else if !cd.IsDynamicArray {
					c.FE.T = cd.FormulaType
					c.FE.Ref = cd.FormulaRef
				}
			}
			if cd.StyleIdx > 0 && styleIndexMap != nil && cd.StyleIdx < len(styleIndexMap) {
				c.S = styleIndexMap[cd.StyleIdx]
			}
			row.Cells = append(row.Cells, c)
		}
		ws.SheetData.Rows = append(ws.SheetData.Rows, row)
	}
	return writeXML(zw, fmt.Sprintf("xl/worksheets/sheet%d.xml", num), ws)
}

func writeSheetRels(zw *zip.Writer, sheetNum int, tables []tableWriteInfo) error {
	rels := xlsxRelationships{Xmlns: NSRelationships}
	for i, tw := range tables {
		rels.Relationships = append(rels.Relationships, xlsxRelationship{
			ID:     fmt.Sprintf("rId%d", i+1),
			Type:   RelTypeTable,
			Target: fmt.Sprintf("../tables/table%d.xml", tw.PartNum),
		})
	}
	return writeXML(zw, fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", sheetNum), rels)
}

func writeTables(zw *zip.Writer, tablePlan map[int][]tableWriteInfo) error {
	for _, sheetTables := range tablePlan {
		for _, tw := range sheetTables {
			if err := writeTableXML(zw, tw.PartNum, tw.Def); err != nil {
				return err
			}
		}
	}
	return nil
}

func buildTableWritePlan(tables []TableDef) map[int][]tableWriteInfo {
	plan := make(map[int][]tableWriteInfo)
	for i, td := range tables {
		plan[td.SheetIndex] = append(plan[td.SheetIndex], tableWriteInfo{
			PartNum: i + 1,
			Def:     td,
		})
	}
	return plan
}

func hasCalcProps(props CalcPropertiesData) bool {
	return props.Mode != "" || props.ID != 0 || props.FullCalcOnLoad || props.ForceFullCalc || props.Completed
}

func boolString(v bool) string {
	if v {
		return "1"
	}
	return ""
}

// workbookNeedsFutureFunctionsMetadata reports whether any serialized formula
// already contains _xlfn. prefixes and therefore needs the workbook-level calc
// feature bundle. The check runs after cell serialization, so formulas such as
// LET(x,5,x+1) are seen here as _xlfn.LET(_xlpm.x,5,_xlpm.x+1).
func workbookNeedsFutureFunctionsMetadata(data *WorkbookData) bool {
	for _, sheet := range data.Sheets {
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				if strings.Contains(strings.ToUpper(cell.Formula), "_XLFN.") {
					return true
				}
			}
		}
	}
	return false
}

func workbookNeedsDynamicArrayMetadata(data *WorkbookData) bool {
	for _, sheet := range data.Sheets {
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				if cell.IsDynamicArray && cell.FormulaRef != "" {
					return true
				}
			}
		}
	}
	return false
}

func writeDynamicArrayMetadata(zw *zip.Writer) error {
	w, err := zw.Create("xl/metadata.xml")
	if err != nil {
		return fmt.Errorf("create xl/metadata.xml: %w", err)
	}
	if _, err := io.WriteString(w, dynamicArrayMetadataXML); err != nil {
		return fmt.Errorf("write xl/metadata.xml: %w", err)
	}
	return nil
}

func writeSST(zw *zip.Writer, sst *SharedStringTable) error {
	return writeXML(zw, "xl/sharedStrings.xml", sst.ToXML())
}

func writeXML(zw *zip.Writer, name string, v any) error {
	w, err := zw.Create(name)
	if err != nil {
		return fmt.Errorf("create %s: %w", name, err)
	}
	if _, err := io.WriteString(w, xmlHeader); err != nil {
		return err
	}
	enc := xml.NewEncoder(w)
	if err := enc.Encode(v); err != nil {
		return fmt.Errorf("encode %s: %w", name, err)
	}
	return enc.Close()
}
