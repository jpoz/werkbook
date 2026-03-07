package ooxml

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
)

const xmlHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"

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

	// Build shared string table from all string cells.
	sst := NewSharedStringTable()
	hasDynamicArrays := false
	for i := range data.Sheets {
		for j := range data.Sheets[i].Rows {
			for k := range data.Sheets[i].Rows[j].Cells {
				c := &data.Sheets[i].Rows[j].Cells[k]
				if c.IsDynamicArray {
					hasDynamicArrays = true
				}
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
	if err := writeContentTypes(zw, sheetCount, len(data.Tables), sst.Len() > 0, hasDynamicArrays); err != nil {
		return err
	}

	// _rels/.rels
	if err := writeRootRels(zw); err != nil {
		return err
	}

	// xl/workbook.xml
	if err := writeWorkbookXML(zw, data); err != nil {
		return err
	}

	// xl/_rels/workbook.xml.rels
	if err := writeWorkbookRels(zw, sheetCount, sst.Len() > 0, hasDynamicArrays); err != nil {
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

	// xl/sharedStrings.xml (only if there are strings)
	if sst.Len() > 0 {
		if err := writeSST(zw, sst); err != nil {
			return err
		}
	}

	if hasDynamicArrays {
		if err := writeDynamicArrayMetadata(zw); err != nil {
			return err
		}
	}

	return zw.Close()
}

func writeContentTypes(zw *zip.Writer, sheetCount, tableCount int, hasSST, hasDynamicArrays bool) error {
	ct := xlsxTypes{
		Xmlns: contentTypesNS,
		Defaults: []xlsxDefault{
			{Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml"},
			{Extension: "xml", ContentType: "application/xml"},
		},
		Overrides: []xlsxOverride{
			{PartName: "/xl/workbook.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"},
			{PartName: "/xl/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"},
		},
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
	if hasDynamicArrays {
		ct.Overrides = append(ct.Overrides, xlsxOverride{
			PartName:    "/xl/metadata.xml",
			ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
		})
	}
	for i := range tableCount {
		ct.Overrides = append(ct.Overrides, xlsxOverride{
			PartName:    fmt.Sprintf("/xl/tables/table%d.xml", i+1),
			ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
		})
	}
	return writeXML(zw, "[Content_Types].xml", ct)
}

func writeRootRels(zw *zip.Writer) error {
	rels := xlsxRelationships{
		Xmlns: NSRelationships,
		Relationships: []xlsxRelationship{
			{ID: "rId1", Type: RelTypeWorkbook, Target: "xl/workbook.xml"},
		},
	}
	return writeXML(zw, "_rels/.rels", rels)
}

func writeWorkbookXML(zw *zip.Writer, data *WorkbookData) error {
	wb := xlsxWorkbook{
		Xmlns:  NSSpreadsheetML,
		XmlnsR: NSOfficeDocument,
	}
	if data.Date1904 {
		wb.WorkbookPr = &xlsxWorkbookPr{Date1904: "1"}
	}
	if hasCalcProps(data.CalcProps) {
		wb.CalcPr = &xlsxCalcPr{
			CalcMode:       data.CalcProps.Mode,
			CalcID:         data.CalcProps.ID,
			FullCalcOnLoad: boolString(data.CalcProps.FullCalcOnLoad),
			ForceFullCalc:  boolString(data.CalcProps.ForceFullCalc),
			CalcCompleted:  boolString(data.CalcProps.Completed),
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
	}
	return writeXML(zw, "xl/workbook.xml", wb)
}

func writeWorkbookRels(zw *zip.Writer, sheetCount int, hasSST, hasDynamicArrays bool) error {
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
	if hasDynamicArrays {
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
				c.FE = &xlsxF{
					T:    cd.FormulaType,
					Ref:  cd.FormulaRef,
					Text: cd.Formula,
				}
			}
			if cd.IsDynamicArray {
				c.CM = 1
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
