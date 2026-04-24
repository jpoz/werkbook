package ooxml

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
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

	hasSST := sst.Len() > 0
	sheetRelIDs := nextRelationshipIDs(data.ExtraRels, len(data.Sheets))
	emittedPaths := buildWriterPaths(data, tablePlan, hasSST, hasCoreProps, hasDynamicArrayMetadata)

	// [Content_Types].xml
	if err := writeContentTypes(zw, data, tablePlan, emittedPaths, hasSST, hasCoreProps, hasDynamicArrayMetadata); err != nil {
		return err
	}

	// _rels/.rels
	if err := writeRootRels(zw, data.ExtraRootRels, hasCoreProps); err != nil {
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
	if err := writeWorkbookXML(zw, data, sheetRelIDs); err != nil {
		return err
	}

	// xl/_rels/workbook.xml.rels
	if err := writeWorkbookRels(zw, data, sheetRelIDs, hasSST, hasDynamicArrayMetadata); err != nil {
		return err
	}

	// xl/styles.xml
	if err := writeXML(zw, "xl/styles.xml", ssb.Build()); err != nil {
		return err
	}

	// xl/worksheets/sheet{N}.xml
	for i, sd := range data.Sheets {
		sheetTables := tablePlan[i]
		tableRelIDs := nextRelationshipIDs(sd.ExtraRels, len(sheetTables))
		if err := writeSheet(zw, i+1, &sd, styleIndexMap, sheetTables, tableRelIDs); err != nil {
			return err
		}
		if len(sheetTables) > 0 || len(sd.ExtraRels) > 0 {
			if err := writeSheetRels(zw, &sd, i+1, sheetTables, tableRelIDs); err != nil {
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
	if hasSST {
		if err := writeSST(zw, sst); err != nil {
			return err
		}
	}

	if err := writeOpaqueEntries(zw, data.OpaqueEntries, emittedPaths); err != nil {
		return err
	}

	return zw.Close()
}

func writeContentTypes(
	zw *zip.Writer,
	data *WorkbookData,
	tablePlan map[int][]tableWriteInfo,
	emittedPaths map[string]struct{},
	hasSST, hasCoreProps, hasDynamicArrayMetadata bool,
) error {
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
	for i := range data.Sheets {
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
	for _, sheetTables := range tablePlan {
		for _, tw := range sheetTables {
			ct.Overrides = append(ct.Overrides, xlsxOverride{
				PartName:    fmt.Sprintf("/xl/tables/table%d.xml", tw.PartNum),
				ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
			})
		}
	}
	if hasDynamicArrayMetadata {
		ct.Overrides = append(ct.Overrides, xlsxOverride{
			PartName:    "/xl/metadata.xml",
			ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
		})
	}

	defaults := map[string]string{
		"rels": "application/vnd.openxmlformats-package.relationships+xml",
		"xml":  "application/xml",
	}
	for _, def := range data.OpaqueDefaults {
		ext := strings.ToLower(def.Extension)
		if ext == "xml" || ext == "rels" {
			continue
		}
		if _, ok := defaults[ext]; ok {
			continue
		}
		defaults[ext] = def.ContentType
		ct.Defaults = append(ct.Defaults, xlsxDefault{
			Extension:   def.Extension,
			ContentType: def.ContentType,
		})
	}

	seenOverrides := make(map[string]struct{}, len(ct.Overrides))
	for _, override := range ct.Overrides {
		seenOverrides[strings.ToLower(override.PartName)] = struct{}{}
	}
	for _, entry := range data.OpaqueEntries {
		if entry.ContentType == "" {
			continue
		}
		if _, known := emittedPaths[entry.Path]; known {
			continue
		}
		partName := "/" + entry.Path
		key := strings.ToLower(partName)
		if _, ok := seenOverrides[key]; ok {
			continue
		}
		seenOverrides[key] = struct{}{}
		ct.Overrides = append(ct.Overrides, xlsxOverride{
			PartName:    partName,
			ContentType: entry.ContentType,
		})
	}

	sort.SliceStable(ct.Defaults, func(i, j int) bool {
		return strings.ToLower(ct.Defaults[i].Extension) < strings.ToLower(ct.Defaults[j].Extension)
	})
	sort.SliceStable(ct.Overrides, func(i, j int) bool {
		return strings.ToLower(ct.Overrides[i].PartName) < strings.ToLower(ct.Overrides[j].PartName)
	})
	return writeXML(zw, "[Content_Types].xml", ct)
}

func writeRootRels(zw *zip.Writer, extra []OpaqueRel, hasCoreProps bool) error {
	rels := xlsxRelationships{
		Xmlns: NSRelationships,
	}
	for _, rel := range extra {
		rels.Relationships = append(rels.Relationships, xlsxRelationship{
			ID:         rel.ID,
			Type:       rel.Type,
			Target:     rel.Target,
			TargetMode: rel.TargetMode,
		})
	}
	ids := nextRelationshipIDs(extra, 1+boolInt(hasCoreProps)+1)
	idx := 0
	rels.Relationships = append(rels.Relationships, xlsxRelationship{
		ID:     ids[idx],
		Type:   RelTypeWorkbook,
		Target: "xl/workbook.xml",
	})
	idx++
	if hasCoreProps {
		rels.Relationships = append(rels.Relationships, xlsxRelationship{
			ID:     ids[idx],
			Type:   RelTypeCoreProps,
			Target: "docProps/core.xml",
		})
		idx++
	}
	rels.Relationships = append(rels.Relationships, xlsxRelationship{
		ID:     ids[idx],
		Type:   RelTypeExtendedApp,
		Target: "docProps/app.xml",
	})
	return writeXML(zw, "_rels/.rels", rels)
}

func writeWorkbookXML(zw *zip.Writer, data *WorkbookData, sheetRelIDs []string) error {
	// Future-function formulas are serialized with _xlfn.* prefixes in sheet XML.
	// Excel expects the workbook part to advertise the matching calc metadata,
	// otherwise it offers to repair the file on open.
	calcProps := data.CalcProps
	if workbookNeedsFutureFunctionsMetadata(data) && calcProps.ID == 0 {
		calcProps.ID = defaultFutureFunctionsCalcID
	}
	requiredAttrs := map[string]string{
		"xmlns":   spreadsheetNamespace(data.RootAttrs),
		"xmlns:r": relationshipsNamespace(data.RootAttrs),
	}
	rootAttrs := mergedRootAttrs(data.RootAttrs, requiredAttrs)

	var fragments []xmlFragment
	if data.Date1904 {
		frag, err := encodeNamedXMLFragment("workbookPr", &xlsxWorkbookPr{Date1904: "1"})
		if err != nil {
			return fmt.Errorf("encode workbookPr: %w", err)
		}
		fragments = append(fragments, xmlFragment{
			OrderKey: workbookElementOrder["workbookPr"],
			Seq:      writerSeq,
			XML:      frag,
		})
	}

	sheets := xlsxSheets{}
	for i, sd := range data.Sheets {
		rid := fmt.Sprintf("rId%d", i+1)
		if i < len(sheetRelIDs) {
			rid = sheetRelIDs[i]
		}
		sheet := xlsxSheet{
			Name:    sd.Name,
			SheetID: i + 1,
			RID:     rid,
		}
		if sd.State != "" {
			sheet.State = sd.State
		}
		sheets.Sheet = append(sheets.Sheet, sheet)
	}
	sheetsFrag, err := encodeNamedXMLFragment("sheets", sheets)
	if err != nil {
		return fmt.Errorf("encode sheets: %w", err)
	}
	fragments = append(fragments, xmlFragment{
		OrderKey: workbookElementOrder["sheets"],
		Seq:      writerSeq,
		XML:      sheetsFrag,
	})

	if hasCalcProps(calcProps) {
		calcFrag, err := encodeNamedXMLFragment("calcPr", &xlsxCalcPr{
			CalcMode:       calcProps.Mode,
			CalcID:         calcProps.ID,
			FullCalcOnLoad: boolString(calcProps.FullCalcOnLoad),
			ForceFullCalc:  boolString(calcProps.ForceFullCalc),
			CalcCompleted:  boolString(calcProps.Completed),
		})
		if err != nil {
			return fmt.Errorf("encode calcPr: %w", err)
		}
		fragments = append(fragments, xmlFragment{
			OrderKey: workbookElementOrder["calcPr"],
			Seq:      writerSeq,
			XML:      calcFrag,
		})
	}

	if len(data.DefinedNames) > 0 {
		defs := &xlsxDefinedNames{}
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
			defs.DefinedName = append(defs.DefinedName, xdn)
		}
		if len(defs.DefinedName) > 0 {
			defsFrag, err := encodeNamedXMLFragment("definedNames", defs)
			if err != nil {
				return fmt.Errorf("encode definedNames: %w", err)
			}
			fragments = append(fragments, xmlFragment{
				OrderKey: workbookElementOrder["definedNames"],
				Seq:      writerSeq,
				XML:      defsFrag,
			})
		}
	}

	for _, extra := range data.ExtraElements {
		fragments = append(fragments, xmlFragment{
			OrderKey: extra.OrderKey,
			Seq:      extra.Seq,
			XML:      []byte(extra.XML),
		})
	}

	// Known limitation: if the original file had an extLst with other
	// extensions (e.g. chart tracking) but no future-function metadata,
	// and the user programmatically adds a LET/LAMBDA formula, the
	// future-function extensions won't be injected because we don't parse
	// extLst content. The file will open with a repair prompt. To fix
	// this properly, extLst would need to be partially parsed to detect
	// whether specific extension URIs are present.
	if workbookNeedsFutureFunctionsMetadata(data) && !hasExtraElementName(data.ExtraElements, "extLst") {
		extFrag, err := encodeNamedXMLFragment("extLst", &xlsxExtLst{InnerXML: futureFunctionsWorkbookExtXML})
		if err != nil {
			return fmt.Errorf("encode workbook extLst: %w", err)
		}
		fragments = append(fragments, xmlFragment{
			OrderKey: workbookElementOrder["extLst"],
			Seq:      writerSeq,
			XML:      extFrag,
		})
	}

	return writePassthroughXML(zw, "xl/workbook.xml", "workbook", rootAttrs, fragments)
}

// isExternalRef returns true if a defined name value references an external
// workbook. External refs start with [n] or '[n] (e.g. "[1]Sheet!$A$1" or
// "'[1]Sheet'!$A$1").
func isExternalRef(value string) bool {
	return strings.HasPrefix(value, "[") || strings.HasPrefix(value, "'[")
}

func writeWorkbookRels(zw *zip.Writer, data *WorkbookData, sheetRelIDs []string, hasSST, hasDynamicArrayMetadata bool) error {
	rels := xlsxRelationships{
		Xmlns: NSRelationships,
	}
	for _, rel := range data.ExtraRels {
		rels.Relationships = append(rels.Relationships, xlsxRelationship{
			ID:         rel.ID,
			Type:       rel.Type,
			Target:     rel.Target,
			TargetMode: rel.TargetMode,
		})
	}
	relNS := relationshipsNamespace(data.RootAttrs)
	worksheetType := strictRelationshipType(relNS, RelTypeWorksheet, RelTypeWorksheetStrict)
	stylesType := strictRelationshipType(relNS, RelTypeStyles, RelTypeStylesStrict)
	sharedStrType := strictRelationshipType(relNS, RelTypeSharedStr, RelTypeSharedStrStrict)
	metadataType := strictRelationshipType(relNS, RelTypeSheetMetadata, RelTypeSheetMetadataStrict)

	// Use the pre-computed sheetRelIDs for worksheet entries, then allocate
	// additional IDs for styles/SST/metadata that don't collide.
	for i := range data.Sheets {
		rels.Relationships = append(rels.Relationships, xlsxRelationship{
			ID:     sheetRelIDs[i],
			Type:   worksheetType,
			Target: fmt.Sprintf("worksheets/sheet%d.xml", i+1),
		})
	}
	tailCount := 1 + boolInt(hasSST) + boolInt(hasDynamicArrayMetadata)
	tailIDs := nextRelationshipIDs(data.ExtraRels, tailCount, sheetRelIDs...)
	idx := 0
	rels.Relationships = append(rels.Relationships, xlsxRelationship{
		ID:     tailIDs[idx],
		Type:   stylesType,
		Target: "styles.xml",
	})
	idx++
	if hasSST {
		rels.Relationships = append(rels.Relationships, xlsxRelationship{
			ID:     tailIDs[idx],
			Type:   sharedStrType,
			Target: "sharedStrings.xml",
		})
		idx++
	}
	if hasDynamicArrayMetadata {
		rels.Relationships = append(rels.Relationships, xlsxRelationship{
			ID:     tailIDs[idx],
			Type:   metadataType,
			Target: "metadata.xml",
		})
	}
	return writeXML(zw, "xl/_rels/workbook.xml.rels", rels)
}

func writeSheet(zw *zip.Writer, num int, sd *SheetData, styleIndexMap []int, tables []tableWriteInfo, tableRelIDs []string) error {
	// Populate column widths.
	var cols *xlsxCols
	if len(sd.ColWidths) > 0 {
		cols = &xlsxCols{}
		for _, cw := range sd.ColWidths {
			cols.Col = append(cols.Col, xlsxCol{
				Min:         cw.Min,
				Max:         cw.Max,
				Width:       cw.Width,
				CustomWidth: 1,
			})
		}
	}
	var mergeCells *xlsxMergeCells
	if len(sd.MergeCells) > 0 {
		mergeCells = &xlsxMergeCells{Count: len(sd.MergeCells)}
		for _, mc := range sd.MergeCells {
			ref := mc.StartAxis
			if mc.EndAxis != "" && mc.EndAxis != mc.StartAxis {
				ref += ":" + mc.EndAxis
			}
			mergeCells.MergeCell = append(mergeCells.MergeCell, xlsxMergeCell{Ref: ref})
		}
	}

	var sheetData xlsxSheetData
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
				} else if cd.IsDynamicArray && cd.FormulaRef == "" {
					// Dynamic array without a spill range (e.g. #SPILL!).
					// Write cm=1 so the formula is recognized as dynamic on
					// reimport, but omit t="array" and ref.
					c.CM = 1
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
		sheetData.Rows = append(sheetData.Rows, row)
	}

	var fragments []xmlFragment
	rootAttrs := mergedRootAttrs(sd.RootAttrs, map[string]string{
		"xmlns": spreadsheetNamespace(sd.RootAttrs),
	})
	if len(tables) > 0 || len(sd.ExtraRels) > 0 || rawAttrValue(rootAttrs, "xmlns:r") != "" {
		rootAttrs = mergedRootAttrs(rootAttrs, map[string]string{
			"xmlns:r": relationshipsNamespace(sd.RootAttrs),
		})
	}

	if cols != nil {
		frag, err := encodeNamedXMLFragment("cols", cols)
		if err != nil {
			return fmt.Errorf("encode cols: %w", err)
		}
		fragments = append(fragments, xmlFragment{
			OrderKey: worksheetElementOrder["cols"],
			Seq:      writerSeq,
			XML:      frag,
		})
	}

	sheetDataFrag, err := encodeNamedXMLFragment("sheetData", sheetData)
	if err != nil {
		return fmt.Errorf("encode sheetData: %w", err)
	}
	fragments = append(fragments, xmlFragment{
		OrderKey: worksheetElementOrder["sheetData"],
		Seq:      writerSeq,
		XML:      sheetDataFrag,
	})

	if mergeCells != nil {
		frag, err := encodeNamedXMLFragment("mergeCells", mergeCells)
		if err != nil {
			return fmt.Errorf("encode mergeCells: %w", err)
		}
		fragments = append(fragments, xmlFragment{
			OrderKey: worksheetElementOrder["mergeCells"],
			Seq:      writerSeq,
			XML:      frag,
		})
	}

	for _, extra := range sd.ExtraElements {
		fragments = append(fragments, xmlFragment{
			OrderKey: extra.OrderKey,
			Seq:      extra.Seq,
			XML:      []byte(extra.XML),
		})
	}

	if len(tables) > 0 {
		tableParts := &xlsxTableParts{Count: len(tables)}
		for i := range tables {
			rid := fmt.Sprintf("rId%d", i+1)
			if i < len(tableRelIDs) {
				rid = tableRelIDs[i]
			}
			tableParts.TablePart = append(tableParts.TablePart, xlsxTablePart{RID: rid})
		}
		frag, err := encodeNamedXMLFragment("tableParts", tableParts)
		if err != nil {
			return fmt.Errorf("encode tableParts: %w", err)
		}
		fragments = append(fragments, xmlFragment{
			OrderKey: worksheetElementOrder["tableParts"],
			Seq:      writerSeq,
			XML:      frag,
		})
	}

	return writePassthroughXML(zw, fmt.Sprintf("xl/worksheets/sheet%d.xml", num), "worksheet", rootAttrs, fragments)
}

func writeSheetRels(zw *zip.Writer, sd *SheetData, sheetNum int, tables []tableWriteInfo, tableRelIDs []string) error {
	rels := xlsxRelationships{Xmlns: NSRelationships}
	for _, rel := range sd.ExtraRels {
		rels.Relationships = append(rels.Relationships, xlsxRelationship{
			ID:         rel.ID,
			Type:       rel.Type,
			Target:     rel.Target,
			TargetMode: rel.TargetMode,
		})
	}
	tableType := strictRelationshipType(relationshipsNamespace(sd.RootAttrs), RelTypeTable, RelTypeTableStrict)
	for i, tw := range tables {
		rid := fmt.Sprintf("rId%d", i+1)
		if i < len(tableRelIDs) {
			rid = tableRelIDs[i]
		}
		rels.Relationships = append(rels.Relationships, xlsxRelationship{
			ID:     rid,
			Type:   tableType,
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

func buildWriterPaths(
	data *WorkbookData,
	tablePlan map[int][]tableWriteInfo,
	hasSST, hasCoreProps, hasDynamicArrayMetadata bool,
) map[string]struct{} {
	paths := map[string]struct{}{
		"[Content_Types].xml":        {},
		"_rels/.rels":                {},
		"docProps/app.xml":           {},
		"xl/workbook.xml":            {},
		"xl/_rels/workbook.xml.rels": {},
		"xl/styles.xml":              {},
	}
	if hasCoreProps {
		paths["docProps/core.xml"] = struct{}{}
	}
	if hasSST {
		paths["xl/sharedStrings.xml"] = struct{}{}
	}
	if hasDynamicArrayMetadata {
		paths["xl/metadata.xml"] = struct{}{}
	}
	for i := range data.Sheets {
		paths[fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1)] = struct{}{}
		if len(tablePlan[i]) > 0 || len(data.Sheets[i].ExtraRels) > 0 {
			paths[fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", i+1)] = struct{}{}
		}
	}
	for _, sheetTables := range tablePlan {
		for _, tw := range sheetTables {
			paths[fmt.Sprintf("xl/tables/table%d.xml", tw.PartNum)] = struct{}{}
		}
	}
	return paths
}

func strictRelationshipType(relNS, transitional, strict string) string {
	if relNS == NSOfficeDocumentStrict {
		return strict
	}
	return transitional
}

func boolInt(v bool) int {
	if v {
		return 1
	}
	return 0
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
				if cell.IsDynamicArray {
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

func writeOpaqueEntries(zw *zip.Writer, entries []OpaqueEntry, emittedPaths map[string]struct{}) error {
	if len(entries) == 0 {
		return nil
	}
	sorted := append([]OpaqueEntry(nil), entries...)
	sort.SliceStable(sorted, func(i, j int) bool {
		return strings.ToLower(sorted[i].Path) < strings.ToLower(sorted[j].Path)
	})
	for _, entry := range sorted {
		if _, ok := emittedPaths[entry.Path]; ok {
			continue
		}
		if err := writeRawFile(zw, entry.Path, entry.Data); err != nil {
			return err
		}
	}
	return nil
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
