package ooxml

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"strconv"
	"strings"

	"github.com/jpoz/werkbook/formula"
)

// ErrEncryptedFile is returned when the input file is encrypted (password-protected).
var ErrEncryptedFile = errors.New("file is password-protected or encrypted; werkbook cannot open encrypted files")

// cfbMagic is the magic number for Microsoft Compound Binary Format (CFB/OLE2) files.
// Encrypted OOXML files are wrapped in this format.
var cfbMagic = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

// ReadWorkbook reads an XLSX file and returns the parsed WorkbookData.
func ReadWorkbook(r io.ReaderAt, size int64) (*WorkbookData, error) {
	zr, err := zip.NewReader(r, size)
	if err != nil {
		// Check if the file is an encrypted CFB/OLE2 container.
		if size >= 8 {
			var header [8]byte
			if _, readErr := r.ReadAt(header[:], 0); readErr == nil {
				match := true
				for i := range cfbMagic {
					if header[i] != cfbMagic[i] {
						match = false
						break
					}
				}
				if match {
					return nil, ErrEncryptedFile
				}
			}
		}
		return nil, fmt.Errorf("open zip: %w", err)
	}

	files := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		files[f.Name] = f
	}

	// Parse workbook relationships to find sheet paths.
	wbRels, err := readXML[xlsxRelationships](files, "xl/_rels/workbook.xml.rels")
	if err != nil {
		return nil, fmt.Errorf("read workbook rels: %w", err)
	}
	sheetRels := make(map[string]string) // rId -> target path
	for _, rel := range wbRels.Relationships {
		if rel.Type == RelTypeWorksheet || rel.Type == RelTypeWorksheetStrict {
			sheetRels[rel.ID] = rel.Target
		}
	}

	// Parse workbook to get sheet names and ordering.
	wb, err := readXML[xlsxWorkbook](files, "xl/workbook.xml")
	if err != nil {
		return nil, fmt.Errorf("read workbook: %w", err)
	}

	// Parse shared string table (may not exist).
	sst, _ := readSST(files)

	// Parse styles (may not exist).
	styles := readStyles(files)

	data := &WorkbookData{Styles: styles}
	if corePropsRaw, err := readFile(files, "docProps/core.xml"); err == nil {
		data.CorePropsRaw = corePropsRaw
		if coreProps, err := parseCoreProperties(corePropsRaw); err == nil {
			data.CoreProps = coreProps
		}
	}

	// Check for the 1904 date system.
	if wb.WorkbookPr != nil {
		v := wb.WorkbookPr.Date1904
		data.Date1904 = v == "1" || v == "true"
	}
	if wb.CalcPr != nil {
		data.CalcProps = CalcPropertiesData{
			Mode:           wb.CalcPr.CalcMode,
			ID:             wb.CalcPr.CalcID,
			FullCalcOnLoad: parseOOXMLBoolString(wb.CalcPr.FullCalcOnLoad),
			ForceFullCalc:  parseOOXMLBoolString(wb.CalcPr.ForceFullCalc),
			Completed:      parseOOXMLBoolString(wb.CalcPr.CalcCompleted),
		}
	}

	// Parse defined names.
	if wb.DefinedNames != nil {
		for _, dn := range wb.DefinedNames.DefinedName {
			localID := -1
			if dn.LocalSheetID != nil {
				localID = *dn.LocalSheetID
			}
			data.DefinedNames = append(data.DefinedNames, DefinedName{
				Name:         dn.Name,
				Value:        dn.Value,
				LocalSheetID: localID,
			})
		}
	}
	for _, s := range wb.Sheets.Sheet {
		target, ok := sheetRels[s.RID]
		if !ok {
			continue
		}
		// Target may be relative (e.g. "worksheets/sheet1.xml") or
		// absolute (e.g. "/xl/worksheets/sheet1.xml"). Absolute paths
		// start with "/" and are relative to the ZIP root.
		var path string
		if strings.HasPrefix(target, "/") {
			path = target[1:]
		} else {
			path = "xl/" + target
		}
		ws, err := readXML[xlsxWorksheet](files, path)
		if err != nil {
			return nil, fmt.Errorf("read sheet %q: %w", s.Name, err)
		}

		sd := SheetData{Name: s.Name, State: s.State}

		// Extract column widths.
		if ws.Cols != nil {
			for _, col := range ws.Cols.Col {
				sd.ColWidths = append(sd.ColWidths, ColWidthData{
					Min: col.Min, Max: col.Max, Width: col.Width,
				})
			}
		}

		for _, xr := range ws.SheetData.Rows {
			rd := RowData{Num: xr.R, Hidden: xr.Hidden == 1}
			if xr.CustomHeight == 1 && xr.Ht != 0 {
				rd.Height = xr.Ht
			}
			for _, xc := range xr.Cells {
				cd := parseCellData(xc, sst)
				rd.Cells = append(rd.Cells, cd)
			}
			if len(rd.Cells) > 0 || rd.Height != 0 || rd.Hidden {
				sd.Rows = append(sd.Rows, rd)
			}
		}
		// Expand shared formulas so child cells get their own formula text.
		expandSharedFormulas(&sd)

		if ws.MergeCells != nil {
			for _, mc := range ws.MergeCells.MergeCell {
				parts := strings.SplitN(mc.Ref, ":", 2)
				start := parts[0]
				end := start
				if len(parts) == 2 {
					end = parts[1]
				}
				sd.MergeCells = append(sd.MergeCells, MergeCellData{
					StartAxis: start,
					EndAxis:   end,
				})
			}
		}

		// Read table definitions associated with this sheet.
		sheetIdx := len(data.Sheets)
		tables := readSheetTables(files, path, sheetIdx)
		data.Tables = append(data.Tables, tables...)

		data.Sheets = append(data.Sheets, sd)
	}
	return data, nil
}

func parseCellData(xc xlsxC, sst []string) CellData {
	isArrayFormula := xc.IsArrayFormula()
	isDynamicArray := false
	if xc.FE != nil && formula.IsDynamicArrayFormula(xc.F()) {
		isDynamicArray = true
		isArrayFormula = false
	}
	cd := CellData{
		Ref:            xc.R,
		Formula:        xc.F(),
		FormulaType:    formulaType(xc.FE),
		FormulaRef:     formulaRef(xc.FE),
		IsArrayFormula: isArrayFormula,
		IsDynamicArray: isDynamicArray,
		SharedIndex:    sharedIndex(xc.FE),
		StyleIdx:       xc.S,
	}

	switch xc.T {
	case "s":
		// Shared string: value is the SST index. Per the OOXML spec,
		// t="s" means the cell value is a string. The formula engine's
		// CoerceNum handles numeric coercion when arithmetic needs it.
		idx, err := strconv.Atoi(xc.V)
		if err == nil && idx >= 0 && idx < len(sst) {
			cd.Type = "s"
			cd.Value = sst[idx]
		}
	case "inlineStr":
		cd.Type = "inlineStr"
		if xc.IS != nil {
			cd.Value = xc.IS.T
		}
	case "str":
		cd.Type = "str"
		cd.Value = xc.V
	case "d":
		cd.Type = "d"
		cd.Value = xc.V
	case "b":
		cd.Type = "b"
		cd.Value = xc.V
	case "e":
		cd.Type = "e"
		cd.Value = xc.V
	default:
		// Number (or empty).
		cd.Value = xc.V
	}
	return cd
}

func parseOOXMLBoolString(v string) bool {
	switch strings.ToLower(v) {
	case "1", "true", "on":
		return true
	default:
		return false
	}
}

func formulaType(fe *xlsxF) string {
	if fe == nil {
		return ""
	}
	return fe.T
}

func formulaRef(fe *xlsxF) string {
	if fe == nil {
		return ""
	}
	return fe.Ref
}

func sharedIndex(fe *xlsxF) int {
	if fe == nil || fe.T != "shared" {
		return -1
	}
	return fe.Si
}

// readStyles parses xl/styles.xml and returns a []StyleData indexed by cellXfs position.
// Returns a single empty StyleData if styles.xml is missing or unparseable.
func readStyles(files map[string]*zip.File) []StyleData {
	ss, err := readXML[xlsxStyleSheet](files, "xl/styles.xml")
	if err != nil {
		return []StyleData{{}}
	}

	// Build lookup tables for fonts, fills, borders, numFmts.
	fonts := ss.Fonts.Font
	fills := ss.Fills.Fill
	borders := ss.Borders.Border
	numFmtMap := make(map[int]string)
	if ss.NumFmts != nil {
		for _, nf := range ss.NumFmts.NumFmt {
			numFmtMap[nf.NumFmtID] = nf.FormatCode
		}
	}

	styles := make([]StyleData, 0, len(ss.CellXfs.Xf))
	for _, xf := range ss.CellXfs.Xf {
		var sd StyleData

		// Font
		if xf.FontID >= 0 && xf.FontID < len(fonts) {
			f := fonts[xf.FontID]
			if f.Name != nil {
				sd.FontName = f.Name.Val
			}
			if f.Sz != nil {
				sd.FontSize = f.Sz.Val
			}
			sd.FontBold = f.B != nil
			sd.FontItalic = f.I != nil
			sd.FontUL = f.U != nil
			if f.Color != nil {
				sd.FontColor = f.Color.RGB
			}
		}

		// Fill
		if xf.FillID >= 0 && xf.FillID < len(fills) {
			pf := fills[xf.FillID].PatternFill
			if pf.PatternType == "solid" && pf.FgColor != nil {
				sd.FillColor = pf.FgColor.RGB
			}
		}

		// Border
		if xf.BorderID >= 0 && xf.BorderID < len(borders) {
			b := borders[xf.BorderID]
			sd.BorderLeftStyle = b.Left.Style
			if b.Left.Color != nil {
				sd.BorderLeftColor = b.Left.Color.RGB
			}
			sd.BorderRightStyle = b.Right.Style
			if b.Right.Color != nil {
				sd.BorderRightColor = b.Right.Color.RGB
			}
			sd.BorderTopStyle = b.Top.Style
			if b.Top.Color != nil {
				sd.BorderTopColor = b.Top.Color.RGB
			}
			sd.BorderBottomStyle = b.Bottom.Style
			if b.Bottom.Color != nil {
				sd.BorderBottomColor = b.Bottom.Color.RGB
			}
		}

		// Number format
		sd.NumFmtID = xf.NumFmtID
		if code, ok := numFmtMap[xf.NumFmtID]; ok {
			sd.NumFmt = code
		}

		// Alignment
		if xf.Alignment != nil {
			sd.HAlign = xf.Alignment.Horizontal
			sd.VAlign = xf.Alignment.Vertical
			sd.WrapText = xf.Alignment.WrapText == 1
		}

		styles = append(styles, sd)
	}

	if len(styles) == 0 {
		return []StyleData{{}}
	}
	return styles
}

func readSST(files map[string]*zip.File) ([]string, error) {
	x, err := readXML[xlsxSST](files, "xl/sharedStrings.xml")
	if err != nil {
		return nil, err
	}
	strs := make([]string, 0, len(x.SI))
	for _, si := range x.SI {
		strs = append(strs, siToString(si))
	}
	return strs, nil
}

// siToString extracts the plain text from a shared string item.
// It handles both simple <t> elements and rich text <r><t> elements.
func siToString(si xlsxSI) string {
	if si.T != nil {
		return *si.T
	}
	// Rich text: concatenate all <r><t> values.
	var sb strings.Builder
	for _, r := range si.R {
		sb.WriteString(r.T)
	}
	return sb.String()
}

// decodeOOXMLEscapes replaces _xHHHH_ escape sequences in OOXML attribute
// values with the corresponding Unicode character.
func decodeOOXMLEscapes(s string) string {
	// Fast path: no escape sequences.
	if !strings.Contains(s, "_x") {
		return s
	}
	var b strings.Builder
	i := 0
	for i < len(s) {
		if i+6 < len(s) && s[i] == '_' && s[i+1] == 'x' && s[i+6] == '_' {
			hex := s[i+2 : i+6]
			code, err := strconv.ParseUint(hex, 16, 16)
			if err == nil {
				b.WriteRune(rune(code))
				i += 7
				continue
			}
		}
		b.WriteByte(s[i])
		i++
	}
	return b.String()
}

// xlsxTable represents the <table> root element in xl/tables/table*.xml.
type xlsxTable struct {
	XMLName        xml.Name            `xml:"table"`
	Name           string              `xml:"name,attr"`
	DisplayName    string              `xml:"displayName,attr"`
	Ref            string              `xml:"ref,attr"`
	HeaderRowCount *int                `xml:"headerRowCount,attr"`
	TotalsRowCount int                 `xml:"totalsRowCount,attr"`
	AutoFilter     *xlsxAutoFilter     `xml:"autoFilter"`
	TableColumns   xlsxTableColumns    `xml:"tableColumns"`
	TableStyleInfo *xlsxTableStyleInfo `xml:"tableStyleInfo"`
}

type xlsxAutoFilter struct {
	Ref           string             `xml:"ref,attr"`
	FilterColumns []xlsxFilterColumn `xml:"filterColumn"`
}

type xlsxFilterColumn struct {
	ColID int `xml:"colId,attr"`
}

type xlsxTableColumns struct {
	Column []xlsxTableColumn `xml:"tableColumn"`
}

type xlsxTableColumn struct {
	Name string `xml:"name,attr"`
}

type xlsxTableStyleInfo struct {
	Name              string    `xml:"name,attr"`
	ShowFirstColumn   ooxmlBool `xml:"showFirstColumn,attr,omitempty"`
	ShowLastColumn    ooxmlBool `xml:"showLastColumn,attr,omitempty"`
	ShowRowStripes    ooxmlBool `xml:"showRowStripes,attr,omitempty"`
	ShowColumnStripes ooxmlBool `xml:"showColumnStripes,attr,omitempty"`
}

// readSheetTables reads table definitions referenced by a sheet's relationship file.
// sheetPath is the path to the sheet XML (e.g. "xl/worksheets/sheet1.xml").
// sheetIndex is the 0-based index of the sheet in WorkbookData.Sheets.
func readSheetTables(files map[string]*zip.File, sheetPath string, sheetIndex int) []TableDef {
	// Determine the relationship file path for this sheet.
	// e.g. "xl/worksheets/sheet1.xml" → "xl/worksheets/_rels/sheet1.xml.rels"
	lastSlash := strings.LastIndex(sheetPath, "/")
	var relsPath string
	if lastSlash >= 0 {
		dir := sheetPath[:lastSlash]
		base := sheetPath[lastSlash+1:]
		relsPath = dir + "/_rels/" + base + ".rels"
	} else {
		relsPath = "_rels/" + sheetPath + ".rels"
	}

	sheetRels, err := readXML[xlsxRelationships](files, relsPath)
	if err != nil {
		return nil
	}

	var tables []TableDef
	for _, rel := range sheetRels.Relationships {
		if rel.Type != RelTypeTable && rel.Type != RelTypeTableStrict {
			continue
		}
		// Resolve the table path relative to the sheet directory.
		var tablePath string
		if strings.HasPrefix(rel.Target, "/") {
			tablePath = rel.Target[1:]
		} else if lastSlash >= 0 {
			tablePath = resolveRelativePath(sheetPath[:lastSlash], rel.Target)
		} else {
			tablePath = rel.Target
		}

		xt, err := readXML[xlsxTable](files, tablePath)
		if err != nil {
			continue
		}

		td := TableDef{
			Name:            xt.Name,
			DisplayName:     xt.DisplayName,
			Ref:             xt.Ref,
			SheetIndex:      sheetIndex,
			TotalsRowCount:  xt.TotalsRowCount,
			HasAutoFilter:   xt.AutoFilter != nil,
			HasActiveFilter: xt.AutoFilter != nil && len(xt.AutoFilter.FilterColumns) > 0,
		}
		// Default headerRowCount is 1 if not specified.
		if xt.HeaderRowCount != nil {
			td.HeaderRowCount = *xt.HeaderRowCount
		} else {
			td.HeaderRowCount = 1
		}
		for _, col := range xt.TableColumns.Column {
			td.Columns = append(td.Columns, decodeOOXMLEscapes(col.Name))
		}
		if xt.TableStyleInfo != nil {
			td.Style = &TableStyleData{
				Name:              xt.TableStyleInfo.Name,
				ShowFirstColumn:   xt.TableStyleInfo.ShowFirstColumn != 0,
				ShowLastColumn:    xt.TableStyleInfo.ShowLastColumn != 0,
				ShowRowStripes:    xt.TableStyleInfo.ShowRowStripes != 0,
				ShowColumnStripes: xt.TableStyleInfo.ShowColumnStripes != 0,
			}
		}
		tables = append(tables, td)
	}
	return tables
}

// resolveRelativePath resolves a relative target path against a base directory.
// Handles "../" prefixes.
func resolveRelativePath(baseDir, target string) string {
	for strings.HasPrefix(target, "../") {
		target = target[3:]
		if idx := strings.LastIndex(baseDir, "/"); idx >= 0 {
			baseDir = baseDir[:idx]
		} else {
			baseDir = ""
		}
	}
	if baseDir == "" {
		return target
	}
	return baseDir + "/" + target
}

func readXML[T any](files map[string]*zip.File, name string) (T, error) {
	var zero T
	data, err := readFile(files, name)
	if err != nil {
		return zero, err
	}
	var v T
	if err := xml.Unmarshal(data, &v); err != nil {
		return zero, fmt.Errorf("unmarshal %s: %w", name, err)
	}
	return v, nil
}

func readFile(files map[string]*zip.File, name string) ([]byte, error) {
	f, ok := files[name]
	if !ok {
		return nil, fmt.Errorf("file %q not found in archive", name)
	}
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	data, err := io.ReadAll(rc)
	if err != nil {
		return nil, fmt.Errorf("read %s: %w", name, err)
	}
	return data, nil
}
