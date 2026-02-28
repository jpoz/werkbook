package ooxml

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// ReadWorkbook reads an XLSX file and returns the parsed WorkbookData.
func ReadWorkbook(r io.ReaderAt, size int64) (*WorkbookData, error) {
	zr, err := zip.NewReader(r, size)
	if err != nil {
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
		if rel.Type == RelTypeWorksheet {
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
	for _, s := range wb.Sheets.Sheet {
		target, ok := sheetRels[s.RID]
		if !ok {
			continue
		}
		// Target is relative to xl/
		path := "xl/" + target
		ws, err := readXML[xlsxWorksheet](files, path)
		if err != nil {
			return nil, fmt.Errorf("read sheet %q: %w", s.Name, err)
		}

		sd := SheetData{Name: s.Name}

		// Extract column widths.
		if ws.Cols != nil {
			for _, col := range ws.Cols.Col {
				sd.ColWidths = append(sd.ColWidths, ColWidthData{
					Min: col.Min, Max: col.Max, Width: col.Width,
				})
			}
		}

		for _, xr := range ws.SheetData.Rows {
			rd := RowData{Num: xr.R}
			if xr.CustomHeight && xr.Ht != 0 {
				rd.Height = xr.Ht
			}
			for _, xc := range xr.Cells {
				cd := parseCellData(xc, sst)
				rd.Cells = append(rd.Cells, cd)
			}
			if len(rd.Cells) > 0 || rd.Height != 0 {
				sd.Rows = append(sd.Rows, rd)
			}
		}
		data.Sheets = append(data.Sheets, sd)
	}
	return data, nil
}

func parseCellData(xc xlsxC, sst []string) CellData {
	cd := CellData{Ref: xc.R, Formula: xc.F, StyleIdx: xc.S}

	switch xc.T {
	case "s":
		// Shared string: value is the SST index.
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
			sd.WrapText = xf.Alignment.WrapText
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

func readXML[T any](files map[string]*zip.File, name string) (T, error) {
	var zero T
	f, ok := files[name]
	if !ok {
		return zero, fmt.Errorf("file %q not found in archive", name)
	}
	rc, err := f.Open()
	if err != nil {
		return zero, err
	}
	defer rc.Close()

	var v T
	dec := xml.NewDecoder(rc)
	if err := dec.Decode(&v); err != nil {
		return zero, fmt.Errorf("decode %s: %w", name, err)
	}
	return v, nil
}
