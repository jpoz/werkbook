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

	data := &WorkbookData{}
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
		for _, xr := range ws.SheetData.Rows {
			rd := RowData{Num: xr.R}
			for _, xc := range xr.Cells {
				cd := parseCellData(xc, sst)
				rd.Cells = append(rd.Cells, cd)
			}
			if len(rd.Cells) > 0 {
				sd.Rows = append(sd.Rows, rd)
			}
		}
		data.Sheets = append(data.Sheets, sd)
	}
	return data, nil
}

func parseCellData(xc xlsxC, sst []string) CellData {
	cd := CellData{Ref: xc.R}

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
