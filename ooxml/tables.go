package ooxml

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"strings"
	"unicode/utf8"
)

type xlsxWrittenTable struct {
	XMLName        xml.Name             `xml:"table"`
	Xmlns          string               `xml:"xmlns,attr"`
	ID             int                  `xml:"id,attr"`
	Name           string               `xml:"name,attr"`
	DisplayName    string               `xml:"displayName,attr"`
	Ref            string               `xml:"ref,attr"`
	HeaderRowCount *int                 `xml:"headerRowCount,attr,omitempty"`
	TotalsRowCount *int                 `xml:"totalsRowCount,attr,omitempty"`
	AutoFilter     *xlsxAutoFilter      `xml:"autoFilter,omitempty"`
	TableColumns   xlsxWrittenTableCols `xml:"tableColumns"`
	TableStyleInfo *xlsxTableStyleInfo  `xml:"tableStyleInfo,omitempty"`
}

type xlsxWrittenTableCols struct {
	Count  int                   `xml:"count,attr"`
	Column []xlsxWrittenTableCol `xml:"tableColumn"`
}

type xlsxWrittenTableCol struct {
	ID   int    `xml:"id,attr"`
	Name string `xml:"name,attr"`
}

func writeTableXML(zw *zip.Writer, partNum int, td TableDef) error {
	name := td.Name
	if name == "" {
		name = td.DisplayName
	}
	displayName := td.DisplayName
	if displayName == "" {
		displayName = name
	}

	xt := xlsxWrittenTable{
		Xmlns:       NSSpreadsheetML,
		ID:          partNum,
		Name:        name,
		DisplayName: displayName,
		Ref:         td.Ref,
		TableColumns: xlsxWrittenTableCols{
			Count: len(td.Columns),
		},
	}
	if td.HeaderRowCount != 1 {
		xt.HeaderRowCount = intPtr(td.HeaderRowCount)
	}
	if td.TotalsRowCount > 0 {
		xt.TotalsRowCount = intPtr(td.TotalsRowCount)
	}
	if td.HasAutoFilter && td.HeaderRowCount != 0 {
		xt.AutoFilter = &xlsxAutoFilter{Ref: td.Ref}
	}
	if td.Style != nil && td.Style.Name != "" {
		xt.TableStyleInfo = &xlsxTableStyleInfo{
			Name:              td.Style.Name,
			ShowFirstColumn:   boolOOXML(td.Style.ShowFirstColumn),
			ShowLastColumn:    boolOOXML(td.Style.ShowLastColumn),
			ShowRowStripes:    boolOOXML(td.Style.ShowRowStripes),
			ShowColumnStripes: boolOOXML(td.Style.ShowColumnStripes),
		}
	}
	for i, col := range td.Columns {
		xt.TableColumns.Column = append(xt.TableColumns.Column, xlsxWrittenTableCol{
			ID:   i + 1,
			Name: encodeOOXMLEscapes(col),
		})
	}
	return writeXML(zw, fmt.Sprintf("xl/tables/table%d.xml", partNum), xt)
}

func intPtr(v int) *int {
	return &v
}

func boolOOXML(v bool) ooxmlBool {
	if v {
		return 1
	}
	return 0
}

func encodeOOXMLEscapes(s string) string {
	var b strings.Builder
	for i := 0; i < len(s); {
		if looksLikeOOXMLEscape(s, i) {
			b.WriteString("_x005F_")
			b.WriteByte(s[i])
			i++
			continue
		}

		r := rune(s[i])
		width := 1
		if r >= 0x80 {
			r, width = utf8.DecodeRuneInString(s[i:])
		}
		switch r {
		case '\t', '\n', '\r':
			fmt.Fprintf(&b, "_x%04X_", r)
		default:
			if r < 0x20 {
				fmt.Fprintf(&b, "_x%04X_", r)
			} else {
				b.WriteRune(r)
			}
		}
		i += width
	}
	return b.String()
}

func looksLikeOOXMLEscape(s string, i int) bool {
	if i+6 >= len(s) || s[i] != '_' || s[i+1] != 'x' || s[i+6] != '_' {
		return false
	}
	for j := i + 2; j < i+6; j++ {
		ch := s[j]
		if !(ch >= '0' && ch <= '9' || ch >= 'a' && ch <= 'f' || ch >= 'A' && ch <= 'F') {
			return false
		}
	}
	return true
}
