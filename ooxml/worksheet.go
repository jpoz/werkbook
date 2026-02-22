package ooxml

import "encoding/xml"

type xlsxWorksheet struct {
	XMLName   xml.Name       `xml:"worksheet"`
	Xmlns     string         `xml:"xmlns,attr"`
	SheetData xlsxSheetData  `xml:"sheetData"`
}

type xlsxSheetData struct {
	Rows []xlsxRow `xml:"row"`
}

type xlsxRow struct {
	R     int     `xml:"r,attr"`
	Cells []xlsxC `xml:"c"`
}

type xlsxC struct {
	R string `xml:"r,attr"`
	T string `xml:"t,attr,omitempty"`
	V string `xml:"v,omitempty"`
	// For inline strings
	IS *xlsxIS `xml:"is,omitempty"`
}

type xlsxIS struct {
	T string `xml:"t"`
}
