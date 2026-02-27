package ooxml

import "encoding/xml"

type xlsxWorksheet struct {
	XMLName   xml.Name      `xml:"worksheet"`
	Xmlns     string        `xml:"xmlns,attr"`
	Cols      *xlsxCols     `xml:"cols,omitempty"`
	SheetData xlsxSheetData `xml:"sheetData"`
}

type xlsxCols struct {
	Col []xlsxCol `xml:"col"`
}

type xlsxCol struct {
	Min         int     `xml:"min,attr"`
	Max         int     `xml:"max,attr"`
	Width       float64 `xml:"width,attr"`
	CustomWidth bool    `xml:"customWidth,attr,omitempty"`
}

type xlsxSheetData struct {
	Rows []xlsxRow `xml:"row"`
}

type xlsxRow struct {
	R            int     `xml:"r,attr"`
	Ht           float64 `xml:"ht,attr,omitempty"`
	CustomHeight bool    `xml:"customHeight,attr,omitempty"`
	Cells        []xlsxC `xml:"c"`
}

type xlsxC struct {
	R  string  `xml:"r,attr"`
	S  int     `xml:"s,attr,omitempty"`
	T  string  `xml:"t,attr,omitempty"`
	F  string  `xml:"f,omitempty"`
	V  string  `xml:"v,omitempty"`
	IS *xlsxIS `xml:"is,omitempty"`
}

type xlsxIS struct {
	T string `xml:"t"`
}
