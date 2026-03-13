package ooxml

import "encoding/xml"

type xlsxExternalLink struct {
	XMLName      xml.Name          `xml:"externalLink"`
	ExternalBook *xlsxExternalBook `xml:"externalBook"`
}

type xlsxExternalBook struct {
	SheetNames   *xlsxExternalSheetNames   `xml:"sheetNames"`
	SheetDataSet *xlsxExternalSheetDataSet `xml:"sheetDataSet"`
}

type xlsxExternalSheetNames struct {
	SheetName []xlsxExternalSheetName `xml:"sheetName"`
}

type xlsxExternalSheetName struct {
	Val string `xml:"val,attr"`
}

type xlsxExternalSheetDataSet struct {
	SheetData []xlsxExternalSheetData `xml:"sheetData"`
}

type xlsxExternalSheetData struct {
	SheetID int               `xml:"sheetId,attr"`
	Rows    []xlsxExternalRow `xml:"row"`
}

type xlsxExternalRow struct {
	R     int                `xml:"r,attr"`
	Cells []xlsxExternalCell `xml:"cell"`
}

type xlsxExternalCell struct {
	R string `xml:"r,attr"`
	T string `xml:"t,attr,omitempty"`
	V string `xml:"v,omitempty"`
}
