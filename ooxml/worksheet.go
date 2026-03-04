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
	R  string   `xml:"r,attr"`
	S  int      `xml:"s,attr,omitempty"`
	T  string   `xml:"t,attr,omitempty"`
	FE *xlsxF   `xml:"f,omitempty"` // formula element with attributes
	V  string   `xml:"v,omitempty"`
	IS *xlsxIS  `xml:"is,omitempty"`
}

// F returns the formula text (for backward-compat convenience).
func (c xlsxC) F() string {
	if c.FE != nil {
		return c.FE.Text
	}
	return ""
}

// IsArrayFormula reports whether the cell's formula is a CSE array formula.
func (c xlsxC) IsArrayFormula() bool {
	return c.FE != nil && c.FE.T == "array"
}

type xlsxF struct {
	T    string `xml:"t,attr,omitempty"`   // "array", "shared", etc.
	Ref  string `xml:"ref,attr,omitempty"` // range for array/shared formulas
	Si   int    `xml:"si,attr,omitempty"`  // shared formula index
	Text string `xml:",chardata"`          // the formula text
}

type xlsxIS struct {
	T string `xml:"t"`
}
