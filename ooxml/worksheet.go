package ooxml

import (
	"encoding/xml"
	"strconv"
)

type xlsxWorksheet struct {
	XMLName    xml.Name        `xml:"worksheet"`
	Xmlns      string          `xml:"xmlns,attr"`
	XmlnsR     string          `xml:"xmlns:r,attr,omitempty"`
	Cols       *xlsxCols       `xml:"cols,omitempty"`
	SheetData  xlsxSheetData   `xml:"sheetData"`
	MergeCells *xlsxMergeCells `xml:"mergeCells,omitempty"`
	TableParts *xlsxTableParts `xml:"tableParts,omitempty"`
}

type xlsxCols struct {
	Col []xlsxCol `xml:"col"`
}

// ooxmlBool is a boolean type for OOXML XML attributes.
// It unmarshals both "true"/"false" and "1"/"0", and always marshals as "1" or
// is omitted when false (via omitempty on int).
type ooxmlBool int

func (b *ooxmlBool) UnmarshalXMLAttr(attr xml.Attr) error {
	switch attr.Value {
	case "1", "true", "on":
		*b = 1
	case "0", "false", "off", "":
		*b = 0
	default:
		// Try parsing as int for any other numeric value.
		n, err := strconv.Atoi(attr.Value)
		if err != nil {
			return err
		}
		if n != 0 {
			*b = 1
		} else {
			*b = 0
		}
	}
	return nil
}

type xlsxCol struct {
	Min         int       `xml:"min,attr"`
	Max         int       `xml:"max,attr"`
	Width       float64   `xml:"width,attr"`
	CustomWidth ooxmlBool `xml:"customWidth,attr,omitempty"`
}

type xlsxMergeCells struct {
	Count     int             `xml:"count,attr,omitempty"`
	MergeCell []xlsxMergeCell `xml:"mergeCell"`
}

type xlsxMergeCell struct {
	Ref string `xml:"ref,attr"`
}

type xlsxSheetData struct {
	Rows []xlsxRow `xml:"row"`
}

type xlsxRow struct {
	R            int       `xml:"r,attr"`
	Ht           float64   `xml:"ht,attr,omitempty"`
	CustomHeight ooxmlBool `xml:"customHeight,attr,omitempty"`
	Hidden       ooxmlBool `xml:"hidden,attr,omitempty"`
	Cells        []xlsxC   `xml:"c"`
}

type xlsxC struct {
	R  string  `xml:"r,attr"`
	S  int     `xml:"s,attr,omitempty"`
	T  string  `xml:"t,attr,omitempty"`
	CM int     `xml:"cm,attr,omitempty"`
	FE *xlsxF  `xml:"f,omitempty"` // formula element with attributes
	V  string  `xml:"v,omitempty"`
	IS *xlsxIS `xml:"is,omitempty"`
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
	return c.FE != nil && c.FE.T == "array" && c.CM == 0
}

// IsDynamicArrayFormula reports whether the cell uses Excel dynamic-array metadata.
func (c xlsxC) IsDynamicArrayFormula() bool {
	return c.FE != nil && c.FE.T == "array" && c.CM != 0
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

type xlsxTableParts struct {
	Count     int             `xml:"count,attr,omitempty"`
	TablePart []xlsxTablePart `xml:"tablePart"`
}

type xlsxTablePart struct {
	RID string
}

// MarshalXML writes the tablePart element with a raw r:id attribute so the
// worksheet-level xmlns:r declaration is reused instead of redefined locally.
func (tp xlsxTablePart) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "tablePart"}
	start.Attr = []xml.Attr{{Name: xml.Name{Local: "r:id"}, Value: tp.RID}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}
