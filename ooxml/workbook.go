package ooxml

import "encoding/xml"

type xlsxWorkbook struct {
	XMLName xml.Name   `xml:"workbook"`
	Xmlns   string     `xml:"xmlns,attr"`
	XmlnsR  string     `xml:"xmlns:r,attr"`
	Sheets  xlsxSheets `xml:"sheets"`
}

type xlsxSheets struct {
	Sheet []xlsxSheet `xml:"sheet"`
}

type xlsxSheet struct {
	Name    string `xml:"name,attr"`
	SheetID int    `xml:"sheetId,attr"`
	RID     string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

// WorkbookData is the internal boundary between the public API and the ooxml package.
type WorkbookData struct {
	Sheets []SheetData
	Styles []StyleData // index 0 = default (empty)
}

// StyleData is the intermediate representation of a cell style,
// passed between the public API and the ooxml layer.
type StyleData struct {
	FontName  string
	FontColor string // 8-char ARGB
	FontSize  float64
	FontBold  bool
	FontItalic bool
	FontUL    bool

	FillColor string // 8-char ARGB

	BorderLeftStyle   string // OOXML border style string
	BorderLeftColor   string // 8-char ARGB
	BorderRightStyle  string
	BorderRightColor  string
	BorderTopStyle    string
	BorderTopColor    string
	BorderBottomStyle string
	BorderBottomColor string

	HAlign   string // OOXML horizontal alignment
	VAlign   string // OOXML vertical alignment
	WrapText bool

	NumFmtID int    // built-in number format ID
	NumFmt   string // custom format string
}

// ColWidthData holds the width for a range of columns.
type ColWidthData struct {
	Min   int     // 1-based first column
	Max   int     // 1-based last column
	Width float64 // column width in character units
}

// SheetData holds the data for a single worksheet.
type SheetData struct {
	Name      string
	Rows      []RowData
	ColWidths []ColWidthData
}

// RowData holds the data for a single row.
type RowData struct {
	Num    int // 1-based
	Height float64
	Cells  []CellData
}

// CellData holds the data for a single cell.
type CellData struct {
	Ref      string // e.g. "A1"
	Type     string // "s" (shared string), "b" (bool), "inlineStr", or "" (number)
	Value    string // raw value (SST index for strings, "0"/"1" for bools, float string for numbers)
	Formula  string // formula text (empty = no formula)
	StyleIdx int    // index into WorkbookData.Styles; 0 = default
}
