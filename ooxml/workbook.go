package ooxml

import (
	"encoding/xml"
	"fmt"
	"strconv"
)

type xlsxWorkbookPr struct {
	Date1904 string `xml:"date1904,attr,omitempty"`
}

type xlsxWorkbook struct {
	XMLName      xml.Name             `xml:"workbook"`
	Xmlns        string               `xml:"xmlns,attr"`
	XmlnsR       string               `xml:"xmlns:r,attr"`
	WorkbookPr   *xlsxWorkbookPr      `xml:"workbookPr,omitempty"`
	Sheets       xlsxSheets           `xml:"sheets"`
	DefinedNames *xlsxDefinedNames    `xml:"definedNames,omitempty"`
}

type xlsxDefinedNames struct {
	DefinedName []xlsxDefinedName `xml:"definedName"`
}

type xlsxDefinedName struct {
	Name         string `xml:"name,attr"`
	LocalSheetID *int   `xml:"localSheetId,attr"`
	Value        string `xml:",chardata"`
}

type xlsxSheets struct {
	Sheet []xlsxSheet `xml:"sheet"`
}

type xlsxSheet struct {
	Name    string
	SheetID int
	RID     string
}

// UnmarshalXML handles both transitional and strict OOXML namespaces for the r:id attribute.
func (s *xlsxSheet) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "name":
			s.Name = attr.Value
		case "sheetId":
			id, err := strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
			s.SheetID = id
		case "id":
			if attr.Name.Space == NSOfficeDocument || attr.Name.Space == NSOfficeDocumentStrict {
				s.RID = attr.Value
			}
		}
	}
	return d.Skip()
}

// MarshalXML writes the sheet element with the r:id attribute in the transitional namespace.
func (s xlsxSheet) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "sheet"}
	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "name"}, Value: s.Name},
		{Name: xml.Name{Local: "sheetId"}, Value: fmt.Sprintf("%d", s.SheetID)},
		{Name: xml.Name{Space: NSOfficeDocument, Local: "id"}, Value: s.RID},
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// DefinedName represents a named range or named formula from the workbook.
type DefinedName struct {
	Name         string // the defined name (e.g. "OneRange")
	Value        string // the reference or expression (e.g. "Sheet1!$A$10")
	LocalSheetID int    // -1 for global; otherwise 0-based sheet index
}

// WorkbookData is the internal boundary between the public API and the ooxml package.
type WorkbookData struct {
	Date1904     bool          // true if the workbook uses the 1904 date system (Mac Excel)
	Sheets       []SheetData
	Styles       []StyleData   // index 0 = default (empty)
	Tables       []TableDef    // table definitions parsed from xl/tables/table*.xml
	DefinedNames []DefinedName // named ranges/formulas from <definedNames>
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

// TableDef holds the definition of an Excel table (ListObject).
type TableDef struct {
	Name           string   // internal name
	DisplayName    string   // display name (used in structured references)
	Ref            string   // range reference, e.g. "A1:E20"
	SheetIndex     int      // 0-based index into WorkbookData.Sheets
	Columns        []string // column names in order
	HeaderRowCount int      // number of header rows (default 1)
	TotalsRowCount int      // number of totals rows (default 0)
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
	Ref            string // e.g. "A1"
	Type           string // "s" (shared string), "b" (bool), "inlineStr", or "" (number)
	Value          string // raw value (SST index for strings, "0"/"1" for bools, float string for numbers)
	Formula        string // formula text (empty = no formula)
	IsArrayFormula bool   // true if the formula is a CSE (Ctrl+Shift+Enter) array formula
	StyleIdx       int    // index into WorkbookData.Styles; 0 = default
}
