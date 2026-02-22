package ooxml

import "encoding/xml"

const (
	NSRelationships  = "http://schemas.openxmlformats.org/package/2006/relationships"
	NSOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	RelTypeWorkbook  = NSOfficeDocument + "/officeDocument"
	RelTypeWorksheet = NSOfficeDocument + "/worksheet"
	RelTypeStyles    = NSOfficeDocument + "/styles"
	RelTypeSharedStr = NSOfficeDocument + "/sharedStrings"
	NSSpreadsheetML  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
)

type xlsxRelationships struct {
	XMLName       xml.Name           `xml:"Relationships"`
	Xmlns         string             `xml:"xmlns,attr"`
	Relationships []xlsxRelationship `xml:"Relationship"`
}

type xlsxRelationship struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}
