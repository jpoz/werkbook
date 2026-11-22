package ooxml

import (
	"encoding/xml"
	"fmt"
)

const (
	NSRelationships      = "http://schemas.openxmlformats.org/package/2006/relationships"
	NSOfficeDocument     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	RelTypeWorkbook      = NSOfficeDocument + "/officeDocument"
	RelTypeWorksheet     = NSOfficeDocument + "/worksheet"
	RelTypeStyles        = NSOfficeDocument + "/styles"
	RelTypeSharedStr     = NSOfficeDocument + "/sharedStrings"
	RelTypeSheetMetadata = NSOfficeDocument + "/sheetMetadata"
	NSSpreadsheetML      = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

	RelTypeTable = NSOfficeDocument + "/table"

	// Strict OOXML namespaces (ISO/IEC 29500 Strict).
	NSOfficeDocumentStrict     = "http://purl.oclc.org/ooxml/officeDocument/relationships"
	RelTypeWorksheetStrict     = NSOfficeDocumentStrict + "/worksheet"
	RelTypeStylesStrict        = NSOfficeDocumentStrict + "/styles"
	RelTypeSharedStrStrict     = NSOfficeDocumentStrict + "/sharedStrings"
	RelTypeSheetMetadataStrict = NSOfficeDocumentStrict + "/sheetMetadata"
	RelTypeTableStrict         = NSOfficeDocumentStrict + "/table"
)

type xlsxRelationships struct {
	XMLName       xml.Name           `xml:"Relationships"`
	Xmlns         string             `xml:"xmlns,attr"`
	Relationships []xlsxRelationship `xml:"Relationship"`
}

type xlsxRelationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

// ParseRelationships parses a relationships part (a .rels file).
func ParseRelationships(data []byte) ([]OpaqueRel, error) {
	var parsed xlsxRelationships
	if err := xml.Unmarshal(data, &parsed); err != nil {
		return nil, fmt.Errorf("unmarshal relationships: %w", err)
	}
	rels := make([]OpaqueRel, len(parsed.Relationships))
	for i, rel := range parsed.Relationships {
		rels[i] = opaqueRelFromXML(rel)
	}
	return rels, nil
}

// MarshalRelationships serializes a relationships part (a .rels file).
func MarshalRelationships(rels []OpaqueRel) ([]byte, error) {
	out := xlsxRelationships{Xmlns: NSRelationships}
	for _, rel := range rels {
		out.Relationships = append(out.Relationships, xlsxRelationship(rel))
	}
	body, err := xml.Marshal(out)
	if err != nil {
		return nil, fmt.Errorf("marshal relationships: %w", err)
	}
	return append([]byte(xmlHeader), selfCloseEmptyElements(body)...), nil
}
