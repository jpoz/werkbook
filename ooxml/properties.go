package ooxml

import (
	"archive/zip"
	"encoding/xml"
	"time"
)

const (
	NSCoreProperties   = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
	NSDublinCore       = "http://purl.org/dc/elements/1.1/"
	NSDublinCoreTerms  = "http://purl.org/dc/terms/"
	NSDublinCoreType   = "http://purl.org/dc/dcmitype/"
	NSXMLSchemaInst    = "http://www.w3.org/2001/XMLSchema-instance"
	NSExtendedProps    = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
	NSDocPropsVTypes   = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
	RelTypeCoreProps   = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
	RelTypeExtendedApp = NSOfficeDocument + "/extended-properties"
)

// CorePropertiesData holds the supported OPC core properties from docProps/core.xml.
type CorePropertiesData struct {
	Title          string
	Subject        string
	Creator        string
	Description    string
	Identifier     string
	Language       string
	Keywords       string
	Category       string
	ContentStatus  string
	Version        string
	Revision       string
	LastModifiedBy string
	Created        time.Time
	Modified       time.Time
}

type xlsxCorePropertiesRead struct {
	XMLName        xml.Name        `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties coreProperties"`
	Title          string          `xml:"http://purl.org/dc/elements/1.1/ title"`
	Subject        string          `xml:"http://purl.org/dc/elements/1.1/ subject"`
	Creator        string          `xml:"http://purl.org/dc/elements/1.1/ creator"`
	Description    string          `xml:"http://purl.org/dc/elements/1.1/ description"`
	Identifier     string          `xml:"http://purl.org/dc/elements/1.1/ identifier"`
	Language       string          `xml:"http://purl.org/dc/elements/1.1/ language"`
	Keywords       string          `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties keywords"`
	Category       string          `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties category"`
	ContentStatus  string          `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties contentStatus"`
	Version        string          `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties version"`
	Revision       string          `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties revision"`
	LastModifiedBy string          `xml:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties lastModifiedBy"`
	Created        *xlsxW3CDTFRead `xml:"http://purl.org/dc/terms/ created"`
	Modified       *xlsxW3CDTFRead `xml:"http://purl.org/dc/terms/ modified"`
}

type xlsxW3CDTFRead struct {
	Type  string `xml:"http://www.w3.org/2001/XMLSchema-instance type,attr"`
	Value string `xml:",chardata"`
}

type xlsxCorePropertiesWrite struct {
	XMLName        xml.Name         `xml:"cp:coreProperties"`
	XmlnsCP        string           `xml:"xmlns:cp,attr"`
	XmlnsDC        string           `xml:"xmlns:dc,attr"`
	XmlnsDCTerms   string           `xml:"xmlns:dcterms,attr"`
	XmlnsDCMIType  string           `xml:"xmlns:dcmitype,attr"`
	XmlnsXSI       string           `xml:"xmlns:xsi,attr"`
	Title          string           `xml:"dc:title,omitempty"`
	Subject        string           `xml:"dc:subject,omitempty"`
	Creator        string           `xml:"dc:creator,omitempty"`
	Description    string           `xml:"dc:description,omitempty"`
	Identifier     string           `xml:"dc:identifier,omitempty"`
	Language       string           `xml:"dc:language,omitempty"`
	Keywords       string           `xml:"cp:keywords,omitempty"`
	Category       string           `xml:"cp:category,omitempty"`
	ContentStatus  string           `xml:"cp:contentStatus,omitempty"`
	Version        string           `xml:"cp:version,omitempty"`
	Revision       string           `xml:"cp:revision,omitempty"`
	LastModifiedBy string           `xml:"cp:lastModifiedBy,omitempty"`
	Created        *xlsxW3CDTFWrite `xml:"dcterms:created,omitempty"`
	Modified       *xlsxW3CDTFWrite `xml:"dcterms:modified,omitempty"`
}

type xlsxW3CDTFWrite struct {
	Type  string `xml:"xsi:type,attr,omitempty"`
	Value string `xml:",chardata"`
}

type xlsxAppProperties struct {
	XMLName           xml.Name          `xml:"Properties"`
	Xmlns             string            `xml:"xmlns,attr"`
	XmlnsVT           string            `xml:"xmlns:vt,attr"`
	Application       string            `xml:"Application"`
	DocSecurity       int               `xml:"DocSecurity"`
	ScaleCrop         bool              `xml:"ScaleCrop"`
	HeadingPairs      xlsxHeadingPairs  `xml:"HeadingPairs"`
	TitlesOfParts     xlsxTitlesOfParts `xml:"TitlesOfParts"`
	Company           string            `xml:"Company"`
	LinksUpToDate     bool              `xml:"LinksUpToDate"`
	SharedDoc         bool              `xml:"SharedDoc"`
	HyperlinksChanged bool              `xml:"HyperlinksChanged"`
	AppVersion        string            `xml:"AppVersion"`
}

type xlsxHeadingPairs struct {
	Vector xlsxVTVectorVariants `xml:"vt:vector"`
}

type xlsxVTVectorVariants struct {
	Size     int             `xml:"size,attr"`
	BaseType string          `xml:"baseType,attr"`
	Variant  []xlsxVTVariant `xml:"vt:variant"`
}

type xlsxVTVariant struct {
	LPSTR *string `xml:"vt:lpstr,omitempty"`
	I4    *int    `xml:"vt:i4,omitempty"`
}

type xlsxTitlesOfParts struct {
	Vector xlsxVTVectorStrings `xml:"vt:vector"`
}

type xlsxVTVectorStrings struct {
	Size     int      `xml:"size,attr"`
	BaseType string   `xml:"baseType,attr"`
	LPSTR    []string `xml:"vt:lpstr"`
}

func parseCoreProperties(data []byte) (CorePropertiesData, error) {
	var x xlsxCorePropertiesRead
	if err := xml.Unmarshal(data, &x); err != nil {
		return CorePropertiesData{}, err
	}

	props := CorePropertiesData{
		Title:          x.Title,
		Subject:        x.Subject,
		Creator:        x.Creator,
		Description:    x.Description,
		Identifier:     x.Identifier,
		Language:       x.Language,
		Keywords:       x.Keywords,
		Category:       x.Category,
		ContentStatus:  x.ContentStatus,
		Version:        x.Version,
		Revision:       x.Revision,
		LastModifiedBy: x.LastModifiedBy,
	}

	var err error
	if x.Created != nil && x.Created.Value != "" {
		props.Created, err = parseW3CDTF(x.Created.Value)
		if err != nil {
			return CorePropertiesData{}, err
		}
	}
	if x.Modified != nil && x.Modified.Value != "" {
		props.Modified, err = parseW3CDTF(x.Modified.Value)
		if err != nil {
			return CorePropertiesData{}, err
		}
	}

	return props, nil
}

func buildCoreProperties(props CorePropertiesData) xlsxCorePropertiesWrite {
	x := xlsxCorePropertiesWrite{
		XmlnsCP:        NSCoreProperties,
		XmlnsDC:        NSDublinCore,
		XmlnsDCTerms:   NSDublinCoreTerms,
		XmlnsDCMIType:  NSDublinCoreType,
		XmlnsXSI:       NSXMLSchemaInst,
		Title:          props.Title,
		Subject:        props.Subject,
		Creator:        props.Creator,
		Description:    props.Description,
		Identifier:     props.Identifier,
		Language:       props.Language,
		Keywords:       props.Keywords,
		Category:       props.Category,
		ContentStatus:  props.ContentStatus,
		Version:        props.Version,
		Revision:       props.Revision,
		LastModifiedBy: props.LastModifiedBy,
	}
	if !props.Created.IsZero() {
		x.Created = &xlsxW3CDTFWrite{
			Type:  "dcterms:W3CDTF",
			Value: formatW3CDTF(props.Created),
		}
	}
	if !props.Modified.IsZero() {
		x.Modified = &xlsxW3CDTFWrite{
			Type:  "dcterms:W3CDTF",
			Value: formatW3CDTF(props.Modified),
		}
	}
	return x
}

func buildAppProperties(sheetNames []string) xlsxAppProperties {
	label := "Worksheets"
	sheetCount := len(sheetNames)

	return xlsxAppProperties{
		Xmlns:       NSExtendedProps,
		XmlnsVT:     NSDocPropsVTypes,
		Application: "Werkbook",
		DocSecurity: 0,
		ScaleCrop:   false,
		HeadingPairs: xlsxHeadingPairs{
			Vector: xlsxVTVectorVariants{
				Size:     2,
				BaseType: "variant",
				Variant: []xlsxVTVariant{
					{LPSTR: stringPtr(label)},
					{I4: intPtr(sheetCount)},
				},
			},
		},
		TitlesOfParts: xlsxTitlesOfParts{
			Vector: xlsxVTVectorStrings{
				Size:     sheetCount,
				BaseType: "lpstr",
				LPSTR:    append([]string(nil), sheetNames...),
			},
		},
		Company:           "",
		LinksUpToDate:     false,
		SharedDoc:         false,
		HyperlinksChanged: false,
		AppVersion:        "1.0",
	}
}

func writeAppProperties(zw *zip.Writer, data *WorkbookData) error {
	sheetNames := make([]string, len(data.Sheets))
	for i, sheet := range data.Sheets {
		sheetNames[i] = sheet.Name
	}
	return writeXML(zw, "docProps/app.xml", buildAppProperties(sheetNames))
}

func writeCoreProperties(zw *zip.Writer, data *WorkbookData) error {
	if len(data.CorePropsRaw) > 0 {
		return writeRawFile(zw, "docProps/core.xml", data.CorePropsRaw)
	}
	return writeXML(zw, "docProps/core.xml", buildCoreProperties(data.CoreProps))
}

func hasCorePropertiesData(props CorePropertiesData) bool {
	return props.Title != "" ||
		props.Subject != "" ||
		props.Creator != "" ||
		props.Description != "" ||
		props.Identifier != "" ||
		props.Language != "" ||
		props.Keywords != "" ||
		props.Category != "" ||
		props.ContentStatus != "" ||
		props.Version != "" ||
		props.Revision != "" ||
		props.LastModifiedBy != "" ||
		!props.Created.IsZero() ||
		!props.Modified.IsZero()
}

func parseW3CDTF(v string) (time.Time, error) {
	layouts := []string{
		time.RFC3339Nano,
		"2006-01-02T15:04:05",
		"2006-01-02",
		"2006-01",
		"2006",
	}
	var err error
	for _, layout := range layouts {
		var t time.Time
		t, err = time.Parse(layout, v)
		if err == nil {
			return t, nil
		}
	}
	return time.Time{}, err
}

func formatW3CDTF(t time.Time) string {
	return t.UTC().Format(time.RFC3339)
}

func stringPtr(v string) *string {
	return &v
}
