package ooxml

import "encoding/xml"

type xlsxTypes struct {
	XMLName   xml.Name       `xml:"Types"`
	Xmlns     string         `xml:"xmlns,attr"`
	Defaults  []xlsxDefault  `xml:"Default"`
	Overrides []xlsxOverride `xml:"Override"`
}

type xlsxDefault struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

type xlsxOverride struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

const contentTypesNS = "http://schemas.openxmlformats.org/package/2006/content-types"
