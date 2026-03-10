package ooxml

import (
	"encoding/xml"
	"strings"
)

type xlsxSST struct {
	XMLName     xml.Name `xml:"sst"`
	Xmlns       string   `xml:"xmlns,attr"`
	Count       int      `xml:"count,attr"`
	UniqueCount int      `xml:"uniqueCount,attr"`
	SI          []xlsxSI `xml:"si"`
}

type xlsxSI struct {
	T *string `xml:"t,omitempty"`
	R []xlsxR `xml:"r,omitempty"`
}

type xlsxR struct {
	T string `xml:"t"`
}

// MarshalXML serializes an xlsxSI item, adding xml:space="preserve" on the <t>
// element when the string has leading or trailing whitespace. Without this
// attribute, OOXML-compliant readers strip the whitespace
// when loading the shared string table.
func (si xlsxSI) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if si.T != nil {
		s := *si.T
		tStart := xml.StartElement{Name: xml.Name{Local: "t"}}
		if strings.HasPrefix(s, " ") || strings.HasSuffix(s, " ") ||
			strings.HasPrefix(s, "\t") || strings.HasSuffix(s, "\t") {
			tStart.Attr = []xml.Attr{{
				Name:  xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"},
				Value: "preserve",
			}}
		}
		if err := e.EncodeToken(tStart); err != nil {
			return err
		}
		if err := e.EncodeToken(xml.CharData(s)); err != nil {
			return err
		}
		if err := e.EncodeToken(tStart.End()); err != nil {
			return err
		}
	}

	for _, r := range si.R {
		if err := e.EncodeElement(r, xml.StartElement{Name: xml.Name{Local: "r"}}); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// SharedStringTable builds a shared string table at write time.
type SharedStringTable struct {
	strings []string
	index   map[string]int
	count   int
}

// NewSharedStringTable creates a new empty shared string table.
func NewSharedStringTable() *SharedStringTable {
	return &SharedStringTable{
		index: make(map[string]int),
	}
}

// Add adds a string to the table and returns its index.
// If the string already exists, it returns the existing index.
func (sst *SharedStringTable) Add(s string) int {
	sst.count++
	if idx, ok := sst.index[s]; ok {
		return idx
	}
	idx := len(sst.strings)
	sst.strings = append(sst.strings, s)
	sst.index[s] = idx
	return idx
}

// Len returns the number of unique strings in the table.
func (sst *SharedStringTable) Len() int {
	return len(sst.strings)
}

// Get returns the string at the given index.
func (sst *SharedStringTable) Get(idx int) string {
	return sst.strings[idx]
}

// Strings returns all strings in the table.
func (sst *SharedStringTable) Strings() []string {
	return sst.strings
}

// ToXML converts the shared string table to XML representation.
func (sst *SharedStringTable) ToXML() xlsxSST {
	x := xlsxSST{
		Xmlns:       NSSpreadsheetML,
		Count:       sst.count,
		UniqueCount: len(sst.strings),
	}
	for _, s := range sst.strings {
		str := s // copy for pointer
		x.SI = append(x.SI, xlsxSI{T: &str})
	}
	return x
}
