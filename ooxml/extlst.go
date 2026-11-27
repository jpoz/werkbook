package ooxml

import (
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"strings"
)

// extensionXML is one <ext> element of an <extLst>, with the URI that
// identifies it.
type extensionXML struct {
	URI string
	XML string
}

// futureFunctionsWorkbookExts are the workbook extensions that Excel expects
// when a sheet formula uses a future function (an _xlfn. prefix).
var futureFunctionsWorkbookExts = []extensionXML{
	{URI: futureFunctionsWorkbookPrExtURI, XML: futureFunctionsWorkbookPrExtXML},
	{URI: futureFunctionsCalcFeaturesExtURI, XML: futureFunctionsCalcFeaturesExtXML},
}

// addMissingExts returns the given <extLst> element with each of the given
// extensions that it does not contain added at its end. Extensions are
// matched by URI, with no difference between upper and lower case. The
// content that the element has is kept byte for byte.
func addMissingExts(extLstXML string, exts []extensionXML) (string, error) {
	present, rootName, closeOffset, err := scanExtLst(extLstXML)
	if err != nil {
		return "", err
	}

	var missing strings.Builder
	for _, ext := range exts {
		if !present[strings.ToUpper(ext.URI)] {
			missing.WriteString(ext.XML)
		}
	}
	if missing.Len() == 0 {
		return extLstXML, nil
	}

	if closeOffset >= 0 {
		return extLstXML[:closeOffset] + missing.String() + extLstXML[closeOffset:], nil
	}
	// The element is self-closing: <extLst/>.
	end := strings.LastIndex(extLstXML, "/>")
	if end < 0 {
		return "", errors.New("extLst has no end tag")
	}
	return extLstXML[:end] + ">" + missing.String() + "</" + rootName + ">" + extLstXML[end+2:], nil
}

// scanExtLst reads an <extLst> element. It returns the URIs of its <ext>
// children in upper case, the name of the element as written (with its
// prefix), and the byte offset of its end tag. The offset is -1 when the
// element is self-closing.
func scanExtLst(extLstXML string) (present map[string]bool, rootName string, closeOffset int, err error) {
	present = make(map[string]bool)
	closeOffset = -1
	dec := xml.NewDecoder(strings.NewReader(extLstXML))
	depth := 0
	sawRoot := false
	for {
		offset := int(dec.InputOffset())
		tok, tokErr := dec.RawToken()
		if tokErr == io.EOF {
			break
		}
		if tokErr != nil {
			return nil, "", -1, fmt.Errorf("read extLst: %w", tokErr)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch depth {
			case 0:
				if sawRoot {
					return nil, "", -1, errors.New("extLst has more than one root element")
				}
				sawRoot = true
				rootName = rawTokenName(t.Name)
			case 1:
				if t.Name.Local == "ext" {
					for _, attr := range t.Attr {
						if attr.Name.Space == "" && attr.Name.Local == "uri" {
							present[strings.ToUpper(attr.Value)] = true
						}
					}
				}
			}
			depth++
		case xml.EndElement:
			depth--
			if depth < 0 {
				return nil, "", -1, errors.New("extLst has an end tag with no start tag")
			}
			// For a self-closing element the decoder gives an end token
			// that has no text of its own.
			if depth == 0 && strings.HasPrefix(extLstXML[offset:], "</") {
				closeOffset = offset
			}
		}
	}
	if !sawRoot || depth != 0 {
		return nil, "", -1, errors.New("extLst is not a complete element")
	}
	return present, rootName, closeOffset, nil
}
