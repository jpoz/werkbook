package ooxml

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strings"
)

// RawAttr preserves an attribute on a root XML element, including namespace
// declarations such as xmlns:mc or prefixed attributes such as mc:Ignorable.
type RawAttr struct {
	Name  string
	Value string
}

// RawElement preserves a direct child element that werkbook does not parse.
// XML stores the element's outer XML verbatim as it appeared in the source part.
type RawElement struct {
	Name     string
	XML      string
	OrderKey int
	Seq      int
}

// rawTokenName returns a colon-joined "prefix:local" name from an xml.Name
// obtained via RawToken (where Space is the prefix, not the namespace URI).
func rawTokenName(name xml.Name) string {
	if name.Space == "" {
		return name.Local
	}
	return name.Space + ":" + name.Local
}

var worksheetElementOrder = map[string]int{
	"sheetPr":               0,
	"dimension":             2,
	"sheetViews":            4,
	"sheetFormatPr":         6,
	"cols":                  8,
	"sheetData":             10,
	"sheetCalcPr":           12,
	"sheetProtection":       14,
	"protectedRanges":       16,
	"scenarios":             18,
	"autoFilter":            20,
	"sortState":             22,
	"dataConsolidate":       24,
	"customSheetViews":      26,
	"mergeCells":            28,
	"phoneticPr":            30,
	"conditionalFormatting": 32,
	"dataValidations":       34,
	"hyperlinks":            36,
	"printOptions":          38,
	"pageMargins":           40,
	"pageSetup":             42,
	"headerFooter":          44,
	"rowBreaks":             46,
	"colBreaks":             48,
	"customProperties":      50,
	"cellWatches":           52,
	"ignoredErrors":         54,
	"smartTags":             56,
	"drawing":               58,
	"legacyDrawing":         60,
	"legacyDrawingHF":       62,
	"drawingHF":             64,
	"picture":               66,
	"oleObjects":            68,
	"controls":              70,
	"webPublishItems":       72,
	"tableParts":            74,
	"extLst":                76,
}

var workbookElementOrder = map[string]int{
	"fileVersion":         0,
	"fileSharing":         2,
	"workbookPr":          4,
	"workbookProtection":  6,
	"bookViews":           8,
	"sheets":              10,
	"functionGroups":      12,
	"externalReferences":  14,
	"definedNames":        16,
	"calcPr":              18,
	"oleSize":             20,
	"customWorkbookViews": 22,
	"pivotCaches":         24,
	"smartTagPr":          26,
	"smartTagTypes":       28,
	"webPublishing":       30,
	"fileRecoveryPr":      32,
	"webPublishObjects":   34,
	"extLst":              36,
}

func worksheetOrderKey(name string, lastKnown int) int {
	return passthroughOrderKey(worksheetElementOrder, name, lastKnown)
}

func workbookOrderKey(name string, lastKnown int) int {
	return passthroughOrderKey(workbookElementOrder, name, lastKnown)
}

func passthroughOrderKey(order map[string]int, name string, lastKnown int) int {
	if key, ok := order[name]; ok {
		return key
	}
	if lastKnown >= 0 {
		return lastKnown + 1
	}
	return -1
}

func captureRootAttrsAndExtras(
	data []byte,
	known func(string) bool,
	orderKey func(string, int) int,
) ([]RawAttr, []RawElement, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	prev := int64(0)
	depth := 0
	lastKnown := -1
	seq := 0

	type captureState struct {
		name     string
		start    int64
		orderKey int
		seq      int
		depth    int
	}
	var capture *captureState
	var rootAttrs []RawAttr
	var extras []RawElement

	for {
		tok, err := dec.RawToken()
		off := dec.InputOffset()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if depth == 0 {
				rootAttrs = make([]RawAttr, 0, len(t.Attr))
				for _, attr := range t.Attr {
					rootAttrs = append(rootAttrs, RawAttr{
						Name:  rawTokenName(attr.Name),
						Value: attr.Value,
					})
				}
			} else {
				if capture != nil {
					capture.depth++
				} else if depth == 1 {
					name := rawTokenName(t.Name)
					if known(name) {
						if key := orderKey(name, lastKnown); key >= 0 {
							lastKnown = key
						}
					} else {
						key := orderKey(name, lastKnown)
						raw := data[prev:off]
						if bytes.HasSuffix(raw, []byte("/>")) {
							extras = append(extras, RawElement{
								Name:     name,
								XML:      string(raw),
								OrderKey: key,
								Seq:      seq,
							})
						} else {
							capture = &captureState{
								name:     name,
								start:    prev,
								orderKey: key,
								seq:      seq,
								depth:    1,
							}
						}
					}
					seq++
				}
			}
			depth++
		case xml.EndElement:
			depth--
			if capture != nil {
				capture.depth--
				if capture.depth == 0 {
					extras = append(extras, RawElement{
						Name:     capture.name,
						XML:      string(data[capture.start:off]),
						OrderKey: capture.orderKey,
						Seq:      capture.seq,
					})
					capture = nil
				}
			}
		}
		prev = off
	}

	return rootAttrs, extras, nil
}

func rawAttrsToXML(attrs []RawAttr) []xml.Attr {
	out := make([]xml.Attr, 0, len(attrs))
	for _, attr := range attrs {
		out = append(out, xml.Attr{
			Name:  xml.Name{Local: attr.Name},
			Value: attr.Value,
		})
	}
	return out
}

func mergedRootAttrs(stored []RawAttr, required map[string]string) []RawAttr {
	attrs := append([]RawAttr(nil), stored...)
	index := make(map[string]int, len(attrs))
	for i, attr := range attrs {
		index[attr.Name] = i
	}
	for name, value := range required {
		if i, ok := index[name]; ok {
			if attrs[i].Value == "" {
				attrs[i].Value = value
			}
			continue
		}
		attrs = append(attrs, RawAttr{Name: name, Value: value})
	}
	return attrs
}

func rawAttrValue(attrs []RawAttr, name string) string {
	for _, attr := range attrs {
		if attr.Name == name {
			return attr.Value
		}
	}
	return ""
}

func spreadsheetNamespace(attrs []RawAttr) string {
	if ns := rawAttrValue(attrs, "xmlns"); ns != "" {
		return ns
	}
	return NSSpreadsheetML
}

func relationshipsNamespace(attrs []RawAttr) string {
	if ns := rawAttrValue(attrs, "xmlns:r"); ns != "" {
		return ns
	}
	if strings.Contains(strings.ToLower(spreadsheetNamespace(attrs)), "purl.oclc.org/ooxml") {
		return NSOfficeDocumentStrict
	}
	return NSOfficeDocument
}

func hasExtraElementName(extras []RawElement, name string) bool {
	for _, extra := range extras {
		if extra.Name == name {
			return true
		}
	}
	return false
}

// writerSeq is the Seq value used for writer-generated fragments. It sorts
// after all reader-captured extras (which use small sequential integers),
// ensuring that when a writer-generated element and an opaque extra share the
// same OrderKey, the extra's original position is respected.
const writerSeq = 1 << 30

type xmlFragment struct {
	OrderKey int
	Seq      int
	XML      []byte
}

func encodeNamedXMLFragment(name string, v any) ([]byte, error) {
	var buf bytes.Buffer
	enc := xml.NewEncoder(&buf)
	start := xml.StartElement{Name: xml.Name{Local: name}}
	if err := enc.EncodeElement(v, start); err != nil {
		return nil, err
	}
	if err := enc.Close(); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func writePassthroughXML(
	zw zipWriter,
	name string,
	root string,
	rootAttrs []RawAttr,
	fragments []xmlFragment,
) error {
	w, err := zw.Create(name)
	if err != nil {
		return fmt.Errorf("create %s: %w", name, err)
	}

	if _, err := io.WriteString(w, xmlHeader); err != nil {
		return err
	}

	var buf bytes.Buffer
	enc := xml.NewEncoder(&buf)
	start := xml.StartElement{
		Name: xml.Name{Local: root},
		Attr: rawAttrsToXML(rootAttrs),
	}
	if err := enc.EncodeToken(start); err != nil {
		return fmt.Errorf("encode start %s: %w", name, err)
	}

	sort.SliceStable(fragments, func(i, j int) bool {
		if fragments[i].OrderKey != fragments[j].OrderKey {
			return fragments[i].OrderKey < fragments[j].OrderKey
		}
		return fragments[i].Seq < fragments[j].Seq
	})

	for _, fragment := range fragments {
		if len(fragment.XML) == 0 {
			continue
		}
		if err := enc.Flush(); err != nil {
			return fmt.Errorf("flush %s: %w", name, err)
		}
		if _, err := buf.Write(fragment.XML); err != nil {
			return fmt.Errorf("buffer %s: %w", name, err)
		}
	}

	if err := enc.EncodeToken(start.End()); err != nil {
		return fmt.Errorf("encode end %s: %w", name, err)
	}
	if err := enc.Close(); err != nil {
		return fmt.Errorf("close %s: %w", name, err)
	}
	if _, err := w.Write(buf.Bytes()); err != nil {
		return fmt.Errorf("write %s: %w", name, err)
	}
	return nil
}

type zipWriter interface {
	Create(name string) (io.Writer, error)
}

// nextRelationshipIDs allocates count new rId strings that do not collide with
// any ID in existing or alreadyUsed.
func nextRelationshipIDs(existing []OpaqueRel, count int, alreadyUsed ...string) []string {
	if count == 0 {
		return nil
	}
	used := make(map[string]struct{}, len(existing)+len(alreadyUsed)+count)
	maxRID := 0
	for _, rel := range existing {
		used[rel.ID] = struct{}{}
		if n, ok := parseRelationshipID(rel.ID); ok && n > maxRID {
			maxRID = n
		}
	}
	for _, id := range alreadyUsed {
		used[id] = struct{}{}
		if n, ok := parseRelationshipID(id); ok && n > maxRID {
			maxRID = n
		}
	}

	ids := make([]string, 0, count)
	next := 1
	if maxRID > 0 {
		next = maxRID + 1
	}
	for len(ids) < count {
		id := fmt.Sprintf("rId%d", next)
		next++
		if _, exists := used[id]; exists {
			continue
		}
		used[id] = struct{}{}
		ids = append(ids, id)
	}
	return ids
}

func parseRelationshipID(id string) (int, bool) {
	if !strings.HasPrefix(id, "rId") || len(id) <= len("rId") {
		return 0, false
	}
	n := 0
	for _, ch := range id[len("rId"):] {
		if ch < '0' || ch > '9' {
			return 0, false
		}
		n = n*10 + int(ch-'0')
	}
	return n, true
}
