package ooxml

import (
	"encoding/xml"
	"math"
	"strconv"
	"strings"
)

// FreezePane is the frozen area of a worksheet view: Rows rows at the top
// and Cols columns at the left stay in view. The zero value is "not frozen".
type FreezePane struct {
	Rows int
	Cols int
}

// rawSheetViews is the minimal view of a preserved <sheetViews> element that
// is necessary to read and to change the frozen pane of its first view.
type rawSheetViews struct {
	Views []rawSheetView `xml:"sheetView"`
}

type rawSheetView struct {
	Attrs []xml.Attr `xml:",any,attr"`
	Pane  *rawPane   `xml:"pane"`
}

type rawPane struct {
	XSplit string `xml:"xSplit,attr"`
	YSplit string `xml:"ySplit,attr"`
	State  string `xml:"state,attr"`
}

// FrozenPane returns the frozen pane of the first sheet view in the
// preserved worksheet elements. Only a pane with the state "frozen" or
// "frozenSplit" counts: a split that is not frozen gives the zero value.
func FrozenPane(extras []RawElement) FreezePane {
	views, ok := parseSheetViews(extras)
	if !ok || len(views.Views) == 0 {
		return FreezePane{}
	}
	return frozenPaneOf(views.Views[0])
}

func parseSheetViews(extras []RawElement) (rawSheetViews, bool) {
	for _, extra := range extras {
		if extra.Name != "sheetViews" {
			continue
		}
		var views rawSheetViews
		if err := xml.Unmarshal([]byte(extra.XML), &views); err != nil {
			return rawSheetViews{}, false
		}
		return views, true
	}
	return rawSheetViews{}, false
}

func frozenPaneOf(view rawSheetView) FreezePane {
	if view.Pane == nil || (view.Pane.State != "frozen" && view.Pane.State != "frozenSplit") {
		return FreezePane{}
	}
	return FreezePane{Rows: splitCount(view.Pane.YSplit), Cols: splitCount(view.Pane.XSplit)}
}

// splitCount reads the xSplit or ySplit value of a frozen pane, which is a
// number of columns or rows. The attribute type is a double.
func splitCount(s string) int {
	if s == "" {
		return 0
	}
	f, err := strconv.ParseFloat(s, 64)
	if err != nil || math.IsNaN(f) || f < 0 || f > 1<<20 {
		return 0
	}
	return int(f)
}

// sheetViewsFragment returns the <sheetViews> XML to write for a sheet.
//
// freeze nil means that the caller does not model the frozen pane: the
// preserved element is written as it is. Else a preserved element is written
// as it is when its frozen pane already agrees with freeze (a split that is
// not frozen agrees with "not frozen", so it survives). When they do not
// agree, only the pane of the first sheet view changes: the view keeps its
// attributes, the selection of the pane that is active with the new freeze,
// and its other children; the selections of panes that the new layout does
// not have are dropped. All other views are written as they are. With no
// preserved element, a minimal element is made for a frozen pane.
//
// It returns nil when no element is necessary. This follows the pattern of
// workbookPrFragment.
func sheetViewsFragment(extras []RawElement, freeze *FreezePane) []byte {
	var raw *RawElement
	for i := range extras {
		if extras[i].Name == "sheetViews" {
			raw = &extras[i]
			break
		}
	}
	if freeze == nil {
		if raw == nil {
			return nil
		}
		return []byte(raw.XML)
	}

	// The parts of the preserved element that the new element keeps: the
	// attributes and the children of the first view, and all other views as
	// they are.
	var (
		viewsAttrs []RawAttr    // attributes of <sheetViews>
		firstAttrs []RawAttr    // attributes of the first <sheetView>
		firstKids  []RawElement // children of the first <sheetView>
		otherViews []RawElement // the views after the first, and other children
	)
	if raw != nil {
		var views rawSheetViews
		if err := xml.Unmarshal([]byte(raw.XML), &views); err != nil {
			// An element that cannot be read cannot be changed safely.
			return []byte(raw.XML)
		}
		if len(views.Views) > 0 && frozenPaneOf(views.Views[0]) == *freeze {
			return []byte(raw.XML)
		}
		all := func(string) bool { return false }
		noOrder := func(string, int) int { return 0 }
		attrs, kids, err := captureRootAttrsAndExtras([]byte(raw.XML), all, noOrder)
		if err != nil {
			return []byte(raw.XML)
		}
		viewsAttrs = attrs
		gotFirst := false
		for _, kid := range kids {
			if kid.Name != "sheetView" || gotFirst {
				otherViews = append(otherViews, kid)
				continue
			}
			gotFirst = true
			a, k, err := captureRootAttrsAndExtras([]byte(kid.XML), all, noOrder)
			if err != nil {
				return []byte(raw.XML)
			}
			firstAttrs, firstKids = a, k
		}
	} else if *freeze == (FreezePane{}) {
		return nil
	}

	// The pane that a selection without a pane attribute belongs to is
	// "topLeft", the only pane of a view that is not frozen.
	active := ""
	switch {
	case freeze.Rows > 0 && freeze.Cols > 0:
		active = "bottomRight"
	case freeze.Rows > 0:
		active = "bottomLeft"
	case freeze.Cols > 0:
		active = "topRight"
	}

	var b strings.Builder
	b.WriteString("<sheetViews")
	for _, attr := range viewsAttrs {
		writeXMLAttr(&b, attr.Name, attr.Value)
	}
	b.WriteString("><sheetView")
	hasViewID := false
	for _, attr := range firstAttrs {
		if attr.Name == "workbookViewId" {
			hasViewID = true
		}
		writeXMLAttr(&b, attr.Name, attr.Value)
	}
	if !hasViewID {
		writeXMLAttr(&b, "workbookViewId", "0")
	}

	// Children of the first view, in the order of the schema: pane,
	// selection, then the others. The old pane goes. A selection stays when
	// it is the selection of the pane that is active now; the selections of
	// panes that do not exist any more go.
	var body strings.Builder
	if active != "" {
		body.WriteString("<pane")
		if freeze.Cols > 0 {
			writeXMLAttr(&body, "xSplit", strconv.Itoa(freeze.Cols))
		}
		if freeze.Rows > 0 {
			writeXMLAttr(&body, "ySplit", strconv.Itoa(freeze.Rows))
		}
		writeXMLAttr(&body, "topLeftCell", columnLetters(freeze.Cols+1)+strconv.Itoa(freeze.Rows+1))
		writeXMLAttr(&body, "activePane", active)
		writeXMLAttr(&body, "state", "frozen")
		body.WriteString("/>")
	}
	for _, kid := range firstKids {
		if kid.Name == "selection" && selectionPane(kid.XML) == active {
			body.WriteString(kid.XML)
		}
	}
	for _, kid := range firstKids {
		if kid.Name != "pane" && kid.Name != "selection" {
			body.WriteString(kid.XML)
		}
	}

	if body.Len() == 0 {
		b.WriteString("/>")
	} else {
		b.WriteByte('>')
		b.WriteString(body.String())
		b.WriteString("</sheetView>")
	}
	for _, view := range otherViews {
		b.WriteString(view.XML)
	}
	b.WriteString("</sheetViews>")
	return []byte(b.String())
}

// selectionPane returns the pane of a <selection> element. A selection with
// no pane attribute is for the top-left pane, which is "" here: the only
// pane of a view that is not frozen.
func selectionPane(raw string) string {
	var sel struct {
		Pane string `xml:"pane,attr"`
	}
	if err := xml.Unmarshal([]byte(raw), &sel); err != nil {
		return "?"
	}
	if sel.Pane == "topLeft" {
		return ""
	}
	return sel.Pane
}

func writeXMLAttr(b *strings.Builder, name, value string) {
	b.WriteByte(' ')
	b.WriteString(name)
	b.WriteString(`="`)
	_ = xml.EscapeText(b, []byte(value))
	b.WriteByte('"')
}

// columnLetters converts a 1-based column number to its letters (1 = A).
func columnLetters(col int) string {
	var out []byte
	for col > 0 {
		col--
		out = append([]byte{byte('A' + col%26)}, out...)
		col /= 26
	}
	return string(out)
}
