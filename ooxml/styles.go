package ooxml

import (
	"encoding/xml"
	"fmt"
)

// XML struct types for xl/styles.xml

type xlsxStyleSheet struct {
	XMLName       xml.Name         `xml:"styleSheet"`
	Xmlns         string           `xml:"xmlns,attr"`
	NumFmts       *xlsxNumFmts     `xml:"numFmts,omitempty"`
	Fonts         xlsxFonts        `xml:"fonts"`
	Fills         xlsxFills        `xml:"fills"`
	Borders       xlsxBorders      `xml:"borders"`
	CellStyleXfs  xlsxCellStyleXfs `xml:"cellStyleXfs"`
	CellXfs       xlsxCellXfs      `xml:"cellXfs"`
}

type xlsxNumFmts struct {
	Count  int          `xml:"count,attr"`
	NumFmt []xlsxNumFmt `xml:"numFmt"`
}

type xlsxNumFmt struct {
	NumFmtID   int    `xml:"numFmtId,attr"`
	FormatCode string `xml:"formatCode,attr"`
}

type xlsxFonts struct {
	Count int        `xml:"count,attr"`
	Font  []xlsxFont `xml:"font"`
}

type xlsxFont struct {
	B     *xlsxVal  `xml:"b,omitempty"`
	I     *xlsxVal  `xml:"i,omitempty"`
	U     *xlsxVal  `xml:"u,omitempty"`
	Sz    *xlsxSz   `xml:"sz,omitempty"`
	Color *xlsxColor `xml:"color,omitempty"`
	Name  *xlsxName `xml:"name,omitempty"`
}

type xlsxVal struct{}

type xlsxSz struct {
	Val float64 `xml:"val,attr"`
}

type xlsxColor struct {
	RGB string `xml:"rgb,attr,omitempty"`
}

type xlsxName struct {
	Val string `xml:"val,attr"`
}

type xlsxFills struct {
	Count int        `xml:"count,attr"`
	Fill  []xlsxFill `xml:"fill"`
}

type xlsxFill struct {
	PatternFill xlsxPatternFill `xml:"patternFill"`
}

type xlsxPatternFill struct {
	PatternType string     `xml:"patternType,attr"`
	FgColor     *xlsxColor `xml:"fgColor,omitempty"`
}

type xlsxBorders struct {
	Count  int           `xml:"count,attr"`
	Border []xlsxBorderE `xml:"border"`
}

type xlsxBorderE struct {
	Left     xlsxBorderSide `xml:"left"`
	Right    xlsxBorderSide `xml:"right"`
	Top      xlsxBorderSide `xml:"top"`
	Bottom   xlsxBorderSide `xml:"bottom"`
	Diagonal xlsxBorderSide `xml:"diagonal"`
}

type xlsxBorderSide struct {
	Style string     `xml:"style,attr,omitempty"`
	Color *xlsxColor `xml:"color,omitempty"`
}

type xlsxCellStyleXfs struct {
	Count int      `xml:"count,attr"`
	Xf    []xlsxXf `xml:"xf"`
}

type xlsxCellXfs struct {
	Count int      `xml:"count,attr"`
	Xf    []xlsxXf `xml:"xf"`
}

type xlsxXf struct {
	NumFmtID        int             `xml:"numFmtId,attr"`
	FontID          int             `xml:"fontId,attr"`
	FillID          int             `xml:"fillId,attr"`
	BorderID        int             `xml:"borderId,attr"`
	XfID            int             `xml:"xfId,attr,omitempty"`
	ApplyFont       bool            `xml:"applyFont,attr,omitempty"`
	ApplyFill       bool            `xml:"applyFill,attr,omitempty"`
	ApplyBorder     bool            `xml:"applyBorder,attr,omitempty"`
	ApplyAlignment  bool            `xml:"applyAlignment,attr,omitempty"`
	ApplyNumberFormat bool          `xml:"applyNumberFormat,attr,omitempty"`
	Alignment       *xlsxAlignment  `xml:"alignment,omitempty"`
}

type xlsxAlignment struct {
	Horizontal string `xml:"horizontal,attr,omitempty"`
	Vertical   string `xml:"vertical,attr,omitempty"`
	WrapText   bool   `xml:"wrapText,attr,omitempty"`
}

// StyleSheetBuilder builds a deduplicated OOXML style sheet.
type StyleSheetBuilder struct {
	fonts    []xlsxFont
	fontKeys map[string]int

	fills    []xlsxFill
	fillKeys map[string]int

	borders    []xlsxBorderE
	borderKeys map[string]int

	numFmts    []xlsxNumFmt
	numFmtKeys map[string]int
	nextFmtID  int

	xfs    []xlsxXf
	xfKeys map[string]int
}

// NewStyleSheetBuilder creates a builder pre-seeded with the required defaults.
func NewStyleSheetBuilder() *StyleSheetBuilder {
	ssb := &StyleSheetBuilder{
		fontKeys:   make(map[string]int),
		fillKeys:   make(map[string]int),
		borderKeys: make(map[string]int),
		numFmtKeys: make(map[string]int),
		xfKeys:     make(map[string]int),
		nextFmtID:  164, // custom IDs start at 164
	}

	// Default font (Calibri 11pt)
	ssb.fonts = append(ssb.fonts, xlsxFont{
		Sz:   &xlsxSz{Val: 11},
		Name: &xlsxName{Val: "Calibri"},
	})
	ssb.fontKeys[fontKey(ssb.fonts[0])] = 0

	// Required fills: none + gray125
	ssb.fills = append(ssb.fills, xlsxFill{PatternFill: xlsxPatternFill{PatternType: "none"}})
	ssb.fills = append(ssb.fills, xlsxFill{PatternFill: xlsxPatternFill{PatternType: "gray125"}})
	ssb.fillKeys[fillKey(ssb.fills[0])] = 0
	ssb.fillKeys[fillKey(ssb.fills[1])] = 1

	// Default border (empty)
	ssb.borders = append(ssb.borders, xlsxBorderE{})
	ssb.borderKeys[borderKey(ssb.borders[0])] = 0

	// Default xf (cellXfs index 0)
	defaultXf := xlsxXf{NumFmtID: 0, FontID: 0, FillID: 0, BorderID: 0, XfID: 0}
	ssb.xfs = append(ssb.xfs, defaultXf)
	ssb.xfKeys[xfKey(defaultXf)] = 0

	return ssb
}

// AddStyle converts a StyleData into deduped font/fill/border/numfmt/xf entries
// and returns the cellXfs index.
func (ssb *StyleSheetBuilder) AddStyle(sd StyleData) int {
	fontID := ssb.addFont(sd)
	fillID := ssb.addFill(sd)
	borderID := ssb.addBorder(sd)
	numFmtID := ssb.addNumFmt(sd)

	xf := xlsxXf{
		NumFmtID: numFmtID,
		FontID:   fontID,
		FillID:   fillID,
		BorderID: borderID,
		XfID:     0,
	}
	if fontID != 0 {
		xf.ApplyFont = true
	}
	if fillID > 1 { // 0=none, 1=gray125 are defaults
		xf.ApplyFill = true
	}
	if borderID != 0 {
		xf.ApplyBorder = true
	}
	if numFmtID != 0 {
		xf.ApplyNumberFormat = true
	}
	if sd.HAlign != "" || sd.VAlign != "" || sd.WrapText {
		xf.ApplyAlignment = true
		xf.Alignment = &xlsxAlignment{
			Horizontal: sd.HAlign,
			Vertical:   sd.VAlign,
			WrapText:   sd.WrapText,
		}
	}

	key := xfKey(xf)
	if idx, ok := ssb.xfKeys[key]; ok {
		return idx
	}
	idx := len(ssb.xfs)
	ssb.xfs = append(ssb.xfs, xf)
	ssb.xfKeys[key] = idx
	return idx
}

func (ssb *StyleSheetBuilder) addFont(sd StyleData) int {
	f := xlsxFont{}
	if sd.FontSize != 0 {
		f.Sz = &xlsxSz{Val: sd.FontSize}
	}
	if sd.FontName != "" {
		f.Name = &xlsxName{Val: sd.FontName}
	}
	if sd.FontBold {
		f.B = &xlsxVal{}
	}
	if sd.FontItalic {
		f.I = &xlsxVal{}
	}
	if sd.FontUL {
		f.U = &xlsxVal{}
	}
	if sd.FontColor != "" {
		f.Color = &xlsxColor{RGB: sd.FontColor}
	}

	// If it matches the default font exactly, return 0.
	key := fontKey(f)
	if idx, ok := ssb.fontKeys[key]; ok {
		return idx
	}
	idx := len(ssb.fonts)
	ssb.fonts = append(ssb.fonts, f)
	ssb.fontKeys[key] = idx
	return idx
}

func (ssb *StyleSheetBuilder) addFill(sd StyleData) int {
	if sd.FillColor == "" {
		return 0 // default (none)
	}
	f := xlsxFill{PatternFill: xlsxPatternFill{
		PatternType: "solid",
		FgColor:     &xlsxColor{RGB: sd.FillColor},
	}}
	key := fillKey(f)
	if idx, ok := ssb.fillKeys[key]; ok {
		return idx
	}
	idx := len(ssb.fills)
	ssb.fills = append(ssb.fills, f)
	ssb.fillKeys[key] = idx
	return idx
}

func (ssb *StyleSheetBuilder) addBorder(sd StyleData) int {
	b := xlsxBorderE{}
	if sd.BorderLeftStyle != "" {
		b.Left = xlsxBorderSide{Style: sd.BorderLeftStyle}
		if sd.BorderLeftColor != "" {
			b.Left.Color = &xlsxColor{RGB: sd.BorderLeftColor}
		}
	}
	if sd.BorderRightStyle != "" {
		b.Right = xlsxBorderSide{Style: sd.BorderRightStyle}
		if sd.BorderRightColor != "" {
			b.Right.Color = &xlsxColor{RGB: sd.BorderRightColor}
		}
	}
	if sd.BorderTopStyle != "" {
		b.Top = xlsxBorderSide{Style: sd.BorderTopStyle}
		if sd.BorderTopColor != "" {
			b.Top.Color = &xlsxColor{RGB: sd.BorderTopColor}
		}
	}
	if sd.BorderBottomStyle != "" {
		b.Bottom = xlsxBorderSide{Style: sd.BorderBottomStyle}
		if sd.BorderBottomColor != "" {
			b.Bottom.Color = &xlsxColor{RGB: sd.BorderBottomColor}
		}
	}

	key := borderKey(b)
	if idx, ok := ssb.borderKeys[key]; ok {
		return idx
	}
	idx := len(ssb.borders)
	ssb.borders = append(ssb.borders, b)
	ssb.borderKeys[key] = idx
	return idx
}

func (ssb *StyleSheetBuilder) addNumFmt(sd StyleData) int {
	if sd.NumFmt != "" {
		key := sd.NumFmt
		if idx, ok := ssb.numFmtKeys[key]; ok {
			return idx
		}
		id := ssb.nextFmtID
		ssb.nextFmtID++
		ssb.numFmts = append(ssb.numFmts, xlsxNumFmt{NumFmtID: id, FormatCode: sd.NumFmt})
		ssb.numFmtKeys[key] = id
		return id
	}
	return sd.NumFmtID
}

// Build returns the serializable xlsxStyleSheet.
func (ssb *StyleSheetBuilder) Build() xlsxStyleSheet {
	ss := xlsxStyleSheet{
		Xmlns: NSSpreadsheetML,
		Fonts: xlsxFonts{Count: len(ssb.fonts), Font: ssb.fonts},
		Fills: xlsxFills{Count: len(ssb.fills), Fill: ssb.fills},
		Borders: xlsxBorders{Count: len(ssb.borders), Border: ssb.borders},
		CellStyleXfs: xlsxCellStyleXfs{
			Count: 1,
			Xf:    []xlsxXf{{NumFmtID: 0, FontID: 0, FillID: 0, BorderID: 0}},
		},
		CellXfs: xlsxCellXfs{Count: len(ssb.xfs), Xf: ssb.xfs},
	}
	if len(ssb.numFmts) > 0 {
		ss.NumFmts = &xlsxNumFmts{Count: len(ssb.numFmts), NumFmt: ssb.numFmts}
	}
	return ss
}

// Dedup key functions

func fontKey(f xlsxFont) string {
	var name string
	var size float64
	var color string
	if f.Name != nil {
		name = f.Name.Val
	}
	if f.Sz != nil {
		size = f.Sz.Val
	}
	if f.Color != nil {
		color = f.Color.RGB
	}
	return fmt.Sprintf("f:%s|%g|%v|%v|%v|%s", name, size, f.B != nil, f.I != nil, f.U != nil, color)
}

func fillKey(f xlsxFill) string {
	var fgColor string
	if f.PatternFill.FgColor != nil {
		fgColor = f.PatternFill.FgColor.RGB
	}
	return fmt.Sprintf("fl:%s|%s", f.PatternFill.PatternType, fgColor)
}

func borderKey(b xlsxBorderE) string {
	return fmt.Sprintf("b:%s|%s|%s|%s|%s|%s|%s|%s",
		b.Left.Style, borderColor(b.Left),
		b.Right.Style, borderColor(b.Right),
		b.Top.Style, borderColor(b.Top),
		b.Bottom.Style, borderColor(b.Bottom))
}

func borderColor(bs xlsxBorderSide) string {
	if bs.Color != nil {
		return bs.Color.RGB
	}
	return ""
}

func xfKey(xf xlsxXf) string {
	var ah, av string
	var wt bool
	if xf.Alignment != nil {
		ah = xf.Alignment.Horizontal
		av = xf.Alignment.Vertical
		wt = xf.Alignment.WrapText
	}
	return fmt.Sprintf("xf:%d|%d|%d|%d|%s|%s|%v", xf.NumFmtID, xf.FontID, xf.FillID, xf.BorderID, ah, av, wt)
}
