package werkbook

import (
	"fmt"

	"github.com/jpoz/werkbook/ooxml"
)

// Style defines visual formatting for a cell.
type Style struct {
	Font      *Font
	Fill      *Fill
	Border    *Border
	Alignment *Alignment
	NumFmt    string // custom format string e.g. "#,##0.00"
	NumFmtID  int    // built-in ID (0-163), ignored when NumFmt is set
}

// Font describes the font properties for a cell.
type Font struct {
	Name      string  // e.g. "Calibri"
	Size      float64 // e.g. 11
	Bold      bool
	Italic    bool
	Underline bool
	Color     string // 6-char RGB hex e.g. "FF0000"
}

// Fill describes a solid fill for a cell.
type Fill struct {
	Color string // 6-char RGB hex for solid fill
}

// Border describes borders for a cell.
type Border struct {
	Left   BorderSide
	Right  BorderSide
	Top    BorderSide
	Bottom BorderSide
}

// BorderSide describes one side of a cell border.
type BorderSide struct {
	Style BorderStyle
	Color string // 6-char RGB hex
}

// BorderStyle represents an OOXML border style.
type BorderStyle int

const (
	BorderNone BorderStyle = iota
	BorderThin
	BorderMedium
	BorderThick
	BorderDashed
	BorderDotted
	BorderDouble
)

var borderStyleToOOXML = map[BorderStyle]string{
	BorderNone:   "",
	BorderThin:   "thin",
	BorderMedium: "medium",
	BorderThick:  "thick",
	BorderDashed: "dashed",
	BorderDotted: "dotted",
	BorderDouble: "double",
}

var ooxmlToBorderStyle = map[string]BorderStyle{
	"":       BorderNone,
	"thin":   BorderThin,
	"medium": BorderMedium,
	"thick":  BorderThick,
	"dashed": BorderDashed,
	"dotted": BorderDotted,
	"double": BorderDouble,
}

// HorizontalAlign represents horizontal alignment of cell content.
type HorizontalAlign int

const (
	HAlignGeneral HorizontalAlign = iota
	HAlignLeft
	HAlignCenter
	HAlignRight
)

var hAlignToOOXML = map[HorizontalAlign]string{
	HAlignGeneral: "",
	HAlignLeft:    "left",
	HAlignCenter:  "center",
	HAlignRight:   "right",
}

var ooxmlToHAlign = map[string]HorizontalAlign{
	"":        HAlignGeneral,
	"general": HAlignGeneral,
	"left":    HAlignLeft,
	"center":  HAlignCenter,
	"right":   HAlignRight,
}

// VerticalAlign represents vertical alignment of cell content.
type VerticalAlign int

const (
	VAlignBottom VerticalAlign = iota
	VAlignCenter
	VAlignTop
)

var vAlignToOOXML = map[VerticalAlign]string{
	VAlignBottom: "",
	VAlignCenter: "center",
	VAlignTop:    "top",
}

var ooxmlToVAlign = map[string]VerticalAlign{
	"":       VAlignBottom,
	"bottom": VAlignBottom,
	"center": VAlignCenter,
	"top":    VAlignTop,
}

// Alignment describes the alignment of cell content.
type Alignment struct {
	Horizontal HorizontalAlign
	Vertical   VerticalAlign
	WrapText   bool
}

// rgbToARGB converts a 6-char RGB hex to 8-char ARGB hex (prepends "FF").
func rgbToARGB(rgb string) string {
	if rgb == "" {
		return ""
	}
	return "FF" + rgb
}

// argbToRGB converts an 8-char ARGB hex to 6-char RGB hex (strips alpha).
func argbToRGB(argb string) string {
	if len(argb) == 8 {
		return argb[2:]
	}
	return argb
}

func borderSideToOOXML(bs BorderSide) (style, color string) {
	return borderStyleToOOXML[bs.Style], rgbToARGB(bs.Color)
}

func ooxmlToBorderSide(style, color string) BorderSide {
	bs, ok := ooxmlToBorderStyle[style]
	if !ok {
		bs = BorderNone
	}
	return BorderSide{Style: bs, Color: argbToRGB(color)}
}

// styleToStyleData converts a public Style to an ooxml.StyleData.
func styleToStyleData(s *Style) ooxml.StyleData {
	if s == nil {
		return ooxml.StyleData{}
	}
	var sd ooxml.StyleData
	if s.Font != nil {
		sd.FontName = s.Font.Name
		sd.FontSize = s.Font.Size
		sd.FontBold = s.Font.Bold
		sd.FontItalic = s.Font.Italic
		sd.FontUL = s.Font.Underline
		sd.FontColor = rgbToARGB(s.Font.Color)
	}
	if s.Fill != nil {
		sd.FillColor = rgbToARGB(s.Fill.Color)
	}
	if s.Border != nil {
		sd.BorderLeftStyle, sd.BorderLeftColor = borderSideToOOXML(s.Border.Left)
		sd.BorderRightStyle, sd.BorderRightColor = borderSideToOOXML(s.Border.Right)
		sd.BorderTopStyle, sd.BorderTopColor = borderSideToOOXML(s.Border.Top)
		sd.BorderBottomStyle, sd.BorderBottomColor = borderSideToOOXML(s.Border.Bottom)
	}
	if s.Alignment != nil {
		sd.HAlign = hAlignToOOXML[s.Alignment.Horizontal]
		sd.VAlign = vAlignToOOXML[s.Alignment.Vertical]
		sd.WrapText = s.Alignment.WrapText
	}
	if s.NumFmt != "" {
		sd.NumFmt = s.NumFmt
	} else if s.NumFmtID != 0 {
		sd.NumFmtID = s.NumFmtID
	}
	return sd
}

// styleDataToStyle converts an ooxml.StyleData to a public Style.
// Returns nil for the zero-value StyleData (default style).
func styleDataToStyle(sd ooxml.StyleData) *Style {
	if sd == (ooxml.StyleData{}) {
		return nil
	}
	s := &Style{}

	hasFont := sd.FontName != "" || sd.FontSize != 0 || sd.FontBold || sd.FontItalic || sd.FontUL || sd.FontColor != ""
	if hasFont {
		s.Font = &Font{
			Name:      sd.FontName,
			Size:      sd.FontSize,
			Bold:      sd.FontBold,
			Italic:    sd.FontItalic,
			Underline: sd.FontUL,
			Color:     argbToRGB(sd.FontColor),
		}
	}

	if sd.FillColor != "" {
		s.Fill = &Fill{Color: argbToRGB(sd.FillColor)}
	}

	hasBorder := sd.BorderLeftStyle != "" || sd.BorderRightStyle != "" ||
		sd.BorderTopStyle != "" || sd.BorderBottomStyle != "" ||
		sd.BorderLeftColor != "" || sd.BorderRightColor != "" ||
		sd.BorderTopColor != "" || sd.BorderBottomColor != ""
	if hasBorder {
		s.Border = &Border{
			Left:   ooxmlToBorderSide(sd.BorderLeftStyle, sd.BorderLeftColor),
			Right:  ooxmlToBorderSide(sd.BorderRightStyle, sd.BorderRightColor),
			Top:    ooxmlToBorderSide(sd.BorderTopStyle, sd.BorderTopColor),
			Bottom: ooxmlToBorderSide(sd.BorderBottomStyle, sd.BorderBottomColor),
		}
	}

	hasAlign := sd.HAlign != "" || sd.VAlign != "" || sd.WrapText
	if hasAlign {
		s.Alignment = &Alignment{
			Horizontal: ooxmlToHAlign[sd.HAlign],
			Vertical:   ooxmlToVAlign[sd.VAlign],
			WrapText:   sd.WrapText,
		}
	}

	if sd.NumFmt != "" {
		s.NumFmt = sd.NumFmt
	} else if sd.NumFmtID != 0 {
		s.NumFmtID = sd.NumFmtID
	}

	return s
}

// styleKey returns a string key for deduplicating styles.
func styleKey(sd ooxml.StyleData) string {
	return fmt.Sprintf("%+v", sd)
}

func cloneStyle(s *Style) *Style {
	if s == nil {
		return nil
	}

	clone := &Style{
		NumFmt:   s.NumFmt,
		NumFmtID: s.NumFmtID,
	}

	if s.Font != nil {
		font := *s.Font
		clone.Font = &font
	}
	if s.Fill != nil {
		fill := *s.Fill
		clone.Fill = &fill
	}
	if s.Border != nil {
		border := *s.Border
		clone.Border = &border
	}
	if s.Alignment != nil {
		alignment := *s.Alignment
		clone.Alignment = &alignment
	}

	return clone
}
