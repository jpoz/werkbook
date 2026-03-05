package main

import (
	"encoding/json"
	"fmt"
	"strings"

	werkbook "github.com/jpoz/werkbook"
)

// styleJSON is the JSON representation of a Style for input/output.
type styleJSON struct {
	Font      *fontJSON      `json:"font,omitempty"`
	Fill      *fillJSON      `json:"fill,omitempty"`
	Border    *borderJSON    `json:"border,omitempty"`
	Alignment *alignmentJSON `json:"alignment,omitempty"`
	NumFmt    string         `json:"num_fmt,omitempty"`
	NumFmtID  int            `json:"num_fmt_id,omitempty"`
}

type fontJSON struct {
	Name      string  `json:"name,omitempty"`
	Size      float64 `json:"size,omitempty"`
	Bold      bool    `json:"bold,omitempty"`
	Italic    bool    `json:"italic,omitempty"`
	Underline bool    `json:"underline,omitempty"`
	Color     string  `json:"color,omitempty"`
}

type fillJSON struct {
	Color string `json:"color,omitempty"`
}

type borderJSON struct {
	Left   *borderSideJSON `json:"left,omitempty"`
	Right  *borderSideJSON `json:"right,omitempty"`
	Top    *borderSideJSON `json:"top,omitempty"`
	Bottom *borderSideJSON `json:"bottom,omitempty"`
}

type borderSideJSON struct {
	Style string `json:"style,omitempty"`
	Color string `json:"color,omitempty"`
}

type alignmentJSON struct {
	Horizontal string `json:"horizontal,omitempty"`
	Vertical   string `json:"vertical,omitempty"`
	WrapText   bool   `json:"wrap_text,omitempty"`
}

// styleToJSON converts a werkbook.Style to a JSON-friendly map.
func styleToJSON(s *werkbook.Style) *styleJSON {
	if s == nil {
		return nil
	}
	sj := &styleJSON{
		NumFmt:   s.NumFmt,
		NumFmtID: s.NumFmtID,
	}
	if s.Font != nil {
		sj.Font = &fontJSON{
			Name:      s.Font.Name,
			Size:      s.Font.Size,
			Bold:      s.Font.Bold,
			Italic:    s.Font.Italic,
			Underline: s.Font.Underline,
			Color:     s.Font.Color,
		}
	}
	if s.Fill != nil {
		sj.Fill = &fillJSON{Color: s.Fill.Color}
	}
	if s.Border != nil {
		sj.Border = &borderJSON{
			Left:   borderSideToJSON(s.Border.Left),
			Right:  borderSideToJSON(s.Border.Right),
			Top:    borderSideToJSON(s.Border.Top),
			Bottom: borderSideToJSON(s.Border.Bottom),
		}
	}
	if s.Alignment != nil {
		sj.Alignment = &alignmentJSON{
			Horizontal: hAlignString(s.Alignment.Horizontal),
			Vertical:   vAlignString(s.Alignment.Vertical),
			WrapText:   s.Alignment.WrapText,
		}
	}
	return sj
}

// jsonToStyle converts JSON bytes into a werkbook.Style.
func jsonToStyle(data json.RawMessage) (*werkbook.Style, error) {
	var sj styleJSON
	if err := strictUnmarshal(data, &sj); err != nil {
		return nil, err
	}
	s := &werkbook.Style{
		NumFmt:   sj.NumFmt,
		NumFmtID: sj.NumFmtID,
	}
	if sj.Font != nil {
		s.Font = &werkbook.Font{
			Name:      sj.Font.Name,
			Size:      sj.Font.Size,
			Bold:      sj.Font.Bold,
			Italic:    sj.Font.Italic,
			Underline: sj.Font.Underline,
			Color:     sj.Font.Color,
		}
	}
	if sj.Fill != nil {
		s.Fill = &werkbook.Fill{Color: sj.Fill.Color}
	}
	if sj.Border != nil {
		s.Border = &werkbook.Border{
			Left:   jsonToBorderSide(sj.Border.Left),
			Right:  jsonToBorderSide(sj.Border.Right),
			Top:    jsonToBorderSide(sj.Border.Top),
			Bottom: jsonToBorderSide(sj.Border.Bottom),
		}
	}
	if sj.Alignment != nil {
		s.Alignment = &werkbook.Alignment{
			Horizontal: parseHAlign(sj.Alignment.Horizontal),
			Vertical:   parseVAlign(sj.Alignment.Vertical),
			WrapText:   sj.Alignment.WrapText,
		}
	}
	return s, nil
}

func borderSideToJSON(bs werkbook.BorderSide) *borderSideJSON {
	if bs.Style == werkbook.BorderNone && bs.Color == "" {
		return nil
	}
	return &borderSideJSON{
		Style: borderStyleString(bs.Style),
		Color: bs.Color,
	}
}

func jsonToBorderSide(bsj *borderSideJSON) werkbook.BorderSide {
	if bsj == nil {
		return werkbook.BorderSide{}
	}
	return werkbook.BorderSide{
		Style: parseBorderStyle(bsj.Style),
		Color: bsj.Color,
	}
}

var borderStyleMap = map[string]werkbook.BorderStyle{
	"":       werkbook.BorderNone,
	"thin":   werkbook.BorderThin,
	"medium": werkbook.BorderMedium,
	"thick":  werkbook.BorderThick,
	"dashed": werkbook.BorderDashed,
	"dotted": werkbook.BorderDotted,
	"double": werkbook.BorderDouble,
}

var borderStyleNames = map[werkbook.BorderStyle]string{
	werkbook.BorderNone:   "",
	werkbook.BorderThin:   "thin",
	werkbook.BorderMedium: "medium",
	werkbook.BorderThick:  "thick",
	werkbook.BorderDashed: "dashed",
	werkbook.BorderDotted: "dotted",
	werkbook.BorderDouble: "double",
}

func parseBorderStyle(s string) werkbook.BorderStyle {
	if bs, ok := borderStyleMap[strings.ToLower(s)]; ok {
		return bs
	}
	return werkbook.BorderNone
}

func borderStyleString(bs werkbook.BorderStyle) string {
	return borderStyleNames[bs]
}

func hAlignString(h werkbook.HorizontalAlign) string {
	switch h {
	case werkbook.HAlignLeft:
		return "left"
	case werkbook.HAlignCenter:
		return "center"
	case werkbook.HAlignRight:
		return "right"
	default:
		return ""
	}
}

func parseHAlign(s string) werkbook.HorizontalAlign {
	switch strings.ToLower(s) {
	case "left":
		return werkbook.HAlignLeft
	case "center":
		return werkbook.HAlignCenter
	case "right":
		return werkbook.HAlignRight
	default:
		return werkbook.HAlignGeneral
	}
}

func vAlignString(v werkbook.VerticalAlign) string {
	switch v {
	case werkbook.VAlignCenter:
		return "center"
	case werkbook.VAlignTop:
		return "top"
	default:
		return ""
	}
}

func parseVAlign(s string) werkbook.VerticalAlign {
	switch strings.ToLower(s) {
	case "center":
		return werkbook.VAlignCenter
	case "top":
		return werkbook.VAlignTop
	default:
		return werkbook.VAlignBottom
	}
}

// styleSummary returns a short human-readable summary of a style, e.g. "bold, fill:#FF0000, 12pt".
func styleSummary(s *werkbook.Style) string {
	if s == nil {
		return ""
	}
	var parts []string
	if s.Font != nil {
		if s.Font.Bold {
			parts = append(parts, "bold")
		}
		if s.Font.Italic {
			parts = append(parts, "italic")
		}
		if s.Font.Underline {
			parts = append(parts, "underline")
		}
		if s.Font.Size > 0 {
			parts = append(parts, fmt.Sprintf("%gpt", s.Font.Size))
		}
		if s.Font.Name != "" {
			parts = append(parts, s.Font.Name)
		}
		if s.Font.Color != "" {
			parts = append(parts, "color:#"+s.Font.Color)
		}
	}
	if s.Fill != nil && s.Fill.Color != "" {
		parts = append(parts, "fill:#"+s.Fill.Color)
	}
	if s.NumFmt != "" {
		parts = append(parts, "fmt:"+s.NumFmt)
	}
	if s.Alignment != nil {
		if h := hAlignString(s.Alignment.Horizontal); h != "" {
			parts = append(parts, "align:"+h)
		}
		if s.Alignment.WrapText {
			parts = append(parts, "wrap")
		}
	}
	return strings.Join(parts, ", ")
}
