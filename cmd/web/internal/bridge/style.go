package bridge

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"strings"

	werkbook "github.com/jpoz/werkbook"
)

const labelFormat = "Format"

// Number format presets of the contract.
const (
	numFmtGeneral  = "general"
	numFmtNumber   = "number"
	numFmtCurrency = "currency"
	numFmtPercent  = "percent"
	numFmtDate     = "date"
	numFmtCustom   = "custom"

	currencyFormatCode = `"$"#,##0.00`
)

// numFmtPresets maps a preset to its stored form. A built-in format is
// stored by its id, as Excel does; currency has no built-in id that is the
// same in all locales, so it is a format code.
var numFmtPresets = map[string]struct {
	id   int
	code string
}{
	numFmtGeneral:  {},
	numFmtNumber:   {id: 4}, // #,##0.00
	numFmtCurrency: {code: currencyFormatCode},
	numFmtPercent:  {id: 9},  // 0%
	numFmtDate:     {id: 14}, // m/d/yyyy
}

// numFmtPreset returns the preset name of a style, "custom" for all other
// formats, and "" for the general format.
func numFmtPreset(st *werkbook.Style) string {
	if st == nil {
		return ""
	}
	if st.NumFmt != "" {
		switch {
		case strings.EqualFold(st.NumFmt, "General"):
			return ""
		case st.NumFmt == currencyFormatCode:
			return numFmtCurrency
		case st.NumFmt == "#,##0.00":
			return numFmtNumber
		case st.NumFmt == "0%":
			return numFmtPercent
		case strings.EqualFold(st.NumFmt, "m/d/yyyy"):
			return numFmtDate
		}
		return numFmtCustom
	}
	switch st.NumFmtID {
	case 0:
		return ""
	case 4:
		return numFmtNumber
	case 9:
		return numFmtPercent
	case 14:
		return numFmtDate
	}
	return numFmtCustom
}

// ColorPatch is a color field of a StylePatch. JSON: absent = no change,
// null = remove the color, "#RRGGBB" = set the color.
type ColorPatch struct {
	Set   bool   // the field is in the patch
	Value string // "RRGGBB" in upper case; "" removes the color
}

// UnmarshalJSON reads null or a "#RRGGBB" string.
func (c *ColorPatch) UnmarshalJSON(data []byte) error {
	if bytes.Equal(bytes.TrimSpace(data), []byte("null")) {
		*c = ColorPatch{Set: true}
		return nil
	}
	var text string
	if err := json.Unmarshal(data, &text); err != nil {
		return errors.New(`a color must be a "#RRGGBB" string or null`)
	}
	css := cssColor(strings.TrimPrefix(text, "#"))
	if len(text) != 7 || text[0] != '#' || css == "" {
		return fmt.Errorf(`color %q is not a "#RRGGBB" string`, text)
	}
	*c = ColorPatch{Set: true, Value: css[1:]}
	return nil
}

// StylePatch changes parts of a cell style. Only the fields that are set
// change; all other parts of the style stay.
type StylePatch struct {
	Bold      *bool      `json:"bold"`
	Italic    *bool      `json:"italic"`
	Underline *bool      `json:"underline"`
	Color     ColorPatch `json:"color"`
	Fill      ColorPatch `json:"fill"`
	// Align is "left", "center", "right" or "general".
	Align *string `json:"align"`
	// NumFmt is a preset name: "general", "number", "currency", "percent"
	// or "date".
	NumFmt *string `json:"numFmt"`
}

func (p StylePatch) validate() error {
	if p.Align != nil {
		switch *p.Align {
		case "left", "center", "right", "general":
		default:
			return fmt.Errorf(`align %q is not "left", "center", "right" or "general"`, *p.Align)
		}
	}
	if p.NumFmt != nil {
		if _, ok := numFmtPresets[*p.NumFmt]; !ok {
			return fmt.Errorf(`numFmt %q is not "general", "number", "currency", "percent" or "date"`, *p.NumFmt)
		}
	}
	return nil
}

// apply returns the style that the patch gives for a cell with the style
// old. The result is old itself when the patch changes nothing, nil when the
// result is the default style, and a new Style in all other cases: styles
// are shared between cells and are never changed in place.
func (p StylePatch) apply(old *werkbook.Style) *werkbook.Style {
	st := cloneStyle(old)
	if st == nil {
		st = &werkbook.Style{}
	}

	if p.Bold != nil || p.Italic != nil || p.Underline != nil || p.Color.Set {
		if st.Font == nil {
			st.Font = &werkbook.Font{}
		}
		if p.Bold != nil {
			st.Font.Bold = *p.Bold
		}
		if p.Italic != nil {
			st.Font.Italic = *p.Italic
		}
		if p.Underline != nil {
			st.Font.Underline = *p.Underline
		}
		if p.Color.Set {
			st.Font.Color = p.Color.Value
		}
	}
	if p.Fill.Set {
		if p.Fill.Value == "" {
			st.Fill = nil
		} else {
			st.Fill = &werkbook.Fill{Color: p.Fill.Value}
		}
	}
	if p.Align != nil {
		if st.Alignment == nil {
			st.Alignment = &werkbook.Alignment{}
		}
		switch *p.Align {
		case "left":
			st.Alignment.Horizontal = werkbook.HAlignLeft
		case "center":
			st.Alignment.Horizontal = werkbook.HAlignCenter
		case "right":
			st.Alignment.Horizontal = werkbook.HAlignRight
		default:
			st.Alignment.Horizontal = werkbook.HAlignGeneral
		}
	}
	if p.NumFmt != nil {
		preset := numFmtPresets[*p.NumFmt]
		st.NumFmt, st.NumFmtID = preset.code, preset.id
	}

	// Parts that hold nothing are not parts of the style.
	if st.Font != nil && *st.Font == (werkbook.Font{}) {
		st.Font = nil
	}
	if st.Fill != nil && st.Fill.Color == "" {
		st.Fill = nil
	}
	if st.Alignment != nil && *st.Alignment == (werkbook.Alignment{}) {
		st.Alignment = nil
	}
	if st.Border != nil && *st.Border == (werkbook.Border{}) {
		st.Border = nil
	}

	if stylesEqual(st, old) {
		return old
	}
	if *st == (werkbook.Style{}) {
		return nil
	}
	return st
}

// cloneStyle makes a deep copy of a style.
func cloneStyle(st *werkbook.Style) *werkbook.Style {
	if st == nil {
		return nil
	}
	out := *st
	if st.Font != nil {
		font := *st.Font
		out.Font = &font
	}
	if st.Fill != nil {
		fill := *st.Fill
		out.Fill = &fill
	}
	if st.Border != nil {
		border := *st.Border
		out.Border = &border
	}
	if st.Alignment != nil {
		align := *st.Alignment
		out.Alignment = &align
	}
	return &out
}

// stylesEqual compares two styles by content. A nil style and a style with
// no part are equal.
func stylesEqual(a, b *werkbook.Style) bool {
	var zero werkbook.Style
	if a == nil {
		a = &zero
	}
	if b == nil {
		b = &zero
	}
	font := func(f *werkbook.Font) werkbook.Font {
		if f == nil {
			return werkbook.Font{}
		}
		return *f
	}
	fill := func(f *werkbook.Fill) werkbook.Fill {
		if f == nil {
			return werkbook.Fill{}
		}
		return *f
	}
	border := func(f *werkbook.Border) werkbook.Border {
		if f == nil {
			return werkbook.Border{}
		}
		return *f
	}
	align := func(f *werkbook.Alignment) werkbook.Alignment {
		if f == nil {
			return werkbook.Alignment{}
		}
		return *f
	}
	return font(a.Font) == font(b.Font) && fill(a.Fill) == fill(b.Fill) &&
		border(a.Border) == border(b.Border) && align(a.Alignment) == align(b.Alignment) &&
		a.NumFmt == b.NumFmt && a.NumFmtID == b.NumFmtID
}

// StyleResult is the result of SetStyle.
type StyleResult struct {
	History HistoryState `json:"history"`
	// Changed is the number of cells whose style changed.
	Changed int `json:"changed"`
	// Range is the range that was formatted.
	Range Range `json:"range"`
	// Clamped is true when a range of full rows or full columns was limited
	// to the used range of the sheet.
	Clamped bool `json:"clamped"`
}

// SetStyle applies a style patch to a range as one undo step. See the rules
// in cmd/web/README.md.
func (s *Session) SetStyle(sheet string, rg Range, patch StylePatch) (StyleResult, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return StyleResult{}, err
	}
	rg, err = normRange(rg)
	if err != nil {
		return StyleResult{}, err
	}
	if err := patch.validate(); err != nil {
		return StyleResult{}, err
	}

	// The library has no column and row styles. A range of full rows or
	// full columns is limited to the cells that the sheet uses.
	res := StyleResult{Range: rg}
	s.ensureCalculated()
	fullRows := rg.Col1 == 1 && rg.Col2 == werkbook.MaxColumns
	fullCols := rg.Row1 == 1 && rg.Row2 == werkbook.MaxRows
	if fullRows || fullCols {
		res.Clamped = true
		// The limit is the used rectangle on both axes: column I of a
		// sheet that uses A1:C4 has no used cell, so nothing changes.
		maxRow, maxCol := s.usedRange(sh)
		rg.Row2 = min(rg.Row2, maxRow)
		rg.Col2 = min(rg.Col2, maxCol)
		if rg.Row2 < rg.Row1 || rg.Col2 < rg.Col1 {
			// No used cell in the rows or columns.
			res.Range = Range{Row1: res.Range.Row1, Col1: res.Range.Col1, Row2: res.Range.Row1, Col2: res.Range.Col1}
			res.History = s.History()
			return res, nil
		}
		res.Range = rg
	}
	if rg.cells() > maxOpCells {
		return StyleResult{}, fmt.Errorf("setStyle: the range has more than %d cells", maxOpCells)
	}

	err = s.mutate(labelFormat, func(rec *recorder) error {
		rec.setSheet(sheet)
		rec.setRange(rg)
		// Cells that had one style pointer get one new pointer, so the
		// styles of the file do not multiply.
		patched := make(map[*werkbook.Style]*werkbook.Style)
		for row := rg.Row1; row <= rg.Row2; row++ {
			for col := rg.Col1; col <= rg.Col2; col++ {
				snap := sh.SnapshotCell(col, row)
				old := snap.Style()
				st, ok := patched[old]
				if !ok {
					st = patch.apply(old)
					patched[old] = st
				}
				if st == old {
					continue
				}
				if err := rec.write(sh, col, row, snap.WithStyle(st)); err != nil {
					return err
				}
				res.Changed++
			}
		}
		return nil
	})
	if err != nil {
		return StyleResult{}, err
	}
	res.History = s.History()
	return res, nil
}
