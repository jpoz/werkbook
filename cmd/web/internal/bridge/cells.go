package bridge

import (
	"fmt"
	"sort"
	"strings"

	werkbook "github.com/jpoz/werkbook"
	"github.com/jpoz/werkbook/formula"
)

// Cell is one cell of a GetRange result.
type Cell struct {
	Row     int        `json:"row"`
	Col     int        `json:"col"`
	Display string     `json:"display"`
	Input   string     `json:"input"`
	Type    string     `json:"type"`
	Formula bool       `json:"formula"`
	Spill   bool       `json:"spill,omitempty"`
	Style   *CellStyle `json:"style,omitempty"`
}

// CellStyle is the subset of the cell style that the grid shows.
type CellStyle struct {
	Bold      bool   `json:"bold,omitempty"`
	Italic    bool   `json:"italic,omitempty"`
	Underline bool   `json:"underline,omitempty"`
	Color     string `json:"color,omitempty"` // font color, "#RRGGBB"
	Fill      string `json:"fill,omitempty"`  // background color, "#RRGGBB"
	Align     string `json:"align,omitempty"` // "left", "center" or "right"
	// NumFmt is a preset name ("number", "currency", "percent", "date") or
	// "custom" for all other formats. It is "" for the general format.
	NumFmt string `json:"numFmt,omitempty"`
}

// visible reports whether the style shows on an empty cell. A number format
// shows nothing without a value.
func (st *CellStyle) visible() bool {
	if st == nil {
		return false
	}
	shown := *st
	shown.NumFmt = ""
	return shown != (CellStyle{})
}

// GetRange returns the cells of the inclusive range in row-major order. Empty
// cells are omitted, with one exception: an empty cell that has a visible
// style (for example a fill) is returned with type "empty" so that the grid
// can paint it. The corners can be given in any order.
func (s *Session) GetRange(sheet string, row1, col1, row2, col2 int) ([]Cell, error) {
	sh, err := s.sheet(sheet)
	if err != nil {
		return nil, err
	}
	if _, err := cellRef(row1, col1); err != nil {
		return nil, err
	}
	if _, err := cellRef(row2, col2); err != nil {
		return nil, err
	}
	if row1 > row2 {
		row1, row2 = row2, row1
	}
	if col1 > col2 {
		col1, col2 = col2, col1
	}
	want := rect{row1: row1, col1: col1, row2: row2, col2: col2}

	s.ensureCalculated()
	spills := s.spillRects(sh)
	fmtr := newFormatter(s.file.Date1904())

	cells := []Cell{}
	type coord struct{ row, col int }
	seen := make(map[coord]struct{})

	// Pass 1: cells that are stored in the sheet.
	for row := range sh.Rows() {
		if row.Num() < row1 {
			continue
		}
		if row.Num() > row2 {
			break
		}
		for _, c := range row.Cells() {
			if c.Col() < col1 || c.Col() > col2 {
				continue
			}
			ref, err := werkbook.CoordinatesToCellName(c.Col(), row.Num())
			if err != nil {
				continue
			}
			// GetValue resolves formulas and reads through empty placeholder
			// cells that lie in a spill range.
			v, err := sh.GetValue(ref)
			if err != nil {
				return nil, fmt.Errorf("read %s: %w", ref, err)
			}
			f := c.Formula()
			style := c.Style()
			out := Cell{
				Row:     row.Num(),
				Col:     c.Col(),
				Display: fmtr.display(v, style),
				Type:    typeName(v),
				Formula: f != "",
				Style:   cellStyle(style),
			}
			if f != "" {
				out.Input = "=" + f
			} else {
				out.Input = inputText(v)
				// A stored cell with its own value blocks a spill, so a
				// placeholder that shows a value is always a spill follower.
				out.Spill = c.Value().IsEmpty() && !v.IsEmpty() && inSpill(spills, row.Num(), c.Col())
			}
			if v.IsEmpty() && f == "" && !out.Style.visible() {
				continue
			}
			seen[coord{row.Num(), c.Col()}] = struct{}{}
			cells = append(cells, out)
		}
	}

	// Pass 2: spill followers. They exist only in the spill overlay of the
	// library, not as stored cells.
	for _, sp := range spills {
		lo := rect{
			row1: max(sp.row1, want.row1), col1: max(sp.col1, want.col1),
			row2: min(sp.row2, want.row2), col2: min(sp.col2, want.col2),
		}
		for r := lo.row1; r <= lo.row2; r++ {
			for c := lo.col1; c <= lo.col2; c++ {
				if r == sp.row1 && c == sp.col1 {
					continue // the anchor is a stored formula cell
				}
				if _, ok := seen[coord{r, c}]; ok {
					continue
				}
				ref, err := werkbook.CoordinatesToCellName(c, r)
				if err != nil {
					continue
				}
				v, err := sh.GetValue(ref)
				if err != nil || v.IsEmpty() {
					continue
				}
				seen[coord{r, c}] = struct{}{}
				cells = append(cells, Cell{
					Row:     r,
					Col:     c,
					Display: fmtr.display(v, nil),
					Input:   inputText(v),
					Type:    typeName(v),
					Spill:   true,
				})
			}
		}
	}

	sort.Slice(cells, func(i, j int) bool {
		if cells[i].Row != cells[j].Row {
			return cells[i].Row < cells[j].Row
		}
		return cells[i].Col < cells[j].Col
	})
	return cells, nil
}

// inSpill reports whether the cell is a follower (not the anchor) of a spill.
func inSpill(spills []rect, row, col int) bool {
	for _, sp := range spills {
		if sp.contains(row, col) && (row != sp.row1 || col != sp.col1) {
			return true
		}
	}
	return false
}

func typeName(v werkbook.Value) string {
	switch v.Type {
	case werkbook.TypeNumber:
		return "number"
	case werkbook.TypeString:
		return "string"
	case werkbook.TypeBool:
		return "bool"
	case werkbook.TypeError:
		return "error"
	default:
		return "empty"
	}
}

// cellStyle maps a library style to the contract subset. It returns nil when
// no field of the subset is set, so the JSON omits the style.
func cellStyle(st *werkbook.Style) *CellStyle {
	if st == nil {
		return nil
	}
	var out CellStyle
	if st.Font != nil {
		out.Bold = st.Font.Bold
		out.Italic = st.Font.Italic
		out.Underline = st.Font.Underline
		out.Color = cssColor(st.Font.Color)
	}
	if st.Fill != nil {
		out.Fill = cssColor(st.Fill.Color)
	}
	if st.Alignment != nil {
		switch st.Alignment.Horizontal {
		case werkbook.HAlignLeft:
			out.Align = "left"
		case werkbook.HAlignCenter:
			out.Align = "center"
		case werkbook.HAlignRight:
			out.Align = "right"
		}
	}
	out.NumFmt = numFmtPreset(st)
	if out == (CellStyle{}) {
		return nil
	}
	return &out
}

// cssColor converts a 6-digit RGB or 8-digit ARGB hex string to "#RRGGBB".
// It returns "" for all other input (for example theme colors).
func cssColor(hex string) string {
	if len(hex) == 8 {
		hex = hex[2:]
	}
	if len(hex) != 6 {
		return ""
	}
	for i := 0; i < len(hex); i++ {
		c := hex[i]
		isHex := (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F')
		if !isHex {
			return ""
		}
	}
	return "#" + strings.ToUpper(hex)
}

// builtinNumFmts maps the built-in number format IDs of OOXML (ECMA-376,
// 18.8.30) to format codes. ID 14 uses the en-US form. IDs that are not in
// the table use general formatting.
var builtinNumFmts = map[int]string{
	1:  "0",
	2:  "0.00",
	3:  "#,##0",
	4:  "#,##0.00",
	9:  "0%",
	10: "0.00%",
	11: "0.00E+00",
	12: "# ?/?",
	13: "# ??/??",
	14: "m/d/yyyy",
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm AM/PM",
	19: "h:mm:ss AM/PM",
	20: "h:mm",
	21: "h:mm:ss",
	22: "m/d/yyyy h:mm",
	37: "#,##0;(#,##0)",
	38: "#,##0;[Red](#,##0)",
	39: "#,##0.00;(#,##0.00)",
	40: "#,##0.00;[Red](#,##0.00)",
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mm:ss.0",
	48: "##0.0E+0",
}

// formatter makes the display text of cell values. It uses the TEXT function
// of the formula engine, which is the public path to the number format
// engine of the library.
type formatter struct {
	textFn int
	ctx    *formula.EvalContext
}

func newFormatter(date1904 bool) formatter {
	return formatter{
		textFn: formula.LookupFunc("TEXT"),
		ctx:    &formula.EvalContext{Date1904: date1904},
	}
}

// display returns the grid text of a value. A number uses the number format
// of the cell when the format engine accepts it, and general formatting
// otherwise.
func (f formatter) display(v werkbook.Value, st *werkbook.Style) string {
	switch v.Type {
	case werkbook.TypeNumber:
		// A negative number has no date or time. A spreadsheet fills the
		// cell with number signs; a date before 1900 would be wrong.
		if v.Number < 0 && st != nil && werkbook.IsDateFormat(st.NumFmt, st.NumFmtID) {
			return "########"
		}
		if code := numFmtCode(st); code != "" {
			if text, ok := f.format(v.Number, code); ok {
				return text
			}
		}
		return formula.ValueToString(formula.NumberVal(v.Number))
	case werkbook.TypeString, werkbook.TypeError:
		return v.String
	case werkbook.TypeBool:
		return boolText(v.Bool)
	default:
		return ""
	}
}

func (f formatter) format(n float64, code string) (text string, ok bool) {
	if f.textFn < 0 {
		return "", false
	}
	// A format code from a file is untrusted input for the format engine.
	// A failure there must not break the read of the whole range.
	defer func() {
		if recover() != nil {
			text, ok = "", false
		}
	}()
	res, err := formula.CallFunc(f.textFn, []formula.Value{
		formula.NumberVal(n),
		formula.StringVal(code),
	}, f.ctx)
	if err != nil || res.Type != formula.ValueString {
		return "", false
	}
	return res.Str, true
}

// numFmtCode returns the number format code of a style, or "" for general.
func numFmtCode(st *werkbook.Style) string {
	if st == nil {
		return ""
	}
	if st.NumFmt != "" {
		if strings.EqualFold(st.NumFmt, "General") || st.NumFmt == "@" {
			return ""
		}
		return st.NumFmt
	}
	return builtinNumFmts[st.NumFmtID]
}
