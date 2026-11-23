package bridge

import (
	"strings"
	"unicode/utf8"

	werkbook "github.com/jpoz/werkbook"
	"github.com/jpoz/werkbook/formula"
)

// The functions of this file help the formula editor of the UI. They do not
// read or change the workbook. All text offsets of the contract are UTF-16
// code units (the unit of selectionStart in a browser) and count the leading
// '=' of the editor text; the formula package uses byte offsets in the text
// without the '='.

// FunctionArg is one argument of a function signature.
type FunctionArg struct {
	Name     string `json:"name"`
	Optional bool   `json:"optional"`
	Repeat   bool   `json:"repeat"`
}

// FunctionInfo describes one function for the autocomplete list and the
// argument hint.
type FunctionInfo struct {
	Name      string        `json:"name"`
	Signature string        `json:"signature"`
	Summary   string        `json:"summary"`
	Args      []FunctionArg `json:"args"`
	MinArgs   int           `json:"minArgs"`
	Variadic  bool          `json:"variadic"`
}

// TextSpan is a part of the editor text.
type TextSpan struct {
	Start int    `json:"start"`
	End   int    `json:"end"`
	Text  string `json:"text"`
}

// CallInfo is the function call that is open at the caret.
type CallInfo struct {
	Name      string `json:"name"`
	ArgIndex  int    `json:"argIndex"`
	NameStart int    `json:"nameStart"`
}

// RefSpan is one reference of the editor text.
type RefSpan struct {
	Start int    `json:"start"`
	End   int    `json:"end"`
	Text  string `json:"text"`
	Sheet string `json:"sheet"`
	// SheetEnd is the second sheet of a 3D reference (Sheet1:Sheet3!A1).
	SheetEnd string `json:"sheetEnd,omitempty"`
	Kind     string `json:"kind"`
	Range    Range  `json:"range"`
	Color    int    `json:"color"`
}

// FormulaContext is the result of Session.FormulaContext.
type FormulaContext struct {
	Prefix *TextSpan `json:"prefix"`
	Call   *CallInfo `json:"call"`
	Refs   []RefSpan `json:"refs"`
}

// CycleResult is the result of CycleRef.
type CycleResult struct {
	Text     string `json:"text"`
	SelStart int    `json:"selStart"`
	SelEnd   int    `json:"selEnd"`
	Changed  bool   `json:"changed"`
}

// Functions returns all functions of the engine, sorted by name.
func (s *Session) Functions() []FunctionInfo {
	docs := formula.FuncDocs()
	out := make([]FunctionInfo, len(docs))
	for i, d := range docs {
		args := make([]FunctionArg, len(d.Args))
		for j, a := range d.Args {
			args[j] = FunctionArg{Name: a.Name, Optional: a.Optional, Repeat: a.Repeat}
		}
		out[i] = FunctionInfo{
			Name:      d.Name,
			Signature: d.Signature,
			Summary:   d.Summary,
			Args:      args,
			MinArgs:   d.MinArgs,
			Variadic:  d.Variadic,
		}
	}
	return out
}

// FormulaContext describes the editor text at the caret: the function-name
// prefix, the open function call, and all references. text is the full
// editor text; caret is a UTF-16 offset. Text that does not start with '='
// gives an empty context.
func (s *Session) FormulaContext(text string, caret int) FormulaContext {
	ctx := FormulaContext{Refs: []RefSpan{}}
	if !strings.HasPrefix(text, "=") {
		return ctx
	}
	src := text[1:]
	// A caret before the '=' is at no name and in no call.
	info := formula.EditContext(src, max(0, utf16ToByte(text, caret)-1))

	// out converts a byte offset in src to a UTF-16 offset in text.
	out := func(off int) int { return byteToUTF16(text, off+1) }

	if p := info.Prefix; p != nil {
		ctx.Prefix = &TextSpan{Start: out(p.Start), End: out(p.End), Text: p.Text}
	}
	if c := info.Call; c != nil {
		ctx.Call = &CallInfo{Name: c.Name, ArgIndex: c.ArgIndex, NameStart: out(c.NameStart)}
	}
	colors := map[string]int{}
	for _, site := range info.Refs {
		span := RefSpan{
			Start:    out(site.Start),
			End:      out(site.End),
			Text:     src[site.Start:site.End],
			Sheet:    site.Sheet,
			SheetEnd: site.SheetEnd,
			Range:    siteRange(site),
		}
		switch site.Kind {
		case formula.RefCell:
			span.Kind = "cell"
		case formula.RefRange:
			span.Kind = "range"
		case formula.RefCols:
			span.Kind = "cols"
		case formula.RefRows:
			span.Kind = "rows"
		}
		// Equal sheet and equal range give the same color, in the sequence
		// of first occurrence.
		key := strings.ToLower(site.Sheet) + "\x00" + strings.ToLower(site.SheetEnd) + "\x00" +
			formula.FormatRef(formula.RefSite{Kind: formula.RefRange,
				From: formula.RefPoint{Col: span.Range.Col1, Row: span.Range.Row1},
				To:   formula.RefPoint{Col: span.Range.Col2, Row: span.Range.Row2}})
		color, ok := colors[key]
		if !ok {
			color = len(colors)
			colors[key] = color
		}
		span.Color = color
		ctx.Refs = append(ctx.Refs, span)
	}
	return ctx
}

// siteRange returns the normalized cell range of a reference. A column-only
// reference covers all rows and a row-only reference covers all columns.
func siteRange(site formula.RefSite) Range {
	rg := Range{Row1: site.From.Row, Col1: site.From.Col, Row2: site.To.Row, Col2: site.To.Col}
	switch site.Kind {
	case formula.RefCols:
		rg.Row1, rg.Row2 = 1, werkbook.MaxRows
	case formula.RefRows:
		rg.Col1, rg.Col2 = 1, werkbook.MaxColumns
	}
	if rg.Row1 > rg.Row2 {
		rg.Row1, rg.Row2 = rg.Row2, rg.Row1
	}
	if rg.Col1 > rg.Col2 {
		rg.Col1, rg.Col2 = rg.Col2, rg.Col1
	}
	return rg
}

// CycleRef changes the absolute marks of the references that the selection
// touches (F4). Offsets are UTF-16 code units in the full editor text. Text
// that does not start with '=' and a selection with no reference are
// returned unchanged.
func (s *Session) CycleRef(text string, selStart, selEnd int) CycleResult {
	if selStart > selEnd {
		selStart, selEnd = selEnd, selStart
	}
	res := CycleResult{Text: text, SelStart: selStart, SelEnd: selEnd}
	if !strings.HasPrefix(text, "=") {
		return res
	}
	src := text[1:]
	bStart := max(0, utf16ToByte(text, selStart)-1)
	bEnd := max(0, utf16ToByte(text, selEnd)-1)
	out, newStart, newEnd, changed := formula.CycleAbsolute(src, bStart, bEnd)
	if !changed {
		return res
	}
	newText := "=" + out
	return CycleResult{
		Text:     newText,
		SelStart: byteToUTF16(newText, newStart+1),
		SelEnd:   byteToUTF16(newText, newEnd+1),
		Changed:  true,
	}
}

// utf16ToByte converts a UTF-16 code unit offset in s to a byte offset. An
// offset out of the text is moved to the nearest end; an offset in the
// middle of a surrogate pair is moved to the start of its character.
func utf16ToByte(s string, units int) int {
	if units <= 0 {
		return 0
	}
	n := 0
	for i, r := range s {
		width := 1
		if r >= 0x10000 {
			width = 2
		}
		if n+width > units {
			return i
		}
		n += width
	}
	return len(s)
}

// byteToUTF16 converts a byte offset in s to a UTF-16 code unit offset. An
// offset in the middle of a character counts that character.
func byteToUTF16(s string, off int) int {
	off = max(0, min(off, len(s)))
	n := 0
	for i := 0; i < off; {
		r, size := utf8.DecodeRuneInString(s[i:])
		if r >= 0x10000 {
			n += 2
		} else {
			n++
		}
		i += size
	}
	return n
}
