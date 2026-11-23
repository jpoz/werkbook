package formula

import "strings"

// This file holds the helpers for a formula that a user is typing: what the
// caret is in (EditContext) and the F4 cycle of absolute marks
// (CycleAbsolute). Both work on the formula text without the leading '=' and
// use byte offsets. Both are tolerant: an incomplete formula is the usual
// input.

// EditPrefix is a name that ends at the caret and can be the start of a
// function name.
type EditPrefix struct {
	Start, End int
	Text       string
}

// EditCall is a function call that is open at the caret.
type EditCall struct {
	Name      string // upper case
	ArgIndex  int    // 0-based index of the argument at the caret
	NameStart int    // byte offset of the function name
}

// EditInfo is the result of EditContext.
type EditInfo struct {
	Prefix *EditPrefix
	Call   *EditCall
	// Refs are the complete references of the full text (ScanRefs).
	Refs []RefSite
}

// EditContext describes the position of the caret in a formula text.
//
// Prefix is set when the name that ends at the caret can be the start of a
// function name: it starts with a letter and has only letters, digits, '.'
// and '_'. It is nil in a string, in the middle of a name, after a number
// and after a sheet qualifier.
//
// Call is the innermost function call that is open at the caret. Commas in
// strings, array literals, nested calls and plain parentheses do not count
// for its argument index. A caret in a text argument is in the call of that
// argument, so that an argument hint stays while the user types the text.
func EditContext(src string, caret int) EditInfo {
	caret = max(0, min(caret, len(src)))
	info := EditInfo{Refs: ScanRefs(src)}

	type frame struct {
		fn        bool // false for plain parentheses
		name      string
		nameStart int
		args      int
	}
	var stack []frame
	braces := 0
	wordStart := -1 // start of the word that ends at the caret
	i := 0
scan:
	for i < caret {
		c := src[i]
		switch {
		case c == '"':
			end, closed := stringLiteralEnd(src, i)
			if !closed || caret < end {
				// The caret is in a string: no name is typed here, but the
				// call that the string is an argument of is still open.
				break scan
			}
			i = end
		case c == '\'':
			// A quoted sheet name. The caret in it is not at a function name.
			end := skipQuotedName(src, i)
			if caret < end || (end == len(src) && (end-1 == i || src[end-1] != '\'')) {
				break scan
			}
			i = end
		case c == '[':
			end := skipBracketGroup(src, i)
			if caret < end || (end == len(src) && src[end-1] != ']') {
				break scan
			}
			i = end
		case c == '#':
			// An error literal (#N/A): its letters are not a name.
			end := skipErrorLiteral(src, i)
			if caret <= end {
				break scan
			}
			i = end
		case c == '(':
			stack = append(stack, frame{})
			i++
		case c == ')':
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
			i++
		case c == '{':
			braces++
			i++
		case c == '}':
			if braces > 0 {
				braces--
			}
			i++
		case c == ',':
			if braces == 0 && len(stack) > 0 && stack[len(stack)-1].fn {
				stack[len(stack)-1].args++
			}
			i++
		case isRefWordByte(c):
			j := refWordEnd(src, i)
			if j >= caret {
				wordStart = i
				break scan
			}
			if src[j] == '(' {
				stack = append(stack, frame{fn: true, name: strings.ToUpper(src[i:j]), nameStart: i})
				i = j + 1
			} else {
				i = j
			}
		default:
			i++
		}
	}

	for k := len(stack) - 1; k >= 0; k-- {
		if stack[k].fn {
			info.Call = &EditCall{Name: stack[k].name, ArgIndex: stack[k].args, NameStart: stack[k].nameStart}
			break
		}
	}
	if wordStart >= 0 && isFunctionPrefix(src, wordStart, caret) {
		info.Prefix = &EditPrefix{Start: wordStart, End: caret, Text: src[wordStart:caret]}
	}
	return info
}

// isFunctionPrefix reports whether src[start:end] can be the start of a
// function name that the user is typing at end.
func isFunctionPrefix(src string, start, end int) bool {
	if end <= start || !isAlpha(src[start]) {
		return false
	}
	for i := start; i < end; i++ {
		c := src[i]
		if !isAlpha(c) && !isDigit(c) && c != '.' && c != '_' {
			return false
		}
	}
	// In the middle of a name, or at a name that is complete.
	if end < len(src) {
		if c := src[end]; isRefWordByte(c) || c == '(' || c == '!' {
			return false
		}
	}
	// After a sheet qualifier the text is a reference.
	if start > 0 && src[start-1] == '!' {
		return false
	}
	return true
}

// stringLiteralEnd returns the index after the string literal that starts at
// i, and whether the literal has its closing quote.
func stringLiteralEnd(src string, i int) (end int, closed bool) {
	i++
	for i < len(src) {
		if src[i] == '"' {
			if i+1 < len(src) && src[i+1] == '"' {
				i += 2
				continue
			}
			return i + 1, true
		}
		i++
	}
	return len(src), false
}

// CycleAbsolute changes the absolute marks of the references that the
// selection [selStart, selEnd] touches, as F4 does in a spreadsheet:
// A1 -> $A$1 -> A$1 -> $A1 -> A1. A reference is touched when the selection
// overlaps it or the caret is at one of its ends.
//
// A range changes both ends together, from the state of its first end. A
// column-only reference (A:A) and a row-only reference (1:3) have two states.
//
// With no selection, the new caret is at the end of the reference. With a
// selection, the new selection covers the changed references. When no
// reference is touched, the input is returned with changed false.
func CycleAbsolute(src string, selStart, selEnd int) (out string, newStart, newEnd int, changed bool) {
	if selStart > selEnd {
		selStart, selEnd = selEnd, selStart
	}
	selStart = max(0, min(selStart, len(src)))
	selEnd = max(0, min(selEnd, len(src)))

	var b strings.Builder
	last := 0
	first := true
	for _, site := range ScanRefs(src) {
		if site.Start > selEnd || site.End < selStart {
			continue
		}
		cycleRefSite(&site)
		b.WriteString(src[last:site.Start])
		if first {
			newStart = b.Len()
			first = false
		}
		b.WriteString(FormatRef(site))
		newEnd = b.Len()
		last = site.End
	}
	if first {
		return src, selStart, selEnd, false
	}
	b.WriteString(src[last:])
	if selStart == selEnd {
		newStart = newEnd
	}
	return b.String(), newStart, newEnd, true
}

// cycleRefSite moves a site to the next state of the F4 cycle.
func cycleRefSite(site *RefSite) {
	var absCol, absRow bool
	switch site.Kind {
	case RefCols:
		absCol = !site.From.AbsCol
	case RefRows:
		absRow = !site.From.AbsRow
	default:
		switch {
		case !site.From.AbsCol && !site.From.AbsRow: // A1 -> $A$1
			absCol, absRow = true, true
		case site.From.AbsCol && site.From.AbsRow: // $A$1 -> A$1
			absCol, absRow = false, true
		case !site.From.AbsCol && site.From.AbsRow: // A$1 -> $A1
			absCol, absRow = true, false
		default: // $A1 -> A1
			absCol, absRow = false, false
		}
	}
	site.From.AbsCol, site.From.AbsRow = absCol, absRow
	site.To.AbsCol, site.To.AbsRow = absCol, absRow
}
