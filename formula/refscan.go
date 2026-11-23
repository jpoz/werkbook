package formula

import (
	"strconv"
	"strings"
)

// This file holds the reference scanner and rewriter. It is the one shared
// mechanism for changes of reference text: the shift of a copy or fill
// (ShiftRefs) and, in later slices, structural edits and the F4 cycle.
//
// The scanner works on the formula text, not on the AST: the AST has no
// source positions and no printer, so a rewrite through it would lose spaces
// and letter case. A rewrite changes only the bytes of the references that
// change; all other text stays byte for byte the same.
//
// The scanner is tolerant. It never fails. Text that it does not understand
// is an opaque span: it is not a reference and it is never changed. Opaque
// spans are:
//
//   - string literals, also an unterminated one (to the end of the text)
//   - a [...] group: structured references (Table1[Col]) and external
//     workbook prefixes ([Book.xlsx]Sheet1!A1); the reference that follows
//     an external prefix is opaque too
//   - error literals (#REF!, #N/A), and the reference text that follows the
//     deleted-sheet form #REF!A1
//   - the spill operator: in C12# the reference C12 is a site and the # stays
//   - function names (LOG10(), table names, defined names, booleans, numbers
//
// Known limits:
//
//   - A LET or LAMBDA parameter that looks like a cell address (x1) is seen
//     as a reference. Excel does not permit such names. A shift renames all
//     its uses in the same way.
//   - A cell range with spaces around the colon (A1 : B2) is one site. When
//     it is replaced, it is written without the spaces. Column-only and
//     row-only ranges are found only without spaces (A:C, 1:3): "x : y" can
//     be two names.

const (
	refMaxCol = 16384   // XFD
	refMaxRow = 1048576 // last row
)

// RefKind is the form of a reference.
type RefKind int

const (
	RefCell  RefKind = iota // A1
	RefRange                // A1:B2
	RefCols                 // A:C
	RefRows                 // 1:3
)

// RefPoint is one corner of a reference. Col is 0 in a row-only reference
// and Row is 0 in a column-only reference.
type RefPoint struct {
	Col, Row       int
	AbsCol, AbsRow bool
}

// RefSite is one reference in a formula text.
type RefSite struct {
	// Start and End are the byte span of the reference in the source,
	// with the sheet qualifier.
	Start, End int
	// Sheet and SheetEnd are the sheet names of the qualifier, without
	// quotes. Both are "" for a reference to the sheet of the formula.
	// SheetEnd is set for a 3D reference (Sheet1:Sheet3!A1).
	Sheet, SheetEnd string
	Kind            RefKind
	// From and To are the corners in the sequence of the text; they are not
	// sorted. To is equal to From for RefCell.
	From, To RefPoint

	// qual is the qualifier text as written, with the '!'. FormatRef writes
	// it again when the sheet names did not change, so that the quote style
	// of the source stays.
	qual                    string
	qualSheet, qualSheetEnd string
}

// RefAction is the decision of a RewriteRefs callback for one site.
type RefAction int

const (
	RefKeep    RefAction = iota // leave the source text unchanged
	RefReplace                  // write the returned site
	RefInvalid                  // write #REF!
)

// ScanRefs returns the references in a formula text, in the sequence of the
// text. src is the formula without the leading '='. It never fails: text that
// is not a reference is skipped.
func ScanRefs(src string) []RefSite {
	var sites []RefSite
	n := len(src)
	i := 0
	for i < n {
		c := src[i]
		switch {
		case c == '"':
			i = skipStringLiteral(src, i)
		case c == '[':
			i = skipExternalTail(src, skipBracketGroup(src, i))
		case c == '\'':
			i = scanQuotedSheet(src, i, &sites)
		case c == '#':
			i = skipErrorLiteral(src, i)
		case isDigit(c) || (c == '.' && i+1 < n && isDigit(src[i+1])):
			i = scanNumber(src, i, &sites)
		case isRefWordByte(c):
			i = scanWord(src, i, &sites)
		default:
			i++
		}
	}
	return sites
}

// RewriteRefs calls fn for each reference in src and returns the text with
// the replacements. Text that is not a replaced reference is unchanged. When
// fn keeps all sites, the result is src.
func RewriteRefs(src string, fn func(RefSite) (RefSite, RefAction)) string {
	sites := ScanRefs(src)
	if len(sites) == 0 {
		return src
	}
	var b strings.Builder
	last := 0
	for _, site := range sites {
		repl, action := fn(site)
		var text string
		end := site.End
		switch action {
		case RefReplace:
			text = FormatRef(repl)
		case RefInvalid:
			text = "#REF!"
			// A spill operator has no meaning after an error.
			if end < len(src) && src[end] == '#' {
				end++
			}
		default:
			continue
		}
		b.WriteString(src[last:site.Start])
		b.WriteString(text)
		last = end
	}
	if last == 0 {
		return src
	}
	b.WriteString(src[last:])
	return b.String()
}

// FormatRef writes a reference. Column letters are upper case. The sheet
// qualifier is written as it was in the source when the sheet names of the
// site did not change; else it is quoted when the names need quotes.
func FormatRef(site RefSite) string {
	var b strings.Builder
	switch {
	case site.Sheet == "" && site.SheetEnd == "":
		// No qualifier.
	case site.qual != "" && site.Sheet == site.qualSheet && site.SheetEnd == site.qualSheetEnd:
		b.WriteString(site.qual)
	default:
		name := site.Sheet
		quote := NeedsQuoting(site.Sheet)
		if site.SheetEnd != "" {
			name += ":" + site.SheetEnd
			quote = quote || NeedsQuoting(site.SheetEnd)
		}
		if quote {
			b.WriteByte('\'')
			b.WriteString(strings.ReplaceAll(name, "'", "''"))
			b.WriteByte('\'')
		} else {
			b.WriteString(name)
		}
		b.WriteByte('!')
	}
	writeRefPoint(&b, site.From)
	if site.Kind != RefCell {
		b.WriteByte(':')
		writeRefPoint(&b, site.To)
	}
	return b.String()
}

func writeRefPoint(b *strings.Builder, p RefPoint) {
	if p.Col > 0 {
		if p.AbsCol {
			b.WriteByte('$')
		}
		b.WriteString(ColNumberToLetters(p.Col))
	}
	if p.Row > 0 {
		if p.AbsRow {
			b.WriteByte('$')
		}
		b.WriteString(strconv.Itoa(p.Row))
	}
}

func isDigit(c byte) bool { return c >= '0' && c <= '9' }

// isRefWordByte reports whether c can be part of a name, a function name, an
// unquoted sheet name or a reference. Bytes of non-ASCII characters count, so
// that a name like "Größe1" is one word and its tail is not read as "e1".
func isRefWordByte(c byte) bool {
	return isAlpha(c) || isDigit(c) || c == '_' || c == '.' || c == '$' || c == '\\' || c >= 0x80
}

func refWordEnd(src string, i int) int {
	for i < len(src) && isRefWordByte(src[i]) {
		i++
	}
	return i
}

// skipStringLiteral returns the index after the string literal that starts
// at i. A doubled quote is an escaped quote. An unterminated literal goes to
// the end of the text.
func skipStringLiteral(src string, i int) int {
	i++
	for i < len(src) {
		if src[i] == '"' {
			if i+1 < len(src) && src[i+1] == '"' {
				i += 2
				continue
			}
			return i + 1
		}
		i++
	}
	return len(src)
}

// skipBracketGroup returns the index after the [...] group that starts at i.
// Groups nest (Table1[[#This Row],[Col]]). In a structured reference the
// apostrophe escapes the next character.
func skipBracketGroup(src string, i int) int {
	depth := 0
	for i < len(src) {
		switch src[i] {
		case '\'':
			if depth > 0 && i+1 < len(src) {
				i++
			}
		case '[':
			depth++
		case ']':
			depth--
			if depth == 0 {
				return i + 1
			}
		}
		i++
	}
	return len(src)
}

// skipExternalTail skips the reference that follows an external workbook
// prefix: in [Book.xlsx]Sheet1!A1:B2 the text after the group refers to a
// different workbook and must not change.
func skipExternalTail(src string, i int) int {
	if i >= len(src) {
		return i
	}
	if src[i] == '\'' {
		i = skipQuotedName(src, i)
	}
	return skipOpaqueRefText(src, i)
}

// skipOpaqueRefText skips reference-like text: word bytes, '!' and ':'.
func skipOpaqueRefText(src string, i int) int {
	for i < len(src) && (isRefWordByte(src[i]) || src[i] == '!' || src[i] == ':') {
		i++
	}
	return i
}

// skipQuotedName returns the index after the single-quoted name that starts
// at i, or len(src) when it has no end.
func skipQuotedName(src string, i int) int {
	i++
	for i < len(src) {
		if src[i] == '\'' {
			if i+1 < len(src) && src[i+1] == '\'' {
				i += 2
				continue
			}
			return i + 1
		}
		i++
	}
	return len(src)
}

// skipErrorLiteral skips an error literal (#REF!, #N/A, #NAME?) or a spill
// operator. In the deleted-sheet form #REF!A1 the reference text after the
// literal is skipped too.
func skipErrorLiteral(src string, i int) int {
	i++
	start := i
	for i < len(src) && (isAlpha(src[i]) || isDigit(src[i]) || src[i] == '/' || src[i] == '_') {
		i++
	}
	if i == start {
		return i // a lone '#': the spill operator
	}
	if i < len(src) && (src[i] == '!' || src[i] == '?') {
		i++
		if src[i-1] == '!' {
			i = skipOpaqueRefText(src, i)
		}
	}
	return i
}

// scanNumber handles text that starts with a digit: a row range (1:3) or a
// number literal.
func scanNumber(src string, i int, sites *[]RefSite) int {
	if site, ok := parseRefBody(src, i); ok && site.Kind == RefRows {
		*sites = append(*sites, site)
		return site.End
	}
	n := len(src)
	for i < n && isDigit(src[i]) {
		i++
	}
	if i < n && src[i] == '.' {
		i++
		for i < n && isDigit(src[i]) {
			i++
		}
	}
	if i < n && (src[i] == 'e' || src[i] == 'E') {
		j := i + 1
		if j < n && (src[j] == '+' || src[j] == '-') {
			j++
		}
		if j < n && isDigit(src[j]) {
			for j < n && isDigit(src[j]) {
				j++
			}
			i = j
		}
	}
	// Word bytes that follow a number are not a reference ("2abc1").
	return refWordEnd(src, i)
}

// scanWord handles a run of word bytes: a reference, an unquoted sheet
// qualifier with its reference, a function name, or a name.
func scanWord(src string, i int, sites *[]RefSite) int {
	n := len(src)
	j := refWordEnd(src, i)
	word := src[i:j]
	var next byte
	if j < n {
		next = src[j]
	}

	switch next {
	case '(', '[':
		// A function name (LOG10() or a table name (Tbl1[Col]).
		return j
	case '!':
		return appendQualified(src, i, j+1, word, "", sites)
	case ':':
		// A 3D reference: Sheet1:Sheet3!A1. As in the lexer, a first word
		// that is a cell address makes the colon a range operator.
		if !LooksLikeCellRef(word) {
			k := refWordEnd(src, j+1)
			if k > j+1 && k < n && src[k] == '!' {
				return appendQualified(src, i, k+1, word, src[j+1:k], sites)
			}
		}
	}

	if site, ok := parseRefBody(src, i); ok {
		*sites = append(*sites, site)
		return site.End
	}
	return j
}

// scanQuotedSheet handles text that starts with a single quote: a quoted
// sheet qualifier with its reference.
func scanQuotedSheet(src string, i int, sites *[]RefSite) int {
	end := skipQuotedName(src, i)
	if end >= len(src) || src[end] != '!' {
		return end
	}
	name := strings.ReplaceAll(src[i+1:end-1], "''", "'")
	if name == "" || strings.ContainsAny(name, "[]") {
		// '[Book.xlsx]Sheet1'!A1 refers to a different workbook.
		return skipOpaqueRefText(src, end)
	}
	sheet, sheetEnd := name, ""
	if colon := strings.IndexByte(name, ':'); colon > 0 {
		sheet, sheetEnd = name[:colon], name[colon+1:]
	}
	return appendQualified(src, i, end+1, sheet, sheetEnd, sites)
}

// appendQualified parses the reference body at bodyStart, which follows a
// sheet qualifier that starts at qualStart. When the body is not a reference
// (Sheet1!Total, Sheet1!#REF!) the text is skipped.
func appendQualified(src string, qualStart, bodyStart int, sheet, sheetEnd string, sites *[]RefSite) int {
	site, ok := parseRefBody(src, bodyStart)
	if !ok {
		return max(refWordEnd(src, bodyStart), bodyStart)
	}
	site.Start = qualStart
	site.Sheet, site.SheetEnd = sheet, sheetEnd
	site.qual = src[qualStart:bodyStart]
	site.qualSheet, site.qualSheetEnd = sheet, sheetEnd
	*sites = append(*sites, site)
	return site.End
}

// parseRefBody parses a reference without qualifier at i: A1, A1:B2, A:C or
// 1:3, each with optional '$' marks.
func parseRefBody(src string, i int) (RefSite, bool) {
	if from, end, ok := parseCellPoint(src, i); ok && refEndsAt(src, end) {
		site := RefSite{Start: i, End: end, Kind: RefCell, From: from, To: from}
		// The range operator of a cell range can have spaces on both sides
		// (A1 : B2). The two cells are then one range, so that a rewrite
		// that makes one end invalid makes the full range #REF!.
		if colon := skipSpaces(src, end); colon < len(src) && src[colon] == ':' {
			if to, end2, ok := parseCellPoint(src, skipSpaces(src, colon+1)); ok && refEndsAt(src, end2) {
				site.Kind, site.To, site.End = RefRange, to, end2
			}
		}
		return site, true
	}
	if from, end, ok := parseColPoint(src, i); ok && end < len(src) && src[end] == ':' {
		if to, end2, ok := parseColPoint(src, end+1); ok && refEndsAt(src, end2) {
			return RefSite{Start: i, End: end2, Kind: RefCols, From: from, To: to}, true
		}
	}
	if from, end, ok := parseRowPoint(src, i); ok && end < len(src) && src[end] == ':' {
		if to, end2, ok := parseRowPoint(src, end+1); ok && refEndsAt(src, end2) {
			return RefSite{Start: i, End: end2, Kind: RefRows, From: from, To: to}, true
		}
	}
	return RefSite{}, false
}

func skipSpaces(src string, i int) int {
	for i < len(src) && src[i] == ' ' {
		i++
	}
	return i
}

// refEndsAt reports whether a reference can end at index i: the next byte
// must not continue a word, open a call or a structured reference, or make
// the text a sheet name.
func refEndsAt(src string, i int) bool {
	if i >= len(src) {
		return true
	}
	c := src[i]
	return !isRefWordByte(c) && c != '(' && c != '[' && c != '!'
}

// parseColPoint parses $?LETTERS at i.
func parseColPoint(src string, i int) (RefPoint, int, bool) {
	var p RefPoint
	if i < len(src) && src[i] == '$' {
		p.AbsCol = true
		i++
	}
	start := i
	for i < len(src) && isAlpha(src[i]) {
		i++
	}
	if i == start || i-start > 3 {
		return RefPoint{}, 0, false
	}
	p.Col = ColLettersToNumber(src[start:i])
	if p.Col < 1 || p.Col > refMaxCol {
		return RefPoint{}, 0, false
	}
	return p, i, true
}

// parseRowPoint parses $?DIGITS at i.
func parseRowPoint(src string, i int) (RefPoint, int, bool) {
	var p RefPoint
	if i < len(src) && src[i] == '$' {
		p.AbsRow = true
		i++
	}
	start := i
	row := 0
	for i < len(src) && isDigit(src[i]) {
		if row > refMaxRow {
			return RefPoint{}, 0, false
		}
		row = row*10 + int(src[i]-'0')
		i++
	}
	if i == start || row < 1 || row > refMaxRow {
		return RefPoint{}, 0, false
	}
	p.Row = row
	return p, i, true
}

// parseCellPoint parses $?LETTERS$?DIGITS at i.
func parseCellPoint(src string, i int) (RefPoint, int, bool) {
	col, end, ok := parseColPoint(src, i)
	if !ok {
		return RefPoint{}, 0, false
	}
	row, end, ok := parseRowPoint(src, end)
	if !ok {
		return RefPoint{}, 0, false
	}
	col.Row, col.AbsRow = row.Row, row.AbsRow
	return col, end, true
}
