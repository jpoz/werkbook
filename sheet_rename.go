package werkbook

import (
	"strings"

	"github.com/jpoz/werkbook/formula"
)

// escapeSheetName doubles any apostrophes in name for use inside a
// single-quoted sheet reference (e.g. Fund's Data → Fund''s Data).
func escapeSheetName(name string) string {
	return strings.ReplaceAll(name, "'", "''")
}

// unescapeSheetName reverses the doubling of apostrophes in a quoted sheet
// name extracted from formula text.
func unescapeSheetName(escaped string) string {
	return strings.ReplaceAll(escaped, "''", "'")
}

// formatSheetRef returns the properly formatted sheet!-prefix for a simple
// (non-3D) sheet reference.
func formatSheetRef(name string) string {
	if formula.NeedsQuoting(name) {
		return "'" + escapeSheetName(name) + "'!"
	}
	return name + "!"
}

// format3DSheetRef returns the properly formatted sheet!-prefix for a 3D
// reference (Start:End!). Each endpoint is independently checked for quoting.
func format3DSheetRef(start, end string) string {
	if formula.NeedsQuoting(start) || formula.NeedsQuoting(end) {
		return "'" + escapeSheetName(start) + ":" + escapeSheetName(end) + "'!"
	}
	return start + ":" + end + "!"
}

// rewriteSheetRefsInFormula rewrites every occurrence of old sheet name
// references in a formula string to use the new name. It handles:
//   - Quoted refs: 'Old Name'!A1 → 'New Name'!A1
//   - Unquoted refs: OldName!A1 → NewName!A1
//   - 3D refs: 'Start:End'!A1 and Start:End!A1
//   - Doubled apostrophes in quoted names: 'Fund''s Data'!A1
//   - Case-insensitive matching (Excel treats sheet refs as case-insensitive)
//
// Double-quoted string literals ("...") are skipped to avoid corrupting
// text content inside formulas.
func rewriteSheetRefsInFormula(src, oldName, newName string) string {
	if src == "" || oldName == newName {
		return src
	}

	oldEscaped := escapeSheetName(oldName)

	var b strings.Builder
	b.Grow(len(src))
	i := 0

	for i < len(src) {
		ch := src[i]

		// Skip double-quoted string literals verbatim.
		if ch == '"' {
			j := i + 1
			for j < len(src) {
				if src[j] == '"' {
					j++
					if j < len(src) && src[j] == '"' {
						j++ // doubled quote escape
						continue
					}
					break
				}
				j++
			}
			b.WriteString(src[i:j])
			i = j
			continue
		}

		// Quoted sheet reference: 'Name'! or 'Start:End'!
		if ch == '\'' {
			end, replacement, ok := rewriteQuotedRef(src, i, oldEscaped, newName)
			if ok {
				b.WriteString(replacement)
				i = end
				continue
			}
			// Not a match — copy through closing quote.
			if end <= len(src) {
				b.WriteString(src[i:end])
			}
			i = end
			continue
		}

		// Unquoted reference: Identifier! or Identifier:Identifier! (3D)
		if isUnquotedSheetStart(ch) {
			end, replacement, ok := rewriteUnquotedRef(src, i, oldName, newName)
			if ok {
				b.WriteString(replacement)
				i = end
				continue
			}
			b.WriteString(src[i:end])
			i = end
			continue
		}

		b.WriteByte(ch)
		i++
	}

	return b.String()
}

// rewriteQuotedRef attempts to rewrite a quoted sheet reference starting at
// src[start] (which must be '\''). Returns the index past the consumed region,
// the replacement string, and whether a rewrite occurred.
func rewriteQuotedRef(src string, start int, oldEscaped, newName string) (int, string, bool) {
	j := start + 1
	var name strings.Builder
	for j < len(src) {
		if src[j] == '\'' {
			if j+1 < len(src) && src[j+1] == '\'' {
				name.WriteString("''")
				j += 2
				continue
			}
			break
		}
		name.WriteByte(src[j])
		j++
	}

	// j at closing quote (or past end).
	if j >= len(src) || src[j] != '\'' {
		return len(src), "", false
	}
	closingQuote := j

	// Must be followed by '!' to be a sheet reference.
	if closingQuote+1 >= len(src) || src[closingQuote+1] != '!' {
		return closingQuote + 1, "", false
	}

	escaped := name.String()
	past := closingQuote + 2 // index past '!

	// Sheet token preceded by ']' indicates an external workbook ref like
	// [Book.xlsx]'Sheet 1'!A1. A ']' inside the quoted token indicates a
	// fully quoted external ref like '[Book.xlsx]Sheet1:Sheet3'!A1. In both
	// cases the sheet name lives in another workbook, so do not rewrite it
	// when renaming a local sheet.
	if precededByExternalRef(src, start) || strings.Contains(escaped, "]") {
		return past, "", false
	}

	// Check for 3D ref (colon inside the quoted name).
	if idx := strings.Index(escaped, ":"); idx >= 0 {
		left := escaped[:idx]
		right := escaped[idx+1:]
		lMatch := strings.EqualFold(left, oldEscaped)
		rMatch := strings.EqualFold(right, oldEscaped)
		if !lMatch && !rMatch {
			return closingQuote + 1, "", false
		}
		lName := unescapeSheetName(left)
		if lMatch {
			lName = newName
		}
		rName := unescapeSheetName(right)
		if rMatch {
			rName = newName
		}
		return past, format3DSheetRef(lName, rName), true
	}

	// Simple quoted ref.
	if !strings.EqualFold(escaped, oldEscaped) {
		return closingQuote + 1, "", false
	}
	return past, formatSheetRef(newName), true
}

// rewriteUnquotedRef attempts to rewrite an unquoted sheet reference starting
// at src[start]. Handles both simple (Identifier!) and 3D (Id1:Id2!) forms.
// Returns the index past the consumed region, the replacement string, and
// whether a rewrite occurred.
func rewriteUnquotedRef(src string, start int, oldName, newName string) (int, string, bool) {
	j := start + 1
	for j < len(src) && isUnquotedSheetCont(src[j]) {
		j++
	}
	word := src[start:j]

	// Sheet token preceded by ']' indicates an external workbook ref like
	// [Book.xlsx]Sheet1!A1 — the name lives in another workbook, so do not
	// rewrite it when renaming a local sheet.
	if precededByExternalRef(src, start) {
		return j, "", false
	}

	// Only match unquoted old names that don't themselves need quoting.
	canMatchUnquoted := !formula.NeedsQuoting(oldName)

	// Check for 3D ref: Word:Word2!
	// To distinguish genuine 3D refs (Other:Sheet1!A1) from cell-range-colon-
	// sheet patterns (A1:Sheet1!B1), skip 3D detection when Word1 looks like a
	// cell reference (letters followed by digits, e.g. A1, BC23). This mirrors
	// the disambiguation in formula/lexer.go:looksLikeCellRef.
	if j < len(src) && src[j] == ':' && canMatchUnquoted && !formula.LooksLikeCellRef(word) {
		k := j + 1
		if k < len(src) && isUnquotedSheetStart(src[k]) {
			m := k + 1
			for m < len(src) && isUnquotedSheetCont(src[m]) {
				m++
			}
			if m < len(src) && src[m] == '!' {
				word2 := src[k:m]
				w1 := strings.EqualFold(word, oldName)
				w2 := strings.EqualFold(word2, oldName)
				if w1 || w2 {
					s, e := word, word2
					if w1 {
						s = newName
					}
					if w2 {
						e = newName
					}
					return m + 1, format3DSheetRef(s, e), true
				}
			}
		}
	}

	// Simple ref: Word!
	if j < len(src) && src[j] == '!' && canMatchUnquoted && strings.EqualFold(word, oldName) {
		return j + 1, formatSheetRef(newName), true
	}

	// Not a sheet ref — just a bareword; return up to end of word only.
	return j, "", false
}

// precededByExternalRef reports whether the sheet token at src[start] is
// part of an external workbook reference. The bracket of [Book.xlsx] /
// [1] / [\Book.xlsx]:
//
//   - immediately precedes the token: [Book.xlsx]Sheet1!A1
//   - precedes a 3D ref: [Book.xlsx]Sheet1:Sheet3!A1 (when scanning the
//     second endpoint Sheet3, the bracket sits past one ':Word' jump).
//
// Returns true when either pattern is detected so the caller skips
// rewriting — the sheet name lives in another workbook.
func precededByExternalRef(src string, start int) bool {
	i := skipBackWhitespace(src, start-1)
	if i >= 0 && src[i] == ']' {
		return true
	}
	if i < 0 || src[i] != ':' {
		return false
	}
	// Walked one ':' back; consume an unquoted word and re-check for ']'.
	i--
	for i >= 0 && isUnquotedSheetCont(src[i]) {
		i--
	}
	i = skipBackWhitespace(src, i)
	return i >= 0 && src[i] == ']'
}

func skipBackWhitespace(src string, i int) int {
	for i >= 0 && (src[i] == ' ' || src[i] == '\t') {
		i--
	}
	return i
}

func isUnquotedSheetStart(ch byte) bool {
	return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || ch == '_'
}

func isUnquotedSheetCont(ch byte) bool {
	return isUnquotedSheetStart(ch) || (ch >= '0' && ch <= '9') || ch == '.'
}

// rewriteSheetNameRefs updates all raw formula text and defined-name Value
// strings that reference oldName to use newName instead. Called by
// SetSheetName before rebuildFormulaState so that the recompiled formulas
// and dep graph reflect the renamed sheet.
func (f *File) rewriteSheetNameRefs(oldName, newName string) {
	for _, s := range f.sheets {
		for _, r := range s.rows {
			for _, c := range r.cells {
				if c.formula == "" {
					continue
				}
				c.formula = rewriteSheetRefsInFormula(c.formula, oldName, newName)
			}
		}
	}
	for i := range f.definedNames {
		f.definedNames[i].Value = rewriteSheetRefsInFormula(f.definedNames[i].Value, oldName, newName)
	}
}
