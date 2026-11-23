package formula

import "strings"

// InvalidateTableColumnRefs returns the formula text with each structured
// reference to a removed column of a table replaced by #REF!, as a
// spreadsheet does when a table column is deleted. Without this, the
// reference would name a column that the table does not have, and the
// formula would give #NAME?.
//
// table is the name of the table and columns are the names of its removed
// columns; both compare without letter case. A reference is replaced as a
// whole when one of its column items is a removed column: Tbl[Qty],
// Tbl[[#This Row],[Qty]], Tbl[@Qty], Tbl[[Qty]:[Price]].
//
// inTable is true for a formula in a cell of the table: there a reference
// can have no table name ([@Qty], [Qty]).
//
// All other text is unchanged. src is the formula without the leading '='.
func InvalidateTableColumnRefs(src, table string, columns []string, inTable bool) (out string, changed bool) {
	if len(columns) == 0 || !strings.Contains(src, "[") {
		return src, false
	}
	removed := make(map[string]bool, len(columns))
	for _, c := range columns {
		removed[strings.ToLower(c)] = true
	}

	var b strings.Builder
	last := 0
	n := len(src)
	i := 0
	for i < n {
		switch src[i] {
		case '"':
			i = skipStringLiteral(src, i)
		case '\'':
			// A quoted sheet name; a bracket in it is an external workbook.
			i = skipQuotedName(src, i)
		case '[':
			end := skipBracketGroup(src, i)
			// The table name is the word immediately before the group.
			nameStart := i
			for nameStart > last && isRefWordByte(src[nameStart-1]) {
				nameStart--
			}
			name := src[nameStart:i]
			ours := strings.EqualFold(name, table) || (name == "" && inTable)
			if ours && structuredRefHasColumn(src[i:end], removed) {
				b.WriteString(src[last:nameStart])
				b.WriteString("#REF!")
				last = end
			}
			i = end
		default:
			i++
		}
	}
	if last == 0 {
		return src, false
	}
	b.WriteString(src[last:])
	return b.String(), true
}

// structuredRefHasColumn reports whether a structured reference group
// ("[Qty]", "[[#This Row],[Qty]]", "[@Qty]") names one of the columns.
func structuredRefHasColumn(group string, columns map[string]bool) bool {
	if len(group) < 2 || group[0] != '[' || group[len(group)-1] != ']' {
		return false // a group with no end
	}
	inner := group[1 : len(group)-1]
	isColumn := func(item string) bool {
		item = strings.TrimSpace(item)
		item = strings.TrimPrefix(item, "@")
		if item == "" || item[0] == '#' {
			return false // a special item: #This Row, #Headers, ...
		}
		// The apostrophe escapes the next character in a column name.
		var name strings.Builder
		for k := 0; k < len(item); k++ {
			if item[k] == '\'' && k+1 < len(item) {
				k++
			}
			name.WriteByte(item[k])
		}
		return columns[strings.ToLower(name.String())]
	}
	// A bracket after an apostrophe is a character of a column name, not the
	// start of an item.
	nested := false
	for k := 0; k < len(inner); k++ {
		if inner[k] == '\'' {
			k++
		} else if inner[k] == '[' {
			nested = true
			break
		}
	}
	if !nested {
		return isColumn(inner)
	}
	// The items of a group with more than one part are in brackets:
	// [#This Row],[Qty] or @[Qty]:[Price].
	for k := 0; k < len(inner); k++ {
		if inner[k] == '\'' {
			k++
			continue
		}
		if inner[k] != '[' {
			continue
		}
		end := skipBracketGroup(inner, k)
		if end > k+1 && inner[end-1] == ']' && isColumn(inner[k+1:end-1]) {
			return true
		}
		k = end - 1
	}
	return false
}
