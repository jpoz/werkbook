package bridge

import "strings"

// TSV is the plain-text clipboard format of spreadsheets: a tab between
// cells, a line end between rows. A field that has a tab, a line end or a
// quote is in double quotes, with each quote doubled.

// writeTSV encodes a grid. Rows end with "\n" except the last row.
func writeTSV(rows [][]string) string {
	var b strings.Builder
	for i, row := range rows {
		if i > 0 {
			b.WriteByte('\n')
		}
		for j, field := range row {
			if j > 0 {
				b.WriteByte('\t')
			}
			if strings.ContainsAny(field, "\t\n\r\"") {
				b.WriteByte('"')
				b.WriteString(strings.ReplaceAll(field, `"`, `""`))
				b.WriteByte('"')
			} else {
				b.WriteString(field)
			}
		}
	}
	return b.String()
}

// parseTSV decodes clipboard text into rows of fields. It accepts "\n",
// "\r\n" and "\r" line ends. One line end at the end of the text does not
// make an empty last row. A field is quoted only when its first character is
// a quote; a quote in other positions is literal. A quoted field with no
// closing quote goes to the end of the text. Empty text gives one empty
// field.
func parseTSV(text string) [][]string {
	var rows [][]string
	var row []string
	var field strings.Builder
	n := len(text)
	i := 0
	endField := func() {
		row = append(row, field.String())
		field.Reset()
	}
	endRow := func() {
		endField()
		rows = append(rows, row)
		row = nil
	}
	for i < n {
		// Start of a field.
		if text[i] == '"' {
			i++
			for i < n {
				if text[i] == '"' {
					if i+1 < n && text[i+1] == '"' {
						field.WriteByte('"')
						i += 2
						continue
					}
					i++
					break
				}
				field.WriteByte(text[i])
				i++
			}
		}
		// The rest of the field, to the next separator.
		for i < n && text[i] != '\t' && text[i] != '\n' && text[i] != '\r' {
			field.WriteByte(text[i])
			i++
		}
		if i >= n {
			endRow()
			return rows
		}
		switch text[i] {
		case '\t':
			endField()
			i++
			if i >= n {
				// A tab at the end of the text is followed by an empty field.
				endRow()
				return rows
			}
		case '\r':
			if i+1 < n && text[i+1] == '\n' {
				i++
			}
			fallthrough
		case '\n':
			endRow()
			i++
		}
	}
	if len(rows) == 0 {
		return [][]string{{""}}
	}
	return rows
}

// normalizeClipboardText makes two clipboard texts comparable: browsers and
// operating systems change line ends, and some add one at the end.
func normalizeClipboardText(text string) string {
	text = strings.ReplaceAll(text, "\r\n", "\n")
	text = strings.ReplaceAll(text, "\r", "\n")
	return strings.TrimSuffix(text, "\n")
}
