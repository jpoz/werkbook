package main

import (
	"fmt"
	"strings"
)

const (
	FormatText     = "text"
	FormatJSON     = "json"
	FormatMarkdown = "markdown"
	FormatCSV      = "csv"
)

// formatTable renders rows as markdown table or CSV.
// headers may be nil (no header row). Each row is a slice of string values.
func formatTable(format string, headers []string, rows [][]string) string {
	switch format {
	case FormatMarkdown:
		return formatMarkdown(headers, rows)
	case FormatCSV:
		return formatCSV(headers, rows)
	default:
		return ""
	}
}

func isTableFormat(format string) bool {
	return format == FormatText || format == FormatMarkdown || format == FormatCSV
}

func displayTableFormat(format string) string {
	if format == FormatText {
		return FormatMarkdown
	}
	return format
}

func renderTextTableSection(title string, headers []string, rows [][]string) string {
	if title == "" && len(headers) == 0 && len(rows) == 0 {
		return ""
	}

	var sb strings.Builder
	if title != "" {
		sb.WriteString(title)
		sb.WriteString("\n\n")
	}

	headers, rows = compactTableColumns(headers, rows)

	if len(headers) == 0 && len(rows) == 0 {
		sb.WriteString("(empty)\n")
		return sb.String()
	}

	sb.WriteString(formatAlignedMarkdownTable(headers, rows))
	if len(rows) == 0 {
		sb.WriteString("(no rows)\n")
	}
	return sb.String()
}

func renderTableTitle(prefix, sheetName, rangeStr string) string {
	if sheetName == "" && rangeStr == "" {
		return prefix
	}
	if rangeStr == "" {
		return fmt.Sprintf("%s: %s", prefix, sheetName)
	}
	if sheetName == "" {
		return fmt.Sprintf("%s: %s", prefix, rangeStr)
	}
	return fmt.Sprintf("%s: %s (%s)", prefix, sheetName, rangeStr)
}

func formatMarkdown(headers []string, rows [][]string) string {
	if len(headers) == 0 && len(rows) == 0 {
		return ""
	}

	var sb strings.Builder

	if len(headers) > 0 {
		sb.WriteString("| ")
		sb.WriteString(strings.Join(headers, " | "))
		sb.WriteString(" |\n")

		sb.WriteString("|")
		for range headers {
			sb.WriteString("---|")
		}
		sb.WriteString("\n")
	}

	for _, row := range rows {
		sb.WriteString("| ")
		sb.WriteString(strings.Join(row, " | "))
		sb.WriteString(" |\n")
	}

	return sb.String()
}

func formatCSV(headers []string, rows [][]string) string {
	var sb strings.Builder
	if len(headers) > 0 {
		sb.WriteString(csvLine(headers))
	}
	for _, row := range rows {
		sb.WriteString(csvLine(row))
	}
	return sb.String()
}

func csvLine(fields []string) string {
	escaped := make([]string, len(fields))
	for i, f := range fields {
		escaped[i] = csvEscape(f)
	}
	return strings.Join(escaped, ",") + "\n"
}

func csvEscape(s string) string {
	if strings.ContainsAny(s, ",\"\n\r") {
		return `"` + strings.ReplaceAll(s, `"`, `""`) + `"`
	}
	return s
}

func formatAlignedMarkdownTable(headers []string, rows [][]string) string {
	width := len(headers)
	for _, row := range rows {
		if len(row) > width {
			width = len(row)
		}
	}
	if width == 0 {
		return ""
	}

	widths := make([]int, width)
	for i := 0; i < width; i++ {
		if len(headers) > i {
			widths[i] = len(headers[i])
		}
	}
	for _, row := range rows {
		for i, value := range row {
			if len(value) > widths[i] {
				widths[i] = len(value)
			}
		}
	}

	var sb strings.Builder
	if len(headers) > 0 {
		writeMarkdownRow(&sb, headers, widths)
		writeMarkdownDivider(&sb, widths)
	}
	for _, row := range rows {
		writeMarkdownRow(&sb, row, widths)
	}
	return sb.String()
}

func writeMarkdownRow(sb *strings.Builder, fields []string, widths []int) {
	sb.WriteString("|")
	for i := range widths {
		value := ""
		if i < len(fields) {
			value = fields[i]
		}
		sb.WriteString(" ")
		sb.WriteString(padRight(value, widths[i]))
		sb.WriteString(" |")
	}
	sb.WriteString("\n")
}

func writeMarkdownDivider(sb *strings.Builder, widths []int) {
	sb.WriteString("|")
	for _, width := range widths {
		if width < 3 {
			width = 3
		}
		sb.WriteString(" ")
		sb.WriteString(strings.Repeat("-", width))
		sb.WriteString(" |")
	}
	sb.WriteString("\n")
}

func padRight(value string, width int) string {
	if len(value) >= width {
		return value
	}
	return value + strings.Repeat(" ", width-len(value))
}

func compactTableColumns(headers []string, rows [][]string) ([]string, [][]string) {
	width := len(headers)
	for _, row := range rows {
		if len(row) > width {
			width = len(row)
		}
	}
	if width == 0 {
		return headers, rows
	}

	keep := make([]bool, width)
	for i := 0; i < width; i++ {
		if strings.TrimSpace(fieldAt(headers, i)) != "" {
			keep[i] = true
			continue
		}
		for _, row := range rows {
			if strings.TrimSpace(fieldAt(row, i)) != "" {
				keep[i] = true
				break
			}
		}
	}

	var compactHeaders []string
	if len(headers) > 0 {
		compactHeaders = make([]string, 0, width)
		for i, ok := range keep {
			if ok {
				compactHeaders = append(compactHeaders, fieldAt(headers, i))
			}
		}
	}

	compactRows := make([][]string, 0, len(rows))
	for _, row := range rows {
		compact := make([]string, 0, width)
		for i, ok := range keep {
			if ok {
				compact = append(compact, fieldAt(row, i))
			}
		}
		compactRows = append(compactRows, compact)
	}

	return compactHeaders, compactRows
}

func fieldAt(fields []string, idx int) string {
	if idx >= 0 && idx < len(fields) {
		return fields[idx]
	}
	return ""
}
