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

	if len(headers) == 0 && len(rows) == 0 {
		sb.WriteString("(empty)\n")
		return sb.String()
	}

	sb.WriteString(formatTable(FormatMarkdown, headers, rows))
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
