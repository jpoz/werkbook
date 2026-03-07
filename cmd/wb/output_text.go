package main

import (
	"fmt"
	"os"
	"strings"
	"text/tabwriter"
)

func ensureFormat(command string, globals globalFlags, supported ...string) bool {
	for _, format := range supported {
		if globals.format == format {
			return true
		}
	}

	writeError(command, &ErrorInfo{
		Code:    ErrCodeInvalidFormat,
		Message: fmt.Sprintf("%s does not support --format %q", command, globals.format),
		Hint:    "Supported formats: " + strings.Join(supported, ", ") + ".",
	}, globals)
	return false
}

func writeTextResponse(resp *Response, toStderr bool) {
	text := renderTextResponse(resp)
	if text == "" {
		text = "ok"
	}
	if !strings.HasSuffix(text, "\n") {
		text += "\n"
	}

	if toStderr {
		fmt.Fprint(os.Stderr, text)
		return
	}
	fmt.Fprint(os.Stdout, text)
}

func renderTextResponse(resp *Response) string {
	var body string
	if resp.OK {
		body = renderTextSuccess(resp.Command, resp.Data)
	} else {
		body = renderTextError(resp)
	}

	if resp.Meta == nil || len(resp.Meta.Warnings) == 0 {
		return body
	}

	var sb strings.Builder
	sb.WriteString(strings.TrimRight(body, "\n"))
	sb.WriteString("\n\nWarnings:\n")
	for _, warning := range resp.Meta.Warnings {
		sb.WriteString("- ")
		sb.WriteString(warning)
		sb.WriteString("\n")
	}
	return strings.TrimRight(sb.String(), "\n")
}

func renderTextSuccess(command string, data any) string {
	switch command {
	case "version":
		if d, ok := data.(versionData); ok {
			return d.Version
		}
	case "info":
		if d, ok := data.(infoData); ok {
			return renderInfoText(d)
		}
	case "edit":
		if d, ok := data.(editData); ok {
			return renderEditText(d)
		}
	case "create":
		if d, ok := data.(createData); ok {
			return renderCreateText(d)
		}
	case "formula":
		if d, ok := data.(formulaListData); ok {
			return renderFormulaListText(d)
		}
	case "capabilities":
		if d, ok := data.(toolSpec); ok {
			return renderCapabilitiesText(d)
		}
	}

	return fmt.Sprintf("%v", data)
}

func renderTextError(resp *Response) string {
	if resp.Error == nil {
		return "error"
	}

	var sb strings.Builder
	if resp.Error.Code != "" {
		sb.WriteString("Error [")
		sb.WriteString(resp.Error.Code)
		sb.WriteString("]: ")
	} else {
		sb.WriteString("Error: ")
	}
	sb.WriteString(resp.Error.Message)

	if resp.Error.Hint != "" {
		sb.WriteString("\nHint: ")
		sb.WriteString(resp.Error.Hint)
	}

	switch data := resp.Data.(type) {
	case editData:
		sb.WriteString("\n\n")
		sb.WriteString(renderEditText(data))
	case createData:
		sb.WriteString("\n\n")
		sb.WriteString(renderCreateText(data))
	}

	return strings.TrimRight(sb.String(), "\n")
}

func renderInfoText(data infoData) string {
	var sb strings.Builder
	sb.WriteString("File: ")
	sb.WriteString(data.File)

	if len(data.Sheets) == 0 {
		sb.WriteString("\nSheets: 0")
		return sb.String()
	}

	rows := make([][]string, 0, len(data.Sheets))
	for _, sheet := range data.Sheets {
		size := "-"
		if sheet.MaxRow > 0 && sheet.MaxCol > 0 {
			size = fmt.Sprintf("%dx%d", sheet.MaxRow, sheet.MaxCol)
		}
		dataRange := sheet.DataRange
		if dataRange == "" {
			dataRange = "-"
		}
		rows = append(rows, []string{
			sheet.Name,
			dataRange,
			size,
			fmt.Sprintf("%d", sheet.NonEmptyCells),
			yesNo(sheet.HasFormulas),
		})
	}

	sb.WriteString("\n\n")
	sb.WriteString(renderTabular(
		[]string{"Sheet", "Range", "Size", "Non-empty", "Formulas"},
		rows,
	))
	return strings.TrimRight(sb.String(), "\n")
}

func renderEditText(data editData) string {
	var sb strings.Builder
	switch {
	case data.ValidateOnly:
		sb.WriteString("Validated workbook changes")
	case data.DryRun:
		sb.WriteString("Planned workbook changes")
	case data.Saved:
		sb.WriteString("Updated workbook")
	default:
		sb.WriteString("Workbook changes were not saved")
	}

	if data.Output != "" && data.Output != data.File && data.Saved {
		sb.WriteString("\nOutput: ")
		sb.WriteString(data.Output)
	} else {
		sb.WriteString("\nFile: ")
		sb.WriteString(data.File)
	}

	sb.WriteString("\nApplied: ")
	sb.WriteString(fmt.Sprintf("%d", data.Applied))
	sb.WriteString("\nFailed: ")
	sb.WriteString(fmt.Sprintf("%d", data.Failed))
	sb.WriteString("\nSaved: ")
	sb.WriteString(yesNo(data.Saved))

	if data.DryRun {
		sb.WriteString("\nDry run: yes")
	}
	if data.ValidateOnly {
		sb.WriteString("\nValidate only: yes")
	}
	if data.Atomic {
		sb.WriteString("\nAtomic: yes")
	} else {
		sb.WriteString("\nAtomic: no")
	}

	if len(data.Operations) > 0 {
		sb.WriteString("\n\nOperations\n")
		rows := make([][]string, 0, len(data.Operations))
		for _, op := range data.Operations {
			rows = append(rows, []string{
				fmt.Sprintf("%d", op.Index+1),
				op.Status,
				op.Action,
				operationTarget(op),
				op.Error,
			})
		}
		sb.WriteString(renderTabular([]string{"#", "Status", "Action", "Target", "Error"}, rows))
	}

	if len(data.Plan) > 0 {
		sb.WriteString("\nPlan\n")
		rows := make([][]string, 0, len(data.Plan))
		for _, op := range data.Plan {
			rows = append(rows, []string{
				fmt.Sprintf("%d", op.Index+1),
				emptyDash(op.Sheet),
				op.Action,
				emptyDash(op.Target),
			})
		}
		sb.WriteString(renderTabular([]string{"#", "Sheet", "Action", "Target"}, rows))
	}

	return strings.TrimRight(sb.String(), "\n")
}

func renderCreateText(data createData) string {
	var sb strings.Builder
	if data.Saved {
		sb.WriteString("Created workbook")
	} else {
		sb.WriteString("Workbook creation did not finish")
	}
	sb.WriteString("\nFile: ")
	sb.WriteString(data.File)
	sb.WriteString("\nSheets: ")
	sb.WriteString(fmt.Sprintf("%d", data.Sheets))
	sb.WriteString("\nCells applied: ")
	sb.WriteString(fmt.Sprintf("%d", data.Applied))
	sb.WriteString("\nFailed: ")
	sb.WriteString(fmt.Sprintf("%d", data.Failed))
	sb.WriteString("\nSaved: ")
	sb.WriteString(yesNo(data.Saved))

	if len(data.Operations) > 0 && data.Failed > 0 {
		sb.WriteString("\n\nOperations\n")
		rows := make([][]string, 0, len(data.Operations))
		for _, op := range data.Operations {
			rows = append(rows, []string{
				fmt.Sprintf("%d", op.Index+1),
				op.Status,
				op.Action,
				operationTarget(op),
				op.Error,
			})
		}
		sb.WriteString(renderTabular([]string{"#", "Status", "Action", "Target", "Error"}, rows))
	}

	return strings.TrimRight(sb.String(), "\n")
}

func renderFormulaListText(data formulaListData) string {
	var sb strings.Builder
	sb.WriteString(fmt.Sprintf("%d %s\n", data.Count, pluralize(data.Count, "function", "functions")))
	for _, fn := range data.Functions {
		sb.WriteString(fn)
		sb.WriteString("\n")
	}
	return strings.TrimRight(sb.String(), "\n")
}

func renderCapabilitiesText(spec toolSpec) string {
	var sb strings.Builder
	sb.WriteString(spec.Name)
	sb.WriteString(" command metadata\n")

	if len(spec.Modes) > 0 {
		sb.WriteString("\nModes\n")
		rows := make([][]string, 0, len(spec.Modes))
		for _, mode := range spec.Modes {
			rows = append(rows, []string{mode.Name, mode.Description})
		}
		sb.WriteString(renderTabular([]string{"Mode", "Description"}, rows))
	}

	if len(spec.Commands) > 0 {
		sb.WriteString("\nCommands\n")
		rows := make([][]string, 0, len(spec.Commands))
		for _, cmd := range spec.Commands {
			rows = append(rows, []string{cmd.Name, cmd.Summary})
		}
		sb.WriteString(renderTabular([]string{"Command", "Summary"}, rows))
	}

	sb.WriteString("\nUse `wb --format json capabilities` or `wb --mode agent capabilities` for structured metadata.")
	return strings.TrimRight(sb.String(), "\n")
}

func renderTabular(headers []string, rows [][]string) string {
	var sb strings.Builder
	tw := tabwriter.NewWriter(&sb, 0, 0, 2, ' ', 0)

	if len(headers) > 0 {
		fmt.Fprintln(tw, strings.Join(headers, "\t"))
	}
	for _, row := range rows {
		fmt.Fprintln(tw, strings.Join(row, "\t"))
	}
	_ = tw.Flush()

	return sb.String()
}

func operationTarget(op opResult) string {
	if op.Cell != "" {
		return op.Cell
	}
	return "-"
}

func yesNo(v bool) string {
	if v {
		return "yes"
	}
	return "no"
}

func emptyDash(s string) string {
	if s == "" {
		return "-"
	}
	return s
}

func pluralize(n int, singular, plural string) string {
	if n == 1 {
		return singular
	}
	return plural
}
