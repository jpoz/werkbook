package main

import (
	"errors"
	"fmt"
	"os"
	"strconv"
	"strings"

	werkbook "github.com/jpoz/werkbook"
)

type readData struct {
	File    string    `json:"file"`
	Sheet   string    `json:"sheet"`
	Range   string    `json:"range"`
	Headers []string  `json:"headers,omitempty"`
	Rows    []rowData `json:"rows"`
}

type readMultiData struct {
	File   string     `json:"file"`
	Sheets []readData `json:"sheets"`
}

type rowData struct {
	Row   int                 `json:"row"`
	Cells map[string]cellData `json:"cells"`
}

type cellData struct {
	Value        any    `json:"value"`
	Type         string `json:"type"`
	Display      string `json:"display,omitempty"`
	Formatted    string `json:"formatted,omitempty"`
	Formula      string `json:"formula,omitempty"`
	HasFormula   bool   `json:"has_formula,omitempty"`
	Style        any    `json:"style,omitempty"`
	StyleSummary string `json:"style_summary,omitempty"`
}

func cmdRead(args []string, globals globalFlags) int {
	cmd := "read"

	if hasHelpFlag(args) {
		return writeHelpTopic([]string{cmd}, globals)
	}
	if !ensureFormat(cmd, globals, FormatText, FormatJSON, FormatMarkdown, FormatCSV) {
		return ExitUsage
	}

	var sheetFlag, rangeFlag string
	var includeFormulas, showFormulas, includeStyles, headersFlag, allSheets, styleSummaryFlag, noDatesFlag bool
	var limitFlag int
	var whereExprs []string

	i := 0
	var filePath string
	for i < len(args) {
		switch args[i] {
		case "--sheet":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--sheet requires a value"), globals)
				return ExitUsage
			}
			sheetFlag = args[i+1]
			i += 2
		case "--all-sheets":
			allSheets = true
			i++
		case "--range":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--range requires a value"), globals)
				return ExitUsage
			}
			rangeFlag = args[i+1]
			i += 2
		case "--limit", "--head":
			if i+1 >= len(args) {
				writeError(cmd, errUsage(args[i]+" requires a value"), globals)
				return ExitUsage
			}
			n, err := strconv.Atoi(args[i+1])
			if err != nil || n < 1 {
				writeError(cmd, errUsage(args[i]+" must be a positive integer"), globals)
				return ExitUsage
			}
			limitFlag = n
			i += 2
		case "--where":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--where requires a value"), globals)
				return ExitUsage
			}
			whereExprs = append(whereExprs, args[i+1])
			i += 2
		case "--include-formulas":
			includeFormulas = true
			i++
		case "--show-formulas":
			showFormulas = true
			i++
		case "--include-styles":
			includeStyles = true
			i++
		case "--style-summary":
			styleSummaryFlag = true
			i++
		case "--headers":
			headersFlag = true
			i++
		case "--no-dates":
			noDatesFlag = true
			i++
		default:
			if filePath == "" && len(args[i]) > 0 && args[i][0] != '-' {
				filePath = args[i]
				i++
			} else {
				writeError(cmd, errUsage("unknown flag: "+args[i]), globals)
				return ExitUsage
			}
		}
	}

	if filePath == "" {
		writeError(cmd, errUsage("file path required"), globals)
		return ExitUsage
	}

	if allSheets && sheetFlag != "" {
		writeError(cmd, errUsage("--all-sheets and --sheet are mutually exclusive"), globals)
		return ExitUsage
	}

	// Parse --where conditions.
	var filters []filterCondition
	for _, expr := range whereExprs {
		fc, err := parseWhere(expr)
		if err != nil {
			writeError(cmd, errUsage(err.Error()), globals)
			return ExitUsage
		}
		filters = append(filters, fc)
	}

	f, err := werkbook.Open(filePath)
	if err != nil {
		if os.IsNotExist(err) {
			writeError(cmd, errFileNotFound(filePath, err), globals)
		} else if errors.Is(err, werkbook.ErrEncryptedFile) {
			writeError(cmd, errEncryptedFile(filePath), globals)
		} else {
			writeError(cmd, errFileOpen(filePath, err), globals)
		}
		return ExitFileIO
	}

	opts := readOpts{
		rangeFlag:       rangeFlag,
		headersFlag:     headersFlag,
		includeFormulas: includeFormulas,
		showFormulas:    showFormulas,
		includeStyles:   includeStyles,
		styleSummary:    styleSummaryFlag,
		noDates:         noDatesFlag,
		limitFlag:       limitFlag,
		filters:         filters,
	}

	if allSheets {
		return readAllSheets(cmd, f, filePath, opts, globals)
	}

	// Resolve sheet.
	sheetName := sheetFlag
	if sheetName == "" {
		names := f.SheetNames()
		if len(names) == 0 {
			writeError(cmd, errInternal(fmt.Errorf("workbook has no sheets")), globals)
			return ExitInternal
		}
		sheetName = names[0]
	}

	s := f.Sheet(sheetName)
	if s == nil {
		writeError(cmd, errSheetNotFound(sheetName), globals)
		return ExitValidate
	}

	return readSingleSheet(cmd, s, filePath, sheetName, opts, globals)
}

type readOpts struct {
	rangeFlag       string
	headersFlag     bool
	includeFormulas bool
	showFormulas    bool
	includeStyles   bool
	styleSummary    bool
	noDates         bool
	limitFlag       int
	filters         []filterCondition
}

func readAllSheets(cmd string, f *werkbook.File, filePath string, opts readOpts, globals globalFlags) int {
	names := f.SheetNames()
	if len(names) == 0 {
		writeError(cmd, errInternal(fmt.Errorf("workbook has no sheets")), globals)
		return ExitInternal
	}

	if isTableFormat(globals.format) {
		var sb strings.Builder
		for i, name := range names {
			s := f.Sheet(name)
			if s == nil {
				continue
			}
			headers, tableRows, rangeStr, err := readSheetTable(s, opts)
			if err != nil {
				writeError(cmd, errValidation(err.Error()), globals)
				return ExitValidate
			}

			switch globals.format {
			case FormatText:
				if i > 0 {
					sb.WriteString("\n")
				}
				sb.WriteString(renderTextTableSection(renderTableTitle("Sheet", name, rangeStr), headers, tableRows))
			case FormatMarkdown:
				if i > 0 {
					sb.WriteString("\n")
				}
				sb.WriteString("## ")
				sb.WriteString(name)
				sb.WriteString("\n\n")
				sb.WriteString(formatTable(globals.format, headers, tableRows))
			case FormatCSV:
				if i > 0 {
					sb.WriteString("\n")
				}
				sb.WriteString("# ")
				sb.WriteString(name)
				sb.WriteString("\n")
				sb.WriteString(formatTable(globals.format, headers, tableRows))
			}
		}
		fmt.Print(sb.String())
		return ExitSuccess
	}

	// JSON output.
	var sheets []readData
	for _, name := range names {
		s := f.Sheet(name)
		if s == nil {
			continue
		}
		rd, exitCode, err := buildReadData(s, filePath, name, opts)
		if exitCode != ExitSuccess {
			writeError(cmd, err, globals)
			return exitCode
		}
		sheets = append(sheets, rd)
	}

	data := readMultiData{
		File:   filePath,
		Sheets: sheets,
	}
	writeSuccess(cmd, data, globals)
	return ExitSuccess
}

func readSingleSheet(cmd string, s *werkbook.Sheet, filePath, sheetName string, opts readOpts, globals globalFlags) int {
	if isTableFormat(globals.format) {
		headers, tableRows, rangeStr, err := readSheetTable(s, opts)
		if err != nil {
			writeError(cmd, errValidation(err.Error()), globals)
			return ExitValidate
		}
		output := formatTable(displayTableFormat(globals.format), headers, tableRows)
		if globals.format == FormatText {
			output = renderTextTableSection(renderTableTitle("Sheet", sheetName, rangeStr), headers, tableRows)
		}
		fmt.Print(output)
		return ExitSuccess
	}

	rd, exitCode, err := buildReadData(s, filePath, sheetName, opts)
	if exitCode != ExitSuccess {
		writeError(cmd, err, globals)
		return exitCode
	}
	writeSuccess(cmd, rd, globals)
	return ExitSuccess
}

// readSheetTable reads a sheet and returns headers and string rows for markdown/CSV output.
func readSheetTable(s *werkbook.Sheet, opts readOpts) (headers []string, tableRows [][]string, rangeStr string, retErr error) {
	col1, row1, col2, row2, err := resolveRange(s, opts.rangeFlag)
	if err != nil {
		return nil, nil, "", fmt.Errorf("invalid range %q: %v", opts.rangeFlag, err)
	}

	rangeStr = buildRangeStr(opts.rangeFlag, col1, row1, col2, row2)

	if opts.headersFlag {
		for c := col1; c <= col2; c++ {
			ref, _ := werkbook.CoordinatesToCellName(c, row1)
			headers = append(headers, displayCellValue(s, ref, opts))
		}
	}

	if opts.styleSummary {
		headers = append(headers, "Style")
	}

	startRow := row1
	if opts.headersFlag {
		startRow = row1 + 1
	}

	// Only cap row2 if there are no filters (filters need to scan all rows).
	if opts.limitFlag > 0 && len(opts.filters) == 0 {
		maxRow := startRow + opts.limitFlag - 1
		if maxRow < row2 {
			row2 = maxRow
		}
	}

	// Resolve filters. Strip the Style header column when resolving filter indices.
	filterHeaders := headers
	if opts.styleSummary && len(headers) > 0 {
		filterHeaders = headers[:len(headers)-1]
	}
	var resolved []resolvedFilter
	for _, fc := range opts.filters {
		idx, err := resolveColumnIndex(fc.Column, filterHeaders, col1)
		if err != nil {
			return nil, nil, "", fmt.Errorf("column %q not found in headers; check header names with: wb read --headers <file>", fc.Column)
		}
		resolved = append(resolved, resolvedFilter{cond: fc, colIdx: idx})
	}

	var count int
	for r := startRow; r <= row2; r++ {
		var row []string
		for c := col1; c <= col2; c++ {
			ref, _ := werkbook.CoordinatesToCellName(c, r)
			row = append(row, displayCellValue(s, ref, opts))
		}

		if len(resolved) > 0 && !matchesFilters(row, resolved) {
			continue
		}

		if opts.styleSummary {
			var summary string
			for c := col1; c <= col2; c++ {
				ref, _ := werkbook.CoordinatesToCellName(c, r)
				style, _ := s.GetStyle(ref)
				if sum := styleSummary(style); sum != "" {
					summary = sum
					break
				}
			}
			row = append(row, summary)
		}

		tableRows = append(tableRows, row)
		count++

		if opts.limitFlag > 0 && count >= opts.limitFlag {
			break
		}
	}

	return headers, tableRows, rangeStr, nil
}

// buildReadData builds JSON readData for a single sheet.
func buildReadData(s *werkbook.Sheet, filePath, sheetName string, opts readOpts) (readData, int, *ErrorInfo) {
	col1, row1, col2, row2, err := resolveRange(s, opts.rangeFlag)
	if err != nil {
		return readData{}, ExitValidate, errInvalidRange(opts.rangeFlag, err)
	}

	rangeStr := buildRangeStr(opts.rangeFlag, col1, row1, col2, row2)

	if col1 == 0 {
		// Empty sheet.
		return readData{File: filePath, Sheet: sheetName, Rows: []rowData{}}, ExitSuccess, nil
	}

	var headers []string
	if opts.headersFlag {
		for c := col1; c <= col2; c++ {
			ref, _ := werkbook.CoordinatesToCellName(c, row1)
			headers = append(headers, displayCellValue(s, ref, opts))
		}
	}

	startRow := row1
	if opts.headersFlag {
		startRow = row1 + 1
	}

	// Resolve filters.
	var resolved []resolvedFilter
	for _, fc := range opts.filters {
		idx, ferr := resolveColumnIndex(fc.Column, headers, col1)
		if ferr != nil {
			return readData{}, ExitValidate, errValidation(
				fmt.Sprintf("column %q not found in headers; check header names with: wb read --headers <file>", fc.Column),
			)
		}
		resolved = append(resolved, resolvedFilter{cond: fc, colIdx: idx})
	}

	var rows []rowData
	var count int
	for r := startRow; r <= row2; r++ {
		// If we have filters, build a string row first to check.
		if len(resolved) > 0 {
			var strRow []string
			for c := col1; c <= col2; c++ {
				ref, _ := werkbook.CoordinatesToCellName(c, r)
				strRow = append(strRow, displayCellValue(s, ref, opts))
			}
			if !matchesFilters(strRow, resolved) {
				continue
			}
		}

		cells := make(map[string]cellData)
		for c := col1; c <= col2; c++ {
			ref, _ := werkbook.CoordinatesToCellName(c, r)
			v, _ := s.GetValue(ref)
			if v.IsEmpty() {
				continue
			}

			cd := cellData{
				Value: v.Raw(),
				Type:  valueTypeName(v),
			}

			if !opts.noDates && v.Type == werkbook.TypeNumber {
				if isDateCell(s, ref, v) {
					cd.Type = "date"
					cd.Formatted = werkbook.SerialToTime(v.Number).Format("2006-01-02")
				}
			}

			formula, _ := s.GetFormula(ref)
			if formula != "" {
				cd.HasFormula = true
				if opts.includeFormulas {
					cd.Formula = formula
				}
			}

			if opts.showFormulas {
				cd.Display = displayCellValue(s, ref, opts)
			}

			if opts.includeStyles {
				style, _ := s.GetStyle(ref)
				if style != nil {
					cd.Style = styleToJSON(style)
				}
			}

			if opts.styleSummary {
				style, _ := s.GetStyle(ref)
				if sum := styleSummary(style); sum != "" {
					cd.StyleSummary = sum
				}
			}

			cells[ref] = cd
		}

		if len(cells) > 0 {
			rows = append(rows, rowData{Row: r, Cells: cells})
		}

		count++
		if opts.limitFlag > 0 && count >= opts.limitFlag {
			break
		}
	}

	if rows == nil {
		rows = []rowData{}
	}

	return readData{
		File:    filePath,
		Sheet:   sheetName,
		Range:   rangeStr,
		Headers: headers,
		Rows:    rows,
	}, ExitSuccess, nil
}

func displayCellValue(s *werkbook.Sheet, ref string, opts readOpts) string {
	if opts.showFormulas {
		formula, _ := s.GetFormula(ref)
		if formula != "" {
			return "=" + formula
		}
	}

	v, _ := s.GetValue(ref)
	if !opts.noDates && v.Type == werkbook.TypeNumber && isDateCell(s, ref, v) {
		return werkbook.SerialToTime(v.Number).Format("2006-01-02")
	}
	return valueToString(v)
}

// resolveRange returns the column/row bounds for a sheet given an optional range flag.
// Returns (0,0,0,0,nil) for an empty sheet with no explicit range.
func resolveRange(s *werkbook.Sheet, rangeFlag string) (col1, row1, col2, row2 int, err error) {
	if rangeFlag != "" {
		col1, row1, col2, row2, err = werkbook.RangeToCoordinates(rangeFlag)
		return
	}
	maxRow := s.MaxRow()
	maxCol := s.MaxCol()
	if maxRow == 0 || maxCol == 0 {
		return 0, 0, 0, 0, nil
	}
	return 1, 1, maxCol, maxRow, nil
}

func buildRangeStr(rangeFlag string, col1, row1, col2, row2 int) string {
	if rangeFlag != "" {
		return rangeFlag
	}
	if col1 == 0 {
		return ""
	}
	start, _ := werkbook.CoordinatesToCellName(col1, row1)
	end, _ := werkbook.CoordinatesToCellName(col2, row2)
	return start + ":" + end
}

func valueToString(v werkbook.Value) string {
	switch v.Type {
	case werkbook.TypeNumber:
		if v.Number == float64(int64(v.Number)) {
			return fmt.Sprintf("%d", int64(v.Number))
		}
		return fmt.Sprintf("%g", v.Number)
	case werkbook.TypeString:
		return v.String
	case werkbook.TypeBool:
		if v.Bool {
			return "TRUE"
		}
		return "FALSE"
	case werkbook.TypeError:
		return v.String
	default:
		return ""
	}
}

func valueTypeName(v werkbook.Value) string {
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

// isDateCell checks whether the cell's style indicates a date number format.
// It also applies a plausibility guard: serial numbers <= 0 or > 2958465 (year 9999)
// are rejected to avoid obviously wrong date interpretations.
func isDateCell(s *werkbook.Sheet, ref string, v werkbook.Value) bool {
	// Plausibility guard: reject implausible serial numbers.
	if v.Type == werkbook.TypeNumber && (v.Number <= 0 || v.Number > 2958465) {
		return false
	}
	style, _ := s.GetStyle(ref)
	if style == nil {
		return false
	}
	return werkbook.IsDateFormat(style.NumFmt, style.NumFmtID)
}
