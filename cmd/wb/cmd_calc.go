package main

import (
	"errors"
	"fmt"
	"os"
	"strings"

	werkbook "github.com/jpoz/werkbook"
)

func cmdCalc(args []string, globals globalFlags) int {
	cmd := "calc"

	if hasHelpFlag(args) {
		return writeHelpTopic([]string{cmd}, globals)
	}
	if !ensureFormat(cmd, globals, FormatText, FormatJSON, FormatMarkdown, FormatCSV) {
		return ExitUsage
	}

	var sheetFlag, rangeFlag, outputFlag string
	var noDatesFlag bool

	i := 0
	var filePath string
	for i < len(args) {
		switch args[i] {
		case "--no-dates":
			noDatesFlag = true
			i++
			continue
		case "--sheet":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--sheet requires a value"), globals)
				return ExitUsage
			}
			sheetFlag = args[i+1]
			i += 2
		case "--range":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--range requires a value"), globals)
				return ExitUsage
			}
			rangeFlag = args[i+1]
			i += 2
		case "--output":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--output requires a value"), globals)
				return ExitUsage
			}
			outputFlag = args[i+1]
			i += 2
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

	f.Recalculate()

	// Save if --output specified.
	if outputFlag != "" {
		if err := f.SaveAs(outputFlag); err != nil {
			writeError(cmd, errFileSave(outputFlag, err), globals)
			return ExitFileIO
		}
	}

	if isTableFormat(globals.format) {
		headers, tableRows, rangeStr, err := readSheetTable(s, readOpts{
			rangeFlag: rangeFlag,
			noDates:   noDatesFlag,
		})
		if err != nil {
			writeError(cmd, errValidation(err.Error()), globals)
			return ExitValidate
		}

		var output string
		switch globals.format {
		case FormatText:
			output = renderTextTableSection(renderTableTitle("Sheet", sheetName, rangeStr), headers, tableRows)
			if outputFlag != "" {
				output = strings.TrimRight(output, "\n") + "\n\nSaved: " + outputFlag + "\n"
			}
		default:
			output = formatTable(globals.format, headers, tableRows)
		}

		fmt.Print(output)
		return ExitSuccess
	}

	// Determine range.
	var col1, row1, col2, row2 int
	if rangeFlag != "" {
		col1, row1, col2, row2, err = werkbook.RangeToCoordinates(rangeFlag)
		if err != nil {
			writeError(cmd, errInvalidRange(rangeFlag, err), globals)
			return ExitValidate
		}
	} else {
		maxRow := s.MaxRow()
		maxCol := s.MaxCol()
		if maxRow == 0 || maxCol == 0 {
			data := readData{File: filePath, Sheet: sheetName, Rows: []rowData{}}
			writeSuccess(cmd, data, globals)
			return ExitSuccess
		}
		col1, row1, col2, row2 = 1, 1, maxCol, maxRow
	}

	rangeStr := rangeFlag
	if rangeStr == "" {
		start, _ := werkbook.CoordinatesToCellName(col1, row1)
		end, _ := werkbook.CoordinatesToCellName(col2, row2)
		rangeStr = start + ":" + end
	}

	// Build row data (same as read, always include formulas for calc output).
	var rows []rowData
	for r := row1; r <= row2; r++ {
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
			// Detect date cells.
			if !noDatesFlag && v.Type == werkbook.TypeNumber {
				if isDateCell(s, ref, v) {
					cd.Type = "date"
					cd.Formatted = werkbook.ExcelSerialToTime(v.Number).Format("2006-01-02")
				}
			}
			formula, _ := s.GetFormula(ref)
			if formula != "" {
				cd.HasFormula = true
				cd.Formula = formula
			}
			cells[ref] = cd
		}
		if len(cells) > 0 {
			rows = append(rows, rowData{Row: r, Cells: cells})
		}
	}

	if rows == nil {
		rows = []rowData{}
	}

	data := readData{
		File:  filePath,
		Sheet: sheetName,
		Range: rangeStr,
		Rows:  rows,
	}

	writeSuccess(cmd, data, globals)
	return ExitSuccess
}
