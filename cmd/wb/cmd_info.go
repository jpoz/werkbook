package main

import (
	"errors"
	"os"

	werkbook "github.com/jpoz/werkbook"
)

type infoData struct {
	File   string      `json:"file"`
	Sheets []sheetInfo `json:"sheets"`
}

type sheetInfo struct {
	Name          string `json:"name"`
	MaxRow        int    `json:"max_row"`
	MaxCol        int    `json:"max_col"`
	MaxColLetter  string `json:"max_col_letter"`
	NonEmptyCells int    `json:"non_empty_cells"`
	HasFormulas   bool   `json:"has_formulas"`
	DataRange     string `json:"data_range"`
}

func cmdInfo(args []string, globals globalFlags) int {
	cmd := "info"

	if hasHelpFlag(args) {
		return writeHelpTopic([]string{cmd}, globals)
	}

	var sheetFlag string

	// Parse flags.
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

	names := f.SheetNames()
	if sheetFlag != "" {
		found := false
		for _, n := range names {
			if n == sheetFlag {
				found = true
				names = []string{n}
				break
			}
		}
		if !found {
			writeError(cmd, errSheetNotFound(sheetFlag), globals)
			return ExitValidate
		}
	}

	data := infoData{File: filePath}
	for _, name := range names {
		s := f.Sheet(name)
		if s == nil {
			continue
		}
		si := buildSheetInfo(s)
		data.Sheets = append(data.Sheets, si)
	}

	writeSuccess(cmd, data, globals)
	return ExitSuccess
}

func buildSheetInfo(s *werkbook.Sheet) sheetInfo {
	maxRow := s.MaxRow()
	maxCol := s.MaxCol()

	var maxColLetter string
	if maxCol > 0 {
		maxColLetter = werkbook.ColumnNumberToName(maxCol)
	}

	nonEmpty := 0
	hasFormulas := false
	for row := range s.Rows() {
		for _, cell := range row.Cells() {
			ref, err := werkbook.CoordinatesToCellName(cell.Col(), row.Num())
			if err != nil {
				continue
			}
			v, _ := s.GetValue(ref)
			if !v.IsEmpty() {
				nonEmpty++
			}
			if !hasFormulas && cell.Formula() != "" {
				hasFormulas = true
			}
		}
	}

	var dataRange string
	if maxRow > 0 && maxCol > 0 {
		end, _ := werkbook.CoordinatesToCellName(maxCol, maxRow)
		dataRange = "A1:" + end
	}

	return sheetInfo{
		Name:          s.Name(),
		MaxRow:        maxRow,
		MaxCol:        maxCol,
		MaxColLetter:  maxColLetter,
		NonEmptyCells: nonEmpty,
		HasFormulas:   hasFormulas,
		DataRange:     dataRange,
	}
}
