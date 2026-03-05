package main

import (
	"encoding/json"
	"fmt"
	"io"
	"os"

	werkbook "github.com/jpoz/werkbook"
)

type createSpec struct {
	Sheets []string    `json:"sheets"`
	Cells  []patchOp   `json:"cells"`
	Rows   []createRow `json:"rows"`
}

type createRow struct {
	Sheet string              `json:"sheet"`
	Start string              `json:"start"`
	Data  [][]json.RawMessage `json:"data"`
}

type createData struct {
	File       string     `json:"file"`
	Sheets     int        `json:"sheets"`
	Cells      int        `json:"cells"`
	Applied    int        `json:"applied"`
	Failed     int        `json:"failed"`
	Saved      bool       `json:"saved"`
	Operations []opResult `json:"operations,omitempty"`
}

func cmdCreate(args []string, globals globalFlags) int {
	cmd := "create"

	if hasHelpFlag(args) {
		return writeHelpTopic([]string{cmd}, globals)
	}

	var specFlag string

	i := 0
	var filePath string
	for i < len(args) {
		switch args[i] {
		case "--spec":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--spec requires a value"), globals)
				return ExitUsage
			}
			specFlag = args[i+1]
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

	// Read spec from flag or stdin.
	var specBytes []byte
	if specFlag != "" {
		specBytes = []byte(specFlag)
	} else {
		var err error
		specBytes, err = io.ReadAll(os.Stdin)
		if err != nil {
			writeError(cmd, errInternal(err), globals)
			return ExitInternal
		}
	}

	var spec createSpec
	if len(specBytes) > 0 {
		if err := strictUnmarshal(specBytes, &spec); err != nil {
			writeError(cmd, errInvalidSpec(err), globals)
			return ExitValidate
		}
	}

	// Create workbook.
	var f *werkbook.File
	if len(spec.Sheets) > 0 {
		f = werkbook.New(werkbook.FirstSheet(spec.Sheets[0]))
		for _, name := range spec.Sheets[1:] {
			f.NewSheet(name)
		}
	} else {
		f = werkbook.New()
	}

	// Convert row-oriented data into patch operations.
	defaultSheet := f.SheetNames()[0]
	rowOps, err := rowsToPatchOps(spec.Rows, defaultSheet)
	if err != nil {
		writeError(cmd, errInvalidSpec(err), globals)
		return ExitValidate
	}

	// Apply cell operations, then row operations.
	allOps := append(spec.Cells, rowOps...)
	var results []opResult
	cellsApplied := 0
	if len(allOps) > 0 {
		results, cellsApplied = applyPatches(f, allOps, defaultSheet)
	}
	failed := len(allOps) - cellsApplied

	data := createData{
		File:       filePath,
		Sheets:     len(f.SheetNames()),
		Cells:      cellsApplied,
		Applied:    cellsApplied,
		Failed:     failed,
		Saved:      false,
		Operations: results,
	}

	if failed > 0 {
		resp := &Response{
			OK:      false,
			Command: cmd,
			Data:    data,
			Error: &ErrorInfo{
				Code:    ErrCodePartialFailure,
				Message: fmt.Sprintf("%d of %d operations failed", failed, len(allOps)),
				Hint:    "Workbook was not saved. Check the 'operations' array for per-operation errors.",
			},
			Meta: buildMeta(cmd, globals),
		}
		writeResponse(resp, globals, true)
		return ExitPartial
	}

	if err := f.SaveAs(filePath); err != nil {
		writeError(cmd, errFileSave(filePath, err), globals)
		return ExitFileIO
	}
	data.Saved = true

	writeSuccess(cmd, data, globals)
	return ExitSuccess
}

// rowsToPatchOps converts row-oriented data blocks into patch operations.
func rowsToPatchOps(rows []createRow, defaultSheet string) ([]patchOp, error) {
	var ops []patchOp
	for _, r := range rows {
		sheet := r.Sheet
		if sheet == "" {
			sheet = defaultSheet
		}
		startCell := r.Start
		if startCell == "" {
			startCell = "A1"
		}
		startCol, startRow, err := werkbook.CellNameToCoordinates(startCell)
		if err != nil {
			return nil, fmt.Errorf("invalid start cell %q: %v", startCell, err)
		}
		for ri, dataRow := range r.Data {
			for ci, rawVal := range dataRow {
				cellRef, err := werkbook.CoordinatesToCellName(startCol+ci, startRow+ri)
				if err != nil {
					return nil, fmt.Errorf("cell coordinates out of range at row %d col %d", ri, ci)
				}
				ops = append(ops, patchOp{
					Cell:  cellRef,
					Sheet: sheet,
					Value: rawVal,
				})
			}
		}
	}
	return ops, nil
}
