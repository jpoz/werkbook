package main

import (
	"bytes"
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
	File         string     `json:"file"`
	Sheets       int        `json:"sheets"`
	Cells        int        `json:"cells"`
	Applied      int        `json:"applied"`
	Failed       int        `json:"failed"`
	Saved        bool       `json:"saved"`
	DryRun       bool       `json:"dry_run,omitempty"`
	ValidateOnly bool       `json:"validate_only,omitempty"`
	Operations   []opResult `json:"operations,omitempty"`
}

func cmdCreate(args []string, globals globalFlags) int {
	cmd := "create"

	if hasHelpFlag(args) {
		return writeHelpTopic([]string{cmd}, globals)
	}
	if !ensureFormat(cmd, globals, FormatText, FormatJSON) {
		return ExitUsage
	}

	var specFlag string
	var dryRun, validateOnly bool

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
		case "--dry-run":
			dryRun = true
			i++
		case "--validate-only":
			validateOnly = true
			dryRun = true
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

	if filePath == "" && !dryRun {
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

	// Apply cell operations, then row operations. The order is fixed: cell
	// ops run first, row ops second. Surface any same-cell collisions so the
	// caller sees that a later op clobbered an earlier one (the bug pattern
	// where a `rows` null placeholder erases a `cells` formula).
	allOps := append(spec.Cells, rowOps...)
	if dupWarnings := detectDuplicateWrites(allOps, defaultSheet); len(dupWarnings) > 0 {
		globals.warnings = append(globals.warnings, dupWarnings...)
	}
	var results []opResult
	cellsApplied := 0
	if len(allOps) > 0 {
		results, cellsApplied = applyPatches(f, allOps, defaultSheet)
	}
	failed := len(allOps) - cellsApplied

	data := createData{
		File:         filePath,
		Sheets:       len(f.SheetNames()),
		Cells:        cellsApplied,
		Applied:      cellsApplied,
		Failed:       failed,
		Saved:        false,
		DryRun:       dryRun,
		ValidateOnly: validateOnly,
		Operations:   results,
	}

	if failed > 0 {
		hint := "Workbook was not saved. Check the 'operations' array for per-operation errors."
		if dryRun {
			hint = "Dry run: workbook was not saved. Check the 'operations' array for per-operation errors."
		}
		resp := &Response{
			OK:      false,
			Command: cmd,
			Data:    data,
			Error: &ErrorInfo{
				Code:    ErrCodePartialFailure,
				Message: fmt.Sprintf("%d of %d operations failed", failed, len(allOps)),
				Hint:    hint,
			},
			Meta: buildMeta(cmd, globals),
		}
		writeResponse(resp, globals, true)
		return ExitPartial
	}

	if !dryRun {
		if err := f.SaveAs(filePath); err != nil {
			writeError(cmd, errFileSave(filePath, err), globals)
			return ExitFileIO
		}
		data.Saved = true
	}

	writeSuccess(cmd, data, globals)
	return ExitSuccess
}

// rowCellOp is the per-cell shape allowed inside a `rows.data` block when an
// element is a JSON object instead of a scalar. It's a strict subset of patch_op:
// `cell` and `sheet` come from the row block itself, and the row geometry is
// what makes the block useful, so accepting only the per-cell knobs (type,
// value, formula, style) keeps row mode coherent without inviting confusion.
type rowCellOp struct {
	Type    string          `json:"type,omitempty"`
	Value   json.RawMessage `json:"value,omitempty"`
	Formula *string         `json:"formula,omitempty"`
	Style   json.RawMessage `json:"style,omitempty"`
}

// rowsToPatchOps converts row-oriented data blocks into patch operations.
// Each element of a row may be a scalar (string/number/bool/null) or a JSON
// object carrying the same per-cell fields as patch_op.
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
				op, err := rowDataElementToPatchOp(rawVal, cellRef, sheet)
				if err != nil {
					return nil, fmt.Errorf("invalid row data at %s: %v", cellRef, err)
				}
				ops = append(ops, op)
			}
		}
	}
	return ops, nil
}

func rowDataElementToPatchOp(raw json.RawMessage, cellRef, sheet string) (patchOp, error) {
	op := patchOp{Cell: cellRef, Sheet: sheet}
	trimmed := bytes.TrimSpace(raw)
	if len(trimmed) > 0 && trimmed[0] == '{' {
		var rco rowCellOp
		if err := strictUnmarshal(trimmed, &rco); err != nil {
			return patchOp{}, err
		}
		op.Type = rco.Type
		op.Value = rco.Value
		op.Formula = rco.Formula
		op.Style = rco.Style
		return op, nil
	}
	op.Value = raw
	return op, nil
}
