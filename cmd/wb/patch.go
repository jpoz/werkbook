package main

import (
	"encoding/json"
	"fmt"
	"strings"

	werkbook "github.com/jpoz/werkbook"
)

// patchOp represents a single operation in a patch array.
type patchOp struct {
	Cell        string          `json:"cell,omitempty"`
	Row         int             `json:"row,omitempty"`
	Sheet       string          `json:"sheet,omitempty"`
	Value       json.RawMessage `json:"value,omitempty"`
	Formula     *string         `json:"formula,omitempty"`
	Style       json.RawMessage `json:"style,omitempty"`
	ColumnWidth *float64        `json:"column_width,omitempty"`
	RowHeight   *float64        `json:"row_height,omitempty"`
	AddSheet    string          `json:"add_sheet,omitempty"`
	DeleteSheet string          `json:"delete_sheet,omitempty"`
	Clear       bool            `json:"clear,omitempty"`
}

// opResult reports the outcome of a single patch operation.
type opResult struct {
	Index  int    `json:"index"`
	Cell   string `json:"cell,omitempty"`
	Action string `json:"action"`
	Status string `json:"status"`
	Error  string `json:"error,omitempty"`
}

// plannedOp is a normalized, non-mutating view of a patch operation.
type plannedOp struct {
	Index  int    `json:"index"`
	Sheet  string `json:"sheet,omitempty"`
	Target string `json:"target,omitempty"`
	Action string `json:"action"`
}

// parsePatchOps parses a JSON array of patch operations.
func parsePatchOps(data []byte) ([]patchOp, error) {
	var ops []patchOp
	if err := strictUnmarshal(data, &ops); err != nil {
		return nil, err
	}
	return ops, nil
}

// applyPatches applies a list of patch operations to the workbook.
// It returns per-operation results and the count of successes.
func applyPatches(f *werkbook.File, ops []patchOp, defaultSheet string) ([]opResult, int) {
	var results []opResult
	applied := 0

	for i, op := range ops {
		res := applyOnePatch(f, op, defaultSheet, i)
		if res.Status == "ok" {
			applied++
		}
		results = append(results, res)
	}

	return results, applied
}

func applyOnePatch(f *werkbook.File, op patchOp, defaultSheet string, index int) opResult {
	// Handle sheet-level operations first.
	if op.AddSheet != "" {
		_, err := f.NewSheet(op.AddSheet)
		if err != nil {
			return opResult{Index: index, Action: "add_sheet", Status: "error", Error: err.Error()}
		}
		return opResult{Index: index, Action: "add_sheet", Status: "ok"}
	}
	if op.DeleteSheet != "" {
		err := f.DeleteSheet(op.DeleteSheet)
		if err != nil {
			return opResult{Index: index, Action: "delete_sheet", Status: "error", Error: err.Error()}
		}
		return opResult{Index: index, Action: "delete_sheet", Status: "ok"}
	}

	// Handle row height.
	if op.Row > 0 && op.RowHeight != nil {
		sheetName := op.Sheet
		if sheetName == "" {
			sheetName = defaultSheet
		}
		s := f.Sheet(sheetName)
		if s == nil {
			return opResult{Index: index, Action: "set_row_height", Status: "error", Error: fmt.Sprintf("sheet %q not found", sheetName)}
		}
		err := s.SetRowHeight(op.Row, *op.RowHeight)
		if err != nil {
			return opResult{Index: index, Action: "set_row_height", Status: "error", Error: err.Error()}
		}
		return opResult{Index: index, Action: "set_row_height", Status: "ok"}
	}

	if op.Cell == "" {
		return opResult{Index: index, Action: "unknown", Status: "error", Error: "missing 'cell' field"}
	}

	sheetName := op.Sheet
	if sheetName == "" {
		sheetName = defaultSheet
	}
	s := f.Sheet(sheetName)
	if s == nil {
		return opResult{Index: index, Cell: op.Cell, Action: "unknown", Status: "error", Error: fmt.Sprintf("sheet %q not found", sheetName)}
	}

	// Column width: cell should be just a column letter like "B".
	if op.ColumnWidth != nil {
		err := s.SetColumnWidth(op.Cell, *op.ColumnWidth)
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_column_width", Status: "error", Error: err.Error()}
		}
		return opResult{Index: index, Cell: op.Cell, Action: "set_column_width", Status: "ok"}
	}

	// Clear cell(s).
	if op.Clear {
		if strings.Contains(op.Cell, ":") {
			col1, row1, col2, row2, err := werkbook.RangeToCoordinates(op.Cell)
			if err != nil {
				return opResult{Index: index, Cell: op.Cell, Action: "clear", Status: "error", Error: err.Error()}
			}
			for r := row1; r <= row2; r++ {
				for c := col1; c <= col2; c++ {
					ref, _ := werkbook.CoordinatesToCellName(c, r)
					_ = s.SetValue(ref, nil)
					_ = s.SetFormula(ref, "")
				}
			}
		} else {
			if err := s.SetValue(op.Cell, nil); err != nil {
				return opResult{Index: index, Cell: op.Cell, Action: "clear", Status: "error", Error: err.Error()}
			}
			_ = s.SetFormula(op.Cell, "")
		}
		return opResult{Index: index, Cell: op.Cell, Action: "clear", Status: "ok"}
	}

	// Style on a range (contains colon).
	if op.Style != nil && len(op.Style) > 0 && string(op.Style) != "null" {
		style, err := jsonToStyle(op.Style)
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_style", Status: "error", Error: err.Error()}
		}
		if strings.Contains(op.Cell, ":") {
			err = s.SetRangeStyle(op.Cell, style)
		} else {
			err = s.SetStyle(op.Cell, style)
		}
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_style", Status: "error", Error: err.Error()}
		}
		// If only style was set (no value or formula), return now.
		if op.Value == nil && op.Formula == nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_style", Status: "ok"}
		}
	}

	// Formula.
	if op.Formula != nil {
		err := s.SetFormula(op.Cell, *op.Formula)
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_formula", Status: "error", Error: err.Error()}
		}
		return opResult{Index: index, Cell: op.Cell, Action: "set_formula", Status: "ok"}
	}

	// Value (including null to clear).
	if op.Value != nil {
		val, err := jsonValueToGo(op.Value)
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_value", Status: "error", Error: err.Error()}
		}
		err = s.SetValue(op.Cell, val)
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_value", Status: "error", Error: err.Error()}
		}
		return opResult{Index: index, Cell: op.Cell, Action: "set_value", Status: "ok"}
	}

	return opResult{Index: index, Cell: op.Cell, Action: "noop", Status: "ok"}
}

func buildPatchPlan(ops []patchOp, defaultSheet string) []plannedOp {
	plan := make([]plannedOp, 0, len(ops))
	for i, op := range ops {
		item := plannedOp{
			Index:  i,
			Sheet:  planSheetName(op, defaultSheet),
			Action: planPatchAction(op),
		}

		switch {
		case op.AddSheet != "":
			item.Target = op.AddSheet
		case op.DeleteSheet != "":
			item.Target = op.DeleteSheet
		case op.Row > 0 && op.RowHeight != nil:
			item.Target = fmt.Sprintf("row:%d", op.Row)
		default:
			item.Target = op.Cell
		}

		plan = append(plan, item)
	}
	return plan
}

func planSheetName(op patchOp, defaultSheet string) string {
	if op.AddSheet != "" || op.DeleteSheet != "" {
		return ""
	}
	if op.Sheet != "" {
		return op.Sheet
	}
	return defaultSheet
}

func planPatchAction(op patchOp) string {
	switch {
	case op.AddSheet != "":
		return "add_sheet"
	case op.DeleteSheet != "":
		return "delete_sheet"
	case op.Row > 0 && op.RowHeight != nil:
		return "set_row_height"
	case op.Cell == "":
		return "unknown"
	case op.ColumnWidth != nil:
		return "set_column_width"
	case op.Clear:
		return "clear"
	case op.Style != nil && len(op.Style) > 0 && string(op.Style) != "null" && op.Formula != nil:
		return "set_style+set_formula"
	case op.Style != nil && len(op.Style) > 0 && string(op.Style) != "null" && op.Value != nil:
		return "set_style+set_value"
	case op.Style != nil && len(op.Style) > 0 && string(op.Style) != "null":
		return "set_style"
	case op.Formula != nil:
		return "set_formula"
	case op.Value != nil:
		return "set_value"
	default:
		return "noop"
	}
}

// jsonValueToGo converts a JSON raw value to a Go value suitable for SetValue.
func jsonValueToGo(raw json.RawMessage) (any, error) {
	if len(raw) == 0 {
		return nil, fmt.Errorf("empty value")
	}
	switch raw[0] {
	case 'n': // null
		return nil, nil
	case '"':
		var s string
		if err := json.Unmarshal(raw, &s); err != nil {
			return nil, err
		}
		return s, nil
	case 't', 'f':
		var b bool
		if err := json.Unmarshal(raw, &b); err != nil {
			return nil, err
		}
		return b, nil
	default: // number
		var num float64
		if err := json.Unmarshal(raw, &num); err != nil {
			return nil, fmt.Errorf("unsupported JSON value: %s", string(raw))
		}
		return num, nil
	}
}
