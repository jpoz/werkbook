package main

import (
	"encoding/json"
	"fmt"
	"sort"
	"strings"
	"time"

	werkbook "github.com/jpoz/werkbook"
)

// patchOp represents a single operation in a patch array.
type patchOp struct {
	Cell        string          `json:"cell,omitempty"`
	Row         int             `json:"row,omitempty"`
	Sheet       string          `json:"sheet,omitempty"`
	Type        string          `json:"type,omitempty"`
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

	// Reject type without a value, or type combined with a formula.
	if op.Type != "" && op.Value == nil {
		return opResult{Index: index, Cell: op.Cell, Action: "set_value", Status: "error", Error: fmt.Sprintf("type %q requires a value", op.Type)}
	}

	// Value (including null to clear).
	if op.Value != nil {
		val, err := jsonValueToGoTyped(op.Value, op.Type)
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_value", Status: "error", Error: err.Error()}
		}
		err = s.SetValue(op.Cell, val)
		if err != nil {
			return opResult{Index: index, Cell: op.Cell, Action: "set_value", Status: "error", Error: err.Error()}
		}
		// Auto-apply a default number format for typed dates/times when the
		// caller did not provide an explicit style. Otherwise Excel renders
		// the underlying serial number rather than a date.
		if op.Type != "" && !hasExplicitStyle(op.Style) {
			if style := defaultStyleForType(op.Type); style != nil {
				_ = s.SetStyle(op.Cell, style)
			}
		}
		return opResult{Index: index, Cell: op.Cell, Action: "set_value", Status: "ok"}
	}

	return opResult{Index: index, Cell: op.Cell, Action: "noop", Status: "ok"}
}

func hasExplicitStyle(raw json.RawMessage) bool {
	return len(raw) > 0 && string(raw) != "null"
}

// defaultStyleForType returns a Style applying the canonical built-in number
// format for typed temporal values, or nil for types that don't need one.
func defaultStyleForType(typ string) *werkbook.Style {
	switch typ {
	case "date":
		return &werkbook.Style{NumFmtID: 14} // m/d/yyyy
	case "datetime":
		return &werkbook.Style{NumFmtID: 22} // m/d/yyyy h:mm
	case "time":
		return &werkbook.Style{NumFmtID: 21} // h:mm:ss
	default:
		return nil
	}
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

// jsonValueToGoTyped converts a JSON raw value to a Go value, honoring an
// optional type tag. Supported types: "date", "datetime", "time" — the value
// must be a JSON string that parses against one of the accepted layouts for
// that type. An empty type tag dispatches to jsonValueToGo for plain decoding.
func jsonValueToGoTyped(raw json.RawMessage, typ string) (any, error) {
	if typ == "" {
		return jsonValueToGo(raw)
	}
	if len(raw) == 0 {
		return nil, fmt.Errorf("type %q requires a value", typ)
	}
	if raw[0] == 'n' { // null with a type is an error — clear should be explicit.
		return nil, fmt.Errorf("type %q is not compatible with null value", typ)
	}
	if raw[0] != '"' {
		return nil, fmt.Errorf("type %q requires a JSON string value, got %s", typ, string(raw))
	}
	var s string
	if err := json.Unmarshal(raw, &s); err != nil {
		return nil, err
	}
	t, err := parseTypedTime(s, typ)
	if err != nil {
		return nil, fmt.Errorf("invalid %s %q: %v", typ, s, err)
	}
	return t, nil
}

// parseTypedTime parses a string into a time.Time using a small whitelist of
// layouts per type. The returned time uses time.UTC so the workbook's serial
// representation is deterministic regardless of host timezone.
func parseTypedTime(s, typ string) (time.Time, error) {
	var layouts []string
	switch typ {
	case "date":
		layouts = []string{"2006-01-02", "2006/01/02", "01/02/2006"}
	case "datetime":
		layouts = []string{
			time.RFC3339,
			"2006-01-02T15:04:05",
			"2006-01-02 15:04:05",
			"2006-01-02 15:04",
		}
	case "time":
		layouts = []string{"15:04:05", "15:04"}
	default:
		return time.Time{}, fmt.Errorf("unknown type (allowed: date, datetime, time)")
	}
	for _, layout := range layouts {
		if t, err := time.ParseInLocation(layout, s, time.UTC); err == nil {
			return t, nil
		}
	}
	return time.Time{}, fmt.Errorf("does not match any accepted layout")
}

// detectDuplicateWrites scans a patch list for value-mutating operations that
// target the same (sheet, cell) more than once. Each such collision becomes a
// warning so callers see last-write-wins clobbers (e.g. a `rows` block that
// overwrites a `cells` formula with a placeholder null) instead of silently
// losing earlier ops.
//
// Style-only ops are excluded: layering style on top of a value is intentional.
// Range ops with a colon are expanded so a clear over A1:B3 is compared
// cell-by-cell against later writes inside that rectangle.
//
// At most maxDuplicateWarnings distinct cells are reported; remaining
// collisions roll up into a single "...and N more" line.
func detectDuplicateWrites(ops []patchOp, defaultSheet string) []string {
	type key struct{ sheet, cell string }
	first := map[key]int{}    // first index that wrote each cell
	indices := map[key][]int{} // all op indices per cell
	var order []key            // insertion order of first writes

	for i, op := range ops {
		if !isCellMutation(op) {
			continue
		}
		sheet := op.Sheet
		if sheet == "" {
			sheet = defaultSheet
		}
		if op.Cell == "" {
			continue
		}
		cells, err := expandCellsForRef(op.Cell)
		if err != nil {
			continue // malformed refs are caught by applyOnePatch
		}
		for _, c := range cells {
			k := key{sheet, c}
			if _, seen := first[k]; !seen {
				first[k] = i
				order = append(order, k)
			}
			indices[k] = append(indices[k], i)
		}
	}

	type dup struct {
		k       key
		indices []int
	}
	var dups []dup
	for _, k := range order {
		if len(indices[k]) > 1 {
			dups = append(dups, dup{k, indices[k]})
		}
	}
	if len(dups) == 0 {
		return nil
	}
	sort.SliceStable(dups, func(i, j int) bool {
		return dups[i].indices[0] < dups[j].indices[0]
	})

	const maxDuplicateWarnings = 20
	var warnings []string
	for i, d := range dups {
		if i >= maxDuplicateWarnings {
			warnings = append(warnings, fmt.Sprintf(
				"...and %d more cells written multiple times", len(dups)-maxDuplicateWarnings))
			break
		}
		warnings = append(warnings, fmt.Sprintf(
			"%s!%s written by %d operations (indexes %s); last write wins",
			d.k.sheet, d.k.cell, len(d.indices), formatIndexList(d.indices)))
	}
	return warnings
}

// isCellMutation reports whether op writes a value, formula, or clears a cell.
// Pure style ops, sheet-level ops, and width/height ops are excluded — they
// don't compete with each other for cell content.
func isCellMutation(op patchOp) bool {
	if op.AddSheet != "" || op.DeleteSheet != "" {
		return false
	}
	if op.ColumnWidth != nil || (op.Row > 0 && op.RowHeight != nil) {
		return false
	}
	if op.Cell == "" {
		return false
	}
	return op.Clear || op.Formula != nil || op.Value != nil
}

// expandCellsForRef expands a cell reference into its constituent cell names.
// A1-style single cells return a one-element slice; ranges like A1:B3 expand
// into every cell they cover. Used for collision detection across range ops.
func expandCellsForRef(ref string) ([]string, error) {
	if !strings.Contains(ref, ":") {
		return []string{ref}, nil
	}
	col1, row1, col2, row2, err := werkbook.RangeToCoordinates(ref)
	if err != nil {
		return nil, err
	}
	out := make([]string, 0, (col2-col1+1)*(row2-row1+1))
	for r := row1; r <= row2; r++ {
		for c := col1; c <= col2; c++ {
			name, err := werkbook.CoordinatesToCellName(c, r)
			if err != nil {
				return nil, err
			}
			out = append(out, name)
		}
	}
	return out, nil
}

func formatIndexList(idx []int) string {
	parts := make([]string, len(idx))
	for i, n := range idx {
		parts[i] = fmt.Sprintf("%d", n)
	}
	return strings.Join(parts, ",")
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
