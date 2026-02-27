package main

import (
	"encoding/json"
	"fmt"
	"path/filepath"

	"github.com/jpoz/werkbook"
)

// buildResult holds werkbook's evaluation result for a single check.
type buildResult struct {
	Ref   string
	Value string
	Type  string
}

// buildXLSX creates an XLSX file from the spec, recalculates, and collects results.
// Returns the XLSX path and werkbook results for each check.
func buildXLSX(spec *TestSpec, dir string) (string, []buildResult, error) {
	f := werkbook.New()

	// Create sheets. The first sheet is created by New(), so rename or reuse it.
	for i, ss := range spec.Sheets {
		var s *werkbook.Sheet
		if i == 0 {
			// werkbook.New() creates "Sheet1"; if spec wants a different name,
			// we delete and recreate. If same name, just use it.
			s = f.Sheet("Sheet1")
			if ss.Name != "Sheet1" {
				// Delete default and create with correct name.
				// But we can't delete the only sheet, so create new first.
				var err error
				s, err = f.NewSheet(ss.Name)
				if err != nil {
					return "", nil, fmt.Errorf("create sheet %q: %w", ss.Name, err)
				}
				if err := f.DeleteSheet("Sheet1"); err != nil {
					return "", nil, fmt.Errorf("delete default sheet: %w", err)
				}
			}
		} else {
			var err error
			s, err = f.NewSheet(ss.Name)
			if err != nil {
				return "", nil, fmt.Errorf("create sheet %q: %w", ss.Name, err)
			}
		}

		// Set cell values and formulas.
		for _, cs := range ss.Cells {
			if cs.Formula != "" {
				if err := s.SetFormula(cs.Ref, cs.Formula); err != nil {
					return "", nil, fmt.Errorf("set formula %s!%s: %w", ss.Name, cs.Ref, err)
				}
			} else {
				v := convertSpecValue(cs)
				if err := s.SetValue(cs.Ref, v); err != nil {
					return "", nil, fmt.Errorf("set value %s!%s: %w", ss.Name, cs.Ref, err)
				}
			}
		}
	}

	// Recalculate all formulas.
	f.Recalculate()

	// Collect results for each check.
	var results []buildResult
	for _, check := range spec.Checks {
		sheet, cellRef := parseCheckRef(check.Ref)
		if sheet == "" && len(spec.Sheets) > 0 {
			sheet = spec.Sheets[0].Name
		}
		s := f.Sheet(sheet)
		if s == nil {
			results = append(results, buildResult{
				Ref:   check.Ref,
				Value: "#SHEET_NOT_FOUND",
				Type:  "error",
			})
			continue
		}

		val, err := s.GetValue(cellRef)
		if err != nil {
			results = append(results, buildResult{
				Ref:   check.Ref,
				Value: fmt.Sprintf("#ERR:%v", err),
				Type:  "error",
			})
			continue
		}

		results = append(results, buildResult{
			Ref:   check.Ref,
			Value: formatValue(val),
			Type:  valueTypeName(val),
		})
	}

	// Save XLSX.
	xlsxPath := filepath.Join(dir, spec.Name+".xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		return "", nil, fmt.Errorf("save xlsx: %w", err)
	}

	return xlsxPath, results, nil
}

// convertSpecValue converts a CellSpec value to a Go value suitable for SetValue.
func convertSpecValue(cs CellSpec) any {
	switch cs.Type {
	case "number":
		switch v := cs.Value.(type) {
		case float64:
			return v
		case json.Number:
			f, _ := v.Float64()
			return f
		default:
			return cs.Value
		}
	case "bool":
		switch v := cs.Value.(type) {
		case bool:
			return v
		default:
			return cs.Value
		}
	case "string":
		switch v := cs.Value.(type) {
		case string:
			return v
		default:
			return fmt.Sprintf("%v", cs.Value)
		}
	default:
		return cs.Value
	}
}

// formatValue converts a werkbook Value to a display string.
func formatValue(v werkbook.Value) string {
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
	case werkbook.TypeEmpty:
		return ""
	default:
		return ""
	}
}

// valueTypeName returns a human-readable type name for a Value.
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
	case werkbook.TypeEmpty:
		return "empty"
	default:
		return "unknown"
	}
}
