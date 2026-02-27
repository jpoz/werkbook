package main

import (
	"encoding/json"
	"fmt"
	"os"
	"strings"
)

// TestSpec defines a complete fuzz test case.
type TestSpec struct {
	Name   string      `json:"name"`
	Sheets []SheetSpec `json:"sheets"`
	Checks []CheckSpec `json:"checks"`
}

// SheetSpec defines one sheet within a test spec.
type SheetSpec struct {
	Name  string     `json:"name"`
	Cells []CellSpec `json:"cells"`
}

// CellSpec defines a single cell in the spec.
type CellSpec struct {
	Ref     string  `json:"ref"`
	Value   any     `json:"value,omitempty"`
	Type    string  `json:"type,omitempty"` // "number", "string", "bool"
	Formula string  `json:"formula,omitempty"`
}

// CheckSpec defines an expected result to verify.
type CheckSpec struct {
	Ref      string `json:"ref"`      // "Sheet1!B1" or "B1" (defaults to first sheet)
	Expected string `json:"expected"`
	Type     string `json:"type"` // "number", "string", "bool", "error", "empty"
}

// excludedFunctions lists functions that should not appear in specs.
var excludedFunctions = map[string]bool{
	// Non-deterministic
	"RAND":        true,
	"RANDBETWEEN": true,
	"NOW":         true,
	"TODAY":       true,
	"INDIRECT":    true,
	// Excel 2019+/365 — not supported by LibreOffice
	"CONCAT":  true,
	"XLOOKUP": true,
	"SORT":    true,
}

// loadSpec reads and validates a test spec from a JSON file.
func loadSpec(path string) (*TestSpec, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("read spec: %w", err)
	}
	var spec TestSpec
	if err := json.Unmarshal(data, &spec); err != nil {
		return nil, fmt.Errorf("parse spec: %w", err)
	}
	if err := validateSpec(&spec); err != nil {
		return nil, err
	}
	return &spec, nil
}

// saveSpec writes a test spec to a JSON file.
func saveSpec(path string, spec *TestSpec) error {
	data, err := json.MarshalIndent(spec, "", "  ")
	if err != nil {
		return fmt.Errorf("marshal spec: %w", err)
	}
	return os.WriteFile(path, data, 0644)
}

// validateSpec checks that a spec is well-formed.
func validateSpec(spec *TestSpec) error {
	if spec.Name == "" {
		return fmt.Errorf("spec missing name")
	}
	if len(spec.Sheets) == 0 {
		return fmt.Errorf("spec has no sheets")
	}
	if len(spec.Checks) == 0 {
		return fmt.Errorf("spec has no checks")
	}
	for i, sheet := range spec.Sheets {
		if sheet.Name == "" {
			return fmt.Errorf("sheet %d missing name", i)
		}
		for j, cell := range sheet.Cells {
			if cell.Ref == "" {
				return fmt.Errorf("sheet %q cell %d missing ref", sheet.Name, j)
			}
			if cell.Formula == "" && cell.Value == nil {
				return fmt.Errorf("sheet %q cell %q has neither value nor formula", sheet.Name, cell.Ref)
			}
			// Check for excluded functions in formulas.
			if cell.Formula != "" {
				upper := strings.ToUpper(cell.Formula)
				for fn := range excludedFunctions {
					if strings.Contains(upper, fn+"(") {
						return fmt.Errorf("sheet %q cell %q uses excluded function %s", sheet.Name, cell.Ref, fn)
					}
				}
			}
		}
	}
	for i, check := range spec.Checks {
		if check.Ref == "" {
			return fmt.Errorf("check %d missing ref", i)
		}
		if check.Type == "" {
			return fmt.Errorf("check %d missing type", i)
		}
	}
	return nil
}

// parseCheckRef splits "Sheet1!B1" into (sheet, cellRef). If no sheet prefix, returns "".
func parseCheckRef(ref string) (sheet, cellRef string) {
	if idx := strings.Index(ref, "!"); idx >= 0 {
		return ref[:idx], ref[idx+1:]
	}
	return "", ref
}
