package fuzz

import (
	"encoding/json"
	"fmt"
	"strings"
)

// MinimizeSpec reduces a failing spec to the minimal reproducing case.
// It keeps only mismatched checks and their referenced cells, removing
// passing checks and unreferenced data cells.
func MinimizeSpec(spec *TestSpec, mismatches []Mismatch) *TestSpec {
	if spec == nil || len(mismatches) == 0 {
		return spec
	}

	// Build set of mismatched refs.
	mismatchRefs := make(map[string]bool)
	for _, m := range mismatches {
		mismatchRefs[m.Ref] = true
		// Also match bare ref without sheet prefix.
		_, cellRef := ParseCheckRef(m.Ref)
		mismatchRefs[cellRef] = true
	}

	// Keep only checks that failed.
	var minChecks []CheckSpec
	for _, check := range spec.Checks {
		if mismatchRefs[check.Ref] {
			minChecks = append(minChecks, check)
			continue
		}
		// Also check bare ref.
		_, cellRef := ParseCheckRef(check.Ref)
		if mismatchRefs[cellRef] {
			minChecks = append(minChecks, check)
		}
	}

	if len(minChecks) == 0 {
		return spec
	}

	// Find all cells referenced by mismatched formulas.
	referencedCells := make(map[string]map[string]bool) // sheet -> set of cell refs
	for _, sheet := range spec.Sheets {
		referencedCells[sheet.Name] = make(map[string]bool)
	}

	// Collect formula cells for mismatched checks.
	formulaCells := make(map[string]string) // "Sheet!Ref" -> formula
	for _, sheet := range spec.Sheets {
		for _, cell := range sheet.Cells {
			if cell.Formula == "" {
				continue
			}
			key := sheet.Name + "!" + cell.Ref
			formulaCells[key] = cell.Formula
			// Also add bare ref for matching.
			formulaCells[cell.Ref] = cell.Formula
		}
	}

	// For each mismatched check, trace its cell dependencies.
	for _, check := range minChecks {
		sheetName, cellRef := ParseCheckRef(check.Ref)
		if sheetName == "" && len(spec.Sheets) > 0 {
			sheetName = spec.Sheets[0].Name
		}

		// Mark the formula cell itself as referenced.
		if _, ok := referencedCells[sheetName]; ok {
			referencedCells[sheetName][cellRef] = true
		}

		// Find the formula and trace its references.
		formula := formulaCells[check.Ref]
		if formula == "" {
			formula = formulaCells[cellRef]
		}
		if formula != "" {
			traceDeps(formula, sheetName, spec, referencedCells)
		}
	}

	// Build minimized spec with only referenced cells.
	var minSheets []SheetSpec
	for _, sheet := range spec.Sheets {
		refs := referencedCells[sheet.Name]
		if len(refs) == 0 {
			continue
		}

		var minCells []CellSpec
		for _, cell := range sheet.Cells {
			if refs[cell.Ref] {
				minCells = append(minCells, cell)
			}
		}

		if len(minCells) > 0 {
			minSheets = append(minSheets, SheetSpec{
				Name:  sheet.Name,
				Cells: minCells,
			})
		}
	}

	if len(minSheets) == 0 {
		return spec
	}

	// Deep copy to avoid modifying original.
	data, err := json.Marshal(&TestSpec{
		Name:   fmt.Sprintf("min_%s", spec.Name),
		Sheets: minSheets,
		Checks: minChecks,
	})
	if err != nil {
		return spec
	}
	var result TestSpec
	if err := json.Unmarshal(data, &result); err != nil {
		return spec
	}

	return &result
}

// traceDeps extracts cell references from a formula and marks them as referenced.
// Recursively traces through referenced formula cells.
func traceDeps(formula string, defaultSheet string, spec *TestSpec, referenced map[string]map[string]bool) {
	// Simple reference extraction: look for patterns like A1, B2, Sheet1!C3, A1:B5.
	upper := strings.ToUpper(formula)

	for _, sheet := range spec.Sheets {
		for _, cell := range sheet.Cells {
			ref := strings.ToUpper(cell.Ref)

			// Check if this cell is referenced in the formula.
			// Check for Sheet!Ref pattern.
			sheetRef := strings.ToUpper(sheet.Name) + "!" + ref
			if strings.Contains(upper, sheetRef) {
				if _, ok := referenced[sheet.Name]; ok {
					if !referenced[sheet.Name][cell.Ref] {
						referenced[sheet.Name][cell.Ref] = true
						// Recursively trace if this is also a formula cell.
						if cell.Formula != "" {
							traceDeps(cell.Formula, sheet.Name, spec, referenced)
						}
					}
				}
				continue
			}

			// Check for bare ref (same sheet).
			if sheet.Name == defaultSheet && containsCellRef(upper, ref) {
				if _, ok := referenced[sheet.Name]; ok {
					if !referenced[sheet.Name][cell.Ref] {
						referenced[sheet.Name][cell.Ref] = true
						if cell.Formula != "" {
							traceDeps(cell.Formula, sheet.Name, spec, referenced)
						}
					}
				}
			}
		}
	}
}

// containsCellRef checks if a formula contains a cell reference,
// avoiding false positives from function names.
func containsCellRef(formula, ref string) bool {
	idx := 0
	for {
		pos := strings.Index(formula[idx:], ref)
		if pos < 0 {
			return false
		}
		pos += idx

		// Check that the ref isn't part of a longer identifier.
		// Before: must be start of string or non-alphanumeric.
		if pos > 0 {
			prev := formula[pos-1]
			if (prev >= 'A' && prev <= 'Z') || (prev >= '0' && prev <= '9') || prev == '_' {
				idx = pos + len(ref)
				continue
			}
		}

		// After: must be end of string or non-alphanumeric (except : for ranges).
		end := pos + len(ref)
		if end < len(formula) {
			next := formula[end]
			if (next >= 'A' && next <= 'Z') || (next >= '0' && next <= '9') || next == '_' {
				idx = end
				continue
			}
		}

		return true
	}
}
