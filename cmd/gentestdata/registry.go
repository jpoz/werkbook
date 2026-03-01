package main

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"github.com/jpoz/werkbook/internal/fuzz"
)

// TestCaseDef defines a single test case with its input data, formulas, and checks.
type TestCaseDef struct {
	FuncNames    []string        // Functions this test covers (for coverage)
	Description  string          // Human-readable label
	InputCells   []fuzz.CellSpec // Data cells (values)
	FormulaCells []fuzz.CellSpec // Formula cells
	Checks       []fuzz.CheckSpec
}

// ScenarioSet groups test cases for a category (e.g. "math", "stat").
type ScenarioSet struct {
	Name   string        // e.g. "math"
	Sheets []SheetDef    // Defaults to single "Sheet1" if nil
	Cases  []TestCaseDef // Test cases (stacked vertically via row offsets)
}

// SheetDef defines a named sheet with its own test cases.
type SheetDef struct {
	Name  string
	Cases []TestCaseDef
}

// offsetRef shifts a cell reference like "A1" by rowOffset rows.
// e.g. offsetRef("A1", 5) -> "A6", offsetRef("AB10", 3) -> "AB13"
func offsetRef(ref string, rowOffset int) string {
	col, row := splitRef(ref)
	return fmt.Sprintf("%s%d", col, row+rowOffset)
}

// splitRef splits "AB12" into ("AB", 12).
func splitRef(ref string) (string, int) {
	i := 0
	for i < len(ref) && ref[i] >= 'A' && ref[i] <= 'Z' {
		i++
	}
	col := ref[:i]
	row := 0
	for _, c := range ref[i:] {
		row = row*10 + int(c-'0')
	}
	return col, row
}

// cellRefPattern matches cell references like A1, AB123, Sheet1!A1, Data!A1:Data!A5
// The pattern requires the cell reference to NOT be preceded by another letter (to avoid
// matching function names like LOG10, DAYS360) and NOT be followed by '(' (functions).
// Group 1: optional sheet prefix (e.g. "Sheet1!")
// Group 2: column letters (A-XFD, 1-3 uppercase letters)
// Group 3: row digits
var cellRefPattern = regexp.MustCompile(`([A-Za-z]+!)?\$?([A-Z]{1,3})\$?(\d+)`)

// isValidColumn returns true if the column string is a valid Excel column (A-XFD).
func isValidColumn(col string) bool {
	if len(col) == 0 || len(col) > 3 {
		return false
	}
	// Max column is XFD = 16384
	if len(col) == 1 {
		return true
	}
	if len(col) == 2 {
		return true // AA-ZZ all valid
	}
	// 3 letters: AAA-XFD
	if col[0] > 'X' {
		return false
	}
	if col[0] == 'X' {
		if col[1] > 'F' {
			return false
		}
		if col[1] == 'F' && col[2] > 'D' {
			return false
		}
	}
	return true
}

// offsetFormula shifts all cell references in a formula by rowOffset rows.
// Handles A1, A1:B5, Sheet1!A1, Sheet1!A1:Sheet1!B5 patterns.
// Does NOT offset references that include a sheet name different from the current sheet
// (cross-sheet references are left unchanged).
// Carefully avoids matching function names (LOG10, DAYS360, etc.) by checking context.
func offsetFormula(formula string, rowOffset int, currentSheet string) string {
	if rowOffset == 0 {
		return formula
	}

	matches := cellRefPattern.FindAllStringSubmatchIndex(formula, -1)
	if len(matches) == 0 {
		return formula
	}

	var result strings.Builder
	lastEnd := 0

	// prevSheetPrefix tracks the sheet context from the left side of a range (e.g. "Data!" in "Data!A1:A5")
	prevSheetPrefix := ""

	for _, m := range matches {
		// m[0..1]: full match, m[2..3]: group 1 (sheet prefix), m[4..5]: group 2 (col), m[6..7]: group 3 (row)
		matchStart := m[0]
		matchEnd := m[1]
		sheetStart, sheetEnd := m[2], m[3]
		colStart, colEnd := m[4], m[5]
		rowStart, rowEnd := m[6], m[7]

		col := formula[colStart:colEnd]
		rowStr := formula[rowStart:rowEnd]

		sheetPrefix := ""
		if sheetStart >= 0 {
			sheetPrefix = formula[sheetStart:sheetEnd]
		}

		// Skip if preceded by a letter (part of a function name like LOG10, DAYS360)
		if sheetPrefix == "" && matchStart > 0 {
			prev := formula[matchStart-1]
			if (prev >= 'A' && prev <= 'Z') || (prev >= 'a' && prev <= 'z') {
				result.WriteString(formula[lastEnd:matchEnd])
				lastEnd = matchEnd
				prevSheetPrefix = ""
				continue
			}
		}

		// Skip if followed by '(' (function name like LOG10(), DAYS360())
		if matchEnd < len(formula) && formula[matchEnd] == '(' {
			result.WriteString(formula[lastEnd:matchEnd])
			lastEnd = matchEnd
			prevSheetPrefix = ""
			continue
		}

		// Skip if not a valid Excel column
		if !isValidColumn(col) {
			result.WriteString(formula[lastEnd:matchEnd])
			lastEnd = matchEnd
			prevSheetPrefix = ""
			continue
		}

		// Determine effective sheet for this reference.
		// If no explicit sheet prefix but preceded by ':', inherit from the left side of the range.
		effectiveSheet := sheetPrefix
		if effectiveSheet == "" && matchStart > 0 && formula[matchStart-1] == ':' && prevSheetPrefix != "" {
			effectiveSheet = prevSheetPrefix
		}

		// If the reference has a sheet prefix that differs from currentSheet, don't offset
		if effectiveSheet != "" {
			sheetName := strings.TrimSuffix(effectiveSheet, "!")
			if sheetName != currentSheet {
				result.WriteString(formula[lastEnd:matchEnd])
				lastEnd = matchEnd
				// Track for potential range right side
				prevSheetPrefix = effectiveSheet
				continue
			}
		}

		// Track the sheet prefix for the next match (in case it's the right side of a range)
		if sheetPrefix != "" {
			prevSheetPrefix = sheetPrefix
		} else if matchEnd < len(formula) && formula[matchEnd] == ':' {
			// This ref is followed by ':', so the next ref inherits nothing new
			// (it could only inherit from an explicit sheet prefix on this ref)
			prevSheetPrefix = ""
		} else {
			prevSheetPrefix = ""
		}

		row, err := strconv.Atoi(rowStr)
		if err != nil {
			result.WriteString(formula[lastEnd:matchEnd])
			lastEnd = matchEnd
			continue
		}

		// Write everything up to the match start, then the offset reference
		result.WriteString(formula[lastEnd:matchStart])
		result.WriteString(fmt.Sprintf("%s%s%d", sheetPrefix, col, row+rowOffset))
		lastEnd = matchEnd
	}

	result.WriteString(formula[lastEnd:])
	return result.String()
}

// cellList generates "A1,A2,A3,...,AN" style references.
func cellList(col string, startRow, endRow int) string {
	parts := make([]string, 0, endRow-startRow+1)
	for r := startRow; r <= endRow; r++ {
		parts = append(parts, fmt.Sprintf("%s%d", col, r))
	}
	return strings.Join(parts, ",")
}

// cellRange generates "A1:AN" style references.
func cellRange(col string, startRow, endRow int) string {
	return fmt.Sprintf("%s%d:%s%d", col, startRow, col, endRow)
}

// buildSpec converts a ScenarioSet into a fuzz.TestSpec by stacking test cases
// vertically with row offsets so they don't collide.
func buildSpec(ss ScenarioSet) *fuzz.TestSpec {
	spec := &fuzz.TestSpec{
		Name: ss.Name,
	}

	// If Sheets are explicitly defined, use them directly (for cross-sheet scenarios)
	if len(ss.Sheets) > 0 {
		for _, sd := range ss.Sheets {
			sheet := fuzz.SheetSpec{Name: sd.Name}
			rowOffset := 0
			for _, tc := range sd.Cases {
				maxRow := 0
				for _, c := range tc.InputCells {
					shifted := fuzz.CellSpec{
						Ref:   offsetRef(c.Ref, rowOffset),
						Value: c.Value,
						Type:  c.Type,
					}
					sheet.Cells = append(sheet.Cells, shifted)
					_, r := splitRef(c.Ref)
					if r > maxRow {
						maxRow = r
					}
				}
				for _, c := range tc.FormulaCells {
					shifted := fuzz.CellSpec{
						Ref:     offsetRef(c.Ref, rowOffset),
						Formula: offsetFormula(c.Formula, rowOffset, sd.Name),
					}
					sheet.Cells = append(sheet.Cells, shifted)
					_, r := splitRef(c.Ref)
					if r > maxRow {
						maxRow = r
					}
				}
				for _, ch := range tc.Checks {
					ref := ch.Ref
					// If check ref includes sheet name, offset the cell part
					if parts := strings.SplitN(ref, "!", 2); len(parts) == 2 {
						ref = parts[0] + "!" + offsetRef(parts[1], rowOffset)
					} else {
						ref = sd.Name + "!" + offsetRef(ref, rowOffset)
					}
					spec.Checks = append(spec.Checks, fuzz.CheckSpec{
						Ref:  ref,
						Type: ch.Type,
					})
				}
				rowOffset += maxRow + 1
			}
			spec.Sheets = append(spec.Sheets, sheet)
		}
		return spec
	}

	// Single-sheet mode: stack all Cases on "Sheet1"
	sheet := fuzz.SheetSpec{Name: "Sheet1"}
	rowOffset := 0
	for _, tc := range ss.Cases {
		maxRow := 0
		for _, c := range tc.InputCells {
			shifted := fuzz.CellSpec{
				Ref:   offsetRef(c.Ref, rowOffset),
				Value: c.Value,
				Type:  c.Type,
			}
			sheet.Cells = append(sheet.Cells, shifted)
			_, r := splitRef(c.Ref)
			if r > maxRow {
				maxRow = r
			}
		}
		for _, c := range tc.FormulaCells {
			shifted := fuzz.CellSpec{
				Ref:     offsetRef(c.Ref, rowOffset),
				Formula: offsetFormula(c.Formula, rowOffset, "Sheet1"),
			}
			sheet.Cells = append(sheet.Cells, shifted)
			_, r := splitRef(c.Ref)
			if r > maxRow {
				maxRow = r
			}
		}
		for _, ch := range tc.Checks {
			ref := ch.Ref
			if !strings.Contains(ref, "!") {
				ref = "Sheet1!" + offsetRef(ref, rowOffset)
			} else {
				parts := strings.SplitN(ref, "!", 2)
				ref = parts[0] + "!" + offsetRef(parts[1], rowOffset)
			}
			spec.Checks = append(spec.Checks, fuzz.CheckSpec{
				Ref:  ref,
				Type: ch.Type,
			})
		}
		rowOffset += maxRow + 1
	}
	spec.Sheets = append(spec.Sheets, sheet)
	return spec
}

// numCheck is a shorthand for creating a numeric check on a cell.
func numCheck(ref string) fuzz.CheckSpec {
	return fuzz.CheckSpec{Ref: ref, Type: "number"}
}

// strCheck is a shorthand for creating a string check on a cell.
func strCheck(ref string) fuzz.CheckSpec {
	return fuzz.CheckSpec{Ref: ref, Type: "string"}
}

// boolCheck is a shorthand for creating a boolean check on a cell.
func boolCheck(ref string) fuzz.CheckSpec {
	return fuzz.CheckSpec{Ref: ref, Type: "bool"}
}

// errCheck is a shorthand for creating an error check on a cell.
func errCheck(ref string) fuzz.CheckSpec {
	return fuzz.CheckSpec{Ref: ref, Type: "error"}
}

// numCell creates a CellSpec with a numeric value.
func numCell(ref string, v float64) fuzz.CellSpec {
	return fuzz.CellSpec{Ref: ref, Value: v, Type: "number"}
}

// strCell creates a CellSpec with a string value.
func strCell(ref string, v string) fuzz.CellSpec {
	return fuzz.CellSpec{Ref: ref, Value: v, Type: "string"}
}

// boolCell creates a CellSpec with a boolean value.
func boolCell(ref string, v bool) fuzz.CellSpec {
	return fuzz.CellSpec{Ref: ref, Value: v, Type: "bool"}
}

// formulaCell creates a CellSpec with a formula.
func formulaCell(ref string, formula string) fuzz.CellSpec {
	return fuzz.CellSpec{Ref: ref, Formula: formula}
}
