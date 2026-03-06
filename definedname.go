package werkbook

import (
	"fmt"
	"strings"
)

// DefinedName represents an Excel defined name, including workbook-scoped
// and sheet-scoped named ranges.
type DefinedName struct {
	Name         string
	Value        string
	LocalSheetID int // -1 for workbook scope; otherwise 0-based sheet index
}

// DefinedNames returns the workbook's defined names in file order.
func (f *File) DefinedNames() []DefinedName {
	out := make([]DefinedName, len(f.definedNames))
	for i, dn := range f.definedNames {
		out[i] = DefinedName{
			Name:         dn.Name,
			Value:        dn.Value,
			LocalSheetID: dn.LocalSheetID,
		}
	}
	return out
}

// ResolveDefinedName looks up a defined name by its name and returns the
// resolved cell values as a 2D grid. For a single-cell reference the result
// is a 1x1 grid. The lookup is case-insensitive. If sheetIndex >= 0, a
// sheet-scoped name for that sheet takes precedence over a global name.
// Pass -1 for sheetIndex to match only workbook-scoped names.
func (f *File) ResolveDefinedName(name string, sheetIndex int) ([][]Value, error) {
	ref, err := f.lookupDefinedName(name, sheetIndex)
	if err != nil {
		return nil, err
	}

	sheetName, cellRef, err := parseDefinedNameRef(ref)
	if err != nil {
		return nil, fmt.Errorf("cannot resolve defined name %q: %w", name, err)
	}

	s := f.Sheet(sheetName)
	if s == nil {
		return nil, fmt.Errorf("cannot resolve defined name %q: sheet %q not found", name, sheetName)
	}

	// Determine if this is a range (contains ":") or a single cell.
	if strings.Contains(cellRef, ":") {
		col1, row1, col2, row2, err := RangeToCoordinates(cellRef)
		if err != nil {
			return nil, fmt.Errorf("cannot resolve defined name %q: %w", name, err)
		}
		rows := make([][]Value, row2-row1+1)
		for r := row1; r <= row2; r++ {
			cols := make([]Value, col2-col1+1)
			for c := col1; c <= col2; c++ {
				ref, _ := CoordinatesToCellName(c, r)
				v, _ := s.GetValue(ref)
				cols[c-col1] = v
			}
			rows[r-row1] = cols
		}
		return rows, nil
	}

	// Single cell.
	v, err := s.GetValue(cellRef)
	if err != nil {
		return nil, fmt.Errorf("cannot resolve defined name %q: %w", name, err)
	}
	return [][]Value{{v}}, nil
}

// lookupDefinedName finds the best-matching defined name, preferring a
// sheet-scoped name for sheetIndex over a global name.
func (f *File) lookupDefinedName(name string, sheetIndex int) (string, error) {
	lower := strings.ToLower(name)
	var globalVal string
	globalFound := false
	for _, dn := range f.definedNames {
		if strings.ToLower(dn.Name) != lower {
			continue
		}
		if dn.LocalSheetID == sheetIndex {
			return dn.Value, nil
		}
		if dn.LocalSheetID == -1 && !globalFound {
			globalVal = dn.Value
			globalFound = true
		}
	}
	if globalFound {
		return globalVal, nil
	}
	return "", fmt.Errorf("defined name %q not found", name)
}

// parseDefinedNameRef parses a defined name value like "Sheet1!$A$1:$C$10"
// into a sheet name and a cell/range reference with $ signs stripped.
func parseDefinedNameRef(ref string) (sheetName, cellRef string, err error) {
	ref = strings.TrimSpace(ref)
	if ref == "" {
		return "", "", fmt.Errorf("empty reference")
	}

	idx := strings.LastIndex(ref, "!")
	if idx < 0 {
		return "", "", fmt.Errorf("reference %q has no sheet qualifier", ref)
	}

	sheetName = ref[:idx]
	cellRef = ref[idx+1:]

	// Strip surrounding quotes from sheet name (e.g. 'My Sheet'!A1).
	if len(sheetName) >= 2 && sheetName[0] == '\'' && sheetName[len(sheetName)-1] == '\'' {
		sheetName = strings.ReplaceAll(sheetName[1:len(sheetName)-1], "''", "'")
	}

	// Strip $ signs from the cell reference.
	cellRef = strings.ReplaceAll(cellRef, "$", "")

	if sheetName == "" || cellRef == "" {
		return "", "", fmt.Errorf("invalid reference %q", ref)
	}
	return sheetName, cellRef, nil
}
