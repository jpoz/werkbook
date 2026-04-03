package main

import (
	"errors"
	"fmt"
	"math"
	"os"
	"strings"

	werkbook "github.com/jpoz/werkbook"
)

type checkDiff struct {
	Sheet   string `json:"sheet"`
	Cell    string `json:"cell"`
	Formula string `json:"formula"`
	Cached  any    `json:"cached"`
	Computed any    `json:"computed"`
}

type checkData struct {
	File       string      `json:"file"`
	Formulas   int         `json:"formulas"`
	Matches    int         `json:"matches"`
	Mismatches int         `json:"mismatches"`
	Diffs      []checkDiff `json:"diffs,omitempty"`
}

func cmdCheck(args []string, globals globalFlags) int {
	cmd := "check"

	if hasHelpFlag(args) {
		return writeHelpTopic([]string{cmd}, globals)
	}
	if !ensureFormat(cmd, globals, FormatText, FormatJSON) {
		return ExitUsage
	}

	var sheetFlag string
	var toleranceFlag float64

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
		case "--tolerance":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--tolerance requires a value"), globals)
				return ExitUsage
			}
			var err error
			toleranceFlag, err = parseFloat(args[i+1])
			if err != nil {
				writeError(cmd, errUsage(fmt.Sprintf("invalid --tolerance value: %s", args[i+1])), globals)
				return ExitUsage
			}
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

	// Determine which sheets to check.
	var sheetNames []string
	if sheetFlag != "" {
		if f.Sheet(sheetFlag) == nil {
			writeError(cmd, errSheetNotFound(sheetFlag), globals)
			return ExitValidate
		}
		sheetNames = []string{sheetFlag}
	} else {
		sheetNames = f.SheetNames()
	}

	// Collect cached values for all formula cells before recalculation.
	type cellID struct {
		sheet string
		ref   string
	}
	type cachedEntry struct {
		formula string
		value   werkbook.Value
	}
	cached := make(map[cellID]cachedEntry)

	for _, name := range sheetNames {
		s := f.Sheet(name)
		if s == nil {
			continue
		}
		for row := range s.Rows() {
			for _, cell := range row.Cells() {
				formula := cell.Formula()
				if formula == "" {
					continue
				}
				ref, err := werkbook.CoordinatesToCellName(cell.Col(), row.Num())
				if err != nil {
					continue
				}
				// Use GetValue to resolve dirty cells whose cached values
				// are stale due to uncached dynamic array spill data.
				v, _ := s.GetValue(ref)
				cached[cellID{sheet: name, ref: ref}] = cachedEntry{
					formula: formula,
					value:   v,
				}
			}
		}
	}

	// Recalculate.
	f.Recalculate()

	// Compare.
	var diffs []checkDiff
	matches := 0

	for _, name := range sheetNames {
		s := f.Sheet(name)
		if s == nil {
			continue
		}
		for row := range s.Rows() {
			for _, cell := range row.Cells() {
				if cell.Formula() == "" {
					continue
				}
				ref, err := werkbook.CoordinatesToCellName(cell.Col(), row.Num())
				if err != nil {
					continue
				}
				id := cellID{sheet: name, ref: ref}
				entry, ok := cached[id]
				if !ok {
					continue
				}

				computed, _ := s.GetValue(ref)

				if valuesEqual(entry.value, computed, toleranceFlag) {
					matches++
				} else {
					diffs = append(diffs, checkDiff{
						Sheet:   name,
						Cell:    ref,
						Formula: entry.formula,
						Cached:  entry.value.Raw(),
						Computed: computed.Raw(),
					})
				}
			}
		}
	}

	if diffs == nil {
		diffs = []checkDiff{}
	}

	data := checkData{
		File:       filePath,
		Formulas:   len(cached),
		Matches:    matches,
		Mismatches: len(diffs),
		Diffs:      diffs,
	}

	writeSuccess(cmd, data, globals)
	return ExitSuccess
}

func valuesEqual(a, b werkbook.Value, tolerance float64) bool {
	if a.Type != b.Type {
		return false
	}
	switch a.Type {
	case werkbook.TypeNumber:
		if tolerance > 0 {
			return math.Abs(a.Number-b.Number) <= tolerance
		}
		return a.Number == b.Number
	case werkbook.TypeString:
		return a.String == b.String
	case werkbook.TypeBool:
		return a.Bool == b.Bool
	case werkbook.TypeError:
		return a.String == b.String
	case werkbook.TypeEmpty:
		return true
	default:
		return a.Raw() == b.Raw()
	}
}

func parseFloat(s string) (float64, error) {
	var f float64
	_, err := fmt.Sscanf(s, "%f", &f)
	return f, err
}

func renderCheckText(data checkData) string {
	var sb strings.Builder
	sb.WriteString("File: ")
	sb.WriteString(data.File)
	sb.WriteString(fmt.Sprintf("\nFormulas: %d", data.Formulas))
	sb.WriteString(fmt.Sprintf("\nMatches: %d", data.Matches))
	sb.WriteString(fmt.Sprintf("\nMismatches: %d", data.Mismatches))

	if len(data.Diffs) > 0 {
		sb.WriteString("\n\nDifferences\n")
		rows := make([][]string, 0, len(data.Diffs))
		for _, d := range data.Diffs {
			rows = append(rows, []string{
				d.Sheet,
				d.Cell,
				d.Formula,
				displayRaw(d.Cached),
				displayRaw(d.Computed),
			})
		}
		sb.WriteString(renderTabular(
			[]string{"Sheet", "Cell", "Formula", "Cached", "Computed"},
			rows,
		))
	}

	return strings.TrimRight(sb.String(), "\n")
}

func displayRaw(v any) string {
	if v == nil {
		return "(empty)"
	}
	return fmt.Sprintf("%v", v)
}
