package main

import (
	"errors"
	"fmt"
	"os"
	"strconv"
	"strings"

	werkbook "github.com/jpoz/werkbook"
	"github.com/jpoz/werkbook/formula"
)

type depData struct {
	File  string         `json:"file"`
	Sheet string         `json:"sheet"`
	Cells []depCellEntry `json:"cells"`
}

type depCellEntry struct {
	Ref        string         `json:"ref"`
	Sheet      string         `json:"sheet"`
	Formula    string         `json:"formula,omitempty"`
	Value      any            `json:"value"`
	Type       string         `json:"type"`
	Precedents []depCellBrief `json:"precedents,omitempty"`
	PrecRanges []depRange     `json:"precedent_ranges,omitempty"`
	Dependents []depCellBrief `json:"dependents,omitempty"`
}

type depRange struct {
	Range string         `json:"range"`
	Cells []depCellBrief `json:"cells"`
}

type depCellBrief struct {
	Ref     string `json:"ref"`
	Sheet   string `json:"sheet"`
	Formula string `json:"formula,omitempty"`
	Value   any    `json:"value"`
	Type    string `json:"type"`
}

func cmdDep(args []string, globals globalFlags) int {
	cmd := "dep"

	if hasHelpFlag(args) {
		return writeHelpTopic([]string{cmd}, globals)
	}
	if !ensureFormat(cmd, globals, FormatText, FormatJSON, FormatMarkdown) {
		return ExitUsage
	}

	var cellFlag, rangeFlag, sheetFlag, directionFlag string
	depthFlag := 1

	i := 0
	var filePath string
	for i < len(args) {
		switch args[i] {
		case "--cell":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--cell requires a value"), globals)
				return ExitUsage
			}
			cellFlag = args[i+1]
			i += 2
		case "--range":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--range requires a value"), globals)
				return ExitUsage
			}
			rangeFlag = args[i+1]
			i += 2
		case "--sheet":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--sheet requires a value"), globals)
				return ExitUsage
			}
			sheetFlag = args[i+1]
			i += 2
		case "--direction":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--direction requires a value"), globals)
				return ExitUsage
			}
			directionFlag = args[i+1]
			i += 2
		case "--depth":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--depth requires a value"), globals)
				return ExitUsage
			}
			n, err := strconv.Atoi(args[i+1])
			if err != nil {
				writeError(cmd, errUsage("--depth must be an integer"), globals)
				return ExitUsage
			}
			depthFlag = n
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

	if cellFlag != "" && rangeFlag != "" {
		writeError(cmd, errUsage("--cell and --range are mutually exclusive"), globals)
		return ExitUsage
	}

	// Validate direction.
	showPrec := true
	showDeps := true
	switch directionFlag {
	case "", "both":
		// default
	case "precedents":
		showDeps = false
	case "dependents":
		showPrec = false
	default:
		writeError(cmd, errUsage(fmt.Sprintf("unknown direction %q; use precedents, dependents, or both", directionFlag)), globals)
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

	// Recalculate so values are current and dep graph is fully populated.
	f.Recalculate()

	// Resolve sheet.
	sheetName := sheetFlag
	if sheetName == "" {
		names := f.SheetNames()
		if len(names) == 0 {
			writeError(cmd, errInternal(fmt.Errorf("workbook has no sheets")), globals)
			return ExitInternal
		}
		sheetName = names[0]
	}

	s := f.Sheet(sheetName)
	if s == nil {
		writeError(cmd, errSheetNotFound(sheetName), globals)
		return ExitValidate
	}

	// Build list of target cells.
	type cellCoord struct {
		col, row int
		ref      string
	}
	var targets []cellCoord

	if cellFlag != "" {
		col, row, cerr := werkbook.CellNameToCoordinates(cellFlag)
		if cerr != nil {
			writeError(cmd, errUsage(fmt.Sprintf("invalid cell %q: %v", cellFlag, cerr)), globals)
			return ExitUsage
		}
		targets = append(targets, cellCoord{col: col, row: row, ref: cellFlag})
	} else if rangeFlag != "" {
		col1, row1, col2, row2, rerr := werkbook.RangeToCoordinates(rangeFlag)
		if rerr != nil {
			writeError(cmd, errInvalidRange(rangeFlag, rerr), globals)
			return ExitValidate
		}
		for r := row1; r <= row2; r++ {
			for c := col1; c <= col2; c++ {
				ref, _ := werkbook.CoordinatesToCellName(c, r)
				frm, _ := s.GetFormula(ref)
				if frm != "" {
					targets = append(targets, cellCoord{col: c, row: r, ref: ref})
				}
			}
		}
	} else {
		// Auto-discover: find all formula cells on the sheet.
		for row := range s.Rows() {
			for _, cell := range row.Cells() {
				if cell.Formula() != "" {
					ref, _ := werkbook.CoordinatesToCellName(cell.Col(), row.Num())
					targets = append(targets, cellCoord{col: cell.Col(), row: row.Num(), ref: ref})
				}
			}
		}
	}

	// resolveValue returns a display-friendly value for a werkbook.Value.
	// Error values use the error string (e.g. "#REF!"); empty values use "".
	resolveValue := func(v werkbook.Value) any {
		switch v.Type {
		case werkbook.TypeError:
			return v.String
		case werkbook.TypeEmpty:
			return ""
		default:
			return v.Raw()
		}
	}

	// Helper to build a depCellBrief for a qualified cell.
	briefCell := func(qc formula.QualifiedCell) depCellBrief {
		ref, _ := werkbook.CoordinatesToCellName(qc.Col, qc.Row)
		sh := f.Sheet(qc.Sheet)
		var val any
		typ := "empty"
		var frm string
		if sh != nil {
			v, _ := sh.GetValue(ref)
			val = resolveValue(v)
			typ = valueTypeName(v)
			frm, _ = sh.GetFormula(ref)
		}
		return depCellBrief{
			Ref:     ref,
			Sheet:   qc.Sheet,
			Formula: frm,
			Value:   val,
			Type:    typ,
		}
	}

	// Collect transitive cells at a given depth using BFS.
	collectTransitive := func(seed []formula.QualifiedCell, nextFn func(formula.QualifiedCell) []formula.QualifiedCell, depth int) []depCellBrief {
		if depth == 0 {
			return nil
		}
		seen := make(map[formula.QualifiedCell]bool)
		current := seed
		var result []depCellBrief
		level := 0
		for len(current) > 0 {
			level++
			if depth > 0 && level > depth {
				break
			}
			var next []formula.QualifiedCell
			for _, qc := range current {
				if seen[qc] {
					continue
				}
				seen[qc] = true
				result = append(result, briefCell(qc))
				if depth < 0 || level < depth {
					next = append(next, nextFn(qc)...)
				}
			}
			current = next
		}
		return result
	}

	var entries []depCellEntry
	for _, tc := range targets {
		v, _ := s.GetValue(tc.ref)
		frm, _ := s.GetFormula(tc.ref)

		entry := depCellEntry{
			Ref:     tc.ref,
			Sheet:   sheetName,
			Formula: frm,
			Value:   resolveValue(v),
			Type:    valueTypeName(v),
		}

		if showPrec {
			precPoints, precRanges, _ := f.Precedents(sheetName, tc.ref)

			// Collect precedent point cells with transitive expansion.
			precSeed := precPoints
			entry.Precedents = collectTransitive(precSeed, func(qc formula.QualifiedCell) []formula.QualifiedCell {
				ref, _ := werkbook.CoordinatesToCellName(qc.Col, qc.Row)
				pts, _, _ := f.Precedents(qc.Sheet, ref)
				return pts
			}, depthFlag)

			// Precedent ranges with expanded cell values.
			for _, rng := range precRanges {
				from, _ := werkbook.CoordinatesToCellName(rng.FromCol, rng.FromRow)
				to, _ := werkbook.CoordinatesToCellName(rng.ToCol, rng.ToRow)
				rangeStr := rng.Sheet + "!" + from + ":" + to

				var cells []depCellBrief
				rngSheet := f.Sheet(rng.Sheet)
				if rngSheet != nil {
					for r := rng.FromRow; r <= rng.ToRow; r++ {
						for c := rng.FromCol; c <= rng.ToCol; c++ {
							ref, _ := werkbook.CoordinatesToCellName(c, r)
							v, _ := rngSheet.GetValue(ref)
							frm, _ := rngSheet.GetFormula(ref)
							if v.Type == werkbook.TypeEmpty && frm == "" {
								continue
							}
							cells = append(cells, depCellBrief{
								Ref:     ref,
								Sheet:   rng.Sheet,
								Formula: frm,
								Value:   resolveValue(v),
								Type:    valueTypeName(v),
							})
						}
					}
				}

				entry.PrecRanges = append(entry.PrecRanges, depRange{
					Range: rangeStr,
					Cells: cells,
				})
			}
		}

		if showDeps {
			directDeps, _ := f.DirectDependents(sheetName, tc.ref)

			entry.Dependents = collectTransitive(directDeps, func(qc formula.QualifiedCell) []formula.QualifiedCell {
				ref, _ := werkbook.CoordinatesToCellName(qc.Col, qc.Row)
				deps, _ := f.DirectDependents(qc.Sheet, ref)
				return deps
			}, depthFlag)
		}

		entries = append(entries, entry)
	}

	if entries == nil {
		entries = []depCellEntry{}
	}

	data := depData{
		File:  filePath,
		Sheet: sheetName,
		Cells: entries,
	}

	if globals.format == FormatText || globals.format == FormatMarkdown {
		fmt.Print(formatDepMarkdown(data))
		return ExitSuccess
	}

	writeSuccess(cmd, data, globals)
	return ExitSuccess
}

func formatDepMarkdown(data depData) string {
	var sb strings.Builder
	for i, cell := range data.Cells {
		if i > 0 {
			sb.WriteString("\n")
		}
		sb.WriteString("# ")
		sb.WriteString(cell.Sheet)
		sb.WriteString("!")
		sb.WriteString(cell.Ref)
		sb.WriteString("\n")

		if cell.Formula != "" {
			sb.WriteString("Formula: =")
			sb.WriteString(cell.Formula)
			sb.WriteString("\n")
		}
		sb.WriteString("Value: ")
		sb.WriteString(fmt.Sprintf("%v", cell.Value))
		sb.WriteString("\n")

		if len(cell.Precedents) > 0 || len(cell.PrecRanges) > 0 {
			sb.WriteString("\n## Precedents (cells this reads from):\n")
			for _, p := range cell.Precedents {
				sb.WriteString("- ")
				sb.WriteString(p.Sheet)
				sb.WriteString("!")
				sb.WriteString(p.Ref)
				sb.WriteString(" = ")
				sb.WriteString(fmt.Sprintf("%v", p.Value))
				if p.Formula != "" {
					sb.WriteString(" (formula: =")
					sb.WriteString(p.Formula)
					sb.WriteString(")")
				}
				sb.WriteString("\n")
			}
			for _, r := range cell.PrecRanges {
				sb.WriteString("- ")
				sb.WriteString(r.Range)
				sb.WriteString(" (range)\n")
				for _, c := range r.Cells {
					sb.WriteString("  - ")
					sb.WriteString(c.Ref)
					sb.WriteString(" = ")
					sb.WriteString(fmt.Sprintf("%v", c.Value))
					if c.Formula != "" {
						sb.WriteString(" (formula: =")
						sb.WriteString(c.Formula)
						sb.WriteString(")")
					}
					sb.WriteString("\n")
				}
			}
		}

		if len(cell.Dependents) > 0 {
			sb.WriteString("\n## Dependents (cells that read this):\n")
			for _, d := range cell.Dependents {
				sb.WriteString("- ")
				sb.WriteString(d.Sheet)
				sb.WriteString("!")
				sb.WriteString(d.Ref)
				sb.WriteString(" = ")
				sb.WriteString(fmt.Sprintf("%v", d.Value))
				if d.Formula != "" {
					sb.WriteString(" (formula: =")
					sb.WriteString(d.Formula)
					sb.WriteString(")")
				}
				sb.WriteString("\n")
			}
		}
	}
	return sb.String()
}
