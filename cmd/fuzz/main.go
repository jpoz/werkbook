package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

func main() {
	rounds := flag.Int("rounds", 10, "number of fuzz rounds")
	fix := flag.Bool("fix", false, "auto-suggest fixes with claude -p")
	specFile := flag.String("spec", "", "run a single saved spec instead of generating")
	keep := flag.Bool("keep", false, "keep generated files")
	seed := flag.String("seed", "", "focus category: math|text|lookup|logic|stat|date|mixed")
	verbose := flag.Bool("v", false, "verbose output")
	flag.Parse()

	// Find LibreOffice.
	soffice, err := findLibreOffice()
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	if *verbose {
		fmt.Printf("Using LibreOffice: %s\n", soffice)
	}

	// Ensure failures directory exists.
	failDir := "failures"
	if err := os.MkdirAll(failDir, 0755); err != nil {
		fmt.Fprintf(os.Stderr, "Error creating failures dir: %v\n", err)
		os.Exit(1)
	}

	// Single spec mode.
	if *specFile != "" {
		spec, err := loadSpec(*specFile)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error loading spec: %v\n", err)
			os.Exit(1)
		}
		ok := runSpec(spec, soffice, failDir, *fix, *keep, *verbose)
		if !ok {
			os.Exit(1)
		}
		return
	}

	// Multi-round fuzz mode.
	passed, failed := 0, 0
	for i := 0; i < *rounds; i++ {
		fmt.Printf("\n=== Round %d/%d ===\n", i+1, *rounds)

		spec, err := generateSpec(seedCategory(*seed), *verbose)
		if err != nil {
			fmt.Printf("  Generation failed: %v\n", err)
			failed++
			continue
		}

		fmt.Printf("  Spec: %s (%d sheets, %d checks)\n", spec.Name, len(spec.Sheets), len(spec.Checks))

		if runSpec(spec, soffice, failDir, *fix, *keep, *verbose) {
			passed++
		} else {
			failed++
		}
	}

	fmt.Printf("\n=== Summary ===\n")
	fmt.Printf("Passed: %d  Failed: %d  Total: %d\n", passed, failed, passed+failed)
	if failed > 0 {
		os.Exit(1)
	}
}

// runSpec executes a single test spec: build, eval with LibreOffice, compare.
// Returns true if all checks pass.
func runSpec(spec *TestSpec, soffice, failDir string, fix, keep, verbose bool) bool {
	// Create a temp dir for this spec.
	tmpDir, err := os.MkdirTemp("", "werkbook-fuzz-*")
	if err != nil {
		fmt.Printf("  Error creating temp dir: %v\n", err)
		return false
	}
	if !keep {
		defer os.RemoveAll(tmpDir)
	} else {
		fmt.Printf("  Working dir: %s\n", tmpDir)
	}

	// Build XLSX and get werkbook results.
	xlsxPath, wbResults, err := buildXLSX(spec, tmpDir)
	if err != nil {
		fmt.Printf("  Build failed: %v\n", err)
		return false
	}

	if verbose {
		fmt.Printf("  XLSX: %s\n", xlsxPath)
		fmt.Printf("  Werkbook results:\n")
		for _, r := range wbResults {
			fmt.Printf("    %s = %q (%s)\n", r.Ref, r.Value, r.Type)
		}
	}

	// Evaluate with LibreOffice.
	loResults, err := libreOfficeEval(soffice, xlsxPath, spec.Checks)
	if err != nil {
		fmt.Printf("  LibreOffice eval failed: %v\n", err)
		return false
	}

	if verbose {
		fmt.Printf("  LibreOffice results:\n")
		for _, r := range loResults {
			fmt.Printf("    %s = %q (%s)\n", r.Ref, r.Value, r.Type)
		}
	}

	// Compare results.
	mismatches := compareResults(spec.Checks, wbResults, loResults)
	if len(mismatches) == 0 {
		fmt.Printf("  PASS: all %d checks match\n", len(spec.Checks))
		return true
	}

	fmt.Printf("  FAIL: %d/%d checks mismatched:\n", len(mismatches), len(spec.Checks))
	fmt.Print(formatMismatches(mismatches))

	// Save failure for reproduction.
	saveFailure(spec, xlsxPath, failDir)

	// Auto-fix mode.
	if fix {
		suggestion, err := suggestFix(spec, mismatches, verbose)
		if err != nil {
			fmt.Printf("  Fix suggestion failed: %v\n", err)
		} else {
			fmt.Printf("\n  Fix suggestion:\n%s\n", indent(suggestion, "    "))
		}
	}

	return false
}

// saveFailure persists a failing spec and its XLSX for later reproduction.
func saveFailure(spec *TestSpec, xlsxPath, failDir string) {
	specPath := filepath.Join(failDir, spec.Name+".json")
	if err := saveSpec(specPath, spec); err != nil {
		fmt.Printf("  Warning: could not save spec: %v\n", err)
		return
	}

	// Copy XLSX to failures dir.
	xlsxDst := filepath.Join(failDir, spec.Name+".xlsx")
	data, err := os.ReadFile(xlsxPath)
	if err == nil {
		os.WriteFile(xlsxDst, data, 0644)
	}

	fmt.Printf("  Saved: %s\n", specPath)
}

// seedCategory maps user-provided seed to a prompt category hint.
func seedCategory(seed string) string {
	switch strings.ToLower(seed) {
	case "math":
		return "math (ABS, CEILING, EXP, FLOOR, INT, LN, LOG, LOG10, MOD, PI, POWER, ROUND, ROUNDDOWN, ROUNDUP, SQRT, SUM, SUMPRODUCT, PRODUCT)"
	case "text":
		return "text (CHAR, CLEAN, CODE, CONCAT, CONCATENATE, EXACT, FIND, LEFT, LEN, LOWER, MID, PROPER, REPLACE, REPT, RIGHT, SEARCH, SUBSTITUTE, TEXT, TRIM, UPPER, VALUE)"
	case "lookup":
		return "lookup and reference (CHOOSE, COLUMN, COLUMNS, HLOOKUP, INDEX, LOOKUP, MATCH, ROW, ROWS, VLOOKUP, XLOOKUP)"
	case "logic":
		return "logical (AND, IF, IFERROR, IFNA, ISBLANK, ISERR, ISERROR, ISNA, ISNUMBER, ISTEXT, NOT, OR, XOR)"
	case "stat":
		return "statistical (AVERAGE, AVERAGEIF, AVERAGEIFS, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, LARGE, MAX, MEDIAN, MIN, SMALL)"
	case "date":
		return "date/time (DATE, DAY, HOUR, MINUTE, MONTH, SECOND, TIME, YEAR)"
	case "mixed", "":
		return ""
	default:
		return seed
	}
}

// indent prefixes each line with the given prefix.
func indent(s, prefix string) string {
	lines := strings.Split(s, "\n")
	for i, line := range lines {
		if line != "" {
			lines[i] = prefix + line
		}
	}
	return strings.Join(lines, "\n")
}

