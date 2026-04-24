package main

import (
	"encoding/json"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"time"

	werkbook "github.com/jpoz/werkbook"
)

// checkConfig holds optional configuration for the check command, typically
// loaded from a JSON file via --config.
type checkConfig struct {
	Tolerance      float64  `json:"tolerance"`
	IgnoreFormulas []string `json:"ignore_formulas"`
	IgnoreFiles    []string `json:"ignore_files"`
}

func loadCheckConfig(path string) (checkConfig, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return checkConfig{}, fmt.Errorf("reading config %q: %v", path, err)
	}
	var cfg checkConfig
	if err := json.Unmarshal(data, &cfg); err != nil {
		return checkConfig{}, fmt.Errorf("parsing config %q: %v", path, err)
	}
	// Normalize ignore_formulas to upper case for case-insensitive matching.
	for i, f := range cfg.IgnoreFormulas {
		cfg.IgnoreFormulas[i] = strings.ToUpper(f)
	}
	return cfg, nil
}

// shouldIgnoreFile returns true if filePath matches any ignore_files glob pattern.
func (c *checkConfig) shouldIgnoreFile(filePath string) bool {
	for _, pattern := range c.IgnoreFiles {
		if matched, _ := filepath.Match(pattern, filepath.Base(filePath)); matched {
			return true
		}
		// Also try matching against the full path for patterns with directories.
		if matched, _ := filepath.Match(pattern, filePath); matched {
			return true
		}
	}
	return false
}

// shouldIgnoreFormula returns true if the formula contains any ignored function name.
func (c *checkConfig) shouldIgnoreFormula(formula string) bool {
	if len(c.IgnoreFormulas) == 0 {
		return false
	}
	upper := strings.ToUpper(formula)
	for _, fn := range c.IgnoreFormulas {
		if strings.Contains(upper, fn+"(") {
			return true
		}
	}
	return false
}

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

type checkMultiData struct {
	Files        int         `json:"files"`
	Skipped      int         `json:"skipped"`
	Formulas     int         `json:"formulas"`
	Matches      int         `json:"matches"`
	Mismatches   int         `json:"mismatches"`
	Errors       int         `json:"errors"`
	Results      []checkData `json:"results"`
	SkippedFiles []string    `json:"skipped_files,omitempty"`
	FileErrors   []fileError `json:"file_errors,omitempty"`
}

type fileError struct {
	File  string `json:"file"`
	Error string `json:"error"`
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
	var toleranceSet bool
	var configPath string
	var verbose bool
	var jobs int

	i := 0
	var paths []string
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
			toleranceSet = true
			i += 2
		case "--config":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--config requires a value"), globals)
				return ExitUsage
			}
			configPath = args[i+1]
			i += 2
		case "--verbose", "-v":
			verbose = true
			i++
		case "--jobs", "-j":
			if i+1 >= len(args) {
				writeError(cmd, errUsage("--jobs requires a value"), globals)
				return ExitUsage
			}
			n, err := parseInt(args[i+1])
			if err != nil || n < 1 {
				writeError(cmd, errUsage(fmt.Sprintf("invalid --jobs value: %s (must be a positive integer)", args[i+1])), globals)
				return ExitUsage
			}
			jobs = n
			i += 2
		default:
			if len(args[i]) > 0 && args[i][0] != '-' {
				paths = append(paths, args[i])
				i++
			} else {
				writeError(cmd, errUsage("unknown flag: "+args[i]), globals)
				return ExitUsage
			}
		}
	}

	// Load config file if provided.
	var cfg checkConfig
	if configPath != "" {
		var err error
		cfg, err = loadCheckConfig(configPath)
		if err != nil {
			writeError(cmd, errUsage(err.Error()), globals)
			return ExitUsage
		}
	}

	// CLI --tolerance overrides config file tolerance.
	if toleranceSet {
		cfg.Tolerance = toleranceFlag
	}

	if len(paths) == 0 {
		writeError(cmd, errUsage("file or directory path required"), globals)
		return ExitUsage
	}

	// Expand paths: directories become all .xlsx files within them.
	filePaths, err := expandXLSXPaths(paths)
	if err != nil {
		writeError(cmd, errUsage(err.Error()), globals)
		return ExitUsage
	}
	if len(filePaths) == 0 {
		writeError(cmd, errUsage("no .xlsx files found in the provided paths"), globals)
		return ExitUsage
	}

	// Single file: use the original single-file output for backward compatibility.
	if len(filePaths) == 1 {
		return cmdCheckSingle(filePaths[0], sheetFlag, &cfg, cmd, globals, verbose)
	}

	// Multiple files: run in parallel, aggregate results in input order so
	// the non-verbose output is byte-identical to the serial implementation.
	if jobs <= 0 {
		jobs = runtime.NumCPU()
	}
	if jobs > len(filePaths) {
		jobs = len(filePaths)
	}

	type checkJob struct {
		idx  int
		path string
	}
	type checkJobResult struct {
		idx     int
		path    string
		ignored bool
		skipped bool
		errMsg  string
		data    checkData
		elapsed time.Duration
	}

	jobCh := make(chan checkJob)
	resCh := make(chan checkJobResult, len(filePaths))

	for range jobs {
		go func() {
			for j := range jobCh {
				if cfg.shouldIgnoreFile(j.path) {
					resCh <- checkJobResult{idx: j.idx, path: j.path, ignored: true}
					continue
				}
				start := time.Now()
				result, skipped, ferr := checkFile(j.path, sheetFlag, &cfg)
				resCh <- checkJobResult{
					idx:     j.idx,
					path:    j.path,
					skipped: skipped,
					errMsg:  ferr,
					data:    result,
					elapsed: time.Since(start),
				}
			}
		}()
	}

	go func() {
		for i, fp := range filePaths {
			jobCh <- checkJob{idx: i, path: fp}
		}
		close(jobCh)
	}()

	results := make([]checkJobResult, len(filePaths))
	for range filePaths {
		r := <-resCh
		results[r.idx] = r
		if verbose {
			// Single consumer — no mutex needed.
			switch {
			case r.ignored:
				fmt.Fprintf(os.Stderr, "%s — ignored\n", r.path)
			case r.errMsg != "":
				fmt.Fprintf(os.Stderr, "%s — error: %s\n", r.path, r.errMsg)
			case r.skipped:
				fmt.Fprintf(os.Stderr, "%s — skipped (uncached) in %s\n", r.path, formatCheckDuration(r.elapsed))
			default:
				fmt.Fprintf(os.Stderr, "%s — %d formulas in %s\n", r.path, r.data.Formulas, formatCheckDuration(r.elapsed))
			}
		}
	}

	multi := checkMultiData{}
	for _, r := range results {
		if r.ignored {
			continue
		}
		if r.errMsg != "" {
			multi.Errors++
			multi.FileErrors = append(multi.FileErrors, fileError{File: r.path, Error: r.errMsg})
			continue
		}
		if r.skipped {
			multi.Skipped++
			multi.SkippedFiles = append(multi.SkippedFiles, r.path)
			continue
		}
		multi.Files++
		multi.Formulas += r.data.Formulas
		multi.Matches += r.data.Matches
		multi.Mismatches += r.data.Mismatches
		multi.Results = append(multi.Results, r.data)
	}

	writeSuccess(cmd, multi, globals)
	if multi.Mismatches > 0 || multi.Errors > 0 {
		return ExitValidate
	}
	return ExitSuccess
}

// expandXLSXPaths takes a list of file/directory paths and returns all .xlsx file paths.
func expandXLSXPaths(paths []string) ([]string, error) {
	var result []string
	for _, p := range paths {
		info, err := os.Stat(p)
		if err != nil {
			return nil, fmt.Errorf("cannot access %q: %v", p, err)
		}
		if !info.IsDir() {
			result = append(result, p)
			continue
		}
		err = filepath.Walk(p, func(path string, fi os.FileInfo, err error) error {
			if err != nil {
				return nil
			}
			if !fi.IsDir() && strings.EqualFold(filepath.Ext(path), ".xlsx") && !strings.HasPrefix(fi.Name(), "~$") {
				result = append(result, path)
			}
			return nil
		})
		if err != nil {
			return nil, fmt.Errorf("walking directory %q: %v", p, err)
		}
	}
	return result, nil
}

func cmdCheckSingle(filePath, sheetFlag string, cfg *checkConfig, cmd string, globals globalFlags, verbose bool) int {
	if cfg.shouldIgnoreFile(filePath) {
		if verbose {
			fmt.Fprintf(os.Stderr, "%s — ignored\n", filePath)
		}
		writeSuccess(cmd, checkData{File: filePath}, globals)
		return ExitSuccess
	}
	start := time.Now()
	result, skipped, ferr := checkFile(filePath, sheetFlag, cfg)
	elapsed := time.Since(start)
	if ferr != "" {
		if verbose {
			fmt.Fprintf(os.Stderr, "%s — error: %s\n", filePath, ferr)
		}
		writeError(cmd, &ErrorInfo{Code: ErrCodeFileOpenFailed, Message: ferr}, globals)
		return ExitFileIO
	}
	if verbose {
		if skipped {
			fmt.Fprintf(os.Stderr, "%s — skipped (uncached) in %s\n", filePath, formatCheckDuration(elapsed))
		} else {
			fmt.Fprintf(os.Stderr, "%s — %d formulas in %s\n", filePath, result.Formulas, formatCheckDuration(elapsed))
		}
	}

	writeSuccess(cmd, result, globals)
	return ExitSuccess
}

// formatCheckDuration renders an elapsed duration in a compact, stable form
// suitable for per-file verbose progress output.
func formatCheckDuration(d time.Duration) string {
	switch {
	case d >= time.Second:
		return fmt.Sprintf("%.2fs", d.Seconds())
	case d >= time.Millisecond:
		return fmt.Sprintf("%.1fms", float64(d)/float64(time.Millisecond))
	default:
		return fmt.Sprintf("%dµs", d.Microseconds())
	}
}

func checkFile(filePath, sheetFlag string, cfg *checkConfig) (result checkData, skipped bool, errMsg string) {
	defer func() {
		if r := recover(); r != nil {
			result = checkData{}
			skipped = false
			errMsg = fmt.Sprintf("panic processing %q: %v", filePath, r)
		}
	}()

	f, err := werkbook.Open(filePath, werkbook.WithSpillCache())
	if err != nil {
		return checkData{}, false, fmt.Sprintf("could not open %q: %v", filePath, err)
	}

	// Determine which sheets to check.
	var sheetNames []string
	if sheetFlag != "" {
		if f.Sheet(sheetFlag) == nil {
			return checkData{}, false, fmt.Sprintf("sheet %q not found in %q", sheetFlag, filePath)
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
				if formula == "" || isVolatileFormula(formula) || cfg.shouldIgnoreFormula(formula) {
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

				// For spill anchors, also cache all cells in the spill range.
				// Skip when the anchor's cached value is #SPILL! — the cells
				// in the range are blockers (user data), not spill shadow results.
				if cell.IsDynamicArraySpill() && !(v.Type == werkbook.TypeError && v.String == "#SPILL!") {
					toCol, toRow, ok := s.SpillBounds(cell.Col(), row.Num())
					if ok {
						for r := row.Num(); r <= toRow; r++ {
							for c := cell.Col(); c <= toCol; c++ {
								if c == cell.Col() && r == row.Num() {
									continue // anchor already cached
								}
								spillRef, err := werkbook.CoordinatesToCellName(c, r)
								if err != nil {
									continue
								}
								sv, _ := s.GetValue(spillRef)
								cached[cellID{sheet: name, ref: spillRef}] = cachedEntry{
									formula: formula,
									value:   sv,
								}
							}
						}
					}
				}
			}
		}
	}

	// Detect files with uncached formula values (e.g. written by exceljs,
	// xlsxwriter). If every formula cell has a zero/empty cached value and
	// there are more than 2 such cells, the file was never calculated, so
	// there is nothing meaningful to compare against.
	if len(cached) > 2 {
		allUncached := true
		for _, entry := range cached {
			v := entry.value
			switch v.Type {
			case werkbook.TypeEmpty:
				// empty counts as uncached
			case werkbook.TypeNumber:
				if v.Number != 0 {
					allUncached = false
				}
			case werkbook.TypeString:
				if v.String != "" {
					allUncached = false
				}
			default:
				allUncached = false
			}
			if !allUncached {
				break
			}
		}
		if allUncached {
			return checkData{File: filePath}, true, ""
		}
	}

	// Clear spill shadow values now that we have cached them, so the
	// engine can properly resolve spill results during recalculation.
	f.ClearSpillShadowValues()

	// Recalculate.
	f.Recalculate()

	// Compare.
	var diffs []checkDiff
	matches := 0

	for id, entry := range cached {
		s := f.Sheet(id.sheet)
		if s == nil {
			continue
		}
		computed, _ := s.GetValue(id.ref)

		if valuesEqual(entry.value, computed, cfg.Tolerance) {
			matches++
		} else {
			diffs = append(diffs, checkDiff{
				Sheet:    id.sheet,
				Cell:     id.ref,
				Formula:  entry.formula,
				Cached:   entry.value.Raw(),
				Computed: computed.Raw(),
			})
		}
	}

	if diffs == nil {
		diffs = []checkDiff{}
	}

	// Sort diffs by sheet, then row, then column for stable output.
	sort.Slice(diffs, func(i, j int) bool {
		if diffs[i].Sheet != diffs[j].Sheet {
			return diffs[i].Sheet < diffs[j].Sheet
		}
		ci, ri, _ := werkbook.CellNameToCoordinates(diffs[i].Cell)
		cj, rj, _ := werkbook.CellNameToCoordinates(diffs[j].Cell)
		if ri != rj {
			return ri < rj
		}
		return ci < cj
	})

	return checkData{
		File:       filePath,
		Formulas:   len(cached),
		Matches:    matches,
		Mismatches: len(diffs),
		Diffs:      diffs,
	}, false, ""
}

// isVolatileFormula returns true if the formula's result is expected to change
// on every recalculation (e.g. RAND, NOW, TODAY).
func isVolatileFormula(formula string) bool {
	upper := strings.ToUpper(formula)
	for _, fn := range volatileFuncs {
		if strings.Contains(upper, fn+"(") {
			return true
		}
	}
	return false
}

var volatileFuncs = []string{
	"RAND",
	"RANDARRAY",
	"RANDBETWEEN",
	"NOW",
	"TODAY",
}

func valuesEqual(a, b werkbook.Value, tolerance float64) bool {
	// Handle cross-type error comparisons: some xlsx writers store error
	// values (e.g. #N/A) as t="str" strings while the formula engine
	// produces TypeError values. Treat them as equal when the strings match.
	if a.Type != b.Type {
		if a.Type == werkbook.TypeError && b.Type == werkbook.TypeString {
			return a.String == b.String
		}
		if a.Type == werkbook.TypeString && b.Type == werkbook.TypeError {
			return a.String == b.String
		}
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

func parseInt(s string) (int, error) {
	var n int
	_, err := fmt.Sscanf(s, "%d", &n)
	return n, err
}

func renderCheckMultiText(data checkMultiData) string {
	var sb strings.Builder
	sb.WriteString(fmt.Sprintf("Files: %d", data.Files))
	if data.Skipped > 0 {
		sb.WriteString(fmt.Sprintf("\nSkipped (uncached): %d", data.Skipped))
		for _, fp := range data.SkippedFiles {
			sb.WriteString(fmt.Sprintf("\n  - %s", fp))
		}
	}
	sb.WriteString(fmt.Sprintf("\nFormulas: %d", data.Formulas))
	sb.WriteString(fmt.Sprintf("\nMatches: %d", data.Matches))
	sb.WriteString(fmt.Sprintf("\nMismatches: %d", data.Mismatches))
	if data.Errors > 0 {
		sb.WriteString(fmt.Sprintf("\nFile errors: %d", data.Errors))
	}

	// Show files with mismatches.
	for _, r := range data.Results {
		if r.Mismatches == 0 {
			continue
		}
		sb.WriteString(fmt.Sprintf("\n\n--- %s ---", r.File))
		sb.WriteString(fmt.Sprintf("\nFormulas: %d  Matches: %d  Mismatches: %d", r.Formulas, r.Matches, r.Mismatches))
		if len(r.Diffs) > 0 {
			sb.WriteString("\n")
			rows := make([][]string, 0, len(r.Diffs))
			for _, d := range r.Diffs {
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
	}

	// Show file errors.
	for _, fe := range data.FileErrors {
		sb.WriteString(fmt.Sprintf("\n\n--- %s ---", fe.File))
		sb.WriteString(fmt.Sprintf("\nError: %s", fe.Error))
	}

	return strings.TrimRight(sb.String(), "\n")
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
