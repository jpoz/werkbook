package fuzz

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"time"

	"github.com/jpoz/werkbook"
)

// FindLocalExcel verifies that Microsoft Excel is installed locally.
func FindLocalExcel() error {
	if _, err := os.Stat("/Applications/Microsoft Excel.app"); err != nil {
		return fmt.Errorf("Microsoft Excel not found at /Applications/Microsoft Excel.app: %w", err)
	}
	return nil
}

// LocalExcelEvaluator uses local Microsoft Excel via AppleScript as a formula oracle.
type LocalExcelEvaluator struct{}

// Name returns "local-excel".
func (e *LocalExcelEvaluator) Name() string { return "local-excel" }

// Eval opens the XLSX in local Excel via AppleScript, saves (triggering recalc),
// then re-reads cached values using werkbook.
func (e *LocalExcelEvaluator) Eval(xlsxPath string, checks []CheckSpec) ([]CellResult, error) {
	// Copy to a temp file with unique name to avoid Excel's "same name" conflict.
	data, err := os.ReadFile(xlsxPath)
	if err != nil {
		return nil, fmt.Errorf("read xlsx: %w", err)
	}
	tmpDir, err := os.MkdirTemp("", "local-excel-*")
	if err != nil {
		return nil, fmt.Errorf("create temp dir: %w", err)
	}
	defer os.RemoveAll(tmpDir)

	tmpPath := filepath.Join(tmpDir, fmt.Sprintf("eval_%d.xlsx", time.Now().UnixNano()))
	if err := os.WriteFile(tmpPath, data, 0644); err != nil {
		return nil, fmt.Errorf("write temp xlsx: %w", err)
	}

	absPath, err := filepath.Abs(tmpPath)
	if err != nil {
		return nil, fmt.Errorf("abs path: %w", err)
	}

	// Run AppleScript to open, recalculate, save, and close.
	script := fmt.Sprintf(`
set xlsxFile to POSIX file "%s"
tell application "Microsoft Excel"
	activate
	open xlsxFile
	delay 2
	save active workbook
	delay 1
	close active workbook saving no
end tell
`, absPath)

	cmd := exec.Command("osascript", "-e", script)
	output, err := cmd.CombinedOutput()
	if err != nil {
		return nil, fmt.Errorf("AppleScript failed: %w\noutput: %s", err, output)
	}

	// Re-read the saved XLSX.
	f, err := werkbook.Open(absPath)
	if err != nil {
		return nil, fmt.Errorf("re-open xlsx: %w", err)
	}

	// Extract cached values for each check.
	var results []CellResult
	for _, check := range checks {
		sheet, cellRef := ParseCheckRef(check.Ref)
		if sheet == "" {
			sheet = "Sheet1"
		}

		s := f.Sheet(sheet)
		if s == nil {
			results = append(results, CellResult{
				Ref:   check.Ref,
				Value: "#SHEET_NOT_FOUND",
				Type:  "error",
			})
			continue
		}

		val, err := s.GetValue(cellRef)
		if err != nil {
			results = append(results, CellResult{
				Ref:   check.Ref,
				Value: fmt.Sprintf("#ERR:%v", err),
				Type:  "error",
			})
			continue
		}

		results = append(results, CellResult{
			Ref:   check.Ref,
			Value: FormatValue(val),
			Type:  ValueTypeName(val),
		})
	}

	return results, nil
}

// ExcludedFunctions returns functions excluded for local Excel.
func (e *LocalExcelEvaluator) ExcludedFunctions() map[string]bool {
	return LocalExcelExcludedFunctions
}
