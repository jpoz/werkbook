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

	ts := time.Now().UnixNano()
	tmpPath := filepath.Join(tmpDir, fmt.Sprintf("eval_%d.xlsx", ts))
	if err := os.WriteFile(tmpPath, data, 0644); err != nil {
		return nil, fmt.Errorf("write temp xlsx: %w", err)
	}

	absPath, err := filepath.Abs(tmpPath)
	if err != nil {
		return nil, fmt.Errorf("abs path: %w", err)
	}

	// Second path for "save as" (avoids issues with repaired/locked files)
	savePath := filepath.Join(tmpDir, fmt.Sprintf("saved_%d.xlsx", ts))

	// First, ensure Excel is in a clean state: dismiss dialogs, close all workbooks.
	cleanupScript := `
tell application "System Events"
	tell process "Microsoft Excel"
		repeat 5 times
			try
				set dlgWindows to every window whose name is ""
				if (count of dlgWindows) = 0 then exit repeat
				repeat with w in dlgWindows
					try
						click button "No" of w
					end try
					try
						click button "Yes" of w
					end try
				end repeat
			on error
				exit repeat
			end try
			delay 1
		end repeat
	end tell
end tell

tell application "Microsoft Excel"
	try
		repeat while (count of workbooks) > 0
			close active workbook saving no
		end repeat
	end try
end tell
`
	cleanCmd := exec.Command("osascript", "-e", cleanupScript)
	_ = cleanCmd.Run() // Ignore errors if Excel isn't running

	// Open the file via shell command (non-blocking), then use AppleScript to
	// dismiss any dialogs, wait for recalc, save, and close.
	openCmd := exec.Command("open", "-a", "Microsoft Excel", absPath)
	if err := openCmd.Run(); err != nil {
		return nil, fmt.Errorf("open Excel: %w", err)
	}

	// AppleScript: dismiss recovery/repair dialogs, then save and close.
	// Excel may show two dialogs:
	//   1. "We found a problem..." with Yes/No -> click Yes to recover
	//   2. "Excel was able to open..." with View/Delete -> click View to keep file
	// We use "save workbook as" to the original path to handle repaired files
	// (Excel renames them to "filename - Repaired" which breaks plain "save").
	script := fmt.Sprintf(`
-- Dismiss any unnamed dialog windows
tell application "System Events"
	tell process "Microsoft Excel"
		repeat 30 times
			delay 1
			try
				set dlgWindows to every window whose name is ""
				if (count of dlgWindows) = 0 then exit repeat
				set dlg to item 1 of dlgWindows
				set btnNames to name of every button of dlg
				if btnNames contains "Yes" then
					click button "Yes" of dlg
				else if btnNames contains "View" then
					click button "View" of dlg
				else if btnNames contains "OK" then
					click button "OK" of dlg
				else if btnNames contains "Close" then
					click button "Close" of dlg
				else
					exit repeat
				end if
			on error
				exit repeat
			end try
		end repeat
	end tell
end tell

delay 5

set savePath to POSIX file "%s"
tell application "Microsoft Excel"
	with timeout of 300 seconds
		if (count of workbooks) > 0 then
			save active workbook in savePath
			delay 2
			close active workbook saving no
		end if
	end timeout
end tell
`, savePath)

	cmd := exec.Command("osascript", "-e", script)
	output, err := cmd.CombinedOutput()
	if err != nil {
		return nil, fmt.Errorf("AppleScript failed: %w\noutput: %s", err, output)
	}

	// Re-read the saved XLSX from the save-as path.
	f, err := werkbook.Open(savePath)
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
