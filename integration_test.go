//go:build integration

package werkbook_test

import (
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"

	"github.com/jpoz/werkbook"
)

func TestLibreOfficeConvert(t *testing.T) {
	// Check if LibreOffice is available.
	lofficePath, err := exec.LookPath("libreoffice")
	if err != nil {
		t.Skip("LibreOffice not found, skipping integration test")
	}
	_ = lofficePath

	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetValue("A1", "hello")
	s.SetValue("B1", 42)
	s.SetValue("C1", 3.14)
	s.SetValue("A2", "world")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "test.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	// Convert to CSV using LibreOffice headless.
	cmd := exec.Command("libreoffice", "--headless", "--convert-to", "csv", "--outdir", dir, xlsxPath)
	output, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("libreoffice convert: %v\noutput: %s", err, output)
	}

	csvPath := filepath.Join(dir, "test.csv")
	csvData, err := os.ReadFile(csvPath)
	if err != nil {
		t.Fatalf("read csv: %v", err)
	}

	csv := string(csvData)
	lines := strings.Split(strings.TrimSpace(csv), "\n")
	if len(lines) < 2 {
		t.Fatalf("expected at least 2 CSV lines, got %d: %q", len(lines), csv)
	}

	// Verify first row contains our values.
	if !strings.Contains(lines[0], "hello") {
		t.Errorf("line 0 missing 'hello': %q", lines[0])
	}
	if !strings.Contains(lines[0], "42") {
		t.Errorf("line 0 missing '42': %q", lines[0])
	}
	if !strings.Contains(lines[1], "world") {
		t.Errorf("line 1 missing 'world': %q", lines[1])
	}
}
