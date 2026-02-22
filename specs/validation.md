# Phase 1 Validation Strategy

## Overview

Use independent XLSX implementations as test oracles to validate werkbook's read/write correctness:

1. **LibreOffice headless** (installed at `/Applications/LibreOffice.app`) - open, validate, and round-trip XLSX files from the CLI
2. **Python openpyxl** (optional, pip installable) - independent reader/writer

The idea: werkbook writes a file, the oracles read it and confirm the values match. An oracle writes a file, werkbook reads it and confirms the values match. If they agree, the file is correct.

## LibreOffice Headless

LibreOffice can run without a GUI and do useful things with spreadsheets.

### Convert to CSV (read validation)

```bash
# Convert an XLSX to CSV - if LibreOffice can read it and produce correct CSV,
# the XLSX is structurally valid
/Applications/LibreOffice.app/Contents/MacOS/soffice \
  --headless \
  --convert-to csv:"Text - txt - csv (StarCalc)":44,34,76,1 \
  --outdir /tmp \
  input.xlsx
```

### Macro-based cell extraction

For more precise validation (types, multiple sheets), use a LibreOffice Basic macro:

```bash
# Run a macro that opens the file and dumps cell values as JSON
/Applications/LibreOffice.app/Contents/MacOS/soffice \
  --headless \
  --calc \
  "macro:///Standard.Module1.DumpCells" \
  input.xlsx
```

We'll ship a small macro or Python-UNO script in `testdata/` for this (see below).

### Recalculate and export (calc validation, Phase 2)

```bash
# Open file, recalculate, save-as CSV
# This forces LibreOffice to evaluate all formulas
/Applications/LibreOffice.app/Contents/MacOS/soffice \
  --headless \
  --calc \
  --convert-to csv \
  --outdir /tmp \
  input_with_formulas.xlsx
```

The CSV output will contain calculated values, not formula text. Compare these against werkbook's CalcValue output.

## Test Harness Design

### Directory Layout

```
werkbook/
  testdata/
    fixtures/              # Pre-built XLSX files from various sources
      libreoffice_basic.xlsx
      excel_basic.xlsx     # If anyone has Excel, add samples
    golden/                # Expected outputs (CSV, JSON)
      basic.csv
    scripts/
      dump_xlsx.py         # openpyxl-based cell dumper (optional)
      validate_xlsx.sh     # LibreOffice headless validation script
```

### validate_xlsx.sh

A reusable script that takes an XLSX file and produces a normalized JSON representation using LibreOffice.

```bash
#!/bin/bash
# Usage: validate_xlsx.sh input.xlsx [output.csv]
# Converts XLSX to CSV via LibreOffice headless, exits non-zero if it fails.
set -euo pipefail

SOFFICE="/Applications/LibreOffice.app/Contents/MacOS/soffice"
INPUT="$1"
OUTDIR="${2:-/tmp/werkbook-validate}"
mkdir -p "$OUTDIR"

# Kill any existing soffice process that might block headless mode
# (LibreOffice only allows one instance per user profile)
"$SOFFICE" \
  --headless \
  --convert-to csv:"Text - txt - csv (StarCalc)":44,34,76,1 \
  --outdir "$OUTDIR" \
  "$INPUT"

echo "$OUTDIR/$(basename "${INPUT%.xlsx}.csv")"
```

### Go Integration Test Pattern

```go
// werkbook/integration_test.go
//go:build integration

package werkbook_test

import (
    "encoding/csv"
    "os"
    "os/exec"
    "path/filepath"
    "testing"

    "github.com/jpoz/werkbook"
)

// requireLibreOffice skips if LibreOffice isn't available
func requireLibreOffice(t *testing.T) string {
    t.Helper()
    paths := []string{
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/libreoffice",
    }
    for _, p := range paths {
        if _, err := os.Stat(p); err == nil {
            return p
        }
    }
    t.Skip("LibreOffice not found, skipping integration test")
    return ""
}

// libreOfficeToCSV converts an XLSX to CSV using LibreOffice headless
func libreOfficeToCSV(t *testing.T, soffice, xlsxPath string) string {
    t.Helper()
    outDir := t.TempDir()
    cmd := exec.Command(soffice,
        "--headless",
        "--convert-to", "csv:Text - txt - csv (StarCalc):44,34,76,1",
        "--outdir", outDir,
        xlsxPath,
    )
    out, err := cmd.CombinedOutput()
    if err != nil {
        t.Fatalf("LibreOffice conversion failed: %v\n%s", err, out)
    }
    base := filepath.Base(xlsxPath)
    csvPath := filepath.Join(outDir, base[:len(base)-5]+".csv")
    return csvPath
}

// readCSV parses a CSV into [][]string
func readCSV(t *testing.T, path string) [][]string {
    t.Helper()
    f, err := os.Open(path)
    if err != nil {
        t.Fatal(err)
    }
    defer f.Close()
    records, err := csv.NewReader(f).ReadAll()
    if err != nil {
        t.Fatal(err)
    }
    return records
}

// TestWriteThenValidateWithLibreOffice writes an XLSX with werkbook,
// opens it with LibreOffice, and confirms the cell values match.
func TestWriteThenValidateWithLibreOffice(t *testing.T) {
    soffice := requireLibreOffice(t)

    // 1. Write with werkbook
    f := werkbook.New()
    sheet := f.Sheet("Sheet1")
    sheet.SetValue("A1", "Name")
    sheet.SetValue("B1", "Age")
    sheet.SetValue("A2", "Alice")
    sheet.SetValue("B2", 30)
    sheet.SetValue("A3", "Bob")
    sheet.SetValue("B3", 25)

    path := filepath.Join(t.TempDir(), "test.xlsx")
    if err := f.SaveAs(path); err != nil {
        t.Fatal(err)
    }

    // 2. Convert to CSV with LibreOffice
    csvPath := libreOfficeToCSV(t, soffice, path)
    records := readCSV(t, csvPath)

    // 3. Assert values match
    expected := [][]string{
        {"Name", "Age"},
        {"Alice", "30"},
        {"Bob", "25"},
    }
    for i, row := range expected {
        for j, want := range row {
            got := records[i][j]
            if got != want {
                t.Errorf("cell [%d][%d]: got %q, want %q", i, j, got, want)
            }
        }
    }
}
```

### Test Categories

#### 1. Write Validation (werkbook writes, oracle reads)

werkbook produces an XLSX, LibreOffice reads it.

| Test Case | What it validates |
|-----------|-------------------|
| Empty workbook | Minimal valid XLSX structure |
| String cells | Shared strings table, cell type "s" |
| Number cells | Numeric values, cell type "n" |
| Bool cells | Boolean encoding |
| Date cells | Date serial number encoding, numfmt |
| Mixed types in one sheet | Type discrimination |
| Multiple sheets | Workbook sheet list, relationship IDs |
| Large sheet (100k rows) | Streaming writer, ZIP64 if needed |
| Column widths | Column definitions |
| Row heights | Row attributes |
| Merged cells | MergeCells element |
| Cell styles (bold, color, numfmt) | Style index references, style XML |
| Unicode / CJK text | UTF-8 encoding, shared strings |
| Empty rows/cols (sparse data) | Sparse representation |
| Sheet with formulas (text preserved) | Formula element in cell XML |

#### 2. Read Validation (oracle writes, werkbook reads)

LibreOffice produces an XLSX, werkbook reads it.

| Test Case | What it validates |
|-----------|-------------------|
| Basic LibreOffice file | Tolerance for LO-specific XML quirks |
| File with named ranges | Defined name parsing |
| File with auto-filter | AutoFilter element |
| File with merged cells | MergeCell parsing |
| File with comments | Comment XML reading |
| File with multiple sheets | Multi-sheet navigation |
| File with styles | Style index resolution |
| Older Excel format quirks | Compatibility attributes |

#### 3. Round-Trip Validation (read -> modify -> write -> read)

Open a file, change some cells, save, reopen with oracle.

| Test Case | What it validates |
|-----------|-------------------|
| Change cell value, re-read | Value update without corruption |
| Add new sheet, re-read | Sheet insertion |
| Delete sheet, re-read | Sheet removal, relationship cleanup |
| Add rows to existing sheet | Row insertion |
| Modify styled cell | Style preservation |

## Running Tests

```bash
# Unit tests (no external deps)
go test ./...

# Integration tests (requires LibreOffice)
go test -tags=integration ./...
```

## CI Considerations

- LibreOffice can be installed in CI (GitHub Actions: `apt-get install libreoffice-calc`)
- Integration tests gated behind `//go:build integration` so `go test ./...` works without LibreOffice
- Fixtures are checked into the repo (they're small XLSX files)
- For CI on macOS, use `brew install --cask libreoffice` or the pre-installed path

```yaml
# .github/workflows/test.yml snippet
- name: Install LibreOffice
  run: sudo apt-get install -y libreoffice-calc

- name: Run integration tests
  run: go test -tags=integration -v ./...
```

## Multi-Sheet CSV Extraction

LibreOffice's `--convert-to csv` only exports the active sheet by default. For multi-sheet validation, use a Python-UNO script or export each sheet individually:

```bash
# Export specific sheet by index using a macro-like filter
/Applications/LibreOffice.app/Contents/MacOS/soffice \
  --headless \
  --infilter="Microsoft Excel 2007-2019 XML (.xlsx)" \
  --convert-to "csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,Sheet2Index" \
  --outdir /tmp \
  input.xlsx
```

Alternatively, convert to FODS (Flat ODS XML) which includes all sheets in a single human-readable XML file:

```bash
/Applications/LibreOffice.app/Contents/MacOS/soffice \
  --headless \
  --convert-to fods \
  --outdir /tmp \
  input.xlsx
```

FODS is XML, easy to parse in Go, and contains every sheet's data, formulas, and styles in one file. This is likely the best format for thorough validation.
