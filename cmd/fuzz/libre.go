package main

import (
	"encoding/xml"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
)

// findLibreOffice returns the path to libreoffice or an error.
func findLibreOffice() (string, error) {
	path, err := exec.LookPath("libreoffice")
	if err != nil {
		// Try macOS default location.
		macPath := "/Applications/LibreOffice.app/Contents/MacOS/soffice"
		if _, serr := os.Stat(macPath); serr == nil {
			return macPath, nil
		}
		return "", fmt.Errorf("libreoffice not found in PATH: %w", err)
	}
	return path, nil
}

// convertToFODS converts an XLSX file to Flat ODS XML and returns the FODS path.
func convertToFODS(soffice, xlsxPath string) (string, error) {
	dir := filepath.Dir(xlsxPath)
	cmd := exec.Command(soffice, "--headless", "--convert-to", "fods", "--outdir", dir, xlsxPath)
	output, err := cmd.CombinedOutput()
	if err != nil {
		return "", fmt.Errorf("libreoffice convert to fods: %w\noutput: %s", err, output)
	}
	base := strings.TrimSuffix(filepath.Base(xlsxPath), filepath.Ext(xlsxPath))
	fodsPath := filepath.Join(dir, base+".fods")
	if _, err := os.Stat(fodsPath); err != nil {
		return "", fmt.Errorf("fods file not created: %w", err)
	}
	return fodsPath, nil
}

// fodsCell is a minimal struct for parsing FODS cell elements.
type fodsCell struct {
	ValueType string `xml:"value-type,attr"`
	Value     string `xml:"value,attr"`
	TextP     string `xml:"p"`
	ColRepeat int    `xml:"number-columns-repeated,attr"`
}

// fodsRow is a minimal struct for parsing FODS row elements.
type fodsRow struct {
	Cells     []fodsCell `xml:"table-cell"`
	RowRepeat int        `xml:"number-rows-repeated,attr"`
}

// fodsTable is a minimal struct for parsing FODS table elements.
type fodsTable struct {
	Name string    `xml:"name,attr"`
	Rows []fodsRow `xml:"table-row"`
}

// fodsBody is the top-level struct for parsing FODS documents.
type fodsBody struct {
	Spreadsheet struct {
		Tables []fodsTable `xml:"table"`
	} `xml:"body>spreadsheet"`
}

// readFODSValues parses a FODS file and returns cell values for a given sheet.
// Returns a map of cell ref (e.g. "A1") to text value.
func readFODSValues(fodsPath, sheetName string) (map[string]string, error) {
	data, err := os.ReadFile(fodsPath)
	if err != nil {
		return nil, fmt.Errorf("read fods: %w", err)
	}

	// Strip namespaces for easier parsing (same approach as integration_test.go).
	content := string(data)
	content = strings.ReplaceAll(content, "office:", "")
	content = strings.ReplaceAll(content, "table:", "")
	content = strings.ReplaceAll(content, "text:", "")

	var doc fodsBody
	if err := xml.Unmarshal([]byte(content), &doc); err != nil {
		return nil, fmt.Errorf("unmarshal fods: %w", err)
	}

	result := make(map[string]string)
	for _, table := range doc.Spreadsheet.Tables {
		if table.Name != sheetName {
			continue
		}
		for rowIdx, row := range table.Rows {
			colIdx := 0
			for _, cell := range row.Cells {
				repeat := cell.ColRepeat
				if repeat == 0 {
					repeat = 1
				}
				if cell.TextP != "" {
					ref := fmt.Sprintf("%s%d", colToLetter(colIdx+1), rowIdx+1)
					result[ref] = cell.TextP
				}
				colIdx += repeat
			}
		}
		break
	}
	return result, nil
}

// colToLetter converts a 1-based column number to Excel column letters.
func colToLetter(col int) string {
	s := ""
	for col > 0 {
		col--
		s = string(rune('A'+col%26)) + s
		col /= 26
	}
	return s
}

// libreOfficeEval runs LibreOffice to evaluate the XLSX and returns
// values for the requested check refs.
func libreOfficeEval(soffice, xlsxPath string, checks []CheckSpec) ([]buildResult, error) {
	fodsPath, err := convertToFODS(soffice, xlsxPath)
	if err != nil {
		return nil, err
	}

	// Collect all unique sheet names needed.
	sheetValues := make(map[string]map[string]string)

	var results []buildResult
	for _, check := range checks {
		sheet, cellRef := parseCheckRef(check.Ref)
		if sheet == "" {
			sheet = "Sheet1"
		}

		// Parse sheet values on demand.
		if _, ok := sheetValues[sheet]; !ok {
			vals, err := readFODSValues(fodsPath, sheet)
			if err != nil {
				return nil, fmt.Errorf("read FODS sheet %q: %w", sheet, err)
			}
			sheetValues[sheet] = vals
		}

		val, ok := sheetValues[sheet][cellRef]
		if !ok {
			results = append(results, buildResult{
				Ref:   check.Ref,
				Value: "",
				Type:  "empty",
			})
		} else {
			results = append(results, buildResult{
				Ref:   check.Ref,
				Value: val,
				Type:  guessType(val),
			})
		}
	}

	return results, nil
}

// guessType infers a type string from a FODS text value.
func guessType(val string) string {
	if val == "" {
		return "empty"
	}
	upper := strings.ToUpper(val)
	if upper == "TRUE" || upper == "FALSE" {
		return "bool"
	}
	if strings.HasPrefix(val, "#") || strings.HasPrefix(val, "Err:") {
		return "error"
	}
	// Try to see if it looks numeric.
	cleaned := strings.TrimRight(val, "0")
	cleaned = strings.TrimRight(cleaned, ".")
	_ = cleaned
	return "number" // Default assumption for FODS numeric outputs.
}
