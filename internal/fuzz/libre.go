package fuzz

import (
	"encoding/xml"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
)

// FindLibreOffice returns the path to libreoffice or an error.
func FindLibreOffice() (string, error) {
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

// ConvertToFODS converts an XLSX file to Flat ODS XML and returns the FODS path.
func ConvertToFODS(soffice, xlsxPath string) (string, error) {
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

// ReadFODSValues parses a FODS file and returns cell values for a given sheet.
// Returns a map of cell ref (e.g. "A1") to text value.
func ReadFODSValues(fodsPath, sheetName string) (map[string]string, error) {
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
					ref := fmt.Sprintf("%s%d", ColToLetter(colIdx+1), rowIdx+1)
					result[ref] = cell.TextP
				}
				colIdx += repeat
			}
		}
		break
	}
	return result, nil
}

// ColToLetter converts a 1-based column number to Excel column letters.
func ColToLetter(col int) string {
	s := ""
	for col > 0 {
		col--
		s = string(rune('A'+col%26)) + s
		col /= 26
	}
	return s
}

// LibreOfficeEval runs LibreOffice to evaluate the XLSX and returns
// values for the requested check refs.
func LibreOfficeEval(soffice, xlsxPath string, checks []CheckSpec) ([]CellResult, error) {
	fodsPath, err := ConvertToFODS(soffice, xlsxPath)
	if err != nil {
		return nil, err
	}

	// Collect all unique sheet names needed.
	sheetValues := make(map[string]map[string]string)

	var results []CellResult
	for _, check := range checks {
		sheet, cellRef := ParseCheckRef(check.Ref)
		if sheet == "" {
			sheet = "Sheet1"
		}

		// Parse sheet values on demand.
		if _, ok := sheetValues[sheet]; !ok {
			vals, err := ReadFODSValues(fodsPath, sheet)
			if err != nil {
				return nil, fmt.Errorf("read FODS sheet %q: %w", sheet, err)
			}
			sheetValues[sheet] = vals
		}

		val, ok := sheetValues[sheet][cellRef]
		if !ok {
			results = append(results, CellResult{
				Ref:   check.Ref,
				Value: "",
				Type:  "empty",
			})
		} else {
			results = append(results, CellResult{
				Ref:   check.Ref,
				Value: val,
				Type:  GuessType(val),
			})
		}
	}

	return results, nil
}

// LibreOfficeEvaluator wraps the LibreOffice evaluation functions as an Evaluator.
type LibreOfficeEvaluator struct {
	SOfficePath string
}

// Name returns "libreoffice".
func (e *LibreOfficeEvaluator) Name() string { return "libreoffice" }

// Eval evaluates the XLSX using LibreOffice and returns cell results.
func (e *LibreOfficeEvaluator) Eval(xlsxPath string, checks []CheckSpec) ([]CellResult, error) {
	return LibreOfficeEval(e.SOfficePath, xlsxPath, checks)
}

// ExcludedFunctions returns functions that LibreOffice cannot evaluate.
func (e *LibreOfficeEvaluator) ExcludedFunctions() map[string]bool {
	return LibreOfficeExcludedFunctions
}

// GuessType infers a type string from a FODS text value.
func GuessType(val string) string {
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
	return "number" // Default assumption for FODS numeric outputs.
}
