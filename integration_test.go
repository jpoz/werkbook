//go:build integration

package werkbook_test

import (
	"encoding/csv"
	"encoding/xml"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"

	"github.com/jpoz/werkbook"
)

// requireLibreOffice returns the soffice path or skips the test.
func requireLibreOffice(t *testing.T) string {
	t.Helper()
	path, err := exec.LookPath("libreoffice")
	if err != nil {
		t.Skip("LibreOffice not found, skipping integration test")
	}
	return path
}

// libreOfficeToCSV converts an XLSX file to CSV and returns the CSV path.
func libreOfficeToCSV(t *testing.T, soffice, xlsxPath string) string {
	t.Helper()
	dir := filepath.Dir(xlsxPath)
	cmd := exec.Command(soffice, "--headless", "--convert-to", "csv", "--outdir", dir, xlsxPath)
	output, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("libreoffice convert to csv: %v\noutput: %s", err, output)
	}
	base := strings.TrimSuffix(filepath.Base(xlsxPath), filepath.Ext(xlsxPath))
	return filepath.Join(dir, base+".csv")
}

// libreOfficeToFODS converts an XLSX file to Flat ODS XML and returns the FODS path.
func libreOfficeToFODS(t *testing.T, soffice, xlsxPath string) string {
	t.Helper()
	dir := filepath.Dir(xlsxPath)
	cmd := exec.Command(soffice, "--headless", "--convert-to", "fods", "--outdir", dir, xlsxPath)
	output, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("libreoffice convert to fods: %v\noutput: %s", err, output)
	}
	base := strings.TrimSuffix(filepath.Base(xlsxPath), filepath.Ext(xlsxPath))
	return filepath.Join(dir, base+".fods")
}

// readCSV reads a CSV file and returns all records.
func readCSV(t *testing.T, path string) [][]string {
	t.Helper()
	f, err := os.Open(path)
	if err != nil {
		t.Fatalf("open csv: %v", err)
	}
	defer f.Close()
	r := csv.NewReader(f)
	records, err := r.ReadAll()
	if err != nil {
		t.Fatalf("read csv: %v", err)
	}
	return records
}

// readFODSCellValues parses a Flat ODS file and returns a map of cell ref -> text value
// for the given sheet name. Cell refs are in A1 style derived from row/col position.
func readFODSCellValues(t *testing.T, fodsPath string, sheetName string) map[string]string {
	t.Helper()
	data, err := os.ReadFile(fodsPath)
	if err != nil {
		t.Fatalf("read fods: %v", err)
	}

	// FODS uses ODF namespaces. We parse with a minimal struct.
	type fodsCell struct {
		ValueType string `xml:"value-type,attr"`
		Value     string `xml:"value,attr"`
		TextP     string `xml:"p"`
		ColRepeat int    `xml:"number-columns-repeated,attr"`
	}
	type fodsRow struct {
		Cells     []fodsCell `xml:"table-cell"`
		RowRepeat int        `xml:"number-rows-repeated,attr"`
	}
	type fodsTable struct {
		Name string    `xml:"name,attr"`
		Rows []fodsRow `xml:"table-row"`
	}
	type fodsBody struct {
		Spreadsheet struct {
			Tables []fodsTable `xml:"table"`
		} `xml:"body>spreadsheet"`
	}

	// Strip namespaces for easier parsing.
	content := string(data)
	content = strings.ReplaceAll(content, "office:", "")
	content = strings.ReplaceAll(content, "table:", "")
	content = strings.ReplaceAll(content, "text:", "")

	var doc fodsBody
	if err := xml.Unmarshal([]byte(content), &doc); err != nil {
		t.Fatalf("unmarshal fods: %v", err)
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
	return result
}

func colToLetter(col int) string {
	s := ""
	for col > 0 {
		col--
		s = string(rune('A'+col%26)) + s
		col /= 26
	}
	return s
}

func TestLibreOfficeConvert(t *testing.T) {
	soffice := requireLibreOffice(t)

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

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	if len(records) < 2 {
		t.Fatalf("expected at least 2 rows, got %d", len(records))
	}
	if !strings.Contains(records[0][0], "hello") {
		t.Errorf("row 0 col 0 missing 'hello': %q", records[0][0])
	}
	if !strings.Contains(records[0][1], "42") {
		t.Errorf("row 0 col 1 missing '42': %q", records[0][1])
	}
	if !strings.Contains(records[1][0], "world") {
		t.Errorf("row 1 col 0 missing 'world': %q", records[1][0])
	}
}

func TestFormulaEvalWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set input values.
	s.SetValue("A1", 10)
	s.SetValue("A2", 20)
	s.SetValue("A3", 30)

	// Set formulas in B column.
	s.SetFormula("B1", "SUM(A1:A3)")
	s.SetFormula("B2", "A1*A2")
	s.SetFormula("B3", `IF(A1>5,"yes","no")`)
	s.SetFormula("B4", "AVERAGE(A1:A3)")
	s.SetFormula("B5", `A1&" items"`)

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "formulas.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := map[int]string{
		0: "60",
		1: "200",
		2: "yes",
		3: "20",
		4: "10 items",
	}

	for row, want := range expected {
		if row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", row, len(records))
			continue
		}
		if len(records[row]) < 2 {
			t.Errorf("row %d: expected at least 2 columns, got %d", row, len(records[row]))
			continue
		}
		got := records[row][1]
		if got != want {
			t.Errorf("B%d = %q, want %q", row+1, got, want)
		}
	}
}

func TestFormulaRoundTripWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetValue("A1", 10)
	s.SetValue("A2", 20)
	s.SetFormula("B1", "SUM(A1:A2)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "roundtrip.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	// Verify formulas preserved when re-read by werkbook.
	f2, err := werkbook.Open(xlsxPath)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	s2 := f2.Sheet("Sheet1")
	formula, _ := s2.GetFormula("B1")
	if formula != "SUM(A1:A2)" {
		t.Errorf("formula after re-read = %q, want %q", formula, "SUM(A1:A2)")
	}

	// Verify LibreOffice evaluates it correctly.
	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)
	if len(records) < 1 || len(records[0]) < 2 {
		t.Fatalf("unexpected CSV shape: %v", records)
	}
	if records[0][1] != "30" {
		t.Errorf("LibreOffice B1 = %q, want %q", records[0][1], "30")
	}
}

func TestMultiSheetFormulaWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s1 := f.Sheet("Sheet1")
	s1.SetValue("A1", 100)

	s2, err := f.NewSheet("Sheet2")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	s2.SetFormula("A1", "Sheet1!A1*2")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "multisheet.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	fodsPath := libreOfficeToFODS(t, soffice, xlsxPath)
	cells := readFODSCellValues(t, fodsPath, "Sheet2")

	if val, ok := cells["A1"]; !ok {
		t.Error("Sheet2 A1 not found in FODS output")
	} else if val != "200" {
		t.Errorf("Sheet2 A1 = %q, want %q", val, "200")
	}
}

func TestASINHWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set input values.
	s.SetValue("A1", 0)
	s.SetValue("A2", 1)
	s.SetValue("A3", -1)

	// Set ASINH formulas.
	s.SetFormula("B1", "ASINH(A1)")
	s.SetFormula("B2", "ASINH(A2)")
	s.SetFormula("B3", "ASINH(A3)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "asinh.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := []struct {
		row  int
		want string
	}{
		{0, "0"},
		{1, "0.881373587019543"},
		{2, "-0.881373587019543"},
	}

	for _, tt := range expected {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 2 {
			t.Errorf("row %d: expected at least 2 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][1]
		if got != tt.want {
			t.Errorf("B%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestATANHWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set input values.
	s.SetValue("A1", 0)
	s.SetValue("A2", 0.5)
	s.SetValue("A3", -0.5)

	// Set ATANH formulas.
	s.SetFormula("B1", "ATANH(A1)")
	s.SetFormula("B2", "ATANH(A2)")
	s.SetFormula("B3", "ATANH(A3)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "atanh.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := []struct {
		row  int
		want string
	}{
		{0, "0"},
		{1, "0.549306144334055"},
		{2, "-0.549306144334055"},
	}

	for _, tt := range expected {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 2 {
			t.Errorf("row %d: expected at least 2 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][1]
		if got != tt.want {
			t.Errorf("B%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestCOSHWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set input values.
	s.SetValue("A1", 0)
	s.SetValue("A2", 1)
	s.SetValue("A3", -1)

	// Set COSH formulas.
	s.SetFormula("B1", "COSH(A1)")
	s.SetFormula("B2", "COSH(A2)")
	s.SetFormula("B3", "COSH(A3)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "cosh.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := []struct {
		row  int
		want string
	}{
		{0, "1"},
		{1, "1.54308063481524"},
		{2, "1.54308063481524"},
	}

	for _, tt := range expected {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 2 {
			t.Errorf("row %d: expected at least 2 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][1]
		if got != tt.want {
			t.Errorf("B%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestSINHWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set input values.
	s.SetValue("A1", 0)
	s.SetValue("A2", 1)
	s.SetValue("A3", -1)

	// Set SINH formulas.
	s.SetFormula("B1", "SINH(A1)")
	s.SetFormula("B2", "SINH(A2)")
	s.SetFormula("B3", "SINH(A3)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "sinh.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := []struct {
		row  int
		want string
	}{
		{0, "0"},
		{1, "1.1752011936438"},
		{2, "-1.1752011936438"},
	}

	for _, tt := range expected {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 2 {
			t.Errorf("row %d: expected at least 2 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][1]
		if got != tt.want {
			t.Errorf("B%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestTANHWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set input values.
	s.SetValue("A1", 0)
	s.SetValue("A2", 1)
	s.SetValue("A3", -1)

	// Set TANH formulas.
	s.SetFormula("B1", "TANH(A1)")
	s.SetFormula("B2", "TANH(A2)")
	s.SetFormula("B3", "TANH(A3)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "tanh.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := []struct {
		row  int
		want string
	}{
		{0, "0"},
		{1, "0.761594155955764"},
		{2, "-0.761594155955764"},
	}

	for _, tt := range expected {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 2 {
			t.Errorf("row %d: expected at least 2 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][1]
		if got != tt.want {
			t.Errorf("B%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestSQRTPIWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set input values.
	s.SetValue("A1", 1)
	s.SetValue("A2", 2)
	s.SetValue("A3", 0)

	// Set SQRTPI formulas.
	s.SetFormula("B1", "SQRTPI(A1)")
	s.SetFormula("B2", "SQRTPI(A2)")
	s.SetFormula("B3", "SQRTPI(A3)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "sqrtpi.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := []struct {
		row  int
		want string
	}{
		{0, "1.7724538509055"},
		{1, "2.50662827463101"},
		{2, "0"},
	}

	for _, tt := range expected {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 2 {
			t.Errorf("row %d: expected at least 2 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][1]
		if got != tt.want {
			t.Errorf("B%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestFACTDOUBLEWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set input values.
	s.SetValue("A1", 6)
	s.SetValue("A2", 7)
	s.SetValue("A3", 0)

	// Set FACTDOUBLE formulas.
	s.SetFormula("B1", "FACTDOUBLE(A1)")
	s.SetFormula("B2", "FACTDOUBLE(A2)")
	s.SetFormula("B3", "FACTDOUBLE(A3)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "factdouble.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := []struct {
		row  int
		want string
	}{
		{0, "48"},
		{1, "105"},
		{2, "1"},
	}

	for _, tt := range expected {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 2 {
			t.Errorf("row %d: expected at least 2 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][1]
		if got != tt.want {
			t.Errorf("B%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestACOSHWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set input values.
	s.SetValue("A1", 1)
	s.SetValue("A2", 2)
	s.SetValue("A3", 10)

	// Set ACOSH formulas.
	s.SetFormula("B1", "ACOSH(A1)")
	s.SetFormula("B2", "ACOSH(A2)")
	s.SetFormula("B3", "ACOSH(A3)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "acosh.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := []struct {
		row  int
		want string
	}{
		{0, "0"},
		{1, "1.31695789692482"},
		{2, "2.99322284612638"},
	}

	for _, tt := range expected {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 2 {
			t.Errorf("row %d: expected at least 2 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][1]
		if got != tt.want {
			t.Errorf("B%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestDECIMALWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetValue("A1", "FF")
	s.SetValue("B1", 16)
	s.SetValue("A2", "111")
	s.SetValue("B2", 2)
	s.SetFormula("C1", `DECIMAL(A1,B1)`)
	s.SetFormula("C2", `DECIMAL(A2,B2)`)

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "decimal.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := []struct {
		row  int
		want string
	}{
		{0, "255"},
		{1, "7"},
	}

	for _, tt := range expected {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 3 {
			t.Errorf("row %d: expected at least 3 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][2]
		if got != tt.want {
			t.Errorf("C%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestMULTINOMIALWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")

	s.SetValue("A1", 2)
	s.SetValue("B1", 3)
	s.SetValue("C1", 4)
	s.SetFormula("D1", "MULTINOMIAL(A1,B1,C1)")

	s.SetValue("A2", 3)
	s.SetValue("B2", 3)
	s.SetFormula("D2", "MULTINOMIAL(A2,B2)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "multinomial.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := []struct {
		row  int
		want string
	}{
		{0, "1260"},
		{1, "20"},
	}

	for _, tt := range expected {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 4 {
			t.Errorf("row %d: expected at least 4 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][3]
		if got != tt.want {
			t.Errorf("D%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestARABICWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetValue("A1", "MCMXCIX")
	s.SetValue("A2", "XLII")
	s.SetFormula("B1", "ARABIC(A1)")
	s.SetFormula("B2", "ARABIC(A2)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "arabic.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := []struct {
		row  int
		want string
	}{
		{0, "1999"},
		{1, "42"},
	}

	for _, tt := range expected {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 2 {
			t.Errorf("row %d: expected at least 2 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][1]
		if got != tt.want {
			t.Errorf("B%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestBASEWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetValue("A1", 255)
	s.SetValue("B1", 16)
	s.SetValue("A2", 7)
	s.SetValue("B2", 2)
	s.SetFormula("C1", "BASE(A1,B1)")
	s.SetFormula("C2", "BASE(A2,B2)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "base.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := []struct {
		row  int
		want string
	}{
		{0, "FF"},
		{1, "111"},
	}

	for _, tt := range expected {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 3 {
			t.Errorf("row %d: expected at least 3 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][2]
		if got != tt.want {
			t.Errorf("C%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestCOMBINAWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set input values.
	s.SetValue("A1", 4)
	s.SetValue("B1", 3)
	s.SetValue("A2", 10)
	s.SetValue("B2", 3)

	// Set COMBINA formulas.
	s.SetFormula("C1", "COMBINA(A1,B1)")
	s.SetFormula("C2", "COMBINA(A2,B2)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "combina.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected2 := []struct {
		row  int
		want string
	}{
		{0, "20"},
		{1, "220"},
	}

	for _, tt := range expected2 {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 3 {
			t.Errorf("row %d: expected at least 3 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][2]
		if got != tt.want {
			t.Errorf("C%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestTrigReciprocalFunctionsWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")

	// Set input value.
	s.SetValue("A1", 1)

	// Set reciprocal trig formulas.
	s.SetFormula("B1", "COT(A1)")
	s.SetFormula("B2", "COTH(A1)")
	s.SetFormula("B3", "CSC(A1)")
	s.SetFormula("B4", "CSCH(A1)")
	s.SetFormula("B5", "SEC(A1)")
	s.SetFormula("B6", "SECH(A1)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "trig_reciprocal.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := []struct {
		row  int
		want string
	}{
		{0, "0.642092615934331"},  // COT(1)
		{1, "1.31303528009535"},   // COTH(1)
		{2, "1.18839510577812"},   // CSC(1)
		{3, "0.850918128239322"},  // CSCH(1)
		{4, "1.85081571768093"},   // SEC(1)
		{5, "0.648054273663885"},  // SECH(1)
	}

	for _, tt := range expected {
		if tt.row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", tt.row, len(records))
			continue
		}
		if len(records[tt.row]) < 2 {
			t.Errorf("row %d: expected at least 2 columns, got %d", tt.row, len(records[tt.row]))
			continue
		}
		got := records[tt.row][1]
		if got != tt.want {
			t.Errorf("B%d = %q, want %q", tt.row+1, got, tt.want)
		}
	}
}

func TestAVEDEVWithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetValue("A1", 1)
	s.SetValue("A2", 2)
	s.SetValue("A3", 3)
	s.SetValue("A4", 4)
	s.SetValue("A5", 5)
	s.SetFormula("B1", "AVEDEV(A1:A5)")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "avedev.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	if len(records) < 1 || len(records[0]) < 2 {
		t.Fatalf("unexpected CSV shape: %v", records)
	}
	got := records[0][1]
	if got != "1.2" {
		t.Errorf("B1 = %q, want %q", got, "1.2")
	}
}

func TestDAYS360WithLibreOffice(t *testing.T) {
	soffice := requireLibreOffice(t)

	f := werkbook.New()
	s := f.Sheet("Sheet1")
	s.SetFormula("A1", "DAYS360(DATE(2024,1,1),DATE(2024,12,31))")
	s.SetFormula("A2", "DAYS360(DATE(2024,1,1),DATE(2024,7,1))")

	dir := t.TempDir()
	xlsxPath := filepath.Join(dir, "days360.xlsx")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	csvPath := libreOfficeToCSV(t, soffice, xlsxPath)
	records := readCSV(t, csvPath)

	expected := map[int]string{
		0: "360",
		1: "180",
	}

	for row, want := range expected {
		if row >= len(records) {
			t.Errorf("row %d: missing (only %d rows)", row, len(records))
			continue
		}
		if len(records[row]) < 1 {
			t.Errorf("row %d: expected at least 1 column, got %d", row, len(records[row]))
			continue
		}
		got := records[row][0]
		if got != want {
			t.Errorf("A%d = %q, want %q", row+1, got, want)
		}
	}
}
