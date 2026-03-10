package main

import (
	"encoding/json"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	werkbook "github.com/jpoz/werkbook"
)

// captureRun calls run() and captures stdout/stderr via a temp redirect.
// Since run() uses fmt.Println/Fprintln directly, we redirect os.Stdout/os.Stderr.
func captureRun(args []string) (stdout string, stderr string, exitCode int) {
	oldStdout := os.Stdout
	oldStderr := os.Stderr
	defer func() {
		os.Stdout = oldStdout
		os.Stderr = oldStderr
	}()

	rOut, wOut, _ := os.Pipe()
	rErr, wErr, _ := os.Pipe()
	os.Stdout = wOut
	os.Stderr = wErr

	exitCode = run(args)

	wOut.Close()
	wErr.Close()

	outBytes, _ := io.ReadAll(rOut)
	stdout = string(outBytes)

	errBytes, _ := io.ReadAll(rErr)
	stderr = string(errBytes)

	return stdout, stderr, exitCode
}

func captureRunJSON(args ...string) (stdout string, stderr string, exitCode int) {
	prefixed := append([]string{"--format", "json"}, args...)
	return captureRun(prefixed)
}

func parseResponse(t *testing.T, s string) Response {
	t.Helper()
	var resp Response
	if err := json.Unmarshal([]byte(s), &resp); err != nil {
		t.Fatalf("failed to parse response JSON: %v\nraw: %s", err, s)
	}
	return resp
}

func createTestFile(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "test.xlsx")
	f := werkbook.New(werkbook.FirstSheet("Sheet1"))
	s := f.Sheet("Sheet1")
	s.SetValue("A1", "Name")
	s.SetValue("B1", "Value")
	s.SetValue("A2", "Alpha")
	s.SetValue("B2", 10.0)
	s.SetValue("A3", "Beta")
	s.SetValue("B3", 20.0)
	s.SetFormula("B4", "SUM(B2:B3)")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("failed to create test file: %v", err)
	}
	return path
}

func createSparseColumnsTestFile(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "sparse.xlsx")
	f := werkbook.New(werkbook.FirstSheet("Data"))
	s := f.Sheet("Data")
	s.SetValue("A1", "Values")
	s.SetValue("C1", "Returns")
	s.SetValue("E1", "Large")
	s.SetValue("A2", 2.0)
	s.SetValue("C2", 1.05)
	s.SetValue("E2", 1000000.0)
	s.SetValue("A3", 4.0)
	s.SetValue("C3", 0.98)
	s.SetValue("E3", 2000000.0)
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("failed to create sparse-column test file: %v", err)
	}
	return path
}

// --- Version ---

func TestVersion(t *testing.T) {
	stdout, _, code := captureRunJSON("version")
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	if resp.Command != "version" {
		t.Fatalf("expected command=version, got %s", resp.Command)
	}
}

func TestVersionTextDefault(t *testing.T) {
	stdout, stderr, code := captureRun([]string{"version"})
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if stderr != "" {
		t.Fatalf("expected no stderr, got %q", stderr)
	}
	if strings.TrimSpace(stdout) != version {
		t.Fatalf("expected plain version output %q, got %q", version, strings.TrimSpace(stdout))
	}
}

// --- Info ---

func TestInfo(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("info", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	data, _ := json.Marshal(resp.Data)
	var info infoData
	json.Unmarshal(data, &info)
	if len(info.Sheets) != 1 {
		t.Fatalf("expected 1 sheet, got %d", len(info.Sheets))
	}
	si := info.Sheets[0]
	if si.Name != "Sheet1" {
		t.Errorf("expected sheet name Sheet1, got %s", si.Name)
	}
	if si.MaxRow != 4 {
		t.Errorf("expected max_row=4, got %d", si.MaxRow)
	}
	if si.MaxCol != 2 {
		t.Errorf("expected max_col=2, got %d", si.MaxCol)
	}
	if si.MaxColLetter != "B" {
		t.Errorf("expected max_col_letter=B, got %s", si.MaxColLetter)
	}
	if !si.HasFormulas {
		t.Error("expected has_formulas=true")
	}
}

func TestInfoSheetFilter(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("info", "--sheet", "Sheet1", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var info infoData
	json.Unmarshal(data, &info)
	if len(info.Sheets) != 1 {
		t.Fatalf("expected 1 sheet, got %d", len(info.Sheets))
	}
}

func TestInfoSheetNotFound(t *testing.T) {
	path := createTestFile(t)
	_, stderr, code := captureRunJSON("info", "--sheet", "Nope", path)
	if code != ExitValidate {
		t.Fatalf("expected exit %d, got %d", ExitValidate, code)
	}
	resp := parseResponse(t, stderr)
	if resp.OK {
		t.Fatal("expected ok=false")
	}
	if resp.Error.Code != ErrCodeSheetNotFound {
		t.Errorf("expected error code %s, got %s", ErrCodeSheetNotFound, resp.Error.Code)
	}
}

func TestInfoFileNotFound(t *testing.T) {
	_, stderr, code := captureRunJSON("info", "/tmp/no_such_file_xyz.xlsx")
	if code != ExitFileIO {
		t.Fatalf("expected exit %d, got %d", ExitFileIO, code)
	}
	resp := parseResponse(t, stderr)
	if resp.Error.Code != ErrCodeFileNotFound {
		t.Errorf("expected error code %s, got %s", ErrCodeFileNotFound, resp.Error.Code)
	}
}

// --- Read ---

func TestRead(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("read", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	data, _ := json.Marshal(resp.Data)
	var rd readData
	json.Unmarshal(data, &rd)
	if rd.Sheet != "Sheet1" {
		t.Errorf("expected sheet=Sheet1, got %s", rd.Sheet)
	}
	if len(rd.Rows) == 0 {
		t.Fatal("expected rows")
	}
}

func TestReadRange(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("read", "--range", "A1:A3", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var rd readData
	json.Unmarshal(data, &rd)
	if rd.Range != "A1:A3" {
		t.Errorf("expected range=A1:A3, got %s", rd.Range)
	}
	// Should have rows 1,2,3 with only column A.
	for _, row := range rd.Rows {
		for cellRef := range row.Cells {
			if !strings.HasPrefix(cellRef, "A") {
				t.Errorf("expected only column A cells, got %s", cellRef)
			}
		}
	}
}

func TestReadHeaders(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("read", "--headers", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var rd readData
	json.Unmarshal(data, &rd)
	if len(rd.Headers) != 2 {
		t.Fatalf("expected 2 headers, got %d", len(rd.Headers))
	}
	if rd.Headers[0] != "Name" || rd.Headers[1] != "Value" {
		t.Errorf("unexpected headers: %v", rd.Headers)
	}
	// Data rows should exclude header row.
	for _, row := range rd.Rows {
		if row.Row == 1 {
			t.Error("header row should not appear in data rows")
		}
	}
}

func TestReadIncludeFormulas(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("read", "--include-formulas", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var rd readData
	json.Unmarshal(data, &rd)
	// B4 should have a formula.
	found := false
	for _, row := range rd.Rows {
		if row.Row == 4 {
			if cd, ok := row.Cells["B4"]; ok {
				if cd.Formula == "SUM(B2:B3)" {
					found = true
				}
			}
		}
	}
	if !found {
		t.Error("expected B4 to have formula SUM(B2:B3)")
	}
}

func TestReadShowFormulasJSON(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("read", "--show-formulas", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var rd readData
	json.Unmarshal(data, &rd)

	found := false
	for _, row := range rd.Rows {
		if row.Row != 4 {
			continue
		}
		cd, ok := row.Cells["B4"]
		if !ok {
			continue
		}
		found = true
		if cd.Display != "=SUM(B2:B3)" {
			t.Fatalf("expected display formula, got %q", cd.Display)
		}
		if cd.Formula != "" {
			t.Fatalf("did not expect formula field without --include-formulas, got %q", cd.Formula)
		}
	}
	if !found {
		t.Fatal("expected B4 in output")
	}
}

func TestReadMarkdown(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRun([]string{"read", "--format", "markdown", "--headers", path})
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stdout, "| Name | Value |") {
		t.Errorf("expected markdown header, got:\n%s", stdout)
	}
	if !strings.Contains(stdout, "|---|---|") {
		t.Errorf("expected markdown separator, got:\n%s", stdout)
	}
}

func TestReadTextDefault(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRun([]string{"read", "--headers", path})
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stdout, "Sheet: Sheet1") {
		t.Errorf("expected sheet heading, got:\n%s", stdout)
	}
	lines := strings.Split(strings.TrimSpace(stdout), "\n")
	if len(lines) < 5 {
		t.Fatalf("expected heading plus table, got:\n%s", stdout)
	}
	if lines[2] != "| Name  | Value |" {
		t.Errorf("expected markdown header row, got %q", lines[2])
	}
	if lines[3] != "| ----- | ----- |" {
		t.Errorf("expected markdown divider row, got %q", lines[3])
	}
}

func TestReadShowFormulasText(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRun([]string{"read", "--show-formulas", "--range", "B4:B4", path})
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stdout, "=SUM(B2:B3)") {
		t.Fatalf("expected formula text in rendered output, got:\n%s", stdout)
	}
}

func TestReadCSV(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRun([]string{"read", "--format", "csv", "--headers", path})
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	lines := strings.Split(strings.TrimSpace(stdout), "\n")
	if len(lines) < 2 {
		t.Fatalf("expected at least 2 lines, got %d", len(lines))
	}
	if lines[0] != "Name,Value" {
		t.Errorf("expected CSV header 'Name,Value', got %q", lines[0])
	}
}

// --- Edit ---

func TestEdit(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("edit", path, "--patch", `[{"cell":"A5","value":"Gamma"},{"cell":"B5","value":30}]`)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	data, _ := json.Marshal(resp.Data)
	var ed editData
	json.Unmarshal(data, &ed)
	if ed.Applied != 2 {
		t.Errorf("expected 2 applied, got %d", ed.Applied)
	}

	// Verify by reading back.
	stdout2, _, code2 := captureRunJSON("read", "--range", "A5:B5", path)
	if code2 != 0 {
		t.Fatalf("expected exit 0 on read, got %d", code2)
	}
	if !strings.Contains(stdout2, "Gamma") {
		t.Error("expected to find Gamma in read output")
	}
}

func TestEditDryRun(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("edit", "--dry-run", path, "--patch", `[{"cell":"A1","value":"Changed"}]`)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}

	// Verify file was NOT changed.
	stdout2, _, _ := captureRunJSON("read", "--range", "A1:A1", path)
	if strings.Contains(stdout2, "Changed") {
		t.Error("dry-run should not modify file")
	}
}

func TestEditFormula(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("edit", path, "--patch", `[{"cell":"C1","formula":"SUM(B2:B3)"}]`)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var ed editData
	json.Unmarshal(data, &ed)
	if ed.Operations[0].Action != "set_formula" {
		t.Errorf("expected action=set_formula, got %s", ed.Operations[0].Action)
	}
}

func TestEditAddDeleteSheet(t *testing.T) {
	path := createTestFile(t)
	// Add a sheet.
	stdout, _, code := captureRunJSON("edit", path, "--patch", `[{"add_sheet":"NewSheet"}]`)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}

	// Verify it exists.
	stdout2, _, _ := captureRunJSON("info", path)
	if !strings.Contains(stdout2, "NewSheet") {
		t.Error("expected NewSheet in info output")
	}

	// Delete it.
	stdout3, _, code3 := captureRunJSON("edit", path, "--patch", `[{"delete_sheet":"NewSheet"}]`)
	if code3 != 0 {
		t.Fatalf("expected exit 0, got %d", code3)
	}
	resp3 := parseResponse(t, stdout3)
	if !resp3.OK {
		t.Fatal("expected ok=true")
	}
}

func TestEditInvalidPatch(t *testing.T) {
	path := createTestFile(t)
	_, stderr, code := captureRunJSON("edit", path, "--patch", `not json`)
	if code != ExitValidate {
		t.Fatalf("expected exit %d, got %d", ExitValidate, code)
	}
	resp := parseResponse(t, stderr)
	if resp.Error.Code != ErrCodeInvalidPatch {
		t.Errorf("expected error code %s, got %s", ErrCodeInvalidPatch, resp.Error.Code)
	}
}

// --- Create ---

func TestCreate(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "new.xlsx")
	stdout, _, code := captureRunJSON("create", path, "--spec", `{"sheets":["Data","Summary"],"cells":[{"cell":"A1","value":"test","sheet":"Data"}]}`)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	data, _ := json.Marshal(resp.Data)
	var cd createData
	json.Unmarshal(data, &cd)
	if cd.Sheets != 2 {
		t.Errorf("expected 2 sheets, got %d", cd.Sheets)
	}
	if cd.Cells != 1 {
		t.Errorf("expected 1 cell, got %d", cd.Cells)
	}

	// Verify the file exists and is readable.
	stdout2, _, code2 := captureRunJSON("info", path)
	if code2 != 0 {
		t.Fatalf("expected exit 0 on info, got %d", code2)
	}
	if !strings.Contains(stdout2, "Data") || !strings.Contains(stdout2, "Summary") {
		t.Error("expected both sheets in info output")
	}
}

// --- Calc ---

func TestCalc(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("calc", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	// B4 should have calculated value 30 (10+20).
	data, _ := json.Marshal(resp.Data)
	var rd readData
	json.Unmarshal(data, &rd)
	for _, row := range rd.Rows {
		if row.Row == 4 {
			if cd, ok := row.Cells["B4"]; ok {
				if v, ok := cd.Value.(float64); ok {
					if v != 30 {
						t.Errorf("expected B4=30, got %v", v)
					}
				} else {
					t.Errorf("expected B4 to be number, got %T", cd.Value)
				}
			} else {
				t.Error("expected B4 in row 4")
			}
		}
	}
}

func TestCalcTextCompactsEmptyColumns(t *testing.T) {
	path := createSparseColumnsTestFile(t)
	stdout, _, code := captureRun([]string{"calc", path})
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	lines := strings.Split(strings.TrimSpace(stdout), "\n")
	if len(lines) < 5 {
		t.Fatalf("expected heading plus compact table, got:\n%s", stdout)
	}
	if lines[2] != "| Values | Returns | Large   |" {
		t.Errorf("expected empty columns to be removed, got header %q", lines[2])
	}
	if lines[3] != "| 2      | 1.05    | 1000000 |" {
		t.Errorf("expected first data row to stay aligned, got %q", lines[3])
	}
}

func TestCalcWithOutput(t *testing.T) {
	path := createTestFile(t)
	dir := t.TempDir()
	outPath := filepath.Join(dir, "calced.xlsx")
	_, _, code := captureRunJSON("calc", path, "--output", outPath)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	// Verify the output file was created.
	if _, err := os.Stat(outPath); err != nil {
		t.Fatalf("expected output file to exist: %v", err)
	}
}

// --- Formula List ---

func TestFormulaList(t *testing.T) {
	stdout, _, code := captureRunJSON("formula", "list")
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	data, _ := json.Marshal(resp.Data)
	var fl formulaListData
	json.Unmarshal(data, &fl)
	if fl.Count == 0 {
		t.Error("expected non-zero function count")
	}
	// Check sorted.
	for i := 1; i < len(fl.Functions); i++ {
		if fl.Functions[i] < fl.Functions[i-1] {
			t.Errorf("functions not sorted: %s before %s", fl.Functions[i-1], fl.Functions[i])
			break
		}
	}
}

// --- Usage / Error cases ---

func TestNoArgs(t *testing.T) {
	_, _, code := captureRun([]string{})
	if code != ExitUsage {
		t.Fatalf("expected exit %d, got %d", ExitUsage, code)
	}
}

func TestUnknownCommand(t *testing.T) {
	_, stderr, code := captureRunJSON("bogus")
	if code != ExitUsage {
		t.Fatalf("expected exit %d, got %d", ExitUsage, code)
	}
	resp := parseResponse(t, stderr)
	if resp.Error.Code != ErrCodeUsage {
		t.Errorf("expected error code %s, got %s", ErrCodeUsage, resp.Error.Code)
	}
}

func TestInvalidFormat(t *testing.T) {
	stdout, stderr, code := captureRun([]string{"--mode", "agent", "--format", "xml", "version"})
	if code != ExitUsage {
		t.Fatalf("expected exit %d, got %d", ExitUsage, code)
	}
	if stderr != "" {
		t.Fatalf("expected no stderr, got %q", stderr)
	}
	resp := parseResponse(t, stdout)
	if resp.Error.Code != ErrCodeInvalidFormat {
		t.Errorf("expected error code %s, got %s", ErrCodeInvalidFormat, resp.Error.Code)
	}
}

func TestMissingFlagValue(t *testing.T) {
	_, stderr, code := captureRunJSON("read", "--sheet")
	if code != ExitUsage {
		t.Fatalf("expected exit %d, got %d", ExitUsage, code)
	}
	resp := parseResponse(t, stderr)
	if resp.Error.Message != "--sheet requires a value" {
		t.Fatalf("expected missing-value message, got %q", resp.Error.Message)
	}
}

func TestFormulaUnknownSubcommand(t *testing.T) {
	_, stderr, code := captureRunJSON("formula", "bogus")
	if code != ExitUsage {
		t.Fatalf("expected exit %d, got %d", ExitUsage, code)
	}
	resp := parseResponse(t, stderr)
	if resp.Error.Message != "unknown subcommand: bogus. Available: list" {
		t.Fatalf("expected normalized subcommand error, got %q", resp.Error.Message)
	}
}

// --- Edit with style ---

func TestEditStyle(t *testing.T) {
	path := createTestFile(t)
	patch := `[{"cell":"A1","style":{"font":{"bold":true,"color":"FF0000"}}}]`
	stdout, _, code := captureRunJSON("edit", path, "--patch", patch)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	data, _ := json.Marshal(resp.Data)
	var ed editData
	json.Unmarshal(data, &ed)
	if ed.Operations[0].Action != "set_style" {
		t.Errorf("expected action=set_style, got %s", ed.Operations[0].Action)
	}
}

func TestEditColumnWidth(t *testing.T) {
	path := createTestFile(t)
	patch := `[{"cell":"A","column_width":25.0}]`
	stdout, _, code := captureRunJSON("edit", path, "--patch", patch)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	data, _ := json.Marshal(resp.Data)
	var ed editData
	json.Unmarshal(data, &ed)
	if ed.Operations[0].Action != "set_column_width" {
		t.Errorf("expected action=set_column_width, got %s", ed.Operations[0].Action)
	}
}

func TestEditRowHeight(t *testing.T) {
	path := createTestFile(t)
	patch := `[{"row":1,"row_height":30.0}]`
	stdout, _, code := captureRunJSON("edit", path, "--patch", patch)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	data, _ := json.Marshal(resp.Data)
	var ed editData
	json.Unmarshal(data, &ed)
	if ed.Operations[0].Action != "set_row_height" {
		t.Errorf("expected action=set_row_height, got %s", ed.Operations[0].Action)
	}
}

func TestEditClearCell(t *testing.T) {
	path := createTestFile(t)
	patch := `[{"cell":"A1","value":null}]`
	stdout, _, code := captureRunJSON("edit", path, "--patch", patch)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}

	// Verify A1 is now empty.
	stdout2, _, _ := captureRunJSON("read", "--range", "A1:A1", path)
	resp2 := parseResponse(t, stdout2)
	data, _ := json.Marshal(resp2.Data)
	var rd readData
	json.Unmarshal(data, &rd)
	// A1 should either not be in rows or be empty.
	for _, row := range rd.Rows {
		if _, ok := row.Cells["A1"]; ok {
			t.Error("expected A1 to be cleared")
		}
	}
}

// --- Help ---

func TestHelpCommand(t *testing.T) {
	_, stderr, code := captureRun([]string{"help", "read"})
	if code != ExitSuccess {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stderr, "Usage: wb read") {
		t.Errorf("expected help text, got:\n%s", stderr)
	}
}

func TestHelpFlag(t *testing.T) {
	_, stderr, code := captureRun([]string{"create", "--help"})
	if code != ExitSuccess {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stderr, "Usage: wb create") {
		t.Errorf("expected help text, got:\n%s", stderr)
	}
}

func TestHelpNestedSubcommand(t *testing.T) {
	_, stderr, code := captureRun([]string{"help", "formula", "list"})
	if code != ExitSuccess {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stderr, "Usage: wb formula list") {
		t.Errorf("expected nested help text, got:\n%s", stderr)
	}
}

func TestAgentHelpCommandJSON(t *testing.T) {
	stdout, stderr, code := captureRun([]string{"--mode", "agent", "help", "read"})
	if code != ExitSuccess {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if stderr != "" {
		t.Fatalf("expected no stderr, got: %s", stderr)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	if resp.Command != "help" {
		t.Fatalf("expected command=help, got %s", resp.Command)
	}
	data, _ := json.Marshal(resp.Data)
	var hd helpData
	json.Unmarshal(data, &hd)
	if hd.Topic != "command" {
		t.Fatalf("expected topic=command, got %s", hd.Topic)
	}
	if hd.Command == nil {
		t.Fatal("expected command help payload")
	}
	if len(hd.Command.Path) != 1 || hd.Command.Path[0] != "read" {
		t.Fatalf("expected path [read], got %v", hd.Command.Path)
	}
	if len(hd.Command.SupportedFormats) != 4 {
		t.Fatalf("expected supported formats for read, got %v", hd.Command.SupportedFormats)
	}
	if len(hd.GlobalFlags) == 0 {
		t.Fatal("expected global flags in agent help payload")
	}
}

func TestCapabilitiesCommand(t *testing.T) {
	stdout, stderr, code := captureRunJSON("capabilities")
	if code != ExitSuccess {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if stderr != "" {
		t.Fatalf("expected no stderr, got: %s", stderr)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	if resp.Command != "capabilities" {
		t.Fatalf("expected command=capabilities, got %s", resp.Command)
	}
	data, _ := json.Marshal(resp.Data)
	var spec toolSpec
	json.Unmarshal(data, &spec)
	if spec.Name != "wb" {
		t.Fatalf("expected tool name wb, got %s", spec.Name)
	}
	foundRead := false
	foundCapabilities := false
	for _, cmd := range spec.Commands {
		if cmd.Name == "read" {
			foundRead = true
		}
		if cmd.Name == "capabilities" {
			foundCapabilities = true
		}
	}
	if !foundRead || !foundCapabilities {
		t.Fatalf("expected read and capabilities commands, got %+v", spec.Commands)
	}
}

// --- Has formula ---

func TestReadHasFormula(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("read", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	// B4 has a formula; has_formula should be true even without --include-formulas.
	if !strings.Contains(stdout, `"has_formula": true`) {
		t.Error("expected has_formula:true in output")
	}
	// Formula text should NOT appear without --include-formulas.
	if strings.Contains(stdout, `"formula": "SUM`) {
		t.Error("did not expect formula text without --include-formulas")
	}
}

// --- Clear ---

func TestEditClearFlag(t *testing.T) {
	path := createTestFile(t)
	patch := `[{"cell":"A1","clear":true}]`
	stdout, _, code := captureRunJSON("edit", path, "--patch", patch)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	data, _ := json.Marshal(resp.Data)
	var ed editData
	json.Unmarshal(data, &ed)
	if ed.Operations[0].Action != "clear" {
		t.Errorf("expected action=clear, got %s", ed.Operations[0].Action)
	}

	// Verify A1 is now empty.
	stdout2, _, _ := captureRunJSON("read", "--range", "A1:A1", path)
	resp2 := parseResponse(t, stdout2)
	data2, _ := json.Marshal(resp2.Data)
	var rd readData
	json.Unmarshal(data2, &rd)
	for _, row := range rd.Rows {
		if _, ok := row.Cells["A1"]; ok {
			t.Error("expected A1 to be cleared")
		}
	}
}

func TestEditClearRange(t *testing.T) {
	path := createTestFile(t)
	patch := `[{"cell":"A1:B2","clear":true}]`
	stdout, _, code := captureRunJSON("edit", path, "--patch", patch)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}

	// Verify A1, B1, A2, B2 are cleared.
	stdout2, _, _ := captureRunJSON("read", "--range", "A1:B2", path)
	resp2 := parseResponse(t, stdout2)
	data2, _ := json.Marshal(resp2.Data)
	var rd readData
	json.Unmarshal(data2, &rd)
	if len(rd.Rows) != 0 {
		t.Errorf("expected no data rows after clearing A1:B2, got %d", len(rd.Rows))
	}
}

// --- Create with rows ---

func TestCreateWithRows(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "rows.xlsx")
	spec := `{"rows":[{"start":"A1","data":[["Name","Age"],["Alice",30],["Bob",25]]}]}`
	stdout, _, code := captureRunJSON("create", path, "--spec", spec)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	data, _ := json.Marshal(resp.Data)
	var cd createData
	json.Unmarshal(data, &cd)
	if cd.Cells != 6 {
		t.Errorf("expected 6 cells applied, got %d", cd.Cells)
	}

	// Verify content.
	stdout2, _, code2 := captureRunJSON("read", "--range", "A1:B3", path)
	if code2 != 0 {
		t.Fatalf("expected exit 0 on read, got %d", code2)
	}
	if !strings.Contains(stdout2, "Alice") || !strings.Contains(stdout2, "Bob") {
		t.Error("expected Alice and Bob in read output")
	}
}

// --- Partial failure ---

func TestEditPartialFailure(t *testing.T) {
	path := createTestFile(t)
	// First op succeeds, second fails (delete non-existent sheet).
	patch := `[{"cell":"A1","value":"OK"},{"delete_sheet":"NoSuchSheet"}]`
	_, stderr, code := captureRunJSON("edit", path, "--patch", patch)
	if code != ExitPartial {
		t.Fatalf("expected exit %d, got %d", ExitPartial, code)
	}
	resp := parseResponse(t, stderr)
	if resp.OK {
		t.Fatal("expected ok=false for partial failure")
	}
	if resp.Error == nil || resp.Error.Code != ErrCodePartialFailure {
		t.Fatalf("expected PARTIAL_FAILURE error code, got %+v", resp.Error)
	}
	data, _ := json.Marshal(resp.Data)
	var ed editData
	json.Unmarshal(data, &ed)
	if ed.Applied != 1 {
		t.Errorf("expected 1 applied, got %d", ed.Applied)
	}
	if ed.Operations[0].Status != "ok" {
		t.Errorf("expected first op ok, got %s", ed.Operations[0].Status)
	}
	if ed.Operations[1].Status != "error" {
		t.Errorf("expected second op error, got %s", ed.Operations[1].Status)
	}
}

func TestEditAtomicDefaultSkipsSaveOnPartialFailure(t *testing.T) {
	path := createTestFile(t)
	patch := `[{"cell":"A1","value":"Changed"},{"delete_sheet":"NoSuchSheet"}]`
	_, stderr, code := captureRunJSON("edit", path, "--patch", patch)
	if code != ExitPartial {
		t.Fatalf("expected exit %d, got %d", ExitPartial, code)
	}
	resp := parseResponse(t, stderr)
	data, _ := json.Marshal(resp.Data)
	var ed editData
	json.Unmarshal(data, &ed)
	if ed.Saved {
		t.Fatal("expected saved=false in atomic partial failure")
	}

	// A1 should remain unchanged because atomic mode skips save.
	stdout2, _, code2 := captureRunJSON("read", "--range", "A1:A1", path)
	if code2 != 0 {
		t.Fatalf("expected read exit 0, got %d", code2)
	}
	if !strings.Contains(stdout2, "Name") {
		t.Fatalf("expected A1 to remain Name, got: %s", stdout2)
	}
}

func TestEditNoAtomicAllowsPartialSave(t *testing.T) {
	path := createTestFile(t)
	patch := `[{"cell":"A1","value":"Changed"},{"delete_sheet":"NoSuchSheet"}]`
	_, stderr, code := captureRunJSON("edit", "--no-atomic", path, "--patch", patch)
	if code != ExitPartial {
		t.Fatalf("expected exit %d, got %d", ExitPartial, code)
	}
	resp := parseResponse(t, stderr)
	data, _ := json.Marshal(resp.Data)
	var ed editData
	json.Unmarshal(data, &ed)
	if !ed.Saved {
		t.Fatal("expected saved=true with --no-atomic")
	}

	stdout2, _, code2 := captureRunJSON("read", "--range", "A1:A1", path)
	if code2 != 0 {
		t.Fatalf("expected read exit 0, got %d", code2)
	}
	if !strings.Contains(stdout2, "Changed") {
		t.Fatalf("expected A1 to be Changed, got: %s", stdout2)
	}
}

func TestEditValidateOnlyDoesNotSave(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("edit", "--validate-only", path, "--patch", `[{"cell":"A1","value":"Changed"}]`)
	if code != ExitSuccess {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var ed editData
	json.Unmarshal(data, &ed)
	if !ed.ValidateOnly || !ed.DryRun {
		t.Fatalf("expected validate_only and dry_run true, got %+v", ed)
	}
	if ed.Saved {
		t.Fatal("expected saved=false with --validate-only")
	}

	stdout2, _, code2 := captureRunJSON("read", "--range", "A1:A1", path)
	if code2 != 0 {
		t.Fatalf("expected read exit 0, got %d", code2)
	}
	if !strings.Contains(stdout2, "Name") {
		t.Fatalf("expected A1 to remain Name, got: %s", stdout2)
	}
}

func TestEditUnknownPatchFieldRejected(t *testing.T) {
	path := createTestFile(t)
	_, stderr, code := captureRunJSON("edit", path, "--patch", `[{"cell":"A1","value":"x","bogus":1}]`)
	if code != ExitValidate {
		t.Fatalf("expected exit %d, got %d", ExitValidate, code)
	}
	resp := parseResponse(t, stderr)
	if resp.Error == nil || resp.Error.Code != ErrCodeInvalidPatch {
		t.Fatalf("expected INVALID_PATCH, got %+v", resp.Error)
	}
}

func TestEditUnknownStyleFieldRejected(t *testing.T) {
	path := createTestFile(t)
	patch := `[{"cell":"A1","style":{"font":{"bold":true,"bogus":1}}}]`
	_, stderr, code := captureRunJSON("edit", path, "--patch", patch)
	if code != ExitPartial {
		t.Fatalf("expected exit %d, got %d", ExitPartial, code)
	}
	resp := parseResponse(t, stderr)
	if resp.Error == nil || resp.Error.Code != ErrCodePartialFailure {
		t.Fatalf("expected PARTIAL_FAILURE, got %+v", resp.Error)
	}
	data, _ := json.Marshal(resp.Data)
	var ed editData
	json.Unmarshal(data, &ed)
	if len(ed.Operations) == 0 || ed.Operations[0].Status != "error" {
		t.Fatalf("expected first operation error, got %+v", ed.Operations)
	}
}

func TestCreateUnknownSpecFieldRejected(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "bad.xlsx")
	_, stderr, code := captureRunJSON("create", path, "--spec", `{"sheets":["S1"],"bogus":1}`)
	if code != ExitValidate {
		t.Fatalf("expected exit %d, got %d", ExitValidate, code)
	}
	resp := parseResponse(t, stderr)
	if resp.Error == nil || resp.Error.Code != ErrCodeInvalidSpec {
		t.Fatalf("expected INVALID_SPEC, got %+v", resp.Error)
	}
}

func TestCreatePartialFailureDoesNotSave(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "partial.xlsx")
	spec := `{"cells":[{"cell":"A1","value":"ok"},{"cell":"ZZZ1","value":1}]}`
	_, stderr, code := captureRunJSON("create", path, "--spec", spec)
	if code != ExitPartial {
		t.Fatalf("expected exit %d, got %d", ExitPartial, code)
	}
	resp := parseResponse(t, stderr)
	data, _ := json.Marshal(resp.Data)
	var cd createData
	json.Unmarshal(data, &cd)
	if cd.Saved {
		t.Fatal("expected saved=false on create partial failure")
	}
	if cd.Failed != 1 {
		t.Fatalf("expected failed=1, got %d", cd.Failed)
	}
	if _, err := os.Stat(path); !os.IsNotExist(err) {
		t.Fatalf("expected output file to not exist, stat err=%v", err)
	}
}

func TestModeAgentForcesJSONEnvelope(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRun([]string{"--mode", "agent", "--format", "markdown", "read", "--headers", path})
	if code != ExitSuccess {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if strings.Contains(stdout, "| Name | Value |") {
		t.Fatalf("expected JSON envelope, got markdown: %s", stdout)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	if resp.Meta == nil {
		t.Fatal("expected meta in response")
	}
	if resp.Meta.Mode != modeAgent {
		t.Fatalf("expected meta.mode=%q, got %q", modeAgent, resp.Meta.Mode)
	}
	if len(resp.Meta.Warnings) == 0 {
		t.Fatal("expected warning for forced json format")
	}
}

func TestCompactJSON(t *testing.T) {
	stdout, _, code := captureRun([]string{"--format", "json", "--compact", "version"})
	if code != ExitSuccess {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if strings.Contains(stdout, "\n  ") {
		t.Fatalf("expected compact json (no indentation), got: %s", stdout)
	}
	resp := parseResponse(t, strings.TrimSpace(stdout))
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
}

// --- Global --help flag ---

func TestGlobalHelpFlag(t *testing.T) {
	_, stderr, code := captureRun([]string{"--help"})
	if code != ExitSuccess {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stderr, "Usage: wb") {
		t.Errorf("expected usage text, got:\n%s", stderr)
	}
}

func TestGlobalHelpShortFlag(t *testing.T) {
	_, stderr, code := captureRun([]string{"-h"})
	if code != ExitSuccess {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stderr, "Usage: wb") {
		t.Errorf("expected usage text, got:\n%s", stderr)
	}
}

func TestHelpNoArgsExitSuccess(t *testing.T) {
	_, _, code := captureRun([]string{"help"})
	if code != ExitSuccess {
		t.Fatalf("expected exit 0 for help with no args, got %d", code)
	}
}

// --- --limit flag ---

func TestReadLimit(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("read", "--headers", "--limit", "1", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var rd readData
	json.Unmarshal(data, &rd)
	if len(rd.Rows) != 1 {
		t.Errorf("expected 1 row with --limit 1, got %d", len(rd.Rows))
	}
}

func TestReadLimitMarkdown(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRun([]string{"read", "--format", "markdown", "--headers", "--limit", "1", path})
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	// Should have header + separator + 1 data row = 3 lines.
	lines := strings.Split(strings.TrimSpace(stdout), "\n")
	if len(lines) != 3 {
		t.Errorf("expected 3 markdown lines (header+sep+1row), got %d:\n%s", len(lines), stdout)
	}
}

func TestReadHeadAlias(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("read", "--headers", "--head", "1", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var rd readData
	json.Unmarshal(data, &rd)
	if len(rd.Rows) != 1 {
		t.Errorf("expected 1 row with --head 1, got %d", len(rd.Rows))
	}
}

// --- --all-sheets flag ---

func createMultiSheetTestFile(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "multi.xlsx")
	f := werkbook.New(werkbook.FirstSheet("Data"))
	s := f.Sheet("Data")
	s.SetValue("A1", "Name")
	s.SetValue("B1", "Score")
	s.SetValue("A2", "Alice")
	s.SetValue("B2", 90.0)
	s2, _ := f.NewSheet("Summary")
	s2.SetValue("A1", "Total")
	s2.SetValue("B1", 90.0)
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("failed to create multi-sheet test file: %v", err)
	}
	return path
}

func TestReadAllSheetsJSON(t *testing.T) {
	path := createMultiSheetTestFile(t)
	stdout, _, code := captureRunJSON("read", "--all-sheets", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	if !resp.OK {
		t.Fatal("expected ok=true")
	}
	data, _ := json.Marshal(resp.Data)
	var md readMultiData
	json.Unmarshal(data, &md)
	if len(md.Sheets) != 2 {
		t.Fatalf("expected 2 sheets, got %d", len(md.Sheets))
	}
	if md.Sheets[0].Sheet != "Data" {
		t.Errorf("expected first sheet=Data, got %s", md.Sheets[0].Sheet)
	}
	if md.Sheets[1].Sheet != "Summary" {
		t.Errorf("expected second sheet=Summary, got %s", md.Sheets[1].Sheet)
	}
}

func TestReadAllSheetsMarkdown(t *testing.T) {
	path := createMultiSheetTestFile(t)
	stdout, _, code := captureRun([]string{"read", "--format", "markdown", "--all-sheets", path})
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stdout, "## Data") {
		t.Errorf("expected '## Data' header, got:\n%s", stdout)
	}
	if !strings.Contains(stdout, "## Summary") {
		t.Errorf("expected '## Summary' header, got:\n%s", stdout)
	}
}

func TestReadAllSheetsCSV(t *testing.T) {
	path := createMultiSheetTestFile(t)
	stdout, _, code := captureRun([]string{"read", "--format", "csv", "--all-sheets", path})
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stdout, "# Data") {
		t.Errorf("expected '# Data' header, got:\n%s", stdout)
	}
	if !strings.Contains(stdout, "# Summary") {
		t.Errorf("expected '# Summary' header, got:\n%s", stdout)
	}
}

func TestReadAllSheetsAndSheetMutuallyExclusive(t *testing.T) {
	path := createMultiSheetTestFile(t)
	_, stderr, code := captureRunJSON("read", "--all-sheets", "--sheet", "Data", path)
	if code != ExitUsage {
		t.Fatalf("expected exit %d, got %d", ExitUsage, code)
	}
	resp := parseResponse(t, stderr)
	if resp.OK {
		t.Fatal("expected ok=false")
	}
}

// --- --where flag ---

func TestReadWhereEquals(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("read", "--headers", "--where", "Name=Alpha", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var rd readData
	json.Unmarshal(data, &rd)
	if len(rd.Rows) != 1 {
		t.Fatalf("expected 1 filtered row, got %d", len(rd.Rows))
	}
	if rd.Rows[0].Row != 2 {
		t.Errorf("expected filtered row to be row 2, got %d", rd.Rows[0].Row)
	}
}

func TestReadWhereNotEquals(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("read", "--headers", "--where", "Name!=Alpha", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var rd readData
	json.Unmarshal(data, &rd)
	// Should get Beta (row 3) and the formula row (row 4).
	for _, row := range rd.Rows {
		if row.Row == 2 {
			t.Error("did not expect row 2 (Alpha) in !=Alpha filter")
		}
	}
}

func TestReadWhereNumericGt(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("read", "--headers", "--where", "Value>15", "--range", "A1:B3", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var rd readData
	json.Unmarshal(data, &rd)
	if len(rd.Rows) != 1 {
		t.Fatalf("expected 1 row with Value>15, got %d", len(rd.Rows))
	}
	if rd.Rows[0].Row != 3 {
		t.Errorf("expected row 3 (Beta/20), got row %d", rd.Rows[0].Row)
	}
}

func TestReadWhereByColumnLetter(t *testing.T) {
	path := createTestFile(t)
	// No --headers, use column letter A.
	stdout, _, code := captureRunJSON("read", "--where", "A=Alpha", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var rd readData
	json.Unmarshal(data, &rd)
	if len(rd.Rows) != 1 {
		t.Fatalf("expected 1 filtered row, got %d", len(rd.Rows))
	}
}

func TestReadWhereMarkdown(t *testing.T) {
	path := createTestFile(t)
	stdout, _, code := captureRun([]string{"read", "--format", "markdown", "--headers", "--where", "Name=Beta", path})
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stdout, "Beta") {
		t.Errorf("expected Beta in output, got:\n%s", stdout)
	}
	if strings.Contains(stdout, "Alpha") {
		t.Errorf("did not expect Alpha in filtered output, got:\n%s", stdout)
	}
}

func TestReadWhereWithLimit(t *testing.T) {
	// Test that --limit applies after --where.
	path := createTestFile(t)
	stdout, _, code := captureRunJSON("read", "--headers", "--where", "Value>=10", "--limit", "1", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	resp := parseResponse(t, stdout)
	data, _ := json.Marshal(resp.Data)
	var rd readData
	json.Unmarshal(data, &rd)
	if len(rd.Rows) != 1 {
		t.Fatalf("expected 1 row (limit after filter), got %d", len(rd.Rows))
	}
}

func TestReadWhereInvalidExpr(t *testing.T) {
	path := createTestFile(t)
	_, stderr, code := captureRunJSON("read", "--where", "badexpr", path)
	if code != ExitUsage {
		t.Fatalf("expected exit %d, got %d", ExitUsage, code)
	}
	resp := parseResponse(t, stderr)
	if resp.OK {
		t.Fatal("expected ok=false for invalid --where")
	}
}

// --- Edit help formula note ---

func TestEditHelpFormulaNote(t *testing.T) {
	_, stderr, code := captureRun([]string{"edit", "--help"})
	if code != ExitSuccess {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stderr, "auto-expand formula ranges") {
		t.Errorf("expected formula range note in edit help, got:\n%s", stderr)
	}
}

// --- --style-summary flag ---

func createStyledTestFile(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "styled.xlsx")
	f := werkbook.New(werkbook.FirstSheet("Sheet1"))
	s := f.Sheet("Sheet1")
	s.SetValue("A1", "Header")
	s.SetValue("A2", "Data")
	s.SetStyle("A1", &werkbook.Style{
		Font: &werkbook.Font{Bold: true, Size: 14},
	})
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("failed to create styled test file: %v", err)
	}
	return path
}

func TestReadStyleSummaryJSON(t *testing.T) {
	path := createStyledTestFile(t)
	stdout, _, code := captureRunJSON("read", "--style-summary", path)
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stdout, `"style_summary"`) {
		t.Errorf("expected style_summary field in JSON output, got:\n%s", stdout)
	}
	if !strings.Contains(stdout, "bold") {
		t.Errorf("expected 'bold' in style summary, got:\n%s", stdout)
	}
}

func TestReadStyleSummaryMarkdown(t *testing.T) {
	path := createStyledTestFile(t)
	stdout, _, code := captureRun([]string{"read", "--format", "markdown", "--headers", "--style-summary", path})
	if code != 0 {
		t.Fatalf("expected exit 0, got %d", code)
	}
	if !strings.Contains(stdout, "Style") {
		t.Errorf("expected Style column header in markdown output, got:\n%s", stdout)
	}
}

// --- filter.go unit tests ---

func TestParseWhere(t *testing.T) {
	tests := []struct {
		input string
		col   string
		op    string
		val   string
	}{
		{"Name=Alice", "Name", "=", "Alice"},
		{"Age>=30", "Age", ">=", "30"},
		{"Score!=0", "Score", "!=", "0"},
		{"A<100", "A", "<", "100"},
		{"B>5", "B", ">", "5"},
		{"X<=10", "X", "<=", "10"},
	}
	for _, tt := range tests {
		fc, err := parseWhere(tt.input)
		if err != nil {
			t.Errorf("parseWhere(%q): unexpected error: %v", tt.input, err)
			continue
		}
		if fc.Column != tt.col || fc.Op != tt.op || fc.Value != tt.val {
			t.Errorf("parseWhere(%q) = {%q, %q, %q}, want {%q, %q, %q}",
				tt.input, fc.Column, fc.Op, fc.Value, tt.col, tt.op, tt.val)
		}
	}
}

func TestParseWhereInvalid(t *testing.T) {
	_, err := parseWhere("justtext")
	if err == nil {
		t.Error("expected error for invalid --where expression")
	}
}

func TestColumnNumberToLetter(t *testing.T) {
	tests := []struct {
		col    int
		expect string
	}{
		{1, "A"},
		{2, "B"},
		{26, "Z"},
		{27, "AA"},
	}
	for _, tt := range tests {
		got := columnNumberToLetter(tt.col)
		if got != tt.expect {
			t.Errorf("columnNumberToLetter(%d) = %q, want %q", tt.col, got, tt.expect)
		}
	}
}

func TestCompareValues(t *testing.T) {
	// Numeric comparison.
	if !compareValues("10", ">", "5") {
		t.Error("expected 10 > 5")
	}
	if compareValues("3", ">", "5") {
		t.Error("did not expect 3 > 5")
	}
	// String comparison (case-insensitive for =).
	if !compareValues("hello", "=", "Hello") {
		t.Error("expected case-insensitive equal")
	}
	if !compareValues("hello", "!=", "world") {
		t.Error("expected hello != world")
	}
}

func TestStyleSummary(t *testing.T) {
	s := &werkbook.Style{
		Font: &werkbook.Font{Bold: true, Size: 12, Color: "FF0000"},
		Fill: &werkbook.Fill{Color: "00FF00"},
	}
	summary := styleSummary(s)
	if !strings.Contains(summary, "bold") {
		t.Errorf("expected 'bold' in summary: %s", summary)
	}
	if !strings.Contains(summary, "12pt") {
		t.Errorf("expected '12pt' in summary: %s", summary)
	}
	if !strings.Contains(summary, "color:#FF0000") {
		t.Errorf("expected 'color:#FF0000' in summary: %s", summary)
	}
	if !strings.Contains(summary, "fill:#00FF00") {
		t.Errorf("expected 'fill:#00FF00' in summary: %s", summary)
	}
}

func TestStyleSummaryNil(t *testing.T) {
	if styleSummary(nil) != "" {
		t.Error("expected empty string for nil style")
	}
}
