package werkbook_test

import (
	"bytes"
	"context"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
	"time"

	"github.com/jpoz/werkbook"
)

const (
	excelSmokeEnv      = "WERKBOOK_EXCEL_SMOKE"
	excelSmokeCasesEnv = "WERKBOOK_EXCEL_SMOKE_CASES"
	excelSmokeKeepEnv  = "WERKBOOK_EXCEL_SMOKE_KEEP"
)

const excelSmokeAppleScript = `on run argv
  set workbookPath to POSIX file (item 1 of argv)
  tell application "Microsoft Excel"
    activate
    set wb to missing value
    try
      set wb to open workbook workbook file name workbookPath
      try
        calculate full
      end try
      save wb
      close wb saving no
    on error errMsg number errNum
      if wb is not missing value then
        try
          close wb saving no
        end try
      end if
      error errMsg number errNum
    end try
  end tell
end run
`

type excelSmokeCase struct {
	name     string
	formulas []excelSmokeFormula
}

type excelSmokeFormula struct {
	label   string
	formula string
}

func TestExcelSmokeFormulaFamilies(t *testing.T) {
	requireExcelSmoke(t)

	for _, tc := range filterExcelSmokeCases(t, excelSmokeCases()) {
		t.Run(tc.name, func(t *testing.T) {
			t.Helper()

			dir := excelSmokeDir(t, tc.name)
			path := filepath.Join(dir, tc.name+".xlsx")
			t.Logf("Excel smoke workbook: %s", path)

			wb := newExcelSmokeWorkbook(t, tc.formulas)
			if err := wb.SaveAs(path); err != nil {
				t.Fatalf("SaveAs(%s): %v", path, err)
			}

			if err := openAndSaveInExcel(path); err != nil {
				t.Fatalf("Excel smoke failed for %s (%s): %v", tc.name, path, err)
			}
		})
	}
}

func excelSmokeDir(t *testing.T, caseName string) string {
	t.Helper()

	if os.Getenv(excelSmokeKeepEnv) == "" {
		return t.TempDir()
	}

	dir, err := os.MkdirTemp("", "werkbook-excel-smoke-"+caseName+"-")
	if err != nil {
		t.Fatalf("MkdirTemp: %v", err)
	}
	t.Logf("keeping Excel smoke artifacts in %s", dir)
	return dir
}

func requireExcelSmoke(t *testing.T) {
	t.Helper()

	if os.Getenv(excelSmokeEnv) == "" {
		t.Skipf("set %s=1 to run Excel smoke tests", excelSmokeEnv)
	}
	if runtime.GOOS != "darwin" {
		t.Skip("Excel smoke tests require macOS")
	}
	if _, err := exec.LookPath("osascript"); err != nil {
		t.Skipf("osascript not available: %v", err)
	}
	cmd := exec.Command("osascript", "-e", `id of application "Microsoft Excel"`)
	if err := cmd.Run(); err != nil {
		t.Skipf("Microsoft Excel is not installed or not scriptable: %v", err)
	}
}

func filterExcelSmokeCases(t *testing.T, cases []excelSmokeCase) []excelSmokeCase {
	t.Helper()

	raw := strings.TrimSpace(os.Getenv(excelSmokeCasesEnv))
	if raw == "" {
		return cases
	}

	want := make(map[string]struct{})
	for _, part := range strings.Split(raw, ",") {
		part = strings.TrimSpace(strings.ToLower(part))
		if part != "" {
			want[part] = struct{}{}
		}
	}

	var filtered []excelSmokeCase
	for _, tc := range cases {
		if _, ok := want[strings.ToLower(tc.name)]; ok {
			filtered = append(filtered, tc)
		}
	}
	if len(filtered) == 0 {
		t.Fatalf("%s=%q matched no smoke cases", excelSmokeCasesEnv, raw)
	}
	return filtered
}

func excelSmokeCases() []excelSmokeCase {
	return []excelSmokeCase{
		{
			name: "baseline-legacy",
			formulas: []excelSmokeFormula{
				{label: "sum", formula: "SUM(Data!A1:A5)"},
				{label: "if", formula: `IF(Data!A1>0,"ok","bad")`},
			},
		},
		{
			name: "scalar-xlfn-acot",
			formulas: []excelSmokeFormula{
				{label: "acot", formula: "ACOT(1)"},
			},
		},
		{
			name: "scalar-xlfn-let",
			formulas: []excelSmokeFormula{
				{label: "let", formula: "LET(x,5,x+1)"},
			},
		},
		{
			name: "scalar-xlfn-maxifs",
			formulas: []excelSmokeFormula{
				{label: "maxifs", formula: `MAXIFS(Data!A1:A5,Data!B1:B5,"b")`},
			},
		},
		{
			name: "scalar-xlfn-textjoin",
			formulas: []excelSmokeFormula{
				{label: "textjoin", formula: `TEXTJOIN(",",TRUE,Data!B1:B5)`},
			},
		},
		{
			name: "scalar-xlfn-xor",
			formulas: []excelSmokeFormula{
				{label: "xor", formula: "XOR(TRUE,FALSE)"},
			},
		},
		{
			name: "spill-core",
			formulas: []excelSmokeFormula{
				{label: "filter", formula: `FILTER(Data!B1:B5,Data!B1:B5<>"")`},
				{label: "sort-unique", formula: `SORT(UNIQUE(Data!B1:B5))`},
				{label: "sequence", formula: "SEQUENCE(2,3,10,5)"},
				{label: "randarray", formula: "RANDARRAY(1,3)"},
				{label: "xlookup", formula: `XLOOKUP("c",Data!B1:B5,Data!A1:A5)`},
			},
		},
		{
			name: "spill-shape",
			formulas: []excelSmokeFormula{
				{label: "textsplit", formula: `TEXTSPLIT("A,B;C,D",",",";")`},
				{label: "take", formula: "TAKE(Data!C1:E3,2,2)"},
				{label: "drop", formula: "DROP(Data!C1:E3,1,1)"},
				{label: "tocol", formula: "TOCOL(Data!C1:E3)"},
				{label: "torow", formula: "TOROW(Data!C1:E3)"},
				{label: "wraprows", formula: "WRAPROWS(Data!A1:A5,2)"},
				{label: "wrapcols", formula: "WRAPCOLS(Data!A1:A5,2)"},
				{label: "vstack", formula: "VSTACK(Data!C1:E1,Data!C2:E2)"},
				{label: "hstack", formula: "HSTACK(Data!C1:C3,Data!D1:D3)"},
				{label: "choosecols", formula: "CHOOSECOLS(Data!C1:E3,1,3)"},
				{label: "chooserows", formula: "CHOOSEROWS(Data!C1:E3,1,3)"},
				{label: "expand", formula: "EXPAND(Data!C1:D2,3,3,0)"},
			},
		},
		{
			name: "spill-lambda",
			formulas: []excelSmokeFormula{
				{label: "byrow", formula: "BYROW(Data!C1:E3,LAMBDA(r,SUM(r)))"},
				{label: "bycol", formula: "BYCOL(Data!C1:E3,LAMBDA(c,SUM(c)))"},
				{label: "map", formula: "MAP(Data!A1:A3,LAMBDA(x,x+1))"},
				{label: "reduce", formula: "REDUCE(0,Data!A1:A3,LAMBDA(a,b,a+b))"},
				{label: "scan", formula: "SCAN(0,Data!A1:A3,LAMBDA(a,b,a+b))"},
				{label: "makearray", formula: "MAKEARRAY(2,3,LAMBDA(r,c,r*c))"},
			},
		},
	}
}

func newExcelSmokeWorkbook(t *testing.T, formulas []excelSmokeFormula) *werkbook.File {
	t.Helper()

	f := werkbook.New(werkbook.FirstSheet("Cases"))
	cases := f.Sheet("Cases")
	if cases == nil {
		t.Fatal("Cases sheet not found")
	}
	data, err := f.NewSheet("Data")
	if err != nil {
		t.Fatalf("NewSheet(Data): %v", err)
	}

	seedExcelSmokeData(data)

	cases.SetValue("A1", "Case")
	cases.SetValue("B1", "Formula")
	for i, entry := range formulas {
		row := i + 2
		labelCell := fmt.Sprintf("A%d", row)
		formulaCell := fmt.Sprintf("B%d", row)
		if err := cases.SetValue(labelCell, entry.label); err != nil {
			t.Fatalf("SetValue(%s): %v", labelCell, err)
		}
		if err := cases.SetFormula(formulaCell, entry.formula); err != nil {
			t.Fatalf("SetFormula(%s=%q): %v", formulaCell, entry.formula, err)
		}
	}

	return f
}

func seedExcelSmokeData(s *werkbook.Sheet) {
	values := map[string]any{
		"A1": 1, "A2": 2, "A3": 3, "A4": 4, "A5": 5,
		"B1": "a", "B2": "b", "B3": "", "B4": "b", "B5": "c",
		"C1": 1, "D1": 2, "E1": 3,
		"C2": 4, "D2": 5, "E2": 6,
		"C3": 7, "D3": 8, "E3": 9,
	}
	for cell, value := range values {
		_ = s.SetValue(cell, value)
	}
}

func openAndSaveInExcel(path string) error {
	absPath, err := filepath.Abs(path)
	if err != nil {
		return fmt.Errorf("abs path: %w", err)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 45*time.Second)
	defer cancel()

	cmd := exec.CommandContext(ctx, "osascript", "-", absPath)
	cmd.Stdin = strings.NewReader(excelSmokeAppleScript)

	var stdout, stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr

	if err := cmd.Run(); err != nil {
		if ctx.Err() == context.DeadlineExceeded {
			return fmt.Errorf("timed out opening workbook in Excel")
		}
		msg := strings.TrimSpace(stderr.String())
		if msg == "" {
			msg = strings.TrimSpace(stdout.String())
		}
		if msg != "" {
			return fmt.Errorf("%w: %s", err, msg)
		}
		return err
	}
	return nil
}
