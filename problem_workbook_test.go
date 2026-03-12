package werkbook_test

import (
	"path/filepath"
	"strconv"
	"testing"

	"github.com/jpoz/werkbook"
)

func TestProblemWorkbookRecalculateMatchesExcel(t *testing.T) {
	f := werkbook.New(werkbook.FirstSheet("Out - Summary"))
	summary := f.Sheet("Out - Summary")
	if summary == nil {
		t.Fatal("Out - Summary sheet not found")
	}

	fundExpenses, err := f.NewSheet("FundExpenses")
	if err != nil {
		t.Fatalf("NewSheet(FundExpenses): %v", err)
	}

	fundExpenses.SetValue("A1", "ID")
	fundExpenses.SetValue("B1", "Category")
	fundExpenses.SetValue("C1", "Recipient")
	fundExpenses.SetValue("D1", "Recipient Name")
	fundExpenses.SetValue("E1", "Estimated Total Cents")
	fundExpenses.SetValue("F1", "Total Paid In Cash Cents")
	fundExpenses.SetValue("G1", "Total Paid Cashless Cents")

	rows := []struct {
		id        string
		category  string
		recipient string
		name      string
		estimate  int
		cash      int
		cashless  int
	}{
		{"e4e7063389984c79823f11faa692bd0a", "administrative", "assure", "", 800000, 800000, 0},
		{"01952a75588475b8b0c523a7f2795832", "cyclical_management", "other", "", 0, 0, 0},
		{"01952a7558b275758483ef7bfb77cfb9", "cyclical_administrative", "alam", "AL Advisors Management, Inc", 0, 0, 0},
	}
	for i, row := range rows {
		r := i + 2
		rowNum := strconv.Itoa(r)
		fundExpenses.SetValue("A"+rowNum, row.id)
		fundExpenses.SetValue("B"+rowNum, row.category)
		fundExpenses.SetValue("C"+rowNum, row.recipient)
		fundExpenses.SetValue("D"+rowNum, row.name)
		fundExpenses.SetValue("E"+rowNum, row.estimate)
		fundExpenses.SetValue("F"+rowNum, row.cash)
		fundExpenses.SetValue("G"+rowNum, row.cashless)
	}

	summary.SetValue("A1", "Total Estimated Expenses")
	summary.SetValue("A2", "Total Paid In Cash")
	summary.SetValue("A3", "Total Paid Cashless")
	summary.SetValue("A4", "Total Paid")
	summary.SetValue("A5", "Total Remaining")
	summary.SetValue("A6", "Expense Count")

	summary.SetFormula("B1", "SUM(FundExpenses!E:E)/100")
	summary.SetFormula("B2", "SUM(FundExpenses!F:F)/100")
	summary.SetFormula("B3", "SUM(FundExpenses!G:G)/100")
	summary.SetFormula("B4", "(SUM(FundExpenses!F:F)+SUM(FundExpenses!G:G))/100")
	summary.SetFormula("B5", "(SUM(FundExpenses!E:E)-SUM(FundExpenses!F:F)-SUM(FundExpenses!G:G))/100")
	summary.SetFormula("B6", "MAX(COUNTA(FundExpenses!A:A)-1,0)")

	path := filepath.Join(t.TempDir(), "problem.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open(%s): %v", path, err)
	}

	f2.Recalculate()

	summary = f2.Sheet("Out - Summary")
	if summary == nil {
		t.Fatal("Out - Summary sheet not found after reopen")
	}

	tests := []struct {
		cell string
		want float64
	}{
		{cell: "B1", want: 8000},
		{cell: "B2", want: 8000},
		{cell: "B3", want: 0},
		{cell: "B4", want: 8000},
		{cell: "B5", want: 0},
		{cell: "B6", want: 3},
	}

	for _, tt := range tests {
		v, err := summary.GetValue(tt.cell)
		if err != nil {
			t.Fatalf("%s: %v", tt.cell, err)
		}
		if v.Type != werkbook.TypeNumber || v.Number != tt.want {
			t.Fatalf("%s = %#v, want %v", tt.cell, v, tt.want)
		}
	}
}

func TestCoolTest4119RecalculateMatchesExcel(t *testing.T) {
	f := werkbook.New(werkbook.FirstSheet("Out - Ledger Totals"))
	totals := f.Sheet("Out - Ledger Totals")
	if totals == nil {
		t.Fatal("Out - Ledger Totals sheet not found")
	}

	if _, err := f.NewSheet("treasury-ledger"); err != nil {
		t.Fatalf("NewSheet(treasury-ledger): %v", err)
	}

	totals.SetValue("A5", "Current Treasury Balance")
	totals.SetFormula("C5", "INDEX('treasury-ledger'!G:G,2)/100")

	if err := f.SetDefinedName(werkbook.DefinedName{
		Name:         "CurrentTreasuryBalance",
		Value:        "'Out - Ledger Totals'!$C$5",
		LocalSheetID: -1,
	}); err != nil {
		t.Fatalf("SetDefinedName(CurrentTreasuryBalance): %v", err)
	}

	path := filepath.Join(t.TempDir(), "cool-test-4119.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs(%s): %v", path, err)
	}

	f, err := werkbook.Open(path)
	if err != nil {
		t.Fatalf("Open(%s): %v", path, err)
	}

	f.Recalculate()

	vals, err := f.ResolveDefinedName("CurrentTreasuryBalance", -1)
	if err != nil {
		t.Fatalf("ResolveDefinedName(CurrentTreasuryBalance): %v", err)
	}
	if len(vals) != 1 || len(vals[0]) != 1 {
		cols := 0
		if len(vals) > 0 {
			cols = len(vals[0])
		}
		t.Fatalf("ResolveDefinedName(CurrentTreasuryBalance) shape = %dx%d, want 1x1", len(vals), cols)
	}
	if vals[0][0].Type != werkbook.TypeNumber || vals[0][0].Number != 0 {
		t.Fatalf("CurrentTreasuryBalance = %#v, want 0", vals[0][0])
	}
}
