package main

import (
	"fmt"
	"log"
	"os"

	"github.com/jpoz/werkbook"
)

func main() {
	wb := werkbook.New(werkbook.FirstSheet("Output"))

	// Create the Ledger sheet with transaction data.
	ledger, err := wb.NewSheet("Ledger")
	if err != nil {
		log.Fatal(err)
	}

	// Headers
	ledger.SetValue("A1", "id")
	ledger.SetValue("B1", "type")
	ledger.SetValue("C1", "amount")
	ledger.SetValue("D1", "description")
	ledger.SetValue("E1", "abs_amount")

	// Sample transactions
	transactions := []struct {
		ID          int
		Type        string
		Amount      float64
		Description string
	}{
		{1, "income", 5000.00, "Monthly salary"},
		{2, "expense", -1200.00, "Rent payment"},
		{3, "income", 250.00, "Freelance work"},
		{4, "expense", -85.50, "Groceries"},
		{5, "transfer", -500.00, "Savings transfer"},
		{6, "income", 12000.00, "Bonus"},
		{7, "expense", -2300.00, "Car payment"},
		{8, "transfer", 1500.00, "Investment return"},
		{9, "expense", -45.00, "Subscription"},
		{10, "income", 800.00, "Side project"},
		{11, "transfer", -3000.00, "Wire to brokerage"},
		{12, "expense", -650.00, "Utilities"},
	}

	for i, t := range transactions {
		row := i + 2
		ledger.SetValue(fmt.Sprintf("A%d", row), t.ID)
		ledger.SetValue(fmt.Sprintf("B%d", row), t.Type)
		ledger.SetValue(fmt.Sprintf("C%d", row), t.Amount)
		ledger.SetValue(fmt.Sprintf("D%d", row), t.Description)
		ledger.SetFormula(fmt.Sprintf("E%d", row), fmt.Sprintf("ABS(C%d)", row))
	}

	// Build the Output sheet.
	output := wb.Sheet("Output")

	output.SetValue("A1", "Type")
	output.SetValue("B1", "Largest Absolute Transaction")

	types := []string{"income", "expense", "transfer"}
	lastRow := len(transactions) + 1

	for i, typ := range types {
		row := i + 2
		output.SetValue(fmt.Sprintf("A%d", row), typ)

		// MAXIFS using the abs_amount helper column.
		formula := fmt.Sprintf(
			"MAXIFS(Ledger!E2:E%d,Ledger!B2:B%d,A%d)",
			lastRow, lastRow, row,
		)
		output.SetFormula(fmt.Sprintf("B%d", row), formula)
	}

	if err := wb.SaveAs("ledger.xlsx"); err != nil {
		log.Fatal(err)
	}

	// Print results
	wb.Recalculate()
	fmt.Println("=== Output ===")
	output.PrintTo(os.Stdout)

	fmt.Println("\nCreated ledger.xlsx")
}
