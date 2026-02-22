package main

import (
	"fmt"
	"log"
	"time"

	"github.com/jpoz/werkbook"
)

func main() {
	wb := werkbook.New(werkbook.FirstSheet("Payroll"))
	sheet := wb.Sheet("Payroll")

	// Header row
	headers := []string{"Name", "Department", "Start Date", "Salary", "Active"}
	for i, h := range headers {
		cell, _ := werkbook.CoordinatesToCellName(i+1, 1)
		if err := sheet.SetValue(cell, h); err != nil {
			log.Fatal(err)
		}
	}

	// Data rows
	employees := []struct {
		Name       string
		Department string
		StartDate  time.Time
		Salary     float64
		Active     bool
	}{
		{"Alice Smith", "Engineering", time.Date(2022, 3, 15, 0, 0, 0, 0, time.UTC), 95000, true},
		{"Bob Jones", "Marketing", time.Date(2021, 7, 1, 0, 0, 0, 0, time.UTC), 82000, true},
		{"Carol Lee", "Engineering", time.Date(2023, 1, 10, 0, 0, 0, 0, time.UTC), 105000, true},
		{"Dave Kim", "Sales", time.Date(2020, 11, 20, 0, 0, 0, 0, time.UTC), 78000, false},
		{"Eve Chen", "Engineering", time.Date(2024, 6, 3, 0, 0, 0, 0, time.UTC), 90000, true},
	}

	for i, emp := range employees {
		row := i + 2
		sheet.SetValue(fmt.Sprintf("A%d", row), emp.Name)
		sheet.SetValue(fmt.Sprintf("B%d", row), emp.Department)
		sheet.SetValue(fmt.Sprintf("C%d", row), emp.StartDate)
		sheet.SetValue(fmt.Sprintf("D%d", row), emp.Salary)
		sheet.SetValue(fmt.Sprintf("E%d", row), emp.Active)
	}

	// Summary formulas
	sheet.SetValue("A8", "Total Salary")
	sheet.SetFormula("D8", "SUM(D2:D6)")

	sheet.SetValue("A9", "Average Salary")
	sheet.SetFormula("D9", "AVERAGE(D2:D6)")

	sheet.SetValue("A10", "Headcount")
	sheet.SetFormula("D10", "COUNTA(A2:A6)")

	// Second sheet
	inventory, err := wb.NewSheet("Inventory")
	if err != nil {
		log.Fatal(err)
	}

	items := [][]any{
		{"Item", "Quantity", "Price", "Total"},
		{"Widget", 150, 9.99, nil},
		{"Gadget", 75, 24.50, nil},
		{"Doohickey", 200, 4.75, nil},
	}

	for r, row := range items {
		for c, val := range row {
			cell, _ := werkbook.CoordinatesToCellName(c+1, r+1)
			if val != nil {
				inventory.SetValue(cell, val)
			}
		}
	}

	// Total = Quantity * Price
	for r := 2; r <= 4; r++ {
		inventory.SetFormula(fmt.Sprintf("D%d", r), fmt.Sprintf("B%d*C%d", r, r))
	}

	inventory.SetValue("A6", "Grand Total")
	inventory.SetFormula("D6", "SUM(D2:D4)")

	if err := wb.SaveAs("example.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("Created example.xlsx")
}
