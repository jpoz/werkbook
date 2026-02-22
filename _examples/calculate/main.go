package main

import (
	"fmt"
	"log"
	"os"

	"github.com/jpoz/werkbook"
)

func main() {
	wb := werkbook.New(werkbook.FirstSheet("Calculations"))
	s := wb.Sheet("Calculations")

	// --- Section 1: Loan amortization inputs (rows 1-6) ---
	s.SetValue("A1", "Loan Amortization")

	s.SetValue("A3", "Principal")
	s.SetValue("B3", 250000.0)

	s.SetValue("A4", "Annual Rate (%)")
	s.SetValue("B4", 6.5)

	s.SetValue("A5", "Term (Years)")
	s.SetValue("B5", 30)

	s.SetValue("A6", "Monthly Payment")
	// PMT = P * r(1+r)^n / ((1+r)^n - 1)  where r = annual/12/100, n = years*12
	s.SetFormula("B6", "B3*(B4/12/100)*POWER(1+B4/12/100,B5*12)/(POWER(1+B4/12/100,B5*12)-1)")

	// --- Section 2: Sales data with conditional aggregation (rows 8-22) ---
	s.SetValue("A8", "Quarterly Sales Analysis")

	headers := []string{"Region", "Q1", "Q2", "Q3", "Q4", "Annual", "Avg", "Best Qtr", "Above Target"}
	for i, h := range headers {
		cell, _ := werkbook.CoordinatesToCellName(i+1, 9)
		s.SetValue(cell, h)
	}

	regions := []struct {
		name string
		q    [4]float64
	}{
		{"North", [4]float64{84500, 91200, 78300, 102400}},
		{"South", [4]float64{67800, 72100, 81900, 69500}},
		{"East", [4]float64{93200, 88400, 95100, 110800}},
		{"West", [4]float64{71000, 65300, 70200, 83600}},
		{"Central", [4]float64{55600, 61200, 58900, 64300}},
	}

	target := 80000.0
	s.SetValue("A16", "Target")
	s.SetValue("B16", target)

	for i, r := range regions {
		row := 10 + i
		s.SetValue(fmt.Sprintf("A%d", row), r.name)
		for q := 0; q < 4; q++ {
			s.SetValue(fmt.Sprintf("%s%d", string(rune('B'+q)), row), r.q[q])
		}
		// Annual total
		s.SetFormula(fmt.Sprintf("F%d", row), fmt.Sprintf("SUM(B%d:E%d)", row, row))
		// Quarterly average
		s.SetFormula(fmt.Sprintf("G%d", row), fmt.Sprintf("AVERAGE(B%d:E%d)", row, row))
		// Best quarter value
		s.SetFormula(fmt.Sprintf("H%d", row), fmt.Sprintf("MAX(B%d:E%d)", row, row))
		// Count of quarters above target
		s.SetFormula(fmt.Sprintf("I%d", row), fmt.Sprintf("COUNTIF(B%d:E%d,\">\"&B16)", row, row))
	}

	// Summary row
	s.SetValue("A17", "All Regions")
	s.SetFormula("F17", "SUM(F10:F14)")
	s.SetFormula("G17", "AVERAGE(G10:G14)")
	s.SetFormula("H17", "MAX(H10:H14)")

	// --- Section 3: Grade book with weighted scoring (rows 19-32) ---
	s.SetValue("A19", "Weighted Grade Book")

	s.SetValue("B20", "Homework")
	s.SetValue("C20", "Midterm")
	s.SetValue("D20", "Final")
	s.SetValue("E20", "Project")
	s.SetValue("F20", "Weighted %")
	s.SetValue("G20", "Letter")

	// Weights
	s.SetValue("A21", "Weight")
	s.SetValue("B21", 0.20)
	s.SetValue("C21", 0.25)
	s.SetValue("D21", 0.35)
	s.SetValue("E21", 0.20)

	students := []struct {
		name string
		hw   float64
		mid  float64
		fin  float64
		proj float64
	}{
		{"Martinez", 88, 76, 82, 91},
		{"Johnson", 95, 89, 93, 87},
		{"Williams", 72, 68, 71, 78},
		{"Brown", 64, 55, 60, 70},
		{"Davis", 91, 94, 88, 96},
		{"Wilson", 83, 79, 85, 80},
	}

	for i, st := range students {
		row := 22 + i
		s.SetValue(fmt.Sprintf("A%d", row), st.name)
		s.SetValue(fmt.Sprintf("B%d", row), st.hw)
		s.SetValue(fmt.Sprintf("C%d", row), st.mid)
		s.SetValue(fmt.Sprintf("D%d", row), st.fin)
		s.SetValue(fmt.Sprintf("E%d", row), st.proj)

		// Weighted score = SUMPRODUCT of scores and weights
		s.SetFormula(fmt.Sprintf("F%d", row), fmt.Sprintf("SUMPRODUCT(B%d:E%d,$B$21:$E$21)", row, row))

		// Letter grade via nested IF
		f := fmt.Sprintf("F%d", row)
		s.SetFormula(fmt.Sprintf("G%d", row),
			fmt.Sprintf("IF(%s>=90,\"A\",IF(%s>=80,\"B\",IF(%s>=70,\"C\",IF(%s>=60,\"D\",\"F\"))))", f, f, f, f))
	}

	// Class statistics
	lastStudent := 22 + len(students) - 1
	statsRow := lastStudent + 2

	s.SetValue(fmt.Sprintf("A%d", statsRow), "Class Mean")
	s.SetFormula(fmt.Sprintf("F%d", statsRow), fmt.Sprintf("AVERAGE(F22:F%d)", lastStudent))

	s.SetValue(fmt.Sprintf("A%d", statsRow+1), "Median")
	s.SetFormula(fmt.Sprintf("F%d", statsRow+1), fmt.Sprintf("MEDIAN(F22:F%d)", lastStudent))

	s.SetValue(fmt.Sprintf("A%d", statsRow+2), "Highest")
	s.SetFormula(fmt.Sprintf("F%d", statsRow+2), fmt.Sprintf("MAX(F22:F%d)", lastStudent))

	s.SetValue(fmt.Sprintf("A%d", statsRow+3), "Lowest")
	s.SetFormula(fmt.Sprintf("F%d", statsRow+3), fmt.Sprintf("MIN(F22:F%d)", lastStudent))

	s.SetValue(fmt.Sprintf("A%d", statsRow+4), "Pass Rate (%)")
	s.SetFormula(fmt.Sprintf("F%d", statsRow+4),
		fmt.Sprintf("COUNTIF(F22:F%d,\">=\"&60)/COUNTA(F22:F%d)*100", lastStudent, lastStudent))

	// --- Section 4: Unit conversion table (rows 36-44) ---
	convRow := statsRow + 7
	s.SetValue(fmt.Sprintf("A%d", convRow), "Unit Conversions")

	s.SetValue(fmt.Sprintf("A%d", convRow+1), "Celsius")
	s.SetValue(fmt.Sprintf("B%d", convRow+1), "Fahrenheit")
	s.SetValue(fmt.Sprintf("C%d", convRow+1), "Kelvin")

	temps := []float64{-40, 0, 20, 37, 100}
	for i, c := range temps {
		row := convRow + 2 + i
		s.SetValue(fmt.Sprintf("A%d", row), c)
		// F = C * 9/5 + 32
		s.SetFormula(fmt.Sprintf("B%d", row), fmt.Sprintf("A%d*9/5+32", row))
		// K = C + 273.15
		s.SetFormula(fmt.Sprintf("C%d", row), fmt.Sprintf("A%d+273.15", row))
	}

	// --- Section 5: Compound interest table (rows 48-58) ---
	ciRow := convRow + 2 + len(temps) + 2
	s.SetValue(fmt.Sprintf("A%d", ciRow), "Compound Interest Growth")

	s.SetValue(fmt.Sprintf("A%d", ciRow+1), "Initial")
	s.SetValue(fmt.Sprintf("B%d", ciRow+1), 10000.0)

	s.SetValue(fmt.Sprintf("A%d", ciRow+2), "Rate (%)")
	s.SetValue(fmt.Sprintf("B%d", ciRow+2), 7.0)

	headerRow := ciRow + 3
	s.SetValue(fmt.Sprintf("A%d", headerRow), "Year")
	s.SetValue(fmt.Sprintf("B%d", headerRow), "Balance")
	s.SetValue(fmt.Sprintf("C%d", headerRow), "Interest Earned")
	s.SetValue(fmt.Sprintf("D%d", headerRow), "Total Interest")

	initRef := fmt.Sprintf("B%d", ciRow+1)
	rateRef := fmt.Sprintf("B%d", ciRow+2)
	for yr := 1; yr <= 10; yr++ {
		row := headerRow + yr
		s.SetValue(fmt.Sprintf("A%d", row), yr)
		// Balance = Principal * (1 + rate/100)^year
		s.SetFormula(fmt.Sprintf("B%d", row),
			fmt.Sprintf("ROUND(%s*POWER(1+%s/100,A%d),2)", initRef, rateRef, row))
		// Interest this year = balance - previous balance
		if yr == 1 {
			s.SetFormula(fmt.Sprintf("C%d", row),
				fmt.Sprintf("B%d-%s", row, initRef))
		} else {
			s.SetFormula(fmt.Sprintf("C%d", row),
				fmt.Sprintf("B%d-B%d", row, row-1))
		}
		// Cumulative interest = balance - initial
		s.SetFormula(fmt.Sprintf("D%d", row),
			fmt.Sprintf("B%d-%s", row, initRef))
	}

	if err := wb.SaveAs("calculate.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("Created calculate.xlsx")
	fmt.Println()

	s.PrintTo(os.Stdout)
}
