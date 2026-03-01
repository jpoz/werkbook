package main

import (
	"fmt"

	"github.com/jpoz/werkbook/internal/fuzz"
)

func dateScenarios() []TestCaseDef {
	var cases []TestCaseDef

	// DATE function
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"DATE"},
		Description:  "DATE construction",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "DATE(2024,6,15)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// DAY
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"DAY"},
		Description:  "DAY extraction",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "DAY(DATE(2024,6,15))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// MONTH
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"MONTH"},
		Description:  "MONTH extraction",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "MONTH(DATE(2024,6,15))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// YEAR
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"YEAR"},
		Description:  "YEAR extraction",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "YEAR(DATE(2024,6,15))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// HOUR
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"HOUR"},
		Description:  "HOUR extraction",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "HOUR(TIME(14,30,45))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// MINUTE
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"MINUTE"},
		Description:  "MINUTE extraction",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "MINUTE(TIME(14,30,45))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// SECOND
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"SECOND"},
		Description:  "SECOND extraction",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "SECOND(TIME(14,30,45))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// TIME
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"TIME"},
		Description:  "TIME construction",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "TIME(14,30,0)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// DAYS360
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"DAYS360"},
		Description: "DAYS360 US method",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 45292), // DATE(2024,1,1) serial
			numCell("A2", 45474), // DATE(2024,7,1) serial
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "DAYS360(DATE(2024,1,1),DATE(2024,7,1))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// Date functions that use serial numbers
	dateFuncs := []struct {
		name    string
		formula string
	}{
		{"DATEDIF", `DATEDIF(DATE(2020,1,1),DATE(2024,6,15),"Y")`},
		{"DAYS", "DAYS(DATE(2024,12,31),DATE(2024,1,1))"},
		{"EDATE", "EDATE(DATE(2024,1,15),3)"},
		{"EOMONTH", "EOMONTH(DATE(2024,1,15),2)"},
		{"ISOWEEKNUM", "ISOWEEKNUM(DATE(2024,1,1))"},
		{"WEEKDAY", "WEEKDAY(DATE(2024,6,15))"},
		{"WEEKNUM", "WEEKNUM(DATE(2024,6,15))"},
	}
	for _, f := range dateFuncs {
		cases = append(cases, TestCaseDef{
			FuncNames:    []string{f.name},
			Description:  fmt.Sprintf("%s date calc", f.name),
			FormulaCells: []fuzz.CellSpec{formulaCell("G1", f.formula)},
			Checks:       []fuzz.CheckSpec{numCheck("G1")},
		})
	}

	// DATEVALUE
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"DATEVALUE"},
		Description:  "DATEVALUE from string",
		InputCells:   []fuzz.CellSpec{strCell("A1", "2024-06-15")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "DATEVALUE(A1)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// NETWORKDAYS
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"NETWORKDAYS"},
		Description:  "NETWORKDAYS working days",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// WORKDAY
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"WORKDAY"},
		Description:  "WORKDAY add working days",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "WORKDAY(DATE(2024,1,1),10)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// YEARFRAC
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"YEARFRAC"},
		Description:  "YEARFRAC fraction of year",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "YEARFRAC(DATE(2024,1,1),DATE(2024,7,1))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	return cases
}
