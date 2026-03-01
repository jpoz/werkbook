package main

import (
	"fmt"

	"github.com/jpoz/werkbook/internal/fuzz"
)

func arrayScenarios() []TestCaseDef {
	var cases []TestCaseDef

	// Shared numeric input: A1:A5 = 1..5
	numericInputs := func() []fuzz.CellSpec {
		cells := make([]fuzz.CellSpec, 5)
		for i := 0; i < 5; i++ {
			cells[i] = numCell(fmt.Sprintf("A%d", i+1), float64(i+1))
		}
		return cells
	}

	// Test representative aggregation functions with both range and cell-list forms
	aggFuncs := []struct {
		name   string
		rangeF string
		listF  string
	}{
		{"SUM", "SUM(A1:A5)", "SUM(" + cellList("A", 1, 5) + ")"},
		{"AVERAGE", "AVERAGE(A1:A5)", "AVERAGE(" + cellList("A", 1, 5) + ")"},
		{"COUNT", "COUNT(A1:A5)", "COUNT(" + cellList("A", 1, 5) + ")"},
		{"MAX", "MAX(A1:A5)", "MAX(" + cellList("A", 1, 5) + ")"},
		{"MIN", "MIN(A1:A5)", "MIN(" + cellList("A", 1, 5) + ")"},
		{"PRODUCT", "PRODUCT(A1:A5)", "PRODUCT(" + cellList("A", 1, 5) + ")"},
		{"STDEV", "STDEV(A1:A5)", "STDEV(" + cellList("A", 1, 5) + ")"},
		{"SUMSQ", "SUMSQ(A1:A5)", "SUMSQ(" + cellList("A", 1, 5) + ")"},
		{"AVEDEV", "AVEDEV(A1:A5)", "AVEDEV(" + cellList("A", 1, 5) + ")"},
		{"DEVSQ", "DEVSQ(A1:A5)", "DEVSQ(" + cellList("A", 1, 5) + ")"},
		{"GEOMEAN", "GEOMEAN(A1:A5)", "GEOMEAN(" + cellList("A", 1, 5) + ")"},
	}

	for _, f := range aggFuncs {
		cases = append(cases, TestCaseDef{
			FuncNames:   []string{f.name},
			Description: fmt.Sprintf("%s range vs cell-list", f.name),
			InputCells:  numericInputs(),
			FormulaCells: []fuzz.CellSpec{
				formulaCell("G1", f.rangeF),
				formulaCell("H1", f.listF),
			},
			Checks: []fuzz.CheckSpec{
				numCheck("G1"),
				numCheck("H1"),
			},
		})
	}

	// LARGE - range form only (first arg must be a range/array)
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"LARGE"},
		Description: "LARGE range form",
		InputCells:  numericInputs(),
		FormulaCells: []fuzz.CellSpec{
			formulaCell("G1", "LARGE(A1:A5,2)"),
		},
		Checks: []fuzz.CheckSpec{numCheck("G1")},
	})

	// SMALL - range form only (first arg must be a range/array)
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"SMALL"},
		Description: "SMALL range form",
		InputCells:  numericInputs(),
		FormulaCells: []fuzz.CellSpec{
			formulaCell("G1", "SMALL(A1:A5,2)"),
		},
		Checks: []fuzz.CheckSpec{numCheck("G1")},
	})

	return cases
}
