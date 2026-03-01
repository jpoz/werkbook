package main

import (
	"fmt"

	"github.com/jpoz/werkbook/internal/fuzz"
)

func statScenarios() []TestCaseDef {
	var cases []TestCaseDef

	// Common numeric input: A1:A5 = 10, 20, 30, 40, 50
	statInputs := func() []fuzz.CellSpec {
		return []fuzz.CellSpec{
			numCell("A1", 10), numCell("A2", 20), numCell("A3", 30),
			numCell("A4", 40), numCell("A5", 50),
		}
	}

	// Basic aggregation functions with range
	aggFuncs := []struct {
		name    string
		formula string
	}{
		{"SUM", "SUM(A1:A5)"},
		{"AVERAGE", "AVERAGE(A1:A5)"},
		{"COUNT", "COUNT(A1:A5)"},
		{"MAX", "MAX(A1:A5)"},
		{"MIN", "MIN(A1:A5)"},
		{"MEDIAN", "MEDIAN(A1:A5)"},
		{"PRODUCT", "PRODUCT(A1:A5)"},
		{"STDEV", "STDEV(A1:A5)"},
		{"STDEVP", "STDEVP(A1:A5)"},
		{"VAR", "VAR(A1:A5)"},
		{"VARP", "VARP(A1:A5)"},
		{"SUMSQ", "SUMSQ(A1:A5)"},
		{"AVEDEV", "AVEDEV(A1:A5)"},
		{"DEVSQ", "DEVSQ(A1:A5)"},
		{"GEOMEAN", "GEOMEAN(A1:A5)"},
	}
	for _, f := range aggFuncs {
		cases = append(cases, TestCaseDef{
			FuncNames:    []string{f.name},
			Description:  fmt.Sprintf("%s basic range", f.name),
			InputCells:   statInputs(),
			FormulaCells: []fuzz.CellSpec{formulaCell("G1", f.formula)},
			Checks:       []fuzz.CheckSpec{numCheck("G1")},
		})
	}

	// COUNTA / COUNTBLANK with mixed data
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"COUNTA"},
		Description: "COUNTA with mixed types",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 10), strCell("A2", "hello"), boolCell("A3", true),
			numCell("A4", 0), numCell("A5", 5),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "COUNTA(A1:A5)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"COUNTBLANK"},
		Description: "COUNTBLANK",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 10),
			// A2 empty
			numCell("A3", 30),
			// A4 empty
			numCell("A5", 50),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "COUNTBLANK(A1:A5)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// MODE
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"MODE"},
		Description: "MODE most frequent",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 3), numCell("A2", 5), numCell("A3", 3),
			numCell("A4", 7), numCell("A5", 3),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "MODE(A1:A5)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// LARGE / SMALL
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"LARGE"},
		Description: "LARGE 2nd largest",
		InputCells:  statInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "LARGE(A1:A5,2)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"SMALL"},
		Description: "SMALL 2nd smallest",
		InputCells:  statInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "SMALL(A1:A5,2)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// PERCENTILE
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"PERCENTILE"},
		Description: "PERCENTILE 75th",
		InputCells:  statInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "PERCENTILE(A1:A5,0.75)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// PERMUT
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"PERMUT"},
		Description:  "PERMUT permutations",
		InputCells:   []fuzz.CellSpec{numCell("A1", 10), numCell("A2", 3)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "PERMUT(A1,A2)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// RANK
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"RANK"},
		Description: "RANK descending",
		InputCells:  statInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "RANK(A3,A1:A5)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// SUMPRODUCT
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"SUMPRODUCT"},
		Description: "SUMPRODUCT two arrays",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 1), numCell("A2", 2), numCell("A3", 3),
			numCell("B1", 4), numCell("B2", 5), numCell("B3", 6),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "SUMPRODUCT(A1:A3,B1:B3)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// SUMIF
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"SUMIF"},
		Description: "SUMIF greater than",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 10), numCell("A2", 20), numCell("A3", 30),
			numCell("A4", 40), numCell("A5", 50),
			numCell("B1", 1), numCell("B2", 2), numCell("B3", 3),
			numCell("B4", 4), numCell("B5", 5),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `SUMIF(A1:A5,">25",B1:B5)`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// SUMIFS
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"SUMIFS"},
		Description: "SUMIFS multiple criteria",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 100), numCell("A2", 200), numCell("A3", 300),
			numCell("A4", 400), numCell("A5", 500),
			numCell("B1", 1), numCell("B2", 2), numCell("B3", 1),
			numCell("B4", 2), numCell("B5", 1),
			numCell("C1", 10), numCell("C2", 20), numCell("C3", 30),
			numCell("C4", 40), numCell("C5", 50),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `SUMIFS(C1:C5,A1:A5,">200",B1:B5,1)`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// COUNTIF
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"COUNTIF"},
		Description: "COUNTIF greater than",
		InputCells:  statInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `COUNTIF(A1:A5,">25")`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// COUNTIFS
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"COUNTIFS"},
		Description: "COUNTIFS multiple criteria",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 10), numCell("A2", 20), numCell("A3", 30),
			numCell("A4", 40), numCell("A5", 50),
			numCell("B1", 1), numCell("B2", 2), numCell("B3", 1),
			numCell("B4", 2), numCell("B5", 1),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `COUNTIFS(A1:A5,">15",B1:B5,1)`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// AVERAGEIF
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"AVERAGEIF"},
		Description: "AVERAGEIF greater than",
		InputCells:  statInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `AVERAGEIF(A1:A5,">20")`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// AVERAGEIFS
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"AVERAGEIFS"},
		Description: "AVERAGEIFS multiple criteria",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 10), numCell("A2", 20), numCell("A3", 30),
			numCell("A4", 40), numCell("A5", 50),
			numCell("B1", 1), numCell("B2", 2), numCell("B3", 1),
			numCell("B4", 2), numCell("B5", 1),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `AVERAGEIFS(A1:A5,A1:A5,">15",B1:B5,1)`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// MAXIFS
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"MAXIFS"},
		Description: "MAXIFS conditional max",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 10), numCell("A2", 20), numCell("A3", 30),
			numCell("A4", 40), numCell("A5", 50),
			numCell("B1", 1), numCell("B2", 2), numCell("B3", 1),
			numCell("B4", 2), numCell("B5", 1),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `MAXIFS(A1:A5,B1:B5,1)`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// MINIFS
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"MINIFS"},
		Description: "MINIFS conditional min",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 10), numCell("A2", 20), numCell("A3", 30),
			numCell("A4", 40), numCell("A5", 50),
			numCell("B1", 1), numCell("B2", 2), numCell("B3", 1),
			numCell("B4", 2), numCell("B5", 1),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `MINIFS(A1:A5,B1:B5,2)`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	return cases
}
