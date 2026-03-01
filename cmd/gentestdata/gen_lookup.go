package main

import (
	"github.com/jpoz/werkbook/internal/fuzz"
)

func lookupScenarios() []TestCaseDef {
	var cases []TestCaseDef

	// VLOOKUP
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"VLOOKUP"},
		Description: "VLOOKUP exact match",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 1), strCell("B1", "Apple"),
			numCell("A2", 2), strCell("B2", "Banana"),
			numCell("A3", 3), strCell("B3", "Cherry"),
			numCell("A4", 4), strCell("B4", "Date"),
			numCell("A5", 5), strCell("B5", "Elderberry"),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "VLOOKUP(3,A1:B5,2,FALSE)")},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// HLOOKUP
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"HLOOKUP"},
		Description: "HLOOKUP exact match",
		InputCells: []fuzz.CellSpec{
			strCell("A1", "Name"), strCell("B1", "Age"), strCell("C1", "Score"),
			strCell("A2", "Alice"), numCell("B2", 30), numCell("C2", 95),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `HLOOKUP("Age",A1:C2,2,FALSE)`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// INDEX
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"INDEX"},
		Description: "INDEX array form",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 10), numCell("B1", 20), numCell("C1", 30),
			numCell("A2", 40), numCell("B2", 50), numCell("C2", 60),
			numCell("A3", 70), numCell("B3", 80), numCell("C3", 90),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "INDEX(A1:C3,2,3)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// MATCH
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"MATCH"},
		Description: "MATCH exact",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 10), numCell("A2", 20), numCell("A3", 30),
			numCell("A4", 40), numCell("A5", 50),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "MATCH(30,A1:A5,0)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// LOOKUP
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"LOOKUP"},
		Description: "LOOKUP vector form",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 10), numCell("A2", 20), numCell("A3", 30),
			numCell("A4", 40), numCell("A5", 50),
			strCell("B1", "ten"), strCell("B2", "twenty"), strCell("B3", "thirty"),
			strCell("B4", "forty"), strCell("B5", "fifty"),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "LOOKUP(30,A1:A5,B1:B5)")},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// XLOOKUP
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"XLOOKUP"},
		Description: "XLOOKUP exact match",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 1), strCell("B1", "Apple"),
			numCell("A2", 2), strCell("B2", "Banana"),
			numCell("A3", 3), strCell("B3", "Cherry"),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `XLOOKUP(2,A1:A3,B1:B3,"Not found")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// CHOOSE
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"CHOOSE"},
		Description:  "CHOOSE from index",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `CHOOSE(2,"apple","banana","cherry")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// ROW / COLUMN
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ROW"},
		Description:  "ROW of reference",
		InputCells:   []fuzz.CellSpec{numCell("A1", 1)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ROW(A3)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"COLUMN"},
		Description:  "COLUMN of reference",
		InputCells:   []fuzz.CellSpec{numCell("A1", 1)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "COLUMN(C1)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// ROWS / COLUMNS
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"ROWS"},
		Description: "ROWS count",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 1), numCell("A2", 2), numCell("A3", 3),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ROWS(A1:A5)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"COLUMNS"},
		Description: "COLUMNS count",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 1), numCell("B1", 2), numCell("C1", 3),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "COLUMNS(A1:C1)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// SORT (wrapped in INDEX to get scalar result)
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"SORT"},
		Description: "SORT ascending via INDEX",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 30), numCell("A2", 10), numCell("A3", 50),
			numCell("A4", 20), numCell("A5", 40),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "INDEX(SORT(A1:A5),1)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// ADDRESS
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ADDRESS"},
		Description:  "ADDRESS cell reference",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ADDRESS(3,2)")},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	return cases
}
