package main

import (
	"github.com/jpoz/werkbook/internal/fuzz"
)

func logicScenarios() []TestCaseDef {
	var cases []TestCaseDef

	// IF
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"IF"},
		Description:  "IF true branch",
		InputCells:   []fuzz.CellSpec{numCell("A1", 10)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `IF(A1>5,"big","small")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"IF"},
		Description:  "IF false branch",
		InputCells:   []fuzz.CellSpec{numCell("A1", 2)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `IF(A1>5,"big","small")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// AND
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"AND"},
		Description:  "AND all true",
		InputCells:   []fuzz.CellSpec{numCell("A1", 10), numCell("A2", 20)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "AND(A1>5,A2>15)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"AND"},
		Description:  "AND one false",
		InputCells:   []fuzz.CellSpec{numCell("A1", 10), numCell("A2", 5)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "AND(A1>5,A2>15)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// OR
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"OR"},
		Description:  "OR one true",
		InputCells:   []fuzz.CellSpec{numCell("A1", 3), numCell("A2", 20)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "OR(A1>5,A2>15)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"OR"},
		Description:  "OR all false",
		InputCells:   []fuzz.CellSpec{numCell("A1", 3), numCell("A2", 5)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "OR(A1>5,A2>15)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// NOT
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"NOT"},
		Description:  "NOT negation",
		InputCells:   []fuzz.CellSpec{boolCell("A1", true)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "NOT(A1)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// XOR
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"XOR"},
		Description:  "XOR exclusive or",
		InputCells:   []fuzz.CellSpec{boolCell("A1", true), boolCell("A2", false), boolCell("A3", true)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "XOR(A1,A2,A3)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// IFERROR
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"IFERROR"},
		Description:  "IFERROR catch division by zero",
		InputCells:   []fuzz.CellSpec{numCell("A1", 10), numCell("A2", 0)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `IFERROR(A1/A2,"error")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"IFERROR"},
		Description:  "IFERROR no error",
		InputCells:   []fuzz.CellSpec{numCell("A1", 10), numCell("A2", 2)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `IFERROR(A1/A2,"error")`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// IFNA
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"IFNA"},
		Description: "IFNA catch NA error",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 1), numCell("A2", 2), numCell("A3", 3),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `IFNA(MATCH(99,A1:A3,0),"not found")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// IFS
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"IFS"},
		Description:  "IFS multiple conditions",
		InputCells:   []fuzz.CellSpec{numCell("A1", 75)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `IFS(A1>=90,"A",A1>=80,"B",A1>=70,"C",TRUE,"F")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// SWITCH
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"SWITCH"},
		Description:  "SWITCH value matching",
		InputCells:   []fuzz.CellSpec{numCell("A1", 2)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `SWITCH(A1,1,"one",2,"two",3,"three","other")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// Nested IF (3-deep)
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"IF"},
		Description:  "Nested IF 3-deep",
		InputCells:   []fuzz.CellSpec{numCell("A1", 50)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `IF(A1>100,"big",IF(A1>10,"medium",IF(A1>0,"small","negative")))`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// IF + AND + OR chain
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"IF", "AND", "OR"},
		Description:  "IF AND OR chain",
		InputCells:   []fuzz.CellSpec{numCell("A1", 10), numCell("A2", 20), numCell("A3", 30)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `IF(AND(OR(A1>5,A2>25),NOT(A3=0)),"yes","no")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	return cases
}
