package main

import (
	"github.com/jpoz/werkbook/internal/fuzz"
)

func infoScenarios() []TestCaseDef {
	var cases []TestCaseDef

	// ISBLANK
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ISBLANK"},
		Description:  "ISBLANK empty cell",
		InputCells:   []fuzz.CellSpec{numCell("A1", 1)}, // A2 is empty
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISBLANK(A2)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ISBLANK"},
		Description:  "ISBLANK non-empty cell",
		InputCells:   []fuzz.CellSpec{numCell("A1", 1)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISBLANK(A1)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// ISNUMBER
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ISNUMBER"},
		Description:  "ISNUMBER with number",
		InputCells:   []fuzz.CellSpec{numCell("A1", 42)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISNUMBER(A1)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ISNUMBER"},
		Description:  "ISNUMBER with string",
		InputCells:   []fuzz.CellSpec{strCell("A1", "hello")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISNUMBER(A1)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// ISTEXT
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ISTEXT"},
		Description:  "ISTEXT with string",
		InputCells:   []fuzz.CellSpec{strCell("A1", "hello")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISTEXT(A1)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ISTEXT"},
		Description:  "ISTEXT with number",
		InputCells:   []fuzz.CellSpec{numCell("A1", 42)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISTEXT(A1)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// ISLOGICAL
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ISLOGICAL"},
		Description:  "ISLOGICAL with bool",
		InputCells:   []fuzz.CellSpec{boolCell("A1", true)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISLOGICAL(A1)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// ISNONTEXT
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ISNONTEXT"},
		Description:  "ISNONTEXT with number",
		InputCells:   []fuzz.CellSpec{numCell("A1", 42)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISNONTEXT(A1)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// ISERROR
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ISERROR"},
		Description:  "ISERROR with error",
		InputCells:   []fuzz.CellSpec{numCell("A1", 1), numCell("A2", 0)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISERROR(A1/A2)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ISERROR"},
		Description:  "ISERROR without error",
		InputCells:   []fuzz.CellSpec{numCell("A1", 10), numCell("A2", 2)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISERROR(A1/A2)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// ISERR (like ISERROR but not for #N/A)
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ISERR"},
		Description:  "ISERR with div/0 error",
		InputCells:   []fuzz.CellSpec{numCell("A1", 1), numCell("A2", 0)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISERR(A1/A2)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// ISNA
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"ISNA"},
		Description: "ISNA with NA",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 1), numCell("A2", 2), numCell("A3", 3),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISNA(MATCH(99,A1:A3,0))")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// ISEVEN / ISODD
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ISEVEN"},
		Description:  "ISEVEN with even number",
		InputCells:   []fuzz.CellSpec{numCell("A1", 4)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISEVEN(A1)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ISODD"},
		Description:  "ISODD with odd number",
		InputCells:   []fuzz.CellSpec{numCell("A1", 7)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ISODD(A1)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// N
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"N"},
		Description:  "N convert to number",
		InputCells:   []fuzz.CellSpec{boolCell("A1", true)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "N(A1)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// NA
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"NA"},
		Description:  "NA produces #N/A error",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "NA()")},
		Checks:       []fuzz.CheckSpec{errCheck("G1")},
	})

	// TYPE
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"TYPE"},
		Description:  "TYPE of number",
		InputCells:   []fuzz.CellSpec{numCell("A1", 42)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "TYPE(A1)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"TYPE"},
		Description:  "TYPE of string",
		InputCells:   []fuzz.CellSpec{strCell("A1", "hello")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "TYPE(A1)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// ERROR.TYPE
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ERROR.TYPE"},
		Description:  "ERROR.TYPE of div/0",
		InputCells:   []fuzz.CellSpec{numCell("A1", 1), numCell("A2", 0)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ERROR.TYPE(A1/A2)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	return cases
}
