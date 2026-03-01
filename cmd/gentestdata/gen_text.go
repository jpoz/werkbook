package main

import (
	"fmt"

	"github.com/jpoz/werkbook/internal/fuzz"
)

func textScenarios() []TestCaseDef {
	var cases []TestCaseDef

	textInputs := func() []fuzz.CellSpec {
		return []fuzz.CellSpec{
			strCell("A1", "  Hello World  "),
			strCell("A2", "abcdef"),
			strCell("A3", "Excel Formulas"),
			numCell("A4", 1234.567),
			strCell("A5", "hello"),
		}
	}

	// Simple text functions
	simpleFuncs := []struct {
		name    string
		formula string
		typ     string
	}{
		{"UPPER", "UPPER(A5)", "string"},
		{"LOWER", "LOWER(A3)", "string"},
		{"PROPER", "PROPER(A5)", "string"},
		{"TRIM", "TRIM(A1)", "string"},
		{"LEN", "LEN(A2)", "number"},
		{"LEFT", "LEFT(A2,3)", "string"},
		{"RIGHT", "RIGHT(A2,3)", "string"},
		{"MID", "MID(A2,2,3)", "string"},
		{"CLEAN", "CLEAN(A1)", "string"},
	}
	for _, f := range simpleFuncs {
		check := fuzz.CheckSpec{Ref: "G1", Type: f.typ}
		cases = append(cases, TestCaseDef{
			FuncNames:    []string{f.name},
			Description:  fmt.Sprintf("%s basic", f.name),
			InputCells:   textInputs(),
			FormulaCells: []fuzz.CellSpec{formulaCell("G1", f.formula)},
			Checks:       []fuzz.CheckSpec{check},
		})
	}

	// CHAR / CODE
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"CHAR"},
		Description:  "CHAR from code",
		InputCells:   []fuzz.CellSpec{numCell("A1", 65)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "CHAR(A1)")},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"CODE"},
		Description:  "CODE of character",
		InputCells:   []fuzz.CellSpec{strCell("A1", "A")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "CODE(A1)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// CONCATENATE
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"CONCATENATE"},
		Description: "CONCATENATE strings",
		InputCells: []fuzz.CellSpec{
			strCell("A1", "Hello"),
			strCell("A2", " "),
			strCell("A3", "World"),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "CONCATENATE(A1,A2,A3)")},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// CONCAT
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"CONCAT"},
		Description: "CONCAT strings",
		InputCells: []fuzz.CellSpec{
			strCell("A1", "Hello"),
			strCell("A2", "World"),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `CONCAT(A1," ",A2)`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// EXACT
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"EXACT"},
		Description: "EXACT comparison",
		InputCells: []fuzz.CellSpec{
			strCell("A1", "hello"),
			strCell("A2", "Hello"),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "EXACT(A1,A2)")},
		Checks:       []fuzz.CheckSpec{boolCheck("G1")},
	})

	// FIND
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"FIND"},
		Description:  "FIND substring position",
		InputCells:   []fuzz.CellSpec{strCell("A1", "Hello World")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `FIND("World",A1)`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// SEARCH (case-insensitive)
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"SEARCH"},
		Description:  "SEARCH case-insensitive",
		InputCells:   []fuzz.CellSpec{strCell("A1", "Hello World")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `SEARCH("world",A1)`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// REPLACE
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"REPLACE"},
		Description:  "REPLACE characters",
		InputCells:   []fuzz.CellSpec{strCell("A1", "abcdef")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `REPLACE(A1,3,2,"XY")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// SUBSTITUTE
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"SUBSTITUTE"},
		Description:  "SUBSTITUTE text",
		InputCells:   []fuzz.CellSpec{strCell("A1", "Hello World Hello")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `SUBSTITUTE(A1,"Hello","Hi")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// REPT
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"REPT"},
		Description:  "REPT repeat text",
		InputCells:   []fuzz.CellSpec{strCell("A1", "ab")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "REPT(A1,3)")},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// TEXT
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"TEXT"},
		Description:  "TEXT format number",
		InputCells:   []fuzz.CellSpec{numCell("A1", 1234.567)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `TEXT(A1,"0.00")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// VALUE
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"VALUE"},
		Description:  "VALUE string to number",
		InputCells:   []fuzz.CellSpec{strCell("A1", "123.45")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "VALUE(A1)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// T
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"T"},
		Description:  "T text extraction",
		InputCells:   []fuzz.CellSpec{strCell("A1", "hello"), numCell("A2", 42)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "T(A1)")},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// FIXED
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"FIXED"},
		Description:  "FIXED format number",
		InputCells:   []fuzz.CellSpec{numCell("A1", 1234567.89)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "FIXED(A1,2)")},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// NUMBERVALUE
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"NUMBERVALUE"},
		Description:  "NUMBERVALUE parse formatted number",
		InputCells:   []fuzz.CellSpec{strCell("A1", "1.234,56")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `NUMBERVALUE(A1,".",",")`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// TEXTJOIN
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"TEXTJOIN"},
		Description: "TEXTJOIN with delimiter",
		InputCells: []fuzz.CellSpec{
			strCell("A1", "Hello"), strCell("A2", "World"), strCell("A3", "Test"),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `TEXTJOIN(", ",TRUE,A1:A3)`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	return cases
}
