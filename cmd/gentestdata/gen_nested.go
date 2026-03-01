package main

import (
	"github.com/jpoz/werkbook/internal/fuzz"
)

func nestedScenarios() []TestCaseDef {
	var cases []TestCaseDef

	// Common inputs: A1:A5 = 15,25,35,45,55 ; B1:B5 = 2,4,6,8,10 ; C1:C5 = 100,200,300,400,500
	commonInputs := func() []fuzz.CellSpec {
		return []fuzz.CellSpec{
			numCell("A1", 15), numCell("A2", 25), numCell("A3", 35),
			numCell("A4", 45), numCell("A5", 55),
			numCell("B1", 2), numCell("B2", 4), numCell("B3", 6),
			numCell("B4", 8), numCell("B5", 10),
			numCell("C1", 100), numCell("C2", 200), numCell("C3", 300),
			numCell("C4", 400), numCell("C5", 500),
		}
	}

	// Extended inputs: A1:A10 for functions needing 10 values
	extendedInputs := func() []fuzz.CellSpec {
		return []fuzz.CellSpec{
			numCell("A1", 15), numCell("A2", 25), numCell("A3", 35),
			numCell("A4", 45), numCell("A5", 55), numCell("A6", 65),
			numCell("A7", 75), numCell("A8", 85), numCell("A9", 95),
			numCell("A10", 105),
			numCell("B1", 2), numCell("B2", 4), numCell("B3", 6),
			numCell("B4", 8), numCell("B5", 10),
			numCell("C1", 100), numCell("C2", 200), numCell("C3", 300),
			numCell("C4", 400), numCell("C5", 500),
		}
	}

	// 1. ROUND(AVERAGE(A1:A5), 2)
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ROUND", "AVERAGE"},
		Description:  "Aggregation in rounding",
		InputCells:   commonInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ROUND(AVERAGE(A1:A5),2)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// 2. IF(AND(A1>0, B1>0), SUM(C1:C5), 0)
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"IF", "AND", "SUM"},
		Description:  "Logic + aggregation",
		InputCells:   commonInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "IF(AND(A1>0,B1>0),SUM(C1:C5),0)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// 3. IF(OR(A1>100, A2>100), MAX(A1:A5), MIN(A1:A5))
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"IF", "OR", "MAX", "MIN"},
		Description:  "Conditional aggregate selection",
		InputCells:   commonInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "IF(OR(A1>100,A2>100),MAX(A1:A5),MIN(A1:A5))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// 4. IFERROR(VLOOKUP(A1, B1:C5, 2, FALSE), "not found") - use matching lookup data
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"IFERROR", "VLOOKUP"},
		Description: "Error handling + lookup",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 99), // lookup value not in B column
			numCell("B1", 1), numCell("C1", 100),
			numCell("B2", 2), numCell("C2", 200),
			numCell("B3", 3), numCell("C3", 300),
			numCell("B4", 4), numCell("C4", 400),
			numCell("B5", 5), numCell("C5", 500),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `IFERROR(VLOOKUP(A1,B1:C5,2,FALSE),"not found")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// 5. SUM(IF(A1>0, A1, 0), IF(A2>0, A2, 0), IF(A3>0, A3, 0))
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"SUM", "IF"},
		Description:  "Nested IF as SUM args",
		InputCells:   commonInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "SUM(IF(A1>0,A1,0),IF(A2>0,A2,0),IF(A3>0,A3,0))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// 6. CONCATENATE(TEXT(A1, "0.00"), " - ", TEXT(A2, "0.00"))
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"CONCATENATE", "TEXT"},
		Description:  "Text formatting chain",
		InputCells:   commonInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `CONCATENATE(TEXT(A1,"0.00")," - ",TEXT(A2,"0.00"))`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// 7. ROUND(STDEV(A1:A5)/AVERAGE(A1:A5)*100, 2) - coefficient of variation
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ROUND", "STDEV", "AVERAGE"},
		Description:  "Coefficient of variation",
		InputCells:   commonInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ROUND(STDEV(A1:A5)/AVERAGE(A1:A5)*100,2)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// 8. SUMPRODUCT(A1:A5,B1:B5)
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"SUMPRODUCT"},
		Description:  "SUMPRODUCT two arrays",
		InputCells:   commonInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "SUMPRODUCT(A1:A5,B1:B5)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// 9. IF(COUNTIF(A1:A10, ">50")>3, "many", "few")
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"IF", "COUNTIF"},
		Description:  "Conditional count check",
		InputCells:   extendedInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `IF(COUNTIF(A1:A10,">50")>3,"many","few")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// 10. INDEX(A1:C5, MATCH(MAX(B1:B5), B1:B5, 0), 3) - INDEX/MATCH pattern
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"INDEX", "MATCH", "MAX"},
		Description:  "INDEX/MATCH pattern",
		InputCells:   commonInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "INDEX(A1:C5,MATCH(MAX(B1:B5),B1:B5,0),3)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// 11. LEFT(UPPER(TRIM(A1)), 5)
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"LEFT", "UPPER", "TRIM"},
		Description:  "Chained text transforms",
		InputCells:   []fuzz.CellSpec{strCell("A1", "  hello world  ")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "LEFT(UPPER(TRIM(A1)),5)")},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// 12. YEAR(DATE(2024,1,1))+MONTH(DATE(2024,6,15))
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"YEAR", "DATE", "MONTH"},
		Description:  "Date construction + extraction",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "YEAR(DATE(2024,1,1))+MONTH(DATE(2024,6,15))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// 13. IF(AND(ISNUMBER(A1), A1>0), SQRT(A1), NA())
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"IF", "AND", "ISNUMBER", "SQRT", "NA"},
		Description:  "Info + math + error",
		InputCells:   []fuzz.CellSpec{numCell("A1", 25)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "IF(AND(ISNUMBER(A1),A1>0),SQRT(A1),NA())")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// 14. AVERAGE(LARGE(A1:A10,1), LARGE(A1:A10,2), LARGE(A1:A10,3))
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"AVERAGE", "LARGE"},
		Description:  "Top-3 average",
		InputCells:   extendedInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "AVERAGE(LARGE(A1:A10,1),LARGE(A1:A10,2),LARGE(A1:A10,3))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// 15. IF(A1>0, IF(A1>10, IF(A1>100, "big", "medium"), "small"), "negative")
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"IF"},
		Description:  "3-deep nested IF",
		InputCells:   []fuzz.CellSpec{numCell("A1", 50)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `IF(A1>0,IF(A1>10,IF(A1>100,"big","medium"),"small"),"negative")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// 16. SUM(A1:A5) + SUM(B1:B5) * SUM(C1:C5)
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"SUM"},
		Description:  "Multiple aggregations in arithmetic",
		InputCells:   commonInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "SUM(A1:A5)+SUM(B1:B5)*SUM(C1:C5)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// 17. ROUND(SUMPRODUCT(A1:A5, B1:B5) / SUM(B1:B5), 4) - weighted average
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ROUND", "SUMPRODUCT", "SUM"},
		Description:  "Weighted average pattern",
		InputCells:   commonInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ROUND(SUMPRODUCT(A1:A5,B1:B5)/SUM(B1:B5),4)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// 18. IF(AND(OR(A1>0, A2>0), NOT(A3=0)), "yes", "no")
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"IF", "AND", "OR", "NOT"},
		Description:  "Mixed logic operators",
		InputCells:   commonInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `IF(AND(OR(A1>0,A2>0),NOT(A3=0)),"yes","no")`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// 19. MID(TEXT(A1*1000, "000000"), 2, 3) - numeric formatting + text extraction
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"MID", "TEXT"},
		Description:  "Numeric formatting + text extraction",
		InputCells:   []fuzz.CellSpec{numCell("A1", 15)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `MID(TEXT(A1*1000,"000000"),2,3)`)},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// 20. MAX(ABS(A1-AVERAGE(A1:A5)), ABS(A2-AVERAGE(A1:A5)))
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"MAX", "ABS", "AVERAGE"},
		Description:  "Deviation comparison",
		InputCells:   commonInputs(),
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "MAX(ABS(A1-AVERAGE(A1:A5)),ABS(A2-AVERAGE(A1:A5)))")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	return cases
}
