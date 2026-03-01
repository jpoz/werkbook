package main

import (
	"github.com/jpoz/werkbook/internal/fuzz"
)

func crosssheetScenarios() ScenarioSet {
	dataSheet := SheetDef{
		Name: "Data",
		Cases: []TestCaseDef{
			{
				Description: "Data sheet values",
				InputCells: []fuzz.CellSpec{
					// Numeric data in A column
					numCell("A1", 75), numCell("A2", 150), numCell("A3", 200),
					numCell("A4", 50), numCell("A5", 300),
					numCell("A6", 10), numCell("A7", 20), numCell("A8", 30),
					numCell("A9", 40), numCell("A10", 50),
					// Secondary numeric data in B column
					numCell("B1", 5), numCell("B2", 15), numCell("B3", 25),
					numCell("B4", 35), numCell("B5", 45),
					numCell("B6", 55), numCell("B7", 65), numCell("B8", 75),
					numCell("B9", 85), numCell("B10", 95),
					// Text data in C column
					strCell("C1", "Hello World"),
				},
			},
		},
	}

	lookupSheet := SheetDef{
		Name: "Lookup",
		Cases: []TestCaseDef{
			{
				Description: "Lookup table",
				InputCells: []fuzz.CellSpec{
					numCell("A1", 75), strCell("B1", "Alpha"), numCell("C1", 1000),
					numCell("A2", 150), strCell("B2", "Beta"), numCell("C2", 2000),
					numCell("A3", 200), strCell("B3", "Gamma"), numCell("C3", 3000),
					numCell("A4", 50), strCell("B4", "Delta"), numCell("C4", 4000),
					numCell("A5", 300), strCell("B5", "Epsilon"), numCell("C5", 5000),
				},
			},
		},
	}

	formulaSheet := SheetDef{
		Name: "Formulas",
		Cases: []TestCaseDef{
			// 1. Simple cross-sheet cell ref
			{
				FuncNames:    []string{},
				Description:  "Simple cross-sheet multiply",
				FormulaCells: []fuzz.CellSpec{formulaCell("G1", "Data!A1*2")},
				Checks:       []fuzz.CheckSpec{numCheck("G1")},
			},
			// 2. Cross-sheet range aggregation
			{
				FuncNames:    []string{"SUM"},
				Description:  "Cross-sheet SUM",
				FormulaCells: []fuzz.CellSpec{formulaCell("G1", "SUM(Data!A1:Data!A5)")},
				Checks:       []fuzz.CheckSpec{numCheck("G1")},
			},
			// 3. Cross-sheet average
			{
				FuncNames:    []string{"AVERAGE"},
				Description:  "Cross-sheet AVERAGE",
				FormulaCells: []fuzz.CellSpec{formulaCell("G1", "AVERAGE(Data!A1:Data!A5)")},
				Checks:       []fuzz.CheckSpec{numCheck("G1")},
			},
			// 4. Cross-sheet VLOOKUP
			{
				FuncNames:    []string{"VLOOKUP"},
				Description:  "Cross-sheet VLOOKUP",
				FormulaCells: []fuzz.CellSpec{formulaCell("G1", "VLOOKUP(200,Lookup!A1:B5,2,FALSE)")},
				Checks:       []fuzz.CheckSpec{strCheck("G1")},
			},
			// 5. Cross-sheet conditional
			{
				FuncNames:    []string{"IF"},
				Description:  "Cross-sheet IF",
				FormulaCells: []fuzz.CellSpec{formulaCell("G1", `IF(Data!A1>50,Data!B1,Data!B2)`)},
				Checks:       []fuzz.CheckSpec{numCheck("G1")},
			},
			// 6. Cross-sheet COUNTIF
			{
				FuncNames:    []string{"COUNTIF"},
				Description:  "Cross-sheet COUNTIF",
				FormulaCells: []fuzz.CellSpec{formulaCell("G1", `COUNTIF(Data!A1:Data!A10,">0")`)},
				Checks:       []fuzz.CheckSpec{numCheck("G1")},
			},
			// 7. Cross-sheet range in expression
			{
				FuncNames:    []string{"MAX", "MIN"},
				Description:  "Cross-sheet MAX-MIN",
				FormulaCells: []fuzz.CellSpec{formulaCell("G1", "MAX(Data!A1:Data!A5)-MIN(Data!A1:Data!A5)")},
				Checks:       []fuzz.CheckSpec{numCheck("G1")},
			},
			// 8. Cross-sheet INDEX/MATCH
			{
				FuncNames:    []string{"INDEX", "MATCH"},
				Description:  "Cross-sheet INDEX/MATCH",
				FormulaCells: []fuzz.CellSpec{formulaCell("G1", "INDEX(Lookup!A1:C5,MATCH(Data!A1,Lookup!A1:A5,0),3)")},
				Checks:       []fuzz.CheckSpec{numCheck("G1")},
			},
			// 9. Cross-sheet LEN
			{
				FuncNames:    []string{"LEN"},
				Description:  "Cross-sheet text operation",
				FormulaCells: []fuzz.CellSpec{formulaCell("G1", "LEN(Data!C1)")},
				Checks:       []fuzz.CheckSpec{numCheck("G1")},
			},
		},
	}

	return ScenarioSet{
		Name:   "crosssheet",
		Sheets: []SheetDef{dataSheet, lookupSheet, formulaSheet},
	}
}
