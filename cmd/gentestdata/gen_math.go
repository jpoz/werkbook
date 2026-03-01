package main

import (
	"fmt"

	"github.com/jpoz/werkbook/internal/fuzz"
)

func mathScenarios() []TestCaseDef {
	var cases []TestCaseDef

	// Common input data: A1=-3.5, A2=0, A3=2.7, A4=100, A5=0.5
	commonInputs := func() []fuzz.CellSpec {
		return []fuzz.CellSpec{
			numCell("A1", -3.5),
			numCell("A2", 0),
			numCell("A3", 2.7),
			numCell("A4", 100),
			numCell("A5", 0.5),
		}
	}

	// Simple single-argument math functions
	simpleUnary := []struct {
		name    string
		formula string
	}{
		{"ABS", "ABS(A1)"},
		{"SIGN", "SIGN(A1)"},
		{"INT", "INT(A1)"},
		{"EVEN", "EVEN(A3)"},
		{"ODD", "ODD(A3)"},
		{"EXP", "EXP(A5)"},
		{"LN", "LN(A3)"},
		{"LOG10", "LOG10(A4)"},
		{"SQRT", "SQRT(A4)"},
		{"SQRTPI", "SQRTPI(A3)"},
		{"FACT", "FACT(5)"},
		{"FACTDOUBLE", "FACTDOUBLE(7)"},
	}
	for _, f := range simpleUnary {
		cases = append(cases, TestCaseDef{
			FuncNames:    []string{f.name},
			Description:  fmt.Sprintf("%s basic", f.name),
			InputCells:   commonInputs(),
			FormulaCells: []fuzz.CellSpec{formulaCell("G1", f.formula)},
			Checks:       []fuzz.CheckSpec{numCheck("G1")},
		})
	}

	// Trig functions (input in radians)
	trigInputs := []fuzz.CellSpec{
		numCell("A1", 0.5),
		numCell("A2", 1.0),
		numCell("A3", 0.785398163), // pi/4
	}

	trigFuncs := []struct {
		name    string
		formula string
	}{
		{"SIN", "SIN(A1)"},
		{"COS", "COS(A1)"},
		{"TAN", "TAN(A3)"},
		{"SINH", "SINH(A1)"},
		{"COSH", "COSH(A1)"},
		{"TANH", "TANH(A1)"},
		{"CSC", "CSC(A2)"},
		{"SEC", "SEC(A1)"},
		{"COT", "COT(A2)"},
		{"CSCH", "CSCH(A2)"},
		{"SECH", "SECH(A1)"},
		{"COTH", "COTH(A2)"},
	}
	for _, f := range trigFuncs {
		cases = append(cases, TestCaseDef{
			FuncNames:    []string{f.name},
			Description:  fmt.Sprintf("%s trig", f.name),
			InputCells:   trigInputs,
			FormulaCells: []fuzz.CellSpec{formulaCell("G1", f.formula)},
			Checks:       []fuzz.CheckSpec{numCheck("G1")},
		})
	}

	// Inverse trig functions
	invTrigInputs := []fuzz.CellSpec{
		numCell("A1", 0.5),
		numCell("A2", 1.5),
	}

	invTrigFuncs := []struct {
		name    string
		formula string
	}{
		{"ACOS", "ACOS(A1)"},
		{"ASIN", "ASIN(A1)"},
		{"ATAN", "ATAN(A1)"},
		{"ACOSH", "ACOSH(A2)"},
		{"ASINH", "ASINH(A1)"},
		{"ATANH", "ATANH(A1)"},
	}
	for _, f := range invTrigFuncs {
		cases = append(cases, TestCaseDef{
			FuncNames:    []string{f.name},
			Description:  fmt.Sprintf("%s inverse trig", f.name),
			InputCells:   invTrigInputs,
			FormulaCells: []fuzz.CellSpec{formulaCell("G1", f.formula)},
			Checks:       []fuzz.CheckSpec{numCheck("G1")},
		})
	}

	// ATAN2
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"ATAN2"},
		Description: "ATAN2 two arguments",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 3),
			numCell("A2", 4),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ATAN2(A1,A2)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// PI
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"PI"},
		Description:  "PI constant",
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "PI()")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// DEGREES / RADIANS
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"DEGREES"},
		Description: "DEGREES from radians",
		InputCells:  []fuzz.CellSpec{numCell("A1", 3.14159265358979)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "DEGREES(A1)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"RADIANS"},
		Description: "RADIANS from degrees",
		InputCells:  []fuzz.CellSpec{numCell("A1", 180)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "RADIANS(A1)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// Rounding functions
	roundInputs := []fuzz.CellSpec{
		numCell("A1", 3.14159),
		numCell("A2", -2.567),
		numCell("A3", 7.5),
	}
	roundFuncs := []struct {
		name    string
		formula string
	}{
		{"ROUND", "ROUND(A1,2)"},
		{"ROUNDDOWN", "ROUNDDOWN(A1,2)"},
		{"ROUNDUP", "ROUNDUP(A2,1)"},
		{"TRUNC", "TRUNC(A1,2)"},
		{"CEILING", "CEILING(A1,0.1)"},
		{"FLOOR", "FLOOR(A1,0.1)"},
		{"MROUND", "MROUND(A3,3)"},
	}
	for _, f := range roundFuncs {
		cases = append(cases, TestCaseDef{
			FuncNames:    []string{f.name},
			Description:  fmt.Sprintf("%s rounding", f.name),
			InputCells:   roundInputs,
			FormulaCells: []fuzz.CellSpec{formulaCell("G1", f.formula)},
			Checks:       []fuzz.CheckSpec{numCheck("G1")},
		})
	}

	// LOG with base
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"LOG"},
		Description: "LOG with custom base",
		InputCells:  []fuzz.CellSpec{numCell("A1", 64), numCell("A2", 2)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "LOG(A1,A2)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// MOD
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"MOD"},
		Description: "MOD remainder",
		InputCells:  []fuzz.CellSpec{numCell("A1", 17), numCell("A2", 5)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "MOD(A1,A2)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// POWER
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"POWER"},
		Description: "POWER exponentiation",
		InputCells:  []fuzz.CellSpec{numCell("A1", 2), numCell("A2", 10)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "POWER(A1,A2)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// QUOTIENT
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"QUOTIENT"},
		Description: "QUOTIENT integer division",
		InputCells:  []fuzz.CellSpec{numCell("A1", 17), numCell("A2", 5)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "QUOTIENT(A1,A2)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// GCD
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"GCD"},
		Description: "GCD of multiple numbers",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 24), numCell("A2", 36), numCell("A3", 48),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "GCD(A1,A2,A3)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// LCM
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"LCM"},
		Description: "LCM of multiple numbers",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 4), numCell("A2", 6), numCell("A3", 10),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "LCM(A1,A2,A3)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// COMBINA
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"COMBINA"},
		Description: "COMBINA combinations with repetition",
		InputCells:  []fuzz.CellSpec{numCell("A1", 10), numCell("A2", 3)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "COMBINA(A1,A2)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// MULTINOMIAL
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"MULTINOMIAL"},
		Description: "MULTINOMIAL",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 2), numCell("A2", 3), numCell("A3", 4),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "MULTINOMIAL(A1,A2,A3)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// DECIMAL
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"DECIMAL"},
		Description:  "DECIMAL from hex",
		InputCells:   []fuzz.CellSpec{strCell("A1", "FF")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", `DECIMAL(A1,16)`)},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// BASE
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"BASE"},
		Description:  "BASE to binary",
		InputCells:   []fuzz.CellSpec{numCell("A1", 255)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "BASE(A1,2)")},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"BASE"},
		Description:  "BASE to hex with padding",
		InputCells:   []fuzz.CellSpec{numCell("A1", 255)},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "BASE(A1,16,4)")},
		Checks:       []fuzz.CheckSpec{strCheck("G1")},
	})

	// SUBTOTAL (function_num 9 = SUM, 1 = AVERAGE)
	cases = append(cases, TestCaseDef{
		FuncNames:   []string{"SUBTOTAL"},
		Description: "SUBTOTAL SUM mode",
		InputCells: []fuzz.CellSpec{
			numCell("A1", 10), numCell("A2", 20), numCell("A3", 30),
		},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "SUBTOTAL(9,A1:A3)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	// ARABIC
	cases = append(cases, TestCaseDef{
		FuncNames:    []string{"ARABIC"},
		Description:  "ARABIC roman numeral",
		InputCells:   []fuzz.CellSpec{strCell("A1", "MCMXCIX")},
		FormulaCells: []fuzz.CellSpec{formulaCell("G1", "ARABIC(A1)")},
		Checks:       []fuzz.CheckSpec{numCheck("G1")},
	})

	return cases
}
