package fuzz

import (
	"os"
	"sort"
	"strings"
)

// isolationTests maps function names to simple test cases that should always work.
// Each entry is a formula and its expected result.
var isolationTests = map[string]struct {
	formula  string
	expected string
}{
	"ABS":           {formula: "ABS(-5)", expected: "5"},
	"ACOS":          {formula: "ACOS(1)", expected: "0"},
	"AND":           {formula: "AND(TRUE,TRUE)", expected: "TRUE"},
	"ASIN":          {formula: "ASIN(0)", expected: "0"},
	"ATAN":          {formula: "ATAN(0)", expected: "0"},
	"AVERAGE":       {formula: "AVERAGE(1,2,3)", expected: "2"},
	"CEILING":       {formula: "CEILING(4.2,1)", expected: "5"},
	"CHAR":          {formula: "CHAR(65)", expected: "A"},
	"CHOOSE":        {formula: "CHOOSE(2,\"a\",\"b\",\"c\")", expected: "b"},
	"CLEAN":         {formula: "CLEAN(\"hello\")", expected: "hello"},
	"CODE":          {formula: "CODE(\"A\")", expected: "65"},
	"CONCATENATE":   {formula: "CONCATENATE(\"a\",\"b\")", expected: "ab"},
	"COS":           {formula: "COS(0)", expected: "1"},
	"COUNT":         {formula: "COUNT(1,2,3)", expected: "3"},
	"COUNTA":        {formula: "COUNTA(1,\"a\",TRUE)", expected: "3"},
	"DATE":          {formula: "DATE(2024,1,15)", expected: "45306"},
	"DAY":           {formula: "DAY(45306)", expected: "15"},
	"EXACT":         {formula: "EXACT(\"a\",\"a\")", expected: "TRUE"},
	"EXP":           {formula: "EXP(0)", expected: "1"},
	"FIND":          {formula: "FIND(\"b\",\"abc\")", expected: "2"},
	"FLOOR":         {formula: "FLOOR(4.8,1)", expected: "4"},
	"HOUR":          {formula: "HOUR(0.5)", expected: "12"},
	"IF":            {formula: "IF(TRUE,1,0)", expected: "1"},
	"IFERROR":       {formula: "IFERROR(1/0,\"err\")", expected: "err"},
	"INT":           {formula: "INT(3.7)", expected: "3"},
	"ISBLANK":       {formula: "ISBLANK(\"\")", expected: "FALSE"},
	"ISERROR":       {formula: "ISERROR(1/0)", expected: "TRUE"},
	"ISNUMBER":      {formula: "ISNUMBER(42)", expected: "TRUE"},
	"ISTEXT":        {formula: "ISTEXT(\"hello\")", expected: "TRUE"},
	"LEFT":          {formula: "LEFT(\"hello\",3)", expected: "hel"},
	"LEN":           {formula: "LEN(\"hello\")", expected: "5"},
	"LN":            {formula: "LN(1)", expected: "0"},
	"LOG":           {formula: "LOG(100,10)", expected: "2"},
	"LOG10":         {formula: "LOG10(100)", expected: "2"},
	"LOWER":         {formula: "LOWER(\"ABC\")", expected: "abc"},
	"MAX":           {formula: "MAX(1,2,3)", expected: "3"},
	"MEDIAN":        {formula: "MEDIAN(1,2,3)", expected: "2"},
	"MID":           {formula: "MID(\"hello\",2,3)", expected: "ell"},
	"MIN":           {formula: "MIN(1,2,3)", expected: "1"},
	"MINUTE":        {formula: "MINUTE(0.5)", expected: "0"},
	"MOD":           {formula: "MOD(10,3)", expected: "1"},
	"MONTH":         {formula: "MONTH(45306)", expected: "1"},
	"NOT":           {formula: "NOT(TRUE)", expected: "FALSE"},
	"OR":            {formula: "OR(FALSE,TRUE)", expected: "TRUE"},
	"PI":            {formula: "PI()", expected: "3.14159265358979"},
	"POWER":         {formula: "POWER(2,3)", expected: "8"},
	"PRODUCT":       {formula: "PRODUCT(2,3,4)", expected: "24"},
	"PROPER":        {formula: "PROPER(\"hello world\")", expected: "Hello World"},
	"REPLACE":       {formula: "REPLACE(\"hello\",2,3,\"X\")", expected: "hXo"},
	"REPT":          {formula: "REPT(\"ab\",3)", expected: "ababab"},
	"RIGHT":         {formula: "RIGHT(\"hello\",3)", expected: "llo"},
	"ROUND":         {formula: "ROUND(3.456,2)", expected: "3.46"},
	"ROUNDDOWN":     {formula: "ROUNDDOWN(3.456,2)", expected: "3.45"},
	"ROUNDUP":       {formula: "ROUNDUP(3.451,2)", expected: "3.46"},
	"ROW":           {formula: "ROW(A5)", expected: "5"},
	"SEARCH":        {formula: "SEARCH(\"b\",\"abc\")", expected: "2"},
	"SECOND":        {formula: "SECOND(0.5)", expected: "0"},
	"SIGN":          {formula: "SIGN(-5)", expected: "-1"},
	"SIN":           {formula: "SIN(0)", expected: "0"},
	"SQRT":          {formula: "SQRT(9)", expected: "3"},
	"SUBSTITUTE":    {formula: "SUBSTITUTE(\"aaa\",\"a\",\"b\")", expected: "bbb"},
	"SUM":           {formula: "SUM(1,2,3)", expected: "6"},
	"TAN":           {formula: "TAN(0)", expected: "0"},
	"TRIM":          {formula: "TRIM(\"  hello  \")", expected: "hello"},
	"TRUNC":         {formula: "TRUNC(3.7)", expected: "3"},
	"UPPER":         {formula: "UPPER(\"abc\")", expected: "ABC"},
	"VALUE":         {formula: "VALUE(\"123\")", expected: "123"},
	"XOR":           {formula: "XOR(TRUE,FALSE)", expected: "TRUE"},
	"YEAR":          {formula: "YEAR(45306)", expected: "2024"},
}

// IsolateFailure takes a spec and its mismatches and determines which specific
// functions are actually broken by testing each function in isolation.
// Returns only the functions that fail in isolation.
func IsolateFailure(spec *TestSpec, mismatches []Mismatch) []string {
	if len(mismatches) == 0 {
		return nil
	}

	// Extract all functions from mismatched formulas.
	candidateFuncs := make(map[string]bool)
	for _, m := range mismatches {
		if m.Formula != "" {
			for _, fn := range ExtractFunctionsFromFormula(m.Formula) {
				candidateFuncs[fn] = true
			}
		}
	}

	// For #NAME? errors, the outermost function is definitely missing.
	var definitelyBroken []string
	seen := make(map[string]bool)
	for _, m := range mismatches {
		if strings.Contains(m.Werkbook, "#NAME?") && m.Formula != "" {
			fn := extractOutermostFunction(m.Formula)
			if fn != "" && !seen[fn] {
				seen[fn] = true
				definitelyBroken = append(definitelyBroken, fn)
			}
		}
	}

	// For value mismatches with multiple functions, test each in isolation.
	for fn := range candidateFuncs {
		if seen[fn] {
			continue
		}

		test, ok := isolationTests[fn]
		if !ok {
			// No isolation test available; assume it could be broken.
			if !seen[fn] {
				seen[fn] = true
				definitelyBroken = append(definitelyBroken, fn)
			}
			continue
		}

		// Build a minimal spec to test this single function.
		testSpec := &TestSpec{
			Name: "isolation_" + fn,
			Sheets: []SheetSpec{{
				Name: "Sheet1",
				Cells: []CellSpec{
					{Ref: "A1", Formula: test.formula},
				},
			}},
			Checks: []CheckSpec{{
				Ref:      "Sheet1!A1",
				Expected: test.expected,
				Type:     "number",
			}},
		}

		// Try to build and check.
		result, err := quickCheck(testSpec)
		if err != nil || result != test.expected {
			if !seen[fn] {
				seen[fn] = true
				definitelyBroken = append(definitelyBroken, fn)
			}
		}
	}

	sort.Strings(definitelyBroken)
	return definitelyBroken
}

// quickCheck builds a spec and returns the computed value for the first check.
func quickCheck(spec *TestSpec) (string, error) {
	tmpDir, err := os.MkdirTemp("", "werkbook-isolate-*")
	if err != nil {
		return "", err
	}
	defer os.RemoveAll(tmpDir)

	_, results, err := BuildXLSX(spec, tmpDir)
	if err != nil {
		return "", err
	}
	if len(results) == 0 {
		return "", nil
	}
	return results[0].Value, nil
}
