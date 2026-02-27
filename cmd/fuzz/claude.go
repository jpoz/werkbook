package main

import (
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"strings"
)

// implementedFunctions is the set of functions werkbook's formula engine currently supports.
// Sourced from formula/compiler.go knownFunctions.
var implementedFunctions = map[string]bool{
	"ABS": true, "ACOS": true, "AND": true, "ASIN": true, "ATAN": true, "ATAN2": true,
	"AVERAGE": true, "AVERAGEIF": true, "AVERAGEIFS": true, "CEILING": true,
	"CHAR": true, "CHOOSE": true, "CLEAN": true, "CODE": true, "COLUMN": true,
	"COLUMNS": true, "CONCAT": true, "CONCATENATE": true, "COS": true,
	"COUNT": true, "COUNTA": true, "COUNTBLANK": true, "COUNTIF": true, "COUNTIFS": true,
	"DATE": true, "DAY": true, "EXACT": true, "EXP": true, "FIND": true, "FLOOR": true,
	"HLOOKUP": true, "HOUR": true, "IF": true, "IFERROR": true, "IFNA": true,
	"INDEX": true, "INDIRECT": true, "INT": true,
	"ISBLANK": true, "ISERR": true, "ISERROR": true, "ISNA": true, "ISNUMBER": true, "ISTEXT": true,
	"LARGE": true, "LEFT": true, "LEN": true, "LN": true, "LOG": true, "LOG10": true,
	"LOOKUP": true, "LOWER": true, "MATCH": true, "MAX": true, "MEDIAN": true,
	"MID": true, "MIN": true, "MINUTE": true, "MOD": true, "MONTH": true,
	"NOT": true, "NOW": true, "OR": true, "PI": true, "POWER": true, "PRODUCT": true,
	"PROPER": true, "RAND": true, "RANDBETWEEN": true,
	"REPLACE": true, "REPT": true, "RIGHT": true,
	"ROUND": true, "ROUNDDOWN": true, "ROUNDUP": true, "ROW": true, "ROWS": true,
	"SEARCH": true, "SECOND": true, "SIN": true, "SMALL": true, "SORT": true,
	"SQRT": true, "SUBSTITUTE": true, "SUM": true, "SUMIF": true, "SUMIFS": true, "SUMPRODUCT": true,
	"TAN": true, "TEXT": true, "TIME": true, "TODAY": true, "TRIM": true, "UPPER": true,
	"VALUE": true, "VLOOKUP": true, "XLOOKUP": true, "XOR": true, "YEAR": true,
}

// knownFunctionsList is the comprehensive list of Excel functions for the generation prompt.
var knownFunctionsList = []string{
	// Math & Trig
	"ABS", "ACOS", "ACOSH", "ACOT", "ACOTH", "ARABIC", "ASIN", "ASINH",
	"ATAN", "ATAN2", "ATANH", "BASE",
	"CEILING", "CEILING.MATH", "CEILING.PRECISE",
	"COMBIN", "COMBINA", "COS", "COSH", "COT", "COTH", "CSC", "CSCH",
	"DECIMAL", "DEGREES", "EVEN", "EXP",
	"FACT", "FACTDOUBLE", "FLOOR", "FLOOR.MATH", "FLOOR.PRECISE",
	"GCD", "INT", "LCM", "LN", "LOG", "LOG10",
	"MDETERM", "MINVERSE", "MMULT", "MOD", "MROUND", "MULTINOMIAL", "MUNIT",
	"ODD", "PI", "POWER", "PRODUCT", "QUOTIENT",
	"RADIANS", "ROMAN", "ROUND", "ROUNDDOWN", "ROUNDUP",
	"SEC", "SECH", "SERIESSUM", "SIGN", "SIN", "SINH", "SQRT", "SQRTPI",
	"SUBTOTAL", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT",
	"SUMSQ", "SUMX2MY2", "SUMX2PY2", "SUMXMY2",
	"TAN", "TANH", "TRUNC",

	// Logical
	"AND", "FALSE", "IF", "IFERROR", "IFNA", "IFS",
	"NOT", "OR", "SWITCH", "TRUE", "XOR",

	// Text
	"CHAR", "CLEAN", "CODE", "CONCATENATE", "DOLLAR",
	"EXACT", "FIND", "FIXED", "LEFT", "LEN", "LOWER", "MID",
	"NUMBERVALUE", "PROPER", "REPLACE", "REPT", "RIGHT",
	"SEARCH", "SUBSTITUTE", "T", "TEXT", "TEXTJOIN",
	"TRIM", "UNICHAR", "UNICODE", "UPPER", "VALUE",

	// Statistical
	"AVEDEV", "AVERAGE", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS",
	"BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.DIST.RANGE", "BINOM.INV",
	"CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST",
	"CONFIDENCE.NORM", "CONFIDENCE.T", "CORREL", "COUNT", "COUNTA", "COUNTBLANK",
	"COUNTIF", "COUNTIFS", "COVARIANCE.P", "COVARIANCE.S",
	"DEVSQ", "EXPON.DIST",
	"F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST",
	"FISHER", "FISHERINV", "FORECAST.LINEAR", "FREQUENCY",
	"GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMALN", "GAMMALN.PRECISE",
	"GAUSS", "GEOMEAN", "GROWTH",
	"HARMEAN", "HYPGEOM.DIST",
	"INTERCEPT", "KURT",
	"LARGE", "LINEST", "LOGEST", "LOGNORM.DIST", "LOGNORM.INV",
	"MAX", "MAXA", "MAXIFS", "MEDIAN", "MIN", "MINA", "MINIFS",
	"MODE.MULT", "MODE.SNGL",
	"NEGBINOM.DIST", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV",
	"PEARSON", "PERCENTILE.EXC", "PERCENTILE.INC",
	"PERCENTRANK.EXC", "PERCENTRANK.INC",
	"PERMUT", "PERMUTATIONA", "PHI", "POISSON.DIST", "PROB",
	"QUARTILE.EXC", "QUARTILE.INC",
	"RANK.AVG", "RANK.EQ", "RSQ",
	"SKEW", "SKEW.P", "SLOPE", "SMALL", "STANDARDIZE",
	"STDEV.P", "STDEV.S", "STDEVA", "STDEVPA", "STEYX",
	"T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST",
	"TREND", "TRIMMEAN",
	"VAR.P", "VAR.S", "VARA", "VARPA",
	"WEIBULL.DIST", "Z.TEST",

	// Lookup & Reference
	"ADDRESS", "AREAS", "CHOOSE", "COLUMN", "COLUMNS",
	"FORMULATEXT", "HLOOKUP", "HYPERLINK",
	"INDEX", "LOOKUP", "MATCH",
	"ROW", "ROWS", "TRANSPOSE", "VLOOKUP",

	// Date & Time
	"DATE", "DATEDIF", "DATEVALUE", "DAY", "DAYS", "DAYS360",
	"EDATE", "EOMONTH", "HOUR", "ISOWEEKNUM",
	"MINUTE", "MONTH",
	"NETWORKDAYS", "NETWORKDAYS.INTL",
	"SECOND", "TIME", "TIMEVALUE",
	"WEEKDAY", "WEEKNUM",
	"WORKDAY", "WORKDAY.INTL",
	"YEAR", "YEARFRAC",

	// Information
	"ERROR.TYPE", "ISBLANK", "ISERR", "ISERROR",
	"ISEVEN", "ISFORMULA", "ISLOGICAL", "ISNA", "ISNONTEXT",
	"ISNUMBER", "ISODD", "ISREF", "ISTEXT",
	"N", "NA", "SHEET", "SHEETS", "TYPE",

	// Financial
	"ACCRINT", "ACCRINTM", "AMORDEGRC", "AMORLINC",
	"COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD",
	"CUMIPMT", "CUMPRINC",
	"DB", "DDB", "DISC", "DOLLARDE", "DOLLARFR", "DURATION",
	"EFFECT", "FV", "FVSCHEDULE",
	"INTRATE", "IPMT", "IRR", "ISPMT",
	"MDURATION", "MIRR",
	"NOMINAL", "NPER", "NPV",
	"ODDFPRICE", "ODDFYIELD", "ODDLPRICE", "ODDLYIELD",
	"PDURATION", "PMT", "PPMT",
	"PRICE", "PRICEDISC", "PRICEMAT", "PV",
	"RATE", "RECEIVED", "RRI",
	"SLN", "SYD",
	"TBILLEQ", "TBILLPRICE", "TBILLYIELD",
	"VDB",
	"XIRR", "XNPV",
	"YIELD", "YIELDDISC", "YIELDMAT",

	// Engineering
	"BESSELI", "BESSELJ", "BESSELK", "BESSELY",
	"BIN2DEC", "BIN2HEX", "BIN2OCT",
	"BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR",
	"COMPLEX", "CONVERT",
	"DEC2BIN", "DEC2HEX", "DEC2OCT",
	"DELTA", "ERF", "ERF.PRECISE", "ERFC", "ERFC.PRECISE",
	"GESTEP",
	"HEX2BIN", "HEX2DEC", "HEX2OCT",
	"IMABS", "IMAGINARY", "IMARGUMENT", "IMCONJUGATE",
	"IMCOS", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH",
	"IMDIV", "IMEXP", "IMLN", "IMLOG10", "IMLOG2",
	"IMPOWER", "IMPRODUCT", "IMREAL",
	"IMSEC", "IMSECH", "IMSIN", "IMSINH", "IMSQRT",
	"IMSUB", "IMSUM", "IMTAN",
	"OCT2BIN", "OCT2DEC", "OCT2HEX",

	// Database
	"DAVERAGE", "DCOUNT", "DCOUNTA", "DGET",
	"DMAX", "DMIN", "DPRODUCT",
	"DSTDEV", "DSTDEVP", "DSUM",
	"DVAR", "DVARP",
}

// generateSpec shells out to `claude -p` to generate a test spec.
func generateSpec(seed string, verbose bool) (*TestSpec, error) {
	prompt := buildGenerationPrompt(seed)

	if verbose {
		fmt.Println("  Generating spec with claude...")
	}

	output, err := runClaude(prompt)
	if err != nil {
		return nil, fmt.Errorf("claude generate: %w", err)
	}

	// Extract JSON from the output (claude may wrap it in markdown).
	jsonStr := extractJSON(output)
	if jsonStr == "" {
		return nil, fmt.Errorf("no JSON found in claude output:\n%s", truncate(output, 500))
	}

	var spec TestSpec
	if err := json.Unmarshal([]byte(jsonStr), &spec); err != nil {
		return nil, fmt.Errorf("parse generated spec: %w\njson: %s", err, truncate(jsonStr, 500))
	}

	if err := validateSpec(&spec); err != nil {
		return nil, fmt.Errorf("invalid generated spec: %w", err)
	}

	return &spec, nil
}

// suggestFix shells out to `claude -p` to suggest a fix for mismatches.
func suggestFix(spec *TestSpec, mismatches []mismatch, verbose bool) (string, error) {
	prompt := buildFixPrompt(spec, mismatches)

	if verbose {
		fmt.Println("  Asking claude for fix suggestions...")
	}

	output, err := runClaude(prompt)
	if err != nil {
		return "", fmt.Errorf("claude fix: %w", err)
	}

	return output, nil
}

// runClaude executes `claude -p` with the given prompt and returns stdout.
func runClaude(prompt string) (string, error) {
	cmd := exec.Command("claude", "-p", prompt)
	// Clear CLAUDECODE env var to allow running inside a Claude Code session.
	cmd.Env = filterEnv(os.Environ(), "CLAUDECODE")
	out, err := cmd.CombinedOutput()
	if err != nil {
		return "", fmt.Errorf("claude command failed: %w\noutput: %s", err, truncate(string(out), 500))
	}
	return string(out), nil
}

// buildGenerationPrompt creates the prompt for generating a test spec.
func buildGenerationPrompt(seed string) string {
	var sb strings.Builder
	sb.WriteString("Generate a JSON test spec for testing Excel formula evaluation in a Go spreadsheet library.\n\n")
	sb.WriteString("The spec must be valid JSON matching this exact schema:\n")
	sb.WriteString(`{
  "name": "test_<descriptive_name>",
  "sheets": [
    {
      "name": "Sheet1",
      "cells": [
        {"ref": "A1", "value": 10, "type": "number"},
        {"ref": "A2", "value": "hello", "type": "string"},
        {"ref": "A3", "value": true, "type": "bool"},
        {"ref": "B1", "formula": "SUM(A1:A3)"}
      ]
    }
  ],
  "checks": [
    {"ref": "Sheet1!B1", "expected": "10", "type": "number"}
  ]
}`)
	sb.WriteString("\n\n")

	if seed != "" {
		sb.WriteString(fmt.Sprintf("Focus on testing %s functions. ", seed))
	}

	// Split into implemented (werkbook supports) and unimplemented.
	var implemented, unimplemented []string
	for _, fn := range knownFunctionsList {
		if implementedFunctions[fn] {
			implemented = append(implemented, fn)
		} else {
			unimplemented = append(unimplemented, fn)
		}
	}

	sb.WriteString("Implemented functions (werkbook already supports these):\n")
	sb.WriteString(strings.Join(implemented, ", "))
	sb.WriteString("\n\n")
	if len(unimplemented) > 0 {
		sb.WriteString("Unimplemented functions (werkbook does NOT support these yet):\n")
		sb.WriteString(strings.Join(unimplemented, ", "))
		sb.WriteString("\n\n")
	}

	sb.WriteString("Requirements:\n")
	sb.WriteString("- Use AT MOST 1 unimplemented function per spec. The rest of the functions MUST be from the implemented list. This lets us test one new function at a time against a backdrop of known-working ones.\n")
	sb.WriteString("- Create edge cases: empty cells, zero values, negative numbers, large numbers, mixed types\n")
	sb.WriteString("- Test nested formulas (e.g., IF(SUM(A1:A3)>10, AVERAGE(A1:A3), 0))\n")
	sb.WriteString("- Test cross-cell references and ranges\n")
	sb.WriteString("- Include 5-15 cells and 3-8 checks\n")
	sb.WriteString("- Use realistic but tricky inputs that might expose bugs\n")
	sb.WriteString("- DO NOT use non-deterministic functions: RAND, RANDBETWEEN, RANDARRAY, NOW, TODAY\n")
	sb.WriteString("- DO NOT use INDIRECT (volatile/indirect resolution)\n")
	sb.WriteString("- DO NOT use Excel 365+ dynamic array functions: CONCAT, XLOOKUP, XMATCH, SORT, SORTBY, FILTER, UNIQUE, SEQUENCE, LET, LAMBDA, MAP, REDUCE, SCAN, BYCOL, BYROW, MAKEARRAY, CHOOSECOLS, CHOOSEROWS, DROP, TAKE, EXPAND, HSTACK, VSTACK, WRAPCOLS, WRAPROWS, TOCOL, TOROW, TEXTBEFORE, TEXTAFTER, TEXTSPLIT, VALUETOTEXT, ARRAYTOTEXT, ISOMITTED, IMAGE, GROUPBY, PIVOTBY, PERCENTOF, TRIMRANGE\n")
	sb.WriteString("- DO NOT use environment-dependent functions: CELL, INFO\n")
	sb.WriteString("- Use CONCATENATE instead of CONCAT\n")
	sb.WriteString("- The formulas will be validated against LibreOffice, so only use functions that LibreOffice supports\n")
	sb.WriteString("- The 'expected' field in checks should be the string representation of the expected result\n")
	sb.WriteString("- For numbers, use the simplest representation (e.g., \"10\" not \"10.0\")\n")
	sb.WriteString("- For booleans, use \"TRUE\" or \"FALSE\"\n")
	sb.WriteString("- Cell refs in formulas should NOT include the sheet name if on the same sheet\n")
	sb.WriteString("- Multi-sheet tests are welcome but keep them simple\n")
	sb.WriteString("\n")
	sb.WriteString("Output ONLY the JSON, no explanation or markdown fences.\n")

	return sb.String()
}

// buildFixPrompt creates the prompt for suggesting a fix.
func buildFixPrompt(spec *TestSpec, mismatches []mismatch) string {
	specJSON, _ := json.MarshalIndent(spec, "", "  ")

	var sb strings.Builder
	sb.WriteString("I'm testing a Go spreadsheet formula engine (github.com/jpoz/werkbook) against LibreOffice.\n\n")
	sb.WriteString("Test spec:\n```json\n")
	sb.Write(specJSON)
	sb.WriteString("\n```\n\n")
	sb.WriteString("Mismatches found (LibreOffice is the ground truth):\n")
	for _, m := range mismatches {
		fmt.Fprintf(&sb, "- %s: werkbook=%q, libreoffice=%q (%s)\n", m.Ref, m.Werkbook, m.LibreOff, m.Reason)
	}
	sb.WriteString("\nThe formula engine code is in the `formula/` package. Key files:\n")
	sb.WriteString("- formula/functions.go: Built-in function implementations\n")
	sb.WriteString("- formula/vm.go: Bytecode VM that executes compiled formulas\n")
	sb.WriteString("- formula/compiler.go: Compiles AST to bytecode\n")
	sb.WriteString("- formula/parser.go: Parses formula strings to AST\n\n")
	sb.WriteString("Analyze the mismatches and suggest what might be wrong in the formula engine.\n")
	sb.WriteString("Focus on the most likely root cause and suggest a specific fix.\n")

	return sb.String()
}

// extractJSON finds the first JSON object in the output.
func extractJSON(s string) string {
	// Try to find raw JSON first.
	start := strings.Index(s, "{")
	if start < 0 {
		return ""
	}

	// Find matching closing brace.
	depth := 0
	for i := start; i < len(s); i++ {
		switch s[i] {
		case '{':
			depth++
		case '}':
			depth--
			if depth == 0 {
				return s[start : i+1]
			}
		}
	}
	return ""
}

// truncate shortens a string to maxLen, adding "..." if truncated.
func truncate(s string, maxLen int) string {
	if len(s) <= maxLen {
		return s
	}
	return s[:maxLen] + "..."
}

// filterEnv returns a copy of env with the named variable removed.
func filterEnv(env []string, name string) []string {
	prefix := name + "="
	out := make([]string, 0, len(env))
	for _, e := range env {
		if !strings.HasPrefix(e, prefix) {
			out = append(out, e)
		}
	}
	return out
}
