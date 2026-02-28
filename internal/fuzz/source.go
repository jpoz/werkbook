package fuzz

import (
	"bufio"
	"fmt"
	"os"
	"strings"
)

// formulaSourceMap maps Excel function names to their Go implementation files.
var formulaSourceMap = map[string]string{
	// Math
	"ABS": "formula/functions_math.go", "ACOS": "formula/functions_math.go",
	"ASIN": "formula/functions_math.go", "ATAN": "formula/functions_math.go",
	"ATAN2": "formula/functions_math.go", "CEILING": "formula/functions_math.go",
	"COMBIN": "formula/functions_math.go", "COS": "formula/functions_math.go",
	"DEGREES": "formula/functions_math.go", "EVEN": "formula/functions_math.go",
	"EXP": "formula/functions_math.go", "FACT": "formula/functions_math.go",
	"FLOOR": "formula/functions_math.go", "GCD": "formula/functions_math.go",
	"INT": "formula/functions_math.go", "LCM": "formula/functions_math.go",
	"LN": "formula/functions_math.go", "LOG": "formula/functions_math.go",
	"LOG10": "formula/functions_math.go", "MOD": "formula/functions_math.go",
	"MROUND": "formula/functions_math.go", "ODD": "formula/functions_math.go",
	"PERMUT": "formula/functions_math.go", "PI": "formula/functions_math.go",
	"POWER": "formula/functions_math.go", "PRODUCT": "formula/functions_math.go",
	"QUOTIENT": "formula/functions_math.go", "RADIANS": "formula/functions_math.go",
	"RAND": "formula/functions_math.go", "RANDBETWEEN": "formula/functions_math.go",
	"ROUND": "formula/functions_math.go", "ROUNDDOWN": "formula/functions_math.go",
	"ROUNDUP": "formula/functions_math.go", "SIGN": "formula/functions_math.go",
	"SIN": "formula/functions_math.go", "SQRT": "formula/functions_math.go",
	"SUBTOTAL": "formula/functions_math.go", "TAN": "formula/functions_math.go",
	"TRUNC": "formula/functions_math.go",
	// Stat
	"AVERAGE": "formula/functions_stat.go", "AVERAGEIF": "formula/functions_stat.go",
	"AVERAGEIFS": "formula/functions_stat.go", "COUNT": "formula/functions_stat.go",
	"COUNTA": "formula/functions_stat.go", "COUNTBLANK": "formula/functions_stat.go",
	"COUNTIF": "formula/functions_stat.go", "COUNTIFS": "formula/functions_stat.go",
	"LARGE": "formula/functions_stat.go", "MAX": "formula/functions_stat.go",
	"MAXIFS": "formula/functions_stat.go", "MEDIAN": "formula/functions_stat.go",
	"MIN": "formula/functions_stat.go", "MINIFS": "formula/functions_stat.go",
	"MODE": "formula/functions_stat.go", "PERCENTILE": "formula/functions_stat.go",
	"RANK": "formula/functions_stat.go", "SMALL": "formula/functions_stat.go",
	"STDEV": "formula/functions_stat.go", "STDEVP": "formula/functions_stat.go",
	"SUM": "formula/functions_stat.go", "SUMIF": "formula/functions_stat.go",
	"SUMIFS": "formula/functions_stat.go", "SUMPRODUCT": "formula/functions_stat.go",
	"SUMSQ": "formula/functions_stat.go", "VAR": "formula/functions_stat.go",
	"VARP": "formula/functions_stat.go",
	// Text
	"CHAR": "formula/functions_text.go", "CHOOSE": "formula/functions_text.go",
	"CLEAN": "formula/functions_text.go", "CODE": "formula/functions_text.go",
	"CONCAT": "formula/functions_text.go", "CONCATENATE": "formula/functions_text.go",
	"EXACT": "formula/functions_text.go", "FIND": "formula/functions_text.go",
	"FIXED": "formula/functions_text.go", "LEFT": "formula/functions_text.go",
	"LEN": "formula/functions_text.go", "LOWER": "formula/functions_text.go",
	"MID": "formula/functions_text.go", "NUMBERVALUE": "formula/functions_text.go",
	"PROPER": "formula/functions_text.go", "REPLACE": "formula/functions_text.go",
	"REPT": "formula/functions_text.go", "RIGHT": "formula/functions_text.go",
	"SEARCH": "formula/functions_text.go", "SUBSTITUTE": "formula/functions_text.go",
	"T": "formula/functions_text.go", "TEXT": "formula/functions_text.go",
	"TEXTJOIN": "formula/functions_text.go", "TRIM": "formula/functions_text.go",
	"UPPER": "formula/functions_text.go", "VALUE": "formula/functions_text.go",
	// Date
	"DATE": "formula/functions_date.go", "DATEDIF": "formula/functions_date.go",
	"DATEVALUE": "formula/functions_date.go", "DAY": "formula/functions_date.go",
	"DAYS": "formula/functions_date.go", "EDATE": "formula/functions_date.go",
	"EOMONTH": "formula/functions_date.go", "HOUR": "formula/functions_date.go",
	"ISOWEEKNUM": "formula/functions_date.go", "MINUTE": "formula/functions_date.go",
	"MONTH": "formula/functions_date.go", "NETWORKDAYS": "formula/functions_date.go",
	"NOW": "formula/functions_date.go", "SECOND": "formula/functions_date.go",
	"TIME": "formula/functions_date.go", "TODAY": "formula/functions_date.go",
	"WEEKDAY": "formula/functions_date.go", "WEEKNUM": "formula/functions_date.go",
	"WORKDAY": "formula/functions_date.go", "YEAR": "formula/functions_date.go",
	"YEARFRAC": "formula/functions_date.go",
	// Logic
	"AND": "formula/functions_logic.go", "IF": "formula/functions_logic.go",
	"IFERROR": "formula/functions_logic.go", "IFS": "formula/functions_logic.go",
	"NOT": "formula/functions_logic.go", "OR": "formula/functions_logic.go",
	"SORT": "formula/functions_logic.go", "SWITCH": "formula/functions_logic.go",
	"XOR": "formula/functions_logic.go",
	// Lookup
	"ADDRESS": "formula/functions_lookup.go", "HLOOKUP": "formula/functions_lookup.go",
	"INDEX": "formula/functions_lookup.go", "LOOKUP": "formula/functions_lookup.go",
	"MATCH": "formula/functions_lookup.go", "VLOOKUP": "formula/functions_lookup.go",
	"XLOOKUP": "formula/functions_lookup.go",
	// Info
	"COLUMN": "formula/functions_info.go", "COLUMNS": "formula/functions_info.go",
	"ERROR.TYPE": "formula/functions_info.go", "IFNA": "formula/functions_info.go",
	"ISBLANK": "formula/functions_info.go", "ISERR": "formula/functions_info.go",
	"ISERROR": "formula/functions_info.go", "ISEVEN": "formula/functions_info.go",
	"ISLOGICAL": "formula/functions_info.go", "ISNA": "formula/functions_info.go",
	"ISNONTEXT": "formula/functions_info.go", "ISNUMBER": "formula/functions_info.go",
	"ISODD": "formula/functions_info.go", "ISTEXT": "formula/functions_info.go",
	"N": "formula/functions_info.go", "NA": "formula/functions_info.go",
	"ROW": "formula/functions_info.go", "ROWS": "formula/functions_info.go",
	"TYPE": "formula/functions_info.go",
}

// goFuncName returns the Go function name for an Excel function.
func goFuncName(excelName string) string {
	name := strings.ReplaceAll(excelName, ".", "")
	switch excelName {
	case "VALUE":
		return "fnVALUEFn"
	default:
		return "fn" + name
	}
}

// extractFunctionSource reads a Go file and extracts a single function body
// by matching `func <funcName>(` and tracking brace depth.
func extractFunctionSource(filePath, funcName string) (string, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return "", err
	}
	defer f.Close()

	scanner := bufio.NewScanner(f)
	scanner.Buffer(make([]byte, 256*1024), 256*1024)

	prefix := "func " + funcName + "("
	var collecting bool
	var lines []string
	depth := 0

	for scanner.Scan() {
		line := scanner.Text()
		if !collecting {
			if strings.HasPrefix(line, prefix) {
				collecting = true
				lines = append(lines, line)
				depth += strings.Count(line, "{") - strings.Count(line, "}")
				if depth == 0 {
					break
				}
				continue
			}
			continue
		}

		lines = append(lines, line)
		depth += strings.Count(line, "{") - strings.Count(line, "}")
		if depth <= 0 {
			break
		}
	}

	if len(lines) == 0 {
		return "", fmt.Errorf("function %s not found in %s", funcName, filePath)
	}
	return strings.Join(lines, "\n"), nil
}

// buildSourceBlock builds a source code block for the given broken functions
// to include in the fix prompt.
func buildSourceBlock(brokenFns []string, failureType FailureType) string {
	if len(brokenFns) == 0 {
		return ""
	}

	const maxSize = 30 * 1024 // 30KB cap
	var sb strings.Builder
	totalSize := 0

	// Group by file for missing functions (include whole file).
	if failureType == FailureMissingFunction {
		included := make(map[string]bool)
		for _, fn := range brokenFns {
			filePath, ok := formulaSourceMap[fn]
			if !ok || included[filePath] {
				continue
			}
			content, err := os.ReadFile(filePath)
			if err != nil {
				continue
			}
			if totalSize+len(content) > maxSize {
				continue
			}
			included[filePath] = true
			fmt.Fprintf(&sb, "--- %s (full file for reference) ---\n", filePath)
			sb.Write(content)
			sb.WriteString("\n---\n\n")
			totalSize += len(content)
		}
	} else {
		// Bug fix: include individual function bodies.
		for _, fn := range brokenFns {
			filePath, ok := formulaSourceMap[fn]
			if !ok {
				continue
			}
			goName := goFuncName(fn)
			body, err := extractFunctionSource(filePath, goName)
			if err != nil {
				continue
			}
			if totalSize+len(body) > maxSize {
				continue
			}
			fmt.Fprintf(&sb, "--- %s (%s) ---\n", filePath, goName)
			sb.WriteString(body)
			sb.WriteString("\n---\n\n")
			totalSize += len(body)
		}
	}

	return sb.String()
}
