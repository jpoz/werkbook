package fuzz

import (
	"bufio"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
)

// IsClaudeCLIError reports whether err looks like a Claude CLI failure
// (crash, auth, nested session, etc.) as opposed to a logic error in
// the generated output.
func IsClaudeCLIError(err error) bool {
	if err == nil {
		return false
	}
	msg := err.Error()
	return strings.Contains(msg, "claude command failed") ||
		strings.Contains(msg, "claude streaming failed") ||
		strings.Contains(msg, "claude apply fix failed") ||
		strings.Contains(msg, "start claude")
}

// ImplementedFunctions is the set of functions werkbook's formula engine currently supports.
// Sourced from formula/compiler.go knownFunctions.
var ImplementedFunctions = map[string]bool{
	"ABS": true, "ACOS": true, "ACOSH": true, "AND": true, "ASIN": true, "ASINH": true, "ATAN": true, "ATAN2": true, "ATANH": true,
	"AVERAGE": true, "AVERAGEIF": true, "AVERAGEIFS": true, "CEILING": true,
	"CHAR": true, "CHOOSE": true, "CLEAN": true, "CODE": true, "COLUMN": true,
	"COLUMNS": true, "CONCAT": true, "CONCATENATE": true, "COS": true, "COSH": true,
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
	"SEARCH": true, "SECOND": true, "SIGN": true, "SIN": true, "SINH": true, "SMALL": true, "SORT": true,
	"SQRT": true, "SQRTPI": true, "SUBSTITUTE": true, "SUM": true, "SUMIF": true, "SUMIFS": true, "SUMPRODUCT": true,
	"TAN": true, "TANH": true, "TEXT": true, "TIME": true, "TODAY": true, "TRIM": true, "TRUNC": true,
	"UPPER": true, "VALUE": true, "VLOOKUP": true, "XLOOKUP": true, "XOR": true, "YEAR": true,
}

// KnownFunctionsList is the comprehensive list of Excel functions for the generation prompt.
var KnownFunctionsList = []string{
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

// FailureType classifies a mismatch into categories for targeted fix prompts.
type FailureType string

const (
	FailureMissingFunction FailureType = "missing_function"
	FailureBugInFunction   FailureType = "bug_in_function"
)

// ClassifyFailure categorizes mismatches into failure types.
func ClassifyFailure(mismatches []Mismatch) FailureType {
	for _, m := range mismatches {
		if strings.Contains(m.Werkbook, "#NAME?") {
			return FailureMissingFunction
		}
	}
	return FailureBugInFunction
}

// ScanImplementedFunctions reads formula/compiler.go to find all functions
// registered in the knownFunctions array.
func ScanImplementedFunctions() map[string]bool {
	result := make(map[string]bool)

	f, err := os.Open("formula/compiler.go")
	if err != nil {
		return result
	}
	defer f.Close()

	scanner := bufio.NewScanner(f)
	inArray := false
	for scanner.Scan() {
		line := strings.TrimSpace(scanner.Text())

		if strings.Contains(line, "knownFunctions") && strings.Contains(line, "[...]string{") {
			inArray = true
			continue
		}

		if inArray {
			if line == "}" {
				break
			}
			// Extract quoted strings from the line.
			for _, part := range strings.Split(line, "\"") {
				name := strings.TrimSpace(part)
				if name == "" || name == "," || strings.HasPrefix(name, "//") {
					continue
				}
				// Valid function names are uppercase letters, digits, dots.
				valid := true
				for _, c := range name {
					if !((c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '.') {
						valid = false
						break
					}
				}
				if valid && len(name) >= 1 {
					result[name] = true
				}
			}
		}
	}

	return result
}

// SyncImplementedFunctions scans formula/compiler.go and updates the
// ImplementedFunctions map. Returns the total number of implemented functions.
func SyncImplementedFunctions() int {
	scanned := ScanImplementedFunctions()
	if len(scanned) == 0 {
		return 0
	}
	for fn := range scanned {
		ImplementedFunctions[fn] = true
	}
	return len(ImplementedFunctions)
}

// GenerateSpec shells out to `claude -p` to generate a test spec.
// If eval is nil, defaults to LibreOffice behavior.
// avoid is an optional list of functions to tell Claude to skip (broken functions).
// prioritize is an optional list of under-tested functions to prioritize.
func GenerateSpec(seed string, complexity *ComplexityConfig, eval Evaluator, avoid []string, prioritize []string, verbose bool) (*TestSpec, error) {
	prompt := BuildGenerationPrompt(seed, complexity, eval, avoid, prioritize)

	output, err := RunClaude(prompt)
	if err != nil {
		return nil, fmt.Errorf("claude generate: %w", err)
	}

	// Extract JSON from the output (claude may wrap it in markdown).
	jsonStr := ExtractJSON(output)
	if jsonStr == "" {
		return nil, fmt.Errorf("no JSON found in claude output:\n%s", Truncate(output, 500))
	}

	var spec TestSpec
	if err := json.Unmarshal([]byte(jsonStr), &spec); err != nil {
		return nil, fmt.Errorf("parse generated spec: %w\njson: %s", err, Truncate(jsonStr, 500))
	}

	var excluded map[string]bool
	if eval != nil {
		excluded = eval.ExcludedFunctions()
	}
	if err := ValidateSpec(&spec, excluded); err != nil {
		return nil, fmt.Errorf("invalid generated spec: %w", err)
	}

	return &spec, nil
}

// SuggestFix shells out to `claude -p` to suggest a fix for mismatches.
func SuggestFix(spec *TestSpec, mismatches []Mismatch, oracleName string, verbose bool) (string, error) {
	prompt := BuildFixPrompt(spec, mismatches, oracleName)

	output, err := RunClaude(prompt)
	if err != nil {
		return "", fmt.Errorf("claude fix: %w", err)
	}

	return output, nil
}

// GenerateTargetedSpec generates a test spec targeting a specific function.
func GenerateTargetedSpec(targetFn string, complexity *ComplexityConfig, eval Evaluator, verbose bool) (*TestSpec, error) {
	prompt := BuildTargetedGenerationPrompt(targetFn, complexity, eval)

	output, err := RunClaude(prompt)
	if err != nil {
		return nil, fmt.Errorf("claude generate targeted: %w", err)
	}

	jsonStr := ExtractJSON(output)
	if jsonStr == "" {
		return nil, fmt.Errorf("no JSON found in claude output:\n%s", Truncate(output, 500))
	}

	var spec TestSpec
	if err := json.Unmarshal([]byte(jsonStr), &spec); err != nil {
		return nil, fmt.Errorf("parse generated spec: %w\njson: %s", err, Truncate(jsonStr, 500))
	}

	var excluded map[string]bool
	if eval != nil {
		excluded = eval.ExcludedFunctions()
	}
	if err := ValidateSpec(&spec, excluded); err != nil {
		return nil, fmt.Errorf("invalid generated spec: %w", err)
	}

	return &spec, nil
}

// streamEvent represents a parsed event from Claude's stream-json output.
type streamEvent struct {
	Type    string     `json:"type"`
	Message *streamMsg `json:"message,omitempty"`
	Result  string     `json:"result,omitempty"`
}

type streamMsg struct {
	Content []streamContentBlock `json:"content,omitempty"`
}

type streamContentBlock struct {
	Type  string          `json:"type"`
	Name  string          `json:"name,omitempty"`
	Input json.RawMessage `json:"input,omitempty"`
}

// runClaudeStreamingFix runs claude with --output-format stream-json and writes
// tool progress to the given writer. Returns the final result text.
func runClaudeStreamingFix(prompt string, progress io.Writer) (string, error) {
	cmd := exec.Command("claude", "-p", "--dangerously-skip-permissions", "--output-format", "stream-json")
	cmd.Stdin = strings.NewReader(prompt)
	cmd.Env = FilterClaudeEnv(os.Environ())

	var stderrBuf strings.Builder
	cmd.Stderr = &stderrBuf

	stdout, err := cmd.StdoutPipe()
	if err != nil {
		return "", fmt.Errorf("create stdout pipe: %w", err)
	}

	if err := cmd.Start(); err != nil {
		return "", fmt.Errorf("start claude: %w", err)
	}

	var result string
	scanner := bufio.NewScanner(stdout)
	scanner.Buffer(make([]byte, 1024*1024), 1024*1024) // 1MB buffer for large JSON lines

	for scanner.Scan() {
		var event streamEvent
		if err := json.Unmarshal(scanner.Bytes(), &event); err != nil {
			continue
		}

		if event.Type == "assistant" && event.Message != nil && progress != nil {
			for _, block := range event.Message.Content {
				if block.Type == "tool_use" {
					writeToolProgress(progress, block.Name, block.Input)
				}
			}
		}

		if event.Type == "result" {
			result = event.Result
		}
	}

	if err := cmd.Wait(); err != nil {
		return "", fmt.Errorf("claude streaming failed: %w\nstderr: %s\nresult so far: %s", err, stderrBuf.String(), Truncate(result, 500))
	}

	return result, nil
}

// writeToolProgress writes a human-readable line describing what tool Claude is using.
func writeToolProgress(w io.Writer, toolName string, input json.RawMessage) {
	switch toolName {
	case "Read":
		var args struct {
			FilePath string `json:"file_path"`
		}
		json.Unmarshal(input, &args)
		if args.FilePath != "" {
			fmt.Fprintf(w, "    Reading %s\n", filepath.Base(args.FilePath))
		}
	case "Edit":
		var args struct {
			FilePath string `json:"file_path"`
		}
		json.Unmarshal(input, &args)
		if args.FilePath != "" {
			fmt.Fprintf(w, "    Editing %s\n", filepath.Base(args.FilePath))
		}
	case "Write":
		var args struct {
			FilePath string `json:"file_path"`
		}
		json.Unmarshal(input, &args)
		if args.FilePath != "" {
			fmt.Fprintf(w, "    Writing %s\n", filepath.Base(args.FilePath))
		}
	case "Bash":
		var args struct {
			Command string `json:"command"`
		}
		json.Unmarshal(input, &args)
		if args.Command != "" {
			cmd := args.Command
			if len(cmd) > 60 {
				cmd = cmd[:60] + "..."
			}
			fmt.Fprintf(w, "    Running: %s\n", cmd)
		}
	case "Glob":
		var args struct {
			Pattern string `json:"pattern"`
		}
		json.Unmarshal(input, &args)
		if args.Pattern != "" {
			fmt.Fprintf(w, "    Searching %s\n", args.Pattern)
		}
	case "Grep":
		var args struct {
			Pattern string `json:"pattern"`
		}
		json.Unmarshal(input, &args)
		if args.Pattern != "" {
			fmt.Fprintf(w, "    Grepping for \"%s\"\n", Truncate(args.Pattern, 40))
		}
	}
}

// ApplyFix shells out to `claude -p` with edit permissions to actually fix the code.
// Uses the failure type and fix history for targeted prompts.
// When progress is non-nil, streams tool activity to the writer for visibility.
func ApplyFix(spec *TestSpec, mismatches []Mismatch, oracleName string, failureType FailureType, fixHistory []FixRecord, verbose bool, progress io.Writer) (string, error) {
	prompt := BuildApplyFixPrompt(spec, mismatches, oracleName, failureType, fixHistory)

	if progress != nil {
		return runClaudeStreamingFix(prompt, progress)
	}

	cmd := exec.Command("claude", "-p", "--dangerously-skip-permissions")
	cmd.Stdin = strings.NewReader(prompt)
	cmd.Env = FilterClaudeEnv(os.Environ())
	if verbose {
		cmd.Stderr = os.Stderr
	}
	out, err := cmd.CombinedOutput()
	if err != nil {
		return "", fmt.Errorf("claude apply fix failed: %w\noutput: %s", err, Truncate(string(out), 500))
	}
	return string(out), nil
}

// ApplyFixRetry shells out to `claude -p` with context from a previous failed attempt.
// When progress is non-nil, streams tool activity to the writer for visibility.
func ApplyFixRetry(spec *TestSpec, mismatches []Mismatch, oracleName string, previousAttempt string, verifyOutput string, verbose bool, progress io.Writer) (string, error) {
	prompt := BuildRetryFixPrompt(spec, mismatches, oracleName, previousAttempt, verifyOutput)

	if progress != nil {
		return runClaudeStreamingFix(prompt, progress)
	}

	cmd := exec.Command("claude", "-p", "--dangerously-skip-permissions")
	cmd.Stdin = strings.NewReader(prompt)
	cmd.Env = FilterClaudeEnv(os.Environ())
	if verbose {
		cmd.Stderr = os.Stderr
	}
	out, err := cmd.CombinedOutput()
	if err != nil {
		return "", fmt.Errorf("claude apply fix retry failed: %w\noutput: %s", err, Truncate(string(out), 500))
	}
	return string(out), nil
}

// BuildRetryFixPrompt builds a retry fix prompt with context from the previous attempt.
// Distinguishes compilation errors from wrong-result failures.
func BuildRetryFixPrompt(spec *TestSpec, mismatches []Mismatch, oracleName string, previousAttempt string, verifyOutput string) string {
	base := BuildApplyFixPrompt(spec, mismatches, oracleName, "", nil)

	prevTruncated := previousAttempt
	if len(prevTruncated) > 3000 {
		prevTruncated = prevTruncated[:3000] + "..."
	}
	verifyTruncated := verifyOutput
	if len(verifyTruncated) > 3000 {
		verifyTruncated = verifyTruncated[:3000] + "..."
	}

	var sb strings.Builder
	sb.WriteString(base)
	sb.WriteString("\n\nIMPORTANT: A previous fix attempt was made but failed verification.\n\n")
	sb.WriteString("Previous attempt output:\n")
	sb.WriteString(prevTruncated)
	sb.WriteString("\n\nVerification result after previous attempt:\n")
	sb.WriteString(verifyTruncated)

	// Classify the verification failure for targeted guidance.
	if isCompilationError(verifyOutput) {
		sb.WriteString("\n\nThe previous fix FAILED TO COMPILE. Common compilation issues:\n")
		sb.WriteString("- Undefined variables or functions — check imports and function signatures\n")
		sb.WriteString("- Type mismatches — ensure return types match the expected Value type\n")
		sb.WriteString("- Missing case in switch statement — make sure the dispatch case uses the correct function name\n")
		sb.WriteString("- Syntax errors — check for missing brackets, semicolons, or parentheses\n")
	} else {
		sb.WriteString("\n\nThe previous fix COMPILED but produced WRONG RESULTS. Common issues:\n")
		sb.WriteString("- The function's argument handling didn't match Excel's conventions (1-based vs 0-based indexing, optional args)\n")
		sb.WriteString("- Edge cases weren't handled (empty cells, non-numeric args, zero, negative numbers)\n")
		sb.WriteString("- The numeric precision or rounding behavior was incorrect\n")
		sb.WriteString("- Type coercion rules differ from Excel (e.g., TRUE=1, FALSE=0, empty=0)\n")
	}

	return sb.String()
}

// isCompilationError checks if a verify output indicates a compilation error.
func isCompilationError(output string) bool {
	lower := strings.ToLower(output)
	return strings.Contains(lower, "cannot") ||
		strings.Contains(lower, "undefined") ||
		strings.Contains(lower, "syntax error") ||
		strings.Contains(lower, "imported and not used") ||
		strings.Contains(lower, "declared and not used") ||
		strings.Contains(lower, "build failed") ||
		strings.Contains(lower, "compilation")
}

// BuildApplyFixPrompt creates a prompt that instructs Claude to actually edit the code.
// Uses failure type classification and fix history for more targeted prompts.
func BuildApplyFixPrompt(spec *TestSpec, mismatches []Mismatch, oracleName string, failureType FailureType, fixHistory []FixRecord) string {
	specJSON, _ := json.MarshalIndent(spec, "", "  ")

	if oracleName == "" {
		oracleName = "libreoffice"
	}

	var sb strings.Builder
	fmt.Fprintf(&sb, "I'm testing a Go spreadsheet formula engine (github.com/jpoz/werkbook) against %s.\n\n", oracleName)
	sb.WriteString("Test spec:\n```json\n")
	sb.Write(specJSON)
	sb.WriteString("\n```\n\n")
	fmt.Fprintf(&sb, "Mismatches found (%s is the ground truth):\n", oracleName)
	for _, m := range mismatches {
		fmt.Fprintf(&sb, "- %s: werkbook=%q, %s=%q (%s)\n", m.Ref, m.Werkbook, oracleName, m.Oracle, m.Reason)
		if m.Formula != "" {
			fmt.Fprintf(&sb, "  Formula: =%s\n", m.Formula)
		}
	}

	// Failure-type-specific guidance.
	if failureType == FailureMissingFunction {
		missingFns := extractMissingFunctions(mismatches)
		sb.WriteString("\nDIAGNOSIS: Missing function implementation.\n")
		if len(missingFns) > 0 {
			fmt.Fprintf(&sb, "The following functions need to be implemented: %s\n\n", strings.Join(missingFns, ", "))
		}
		sb.WriteString("To implement a new function, you need to:\n")
		sb.WriteString("1. Add the function name to the knownFunctions array in formula/compiler.go\n")
		sb.WriteString("2. Add a dispatch case in formula/eval.go's function switch statement\n")
		sb.WriteString("3. Implement the function (e.g., fnFUNCNAME) in the appropriate formula/functions_*.go file\n\n")
		sb.WriteString("Look at similar existing implementations for patterns. For example:\n")
		sb.WriteString("- Math functions: see fnABS, fnSIN, fnROUND in formula/functions_math.go\n")
		sb.WriteString("- Text functions: see fnLEFT, fnMID in formula/functions_text.go\n")
		sb.WriteString("- Stat functions: see fnAVERAGE, fnCOUNT in formula/functions_stat.go\n")
		sb.WriteString("- Date functions: see fnDATE, fnDAY in formula/functions_date.go\n\n")
	} else if failureType == FailureBugInFunction {
		sb.WriteString("\nDIAGNOSIS: Bug in existing function implementation.\n")
		sb.WriteString("The function(s) exist but produce incorrect results.\n\n")
		sb.WriteString("Common bug patterns:\n")
		sb.WriteString("- Off-by-one errors in argument indexing\n")
		sb.WriteString("- Wrong type coercion (Excel treats TRUE as 1, empty cells as 0 in numeric context)\n")
		sb.WriteString("- Incorrect handling of optional arguments\n")
		sb.WriteString("- Precision/rounding differences\n")
		sb.WriteString("- Wrong error type returned\n\n")
	}

	sb.WriteString("The formula engine code is in the `formula/` package. Key files:\n")
	sb.WriteString("- formula/eval.go: Function dispatch (switch on function name, calls fnXXX)\n")
	sb.WriteString("- formula/functions_math.go: Math function implementations (fnABS, fnSIN, etc.)\n")
	sb.WriteString("- formula/functions_text.go: Text function implementations\n")
	sb.WriteString("- formula/functions_stat.go: Statistical function implementations\n")
	sb.WriteString("- formula/functions_logic.go: Logical function implementations\n")
	sb.WriteString("- formula/functions_lookup.go: Lookup function implementations\n")
	sb.WriteString("- formula/functions_date.go: Date function implementations\n")
	sb.WriteString("- formula/functions_info.go: Info function implementations\n")
	sb.WriteString("- formula/compiler.go: Compiles AST to bytecode\n")
	sb.WriteString("- formula/parser.go: Parses formula strings to AST\n\n")

	// Pre-include source code for broken functions to avoid tool call overhead.
	brokenForSource := extractBrokenFunctions(mismatches)
	if sourceBlock := buildSourceBlock(brokenForSource, failureType); sourceBlock != "" {
		sb.WriteString("Current source for the affected function(s) — you do NOT need to re-read these files:\n\n")
		sb.WriteString(sourceBlock)
	}

	// Include relevant fix history.
	if len(fixHistory) > 0 {
		relevantHistory := findRelevantHistory(mismatches, fixHistory)
		if len(relevantHistory) > 0 {
			sb.WriteString("Previous successful fixes for similar issues:\n")
			for _, h := range relevantHistory {
				fmt.Fprintf(&sb, "- %s (%s): %s\n", h.Function, h.FailureType, h.FixSummary)
			}
			sb.WriteString("\n")
		}
	}

	sb.WriteString("Your task: READ the relevant source files, LOOK UP the official documentation, DIAGNOSE the root cause, and EDIT the code to fix the issue.\n\n")

	// Instruct Claude to look up official Microsoft documentation for the broken function(s).
	brokenFns := extractBrokenFunctions(mismatches)
	if len(brokenFns) > 0 {
		sb.WriteString("IMPORTANT — Look up official documentation FIRST:\n")
		sb.WriteString("Before implementing or fixing any function, run the exceldoc tool to get the official Microsoft Excel specification:\n")
		for _, fn := range brokenFns {
			fmt.Fprintf(&sb, "  go run cmd/exceldoc/main.go %s\n", fn)
		}
		sb.WriteString("\nUse the documentation (Description, Syntax, Remarks, Examples) to ensure your implementation fully matches Excel's specification.\n")
		sb.WriteString("Pay close attention to:\n")
		sb.WriteString("- The exact syntax and argument order\n")
		sb.WriteString("- Which arguments are required vs optional, and their default values\n")
		sb.WriteString("- Edge case behavior described in the Remarks section\n")
		sb.WriteString("- Error conditions (e.g., #VALUE!, #NUM!, #DIV/0!) and when they are returned\n\n")
	}

	sb.WriteString("Instructions:\n")
	sb.WriteString("- If a function is missing, implement it by following the patterns in the existing code.\n")
	sb.WriteString("- Add the function implementation in the appropriate functions_*.go file.\n")
	sb.WriteString("- Add the dispatch case in formula/eval.go.\n")
	sb.WriteString("- If it's a bug in an existing function, fix the bug.\n")
	sb.WriteString("- Run `gotestsum ./formula/...` after your changes to make sure existing tests still pass.\n")
	sb.WriteString("- Do NOT modify test files or add new tests — just fix the implementation.\n")
	sb.WriteString("- Do NOT modify files outside the formula/ package.\n")
	sb.WriteString("- Keep changes minimal and focused on the fix.\n")

	return sb.String()
}

// extractMissingFunctions returns function names from #NAME? mismatches.
func extractMissingFunctions(mismatches []Mismatch) []string {
	seen := make(map[string]bool)
	var fns []string
	for _, m := range mismatches {
		if !strings.Contains(m.Werkbook, "#NAME?") {
			continue
		}
		if m.Formula != "" {
			fn := extractOutermostFunction(m.Formula)
			if fn != "" && !seen[fn] {
				seen[fn] = true
				fns = append(fns, fn)
			}
		}
	}
	return fns
}

// extractBrokenFunctions returns function names from all mismatches (both missing and buggy).
func extractBrokenFunctions(mismatches []Mismatch) []string {
	seen := make(map[string]bool)
	var fns []string
	for _, m := range mismatches {
		if m.Formula != "" {
			fn := extractOutermostFunction(m.Formula)
			if fn != "" && !seen[fn] {
				seen[fn] = true
				fns = append(fns, fn)
			}
		}
	}
	return fns
}

// findRelevantHistory finds fix records relevant to the current mismatches.
func findRelevantHistory(mismatches []Mismatch, history []FixRecord) []FixRecord {
	// Extract function names from current mismatches.
	currentFuncs := make(map[string]bool)
	for _, m := range mismatches {
		if m.Formula != "" {
			for _, fn := range ExtractFunctionsFromFormula(m.Formula) {
				currentFuncs[fn] = true
			}
		}
	}

	var relevant []FixRecord
	seen := make(map[string]bool)
	for _, h := range history {
		if currentFuncs[h.Function] && !seen[h.Function] {
			seen[h.Function] = true
			relevant = append(relevant, h)
		}
	}

	// Limit to 3 most relevant.
	if len(relevant) > 3 {
		relevant = relevant[:3]
	}
	return relevant
}

// RunClaude executes `claude -p` with the given prompt and returns stdout.
func RunClaude(prompt string) (string, error) {
	cmd := exec.Command("claude", "-p", prompt)
	// Clear CLAUDECODE env var to allow running inside a Claude Code session.
	cmd.Env = FilterClaudeEnv(os.Environ())
	out, err := cmd.CombinedOutput()
	if err != nil {
		return "", fmt.Errorf("claude command failed: %w\noutput: %s", err, Truncate(string(out), 500))
	}
	return string(out), nil
}

// BuildGenerationPrompt creates the prompt for generating a test spec.
// If eval is nil, defaults to LibreOffice behavior.
// avoid is an optional list of functions to tell Claude to skip.
// prioritize is an optional list of under-tested functions to prioritize.
func BuildGenerationPrompt(seed string, complexity *ComplexityConfig, eval Evaluator, avoid []string, prioritize []string) string {
	oracleName := "libreoffice"
	if eval != nil {
		oracleName = eval.Name()
	}
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

	// Apply complexity constraints if provided.
	if complexity != nil {
		sb.WriteString(fmt.Sprintf("Complexity level %d constraints:\n", complexity.Level))
		sb.WriteString(fmt.Sprintf("- Use %d-%d cells\n", complexity.MinCells, complexity.MaxCells))
		sb.WriteString(fmt.Sprintf("- Include %d-%d checks\n", complexity.MinChecks, complexity.MaxChecks))
		sb.WriteString(fmt.Sprintf("- Use at most %d sheets\n", complexity.MaxSheets))
		sb.WriteString(fmt.Sprintf("- Formula nesting depth: up to %d levels\n", complexity.MaxNesting))
		if complexity.CrossSheet {
			sb.WriteString("- Cross-sheet references are encouraged\n")
		} else {
			sb.WriteString("- Keep all references within the same sheet\n")
		}
		if complexity.EdgeCases == "aggressive" {
			sb.WriteString("- Aggressively test edge cases: empty cells, zero values, negative numbers, very large/small numbers, mixed types, error conditions\n")
		} else if complexity.EdgeCases == "basic" {
			sb.WriteString("- Include some edge cases: empty cells, zero values, negative numbers\n")
		}
		sb.WriteString(fmt.Sprintf("- Focus on these function categories: %s\n", strings.Join(complexity.Functions, ", ")))
		sb.WriteString("\n")
	}

	// Filter available functions based on complexity if provided.
	funcs := KnownFunctionsList
	if complexity != nil {
		funcs = complexity.FilteredFunctions()
	}

	// Split into implemented (werkbook supports) and unimplemented.
	var implemented, unimplemented []string
	for _, fn := range funcs {
		if ImplementedFunctions[fn] {
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

	if len(avoid) > 0 {
		fmt.Fprintf(&sb, "AVOID these functions (they have known bugs being worked on): %s\n\n", strings.Join(avoid, ", "))
	}

	// Coverage-guided prioritization (Task 8).
	if len(prioritize) > 0 {
		fmt.Fprintf(&sb, "PRIORITY: prefer using these under-tested functions to improve coverage: %s\n\n", strings.Join(prioritize, ", "))
	}

	sb.WriteString("Requirements:\n")
	sb.WriteString("- Use AT MOST 1 unimplemented function per spec. The rest of the functions MUST be from the implemented list. This lets us test one new function at a time against a backdrop of known-working ones.\n")
	sb.WriteString("- Create edge cases: empty cells, zero values, negative numbers, large numbers, mixed types\n")
	sb.WriteString("- Test nested formulas (e.g., IF(SUM(A1:A3)>10, AVERAGE(A1:A3), 0))\n")
	sb.WriteString("- Test cross-cell references and ranges\n")
	if complexity == nil {
		sb.WriteString("- Include 5-15 cells and 3-8 checks\n")
	}
	sb.WriteString("- Use realistic but tricky inputs that might expose bugs\n")
	sb.WriteString("- DO NOT use non-deterministic functions: RAND, RANDBETWEEN, RANDARRAY, NOW, TODAY\n")
	sb.WriteString("- DO NOT use INDIRECT (volatile/indirect resolution)\n")

	// Build excluded functions list from the evaluator.
	var excludedList []string
	var excl map[string]bool
	if eval != nil {
		excl = eval.ExcludedFunctions()
	} else {
		excl = ExcludedFunctions
	}
	for fn := range excl {
		excludedList = append(excludedList, fn)
	}
	if len(excludedList) > 0 {
		fmt.Fprintf(&sb, "- DO NOT use these excluded functions: %s\n", strings.Join(excludedList, ", "))
	}

	sb.WriteString("- DO NOT use environment-dependent functions: CELL, INFO\n")
	if oracleName == "libreoffice" {
		sb.WriteString("- Use CONCATENATE instead of CONCAT\n")
	}
	fmt.Fprintf(&sb, "- The formulas will be validated against %s, so only use functions that %s supports\n", oracleName, oracleName)
	sb.WriteString("- The 'expected' field in checks should be the string representation of the expected result\n")
	sb.WriteString("- For numbers, use the simplest representation (e.g., \"10\" not \"10.0\")\n")
	sb.WriteString("- For booleans, use \"TRUE\" or \"FALSE\"\n")
	sb.WriteString("- Cell refs in formulas should NOT include the sheet name if on the same sheet\n")
	sb.WriteString("- Multi-sheet tests are welcome but keep them simple\n")
	sb.WriteString("\n")
	sb.WriteString("Output ONLY the JSON, no explanation or markdown fences.\n")

	return sb.String()
}

// BuildTargetedGenerationPrompt creates a prompt focused on testing one specific function.
func BuildTargetedGenerationPrompt(targetFn string, complexity *ComplexityConfig, eval Evaluator) string {
	oracleName := "libreoffice"
	if eval != nil {
		oracleName = eval.Name()
	}

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

	fmt.Fprintf(&sb, "Generate a test spec that specifically tests the %s function.\n", targetFn)
	fmt.Fprintf(&sb, "Include 2-3 checks that exercise %s with different inputs.\n", targetFn)
	sb.WriteString("Also include 2-3 checks using known-working functions as a control group.\n\n")

	var implemented []string
	for _, fn := range KnownFunctionsList {
		if ImplementedFunctions[fn] {
			implemented = append(implemented, fn)
		}
	}

	sb.WriteString("Implemented functions (werkbook already supports these - use for control group checks):\n")
	sb.WriteString(strings.Join(implemented, ", "))
	sb.WriteString("\n\n")

	fmt.Fprintf(&sb, "The target function %s is NOT yet implemented. Test it thoroughly.\n\n", targetFn)

	if complexity != nil {
		sb.WriteString(fmt.Sprintf("Complexity level %d constraints:\n", complexity.Level))
		sb.WriteString(fmt.Sprintf("- Use %d-%d cells\n", complexity.MinCells, complexity.MaxCells))
		sb.WriteString(fmt.Sprintf("- Include %d-%d checks\n", complexity.MinChecks, complexity.MaxChecks))
		sb.WriteString("\n")
	}

	sb.WriteString("Requirements:\n")
	fmt.Fprintf(&sb, "- The spec MUST use %s in at least 2-3 formula cells with different argument patterns\n", targetFn)
	sb.WriteString("- Include 2-3 additional checks using implemented functions as a control group\n")
	sb.WriteString("- Create edge cases: empty cells, zero values, negative numbers, large numbers, mixed types\n")
	sb.WriteString("- Include 5-15 cells and 4-6 checks total\n")
	sb.WriteString("- Use realistic but tricky inputs that might expose bugs\n")
	sb.WriteString("- DO NOT use non-deterministic functions: RAND, RANDBETWEEN, RANDARRAY, NOW, TODAY\n")
	sb.WriteString("- DO NOT use INDIRECT (volatile/indirect resolution)\n")

	var excludedList []string
	var excl map[string]bool
	if eval != nil {
		excl = eval.ExcludedFunctions()
	} else {
		excl = ExcludedFunctions
	}
	for fn := range excl {
		excludedList = append(excludedList, fn)
	}
	if len(excludedList) > 0 {
		fmt.Fprintf(&sb, "- DO NOT use these excluded functions: %s\n", strings.Join(excludedList, ", "))
	}

	sb.WriteString("- DO NOT use environment-dependent functions: CELL, INFO\n")
	if oracleName == "libreoffice" {
		sb.WriteString("- Use CONCATENATE instead of CONCAT\n")
	}
	fmt.Fprintf(&sb, "- The formulas will be validated against %s, so only use functions that %s supports\n", oracleName, oracleName)
	sb.WriteString("- The 'expected' field in checks should be the string representation of the expected result\n")
	sb.WriteString("- For numbers, use the simplest representation (e.g., \"10\" not \"10.0\")\n")
	sb.WriteString("- For booleans, use \"TRUE\" or \"FALSE\"\n")
	sb.WriteString("- Cell refs in formulas should NOT include the sheet name if on the same sheet\n")
	sb.WriteString("\n")
	sb.WriteString("Output ONLY the JSON, no explanation or markdown fences.\n")

	return sb.String()
}

// BuildGenerateTestsPrompt creates a prompt asking Claude to add unit tests for
// functions that were just fixed. The prompt constrains edits to test files only.
func BuildGenerateTestsPrompt(spec *TestSpec, mismatches []Mismatch, oracleName string, fixedFunctions []string) string {
	specJSON, _ := json.MarshalIndent(spec, "", "  ")

	if oracleName == "" {
		oracleName = "libreoffice"
	}

	var sb strings.Builder
	fmt.Fprintf(&sb, "I just fixed the Go spreadsheet formula engine (github.com/jpoz/werkbook) so it matches %s.\n\n", oracleName)
	sb.WriteString("The fix was based on this test spec:\n```json\n")
	sb.Write(specJSON)
	sb.WriteString("\n```\n\n")

	fmt.Fprintf(&sb, "Mismatches that were fixed (%s was the ground truth):\n", oracleName)
	for _, m := range mismatches {
		fmt.Fprintf(&sb, "- %s: werkbook=%q, %s=%q (%s)\n", m.Ref, m.Werkbook, oracleName, m.Oracle, m.Reason)
		if m.Formula != "" {
			fmt.Fprintf(&sb, "  Formula: =%s\n", m.Formula)
		}
	}

	fmt.Fprintf(&sb, "\nFixed functions: %s\n\n", strings.Join(fixedFunctions, ", "))

	// Instruct Claude to look up official documentation for thorough test coverage.
	if len(fixedFunctions) > 0 {
		sb.WriteString("IMPORTANT — Look up official documentation to write thorough tests:\n")
		sb.WriteString("Before writing tests, run the exceldoc tool to get the full Microsoft Excel specification for each function:\n")
		for _, fn := range fixedFunctions {
			fmt.Fprintf(&sb, "  go run cmd/exceldoc/main.go %s\n", fn)
		}
		sb.WriteString("\nUse the documentation to write comprehensive tests that cover the FULL specification, not just the mismatches above.\n\n")
	}

	sb.WriteString("Your task: Add thorough unit test cases for the fixed function(s) to lock in correct behavior.\n\n")
	sb.WriteString("Instructions:\n")
	sb.WriteString("- Read the existing test files in the formula/ package to understand the test patterns used.\n")
	sb.WriteString("- The tests use table-driven patterns (e.g., numTests in functions_math_test.go, textTests in functions_text_test.go, etc.).\n")
	sb.WriteString("- Add test cases to the EXISTING table in the appropriate formula/functions_*_test.go file.\n")
	sb.WriteString("- Each test case should cover the scenarios from the mismatches above.\n")
	sb.WriteString("- Go BEYOND the mismatches — use the official documentation to test the full behavior:\n")
	sb.WriteString("  - Every documented argument combination (required args, optional args, default values)\n")
	sb.WriteString("  - All examples from the documentation's Example section\n")
	sb.WriteString("  - Edge cases from the Remarks section (special values, boundary conditions)\n")
	sb.WriteString("  - Error conditions: what inputs should produce #VALUE!, #NUM!, #DIV/0!, #N/A, #REF!\n")
	sb.WriteString("  - Type coercion: booleans as numbers (TRUE=1, FALSE=0), strings that look like numbers, empty cells\n")
	sb.WriteString("  - Boundary values: zero, negative numbers, very large/small numbers, empty strings, max length strings\n")
	sb.WriteString("  - If the function takes a range, test single-cell ranges, multi-cell ranges, and ranges with mixed types\n")
	sb.WriteString("- Aim for at least 10-15 test cases per function to ensure solid coverage.\n")
	sb.WriteString("- You may ONLY edit formula/*_test.go files. Do NOT modify implementation files.\n")
	sb.WriteString("- Do NOT modify files outside the formula/ package.\n")
	sb.WriteString("- Run `gotestsum ./formula/...` after your changes to confirm the tests pass.\n")

	return sb.String()
}

// GenerateTests shells out to `claude -p --dangerously-skip-permissions` to generate
// unit tests for functions that were just fixed.
// When progress is non-nil, streams tool activity to the writer for visibility.
func GenerateTests(spec *TestSpec, mismatches []Mismatch, oracleName string, fixedFunctions []string, verbose bool, progress io.Writer) (string, error) {
	prompt := BuildGenerateTestsPrompt(spec, mismatches, oracleName, fixedFunctions)

	if progress != nil {
		return runClaudeStreamingFix(prompt, progress)
	}

	cmd := exec.Command("claude", "-p", "--dangerously-skip-permissions")
	cmd.Stdin = strings.NewReader(prompt)
	cmd.Env = FilterClaudeEnv(os.Environ())
	if verbose {
		cmd.Stderr = os.Stderr
	}
	out, err := cmd.CombinedOutput()
	if err != nil {
		return "", fmt.Errorf("claude generate tests failed: %w\noutput: %s", err, Truncate(string(out), 500))
	}
	return string(out), nil
}

// BuildFixPrompt creates the prompt for suggesting a fix.
func BuildFixPrompt(spec *TestSpec, mismatches []Mismatch, oracleName string) string {
	specJSON, _ := json.MarshalIndent(spec, "", "  ")

	if oracleName == "" {
		oracleName = "libreoffice"
	}

	var sb strings.Builder
	fmt.Fprintf(&sb, "I'm testing a Go spreadsheet formula engine (github.com/jpoz/werkbook) against %s.\n\n", oracleName)
	sb.WriteString("Test spec:\n```json\n")
	sb.Write(specJSON)
	sb.WriteString("\n```\n\n")
	fmt.Fprintf(&sb, "Mismatches found (%s is the ground truth):\n", oracleName)
	for _, m := range mismatches {
		fmt.Fprintf(&sb, "- %s: werkbook=%q, %s=%q (%s)\n", m.Ref, m.Werkbook, oracleName, m.Oracle, m.Reason)
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

// ExtractJSON finds the first JSON object in the output.
func ExtractJSON(s string) string {
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

// Truncate shortens a string to maxLen, adding "..." if truncated.
func Truncate(s string, maxLen int) string {
	if len(s) <= maxLen {
		return s
	}
	return s[:maxLen] + "..."
}

// FilterClaudeEnv returns a copy of env with all Claude Code session
// variables removed so that subprocess invocations of the claude CLI
// don't detect a nested session and refuse to start.
func FilterClaudeEnv(env []string) []string {
	out := make([]string, 0, len(env))
	for _, e := range env {
		if strings.HasPrefix(e, "CLAUDECODE=") ||
			strings.HasPrefix(e, "CLAUDE_CODE_") {
			continue
		}
		out = append(out, e)
	}
	return out
}
