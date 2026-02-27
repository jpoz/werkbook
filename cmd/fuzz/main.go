package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/jpoz/werkbook/internal/fuzz"
)

func main() {
	rounds := flag.Int("rounds", 10, "number of fuzz rounds")
	fix := flag.Bool("fix", false, "auto-suggest fixes with claude -p")
	specFile := flag.String("spec", "", "run a single saved spec instead of generating")
	keep := flag.Bool("keep", false, "keep generated files")
	seed := flag.String("seed", "", "focus category: math|text|lookup|logic|stat|date|info|financial|engineering|database|mixed")
	oracle := flag.String("oracle", "libreoffice", "evaluation oracle: libreoffice or excel")
	verbose := flag.Bool("v", false, "verbose output")
	flag.Parse()

	// Create evaluator.
	eval, err := fuzz.NewEvaluator(*oracle)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	if *verbose {
		fmt.Printf("Using oracle: %s\n", eval.Name())
	}

	// Ensure failures directory exists.
	failDir := "failures"
	if err := os.MkdirAll(failDir, 0755); err != nil {
		fmt.Fprintf(os.Stderr, "Error creating failures dir: %v\n", err)
		os.Exit(1)
	}

	// Single spec mode.
	if *specFile != "" {
		spec, err := loadSpec(*specFile)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error loading spec: %v\n", err)
			os.Exit(1)
		}
		ok := runSpec(spec, eval, failDir, *fix, *keep, *verbose)
		if !ok {
			os.Exit(1)
		}
		return
	}

	// Multi-round fuzz mode.
	passed, failed := 0, 0
	for i := 0; i < *rounds; i++ {
		fmt.Printf("\n=== Round %d/%d ===\n", i+1, *rounds)

		spec, err := generateSpec(seedCategory(*seed), *verbose)
		if err != nil {
			fmt.Printf("  Generation failed: %v\n", err)
			failed++
			continue
		}

		fmt.Printf("  Spec: %s (%d sheets, %d checks)\n", spec.Name, len(spec.Sheets), len(spec.Checks))

		if runSpec(spec, eval, failDir, *fix, *keep, *verbose) {
			passed++
		} else {
			failed++
		}
	}

	fmt.Printf("\n=== Summary ===\n")
	fmt.Printf("Passed: %d  Failed: %d  Total: %d\n", passed, failed, passed+failed)
	if failed > 0 {
		os.Exit(1)
	}
}

// runSpec executes a single test spec: build, eval with oracle, compare.
// Returns true if all checks pass.
func runSpec(spec *TestSpec, eval fuzz.Evaluator, failDir string, fix, keep, verbose bool) bool {
	// Create a temp dir for this spec.
	tmpDir, err := os.MkdirTemp("", "werkbook-fuzz-*")
	if err != nil {
		fmt.Printf("  Error creating temp dir: %v\n", err)
		return false
	}
	if !keep {
		defer os.RemoveAll(tmpDir)
	} else {
		fmt.Printf("  Working dir: %s\n", tmpDir)
	}

	// Build XLSX and get werkbook results.
	xlsxPath, wbResults, err := buildXLSX(spec, tmpDir)
	if err != nil {
		fmt.Printf("  Build failed: %v\n", err)
		return false
	}

	if verbose {
		fmt.Printf("  XLSX: %s\n", xlsxPath)
		fmt.Printf("  Werkbook results:\n")
		for _, r := range wbResults {
			fmt.Printf("    %s = %q (%s)\n", r.Ref, r.Value, r.Type)
		}
	}

	// Evaluate with oracle.
	fuzzChecks := toFuzzChecks(spec.Checks)
	oracleResults, err := eval.Eval(xlsxPath, fuzzChecks)
	if err != nil {
		fmt.Printf("  %s eval failed: %v\n", eval.Name(), err)
		return false
	}
	loResults := fromFuzzResults(oracleResults)

	if verbose {
		fmt.Printf("  %s results:\n", eval.Name())
		for _, r := range loResults {
			fmt.Printf("    %s = %q (%s)\n", r.Ref, r.Value, r.Type)
		}
	}

	// Compare results.
	mismatches := compareResults(spec.Checks, wbResults, loResults)
	if len(mismatches) == 0 {
		fmt.Printf("  PASS: all %d checks match\n", len(spec.Checks))
		return true
	}

	fmt.Printf("  FAIL: %d/%d checks mismatched:\n", len(mismatches), len(spec.Checks))
	fmt.Print(formatMismatches(mismatches, eval.Name()))

	// Save failure for reproduction.
	saveFailure(spec, xlsxPath, failDir)

	// Auto-fix mode.
	if fix {
		suggestion, err := suggestFix(spec, mismatches, verbose)
		if err != nil {
			fmt.Printf("  Fix suggestion failed: %v\n", err)
		} else {
			fmt.Printf("\n  Fix suggestion:\n%s\n", indent(suggestion, "    "))
		}
	}

	return false
}

// saveFailure persists a failing spec and its XLSX for later reproduction.
func saveFailure(spec *TestSpec, xlsxPath, failDir string) {
	specPath := filepath.Join(failDir, spec.Name+".json")
	if err := saveSpec(specPath, spec); err != nil {
		fmt.Printf("  Warning: could not save spec: %v\n", err)
		return
	}

	// Copy XLSX to failures dir.
	xlsxDst := filepath.Join(failDir, spec.Name+".xlsx")
	data, err := os.ReadFile(xlsxPath)
	if err == nil {
		os.WriteFile(xlsxDst, data, 0644)
	}

	fmt.Printf("  Saved: %s\n", specPath)
}

// seedCategory maps user-provided seed to a prompt category hint.
func seedCategory(seed string) string {
	switch strings.ToLower(seed) {
	case "math":
		return "math (ABS, CEILING, CEILING.MATH, COMBIN, COS, COSH, DEGREES, EVEN, EXP, FACT, FACTDOUBLE, FLOOR, FLOOR.MATH, GCD, INT, LCM, LN, LOG, LOG10, MDETERM, MOD, MROUND, MULTINOMIAL, ODD, PI, POWER, PRODUCT, QUOTIENT, RADIANS, ROUND, ROUNDDOWN, ROUNDUP, SIGN, SIN, SINH, SQRT, SQRTPI, SUBTOTAL, SUM, SUMIF, SUMIFS, SUMPRODUCT, SUMSQ, TAN, TANH, TRUNC)"
	case "text":
		return "text (CHAR, CLEAN, CODE, CONCATENATE, DOLLAR, EXACT, FIND, FIXED, LEFT, LEN, LOWER, MID, NUMBERVALUE, PROPER, REPLACE, REPT, RIGHT, SEARCH, SUBSTITUTE, T, TEXT, TEXTJOIN, TRIM, UNICHAR, UNICODE, UPPER, VALUE)"
	case "lookup":
		return "lookup and reference (ADDRESS, AREAS, CHOOSE, COLUMN, COLUMNS, FORMULATEXT, HLOOKUP, INDEX, LOOKUP, MATCH, ROW, ROWS, TRANSPOSE, VLOOKUP)"
	case "logic":
		return "logical (AND, FALSE, IF, IFERROR, IFNA, IFS, ISBLANK, ISERR, ISERROR, ISNA, ISNUMBER, ISTEXT, NOT, OR, SWITCH, TRUE, XOR)"
	case "stat":
		return "statistical (AVEDEV, AVERAGE, AVERAGEA, AVERAGEIF, AVERAGEIFS, BINOM.DIST, CHISQ.DIST, CONFIDENCE.NORM, CORREL, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, COVARIANCE.P, EXPON.DIST, F.DIST, FISHER, FORECAST.LINEAR, GAMMA.DIST, GEOMEAN, HARMEAN, LARGE, MAX, MAXA, MAXIFS, MEDIAN, MIN, MINA, MINIFS, MODE.SNGL, NORM.DIST, NORM.INV, PEARSON, PERCENTILE.INC, PERMUT, POISSON.DIST, RANK.AVG, RANK.EQ, SMALL, STDEV.P, STDEV.S, T.DIST, TRIMMEAN, VAR.P, VAR.S, WEIBULL.DIST, Z.TEST)"
	case "date":
		return "date/time (DATE, DATEDIF, DATEVALUE, DAY, DAYS, DAYS360, EDATE, EOMONTH, HOUR, ISOWEEKNUM, MINUTE, MONTH, NETWORKDAYS, SECOND, TIME, TIMEVALUE, WEEKDAY, WEEKNUM, WORKDAY, YEAR, YEARFRAC)"
	case "info":
		return "information (ERROR.TYPE, ISBLANK, ISERR, ISERROR, ISEVEN, ISFORMULA, ISLOGICAL, ISNA, ISNONTEXT, ISNUMBER, ISODD, ISREF, ISTEXT, N, NA, TYPE)"
	case "financial":
		return "financial (ACCRINT, CUMIPMT, CUMPRINC, DB, DDB, DISC, DOLLARDE, DOLLARFR, DURATION, EFFECT, FV, FVSCHEDULE, INTRATE, IPMT, IRR, ISPMT, MDURATION, MIRR, NOMINAL, NPER, NPV, PDURATION, PMT, PPMT, PRICE, PRICEDISC, PV, RATE, RECEIVED, RRI, SLN, SYD, TBILLEQ, TBILLPRICE, TBILLYIELD, VDB, XIRR, XNPV, YIELD, YIELDDISC, YIELDMAT)"
	case "engineering":
		return "engineering (BESSELI, BESSELJ, BIN2DEC, BIN2HEX, BIN2OCT, BITAND, BITLSHIFT, BITOR, BITRSHIFT, BITXOR, COMPLEX, CONVERT, DEC2BIN, DEC2HEX, DEC2OCT, DELTA, ERF, ERFC, GESTEP, HEX2BIN, HEX2DEC, HEX2OCT, IMABS, IMAGINARY, IMARGUMENT, IMCONJUGATE, IMCOS, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMREAL, IMSIN, IMSQRT, IMSUB, IMSUM, OCT2BIN, OCT2DEC, OCT2HEX)"
	case "database":
		return "database (DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV, DSTDEVP, DSUM, DVAR, DVARP)"
	case "mixed", "":
		return ""
	default:
		return seed
	}
}

// toFuzzChecks converts local CheckSpec to fuzz.CheckSpec.
func toFuzzChecks(checks []CheckSpec) []fuzz.CheckSpec {
	out := make([]fuzz.CheckSpec, len(checks))
	for i, c := range checks {
		out[i] = fuzz.CheckSpec{
			Ref:      c.Ref,
			Expected: c.Expected,
			Type:     c.Type,
		}
	}
	return out
}

// fromFuzzResults converts fuzz.CellResult to local buildResult.
func fromFuzzResults(results []fuzz.CellResult) []buildResult {
	out := make([]buildResult, len(results))
	for i, r := range results {
		out[i] = buildResult{
			Ref:   r.Ref,
			Value: r.Value,
			Type:  r.Type,
		}
	}
	return out
}

// indent prefixes each line with the given prefix.
func indent(s, prefix string) string {
	lines := strings.Split(s, "\n")
	for i, line := range lines {
		if line != "" {
			lines[i] = prefix + line
		}
	}
	return strings.Join(lines, "\n")
}

