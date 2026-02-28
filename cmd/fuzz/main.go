package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/jpoz/werkbook/internal/fuzz"
)

// SystematicState tracks progress through the unimplemented function queue.
type SystematicState struct {
	Queue   []string // functions to test
	Current int      // index in queue
	Phase   string   // "targeted" or "regression"
}

func main() {
	rounds := flag.Int("rounds", 10, "number of fuzz rounds")
	fix := flag.Bool("fix", false, "auto-suggest fixes with claude -p")
	specFile := flag.String("spec", "", "run a single saved spec instead of generating")
	keep := flag.Bool("keep", false, "keep generated files")
	seed := flag.String("seed", "", "focus category: math|text|lookup|logic|stat|date|info|financial|engineering|database|mixed")
	oracle := flag.String("oracle", "libreoffice", "evaluation oracle: libreoffice or excel")
	systematic := flag.Bool("systematic", false, "systematically test unimplemented functions in order")
	coverageInterval := flag.Int("coverage-interval", 10, "print coverage summary every N rounds (0 to disable)")
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

	// Create function coverage tracker.
	coverage := fuzz.NewFunctionCoverage()

	// Single spec mode.
	if *specFile != "" {
		spec, err := loadSpec(*specFile)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error loading spec: %v\n", err)
			os.Exit(1)
		}
		ok := runSpec(spec, eval, coverage, failDir, *fix, *keep, *verbose)
		fmt.Println()
		fmt.Print(coverage.Summary())
		if !ok {
			os.Exit(1)
		}
		return
	}

	// Systematic mode.
	if *systematic {
		runSystematic(*rounds, eval, coverage, failDir, *fix, *keep, *coverageInterval, *verbose)
		return
	}

	// Multi-round fuzz mode.
	passed, failed := 0, 0
	var totalTime time.Duration
	for i := 0; i < *rounds; i++ {
		fmt.Printf("\n=== Round %d/%d ===\n", i+1, *rounds)

		roundStart := time.Now()

		fmt.Printf("  Generating spec with Claude...")
		genStart := time.Now()
		spec, err := generateSpec(seedCategory(*seed), *verbose)
		genElapsed := time.Since(genStart)
		if err != nil {
			fmt.Printf(" failed (%.1fs)\n", genElapsed.Seconds())
			fmt.Printf("  Generation failed: %v\n", err)
			failed++
			roundElapsed := time.Since(roundStart)
			totalTime += roundElapsed
			fmt.Printf("  Round time: %.1fs\n", roundElapsed.Seconds())
			continue
		}
		fmt.Printf(" ok (%.1fs)\n", genElapsed.Seconds())

		fmt.Printf("  Spec: %s (%d sheets, %d checks)\n", spec.Name, len(spec.Sheets), len(spec.Checks))

		if runSpec(spec, eval, coverage, failDir, *fix, *keep, *verbose) {
			passed++
		} else {
			failed++
		}

		roundElapsed := time.Since(roundStart)
		totalTime += roundElapsed
		fmt.Printf("  Round time: %.1fs\n", roundElapsed.Seconds())

		generated := passed + failed
		avg := totalTime / time.Duration(generated)
		fmt.Printf("  Generated: %d | Passed: %d | Failed: %d | Avg: %.1fs/round\n",
			generated, passed, failed, avg.Seconds())

		// Print coverage summary at the configured interval.
		if *coverageInterval > 0 && (i+1)%*coverageInterval == 0 {
			fmt.Println()
			fmt.Print(coverage.Summary())
		}
	}

	fmt.Printf("\n=== Summary ===\n")
	fmt.Printf("Passed: %d  Failed: %d  Total: %d\n", passed, failed, passed+failed)
	if passed+failed > 0 {
		avg := totalTime / time.Duration(passed+failed)
		fmt.Printf("Total time: %.1fs  Avg: %.1fs/round\n", totalTime.Seconds(), avg.Seconds())
	}
	fmt.Println()
	fmt.Print(coverage.Summary())
	if failed > 0 {
		os.Exit(1)
	}
}

// runSystematic runs the systematic function targeting mode.
func runSystematic(maxRounds int, eval fuzz.Evaluator, coverage *fuzz.FunctionCoverage, failDir string, fix, keep bool, coverageInterval int, verbose bool) {
	state := &SystematicState{
		Queue:   buildUnimplementedQueue(),
		Current: 0,
		Phase:   "targeted",
	}

	totalUnimplemented := len(state.Queue)
	fmt.Printf("Systematic mode: %d unimplemented functions to test\n", totalUnimplemented)

	passed, failed, fixed := 0, 0, 0
	var totalTime time.Duration
	for i := 0; i < maxRounds; i++ {
		roundStart := time.Now()

		if state.Current < len(state.Queue) {
			// Targeted phase: test the next unimplemented function.
			targetFn := state.Queue[state.Current]
			remaining := len(state.Queue) - state.Current
			fmt.Printf("\n=== Round %d (targeting %s, %d remaining) ===\n", i+1, targetFn, remaining)

			fmt.Printf("  Generating spec for %s...", targetFn)
			genStart := time.Now()
			spec, err := generateTargetedSpec(targetFn, verbose)
			genElapsed := time.Since(genStart)
			if err != nil {
				fmt.Printf(" failed (%.1fs)\n", genElapsed.Seconds())
				fmt.Printf("  Generation failed: %v\n", err)
				failed++
				state.Current++
				roundElapsed := time.Since(roundStart)
				totalTime += roundElapsed
				fmt.Printf("  Round time: %.1fs\n", roundElapsed.Seconds())
				continue
			}
			fmt.Printf(" ok (%.1fs)\n", genElapsed.Seconds())

			fmt.Printf("  Spec: %s (%d sheets, %d checks)\n", spec.Name, len(spec.Sheets), len(spec.Checks))

			ok := runSpec(spec, eval, coverage, failDir, fix, keep, verbose)
			if ok {
				fmt.Printf("  PASS - %s works correctly\n", targetFn)
				passed++
			} else {
				fmt.Printf("  FAIL - %s has mismatches\n", targetFn)
				failed++
			}

			state.Current++

			// After a fix attempt, re-scan implemented functions to see if the
			// queue should be updated (the fix may have added functions).
			if fix && !ok {
				oldLen := len(state.Queue)
				newQueue := buildUnimplementedQueue()
				if len(newQueue) < oldLen {
					fixed += oldLen - len(newQueue)
					fmt.Printf("  Queue updated: %d functions newly implemented\n", oldLen-len(newQueue))
					state.Queue = newQueue
					if state.Current > len(state.Queue) {
						state.Current = len(state.Queue)
					}
				}
			}
		} else {
			// Regression phase: random generation.
			if state.Phase != "regression" {
				state.Phase = "regression"
				fmt.Printf("\n--- All unimplemented functions tested, switching to regression mode ---\n")
			}
			fmt.Printf("\n=== Round %d (regression mode, random) ===\n", i+1)

			fmt.Printf("  Generating spec with Claude...")
			genStart := time.Now()
			spec, err := generateSpec("", verbose)
			genElapsed := time.Since(genStart)
			if err != nil {
				fmt.Printf(" failed (%.1fs)\n", genElapsed.Seconds())
				fmt.Printf("  Generation failed: %v\n", err)
				failed++
				roundElapsed := time.Since(roundStart)
				totalTime += roundElapsed
				fmt.Printf("  Round time: %.1fs\n", roundElapsed.Seconds())
				continue
			}
			fmt.Printf(" ok (%.1fs)\n", genElapsed.Seconds())

			fmt.Printf("  Spec: %s (%d sheets, %d checks)\n", spec.Name, len(spec.Sheets), len(spec.Checks))

			if runSpec(spec, eval, coverage, failDir, fix, keep, verbose) {
				passed++
			} else {
				failed++
			}
		}

		roundElapsed := time.Since(roundStart)
		totalTime += roundElapsed
		fmt.Printf("  Round time: %.1fs\n", roundElapsed.Seconds())

		// Print coverage summary at the configured interval.
		if coverageInterval > 0 && (i+1)%coverageInterval == 0 {
			fmt.Println()
			fmt.Print(coverage.Summary())
		}
	}

	fmt.Printf("\n=== Summary ===\n")
	fmt.Printf("Passed: %d  Failed: %d  Fixed: %d  Total: %d\n", passed, failed, fixed, passed+failed)
	fmt.Printf("Functions tested: %d/%d unimplemented\n", state.Current, totalUnimplemented)
	if state.Phase == "regression" {
		fmt.Printf("Regression rounds: %d\n", maxRounds-totalUnimplemented)
	}
	if passed+failed > 0 {
		avg := totalTime / time.Duration(passed+failed)
		fmt.Printf("Total time: %.1fs  Avg: %.1fs/round\n", totalTime.Seconds(), avg.Seconds())
	}
	fmt.Println()
	fmt.Print(coverage.Summary())
	if failed > 0 {
		os.Exit(1)
	}
}

// runSpec executes a single test spec: build, eval with oracle, compare.
// Returns true if all checks pass.
func runSpec(spec *TestSpec, eval fuzz.Evaluator, coverage *fuzz.FunctionCoverage, failDir string, fix, keep, verbose bool) bool {
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
	fmt.Printf("  Building XLSX...")
	buildStart := time.Now()
	xlsxPath, wbResults, err := buildXLSX(spec, tmpDir)
	buildElapsed := time.Since(buildStart)
	if err != nil {
		fmt.Printf(" failed (%.1fs)\n", buildElapsed.Seconds())
		fmt.Printf("  Build failed: %v\n", err)
		return false
	}
	fmt.Printf(" ok (%.1fs)\n", buildElapsed.Seconds())

	if verbose {
		fmt.Printf("  XLSX: %s\n", xlsxPath)
		fmt.Printf("  Werkbook results:\n")
		for _, r := range wbResults {
			fmt.Printf("    %s = %q (%s)\n", r.Ref, r.Value, r.Type)
		}
	}

	// Evaluate with oracle.
	fmt.Printf("  Evaluating with %s...", eval.Name())
	evalStart := time.Now()
	fuzzChecks := toFuzzChecks(spec.Checks)
	oracleResults, err := eval.Eval(xlsxPath, fuzzChecks)
	evalElapsed := time.Since(evalStart)
	if err != nil {
		fmt.Printf(" failed (%.1fs)\n", evalElapsed.Seconds())
		fmt.Printf("  %s eval failed: %v\n", eval.Name(), err)
		return false
	}
	fmt.Printf(" ok (%.1fs)\n", evalElapsed.Seconds())
	loResults := fromFuzzResults(oracleResults)

	if verbose {
		fmt.Printf("  %s results:\n", eval.Name())
		for _, r := range loResults {
			fmt.Printf("    %s = %q (%s)\n", r.Ref, r.Value, r.Type)
		}
	}

	// Compare results.
	fmt.Printf("  Comparing results...")
	compareStart := time.Now()
	mismatches := compareResults(spec.Checks, wbResults, loResults, spec)
	compareElapsed := time.Since(compareStart)

	// Extract functions from the spec and record coverage.
	funcs := fuzz.ExtractFunctions(toFuzzSpec(spec))
	if len(funcs) > 0 {
		fmt.Printf(" done (%.1fs)\n", compareElapsed.Seconds())
		fmt.Printf("  Functions tested: %s\n", strings.Join(funcs, ", "))
	} else {
		fmt.Printf(" done (%.1fs)\n", compareElapsed.Seconds())
	}
	coverage.Record(funcs, len(mismatches) == 0)

	if len(mismatches) == 0 {
		fmt.Printf("  PASS: all %d checks match\n", len(spec.Checks))
		return true
	}

	fmt.Printf("  %d/%d checks mismatched:\n", len(mismatches), len(spec.Checks))
	fmt.Print(formatMismatches(mismatches, eval.Name()))

	// Save failure for reproduction.
	saveFailure(spec, xlsxPath, failDir)

	// Auto-fix mode.
	if fix {
		fmt.Printf("  Applying fix with Claude...")
		fixStart := time.Now()
		suggestion, err := suggestFix(spec, mismatches, verbose)
		fixElapsed := time.Since(fixStart)
		if err != nil {
			fmt.Printf(" failed (%.1fs)\n", fixElapsed.Seconds())
			fmt.Printf("  Fix suggestion failed: %v\n", err)
		} else {
			fmt.Printf(" ok (%.1fs)\n", fixElapsed.Seconds())
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

// toFuzzSpec converts a local TestSpec to a fuzz.TestSpec for use with
// fuzz package functions like ExtractFunctions.
func toFuzzSpec(spec *TestSpec) *fuzz.TestSpec {
	fs := &fuzz.TestSpec{
		Name: spec.Name,
	}
	for _, s := range spec.Sheets {
		sheet := fuzz.SheetSpec{Name: s.Name}
		for _, c := range s.Cells {
			sheet.Cells = append(sheet.Cells, fuzz.CellSpec{
				Ref:     c.Ref,
				Value:   c.Value,
				Type:    c.Type,
				Formula: c.Formula,
			})
		}
		fs.Sheets = append(fs.Sheets, sheet)
	}
	for _, c := range spec.Checks {
		fs.Checks = append(fs.Checks, fuzz.CheckSpec{
			Ref:      c.Ref,
			Expected: c.Expected,
			Type:     c.Type,
		})
	}
	return fs
}
