package main

import (
	"flag"
	"fmt"
	"os"
	"os/exec"
	"os/signal"
	"path/filepath"
	"sort"
	"strings"
	"time"

	"github.com/jpoz/werkbook/internal/fuzz"
)

// OrchestratorState tracks the orchestrator's progress.
type OrchestratorState struct {
	Level             int
	ConsecutivePasses int
	TotalGenerated    int
	TotalPassed       int
	TotalFixed        int
	TotalFailed       int
	TotalRetries      int
	FixedOnRetry      int
	TotalTime         time.Duration
	Coverage          *fuzz.FunctionCoverage
}

// SystematicState tracks progress through the unimplemented function queue.
type SystematicState struct {
	Queue   []string // functions to test
	Current int      // index in queue
	Phase   string   // "targeted" or "regression"
}

func main() {
	startLevel := flag.Int("start-level", 1, "starting complexity level")
	passesToEscalate := flag.Int("passes-to-escalate", 3, "consecutive passes before escalating")
	maxRounds := flag.Int("max-rounds", 0, "maximum number of rounds (0 = unlimited)")
	maxFixAttempts := flag.Int("max-fix-attempts", 2, "maximum fix attempts per round")
	coverageInterval := flag.Int("coverage-interval", 10, "print coverage summary every N rounds (0 to disable)")
	systematic := flag.Bool("systematic", false, "systematically test unimplemented functions in order")
	replay := flag.Bool("replay", false, "replay saved failures from failures/ directory before main loop")
	mutate := flag.Bool("mutate", false, "enable mutation-based spec generation (alternates with Claude generation)")
	seed := flag.String("seed", "", "focus category: math|text|lookup|logic|stat|date|info|financial|engineering|database|mixed")
	oracle := flag.String("oracle", "libreoffice", "validation oracle: libreoffice or excel")
	verbose := flag.Bool("v", false, "verbose output")
	flag.Parse()

	if err := run(*startLevel, *passesToEscalate, *maxRounds, *maxFixAttempts, *coverageInterval, *systematic, *replay, *mutate, *seed, *oracle, *verbose); err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
}

func run(startLevel, passesToEscalate, maxRounds, maxFixAttempts, coverageInterval int, systematic, replay, mutate bool, seed, oracleName string, verbose bool) error {
	// Create evaluator.
	eval, err := fuzz.NewEvaluator(oracleName)
	if err != nil {
		return err
	}
	if verbose {
		fmt.Printf("Using oracle: %s\n", eval.Name())
	}

	// Ensure output directories exist.
	testcaseDir := "testcases"
	failureDir := "failures"
	for _, dir := range []string{testcaseDir, failureDir} {
		if err := os.MkdirAll(dir, 0755); err != nil {
			return fmt.Errorf("create directory %s: %w", dir, err)
		}
	}

	// Set up signal handling for graceful shutdown.
	sigCh := make(chan os.Signal, 1)
	signal.Notify(sigCh, os.Interrupt)

	state := &OrchestratorState{
		Level:    startLevel,
		Coverage: fuzz.NewFunctionCoverage(),
	}

	fmt.Printf("Fuzz orchestrator starting at level %d (escalate after %d consecutive passes)\n", state.Level, passesToEscalate)
	fmt.Printf("Oracle: %s\n", eval.Name())
	if maxRounds > 0 {
		fmt.Printf("Max rounds: %d\n", maxRounds)
	}
	fmt.Println("Press Ctrl+C to stop.")
	fmt.Println()

	// Replay mode: re-run saved failures before the main loop.
	if replay {
		runReplay(state, maxFixAttempts, eval, testcaseDir, failureDir, verbose)
	}

	// Systematic mode.
	if systematic {
		return runSystematic(state, eval, maxRounds, maxFixAttempts, coverageInterval, testcaseDir, failureDir, verbose, sigCh)
	}

	// Track last passing spec for mutation mode.
	var lastPassingSpec *fuzz.TestSpec
	roundNum := 0

	for maxRounds <= 0 || state.TotalGenerated < maxRounds {
		// Check for interrupt.
		select {
		case <-sigCh:
			fmt.Println("\nInterrupted. Printing final summary.")
			printSummary(state)
			fmt.Println()
			fmt.Print(state.Coverage.Summary())
			return nil
		default:
		}

		state.TotalGenerated++
		roundNum++
		fmt.Printf("=== Round %d (Level %d, %d consecutive passes) ===\n",
			state.TotalGenerated, state.Level, state.ConsecutivePasses)

		roundStart := time.Now()

		// Decide whether to mutate or generate fresh.
		var spec *fuzz.TestSpec
		if mutate && lastPassingSpec != nil && roundNum%2 == 0 {
			fmt.Printf("  Mutating previous passing spec... ")
			spec = fuzz.MutateSpec(lastPassingSpec)
			fmt.Printf("ok (%s, %d sheets, %d checks)\n", spec.Name, len(spec.Sheets), len(spec.Checks))
		} else {
			complexity := fuzz.GetLevel(state.Level)
			seedCat := fuzz.SeedCategory(seed)
			leastTested := state.Coverage.LeastTestedFunctions(5)

			fmt.Printf("  Generating spec with Claude... ")
			genStart := time.Now()
			spec, err = fuzz.GenerateSpec(seedCat, complexity, eval, state.Coverage.BrokenList(), leastTested, verbose)
			genElapsed := time.Since(genStart)
			if err != nil {
				fmt.Printf("failed (%.1fs)\n", genElapsed.Seconds())
				fmt.Printf("  Generation failed: %v\n", err)
				roundElapsed := time.Since(roundStart)
				state.TotalTime += roundElapsed
				state.TotalFailed++
				fmt.Printf("  Round time: %.1fs\n", roundElapsed.Seconds())
				printSummary(state)
				fmt.Println()
				continue
			}
			fmt.Printf("ok (%.1fs)\n", genElapsed.Seconds())
			fmt.Printf("  Spec: %s (%d sheets, %d checks)\n",
				spec.Name, len(spec.Sheets), len(spec.Checks))
		}

		result, fi := executeSpec(state, maxFixAttempts, spec, eval, testcaseDir, failureDir, verbose)
		roundElapsed := time.Since(roundStart)
		state.TotalTime += roundElapsed
		fmt.Printf("  Round time: %.1fs\n", roundElapsed.Seconds())

		switch result {
		case resultPass:
			state.TotalPassed++
			state.ConsecutivePasses++
			lastPassingSpec = spec
			fmt.Printf("  PASS (%d/%d to escalate)\n", state.ConsecutivePasses, passesToEscalate)

			if state.ConsecutivePasses >= passesToEscalate {
				state.Level++
				state.ConsecutivePasses = 0
				fmt.Printf("  >>> Escalating to level %d <<<\n", state.Level)
			}

		case resultFixed:
			state.ConsecutivePasses = 0
			fmt.Println("  FIXED - running regression tests...")

			if err := runTests(verbose); err != nil {
				fmt.Printf("  Regression tests failed after fix, reverting: %v\n", err)
				revertFormula(verbose)
				state.TotalFailed++
			} else {
				fmt.Println("  Regression tests passed. Fix kept.")
				state.TotalFixed++
				// Auto-sync implemented functions after successful fix.
				if synced := fuzz.SyncImplementedFunctions(); synced > 0 {
					fmt.Printf("  Auto-synced implemented functions (%d total)\n", synced)
				}
				if fi != nil {
					generateTests(spec, fi.Mismatches, eval, fi.BrokenFuncs, verbose)
				}
			}

		case resultFailed:
			state.TotalFailed++
			state.ConsecutivePasses = 0
			fmt.Println("  FAILED (unfixed)")

		case resultError:
			state.TotalFailed++
		}

		printSummary(state)

		// Print coverage summary at interval.
		if coverageInterval > 0 && state.TotalGenerated%coverageInterval == 0 {
			fmt.Println()
			fmt.Print(state.Coverage.Summary())
		}
		fmt.Println()
	}

	fmt.Println("Max rounds reached. Final summary:")
	printSummary(state)
	fmt.Println()
	fmt.Print(state.Coverage.Summary())
	return nil
}

type roundResult int

const (
	resultPass   roundResult = iota
	resultFixed
	resultFailed
	resultError
)

// fixInfo holds data about a successful fix for downstream use (e.g., test generation).
type fixInfo struct {
	Mismatches  []fuzz.Mismatch
	BrokenFuncs []string
}

// executeSpec runs the build->eval->compare->fix pipeline on an already-generated spec.
// When the result is resultFixed, fixData contains the mismatches and broken functions.
func executeSpec(state *OrchestratorState, maxFixAttempts int, spec *fuzz.TestSpec, eval fuzz.Evaluator, testcaseDir, failureDir string, verbose bool) (roundResult, *fixInfo) {
	// Build XLSX.
	fmt.Printf("  Building XLSX... ")
	buildStart := time.Now()
	tmpDir, err := os.MkdirTemp("", "werkbook-fuzzorch-*")
	if err != nil {
		fmt.Printf("  Error creating temp dir: %v\n", err)
		return resultError, nil
	}
	defer os.RemoveAll(tmpDir)

	xlsxPath, wbResults, err := fuzz.BuildXLSX(spec, tmpDir)
	buildElapsed := time.Since(buildStart)
	if err != nil {
		fmt.Printf("failed (%.1fs)\n", buildElapsed.Seconds())
		fmt.Printf("  Build failed: %v\n", err)
		return resultError, nil
	}
	fmt.Printf("ok (%.1fs)\n", buildElapsed.Seconds())

	if verbose {
		fmt.Printf("  Werkbook results:\n")
		for _, r := range wbResults {
			fmt.Printf("    %s = %q (%s)\n", r.Ref, r.Value, r.Type)
		}
	}

	// Get oracle ground truth.
	fmt.Printf("  Evaluating with %s... ", eval.Name())
	evalStart := time.Now()
	groundTruth, err := eval.Eval(xlsxPath, spec.Checks)
	evalElapsed := time.Since(evalStart)
	if err != nil {
		fmt.Printf("failed (%.1fs)\n", evalElapsed.Seconds())
		fmt.Printf("  %s eval failed: %v\n", eval.Name(), err)
		return resultError, nil
	}
	fmt.Printf("ok (%.1fs)\n", evalElapsed.Seconds())

	if verbose {
		fmt.Printf("  %s results:\n", eval.Name())
		for _, r := range groundTruth {
			fmt.Printf("    %s = %q (%s)\n", r.Ref, r.Value, r.Type)
		}
	}

	// Save the test case.
	tcDir, err := fuzz.NewTestCaseDir(testcaseDir)
	if err != nil {
		fmt.Printf("  Error creating test case dir: %v\n", err)
		return resultError, nil
	}
	if err := fuzz.SaveTestCase(tcDir, spec, xlsxPath, groundTruth, eval.Name()); err != nil {
		fmt.Printf("  Error saving test case: %v\n", err)
		return resultError, nil
	}

	// Compare results.
	fmt.Printf("  Comparing results... ")
	compareStart := time.Now()
	mismatches := fuzz.CompareResults(spec.Checks, wbResults, groundTruth, spec)
	compareElapsed := time.Since(compareStart)

	// Track function coverage.
	specFunctions := fuzz.ExtractFunctions(spec)
	if len(specFunctions) > 0 {
		fmt.Printf("done (%.1fs)\n", compareElapsed.Seconds())
		fmt.Printf("  Functions tested: %s\n", strings.Join(specFunctions, ", "))
	} else {
		fmt.Printf("done (%.1fs)\n", compareElapsed.Seconds())
	}

	if len(mismatches) == 0 {
		state.Coverage.Record(specFunctions, true)
		fmt.Printf("  All %d checks match.\n", len(spec.Checks))
		result := &fuzz.CheckResult{Passed: true}
		fuzz.SaveCheckResult(tcDir, result)
		return resultPass, nil
	}

	state.Coverage.Record(specFunctions, false)
	fmt.Printf("  %d/%d checks mismatched:\n", len(mismatches), len(spec.Checks))
	fmt.Print(fuzz.FormatMismatches(mismatches, eval.Name()))

	// Save failure for manual review.
	saveFailure(spec, xlsxPath, failureDir)

	// Minimize spec for better fix prompts.
	minimized := fuzz.MinimizeSpec(spec, mismatches)

	// Classify the failure type.
	failureType := fuzz.ClassifyFailure(mismatches)
	fmt.Printf("  Failure type: %s\n", failureType)

	// Isolate broken functions for multi-function formulas.
	brokenFuncs := fuzz.IsolateFailure(spec, mismatches)
	if len(brokenFuncs) > 0 {
		fmt.Printf("  Isolated broken functions: %s\n", strings.Join(brokenFuncs, ", "))
	}

	// Load fix history for context.
	fixHistory := fuzz.LoadFixHistory("fix_history")

	// Attempt fix with Claude (with retry loop).
	var lastAttemptOutput, lastVerifyOutput string
	for attempt := 1; attempt <= maxFixAttempts; attempt++ {
		if attempt == 1 {
			fmt.Println("  Applying fix with Claude...")
		} else {
			fmt.Printf("  Retry fix attempt %d/%d...\n", attempt, maxFixAttempts)
			state.TotalRetries++
		}

		fixStart := time.Now()
		var output string
		if attempt == 1 {
			output, err = fuzz.ApplyFix(minimized, mismatches, eval.Name(), failureType, fixHistory, verbose, os.Stdout)
		} else {
			output, err = fuzz.ApplyFixRetry(minimized, mismatches, eval.Name(), lastAttemptOutput, lastVerifyOutput, verbose, os.Stdout)
		}
		fixElapsed := time.Since(fixStart)

		if err != nil {
			fmt.Printf("  Fix failed (%.1fs): %v\n", fixElapsed.Seconds(), err)
			revertFormula(verbose)
			break
		}
		fmt.Printf("  Fix applied (%.1fs)\n", fixElapsed.Seconds())

		if verbose {
			fmt.Printf("  Claude output:\n%s\n", output)
		}

		// Verify the fix by building and running a subprocess.
		fmt.Printf("  Verifying fix... ")
		verifyStart := time.Now()

		verifyDir, err := os.MkdirTemp("", "werkbook-fuzzorch-verify-*")
		if err != nil {
			fmt.Printf("  Error creating temp dir: %v\n", err)
			revertFormula(verbose)
			break
		}

		verifySpecPath := filepath.Join(verifyDir, "spec.json")
		if err := fuzz.SaveSpec(verifySpecPath, spec); err != nil {
			fmt.Printf("  Error saving spec for verification: %v\n", err)
			os.RemoveAll(verifyDir)
			revertFormula(verbose)
			break
		}

		// Build the verify binary.
		verifyBin := filepath.Join(verifyDir, "fuzz-verify")
		buildCmd := exec.Command("go", "build", "-o", verifyBin, "./cmd/fuzz/")
		buildOut, buildErr := buildCmd.CombinedOutput()
		if buildErr != nil {
			verifyElapsed := time.Since(verifyStart)
			fmt.Printf("build failed (%.1fs)\n", verifyElapsed.Seconds())
			fmt.Printf("  Build failed after fix:\n")
			for _, line := range strings.Split(strings.TrimSpace(string(buildOut)), "\n") {
				fmt.Printf("    %s\n", line)
			}
			fmt.Println("  Fix introduced compilation errors, reverting.")
			os.RemoveAll(verifyDir)
			revertFormula(verbose)

			lastAttemptOutput = output
			lastVerifyOutput = string(buildOut)
			if attempt < maxFixAttempts {
				continue
			}
			break
		}

		// Run the pre-built binary.
		verifyCmd := exec.Command(verifyBin, "-spec", verifySpecPath, "-oracle", eval.Name())
		verifyOut, verifyErr := verifyCmd.CombinedOutput()
		verifyElapsed := time.Since(verifyStart)
		os.RemoveAll(verifyDir)

		if verifyErr == nil {
			fmt.Printf("ok (%.1fs)\n", verifyElapsed.Seconds())
			if attempt > 1 {
				fmt.Printf("  Fix verified on attempt %d!\n", attempt)
				state.FixedOnRetry++
			} else {
				fmt.Println("  Fix verified: all checks now pass!")
			}

			// Mark functions as fixed.
			fixedFuncs := fuzz.ExtractBrokenFunctionsFromMismatches(spec, mismatches)
			state.Coverage.MarkFixed(fixedFuncs)

			// Save fix record for future reference.
			for _, fn := range brokenFuncs {
				record := fuzz.FixRecord{
					Function:    fn,
					FailureType: string(failureType),
					ErrorPattern: func() string {
						for _, m := range mismatches {
							return m.Reason
						}
						return ""
					}(),
					FixSummary: fuzz.Truncate(output, 500),
					Timestamp:  time.Now(),
				}
				fuzz.SaveFixRecord("fix_history", record)
			}

			result := &fuzz.CheckResult{Passed: false, FixApplied: true}
			fuzz.SaveCheckResult(tcDir, result)
			return resultFixed, &fixInfo{Mismatches: mismatches, BrokenFuncs: brokenFuncs}
		}

		fmt.Printf("failed (%.1fs)\n", verifyElapsed.Seconds())
		if outputStr := strings.TrimSpace(string(verifyOut)); outputStr != "" {
			for _, line := range strings.Split(outputStr, "\n") {
				fmt.Printf("    %s\n", line)
			}
		}

		lastAttemptOutput = output
		lastVerifyOutput = string(verifyOut)

		if attempt < maxFixAttempts {
			fmt.Println("  Fix did not resolve mismatches, retrying...")
			revertFormula(verbose)
		} else {
			fmt.Println("  Fix did not resolve mismatches, reverting.")
			revertFormula(verbose)
		}
	}

	// Mark functions as broken after all attempts failed.
	if len(brokenFuncs) > 0 {
		state.Coverage.MarkBroken(brokenFuncs)
	} else {
		allBroken := fuzz.ExtractBrokenFunctionsFromMismatches(spec, mismatches)
		state.Coverage.MarkBroken(allBroken)
	}

	result := &fuzz.CheckResult{Passed: false, Mismatches: mismatches}
	fuzz.SaveCheckResult(tcDir, result)
	return resultFailed, nil
}

// runReplay re-runs saved failures from the failures/ directory.
func runReplay(state *OrchestratorState, maxFixAttempts int, eval fuzz.Evaluator, testcaseDir, failureDir string, verbose bool) {
	matches, err := filepath.Glob(filepath.Join(failureDir, "*.json"))
	if err != nil || len(matches) == 0 {
		fmt.Println("No saved failures to replay.")
		return
	}

	resolvedDir := filepath.Join(failureDir, "resolved")
	os.MkdirAll(resolvedDir, 0755)

	fmt.Printf("=== Replaying %d saved failures ===\n", len(matches))
	passed, fixed, failed := 0, 0, 0

	for _, specPath := range matches {
		spec, err := fuzz.LoadSpec(specPath)
		if err != nil {
			if verbose {
				fmt.Printf("  Skipping %s: %v\n", filepath.Base(specPath), err)
			}
			continue
		}

		fmt.Printf("  Replaying %s... ", spec.Name)
		result, fi := executeSpec(state, maxFixAttempts, spec, eval, testcaseDir, failureDir, verbose)

		switch result {
		case resultPass:
			passed++
			fmt.Printf("  Now passes! Moving to resolved/\n")
			baseName := filepath.Base(specPath)
			os.Rename(specPath, filepath.Join(resolvedDir, baseName))
			// Also move matching XLSX if present.
			xlsxPath := strings.TrimSuffix(specPath, ".json") + ".xlsx"
			if _, err := os.Stat(xlsxPath); err == nil {
				os.Rename(xlsxPath, filepath.Join(resolvedDir, strings.TrimSuffix(baseName, ".json")+".xlsx"))
			}

		case resultFixed:
			fixed++
			fmt.Println("  FIXED - running regression tests...")
			if err := runTests(verbose); err != nil {
				fmt.Printf("  Regression tests failed after fix, reverting: %v\n", err)
				revertFormula(verbose)
				failed++
			} else {
				fmt.Println("  Regression tests passed. Fix kept.")
				state.TotalFixed++
				if synced := fuzz.SyncImplementedFunctions(); synced > 0 {
					fmt.Printf("  Auto-synced implemented functions (%d total)\n", synced)
				}
				if fi != nil {
					generateTests(spec, fi.Mismatches, eval, fi.BrokenFuncs, verbose)
				}
				// Move to resolved.
				baseName := filepath.Base(specPath)
				os.Rename(specPath, filepath.Join(resolvedDir, baseName))
				xlsxPath := strings.TrimSuffix(specPath, ".json") + ".xlsx"
				if _, err := os.Stat(xlsxPath); err == nil {
					os.Rename(xlsxPath, filepath.Join(resolvedDir, strings.TrimSuffix(baseName, ".json")+".xlsx"))
				}
			}

		case resultFailed, resultError:
			failed++
		}
	}

	fmt.Printf("=== Replay summary: %d passed, %d fixed, %d still failing ===\n\n", passed, fixed, failed)
}

// runSystematic runs the systematic function targeting mode.
func runSystematic(state *OrchestratorState, eval fuzz.Evaluator, maxRounds, maxFixAttempts, coverageInterval int, testcaseDir, failureDir string, verbose bool, sigCh chan os.Signal) error {
	sysState := &SystematicState{
		Queue:   buildUnimplementedQueue(eval),
		Current: 0,
		Phase:   "targeted",
	}

	totalUnimplemented := len(sysState.Queue)
	fmt.Printf("Systematic mode: %d unimplemented functions to test\n", totalUnimplemented)

	for maxRounds <= 0 || state.TotalGenerated < maxRounds {
		select {
		case <-sigCh:
			fmt.Println("\nInterrupted. Printing final summary.")
			printSummary(state)
			fmt.Println()
			fmt.Print(state.Coverage.Summary())
			return nil
		default:
		}

		state.TotalGenerated++
		roundStart := time.Now()

		if sysState.Current < len(sysState.Queue) {
			targetFn := sysState.Queue[sysState.Current]
			remaining := len(sysState.Queue) - sysState.Current
			fmt.Printf("=== Round %d (targeting %s, %d remaining) ===\n", state.TotalGenerated, targetFn, remaining)

			complexity := fuzz.GetLevel(state.Level)

			fmt.Printf("  Generating spec for %s... ", targetFn)
			genStart := time.Now()
			spec, err := fuzz.GenerateTargetedSpec(targetFn, complexity, eval, verbose)
			genElapsed := time.Since(genStart)
			if err != nil {
				fmt.Printf("failed (%.1fs)\n", genElapsed.Seconds())
				fmt.Printf("  Generation failed: %v\n", err)
				state.TotalFailed++
				sysState.Current++
				roundElapsed := time.Since(roundStart)
				state.TotalTime += roundElapsed
				fmt.Printf("  Round time: %.1fs\n", roundElapsed.Seconds())
				printSummary(state)
				fmt.Println()
				continue
			}
			fmt.Printf("ok (%.1fs)\n", genElapsed.Seconds())

			fmt.Printf("  Spec: %s (%d sheets, %d checks)\n", spec.Name, len(spec.Sheets), len(spec.Checks))

			result, fi := executeSpec(state, maxFixAttempts, spec, eval, testcaseDir, failureDir, verbose)

			switch result {
			case resultPass:
				state.TotalPassed++
				fmt.Printf("  PASS - %s works correctly\n", targetFn)
			case resultFixed:
				fmt.Println("  FIXED - running regression tests...")
				if err := runTests(verbose); err != nil {
					fmt.Printf("  Regression tests failed after fix, reverting: %v\n", err)
					revertFormula(verbose)
					state.TotalFailed++
				} else {
					fmt.Println("  Regression tests passed. Fix kept.")
					state.TotalFixed++
					if synced := fuzz.SyncImplementedFunctions(); synced > 0 {
						fmt.Printf("  Auto-synced implemented functions (%d total)\n", synced)
					}
					if fi != nil {
						generateTests(spec, fi.Mismatches, eval, fi.BrokenFuncs, verbose)
					}
				}
			case resultFailed:
				state.TotalFailed++
				fmt.Printf("  FAILED - %s has mismatches\n", targetFn)
			case resultError:
				state.TotalFailed++
			}

			sysState.Current++
		} else {
			if sysState.Phase != "regression" {
				sysState.Phase = "regression"
				fmt.Printf("\n--- All unimplemented functions tested, switching to regression mode ---\n")
			}
			fmt.Printf("=== Round %d (regression mode, random) ===\n", state.TotalGenerated)

			complexity := fuzz.GetLevel(state.Level)
			seedCat := fuzz.SeedCategory("")
			leastTested := state.Coverage.LeastTestedFunctions(5)

			fmt.Printf("  Generating spec with Claude... ")
			genStart := time.Now()
			spec, err := fuzz.GenerateSpec(seedCat, complexity, eval, state.Coverage.BrokenList(), leastTested, verbose)
			genElapsed := time.Since(genStart)
			if err != nil {
				fmt.Printf("failed (%.1fs)\n", genElapsed.Seconds())
				fmt.Printf("  Generation failed: %v\n", err)
				state.TotalFailed++
				roundElapsed := time.Since(roundStart)
				state.TotalTime += roundElapsed
				fmt.Printf("  Round time: %.1fs\n", roundElapsed.Seconds())
				printSummary(state)
				fmt.Println()
				continue
			}
			fmt.Printf("ok (%.1fs)\n", genElapsed.Seconds())
			fmt.Printf("  Spec: %s (%d sheets, %d checks)\n", spec.Name, len(spec.Sheets), len(spec.Checks))

			result, fi := executeSpec(state, maxFixAttempts, spec, eval, testcaseDir, failureDir, verbose)
			switch result {
			case resultPass:
				state.TotalPassed++
			case resultFixed:
				fmt.Println("  FIXED - running regression tests...")
				if err := runTests(verbose); err != nil {
					fmt.Printf("  Regression tests failed after fix, reverting: %v\n", err)
					revertFormula(verbose)
					state.TotalFailed++
				} else {
					fmt.Println("  Regression tests passed. Fix kept.")
					state.TotalFixed++
					if synced := fuzz.SyncImplementedFunctions(); synced > 0 {
						fmt.Printf("  Auto-synced implemented functions (%d total)\n", synced)
					}
					if fi != nil {
						generateTests(spec, fi.Mismatches, eval, fi.BrokenFuncs, verbose)
					}
				}
			case resultFailed:
				state.TotalFailed++
			case resultError:
				state.TotalFailed++
			}
		}

		roundElapsed := time.Since(roundStart)
		state.TotalTime += roundElapsed
		fmt.Printf("  Round time: %.1fs\n", roundElapsed.Seconds())

		printSummary(state)
		if coverageInterval > 0 && state.TotalGenerated%coverageInterval == 0 {
			fmt.Println()
			fmt.Print(state.Coverage.Summary())
		}
		fmt.Println()
	}

	fmt.Println("Finished. Final summary:")
	printSummary(state)
	fmt.Printf("Functions tested: %d/%d unimplemented\n", sysState.Current, totalUnimplemented)
	fmt.Println()
	fmt.Print(state.Coverage.Summary())
	return nil
}

// buildUnimplementedQueue returns unimplemented functions sorted by category.
func buildUnimplementedQueue(eval fuzz.Evaluator) []string {
	categoryOrder := map[string]int{
		"math": 0, "logical": 1, "text": 2, "statistical": 3,
		"lookup": 4, "date": 5, "information": 6, "financial": 7,
		"engineering": 8, "database": 9,
	}

	// Build category map.
	categories := map[string]string{}
	catFuncs := map[string][]string{
		"math":        {"ABS", "ACOS", "ACOSH", "ACOT", "ACOTH", "ARABIC", "ASIN", "ASINH", "ATAN", "ATAN2", "ATANH", "BASE", "CEILING", "CEILING.MATH", "CEILING.PRECISE", "COMBIN", "COMBINA", "COS", "COSH", "COT", "COTH", "CSC", "CSCH", "DECIMAL", "DEGREES", "EVEN", "EXP", "FACT", "FACTDOUBLE", "FLOOR", "FLOOR.MATH", "FLOOR.PRECISE", "GCD", "INT", "LCM", "LN", "LOG", "LOG10", "MDETERM", "MINVERSE", "MMULT", "MOD", "MROUND", "MULTINOMIAL", "MUNIT", "ODD", "PI", "POWER", "PRODUCT", "QUOTIENT", "RADIANS", "ROMAN", "ROUND", "ROUNDDOWN", "ROUNDUP", "SEC", "SECH", "SERIESSUM", "SIGN", "SIN", "SINH", "SQRT", "SQRTPI", "SUBTOTAL", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ", "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "TAN", "TANH", "TRUNC"},
		"logical":     {"AND", "FALSE", "IF", "IFERROR", "IFNA", "IFS", "NOT", "OR", "SWITCH", "TRUE", "XOR"},
		"text":        {"CHAR", "CLEAN", "CODE", "CONCATENATE", "DOLLAR", "EXACT", "FIND", "FIXED", "LEFT", "LEN", "LOWER", "MID", "NUMBERVALUE", "PROPER", "REPLACE", "REPT", "RIGHT", "SEARCH", "SUBSTITUTE", "T", "TEXT", "TEXTJOIN", "TRIM", "UNICHAR", "UNICODE", "UPPER", "VALUE"},
		"statistical": {"AVEDEV", "AVERAGE", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS", "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.DIST.RANGE", "BINOM.INV", "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST", "CONFIDENCE.NORM", "CONFIDENCE.T", "CORREL", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "COVARIANCE.P", "COVARIANCE.S", "DEVSQ", "EXPON.DIST", "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST", "FISHER", "FISHERINV", "FORECAST.LINEAR", "FREQUENCY", "GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMALN", "GAMMALN.PRECISE", "GAUSS", "GEOMEAN", "GROWTH", "HARMEAN", "HYPGEOM.DIST", "INTERCEPT", "KURT", "LARGE", "LINEST", "LOGEST", "LOGNORM.DIST", "LOGNORM.INV", "MAX", "MAXA", "MAXIFS", "MEDIAN", "MIN", "MINA", "MINIFS", "MODE.MULT", "MODE.SNGL", "NEGBINOM.DIST", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "PEARSON", "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTRANK.EXC", "PERCENTRANK.INC", "PERMUT", "PERMUTATIONA", "PHI", "POISSON.DIST", "PROB", "QUARTILE.EXC", "QUARTILE.INC", "RANK.AVG", "RANK.EQ", "RSQ", "SKEW", "SKEW.P", "SLOPE", "SMALL", "STANDARDIZE", "STDEV.P", "STDEV.S", "STDEVA", "STDEVPA", "STEYX", "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST", "TREND", "TRIMMEAN", "VAR.P", "VAR.S", "VARA", "VARPA", "WEIBULL.DIST", "Z.TEST"},
		"lookup":      {"ADDRESS", "AREAS", "CHOOSE", "COLUMN", "COLUMNS", "FORMULATEXT", "HLOOKUP", "HYPERLINK", "INDEX", "LOOKUP", "MATCH", "ROW", "ROWS", "TRANSPOSE", "VLOOKUP"},
		"date":        {"DATE", "DATEDIF", "DATEVALUE", "DAY", "DAYS", "DAYS360", "EDATE", "EOMONTH", "HOUR", "ISOWEEKNUM", "MINUTE", "MONTH", "NETWORKDAYS", "NETWORKDAYS.INTL", "SECOND", "TIME", "TIMEVALUE", "WEEKDAY", "WEEKNUM", "WORKDAY", "WORKDAY.INTL", "YEAR", "YEARFRAC"},
		"information": {"ERROR.TYPE", "ISBLANK", "ISERR", "ISERROR", "ISEVEN", "ISFORMULA", "ISLOGICAL", "ISNA", "ISNONTEXT", "ISNUMBER", "ISODD", "ISREF", "ISTEXT", "N", "NA", "SHEET", "SHEETS", "TYPE"},
		"financial":   {"ACCRINT", "ACCRINTM", "AMORDEGRC", "AMORLINC", "COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD", "CUMIPMT", "CUMPRINC", "DB", "DDB", "DISC", "DOLLARDE", "DOLLARFR", "DURATION", "EFFECT", "FV", "FVSCHEDULE", "INTRATE", "IPMT", "IRR", "ISPMT", "MDURATION", "MIRR", "NOMINAL", "NPER", "NPV", "ODDFPRICE", "ODDFYIELD", "ODDLPRICE", "ODDLYIELD", "PDURATION", "PMT", "PPMT", "PRICE", "PRICEDISC", "PRICEMAT", "PV", "RATE", "RECEIVED", "RRI", "SLN", "SYD", "TBILLEQ", "TBILLPRICE", "TBILLYIELD", "VDB", "XIRR", "XNPV", "YIELD", "YIELDDISC", "YIELDMAT"},
		"engineering": {"BESSELI", "BESSELJ", "BESSELK", "BESSELY", "BIN2DEC", "BIN2HEX", "BIN2OCT", "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "COMPLEX", "CONVERT", "DEC2BIN", "DEC2HEX", "DEC2OCT", "DELTA", "ERF", "ERF.PRECISE", "ERFC", "ERFC.PRECISE", "GESTEP", "HEX2BIN", "HEX2DEC", "HEX2OCT", "IMABS", "IMAGINARY", "IMARGUMENT", "IMCONJUGATE", "IMCOS", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMDIV", "IMEXP", "IMLN", "IMLOG10", "IMLOG2", "IMPOWER", "IMPRODUCT", "IMREAL", "IMSEC", "IMSECH", "IMSIN", "IMSINH", "IMSQRT", "IMSUB", "IMSUM", "IMTAN", "OCT2BIN", "OCT2DEC", "OCT2HEX"},
		"database":    {"DAVERAGE", "DCOUNT", "DCOUNTA", "DGET", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM", "DVAR", "DVARP"},
	}
	for cat, fns := range catFuncs {
		for _, fn := range fns {
			categories[fn] = cat
		}
	}

	var excluded map[string]bool
	if eval != nil {
		excluded = eval.ExcludedFunctions()
	}

	var queue []string
	for _, fn := range fuzz.KnownFunctionsList {
		if fuzz.ImplementedFunctions[fn] {
			continue
		}
		if excluded != nil && excluded[fn] {
			continue
		}
		queue = append(queue, fn)
	}

	sort.Slice(queue, func(i, j int) bool {
		orderI := categoryOrder[categories[queue[i]]]
		orderJ := categoryOrder[categories[queue[j]]]
		if orderI != orderJ {
			return orderI < orderJ
		}
		return queue[i] < queue[j]
	})

	return queue
}

// saveFailure persists a failing spec and its XLSX for later reproduction.
func saveFailure(spec *fuzz.TestSpec, xlsxPath, failDir string) {
	name := fmt.Sprintf("%s_%s", spec.Name, time.Now().Format("20060102_150405"))
	specPath := filepath.Join(failDir, name+".json")
	if err := fuzz.SaveSpec(specPath, spec); err != nil {
		fmt.Printf("  Warning: could not save spec: %v\n", err)
		return
	}

	xlsxDst := filepath.Join(failDir, name+".xlsx")
	data, err := os.ReadFile(xlsxPath)
	if err == nil {
		os.WriteFile(xlsxDst, data, 0644)
	}

	fmt.Printf("  Saved failure: %s\n", specPath)
}

// generateTests calls Claude to generate unit tests for the fixed function(s).
// This is non-blocking: failure does not revert the implementation fix.
func generateTests(spec *fuzz.TestSpec, mismatches []fuzz.Mismatch, eval fuzz.Evaluator, brokenFuncs []string, verbose bool) {
	if len(brokenFuncs) == 0 {
		return
	}

	fmt.Printf("  Generating tests for %s...\n", strings.Join(brokenFuncs, ", "))
	output, err := fuzz.GenerateTests(spec, mismatches, eval.Name(), brokenFuncs, verbose, os.Stdout)
	if err != nil {
		fmt.Printf("  Test generation failed: %v\n", err)
		return
	}

	if verbose {
		fmt.Printf("  Claude output:\n%s\n", output)
	}
	fmt.Println("  Tests generated.")

	// Verify the new tests pass.
	if err := runTests(verbose); err != nil {
		fmt.Printf("  Generated tests failed, reverting test changes: %v\n", err)
		revertTestFiles(verbose)
		return
	}

	fmt.Println("  Generated tests pass.")
}

// revertTestFiles reverts only test file changes in the formula/ package.
func revertTestFiles(verbose bool) {
	if verbose {
		fmt.Println("  Reverting formula/*_test.go changes...")
	}
	cmd := exec.Command("git", "checkout", "--", "formula/*_test.go")
	cmd.Stderr = os.Stderr
	if err := cmd.Run(); err != nil {
		// Fallback: try with glob expansion via shell.
		cmd2 := exec.Command("bash", "-c", "git checkout -- formula/*_test.go")
		cmd2.Stderr = os.Stderr
		cmd2.Run()
	}
}

// revertFormula reverts any changes Claude made to the formula/ package.
func revertFormula(verbose bool) {
	if verbose {
		fmt.Println("  Reverting formula/ changes...")
	}
	cmd := exec.Command("git", "checkout", "--", "formula/")
	cmd.Stderr = os.Stderr
	if err := cmd.Run(); err != nil {
		fmt.Printf("  Warning: git checkout formula/ failed: %v\n", err)
	}
}

// runTests runs `gotestsum ./...` to check for regressions.
func runTests(verbose bool) error {
	cmd := exec.Command("gotestsum", "./...")
	if verbose {
		cmd.Stdout = os.Stdout
		cmd.Stderr = os.Stderr
	}
	return cmd.Run()
}

func printSummary(state *OrchestratorState) {
	brokenList := state.Coverage.BrokenList()
	avg := time.Duration(0)
	if state.TotalGenerated > 0 {
		avg = state.TotalTime / time.Duration(state.TotalGenerated)
	}

	parts := []string{
		fmt.Sprintf("Generated: %d", state.TotalGenerated),
		fmt.Sprintf("Passed: %d", state.TotalPassed),
		fmt.Sprintf("Fixed: %d", state.TotalFixed),
		fmt.Sprintf("Failed: %d", state.TotalFailed),
	}
	if state.TotalRetries > 0 {
		parts = append(parts, fmt.Sprintf("Retries: %d", state.TotalRetries))
	}
	if state.FixedOnRetry > 0 {
		parts = append(parts, fmt.Sprintf("Fixed on retry: %d", state.FixedOnRetry))
	}
	if len(brokenList) > 0 {
		parts = append(parts, fmt.Sprintf("Broken: %s", strings.Join(brokenList, ", ")))
	}
	parts = append(parts, fmt.Sprintf("Avg: %.1fs/round", avg.Seconds()))

	fmt.Printf("  [Level %d] %s\n", state.Level, strings.Join(parts, " | "))
}
