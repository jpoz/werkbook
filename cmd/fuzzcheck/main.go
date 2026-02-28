package main

import (
	"flag"
	"fmt"
	"os"

	"github.com/jpoz/werkbook/internal/fuzz"
)

func main() {
	testcaseDir := flag.String("testcase", "", "path to test case directory")
	noFix := flag.Bool("no-fix", false, "just report mismatches, don't attempt fixes")
	verbose := flag.Bool("v", false, "verbose output")
	flag.Parse()

	if *testcaseDir == "" {
		fmt.Fprintf(os.Stderr, "Error: --testcase is required\n")
		flag.Usage()
		os.Exit(2)
	}

	exitCode := run(*testcaseDir, *noFix, *verbose)
	os.Exit(exitCode)
}

// run executes the check/fix cycle. Returns exit code:
//
//	0 = pass
//	1 = fail, fix applied
//	2 = fail, no fix or fix failed
func run(tcDir string, noFix, verbose bool) int {
	// Load the test case.
	fmt.Printf("Loading test case from %s... ", tcDir)
	tc, err := fuzz.LoadTestCase(tcDir)
	if err != nil {
		fmt.Printf("failed\n")
		fmt.Fprintf(os.Stderr, "Error loading test case: %v\n", err)
		return 2
	}
	fmt.Printf("ok\n")
	fmt.Printf("  Spec: %s (%d sheets, %d checks)\n",
		tc.Spec.Name, len(tc.Spec.Sheets), len(tc.Spec.Checks))

	// Build XLSX and evaluate with werkbook.
	fmt.Printf("Evaluating with werkbook... ")
	tmpDir, err := os.MkdirTemp("", "werkbook-fuzzcheck-*")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error creating temp dir: %v\n", err)
		return 2
	}
	defer os.RemoveAll(tmpDir)

	_, wbResults, err := fuzz.BuildXLSX(tc.Spec, tmpDir)
	if err != nil {
		fmt.Printf("failed\n")
		fmt.Fprintf(os.Stderr, "Error building XLSX: %v\n", err)
		return 2
	}
	fmt.Printf("ok\n")

	if verbose {
		fmt.Printf("Werkbook results:\n")
		for _, r := range wbResults {
			fmt.Printf("  %s = %q (%s)\n", r.Ref, r.Value, r.Type)
		}
	}

	// Compare against ground truth.
	fmt.Printf("Comparing against ground truth... ")
	mismatches := fuzz.CompareResults(tc.Spec.Checks, wbResults, tc.GroundTruth, tc.Spec)

	if len(mismatches) == 0 {
		fmt.Printf("ok\n")
		fmt.Printf("PASS: all %d checks match\n", len(tc.Spec.Checks))
		result := &fuzz.CheckResult{Passed: true}
		if err := fuzz.SaveCheckResult(tcDir, result); err != nil {
			fmt.Fprintf(os.Stderr, "Warning: could not save result: %v\n", err)
		}
		return 0
	}
	fmt.Printf("done\n")

	fmt.Printf("FAIL: %d/%d checks mismatched:\n", len(mismatches), len(tc.Spec.Checks))
	oracleName := tc.OracleName
	if oracleName == "" {
		oracleName = "libreoffice"
	}
	fmt.Print(fuzz.FormatMismatches(mismatches, oracleName))

	if noFix {
		result := &fuzz.CheckResult{Passed: false, Mismatches: mismatches}
		if err := fuzz.SaveCheckResult(tcDir, result); err != nil {
			fmt.Fprintf(os.Stderr, "Warning: could not save result: %v\n", err)
		}
		return 2
	}

	// Attempt fix with Claude.
	fmt.Println("Asking Claude for fix suggestions (this may take a moment)...")
	suggestion, err := fuzz.SuggestFix(tc.Spec, mismatches, oracleName, verbose)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Fix suggestion failed: %v\n", err)
		result := &fuzz.CheckResult{Passed: false, Mismatches: mismatches}
		if err := fuzz.SaveCheckResult(tcDir, result); err != nil {
			fmt.Fprintf(os.Stderr, "Warning: could not save result: %v\n", err)
		}
		return 2
	}

	fmt.Printf("\nFix suggestion:\n%s\n", suggestion)

	// Re-build and re-evaluate to verify the fix.
	fmt.Printf("Verifying fix... ")
	tmpDir2, err := os.MkdirTemp("", "werkbook-fuzzcheck-verify-*")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error creating temp dir: %v\n", err)
		return 2
	}
	defer os.RemoveAll(tmpDir2)

	_, wbResults2, err := fuzz.BuildXLSX(tc.Spec, tmpDir2)
	if err != nil {
		fmt.Printf("failed\n")
		fmt.Fprintf(os.Stderr, "Error re-building XLSX: %v\n", err)
		result := &fuzz.CheckResult{Passed: false, Mismatches: mismatches}
		if err := fuzz.SaveCheckResult(tcDir, result); err != nil {
			fmt.Fprintf(os.Stderr, "Warning: could not save result: %v\n", err)
		}
		return 2
	}

	mismatches2 := fuzz.CompareResults(tc.Spec.Checks, wbResults2, tc.GroundTruth, tc.Spec)
	if len(mismatches2) == 0 {
		fmt.Printf("ok\n")
		fmt.Println("Fix verified: all checks now pass!")
		result := &fuzz.CheckResult{Passed: false, FixApplied: true}
		if err := fuzz.SaveCheckResult(tcDir, result); err != nil {
			fmt.Fprintf(os.Stderr, "Warning: could not save result: %v\n", err)
		}
		return 1
	}

	fmt.Printf("done\n")
	fmt.Printf("Fix did not resolve all mismatches (%d remaining)\n", len(mismatches2))
	result := &fuzz.CheckResult{Passed: false, Mismatches: mismatches2}
	if err := fuzz.SaveCheckResult(tcDir, result); err != nil {
		fmt.Fprintf(os.Stderr, "Warning: could not save result: %v\n", err)
	}
	return 2
}
