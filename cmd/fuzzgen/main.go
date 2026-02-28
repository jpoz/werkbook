package main

import (
	"flag"
	"fmt"
	"os"

	"github.com/jpoz/werkbook/internal/fuzz"
)

func main() {
	level := flag.Int("level", 1, "complexity level (1-8+)")
	seed := flag.String("seed", "", "focus category: math|text|lookup|logic|stat|date|info|financial|engineering|database|mixed")
	outDir := flag.String("outdir", "testcases", "output directory for test cases")
	oracle := flag.String("oracle", "libreoffice", "validation oracle: libreoffice or excel")
	verbose := flag.Bool("v", false, "verbose output")
	flag.Parse()

	if err := run(*level, *seed, *outDir, *oracle, *verbose); err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
}

func run(level int, seed, outDir, oracleName string, verbose bool) error {
	// Create evaluator.
	fmt.Fprintf(os.Stderr, "Initializing %s oracle... ", oracleName)
	eval, err := fuzz.NewEvaluator(oracleName)
	if err != nil {
		fmt.Fprintf(os.Stderr, "failed\n")
		return err
	}
	fmt.Fprintf(os.Stderr, "ok\n")

	// Get complexity config for this level.
	complexity := fuzz.GetLevel(level)
	fmt.Fprintf(os.Stderr, "Level %d: %d-%d cells, %d-%d checks, %d sheets max\n",
		complexity.Level, complexity.MinCells, complexity.MaxCells,
		complexity.MinChecks, complexity.MaxChecks, complexity.MaxSheets)

	// Generate spec using Claude.
	seedCat := fuzz.SeedCategory(seed)
	fmt.Fprintf(os.Stderr, "Generating spec with Claude (this may take a moment)... ")
	spec, err := fuzz.GenerateSpec(seedCat, complexity, eval, nil, nil, verbose)
	if err != nil {
		fmt.Fprintf(os.Stderr, "failed\n")
		return fmt.Errorf("generate spec: %w", err)
	}
	fmt.Fprintf(os.Stderr, "ok\n")
	fmt.Fprintf(os.Stderr, "  Spec: %s (%d sheets, %d checks)\n",
		spec.Name, len(spec.Sheets), len(spec.Checks))

	// Build XLSX and get werkbook results (we need the XLSX for the oracle).
	fmt.Fprintf(os.Stderr, "Building XLSX with werkbook... ")
	tmpDir, err := os.MkdirTemp("", "werkbook-fuzzgen-*")
	if err != nil {
		return fmt.Errorf("create temp dir: %w", err)
	}
	defer os.RemoveAll(tmpDir)

	xlsxPath, _, err := fuzz.BuildXLSX(spec, tmpDir)
	if err != nil {
		fmt.Fprintf(os.Stderr, "failed\n")
		return fmt.Errorf("build xlsx: %w", err)
	}
	fmt.Fprintf(os.Stderr, "ok\n")

	// Evaluate with the oracle to get ground truth.
	fmt.Fprintf(os.Stderr, "Evaluating with %s... ", eval.Name())
	groundTruth, err := eval.Eval(xlsxPath, spec.Checks)
	if err != nil {
		fmt.Fprintf(os.Stderr, "failed\n")
		return fmt.Errorf("%s eval: %w", eval.Name(), err)
	}
	fmt.Fprintf(os.Stderr, "ok (%d results)\n", len(groundTruth))

	if verbose {
		for _, r := range groundTruth {
			fmt.Fprintf(os.Stderr, "  %s = %q (%s)\n", r.Ref, r.Value, r.Type)
		}
	}

	// Create test case directory and save everything.
	fmt.Fprintf(os.Stderr, "Saving test case... ")
	tcDir, err := fuzz.NewTestCaseDir(outDir)
	if err != nil {
		return fmt.Errorf("create test case dir: %w", err)
	}

	if err := fuzz.SaveTestCase(tcDir, spec, xlsxPath, groundTruth, eval.Name()); err != nil {
		return fmt.Errorf("save test case: %w", err)
	}
	fmt.Fprintf(os.Stderr, "ok\n")

	// Print the test case directory path to stdout (for orchestrator to consume).
	fmt.Println(tcDir)
	return nil
}
