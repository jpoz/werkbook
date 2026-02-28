package fuzz

import (
	"encoding/json"
	"os"
	"regexp"
	"sync"
	"time"
)

// LearnedIsolationTest is a test case auto-captured from passing specs.
type LearnedIsolationTest struct {
	Formula    string    `json:"formula"`
	Expected   string    `json:"expected"`
	Type       string    `json:"type"`
	SourceSpec string    `json:"source_spec,omitempty"`
	Oracle     string    `json:"oracle,omitempty"`
	CapturedAt time.Time `json:"captured_at"`
}

// LearnedIsolationTestsPath is the file path for storing learned isolation tests.
var LearnedIsolationTestsPath = "isolate_learned.json"

var (
	learnedIsolationTests map[string]LearnedIsolationTest
	learnedMu             sync.RWMutex
	learnedLoaded         bool
)

var cellRefRegex = regexp.MustCompile(`[A-Z]+[0-9]+`)

// loadLearnedTests loads learned tests from disk once.
func loadLearnedTests() {
	learnedMu.Lock()
	defer learnedMu.Unlock()
	if learnedLoaded {
		return
	}
	learnedLoaded = true
	learnedIsolationTests = readLearnedTestsFromDisk()
}

func readLearnedTestsFromDisk() map[string]LearnedIsolationTest {
	b, err := os.ReadFile(LearnedIsolationTestsPath)
	if err != nil {
		return make(map[string]LearnedIsolationTest)
	}
	var m map[string]LearnedIsolationTest
	if err := json.Unmarshal(b, &m); err != nil {
		return make(map[string]LearnedIsolationTest)
	}
	return m
}

func writeLearnedTestsToDisk(m map[string]LearnedIsolationTest) error {
	b, err := json.MarshalIndent(m, "", "  ")
	if err != nil {
		return err
	}
	tmp := LearnedIsolationTestsPath + ".tmp"
	if err := os.WriteFile(tmp, b, 0644); err != nil {
		return err
	}
	return os.Rename(tmp, LearnedIsolationTestsPath)
}

// GetLearnedIsolationTest returns a learned test for the given function, if any.
func GetLearnedIsolationTest(fn string) (LearnedIsolationTest, bool) {
	loadLearnedTests()
	learnedMu.RLock()
	defer learnedMu.RUnlock()
	t, ok := learnedIsolationTests[fn]
	return t, ok
}

// LearnIsolationTests extracts good isolation test candidates from a passing
// spec and saves them. Returns the count of newly learned tests.
func LearnIsolationTests(spec *TestSpec, groundTruth []CellResult, oracleName string) int {
	if spec == nil || len(groundTruth) == 0 {
		return 0
	}

	// Build a map from check ref to oracle result.
	oracleMap := make(map[string]CellResult)
	for _, r := range groundTruth {
		oracleMap[r.Ref] = r
	}

	// Build a map from check ref to formula.
	formulaMap := buildFormulaMap(spec)

	loadLearnedTests()
	learnedMu.Lock()
	defer learnedMu.Unlock()

	newCount := 0

	for _, check := range spec.Checks {
		formula, ok := formulaMap[check.Ref]
		if !ok || formula == "" {
			continue
		}

		oracleResult, ok := oracleMap[check.Ref]
		if !ok || oracleResult.Value == "" {
			continue
		}

		// Skip error results.
		if oracleResult.Type == "error" {
			continue
		}

		// Extract functions from this formula.
		funcs := ExtractFunctionsFromFormula(formula)
		if len(funcs) == 0 || len(funcs) > 2 {
			continue
		}

		// Skip long formulas.
		if len(formula) > 80 {
			continue
		}

		// Skip formulas with cell references (they depend on data cells).
		// Strip function names first, then check for cell-ref patterns.
		stripped := funcNameRegex.ReplaceAllString(formula, "")
		if cellRefRegex.MatchString(stripped) {
			continue
		}

		// Skip non-deterministic functions.
		skip := false
		for _, fn := range funcs {
			if commonExcludedFunctions[fn] {
				skip = true
				break
			}
		}
		if skip {
			continue
		}

		// The primary function is the first (or only) one.
		primaryFn := funcs[0]

		// Don't overwrite hardcoded tests.
		if _, hasHardcoded := isolationTests[primaryFn]; hasHardcoded {
			continue
		}

		// Check if we already have a learned test — prefer simpler ones.
		if existing, exists := learnedIsolationTests[primaryFn]; exists {
			existingFuncs := ExtractFunctionsFromFormula(existing.Formula)
			if len(funcs) > len(existingFuncs) {
				continue
			}
			if len(funcs) == len(existingFuncs) && len(formula) >= len(existing.Formula) {
				continue
			}
		}

		learnedIsolationTests[primaryFn] = LearnedIsolationTest{
			Formula:    formula,
			Expected:   oracleResult.Value,
			Type:       oracleResult.Type,
			SourceSpec: spec.Name,
			Oracle:     oracleName,
			CapturedAt: time.Now(),
		}
		newCount++
	}

	if newCount > 0 {
		writeLearnedTestsToDisk(learnedIsolationTests)
	}

	return newCount
}
