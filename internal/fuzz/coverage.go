package fuzz

import (
	"fmt"
	"regexp"
	"sort"
	"strings"
	"sync"
)

const hardBrokenThreshold = 3

// FunctionCoverage tracks which functions have been tested and their pass/fail rates.
type FunctionCoverage struct {
	Tested map[string]*FunctionStats
	Broken map[string]int // function -> consecutive failed fix attempts
	mu     sync.Mutex
}

// FunctionStats records pass/fail counts for a single function.
type FunctionStats struct {
	Passed int
	Failed int
}

// NewFunctionCoverage creates a new coverage tracker.
func NewFunctionCoverage() *FunctionCoverage {
	return &FunctionCoverage{
		Tested: make(map[string]*FunctionStats),
		Broken: make(map[string]int),
	}
}

var funcNameRegex = regexp.MustCompile(`([A-Z][A-Z0-9.]+)\(`)

// ExtractFunctionsFromFormula extracts function names from a single formula string.
// Returns deduplicated, sorted list.
func ExtractFunctionsFromFormula(formula string) []string {
	upper := strings.ToUpper(formula)
	matches := funcNameRegex.FindAllStringSubmatch(upper, -1)
	seen := make(map[string]bool)
	var funcs []string
	for _, m := range matches {
		name := m[1]
		if !seen[name] {
			seen[name] = true
			funcs = append(funcs, name)
		}
	}
	sort.Strings(funcs)
	return funcs
}

// ExtractFunctions extracts all function names from all formulas in a test spec.
// Returns deduplicated, sorted list.
func ExtractFunctions(spec *TestSpec) []string {
	if spec == nil {
		return nil
	}
	seen := make(map[string]bool)
	var funcs []string
	for _, sheet := range spec.Sheets {
		for _, cell := range sheet.Cells {
			if cell.Formula == "" {
				continue
			}
			for _, fn := range ExtractFunctionsFromFormula(cell.Formula) {
				if !seen[fn] {
					seen[fn] = true
					funcs = append(funcs, fn)
				}
			}
		}
	}
	sort.Strings(funcs)
	return funcs
}

// Record increments pass/fail counters for each function in the list.
func (fc *FunctionCoverage) Record(functions []string, passed bool) {
	fc.mu.Lock()
	defer fc.mu.Unlock()
	for _, fn := range functions {
		stats, ok := fc.Tested[fn]
		if !ok {
			stats = &FunctionStats{}
			fc.Tested[fn] = stats
		}
		if passed {
			stats.Passed++
		} else {
			stats.Failed++
		}
	}
}

// FailingFunctions returns sorted list of functions with >0 failures and 0 passes.
func (fc *FunctionCoverage) FailingFunctions() []string {
	fc.mu.Lock()
	defer fc.mu.Unlock()
	var failing []string
	for fn, stats := range fc.Tested {
		if stats.Failed > 0 && stats.Passed == 0 {
			failing = append(failing, fn)
		}
	}
	sort.Strings(failing)
	return failing
}

// MarkBroken increments the broken counter for each function.
func (fc *FunctionCoverage) MarkBroken(functions []string) {
	fc.mu.Lock()
	defer fc.mu.Unlock()
	for _, fn := range functions {
		fc.Broken[fn]++
	}
}

// MarkFixed removes functions from the broken set.
func (fc *FunctionCoverage) MarkFixed(functions []string) {
	fc.mu.Lock()
	defer fc.mu.Unlock()
	for _, fn := range functions {
		delete(fc.Broken, fn)
	}
}

// BrokenList returns currently broken functions, sorted alphabetically.
func (fc *FunctionCoverage) BrokenList() []string {
	fc.mu.Lock()
	defer fc.mu.Unlock()
	var broken []string
	for fn := range fc.Broken {
		broken = append(broken, fn)
	}
	sort.Strings(broken)
	return broken
}

// IsHardBroken returns true if a function has hit the hard broken threshold.
func (fc *FunctionCoverage) IsHardBroken(fn string) bool {
	fc.mu.Lock()
	defer fc.mu.Unlock()
	return fc.Broken[fn] >= hardBrokenThreshold
}

// ExtractBrokenFunctionsFromMismatches identifies which functions are broken
// based on mismatch results.
// For #NAME? errors: extracts the outermost function (the unimplemented one).
// For value mismatches: extracts all functions from the formula.
func ExtractBrokenFunctionsFromMismatches(spec *TestSpec, mismatches []Mismatch) []string {
	formulaMap := buildFormulaMap(spec)
	seen := make(map[string]bool)
	var broken []string

	for _, m := range mismatches {
		formula, ok := formulaMap[m.Ref]
		if !ok {
			continue
		}

		if strings.Contains(m.Werkbook, "#NAME?") {
			// Outermost function is likely unimplemented.
			if fn := extractOutermostFunction(formula); fn != "" && !seen[fn] {
				seen[fn] = true
				broken = append(broken, fn)
			}
		} else {
			// Value mismatch — all functions might be suspect.
			for _, fn := range ExtractFunctionsFromFormula(formula) {
				if !seen[fn] {
					seen[fn] = true
					broken = append(broken, fn)
				}
			}
		}
	}

	sort.Strings(broken)
	return broken
}

// buildFormulaMap creates a map from check ref to formula string.
func buildFormulaMap(spec *TestSpec) map[string]string {
	m := make(map[string]string)
	if spec == nil {
		return m
	}
	for _, sheet := range spec.Sheets {
		for _, cell := range sheet.Cells {
			if cell.Formula == "" {
				continue
			}
			// Map both "Sheet1!A1" and "A1" forms.
			m[sheet.Name+"!"+cell.Ref] = cell.Formula
			m[cell.Ref] = cell.Formula
		}
	}
	return m
}

// extractOutermostFunction returns the first function name in a formula.
func extractOutermostFunction(formula string) string {
	upper := strings.ToUpper(formula)
	matches := funcNameRegex.FindStringSubmatch(upper)
	if len(matches) >= 2 {
		return matches[1]
	}
	return ""
}

// LeastTestedFunctions returns the n functions that have been tested the fewest times.
// Considers only implemented functions. Returns sorted by test count (ascending).
func (fc *FunctionCoverage) LeastTestedFunctions(n int) []string {
	fc.mu.Lock()
	defer fc.mu.Unlock()

	type funcCount struct {
		name  string
		count int
	}

	var counts []funcCount
	for _, fn := range KnownFunctionsList {
		if !ImplementedFunctions[fn] {
			continue
		}
		total := 0
		if stats, ok := fc.Tested[fn]; ok {
			total = stats.Passed + stats.Failed
		}
		counts = append(counts, funcCount{name: fn, count: total})
	}

	sort.Slice(counts, func(i, j int) bool {
		if counts[i].count != counts[j].count {
			return counts[i].count < counts[j].count
		}
		return counts[i].name < counts[j].name
	})

	result := make([]string, 0, n)
	for i := 0; i < n && i < len(counts); i++ {
		result = append(result, counts[i].name)
	}
	return result
}

// Summary formats a comprehensive coverage report.
func (fc *FunctionCoverage) Summary() string {
	fc.mu.Lock()
	defer fc.mu.Unlock()

	var passing, failing, mixed []string
	for fn, stats := range fc.Tested {
		switch {
		case stats.Failed == 0 && stats.Passed > 0:
			passing = append(passing, fmt.Sprintf("%s(%d/%d)", fn, stats.Passed, stats.Passed))
		case stats.Passed == 0 && stats.Failed > 0:
			failing = append(failing, fmt.Sprintf("%s(%d/%d)", fn, 0, stats.Failed))
		default:
			mixed = append(mixed, fmt.Sprintf("%s(%d/%d)", fn, stats.Passed, stats.Passed+stats.Failed))
		}
	}
	sort.Strings(passing)
	sort.Strings(failing)
	sort.Strings(mixed)

	// Find untested functions.
	var untested []string
	for _, fn := range KnownFunctionsList {
		if _, ok := fc.Tested[fn]; !ok {
			untested = append(untested, fn)
		}
	}

	totalKnown := len(KnownFunctionsList)
	testedCount := len(fc.Tested)

	var sb strings.Builder
	fmt.Fprintf(&sb, "=== Function Coverage (%d/%d tested) ===\n", testedCount, totalKnown)
	if len(passing) > 0 {
		fmt.Fprintf(&sb, "Passing (%d): %s\n", len(passing), strings.Join(passing, " "))
	}
	if len(failing) > 0 {
		fmt.Fprintf(&sb, "Failing (%d): %s\n", len(failing), strings.Join(failing, " "))
	}
	if len(mixed) > 0 {
		fmt.Fprintf(&sb, "Mixed (%d): %s\n", len(mixed), strings.Join(mixed, " "))
	}
	if len(untested) > 0 {
		fmt.Fprintf(&sb, "Untested (%d): %s\n", len(untested), strings.Join(untested, ", "))
	}
	return sb.String()
}
