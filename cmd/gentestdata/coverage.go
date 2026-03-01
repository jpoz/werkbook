package main

import (
	"fmt"
	"sort"

	"github.com/jpoz/werkbook/internal/fuzz"
)

// coverageReport checks all ImplementedFunctions against test case coverage.
func coverageReport(scenarios []ScenarioSet) {
	// Collect all function names from test cases
	covered := make(map[string]bool)
	for _, ss := range scenarios {
		for _, tc := range ss.Cases {
			for _, fn := range tc.FuncNames {
				covered[fn] = true
			}
		}
		for _, sd := range ss.Sheets {
			for _, tc := range sd.Cases {
				for _, fn := range tc.FuncNames {
					covered[fn] = true
				}
			}
		}
	}

	excluded := fuzz.LocalExcelExcludedFunctions

	var testedFuncs []string
	var excludedFuncs []string
	var missingFuncs []string

	for fn := range fuzz.ImplementedFunctions {
		if excluded[fn] {
			excludedFuncs = append(excludedFuncs, fn)
		} else if covered[fn] {
			testedFuncs = append(testedFuncs, fn)
		} else {
			missingFuncs = append(missingFuncs, fn)
		}
	}

	sort.Strings(testedFuncs)
	sort.Strings(excludedFuncs)
	sort.Strings(missingFuncs)

	total := len(fuzz.ImplementedFunctions)
	testable := total - len(excludedFuncs)

	fmt.Printf("\n=== Coverage Report ===\n")
	fmt.Printf("Coverage: %d/%d testable functions covered\n", len(testedFuncs), testable)
	fmt.Printf("Total implemented: %d\n", total)

	if len(excludedFuncs) > 0 {
		fmt.Printf("\nExcluded (non-deterministic/volatile): %d\n", len(excludedFuncs))
		for _, fn := range excludedFuncs {
			fmt.Printf("  - %s\n", fn)
		}
	}

	if len(missingFuncs) > 0 {
		fmt.Printf("\nMissing (need test cases): %d\n", len(missingFuncs))
		for _, fn := range missingFuncs {
			fmt.Printf("  - %s\n", fn)
		}
	}

	if len(missingFuncs) == 0 {
		fmt.Println("\nAll testable functions are covered!")
	}
}
