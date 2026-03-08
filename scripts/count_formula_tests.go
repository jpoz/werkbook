// Script to regenerate FORMULAS.md from the source of truth:
//   1. Supported functions = all Register() calls in formula/*.go
//   2. Test counts = formula strings + TestFUNC_ patterns in *_test.go
//   3. Unsupported functions = those in the known Excel list but not registered
//
// Usage: go run scripts/count_formula_tests.go [--check]
//   (default) regenerate FORMULAS.md
//   --check   exit 1 if FORMULAS.md would change (for CI)
package main

import (
	"bufio"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

// categoryMap maps source file suffix to display category.
var categoryMap = map[string]string{
	"date":    "Date & Time",
	"eng":     "Engineering",
	"finance": "Financial",
	"info":    "Information",
	"logic":   "Logical",
	"lookup":  "Lookup & Reference",
	"math":    "Math & Trig",
	"stat":    "Statistical",
	"text":    "Text",
}

type funcInfo struct {
	Name     string
	Category string
	Tests    int
}

func main() {
	check := len(os.Args) > 1 && os.Args[1] == "--check"

	// Step 1: discover supported functions from Register() calls
	supported := discoverSupported()

	// Step 2: count tests
	testCounts := countTests(supported)
	for name := range testCounts {
		if fi, ok := supported[name]; ok {
			fi.Tests = testCounts[name]
			supported[name] = fi
		}
	}

	// Step 3: load known Excel functions and compute unsupported
	unsupported := computeUnsupported(supported)

	// Step 4: generate FORMULAS.md content
	content := generateMarkdown(supported, unsupported)

	if check {
		existing, err := os.ReadFile("FORMULAS.md")
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error reading FORMULAS.md: %v\n", err)
			os.Exit(1)
		}
		if string(existing) == content {
			fmt.Println("FORMULAS.md is up to date.")
			os.Exit(0)
		}
		fmt.Println("FORMULAS.md is out of date. Run: go run scripts/count_formula_tests.go")
		os.Exit(1)
	}

	if err := os.WriteFile("FORMULAS.md", []byte(content), 0644); err != nil {
		fmt.Fprintf(os.Stderr, "Error writing FORMULAS.md: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("Updated FORMULAS.md: %d supported, %d unsupported\n",
		len(supported), len(unsupported))
}

// discoverSupported scans formula/*.go (non-test) for Register("NAME", ...) calls.
func discoverSupported() map[string]funcInfo {
	re := regexp.MustCompile(`Register\("([A-Z][A-Z0-9.]*)"`)
	result := map[string]funcInfo{}

	files, _ := filepath.Glob("formula/functions_*.go")
	for _, path := range files {
		if strings.HasSuffix(path, "_test.go") {
			continue
		}
		// Extract category from filename: functions_math.go -> "math"
		base := filepath.Base(path)
		cat := strings.TrimPrefix(base, "functions_")
		cat = strings.TrimSuffix(cat, ".go")
		displayCat := categoryMap[cat]
		if displayCat == "" {
			displayCat = strings.ToUpper(cat[:1]) + cat[1:]
		}

		f, err := os.Open(path)
		if err != nil {
			continue
		}
		scanner := bufio.NewScanner(f)
		for scanner.Scan() {
			m := re.FindStringSubmatch(scanner.Text())
			if m != nil {
				result[m[1]] = funcInfo{Name: m[1], Category: displayCat}
			}
		}
		f.Close()
	}
	return result
}

// countTests counts test cases per formula function.
func countTests(supported map[string]funcInfo) map[string]int {
	counts := map[string]int{}

	// Strategy 1: formula strings in test files
	countFormulaStrings(counts, supported)

	// Strategy 2: TestFUNC_Name test function names
	countTestFunctions(counts, supported)

	// Strategy 3: count t.Run sub-tests inside TestFUNCNAME functions for
	// functions not yet counted (e.g. TRUE, FALSE, MODE — which are either
	// skipped by strategy 1 or use dynamic formula construction).
	countSubTests(counts, supported)

	return counts
}

func countFormulaStrings(counts map[string]int, supported map[string]funcInfo) {
	reFormula := regexp.MustCompile("[\"` ]{1}=?([A-Z][A-Z0-9.]*)" + `\(`)

	skip := map[string]bool{
		"TRUE": true, "FALSE": true, "DIV": true,
		"REF": true, "NULL": true, "NAME": true,
	}

	testFiles, _ := filepath.Glob("formula/*_test.go")
	rootFiles, _ := filepath.Glob("*_test.go")
	testFiles = append(testFiles, rootFiles...)

	for _, path := range testFiles {
		f, err := os.Open(path)
		if err != nil {
			continue
		}
		scanner := bufio.NewScanner(f)
		for scanner.Scan() {
			line := scanner.Text()
			trimmed := strings.TrimSpace(line)
			if strings.HasPrefix(trimmed, "//") {
				continue
			}

			matches := reFormula.FindAllStringSubmatch(trimmed, -1)
			if len(matches) == 0 {
				continue
			}

			isTestCase := strings.Contains(trimmed, "evalCompile") ||
				strings.Contains(trimmed, `formula:`) ||
				strings.Contains(trimmed, `formula :`) ||
				strings.Contains(trimmed, "SetFormula") ||
				(strings.HasPrefix(trimmed, "{") && (strings.Contains(trimmed, `"`) || strings.Contains(trimmed, "`"))) ||
				strings.Contains(trimmed, `name:`)

			if !isTestCase {
				continue
			}

			funcName := matches[0][1]
			if skip[funcName] || !isSupported(funcName, supported) {
				continue
			}
			counts[funcName]++
		}
		f.Close()
	}
}

func countTestFunctions(counts map[string]int, supported map[string]funcInfo) {
	reTestFunc := regexp.MustCompile(`^func\s+Test([A-Z][A-Z0-9_]*?)_\w+\s*\(`)

	testFiles, _ := filepath.Glob("formula/*_test.go")
	for _, path := range testFiles {
		f, err := os.Open(path)
		if err != nil {
			continue
		}
		scanner := bufio.NewScanner(f)
		for scanner.Scan() {
			m := reTestFunc.FindStringSubmatch(scanner.Text())
			if m == nil {
				continue
			}
			funcName := m[1]
			if isSupported(funcName, supported) {
				counts[funcName]++
				continue
			}
			dotted := strings.ReplaceAll(funcName, "_", ".")
			if isSupported(dotted, supported) {
				counts[dotted]++
			}
		}
		f.Close()
	}
}

// countSubTests finds TestFUNCNAME functions (no underscore suffix) and counts
// test cases inside them. This catches functions skipped by strategy 1
// (TRUE, FALSE) and functions that build formula strings dynamically (MODE).
// Only adds counts for functions that have zero tests from earlier strategies.
//
// It counts test cases by looking for:
//   - Table-driven test entries: lines starting with { that contain a string literal
//   - Named t.Run calls with string-literal names that directly test (not containers)
//
// A t.Run is considered a "container" (not counted) if its body includes a
// []struct test table or a for-range loop. Only leaf t.Run calls count.
//
// For functions tested together (e.g. MODE and MODE.SNGL in a single TestMODE),
// it also detects for-loop multipliers like: for _, fn := range []string{"MODE", "MODE.SNGL"}
func countSubTests(counts map[string]int, supported map[string]funcInfo) {
	reFunc := regexp.MustCompile(`^func\s+Test([A-Z][A-Z0-9]*)\s*\(`)
	// Matches for loops that iterate over function names to test multiple aliases
	reFuncLoop := regexp.MustCompile(`for\s+.+range\s+\[\]string\{([^}]+)\}`)

	testFiles, _ := filepath.Glob("formula/*_test.go")
	for _, path := range testFiles {
		data, err := os.ReadFile(path)
		if err != nil {
			continue
		}
		lines := strings.Split(string(data), "\n")

		for i := 0; i < len(lines); i++ {
			m := reFunc.FindStringSubmatch(lines[i])
			if m == nil {
				continue
			}
			testName := m[1]

			// Resolve to a supported function name
			funcName := ""
			if isSupported(testName, supported) {
				funcName = testName
			} else {
				dotted := strings.ReplaceAll(testName, "_", ".")
				if isSupported(dotted, supported) {
					funcName = dotted
				}
			}
			if funcName == "" {
				continue
			}

			// Only fill in if strategy 1+2 found nothing
			if counts[funcName] > 0 {
				continue
			}

			// Find the end of the function body using brace depth
			depth := 0
			started := false
			endLine := len(lines) - 1
			for j := i; j < len(lines); j++ {
				line := lines[j]
				depth += strings.Count(line, "{") - strings.Count(line, "}")
				if !started && depth > 0 {
					started = true
				}
				if started && depth <= 0 {
					endLine = j
					break
				}
			}

			body := lines[i : endLine+1]
			bodyStr := strings.Join(body, "\n")

			// Check if there's a function-name loop multiplier
			extraFuncs := []string{}
			if lm := reFuncLoop.FindStringSubmatch(bodyStr); lm != nil {
				names := strings.Split(lm[1], ",")
				for _, n := range names {
					n = strings.TrimSpace(n)
					n = strings.Trim(n, `"`)
					n = strings.Trim(n, "`")
					if n != "" && n != funcName {
						extraFuncs = append(extraFuncs, n)
					}
				}
			}

			total := countTestCasesInBody(body)
			if total > 0 {
				counts[funcName] += total
				for _, extra := range extraFuncs {
					if isSupported(extra, supported) && counts[extra] == 0 {
						counts[extra] += total
					}
				}
			}
		}
	}
}

// countTestCasesInBody counts test cases in a function body by walking through
// t.Run calls and test table entries. It distinguishes between:
//   - Container t.Run calls (those wrapping a test table + for loop): not counted,
//     but their table entries ARE counted
//   - Leaf t.Run calls (those containing direct test logic): counted as 1 each
//   - Table entries ({...} lines inside a []struct literal): counted as 1 each
func countTestCasesInBody(lines []string) int {
	reNamedRun := regexp.MustCompile(`t\.Run\(["\x60]`)
	reTableEntry := regexp.MustCompile(`^\s*\{["\x60]`)
	reStructDecl := regexp.MustCompile(`\[\]struct`)

	total := 0

	for idx := 0; idx < len(lines); idx++ {
		trimmed := strings.TrimSpace(lines[idx])
		if strings.HasPrefix(trimmed, "//") {
			continue
		}

		// When we find a t.Run with a string literal name, check if it's
		// a container (has a []struct inside) or a leaf test.
		if reNamedRun.MatchString(trimmed) {
			// Find the extent of this t.Run block
			blockStart := idx
			depth := 0
			blockStarted := false
			blockEnd := idx

			for j := blockStart; j < len(lines); j++ {
				line := lines[j]
				depth += strings.Count(line, "{") - strings.Count(line, "}")
				if !blockStarted && depth > 0 {
					blockStarted = true
				}
				if blockStarted && depth <= 0 {
					blockEnd = j
					break
				}
			}

			// Check if this block contains a []struct (table-driven test container)
			blockLines := lines[blockStart : blockEnd+1]
			isContainer := false
			for _, bl := range blockLines {
				if reStructDecl.MatchString(bl) {
					isContainer = true
					break
				}
			}

			if isContainer {
				// Count table entries inside this container
				for _, bl := range blockLines {
					bt := strings.TrimSpace(bl)
					if strings.HasPrefix(bt, "//") {
						continue
					}
					if reTableEntry.MatchString(bt) {
						total++
					}
				}
			} else {
				// Leaf t.Run — counts as one test case
				total++
			}

			// Skip past this block
			idx = blockEnd
			continue
		}
	}

	return total
}

func isSupported(name string, supported map[string]funcInfo) bool {
	_, ok := supported[name]
	return ok
}

// computeUnsupported reads the existing FORMULAS.md unsupported list and removes
// any that are now supported.
func computeUnsupported(supported map[string]funcInfo) []funcInfo {
	f, err := os.Open("FORMULAS.md")
	if err != nil {
		return nil
	}
	defer f.Close()

	var unsupported []funcInfo
	inUnsupported := false
	reRow := regexp.MustCompile(`^\|\s*([A-Z][A-Z0-9.]*)\s*\|\s*([^|]+?)\s*\|`)

	scanner := bufio.NewScanner(f)
	for scanner.Scan() {
		line := scanner.Text()
		if strings.HasPrefix(line, "# Unsupported") {
			inUnsupported = true
			continue
		}
		if !inUnsupported {
			continue
		}
		m := reRow.FindStringSubmatch(line)
		if m == nil {
			continue
		}
		name := m[1]
		cat := strings.TrimSpace(m[2])
		if _, ok := supported[name]; ok {
			continue // now supported, skip
		}
		unsupported = append(unsupported, funcInfo{Name: name, Category: cat})
	}
	return unsupported
}

func generateMarkdown(supported map[string]funcInfo, unsupported []funcInfo) string {
	var b strings.Builder

	// Sort supported by name
	var funcs []funcInfo
	for _, fi := range supported {
		funcs = append(funcs, fi)
	}
	sort.Slice(funcs, func(i, j int) bool {
		return funcs[i].Name < funcs[j].Name
	})

	b.WriteString("# Supported Formulas\n\n")
	b.WriteString(fmt.Sprintf("Werkbook supports **%d** Excel formula functions.\n\n", len(funcs)))
	b.WriteString("| Function | Category | Tests |\n")
	b.WriteString("|----------|----------|------:|\n")
	for _, fi := range funcs {
		tests := "-"
		if fi.Tests > 0 {
			tests = strconv.Itoa(fi.Tests)
		}
		b.WriteString(fmt.Sprintf("| %s | %s | %s |\n", fi.Name, fi.Category, tests))
	}

	if len(unsupported) > 0 {
		sort.Slice(unsupported, func(i, j int) bool {
			return unsupported[i].Name < unsupported[j].Name
		})
		b.WriteString(fmt.Sprintf("\n# Unsupported Formulas\n\n"))
		b.WriteString(fmt.Sprintf("The following **%d** Excel functions are not yet supported.\n\n", len(unsupported)))
		b.WriteString("| Function | Category |\n")
		b.WriteString("|----------|----------|\n")
		for _, fi := range unsupported {
			b.WriteString(fmt.Sprintf("| %s | %s |\n", fi.Name, fi.Category))
		}
	}

	return b.String()
}
