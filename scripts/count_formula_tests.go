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
