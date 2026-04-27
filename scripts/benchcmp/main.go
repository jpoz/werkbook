package main

import (
	"bufio"
	"flag"
	"fmt"
	"os"
	"sort"
	"strconv"
	"strings"
)

type benchMetrics struct {
	nsPerOp   float64
	bytesOp   float64
	allocsOp  float64
	hasNS     bool
	hasBytes  bool
	hasAllocs bool
}

type regression struct {
	name    string
	metric  string
	base    float64
	current float64
	ratio   float64
}

func main() {
	basePath := flag.String("base", "", "base benchmark output path")
	currentPath := flag.String("current", "", "current benchmark output path")
	threshold := flag.Float64("threshold", 1.5, "regression threshold as a ratio")
	outPath := flag.String("out", "", "optional markdown output path")
	flag.Parse()

	if *basePath == "" || *currentPath == "" {
		exitWith("base and current are required")
	}
	if *threshold <= 1 {
		exitWith("threshold must be greater than 1")
	}

	base, err := parseBenchmarks(*basePath)
	if err != nil {
		exitWith("parse base: %v", err)
	}
	current, err := parseBenchmarks(*currentPath)
	if err != nil {
		exitWith("parse current: %v", err)
	}
	if len(base) == 0 {
		exitWith("no base benchmarks found in %s", *basePath)
	}
	if len(current) == 0 {
		exitWith("no current benchmarks found in %s", *currentPath)
	}

	regressions := compareBenchmarks(base, current, *threshold)
	summary := renderMarkdown(base, current, regressions, *threshold)
	fmt.Print(summary)
	if *outPath != "" {
		if err := os.WriteFile(*outPath, []byte(summary), 0o644); err != nil {
			exitWith("write summary: %v", err)
		}
	}
	if len(regressions) > 0 {
		os.Exit(1)
	}
}

func parseBenchmarks(path string) (map[string]benchMetrics, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	out := make(map[string]benchMetrics)
	scanner := bufio.NewScanner(f)
	for scanner.Scan() {
		fields := strings.Fields(scanner.Text())
		if len(fields) == 0 || !strings.HasPrefix(fields[0], "Benchmark") {
			continue
		}
		name := normalizeBenchmarkName(fields[0])
		var metrics benchMetrics
		for i := 1; i < len(fields); i++ {
			switch fields[i] {
			case "ns/op":
				v, ok := parsePrevFloat(fields, i)
				if ok {
					metrics.nsPerOp = v
					metrics.hasNS = true
				}
			case "B/op":
				v, ok := parsePrevFloat(fields, i)
				if ok {
					metrics.bytesOp = v
					metrics.hasBytes = true
				}
			case "allocs/op":
				v, ok := parsePrevFloat(fields, i)
				if ok {
					metrics.allocsOp = v
					metrics.hasAllocs = true
				}
			}
		}
		if metrics.hasNS || metrics.hasBytes || metrics.hasAllocs {
			out[name] = metrics
		}
	}
	return out, scanner.Err()
}

func normalizeBenchmarkName(name string) string {
	if idx := strings.LastIndex(name, "-"); idx > len("Benchmark") {
		suffix := name[idx+1:]
		if _, err := strconv.Atoi(suffix); err == nil {
			return name[:idx]
		}
	}
	return name
}

func parsePrevFloat(fields []string, idx int) (float64, bool) {
	if idx == 0 {
		return 0, false
	}
	v, err := strconv.ParseFloat(strings.ReplaceAll(fields[idx-1], ",", ""), 64)
	return v, err == nil
}

func compareBenchmarks(base, current map[string]benchMetrics, threshold float64) []regression {
	var out []regression
	for _, name := range sortedBenchmarkNames(base) {
		baseMetrics := base[name]
		currentMetrics, ok := current[name]
		if !ok {
			continue
		}
		out = appendMetricRegression(out, name, "ns/op", baseMetrics.nsPerOp, currentMetrics.nsPerOp, baseMetrics.hasNS && currentMetrics.hasNS, threshold)
		out = appendMetricRegression(out, name, "B/op", baseMetrics.bytesOp, currentMetrics.bytesOp, baseMetrics.hasBytes && currentMetrics.hasBytes, threshold)
		out = appendMetricRegression(out, name, "allocs/op", baseMetrics.allocsOp, currentMetrics.allocsOp, baseMetrics.hasAllocs && currentMetrics.hasAllocs, threshold)
	}
	return out
}

func sortedBenchmarkNames(values map[string]benchMetrics) []string {
	names := make([]string, 0, len(values))
	for name := range values {
		names = append(names, name)
	}
	sort.Strings(names)
	return names
}

func appendMetricRegression(out []regression, name, metric string, base, current float64, ok bool, threshold float64) []regression {
	if !ok || base <= 0 {
		return out
	}
	ratio := current / base
	if ratio <= threshold {
		return out
	}
	return append(out, regression{name: name, metric: metric, base: base, current: current, ratio: ratio})
}

func renderMarkdown(base, current map[string]benchMetrics, regressions []regression, threshold float64) string {
	var b strings.Builder
	common := commonBenchmarkCount(base, current)
	fmt.Fprintf(&b, "\n### Benchmark Comparison\n\n")
	fmt.Fprintf(&b, "Regression threshold: %.2fx\n\n", threshold)
	if common == 0 {
		fmt.Fprintf(&b, "No common benchmarks were found between base and current outputs.\n")
		return b.String()
	}
	if len(regressions) == 0 {
		fmt.Fprintf(&b, "No benchmark metric exceeded the threshold across %d common benchmarks.\n", common)
		return b.String()
	}
	fmt.Fprintf(&b, "| Benchmark | Metric | Base | Current | Ratio |\n")
	fmt.Fprintf(&b, "| --- | ---: | ---: | ---: | ---: |\n")
	for _, r := range regressions {
		fmt.Fprintf(&b, "| `%s` | `%s` | %.0f | %.0f | %.2fx |\n", r.name, r.metric, r.base, r.current, r.ratio)
	}
	fmt.Fprintf(&b, "\nCompared %d common benchmarks.\n", common)
	return b.String()
}

func commonBenchmarkCount(base, current map[string]benchMetrics) int {
	count := 0
	for name := range base {
		if _, ok := current[name]; ok {
			count++
		}
	}
	return count
}

func exitWith(format string, args ...any) {
	fmt.Fprintf(os.Stderr, format+"\n", args...)
	os.Exit(2)
}
