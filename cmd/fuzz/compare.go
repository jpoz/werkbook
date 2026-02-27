package main

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// mismatch describes a single comparison failure.
type mismatch struct {
	Ref      string
	Werkbook string
	Oracle   string
	Reason   string
}

// compareResults compares werkbook results against LibreOffice results.
// Returns nil if all match.
func compareResults(checks []CheckSpec, wb, lo []buildResult) []mismatch {
	var mismatches []mismatch
	for i, check := range checks {
		if i >= len(wb) || i >= len(lo) {
			mismatches = append(mismatches, mismatch{
				Ref:    check.Ref,
				Reason: "missing result",
			})
			continue
		}

		wbVal := wb[i].Value
		loVal := lo[i].Value

		if match, reason := valuesMatch(wbVal, loVal); !match {
			mismatches = append(mismatches, mismatch{
				Ref:      check.Ref,
				Werkbook: wbVal,
				Oracle:   loVal,
				Reason:   reason,
			})
		}
	}
	return mismatches
}

// valuesMatch applies layered comparison strategies.
// Returns (true, "") if match, or (false, reason) if not.
func valuesMatch(wb, lo string) (bool, string) {
	// Exact match.
	if wb == lo {
		return true, ""
	}

	// Normalize and compare.
	wbNorm := normalizeValue(wb)
	loNorm := normalizeValue(lo)
	if wbNorm == loNorm {
		return true, ""
	}

	// Boolean normalization.
	if boolMatch(wb, lo) {
		return true, ""
	}

	// Error normalization.
	wbErr := normalizeError(wb)
	loErr := normalizeError(lo)
	if wbErr != "" && loErr != "" && wbErr == loErr {
		return true, ""
	}

	// Numeric tolerance.
	wbNum, wbOk := parseNumber(wbNorm)
	loNum, loOk := parseNumber(loNorm)
	if wbOk && loOk {
		if numericMatch(wbNum, loNum) {
			return true, ""
		}
		return false, fmt.Sprintf("numeric mismatch: werkbook=%g oracle=%g (diff=%g)", wbNum, loNum, math.Abs(wbNum-loNum))
	}

	return false, fmt.Sprintf("value mismatch: werkbook=%q oracle=%q", wb, lo)
}

// normalizeValue strips trailing zeros and whitespace.
func normalizeValue(s string) string {
	s = strings.TrimSpace(s)
	// Strip trailing zeros from decimal numbers.
	if strings.Contains(s, ".") {
		s = strings.TrimRight(s, "0")
		s = strings.TrimRight(s, ".")
	}
	return s
}

// boolMatch returns true if both values represent the same boolean.
func boolMatch(a, b string) bool {
	ab := parseBool(a)
	bb := parseBool(b)
	if ab == nil || bb == nil {
		return false
	}
	return *ab == *bb
}

// parseBool attempts to parse a value as a boolean.
func parseBool(s string) *bool {
	upper := strings.ToUpper(strings.TrimSpace(s))
	var b bool
	switch upper {
	case "TRUE", "1":
		b = true
	case "FALSE", "0":
		b = false
	default:
		return nil
	}
	return &b
}

// normalizeError converts LibreOffice error codes to Excel-style error strings.
func normalizeError(s string) string {
	s = strings.TrimSpace(s)

	// Already in Excel format.
	if strings.HasPrefix(s, "#") {
		return strings.ToUpper(s)
	}

	// LibreOffice Err:NNN format.
	if strings.HasPrefix(s, "Err:") {
		code := strings.TrimPrefix(s, "Err:")
		switch code {
		case "502":
			return "#VALUE!"
		case "504":
			return "#REF!"
		case "508":
			return "#NULL!"
		case "511":
			return "#DIV/0!"
		case "519":
			return "#VALUE!"
		case "521":
			return "#DIV/0!"
		case "522":
			return "#REF!"
		case "524":
			return "#REF!"
		case "525":
			return "#NAME?"
		case "532":
			return "#DIV/0!"
		case "533":
			return "#N/A"
		default:
			return "#VALUE!" // Fallback.
		}
	}
	return ""
}

// parseNumber tries to parse a string as float64.
func parseNumber(s string) (float64, bool) {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, false
	}
	f, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return 0, false
	}
	return f, true
}

// numericMatch checks if two numbers are within acceptable tolerance.
// Uses max(1e-9, 1e-6 * max(|a|, |b|)) as the tolerance.
func numericMatch(a, b float64) bool {
	diff := math.Abs(a - b)
	tol := math.Max(1e-9, 1e-6*math.Max(math.Abs(a), math.Abs(b)))
	return diff <= tol
}

// formatMismatches produces a human-readable summary of comparison failures.
// oracleName is used to label the oracle column (e.g. "libreoffice", "local-excel").
func formatMismatches(ms []mismatch, oracleName string) string {
	var sb strings.Builder
	for _, m := range ms {
		fmt.Fprintf(&sb, "  %s: %s\n", m.Ref, m.Reason)
		if m.Werkbook != "" || m.Oracle != "" {
			fmt.Fprintf(&sb, "    werkbook:    %q\n", m.Werkbook)
			fmt.Fprintf(&sb, "    %-12s %q\n", oracleName+":", m.Oracle)
		}
	}
	return sb.String()
}
