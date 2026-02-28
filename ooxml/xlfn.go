package ooxml

import (
	"regexp"
	"strings"
)

// xlfnPrefix maps function names to their required OOXML prefix.
// These are Excel functions added after the original OOXML specification
// that require a special prefix in the XML to be recognized by Excel.
var xlfnPrefix = map[string]string{
	"CONCAT":   "_xlfn.",
	"DAYS":     "_xlfn.",
	"IFERROR":  "_xlfn.",
	"IFNA":     "_xlfn.",
	"IFS":      "_xlfn.",
	"MAXIFS":   "_xlfn.",
	"MINIFS":   "_xlfn.",
	"SORT":     "_xlfn._xlws.",
	"SWITCH":   "_xlfn.",
	"TEXTJOIN": "_xlfn.",
	"XLOOKUP":  "_xlfn.",
	"XOR":      "_xlfn.",
}

// xlfnPattern matches function names that need an OOXML prefix, followed by '('.
var xlfnPattern = regexp.MustCompile(
	`(?i)\b(CONCAT|DAYS|IFERROR|IFNA|IFS|MAXIFS|MINIFS|SORT|SWITCH|TEXTJOIN|XLOOKUP|XOR)\s*\(`,
)

// AddFormulaPrefix inserts OOXML-required _xlfn. prefixes into a formula string.
func AddFormulaPrefix(formula string) string {
	return xlfnPattern.ReplaceAllStringFunc(formula, func(match string) string {
		parenIdx := strings.Index(match, "(")
		name := strings.ToUpper(strings.TrimSpace(match[:parenIdx]))
		prefix, ok := xlfnPrefix[name]
		if !ok {
			return match
		}
		return prefix + name + match[parenIdx:]
	})
}

// StripFormulaPrefix removes OOXML _xlfn. prefixes from a formula string.
func StripFormulaPrefix(formula string) string {
	// Strip the longer prefix first to avoid partial replacement.
	formula = strings.ReplaceAll(formula, "_xlfn._xlws.", "")
	formula = strings.ReplaceAll(formula, "_xlfn.", "")
	return formula
}
