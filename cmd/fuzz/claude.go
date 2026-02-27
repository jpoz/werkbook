package main

import (
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"strings"
)

// knownFunctionsList is the full list from formula/compiler.go for the generation prompt.
var knownFunctionsList = []string{
	"ABS", "ACOS", "AND", "ASIN", "ATAN", "ATAN2", "AVERAGE", "AVERAGEIF",
	"AVERAGEIFS", "CEILING", "CHAR", "CHOOSE", "CLEAN", "CODE", "COLUMN",
	"COLUMNS", "CONCAT", "CONCATENATE", "COS", "COUNT", "COUNTA", "COUNTBLANK",
	"COUNTIF", "COUNTIFS", "DATE", "DAY", "EXACT", "EXP", "FIND", "FLOOR",
	"HLOOKUP", "HOUR", "IF", "IFERROR", "IFNA", "INDEX", "INT",
	"ISBLANK", "ISERR", "ISERROR", "ISNA", "ISNUMBER", "ISTEXT", "LARGE",
	"LEFT", "LEN", "LN", "LOG", "LOG10", "LOOKUP", "LOWER", "MATCH", "MAX",
	"MEDIAN", "MID", "MIN", "MINUTE", "MOD", "MONTH", "NOT", "OR",
	"PI", "POWER", "PRODUCT", "PROPER", "REPLACE",
	"REPT", "RIGHT", "ROUND", "ROUNDDOWN", "ROUNDUP", "ROW", "ROWS", "SEARCH",
	"SECOND", "SIN", "SMALL", "SORT", "SQRT", "SUBSTITUTE", "SUM", "SUMIF",
	"SUMIFS", "SUMPRODUCT", "TAN", "TEXT", "TIME", "TRIM", "UPPER",
	"VALUE", "VLOOKUP", "XLOOKUP", "XOR", "YEAR",
}

// generateSpec shells out to `claude -p` to generate a test spec.
func generateSpec(seed string, verbose bool) (*TestSpec, error) {
	prompt := buildGenerationPrompt(seed)

	if verbose {
		fmt.Println("  Generating spec with claude...")
	}

	output, err := runClaude(prompt)
	if err != nil {
		return nil, fmt.Errorf("claude generate: %w", err)
	}

	// Extract JSON from the output (claude may wrap it in markdown).
	jsonStr := extractJSON(output)
	if jsonStr == "" {
		return nil, fmt.Errorf("no JSON found in claude output:\n%s", truncate(output, 500))
	}

	var spec TestSpec
	if err := json.Unmarshal([]byte(jsonStr), &spec); err != nil {
		return nil, fmt.Errorf("parse generated spec: %w\njson: %s", err, truncate(jsonStr, 500))
	}

	if err := validateSpec(&spec); err != nil {
		return nil, fmt.Errorf("invalid generated spec: %w", err)
	}

	return &spec, nil
}

// suggestFix shells out to `claude -p` to suggest a fix for mismatches.
func suggestFix(spec *TestSpec, mismatches []mismatch, verbose bool) (string, error) {
	prompt := buildFixPrompt(spec, mismatches)

	if verbose {
		fmt.Println("  Asking claude for fix suggestions...")
	}

	output, err := runClaude(prompt)
	if err != nil {
		return "", fmt.Errorf("claude fix: %w", err)
	}

	return output, nil
}

// runClaude executes `claude -p` with the given prompt and returns stdout.
func runClaude(prompt string) (string, error) {
	cmd := exec.Command("claude", "-p", prompt)
	// Clear CLAUDECODE env var to allow running inside a Claude Code session.
	cmd.Env = filterEnv(os.Environ(), "CLAUDECODE")
	out, err := cmd.CombinedOutput()
	if err != nil {
		return "", fmt.Errorf("claude command failed: %w\noutput: %s", err, truncate(string(out), 500))
	}
	return string(out), nil
}

// buildGenerationPrompt creates the prompt for generating a test spec.
func buildGenerationPrompt(seed string) string {
	var sb strings.Builder
	sb.WriteString("Generate a JSON test spec for testing Excel formula evaluation in a Go spreadsheet library.\n\n")
	sb.WriteString("The spec must be valid JSON matching this exact schema:\n")
	sb.WriteString(`{
  "name": "test_<descriptive_name>",
  "sheets": [
    {
      "name": "Sheet1",
      "cells": [
        {"ref": "A1", "value": 10, "type": "number"},
        {"ref": "A2", "value": "hello", "type": "string"},
        {"ref": "A3", "value": true, "type": "bool"},
        {"ref": "B1", "formula": "SUM(A1:A3)"}
      ]
    }
  ],
  "checks": [
    {"ref": "Sheet1!B1", "expected": "10", "type": "number"}
  ]
}`)
	sb.WriteString("\n\n")

	if seed != "" {
		sb.WriteString(fmt.Sprintf("Focus on testing %s functions. ", seed))
	}

	sb.WriteString("Available functions (pick 3-8 to test per spec):\n")
	sb.WriteString(strings.Join(knownFunctionsList, ", "))
	sb.WriteString("\n\n")

	sb.WriteString("Requirements:\n")
	sb.WriteString("- Create edge cases: empty cells, zero values, negative numbers, large numbers, mixed types\n")
	sb.WriteString("- Test nested formulas (e.g., IF(SUM(A1:A3)>10, AVERAGE(A1:A3), 0))\n")
	sb.WriteString("- Test cross-cell references and ranges\n")
	sb.WriteString("- Include 5-15 cells and 3-8 checks\n")
	sb.WriteString("- Use realistic but tricky inputs that might expose bugs\n")
	sb.WriteString("- DO NOT use RAND, RANDBETWEEN, NOW, TODAY, or INDIRECT\n")
	sb.WriteString("- The 'expected' field in checks should be the string representation of the expected result\n")
	sb.WriteString("- For numbers, use the simplest representation (e.g., \"10\" not \"10.0\")\n")
	sb.WriteString("- For booleans, use \"TRUE\" or \"FALSE\"\n")
	sb.WriteString("- Cell refs in formulas should NOT include the sheet name if on the same sheet\n")
	sb.WriteString("- Multi-sheet tests are welcome but keep them simple\n")
	sb.WriteString("\n")
	sb.WriteString("Output ONLY the JSON, no explanation or markdown fences.\n")

	return sb.String()
}

// buildFixPrompt creates the prompt for suggesting a fix.
func buildFixPrompt(spec *TestSpec, mismatches []mismatch) string {
	specJSON, _ := json.MarshalIndent(spec, "", "  ")

	var sb strings.Builder
	sb.WriteString("I'm testing a Go spreadsheet formula engine (github.com/jpoz/werkbook) against LibreOffice.\n\n")
	sb.WriteString("Test spec:\n```json\n")
	sb.Write(specJSON)
	sb.WriteString("\n```\n\n")
	sb.WriteString("Mismatches found (LibreOffice is the ground truth):\n")
	for _, m := range mismatches {
		fmt.Fprintf(&sb, "- %s: werkbook=%q, libreoffice=%q (%s)\n", m.Ref, m.Werkbook, m.LibreOff, m.Reason)
	}
	sb.WriteString("\nThe formula engine code is in the `formula/` package. Key files:\n")
	sb.WriteString("- formula/functions.go: Built-in function implementations\n")
	sb.WriteString("- formula/vm.go: Bytecode VM that executes compiled formulas\n")
	sb.WriteString("- formula/compiler.go: Compiles AST to bytecode\n")
	sb.WriteString("- formula/parser.go: Parses formula strings to AST\n\n")
	sb.WriteString("Analyze the mismatches and suggest what might be wrong in the formula engine.\n")
	sb.WriteString("Focus on the most likely root cause and suggest a specific fix.\n")

	return sb.String()
}

// extractJSON finds the first JSON object in the output.
func extractJSON(s string) string {
	// Try to find raw JSON first.
	start := strings.Index(s, "{")
	if start < 0 {
		return ""
	}

	// Find matching closing brace.
	depth := 0
	for i := start; i < len(s); i++ {
		switch s[i] {
		case '{':
			depth++
		case '}':
			depth--
			if depth == 0 {
				return s[start : i+1]
			}
		}
	}
	return ""
}

// truncate shortens a string to maxLen, adding "..." if truncated.
func truncate(s string, maxLen int) string {
	if len(s) <= maxLen {
		return s
	}
	return s[:maxLen] + "..."
}

// filterEnv returns a copy of env with the named variable removed.
func filterEnv(env []string, name string) []string {
	prefix := name + "="
	out := make([]string, 0, len(env))
	for _, e := range env {
		if !strings.HasPrefix(e, prefix) {
			out = append(out, e)
		}
	}
	return out
}
