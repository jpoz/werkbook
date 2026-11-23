package bridge

import (
	"math"
	"strconv"
	"strings"

	werkbook "github.com/jpoz/werkbook"
)

// inputKind is the result class of the input rules.
type inputKind int

const (
	inputClear inputKind = iota
	inputFormula
	inputBool
	inputNumber
	inputString
)

// parsedInput is editor text after the input rules are applied.
type parsedInput struct {
	kind    inputKind
	formula string // with the leading '='
	b       bool
	num     float64
	str     string
}

// parseInput applies the input rules of cmd/web/README.md, in order:
// empty text clears, "=..." is a formula, TRUE/FALSE is a boolean, numeric
// text is a number, a leading apostrophe forces a string, all other text is a
// string. Booleans and numbers tolerate surrounding spaces, as a spreadsheet
// does; the string result keeps the text as typed.
func parseInput(input string) parsedInput {
	if input == "" {
		return parsedInput{kind: inputClear}
	}
	if strings.HasPrefix(input, "=") && len(input) > 1 {
		return parsedInput{kind: inputFormula, formula: input}
	}
	trimmed := strings.TrimSpace(input)
	switch {
	case strings.EqualFold(trimmed, "TRUE"):
		return parsedInput{kind: inputBool, b: true}
	case strings.EqualFold(trimmed, "FALSE"):
		return parsedInput{kind: inputBool, b: false}
	}
	if n, ok := parseNumber(trimmed); ok {
		return parsedInput{kind: inputNumber, num: n}
	}
	if rest, ok := strings.CutPrefix(input, "'"); ok {
		return parsedInput{kind: inputString, str: rest}
	}
	return parsedInput{kind: inputString, str: input}
}

// parseNumber accepts decimal and exponent notation only. strconv.ParseFloat
// also accepts "inf", "nan", hexadecimal floats and digit separators; a
// spreadsheet user means text when they type those.
func parseNumber(s string) (float64, bool) {
	if s == "" {
		return 0, false
	}
	for i := 0; i < len(s); i++ {
		switch c := s[i]; {
		case c >= '0' && c <= '9':
		case c == '.', c == '+', c == '-', c == 'e', c == 'E':
		default:
			return 0, false
		}
	}
	n, err := strconv.ParseFloat(s, 64)
	if err != nil || math.IsInf(n, 0) || math.IsNaN(n) {
		return 0, false
	}
	return n, true
}

// inputSnapshot returns the cell state that editor text gives when it is
// written to a cell with the state base. The style of base is kept, so a
// cleared cell keeps its style.
func inputSnapshot(base werkbook.CellSnapshot, input string) werkbook.CellSnapshot {
	p := parseInput(input)
	switch p.kind {
	case inputClear:
		return base.WithFormula("")
	case inputFormula:
		// WithFormula removes the single leading '='.
		return base.WithFormula(p.formula)
	case inputBool:
		return base.WithValue(werkbook.Value{Type: werkbook.TypeBool, Bool: p.b})
	case inputNumber:
		return base.WithValue(werkbook.Value{Type: werkbook.TypeNumber, Number: p.num})
	default:
		return base.WithValue(werkbook.Value{Type: werkbook.TypeString, String: p.str})
	}
}

// inputText returns the editor text for a value: the text that gives the same
// value when it goes through parseInput again. A string that the input rules
// would read as a different type gets a leading apostrophe.
func inputText(v werkbook.Value) string {
	switch v.Type {
	case werkbook.TypeNumber:
		return numberText(v.Number)
	case werkbook.TypeBool:
		return boolText(v.Bool)
	case werkbook.TypeString:
		if p := parseInput(v.String); p.kind != inputString || p.str != v.String {
			return "'" + v.String
		}
		return v.String
	case werkbook.TypeError:
		return v.String
	default:
		return ""
	}
}

// numberText gives the shortest text that parses back to the same number.
func numberText(n float64) string {
	if abs := math.Abs(n); abs != 0 && (abs >= 1e15 || abs < 1e-7) {
		return strconv.FormatFloat(n, 'E', -1, 64)
	}
	return strconv.FormatFloat(n, 'f', -1, 64)
}

func boolText(b bool) string {
	if b {
		return "TRUE"
	}
	return "FALSE"
}
