package formula

import (
	"strconv"
	"strings"
)

// excelParseNumber parses text into a float64 using Excel's en-US rules for
// implicit and explicit numeric coercion. It accepts:
//   - Plain numbers and scientific notation.
//   - A leading "$" currency marker, optionally adjacent to a minus sign
//     ("$1,234.56", "-$99", "$-99").
//   - Accounting parens — "(1,234)" becomes -1234.
//   - US-style thousand separators (commas appearing before the decimal point).
//   - A trailing "%" — the value is divided by 100.
//
// It deliberately rejects the things Excel rejects in en-US locale:
//   - EU-style format where the comma is the decimal separator ("1.234,56").
//   - Non-"$" currency prefixes like "€".
//   - Trailing minus signs ("1234-").
//   - Non-ASCII minus signs such as U+2212 or U+FE63.
//   - Strings with internal whitespace ("12 34").
func excelParseNumber(s string) (float64, bool) {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, false
	}

	// Accounting parens negate the parsed value.
	negate := false
	if len(s) >= 2 && s[0] == '(' && s[len(s)-1] == ')' {
		s = strings.TrimSpace(s[1 : len(s)-1])
		negate = true
	}

	// Trailing "%" scales the parsed value by 1/100.
	percent := false
	if strings.HasSuffix(s, "%") {
		s = strings.TrimSpace(strings.TrimSuffix(s, "%"))
		percent = true
	}

	// Accept "$" prefix alone or adjacent to a sign. Anything else (€, ¥, …)
	// falls through and is rejected by ParseFloat below.
	switch {
	case strings.HasPrefix(s, "-$"), strings.HasPrefix(s, "$-"):
		s = "-" + s[2:]
	case strings.HasPrefix(s, "+$"), strings.HasPrefix(s, "$+"):
		s = "+" + s[2:]
	case strings.HasPrefix(s, "$"):
		s = s[1:]
	}

	// Reject EU-style decimals: if both "," and "." appear and the last "," is
	// after the first ".", the comma is being used as a decimal separator
	// rather than a thousand separator.
	if dot := strings.IndexByte(s, '.'); dot >= 0 {
		if com := strings.LastIndexByte(s, ','); com > dot {
			return 0, false
		}
	}

	// Strip thousand separators.
	s = strings.ReplaceAll(s, ",", "")

	n, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return 0, false
	}
	if negate {
		n = -n
	}
	if percent {
		n /= 100
	}
	return n, true
}
