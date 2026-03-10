package formula

import (
	"errors"
	"fmt"
	"strings"
)

// MaxExpandedFormulaBytes caps the size of a formula after structured
// references and defined names have been expanded. The limit is set well above
// Excel's cell formula size ceiling, but low enough to prevent a tiny workbook
// from ballooning into large intermediate allocations during parse-time
// dependency registration.
const MaxExpandedFormulaBytes = 64 * 1024

// ErrFormulaTooLarge reports that a formula or formula expansion exceeded the
// engine's parse-time size budget.
var ErrFormulaTooLarge = errors.New("formula exceeds expansion limit")

// DefinedNameInfo holds a resolved defined name for expansion in formulas.
type DefinedNameInfo struct {
	Name         string // the name (e.g. "OneRange")
	Value        string // the reference or expression (e.g. "Sheet1!$A$10")
	LocalSheetID int    // -1 for global; otherwise 0-based sheet index
}

// ExpandDefinedNames replaces occurrences of defined names in a formula string
// with their referenced values. It skips names inside string literals and avoids
// replacing substrings that are part of function names or cell references.
// currentSheetIdx is the 0-based index of the sheet containing the formula (-1 if unknown).
// sheetNames maps localSheetId values to workbook sheet names for qualified local-name refs.
func ExpandDefinedNames(src string, names []DefinedNameInfo, currentSheetIdx int, sheetNames []string) string {
	expanded, err := ExpandDefinedNamesBounded(src, names, currentSheetIdx, sheetNames, MaxExpandedFormulaBytes)
	if err != nil {
		return src
	}
	return expanded
}

// ExpandDefinedNamesBounded is like ExpandDefinedNames, but returns an error if
// the expanded formula would exceed maxBytes. A non-positive maxBytes disables
// the size check.
func ExpandDefinedNamesBounded(src string, names []DefinedNameInfo, currentSheetIdx int, sheetNames []string, maxBytes int) (string, error) {
	if err := checkExpandedFormulaSize(len(src), maxBytes); err != nil {
		return "", err
	}
	if len(names) == 0 {
		return src, nil
	}

	// Build case-insensitive lookups. Local names for the current sheet take
	// precedence over global names, while qualified local names stay addressable
	// as SheetName!LocalName from any sheet.
	lookup := make(map[string]string, len(names))
	qualified := make(map[string]string, len(names))
	for _, n := range names {
		key := strings.ToLower(n.Name)
		if n.LocalSheetID >= 0 {
			if n.LocalSheetID == currentSheetIdx {
				lookup[key] = n.Value
			}
			if n.LocalSheetID < len(sheetNames) {
				qualified[qualifiedDefinedNameKey(sheetNames[n.LocalSheetID], n.Name)] = n.Value
			}
			continue
		}
		if _, exists := lookup[key]; !exists {
			lookup[key] = n.Value
		}
	}
	if len(lookup) == 0 && len(qualified) == 0 {
		return src, nil
	}

	var result strings.Builder
	written := 0
	i := 0
	for i < len(src) {
		if val, consumed, ok := matchQualifiedDefinedName(src, i, qualified); ok {
			if err := writeExpandedString(&result, &written, val, maxBytes); err != nil {
				return "", err
			}
			i += consumed
			continue
		}

		// Skip string literals.
		if src[i] == '"' {
			j := i + 1
			for j < len(src) {
				if src[j] == '"' {
					j++
					if j < len(src) && src[j] == '"' {
						j++ // escaped quote
						continue
					}
					break
				}
				j++
			}
			if err := writeExpandedString(&result, &written, src[i:j], maxBytes); err != nil {
				return "", err
			}
			i = j
			continue
		}

		// Look for identifiers.
		if isIdentStartByte(src[i]) {
			nameStart := i
			for i < len(src) && isIdentOrDotByte(src[i]) {
				i++
			}
			ident := src[nameStart:i]

			// Don't replace if followed by '(' — it's a function call.
			if i < len(src) && src[i] == '(' {
				if err := writeExpandedString(&result, &written, ident, maxBytes); err != nil {
					return "", err
				}
				continue
			}

			// Don't replace if followed by '!' — it's a sheet prefix.
			if i < len(src) && src[i] == '!' {
				if err := writeExpandedString(&result, &written, ident, maxBytes); err != nil {
					return "", err
				}
				continue
			}

			// Don't replace if preceded by '!' (sheet-qualified ref) or ':' (range part).
			if nameStart > 0 && (src[nameStart-1] == '!' || src[nameStart-1] == ':') {
				if err := writeExpandedString(&result, &written, ident, maxBytes); err != nil {
					return "", err
				}
				continue
			}

			// Check if this identifier is a defined name.
			if val, ok := lookup[strings.ToLower(ident)]; ok {
				// Wrap in parentheses to avoid precedence issues
				// when the defined name value is a complex expression.
				if err := writeExpandedString(&result, &written, val, maxBytes); err != nil {
					return "", err
				}
			} else {
				if err := writeExpandedString(&result, &written, ident, maxBytes); err != nil {
					return "", err
				}
			}
			continue
		}

		if err := writeExpandedByte(&result, &written, src[i], maxBytes); err != nil {
			return "", err
		}
		i++
	}
	return result.String(), nil
}

func qualifiedDefinedNameKey(sheet, name string) string {
	return strings.ToLower(sheet) + "\x00" + strings.ToLower(name)
}

func matchQualifiedDefinedName(src string, start int, lookup map[string]string) (string, int, bool) {
	if len(lookup) == 0 {
		return "", 0, false
	}

	sheetName, next, ok := scanSheetQualifier(src, start)
	if !ok {
		return "", 0, false
	}
	if next >= len(src) || !isIdentStartByte(src[next]) {
		return "", 0, false
	}

	nameStart := next
	for next < len(src) && isIdentOrDotByte(src[next]) {
		next++
	}
	if next < len(src) && src[next] == '(' {
		return "", 0, false
	}

	val, ok := lookup[qualifiedDefinedNameKey(sheetName, src[nameStart:next])]
	if !ok {
		return "", 0, false
	}
	return val, next - start, true
}

func scanSheetQualifier(src string, start int) (string, int, bool) {
	if start >= len(src) {
		return "", 0, false
	}

	if src[start] == '\'' {
		var sheet strings.Builder
		i := start + 1
		for i < len(src) {
			if src[i] != '\'' {
				sheet.WriteByte(src[i])
				i++
				continue
			}
			if i+1 < len(src) && src[i+1] == '\'' {
				sheet.WriteByte('\'')
				i += 2
				continue
			}
			if i+1 >= len(src) || src[i+1] != '!' {
				return "", 0, false
			}
			return sheet.String(), i + 2, true
		}
		return "", 0, false
	}

	if !isIdentStartByte(src[start]) {
		return "", 0, false
	}
	i := start
	for i < len(src) && isIdentOrDotByte(src[i]) {
		i++
	}
	if i >= len(src) || src[i] != '!' {
		return "", 0, false
	}
	return src[start:i], i + 1, true
}

// isIdentOrDotByte returns true for characters that can appear in an identifier
// (letters, digits, underscore, dot).
func isIdentOrDotByte(c byte) bool {
	return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') ||
		(c >= '0' && c <= '9') || c == '_' || c == '.'
}

func writeExpandedString(dst *strings.Builder, written *int, s string, maxBytes int) error {
	if err := checkExpandedFormulaSize(*written+len(s), maxBytes); err != nil {
		return err
	}
	dst.WriteString(s)
	*written += len(s)
	return nil
}

func writeExpandedByte(dst *strings.Builder, written *int, b byte, maxBytes int) error {
	if err := checkExpandedFormulaSize(*written+1, maxBytes); err != nil {
		return err
	}
	dst.WriteByte(b)
	*written += 1
	return nil
}

func checkExpandedFormulaSize(size, maxBytes int) error {
	if maxBytes > 0 && size > maxBytes {
		return fmt.Errorf("%w: expanded formula exceeds %d bytes", ErrFormulaTooLarge, maxBytes)
	}
	return nil
}
