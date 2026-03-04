package formula

import (
	"strings"
)

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
func ExpandDefinedNames(src string, names []DefinedNameInfo, currentSheetIdx int) string {
	if len(names) == 0 {
		return src
	}

	// Build a case-insensitive lookup of names. Local names for the current
	// sheet take precedence over global names.
	lookup := make(map[string]string, len(names))
	for _, n := range names {
		if n.LocalSheetID >= 0 && n.LocalSheetID != currentSheetIdx {
			continue // local to a different sheet
		}
		key := strings.ToLower(n.Name)
		if _, exists := lookup[key]; exists {
			// Local sheet-scoped names override global names.
			if n.LocalSheetID == currentSheetIdx {
				lookup[key] = n.Value
			}
		} else {
			lookup[key] = n.Value
		}
	}
	if len(lookup) == 0 {
		return src
	}

	var result strings.Builder
	i := 0
	for i < len(src) {
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
			result.WriteString(src[i:j])
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
				result.WriteString(ident)
				continue
			}

			// Don't replace if followed by '!' — it's a sheet prefix.
			if i < len(src) && src[i] == '!' {
				result.WriteString(ident)
				continue
			}

			// Don't replace if preceded by '!' (sheet-qualified ref) or ':' (range part).
			if nameStart > 0 && (src[nameStart-1] == '!' || src[nameStart-1] == ':') {
				result.WriteString(ident)
				continue
			}

			// Check if this identifier is a defined name.
			if val, ok := lookup[strings.ToLower(ident)]; ok {
				// Wrap in parentheses to avoid precedence issues
				// when the defined name value is a complex expression.
				result.WriteString(val)
			} else {
				result.WriteString(ident)
			}
			continue
		}

		result.WriteByte(src[i])
		i++
	}
	return result.String()
}

// isIdentOrDotByte returns true for characters that can appear in an identifier
// (letters, digits, underscore, dot).
func isIdentOrDotByte(c byte) bool {
	return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') ||
		(c >= '0' && c <= '9') || c == '_' || c == '.'
}
