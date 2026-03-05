package formula

import (
	"math"
	"sort"
	"strconv"
	"strings"
)

func init() {
	Register("AVERAGE", NoCtx(fnAVERAGE))
	Register("AVERAGEA", NoCtx(fnAVERAGEA))
	Register("AVEDEV", NoCtx(fnAVEDEV))
	Register("AVERAGEIF", NoCtx(fnAVERAGEIF))
	Register("AVERAGEIFS", NoCtx(fnAVERAGEIFS))
	Register("COUNT", NoCtx(fnCOUNT))
	Register("COUNTA", NoCtx(fnCOUNTA))
	Register("CORREL", NoCtx(fnCORREL))
	Register("COVAR", NoCtx(fnCOVARIANCEP))
	Register("COVARIANCE.P", NoCtx(fnCOVARIANCEP))
	Register("COVARIANCE.S", NoCtx(fnCOVARIANCES))
	Register("INTERCEPT", NoCtx(fnINTERCEPT))
	Register("COUNTBLANK", NoCtx(fnCOUNTBLANK))
	Register("COUNTIF", NoCtx(fnCOUNTIF))
	Register("COUNTIFS", NoCtx(fnCOUNTIFS))
	Register("DEVSQ", NoCtx(fnDEVSQ))
	Register("FISHER", NoCtx(fnFISHER))
	Register("FISHERINV", NoCtx(fnFISHERINV))
	Register("FORECAST", NoCtx(fnFORECAST))
	Register("FORECAST.LINEAR", NoCtx(fnFORECAST))
	Register("GAMMALN", NoCtx(fnGAMMALN))
	Register("GAMMALN.PRECISE", NoCtx(fnGAMMALN))
	Register("GEOMEAN", NoCtx(fnGEOMEAN))
	Register("HARMEAN", NoCtx(fnHARMEAN))
	Register("LARGE", NoCtx(fnLARGE))
	Register("MAX", NoCtx(fnMAX))
	Register("MAXA", NoCtx(fnMAXA))
	Register("MAXIFS", NoCtx(fnMAXIFS))
	Register("MEDIAN", NoCtx(fnMEDIAN))
	Register("MIN", NoCtx(fnMIN))
	Register("MINA", NoCtx(fnMINA))
	Register("MINIFS", NoCtx(fnMINIFS))
	Register("MODE", NoCtx(fnMODE))
	Register("MODE.SNGL", NoCtx(fnMODE))
	Register("PERCENTILE", NoCtx(fnPERCENTILE))
	Register("PERCENTILE.EXC", NoCtx(fnPERCENTILEEXC))
	Register("QUARTILE", NoCtx(fnQUARTILE))
	Register("QUARTILE.EXC", NoCtx(fnQUARTILEEXC))
	Register("PERCENTRANK", NoCtx(fnPERCENTRANK))
	Register("PERCENTRANK.INC", NoCtx(fnPERCENTRANK))
	Register("PERCENTRANK.EXC", NoCtx(fnPERCENTRANKEXC))
	Register("RANK", NoCtx(fnRANK))
	Register("RANK.EQ", NoCtx(fnRANK))
	Register("RANK.AVG", NoCtx(fnRANKAVG))
	Register("SLOPE", NoCtx(fnSLOPE))
	Register("SMALL", NoCtx(fnSMALL))
	Register("STDEV", NoCtx(fnSTDEV))
	Register("STDEV.S", NoCtx(fnSTDEV))
	Register("STDEVP", NoCtx(fnSTDEVP))
	Register("STDEV.P", NoCtx(fnSTDEVP))
	Register("SUM", NoCtx(fnSUM))
	Register("SUMIF", NoCtx(fnSUMIF))
	Register("SUMIFS", NoCtx(fnSUMIFS))
	Register("SUMPRODUCT", NoCtx(fnSUMPRODUCT))
	Register("SUMSQ", NoCtx(fnSUMSQ))
	Register("VAR", NoCtx(fnVAR))
	Register("VAR.S", NoCtx(fnVAR))
	Register("TRIMMEAN", NoCtx(fnTRIMMEAN))
	Register("SKEW", NoCtx(fnSKEW))
	Register("VARP", NoCtx(fnVARP))
	Register("VAR.P", NoCtx(fnVARP))
}

func fnSUM(args []Value) (Value, error) {
	sum := 0.0
	if e := IterateNumeric(args, func(n float64) { sum += n }); e != nil {
		return *e, nil
	}
	return NumberVal(sum), nil
}

func fnAVERAGE(args []Value) (Value, error) {
	sum := 0.0
	count := 0
	if e := IterateNumeric(args, func(n float64) { sum += n; count++ }); e != nil {
		return *e, nil
	}
	if count == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(sum / float64(count)), nil
}

// fnAVERAGEA calculates the average of its arguments, including text and
// logical values.  In arrays/ranges: numbers count as their value, TRUE=1,
// FALSE=0, text strings=0, empty cells are ignored.  For direct (non-array)
// arguments: booleans and numbers are coerced normally; text that cannot be
// parsed as a number returns #VALUE!.
func fnAVERAGEA(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	sum := 0.0
	count := 0
	for _, arg := range args {
		switch arg.Type {
		case ValueArray:
			for _, row := range arg.Array {
				for _, cell := range row {
					switch cell.Type {
					case ValueError:
						return cell, nil
					case ValueNumber:
						sum += cell.Num
						count++
					case ValueBool:
						if cell.Bool {
							sum += 1
						}
						count++
					case ValueString:
						// Text in a range counts as 0.
						sum += 0
						count++
					case ValueEmpty:
						// Empty cells are ignored.
					}
				}
			}
		case ValueError:
			return arg, nil
		case ValueNumber:
			sum += arg.Num
			count++
		case ValueBool:
			if arg.Bool {
				sum += 1
			}
			count++
		case ValueString:
			// Direct text argument: try to coerce to number.
			n, e := CoerceNum(arg)
			if e != nil {
				return *e, nil
			}
			sum += n
			count++
		case ValueEmpty:
			// Empty direct arguments are ignored.
		}
	}
	if count == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(sum / float64(count)), nil
}

func fnCOUNT(args []Value) (Value, error) {
	count := 0
	for _, arg := range args {
		switch arg.Type {
		case ValueArray:
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueNumber {
						count++
					}
				}
			}
		case ValueNumber:
			count++
		case ValueBool:
			// Direct boolean args (e.g. COUNT(TRUE)) are counted,
			// but booleans from cell references (e.g. COUNT(A1) where
			// A1=TRUE) are not — matching Excel behavior.
			if !arg.FromCell {
				count++
			}
		case ValueString:
			if _, err := strconv.ParseFloat(arg.Str, 64); err == nil {
				count++
			}
		}
	}
	return NumberVal(float64(count)), nil
}

func fnCOUNTA(args []Value) (Value, error) {
	count := 0
	for _, arg := range args {
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type != ValueEmpty {
						count++
					}
				}
			}
		} else if arg.Type != ValueEmpty {
			count++
		}
	}
	return NumberVal(float64(count)), nil
}

func fnMIN(args []Value) (Value, error) {
	min := math.MaxFloat64
	found := false
	if e := IterateNumeric(args, func(n float64) {
		if !found || n < min {
			min = n
			found = true
		}
	}); e != nil {
		return *e, nil
	}
	if !found {
		return NumberVal(0), nil
	}
	return NumberVal(min), nil
}

func fnMAX(args []Value) (Value, error) {
	max := -math.MaxFloat64
	found := false
	if e := IterateNumeric(args, func(n float64) {
		if !found || n > max {
			max = n
			found = true
		}
	}); e != nil {
		return *e, nil
	}
	if !found {
		return NumberVal(0), nil
	}
	return NumberVal(max), nil
}

// fnMAXA returns the largest value in args, including text and logical values.
// In arrays/ranges: numbers count as their value, TRUE=1, FALSE=0, text
// strings=0, empty cells are ignored.  For direct (non-array) arguments:
// booleans and numbers are coerced normally; text that cannot be parsed as a
// number returns #VALUE!.  Returns 0 when no values are found.
func fnMAXA(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	max := 0.0
	found := false
	for _, arg := range args {
		switch arg.Type {
		case ValueArray:
			for _, row := range arg.Array {
				for _, cell := range row {
					switch cell.Type {
					case ValueError:
						return cell, nil
					case ValueNumber:
						if !found || cell.Num > max {
							max = cell.Num
							found = true
						}
					case ValueBool:
						v := 0.0
						if cell.Bool {
							v = 1
						}
						if !found || v > max {
							max = v
							found = true
						}
					case ValueString:
						// Text in a range counts as 0.
						if !found || 0 > max {
							max = 0
							found = true
						}
					case ValueEmpty:
						// Empty cells are ignored.
					}
				}
			}
		case ValueError:
			return arg, nil
		case ValueNumber:
			if !found || arg.Num > max {
				max = arg.Num
				found = true
			}
		case ValueBool:
			v := 0.0
			if arg.Bool {
				v = 1
			}
			if !found || v > max {
				max = v
				found = true
			}
		case ValueString:
			// Direct text argument: try to coerce to number.
			n, e := CoerceNum(arg)
			if e != nil {
				return *e, nil
			}
			if !found || n > max {
				max = n
				found = true
			}
		case ValueEmpty:
			// Empty direct arguments are ignored.
		}
	}
	if !found {
		return NumberVal(0), nil
	}
	return NumberVal(max), nil
}

// fnMINA returns the smallest value in args, including text and logical values.
// In arrays/ranges: numbers count as their value, TRUE=1, FALSE=0, text
// strings=0, empty cells are ignored.  For direct (non-array) arguments:
// booleans and numbers are coerced normally; text that cannot be parsed as a
// number returns #VALUE!.  Returns 0 when no values are found.
func fnMINA(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	min := 0.0
	found := false
	for _, arg := range args {
		switch arg.Type {
		case ValueArray:
			for _, row := range arg.Array {
				for _, cell := range row {
					switch cell.Type {
					case ValueError:
						return cell, nil
					case ValueNumber:
						if !found || cell.Num < min {
							min = cell.Num
							found = true
						}
					case ValueBool:
						v := 0.0
						if cell.Bool {
							v = 1
						}
						if !found || v < min {
							min = v
							found = true
						}
					case ValueString:
						// Text in a range counts as 0.
						if !found || 0 < min {
							min = 0
							found = true
						}
					case ValueEmpty:
						// Empty cells are ignored.
					}
				}
			}
		case ValueError:
			return arg, nil
		case ValueNumber:
			if !found || arg.Num < min {
				min = arg.Num
				found = true
			}
		case ValueBool:
			v := 0.0
			if arg.Bool {
				v = 1
			}
			if !found || v < min {
				min = v
				found = true
			}
		case ValueString:
			// Direct text argument: try to coerce to number.
			n, e := CoerceNum(arg)
			if e != nil {
				return *e, nil
			}
			if !found || n < min {
				min = n
				found = true
			}
		case ValueEmpty:
			// Empty direct arguments are ignored.
		}
	}
	if !found {
		return NumberVal(0), nil
	}
	return NumberVal(min), nil
}

func fnLARGE(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	var nums []float64
	arr := args[0]
	if arr.Type == ValueArray {
		for _, row := range arr.Array {
			for _, cell := range row {
				if cell.Type == ValueError {
					return cell, nil
				}
				if cell.Type == ValueNumber {
					nums = append(nums, cell.Num)
				}
			}
		}
	} else {
		n, e := CoerceNum(arr)
		if e != nil {
			return *e, nil
		}
		nums = append(nums, n)
	}
	k, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	ki := int(k)
	if ki < 1 || ki > len(nums) {
		return ErrorVal(ErrValNUM), nil
	}
	sort.Float64s(nums)
	return NumberVal(nums[len(nums)-ki]), nil
}

func fnSMALL(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	var nums []float64
	arr := args[0]
	if arr.Type == ValueArray {
		for _, row := range arr.Array {
			for _, cell := range row {
				if cell.Type == ValueError {
					return cell, nil
				}
				if cell.Type == ValueNumber {
					nums = append(nums, cell.Num)
				}
			}
		}
	} else {
		n, e := CoerceNum(arr)
		if e != nil {
			return *e, nil
		}
		nums = append(nums, n)
	}
	k, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	ki := int(k)
	if ki < 1 || ki > len(nums) {
		return ErrorVal(ErrValNUM), nil
	}
	sort.Float64s(nums)
	return NumberVal(nums[ki-1]), nil
}

func fnCOUNTBLANK(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	count := 0
	arg := args[0]
	switch arg.Type {
	case ValueArray:
		for _, row := range arg.Array {
			for _, cell := range row {
				if cell.Type == ValueEmpty || (cell.Type == ValueString && cell.Str == "") {
					count++
				}
			}
		}
	case ValueEmpty:
		count = 1
	case ValueString:
		if arg.Str == "" {
			count = 1
		}
	}
	return NumberVal(float64(count)), nil
}

// MatchesCriteria checks if a value matches Excel-style criteria.
//
// In Excel, booleans and numbers are distinct types for criteria matching:
//   - Numeric criteria (e.g. 1) only match numeric cells, not booleans.
//   - Boolean criteria (e.g. TRUE) only match boolean cells, not numbers.
//   - String criteria with comparison operators (e.g. ">0") only compare
//     against numeric cells; booleans are excluded from numeric comparisons.
//   - String criteria "TRUE"/"FALSE" are coerced to boolean for matching.
func MatchesCriteria(v Value, criteria Value) bool {
	critStr := ValueToString(criteria)

	if len(critStr) >= 1 {
		// Extract operator and operand.
		var op, operand string
		switch {
		case strings.HasPrefix(critStr, ">="):
			op, operand = ">=", critStr[2:]
		case strings.HasPrefix(critStr, "<="):
			op, operand = "<=", critStr[2:]
		case strings.HasPrefix(critStr, "<>"):
			op, operand = "<>", critStr[2:]
		case strings.HasPrefix(critStr, ">"):
			op, operand = ">", critStr[1:]
		case strings.HasPrefix(critStr, "<"):
			op, operand = "<", critStr[1:]
		case strings.HasPrefix(critStr, "="):
			op, operand = "=", critStr[1:]
		}
		if op != "" {
			return matchOperator(v, op, operand)
		}
	}

	switch classifyWildcard(critStr) {
	case wildcardFull:
		return WildcardMatch(ValueToString(v), critStr)
	case wildcardEscape:
		// Pattern has escape sequences but no unescaped wildcards (e.g. "~**").
		// Do a case-insensitive literal comparison after unescaping.
		return strings.EqualFold(ValueToString(v), unescapePattern(critStr))
	}

	// Boolean criteria: only match boolean cells.
	if criteria.Type == ValueBool {
		return v.Type == ValueBool && v.Bool == criteria.Bool
	}

	// Numeric criteria: match numeric cells, and also string cells whose
	// content parses as the same number (Excel coerces strings to numbers
	// for COUNTIF/SUMIF criteria matching). Booleans are excluded.
	if criteria.Type == ValueNumber {
		if v.Type == ValueNumber {
			return v.Num == criteria.Num
		}
		if v.Type == ValueString {
			if n, err := strconv.ParseFloat(v.Str, 64); err == nil {
				return n == criteria.Num
			}
		}
		return false
	}

	// String criteria "TRUE"/"FALSE": coerce to boolean for matching.
	// In Excel, =COUNTIF(range,"TRUE") counts only boolean TRUE cells,
	// not cells containing the string "TRUE".
	upper := strings.ToUpper(strings.TrimSpace(critStr))
	if upper == "TRUE" {
		return v.Type == ValueBool && v.Bool
	}
	if upper == "FALSE" {
		return v.Type == ValueBool && !v.Bool
	}

	// Other string criteria: use case-insensitive comparison.
	return strings.EqualFold(ValueToString(v), critStr)
}

// matchOperator applies a comparison operator to a cell value and an operand
// string, respecting Excel's type separation between booleans and numbers.
func matchOperator(v Value, op, operand string) bool {
	// "=" with empty operand matches only truly empty cells (TypeEmpty).
	// This differs from bare "" criteria which matches both empty and empty-string cells.
	if op == "=" && operand == "" {
		return v.Type == ValueEmpty
	}
	// "<>" with empty operand: everything except truly empty cells.
	// In Excel, empty-string cells are considered non-blank for <>.
	if op == "<>" && operand == "" {
		return v.Type != ValueEmpty
	}
	upper := strings.ToUpper(strings.TrimSpace(operand))

	// If the operand is a boolean literal, compare against boolean type only.
	if upper == "TRUE" || upper == "FALSE" {
		critBool := upper == "TRUE"
		if v.Type == ValueBool {
			cmp := 0
			if v.Bool != critBool {
				if !v.Bool {
					cmp = -1 // FALSE < TRUE
				} else {
					cmp = 1
				}
			}
			return evalOp(op, cmp)
		}
		// Non-boolean value compared to boolean operand:
		// For <> they are not equal; for = they are not equal;
		// for ordering operators they don't match.
		if op == "<>" {
			return true
		}
		return false
	}

	// If the operand is numeric, compare against numeric cells.
	// For the "=" operator only, also match string cells whose content
	// parses as the same number (Excel coerces text-numbers for equality
	// in *IF criteria, but not for ordering operators like >, <, etc.).
	// Booleans are excluded from numeric comparisons.
	if cn, err := strconv.ParseFloat(operand, 64); err == nil {
		if v.Type == ValueNumber {
			return evalOp(op, cmpFloat(v.Num, cn))
		}
		if op == "=" && v.Type == ValueString {
			if n, err2 := strconv.ParseFloat(v.Str, 64); err2 == nil {
				return cmpFloat(n, cn) == 0
			}
		}
		// Boolean or other non-numeric type vs numeric operand:
		// For <> they are not equal; otherwise no match.
		if op == "<>" {
			return true
		}
		return false
	}

	// Non-numeric, non-boolean operand: plain string comparison.
	cmp := strings.Compare(strings.ToLower(ValueToString(v)), strings.ToLower(operand))
	return evalOp(op, cmp)
}

// evalOp applies a comparison operator to a cmp result (-1, 0, 1).
func evalOp(op string, cmp int) bool {
	switch op {
	case ">=":
		return cmp >= 0
	case "<=":
		return cmp <= 0
	case "<>":
		return cmp != 0
	case ">":
		return cmp > 0
	case "<":
		return cmp < 0
	case "=":
		return cmp == 0
	}
	return false
}

// CompareToCriteria compares a value to a criteria string. This is the
// original type-agnostic version kept for backward compatibility with
// external callers. Internal *IF functions use compareToCriteriaTyped.
func CompareToCriteria(v Value, critValStr string) int {
	if n, e := CoerceNum(v); e == nil {
		if cn, err := strconv.ParseFloat(critValStr, 64); err == nil {
			return cmpFloat(n, cn)
		}
	}
	return strings.Compare(strings.ToLower(ValueToString(v)), strings.ToLower(critValStr))
}

// wildcardMode describes what wildcard processing a criteria string needs.
type wildcardMode int

const (
	wildcardNone   wildcardMode = iota // no wildcards, no escapes
	wildcardEscape                     // escape sequences only (e.g. "~**"), no unescaped wildcards
	wildcardFull                       // has at least one unescaped wildcard (* or ?)
)

// classifyWildcard examines a criteria string and returns what kind of
// wildcard processing it needs.
//
// In Excel, ~* escapes a literal * and also absorbs any immediately
// following identical wildcard characters.  So "~**" has no unescaped
// wildcards (it matches literal "**"), whereas "*~**" still has unescaped
// wildcards at the leading and trailing positions.
func classifyWildcard(s string) wildcardMode {
	hasEscape := false
	for i := 0; i < len(s); i++ {
		switch s[i] {
		case '~':
			if i+1 < len(s) && (s[i+1] == '*' || s[i+1] == '?' || s[i+1] == '~') {
				hasEscape = true
				escaped := s[i+1]
				i++ // skip the escaped char
				// Also skip any immediately following identical wildcard chars.
				if escaped == '*' || escaped == '?' {
					for i+1 < len(s) && s[i+1] == escaped {
						i++
					}
				}
			}
		case '*', '?':
			return wildcardFull
		}
	}
	if hasEscape {
		return wildcardEscape
	}
	return wildcardNone
}

// unescapePattern removes tilde escape sequences from a pattern,
// returning the literal string it represents.  This is used when
// the pattern has escape sequences but no unescaped wildcards.
func unescapePattern(s string) string {
	var b strings.Builder
	for i := 0; i < len(s); i++ {
		if s[i] == '~' && i+1 < len(s) && (s[i+1] == '*' || s[i+1] == '?' || s[i+1] == '~') {
			b.WriteByte(s[i+1])
			i++ // skip escaped char
			// Also emit any immediately following identical wildcard chars.
			if s[i] == '*' || s[i] == '?' {
				for i+1 < len(s) && s[i+1] == s[i] {
					i++
					b.WriteByte(s[i])
				}
			}
		} else {
			b.WriteByte(s[i])
		}
	}
	return b.String()
}

func WildcardMatch(s, pattern string) bool {
	return WildcardHelper(strings.ToLower(s), strings.ToLower(pattern))
}

func WildcardHelper(s, p string) bool {
	for len(p) > 0 {
		switch p[0] {
		case '~':
			// Escape sequence: ~* means literal *, ~? means literal ?, ~~ means literal ~
			if len(p) >= 2 && (p[1] == '*' || p[1] == '?' || p[1] == '~') {
				if len(s) == 0 || s[0] != p[1] {
					return false
				}
				s = s[1:]
				p = p[2:]
			} else {
				// Lone ~ at end or before non-special char: treat as literal ~
				if len(s) == 0 || s[0] != '~' {
					return false
				}
				s = s[1:]
				p = p[1:]
			}
		case '*':
			for len(p) > 0 && p[0] == '*' {
				p = p[1:]
			}
			if len(p) == 0 {
				return true
			}
			for i := 0; i <= len(s); i++ {
				if WildcardHelper(s[i:], p) {
					return true
				}
			}
			return false
		case '?':
			if len(s) == 0 {
				return false
			}
			s = s[1:]
			p = p[1:]
		default:
			if len(s) == 0 || s[0] != p[0] {
				return false
			}
			s = s[1:]
			p = p[1:]
		}
	}
	return len(s) == 0
}

func fnSUMIF(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	rangeArg := args[0]
	if rangeArg.Type == ValueError {
		return rangeArg, nil
	}
	criteria := args[1]
	sumRange := rangeArg
	if len(args) == 3 {
		sumRange = args[2]
	}

	if rangeArg.Type != ValueArray {
		if MatchesCriteria(rangeArg, criteria) {
			n, e := CoerceNum(sumRange)
			if e != nil {
				return NumberVal(0), nil
			}
			return NumberVal(n), nil
		}
		return NumberVal(0), nil
	}

	sum := 0.0
	for i, row := range rangeArg.Array {
		for j, cell := range row {
			if MatchesCriteria(cell, criteria) {
				var sv Value
				if sumRange.Type == ValueArray && i < len(sumRange.Array) && j < len(sumRange.Array[i]) {
					sv = sumRange.Array[i][j]
				} else if sumRange.Type != ValueArray {
					sv = sumRange
				}
				if n, e := CoerceNum(sv); e == nil {
					sum += n
				}
			}
		}
	}
	return NumberVal(sum), nil
}

func fnSUMIFS(args []Value) (Value, error) {
	if len(args) < 3 || (len(args)-1)%2 != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	sumRange := args[0]
	if sumRange.Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}

	sum := 0.0
	for r, row := range sumRange.Array {
		for c := range row {
			allMatch := true
			for k := 1; k < len(args); k += 2 {
				critRange := args[k]
				criteria := args[k+1]
				var cellVal Value
				if critRange.Type == ValueArray && r < len(critRange.Array) && c < len(critRange.Array[r]) {
					cellVal = critRange.Array[r][c]
				}
				if !MatchesCriteria(cellVal, criteria) {
					allMatch = false
					break
				}
			}
			if allMatch {
				if n, e := CoerceNum(sumRange.Array[r][c]); e == nil {
					sum += n
				}
			}
		}
	}
	return NumberVal(sum), nil
}

func fnCOUNTIF(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	rangeArg := args[0]
	if rangeArg.Type == ValueError {
		return rangeArg, nil
	}
	criteria := args[1]
	count := 0
	if rangeArg.Type == ValueArray {
		for _, row := range rangeArg.Array {
			for _, cell := range row {
				if MatchesCriteria(cell, criteria) {
					count++
				}
			}
		}
	} else if MatchesCriteria(rangeArg, criteria) {
		count = 1
	}
	return NumberVal(float64(count)), nil
}

func fnCOUNTIFS(args []Value) (Value, error) {
	if len(args) < 2 || len(args)%2 != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	firstRange := args[0]
	if firstRange.Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}

	count := 0
	for r, row := range firstRange.Array {
		for c := range row {
			allMatch := true
			for k := 0; k < len(args); k += 2 {
				critRange := args[k]
				criteria := args[k+1]
				var cellVal Value
				if critRange.Type == ValueArray && r < len(critRange.Array) && c < len(critRange.Array[r]) {
					cellVal = critRange.Array[r][c]
				}
				if !MatchesCriteria(cellVal, criteria) {
					allMatch = false
					break
				}
			}
			if allMatch {
				count++
			}
		}
	}
	return NumberVal(float64(count)), nil
}

func fnAVERAGEIF(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	rangeArg := args[0]
	criteria := args[1]
	avgRange := rangeArg
	if len(args) == 3 {
		avgRange = args[2]
	}

	sum := 0.0
	count := 0
	if rangeArg.Type == ValueArray {
		for i, row := range rangeArg.Array {
			for j, cell := range row {
				if MatchesCriteria(cell, criteria) {
					var sv Value
					if avgRange.Type == ValueArray && i < len(avgRange.Array) && j < len(avgRange.Array[i]) {
						sv = avgRange.Array[i][j]
					}
					if n, e := CoerceNum(sv); e == nil {
						sum += n
						count++
					}
				}
			}
		}
	}
	if count == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(sum / float64(count)), nil
}

func fnSUMPRODUCT(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}
	firstArr := args[0].Array
	rows := len(firstArr)
	cols := 0
	if rows > 0 {
		cols = len(firstArr[0])
	}

	for _, arg := range args[1:] {
		if arg.Type != ValueArray {
			return ErrorVal(ErrValVALUE), nil
		}
		if len(arg.Array) != rows {
			return ErrorVal(ErrValVALUE), nil
		}
		for _, row := range arg.Array {
			if len(row) != cols {
				return ErrorVal(ErrValVALUE), nil
			}
		}
	}

	sum := 0.0
	for r := 0; r < rows; r++ {
		for c := 0; c < cols; c++ {
			product := 1.0
			for _, arg := range args {
				cell := arg.Array[r][c]
				if cell.Type == ValueError {
					return cell, nil
				}
				n, e := CoerceNum(cell)
				if e != nil {
					n = 0
				}
				product *= n
			}
			sum += product
		}
	}
	return NumberVal(sum), nil
}

// collectNumeric gathers all numeric values from args into a slice.
func collectNumeric(args []Value) ([]float64, *Value) {
	// Pre-count capacity to avoid repeated reallocation.
	cap := 0
	for _, arg := range args {
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				cap += len(row)
			}
		} else {
			cap++
		}
	}
	nums := make([]float64, 0, cap)
	for _, arg := range args {
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return nil, &cell
					}
					if cell.Type == ValueNumber {
						nums = append(nums, cell.Num)
					}
				}
			}
		} else {
			if arg.Type == ValueError {
				return nil, &arg
			}
			n, e := CoerceNum(arg)
			if e != nil {
				return nil, e
			}
			nums = append(nums, n)
		}
	}
	return nums, nil
}

// meanAndSumSq computes the mean and sum of squared deviations from the mean
// for a set of numeric values collected from formula arguments.
func meanAndSumSq(args []Value) (nums []float64, mean, ssq float64, ev *Value) {
	nums, ev = collectNumeric(args)
	if ev != nil {
		return
	}
	n := len(nums)
	if n == 0 {
		return
	}
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean = sum / float64(n)
	for _, v := range nums {
		d := v - mean
		ssq += d * d
	}
	return
}

func fnAVEDEV(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean := sum / float64(n)
	absDevSum := 0.0
	for _, v := range nums {
		d := v - mean
		if d < 0 {
			d = -d
		}
		absDevSum += d
	}
	return NumberVal(absDevSum / float64(n)), nil
}

func fnDEVSQ(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	_, _, ssq, e := meanAndSumSq(args)
	if e != nil {
		return *e, nil
	}
	return NumberVal(ssq), nil
}

func fnAVERAGEIFS(args []Value) (Value, error) {
	if len(args) < 3 || (len(args)-1)%2 != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	avgRange := args[0]
	if avgRange.Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}

	sum := 0.0
	count := 0
	for r, row := range avgRange.Array {
		for c := range row {
			allMatch := true
			for k := 1; k < len(args); k += 2 {
				critRange := args[k]
				criteria := args[k+1]
				var cellVal Value
				if critRange.Type == ValueArray && r < len(critRange.Array) && c < len(critRange.Array[r]) {
					cellVal = critRange.Array[r][c]
				}
				if !MatchesCriteria(cellVal, criteria) {
					allMatch = false
					break
				}
			}
			if allMatch {
				if n, e := CoerceNum(avgRange.Array[r][c]); e == nil {
					sum += n
					count++
				}
			}
		}
	}
	if count == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(sum / float64(count)), nil
}

func fnMAXIFS(args []Value) (Value, error) {
	if len(args) < 3 || (len(args)-1)%2 != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	maxRange := args[0]
	if maxRange.Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}

	maxVal := -math.MaxFloat64
	found := false
	for r, row := range maxRange.Array {
		for c := range row {
			allMatch := true
			for k := 1; k < len(args); k += 2 {
				critRange := args[k]
				criteria := args[k+1]
				var cellVal Value
				if critRange.Type == ValueArray && r < len(critRange.Array) && c < len(critRange.Array[r]) {
					cellVal = critRange.Array[r][c]
				}
				if !MatchesCriteria(cellVal, criteria) {
					allMatch = false
					break
				}
			}
			if allMatch {
				if n, e := CoerceNum(maxRange.Array[r][c]); e == nil {
					if !found || n > maxVal {
						maxVal = n
						found = true
					}
				}
			}
		}
	}
	if !found {
		return NumberVal(0), nil
	}
	return NumberVal(maxVal), nil
}

func fnMINIFS(args []Value) (Value, error) {
	if len(args) < 3 || (len(args)-1)%2 != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	minRange := args[0]
	if minRange.Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}

	minVal := math.MaxFloat64
	found := false
	for r, row := range minRange.Array {
		for c := range row {
			allMatch := true
			for k := 1; k < len(args); k += 2 {
				critRange := args[k]
				criteria := args[k+1]
				var cellVal Value
				if critRange.Type == ValueArray && r < len(critRange.Array) && c < len(critRange.Array[r]) {
					cellVal = critRange.Array[r][c]
				}
				if !MatchesCriteria(cellVal, criteria) {
					allMatch = false
					break
				}
			}
			if allMatch {
				if n, e := CoerceNum(minRange.Array[r][c]); e == nil {
					if !found || n < minVal {
						minVal = n
						found = true
					}
				}
			}
		}
	}
	if !found {
		return NumberVal(0), nil
	}
	return NumberVal(minVal), nil
}

func fnMEDIAN(args []Value) (Value, error) {
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	if len(nums) == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	sort.Float64s(nums)
	n := len(nums)
	if n%2 == 1 {
		return NumberVal(nums[n/2]), nil
	}
	return NumberVal((nums[n/2-1] + nums[n/2]) / 2), nil
}

func fnMODE(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	if len(nums) == 0 {
		return ErrorVal(ErrValNA), nil
	}
	freq := make(map[float64]int)
	order := make([]float64, 0)
	for _, n := range nums {
		if freq[n] == 0 {
			order = append(order, n)
		}
		freq[n]++
	}
	bestVal := 0.0
	bestCount := 1
	for _, v := range order {
		if freq[v] > bestCount {
			bestCount = freq[v]
			bestVal = v
		}
	}
	if bestCount < 2 {
		return ErrorVal(ErrValNA), nil
	}
	return NumberVal(bestVal), nil
}

func fnPERCENTILE(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args[:1])
	if e != nil {
		return *e, nil
	}
	if len(nums) == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	k, e2 := CoerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}
	if k < 0 || k > 1 {
		return ErrorVal(ErrValNUM), nil
	}
	sort.Float64s(nums)
	n := len(nums)
	if n == 1 {
		return NumberVal(nums[0]), nil
	}
	rank := k * float64(n-1)
	intPart := int(rank)
	frac := rank - float64(intPart)
	if intPart >= n-1 {
		return NumberVal(nums[n-1]), nil
	}
	result := nums[intPart] + frac*(nums[intPart+1]-nums[intPart])
	return NumberVal(result), nil
}

func fnQUARTILE(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	q, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	q = math.Trunc(q)
	if q < 0 || q > 4 {
		return ErrorVal(ErrValNUM), nil
	}
	return fnPERCENTILE([]Value{args[0], NumberVal(q * 0.25)})
}

// fnPERCENTILEEXC implements PERCENTILE.EXC which returns the k-th percentile
// using exclusive interpolation. k must be strictly between 0 and 1 (exclusive).
// The rank is computed as k*(n+1), and if the rank falls outside [1, n] it
// returns #NUM!.
func fnPERCENTILEEXC(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args[:1])
	if e != nil {
		return *e, nil
	}
	if len(nums) == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	k, e2 := CoerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}
	if k <= 0 || k >= 1 {
		return ErrorVal(ErrValNUM), nil
	}
	sort.Float64s(nums)
	n := len(nums)
	rank := k * float64(n+1)
	if rank < 1 || rank > float64(n) {
		return ErrorVal(ErrValNUM), nil
	}
	intPart := int(rank)
	frac := rank - float64(intPart)
	if intPart >= n {
		return NumberVal(nums[n-1]), nil
	}
	result := nums[intPart-1] + frac*(nums[intPart]-nums[intPart-1])
	return NumberVal(result), nil
}

// fnQUARTILEEXC implements QUARTILE.EXC which returns the exclusive quartile.
// quart must be 1, 2, or 3 (0 and 4 return #NUM!, unlike QUARTILE.INC).
func fnQUARTILEEXC(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	q, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	q = math.Trunc(q)
	if q <= 0 || q >= 4 {
		return ErrorVal(ErrValNUM), nil
	}
	return fnPERCENTILEEXC([]Value{args[0], NumberVal(q * 0.25)})
}

func fnRANK(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	num, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	nums, e2 := collectNumeric(args[1:2])
	if e2 != nil {
		return *e2, nil
	}
	ascending := false
	if len(args) == 3 {
		order, e3 := CoerceNum(args[2])
		if e3 != nil {
			return *e3, nil
		}
		ascending = order != 0
	}
	found := false
	for _, v := range nums {
		if v == num {
			found = true
			break
		}
	}
	if !found {
		return ErrorVal(ErrValNA), nil
	}
	rank := 1
	for _, v := range nums {
		if ascending {
			if v < num {
				rank++
			}
		} else {
			if v > num {
				rank++
			}
		}
	}
	return NumberVal(float64(rank)), nil
}

// fnRANKAVG implements RANK.AVG. It behaves like RANK except that tied values
// receive the average of the ranks they would span.
func fnRANKAVG(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	num, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	nums, e2 := collectNumeric(args[1:2])
	if e2 != nil {
		return *e2, nil
	}
	ascending := false
	if len(args) == 3 {
		order, e3 := CoerceNum(args[2])
		if e3 != nil {
			return *e3, nil
		}
		ascending = order != 0
	}
	// Count how many values match (ties) and how many are strictly better.
	ties := 0
	better := 0
	for _, v := range nums {
		if v == num {
			ties++
		} else if ascending {
			if v < num {
				better++
			}
		} else {
			if v > num {
				better++
			}
		}
	}
	if ties == 0 {
		return ErrorVal(ErrValNA), nil
	}
	// The tied values span ranks (better+1) through (better+ties).
	// Average = better + 1 + (ties-1)/2 = better + (ties+1)/2.
	avg := float64(better) + (float64(ties)+1)/2
	return NumberVal(avg), nil
}

func fnPERCENTRANK(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Collect numeric values from the array argument.
	nums, e := collectNumeric(args[:1])
	if e != nil {
		return *e, nil
	}
	if len(nums) == 0 {
		return ErrorVal(ErrValNUM), nil
	}

	// x must be numeric.
	x, e2 := CoerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}

	// Optional significance (default 3).
	sig := 3.0
	if len(args) == 3 {
		s, e3 := CoerceNum(args[2])
		if e3 != nil {
			return *e3, nil
		}
		sig = math.Trunc(s)
		if sig < 1 {
			return ErrorVal(ErrValNUM), nil
		}
	}

	sort.Float64s(nums)
	n := len(nums)

	// x outside data range → #N/A
	if x < nums[0] || x > nums[n-1] {
		return ErrorVal(ErrValNA), nil
	}

	// Single element: if x matches, return 1.
	if n == 1 {
		return NumberVal(truncToSig(1, int(sig))), nil
	}

	// Find position of x in sorted data.
	var rank float64
	if x <= nums[0] {
		rank = 0
	} else if x >= nums[n-1] {
		rank = 1
	} else {
		// Find the two adjacent values x falls between (or equals).
		lo := 0
		for i := 0; i < n; i++ {
			if nums[i] == x {
				rank = float64(i) / float64(n-1)
				return NumberVal(truncToSig(rank, int(sig))), nil
			}
			if nums[i] < x {
				lo = i
			}
		}
		// Interpolate between nums[lo] and nums[lo+1].
		loRank := float64(lo) / float64(n-1)
		hiRank := float64(lo+1) / float64(n-1)
		frac := (x - nums[lo]) / (nums[lo+1] - nums[lo])
		rank = loRank + frac*(hiRank-loRank)
	}

	return NumberVal(truncToSig(rank, int(sig))), nil
}

func fnPERCENTRANKEXC(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Collect numeric values from the array argument.
	nums, e := collectNumeric(args[:1])
	if e != nil {
		return *e, nil
	}
	if len(nums) == 0 {
		return ErrorVal(ErrValNUM), nil
	}

	// x must be numeric.
	x, e2 := CoerceNum(args[1])
	if e2 != nil {
		return *e2, nil
	}

	// Optional significance (default 3).
	sig := 3.0
	if len(args) == 3 {
		s, e3 := CoerceNum(args[2])
		if e3 != nil {
			return *e3, nil
		}
		sig = math.Trunc(s)
		if sig < 1 {
			return ErrorVal(ErrValNUM), nil
		}
	}

	sort.Float64s(nums)
	n := len(nums)

	// x outside data range → #N/A
	if x < nums[0] || x > nums[n-1] {
		return ErrorVal(ErrValNA), nil
	}

	// EXC uses rank/(n+1) where rank is 1-based position.
	// With n values, results range from 1/(n+1) to n/(n+1), exclusive of 0 and 1.
	denom := float64(n + 1)

	// Find position of x in sorted data.
	for i := 0; i < n; i++ {
		if nums[i] == x {
			rank := float64(i+1) / denom
			return NumberVal(truncToSig(rank, int(sig))), nil
		}
		if nums[i] > x {
			// Interpolate between nums[i-1] and nums[i].
			lo := i - 1
			loRank := float64(lo+1) / denom
			hiRank := float64(i+1) / denom
			frac := (x - nums[lo]) / (nums[i] - nums[lo])
			rank := loRank + frac*(hiRank-loRank)
			return NumberVal(truncToSig(rank, int(sig))), nil
		}
	}

	// x equals the last element (already handled in loop, but just in case).
	rank := float64(n) / denom
	return NumberVal(truncToSig(rank, int(sig))), nil
}

// truncToSig truncates a float to sig decimal digits.
func truncToSig(v float64, sig int) float64 {
	// Round to 15 significant digits first (matching Excel precision)
	// to eliminate FP noise before truncating.
	v = roundTo15SigFigs(v)
	pow := math.Pow(10, float64(sig))
	return math.Floor(v*pow) / pow
}

func fnSUMSQ(args []Value) (Value, error) {
	sum := 0.0
	if e := IterateNumeric(args, func(n float64) { sum += n * n }); e != nil {
		return *e, nil
	}
	return NumberVal(sum), nil
}

func fnSTDEV(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, _, ssq, e := meanAndSumSq(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n < 2 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(math.Sqrt(ssq / float64(n-1))), nil
}

func fnSTDEVP(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, _, ssq, e := meanAndSumSq(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n < 1 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(math.Sqrt(ssq / float64(n))), nil
}

func fnVAR(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, _, ssq, e := meanAndSumSq(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n < 2 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(ssq / float64(n-1)), nil
}

func fnVARP(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, _, ssq, e := meanAndSumSq(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n < 1 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(ssq / float64(n)), nil
}

func fnTRIMMEAN(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	percent, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	if percent < 0 || percent >= 1 {
		return ErrorVal(ErrValNUM), nil
	}
	nums, ev := collectNumeric(args[:1])
	if ev != nil {
		return *ev, nil
	}
	n := len(nums)
	if n == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	sort.Float64s(nums)
	trim := int(math.Floor(float64(n) * percent / 2))
	remaining := nums[trim : n-trim]
	if len(remaining) == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	sum := 0.0
	for _, v := range remaining {
		sum += v
	}
	return NumberVal(sum / float64(len(remaining))), nil
}

func fnCORREL(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Flatten both arrays to 1D slices of Value.
	flat1 := flattenValuesGeneric(args[0])
	flat2 := flattenValuesGeneric(args[1])

	// Arrays must have the same number of positions.
	if len(flat1) != len(flat2) {
		return ErrorVal(ErrValNA), nil
	}

	// Walk paired positions; keep only pairs where BOTH values are numeric.
	var xs, ys []float64
	for i := range flat1 {
		v1, v2 := flat1[i], flat2[i]
		if v1.Type == ValueError {
			return v1, nil
		}
		if v2.Type == ValueError {
			return v2, nil
		}
		if v1.Type != ValueNumber || v2.Type != ValueNumber {
			continue
		}
		xs = append(xs, v1.Num)
		ys = append(ys, v2.Num)
	}

	n := len(xs)
	if n == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	// Compute means.
	sumX, sumY := 0.0, 0.0
	for i := 0; i < n; i++ {
		sumX += xs[i]
		sumY += ys[i]
	}
	meanX := sumX / float64(n)
	meanY := sumY / float64(n)

	// Compute covariance numerator and both sum-of-squared-deviations.
	cov := 0.0
	ssqX := 0.0
	ssqY := 0.0
	for i := 0; i < n; i++ {
		dx := xs[i] - meanX
		dy := ys[i] - meanY
		cov += dx * dy
		ssqX += dx * dx
		ssqY += dy * dy
	}

	denom := math.Sqrt(ssqX * ssqY)
	if denom == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	return NumberVal(cov / denom), nil
}

// flattenValuesGeneric flattens a Value (possibly an array) into a 1D slice of Value.
func flattenValuesGeneric(arg Value) []Value {
	if arg.Type == ValueArray {
		total := 0
		for _, row := range arg.Array {
			total += len(row)
		}
		out := make([]Value, 0, total)
		for _, row := range arg.Array {
			out = append(out, row...)
		}
		return out
	}
	return []Value{arg}
}

// fnCOVARIANCEP implements COVARIANCE.P (and COVAR, which is identical).
func fnCOVARIANCEP(args []Value) (Value, error) {
	return covarianceImpl(args, false)
}

// fnCOVARIANCES implements COVARIANCE.S (sample covariance).
func fnCOVARIANCES(args []Value) (Value, error) {
	return covarianceImpl(args, true)
}

// covarianceImpl computes population or sample covariance for two arrays.
func covarianceImpl(args []Value, sample bool) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	flat1 := flattenValuesGeneric(args[0])
	flat2 := flattenValuesGeneric(args[1])

	if len(flat1) != len(flat2) {
		return ErrorVal(ErrValNA), nil
	}

	var xs, ys []float64
	for i := range flat1 {
		v1, v2 := flat1[i], flat2[i]
		if v1.Type == ValueError {
			return v1, nil
		}
		if v2.Type == ValueError {
			return v2, nil
		}
		if v1.Type != ValueNumber || v2.Type != ValueNumber {
			continue
		}
		xs = append(xs, v1.Num)
		ys = append(ys, v2.Num)
	}

	n := len(xs)
	if n == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	if sample && n < 2 {
		return ErrorVal(ErrValDIV0), nil
	}

	sumX, sumY := 0.0, 0.0
	for i := 0; i < n; i++ {
		sumX += xs[i]
		sumY += ys[i]
	}
	meanX := sumX / float64(n)
	meanY := sumY / float64(n)

	cov := 0.0
	for i := 0; i < n; i++ {
		cov += (xs[i] - meanX) * (ys[i] - meanY)
	}

	if sample {
		return NumberVal(cov / float64(n-1)), nil
	}
	return NumberVal(cov / float64(n)), nil
}

// linearRegression computes the slope and intercept of the least-squares
// regression line for paired y/x arrays.  It returns (slope, intercept, ok).
// On error it returns (0, 0, false) and an error Value.
func linearRegression(args []Value) (slope, intercept float64, errVal Value, ok bool) {
	if len(args) != 2 {
		return 0, 0, ErrorVal(ErrValVALUE), false
	}

	flatY := flattenValuesGeneric(args[0])
	flatX := flattenValuesGeneric(args[1])

	if len(flatY) != len(flatX) {
		return 0, 0, ErrorVal(ErrValNA), false
	}

	// Walk paired positions; keep only pairs where BOTH values are numeric.
	var xs, ys []float64
	for i := range flatY {
		vy, vx := flatY[i], flatX[i]
		if vy.Type == ValueError {
			return 0, 0, vy, false
		}
		if vx.Type == ValueError {
			return 0, 0, vx, false
		}
		if vy.Type != ValueNumber || vx.Type != ValueNumber {
			continue
		}
		xs = append(xs, vx.Num)
		ys = append(ys, vy.Num)
	}

	n := len(xs)
	if n == 0 {
		return 0, 0, ErrorVal(ErrValDIV0), false
	}

	// Compute means.
	sumX, sumY := 0.0, 0.0
	for i := 0; i < n; i++ {
		sumX += xs[i]
		sumY += ys[i]
	}
	meanX := sumX / float64(n)
	meanY := sumY / float64(n)

	// Compute covariance numerator and sum-of-squared-deviations for x.
	cov := 0.0
	ssqX := 0.0
	for i := 0; i < n; i++ {
		dx := xs[i] - meanX
		dy := ys[i] - meanY
		cov += dx * dy
		ssqX += dx * dx
	}

	if ssqX == 0 {
		return 0, 0, ErrorVal(ErrValDIV0), false
	}

	slope = cov / ssqX
	intercept = meanY - slope*meanX
	return slope, intercept, Value{}, true
}

func fnSLOPE(args []Value) (Value, error) {
	slope, _, errVal, ok := linearRegression(args)
	if !ok {
		return errVal, nil
	}
	return NumberVal(slope), nil
}

func fnINTERCEPT(args []Value) (Value, error) {
	_, intercept, errVal, ok := linearRegression(args)
	if !ok {
		return errVal, nil
	}
	return NumberVal(intercept), nil
}

func fnFORECAST(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	slope, intercept, errVal, ok := linearRegression([]Value{args[1], args[2]})
	if !ok {
		return errVal, nil
	}
	return NumberVal(intercept + slope*x), nil
}

func fnGEOMEAN(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	sumLn := 0.0
	for _, v := range nums {
		if v <= 0 {
			return ErrorVal(ErrValNUM), nil
		}
		sumLn += math.Log(v)
	}
	return NumberVal(math.Exp(sumLn / float64(n))), nil
}

func fnHARMEAN(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n == 0 {
		return ErrorVal(ErrValNUM), nil
	}
	sumReciprocals := 0.0
	for _, v := range nums {
		if v <= 0 {
			return ErrorVal(ErrValNUM), nil
		}
		sumReciprocals += 1.0 / v
	}
	return NumberVal(float64(n) / sumReciprocals), nil
}

func fnSKEW(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n < 3 {
		return ErrorVal(ErrValDIV0), nil
	}

	// Compute mean.
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean := sum / float64(n)

	// Compute sample standard deviation (n-1 denominator).
	ssq := 0.0
	for _, v := range nums {
		d := v - mean
		ssq += d * d
	}
	s := math.Sqrt(ssq / float64(n-1))
	if s == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	// Compute skewness: (n / ((n-1)*(n-2))) * sum((xi - mean) / s)^3
	sumCubed := 0.0
	for _, v := range nums {
		z := (v - mean) / s
		sumCubed += z * z * z
	}
	nf := float64(n)
	skew := (nf / ((nf - 1) * (nf - 2))) * sumCubed
	return NumberVal(skew), nil
}

func fnFISHER(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			n, e := CoerceNum(v)
			if e != nil {
				return *e
			}
			if n <= -1 || n >= 1 {
				return ErrorVal(ErrValNUM)
			}
			return NumberVal(0.5 * math.Log((1+n)/(1-n)))
		}), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if x <= -1 || x >= 1 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(0.5 * math.Log((1+x)/(1-x))), nil
}

func fnGAMMALN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			n, e := CoerceNum(v)
			if e != nil {
				return *e
			}
			if n <= 0 {
				return ErrorVal(ErrValNUM)
			}
			lg, _ := math.Lgamma(n)
			return NumberVal(lg)
		}), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if x <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	lg, _ := math.Lgamma(x)
	return NumberVal(lg), nil
}

func fnFISHERINV(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			n, e := CoerceNum(v)
			if e != nil {
				return *e
			}
			e2y := math.Exp(2 * n)
			return NumberVal((e2y - 1) / (e2y + 1))
		}), nil
	}
	y, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	e2y := math.Exp(2 * y)
	return NumberVal((e2y - 1) / (e2y + 1)), nil
}
