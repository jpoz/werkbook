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
	Register("CONFIDENCE.NORM", NoCtx(fnConfidenceNorm))
	Register("CONFIDENCE.T", NoCtx(fnConfidenceT))
	Register("COVAR", NoCtx(fnCOVARIANCEP))
	Register("COVARIANCE.P", NoCtx(fnCOVARIANCEP))
	Register("COVARIANCE.S", NoCtx(fnCOVARIANCES))
	Register("INTERCEPT", NoCtx(fnINTERCEPT))
	Register("COUNTBLANK", NoCtx(fnCOUNTBLANK))
	Register("COUNTIF", NoCtx(fnCOUNTIF))
	Register("COUNTIFS", NoCtx(fnCOUNTIFS))
	Register("DEVSQ", NoCtx(fnDEVSQ))
	Register("FISHER", NoCtx(fnFISHER))
	Register("FREQUENCY", NoCtx(fnFREQUENCY))
	Register("FISHERINV", NoCtx(fnFISHERINV))
	Register("FORECAST", NoCtx(fnFORECAST))
	Register("FORECAST.LINEAR", NoCtx(fnFORECAST))
	Register("GAMMALN", NoCtx(fnGAMMALN))
	Register("GAMMALN.PRECISE", NoCtx(fnGAMMALN))
	Register("GEOMEAN", NoCtx(fnGEOMEAN))
	Register("GROWTH", NoCtx(fnGROWTH))
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
	Register("MODE.MULT", NoCtx(fnModeMult))
	Register("MODE.SNGL", NoCtx(fnMODE))
	Register("PERMUTATIONA", NoCtx(fnPERMUTATIONA))
	Register("PEARSON", NoCtx(fnCORREL))
	Register("PERCENTILE", NoCtx(fnPERCENTILE))
	Register("PERCENTILE.EXC", NoCtx(fnPERCENTILEEXC))
	Register("PERCENTILE.INC", NoCtx(fnPERCENTILE))
	Register("QUARTILE", NoCtx(fnQUARTILE))
	Register("QUARTILE.EXC", NoCtx(fnQUARTILEEXC))
	Register("QUARTILE.INC", NoCtx(fnQUARTILE))
	Register("PERCENTRANK", NoCtx(fnPERCENTRANK))
	Register("PERCENTRANK.INC", NoCtx(fnPERCENTRANK))
	Register("PERCENTRANK.EXC", NoCtx(fnPERCENTRANKEXC))
	Register("RSQ", NoCtx(fnRSQ))
	Register("RANK", NoCtx(fnRANK))
	Register("RANK.EQ", NoCtx(fnRANK))
	Register("RANK.AVG", NoCtx(fnRANKAVG))
	Register("SLOPE", NoCtx(fnSLOPE))
	Register("SMALL", NoCtx(fnSMALL))
	Register("STDEV", NoCtx(fnSTDEV))
	Register("STDEV.S", NoCtx(fnSTDEV))
	Register("STDEVA", NoCtx(fnSTDEVA))
	Register("STDEVP", NoCtx(fnSTDEVP))
	Register("STDEV.P", NoCtx(fnSTDEVP))
	Register("STDEVPA", NoCtx(fnSTDEVPA))
	Register("STANDARDIZE", NoCtx(fnSTANDARDIZE))
	Register("STEYX", NoCtx(fnSTEYX))
	Register("SUM", NoCtx(fnSUM))
	Register("SUMIF", NoCtx(fnSUMIF))
	Register("SUMIFS", NoCtx(fnSUMIFS))
	Register("SUMPRODUCT", NoCtx(fnSUMPRODUCT))
	Register("SUMSQ", NoCtx(fnSUMSQ))
	Register("VAR", NoCtx(fnVAR))
	Register("VAR.S", NoCtx(fnVAR))
	Register("TRIMMEAN", NoCtx(fnTRIMMEAN))
	Register("SKEW", NoCtx(fnSKEW))
	Register("SKEW.P", NoCtx(fnSkewP))
	Register("KURT", NoCtx(fnKURT))
	Register("VARA", NoCtx(fnVARA))
	Register("VARP", NoCtx(fnVARP))
	Register("VAR.P", NoCtx(fnVARP))
	Register("VARPA", NoCtx(fnVARPA))
	Register("NORM.DIST", NoCtx(fnNormDist))
	Register("NORM.INV", NoCtx(fnNormInv))
	Register("NORM.S.DIST", NoCtx(fnNormSDist))
	Register("NORM.S.INV", NoCtx(fnNormSInv))
	Register("BINOM.DIST", NoCtx(fnBinomDist))
	Register("BINOM.DIST.RANGE", NoCtx(fnBinomDistRange))
	Register("BINOM.INV", NoCtx(fnBinomInv))
	Register("POISSON.DIST", NoCtx(fnPoissonDist))
	Register("EXPON.DIST", NoCtx(fnExponDist))
	Register("WEIBULL.DIST", NoCtx(fnWeibullDist))
	Register("LOGNORM.DIST", NoCtx(fnLognormDist))
	Register("LOGNORM.INV", NoCtx(fnLognormInv))
	Register("CHISQ.DIST", NoCtx(fnChisqDist))
	Register("CHISQ.INV", NoCtx(fnChisqInv))
	Register("GAMMA.DIST", NoCtx(fnGammaDist))
	Register("GAMMA.INV", NoCtx(fnGammaInv))
	Register("T.DIST", NoCtx(fnTDist))
	Register("T.DIST.RT", NoCtx(fnTDistRT))
	Register("T.DIST.2T", NoCtx(fnTDist2T))
	Register("T.INV", NoCtx(fnTInv))
	Register("T.INV.2T", NoCtx(fnTInv2T))
	Register("BETA.DIST", NoCtx(fnBetaDist))
	Register("BETA.INV", NoCtx(fnBetaInv))
	Register("CHISQ.DIST.RT", NoCtx(fnChisqDistRT))
	Register("CHISQ.INV.RT", NoCtx(fnChisqInvRT))
	Register("CHISQ.TEST", NoCtx(fnChisqTest))
	Register("F.TEST", NoCtx(fnFTest))
	Register("F.DIST", NoCtx(fnFDist))
	Register("F.DIST.RT", NoCtx(fnFDistRT))
	Register("F.INV", NoCtx(fnFInv))
	Register("F.INV.RT", NoCtx(fnFInvRT))
	Register("GAUSS", NoCtx(fnGauss))
	Register("HYPGEOM.DIST", NoCtx(fnHypgeomDist))
	Register("NEGBINOM.DIST", NoCtx(fnNegbinomDist))
	Register("PHI", NoCtx(fnPhi))
	Register("PROB", NoCtx(fnPROB))
	Register("T.TEST", NoCtx(fnTTest))
	Register("Z.TEST", NoCtx(fnZTEST))
	Register("AGGREGATE", NoCtx(fnAggregate))
	Register("TREND", NoCtx(fnTREND))
	Register("LINEST", NoCtx(fnLINEST))
	Register("LOGEST", NoCtx(fnLOGEST))
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
			// A1=TRUE) are not — matching expected behavior.
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

// MatchesCriteria checks if a value matches criteria.
//
// Booleans and numbers are distinct types for criteria matching:
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
	// content parses as the same number (strings are coerced to numbers
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
	// =COUNTIF(range,"TRUE") counts only boolean TRUE cells,
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
// string, respecting the type separation between booleans and numbers.
func matchOperator(v Value, op, operand string) bool {
	// "=" with empty operand matches only truly empty cells (TypeEmpty).
	// This differs from bare "" criteria which matches both empty and empty-string cells.
	if op == "=" && operand == "" {
		return v.Type == ValueEmpty
	}
	// "<>" with empty operand: everything except truly empty cells.
	// Empty-string cells are considered non-blank for <>.
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
	// parses as the same number (text-numbers are coerced for equality
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
// ~* escapes a single literal *, ~? escapes a single literal ?,
// and ~~ escapes a single literal ~.  So "~**" has an escaped * followed
// by an unescaped wildcard *, meaning it matches strings starting with
// a literal '*'.
func classifyWildcard(s string) wildcardMode {
	hasEscape := false
	for i := 0; i < len(s); i++ {
		switch s[i] {
		case '~':
			if i+1 < len(s) && (s[i+1] == '*' || s[i+1] == '?' || s[i+1] == '~') {
				hasEscape = true
				i++ // skip the escaped char
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
			if sumRange.Type == ValueError {
				return sumRange, nil
			}
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
				if sv.Type == ValueError {
					return sv, nil
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

	// Check if any criteria argument is an array (e.g. from ANCHORARRAY).
	// If so, produce an array result by evaluating once per criteria element.
	for k := 2; k < len(args); k += 2 {
		if args[k].Type == ValueArray {
			return sumIFSArray(args)
		}
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
				sv := sumRange.Array[r][c]
				if sv.Type == ValueError {
					return sv, nil
				}
				if n, e := CoerceNum(sv); e == nil {
					sum += n
				}
			}
		}
	}
	return NumberVal(sum), nil
}

// sumIFSArray handles SUMIFS when one or more criteria arguments are arrays
// (dynamic array spill). It iterates over the criteria array elements and
// produces an array of sums.
func sumIFSArray(args []Value) (Value, error) {
	// Find the first array criteria to determine the output shape.
	var arrCrit Value
	for k := 2; k < len(args); k += 2 {
		if args[k].Type == ValueArray {
			arrCrit = args[k]
			break
		}
	}
	rows := len(arrCrit.Array)
	if rows == 0 {
		return NumberVal(0), nil
	}
	cols := len(arrCrit.Array[0])
	result := make([][]Value, rows)
	for i := 0; i < rows; i++ {
		result[i] = make([]Value, cols)
		for j := 0; j < cols; j++ {
			scalarArgs := make([]Value, len(args))
			copy(scalarArgs, args)
			for k := 2; k < len(args); k += 2 {
				if args[k].Type == ValueArray {
					scalarArgs[k] = ArrayElement(args[k], i, j)
				}
			}
			v, _ := fnSUMIFS(scalarArgs)
			result[i][j] = v
		}
	}
	return Value{Type: ValueArray, Array: result}, nil
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

	// Check if any criteria argument is an array (e.g. from ANCHORARRAY).
	// If so, produce an array result by evaluating once per criteria element.
	for k := 1; k < len(args); k += 2 {
		if args[k].Type == ValueArray {
			return countIFSArray(args), nil
		}
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

// countIFSArray handles COUNTIFS when one or more criteria arguments are arrays
// (dynamic array spill). It iterates over the criteria array elements and
// produces an array of counts.
func countIFSArray(args []Value) Value {
	// Find the first array criteria to determine the output shape.
	var arrCrit Value
	for k := 1; k < len(args); k += 2 {
		if args[k].Type == ValueArray {
			arrCrit = args[k]
			break
		}
	}
	rows := len(arrCrit.Array)
	if rows == 0 {
		return NumberVal(0)
	}
	cols := len(arrCrit.Array[0])
	result := make([][]Value, rows)
	for i := 0; i < rows; i++ {
		result[i] = make([]Value, cols)
		for j := 0; j < cols; j++ {
			// Build scalar args for this element.
			scalarArgs := make([]Value, len(args))
			copy(scalarArgs, args)
			for k := 1; k < len(args); k += 2 {
				if args[k].Type == ValueArray {
					scalarArgs[k] = ArrayElement(args[k], i, j)
				}
			}
			v, _ := fnCOUNTIFS(scalarArgs)
			result[i][j] = v
		}
	}
	return Value{Type: ValueArray, Array: result}
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
					} else if avgRange.Type != ValueArray {
						sv = avgRange
					}
					if sv.Type == ValueError {
						return sv, nil
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
	// Propagate any error argument; promote scalar numeric/bool/string args to
	// 1x1 arrays so SUMPRODUCT(scalar) matches Excel (e.g. SUMPRODUCT(5)=5,
	// SUMPRODUCT(#N/A)=#N/A).
	promoted := make([]Value, len(args))
	for i, arg := range args {
		if arg.Type == ValueError {
			return arg, nil
		}
		if arg.Type == ValueArray {
			promoted[i] = arg
			continue
		}
		promoted[i] = Value{Type: ValueArray, Array: [][]Value{{arg}}}
	}
	args = promoted
	firstArr := args[0].Array
	rows := len(firstArr)
	cols := 0
	if rows > 0 {
		cols = len(firstArr[0])
	}

	for _, arg := range args[1:] {
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
				// Excel treats text and boolean cell values as 0 in
				// SUMPRODUCT.  Computed booleans (e.g. from A1:A5>3)
				// will already have been coerced to numbers by the
				// arithmetic operators before reaching this function.
				if cell.Type == ValueString || cell.Type == ValueBool {
					product = 0
					continue
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
			// Single-cell references to text, booleans, or empty cells
			// are ignored, matching the range/array behavior above.
			if arg.FromCell && (arg.Type == ValueString || arg.Type == ValueBool || arg.Type == ValueEmpty) {
				continue
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
				sv := avgRange.Array[r][c]
				if sv.Type == ValueError {
					return sv, nil
				}
				if n, e := CoerceNum(sv); e == nil {
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
				sv := maxRange.Array[r][c]
				if sv.Type == ValueError {
					return sv, nil
				}
				if n, e := CoerceNum(sv); e == nil {
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
				sv := minRange.Array[r][c]
				if sv.Type == ValueError {
					return sv, nil
				}
				if n, e := CoerceNum(sv); e == nil {
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

func fnModeMult(args []Value) (Value, error) {
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

	// Count frequency and track insertion order.
	freq := make(map[float64]int)
	for _, n := range nums {
		freq[n]++
	}

	// Find the maximum frequency.
	maxFreq := 0
	for _, c := range freq {
		if c > maxFreq {
			maxFreq = c
		}
	}
	if maxFreq < 2 {
		return ErrorVal(ErrValNA), nil
	}

	// Collect all values with the maximum frequency.
	modes := make([]float64, 0)
	for v, c := range freq {
		if c == maxFreq {
			modes = append(modes, v)
		}
	}

	// Sort in ascending order (matches Excel behaviour).
	sort.Float64s(modes)

	// Return as a vertical array (each mode in its own row).
	result := make([][]Value, len(modes))
	for i, m := range modes {
		result[i] = []Value{NumberVal(m)}
	}
	return Value{Type: ValueArray, Array: result}, nil
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
	// Round to 15 significant digits first (matching expected precision)
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

// collectNumericA collects numeric values using the "A" variant rules:
// In arrays/ranges: numbers kept, TRUE→1, FALSE→0, text→0, empty ignored, errors propagated.
// Direct args: numbers kept, TRUE→1, FALSE→0, text coerced to number (error if not numeric), empty ignored.
func collectNumericA(args []Value) ([]float64, *Value) {
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
					switch cell.Type {
					case ValueError:
						return nil, &cell
					case ValueNumber:
						nums = append(nums, cell.Num)
					case ValueBool:
						if cell.Bool {
							nums = append(nums, 1)
						} else {
							nums = append(nums, 0)
						}
					case ValueString:
						nums = append(nums, 0)
					case ValueEmpty:
						// ignored
					}
				}
			}
		} else {
			switch arg.Type {
			case ValueError:
				return nil, &arg
			case ValueEmpty:
				// ignored
			case ValueString:
				n, e := CoerceNum(arg)
				if e != nil {
					return nil, e
				}
				nums = append(nums, n)
			default:
				n, e := CoerceNum(arg)
				if e != nil {
					return nil, e
				}
				nums = append(nums, n)
			}
		}
	}
	return nums, nil
}

func fnVARA(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, ev := collectNumericA(args)
	if ev != nil {
		return *ev, nil
	}
	n := len(nums)
	if n < 2 {
		return ErrorVal(ErrValDIV0), nil
	}
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean := sum / float64(n)
	ssq := 0.0
	for _, v := range nums {
		d := v - mean
		ssq += d * d
	}
	return NumberVal(ssq / float64(n-1)), nil
}

func fnSTDEVA(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, ev := collectNumericA(args)
	if ev != nil {
		return *ev, nil
	}
	n := len(nums)
	if n < 2 {
		return ErrorVal(ErrValDIV0), nil
	}
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean := sum / float64(n)
	ssq := 0.0
	for _, v := range nums {
		d := v - mean
		ssq += d * d
	}
	return NumberVal(math.Sqrt(ssq / float64(n-1))), nil
}

func fnSTDEVPA(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, ev := collectNumericA(args)
	if ev != nil {
		return *ev, nil
	}
	n := len(nums)
	if n < 1 {
		return ErrorVal(ErrValDIV0), nil
	}
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean := sum / float64(n)
	ssq := 0.0
	for _, v := range nums {
		d := v - mean
		ssq += d * d
	}
	return NumberVal(math.Sqrt(ssq / float64(n))), nil
}

func fnVARPA(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, ev := collectNumericA(args)
	if ev != nil {
		return *ev, nil
	}
	n := len(nums)
	if n < 1 {
		return ErrorVal(ErrValDIV0), nil
	}
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean := sum / float64(n)
	ssq := 0.0
	for _, v := range nums {
		d := v - mean
		ssq += d * d
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

func fnRSQ(args []Value) (Value, error) {
	r, err := fnCORREL(args)
	if err != nil {
		return r, err
	}
	if r.Type != ValueNumber {
		return r, nil
	}
	return NumberVal(r.Num * r.Num), nil
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

// fnSTEYX returns the standard error of the predicted y-value for each x in
// a linear regression.  STEYX(known_y's, known_x's).
func fnSTEYX(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	flatY := flattenValuesGeneric(args[0])
	flatX := flattenValuesGeneric(args[1])

	if len(flatY) != len(flatX) {
		return ErrorVal(ErrValNA), nil
	}

	// Walk paired positions; keep only pairs where BOTH values are numeric.
	var xs, ys []float64
	for i := range flatY {
		vy, vx := flatY[i], flatX[i]
		if vy.Type == ValueError {
			return vy, nil
		}
		if vx.Type == ValueError {
			return vx, nil
		}
		if vy.Type != ValueNumber || vx.Type != ValueNumber {
			continue
		}
		xs = append(xs, vx.Num)
		ys = append(ys, vy.Num)
	}

	n := len(xs)
	if n < 3 {
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

	// Compute SSx, SSy, SSxy.
	ssX, ssY, ssXY := 0.0, 0.0, 0.0
	for i := 0; i < n; i++ {
		dx := xs[i] - meanX
		dy := ys[i] - meanY
		ssX += dx * dx
		ssY += dy * dy
		ssXY += dx * dy
	}

	if ssX == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	result := math.Sqrt((1 / float64(n-2)) * (ssY - ssXY*ssXY/ssX))
	return NumberVal(result), nil
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

func fnSkewP(args []Value) (Value, error) {
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

	// Compute population standard deviation (n denominator).
	ssq := 0.0
	for _, v := range nums {
		d := v - mean
		ssq += d * d
	}
	sigma := math.Sqrt(ssq / float64(n))
	if sigma == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	// Compute population skewness: (1/n) * sum((xi - mean) / sigma)^3
	sumCubed := 0.0
	for _, v := range nums {
		z := (v - mean) / sigma
		sumCubed += z * z * z
	}
	skewP := sumCubed / float64(n)
	return NumberVal(skewP), nil
}

func fnKURT(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	nums, e := collectNumeric(args)
	if e != nil {
		return *e, nil
	}
	n := len(nums)
	if n < 4 {
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

	// Compute kurtosis:
	// [n(n+1)/((n-1)(n-2)(n-3)) * sum((xi-mean)/s)^4] - 3(n-1)^2/((n-2)(n-3))
	sumFourth := 0.0
	for _, v := range nums {
		z := (v - mean) / s
		z2 := z * z
		sumFourth += z2 * z2
	}
	nf := float64(n)
	kurt := (nf*(nf+1))/((nf-1)*(nf-2)*(nf-3))*sumFourth - 3*(nf-1)*(nf-1)/((nf-2)*(nf-3))
	return NumberVal(kurt), nil
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

func fnSTANDARDIZE(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	mean, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	stddev, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	if stddev <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal((x - mean) / stddev), nil
}

// fnPERMUTATIONA returns the number of permutations with repetitions: number^number_chosen.
func fnPERMUTATIONA(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	number, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	numberChosen, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	number = math.Trunc(number)
	numberChosen = math.Trunc(numberChosen)
	if number < 0 || numberChosen < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if number == 0 && numberChosen > 0 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(math.Pow(number, numberChosen)), nil
}

// fnFREQUENCY implements the FREQUENCY(data_array, bins_array) function.
// It counts how many values in data_array fall into each interval defined by
// bins_array. The result is a vertical array with len(bins)+1 rows.
func fnFREQUENCY(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	// collectNums flattens an argument into a slice of float64 values,
	// ignoring text, booleans, and empty cells (matching expected behaviour
	// for array arguments). Propagates errors.
	collectNums := func(v Value) ([]float64, *Value) {
		var nums []float64
		if v.Type == ValueArray {
			for _, row := range v.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return nil, &cell
					}
					if cell.Type == ValueNumber {
						nums = append(nums, cell.Num)
					}
					// text, bool, empty → skip
				}
			}
		} else {
			if v.Type == ValueError {
				return nil, &v
			}
			if v.Type == ValueNumber {
				nums = append(nums, v.Num)
			}
		}
		return nums, nil
	}

	data, errVal := collectNums(args[0])
	if errVal != nil {
		return *errVal, nil
	}
	bins, errVal := collectNums(args[1])
	if errVal != nil {
		return *errVal, nil
	}

	// If bins_array is empty, return single-element array with count of data.
	if len(bins) == 0 {
		result := [][]Value{{NumberVal(float64(len(data)))}}
		return Value{Type: ValueArray, Array: result}, nil
	}

	// Sort bins in ascending order.
	sort.Float64s(bins)

	// Build frequency counts: len(bins)+1 buckets.
	counts := make([]float64, len(bins)+1)
	for _, v := range data {
		placed := false
		for i, b := range bins {
			if v <= b {
				counts[i]++
				placed = true
				break
			}
		}
		if !placed {
			counts[len(bins)]++
		}
	}

	// Return as vertical array (n+1 rows, 1 column).
	result := make([][]Value, len(counts))
	for i, c := range counts {
		result[i] = []Value{NumberVal(c)}
	}
	return Value{Type: ValueArray, Array: result}, nil
}

// ---------------------------------------------------------------------------
// normSDistCDF / normSDistPDF — internal helpers for standard normal
// ---------------------------------------------------------------------------

// normSDistCDF returns the CDF of the standard normal distribution: Φ(z).
func normSDistCDF(z float64) float64 {
	return 0.5 * (1 + math.Erf(z/math.Sqrt(2)))
}

// normSDistPDF returns the PDF of the standard normal distribution: φ(z).
func normSDistPDF(z float64) float64 {
	return (1.0 / math.Sqrt(2*math.Pi)) * math.Exp(-z*z/2)
}

// ---------------------------------------------------------------------------
// NORM.DIST — Normal distribution (PDF or CDF)
// ---------------------------------------------------------------------------

func fnNormDist(args []Value) (Value, error) {
	if len(args) != 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	mean, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	stdev, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	if stdev <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	cum, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	if cum != 0 {
		// CDF: standardize then use Φ
		z := (x - mean) / stdev
		return NumberVal(normSDistCDF(z)), nil
	}
	// PDF: (1/(stdev*√(2π))) * exp(-((x-mean)/stdev)²/2)
	z := (x - mean) / stdev
	return NumberVal(normSDistPDF(z) / stdev), nil
}

// ---------------------------------------------------------------------------
// NORM.INV — Inverse of the normal CDF
// ---------------------------------------------------------------------------

func fnNormInv(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	p, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	mean, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	stdev, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	if p <= 0 || p >= 1 {
		return ErrorVal(ErrValNUM), nil
	}
	if stdev <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	x := mean + stdev*normSInv(p)
	// Refine with Newton-Raphson on the full normal distribution.
	// normSInv already provides high precision, but this ensures
	// consistency across all parameter combinations.
	for range 2 {
		z := (x - mean) / stdev
		cdf := normSDistCDF(z)
		pdf := normSDistPDF(z)
		if pdf > 0 {
			x -= (cdf - p) / (pdf / stdev)
		}
	}
	return NumberVal(x), nil
}

// ---------------------------------------------------------------------------
// NORM.S.DIST — Standard normal distribution (PDF or CDF)
// ---------------------------------------------------------------------------

func fnNormSDist(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	z, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	cum, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	if cum != 0 {
		return NumberVal(normSDistCDF(z)), nil
	}
	return NumberVal(normSDistPDF(z)), nil
}

// ---------------------------------------------------------------------------
// PHI — Standard normal PDF φ(x)
// ---------------------------------------------------------------------------

func fnPhi(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(normSDistPDF(x)), nil
}

// ---------------------------------------------------------------------------
// GAUSS — P(0 < Z < z) = NORM.S.DIST(z, TRUE) - 0.5
// ---------------------------------------------------------------------------

func fnGauss(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	z, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return NumberVal(normSDistCDF(z) - 0.5), nil
}

// ---------------------------------------------------------------------------
// NORM.S.INV — Inverse of the standard normal CDF
// ---------------------------------------------------------------------------

func fnNormSInv(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	p, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	if p <= 0 || p >= 1 {
		return ErrorVal(ErrValNUM), nil
	}
	return NumberVal(normSInv(p)), nil
}

// normSInv computes the inverse of the standard normal CDF using
// Peter Acklam's rational approximation algorithm.
func normSInv(p float64) float64 {
	const (
		pLow  = 0.02425
		pHigh = 1 - pLow
	)

	// Coefficients for the rational approximation.
	var (
		a = [6]float64{
			-3.969683028665376e+01,
			2.209460984245205e+02,
			-2.759285104469687e+02,
			1.383577518672690e+02,
			-3.066479806614716e+01,
			2.506628277459239e+00,
		}
		b = [5]float64{
			-5.447609879822406e+01,
			1.615858368580409e+02,
			-1.556989798598866e+02,
			6.680131188771972e+01,
			-1.328068155288572e+01,
		}
		c = [6]float64{
			-7.784894002430293e-03,
			-3.223964580411365e-01,
			-2.400758277161838e+00,
			-2.549732539343734e+00,
			4.374664141464968e+00,
			2.938163982698783e+00,
		}
		d = [4]float64{
			7.784695709041462e-03,
			3.224671290700398e-01,
			2.445134137142996e+00,
			3.754408661907416e+00,
		}
	)

	var x float64
	if p < pLow {
		// Lower region.
		q := math.Sqrt(-2 * math.Log(p))
		x = (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q + c[5]) /
			((((d[0]*q+d[1])*q+d[2])*q+d[3])*q + 1)
	} else if p <= pHigh {
		// Central region.
		q := p - 0.5
		r := q * q
		x = (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r + a[5]) * q /
			(((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+b[4])*r + 1)
	} else {
		// Upper region.
		q := math.Sqrt(-2 * math.Log(1-p))
		x = -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q + c[5]) /
			((((d[0]*q+d[1])*q+d[2])*q+d[3])*q + 1)
	}

	// One Halley step closes most of the gap to full float64 precision.
	e := 0.5*math.Erfc(-x/math.Sqrt2) - p
	u := e * math.Sqrt(2*math.Pi) * math.Exp(x*x/2)
	x = x - u/(1+x*u/2)

	// Two Newton-Raphson iterations bring us to ~15+ digits of precision,
	// matching Excel's output for NORM.S.INV.
	for range 2 {
		cdf := normSDistCDF(x)
		pdf := normSDistPDF(x)
		if pdf > 0 {
			x -= (cdf - p) / pdf
		}
	}

	return x
}

// ---------------------------------------------------------------------------
// BINOM.DIST — Binomial distribution (PMF or CDF)
// ---------------------------------------------------------------------------

func fnBinomDist(args []Value) (Value, error) {
	if len(args) != 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	sf, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	tf, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	p, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	cum, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}

	// Truncate to integers.
	k := int(sf)
	n := int(tf)

	if k < 0 || k > n {
		return ErrorVal(ErrValNUM), nil
	}
	if p < 0 || p > 1 {
		return ErrorVal(ErrValNUM), nil
	}

	if cum != 0 {
		// CDF: P(X <= k) = sum from i=0 to k of PMF(i)
		sum := 0.0
		for i := 0; i <= k; i++ {
			sum += binomPMF(n, i, p)
		}
		return NumberVal(sum), nil
	}
	return NumberVal(binomPMF(n, k, p)), nil
}

// binomPMF returns the binomial probability mass function:
// P(X=k) = C(n,k) * p^k * (1-p)^(n-k)
// Uses log-gamma for numerical stability with large n.
func binomPMF(n, k int, p float64) float64 {
	if p == 0 {
		if k == 0 {
			return 1
		}
		return 0
	}
	if p == 1 {
		if k == n {
			return 1
		}
		return 0
	}
	// log(C(n,k)) = lgamma(n+1) - lgamma(k+1) - lgamma(n-k+1)
	logC, _ := math.Lgamma(float64(n + 1))
	logK, _ := math.Lgamma(float64(k + 1))
	logNK, _ := math.Lgamma(float64(n - k + 1))
	logBinom := logC - logK - logNK
	logProb := logBinom + float64(k)*math.Log(p) + float64(n-k)*math.Log(1-p)
	return math.Exp(logProb)
}

// ---------------------------------------------------------------------------
// BINOM.DIST.RANGE — Binomial distribution range probability
// ---------------------------------------------------------------------------

// fnBinomDistRange implements BINOM.DIST.RANGE(trials, probability_s, number_s, [number_s2]).
// Returns the probability of a trial result using a binomial distribution.
// When number_s2 is omitted it equals number_s (single-point probability).
func fnBinomDistRange(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	nf, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	p, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	sf, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	n := int(nf)
	s := int(sf)
	s2 := s
	if len(args) == 4 {
		s2f, e := CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
		s2 = int(s2f)
	}
	// Validate constraints.
	if n < 0 || p < 0 || p > 1 || s < 0 || s > n || s2 < s || s2 > n {
		return ErrorVal(ErrValNUM), nil
	}
	sum := 0.0
	for k := s; k <= s2; k++ {
		sum += binomPMF(n, k, p)
	}
	return NumberVal(sum), nil
}

// ---------------------------------------------------------------------------
// BINOM.INV — Inverse binomial distribution
// ---------------------------------------------------------------------------

// fnBinomInv implements BINOM.INV(trials, probability_s, alpha).
// Returns the smallest k such that BINOM.DIST(k, trials, p, TRUE) >= alpha.
func fnBinomInv(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	tf, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	p, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	alpha, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	// Truncate trials to integer.
	n := int(tf)

	if n < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if p <= 0 || p >= 1 {
		return ErrorVal(ErrValNUM), nil
	}
	if alpha <= 0 || alpha >= 1 {
		return ErrorVal(ErrValNUM), nil
	}

	// Accumulate the binomial CDF and return the first k where CDF >= alpha.
	cum := 0.0
	for k := 0; k <= n; k++ {
		cum += binomPMF(n, k, p)
		if cum >= alpha {
			return NumberVal(float64(k)), nil
		}
	}
	// Fallback: should not be reached for valid inputs since CDF(n) = 1.
	return NumberVal(float64(n)), nil
}

// fnPoissonDist implements POISSON.DIST(x, mean, cumulative).
func fnPoissonDist(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	xf, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	mean, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	cum, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	// Truncate x to integer.
	k := int(xf)

	if k < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if mean < 0 {
		return ErrorVal(ErrValNUM), nil
	}

	// Special case: mean == 0.
	if mean == 0 {
		if k == 0 {
			return NumberVal(1), nil
		}
		if cum != 0 {
			return NumberVal(1), nil
		}
		return NumberVal(0), nil
	}

	if cum != 0 {
		// CDF: P(X <= k) = sum from i=0 to k of PMF(i)
		sum := 0.0
		for i := 0; i <= k; i++ {
			sum += poissonPMF(i, mean)
		}
		return NumberVal(sum), nil
	}
	return NumberVal(poissonPMF(k, mean)), nil
}

// poissonPMF returns the Poisson probability mass function:
// P(X=k) = (mean^k * e^(-mean)) / k!
// Uses log-gamma for numerical stability.
func poissonPMF(k int, mean float64) float64 {
	lg, _ := math.Lgamma(float64(k + 1))
	return math.Exp(float64(k)*math.Log(mean) - mean - lg)
}

// ---------------------------------------------------------------------------
// EXPON.DIST — Exponential distribution
// ---------------------------------------------------------------------------

func fnExponDist(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	lambda, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	cum, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	if x < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if lambda <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	if cum != 0 {
		// CDF: F(x) = 1 - exp(-lambda * x)
		return NumberVal(1 - math.Exp(-lambda*x)), nil
	}
	// PDF: f(x) = lambda * exp(-lambda * x)
	return NumberVal(lambda * math.Exp(-lambda*x)), nil
}

// WEIBULL.DIST — Weibull distribution (PDF or CDF)
// ---------------------------------------------------------------------------

func fnWeibullDist(args []Value) (Value, error) {
	if len(args) != 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	alpha, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	beta, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	cum, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}

	if x < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if alpha <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if beta <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	if cum != 0 {
		// CDF: F(x) = 1 - exp(-(x/beta)^alpha)
		return NumberVal(1 - math.Exp(-math.Pow(x/beta, alpha))), nil
	}
	// PDF: f(x) = (alpha/beta) * (x/beta)^(alpha-1) * exp(-(x/beta)^alpha)
	if x == 0 {
		// Returns 0 for all PDF evaluations at x=0.
		return NumberVal(0), nil
	}
	return NumberVal((alpha / beta) * math.Pow(x/beta, alpha-1) * math.Exp(-math.Pow(x/beta, alpha))), nil
}

// ---------------------------------------------------------------------------
// LOGNORM.DIST — Lognormal distribution (PDF or CDF)
// ---------------------------------------------------------------------------

func fnLognormDist(args []Value) (Value, error) {
	if len(args) != 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	mean, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	stdev, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	cum, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}

	if x <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if stdev <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	if cum != 0 {
		// CDF: Φ((ln(x) - μ) / σ)
		z := (math.Log(x) - mean) / stdev
		return NumberVal(normSDistCDF(z)), nil
	}
	// PDF: (1 / (x * σ * √(2π))) * exp(-((ln(x) - μ)² / (2σ²)))
	lnx := math.Log(x)
	return NumberVal((1 / (x * stdev * math.Sqrt(2*math.Pi))) * math.Exp(-(lnx-mean)*(lnx-mean)/(2*stdev*stdev))), nil
}

// ---------------------------------------------------------------------------
// LOGNORM.INV — Inverse of the lognormal CDF
// ---------------------------------------------------------------------------

func fnLognormInv(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	p, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	mean, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	stdev, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	if p <= 0 || p >= 1 {
		return ErrorVal(ErrValNUM), nil
	}
	if stdev <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	// LOGNORM.INV(p, μ, σ) = exp(μ + σ * NORM.S.INV(p))
	return NumberVal(math.Exp(mean + stdev*normSInv(p))), nil
}

// ---------------------------------------------------------------------------
// GAMMA.DIST — Gamma distribution (PDF or CDF)
// ---------------------------------------------------------------------------

// regLowerGamma returns the regularized lower incomplete gamma function P(a, x).
// It uses the series expansion for x < a+1 and the continued fraction
// (via the complementary function Q) for x >= a+1.
func regLowerGamma(a, x float64) float64 {
	if x == 0 {
		return 0
	}
	if x < a+1 {
		return regLowerGammaSeries(a, x)
	}
	// Use complementary: P(a,x) = 1 - Q(a,x)
	return 1 - regUpperGammaCF(a, x)
}

// regLowerGammaSeries computes P(a, x) via the series expansion:
//
//	P(a,x) = e^(-x) * x^a / Γ(a) * Σ_{n=0}^∞ x^n / (a*(a+1)*...*(a+n))
func regLowerGammaSeries(a, x float64) float64 {
	sum := 1.0 / a
	term := 1.0 / a
	for n := 1; n < 1000; n++ {
		term *= x / (a + float64(n))
		sum += term
		if math.Abs(term) < 1e-15*math.Abs(sum) {
			break
		}
	}
	lgA, _ := math.Lgamma(a)
	return math.Exp(-x+a*math.Log(x)-lgA) * sum
}

// regUpperGammaCF computes Q(a, x) = 1 - P(a, x) via the Lentz continued fraction.
func regUpperGammaCF(a, x float64) float64 {
	const eps = 1e-15
	const tiny = 1e-30

	// Modified Lentz's method for the CF representation of Q(a,x).
	// CF: Q(a,x) = e^(-x)*x^a/Γ(a) * 1/(x+1-a- 1*(1-a)/(x+3-a- 2*(2-a)/(x+5-a- ...)))
	// Using the standard form: b0=0, a1=1, b1=x+1-a, then an = -n*(n-a), bn = x+2n+1-a
	f := tiny
	c := f
	d := 0.0
	for n := 1; n < 1000; n++ {
		an := float64(0)
		bn := float64(0)
		if n == 1 {
			an = 1.0
			bn = x + 1 - a
		} else {
			nf := float64(n - 1)
			an = -nf * (nf - a)
			bn = x + 2*nf + 1 - a
		}

		d = bn + an*d
		if math.Abs(d) < tiny {
			d = tiny
		}
		c = bn + an/c
		if math.Abs(c) < tiny {
			c = tiny
		}
		d = 1.0 / d
		delta := c * d
		f *= delta
		if math.Abs(delta-1.0) < eps {
			break
		}
	}

	lgA, _ := math.Lgamma(a)
	return math.Exp(-x+a*math.Log(x)-lgA) * f
}

func fnGammaDist(args []Value) (Value, error) {
	if len(args) != 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	alpha, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	beta, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	cum, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}

	if x < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if alpha <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if beta <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	if cum != 0 {
		// CDF: regularized lower incomplete gamma function P(alpha, x/beta)
		return NumberVal(regLowerGamma(alpha, x/beta)), nil
	}

	// PDF: f(x) = (1 / (beta^alpha * Γ(alpha))) * x^(alpha-1) * exp(-x/beta)
	if x == 0 {
		if alpha > 1 {
			return NumberVal(0), nil
		}
		// alpha <= 1: returns #NUM! at x=0 (PDF diverges for alpha<1,
		// and also returns #NUM! for the alpha==1 boundary case).
		return ErrorVal(ErrValNUM), nil
	}

	lgA, _ := math.Lgamma(alpha)
	logPdf := (alpha-1)*math.Log(x) - x/beta - alpha*math.Log(beta) - lgA
	return NumberVal(math.Exp(logPdf)), nil
}

// ---------------------------------------------------------------------------
// GAMMA.INV — Inverse of the gamma cumulative distribution function
// ---------------------------------------------------------------------------
// Given p, alpha, beta it finds x such that GAMMA.DIST(x, alpha, beta, TRUE) = p.

func fnGammaInv(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	p, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	alpha, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	beta, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	// Validate ranges.
	if p < 0 || p > 1 {
		return ErrorVal(ErrValNUM), nil
	}
	if alpha <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if beta <= 0 {
		return ErrorVal(ErrValNUM), nil
	}

	// Edge cases.
	if p == 0 {
		return NumberVal(0), nil
	}
	// Returns #NUM! for probability = 1.
	if p == 1 {
		return ErrorVal(ErrValNUM), nil
	}

	// Initial guess using the Wilson-Hilferty normal approximation:
	//   x/alpha ≈ (1 - 1/(9*alpha) + z * sqrt(1/(9*alpha)))^3
	// where z = NORM.S.INV(p).
	z := normSInv(p)
	t := 1.0 / (9 * alpha)
	wh := 1 - t + z*math.Sqrt(t)
	var x float64
	if wh > 0 {
		x = alpha * beta * wh * wh * wh
	} else {
		// Fallback for cases where the WH approximation gives non-positive.
		x = alpha * beta * 0.5
	}
	if x <= 0 {
		x = beta * 0.001
	}

	// Newton-Raphson iteration.
	// f(x) = regLowerGamma(alpha, x/beta) - p
	// f'(x) = gammaPDF(x, alpha, beta)
	//       = x^(alpha-1) * exp(-x/beta) / (beta^alpha * Γ(alpha))
	lgA, _ := math.Lgamma(alpha)
	const maxIter = 200
	const tol = 1e-12

	for i := 0; i < maxIter; i++ {
		cdf := regLowerGamma(alpha, x/beta)
		f := cdf - p

		if math.Abs(f) < tol {
			return NumberVal(x), nil
		}

		// Gamma PDF as the derivative of the CDF.
		logPdf := (alpha-1)*math.Log(x) - x/beta - alpha*math.Log(beta) - lgA
		pdf := math.Exp(logPdf)

		if pdf < 1e-300 {
			// PDF too small for Newton step; use bisection fallback.
			break
		}

		step := f / pdf
		xNew := x - step
		// Ensure x stays positive.
		if xNew <= 0 {
			x = x / 2
		} else {
			x = xNew
		}
	}

	// If Newton didn't converge, fall back to bisection.
	lo := 0.0
	hi := x
	// Expand hi until CDF(hi) > p.
	for regLowerGamma(alpha, hi/beta) < p {
		hi *= 2
		if hi > 1e308 {
			return ErrorVal(ErrValNA), nil
		}
	}

	for i := 0; i < maxIter; i++ {
		mid := (lo + hi) / 2
		cdf := regLowerGamma(alpha, mid/beta)
		if math.Abs(cdf-p) < tol {
			return NumberVal(mid), nil
		}
		if cdf < p {
			lo = mid
		} else {
			hi = mid
		}
		if (hi - lo) < tol*hi {
			return NumberVal((lo + hi) / 2), nil
		}
	}

	return ErrorVal(ErrValNA), nil
}

// ---------------------------------------------------------------------------
// CHISQ.DIST — Chi-squared distribution (PDF or CDF)
// ---------------------------------------------------------------------------

func fnChisqDist(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	dfRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	cum, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	if x < 0 {
		return ErrorVal(ErrValNUM), nil
	}

	// Truncate deg_freedom to integer.
	df := math.Trunc(dfRaw)
	if df < 1 || df > 1e10 {
		return ErrorVal(ErrValNUM), nil
	}

	alpha := df / 2.0

	if cum != 0 {
		// CDF: regularized lower incomplete gamma P(df/2, x/2)
		return NumberVal(regLowerGamma(alpha, x/2.0)), nil
	}

	// PDF: gamma PDF with alpha=df/2, beta=2
	// f(x) = x^(alpha-1) * exp(-x/2) / (2^alpha * Γ(alpha))
	if x == 0 {
		if df == 1 {
			// PDF diverges at x=0 for df=1 (alpha=0.5).
			// Returns Inf.
			return NumberVal(math.Inf(1)), nil
		}
		if df == 2 {
			// alpha=1, PDF = exp(0)/(2*1) = 0.5
			return NumberVal(0.5), nil
		}
		// df > 2 ⇒ alpha > 1 ⇒ PDF = 0
		return NumberVal(0), nil
	}

	lgA, _ := math.Lgamma(alpha)
	logPdf := (alpha-1)*math.Log(x) - x/2.0 - alpha*math.Log(2) - lgA
	return NumberVal(math.Exp(logPdf)), nil
}

// ---------------------------------------------------------------------------
// CHISQ.INV — Inverse of the left-tailed chi-squared distribution
// ---------------------------------------------------------------------------
// CHISQ.INV(probability, deg_freedom)
// Since chi-squared is gamma(alpha=df/2, beta=2):
//   CHISQ.INV(p, df) = GAMMA.INV(p, df/2, 2)

func fnChisqInv(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	p, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	dfRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	if p < 0 || p > 1 {
		return ErrorVal(ErrValNUM), nil
	}

	// Truncate deg_freedom to integer.
	df := math.Trunc(dfRaw)
	if df < 1 || df > 1e10 {
		return ErrorVal(ErrValNUM), nil
	}

	// Edge cases.
	if p == 0 {
		return NumberVal(0), nil
	}
	if p == 1 {
		return ErrorVal(ErrValNUM), nil
	}

	// Delegate to GAMMA.INV(p, df/2, 2).
	return fnGammaInv([]Value{NumberVal(p), NumberVal(df / 2), NumberVal(2)})
}

// ---------------------------------------------------------------------------
// CHISQ.DIST.RT — Right-tailed chi-squared distribution
// ---------------------------------------------------------------------------
// CHISQ.DIST.RT(x, deg_freedom)
//   x           – must be >= 0 (if x < 0, #NUM!)
//   deg_freedom – truncated to integer, must be >= 1
// Returns 1 - CHISQ.DIST(x, df, TRUE).

func fnChisqDistRT(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	dfRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	if x < 0 {
		return ErrorVal(ErrValNUM), nil
	}

	df := math.Trunc(dfRaw)
	if df < 1 || df > 1e10 {
		return ErrorVal(ErrValNUM), nil
	}

	alpha := df / 2.0
	return NumberVal(1 - regLowerGamma(alpha, x/2.0)), nil
}

// ---------------------------------------------------------------------------
// CHISQ.INV.RT — Right-tailed inverse chi-squared distribution
// ---------------------------------------------------------------------------
// CHISQ.INV.RT(probability, deg_freedom)
//   probability  – 0 <= p <= 1
//   deg_freedom  – truncated to integer, must be >= 1
// Returns CHISQ.INV(1 - probability, df).

func fnChisqInvRT(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	p, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	dfRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	if p < 0 || p > 1 {
		return ErrorVal(ErrValNUM), nil
	}

	df := math.Trunc(dfRaw)
	if df < 1 || df > 1e10 {
		return ErrorVal(ErrValNUM), nil
	}

	// Edge cases.
	if p == 1 {
		return NumberVal(0), nil
	}
	if p == 0 {
		return ErrorVal(ErrValNUM), nil
	}

	// Delegate to CHISQ.INV(1-p, df) = GAMMA.INV(1-p, df/2, 2).
	return fnGammaInv([]Value{NumberVal(1 - p), NumberVal(df / 2), NumberVal(2)})
}

// ---------------------------------------------------------------------------
// CHISQ.TEST — Chi-squared test for independence
// ---------------------------------------------------------------------------
// CHISQ.TEST(actual_range, expected_range)
//   actual_range   – observed data (2D range)
//   expected_range – expected data (2D range, same dimensions)
// Returns the p-value from the chi-squared goodness-of-fit test.
//   χ² = Σ (Aij - Eij)² / Eij
//   df = (r-1)(c-1) when both r>1 and c>1; otherwise max(r,c)-1.
//   p  = 1 - regLowerGamma(df/2, χ²/2)

func fnChisqTest(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	actual, ev := normalizeToGrid(args[0])
	if ev != nil {
		return *ev, nil
	}
	expected, ev := normalizeToGrid(args[1])
	if ev != nil {
		return *ev, nil
	}
	ar, ac := gridDims(actual)
	er, ec := gridDims(expected)
	if ar != er || ac != ec {
		return ErrorVal(ErrValNA), nil
	}
	if ar == 1 && ac == 1 {
		return ErrorVal(ErrValNA), nil
	}
	// Compute chi-squared statistic.
	chiSq := 0.0
	for i := 0; i < ar; i++ {
		for j := 0; j < ac; j++ {
			a, e1 := CoerceNum(actual[i][j])
			if e1 != nil {
				return *e1, nil
			}
			e, e2 := CoerceNum(expected[i][j])
			if e2 != nil {
				return *e2, nil
			}
			if e == 0 {
				return ErrorVal(ErrValDIV0), nil
			}
			d := a - e
			chiSq += d * d / e
		}
	}
	// Degrees of freedom.
	var df int
	if ar > 1 && ac > 1 {
		df = (ar - 1) * (ac - 1)
	} else if ar == 1 {
		df = ac - 1
	} else {
		df = ar - 1
	}
	// Right-tail probability using regularized lower incomplete gamma.
	alpha := float64(df) / 2.0
	p := 1 - regLowerGamma(alpha, chiSq/2.0)
	return NumberVal(p), nil
}

// ---------------------------------------------------------------------------
// T.DIST — Student's t-distribution (left-tailed)
// ---------------------------------------------------------------------------
// T.DIST(x, deg_freedom, cumulative)
//   x           – numeric value at which to evaluate
//   deg_freedom – degrees of freedom (truncated to integer, must be >= 1)
//   cumulative  – TRUE for CDF, FALSE for PDF

func fnTDist(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	dfRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	cum, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	// Truncate deg_freedom to integer.
	df := math.Trunc(dfRaw)
	if df < 1 {
		return ErrorVal(ErrValNUM), nil
	}

	if cum != 0 {
		return NumberVal(tDistCDF(x, df)), nil
	}
	return NumberVal(tDistPDF(x, df)), nil
}

// tDistCDF computes the CDF of the Student's t-distribution at x with df
// degrees of freedom: P(T <= x).
func tDistCDF(x, df float64) float64 {
	bx := df / (df + x*x)
	beta := regBetaInc(bx, df/2, 0.5)
	if x >= 0 {
		return 1 - 0.5*beta
	}
	return 0.5 * beta
}

// tDistPDF computes the PDF of the Student's t-distribution at x with df
// degrees of freedom, evaluated in log space for numerical stability.
func tDistPDF(x, df float64) float64 {
	lgNum, _ := math.Lgamma((df + 1) / 2)
	lgDen, _ := math.Lgamma(df / 2)
	logPdf := lgNum - lgDen - 0.5*math.Log(df*math.Pi) - ((df+1)/2)*math.Log(1+x*x/df)
	return math.Exp(logPdf)
}

// ---------------------------------------------------------------------------
// T.INV — Inverse of the Student's t-distribution (left-tailed)
// ---------------------------------------------------------------------------
// T.INV(probability, deg_freedom)
//   probability  – 0 < p < 1 (p=0 and p=1 return #NUM!)
//   deg_freedom  – truncated to integer, must be >= 1

func fnTInv(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	p, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	dfRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	// Truncate deg_freedom to integer.
	df := math.Trunc(dfRaw)
	if df < 1 {
		return ErrorVal(ErrValNUM), nil
	}

	if p <= 0 || p >= 1 {
		return ErrorVal(ErrValNUM), nil
	}

	result, ok := tInv(p, df)
	if !ok {
		return ErrorVal(ErrValNA), nil
	}
	return NumberVal(result), nil
}

// tInv computes the inverse of the Student's t-distribution (left-tailed).
// It returns the t-value such that CDF(t, df) = p, plus a boolean indicating
// convergence. Caller must ensure 0 < p < 1 and df >= 1.
func tInv(p, df float64) (float64, bool) {
	// By symmetry, p = 0.5 always returns 0.
	if p == 0.5 {
		return 0, true
	}

	// Initial guess from the standard normal inverse.
	t := normSInv(p)

	// For small df the normal approximation is poor; clamp the initial
	// guess to keep it in a reasonable range.
	if df <= 2 && math.Abs(t) > 10 {
		if t > 0 {
			t = 10
		} else {
			t = -10
		}
	}

	// Newton-Raphson iteration: t_new = t - (CDF(t) - p) / PDF(t)
	const maxIter = 100
	const tol = 1e-12

	for i := 0; i < maxIter; i++ {
		cdf := tDistCDF(t, df)
		f := cdf - p
		if math.Abs(f) < tol {
			return t, true
		}

		pdf := tDistPDF(t, df)
		if pdf < 1e-300 {
			// PDF too small; fall back to bisection.
			break
		}

		step := f / pdf
		tNew := t - step
		t = tNew
	}

	// Bisection fallback.
	lo := -1000.0
	hi := 1000.0
	// Adjust bounds so that CDF(lo) < p < CDF(hi).
	for tDistCDF(lo, df) > p {
		lo *= 2
	}
	for tDistCDF(hi, df) < p {
		hi *= 2
	}

	for i := 0; i < 200; i++ {
		mid := (lo + hi) / 2
		cdf := tDistCDF(mid, df)
		if math.Abs(cdf-p) < tol {
			return mid, true
		}
		if cdf < p {
			lo = mid
		} else {
			hi = mid
		}
		if (hi - lo) < tol*math.Abs(hi) {
			return (lo + hi) / 2, true
		}
	}

	return 0, false
}

// ---------------------------------------------------------------------------
// T.DIST.RT — Right-tailed Student's t-distribution
// ---------------------------------------------------------------------------
// T.DIST.RT(x, deg_freedom)
//   x           – any numeric value
//   deg_freedom – truncated to integer, must be >= 1
// Returns 1 - T.DIST(x, df, TRUE).

func fnTDistRT(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	dfRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	df := math.Trunc(dfRaw)
	if df < 1 {
		return ErrorVal(ErrValNUM), nil
	}

	return NumberVal(1 - tDistCDF(x, df)), nil
}

// ---------------------------------------------------------------------------
// T.DIST.2T — Two-tailed Student's t-distribution
// ---------------------------------------------------------------------------
// T.DIST.2T(x, deg_freedom)
//   x           – must be >= 0 (if x < 0, return #NUM!)
//   deg_freedom – truncated to integer, must be >= 1
// Returns 2 * (1 - T.DIST(|x|, df, TRUE)).

func fnTDist2T(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	dfRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	if x < 0 {
		return ErrorVal(ErrValNUM), nil
	}

	df := math.Trunc(dfRaw)
	if df < 1 {
		return ErrorVal(ErrValNUM), nil
	}

	return NumberVal(2 * (1 - tDistCDF(x, df))), nil
}

// ---------------------------------------------------------------------------
// T.INV.2T — Two-tailed inverse of Student's t-distribution
// ---------------------------------------------------------------------------
// T.INV.2T(probability, deg_freedom)
//   probability  – 0 < p <= 1 (p <= 0 or p > 1 → #NUM!)
//   deg_freedom  – truncated to integer, must be >= 1
// Returns T.INV(1 - probability/2, df), i.e. the positive t-value such that
// P(|T| >= t) = probability.

func fnTInv2T(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	p, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	dfRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}

	df := math.Trunc(dfRaw)
	if df < 1 {
		return ErrorVal(ErrValNUM), nil
	}

	if p <= 0 || p > 1 {
		return ErrorVal(ErrValNUM), nil
	}

	// T.INV.2T(p, df) = T.INV(1 - p/2, df)
	// For p=1, 1 - p/2 = 0.5, so T.INV(0.5, df) = 0.
	result, ok := tInv(1-p/2, df)
	if !ok {
		return ErrorVal(ErrValNA), nil
	}
	return NumberVal(result), nil
}

// regBetaInc returns the regularized incomplete beta function I_x(a, b).
// It uses Lentz's continued fraction method for efficient evaluation.
func regBetaInc(x, a, b float64) float64 {
	if x <= 0 {
		return 0
	}
	if x >= 1 {
		return 1
	}

	// Use the symmetry relation: I_x(a,b) = 1 - I_{1-x}(b,a)
	// when x > (a+1)/(a+b+2) for better convergence of the CF.
	if x > (a+1)/(a+b+2) {
		return 1 - regBetaInc(1-x, b, a)
	}

	// Log of the front factor: x^a * (1-x)^b / B(a,b)
	lgA, _ := math.Lgamma(a)
	lgB, _ := math.Lgamma(b)
	lgAB, _ := math.Lgamma(a + b)
	lbeta := lgA + lgB - lgAB
	front := math.Exp(a*math.Log(x) + b*math.Log(1-x) - lbeta)

	// Evaluate the continued fraction using the modified Lentz's method.
	// From Numerical Recipes, the CF for I_x(a,b) / front is:
	//   1/(1+ d1/(1+ d2/(1+ ...)))
	// where the coefficients are:
	//   d_{2m+1} = -(a+m)(a+b+m) x / ((a+2m)(a+2m+1))
	//   d_{2m}   = m(b-m) x / ((a+2m-1)(a+2m))
	//
	// This evaluates to: betacf(a,b,x) and I_x(a,b) = front * betacf / a
	return front * betacf(a, b, x) / a
}

// betacf evaluates the continued fraction for the incomplete beta function.
// This implements the algorithm from Numerical Recipes (betacf).
func betacf(a, b, x float64) float64 {
	const eps = 1e-15
	const tiny = 1e-30
	const maxIter = 1000

	qab := a + b
	qap := a + 1
	qam := a - 1

	// Initial setup for modified Lentz's method.
	c := 1.0
	d := 1.0 - qab*x/qap
	if math.Abs(d) < tiny {
		d = tiny
	}
	d = 1.0 / d
	h := d

	for m := 1; m <= maxIter; m++ {
		mf := float64(m)
		m2 := 2.0 * mf

		// Even coefficient: d_{2m} = m(b-m)x / ((a+2m-1)(a+2m))
		aa := mf * (b - mf) * x / ((qam + m2) * (a + m2))
		d = 1.0 + aa*d
		if math.Abs(d) < tiny {
			d = tiny
		}
		c = 1.0 + aa/c
		if math.Abs(c) < tiny {
			c = tiny
		}
		d = 1.0 / d
		h *= d * c

		// Odd coefficient: d_{2m+1} = -(a+m)(a+b+m)x / ((a+2m)(a+2m+1))
		aa = -(a + mf) * (qab + mf) * x / ((a + m2) * (qap + m2))
		d = 1.0 + aa*d
		if math.Abs(d) < tiny {
			d = tiny
		}
		c = 1.0 + aa/c
		if math.Abs(c) < tiny {
			c = tiny
		}
		d = 1.0 / d
		delta := d * c
		h *= delta

		if math.Abs(delta-1.0) < eps {
			break
		}
	}
	return h
}

// ---------------------------------------------------------------------------
// F.DIST — F probability distribution
// ---------------------------------------------------------------------------
// F.DIST(x, deg_freedom1, deg_freedom2, cumulative)
//
//	x            – value at which to evaluate (must be >= 0)
//	deg_freedom1 – numerator degrees of freedom (truncated to integer, >= 1)
//	deg_freedom2 – denominator degrees of freedom (truncated to integer, >= 1)
//	cumulative   – TRUE for CDF, FALSE for PDF

func fnFDist(args []Value) (Value, error) {
	if len(args) != 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	xRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	df1Raw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	df2Raw, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	cum, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}

	if xRaw < 0 {
		return ErrorVal(ErrValNUM), nil
	}

	// Truncate degrees of freedom to integers.
	d1 := math.Trunc(df1Raw)
	d2 := math.Trunc(df2Raw)
	if d1 < 1 || d2 < 1 {
		return ErrorVal(ErrValNUM), nil
	}

	if cum != 0 {
		// CDF: I(d1*x/(d1*x+d2), d1/2, d2/2) using the regularized
		// incomplete beta function.
		if xRaw == 0 {
			return NumberVal(0), nil
		}
		z := d1 * xRaw / (d1*xRaw + d2)
		return NumberVal(regBetaInc(z, d1/2.0, d2/2.0)), nil
	}

	// PDF
	if xRaw == 0 {
		if d1 < 2 {
			// df1 = 1: PDF diverges at x=0.
			return ErrorVal(ErrValNUM), nil
		}
		if d1 == 2 {
			return NumberVal(1), nil
		}
		// df1 > 2: PDF is 0 at x=0.
		return NumberVal(0), nil
	}

	// Use log form for numerical stability:
	// log(f(x)) = 0.5*d1*log(d1) + 0.5*d2*log(d2) +
	//             (0.5*d1-1)*log(x) - 0.5*(d1+d2)*log(d1*x+d2) -
	//             lbeta(d1/2, d2/2)
	lgA, _ := math.Lgamma(d1 / 2)
	lgB, _ := math.Lgamma(d2 / 2)
	lgAB, _ := math.Lgamma((d1 + d2) / 2)
	lb := lgA + lgB - lgAB

	logPdf := 0.5*d1*math.Log(d1) + 0.5*d2*math.Log(d2) +
		(0.5*d1-1)*math.Log(xRaw) - 0.5*(d1+d2)*math.Log(d1*xRaw+d2) - lb
	return NumberVal(math.Exp(logPdf)), nil
}

// ---------------------------------------------------------------------------
// F.INV — Inverse of the F probability distribution
// ---------------------------------------------------------------------------
// F.INV(probability, deg_freedom1, deg_freedom2)
//
//	probability  – 0 <= p <= 1
//	deg_freedom1 – numerator degrees of freedom (truncated to integer, >= 1)
//	deg_freedom2 – denominator degrees of freedom (truncated to integer, >= 1)
//
// Returns x such that F.DIST(x, df1, df2, TRUE) = probability.

func fnFInv(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	p, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	df1Raw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	df2Raw, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	if p < 0 || p > 1 {
		return ErrorVal(ErrValNUM), nil
	}

	// Truncate degrees of freedom to integers.
	d1 := math.Trunc(df1Raw)
	d2 := math.Trunc(df2Raw)
	if d1 < 1 || d2 < 1 {
		return ErrorVal(ErrValNUM), nil
	}

	// Edge cases.
	if p == 0 {
		return NumberVal(0), nil
	}
	if p == 1 {
		return ErrorVal(ErrValNUM), nil
	}

	// fDistCDF computes the CDF of the F-distribution at x:
	//   I(d1*x/(d1*x+d2), d1/2, d2/2)
	fDistCDF := func(x float64) float64 {
		if x <= 0 {
			return 0
		}
		z := d1 * x / (d1*x + d2)
		return regBetaInc(z, d1/2.0, d2/2.0)
	}

	// fDistPDF computes the PDF of the F-distribution at x (log form for stability).
	lgA, _ := math.Lgamma(d1 / 2)
	lgB, _ := math.Lgamma(d2 / 2)
	lgAB, _ := math.Lgamma((d1 + d2) / 2)
	lb := lgA + lgB - lgAB

	fDistPDF := func(x float64) float64 {
		if x <= 0 {
			return 0
		}
		logPdf := 0.5*d1*math.Log(d1) + 0.5*d2*math.Log(d2) +
			(0.5*d1-1)*math.Log(x) - 0.5*(d1+d2)*math.Log(d1*x+d2) - lb
		return math.Exp(logPdf)
	}

	// Initial guess: use the mean of the F-distribution (d2/(d2-2)) scaled by p,
	// or a simple heuristic for a starting point.
	var x float64
	if d2 > 2 {
		mean := d2 / (d2 - 2)
		// Use the normal inverse to adjust the initial guess.
		z := normSInv(p)
		// Approximate: x ≈ mean * exp(z * sqrt(2/d1))
		x = mean * math.Exp(z*math.Sqrt(2/d1))
		if x <= 0 {
			x = 0.001
		}
	} else {
		// For small d2, start with 1.0 and adjust.
		x = 1.0
		if p < 0.5 {
			x = 0.5 * p
			if x < 0.001 {
				x = 0.001
			}
		} else if p > 0.9 {
			x = 10.0
		}
	}

	// Newton-Raphson iteration.
	const maxIter = 100
	const tol = 1e-12

	for i := 0; i < maxIter; i++ {
		cdf := fDistCDF(x)
		f := cdf - p

		if math.Abs(f) < tol {
			return NumberVal(x), nil
		}

		pdf := fDistPDF(x)
		if pdf < 1e-300 {
			break
		}

		step := f / pdf
		xNew := x - step
		// Ensure x stays positive.
		if xNew <= 0 {
			x = x / 2
		} else {
			x = xNew
		}
	}

	// Bisection fallback.
	lo := 0.0
	hi := x
	if hi <= 0 {
		hi = 1.0
	}
	// Expand hi until CDF(hi) > p.
	for fDistCDF(hi) < p {
		hi *= 2
		if hi > 1e100 {
			return ErrorVal(ErrValNUM), nil
		}
	}

	for i := 0; i < 200; i++ {
		mid := (lo + hi) / 2
		cdf := fDistCDF(mid)
		if math.Abs(cdf-p) < tol {
			return NumberVal(mid), nil
		}
		if cdf < p {
			lo = mid
		} else {
			hi = mid
		}
		if (hi - lo) < tol*math.Abs(hi) {
			return NumberVal((lo + hi) / 2), nil
		}
	}

	return ErrorVal(ErrValNA), nil
}

// ---------------------------------------------------------------------------
// F.DIST.RT — Right-tailed F-distribution
// ---------------------------------------------------------------------------
// F.DIST.RT(x, deg_freedom1, deg_freedom2)
//   x            – must be >= 0 (if x < 0, #NUM!)
//   deg_freedom1 – numerator degrees of freedom (truncated to integer, >= 1)
//   deg_freedom2 – denominator degrees of freedom (truncated to integer, >= 1)
// Returns 1 - F.DIST(x, df1, df2, TRUE).

func fnFDistRT(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	xRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	df1Raw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	df2Raw, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	if xRaw < 0 {
		return ErrorVal(ErrValNUM), nil
	}

	d1 := math.Trunc(df1Raw)
	d2 := math.Trunc(df2Raw)
	if d1 < 1 || d2 < 1 {
		return ErrorVal(ErrValNUM), nil
	}

	if xRaw == 0 {
		return NumberVal(1), nil
	}
	z := d1 * xRaw / (d1*xRaw + d2)
	return NumberVal(1 - regBetaInc(z, d1/2.0, d2/2.0)), nil
}

// ---------------------------------------------------------------------------
// F.INV.RT — Right-tailed inverse of the F probability distribution
// ---------------------------------------------------------------------------
// F.INV.RT(probability, deg_freedom1, deg_freedom2)
//   probability  – 0 <= p <= 1
//   deg_freedom1 – numerator degrees of freedom (truncated to integer, >= 1)
//   deg_freedom2 – denominator degrees of freedom (truncated to integer, >= 1)
// Returns F.INV(1 - probability, df1, df2).

func fnFInvRT(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	p, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	df1Raw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	df2Raw, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	if p < 0 || p > 1 {
		return ErrorVal(ErrValNUM), nil
	}

	d1 := math.Trunc(df1Raw)
	d2 := math.Trunc(df2Raw)
	if d1 < 1 || d2 < 1 {
		return ErrorVal(ErrValNUM), nil
	}

	// Edge cases.
	if p == 1 {
		return NumberVal(0), nil
	}
	if p == 0 {
		return ErrorVal(ErrValNUM), nil
	}

	// Delegate to F.INV(1-p, df1, df2).
	return fnFInv([]Value{NumberVal(1 - p), NumberVal(d1), NumberVal(d2)})
}

// ---------------------------------------------------------------------------
// CONFIDENCE.NORM — Confidence interval for a population mean (normal dist)
// ---------------------------------------------------------------------------
// CONFIDENCE.NORM(alpha, standard_dev, size)
//
//	alpha        – significance level, 0 < alpha < 1
//	standard_dev – population standard deviation, must be > 0
//	size         – sample size, truncated to integer, must be >= 1
//
// Returns NORM.S.INV(1 - alpha/2) * standard_dev / SQRT(size).

func fnConfidenceNorm(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	alpha, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	stddev, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	sizeRaw, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	size := math.Trunc(sizeRaw)
	if alpha <= 0 || alpha >= 1 {
		return ErrorVal(ErrValNUM), nil
	}
	if stddev <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if size < 1 {
		return ErrorVal(ErrValNUM), nil
	}

	z := normSInv(1 - alpha/2)
	return NumberVal(z * stddev / math.Sqrt(size)), nil
}

// ---------------------------------------------------------------------------
// CONFIDENCE.T — Confidence interval for a population mean (t-distribution)
// ---------------------------------------------------------------------------
// CONFIDENCE.T(alpha, standard_dev, size)
//
//	alpha        – significance level, 0 < alpha < 1
//	standard_dev – population standard deviation, must be > 0
//	size         – sample size, truncated to integer, must be >= 1
//	              (size = 1 returns #DIV/0! because df = 0)
//
// Returns T.INV(1 - alpha/2, size-1) * standard_dev / SQRT(size).

func fnConfidenceT(args []Value) (Value, error) {
	if len(args) != 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	alpha, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	stddev, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	sizeRaw, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	size := math.Trunc(sizeRaw)
	if alpha <= 0 || alpha >= 1 {
		return ErrorVal(ErrValNUM), nil
	}
	if stddev <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if size < 1 {
		return ErrorVal(ErrValNUM), nil
	}
	if size == 1 {
		return ErrorVal(ErrValDIV0), nil
	}

	df := size - 1
	t, ok := tInv(1-alpha/2, df)
	if !ok {
		return ErrorVal(ErrValNA), nil
	}
	return NumberVal(t * stddev / math.Sqrt(size)), nil
}

// ---------------------------------------------------------------------------
// BETA.DIST — Beta distribution (CDF or PDF)
// ---------------------------------------------------------------------------
// BETA.DIST(x, alpha, beta, cumulative, [A], [B])
//
//	x          – value at which to evaluate (must be between A and B)
//	alpha      – first shape parameter (> 0)
//	beta       – second shape parameter (> 0)
//	cumulative – TRUE for CDF, FALSE for PDF
//	A          – optional lower bound (default 0)
//	B          – optional upper bound (default 1)

func fnBetaDist(args []Value) (Value, error) {
	if len(args) < 4 || len(args) > 6 {
		return ErrorVal(ErrValVALUE), nil
	}

	x, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	alpha, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	beta, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	cum, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}

	a := 0.0 // lower bound
	b := 1.0 // upper bound
	if len(args) >= 5 {
		a, e = CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
	}
	if len(args) >= 6 {
		b, e = CoerceNum(args[5])
		if e != nil {
			return *e, nil
		}
	}

	// Validate parameters.
	if alpha <= 0 || beta <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if a == b {
		return ErrorVal(ErrValNUM), nil
	}
	if x < a || x > b {
		return ErrorVal(ErrValNUM), nil
	}

	// Transform to standard [0,1] range.
	z := (x - a) / (b - a)

	if cum != 0 {
		// CDF: regularized incomplete beta function I_z(alpha, beta).
		return NumberVal(regBetaInc(z, alpha, beta)), nil
	}

	// PDF: z^(alpha-1) * (1-z)^(beta-1) / (B(alpha,beta) * (b-a))
	// Handle boundary cases.
	if z == 0 {
		if alpha < 1 {
			// PDF diverges.
			return ErrorVal(ErrValNUM), nil
		}
		if alpha == 1 {
			// PDF = (1-0)^(beta-1) / (B(1,beta) * (b-a)) = 1 / (B(1,beta) * (b-a))
			// B(1, beta) = 1/beta, so PDF = beta / (b-a)
			lgA, _ := math.Lgamma(alpha)
			lgB, _ := math.Lgamma(beta)
			lgAB, _ := math.Lgamma(alpha + beta)
			lb := lgA + lgB - lgAB
			logPdf := (beta-1)*math.Log(1) - lb - math.Log(b-a)
			return NumberVal(math.Exp(logPdf)), nil
		}
		// alpha > 1: PDF = 0
		return NumberVal(0), nil
	}
	if z == 1 {
		if beta < 1 {
			// PDF diverges.
			return ErrorVal(ErrValNUM), nil
		}
		if beta == 1 {
			lgA, _ := math.Lgamma(alpha)
			lgB, _ := math.Lgamma(beta)
			lgAB, _ := math.Lgamma(alpha + beta)
			lb := lgA + lgB - lgAB
			logPdf := (alpha-1)*math.Log(1) - lb - math.Log(b-a)
			return NumberVal(math.Exp(logPdf)), nil
		}
		// beta > 1: PDF = 0
		return NumberVal(0), nil
	}

	lgA, _ := math.Lgamma(alpha)
	lgB, _ := math.Lgamma(beta)
	lgAB, _ := math.Lgamma(alpha + beta)
	lb := lgA + lgB - lgAB

	logPdf := (alpha-1)*math.Log(z) + (beta-1)*math.Log(1-z) - lb - math.Log(b-a)
	return NumberVal(math.Exp(logPdf)), nil
}

// ---------------------------------------------------------------------------
// BETA.INV — Inverse of the beta cumulative distribution function
// ---------------------------------------------------------------------------
// BETA.INV(probability, alpha, beta, [A], [B])
//
//	probability – value at which to evaluate the inverse (0 < p <= 1;
//	              p=0 returns A, p<=0 or p>1 ⇒ #NUM!)
//	alpha       – first shape parameter (> 0)
//	beta        – second shape parameter (> 0)
//	A           – optional lower bound (default 0)
//	B           – optional upper bound (default 1)
//
// Returns x such that BETA.DIST(x, alpha, beta, TRUE, A, B) = probability.

func fnBetaInv(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}

	p, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	alpha, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	bt, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	a := 0.0 // lower bound
	b := 1.0 // upper bound
	if len(args) >= 4 {
		a, e = CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
	}
	if len(args) >= 5 {
		b, e = CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
	}

	// Validate parameters.
	if alpha <= 0 || bt <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if p <= 0 || p > 1 {
		// probability <= 0 or > 1 ⇒ #NUM!
		// But p == 0 returns A in practice.
		if p == 0 {
			return NumberVal(a), nil
		}
		return ErrorVal(ErrValNUM), nil
	}
	if a >= b {
		return ErrorVal(ErrValNUM), nil
	}

	// p == 1 → return B
	if p == 1 {
		return NumberVal(b), nil
	}

	// Find z in [0,1] such that regBetaInc(z, alpha, bt) = p,
	// then transform back: x = a + z*(b-a).

	// Initial guess: use the mean of the beta distribution as starting point.
	z := alpha / (alpha + bt)

	// Newton-Raphson iteration.
	// f(z)  = regBetaInc(z, alpha, bt) - p
	// f'(z) = betaPDF(z) = z^(alpha-1) * (1-z)^(bt-1) / B(alpha,bt)
	lgA, _ := math.Lgamma(alpha)
	lgB, _ := math.Lgamma(bt)
	lgAB, _ := math.Lgamma(alpha + bt)
	lbeta := lgA + lgB - lgAB

	const maxIter = 200
	const tol = 1e-12

	for i := 0; i < maxIter; i++ {
		cdf := regBetaInc(z, alpha, bt)
		f := cdf - p

		if math.Abs(f) < tol {
			return NumberVal(a + z*(b-a)), nil
		}

		// Beta PDF as derivative of the CDF.
		if z <= 0 || z >= 1 {
			break // can't compute PDF at boundary; fall to bisection
		}
		logPdf := (alpha-1)*math.Log(z) + (bt-1)*math.Log(1-z) - lbeta
		pdf := math.Exp(logPdf)

		if pdf < 1e-300 {
			break // PDF too small for Newton step; bisection fallback
		}

		step := f / pdf
		zNew := z - step

		// Keep z strictly in (0,1).
		if zNew <= 0 {
			z = z / 2
		} else if zNew >= 1 {
			z = (z + 1) / 2
		} else {
			z = zNew
		}
	}

	// Bisection fallback on [0, 1].
	lo := 0.0
	hi := 1.0

	for i := 0; i < maxIter; i++ {
		mid := (lo + hi) / 2
		cdf := regBetaInc(mid, alpha, bt)
		if math.Abs(cdf-p) < tol {
			return NumberVal(a + mid*(b-a)), nil
		}
		if cdf < p {
			lo = mid
		} else {
			hi = mid
		}
		if (hi - lo) < tol {
			return NumberVal(a + (lo+hi)/2*(b-a)), nil
		}
	}

	return NumberVal(a + (lo+hi)/2*(b-a)), nil
}

// ---------------------------------------------------------------------------
// HYPGEOM.DIST — Hypergeometric distribution (PMF or CDF)
// ---------------------------------------------------------------------------

func fnHypgeomDist(args []Value) (Value, error) {
	if len(args) != 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	sampleSF, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	numberSampleF, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	populationSF, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	numberPopF, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}
	cumF, e := CoerceNum(args[4])
	if e != nil {
		return *e, nil
	}

	// Truncate all to integers.
	k := int(sampleSF)        // sample_s
	n := int(numberSampleF)   // number_sample
	bigM := int(populationSF) // population_s (M)
	bigN := int(numberPopF)   // number_pop (N)

	// Validate constraints.
	if bigN <= 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if n <= 0 || n > bigN {
		return ErrorVal(ErrValNUM), nil
	}
	if bigM <= 0 || bigM > bigN {
		return ErrorVal(ErrValNUM), nil
	}
	// k must be >= 0 and <= min(n, M)
	minNM := n
	if bigM < minNM {
		minNM = bigM
	}
	if k < 0 || k > minNM {
		return ErrorVal(ErrValNUM), nil
	}
	// k must be >= max(0, n + M - N)
	lowerBound := 0
	if n+bigM-bigN > 0 {
		lowerBound = n + bigM - bigN
	}
	if k < lowerBound {
		return ErrorVal(ErrValNUM), nil
	}

	if cumF != 0 {
		// CDF: sum PMF from lowerBound to k.
		sum := 0.0
		for i := lowerBound; i <= k; i++ {
			sum += hypgeomPMF(i, n, bigM, bigN)
		}
		return NumberVal(sum), nil
	}
	return NumberVal(hypgeomPMF(k, n, bigM, bigN)), nil
}

// hypgeomPMF returns the hypergeometric probability mass function:
// P(X=k) = C(M,k) * C(N-M, n-k) / C(N, n)
// Uses log-gamma for numerical stability with large values.
func hypgeomPMF(k, n, bigM, bigN int) float64 {
	// log(C(M,k))
	lgM1, _ := math.Lgamma(float64(bigM + 1))
	lgK1, _ := math.Lgamma(float64(k + 1))
	lgMK1, _ := math.Lgamma(float64(bigM - k + 1))

	// log(C(N-M, n-k))
	lgNM1, _ := math.Lgamma(float64(bigN - bigM + 1))
	lgNK1, _ := math.Lgamma(float64(n - k + 1))
	lgNMNK1, _ := math.Lgamma(float64(bigN - bigM - n + k + 1))

	// log(C(N, n))
	lgN1, _ := math.Lgamma(float64(bigN + 1))
	lgn1, _ := math.Lgamma(float64(n + 1))
	lgNn1, _ := math.Lgamma(float64(bigN - n + 1))

	logP := (lgM1 - lgK1 - lgMK1) + (lgNM1 - lgNK1 - lgNMNK1) - (lgN1 - lgn1 - lgNn1)
	return math.Exp(logP)
}

// fnNegbinomDist implements NEGBINOM.DIST(number_f, number_s, probability_s, cumulative).
// Returns the negative binomial distribution — the probability of number_f failures
// before the number_s-th success, with probability_s chance of success on each trial.
func fnNegbinomDist(args []Value) (Value, error) {
	if len(args) != 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	ff, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	rf, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	p, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	cum, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}

	// Truncate to integers.
	f := int(ff)
	r := int(rf)

	if f < 0 {
		return ErrorVal(ErrValNUM), nil
	}
	if r < 1 {
		return ErrorVal(ErrValNUM), nil
	}
	if p < 0 || p > 1 {
		return ErrorVal(ErrValNUM), nil
	}

	if cum != 0 {
		// CDF: use the regularized incomplete beta function.
		// CDF of NegBinom(f; r, p) = I_p(r, f+1)
		return NumberVal(regBetaInc(p, float64(r), float64(f+1))), nil
	}

	// PMF: P(X=f) = C(f+r-1, r-1) * p^r * (1-p)^f
	return NumberVal(negbinomPMF(f, r, p)), nil
}

// negbinomPMF returns the negative binomial PMF:
// P(X=f) = C(f+r-1, r-1) * p^r * (1-p)^f
// Uses log-gamma for numerical stability.
func negbinomPMF(f, r int, p float64) float64 {
	if p == 0 {
		if f == 0 {
			return 1
		}
		return 0
	}
	if p == 1 {
		if f == 0 {
			return 1
		}
		return 0
	}
	// log(C(f+r-1, r-1)) = lgamma(f+r) - lgamma(r) - lgamma(f+1)
	lgFR, _ := math.Lgamma(float64(f + r))
	lgR, _ := math.Lgamma(float64(r))
	lgF1, _ := math.Lgamma(float64(f + 1))
	logC := lgFR - lgR - lgF1
	logProb := logC + float64(r)*math.Log(p) + float64(f)*math.Log(1-p)
	return math.Exp(logProb)
}

// fnPROB implements the PROB function.
// PROB(x_range, prob_range, lower_limit, [upper_limit])
// Returns the probability that values in x_range fall between lower_limit and
// upper_limit (inclusive). If upper_limit is omitted, returns the probability
// of being exactly equal to lower_limit.
func fnPROB(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Flatten x_range and prob_range into 1-D slices.
	flatX := flattenValuesGeneric(args[0])
	flatP := flattenValuesGeneric(args[1])

	// x_range and prob_range must have the same number of data points.
	if len(flatX) != len(flatP) {
		return ErrorVal(ErrValNA), nil
	}

	// Both ranges must be non-empty.
	if len(flatX) == 0 {
		return ErrorVal(ErrValNA), nil
	}

	// Extract numeric values, propagating errors.
	xs := make([]float64, 0, len(flatX))
	ps := make([]float64, 0, len(flatP))
	probSum := 0.0
	for i := range flatX {
		xv, yv := flatX[i], flatP[i]
		if xv.Type == ValueError {
			return xv, nil
		}
		if yv.Type == ValueError {
			return yv, nil
		}
		// Both must be numeric.
		if xv.Type != ValueNumber || yv.Type != ValueNumber {
			return ErrorVal(ErrValVALUE), nil
		}
		xs = append(xs, xv.Num)
		ps = append(ps, yv.Num)
		probSum += yv.Num
	}

	// Sum of probabilities must equal 1 (with floating-point tolerance).
	if math.Abs(probSum-1.0) > 1e-10 {
		return ErrorVal(ErrValNUM), nil
	}

	// Parse lower_limit.
	lower, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}

	// Parse optional upper_limit; if omitted, match exactly lower_limit.
	upper := lower
	if len(args) == 4 {
		u, e2 := CoerceNum(args[3])
		if e2 != nil {
			return *e2, nil
		}
		upper = u
	}

	// Sum probabilities for x values in [lower, upper].
	result := 0.0
	for i, x := range xs {
		if x >= lower && x <= upper {
			result += ps[i]
		}
	}

	return NumberVal(result), nil
}

// aggregateFuncNames maps AGGREGATE function_num (1-19) to the sub-function
// name. Indices 1-13 are reference-form functions; 14-19 are array-form
// functions that require a k/quart argument.
var aggregateFuncNames = [20]string{
	1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN",
	6: "PRODUCT", 7: "STDEV.S", 8: "STDEV.P", 9: "SUM", 10: "VAR.S", 11: "VAR.P",
	12: "MEDIAN", 13: "MODE.SNGL", 14: "LARGE", 15: "SMALL",
	16: "PERCENTILE.INC", 17: "QUARTILE.INC", 18: "PERCENTILE.EXC", 19: "QUARTILE.EXC",
}

// fnAggregate implements the AGGREGATE function.
// AGGREGATE(function_num, options, ref1, [ref2], ...) — reference form
// AGGREGATE(function_num, options, array, [k])          — array form
//
// It applies one of 19 aggregate functions to the supplied data, with the
// ability to ignore error values (hidden-row options are accepted but treated
// as no-ops because the formula engine has no UI state).
func fnAggregate(args []Value) (Value, error) {
	if len(args) < 3 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Parse function_num (must be 1-19).
	fnRaw, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	fnNum := int(fnRaw)
	if fnNum < 1 || fnNum > 19 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Parse options (must be 0-7).
	optRaw, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	opt := int(optRaw)
	if opt < 0 || opt > 7 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Determine whether to ignore error values (options 2, 3, 6, 7).
	ignoreErrors := opt == 2 || opt == 3 || opt == 6 || opt == 7

	// Functions 14-19 require a k/quart argument (the last positional arg).
	needsK := fnNum >= 14

	dataArgs := args[2:]
	var kArg Value
	if needsK {
		if len(dataArgs) < 2 {
			return ErrorVal(ErrValVALUE), nil
		}
		kArg = dataArgs[len(dataArgs)-1]
		dataArgs = dataArgs[:len(dataArgs)-1]
	}

	// Flatten all data arguments into a single list of values, optionally
	// filtering out error values.
	var vals []Value
	for _, arg := range dataArgs {
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						if ignoreErrors {
							continue
						}
						return cell, nil
					}
					vals = append(vals, cell)
				}
			}
		} else {
			if arg.Type == ValueError {
				if ignoreErrors {
					continue
				}
				return arg, nil
			}
			vals = append(vals, arg)
		}
	}

	// Build a single-row ValueArray from the filtered values.
	arr := Value{Type: ValueArray, Array: [][]Value{vals}}

	// Look up and call the sub-function.
	name := aggregateFuncNames[fnNum]
	id := LookupFunc(name)
	if id < 0 {
		return ErrorVal(ErrValVALUE), nil
	}

	var callArgs []Value
	if needsK {
		callArgs = []Value{arr, kArg}
	} else {
		callArgs = []Value{arr}
	}
	return CallFunc(id, callArgs, nil)
}

// ---------------------------------------------------------------------------
// T.TEST — Student's t-Test
// ---------------------------------------------------------------------------
// T.TEST(array1, array2, tails, type)
//   array1 – first data set
//   array2 – second data set
//   tails  – 1 = one-tailed, 2 = two-tailed
//   type   – 1 = Paired, 2 = Two-sample equal variance (homoscedastic),
//            3 = Two-sample unequal variance (heteroscedastic)

func fnTTest(args []Value) (Value, error) {
	if len(args) != 4 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Extract tails and type, coerce to numeric.
	tailsRaw, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	typeRaw, e := CoerceNum(args[3])
	if e != nil {
		return *e, nil
	}

	// Truncate to integers.
	tails := int(math.Trunc(tailsRaw))
	ttype := int(math.Trunc(typeRaw))

	// Validate tails and type.
	if tails != 1 && tails != 2 {
		return ErrorVal(ErrValNUM), nil
	}
	if ttype < 1 || ttype > 3 {
		return ErrorVal(ErrValNUM), nil
	}

	var tStat, df float64

	switch ttype {
	case 1: // Paired t-test.
		// For paired tests, we must iterate both arrays in lockstep and
		// only include pairs where BOTH values are numeric.
		flat1 := flattenValues(args[0])
		flat2 := flattenValues(args[1])
		// Check for errors in either array.
		for i := range flat1 {
			if flat1[i].Type == ValueError {
				return flat1[i], nil
			}
		}
		for i := range flat2 {
			if flat2[i].Type == ValueError {
				return flat2[i], nil
			}
		}
		if len(flat1) != len(flat2) {
			return ErrorVal(ErrValNA), nil
		}
		// Collect paired numeric values.
		var nums1, nums2 []float64
		for i := range flat1 {
			if flat1[i].Type == ValueNumber && flat2[i].Type == ValueNumber {
				nums1 = append(nums1, flat1[i].Num)
				nums2 = append(nums2, flat2[i].Num)
			}
		}
		n := len(nums1)
		if n < 2 {
			return ErrorVal(ErrValDIV0), nil
		}
		// Compute differences and their mean/stdev.
		diffs := make([]float64, n)
		sumD := 0.0
		for i := 0; i < n; i++ {
			diffs[i] = nums1[i] - nums2[i]
			sumD += diffs[i]
		}
		meanD := sumD / float64(n)
		ssq := 0.0
		for _, d := range diffs {
			ssq += (d - meanD) * (d - meanD)
		}
		sdD := math.Sqrt(ssq / float64(n-1))
		if sdD == 0 {
			return ErrorVal(ErrValDIV0), nil
		}
		tStat = meanD / (sdD / math.Sqrt(float64(n)))
		df = float64(n - 1)

	case 2, 3: // Two-sample tests: collect numeric values independently.
		nums1, ev := collectNumeric(args[:1])
		if ev != nil {
			return *ev, nil
		}
		nums2, ev := collectNumeric(args[1:2])
		if ev != nil {
			return *ev, nil
		}
		n1 := len(nums1)
		n2 := len(nums2)

		if ttype == 2 { // Equal variance (homoscedastic).
			if n1 < 2 || n2 < 2 {
				return ErrorVal(ErrValDIV0), nil
			}
			mean1 := mean(nums1)
			mean2 := mean(nums2)
			ssq1 := sumSqDev(nums1, mean1)
			ssq2 := sumSqDev(nums2, mean2)
			// Pooled variance.
			sp2 := (ssq1 + ssq2) / float64(n1+n2-2)
			denom := math.Sqrt(sp2 * (1.0/float64(n1) + 1.0/float64(n2)))
			if denom == 0 {
				return ErrorVal(ErrValDIV0), nil
			}
			tStat = (mean1 - mean2) / denom
			df = float64(n1 + n2 - 2)
		} else { // Unequal variance (Welch's t-test).
			if n1 < 2 || n2 < 2 {
				return ErrorVal(ErrValDIV0), nil
			}
			mean1 := mean(nums1)
			mean2 := mean(nums2)
			var1 := sumSqDev(nums1, mean1) / float64(n1-1)
			var2 := sumSqDev(nums2, mean2) / float64(n2-1)
			se := math.Sqrt(var1/float64(n1) + var2/float64(n2))
			if se == 0 {
				return ErrorVal(ErrValDIV0), nil
			}
			tStat = (mean1 - mean2) / se
			// Welch-Satterthwaite degrees of freedom.
			vn1 := var1 / float64(n1)
			vn2 := var2 / float64(n2)
			num := (vn1 + vn2) * (vn1 + vn2)
			den := vn1*vn1/float64(n1-1) + vn2*vn2/float64(n2-1)
			df = num / den
		}
	}

	// Compute p-value from the t-distribution.
	// p = P(|T| >= |t|) for two-tailed, or P(T >= |t|) for one-tailed.
	// Using: P(T >= |t|) = 1 - tDistCDF(|t|, df) = 0.5 * I_x(df/2, 0.5)
	// where x = df / (df + t²).
	absT := math.Abs(tStat)
	bx := df / (df + absT*absT)
	p := 0.5 * regBetaInc(bx, df/2, 0.5)

	if tails == 2 {
		p = 2 * p
	}

	return NumberVal(p), nil
}

// mean computes the arithmetic mean of a slice of floats.
func mean(vals []float64) float64 {
	s := 0.0
	for _, v := range vals {
		s += v
	}
	return s / float64(len(vals))
}

// sumSqDev computes the sum of squared deviations from the mean.
func sumSqDev(vals []float64, m float64) float64 {
	s := 0.0
	for _, v := range vals {
		d := v - m
		s += d * d
	}
	return s
}

// fnZTEST implements Z.TEST(array, x, [sigma]).
// Returns the one-tailed p-value of a z-test.
func fnZTEST(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	// Collect numeric values from the array (first arg only).
	nums, ev := collectNumeric(args[:1])
	if ev != nil {
		return *ev, nil
	}
	n := len(nums)
	if n == 0 {
		return ErrorVal(ErrValNA), nil
	}
	x, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	// Compute mean.
	sum := 0.0
	for _, v := range nums {
		sum += v
	}
	mean := sum / float64(n)

	var sigma float64
	if len(args) == 3 {
		sigma, e = CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
	} else {
		// Sample standard deviation.
		if n < 2 {
			return ErrorVal(ErrValDIV0), nil
		}
		ssq := 0.0
		for _, v := range nums {
			d := v - mean
			ssq += d * d
		}
		sigma = math.Sqrt(ssq / float64(n-1))
	}
	if sigma == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	z := (mean - x) / (sigma / math.Sqrt(float64(n)))
	// Use Erfc directly to avoid precision loss from 1 - CDF when CDF ≈ 1.
	return NumberVal(0.5 * math.Erfc(z/math.Sqrt(2))), nil
}

// ---------------------------------------------------------------------------
// F.TEST — F-test (two-tailed probability that variances are not different)
// ---------------------------------------------------------------------------
// F.TEST(array1, array2)
//
//	array1 – first array or range of data
//	array2 – second array or range of data
//
// Returns the two-tailed probability of the F-distribution. Text, logical
// values, and empty cells in the arrays are ignored; zeros are included.
// If either array has fewer than 2 numeric data points or either variance
// is zero, returns #DIV/0!.

func fnFTest(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	nums1, ev := collectNumeric(args[:1])
	if ev != nil {
		return *ev, nil
	}
	nums2, ev := collectNumeric(args[1:2])
	if ev != nil {
		return *ev, nil
	}

	n1 := len(nums1)
	n2 := len(nums2)
	if n1 < 2 || n2 < 2 {
		return ErrorVal(ErrValDIV0), nil
	}

	// Compute sample variances (with n-1 denominator).
	mean1 := mean(nums1)
	mean2 := mean(nums2)
	var1 := sumSqDev(nums1, mean1) / float64(n1-1)
	var2 := sumSqDev(nums2, mean2) / float64(n2-1)

	if var1 == 0 || var2 == 0 {
		return ErrorVal(ErrValDIV0), nil
	}

	// F-statistic: var1/var2 with df1=n1-1, df2=n2-1.
	f := var1 / var2
	df1 := float64(n1 - 1)
	df2 := float64(n2 - 1)

	// CDF of the F-distribution at f: I(d1*f/(d1*f+d2), d1/2, d2/2).
	z := df1 * f / (df1*f + df2)
	cdf := regBetaInc(z, df1/2.0, df2/2.0)

	// Two-tailed p-value.
	tail := cdf
	if 1-cdf < tail {
		tail = 1 - cdf
	}
	p := 2 * tail
	if p > 1 {
		p = 1
	}

	return NumberVal(p), nil
}

// fnGROWTH implements GROWTH(known_y's, [known_x's], [new_x's], [const]).
// It returns predicted y-values for new x-values based on an exponential
// regression model y = b * m^x (equivalently, linear regression on ln(y) vs x).
func fnGROWTH(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}

	// ---- Helper: extract a flat slice of float64 from a Value (array or scalar). ----
	flattenNums := func(v Value) ([]float64, *Value) {
		if v.Type == ValueArray {
			var nums []float64
			for _, row := range v.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return nil, &cell
					}
					n, e := CoerceNum(cell)
					if e != nil {
						return nil, e
					}
					nums = append(nums, n)
				}
			}
			return nums, nil
		}
		if v.Type == ValueError {
			return nil, &v
		}
		n, e := CoerceNum(v)
		if e != nil {
			return nil, e
		}
		return []float64{n}, nil
	}

	// ---- arrayDims returns (rows, cols) for an array, or (1,1) for scalar. ----
	arrayDims := func(v Value) (int, int) {
		if v.Type == ValueArray {
			rows := len(v.Array)
			if rows == 0 {
				return 0, 0
			}
			return rows, len(v.Array[0])
		}
		return 1, 1
	}

	// ---- 1. known_y's ----
	knownY, ev := flattenNums(args[0])
	if ev != nil {
		return *ev, nil
	}
	n := len(knownY)
	if n == 0 {
		return ErrorVal(ErrValNA), nil
	}
	// All known_y must be > 0.
	for _, y := range knownY {
		if y <= 0 {
			return ErrorVal(ErrValNUM), nil
		}
	}

	// ---- 2. known_x's ----
	var knownX []float64
	if len(args) >= 2 && args[1].Type != ValueEmpty {
		knownX, ev = flattenNums(args[1])
		if ev != nil {
			return *ev, nil
		}
		if len(knownX) != n {
			return ErrorVal(ErrValREF), nil
		}
	} else {
		knownX = make([]float64, n)
		for i := range knownX {
			knownX[i] = float64(i + 1)
		}
	}

	// ---- 3. new_x's ----
	var newX []float64
	newXRows, newXCols := 1, 1
	isNewXScalar := true
	if len(args) >= 3 && args[2].Type != ValueEmpty {
		newX, ev = flattenNums(args[2])
		if ev != nil {
			return *ev, nil
		}
		newXRows, newXCols = arrayDims(args[2])
		isNewXScalar = args[2].Type != ValueArray
	} else {
		newX = make([]float64, len(knownX))
		copy(newX, knownX)
		// Shape matches known_x's shape.
		if len(args) >= 2 && args[1].Type != ValueEmpty {
			newXRows, newXCols = arrayDims(args[1])
			isNewXScalar = args[1].Type != ValueArray
		} else {
			newXRows, newXCols = arrayDims(args[0])
			isNewXScalar = args[0].Type != ValueArray
		}
	}

	// ---- 4. const ----
	useConst := true
	if len(args) >= 4 && args[3].Type != ValueEmpty {
		cv, e := CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
		useConst = cv != 0
	}

	// ---- 5. Compute ln(y) ----
	lnY := make([]float64, n)
	for i, y := range knownY {
		lnY[i] = math.Log(y)
	}

	// ---- 6. Linear regression of lnY on knownX ----
	var slope, intercept float64
	if useConst {
		// Compute means.
		sumX, sumLnY := 0.0, 0.0
		for i := 0; i < n; i++ {
			sumX += knownX[i]
			sumLnY += lnY[i]
		}
		meanX := sumX / float64(n)
		meanLnY := sumLnY / float64(n)

		// Compute slope and intercept.
		cov := 0.0
		ssqX := 0.0
		for i := 0; i < n; i++ {
			dx := knownX[i] - meanX
			cov += dx * (lnY[i] - meanLnY)
			ssqX += dx * dx
		}

		if ssqX == 0 {
			// All x values are the same. slope=0, intercept=meanLnY.
			slope = 0
			intercept = meanLnY
		} else {
			slope = cov / ssqX
			intercept = meanLnY - slope*meanX
		}
	} else {
		// const=FALSE: force intercept=0, so lnY = slope*x.
		// slope = Σ(x*lnY) / Σ(x²)
		sumXLnY := 0.0
		sumXX := 0.0
		for i := 0; i < n; i++ {
			sumXLnY += knownX[i] * lnY[i]
			sumXX += knownX[i] * knownX[i]
		}
		if sumXX == 0 {
			return ErrorVal(ErrValDIV0), nil
		}
		slope = sumXLnY / sumXX
		intercept = 0
	}

	// ---- 7. Predict ----
	predicted := make([]float64, len(newX))
	for i, x := range newX {
		predicted[i] = math.Exp(intercept + slope*x)
	}

	// ---- 8. Return result ----
	if isNewXScalar && len(predicted) == 1 {
		return NumberVal(predicted[0]), nil
	}

	// Build array matching the shape of new_x's.
	result := make([][]Value, newXRows)
	idx := 0
	for r := 0; r < newXRows; r++ {
		row := make([]Value, newXCols)
		for c := 0; c < newXCols; c++ {
			if idx < len(predicted) {
				row[c] = NumberVal(predicted[idx])
				idx++
			}
		}
		result[r] = row
	}
	return Value{Type: ValueArray, Array: result}, nil
}

// fnTREND implements TREND(known_y's, [known_x's], [new_x's], [const]).
// It returns predicted y-values for new x-values based on a linear
// regression model y = slope*x + intercept.
func fnTREND(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}

	// ---- Helper: extract a flat slice of float64 from a Value (array or scalar). ----
	flattenNums := func(v Value) ([]float64, *Value) {
		if v.Type == ValueArray {
			var nums []float64
			for _, row := range v.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return nil, &cell
					}
					n, e := CoerceNum(cell)
					if e != nil {
						return nil, e
					}
					nums = append(nums, n)
				}
			}
			return nums, nil
		}
		if v.Type == ValueError {
			return nil, &v
		}
		n, e := CoerceNum(v)
		if e != nil {
			return nil, e
		}
		return []float64{n}, nil
	}

	// ---- arrayDims returns (rows, cols) for an array, or (1,1) for scalar. ----
	arrayDims := func(v Value) (int, int) {
		if v.Type == ValueArray {
			rows := len(v.Array)
			if rows == 0 {
				return 0, 0
			}
			return rows, len(v.Array[0])
		}
		return 1, 1
	}

	// ---- 1. known_y's ----
	knownY, ev := flattenNums(args[0])
	if ev != nil {
		return *ev, nil
	}
	n := len(knownY)
	if n == 0 {
		return ErrorVal(ErrValNA), nil
	}
	// Unlike GROWTH, y-values can be negative or zero.

	// ---- 2. known_x's ----
	var knownX []float64
	if len(args) >= 2 && args[1].Type != ValueEmpty {
		knownX, ev = flattenNums(args[1])
		if ev != nil {
			return *ev, nil
		}
		if len(knownX) != n {
			return ErrorVal(ErrValREF), nil
		}
	} else {
		knownX = make([]float64, n)
		for i := range knownX {
			knownX[i] = float64(i + 1)
		}
	}

	// ---- 3. new_x's ----
	var newX []float64
	newXRows, newXCols := 1, 1
	isNewXScalar := true
	if len(args) >= 3 && args[2].Type != ValueEmpty {
		newX, ev = flattenNums(args[2])
		if ev != nil {
			return *ev, nil
		}
		newXRows, newXCols = arrayDims(args[2])
		isNewXScalar = args[2].Type != ValueArray
	} else {
		newX = make([]float64, len(knownX))
		copy(newX, knownX)
		// Shape matches known_x's shape.
		if len(args) >= 2 && args[1].Type != ValueEmpty {
			newXRows, newXCols = arrayDims(args[1])
			isNewXScalar = args[1].Type != ValueArray
		} else {
			newXRows, newXCols = arrayDims(args[0])
			isNewXScalar = args[0].Type != ValueArray
		}
	}

	// ---- 4. const ----
	useConst := true
	if len(args) >= 4 && args[3].Type != ValueEmpty {
		cv, e := CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
		useConst = cv != 0
	}

	// ---- 5. Linear regression of knownY on knownX ----
	var slope, intercept float64
	if useConst {
		// Compute means.
		sumX, sumY := 0.0, 0.0
		for i := 0; i < n; i++ {
			sumX += knownX[i]
			sumY += knownY[i]
		}
		meanX := sumX / float64(n)
		meanY := sumY / float64(n)

		// Compute slope and intercept.
		cov := 0.0
		ssqX := 0.0
		for i := 0; i < n; i++ {
			dx := knownX[i] - meanX
			cov += dx * (knownY[i] - meanY)
			ssqX += dx * dx
		}

		if ssqX == 0 {
			// All x values are the same. slope=0, intercept=meanY.
			slope = 0
			intercept = meanY
		} else {
			slope = cov / ssqX
			intercept = meanY - slope*meanX
		}
	} else {
		// const=FALSE: force intercept=0, so y = slope*x.
		// slope = Σ(x*y) / Σ(x²)
		sumXY := 0.0
		sumXX := 0.0
		for i := 0; i < n; i++ {
			sumXY += knownX[i] * knownY[i]
			sumXX += knownX[i] * knownX[i]
		}
		if sumXX == 0 {
			return ErrorVal(ErrValDIV0), nil
		}
		slope = sumXY / sumXX
		intercept = 0
	}

	// ---- 6. Predict ----
	predicted := make([]float64, len(newX))
	for i, x := range newX {
		predicted[i] = intercept + slope*x
	}

	// ---- 7. Return result ----
	if isNewXScalar && len(predicted) == 1 {
		return NumberVal(predicted[0]), nil
	}

	// Build array matching the shape of new_x's.
	result := make([][]Value, newXRows)
	idx := 0
	for r := 0; r < newXRows; r++ {
		row := make([]Value, newXCols)
		for c := 0; c < newXCols; c++ {
			if idx < len(predicted) {
				row[c] = NumberVal(predicted[idx])
				idx++
			}
		}
		result[r] = row
	}
	return Value{Type: ValueArray, Array: result}, nil
}

// linestQRResidualSS computes the residual sum of squares using Householder
// QR decomposition. When the direct computation gives ssResid == 0 for a
// perfect (or near-perfect) linear fit, the QR approach may still produce a
// tiny non-zero residual due to intermediate floating-point rounding in the
// Householder reflections. This matches the behaviour of Excel's LINEST,
// which uses an analogous QR-based algorithm internally.
func linestQRResidualSS(knownX, knownY []float64, useConst bool) float64 {
	n := len(knownY)
	var p, cols int
	if useConst {
		p, cols = 2, 3
	} else {
		p, cols = 1, 2
	}

	// Build augmented matrix [A | y].
	M := make([][]float64, n)
	for i := 0; i < n; i++ {
		M[i] = make([]float64, cols)
		if useConst {
			M[i][0] = 1.0       // intercept column
			M[i][1] = knownX[i] // x column
			M[i][2] = knownY[i] // response
		} else {
			M[i][0] = knownX[i]
			M[i][1] = knownY[i]
		}
	}

	// Householder QR factorisation applied to the augmented matrix.
	for j := 0; j < p; j++ {
		sigma := 0.0
		for i := j; i < n; i++ {
			sigma += M[i][j] * M[i][j]
		}
		sigma = math.Sqrt(sigma)
		if sigma == 0 {
			continue
		}
		if M[j][j] > 0 {
			sigma = -sigma
		}
		M[j][j] -= sigma

		vtv := 0.0
		for i := j; i < n; i++ {
			vtv += M[i][j] * M[i][j]
		}
		if vtv == 0 {
			continue
		}
		beta := 2.0 / vtv

		for k := j + 1; k < cols; k++ {
			dot := 0.0
			for i := j; i < n; i++ {
				dot += M[i][j] * M[i][k]
			}
			dot *= beta
			for i := j; i < n; i++ {
				M[i][k] -= dot * M[i][j]
			}
		}
		M[j][j] = sigma
	}

	// The residual SS is the sum of squares of the tail of Q^T * y.
	yCol := cols - 1
	ss := 0.0
	for i := p; i < n; i++ {
		ss += M[i][yCol] * M[i][yCol]
	}
	return ss
}

// fnLINEST implements LINEST(known_y's, [known_x's], [const], [stats]).
// It calculates statistics for a line using the least squares method and
// returns an array describing the line.
func fnLINEST(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}

	// ---- Helper: extract a flat slice of float64 from a Value (array or scalar). ----
	flattenNums := func(v Value) ([]float64, *Value) {
		if v.Type == ValueArray {
			var nums []float64
			for _, row := range v.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return nil, &cell
					}
					n, e := CoerceNum(cell)
					if e != nil {
						return nil, e
					}
					nums = append(nums, n)
				}
			}
			return nums, nil
		}
		if v.Type == ValueError {
			return nil, &v
		}
		n, e := CoerceNum(v)
		if e != nil {
			return nil, e
		}
		return []float64{n}, nil
	}

	// ---- 1. known_y's ----
	knownY, ev := flattenNums(args[0])
	if ev != nil {
		return *ev, nil
	}
	n := len(knownY)
	if n == 0 {
		return ErrorVal(ErrValNA), nil
	}

	// ---- 2. known_x's ----
	var knownX []float64
	if len(args) >= 2 && args[1].Type != ValueEmpty {
		knownX, ev = flattenNums(args[1])
		if ev != nil {
			return *ev, nil
		}
		if len(knownX) != n {
			return ErrorVal(ErrValREF), nil
		}
	} else {
		knownX = make([]float64, n)
		for i := range knownX {
			knownX[i] = float64(i + 1)
		}
	}

	// ---- 3. const ----
	useConst := true
	if len(args) >= 3 && args[2].Type != ValueEmpty {
		cv, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		useConst = cv != 0
	}

	// ---- 4. stats ----
	wantStats := false
	if len(args) >= 4 && args[3].Type != ValueEmpty {
		sv, e := CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
		wantStats = sv != 0
	}

	// ---- 5. Compute regression ----
	nf := float64(n)
	var slope, intercept float64
	var ssResid, ssTotal, ssReg float64
	var df float64
	var sumXX float64 // Σ(xi²)   — used for se_intercept
	var ssqX float64  // Σ(xi-x̄)² — used for se_slope (const=TRUE) or Σ(xi²) (const=FALSE)

	if useConst {
		// Compute means.
		sumX, sumY := 0.0, 0.0
		for i := 0; i < n; i++ {
			sumX += knownX[i]
			sumY += knownY[i]
		}
		meanX := sumX / nf
		meanY := sumY / nf

		// Compute slope and intercept.
		cov := 0.0
		ssqX = 0.0
		sumXX = 0.0
		for i := 0; i < n; i++ {
			dx := knownX[i] - meanX
			cov += dx * (knownY[i] - meanY)
			ssqX += dx * dx
			sumXX += knownX[i] * knownX[i]
		}

		if ssqX == 0 {
			slope = 0
			intercept = meanY
		} else {
			slope = cov / ssqX
			intercept = meanY - slope*meanX
		}

		// ss_total = Σ(yi - ȳ)²
		for i := 0; i < n; i++ {
			dy := knownY[i] - meanY
			ssTotal += dy * dy
		}
		// ss_resid = Σ(yi - predicted)²
		for i := 0; i < n; i++ {
			pred := slope*knownX[i] + intercept
			r := knownY[i] - pred
			ssResid += r * r
		}
		ssReg = ssTotal - ssResid
		df = nf - 2 // n - k - 1, k=1
	} else {
		// const=FALSE: force intercept=0.
		// slope = Σ(xi*yi) / Σ(xi²)
		sumXY := 0.0
		sumXX = 0.0
		for i := 0; i < n; i++ {
			sumXY += knownX[i] * knownY[i]
			sumXX += knownX[i] * knownX[i]
		}
		if sumXX == 0 {
			return ErrorVal(ErrValDIV0), nil
		}
		slope = sumXY / sumXX
		intercept = 0
		ssqX = sumXX // For const=FALSE, se_slope uses Σxi² directly.

		// ss_total = Σ(yi²) (not centered)
		for i := 0; i < n; i++ {
			ssTotal += knownY[i] * knownY[i]
		}
		// ss_resid = Σ(yi - slope*xi)²
		for i := 0; i < n; i++ {
			pred := slope * knownX[i]
			r := knownY[i] - pred
			ssResid += r * r
		}
		ssReg = ssTotal - ssResid
		df = nf - 1 // n - k, k=1
	}

	// ---- 6. Build result ----
	if !wantStats {
		// Return 1×2 array: {slope, intercept}
		row := []Value{NumberVal(slope), NumberVal(intercept)}
		return Value{Type: ValueArray, Array: [][]Value{row}}, nil
	}

	// stats=TRUE: build 5×2 array
	var r2, seY, seSlope, seIntercept Value
	var fStat Value

	if df <= 0 {
		// Not enough degrees of freedom.
		seY = ErrorVal(ErrValNA)
		seSlope = ErrorVal(ErrValNA)
		seIntercept = ErrorVal(ErrValNA)
		fStat = ErrorVal(ErrValNA)
		if ssResid == 0 {
			r2 = NumberVal(1) // perfect fit trivially
		} else {
			r2 = ErrorVal(ErrValNA)
		}
	} else {
		// r²
		if ssTotal == 0 {
			if ssResid == 0 {
				r2 = NumberVal(1)
			} else {
				r2 = NumberVal(0)
			}
		} else {
			r2 = NumberVal(ssReg / ssTotal)
		}

		// se_y = sqrt(ss_resid / df)
		seYVal := math.Sqrt(ssResid / df)
		seY = NumberVal(seYVal)

		// se_slope
		if ssqX == 0 {
			seSlope = ErrorVal(ErrValNA)
		} else {
			seSlope = NumberVal(seYVal / math.Sqrt(ssqX))
		}

		// se_intercept
		if useConst {
			if ssqX == 0 {
				seIntercept = ErrorVal(ErrValNA)
			} else {
				seIntercept = NumberVal(seYVal * math.Sqrt(sumXX/(nf*ssqX)))
			}
		} else {
			seIntercept = ErrorVal(ErrValNA)
		}

		// F-statistic = (ss_reg / k) / (ss_resid / df)
		if ssResid == 0 && ssReg > 0 {
			// Direct computation gave exactly zero residual (perfect fit).
			// Fall back to QR decomposition which may produce a tiny non-zero
			// residual from floating-point rounding, matching Excel's behaviour.
			ssResidQR := linestQRResidualSS(knownX, knownY, useConst)
			if ssResidQR > 0 {
				fStat = NumberVal((ssReg / 1) / (ssResidQR / df))
			} else {
				fStat = ErrorVal(ErrValNUM)
			}
		} else if ssResid == 0 {
			fStat = ErrorVal(ErrValNA)
		} else {
			fStat = NumberVal((ssReg / 1) / (ssResid / df))
		}
	}

	// Row 1: {slope, intercept}
	row1 := []Value{NumberVal(slope), NumberVal(intercept)}
	// Row 2: {se_slope, se_intercept}
	row2 := []Value{seSlope, seIntercept}
	// Row 3: {r², se_y}
	row3 := []Value{r2, seY}
	// Row 4: {F, df}
	row4 := []Value{fStat, NumberVal(df)}
	// Row 5: {ss_reg, ss_resid}
	row5 := []Value{NumberVal(ssReg), NumberVal(ssResid)}

	return Value{Type: ValueArray, Array: [][]Value{row1, row2, row3, row4, row5}}, nil
}

// fnLOGEST implements LOGEST(known_y's, [known_x's], [const], [stats]).
// It fits an exponential curve y = b * m^x to the data by performing
// linear regression on ln(y) vs x (i.e., it is the exponential counterpart
// of LINEST).
func fnLOGEST(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}

	// ---- Helper: extract a flat slice of float64 from a Value (array or scalar). ----
	flattenNums := func(v Value) ([]float64, *Value) {
		if v.Type == ValueArray {
			var nums []float64
			for _, row := range v.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return nil, &cell
					}
					n, e := CoerceNum(cell)
					if e != nil {
						return nil, e
					}
					nums = append(nums, n)
				}
			}
			return nums, nil
		}
		if v.Type == ValueError {
			return nil, &v
		}
		n, e := CoerceNum(v)
		if e != nil {
			return nil, e
		}
		return []float64{n}, nil
	}

	// ---- 1. known_y's ----
	knownY, ev := flattenNums(args[0])
	if ev != nil {
		return *ev, nil
	}
	n := len(knownY)
	if n == 0 {
		return ErrorVal(ErrValNA), nil
	}
	// All known_y must be > 0 (since we take ln).
	for _, y := range knownY {
		if y <= 0 {
			return ErrorVal(ErrValNUM), nil
		}
	}

	// ---- 2. known_x's ----
	var knownX []float64
	if len(args) >= 2 && args[1].Type != ValueEmpty {
		knownX, ev = flattenNums(args[1])
		if ev != nil {
			return *ev, nil
		}
		if len(knownX) != n {
			return ErrorVal(ErrValREF), nil
		}
	} else {
		knownX = make([]float64, n)
		for i := range knownX {
			knownX[i] = float64(i + 1)
		}
	}

	// ---- 3. const ----
	useConst := true
	if len(args) >= 3 && args[2].Type != ValueEmpty {
		cv, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		useConst = cv != 0
	}

	// ---- 4. stats ----
	wantStats := false
	if len(args) >= 4 && args[3].Type != ValueEmpty {
		sv, e := CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
		wantStats = sv != 0
	}

	// ---- 5. Compute ln(y) ----
	lnY := make([]float64, n)
	for i, y := range knownY {
		lnY[i] = math.Log(y)
	}

	// ---- 6. Compute regression on ln(y) vs x (same as LINEST) ----
	nf := float64(n)
	var slope, intercept float64
	var ssResid, ssTotal, ssReg float64
	var df float64
	var sumXX float64 // Σ(xi²) — used for se_intercept
	var ssqX float64  // Σ(xi-x̄)² or Σ(xi²) for const=FALSE

	if useConst {
		// Compute means.
		sumX, sumLnY := 0.0, 0.0
		for i := 0; i < n; i++ {
			sumX += knownX[i]
			sumLnY += lnY[i]
		}
		meanX := sumX / nf
		meanLnY := sumLnY / nf

		// Compute slope and intercept.
		cov := 0.0
		ssqX = 0.0
		sumXX = 0.0
		for i := 0; i < n; i++ {
			dx := knownX[i] - meanX
			cov += dx * (lnY[i] - meanLnY)
			ssqX += dx * dx
			sumXX += knownX[i] * knownX[i]
		}

		if ssqX == 0 {
			slope = 0
			intercept = meanLnY
		} else {
			slope = cov / ssqX
			intercept = meanLnY - slope*meanX
		}

		// ss_total = Σ(lnYi - mean(lnY))²
		for i := 0; i < n; i++ {
			dy := lnY[i] - meanLnY
			ssTotal += dy * dy
		}
		// ss_resid = Σ(lnYi - predicted)²
		for i := 0; i < n; i++ {
			pred := slope*knownX[i] + intercept
			r := lnY[i] - pred
			ssResid += r * r
		}
		ssReg = ssTotal - ssResid
		df = nf - 2 // n - k - 1, k=1
	} else {
		// const=FALSE: force intercept=0, so lnY = slope*x.
		// slope = Σ(xi*lnYi) / Σ(xi²)
		sumXLnY := 0.0
		sumXX = 0.0
		for i := 0; i < n; i++ {
			sumXLnY += knownX[i] * lnY[i]
			sumXX += knownX[i] * knownX[i]
		}
		if sumXX == 0 {
			return ErrorVal(ErrValDIV0), nil
		}
		slope = sumXLnY / sumXX
		intercept = 0
		ssqX = sumXX

		// ss_total = Σ(lnYi²) (not centered)
		for i := 0; i < n; i++ {
			ssTotal += lnY[i] * lnY[i]
		}
		// ss_resid = Σ(lnYi - slope*xi)²
		for i := 0; i < n; i++ {
			pred := slope * knownX[i]
			r := lnY[i] - pred
			ssResid += r * r
		}
		ssReg = ssTotal - ssResid
		df = nf - 1 // n - k, k=1
	}

	// ---- 7. Build result ----
	// Row 1: {exp(slope), exp(intercept)} = {m, b} where y = b*m^x
	m := math.Exp(slope)
	b := math.Exp(intercept)

	if !wantStats {
		// Return 1×2 array: {m, b}
		row := []Value{NumberVal(m), NumberVal(b)}
		return Value{Type: ValueArray, Array: [][]Value{row}}, nil
	}

	// stats=TRUE: build 5×2 array
	var r2, seY, seSlope, seIntercept Value
	var fStat Value

	if df <= 0 {
		// Not enough degrees of freedom.
		seY = ErrorVal(ErrValNA)
		seSlope = ErrorVal(ErrValNA)
		seIntercept = ErrorVal(ErrValNA)
		fStat = ErrorVal(ErrValNA)
		if ssResid == 0 {
			r2 = NumberVal(1) // perfect fit trivially
		} else {
			r2 = ErrorVal(ErrValNA)
		}
	} else {
		// r²
		r2Val := 0.0
		if ssTotal == 0 {
			if ssResid == 0 {
				r2Val = 1
			}
		} else {
			r2Val = ssReg / ssTotal
		}
		r2 = NumberVal(r2Val)

		// Detect near-perfect exponential fit. The log transformation
		// introduces floating-point noise that makes ssResid tiny but
		// nonzero even for mathematically perfect exponential data
		// (e.g., y = 2^x). When R² rounds to exactly 1.0 in float64,
		// ssResid is dominated by this noise and must be treated as
		// zero, matching Excel's behaviour.
		perfectFit := r2Val == 1.0 && ssReg > 0

		// For a perfect fit, recompute ssResid so that downstream
		// statistics (se_y, se_slope, se_intercept, F) are consistent.
		if perfectFit {
			ssResid = 0
		}

		// se_y = sqrt(ss_resid / df) — from the ln(y) regression, as-is
		seYVal := math.Sqrt(ssResid / df)
		seY = NumberVal(seYVal)

		// se_slope — raw from ln(y) regression (NOT exponentiated)
		if ssqX == 0 {
			seSlope = ErrorVal(ErrValNA)
		} else {
			seSlope = NumberVal(seYVal / math.Sqrt(ssqX))
		}

		// se_intercept — raw from ln(y) regression (NOT exponentiated)
		if useConst {
			if ssqX == 0 {
				seIntercept = ErrorVal(ErrValNA)
			} else {
				seIntercept = NumberVal(seYVal * math.Sqrt(sumXX/(nf*ssqX)))
			}
		} else {
			seIntercept = ErrorVal(ErrValNA)
		}

		// F-statistic = (ss_reg / k) / (ss_resid / df).
		if ssResid == 0 && ssReg > 0 {
			if useConst {
				// Perfect fit with intercept: fall back to QR
				// decomposition to obtain a tiny non-zero residual
				// from floating-point rounding, matching the
				// approach LINEST uses for perfect linear fits.
				ssResidQR := linestQRResidualSS(knownX, lnY, useConst)
				if ssResidQR > 0 {
					fStat = NumberVal((ssReg / 1) / (ssResidQR / df))
				} else {
					fStat = ErrorVal(ErrValNUM)
				}
			} else {
				// Perfect fit without intercept: infinite F.
				fStat = ErrorVal(ErrValNUM)
			}
		} else if ssResid == 0 {
			fStat = ErrorVal(ErrValNA)
		} else {
			fStat = NumberVal((ssReg / 1) / (ssResid / df))
		}
	}

	// Row 1: {m, b} = {exp(slope), exp(intercept)}
	row1 := []Value{NumberVal(m), NumberVal(b)}
	// Row 2: {se_slope, se_intercept} — raw standard errors from ln(y) regression
	row2 := []Value{seSlope, seIntercept}
	// Row 3: {r², se_y} — from ln(y) regression, as-is
	row3 := []Value{r2, seY}
	// Row 4: {F, df} — from ln(y) regression, as-is
	row4 := []Value{fStat, NumberVal(df)}
	// Row 5: {ss_reg, ss_resid} — from ln(y) regression, as-is
	row5 := []Value{NumberVal(ssReg), NumberVal(ssResid)}

	return Value{Type: ValueArray, Array: [][]Value{row1, row2, row3, row4, row5}}, nil
}
