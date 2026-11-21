package formula

import (
	"math"
	"strconv"
	"strings"
)

func init() {
	Register("DSUM", NoCtx(fnDSum))
	Register("DAVERAGE", NoCtx(fnDAverage))
	Register("DCOUNT", NoCtx(fnDCount))
	Register("DCOUNTA", NoCtx(fnDCountA))
	Register("DGET", NoCtx(fnDGet))
	Register("DMAX", NoCtx(fnDMax))
	Register("DMIN", NoCtx(fnDMin))
	Register("DPRODUCT", NoCtx(fnDProduct))
	Register("DSTDEV", NoCtx(fnDStdev))
	Register("DSTDEVP", NoCtx(fnDStdevP))
	Register("DVAR", NoCtx(fnDVar))
	Register("DVARP", NoCtx(fnDVarP))

	// Register all D-functions as array-forcing so that range arguments
	// are not implicitly intersected to a single cell.
	for _, name := range []string{
		"DSUM", "DAVERAGE", "DCOUNT", "DCOUNTA", "DGET",
		"DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DVAR", "DVARP",
	} {
		arrayForcingFuncs[name] = true
	}
}

// dFunc is the shared infrastructure for D-functions (DSUM, DAVERAGE, DCOUNT,
// DMAX, DMIN, etc.). It extracts the database, resolves the field, evaluates
// criteria, and returns the matching field values.
//
// Arguments:
//
//	args[0] – database  (ValueArray): first row = headers, remaining rows = records
//	args[1] – field     (number or string): 1-based column index or header label
//	args[2] – criteria  (ValueArray): first row = criteria headers, remaining rows = conditions
//
// Returns the Values from the target field column for every record that
// matches the criteria, or an error Value on failure.
func dFunc(args []Value) ([]Value, *Value) {
	if len(args) != 3 {
		e := ErrorVal(ErrValVALUE)
		return nil, &e
	}

	dbArg := args[0]
	fieldArg := args[1]
	critArg := args[2]

	// --- database ---
	if dbArg.Type == ValueError {
		return nil, &dbArg
	}
	dbSource, dbRows, dbCols, errVal := dGridSource(dbArg)
	if errVal != nil {
		return nil, errVal
	}
	headers := dbSource.rowValues(0)

	// --- field ---
	fieldIdx := -1 // 0-based column index into the database
	switch fieldArg.Type {
	case ValueError:
		return nil, &fieldArg
	case ValueNumber:
		// 1-based column index
		idx := int(math.Floor(fieldArg.Num))
		if idx < 1 || idx > dbCols {
			e := ErrorVal(ErrValVALUE)
			return nil, &e
		}
		fieldIdx = idx - 1
	case ValueString:
		label := fieldArg.Str
		for i, h := range headers {
			if strings.EqualFold(ValueToString(h), label) {
				fieldIdx = i
				break
			}
		}
		if fieldIdx < 0 {
			e := ErrorVal(ErrValVALUE)
			return nil, &e
		}
	default:
		// Empty or bool field: try to coerce to a number or string.
		if fieldArg.Type == ValueEmpty {
			e := ErrorVal(ErrValVALUE)
			return nil, &e
		}
		// Bool: coerce to number (TRUE=1, FALSE=0)
		n, ev := CoerceNum(fieldArg)
		if ev != nil {
			return nil, ev
		}
		idx := int(math.Floor(n))
		if idx < 1 || idx > dbCols {
			e := ErrorVal(ErrValVALUE)
			return nil, &e
		}
		fieldIdx = idx - 1
	}

	// --- criteria ---
	if critArg.Type == ValueError {
		return nil, &critArg
	}
	critSource, critRows, critCols, errVal := dGridSource(critArg)
	if errVal != nil {
		return nil, errVal
	}
	critHeaders := critSource.rowValues(0)

	// Map each criteria column to a database column index.
	critColMap := make([]int, critCols)
	for i, ch := range critHeaders {
		chStr := strings.TrimSpace(ValueToString(ch))
		critColMap[i] = -1
		for j, h := range headers {
			if strings.EqualFold(strings.TrimSpace(ValueToString(h)), chStr) {
				critColMap[i] = j
				break
			}
		}
		// If a non-empty criteria header doesn't match any database column,
		// it can never match, so those criteria rows always fail.
	}

	// --- match records ---
	result := make([]Value, 0, max(0, dbRows-1))
	for row := 1; row < dbRows; row++ {
		if dbCriteriaMatch(dbSource, row, critSource, critRows, critCols, critColMap) {
			result = append(result, dbSource.cell(row, fieldIdx))
		}
	}
	return result, nil
}

func dGridSource(v Value) (gridValueSource, int, int, *Value) {
	if v.Type != ValueArray {
		errVal := ErrorVal(ErrValVALUE)
		return gridValueSource{}, 0, 0, &errVal
	}
	source := newGridValueSource(ValueToEvalValue(v))
	rows, cols := source.dims()
	if rows < 1 {
		errVal := ErrorVal(ErrValVALUE)
		return gridValueSource{}, 0, 0, &errVal
	}
	return source, rows, cols, nil
}

// dbCriteriaMatch checks whether a single database record matches the
// criteria.  Criteria rows are OR-ed; within each row, conditions are AND-ed.
func dbCriteriaMatch(
	recordSource gridValueSource,
	recordRow int,
	critSource gridValueSource,
	critRows int,
	critCols int,
	critColMap []int,
) bool {
	// If there are no criteria rows, treat as "match all".
	if critRows <= 1 {
		return true
	}

	for critRow := 1; critRow < critRows; critRow++ {
		if dbRowMatch(recordSource, recordRow, critSource, critRow, critCols, critColMap) {
			return true // OR: any row matching is sufficient
		}
	}
	return false
}

// dbRowMatch checks whether a record matches all conditions in a single
// criteria row (AND logic).
func dbRowMatch(
	recordSource gridValueSource,
	recordRow int,
	critSource gridValueSource,
	critRow int,
	critCols int,
	critColMap []int,
) bool {
	for ci := 0; ci < critCols; ci++ {
		crit := critSource.cell(critRow, ci)

		// Skip blank criteria cells (match all for this column).
		if isBlankCriterion(crit) {
			continue
		}

		dbCol := critColMap[ci]
		if dbCol < 0 {
			// Criteria header didn't match any database column —
			// non-blank criterion can never match.
			return false
		}

		if !dbMatchCell(recordSource.cell(recordRow, dbCol), crit) {
			return false
		}
	}
	return true
}

// isBlankCriterion returns true when a criteria cell should be treated as
// "match all" (i.e. it's empty or an empty string).
func isBlankCriterion(v Value) bool {
	switch v.Type {
	case ValueEmpty:
		return true
	case ValueString:
		return strings.TrimSpace(v.Str) == ""
	default:
		return false
	}
}

// dbMatchCell compares a database cell value against a single criterion.
// It handles numeric comparison operators, exact-match prefix "=", and
// wildcard patterns.
func dbMatchCell(cellVal Value, crit Value) bool {
	critStr := ValueToString(crit)

	// Try to parse comparison operators: >=, <=, <>, >, <, =
	if len(critStr) >= 1 {
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
			return dbMatchOperator(cellVal, op, operand)
		}
	}

	// Wildcard matching for text criteria.
	switch classifyWildcard(critStr) {
	case wildcardFull:
		return WildcardMatch(ValueToString(cellVal), critStr)
	case wildcardEscape:
		return strings.EqualFold(ValueToString(cellVal), unescapePattern(critStr))
	}

	// Numeric criteria value: match numeric cells.
	if crit.Type == ValueNumber {
		if cellVal.Type == ValueNumber {
			return cellVal.Num == crit.Num
		}
		if cellVal.Type == ValueString {
			if n, err := strconv.ParseFloat(cellVal.Str, 64); err == nil {
				return n == crit.Num
			}
		}
		return false
	}

	// Boolean criteria.
	if crit.Type == ValueBool {
		return cellVal.Type == ValueBool && cellVal.Bool == crit.Bool
	}

	// Plain text: case-insensitive exact match.
	return strings.EqualFold(ValueToString(cellVal), critStr)
}

// dbMatchOperator applies a comparison operator in the context of D-function
// criteria matching. The semantics mirror MatchesCriteria's matchOperator.
func dbMatchOperator(cellVal Value, op, operand string) bool {
	// "=" with empty operand matches empty cells.
	if op == "=" && operand == "" {
		return cellVal.Type == ValueEmpty
	}
	if op == "<>" && operand == "" {
		return cellVal.Type != ValueEmpty
	}

	upper := strings.ToUpper(strings.TrimSpace(operand))

	// Boolean operand.
	if upper == "TRUE" || upper == "FALSE" {
		critBool := upper == "TRUE"
		if cellVal.Type == ValueBool {
			cmp := 0
			if cellVal.Bool != critBool {
				if !cellVal.Bool {
					cmp = -1
				} else {
					cmp = 1
				}
			}
			return evalOp(op, cmp)
		}
		if op == "<>" {
			return true
		}
		return false
	}

	// Numeric operand.
	if cn, err := strconv.ParseFloat(operand, 64); err == nil {
		if cellVal.Type == ValueNumber {
			return evalOp(op, cmpFloat(cellVal.Num, cn))
		}
		// For "=" also try coercing string cells to numbers.
		if op == "=" && cellVal.Type == ValueString {
			if n, err2 := strconv.ParseFloat(cellVal.Str, 64); err2 == nil {
				return cmpFloat(n, cn) == 0
			}
		}
		if op == "<>" {
			return true
		}
		return false
	}

	// Wildcard matching inside operator context (e.g., "=App*").
	if op == "=" {
		switch classifyWildcard(operand) {
		case wildcardFull:
			return WildcardMatch(ValueToString(cellVal), operand)
		case wildcardEscape:
			return strings.EqualFold(ValueToString(cellVal), unescapePattern(operand))
		}
	}

	// String comparison.
	cmp := strings.Compare(strings.ToLower(ValueToString(cellVal)), strings.ToLower(operand))
	return evalOp(op, cmp)
}

// ---------------------------------------------------------------------------
// DSUM
// ---------------------------------------------------------------------------

func fnDSum(args []Value) (Value, error) {
	vals, errv := dFunc(args)
	if errv != nil {
		return *errv, nil
	}
	sum := 0.0
	for _, v := range vals {
		switch v.Type {
		case ValueNumber:
			sum += v.Num
		case ValueError:
			return v, nil
		}
		// Ignore text, booleans, empty cells.
	}
	return NumberVal(sum), nil
}

// ---------------------------------------------------------------------------
// DAVERAGE
// ---------------------------------------------------------------------------

func fnDAverage(args []Value) (Value, error) {
	vals, errv := dFunc(args)
	if errv != nil {
		return *errv, nil
	}
	sum := 0.0
	count := 0
	for _, v := range vals {
		switch v.Type {
		case ValueNumber:
			sum += v.Num
			count++
		case ValueError:
			return v, nil
		}
	}
	if count == 0 {
		return ErrorVal(ErrValDIV0), nil
	}
	return NumberVal(sum / float64(count)), nil
}

// ---------------------------------------------------------------------------
// DCOUNT
// ---------------------------------------------------------------------------

func fnDCount(args []Value) (Value, error) {
	vals, errv := dFunc(args)
	if errv != nil {
		return *errv, nil
	}
	count := 0
	for _, v := range vals {
		switch v.Type {
		case ValueNumber:
			count++
		case ValueError:
			return v, nil
		}
	}
	return NumberVal(float64(count)), nil
}

// ---------------------------------------------------------------------------
// DCOUNTA
// ---------------------------------------------------------------------------

func fnDCountA(args []Value) (Value, error) {
	vals, errv := dFunc(args)
	if errv != nil {
		return *errv, nil
	}
	count := 0
	for _, v := range vals {
		switch v.Type {
		case ValueError:
			return v, nil
		case ValueEmpty:
			// skip empty
		default:
			count++
		}
	}
	return NumberVal(float64(count)), nil
}

// ---------------------------------------------------------------------------
// DGET
// ---------------------------------------------------------------------------

func fnDGet(args []Value) (Value, error) {
	vals, errv := dFunc(args)
	if errv != nil {
		return *errv, nil
	}
	if len(vals) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	if len(vals) > 1 {
		return ErrorVal(ErrValNUM), nil
	}
	v := vals[0]
	if v.Type == ValueError {
		return v, nil
	}
	return v, nil
}

// ---------------------------------------------------------------------------
// DMAX
// ---------------------------------------------------------------------------

func fnDMax(args []Value) (Value, error) {
	vals, errv := dFunc(args)
	if errv != nil {
		return *errv, nil
	}
	max := 0.0
	found := false
	for _, v := range vals {
		switch v.Type {
		case ValueNumber:
			if !found || v.Num > max {
				max = v.Num
				found = true
			}
		case ValueError:
			return v, nil
		}
	}
	return NumberVal(max), nil
}

// ---------------------------------------------------------------------------
// DMIN
// ---------------------------------------------------------------------------

func fnDMin(args []Value) (Value, error) {
	vals, errv := dFunc(args)
	if errv != nil {
		return *errv, nil
	}
	min := 0.0
	found := false
	for _, v := range vals {
		switch v.Type {
		case ValueNumber:
			if !found || v.Num < min {
				min = v.Num
				found = true
			}
		case ValueError:
			return v, nil
		}
	}
	return NumberVal(min), nil
}

// ---------------------------------------------------------------------------
// DPRODUCT
// ---------------------------------------------------------------------------

func fnDProduct(args []Value) (Value, error) {
	vals, errv := dFunc(args)
	if errv != nil {
		return *errv, nil
	}
	product := 0.0
	found := false
	for _, v := range vals {
		switch v.Type {
		case ValueNumber:
			if !found {
				product = v.Num
				found = true
			} else {
				product *= v.Num
			}
		case ValueError:
			return v, nil
		}
	}
	return NumberVal(product), nil
}

// ---------------------------------------------------------------------------
// DSTDEV (sample standard deviation, n-1 denominator)
// ---------------------------------------------------------------------------

func fnDStdev(args []Value) (Value, error) {
	vals, errv := dFunc(args)
	if errv != nil {
		return *errv, nil
	}
	variance, ev := dVariance(vals, false)
	if ev != nil {
		return *ev, nil
	}
	return NumberVal(math.Sqrt(variance)), nil
}

// ---------------------------------------------------------------------------
// DSTDEVP (population standard deviation, n denominator)
// ---------------------------------------------------------------------------

func fnDStdevP(args []Value) (Value, error) {
	vals, errv := dFunc(args)
	if errv != nil {
		return *errv, nil
	}
	variance, ev := dVariance(vals, true)
	if ev != nil {
		return *ev, nil
	}
	return NumberVal(math.Sqrt(variance)), nil
}

// ---------------------------------------------------------------------------
// DVAR (sample variance, n-1 denominator)
// ---------------------------------------------------------------------------

func fnDVar(args []Value) (Value, error) {
	vals, errv := dFunc(args)
	if errv != nil {
		return *errv, nil
	}
	variance, ev := dVariance(vals, false)
	if ev != nil {
		return *ev, nil
	}
	return NumberVal(variance), nil
}

// ---------------------------------------------------------------------------
// DVARP (population variance, n denominator)
// ---------------------------------------------------------------------------

func fnDVarP(args []Value) (Value, error) {
	vals, errv := dFunc(args)
	if errv != nil {
		return *errv, nil
	}
	variance, ev := dVariance(vals, true)
	if ev != nil {
		return *ev, nil
	}
	return NumberVal(variance), nil
}

// dVariance computes variance over numeric values. If population is true it
// uses n as the denominator; otherwise it uses n-1 (sample variance).
// Returns #DIV/0! if there are insufficient numeric values.
func dVariance(vals []Value, population bool) (float64, *Value) {
	var nums []float64
	for _, v := range vals {
		switch v.Type {
		case ValueNumber:
			nums = append(nums, v.Num)
		case ValueError:
			return 0, &v
		}
	}
	n := len(nums)
	if population {
		if n < 1 {
			e := ErrorVal(ErrValDIV0)
			return 0, &e
		}
	} else {
		if n < 2 {
			e := ErrorVal(ErrValDIV0)
			return 0, &e
		}
	}

	// Compute mean.
	sum := 0.0
	for _, x := range nums {
		sum += x
	}
	mean := sum / float64(n)

	// Compute sum of squared deviations.
	ss := 0.0
	for _, x := range nums {
		d := x - mean
		ss += d * d
	}

	denom := float64(n)
	if !population {
		denom = float64(n - 1)
	}
	return ss / denom, nil
}
