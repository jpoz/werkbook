package formula

import (
	"fmt"
	"math"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

func init() {
	Register("ADDRESS", NoCtx(fnADDRESS))
	Register("ANCHORARRAY", fnANCHORARRAY)
	RegisterWithSpec("FILTER", NoCtx(fnFILTER), gridShapeFuncSpec(evalFILTER))
	RegisterWithSpec("HLOOKUP", NoCtx(fnHLOOKUP), hlookupFuncSpec())
	RegisterWithSpec("INDEX", NoCtx(fnINDEX), indexFuncSpec(evalINDEXSelector))
	RegisterWithSpec("INDIRECT", fnINDIRECT, refProducerFuncSpec(evalINDIRECT))
	RegisterWithSpec("LOOKUP", NoCtx(fnLOOKUP), lookupFuncSpec())
	RegisterWithSpec("MATCH", NoCtx(fnMATCH), matchFuncSpec())
	RegisterWithSpec("OFFSET", fnOFFSET, refProducerFuncSpec(evalOFFSET))
	Register("SINGLE", fnSINGLE)
	RegisterWithSpec("VLOOKUP", NoCtx(fnVLOOKUP), vlookupFuncSpec())
	RegisterWithSpec("TAKE", NoCtx(fnTAKE), selectorFuncSpec(evalTAKESelector))
	RegisterWithSpec("DROP", NoCtx(fnDROP), selectorFuncSpec(evalDROPSelector))
	Register("EXPAND", NoCtx(fnEXPAND))
	RegisterWithSpec("CHOOSECOLS", NoCtx(fnCHOOSECOLS), selectorFuncSpec(evalCHOOSECOLSSelector))
	RegisterWithSpec("CHOOSEROWS", NoCtx(fnCHOOSEROWS), selectorFuncSpec(evalCHOOSEROWSSelector))
	Register("TOCOL", NoCtx(fnTOCOL))
	Register("TOROW", NoCtx(fnTOROW))
	Register("TRANSPOSE", NoCtx(fnTRANSPOSE))
	RegisterWithSpec("UNIQUE", NoCtx(fnUNIQUE), gridShapeFuncSpec(evalUNIQUE))
	Register("WRAPCOLS", NoCtx(fnWRAPCOLS))
	Register("WRAPROWS", NoCtx(fnWRAPROWS))
	Register("HSTACK", NoCtx(fnHSTACK))
	Register("VSTACK", NoCtx(fnVSTACK))
	Register("HYPERLINK", NoCtx(fnHyperlink))
	RegisterWithSpec("XLOOKUP", NoCtx(fnXLOOKUP), xlookupFuncSpec())
	RegisterWithSpec("XMATCH", NoCtx(fnXMATCH), xmatchFuncSpec())
}

// lookupArg0 is the shared ArgSpec for the lookup_value argument: scalar
// context collapses an array to its top-left (or range-aligned) cell via
// ArgAdaptScalarizeAny, while array context preserves the array so the
// FnKindLookupArrayLift dispatch in callFuncWithSpec can fan it out.
var lookupArg0 = ArgSpec{Load: ArgLoadPassthrough, Adapt: ArgAdaptScalarizeAny}

// lookupPassRef is the shared ArgSpec for trailing arguments (table/range,
// column index, match type, etc.). Arrays and ranges pass through unchanged.
var lookupPassRef = ArgSpec{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough}

// xlookupFuncSpec wires XLOOKUP into the Phase 2 contract system: scalar
// context scalarizes the first argument via legacy implicit intersection,
// and array context fans it out element-wise through FnKindLookupArrayLift.
func xlookupFuncSpec() FuncSpec {
	return FuncSpec{
		Kind:   FnKindLookupArrayLift,
		Args:   []ArgSpec{lookupArg0, lookupPassRef, lookupPassRef},
		VarArg: func(_ int) ArgSpec { return lookupPassRef },
		Return: ReturnModePassThrough,
	}
}

// xmatchFuncSpec mirrors xlookupFuncSpec with two positional args.
func xmatchFuncSpec() FuncSpec {
	return FuncSpec{
		Kind:   FnKindLookupArrayLift,
		Args:   []ArgSpec{lookupArg0, lookupPassRef},
		VarArg: func(_ int) ArgSpec { return lookupPassRef },
		Return: ReturnModePassThrough,
	}
}

// vlookupFuncSpec wires VLOOKUP. Excel fans out the lookup_value array in
// array context (e.g. inside FILTER's include argument) and collapses it in
// scalar context (via implicit intersection on range-backed arrays).
func vlookupFuncSpec() FuncSpec {
	return FuncSpec{
		Kind:   FnKindLookupArrayLift,
		Args:   []ArgSpec{lookupArg0, lookupPassRef, lookupPassRef},
		VarArg: func(_ int) ArgSpec { return lookupPassRef },
		Return: ReturnModePassThrough,
	}
}

// hlookupFuncSpec wires HLOOKUP with the same shape as VLOOKUP.
func hlookupFuncSpec() FuncSpec {
	return FuncSpec{
		Kind:   FnKindLookupArrayLift,
		Args:   []ArgSpec{lookupArg0, lookupPassRef, lookupPassRef},
		VarArg: func(_ int) ArgSpec { return lookupPassRef },
		Return: ReturnModePassThrough,
	}
}

// matchFuncSpec wires MATCH with two positional args (lookup_value,
// lookup_array) and optional match_type.
func matchFuncSpec() FuncSpec {
	return FuncSpec{
		Kind:   FnKindLookupArrayLift,
		Args:   []ArgSpec{lookupArg0, lookupPassRef},
		VarArg: func(_ int) ArgSpec { return lookupPassRef },
		Return: ReturnModePassThrough,
	}
}

// lookupFuncSpec wires the legacy LOOKUP function.
func lookupFuncSpec() FuncSpec {
	return FuncSpec{
		Kind:   FnKindLookupArrayLift,
		Args:   []ArgSpec{lookupArg0, lookupPassRef},
		VarArg: func(_ int) ArgSpec { return lookupPassRef },
		Return: ReturnModePassThrough,
	}
}

func fnSINGLE(args []Value, ctx *EvalContext) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return explicitIntersect(args[0], ctx), nil
}

// fnANCHORARRAY implements ANCHORARRAY(ref). It returns the full dynamic
// array (spilled range) produced by the formula in the anchor cell ref.
// If the referenced cell has no formula or produces a scalar, the scalar
// value is returned.
func fnANCHORARRAY(args []Value, ctx *EvalContext) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type != ValueRef {
		return ErrorVal(ErrValVALUE), nil
	}
	if ctx == nil || ctx.Resolver == nil {
		return ErrorVal(ErrValREF), nil
	}

	row := int(args[0].Num) / 100_000
	col := int(args[0].Num) % 100_000
	sheet := args[0].Str

	fae, ok := ctx.Resolver.(FormulaArrayEvaluator)
	if !ok {
		// Resolver does not support formula array evaluation; fall back to
		// loading the scalar cell value.
		return ctx.Resolver.GetCellValue(CellAddr{Sheet: sheet, Col: col, Row: row}), nil
	}

	return fae.EvalCellFormula(sheet, col, row), nil
}

// fnFILTER implements FILTER(array, include, [if_empty]).
// It filters rows or columns of an array based on a Boolean array.
func fnFILTER(args []Value) (Value, error) {
	return filterCore(args, nil)
}

func evalFILTER(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	return ValueToEvalValue(filterCoreEval(args, nil)), nil
}

func filterCore(args []Value, evalArgs []EvalValue) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}

	arrSource, errVal := normalizeGridShapeArg(args[0], evalArgAt(evalArgs, 0))
	if errVal != nil {
		if errVal.Err == ErrValVALUE {
			return ErrorVal(ErrValCALC), nil
		}
		return *errVal, nil
	}
	numRows, numCols := arrSource.dims()
	if numRows == 0 || numCols == 0 {
		return ErrorVal(ErrValCALC), nil
	}

	incSource, errVal := normalizeGridShapeArg(args[1], evalArgAt(evalArgs, 1))
	if errVal != nil {
		return *errVal, nil
	}
	incRows, incCols := incSource.dims()

	// Flatten include to a single list of values.
	filterByCol := false
	var includeVals []Value
	if incRows == numRows && (incCols == 1 || incRows == 1 && incCols == 1) {
		// Row filtering: include is a column vector with same row count.
		includeVals = make([]Value, numRows)
		for row := 0; row < numRows; row++ {
			includeVals[row] = incSource.cell(row, 0)
		}
	} else if incRows == 1 && incCols == numCols {
		// Column filtering: include is a row vector with same column count.
		filterByCol = true
		includeVals = make([]Value, numCols)
		for col := 0; col < numCols; col++ {
			includeVals[col] = incSource.cell(0, col)
		}
	} else if incRows == numRows {
		// Multi-column include with same rows — flatten first column.
		includeVals = make([]Value, numRows)
		for row := 0; row < numRows; row++ {
			includeVals[row] = incSource.cell(row, 0)
		}
	} else {
		return ErrorVal(ErrValVALUE), nil
	}

	if !filterByCol {
		// Row filtering.
		if len(includeVals) != numRows {
			return ErrorVal(ErrValVALUE), nil
		}
		// Excel's FILTER, when the include argument contains errors, does
		// not abort to a scalar — the error spills across the whole value
		// shape, so COUNTA over the result sees one cell per input row and
		// SUM propagates the error cleanly. Mirror that here by replicating
		// the first error into a grid the size of the value argument.
		for _, iv := range includeVals {
			if iv.Type == ValueError {
				return arrSource.spilledError(iv), nil
			}
		}
		var keepRows []int
		for i, iv := range includeVals {
			n, e := CoerceNum(iv)
			if e != nil {
				return *e, nil
			}
			if n != 0 {
				keepRows = append(keepRows, i)
			}
		}
		if len(keepRows) == 0 {
			if len(args) == 3 {
				return legacyArgValue(args[2], evalArgAt(evalArgs, 2)), nil
			}
			return ErrorVal(ErrValCALC), nil
		}
		return collapseArrayResult(arrSource.materializeRows(keepRows)), nil
	}

	// Column filtering.
	if len(includeVals) != numCols {
		return ErrorVal(ErrValVALUE), nil
	}
	for _, iv := range includeVals {
		if iv.Type == ValueError {
			return arrSource.spilledError(iv), nil
		}
	}
	var keepCols []int
	for col, iv := range includeVals {
		n, e := CoerceNum(iv)
		if e != nil {
			return *e, nil
		}
		if n != 0 {
			keepCols = append(keepCols, col)
		}
	}
	if len(keepCols) == 0 {
		if len(args) == 3 {
			return legacyArgValue(args[2], evalArgAt(evalArgs, 2)), nil
		}
		return ErrorVal(ErrValCALC), nil
	}
	return collapseArrayResult(arrSource.materializeCols(keepCols)), nil
}

func fnVLOOKUP(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	lookup := args[0]
	table := args[1]
	if table.Type == ValueError {
		return table, nil
	}
	if table.Type != ValueArray {
		return ErrorVal(ErrValVALUE), nil
	}
	colIdx, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	ci := int(colIdx)
	if ci < 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	rangeLookup := true
	if len(args) == 4 {
		rangeLookup = IsTruthy(args[3])
	}

	if rangeLookup {
		// Excel uses binary search for approximate match. This matters
		// on unsorted data where a linear scan and binary search diverge.
		lo, hi := 0, len(table.Array)-1
		result := -1
		for lo <= hi {
			mid := (lo + hi) / 2
			row := table.Array[mid]
			if len(row) == 0 {
				// Skip empty rows by shrinking the window.
				hi = mid - 1
				continue
			}
			cmp := CompareValues(row[0], lookup)
			if cmp == 0 {
				result = mid
				break
			} else if cmp < 0 {
				result = mid
				lo = mid + 1
			} else {
				hi = mid - 1
			}
		}
		if result < 0 {
			return ErrorVal(ErrValNA), nil
		}
		if ci > len(table.Array[result]) {
			return ErrorVal(ErrValREF), nil
		}
		return table.Array[result][ci-1], nil
	}

	// Determine if wildcard matching is needed (only for string lookups).
	useWildcard := false
	if lookup.Type == ValueString {
		wm := classifyWildcard(lookup.Str)
		if wm == wildcardFull {
			useWildcard = true
		} else if wm == wildcardEscape {
			// Tilde escapes with no unescaped wildcards: compare against
			// the unescaped literal string.
			lookup = StringVal(unescapePattern(lookup.Str))
		}
	}

	for _, row := range table.Array {
		if len(row) == 0 {
			continue
		}
		cell := row[0]
		// In Excel, VLOOKUP exact match skips truly empty cells.
		if cell.Type == ValueEmpty {
			continue
		}
		matched := false
		if useWildcard {
			if cell.Type == ValueString {
				matched = WildcardMatch(cell.Str, lookup.Str)
			}
		} else {
			matched = CompareValuesExact(cell, lookup) == 0
		}
		if matched {
			if ci > len(row) {
				return ErrorVal(ErrValREF), nil
			}
			return row[ci-1], nil
		}
	}
	return ErrorVal(ErrValNA), nil
}

func fnHLOOKUP(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	lookup := args[0]
	table := args[1]
	if table.Type != ValueArray || len(table.Array) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	rowIdx, e := CoerceNum(args[2])
	if e != nil {
		return *e, nil
	}
	ri := int(rowIdx)
	if ri < 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if ri > len(table.Array) {
		return ErrorVal(ErrValREF), nil
	}

	rangeLookup := true
	if len(args) == 4 {
		rangeLookup = IsTruthy(args[3])
	}

	firstRow := table.Array[0]

	if rangeLookup {
		// Excel uses binary search for approximate match.
		lo, hi := 0, len(firstRow)-1
		result := -1
		for lo <= hi {
			mid := (lo + hi) / 2
			cmp := CompareValues(firstRow[mid], lookup)
			if cmp == 0 {
				result = mid
				break
			} else if cmp < 0 {
				result = mid
				lo = mid + 1
			} else {
				hi = mid - 1
			}
		}
		if result < 0 {
			return ErrorVal(ErrValNA), nil
		}
		if result >= len(table.Array[ri-1]) {
			return ErrorVal(ErrValREF), nil
		}
		return table.Array[ri-1][result], nil
	}

	// Determine if wildcard matching is needed (only for string lookups).
	useWildcard := false
	if lookup.Type == ValueString {
		wm := classifyWildcard(lookup.Str)
		if wm == wildcardFull {
			useWildcard = true
		} else if wm == wildcardEscape {
			// Tilde escapes with no unescaped wildcards: compare against
			// the unescaped literal string.
			lookup = StringVal(unescapePattern(lookup.Str))
		}
	}

	for i, cell := range firstRow {
		if cell.Type == ValueEmpty {
			continue
		}
		matched := false
		if useWildcard {
			if cell.Type == ValueString {
				matched = WildcardMatch(cell.Str, lookup.Str)
			}
		} else {
			matched = CompareValuesExact(cell, lookup) == 0
		}
		if matched {
			if i >= len(table.Array[ri-1]) {
				return ErrorVal(ErrValREF), nil
			}
			return table.Array[ri-1][i], nil
		}
	}
	return ErrorVal(ErrValNA), nil
}

func fnINDEX(args []Value) (Value, error) {
	return callSelectorEval(evalINDEXSelector, args)
}

func callSelectorEval(eval EvalFunc, args []Value) (Value, error) {
	evalArgs := make([]EvalValue, len(args))
	for i, arg := range args {
		evalArgs[i] = ValueToEvalValue(arg)
	}
	result, err := eval(evalArgs, nil)
	if err != nil {
		return Value{}, err
	}
	return EvalValueToValue(result), nil
}

func normalizeIndexSelector(idx, max int) (int, *Value) {
	if idx < 0 {
		errVal := ErrorVal(ErrValVALUE)
		return 0, &errVal
	}
	if idx > max {
		errVal := ErrorVal(ErrValREF)
		return 0, &errVal
	}
	return idx, nil
}

func indexArrayValue(arr Value, rowIdx, colIdx int) Value {
	if rowIdx >= 0 && rowIdx < len(arr.Array) && colIdx >= 0 && colIdx < len(arr.Array[rowIdx]) {
		return arr.Array[rowIdx][colIdx]
	}
	return EmptyVal()
}

func fnMATCH(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	lookup := args[0]
	arr := args[1]
	if lookup.Type == ValueError {
		return lookup, nil
	}
	if arr.Type == ValueError {
		return arr, nil
	}
	matchType := 1
	if len(args) == 3 {
		mt, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		matchType = int(mt)
	}

	var values []Value
	if arr.Type == ValueArray {
		grid, errVal := normalizeToArrayGrid(arr)
		if errVal != nil {
			return *errVal, nil
		}
		values = make([]Value, 0, grid.rowCount*grid.colCount)
		for r := 0; r < grid.rowCount; r++ {
			values = append(values, grid.row(r)...)
		}
	} else {
		values = []Value{arr}
	}

	// For exact match, support wildcard matching on string lookups.
	useWildcard := false
	if matchType == 0 && lookup.Type == ValueString {
		wm := classifyWildcard(lookup.Str)
		if wm == wildcardFull {
			useWildcard = true
		} else if wm == wildcardEscape {
			lookup = StringVal(unescapePattern(lookup.Str))
		}
	}

	switch matchType {
	case 0:
		for i, v := range values {
			matched := false
			if useWildcard {
				if v.Type == ValueString {
					matched = WildcardMatch(v.Str, lookup.Str)
				}
			} else {
				matched = CompareValuesExact(v, lookup) == 0
			}
			if matched {
				return NumberVal(float64(i + 1)), nil
			}
		}
		return ErrorVal(ErrValNA), nil

	case 1:
		// Approximate match (ascending). Empty cells are skipped so that
		// whole-column references (e.g. Q:Q) with sparse data work
		// correctly. We scan all non-empty values and track the last
		// position where value <= lookup, without breaking early. This
		// handles unsorted data gracefully (common in real-world
		// spreadsheets that omit the match_type argument) while producing
		// identical results for properly sorted data.
		last := -1
		for i, v := range values {
			if v.Type == ValueEmpty {
				continue
			}
			cmp := CompareValues(v, lookup)
			if cmp <= 0 {
				last = i
			}
		}
		if last < 0 {
			return ErrorVal(ErrValNA), nil
		}
		return NumberVal(float64(last + 1)), nil

	case -1:
		// Approximate match (descending). Uses a binary search
		// expecting descending-sorted data. We replicate that binary
		// search so that unsorted data produces the same result (often
		// #N/A) as expected.
		n := len(values)
		if n == 0 {
			return ErrorVal(ErrValNA), nil
		}
		lo, hi := 0, n-1
		result := -1
		for lo <= hi {
			mid := lo + (hi-lo)/2
			v := values[mid]
			if v.Type == ValueEmpty {
				// Treat empty as less than any lookup value, search left.
				hi = mid - 1
				continue
			}
			cmp := CompareValues(v, lookup)
			if cmp == 0 {
				// Exact match – smallest value >= lookup.
				return NumberVal(float64(mid + 1)), nil
			} else if cmp > 0 {
				// v > lookup – this is a candidate (>= lookup), but
				// there might be a smaller one further right (descending).
				result = mid
				lo = mid + 1
			} else {
				// v < lookup – need larger values, go left (descending).
				hi = mid - 1
			}
		}
		if result < 0 {
			return ErrorVal(ErrValNA), nil
		}
		return NumberVal(float64(result + 1)), nil
	}

	return ErrorVal(ErrValVALUE), nil
}

func fnADDRESS(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 5 {
		return ErrorVal(ErrValVALUE), nil
	}
	rowNum, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	colNum, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	row := int(rowNum)
	col := int(colNum)
	if row < 1 || col < 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	absNum := 1
	if len(args) >= 3 {
		a, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		absNum = int(a)
	}

	a1Style := true
	if len(args) >= 4 {
		a1Style = IsTruthy(args[3])
	}

	sheetText := ""
	if len(args) >= 5 {
		sheetText = ValueToString(args[4])
	}

	var result string
	if a1Style {
		colName := ColNumberToLetters(col)
		switch absNum {
		case 1:
			result = fmt.Sprintf("$%s$%d", colName, row)
		case 2:
			result = fmt.Sprintf("%s$%d", colName, row)
		case 3:
			result = fmt.Sprintf("$%s%d", colName, row)
		case 4:
			result = fmt.Sprintf("%s%d", colName, row)
		default:
			return ErrorVal(ErrValVALUE), nil
		}
	} else {
		switch absNum {
		case 1:
			result = fmt.Sprintf("R%dC%d", row, col)
		case 2:
			result = fmt.Sprintf("R%dC[%d]", row, col)
		case 3:
			result = fmt.Sprintf("R[%d]C%d", row, col)
		case 4:
			result = fmt.Sprintf("R[%d]C[%d]", row, col)
		default:
			return ErrorVal(ErrValVALUE), nil
		}
	}

	if sheetText != "" {
		needsQuote := strings.ContainsAny(sheetText, " '[")
		if needsQuote {
			result = "'" + sheetText + "'!" + result
		} else {
			result = sheetText + "!" + result
		}
	}

	return StringVal(result), nil
}

func fnLOOKUP(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	lookup := args[0]
	lookupArr := args[1]
	resultArr := lookupArr
	if len(args) == 3 {
		resultArr = args[2]
	}

	var lookupValues []Value
	if lookupArr.Type == ValueArray {
		for _, row := range lookupArr.Array {
			lookupValues = append(lookupValues, row...)
		}
	} else {
		lookupValues = []Value{lookupArr}
	}

	var resultValues []Value
	if resultArr.Type == ValueArray {
		for _, row := range resultArr.Array {
			resultValues = append(resultValues, row...)
		}
	} else {
		resultValues = []Value{resultArr}
	}

	lastMatch := -1
	for i, v := range lookupValues {
		cmp := CompareValues(v, lookup)
		if cmp <= 0 {
			lastMatch = i
		}
		if cmp > 0 {
			break
		}
	}

	if lastMatch < 0 || lastMatch >= len(resultValues) {
		return ErrorVal(ErrValNA), nil
	}
	return resultValues[lastMatch], nil
}

func fnXLOOKUP(args []Value) (Value, error) {
	if len(args) < 3 || len(args) > 6 {
		return ErrorVal(ErrValVALUE), nil
	}
	lookup := args[0]
	lookupArr := args[1]
	returnArr := args[2]

	notFound := ErrorVal(ErrValNA)
	if len(args) >= 4 && args[3].Type != ValueEmpty {
		notFound = args[3]
	}

	matchMode := 0
	if len(args) >= 5 {
		mm, e := CoerceNum(args[4])
		if e != nil {
			return *e, nil
		}
		matchMode = int(mm)
	}

	searchMode := 1
	if len(args) >= 6 {
		sm, e := CoerceNum(args[5])
		if e != nil {
			return *e, nil
		}
		searchMode = int(sm)
	}

	var lookupValues []Value
	if lookupArr.Type == ValueArray {
		for _, row := range lookupArr.Array {
			lookupValues = append(lookupValues, row...)
		}
	} else {
		lookupValues = []Value{lookupArr}
	}

	// Determine lookup orientation: row-oriented if the lookup array is a
	// single row with multiple columns; column-oriented otherwise.
	isRowOriented := lookupArr.Type == ValueArray &&
		len(lookupArr.Array) == 1 && len(lookupArr.Array[0]) > 1

	n := len(lookupValues)

	xlookupReturn := func(i int) (Value, error) {
		if returnArr.Type != ValueArray {
			if i == 0 {
				return returnArr, nil
			}
			return ErrorVal(ErrValNA), nil
		}
		if isRowOriented {
			// Row-oriented lookup: index i is a column index in the
			// return array. Extract column i across all rows.
			rows := returnArr.Array
			if len(rows) == 1 {
				// Single-row return → scalar
				if i >= 0 && i < len(rows[0]) {
					return rows[0][i], nil
				}
				return ErrorVal(ErrValNA), nil
			}
			col := make([][]Value, len(rows))
			for r, row := range rows {
				if i >= 0 && i < len(row) {
					col[r] = []Value{row[i]}
				} else {
					col[r] = []Value{EmptyVal()}
				}
			}
			return Value{Type: ValueArray, Array: col}, nil
		}
		// Column-oriented lookup: index i is a row index.
		if i < 0 || i >= len(returnArr.Array) {
			return ErrorVal(ErrValNA), nil
		}
		row := returnArr.Array[i]
		if len(row) <= 1 {
			if len(row) == 1 {
				return row[0], nil
			}
			return EmptyVal(), nil
		}
		return Value{Type: ValueArray, Array: [][]Value{row}}, nil
	}

	// --- Binary search modes (search_mode 2 or -2) ---
	if searchMode == 2 || searchMode == -2 {
		ascending := searchMode == 2
		idx := xlookupBinarySearch(lookupValues, lookup, matchMode, ascending)
		if idx >= 0 {
			return xlookupReturn(idx)
		}
		return notFound, nil
	}

	// --- Linear search modes (search_mode 1 or -1) ---
	// Determine iteration order: forward (1) or reverse (-1).
	start, end, step := 0, n, 1
	if searchMode == -1 {
		start, end, step = n-1, -1, -1
	}

	switch matchMode {
	case 0: // Exact match
		for i := start; i != end; i += step {
			if CompareValuesExact(lookupValues[i], lookup) == 0 {
				return xlookupReturn(i)
			}
		}

	case -1: // Exact match or next smaller item
		// Scan in the specified direction, keeping the last index where v <= lookup.
		// On sorted ascending data this finds the correct "next smaller" value.
		lastMatch := -1
		for i := start; i != end; i += step {
			if CompareValues(lookupValues[i], lookup) <= 0 {
				lastMatch = i
			}
		}
		if lastMatch >= 0 {
			return xlookupReturn(lastMatch)
		}

	case 1: // Exact match or next larger item
		// Scan in the specified direction, returning the first index where v >= lookup.
		for i := start; i != end; i += step {
			if CompareValues(lookupValues[i], lookup) >= 0 {
				return xlookupReturn(i)
			}
		}

	case 2: // Wildcard match
		pattern := ValueToString(lookup)
		re, err := patternToRegexp(pattern)
		if err != nil {
			return ErrorVal(ErrValVALUE), nil
		}
		anchored, err := regexp.Compile("(?i)^" + re.String() + "$")
		if err != nil {
			return ErrorVal(ErrValVALUE), nil
		}
		for i := start; i != end; i += step {
			s := ValueToString(lookupValues[i])
			if anchored.MatchString(s) {
				return xlookupReturn(i)
			}
		}
	}

	return notFound, nil
}

// xlookupBinarySearch performs a binary search on lookupValues for the given
// lookup value, respecting matchMode (0=exact, -1=next smaller, 1=next larger)
// and whether the data is sorted ascending or descending.
func xlookupBinarySearch(lookupValues []Value, lookup Value, matchMode int, ascending bool) int {
	n := len(lookupValues)
	if n == 0 {
		return -1
	}

	// Binary search for exact match position.
	// For ascending data, sort.Search finds the first index where lookupValues[i] >= lookup.
	// For descending data, sort.Search finds the first index where lookupValues[i] <= lookup.
	idx := sort.Search(n, func(i int) bool {
		cmp := CompareValues(lookupValues[i], lookup)
		if ascending {
			return cmp >= 0
		}
		return cmp <= 0
	})

	switch matchMode {
	case 0: // Exact match only
		if idx < n && CompareValues(lookupValues[idx], lookup) == 0 {
			return idx
		}
		return -1

	case -1: // Exact or next smaller
		if idx < n && CompareValues(lookupValues[idx], lookup) == 0 {
			return idx
		}
		if ascending {
			// The element before idx is the largest element < lookup.
			if idx > 0 {
				return idx - 1
			}
		} else {
			// In descending data, idx points at first element <= lookup.
			// If it's not exact, it's the next smaller element.
			if idx < n {
				return idx
			}
		}
		return -1

	case 1: // Exact or next larger
		if idx < n && CompareValues(lookupValues[idx], lookup) == 0 {
			return idx
		}
		if ascending {
			// idx points at first element >= lookup; if not exact, it's next larger.
			if idx < n {
				return idx
			}
		} else {
			// In descending data, the element before idx is the first > lookup.
			if idx > 0 {
				return idx - 1
			}
		}
		return -1
	}

	return -1
}

// fnINDIRECT implements INDIRECT(ref_text, [a1]).
// It converts a text string into a cell or range reference and resolves it.
func fnINDIRECT(args []Value, ctx *EvalContext) (Value, error) {
	return callLegacyRefEval(evalINDIRECT, args, ctx)
}

// indirectParseCell parses a cell reference like "A1" or "B3" into (col, row).
func indirectParseCell(s string) (col, row int, err error) {
	if s == "" {
		return 0, 0, fmt.Errorf("empty cell reference")
	}
	i := 0
	for i < len(s) && ((s[i] >= 'A' && s[i] <= 'Z') || (s[i] >= 'a' && s[i] <= 'z')) {
		i++
	}
	if i == 0 || i == len(s) {
		return 0, 0, fmt.Errorf("invalid cell reference %q", s)
	}
	col = ColLettersToNumber(s[:i])
	if col < 1 || col > maxCols {
		return 0, 0, fmt.Errorf("column out of range in %q", s)
	}
	row, err = strconv.Atoi(s[i:])
	if err != nil || row < 1 || row > maxRows {
		return 0, 0, fmt.Errorf("invalid row in %q", s)
	}
	return col, row, nil
}

// indirectParseRange parses the left and right parts of a range reference.
// Supports cell:cell (A1:C5), row:row (1:20), and col:col (A:C).
func indirectParseRange(left, right, sheet string) (RangeAddr, error) {
	leftIsRowOnly := isAllDigits(left)
	rightIsRowOnly := isAllDigits(right)
	leftIsColOnly := isAllLetters(left)
	rightIsColOnly := isAllLetters(right)

	// Row-only range like "1:20"
	if leftIsRowOnly && rightIsRowOnly {
		r1, err := strconv.Atoi(left)
		if err != nil || r1 < 1 {
			return RangeAddr{}, fmt.Errorf("invalid row %q", left)
		}
		r2, err := strconv.Atoi(right)
		if err != nil || r2 < 1 {
			return RangeAddr{}, fmt.Errorf("invalid row %q", right)
		}
		if r1 > r2 {
			r1, r2 = r2, r1
		}
		return RangeAddr{
			Sheet:   sheet,
			FromCol: 1, FromRow: r1,
			ToCol: maxCols, ToRow: r2,
		}, nil
	}

	// Column-only range like "A:C"
	if leftIsColOnly && rightIsColOnly {
		c1 := ColLettersToNumber(left)
		c2 := ColLettersToNumber(right)
		if c1 < 1 || c2 < 1 {
			return RangeAddr{}, fmt.Errorf("invalid column range")
		}
		if c1 > c2 {
			c1, c2 = c2, c1
		}
		return RangeAddr{
			Sheet:   sheet,
			FromCol: c1, FromRow: 1,
			ToCol: c2, ToRow: maxRows,
		}, nil
	}

	// Standard cell:cell range like "A1:C5"
	c1, r1, err := indirectParseCell(left)
	if err != nil {
		return RangeAddr{}, err
	}
	c2, r2, err := indirectParseCell(right)
	if err != nil {
		return RangeAddr{}, err
	}
	if c1 > c2 {
		c1, c2 = c2, c1
	}
	if r1 > r2 {
		r1, r2 = r2, r1
	}
	return RangeAddr{
		Sheet:   sheet,
		FromCol: c1, FromRow: r1,
		ToCol: c2, ToRow: r2,
	}, nil
}

// parseR1C1Cell parses a single R1C1-style cell reference like "R1C1" or "R5C3"
// and returns (col, row). The input is case-insensitive.
func parseR1C1Cell(s string) (col, row int, err error) {
	s = strings.ToUpper(s)
	if len(s) < 4 || s[0] != 'R' {
		return 0, 0, fmt.Errorf("invalid R1C1 reference %q", s)
	}
	cIdx := strings.IndexByte(s[1:], 'C')
	if cIdx < 0 {
		return 0, 0, fmt.Errorf("invalid R1C1 reference %q: missing C", s)
	}
	cIdx++ // adjust for the slice offset
	rowStr := s[1:cIdx]
	colStr := s[cIdx+1:]
	if rowStr == "" || colStr == "" {
		return 0, 0, fmt.Errorf("invalid R1C1 reference %q: empty row or col", s)
	}
	row, err = strconv.Atoi(rowStr)
	if err != nil || row < 1 || row > maxRows {
		return 0, 0, fmt.Errorf("invalid row in R1C1 reference %q", s)
	}
	col, err = strconv.Atoi(colStr)
	if err != nil || col < 1 || col > maxCols {
		return 0, 0, fmt.Errorf("invalid col in R1C1 reference %q", s)
	}
	return col, row, nil
}

// r1c1ToA1 converts an R1C1-style reference string to A1-style.
// Supports single cell (R1C1), ranges (R1C1:R5C3), and optional sheet prefixes.
func r1c1ToA1(ref string) (string, error) {
	// Preserve sheet prefix. splitSheetPrefix is quote-aware so a sheet
	// name quoted like 'Bob's-Sheet' (with an escaped '') does not get
	// mis-split at an embedded '!'.
	prefix, cellPart := splitSheetPrefix(ref)

	// Check if it's a range.
	if colonIdx := strings.IndexByte(cellPart, ':'); colonIdx >= 0 {
		left := cellPart[:colonIdx]
		right := cellPart[colonIdx+1:]
		c1, r1, err := parseR1C1Cell(left)
		if err != nil {
			return "", err
		}
		c2, r2, err := parseR1C1Cell(right)
		if err != nil {
			return "", err
		}
		return prefix + ColNumberToLetters(c1) + strconv.Itoa(r1) + ":" + ColNumberToLetters(c2) + strconv.Itoa(r2), nil
	}

	// Single cell.
	c, r, err := parseR1C1Cell(cellPart)
	if err != nil {
		return "", err
	}
	return prefix + ColNumberToLetters(c) + strconv.Itoa(r), nil
}

// splitSheetPrefix separates a reference string like "Sheet1!A1" or
// "'Sheet Name'!A1" into its sheet-qualifier prefix (including the
// trailing '!') and the cell portion. When the sheet name is quoted,
// embedded '!' characters (and escaped ” quotes) inside the quotes
// are preserved and do not split the reference. If no sheet qualifier
// is present, prefix is empty and rest is the whole input.
func splitSheetPrefix(ref string) (prefix, rest string) {
	if ref == "" {
		return "", ref
	}
	if ref[0] == '\'' {
		for i := 1; i < len(ref); i++ {
			if ref[i] != '\'' {
				continue
			}
			if i+1 < len(ref) && ref[i+1] == '\'' {
				i++ // step over the second '; outer loop increment skips it
				continue
			}
			if i+1 < len(ref) && ref[i+1] == '!' {
				return ref[:i+2], ref[i+2:]
			}
			return "", ref
		}
		return "", ref
	}
	if idx := strings.IndexByte(ref, '!'); idx >= 0 {
		return ref[:idx+1], ref[idx+1:]
	}
	return "", ref
}

func isAllDigits(s string) bool {
	if s == "" {
		return false
	}
	for _, c := range s {
		if c < '0' || c > '9' {
			return false
		}
	}
	return true
}

func isAllLetters(s string) bool {
	if s == "" {
		return false
	}
	for _, c := range s {
		if !((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')) {
			return false
		}
	}
	return true
}

func fnTRANSPOSE(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	v := args[0]
	if v.Type != ValueArray {
		return v, nil
	}

	rows, cols := effectiveArrayBounds(v)
	if rows == 0 {
		return Value{Type: ValueArray, Array: nil}, nil
	}
	if cols == 0 {
		return Value{Type: ValueArray, Array: nil}, nil
	}

	// Transpose: result has cols rows and rows columns.
	result := make([][]Value, cols)
	for c := 0; c < cols; c++ {
		result[c] = make([]Value, rows)
		for r := 0; r < rows; r++ {
			result[c][r] = indexArrayValue(v, r, c)
		}
	}
	return Value{Type: ValueArray, Array: result}, nil
}

// fnUNIQUE implements UNIQUE(array, [by_col], [exactly_once]).
func fnUNIQUE(args []Value) (Value, error) {
	return uniqueCore(args, nil)
}

func evalUNIQUE(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	return ValueToEvalValue(uniqueCoreEval(args, nil)), nil
}

func uniqueCore(args []Value, evalArgs []EvalValue) (Value, error) {
	if len(args) < 1 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}

	src, errVal := normalizeGridShapeArg(args[0], evalArgAt(evalArgs, 0))
	if errVal != nil {
		if errVal.Err == ErrValVALUE {
			return ErrorVal(ErrValCALC), nil
		}
		return *errVal, nil
	}

	// by_col: default FALSE.
	byCol := false
	if len(args) >= 2 {
		bc, e := CoerceNum(legacyArgValue(args[1], evalArgAt(evalArgs, 1)))
		if e != nil {
			return *e, nil
		}
		byCol = bc != 0
	}

	// exactly_once: default FALSE.
	exactlyOnce := false
	if len(args) >= 3 {
		eo, e := CoerceNum(legacyArgValue(args[2], evalArgAt(evalArgs, 2)))
		if e != nil {
			return *e, nil
		}
		exactlyOnce = eo != 0
	}

	// Build a key for each row and track counts / first-seen order.
	type rowEntry struct {
		index int
		key   string
	}
	seen := make(map[string]int) // key → count
	var order []rowEntry
	itemCount := uniqueAxisCount(src, byCol)
	for i := 0; i < itemCount; i++ {
		k := uniqueAxisKey(src, i, byCol)
		seen[k]++
		if seen[k] == 1 {
			order = append(order, rowEntry{index: i, key: k})
		}
	}

	// Collect result rows. Anonymous-array output positions don't have the
	// "truly blank" concept that ranges do — Excel renders blank source
	// rows as empty strings in UNIQUE's spill so COUNTA counts them
	// (COUNT still ignores them, since they aren't numeric). Mirror that
	// by normalising ValueEmpty to StringVal("") in output while keeping
	// the row keyed distinctly via rowKey's "E:" prefix.
	var keep []int
	for _, entry := range order {
		if exactlyOnce && seen[entry.key] != 1 {
			continue
		}
		keep = append(keep, entry.index)
	}

	// If exactly_once filtered everything out, return #CALC!.
	if len(keep) == 0 {
		return ErrorVal(ErrValCALC), nil
	}

	result := uniqueMaterialize(src, keep, byCol)

	// Return: single value, 1D column array, or 2D array.
	return collapseArrayResult(result), nil
}

// rowKey produces a string key for a row of Values, encoding type and value
// so that different types with the same string representation (e.g. 1 vs "1")
// are distinguishable.
func rowKey(row []Value) string {
	var b strings.Builder
	for i, v := range row {
		if i > 0 {
			b.WriteByte('|')
		}
		switch v.Type {
		case ValueEmpty:
			b.WriteString("E:")
		case ValueNumber:
			b.WriteString("N:")
			b.WriteString(strconv.FormatFloat(v.Num, 'g', -1, 64))
		case ValueString:
			// Excel UNIQUE is case-insensitive: "apple" and "Apple" dedup.
			b.WriteString("S:")
			b.WriteString(strings.ToLower(v.Str))
		case ValueBool:
			b.WriteString("B:")
			if v.Bool {
				b.WriteString("1")
			} else {
				b.WriteString("0")
			}
		case ValueError:
			b.WriteString("R:")
			b.WriteString(strconv.Itoa(int(v.Err)))
		default:
			b.WriteString("?:")
		}
	}
	return b.String()
}

// normalizeToGrid converts a Value into a 2D grid.
// Scalars become {{value}}, arrays are used as-is.
func normalizeToGrid(v Value) ([][]Value, *Value) {
	switch v.Type {
	case ValueArray:
		if len(v.Array) == 0 {
			e := ErrorVal(ErrValVALUE)
			return nil, &e
		}
		return v.Array, nil
	case ValueError:
		return nil, &v
	default:
		return [][]Value{{v}}, nil
	}
}

// spilledErrorMatchingGrid returns a Value whose shape matches the given
// grid, with every cell set to err. Used by FILTER when the include argument
// has errors: Excel propagates one error cell per value-row so that COUNTA
// counts them and SUM propagates the first error — returning a bare scalar
// error collapses both downstream consumers to a single cell.
func spilledErrorMatchingGrid(grid [][]Value, err Value) Value {
	if len(grid) == 0 {
		return err
	}
	out := make([][]Value, len(grid))
	for r, row := range grid {
		cols := len(row)
		if cols == 0 {
			cols = 1
		}
		cells := make([]Value, cols)
		for c := range cells {
			cells[c] = err
		}
		out[r] = cells
	}
	if len(out) == 1 && len(out[0]) == 1 {
		return out[0][0]
	}
	return Value{Type: ValueArray, Array: out}
}

// gridDims returns (rows, maxCols) for a 2D grid.
func gridDims(grid [][]Value) (int, int) {
	rows := len(grid)
	cols := 0
	for _, row := range grid {
		if len(row) > cols {
			cols = len(row)
		}
	}
	return rows, cols
}

// fnTAKE implements TAKE(array, rows, [columns]).
// Returns a specified number of contiguous rows or columns from the start or end of an array.
func fnTAKE(args []Value) (Value, error) {
	return callSelectorEval(evalTAKESelector, args)
}

// fnDROP implements DROP(array, rows, [columns]).
// Excludes a specified number of rows or columns from the start or end of an array.
func fnDROP(args []Value) (Value, error) {
	return callSelectorEval(evalDROPSelector, args)
}

// fnEXPAND implements EXPAND(array, rows, [columns], [pad_with]).
// It expands an array to specified dimensions, padding new cells with pad_with
// (default #N/A).
func fnEXPAND(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}

	grid, errVal := normalizeToArrayGrid(args[0])
	if errVal != nil {
		return *errVal, nil
	}
	srcRows, srcCols := grid.rowCount, grid.colCount

	// Parse rows argument.
	targetRows := srcRows
	if args[1].Type != ValueEmpty {
		r, e := CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		targetRows = int(math.Trunc(r))
	}

	// Parse optional columns argument.
	targetCols := srcCols
	if len(args) >= 3 && args[2].Type != ValueEmpty {
		c, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		targetCols = int(math.Trunc(c))
	}

	// Validate dimensions.
	if targetRows <= 0 || targetCols <= 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	if targetRows < srcRows || targetCols < srcCols {
		return ErrorVal(ErrValVALUE), nil
	}

	// Determine pad value.
	pad := Value{Type: ValueError, Err: ErrValNA}
	if len(args) >= 4 {
		pad = args[3]
	}

	// If no expansion needed, return original.
	if targetRows == srcRows && targetCols == srcCols {
		if srcRows == 1 && srcCols == 1 {
			return grid.cell(0, 0), nil
		}
		return Value{Type: ValueArray, Array: grid.matrix()}, nil
	}

	// Build expanded grid.
	result := make([][]Value, targetRows)
	for r := 0; r < targetRows; r++ {
		row := make([]Value, targetCols)
		for c := 0; c < targetCols; c++ {
			if r < srcRows && c < srcCols {
				row[c] = grid.cell(r, c)
			} else {
				row[c] = pad
			}
		}
		result[r] = row
	}
	return Value{Type: ValueArray, Array: result}, nil
}

// fnCHOOSECOLS implements CHOOSECOLS(array, col_num1, [col_num2], ...).
func fnCHOOSECOLS(args []Value) (Value, error) {
	return callSelectorEval(evalCHOOSECOLSSelector, args)
}

// fnCHOOSEROWS implements CHOOSEROWS(array, row_num1, [row_num2], ...).
func fnCHOOSEROWS(args []Value) (Value, error) {
	return callSelectorEval(evalCHOOSEROWSSelector, args)
}

// fnTOCOL implements TOCOL(array, [ignore], [scan_by_column]).
// Returns all values from a 2D array as a single column.
func fnTOCOL(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}

	grid, errVal := normalizeToArrayGrid(args[0])
	if errVal != nil {
		return *errVal, nil
	}

	ignore := 0
	if len(args) >= 2 {
		ig, e := CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		ignore = int(ig)
		if ignore < 0 || ignore > 3 {
			return ErrorVal(ErrValVALUE), nil
		}
	}

	scanByCol := false
	if len(args) >= 3 {
		scanByCol = IsTruthy(args[2])
	}

	flat := flattenArrayGrid(grid, scanByCol, ignore)
	if len(flat) == 0 {
		return ErrorVal(ErrValCALC), nil
	}

	// Build a single-column array (n rows x 1 col).
	result := make([][]Value, len(flat))
	for i, v := range flat {
		result[i] = []Value{v}
	}
	if len(result) == 1 {
		return result[0][0], nil
	}
	return Value{Type: ValueArray, Array: result}, nil
}

// fnTOROW implements TOROW(array, [ignore], [scan_by_column]).
// Returns all values from a 2D array as a single row.
func fnTOROW(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}

	grid, errVal := normalizeToArrayGrid(args[0])
	if errVal != nil {
		return *errVal, nil
	}

	ignore := 0
	if len(args) >= 2 {
		ig, e := CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		ignore = int(ig)
		if ignore < 0 || ignore > 3 {
			return ErrorVal(ErrValVALUE), nil
		}
	}

	scanByCol := false
	if len(args) >= 3 {
		scanByCol = IsTruthy(args[2])
	}

	flat := flattenArrayGrid(grid, scanByCol, ignore)
	if len(flat) == 0 {
		return ErrorVal(ErrValCALC), nil
	}

	// Build a single-row array (1 row x n cols).
	if len(flat) == 1 {
		return flat[0], nil
	}
	return Value{Type: ValueArray, Array: [][]Value{flat}}, nil
}

// flattenArrayGrid flattens a grid into a 1D slice, optionally scanning by
// column and filtering based on ignore flags (0=keep all, 1=ignore blanks,
// 2=ignore errors, 3=ignore blanks and errors).
func flattenArrayGrid(grid arrayGrid, scanByCol bool, ignore int) []Value {
	numCols := grid.colCount
	numRows := grid.rowCount

	var flat []Value
	if scanByCol {
		for c := 0; c < numCols; c++ {
			for r := 0; r < numRows; r++ {
				v := grid.cell(r, c)
				if shouldInclude(v, ignore) {
					flat = append(flat, v)
				}
			}
		}
	} else {
		for r := 0; r < numRows; r++ {
			for c := 0; c < numCols; c++ {
				v := grid.cell(r, c)
				if shouldInclude(v, ignore) {
					flat = append(flat, v)
				}
			}
		}
	}
	return flat
}

// shouldInclude returns true if the value should be kept given the ignore flag.
func shouldInclude(v Value, ignore int) bool {
	switch ignore {
	case 1: // ignore blanks
		return v.Type != ValueEmpty
	case 2: // ignore errors
		return v.Type != ValueError
	case 3: // ignore blanks and errors
		return v.Type != ValueEmpty && v.Type != ValueError
	default:
		return true
	}
}

// fnWRAPROWS implements WRAPROWS(vector, wrap_count, [pad_with]).
// Wraps a row or column vector into a 2D array with wrap_count columns per row.
func fnWRAPROWS(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Flatten input to a 1D vector.
	grid, errVal := normalizeToArrayGrid(args[0])
	if errVal != nil {
		return *errVal, nil
	}
	flat := flattenArrayGrid(grid, false, 0)
	if len(flat) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}

	wrapArg, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	wrapCount := int(wrapArg)
	if wrapCount < 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	padWith := ErrorVal(ErrValNA)
	if len(args) == 3 {
		padWith = args[2]
	}

	// Build 2D array.
	numRows := (len(flat) + wrapCount - 1) / wrapCount
	result := make([][]Value, numRows)
	for i := 0; i < numRows; i++ {
		row := make([]Value, wrapCount)
		for j := 0; j < wrapCount; j++ {
			idx := i*wrapCount + j
			if idx < len(flat) {
				row[j] = flat[idx]
			} else {
				row[j] = padWith
			}
		}
		result[i] = row
	}

	if len(result) == 1 && len(result[0]) == 1 {
		return result[0][0], nil
	}
	return Value{Type: ValueArray, Array: result}, nil
}

// fnWRAPCOLS implements WRAPCOLS(vector, wrap_count, [pad_with]).
// Wraps a row or column vector into a 2D array with wrap_count rows per column.
func fnWRAPCOLS(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Flatten input to a 1D vector.
	grid, errVal := normalizeToArrayGrid(args[0])
	if errVal != nil {
		return *errVal, nil
	}
	flat := flattenArrayGrid(grid, false, 0)
	if len(flat) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}

	wrapArg, e := CoerceNum(args[1])
	if e != nil {
		return *e, nil
	}
	wrapCount := int(wrapArg)
	if wrapCount < 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	padWith := ErrorVal(ErrValNA)
	if len(args) == 3 {
		padWith = args[2]
	}

	// Build 2D array filling column-first.
	numCols := (len(flat) + wrapCount - 1) / wrapCount
	result := make([][]Value, wrapCount)
	for r := 0; r < wrapCount; r++ {
		result[r] = make([]Value, numCols)
		for c := 0; c < numCols; c++ {
			idx := c*wrapCount + r
			if idx < len(flat) {
				result[r][c] = flat[idx]
			} else {
				result[r][c] = padWith
			}
		}
	}

	if len(result) == 1 && len(result[0]) == 1 {
		return result[0][0], nil
	}
	return Value{Type: ValueArray, Array: result}, nil
}

// fnHSTACK implements HSTACK(array1, [array2], ...).
// Horizontally stacks arrays side by side.
func fnHSTACK(args []Value) (Value, error) {
	if len(args) < 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Normalize all arguments to grids and find the max row count.
	grids := make([]arrayGrid, len(args))
	maxRows := 0
	for i, arg := range args {
		g, errVal := normalizeToArrayGrid(arg)
		if errVal != nil {
			return *errVal, nil
		}
		grids[i] = g
		if g.rowCount > maxRows {
			maxRows = g.rowCount
		}
	}

	// Build result by concatenating columns from each grid.
	result := make([][]Value, maxRows)
	for r := 0; r < maxRows; r++ {
		var row []Value
		for _, g := range grids {
			for c := 0; c < g.colCount; c++ {
				if r < g.rowCount {
					if g.rangeOrigin == nil && !g.hasMaterializedCell(r, c) {
						row = append(row, ErrorVal(ErrValNA))
						continue
					}
					row = append(row, g.cell(r, c))
					continue
				}
				row = append(row, ErrorVal(ErrValNA))
			}
		}
		result[r] = row
	}

	if len(result) == 1 && len(result[0]) == 1 {
		return result[0][0], nil
	}
	return Value{Type: ValueArray, Array: result}, nil
}

// fnVSTACK implements VSTACK(array1, [array2], ...).
// Vertically stacks arrays on top of each other.
func fnVSTACK(args []Value) (Value, error) {
	if len(args) < 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Normalize all arguments to grids and find the max column count.
	grids := make([]arrayGrid, len(args))
	maxCols := 0
	for i, arg := range args {
		g, errVal := normalizeToArrayGrid(arg)
		if errVal != nil {
			return *errVal, nil
		}
		grids[i] = g
		if g.colCount > maxCols {
			maxCols = g.colCount
		}
	}

	// Build result by stacking rows from each grid vertically.
	var result [][]Value
	for _, g := range grids {
		for r := 0; r < g.rowCount; r++ {
			row := make([]Value, maxCols)
			for c := 0; c < maxCols; c++ {
				if c < g.colCount {
					if g.rangeOrigin == nil && !g.hasMaterializedCell(r, c) {
						row[c] = ErrorVal(ErrValNA)
						continue
					}
					row[c] = g.cell(r, c)
				} else {
					row[c] = ErrorVal(ErrValNA)
				}
			}
			result = append(result, row)
		}
	}

	if len(result) == 1 && len(result[0]) == 1 {
		return result[0][0], nil
	}
	return Value{Type: ValueArray, Array: result}, nil
}

// transposeGrid transposes a 2D grid of Values.
func transposeGrid(grid [][]Value) [][]Value {
	if len(grid) == 0 {
		return nil
	}
	cols := 0
	for _, row := range grid {
		if len(row) > cols {
			cols = len(row)
		}
	}
	if cols == 0 {
		return nil
	}
	result := make([][]Value, cols)
	for c := 0; c < cols; c++ {
		result[c] = make([]Value, len(grid))
		for r := 0; r < len(grid); r++ {
			if c < len(grid[r]) {
				result[c][r] = grid[r][c]
			} else {
				result[c][r] = EmptyVal()
			}
		}
	}
	return result
}

// fnXMATCH implements XMATCH(lookup_value, lookup_array, [match_mode], [search_mode]).
// It returns the 1-based relative position of an item in an array.
func fnXMATCH(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}

	lookup := args[0]
	if lookup.Type == ValueError {
		return lookup, nil
	}

	arr := args[1]
	if arr.Type == ValueError {
		return arr, nil
	}

	matchMode := 0
	if len(args) >= 3 {
		if args[2].Type == ValueError {
			return args[2], nil
		}
		mm, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		matchMode = int(mm)
	}

	searchMode := 1
	if len(args) >= 4 {
		if args[3].Type == ValueError {
			return args[3], nil
		}
		sm, e := CoerceNum(args[3])
		if e != nil {
			return *e, nil
		}
		searchMode = int(sm)
	}

	// Validate match_mode.
	switch matchMode {
	case 0, -1, 1, 2:
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Validate search_mode.
	switch searchMode {
	case 1, -1, 2, -2:
	default:
		return ErrorVal(ErrValVALUE), nil
	}

	// Flatten lookup_array into a single slice.
	var values []Value
	if arr.Type == ValueArray {
		for _, row := range arr.Array {
			values = append(values, row...)
		}
	} else {
		values = []Value{arr}
	}

	n := len(values)
	if n == 0 {
		return ErrorVal(ErrValNA), nil
	}

	switch matchMode {
	case 0:
		// Exact match.
		return xmatchExact(lookup, values, searchMode), nil

	case 2:
		// Wildcard match.
		return xmatchWildcard(lookup, values, searchMode), nil

	case -1:
		// Exact match or next smallest.
		return xmatchNextSmallest(lookup, values, searchMode), nil

	case 1:
		// Exact match or next largest.
		return xmatchNextLargest(lookup, values, searchMode), nil
	}

	return ErrorVal(ErrValVALUE), nil
}

// xmatchExact performs exact match (match_mode 0) with the given search_mode.
func xmatchExact(lookup Value, values []Value, searchMode int) Value {
	n := len(values)
	switch searchMode {
	case 1:
		// First-to-last.
		for i := 0; i < n; i++ {
			if CompareValuesExact(values[i], lookup) == 0 {
				return NumberVal(float64(i + 1))
			}
		}
	case -1:
		// Last-to-first.
		for i := n - 1; i >= 0; i-- {
			if CompareValuesExact(values[i], lookup) == 0 {
				return NumberVal(float64(i + 1))
			}
		}
	case 2:
		// Binary search ascending.
		lo, hi := 0, n-1
		for lo <= hi {
			mid := lo + (hi-lo)/2
			cmp := CompareValuesExact(values[mid], lookup)
			if cmp == 0 {
				return NumberVal(float64(mid + 1))
			} else if cmp < 0 {
				lo = mid + 1
			} else {
				hi = mid - 1
			}
		}
	case -2:
		// Binary search descending.
		lo, hi := 0, n-1
		for lo <= hi {
			mid := lo + (hi-lo)/2
			cmp := CompareValuesExact(values[mid], lookup)
			if cmp == 0 {
				return NumberVal(float64(mid + 1))
			} else if cmp > 0 {
				lo = mid + 1
			} else {
				hi = mid - 1
			}
		}
	}
	return ErrorVal(ErrValNA)
}

// xmatchWildcard performs wildcard match (match_mode 2) with the given search_mode.
func xmatchWildcard(lookup Value, values []Value, searchMode int) Value {
	pattern := ValueToString(lookup)
	re, err := patternToRegexp(pattern)
	if err != nil {
		return ErrorVal(ErrValVALUE)
	}
	anchored, err := regexp.Compile("(?i)^" + re.String() + "$")
	if err != nil {
		return ErrorVal(ErrValVALUE)
	}

	n := len(values)
	switch searchMode {
	case 1, 2:
		// First-to-last (binary search not meaningful for wildcard, use linear).
		for i := 0; i < n; i++ {
			if anchored.MatchString(ValueToString(values[i])) {
				return NumberVal(float64(i + 1))
			}
		}
	case -1, -2:
		// Last-to-first.
		for i := n - 1; i >= 0; i-- {
			if anchored.MatchString(ValueToString(values[i])) {
				return NumberVal(float64(i + 1))
			}
		}
	}
	return ErrorVal(ErrValNA)
}

// xmatchNextSmallest performs exact match or next smallest (match_mode -1).
func xmatchNextSmallest(lookup Value, values []Value, searchMode int) Value {
	n := len(values)
	switch searchMode {
	case 1:
		// Linear first-to-last: find best match <= lookup.
		best := -1
		for i := 0; i < n; i++ {
			if values[i].Type == ValueEmpty {
				continue
			}
			cmp := CompareValues(values[i], lookup)
			if cmp == 0 {
				return NumberVal(float64(i + 1))
			}
			if cmp < 0 {
				if best < 0 || CompareValues(values[i], values[best]) > 0 {
					best = i
				}
			}
		}
		if best >= 0 {
			return NumberVal(float64(best + 1))
		}
	case -1:
		// Linear last-to-first: find best match <= lookup.
		best := -1
		for i := n - 1; i >= 0; i-- {
			if values[i].Type == ValueEmpty {
				continue
			}
			cmp := CompareValues(values[i], lookup)
			if cmp == 0 {
				return NumberVal(float64(i + 1))
			}
			if cmp < 0 {
				if best < 0 || CompareValues(values[i], values[best]) > 0 {
					best = i
				}
			}
		}
		if best >= 0 {
			return NumberVal(float64(best + 1))
		}
	case 2:
		// Binary search ascending: data sorted ascending.
		lo, hi := 0, n-1
		result := -1
		for lo <= hi {
			mid := lo + (hi-lo)/2
			cmp := CompareValues(values[mid], lookup)
			if cmp == 0 {
				return NumberVal(float64(mid + 1))
			} else if cmp < 0 {
				result = mid
				lo = mid + 1
			} else {
				hi = mid - 1
			}
		}
		if result >= 0 {
			return NumberVal(float64(result + 1))
		}
	case -2:
		// Binary search descending: data sorted descending.
		// We want the largest value <= lookup (next smallest).
		// In descending order, values decrease left-to-right.
		lo, hi := 0, n-1
		result := -1
		for lo <= hi {
			mid := lo + (hi-lo)/2
			cmp := CompareValues(values[mid], lookup)
			if cmp == 0 {
				return NumberVal(float64(mid + 1))
			} else if cmp > 0 {
				// values[mid] > lookup → need smaller values → go right
				lo = mid + 1
			} else {
				// values[mid] < lookup → candidate; look left for closer
				result = mid
				hi = mid - 1
			}
		}
		if result >= 0 {
			return NumberVal(float64(result + 1))
		}
	}
	return ErrorVal(ErrValNA)
}

// xmatchNextLargest performs exact match or next largest (match_mode 1).
func xmatchNextLargest(lookup Value, values []Value, searchMode int) Value {
	n := len(values)
	switch searchMode {
	case 1:
		// Linear first-to-last: find best match >= lookup.
		best := -1
		for i := 0; i < n; i++ {
			if values[i].Type == ValueEmpty {
				continue
			}
			cmp := CompareValues(values[i], lookup)
			if cmp == 0 {
				return NumberVal(float64(i + 1))
			}
			if cmp > 0 {
				if best < 0 || CompareValues(values[i], values[best]) < 0 {
					best = i
				}
			}
		}
		if best >= 0 {
			return NumberVal(float64(best + 1))
		}
	case -1:
		// Linear last-to-first: find best match >= lookup.
		best := -1
		for i := n - 1; i >= 0; i-- {
			if values[i].Type == ValueEmpty {
				continue
			}
			cmp := CompareValues(values[i], lookup)
			if cmp == 0 {
				return NumberVal(float64(i + 1))
			}
			if cmp > 0 {
				if best < 0 || CompareValues(values[i], values[best]) < 0 {
					best = i
				}
			}
		}
		if best >= 0 {
			return NumberVal(float64(best + 1))
		}
	case 2:
		// Binary search ascending: data sorted ascending.
		lo, hi := 0, n-1
		result := -1
		for lo <= hi {
			mid := lo + (hi-lo)/2
			cmp := CompareValues(values[mid], lookup)
			if cmp == 0 {
				return NumberVal(float64(mid + 1))
			} else if cmp < 0 {
				lo = mid + 1
			} else {
				result = mid
				hi = mid - 1
			}
		}
		if result >= 0 {
			return NumberVal(float64(result + 1))
		}
	case -2:
		// Binary search descending: data sorted descending.
		// We want the smallest value >= lookup (next largest).
		// In descending order, values decrease left-to-right.
		lo, hi := 0, n-1
		result := -1
		for lo <= hi {
			mid := lo + (hi-lo)/2
			cmp := CompareValues(values[mid], lookup)
			if cmp == 0 {
				return NumberVal(float64(mid + 1))
			} else if cmp > 0 {
				// values[mid] > lookup → candidate; look right for closer
				result = mid
				lo = mid + 1
			} else {
				// values[mid] < lookup → need larger values → go left
				hi = mid - 1
			}
		}
		if result >= 0 {
			return NumberVal(float64(result + 1))
		}
	}
	return ErrorVal(ErrValNA)
}

// fnHyperlink implements HYPERLINK(link_location, [friendly_name]).
// It returns the friendly_name (or link_location if omitted) as the display value.
func fnHyperlink(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Evaluate link_location — propagate errors.
	loc := args[0]
	if loc.Type == ValueError {
		return loc, nil
	}

	// If friendly_name is provided, return it as-is (propagating errors).
	if len(args) == 2 {
		fn := args[1]
		if fn.Type == ValueError {
			return fn, nil
		}
		return fn, nil
	}

	// No friendly_name — return link_location coerced to string.
	return StringVal(ValueToString(loc)), nil
}

// fnOFFSET implements OFFSET(reference, rows, cols, [height], [width]).
// It returns a reference to a range offset from the given reference.
func fnOFFSET(args []Value, ctx *EvalContext) (Value, error) {
	return callLegacyRefEval(evalOFFSET, args, ctx)
}
