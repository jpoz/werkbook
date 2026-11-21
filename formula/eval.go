package formula

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

const (
	maxRows = 1048576 // maximum rows in a worksheet
	maxCols = 16384   // maximum columns in a worksheet (XFD)
)

// CellResolver abstracts cell/range lookups so the VM has no dependency on Sheet.
type CellResolver interface {
	GetCellValue(addr CellAddr) Value
	GetRangeValues(addr RangeAddr) [][]Value
}

// DefinedNameResolver is an optional CellResolver extension used by
// ref-producing functions like INDIRECT to resolve runtime named ranges.
type DefinedNameResolver interface {
	ResolveDefinedNameValue(name, scopeSheet string) (Value, bool)
}

// EvalContext provides context about the current evaluation environment.
type EvalContext struct {
	CurrentCol     int
	CurrentRow     int
	CurrentSheet   string
	IsArrayFormula bool         // true for CSE (Ctrl+Shift+Enter) array formulas
	InheritedArray bool         // internal: true while a call executes inside inherited array context
	Date1904       bool         // true if the workbook uses the 1904 date system
	Resolver       CellResolver // the active resolver; used by SUBTOTAL to inspect cells
	Tracer         EvalTracer   // optional; nil means no tracing
}

// SubtotalChecker is an optional interface that a CellResolver may implement
// to allow SUBTOTAL to skip cells that themselves contain SUBTOTAL formulas,
// preventing double-counting of nested subtotals.
type SubtotalChecker interface {
	IsSubtotalCell(sheet string, col, row int) bool
}

// HiddenRowChecker is an optional interface that a CellResolver may implement
// to allow SUBTOTAL to exclude hidden rows. IsRowHidden returns true if the
// row is hidden for any reason. IsRowFilteredByAutoFilter returns true only
// if the row is hidden AND falls within a table with an active autoFilter
// (used by SUBTOTAL function numbers 1-11 to exclude auto-filtered rows).
type HiddenRowChecker interface {
	IsRowHidden(sheet string, row int) bool
	IsRowFilteredByAutoFilter(sheet string, row int) bool
}

// SheetListProvider is an optional interface that a CellResolver may implement
// to support 3D sheet references (e.g. Sheet2:Sheet5!A1). It returns the
// ordered list of all sheet names in the workbook.
type SheetListProvider interface {
	GetSheetNames() []string
}

// FormulaIntrospector is an optional interface that a CellResolver may
// implement to support ISFORMULA and FORMULATEXT. It allows the formula
// engine to check whether a cell contains a formula and retrieve its text.
type FormulaIntrospector interface {
	// HasFormula reports whether the cell at the given address contains a formula.
	HasFormula(sheet string, col, row int) bool
	// GetFormulaText returns the formula text (without leading '=') for the
	// cell at the given address, or "" if the cell has no formula.
	GetFormulaText(sheet string, col, row int) string
}

// FormulaArrayEvaluator is an optional interface that a CellResolver may
// implement to support ANCHORARRAY. It evaluates the formula in the given
// cell and returns the full array result (not just the top-left element).
// If the cell has no formula or the formula does not produce an array,
// it returns the cell's scalar value wrapped in a 1x1 array.
type FormulaArrayEvaluator interface {
	// EvalCellFormula evaluates the formula in the cell at (sheet, col, row)
	// and returns the full result. For dynamic array formulas, this returns
	// the complete ValueArray rather than just the anchor cell's value.
	EvalCellFormula(sheet string, col, row int) Value
}

// Eval executes a compiled formula and returns the result.
func Eval(cf *CompiledFormula, resolver CellResolver, ctx *EvalContext) (Value, error) {
	return evalWithParams(cf, resolver, ctx, nil)
}

func evalWithParams(cf *CompiledFormula, resolver CellResolver, ctx *EvalContext, params []Value) (Value, error) {
	stack := make([]Value, 0, 16)
	arrayCtxDepth := 0 // >0 means we're inside an array-forcing function's arguments

	push := func(v Value) { stack = append(stack, v) }
	pop := func() (Value, error) {
		if len(stack) == 0 {
			return Value{}, fmt.Errorf("stack underflow")
		}
		v := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		return v, nil
	}

	tracer := ctx != nil && ctx.Tracer != nil
	for instIdx, inst := range cf.Code {
		_ = instIdx
		switch inst.Op {
		case OpPushNum:
			push(cf.Consts[inst.Operand])
		case OpPushStr:
			push(cf.Consts[inst.Operand])
		case OpPushBool:
			push(BoolVal(inst.Operand != 0))
		case OpPushError:
			push(ErrorVal(ErrorValue(inst.Operand)))
		case OpPushEmpty:
			push(EmptyVal())

		case OpLoadCell:
			addr := cf.Refs[inst.Operand]
			v := resolver.GetCellValue(addr)
			v.FromCell = true
			push(v)

		case OpLoadRange:
			addr := cf.Ranges[inst.Operand]
			isFullCol := addr.FromRow == 1 && addr.ToRow >= maxRows
			isFullRow := addr.FromCol == 1 && addr.ToCol >= maxCols
			// Implicit intersection: when a full-column or full-row range is
			// used in a non-array formula, reduce to the single cell at the
			// formula's own row/column rather than loading the entire range.
			// Skip implicit intersection when inside an array-forcing function
			// (arrayCtxDepth > 0), since those functions need the full range.
			if ctx != nil && !ctx.IsArrayFormula && arrayCtxDepth == 0 {
				if isFullCol && addr.FromCol == addr.ToCol && ctx.CurrentRow >= addr.FromRow {
					// Full-column ref like F:F → intersect at current row
					push(resolver.GetCellValue(CellAddr{
						Sheet: addr.Sheet,
						Col:   addr.FromCol,
						Row:   ctx.CurrentRow,
					}))
					continue
				}
				if isFullRow && addr.FromRow == addr.ToRow && ctx.CurrentCol >= addr.FromCol {
					// Full-row ref like 1:1 → intersect at current column
					push(resolver.GetCellValue(CellAddr{
						Sheet: addr.Sheet,
						Col:   ctx.CurrentCol,
						Row:   addr.FromRow,
					}))
					continue
				}
			}
			if !isFullCol && !isFullRow &&
				RangeCellCountExceedsLimit(addr.ToRow-addr.FromRow+1, addr.ToCol-addr.FromCol+1) {
				push(ErrorVal(ErrValREF))
				continue
			}
			rows := resolver.GetRangeValues(addr)
			if isRangeOverflowMatrix(rows) {
				push(ErrorVal(ErrValREF))
				continue
			}
			rows = normalizeResolverRangeRows(addr, rows)
			origin := addr // capture for the Value
			push(Value{Type: ValueArray, Array: rows, RangeOrigin: &origin})

		case OpLoad3DRange:
			addr := cf.Ranges[inst.Operand]
			// Resolve 3D sheet reference: collect values from all sheets
			// between addr.Sheet and addr.SheetEnd.
			slp, ok := resolver.(SheetListProvider)
			if !ok {
				push(ErrorVal(ErrValREF))
				continue
			}
			sheets := resolveSheetRange(slp.GetSheetNames(), addr.Sheet, addr.SheetEnd)
			if len(sheets) == 0 {
				push(ErrorVal(ErrValREF))
				continue
			}
			// Collect all values from all sheets into a flat array.
			var allValues [][]Value
			for _, sheetName := range sheets {
				singleAddr := RangeAddr{
					Sheet:   sheetName,
					FromCol: addr.FromCol, FromRow: addr.FromRow,
					ToCol: addr.ToCol, ToRow: addr.ToRow,
				}
				sheetRows := resolver.GetRangeValues(singleAddr)
				allValues = append(allValues, sheetRows...)
			}
			if len(allValues) == 0 {
				push(EmptyVal())
			} else {
				push(Value{Type: ValueArray, Array: allValues})
			}

		case OpLoadCellRef:
			addr := cf.Refs[inst.Operand]
			// Encode col and row into Num: col + row*100_000.
			// Max col = 16384 < 100_000, max row = 1_048_576, product < 2^53.
			// Store the sheet name in Str so cross-sheet refs are available.
			sheet := addr.Sheet
			if sheet == "" && ctx != nil {
				sheet = ctx.CurrentSheet
			}
			push(Value{Type: ValueRef, Num: float64(addr.Col + addr.Row*100_000), Str: sheet})

		case OpAdd, OpSub, OpMul, OpDiv, OpPow:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			if ctx != nil && !ctx.IsArrayFormula && arrayCtxDepth == 0 {
				_, aWasRange := legacyArrayRef(a)
				_, bWasRange := legacyArrayRef(b)
				a = implicitIntersect(a, ctx)
				b = implicitIntersect(b, ctx)
				// If one side was a range and got intersected to a scalar but
				// the other is still an anonymous array (e.g. a LET-bound
				// SCAN/SEQUENCE result), Excel intersects it to its top-left
				// element so both operands match. Leaving it as an array
				// would cause lop-sided broadcasts like A1:A8 - {rm}.
				if aWasRange && a.Type != ValueArray && b.Type == ValueArray {
					if _, ok := legacyArrayRef(b); !ok {
						b = arrayTopLeft(b)
					}
				}
				if bWasRange && b.Type != ValueArray && a.Type == ValueArray {
					if _, ok := legacyArrayRef(a); !ok {
						a = arrayTopLeft(a)
					}
				}
			}
			var fn func(float64, float64) Value
			switch inst.Op {
			case OpAdd:
				fn = func(an, bn float64) Value { return NumberVal(an + bn) }
			case OpSub:
				fn = func(an, bn float64) Value { return NumberVal(an - bn) }
			case OpMul:
				fn = func(an, bn float64) Value { return NumberVal(an * bn) }
			case OpDiv:
				fn = func(an, bn float64) Value {
					if bn == 0 {
						return ErrorVal(ErrValDIV0)
					}
					return NumberVal(an / bn)
				}
			case OpPow:
				fn = func(an, bn float64) Value {
					result := math.Pow(an, bn)
					if math.IsNaN(result) || math.IsInf(result, 0) {
						return ErrorVal(ErrValNUM)
					}
					return NumberVal(result)
				}
			}
			result := binaryArith(a, b, fn)
			if tracer {
				ctx.Tracer.OnBinaryOp(instIdx, inst.Op, a, b, result)
			}
			push(result)

		case OpNeg:
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			if a.Type == ValueArray {
				push(LiftUnary(a, func(v Value) Value {
					n, e := CoerceNum(v)
					if e != nil {
						return *e
					}
					return NumberVal(-n)
				}))
			} else {
				an, ae := CoerceNum(a)
				if ae != nil {
					push(*ae)
				} else {
					push(NumberVal(-an))
				}
			}

		case OpPercent:
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			if a.Type == ValueArray {
				push(LiftUnary(a, func(v Value) Value {
					n, e := CoerceNum(v)
					if e != nil {
						return *e
					}
					return NumberVal(n / 100)
				}))
			} else {
				an, ae := CoerceNum(a)
				if ae != nil {
					push(*ae)
				} else {
					push(NumberVal(an / 100))
				}
			}

		case OpConcat:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			// Error propagation: if either operand is an error, return that error.
			if a.Type == ValueError {
				push(a)
			} else if b.Type == ValueError {
				push(b)
			} else {
				push(StringVal(ValueToString(a) + ValueToString(b)))
			}

		case OpEq:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			{
				r := binaryCompare(a, b, func(c int) bool { return c == 0 })
				if tracer {
					ctx.Tracer.OnBinaryOp(instIdx, inst.Op, a, b, r)
				}
				push(r)
			}

		case OpNe:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			{
				r := binaryCompare(a, b, func(c int) bool { return c != 0 })
				if tracer {
					ctx.Tracer.OnBinaryOp(instIdx, inst.Op, a, b, r)
				}
				push(r)
			}

		case OpLt:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			{
				r := binaryCompare(a, b, func(c int) bool { return c < 0 })
				if tracer {
					ctx.Tracer.OnBinaryOp(instIdx, inst.Op, a, b, r)
				}
				push(r)
			}

		case OpLe:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			{
				r := binaryCompare(a, b, func(c int) bool { return c <= 0 })
				if tracer {
					ctx.Tracer.OnBinaryOp(instIdx, inst.Op, a, b, r)
				}
				push(r)
			}

		case OpGt:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			{
				r := binaryCompare(a, b, func(c int) bool { return c > 0 })
				if tracer {
					ctx.Tracer.OnBinaryOp(instIdx, inst.Op, a, b, r)
				}
				push(r)
			}

		case OpGe:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			{
				r := binaryCompare(a, b, func(c int) bool { return c >= 0 })
				if tracer {
					ctx.Tracer.OnBinaryOp(instIdx, inst.Op, a, b, r)
				}
				push(r)
			}

		case OpCall:
			funcID := int((inst.Operand >> 8) & funcIDMask)
			inheritedArrayCtx := inst.Operand&callFlagInheritedArrayCtx != 0
			argc := int(inst.Operand & 0xFF)
			if argc > len(stack) {
				return Value{}, fmt.Errorf("stack underflow in function call")
			}
			args := make([]Value, argc)
			copy(args, stack[len(stack)-argc:])
			stack = stack[:len(stack)-argc]

			// Legacy implicit intersection for scalar-lifted functions.
			// When a function in elementWiseCallFuncs is invoked in true
			// scalar context — not CSE, not inside any array-forcing caller
			// (even a suspended one like IF inside SUMPRODUCT) — Excel
			// applies implicit intersection to each range argument instead
			// of broadcasting. Example: TEXTJOIN(",", FALSE,
			// SUBSTITUTE(A1:A4, …)) collapses SUBSTITUTE to the row-aligned
			// cell rather than lifting over the whole range.
			//
			// The inheritedArrayCtx flag (set at compile time when the call
			// site was inside an array-forcing function) suppresses this
			// gate so broadcast behaviour still applies inside IF arms
			// inside SUMPRODUCT, etc.
			if ctx != nil && !ctx.IsArrayFormula && arrayCtxDepth == 0 &&
				!inheritedArrayCtx &&
				funcID >= 0 && funcID < len(idToName) &&
				functionNeedsLegacyElementwisePreIntersect(idToName[funcID]) {
				for i := range args {
					if _, ok := legacyArrayRef(args[i]); ok {
						args[i] = implicitIntersect(args[i], ctx)
					}
				}
			}

			restoreInheritedArray := false
			prevInheritedArray := false
			if ctx != nil {
				restoreInheritedArray = true
				prevInheritedArray = ctx.InheritedArray
				ctx.InheritedArray = inheritedArrayCtx
			}
			result, err := CallFunc(funcID, args, ctx)
			if restoreInheritedArray {
				ctx.InheritedArray = prevInheritedArray
			}
			if err != nil {
				return Value{}, err
			}
			if tracer {
				name := ""
				if funcID >= 0 && funcID < len(idToName) {
					name = idToName[funcID]
				}
				ctx.Tracer.OnCallFunc(instIdx, name, args, result)
			}
			push(result)

		case OpMakeArray:
			rows := int(inst.Operand >> 16)
			cols := int(inst.Operand & 0xFFFF)
			total := rows * cols
			if total > len(stack) {
				return Value{}, fmt.Errorf("stack underflow in array construction")
			}
			elems := make([]Value, total)
			copy(elems, stack[len(stack)-total:])
			stack = stack[:len(stack)-total]

			arr := make([][]Value, rows)
			for r := 0; r < rows; r++ {
				arr[r] = elems[r*cols : (r+1)*cols]
			}
			push(Value{Type: ValueArray, Array: arr})

		case OpEnterArrayCtx:
			arrayCtxDepth++

		case OpLeaveArrayCtx:
			if arrayCtxDepth > 0 {
				arrayCtxDepth--
			}

		case OpRefResultToBool:
			// Used by ISREF wrapping ref-returning functions (e.g. INDIRECT).
			// A ref-returning function produces a reference on success and an
			// error on failure.  ISREF should return TRUE for non-error results
			// and FALSE for errors.
			v, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(BoolVal(v.Type != ValueError))

		case OpImplicitIntersect:
			v, err := pop()
			if err != nil {
				return Value{}, err
			}
			// Scalarize arrays for explicit @/SINGLE and the legacy IFERROR/
			// IFNA compatibility path. CSE array formulas opt out.
			if ctx != nil && !ctx.IsArrayFormula && v.Type == ValueArray {
				v = explicitIntersect(v, ctx)
			}
			push(v)

		case OpImplicitIntersectRefOnly:
			v, err := pop()
			if err != nil {
				return Value{}, err
			}
			// Legacy IFERROR/IFNA compatibility: intersect range-backed
			// arrays so SUMPRODUCT(IFERROR(range, …), …) still collapses
			// IFERROR to the row-aligned cell, but leave anonymous arrays
			// (SEQUENCE/FILTER/MAP output used as an IFERROR fallback)
			// alone so their dynamic-array shape reaches the wrapping
			// function unchanged.
			if ctx != nil && !ctx.IsArrayFormula {
				if _, ok := legacyArrayRef(v); ok {
					v = implicitIntersect(v, ctx)
				}
			}
			push(v)

		case OpIntersect:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			result := rangeIntersect(a, b, resolver, ctx, arrayCtxDepth > 0)
			// In non-array contexts a multi-cell intersection is subject to
			// implicit intersection based on the formula position.
			if arrayCtxDepth == 0 {
				if _, ok := legacyArrayRef(result); ok {
					result = implicitIntersect(result, ctx)
				}
			}
			push(result)

		case OpUnion:
			n := int(inst.Operand)
			if n > len(stack) {
				return Value{}, fmt.Errorf("stack underflow in union")
			}
			areas := make([]Value, n)
			copy(areas, stack[len(stack)-n:])
			stack = stack[:len(stack)-n]
			push(unionAreas(areas))

		case OpBuildRange:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(buildRangeFromRefs(a, b, resolver, ctx))

		case OpLoadParam:
			if params == nil || int(inst.Operand) >= len(params) {
				return Value{}, fmt.Errorf("parameter index %d out of range", inst.Operand)
			}
			push(params[inst.Operand])

		case OpReduce:
			subIdx := int(inst.Operand)
			if subIdx >= len(cf.SubFormulas) {
				return Value{}, fmt.Errorf("sub-formula index %d out of range", subIdx)
			}
			subFormula := cf.SubFormulas[subIdx]

			// Pop array, then initial value
			arr, err := pop()
			if err != nil {
				return Value{}, err
			}
			initialVal, err := pop()
			if err != nil {
				return Value{}, err
			}

			// Flatten row-major using shared array-grid semantics so
			// ref-derived arrays keep their logical shape.
			elements := valueFlattenRowMajor(arr)

			// Determine starting accumulator
			acc := initialVal
			startIdx := 0
			if acc.Type == ValueEmpty {
				if len(elements) == 0 {
					push(ErrorVal(ErrValCALC))
					continue
				}
				acc = elements[0]
				startIdx = 1
			}

			// If array is empty and we have initial value, return it
			if len(elements) == 0 {
				push(acc)
				continue
			}

			// Fold: for each element, call lambda(accumulator, element)
			paramVals := make([]Value, 2)
			for i := startIdx; i < len(elements); i++ {
				paramVals[0] = acc
				paramVals[1] = elements[i]
				result, err := evalWithParams(subFormula, resolver, ctx, paramVals)
				if err != nil {
					return Value{}, err
				}
				// If lambda returns an error, propagate immediately
				if result.Type == ValueError {
					acc = result
					break
				}
				acc = result
			}
			push(acc)

		case OpScan:
			subIdx := int(inst.Operand)
			if subIdx >= len(cf.SubFormulas) {
				return Value{}, fmt.Errorf("sub-formula index %d out of range", subIdx)
			}
			subFormula := cf.SubFormulas[subIdx]

			// Pop array, then initial value
			arr, err := pop()
			if err != nil {
				return Value{}, err
			}
			initialVal, err := pop()
			if err != nil {
				return Value{}, err
			}

			// Get array dimensions for output shape
			var scanRows, scanCols int
			if arr.Type == ValueArray {
				scanRows, scanCols = arrayOpBounds(arr)
			} else {
				scanRows, scanCols = 1, 1
			}

			// Handle empty array
			if scanRows == 0 || scanCols == 0 {
				if initialVal.Type == ValueEmpty {
					push(ErrorVal(ErrValCALC))
				} else {
					push(Value{Type: ValueArray, Array: nil})
				}
				continue
			}

			// Build output array with same shape as input
			scanResult := newValueMatrix(scanRows, scanCols)
			acc := initialVal
			paramVals := make([]Value, 2)
			first := true

			for i := 0; i < scanRows; i++ {
				for j := 0; j < scanCols; j++ {
					elem := ArrayElement(arr, i, j)

					if acc.Type == ValueEmpty && first {
						// No initial value: first element becomes accumulator
						acc = elem
						scanResult[i][j] = acc
						first = false
						continue
					}
					first = false

					// If accumulator is an error, propagate to remaining cells
					if acc.Type == ValueError {
						scanResult[i][j] = acc
						continue
					}

					paramVals[0] = acc
					paramVals[1] = elem
					res, err := evalWithParams(subFormula, resolver, ctx, paramVals)
					if err != nil {
						return Value{}, err
					}
					acc = res
					scanResult[i][j] = acc
				}
			}

			push(Value{Type: ValueArray, Array: scanResult})

		case OpMap:
			subIdx := int(inst.Operand >> 8)
			numArrays := int(inst.Operand & 0xFF)
			if subIdx >= len(cf.SubFormulas) {
				return Value{}, fmt.Errorf("sub-formula index %d out of range", subIdx)
			}
			subFormula := cf.SubFormulas[subIdx]

			// Pop arrays from stack (in reverse order since last pushed is on top)
			arrays := make([]Value, numArrays)
			for i := numArrays - 1; i >= 0; i-- {
				v, err := pop()
				if err != nil {
					return Value{}, err
				}
				arrays[i] = v
			}

			// Determine output dimensions (max rows x max cols across all arrays)
			rows, cols := 1, 1
			for _, arr := range arrays {
				if arr.Type == ValueArray {
					r, c := arrayOpBounds(arr)
					if r > rows {
						rows = r
					}
					if c > cols {
						cols = c
					}
				}
			}

			// Precompute per-array bounds to avoid O(rows) scans per element.
			type arrBounds struct{ rows, cols int }
			arrBoundsCache := make([]arrBounds, len(arrays))
			for k, arr := range arrays {
				r, c := arrayOpBoundsOrScalar(arr)
				arrBoundsCache[k] = arrBounds{r, c}
			}

			// For each element position, bind params and eval the sub-formula
			result := newValueMatrix(rows, cols)
			paramVals := make([]Value, numArrays)
			for i := 0; i < rows; i++ {
				for j := 0; j < cols; j++ {
					for k, arr := range arrays {
						paramVals[k] = arrayElementDirect(arr, arrBoundsCache[k].rows, arrBoundsCache[k].cols, i, j)
					}
					cellResult, err := evalWithParams(subFormula, resolver, ctx, paramVals)
					if err != nil {
						return Value{}, err
					}
					result[i][j] = cellResult
				}
			}
			push(Value{Type: ValueArray, Array: result})

		case OpByRow:
			subIdx := int(inst.Operand)
			if subIdx >= len(cf.SubFormulas) {
				return Value{}, fmt.Errorf("sub-formula index %d out of range", subIdx)
			}
			subFormula := cf.SubFormulas[subIdx]

			// Pop array
			arr, err := pop()
			if err != nil {
				return Value{}, err
			}

			// Determine dimensions
			var byrowRows int
			if arr.Type == ValueArray {
				byrowRows, _ = arrayOpBounds(arr)
			} else {
				byrowRows = 1
			}

			// For each row, create a 1-row array and call lambda
			byrowResult := make([][]Value, byrowRows)
			byrowParamVals := make([]Value, 1)
			for i := 0; i < byrowRows; i++ {
				byrowParamVals[0] = valueProjectRowArray(arr, i)
				res, err := evalWithParams(subFormula, resolver, ctx, byrowParamVals)
				if err != nil {
					return Value{}, err
				}

				// Per Excel: BYROW returns #VALUE! if the lambda produces
				// anything but a single value. A 1x1 array is still scalar.
				if res.Type == ValueArray {
					if r, c := effectiveArrayBounds(res); r == 1 && c == 1 {
						res = ArrayElement(res, 0, 0)
					} else {
						res = ErrorVal(ErrValVALUE)
					}
				}

				byrowResult[i] = []Value{res}
			}

			push(Value{Type: ValueArray, Array: byrowResult})

		case OpMakeArrayLambda:
			subIdx := int(inst.Operand)
			if subIdx >= len(cf.SubFormulas) {
				return Value{}, fmt.Errorf("sub-formula index %d out of range", subIdx)
			}
			subFormula := cf.SubFormulas[subIdx]

			// Pop cols, then rows
			colsVal, err := pop()
			if err != nil {
				return Value{}, err
			}
			rowsVal, err := pop()
			if err != nil {
				return Value{}, err
			}

			// Coerce to numbers
			rowsNum, re := CoerceNum(rowsVal)
			if re != nil {
				push(*re)
				continue
			}
			colsNum, ce := CoerceNum(colsVal)
			if ce != nil {
				push(*ce)
				continue
			}

			// Must be positive integers
			maRows := int(rowsNum)
			maCols := int(colsNum)
			if maRows < 1 || maCols < 1 {
				push(ErrorVal(ErrValVALUE))
				continue
			}

			// Build array by calling lambda(row, col) with 1-based indices
			maResult := newValueMatrix(maRows, maCols)
			maParamVals := make([]Value, 2)
			for i := 0; i < maRows; i++ {
				for j := 0; j < maCols; j++ {
					maParamVals[0] = NumberVal(float64(i + 1)) // 1-based row
					maParamVals[1] = NumberVal(float64(j + 1)) // 1-based col
					res, evalErr := evalWithParams(subFormula, resolver, ctx, maParamVals)
					if evalErr != nil {
						return Value{}, evalErr
					}
					maResult[i][j] = res
				}
			}

			push(Value{Type: ValueArray, Array: maResult})

		case OpByCol:
			subIdx := int(inst.Operand)
			if subIdx >= len(cf.SubFormulas) {
				return Value{}, fmt.Errorf("sub-formula index %d out of range", subIdx)
			}
			subFormula := cf.SubFormulas[subIdx]

			// Pop array
			arr, err := pop()
			if err != nil {
				return Value{}, err
			}

			// Determine dimensions
			var bycolCols int
			if arr.Type == ValueArray {
				_, bycolCols = arrayOpBounds(arr)
			} else {
				bycolCols = 1
			}

			// For each column, create a column vector (rows x 1) and call lambda
			bycolResult := make([][]Value, 1) // single row
			bycolResult[0] = make([]Value, bycolCols)
			bycolParamVals := make([]Value, 1)

			for j := 0; j < bycolCols; j++ {
				bycolParamVals[0] = valueProjectColArray(arr, j)
				res, err := evalWithParams(subFormula, resolver, ctx, bycolParamVals)
				if err != nil {
					return Value{}, err
				}

				// Per Excel: BYCOL returns #VALUE! if the lambda produces
				// anything but a single value. A 1x1 array is still scalar.
				if res.Type == ValueArray {
					if r, c := effectiveArrayBounds(res); r == 1 && c == 1 {
						res = ArrayElement(res, 0, 0)
					} else {
						res = ErrorVal(ErrValVALUE)
					}
				}

				bycolResult[0][j] = res
			}

			push(Value{Type: ValueArray, Array: bycolResult})

		default:
			return Value{}, fmt.Errorf("unknown opcode %d", inst.Op)
		}
	}

	if len(stack) != 1 {
		return Value{}, fmt.Errorf("expected 1 value on stack, got %d", len(stack))
	}
	return stack[0], nil
}

// CoerceNum converts a Value to float64 for arithmetic.
// Returns the number and nil on success, or 0 and a pointer to an error Value.
func CoerceNum(v Value) (float64, *Value) {
	switch v.Type {
	case ValueNumber:
		return v.Num, nil
	case ValueEmpty:
		return 0, nil
	case ValueBool:
		if v.Bool {
			return 1, nil
		}
		return 0, nil
	case ValueString:
		trimmed := strings.TrimSpace(v.Str)
		if trimmed == "" {
			e := ErrorVal(ErrValVALUE)
			return 0, &e
		}
		if n, ok := excelParseNumber(trimmed); ok {
			return n, nil
		}
		// Excel coerces text dates (and datetimes) to their serial value
		// during arithmetic, so anything parseDateTimeString accepts
		// behaves like a number here.
		if serial, _, ok := parseDateTimeString(trimmed); ok {
			return serial, nil
		}
		e := ErrorVal(ErrValVALUE)
		return 0, &e
	case ValueError:
		return 0, &v
	default:
		e := ErrorVal(ErrValVALUE)
		return 0, &e
	}
}

// numberToString formats a number for concatenation using Excel's rules:
//   - At most 15 significant digits (via Go's 'G' format with precision 15)
//   - Prefer plain decimal notation; only use scientific notation for
//     extremely large numbers (exponent > 20) or extremely small numbers
//     (more than 9 leading zeros after the decimal point, i.e. exponent < -9).
func numberToString(f float64) string {
	if f == 0 {
		return "0"
	}
	if math.IsInf(f, 0) || math.IsNaN(f) {
		return strconv.FormatFloat(f, 'G', 15, 64)
	}

	// Format with 15 significant digits. Go's 'G' may produce scientific
	// notation (e.g. "1E+15") for numbers outside [1e-4, 1e15).
	s := strconv.FormatFloat(f, 'G', 15, 64)

	// If already in plain decimal, nothing more to do.
	eIdx := strings.IndexByte(s, 'E')
	if eIdx < 0 {
		return s
	}

	// Parse the exponent to decide whether to expand to decimal form.
	exp, err := strconv.Atoi(s[eIdx+1:])
	if err != nil {
		return s
	}

	// Keep scientific notation for very large or very small values.
	if exp > 20 || exp < -9 {
		return s
	}

	// Convert the G-formatted scientific notation to plain decimal.
	return sciToDecimal(s[:eIdx], exp)
}

// sciToDecimal expands a mantissa string (e.g. "1.23456") with the given
// base-10 exponent into plain decimal notation. It assumes the mantissa
// has already been rounded to the desired number of significant digits.
func sciToDecimal(mantissa string, exp int) string {
	neg := len(mantissa) > 0 && mantissa[0] == '-'
	if neg {
		mantissa = mantissa[1:]
	}

	// Strip the decimal point to get a pure digit string and record where
	// the original decimal point was.
	dotIdx := strings.IndexByte(mantissa, '.')
	var digits string
	if dotIdx >= 0 {
		digits = mantissa[:dotIdx] + mantissa[dotIdx+1:]
	} else {
		digits = mantissa
	}

	// The decimal point position (counted from the left of digits) is 1 + exp
	// because the mantissa is normalised as d.ddd...
	decPos := 1 + exp
	n := len(digits)

	var result string
	switch {
	case decPos >= n:
		// All digits sit before the decimal point; pad with trailing zeros.
		result = digits + strings.Repeat("0", decPos-n)
	case decPos <= 0:
		// All digits sit after the decimal point; pad with leading zeros.
		result = "0." + strings.Repeat("0", -decPos) + digits
	default:
		result = digits[:decPos] + "." + digits[decPos:]
	}

	// Trim unnecessary trailing zeros / decimal point.
	if strings.IndexByte(result, '.') >= 0 {
		result = strings.TrimRight(result, "0")
		result = strings.TrimRight(result, ".")
	}

	if neg {
		return "-" + result
	}
	return result
}

func ValueToString(v Value) string {
	switch v.Type {
	case ValueNumber:
		return numberToString(v.Num)
	case ValueString:
		return v.Str
	case ValueBool:
		if v.Bool {
			return "TRUE"
		}
		return "FALSE"
	case ValueError:
		return errorValueToString(v.Err)
	default:
		return ""
	}
}

func errorValueToString(e ErrorValue) string {
	switch e {
	case ErrValDIV0:
		return "#DIV/0!"
	case ErrValNA:
		return "#N/A"
	case ErrValNAME:
		return "#NAME?"
	case ErrValNULL:
		return "#NULL!"
	case ErrValNUM:
		return "#NUM!"
	case ErrValREF:
		return "#REF!"
	case ErrValVALUE:
		return "#VALUE!"
	case ErrValSPILL:
		return "#SPILL!"
	case ErrValCALC:
		return "#CALC!"
	case ErrValGETTINGDATA:
		return "#GETTING_DATA"
	default:
		return "#VALUE!"
	}
}

// CompareValues compares two values for ordering. Returns -1, 0, or 1.
func CompareValues(a, b Value) int {
	// Empty cells adapt to the type of the other operand:
	//   empty = "" → TRUE,  empty = 0 → TRUE,  empty = FALSE → TRUE
	if a.Type == ValueEmpty && b.Type == ValueEmpty {
		return 0
	}
	if a.Type == ValueEmpty {
		switch b.Type {
		case ValueString:
			a = StringVal("")
		case ValueBool:
			a = BoolVal(false)
		default:
			a = NumberVal(0)
		}
	}
	if b.Type == ValueEmpty {
		switch a.Type {
		case ValueString:
			b = StringVal("")
		case ValueBool:
			b = BoolVal(false)
		default:
			b = NumberVal(0)
		}
	}

	if a.Type == b.Type {
		switch a.Type {
		case ValueNumber:
			return cmpFloat(a.Num, b.Num)
		case ValueString:
			sa := stripDefaultIgnorable(a.Str)
			sb := stripDefaultIgnorable(b.Str)
			return strings.Compare(strings.ToLower(sa), strings.ToLower(sb))
		case ValueBool:
			if a.Bool == b.Bool {
				return 0
			}
			if !a.Bool {
				return -1
			}
			return 1
		}
	}

	return typeRank(a.Type) - typeRank(b.Type)
}

// stripDefaultIgnorable removes Unicode default-ignorable code points. Excel's
// `=`, `<`, `>` operators normalise characters like ZWSP (U+200B) and BOM
// (U+FEFF) out before comparing text, even though LEN/FIND/SUBSTITUTE still
// see them. Lookup functions use CompareValuesExact, which deliberately does
// NOT call this — VLOOKUP stays strict.
func stripDefaultIgnorable(s string) string {
	needs := false
	for _, r := range s {
		if isDefaultIgnorable(r) {
			needs = true
			break
		}
	}
	if !needs {
		return s
	}
	var b strings.Builder
	b.Grow(len(s))
	for _, r := range s {
		if !isDefaultIgnorable(r) {
			b.WriteRune(r)
		}
	}
	return b.String()
}

func isDefaultIgnorable(r rune) bool {
	switch {
	case r == 0x00AD: // SOFT HYPHEN
		return true
	case r >= 0x200B && r <= 0x200F: // ZWSP, ZWNJ, ZWJ, LRM, RLM
		return true
	case r >= 0x202A && r <= 0x202E: // bidi embedding/override markers
		return true
	case r >= 0x2060 && r <= 0x2064: // WJ and invisible math operators
		return true
	case r >= 0xFE00 && r <= 0xFE0F: // variation selectors 1–16
		return true
	case r == 0xFEFF: // BOM / ZWNBSP
		return true
	}
	return false
}

// CompareValuesExact is like CompareValues but uses bit-exact float
// comparison (no tolerance). Used by lookup functions for exact-match
// mode, which does not apply the ≈1e-15 tolerance that the =
// operator uses.
func CompareValuesExact(a, b Value) int {
	if a.Type == ValueEmpty && b.Type == ValueEmpty {
		return 0
	}
	if a.Type == ValueEmpty {
		switch b.Type {
		case ValueString:
			a = StringVal("")
		case ValueBool:
			a = BoolVal(false)
		default:
			a = NumberVal(0)
		}
	}
	if b.Type == ValueEmpty {
		switch a.Type {
		case ValueString:
			b = StringVal("")
		case ValueBool:
			b = BoolVal(false)
		default:
			b = NumberVal(0)
		}
	}

	if a.Type == b.Type {
		switch a.Type {
		case ValueNumber:
			return cmpFloatExact(a.Num, b.Num)
		case ValueString:
			return strings.Compare(strings.ToLower(a.Str), strings.ToLower(b.Str))
		case ValueBool:
			if a.Bool == b.Bool {
				return 0
			}
			if !a.Bool {
				return -1
			}
			return 1
		}
	}

	return typeRank(a.Type) - typeRank(b.Type)
}

func typeRank(t ValueType) int {
	// Excel sort order: numbers first, then text, then booleans, then errors.
	switch t {
	case ValueNumber, ValueEmpty:
		return 0
	case ValueString:
		return 1
	case ValueBool:
		return 2
	case ValueError:
		return 3
	default:
		return 4
	}
}

// roundTo15SigFigs rounds a float64 to 15 significant decimal digits,
// matching the expected internal precision model.
func roundTo15SigFigs(f float64) float64 {
	if f == 0 || math.IsNaN(f) || math.IsInf(f, 0) {
		return f
	}
	a := math.Abs(f)
	if a > 1e292 || a < 1e-292 {
		return f // avoid overflow in the computation
	}
	d := math.Ceil(math.Log10(a))
	pow := math.Pow(10, 15-d)
	rounded := math.Round(a*pow) / pow
	if f < 0 {
		return -rounded
	}
	return rounded
}

// roundArithResult rounds a numeric Value to 15 significant digits.
func roundArithResult(v Value) Value {
	if v.Type == ValueNumber {
		return NumberVal(roundTo15SigFigs(v.Num))
	}
	return v
}

func cmpFloat(a, b float64) int {
	// Numbers are compared after rounding both to 15 significant digits.
	// This makes (1/3*3)=1 evaluate to TRUE while (1-1e-15)=1 is FALSE.
	ra := roundTo15SigFigs(a)
	rb := roundTo15SigFigs(b)
	if ra < rb {
		return -1
	}
	if ra > rb {
		return 1
	}
	return 0
}

// cmpFloatExact compares two float64 values without tolerance.
// Used by lookup functions (MATCH, VLOOKUP, etc.) for exact-match mode,
// where bit-exact equality is required.
func cmpFloatExact(a, b float64) int {
	if a < b {
		return -1
	}
	if a > b {
		return 1
	}
	return 0
}

func IsTruthy(v Value) bool {
	switch v.Type {
	case ValueBool:
		return v.Bool
	case ValueNumber:
		return v.Num != 0
	case ValueString:
		return v.Str != ""
	default:
		return false
	}
}

// implicitIntersect scalarizes an array for legacy implicit intersection and
// explicit @/SINGLE semantics. Anonymous arrays are left unchanged; explicit
// @/SINGLE scalarization uses explicitIntersect instead.
func implicitIntersect(v Value, ctx *EvalContext) Value {
	if v.Type != ValueArray || ctx == nil {
		return v
	}
	ref, ok := legacyArrayRef(v)
	if !ok {
		return v
	}
	ro := ref.Bounds()
	isSingleCol := ro.FromCol == ro.ToCol
	isSingleRow := ro.FromRow == ro.ToRow
	if isSingleCol {
		r := ctx.CurrentRow
		if r >= ro.FromRow && r <= ro.ToRow {
			idx := r - ro.FromRow
			return legacyRefCellValue(ref, idx, 0)
		}
		return ErrorVal(ErrValVALUE)
	}
	if isSingleRow {
		c := ctx.CurrentCol
		if c >= ro.FromCol && c <= ro.ToCol {
			idx := c - ro.FromCol
			return legacyRefCellValue(ref, 0, idx)
		}
		return ErrorVal(ErrValVALUE)
	}
	return ErrorVal(ErrValVALUE)
}

func explicitIntersect(v Value, ctx *EvalContext) Value {
	if v.Type != ValueArray {
		return v
	}
	if _, ok := legacyArrayRef(v); !ok {
		return arrayTopLeft(v)
	}
	return implicitIntersect(v, ctx)
}

// rangeIntersect computes the rectangular intersection of two ref-like values
// for Excel's intersection operator (space between references). The operands
// may arrive as raw refs, range-backed arrays, or evalRef-backed legacy
// values. Returns #VALUE! if either operand does not still resolve to live
// range bounds or the sheets differ, and #NULL! if the rectangles don't
// overlap.
func rangeIntersect(a, b Value, resolver CellResolver, ctx *EvalContext, inArrayCtx bool) Value {
	// Propagate existing errors.
	if a.Type == ValueError {
		return a
	}
	if b.Type == ValueError {
		return b
	}
	aRef, aOk := legacyValueRef(a)
	bRef, bOk := legacyValueRef(b)
	if !aOk || !bOk {
		return ErrorVal(ErrValVALUE)
	}
	aRo := aRef.Bounds()
	bRo := bRef.Bounds()
	if aRo.Sheet != bRo.Sheet || aRo.SheetEnd != "" || bRo.SheetEnd != "" {
		return ErrorVal(ErrValVALUE)
	}
	fromCol := aRo.FromCol
	if bRo.FromCol > fromCol {
		fromCol = bRo.FromCol
	}
	fromRow := aRo.FromRow
	if bRo.FromRow > fromRow {
		fromRow = bRo.FromRow
	}
	toCol := aRo.ToCol
	if bRo.ToCol < toCol {
		toCol = bRo.ToCol
	}
	toRow := aRo.ToRow
	if bRo.ToRow < toRow {
		toRow = bRo.ToRow
	}
	if fromCol > toCol || fromRow > toRow {
		return ErrorVal(ErrValNULL)
	}
	addr := RangeAddr{
		Sheet:   aRo.Sheet,
		FromCol: fromCol, FromRow: fromRow,
		ToCol: toCol, ToRow: toRow,
	}
	// Single-cell intersection when not in array context: return the scalar.
	if !inArrayCtx && fromCol == toCol && fromRow == toRow {
		return resolver.GetCellValue(CellAddr{
			Sheet: addr.Sheet, Col: fromCol, Row: fromRow,
		})
	}
	rows := resolver.GetRangeValues(addr)
	origin := addr
	return Value{Type: ValueArray, Array: rows, RangeOrigin: &origin}
}

// buildRangeFromRefs constructs the rectangular range spanning two single-cell
// references for the dynamic range operator (e.g. A1:INDEX(A:A,n)). Each
// operand must still resolve to live single-cell bounds after crossing the
// legacy Value boundary, whether it arrived as a raw ref, a cell-origin
// scalar, a single-cell range, or an evalRef-backed value. Mismatched sheets
// or unresolvable operands yield #REF!.
func buildRangeFromRefs(a, b Value, resolver CellResolver, ctx *EvalContext) Value {
	if a.Type == ValueError {
		return a
	}
	if b.Type == ValueError {
		return b
	}
	aRef, aOk := legacyValueRef(a)
	bRef, bOk := legacyValueRef(b)
	if !aOk || !bOk {
		return ErrorVal(ErrValREF)
	}
	aRo := aRef.Bounds()
	bRo := bRef.Bounds()
	if aRo.SheetEnd != "" || bRo.SheetEnd != "" {
		return ErrorVal(ErrValREF)
	}
	// Resolve empty sheet names to the current sheet. OpLoadRange leaves the
	// sheet empty for unqualified refs while OpLoadCellRef substitutes the
	// current sheet, so the two operands of A1:INDEX(A:A,n) would otherwise
	// disagree even though both refer to the same sheet.
	currentSheet := ""
	if ctx != nil {
		currentSheet = ctx.CurrentSheet
	}
	aSheet := aRo.Sheet
	if aSheet == "" {
		aSheet = currentSheet
	}
	bSheet := bRo.Sheet
	if bSheet == "" {
		bSheet = currentSheet
	}
	if aSheet != bSheet {
		return ErrorVal(ErrValREF)
	}
	aRo.Sheet = aSheet
	fromCol := aRo.FromCol
	if bRo.FromCol < fromCol {
		fromCol = bRo.FromCol
	}
	fromRow := aRo.FromRow
	if bRo.FromRow < fromRow {
		fromRow = bRo.FromRow
	}
	toCol := aRo.ToCol
	if bRo.ToCol > toCol {
		toCol = bRo.ToCol
	}
	toRow := aRo.ToRow
	if bRo.ToRow > toRow {
		toRow = bRo.ToRow
	}
	addr := RangeAddr{
		Sheet:   aRo.Sheet,
		FromCol: fromCol, FromRow: fromRow,
		ToCol: toCol, ToRow: toRow,
	}
	if fromCol == toCol && fromRow == toRow {
		v := resolver.GetCellValue(CellAddr{Sheet: addr.Sheet, Col: fromCol, Row: fromRow})
		v.FromCell = true
		v.CellOrigin = &CellAddr{Sheet: addr.Sheet, Col: fromCol, Row: fromRow}
		return v
	}
	rows := resolver.GetRangeValues(addr)
	origin := addr
	return Value{Type: ValueArray, Array: rows, RangeOrigin: &origin}
}

// unionAreas flattens a sequence of area Values into a single ValueArray by
// concatenating rows row-by-row. Scalars become 1x1 elements. The result has
// no RangeOrigin because it no longer represents a single rectangle.
// Errors are propagated: if any area is an error, that error becomes the
// result.
func unionAreas(areas []Value) Value {
	var out [][]Value
	for _, v := range areas {
		if v.Type == ValueError {
			return v
		}
		if v.Type == ValueArray {
			rows, cols := arrayOpBounds(v)
			for r := 0; r < rows; r++ {
				row := make([]Value, cols)
				for c := 0; c < cols; c++ {
					row[c] = arrayElementDirect(v, rows, cols, r, c)
				}
				out = append(out, row)
			}
			continue
		}
		out = append(out, []Value{v})
	}
	if len(out) == 0 {
		return EmptyVal()
	}
	return Value{Type: ValueArray, Array: out}
}

// arrayDims returns the maximum row and column dimensions across two values,
// treating scalars as 1×1.
func arrayDims(a, b Value) (rows, cols int) {
	rows, cols = 1, 1
	if a.Type == ValueArray {
		rows, cols = arrayOpBounds(a)
	}
	if b.Type == ValueArray {
		if r, c := arrayOpBounds(b); r > rows {
			rows = r
			cols = max(cols, c)
		} else if c > cols {
			cols = c
		}
	}
	return
}

// binaryArith performs a binary arithmetic operation on two Values,
// supporting element-wise array operations when one or both operands are arrays.
func binaryArith(a, b Value, op func(float64, float64) Value) Value {
	if a.Type != ValueArray && b.Type != ValueArray {
		// Scalar case.
		an, ae := CoerceNum(a)
		bn, be := CoerceNum(b)
		if ae != nil {
			return *ae
		}
		if be != nil {
			return *be
		}
		return op(an, bn)
	}

	// At least one operand is an array — do element-wise computation.
	rows, cols := arrayDims(a, b)
	aRows, aCols := arrayOpBoundsOrScalar(a)
	bRows, bCols := arrayOpBoundsOrScalar(b)
	result := newValueMatrix(rows, cols)
	for i := 0; i < rows; i++ {
		for j := 0; j < cols; j++ {
			av := arrayElementDirect(a, aRows, aCols, i, j)
			bv := arrayElementDirect(b, bRows, bCols, i, j)
			an, ae := CoerceNum(av)
			bn, be := CoerceNum(bv)
			if ae != nil {
				result[i][j] = *ae
			} else if be != nil {
				result[i][j] = *be
			} else {
				result[i][j] = op(an, bn)
			}
		}
	}
	out := Value{Type: ValueArray, Array: result}
	out.RangeOrigin = combinedArrayOpRangeOrigin(rows, cols, a, b)
	return out
}

// binaryCompare performs a comparison operation on two Values, supporting
// element-wise array operations when one or both operands are arrays.
func binaryCompare(a, b Value, op func(int) bool) Value {
	if a.Type != ValueArray && b.Type != ValueArray {
		if a.Type == ValueError {
			return a
		}
		if b.Type == ValueError {
			return b
		}
		return BoolVal(op(CompareValues(a, b)))
	}

	// At least one operand is an array — do element-wise comparison.
	rows, cols := arrayDims(a, b)
	aRows, aCols := arrayOpBoundsOrScalar(a)
	bRows, bCols := arrayOpBoundsOrScalar(b)
	result := newValueMatrix(rows, cols)
	for i := 0; i < rows; i++ {
		for j := 0; j < cols; j++ {
			av := arrayElementDirect(a, aRows, aCols, i, j)
			bv := arrayElementDirect(b, bRows, bCols, i, j)
			if av.Type == ValueError {
				result[i][j] = av
			} else if bv.Type == ValueError {
				result[i][j] = bv
			} else {
				result[i][j] = BoolVal(op(CompareValues(av, bv)))
			}
		}
	}
	out := Value{Type: ValueArray, Array: result}
	out.RangeOrigin = combinedArrayOpRangeOrigin(rows, cols, a, b)
	return out
}

// callFunction is replaced by CallFunc in registry.go.

// LiftUnary applies a scalar function element-wise over a ValueArray,
// returning a new ValueArray of the same shape. Used for array-formula
// evaluation of functions like ABS, ISNUMBER, etc.
func LiftUnary(arr Value, fn func(Value) Value) Value {
	rows, cols := arrayOpBounds(arr)
	result := newValueMatrix(rows, cols)
	for i := 0; i < rows; i++ {
		for j := 0; j < cols; j++ {
			result[i][j] = fn(arrayElementDirect(arr, rows, cols, i, j))
		}
	}
	out := Value{Type: ValueArray, Array: result}
	out.RangeOrigin = combinedArrayOpRangeOrigin(rows, cols, arr)
	return out
}

// ArrayElement returns element [i][j] from arr if it is an array,
// or returns the scalar arr otherwise. Used for broadcasting scalars
// alongside arrays in element-wise operations.
func ArrayElement(v Value, i, j int) Value {
	if v.Type != ValueArray {
		return v
	}
	rows, cols := effectiveArrayBounds(v)
	return arrayElementDirect(v, rows, cols, i, j)
}

// resolveSheetRange returns the slice of sheet names from startSheet to endSheet
// inclusive, based on the ordering in allSheets. If either sheet is not found,
// returns nil. Comparison is case-insensitive.
func resolveSheetRange(allSheets []string, startSheet, endSheet string) []string {
	startIdx := -1
	endIdx := -1
	startLower := strings.ToLower(startSheet)
	endLower := strings.ToLower(endSheet)
	for i, name := range allSheets {
		nameLower := strings.ToLower(name)
		if nameLower == startLower {
			startIdx = i
		}
		if nameLower == endLower {
			endIdx = i
		}
	}
	if startIdx < 0 || endIdx < 0 {
		return nil
	}
	if startIdx > endIdx {
		startIdx, endIdx = endIdx, startIdx
	}
	return allSheets[startIdx : endIdx+1]
}

// IterateNumeric calls fn for each numeric value in args, expanding arrays.
// Non-numeric values in ranges are skipped; non-numeric scalar args cause #VALUE!.
func IterateNumeric(args []Value, fn func(float64)) *Value {
	for _, arg := range args {
		if arg.Type != ValueArray {
			if arg.Type == ValueError {
				return &arg
			}
			n, e := CoerceNum(arg)
			if e != nil {
				return e
			}
			fn(n)
			continue
		}
		var err *Value
		iterateValueElements(arg, func(cell Value) bool {
			if cell.Type == ValueError {
				cellErr := cell
				err = &cellErr
				return false
			}
			if cell.Type == ValueNumber {
				fn(cell.Num)
			}
			return true
		})
		if err != nil {
			return err
		}
	}
	return nil
}
