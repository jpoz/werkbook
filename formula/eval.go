package formula

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

const (
	maxExcelRows = 1048576 // maximum rows in an Excel worksheet
	maxExcelCols = 16384   // maximum columns in an Excel worksheet (XFD)
)

// CellResolver abstracts cell/range lookups so the VM has no dependency on Sheet.
type CellResolver interface {
	GetCellValue(addr CellAddr) Value
	GetRangeValues(addr RangeAddr) [][]Value
}

// EvalContext provides context about the current evaluation environment.
type EvalContext struct {
	CurrentCol     int
	CurrentRow     int
	CurrentSheet   string
	IsArrayFormula bool // true for CSE (Ctrl+Shift+Enter) array formulas
	Date1904       bool // true if the workbook uses the 1904 date system
	Resolver       CellResolver // the active resolver; used by SUBTOTAL to inspect cells
}

// SubtotalChecker is an optional interface that a CellResolver may implement
// to allow SUBTOTAL to skip cells that themselves contain SUBTOTAL formulas,
// preventing double-counting of nested subtotals.
type SubtotalChecker interface {
	IsSubtotalCell(sheet string, col, row int) bool
}

// SheetListProvider is an optional interface that a CellResolver may implement
// to support 3D sheet references (e.g. Sheet2:Sheet5!A1). It returns the
// ordered list of all sheet names in the workbook.
type SheetListProvider interface {
	GetSheetNames() []string
}

// Eval executes a compiled formula and returns the result.
func Eval(cf *CompiledFormula, resolver CellResolver, ctx *EvalContext) (Value, error) {
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

	for _, inst := range cf.Code {
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
			push(resolver.GetCellValue(addr))

		case OpLoadRange:
			addr := cf.Ranges[inst.Operand]
			// Implicit intersection: when a full-column or full-row range is
			// used in a non-array formula, reduce to the single cell at the
			// formula's own row/column rather than loading the entire range.
			if ctx != nil && !ctx.IsArrayFormula {
				isFullCol := addr.FromRow == 1 && addr.ToRow >= maxExcelRows
				isFullRow := addr.FromCol == 1 && addr.ToCol >= maxExcelCols
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
			rows := resolver.GetRangeValues(addr)
			// Pad trailing blank rows for bounded ranges. GetRangeValues
			// clamps toRow to MaxRow to avoid huge allocations for
			// full-column refs, but bounded ranges like A1:A5 need all
			// requested rows so functions like COUNTBLANK see every blank.
			isFullCol := addr.FromRow == 1 && addr.ToRow >= maxExcelRows
			isFullRow := addr.FromCol == 1 && addr.ToCol >= maxExcelCols
			if !isFullCol && !isFullRow {
				expectedRows := addr.ToRow - addr.FromRow + 1
				cols := addr.ToCol - addr.FromCol + 1
				for len(rows) < expectedRows {
					emptyRow := make([]Value, cols)
					for j := range emptyRow {
						emptyRow[j] = EmptyVal()
					}
					rows = append(rows, emptyRow)
				}
			}
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
			push(Value{Type: ValueRef, Num: float64(addr.Col + addr.Row*100_000)})

		case OpAdd:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			if ctx != nil && !ctx.IsArrayFormula && arrayCtxDepth == 0 {
				a = implicitIntersect(a, ctx)
				b = implicitIntersect(b, ctx)
			}
			push(binaryArith(a, b, func(an, bn float64) Value {
				return NumberVal(an + bn)
			}))

		case OpSub:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			if ctx != nil && !ctx.IsArrayFormula && arrayCtxDepth == 0 {
				a = implicitIntersect(a, ctx)
				b = implicitIntersect(b, ctx)
			}
			push(binaryArith(a, b, func(an, bn float64) Value {
				return NumberVal(an - bn)
			}))

		case OpMul:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			if ctx != nil && !ctx.IsArrayFormula && arrayCtxDepth == 0 {
				a = implicitIntersect(a, ctx)
				b = implicitIntersect(b, ctx)
			}
			push(binaryArith(a, b, func(an, bn float64) Value {
				return NumberVal(an * bn)
			}))

		case OpDiv:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			if ctx != nil && !ctx.IsArrayFormula && arrayCtxDepth == 0 {
				a = implicitIntersect(a, ctx)
				b = implicitIntersect(b, ctx)
			}
			push(binaryArith(a, b, func(an, bn float64) Value {
				if bn == 0 {
					return ErrorVal(ErrValDIV0)
				}
				return NumberVal(an / bn)
			}))

		case OpPow:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			if ctx != nil && !ctx.IsArrayFormula && arrayCtxDepth == 0 {
				a = implicitIntersect(a, ctx)
				b = implicitIntersect(b, ctx)
			}
			push(binaryArith(a, b, func(an, bn float64) Value {
				result := math.Pow(an, bn)
				if math.IsNaN(result) || math.IsInf(result, 0) {
					return ErrorVal(ErrValNUM)
				}
				return NumberVal(result)
			}))

		case OpNeg:
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			an, ae := CoerceNum(a)
			if ae != nil {
				push(*ae)
			} else {
				push(NumberVal(-an))
			}

		case OpPercent:
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			an, ae := CoerceNum(a)
			if ae != nil {
				push(*ae)
			} else {
				push(NumberVal(an / 100))
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
			push(StringVal(ValueToString(a) + ValueToString(b)))

		case OpEq:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(binaryCompare(a, b, func(c int) bool { return c == 0 }))

		case OpNe:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(binaryCompare(a, b, func(c int) bool { return c != 0 }))

		case OpLt:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(binaryCompare(a, b, func(c int) bool { return c < 0 }))

		case OpLe:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(binaryCompare(a, b, func(c int) bool { return c <= 0 }))

		case OpGt:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(binaryCompare(a, b, func(c int) bool { return c > 0 }))

		case OpGe:
			b, err := pop()
			if err != nil {
				return Value{}, err
			}
			a, err := pop()
			if err != nil {
				return Value{}, err
			}
			push(binaryCompare(a, b, func(c int) bool { return c >= 0 }))

		case OpCall:
			funcID := int(inst.Operand >> 8)
			argc := int(inst.Operand & 0xFF)
			if argc > len(stack) {
				return Value{}, fmt.Errorf("stack underflow in function call")
			}
			args := make([]Value, argc)
			copy(args, stack[len(stack)-argc:])
			stack = stack[:len(stack)-argc]

			result, err := CallFunc(funcID, args, ctx)
			if err != nil {
				return Value{}, err
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
		if v.Str == "" {
			return 0, nil
		}
		n, err := strconv.ParseFloat(v.Str, 64)
		if err != nil {
			e := ErrorVal(ErrValVALUE)
			return 0, &e
		}
		return n, nil
	case ValueError:
		return 0, &v
	default:
		e := ErrorVal(ErrValVALUE)
		return 0, &e
	}
}

func ValueToString(v Value) string {
	switch v.Type {
	case ValueNumber:
		return strconv.FormatFloat(v.Num, 'f', -1, 64)
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
	// In Excel, empty cells adapt to the type of the other operand:
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
	switch t {
	case ValueError:
		return 0
	case ValueNumber, ValueEmpty:
		return 1
	case ValueString:
		return 2
	case ValueBool:
		return 3
	default:
		return 4
	}
}

func cmpFloat(a, b float64) int {
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

// implicitIntersect reduces a ValueArray loaded from a worksheet range to a
// scalar value using Excel's implicit intersection rules (legacy non-array
// formula behaviour).  For a single-column range the value at the formula's
// row is returned; for a single-row range the value at the formula's column
// is returned.  If the range is multi-row and multi-column, or the formula
// position falls outside the range, #VALUE! is returned.  Values that are
// not range-origin arrays are returned unchanged.
func implicitIntersect(v Value, ctx *EvalContext) Value {
	if v.Type != ValueArray || ctx == nil || v.RangeOrigin == nil {
		return v
	}
	ro := v.RangeOrigin
	isSingleCol := ro.FromCol == ro.ToCol
	isSingleRow := ro.FromRow == ro.ToRow
	if isSingleCol {
		r := ctx.CurrentRow
		if r >= ro.FromRow && r <= ro.ToRow && len(v.Array) > 0 {
			idx := r - ro.FromRow
			if idx < len(v.Array) {
				return v.Array[idx][0]
			}
		}
		return ErrorVal(ErrValVALUE)
	}
	if isSingleRow {
		c := ctx.CurrentCol
		if c >= ro.FromCol && c <= ro.ToCol && len(v.Array) > 0 {
			idx := c - ro.FromCol
			if idx < len(v.Array[0]) {
				return v.Array[0][idx]
			}
		}
		return ErrorVal(ErrValVALUE)
	}
	return ErrorVal(ErrValVALUE)
}

// arrayDims returns the maximum row and column dimensions across two values,
// treating scalars as 1×1.
func arrayDims(a, b Value) (rows, cols int) {
	rows, cols = 1, 1
	if a.Type == ValueArray {
		rows = len(a.Array)
		if rows > 0 {
			cols = len(a.Array[0])
		}
	}
	if b.Type == ValueArray {
		if r := len(b.Array); r > rows {
			rows = r
		}
		if len(b.Array) > 0 {
			if c := len(b.Array[0]); c > cols {
				cols = c
			}
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
	result := make([][]Value, rows)
	for i := 0; i < rows; i++ {
		result[i] = make([]Value, cols)
		for j := 0; j < cols; j++ {
			av := ArrayElement(a, i, j)
			bv := ArrayElement(b, i, j)
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
	return Value{Type: ValueArray, Array: result}
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
	result := make([][]Value, rows)
	for i := 0; i < rows; i++ {
		result[i] = make([]Value, cols)
		for j := 0; j < cols; j++ {
			av := ArrayElement(a, i, j)
			bv := ArrayElement(b, i, j)
			if av.Type == ValueError {
				result[i][j] = av
			} else if bv.Type == ValueError {
				result[i][j] = bv
			} else {
				result[i][j] = BoolVal(op(CompareValues(av, bv)))
			}
		}
	}
	return Value{Type: ValueArray, Array: result}
}

// callFunction is replaced by CallFunc in registry.go.

// LiftUnary applies a scalar function element-wise over a ValueArray,
// returning a new ValueArray of the same shape. Used for array-formula
// evaluation of functions like ABS, ISNUMBER, etc.
func LiftUnary(arr Value, fn func(Value) Value) Value {
	rows := make([][]Value, len(arr.Array))
	for i, row := range arr.Array {
		out := make([]Value, len(row))
		for j, cell := range row {
			out[j] = fn(cell)
		}
		rows[i] = out
	}
	return Value{Type: ValueArray, Array: rows}
}

// ArrayElement returns element [i][j] from arr if it is an array,
// or returns the scalar arr otherwise. Used for broadcasting scalars
// alongside arrays in element-wise operations.
func ArrayElement(v Value, i, j int) Value {
	if v.Type != ValueArray {
		return v
	}
	if i < len(v.Array) && j < len(v.Array[i]) {
		return v.Array[i][j]
	}
	return ErrorVal(ErrValNA)
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
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return &cell
					}
					if cell.Type == ValueNumber {
						fn(cell.Num)
					}
				}
			}
		} else {
			if arg.Type == ValueError {
				return &arg
			}
			n, e := CoerceNum(arg)
			if e != nil {
				return e
			}
			fn(n)
		}
	}
	return nil
}
