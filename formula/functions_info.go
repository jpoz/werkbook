package formula

import (
	"math"
	"strings"
)

func init() {
	Register("AREAS", NoCtx(fnAREAS))
	Register("COLUMN", fnCOLUMN)
	Register("COLUMNS", NoCtx(fnCOLUMNS))
	Register("ERROR.TYPE", NoCtx(fnERRORTYPE))
	RegisterWithMeta("IFNA", NoCtx(fnIFNA), FuncMeta{
		Kind:               FnKindScalarLifted,
		InheritedArrayArgs: map[int]bool{0: true, 1: true},
	})
	RegisterScalarLiftedUnarySpec("ISBLANK", NoCtx(fnISBLANK))
	RegisterScalarLiftedUnarySpec("ISERR", NoCtx(fnISERR))
	RegisterScalarLiftedUnarySpec("ISERROR", NoCtx(fnISERROR))
	Register("ISEVEN", NoCtx(fnISEVEN))
	Register("ISLOGICAL", NoCtx(fnISLOGICAL))
	RegisterScalarLiftedUnarySpec("ISNA", NoCtx(fnISNA))
	Register("ISNONTEXT", NoCtx(fnISNONTEXT))
	RegisterScalarLiftedUnarySpec("ISNUMBER", NoCtx(fnISNUMBER))
	Register("ISODD", NoCtx(fnISODD))
	Register("ISREF", NoCtx(fnISREF))
	RegisterScalarLiftedUnarySpec("ISTEXT", NoCtx(fnISTEXT))
	RegisterScalarLiftedUnarySpec("N", NoCtx(fnN))
	Register("NA", NoCtx(fnNA))
	Register("ROW", fnROW)
	Register("ROWS", NoCtx(fnROWS))
	RegisterScalarLifted("TYPE", NoCtx(fnTYPE))
	Register("ISFORMULA", fnISFORMULA)
	Register("FORMULATEXT", fnFORMULATEXT)
	Register("SHEET", fnSHEET)
	Register("SHEETS", fnSHEETS)
}

func fnIFNA(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError && args[0].Err == ErrValNA {
		return args[1], nil
	}
	return args[0], nil
}

func fnISBLANK(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(args[0].Type == ValueEmpty), nil
}

func fnISERR(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(args[0].Type == ValueError && args[0].Err != ErrValNA), nil
}

func fnISERROR(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(args[0].Type == ValueError), nil
}

func fnISEVEN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueBool {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return BoolVal(int(math.Trunc(n))%2 == 0), nil
}

func fnISODD(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueBool {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return *e, nil
	}
	return BoolVal(int(math.Trunc(n))%2 != 0), nil
}

// fnISREF implements ISREF(value). It returns TRUE if the argument is a cell
// or range reference, FALSE otherwise. Errors are not propagated — ISREF(1/0)
// returns FALSE, not #DIV/0!.
func fnISREF(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	_, ok := infoRefBounds(args[0])
	return BoolVal(ok), nil
}

func fnISNA(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(args[0].Type == ValueError && args[0].Err == ErrValNA), nil
}

func fnISNUMBER(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueArray {
		return LiftUnary(args[0], func(v Value) Value {
			return BoolVal(v.Type == ValueNumber)
		}), nil
	}
	return BoolVal(args[0].Type == ValueNumber), nil
}

func fnISTEXT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(args[0].Type == ValueString), nil
}

func fnISLOGICAL(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(args[0].Type == ValueBool), nil
}

func fnISNONTEXT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(args[0].Type != ValueString), nil
}

func fnERRORTYPE(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type != ValueError {
		return ErrorVal(ErrValNA), nil
	}
	switch args[0].Err {
	case ErrValNULL:
		return NumberVal(1), nil
	case ErrValDIV0:
		return NumberVal(2), nil
	case ErrValVALUE:
		return NumberVal(3), nil
	case ErrValREF:
		return NumberVal(4), nil
	case ErrValNAME:
		return NumberVal(5), nil
	case ErrValNUM:
		return NumberVal(6), nil
	case ErrValNA:
		return NumberVal(7), nil
	default:
		return ErrorVal(ErrValNA), nil
	}
}

func fnNA(args []Value) (Value, error) {
	if len(args) != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return ErrorVal(ErrValNA), nil
}

func fnN(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	switch args[0].Type {
	case ValueNumber:
		return args[0], nil
	case ValueBool:
		if args[0].Bool {
			return NumberVal(1), nil
		}
		return NumberVal(0), nil
	case ValueError:
		return args[0], nil
	case ValueString, ValueEmpty:
		return NumberVal(0), nil
	default:
		return NumberVal(0), nil
	}
}

func fnAREAS(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if areas, ok := areasCount(args[0]); ok {
		return NumberVal(float64(areas)), nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	return ErrorVal(ErrValVALUE), nil
}

func infoRefBounds(v Value) (RangeAddr, bool) {
	ref, ok := legacyValueRef(v)
	if !ok {
		return RangeAddr{}, false
	}
	return ref.Bounds(), true
}

func areasCount(v Value) (int, bool) {
	if _, ok := infoRefBounds(v); ok {
		return 1, true
	}
	return 0, false
}

func fnCOLUMN(args []Value, ctx *EvalContext) (Value, error) {
	if len(args) == 0 {
		if ctx == nil {
			return ErrorVal(ErrValVALUE), nil
		}
		return NumberVal(float64(ctx.CurrentCol)), nil
	}
	if len(args) == 1 {
		if bounds, ok := infoRefBounds(args[0]); ok {
			if args[0].Type == ValueArray && args[0].RangeOrigin != nil {
				return ErrorVal(ErrValVALUE), nil
			}
			return NumberVal(float64(bounds.FromCol)), nil
		}
		if args[0].Type == ValueError {
			return args[0], nil
		}
	}
	return ErrorVal(ErrValVALUE), nil
}

func fnROW(args []Value, ctx *EvalContext) (Value, error) {
	if len(args) == 0 {
		if ctx == nil {
			return ErrorVal(ErrValVALUE), nil
		}
		return NumberVal(float64(ctx.CurrentRow)), nil
	}
	if len(args) == 1 {
		if bounds, ok := infoRefBounds(args[0]); ok {
			if args[0].Type != ValueArray || args[0].RangeOrigin == nil {
				return NumberVal(float64(bounds.FromRow)), nil
			}
			nRows := bounds.ToRow - bounds.FromRow + 1
			rows := make([][]Value, nRows)
			for i := 0; i < nRows; i++ {
				rows[i] = []Value{NumberVal(float64(bounds.FromRow + i))}
			}
			return Value{Type: ValueArray, Array: rows}, nil
		}
		if args[0].Type == ValueError {
			return args[0], nil
		}
	}
	return ErrorVal(ErrValVALUE), nil
}

func fnCOLUMNS(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		if args[0].Err == ErrValCALC {
			return ErrorVal(ErrValVALUE), nil
		}
		return args[0], nil
	}
	if args[0].Type == ValueArray {
		_, cols := effectiveArrayBounds(args[0])
		if cols > 0 {
			return NumberVal(float64(cols)), nil
		}
	}
	return NumberVal(1), nil
}

func fnROWS(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		if args[0].Err == ErrValCALC {
			return ErrorVal(ErrValVALUE), nil
		}
		return args[0], nil
	}
	if args[0].Type == ValueArray {
		rows, _ := effectiveArrayBounds(args[0])
		if rows > 0 {
			return NumberVal(float64(rows)), nil
		}
	}
	return NumberVal(1), nil
}

func fnTYPE(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	switch args[0].Type {
	case ValueNumber:
		return NumberVal(1), nil
	case ValueString:
		return NumberVal(2), nil
	case ValueBool:
		return NumberVal(4), nil
	case ValueError:
		return NumberVal(16), nil
	case ValueArray:
		return NumberVal(64), nil
	case ValueEmpty:
		return NumberVal(1), nil
	default:
		return NumberVal(1), nil
	}
}

// fnISFORMULA implements ISFORMULA(reference). It returns TRUE if the
// referenced cell contains a formula, FALSE otherwise. The argument must
// be a cell reference (ValueRef); non-reference arguments return #VALUE!.
func fnISFORMULA(args []Value, ctx *EvalContext) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type != ValueRef {
		return ErrorVal(ErrValVALUE), nil
	}
	if ctx == nil || ctx.Resolver == nil {
		return BoolVal(false), nil
	}
	fi, ok := ctx.Resolver.(FormulaIntrospector)
	if !ok {
		// Resolver does not support formula introspection; fall back to FALSE.
		return BoolVal(false), nil
	}
	row := int(args[0].Num) / 100_000
	col := int(args[0].Num) % 100_000
	return BoolVal(fi.HasFormula(ctx.CurrentSheet, col, row)), nil
}

// fnFORMULATEXT implements FORMULATEXT(reference). It returns the formula
// text (with leading '=') if the referenced cell contains a formula, or
// #N/A if it does not. The argument must be a cell reference (ValueRef);
// non-reference arguments return #VALUE!.
func fnFORMULATEXT(args []Value, ctx *EvalContext) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type != ValueRef {
		return ErrorVal(ErrValVALUE), nil
	}
	if ctx == nil || ctx.Resolver == nil {
		return ErrorVal(ErrValNA), nil
	}
	fi, ok := ctx.Resolver.(FormulaIntrospector)
	if !ok {
		// Resolver does not support formula introspection; return #N/A.
		return ErrorVal(ErrValNA), nil
	}
	row := int(args[0].Num) / 100_000
	col := int(args[0].Num) % 100_000
	sheet := args[0].Str
	if sheet == "" {
		sheet = ctx.CurrentSheet
	}
	text := fi.GetFormulaText(sheet, col, row)
	if text == "" {
		return ErrorVal(ErrValNA), nil
	}
	return StringVal("=" + text), nil
}

// fnSHEET implements SHEET(value). It returns the 1-based sheet index of
// the given sheet reference or name. With no arguments it returns the index
// of the current sheet.
func fnSHEET(args []Value, ctx *EvalContext) (Value, error) {
	if len(args) > 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// No args: return current sheet index.
	if len(args) == 0 {
		if ctx == nil || ctx.Resolver == nil {
			return ErrorVal(ErrValNA), nil
		}
		slp, ok := ctx.Resolver.(SheetListProvider)
		if !ok {
			return ErrorVal(ErrValNA), nil
		}
		sheets := slp.GetSheetNames()
		for i, name := range sheets {
			if strings.EqualFold(name, ctx.CurrentSheet) {
				return NumberVal(float64(i + 1)), nil
			}
		}
		return ErrorVal(ErrValNA), nil
	}

	arg := args[0]

	if bounds, ok := infoRefBounds(arg); ok {
		sheetName := bounds.Sheet
		if sheetName == "" && ctx != nil {
			sheetName = ctx.CurrentSheet
		}
		if ctx == nil || ctx.Resolver == nil {
			return ErrorVal(ErrValNA), nil
		}
		slp, ok := ctx.Resolver.(SheetListProvider)
		if !ok {
			return ErrorVal(ErrValNA), nil
		}
		sheets := slp.GetSheetNames()
		for i, name := range sheets {
			if strings.EqualFold(name, sheetName) {
				return NumberVal(float64(i + 1)), nil
			}
		}
		return ErrorVal(ErrValREF), nil
	}

	// Error propagation.
	if arg.Type == ValueError {
		return arg, nil
	}

	// String arg: look up sheet by name.
	if arg.Type == ValueString {
		if ctx == nil || ctx.Resolver == nil {
			return ErrorVal(ErrValNA), nil
		}
		slp, ok := ctx.Resolver.(SheetListProvider)
		if !ok {
			return ErrorVal(ErrValNA), nil
		}
		sheets := slp.GetSheetNames()
		for i, name := range sheets {
			if strings.EqualFold(name, arg.Str) {
				return NumberVal(float64(i + 1)), nil
			}
		}
		return ErrorVal(ErrValNA), nil
	}

	// Other types: #VALUE!
	return ErrorVal(ErrValVALUE), nil
}

// fnSHEETS implements SHEETS(reference). It returns the number of sheets
// in a reference. With no arguments it returns the total number of sheets
// in the workbook.
func fnSHEETS(args []Value, ctx *EvalContext) (Value, error) {
	if len(args) > 1 {
		return ErrorVal(ErrValVALUE), nil
	}

	// No args: return total sheet count.
	if len(args) == 0 {
		if ctx == nil || ctx.Resolver == nil {
			return ErrorVal(ErrValNA), nil
		}
		slp, ok := ctx.Resolver.(SheetListProvider)
		if !ok {
			return ErrorVal(ErrValNA), nil
		}
		return NumberVal(float64(len(slp.GetSheetNames()))), nil
	}

	arg := args[0]

	if bounds, ok := infoRefBounds(arg); ok {
		if bounds.SheetEnd != "" {
			if ctx == nil || ctx.Resolver == nil {
				return ErrorVal(ErrValREF), nil
			}
			slp, ok := ctx.Resolver.(SheetListProvider)
			if !ok {
				return ErrorVal(ErrValREF), nil
			}
			sheets := slp.GetSheetNames()
			startIdx, endIdx := -1, -1
			startLower := strings.ToLower(bounds.Sheet)
			endLower := strings.ToLower(bounds.SheetEnd)
			for i, name := range sheets {
				nameLower := strings.ToLower(name)
				if nameLower == startLower {
					startIdx = i
				}
				if nameLower == endLower {
					endIdx = i
				}
			}
			if startIdx < 0 || endIdx < 0 {
				return ErrorVal(ErrValREF), nil
			}
			if startIdx > endIdx {
				startIdx, endIdx = endIdx, startIdx
			}
			return NumberVal(float64(endIdx - startIdx + 1)), nil
		}
		return NumberVal(1), nil
	}

	if arg.Type == ValueError {
		return arg, nil
	}

	// Other types: #VALUE!
	return ErrorVal(ErrValVALUE), nil
}
