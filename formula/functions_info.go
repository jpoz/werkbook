package formula

import "math"

func init() {
	Register("COLUMN", fnCOLUMN)
	Register("COLUMNS", NoCtx(fnCOLUMNS))
	Register("ERROR.TYPE", NoCtx(fnERRORTYPE))
	Register("IFNA", NoCtx(fnIFNA))
	Register("ISBLANK", NoCtx(fnISBLANK))
	Register("ISERR", NoCtx(fnISERR))
	Register("ISERROR", NoCtx(fnISERROR))
	Register("ISEVEN", NoCtx(fnISEVEN))
	Register("ISLOGICAL", NoCtx(fnISLOGICAL))
	Register("ISNA", NoCtx(fnISNA))
	Register("ISNONTEXT", NoCtx(fnISNONTEXT))
	Register("ISNUMBER", NoCtx(fnISNUMBER))
	Register("ISODD", NoCtx(fnISODD))
	Register("ISTEXT", NoCtx(fnISTEXT))
	Register("N", NoCtx(fnN))
	Register("NA", NoCtx(fnNA))
	Register("ROW", fnROW)
	Register("ROWS", NoCtx(fnROWS))
	Register("TYPE", NoCtx(fnTYPE))
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
	n, e := CoerceNum(args[0])
	if e != nil {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(int(math.Trunc(n))%2 == 0), nil
}

func fnISODD(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	n, e := CoerceNum(args[0])
	if e != nil {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(int(math.Trunc(n))%2 != 0), nil
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

func fnCOLUMN(args []Value, ctx *EvalContext) (Value, error) {
	if len(args) == 0 {
		if ctx == nil {
			return ErrorVal(ErrValVALUE), nil
		}
		return NumberVal(float64(ctx.CurrentCol)), nil
	}
	if len(args) == 1 && args[0].Type == ValueRef {
		col := int(args[0].Num) % 100_000
		return NumberVal(float64(col)), nil
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
	if len(args) == 1 && args[0].Type == ValueRef {
		row := int(args[0].Num) / 100_000
		return NumberVal(float64(row)), nil
	}
	return ErrorVal(ErrValVALUE), nil
}

func fnCOLUMNS(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueArray && len(args[0].Array) > 0 {
		return NumberVal(float64(len(args[0].Array[0]))), nil
	}
	return NumberVal(1), nil
}

func fnROWS(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueArray {
		return NumberVal(float64(len(args[0].Array))), nil
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
