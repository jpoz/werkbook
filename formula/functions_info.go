package formula

func fnISBLANK(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(args[0].Type == ValueEmpty), nil
}

func fnISERROR(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(args[0].Type == ValueError), nil
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
		return liftUnary(args[0], func(v Value) Value {
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

func fnIFNA(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError && args[0].Err == ErrValNA {
		return args[1], nil
	}
	return args[0], nil
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
