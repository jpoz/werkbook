package formula

func init() {
	Register("AND", NoCtx(fnAND))
	Register("IF", NoCtx(fnIF))
	Register("IFERROR", NoCtx(fnIFERROR))
	Register("NOT", NoCtx(fnNOT))
	Register("OR", NoCtx(fnOR))
}

func fnIF(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	// Array formula: when the condition is an array, apply IF element-wise.
	if args[0].Type == ValueArray {
		cond := args[0]
		rows := make([][]Value, len(cond.Array))
		for i, row := range cond.Array {
			out := make([]Value, len(row))
			for j, cell := range row {
				if IsTruthy(cell) {
					out[j] = ArrayElement(args[1], i, j)
				} else if len(args) == 3 {
					out[j] = ArrayElement(args[2], i, j)
				} else {
					out[j] = BoolVal(false)
				}
			}
			rows[i] = out
		}
		return Value{Type: ValueArray, Array: rows}, nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	if IsTruthy(args[0]) {
		return args[1], nil
	}
	if len(args) == 3 {
		return args[2], nil
	}
	return BoolVal(false), nil
}

func fnIFERROR(args []Value) (Value, error) {
	if len(args) != 2 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[1], nil
	}
	return args[0], nil
}

func fnAND(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	for _, arg := range args {
		if arg.Type == ValueError {
			return arg, nil
		}
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return cell, nil
					}
					if !IsTruthy(cell) {
						return BoolVal(false), nil
					}
				}
			}
			continue
		}
		if !IsTruthy(arg) {
			return BoolVal(false), nil
		}
	}
	return BoolVal(true), nil
}

func fnOR(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	for _, arg := range args {
		if arg.Type == ValueError {
			return arg, nil
		}
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return cell, nil
					}
					if IsTruthy(cell) {
						return BoolVal(true), nil
					}
				}
			}
			continue
		}
		if IsTruthy(arg) {
			return BoolVal(true), nil
		}
	}
	return BoolVal(false), nil
}

func fnNOT(args []Value) (Value, error) {
	if len(args) != 1 {
		return ErrorVal(ErrValVALUE), nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	return BoolVal(!IsTruthy(args[0])), nil
}
