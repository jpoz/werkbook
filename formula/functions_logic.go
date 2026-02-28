package formula

import "sort"

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
				if isTruthy(cell) {
					out[j] = arrayElement(args[1], i, j)
				} else if len(args) == 3 {
					out[j] = arrayElement(args[2], i, j)
				} else {
					out[j] = BoolVal(false)
				}
			}
			rows[i] = out
		}
		return Value{Type: ValueArray, Array: rows}, nil
	}
	if isTruthy(args[0]) {
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
		if !isTruthy(arg) {
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
		if isTruthy(arg) {
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
	return BoolVal(!isTruthy(args[0])), nil
}

func fnXOR(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	count := 0
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
					if isTruthy(cell) {
						count++
					}
				}
			}
		} else if isTruthy(arg) {
			count++
		}
	}
	return BoolVal(count%2 == 1), nil
}

func fnIFS(args []Value) (Value, error) {
	// Must have an even number of args >= 2 (condition/value pairs).
	if len(args) < 2 || len(args)%2 != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	// Iterate condition/value pairs; return value for first TRUE condition.
	for i := 0; i < len(args); i += 2 {
		if args[i].Type == ValueError {
			return args[i], nil
		}
		if isTruthy(args[i]) {
			return args[i+1], nil
		}
	}
	// No TRUE condition found.
	return ErrorVal(ErrValNA), nil
}

func fnSWITCH(args []Value) (Value, error) {
	// SWITCH(expression, value1, result1, [default_or_value2, result2], ...)
	// Minimum 3 args: expression, value1, result1
	if len(args) < 3 {
		return ErrorVal(ErrValVALUE), nil
	}

	expr := args[0]
	rest := args[1:] // everything after the expression

	// Determine if there's a default value.
	// If odd count of remaining args → last is default; pairs before it are value/result.
	// If even count → all are value/result pairs; no default.
	hasDefault := len(rest)%2 == 1
	pairCount := len(rest) / 2

	// Check each value/result pair
	for i := 0; i < pairCount; i++ {
		val := rest[i*2]
		result := rest[i*2+1]
		if compareValues(expr, val) == 0 {
			return result, nil
		}
	}

	// No match found
	if hasDefault {
		return rest[len(rest)-1], nil
	}
	return ErrorVal(ErrValNA), nil
}

func fnSORT(args []Value) (Value, error) {
	if len(args) < 1 || len(args) > 4 {
		return ErrorVal(ErrValVALUE), nil
	}
	arr := args[0]
	if arr.Type != ValueArray || len(arr.Array) == 0 {
		return arr, nil
	}

	sortIndex := 1
	if len(args) >= 2 {
		si, e := coerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		sortIndex = int(si)
	}

	sortOrder := 1 // 1 = ascending, -1 = descending
	if len(args) >= 3 {
		so, e := coerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		sortOrder = int(so)
	}

	// Make a copy of the array
	rows := make([][]Value, len(arr.Array))
	for i, row := range arr.Array {
		rows[i] = make([]Value, len(row))
		copy(rows[i], row)
	}

	si := sortIndex - 1
	sort.SliceStable(rows, func(i, j int) bool {
		var vi, vj Value
		if si >= 0 && si < len(rows[i]) {
			vi = rows[i][si]
		}
		if si >= 0 && si < len(rows[j]) {
			vj = rows[j][si]
		}
		cmp := compareValues(vi, vj)
		if sortOrder < 0 {
			return cmp > 0
		}
		return cmp < 0
	})

	return Value{Type: ValueArray, Array: rows}, nil
}
