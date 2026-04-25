package formula

import "sort"

func init() {
	Register("AND", NoCtx(fnAND))
	Register("FALSE", NoCtx(fnFALSE))
	Register("IF", NoCtx(fnIF))
	RegisterWithMeta("IFERROR", NoCtx(fnIFERROR), FuncMeta{
		Kind:               FnKindScalarLifted,
		InheritedArrayArgs: map[int]bool{0: true, 1: true},
	})
	Register("IFS", NoCtx(fnIFS))
	RegisterScalarLifted("NOT", NoCtx(fnNOT))
	Register("OR", NoCtx(fnOR))
	Register("SORT", NoCtx(fnSORT))
	RegisterWithSpec("SORTBY", NoCtx(fnSORTBY), gridShapeFuncSpec(evalSORTBY))
	Register("SWITCH", NoCtx(fnSWITCH))
	Register("TRUE", NoCtx(fnTRUE))
	Register("XOR", NoCtx(fnXOR))
}

func fnTRUE(args []Value) (Value, error) {
	if len(args) != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(true), nil
}

func fnFALSE(args []Value) (Value, error) {
	if len(args) != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	return BoolVal(false), nil
}

func fnIF(args []Value) (Value, error) {
	if len(args) < 2 || len(args) > 3 {
		return ErrorVal(ErrValVALUE), nil
	}
	// Array formula: when the condition is an array, apply IF element-wise.
	if args[0].Type == ValueArray {
		cond := args[0]
		rowCount, colCount := arrayOpBounds(cond)
		rows := newValueMatrix(rowCount, colCount)
		for i := 0; i < rowCount; i++ {
			for j := 0; j < colCount; j++ {
				cell := ArrayElement(cond, i, j)
				if IsTruthy(cell) {
					rows[i][j] = ArrayElement(args[1], i, j)
				} else if len(args) == 3 {
					rows[i][j] = ArrayElement(args[2], i, j)
				} else {
					rows[i][j] = BoolVal(false)
				}
			}
		}
		out := Value{Type: ValueArray, Array: rows}
		out.RangeOrigin = combinedArrayOpRangeOrigin(rowCount, colCount, cond)
		return out, nil
	}
	if args[0].Type == ValueError {
		return args[0], nil
	}
	// IF requires a numeric or boolean condition.
	// Strings that can be coerced to numbers are allowed; others cause #VALUE!.
	if args[0].Type == ValueString {
		n, e := CoerceNum(args[0])
		if e != nil {
			return *e, nil
		}
		if n != 0 {
			return args[1], nil
		}
		if len(args) == 3 {
			return args[2], nil
		}
		return BoolVal(false), nil
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
		// Direct string argument → #VALUE! (expected behaviour).
		if arg.Type == ValueString {
			return ErrorVal(ErrValVALUE), nil
		}
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return cell, nil
					}
					// Skip strings in ranges.
					if cell.Type == ValueString {
						continue
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
		// Direct string argument → #VALUE! (expected behaviour).
		if arg.Type == ValueString {
			return ErrorVal(ErrValVALUE), nil
		}
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return cell, nil
					}
					// Skip strings in ranges.
					if cell.Type == ValueString {
						continue
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

func fnXOR(args []Value) (Value, error) {
	if len(args) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	count := 0
	for _, arg := range args {
		if arg.Type == ValueError {
			return arg, nil
		}
		// XOR skips strings whether direct or from ranges.
		if arg.Type == ValueString {
			continue
		}
		if arg.Type == ValueArray {
			for _, row := range arg.Array {
				for _, cell := range row {
					if cell.Type == ValueError {
						return cell, nil
					}
					// Skip strings in ranges.
					if cell.Type == ValueString {
						continue
					}
					if IsTruthy(cell) {
						count++
					}
				}
			}
		} else if IsTruthy(arg) {
			count++
		}
	}
	return BoolVal(count%2 == 1), nil
}

func fnIFS(args []Value) (Value, error) {
	if len(args) < 2 || len(args)%2 != 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	for i := 0; i < len(args); i += 2 {
		if args[i].Type == ValueError {
			return args[i], nil
		}
		if IsTruthy(args[i]) {
			return args[i+1], nil
		}
	}
	return ErrorVal(ErrValNA), nil
}

func fnSWITCH(args []Value) (Value, error) {
	if len(args) < 3 {
		return ErrorVal(ErrValVALUE), nil
	}

	expr := args[0]
	if expr.Type == ValueError {
		return expr, nil
	}
	rest := args[1:]

	hasDefault := len(rest)%2 == 1
	pairCount := len(rest) / 2

	for i := 0; i < pairCount; i++ {
		val := rest[i*2]
		result := rest[i*2+1]
		if CompareValues(expr, val) == 0 {
			return result, nil
		}
	}

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
	grid, errVal := normalizeToArrayGrid(arr)
	if errVal != nil {
		return *errVal, nil
	}

	sortIndex := 1
	if len(args) >= 2 {
		si, e := CoerceNum(args[1])
		if e != nil {
			return *e, nil
		}
		sortIndex = int(si)
	}

	sortOrder := 1
	if len(args) >= 3 {
		so, e := CoerceNum(args[2])
		if e != nil {
			return *e, nil
		}
		sortOrder = int(so)
	}

	rows := grid.matrix()

	si := sortIndex - 1
	sort.SliceStable(rows, func(i, j int) bool {
		var vi, vj Value
		if si >= 0 && si < len(rows[i]) {
			vi = rows[i][si]
		}
		if si >= 0 && si < len(rows[j]) {
			vj = rows[j][si]
		}
		cmp := CompareValues(vi, vj)
		if sortOrder < 0 {
			return cmp > 0
		}
		return cmp < 0
	})

	return Value{Type: ValueArray, Array: rows}, nil
}

func fnSORTBY(args []Value) (Value, error) {
	return sortbyCore(args, nil)
}

func evalSORTBY(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	return ValueToEvalValue(sortbyCoreEval(args, nil)), nil
}

func sortbyCore(args []Value, evalArgs []EvalValue) (Value, error) {
	if len(args) < 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	arrSource, errVal := normalizeGridShapeArg(args[0], evalArgAt(evalArgs, 0))
	if errVal != nil {
		return *errVal, nil
	}
	numRows, numCols := arrSource.dims()

	// Parse (by_array, sort_order) pairs from remaining args.
	// After args[0], remaining args are grouped in pairs of 2: (by_array, sort_order).
	// If the last group has only 1 element, sort_order defaults to 1.
	remaining := args[1:]
	type sortKey struct {
		valueAt func(int) Value
		order   int // 1 = ascending, -1 = descending
	}
	var keys []sortKey
	sortByCols := false
	sortByColsSet := false

	for i := 0; i < len(remaining); i += 2 {
		bySource, errVal := normalizeGridShapeArg(remaining[i], evalArgAt(evalArgs, i+1))
		if errVal != nil {
			return *errVal, nil
		}

		// Validate: by_array must be a vector (single row or single column).
		byRows, byCols := bySource.dims()
		if byRows > 1 && byCols > 1 {
			return ErrorVal(ErrValVALUE), nil
		}

		keySortByCols := byRows == 1 && byCols > 1
		if !sortByColsSet {
			sortByCols = keySortByCols
			sortByColsSet = true
		} else if sortByCols != keySortByCols {
			return ErrorVal(ErrValVALUE), nil
		}

		var valueAt func(int) Value
		keyLen := byRows
		expectedLen := numRows
		if sortByCols {
			keyLen = byCols
			expectedLen = numCols
			valueAt = func(col int) Value { return bySource.cell(0, col) }
		} else {
			valueAt = func(row int) Value { return bySource.cell(row, 0) }
		}

		// Validate: by_array length must match the sorted axis.
		if keyLen != expectedLen {
			return ErrorVal(ErrValVALUE), nil
		}

		// Parse sort_order (defaults to 1 if not provided).
		order := 1
		if i+1 < len(remaining) {
			so, e := CoerceNum(legacyArgValue(remaining[i+1], evalArgAt(evalArgs, i+2)))
			if e != nil {
				return *e, nil
			}
			order = int(so)
			if order != 1 && order != -1 {
				return ErrorVal(ErrValVALUE), nil
			}
		}

		keys = append(keys, sortKey{valueAt: valueAt, order: order})
	}

	if len(keys) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Build index slice and sort using stable sort.
	count := numRows
	if sortByCols {
		count = numCols
	}
	indices := make([]int, count)
	for i := range indices {
		indices[i] = i
	}

	sort.SliceStable(indices, func(a, b int) bool {
		for _, k := range keys {
			cmp := CompareValues(k.valueAt(indices[a]), k.valueAt(indices[b]))
			if cmp != 0 {
				if k.order < 0 {
					return cmp > 0
				}
				return cmp < 0
			}
		}
		return false
	})

	// Build result array by reordering the selected axis.
	if sortByCols {
		return Value{Type: ValueArray, Array: arrSource.materializeCols(indices)}, nil
	}
	return Value{Type: ValueArray, Array: arrSource.materializeRows(indices)}, nil
}
