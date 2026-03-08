package formula

import "sort"

func init() {
	Register("AND", NoCtx(fnAND))
	Register("FALSE", NoCtx(fnFALSE))
	Register("IF", NoCtx(fnIF))
	Register("IFERROR", NoCtx(fnIFERROR))
	Register("IFS", NoCtx(fnIFS))
	Register("NOT", NoCtx(fnNOT))
	Register("OR", NoCtx(fnOR))
	Register("SORT", NoCtx(fnSORT))
	Register("SORTBY", NoCtx(fnSORTBY))
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
	// Excel's IF requires a numeric or boolean condition.
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
		// Direct string argument → #VALUE! (Excel behaviour).
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
		// Direct string argument → #VALUE! (Excel behaviour).
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
		cmp := CompareValues(vi, vj)
		if sortOrder < 0 {
			return cmp > 0
		}
		return cmp < 0
	})

	return Value{Type: ValueArray, Array: rows}, nil
}

func fnSORTBY(args []Value) (Value, error) {
	if len(args) < 2 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Normalize array argument to a 2D grid.
	arr := args[0]
	var grid [][]Value
	switch arr.Type {
	case ValueArray:
		grid = arr.Array
	case ValueError:
		return arr, nil
	default:
		grid = [][]Value{{arr}}
	}
	if len(grid) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	numRows := len(grid)

	// Parse (by_array, sort_order) pairs from remaining args.
	// After args[0], remaining args are grouped in pairs of 2: (by_array, sort_order).
	// If the last group has only 1 element, sort_order defaults to 1.
	remaining := args[1:]
	type sortKey struct {
		values []Value
		order  int // 1 = ascending, -1 = descending
	}
	var keys []sortKey

	for i := 0; i < len(remaining); i += 2 {
		byArg := remaining[i]

		// Propagate errors from by_array.
		if byArg.Type == ValueError {
			return byArg, nil
		}

		// Normalize by_array to a 2D grid, then flatten to a 1D vector.
		var byGrid [][]Value
		switch byArg.Type {
		case ValueArray:
			byGrid = byArg.Array
		default:
			byGrid = [][]Value{{byArg}}
		}

		// Validate: by_array must be a vector (single row or single column).
		byRows := len(byGrid)
		byCols := 0
		for _, row := range byGrid {
			if len(row) > byCols {
				byCols = len(row)
			}
		}
		if byRows > 1 && byCols > 1 {
			return ErrorVal(ErrValVALUE), nil
		}

		// Flatten to 1D.
		var flat []Value
		if byCols <= 1 {
			// Column vector or single cell: one value per row.
			for _, row := range byGrid {
				if len(row) > 0 {
					flat = append(flat, row[0])
				} else {
					flat = append(flat, EmptyVal())
				}
			}
		} else {
			// Row vector: one value per column.
			if len(byGrid) > 0 {
				flat = byGrid[0]
			}
		}

		// Validate: by_array length must match number of rows in array.
		if len(flat) != numRows {
			return ErrorVal(ErrValVALUE), nil
		}

		// Check for errors in the by_array values.
		for _, v := range flat {
			if v.Type == ValueError {
				return v, nil
			}
		}

		// Parse sort_order (defaults to 1 if not provided).
		order := 1
		if i+1 < len(remaining) {
			so, e := CoerceNum(remaining[i+1])
			if e != nil {
				return *e, nil
			}
			order = int(so)
			if order != 1 && order != -1 {
				return ErrorVal(ErrValVALUE), nil
			}
		}

		keys = append(keys, sortKey{values: flat, order: order})
	}

	if len(keys) == 0 {
		return ErrorVal(ErrValVALUE), nil
	}

	// Build index slice and sort using stable sort.
	indices := make([]int, numRows)
	for i := range indices {
		indices[i] = i
	}

	sort.SliceStable(indices, func(a, b int) bool {
		for _, k := range keys {
			cmp := CompareValues(k.values[indices[a]], k.values[indices[b]])
			if cmp != 0 {
				if k.order < 0 {
					return cmp > 0
				}
				return cmp < 0
			}
		}
		return false
	})

	// Build result array by reordering rows.
	result := make([][]Value, numRows)
	for i, idx := range indices {
		row := make([]Value, len(grid[idx]))
		copy(row, grid[idx])
		result[i] = row
	}

	return Value{Type: ValueArray, Array: result}, nil
}
