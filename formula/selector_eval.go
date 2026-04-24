package formula

func indexFuncSpec(eval EvalFunc) FuncSpec {
	return FuncSpec{
		Kind: FnKindLookup,
		Args: []ArgSpec{
			{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
		},
		VarArg: func(_ int) ArgSpec {
			return ArgSpec{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough}
		},
		Return: ReturnModePassThrough,
		Eval:   eval,
	}
}

func selectorFuncSpec(eval EvalFunc) FuncSpec {
	return indexFuncSpec(eval)
}

type selectorSource struct {
	grid        Grid
	rows        int
	cols        int
	matRows     int
	matCols     int
	rangeOrigin *RangeAddr
	directRef   bool
}

func newSelectorSource(v EvalValue) selectorSource {
	switch v.Kind {
	case EvalRef:
		rows, cols := 0, 0
		var origin *RangeAddr
		var grid Grid
		if v.Ref != nil {
			rows = v.Ref.ToRow - v.Ref.FromRow + 1
			cols = v.Ref.ToCol - v.Ref.FromCol + 1
			origin = ptrRange(v.Ref.Bounds())
			grid = v.Ref.Materialized
		}
		src := selectorSource{
			grid:        grid,
			rows:        rows,
			cols:        cols,
			rangeOrigin: origin,
			directRef:   true,
		}
		if grid != nil {
			src.matRows = grid.Rows()
			src.matCols = grid.Cols()
		}
		return src
	case EvalArray:
		if v.Array == nil {
			return selectorSource{}
		}
		src := selectorSource{
			grid: v.Array.Grid,
			rows: v.Array.Rows,
			cols: v.Array.Cols,
		}
		if v.Array.Origin != nil && v.Array.Origin.Range != nil {
			src.rangeOrigin = ptrRange(*v.Array.Origin.Range)
		}
		if src.grid != nil {
			src.matRows = src.grid.Rows()
			src.matCols = src.grid.Cols()
		}
		return src
	case EvalScalar:
		grid := newLegacyValueGrid([][]Value{{v.Scalar}})
		return selectorSource{
			grid:    grid,
			rows:    1,
			cols:    1,
			matRows: 1,
			matCols: 1,
		}
	default:
		return selectorSource{}
	}
}

func (s selectorSource) cell(row, col int) Value {
	if row < 0 || col < 0 || row >= s.rows || col >= s.cols {
		return ErrorVal(ErrValNA)
	}
	return s.gridCell(row, col, EmptyVal())
}

func (s selectorSource) gridCell(row, col int, fallback Value) Value {
	if s.grid == nil {
		return fallback
	}
	if row < 0 || col < 0 || row >= s.matRows || col >= s.matCols {
		return fallback
	}
	return EvalValueToValue(s.grid.Cell(row, col))
}

func (s selectorSource) selectorDims() (rows, cols int) {
	rows, cols = s.rows, s.cols
	if s.directRef && s.rangeOrigin != nil && rangeAddrUsesFullSheetAxis(*s.rangeOrigin) {
		return s.matRows, s.matCols
	}
	return rows, cols
}

func rangeAddrUsesFullSheetAxis(addr RangeAddr) bool {
	return addr.ToRow >= maxRows || addr.ToCol >= maxCols
}

func selectorScalarInt(arg EvalValue) (int, *Value) {
	switch arg.Kind {
	case EvalKindError:
		errVal := ErrorVal(arg.Err)
		return 0, &errVal
	case EvalScalar:
		n, e := CoerceNum(arg.Scalar)
		if e != nil {
			return 0, e
		}
		return int(n), nil
	default:
		errVal := ErrorVal(ErrValVALUE)
		return 0, &errVal
	}
}

func selectorTakeBounds(count, max int) (start, length int, errVal *Value) {
	if count == 0 {
		err := ErrorVal(ErrValVALUE)
		return 0, 0, &err
	}
	if count > 0 {
		if count > max {
			err := ErrorVal(ErrValVALUE)
			return 0, 0, &err
		}
		return 0, count, nil
	}
	if -count > max {
		err := ErrorVal(ErrValVALUE)
		return 0, 0, &err
	}
	return max + count, -count, nil
}

func selectorDropBounds(count, max int) (start, length int, errVal *Value) {
	if count >= 0 {
		start = count
		length = max - count
	} else {
		start = 0
		length = max + count
	}
	if length <= 0 {
		err := ErrorVal(ErrValVALUE)
		return 0, 0, &err
	}
	return start, length, nil
}

func normalizeChooserSelector(idx, max int) (int, *Value) {
	if idx == 0 || idx > max || idx < -max {
		errVal := ErrorVal(ErrValVALUE)
		return 0, &errVal
	}
	if idx < 0 {
		idx = max + idx + 1
	}
	return idx - 1, nil
}

func chooserSelectorEvalValues(selector EvalValue, max int) ([]int, *Value) {
	switch selector.Kind {
	case EvalKindError:
		errVal := ErrorVal(selector.Err)
		return nil, &errVal
	case EvalScalar:
		n, e := CoerceNum(selector.Scalar)
		if e != nil {
			return nil, e
		}
		idx, errVal := normalizeChooserSelector(int(n), max)
		if errVal != nil {
			return nil, errVal
		}
		return []int{idx}, nil
	}

	src := newSelectorSource(selector)
	rows, cols := src.selectorDims()
	if rows == 0 || cols == 0 {
		errVal := ErrorVal(ErrValVALUE)
		return nil, &errVal
	}

	out := make([]int, 0, rows*cols)
	for i := 0; i < rows; i++ {
		for j := 0; j < cols; j++ {
			cell := src.cell(i, j)
			if cell.Type == ValueError {
				return nil, &cell
			}
			n, e := CoerceNum(cell)
			if e != nil {
				return nil, e
			}
			idx, errVal := normalizeChooserSelector(int(n), max)
			if errVal != nil {
				return nil, errVal
			}
			out = append(out, idx)
		}
	}
	if len(out) == 0 {
		errVal := ErrorVal(ErrValVALUE)
		return nil, &errVal
	}
	return out, nil
}

func selectorContiguousRun(indices []int) (start int, ok bool) {
	if len(indices) == 0 {
		return 0, false
	}
	start = indices[0]
	for i := 1; i < len(indices); i++ {
		if indices[i] != start+i {
			return 0, false
		}
	}
	return start, true
}

func selectorWindowOrigin(src selectorSource, rowStart, logicalRows, colStart, logicalCols int) *RangeAddr {
	if src.rangeOrigin == nil {
		return nil
	}
	return ptrRange(RangeAddr{
		Sheet:   src.rangeOrigin.Sheet,
		FromCol: src.rangeOrigin.FromCol + colStart,
		FromRow: src.rangeOrigin.FromRow + rowStart,
		ToCol:   src.rangeOrigin.FromCol + colStart + logicalCols - 1,
		ToRow:   src.rangeOrigin.FromRow + rowStart + logicalRows - 1,
	})
}

func selectorWindowSpillClass(origin *RangeAddr, preserveRef bool) SpillClass {
	if preserveRef {
		return SpillScalarOnly
	}
	if origin != nil && rangeAddrUsesFullSheetAxis(*origin) {
		return SpillUnbounded
	}
	return SpillBounded
}

func selectorWindowEval(src selectorSource, rowStart, logicalRows, colStart, logicalCols int, preserveRef bool) EvalValue {
	if logicalRows == 1 && logicalCols == 1 {
		return ValueToEvalValue(indexScalarCellValue(src, rowStart, colStart))
	}

	origin := selectorWindowOrigin(src, rowStart, logicalRows, colStart, logicalCols)
	matrix := selectorWindowMatrix(src, rowStart, logicalRows, colStart, logicalCols, preserveRef)
	out := EvalValue{
		Kind: EvalArray,
		Array: &ArrayValue{
			Rows:       logicalRows,
			Cols:       logicalCols,
			Grid:       newLegacyValueGrid(matrix),
			SpillClass: selectorWindowSpillClass(origin, preserveRef),
		},
	}
	if origin != nil {
		out.Array.Origin = &ArrayOrigin{Range: origin}
	}
	return out
}

func selectorWindowMatrix(src selectorSource, rowStart, logicalRows, colStart, logicalCols int, preserveRef bool) [][]Value {
	if preserveRef {
		return indexMaterializedWindow(src, rowStart, logicalRows, colStart, logicalCols)
	}
	rows := newValueMatrix(logicalRows, logicalCols)
	for r := 0; r < logicalRows; r++ {
		for c := 0; c < logicalCols; c++ {
			rows[r][c] = src.cell(rowStart+r, colStart+c)
		}
	}
	return rows
}

func selectorSliceEval(src selectorSource, rowStart, logicalRows, colStart, logicalCols int) EvalValue {
	if src.rangeOrigin != nil {
		return selectorWindowEval(src, rowStart, logicalRows, colStart, logicalCols, false)
	}
	return indexMaterializedSelectionEval(src, logicalRows, logicalCols, func(r, c int) Value {
		return src.cell(rowStart+r, colStart+c)
	})
}

func evalTAKESelector(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) < 2 || len(args) > 3 {
		return ValueToEvalValue(ErrorVal(ErrValVALUE)), nil
	}
	if args[0].Kind == EvalKindError {
		return args[0], nil
	}

	src := newSelectorSource(args[0])
	if src.rows == 0 || src.cols == 0 {
		return ValueToEvalValue(ErrorVal(ErrValVALUE)), nil
	}

	rowStart, outRows, errVal := func() (int, int, *Value) {
		count, errVal := selectorScalarInt(args[1])
		if errVal != nil {
			return 0, 0, errVal
		}
		return selectorTakeBounds(count, src.rows)
	}()
	if errVal != nil {
		return ValueToEvalValue(*errVal), nil
	}

	colStart, outCols := 0, src.cols
	if len(args) == 3 {
		count, errVal := selectorScalarInt(args[2])
		if errVal != nil {
			return ValueToEvalValue(*errVal), nil
		}
		colStart, outCols, errVal = selectorTakeBounds(count, src.cols)
		if errVal != nil {
			return ValueToEvalValue(*errVal), nil
		}
	}

	return selectorSliceEval(src, rowStart, outRows, colStart, outCols), nil
}

func evalDROPSelector(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) < 2 || len(args) > 3 {
		return ValueToEvalValue(ErrorVal(ErrValVALUE)), nil
	}
	if args[0].Kind == EvalKindError {
		return args[0], nil
	}

	src := newSelectorSource(args[0])
	if src.rows == 0 || src.cols == 0 {
		return ValueToEvalValue(ErrorVal(ErrValVALUE)), nil
	}

	rowStart, outRows, errVal := func() (int, int, *Value) {
		count, errVal := selectorScalarInt(args[1])
		if errVal != nil {
			return 0, 0, errVal
		}
		return selectorDropBounds(count, src.rows)
	}()
	if errVal != nil {
		return ValueToEvalValue(*errVal), nil
	}

	colStart, outCols := 0, src.cols
	if len(args) == 3 {
		count, errVal := selectorScalarInt(args[2])
		if errVal != nil {
			return ValueToEvalValue(*errVal), nil
		}
		colStart, outCols, errVal = selectorDropBounds(count, src.cols)
		if errVal != nil {
			return ValueToEvalValue(*errVal), nil
		}
	}

	return selectorSliceEval(src, rowStart, outRows, colStart, outCols), nil
}

func evalCHOOSECOLSSelector(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) < 2 {
		return ValueToEvalValue(ErrorVal(ErrValVALUE)), nil
	}
	if args[0].Kind == EvalKindError {
		return args[0], nil
	}

	src := newSelectorSource(args[0])
	if src.rows == 0 || src.cols == 0 {
		return ValueToEvalValue(ErrorVal(ErrValVALUE)), nil
	}

	selectCols := make([]int, 0, len(args)-1)
	for _, arg := range args[1:] {
		cols, errVal := chooserSelectorEvalValues(arg, src.cols)
		if errVal != nil {
			return ValueToEvalValue(*errVal), nil
		}
		selectCols = append(selectCols, cols...)
	}

	if start, ok := selectorContiguousRun(selectCols); ok && src.rangeOrigin != nil {
		return selectorWindowEval(src, 0, src.rows, start, len(selectCols), false), nil
	}
	return indexMaterializedSelectionEval(src, src.rows, len(selectCols), func(r, c int) Value {
		return src.cell(r, selectCols[c])
	}), nil
}

func evalCHOOSEROWSSelector(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) < 2 {
		return ValueToEvalValue(ErrorVal(ErrValVALUE)), nil
	}
	if args[0].Kind == EvalKindError {
		return args[0], nil
	}

	src := newSelectorSource(args[0])
	if src.rows == 0 || src.cols == 0 {
		return ValueToEvalValue(ErrorVal(ErrValVALUE)), nil
	}

	selectRows := make([]int, 0, len(args)-1)
	for _, arg := range args[1:] {
		rows, errVal := chooserSelectorEvalValues(arg, src.rows)
		if errVal != nil {
			return ValueToEvalValue(*errVal), nil
		}
		selectRows = append(selectRows, rows...)
	}

	if start, ok := selectorContiguousRun(selectRows); ok && src.rangeOrigin != nil {
		return selectorWindowEval(src, start, len(selectRows), 0, src.cols, false), nil
	}
	return indexMaterializedSelectionEval(src, len(selectRows), src.cols, func(r, c int) Value {
		return src.cell(selectRows[r], c)
	}), nil
}

func evalINDEXSelector(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) < 2 || len(args) > 3 {
		return ValueToEvalValue(ErrorVal(ErrValVALUE)), nil
	}
	switch args[0].Kind {
	case EvalKindError:
		return args[0], nil
	case EvalScalar:
		return args[0], nil
	}

	src := newSelectorSource(args[0])
	rows, cols := src.rows, src.cols

	rowSelector := args[1]
	colSelector := EvalValue{Kind: EvalScalar, Scalar: NumberVal(1)}
	if len(args) == 3 {
		colSelector = args[2]
	} else if rows == 1 {
		rowSelector = EvalValue{Kind: EvalScalar, Scalar: NumberVal(1)}
		colSelector = args[1]
	}

	rowVals, rowShapeRows, rowShapeCols, errVal := indexSelectorEvalValues(rowSelector, rows)
	if errVal != nil {
		return ValueToEvalValue(*errVal), nil
	}
	colVals, colShapeRows, colShapeCols, errVal := indexSelectorEvalValues(colSelector, cols)
	if errVal != nil {
		return ValueToEvalValue(*errVal), nil
	}

	if len(rowVals) > 1 {
		for _, v := range rowVals {
			if v == 0 {
				return ValueToEvalValue(ErrorVal(ErrValVALUE)), nil
			}
		}
	}
	if len(colVals) > 1 {
		for _, v := range colVals {
			if v == 0 {
				return ValueToEvalValue(ErrorVal(ErrValVALUE)), nil
			}
		}
	}

	rowScalar := len(rowVals) == 1
	colScalar := len(colVals) == 1

	if rowScalar && colScalar {
		return indexScalarSelectionEval(src, rowVals[0], colVals[0]), nil
	}
	if !rowScalar && colScalar {
		if colVals[0] == 0 {
			return indexMaterializedSelectionEval(src, len(rowVals), cols, func(r, c int) Value {
				return src.cell(rowVals[r]-1, c)
			}), nil
		}
		return indexMaterializedSelectionEval(src, rowShapeRows, rowShapeCols, func(r, c int) Value {
			return src.cell(rowVals[r*rowShapeCols+c]-1, colVals[0]-1)
		}), nil
	}
	if rowScalar && !colScalar {
		if rowVals[0] == 0 {
			return indexMaterializedSelectionEval(src, rows, len(colVals), func(r, c int) Value {
				return src.cell(r, colVals[c]-1)
			}), nil
		}
		return indexMaterializedSelectionEval(src, colShapeRows, colShapeCols, func(r, c int) Value {
			return src.cell(rowVals[0]-1, colVals[r*colShapeCols+c]-1)
		}), nil
	}
	return indexMaterializedSelectionEval(src, len(rowVals), len(colVals), func(r, c int) Value {
		return src.cell(rowVals[r]-1, colVals[c]-1)
	}), nil
}

func indexSelectorEvalValues(selector EvalValue, max int) ([]int, int, int, *Value) {
	switch selector.Kind {
	case EvalKindError:
		errVal := ErrorVal(selector.Err)
		return nil, 0, 0, &errVal
	case EvalScalar:
		n, e := CoerceNum(selector.Scalar)
		if e != nil {
			return nil, 0, 0, e
		}
		idx, errVal := normalizeIndexSelector(int(n), max)
		if errVal != nil {
			return nil, 0, 0, errVal
		}
		return []int{idx}, 1, 1, nil
	}

	src := newSelectorSource(selector)
	rows, cols := src.selectorDims()
	if rows == 0 || cols == 0 {
		errVal := ErrorVal(ErrValVALUE)
		return nil, 0, 0, &errVal
	}

	out := make([]int, 0, 16)
	for i := 0; i < rows; i++ {
		for j := 0; j < cols; j++ {
			cell := src.cell(i, j)
			if cell.Type == ValueError {
				return nil, 0, 0, &cell
			}
			n, e := CoerceNum(cell)
			if e != nil {
				return nil, 0, 0, e
			}
			idx, errVal := normalizeIndexSelector(int(n), max)
			if errVal != nil {
				return nil, 0, 0, errVal
			}
			out = append(out, idx)
		}
	}
	return out, rows, cols, nil
}

func indexScalarSelectionEval(src selectorSource, rowNum, colNum int) EvalValue {
	if rowNum == 0 && colNum == 0 {
		if src.directRef {
			return indexDirectRefArrayResult(src, 0, src.rows, 0, src.cols)
		}
		return indexMaterializedSelectionEval(src, src.rows, src.cols, func(r, c int) Value {
			return src.cell(r, c)
		})
	}
	if rowNum == 0 {
		if src.directRef {
			return indexDirectRefArrayResult(src, 0, src.rows, colNum-1, 1)
		}
		return indexMaterializedSelectionEval(src, src.rows, 1, func(r, _ int) Value {
			return src.cell(r, colNum-1)
		})
	}
	if colNum == 0 {
		if src.directRef {
			return indexDirectRefArrayResult(src, rowNum-1, 1, 0, src.cols)
		}
		return indexMaterializedSelectionEval(src, 1, src.cols, func(_ int, c int) Value {
			return src.cell(rowNum-1, c)
		})
	}
	return ValueToEvalValue(indexScalarCellValue(src, rowNum-1, colNum-1))
}

func indexScalarCellValue(src selectorSource, rowIdx, colIdx int) Value {
	v := src.cell(rowIdx, colIdx)
	if src.rangeOrigin != nil {
		v.CellOrigin = &CellAddr{
			Sheet: src.rangeOrigin.Sheet,
			Col:   src.rangeOrigin.FromCol + colIdx,
			Row:   src.rangeOrigin.FromRow + rowIdx,
		}
	}
	return v
}

func indexDirectRefArrayResult(src selectorSource, rowStart, logicalRows, colStart, logicalCols int) EvalValue {
	return selectorWindowEval(src, rowStart, logicalRows, colStart, logicalCols, true)
}

func indexMaterializedWindow(src selectorSource, rowStart, logicalRows, colStart, logicalCols int) [][]Value {
	matRows := 0
	if rowStart < src.matRows {
		matRows = src.matRows - rowStart
		if logicalRows < matRows {
			matRows = logicalRows
		}
	}
	matCols := 0
	if colStart < src.matCols {
		matCols = src.matCols - colStart
		if logicalCols < matCols {
			matCols = logicalCols
		}
	}
	if matRows <= 0 || matCols <= 0 {
		return nil
	}
	rows := newValueMatrix(matRows, matCols)
	for r := 0; r < matRows; r++ {
		for c := 0; c < matCols; c++ {
			rows[r][c] = src.gridCell(rowStart+r, colStart+c, EmptyVal())
		}
	}
	return rows
}

func indexMaterializedSelectionEval(src selectorSource, outRows, outCols int, fill func(r, c int) Value) EvalValue {
	if outRows == 1 && outCols == 1 {
		return ValueToEvalValue(fill(0, 0))
	}
	rows := newValueMatrix(outRows, outCols)
	for r := 0; r < outRows; r++ {
		for c := 0; c < outCols; c++ {
			rows[r][c] = fill(r, c)
		}
	}
	return EvalValue{
		Kind: EvalArray,
		Array: &ArrayValue{
			Rows:       outRows,
			Cols:       outCols,
			Grid:       newLegacyValueGrid(rows),
			SpillClass: SpillBounded,
		},
	}
}
