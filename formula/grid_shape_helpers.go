package formula

type gridValueSource struct {
	grid        Grid
	rows        int
	cols        int
	matRows     int
	matCols     int
	rangeOrigin *RangeAddr
	fullAxisRef bool
	scalar      *Value
}

func gridShapeFuncSpec(eval EvalFunc) FuncSpec {
	return FuncSpec{
		Kind: FnKindLookup,
		Args: []ArgSpec{{
			Load:  ArgLoadPassthrough,
			Adapt: ArgAdaptPassThrough,
		}},
		VarArg: func(_ int) ArgSpec {
			return ArgSpec{
				Load:  ArgLoadPassthrough,
				Adapt: ArgAdaptPassThrough,
			}
		},
		Return: ReturnModePassThrough,
		Eval:   eval,
	}
}

func newGridValueSource(v EvalValue) gridValueSource {
	switch v.Kind {
	case EvalScalar:
		scalar := v.Scalar
		return gridValueSource{scalar: &scalar}
	case EvalArray:
		if v.Array == nil {
			return gridValueSource{}
		}
		src := gridValueSource{
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
	case EvalRef:
		if v.Ref == nil {
			return gridValueSource{}
		}
		src := gridValueSource{
			grid:        criteriaRefGrid(v.Ref),
			rows:        v.Ref.ToRow - v.Ref.FromRow + 1,
			cols:        v.Ref.ToCol - v.Ref.FromCol + 1,
			rangeOrigin: ptrRange(v.Ref.Bounds()),
		}
		if src.rangeOrigin != nil {
			src.fullAxisRef = rangeAddrUsesFullSheetAxis(*src.rangeOrigin)
		}
		if src.grid != nil {
			src.matRows = src.grid.Rows()
			src.matCols = src.grid.Cols()
		}
		return src
	default:
		return gridValueSource{}
	}
}

func normalizeEvalGridValueSource(v EvalValue) (gridValueSource, *Value) {
	if v.Kind == EvalKindError {
		errVal := ErrorVal(v.Err)
		return gridValueSource{}, &errVal
	}
	src := newGridValueSource(v)
	rows, cols := src.dims()
	if rows == 0 || cols == 0 {
		errVal := ErrorVal(ErrValVALUE)
		return gridValueSource{}, &errVal
	}
	return src, nil
}

func normalizeValueGridSource(v Value) (gridValueSource, *Value) {
	return normalizeEvalGridValueSource(ValueToEvalValue(v))
}

func (s gridValueSource) dims() (rows, cols int) {
	if s.scalar != nil {
		return 1, 1
	}
	if s.fullAxisRef {
		return s.matRows, s.matCols
	}
	if s.rows != 0 || s.cols != 0 {
		return s.rows, s.cols
	}
	if s.grid == nil {
		return 0, 0
	}
	return s.grid.Rows(), s.grid.Cols()
}

func (s gridValueSource) materializedDims() (rows, cols int) {
	if s.scalar != nil {
		return 1, 1
	}
	if s.grid == nil {
		return 0, 0
	}
	return s.matRows, s.matCols
}

func (s gridValueSource) cell(row, col int) Value {
	rows, cols := s.dims()
	if row < 0 || col < 0 || row >= rows || col >= cols {
		return ErrorVal(ErrValNA)
	}
	return s.materializedCell(row, col, EmptyVal())
}

func (s gridValueSource) materializedCell(row, col int, fallback Value) Value {
	if s.scalar != nil {
		if row == 0 && col == 0 {
			return *s.scalar
		}
		return fallback
	}
	if s.grid == nil {
		return fallback
	}
	if row < 0 || col < 0 || row >= s.matRows || col >= s.matCols {
		return fallback
	}
	return EvalValueToValue(s.grid.Cell(row, col))
}

func (s gridValueSource) hasMaterializedCell(row, col int) bool {
	if s.scalar != nil {
		return row == 0 && col == 0
	}
	return s.grid != nil && row >= 0 && col >= 0 && row < s.matRows && col < s.matCols
}

func (s gridValueSource) matrix() [][]Value {
	rows, cols := s.dims()
	if rows == 0 || cols == 0 {
		return nil
	}
	out := newValueMatrix(rows, cols)
	for row := 0; row < rows; row++ {
		for col := 0; col < cols; col++ {
			out[row][col] = s.cell(row, col)
		}
	}
	return out
}

func (s gridValueSource) rowValues(row int) []Value {
	_, cols := s.dims()
	values := make([]Value, cols)
	for col := 0; col < cols; col++ {
		values[col] = s.cell(row, col)
	}
	return values
}

func (s gridValueSource) columnValues(col int) []Value {
	rows, _ := s.dims()
	values := make([]Value, rows)
	for row := 0; row < rows; row++ {
		values[row] = s.cell(row, col)
	}
	return values
}

func flattenGridValueSource(src gridValueSource, scanByCol bool, ignore int) []Value {
	numRows, numCols := src.dims()

	var flat []Value
	if scanByCol {
		for col := 0; col < numCols; col++ {
			for row := 0; row < numRows; row++ {
				v := src.cell(row, col)
				if shouldInclude(v, ignore) {
					flat = append(flat, v)
				}
			}
		}
		return flat
	}
	for row := 0; row < numRows; row++ {
		for col := 0; col < numCols; col++ {
			v := src.cell(row, col)
			if shouldInclude(v, ignore) {
				flat = append(flat, v)
			}
		}
	}
	return flat
}

func (s gridValueSource) flattenRowMajor() []Value {
	return flattenGridValueSource(s, false, 0)
}

func (s gridValueSource) projectRow(row int) Value {
	rows, cols := s.dims()
	if row < 0 || row >= rows {
		return ErrorVal(ErrValNA)
	}
	if cols <= 1 {
		return s.cell(row, 0)
	}
	return Value{Type: ValueArray, Array: [][]Value{s.rowValues(row)}}
}

func (s gridValueSource) projectCol(col int) Value {
	rows, cols := s.dims()
	if col < 0 || col >= cols {
		return ErrorVal(ErrValNA)
	}
	if rows <= 1 {
		return s.cell(0, col)
	}
	values := s.columnValues(col)
	out := make([][]Value, len(values))
	for i, v := range values {
		out[i] = []Value{v}
	}
	return Value{Type: ValueArray, Array: out}
}

func (s gridValueSource) materializeRows(indices []int) [][]Value {
	if len(indices) == 0 {
		return nil
	}
	_, cols := s.dims()
	out := newValueMatrix(len(indices), cols)
	for i, row := range indices {
		for col := 0; col < cols; col++ {
			out[i][col] = s.cell(row, col)
		}
	}
	return out
}

func (s gridValueSource) materializeCols(indices []int) [][]Value {
	rows, _ := s.dims()
	if rows == 0 || len(indices) == 0 {
		return nil
	}
	out := newValueMatrix(rows, len(indices))
	for row := 0; row < rows; row++ {
		for i, col := range indices {
			out[row][i] = s.cell(row, col)
		}
	}
	return out
}

func (s gridValueSource) spilledError(err Value) Value {
	rows, cols := s.dims()
	if rows == 0 || cols == 0 {
		return err
	}
	out := newValueMatrix(rows, cols)
	for row := 0; row < rows; row++ {
		for col := 0; col < cols; col++ {
			out[row][col] = err
		}
	}
	if rows == 1 && cols == 1 {
		return out[0][0]
	}
	return Value{Type: ValueArray, Array: out}
}

func collapseArrayResult(rows [][]Value) Value {
	if len(rows) == 1 && len(rows[0]) == 1 {
		return rows[0][0]
	}
	return Value{Type: ValueArray, Array: rows}
}

func evalArgAt(args []EvalValue, index int) *EvalValue {
	if index < 0 || index >= len(args) {
		return nil
	}
	return &args[index]
}

func normalizeGridShapeArg(legacy Value, evalArg *EvalValue) (gridValueSource, *Value) {
	if evalArg != nil {
		return normalizeEvalGridValueSource(*evalArg)
	}
	return normalizeValueGridSource(legacy)
}

func scalarLegacyArgsFromEval(args []EvalValue) []Value {
	values := make([]Value, len(args))
	for i, arg := range args {
		switch arg.Kind {
		case EvalScalar, EvalKindError:
			values[i] = EvalValueToValue(arg)
		default:
			values[i] = EmptyVal()
		}
	}
	return values
}

func legacyArgValue(legacy Value, evalArg *EvalValue) Value {
	if evalArg == nil {
		return legacy
	}
	return EvalValueToValue(*evalArg)
}

func argTopLevelError(legacy Value, evalArg *EvalValue) *Value {
	if evalArg != nil {
		switch evalArg.Kind {
		case EvalKindError:
			errVal := ErrorVal(evalArg.Err)
			return &errVal
		case EvalScalar:
			if evalArg.Scalar.Type == ValueError {
				errVal := evalArg.Scalar
				return &errVal
			}
		}
		return nil
	}
	if legacy.Type == ValueError {
		errVal := legacy
		return &errVal
	}
	return nil
}

func evalGridShapeCore(args []EvalValue, core func([]Value, []EvalValue) (Value, error)) (EvalValue, error) {
	got, err := core(scalarLegacyArgsFromEval(args), args)
	if err != nil {
		return EvalValue{}, err
	}
	return ValueToEvalValue(got), nil
}

func filterCoreEval(args []EvalValue, legacyArgs []Value) Value {
	values := legacyArgs
	if values == nil {
		values = scalarLegacyArgsFromEval(args)
	}
	got, _ := filterCore(values, args)
	return got
}

func uniqueCoreEval(args []EvalValue, legacyArgs []Value) Value {
	values := legacyArgs
	if values == nil {
		values = scalarLegacyArgsFromEval(args)
	}
	got, _ := uniqueCore(values, args)
	return got
}

func uniqueAxisCount(src gridValueSource, byCol bool) int {
	rows, cols := src.dims()
	if byCol {
		return cols
	}
	return rows
}

func uniqueAxisKey(src gridValueSource, index int, byCol bool) string {
	if byCol {
		return rowKey(src.columnValues(index))
	}
	return rowKey(src.rowValues(index))
}

func uniqueMaterialize(src gridValueSource, keep []int, byCol bool) [][]Value {
	var result [][]Value
	if byCol {
		result = src.materializeCols(keep)
	} else {
		result = src.materializeRows(keep)
	}
	for row := range result {
		for col := range result[row] {
			if result[row][col].Type == ValueEmpty {
				result[row][col] = StringVal("")
			}
		}
	}
	return result
}

func sortbyCoreEval(args []EvalValue, legacyArgs []Value) Value {
	values := legacyArgs
	if values == nil {
		values = scalarLegacyArgsFromEval(args)
	}
	got, _ := sortbyCore(values, args)
	return got
}
