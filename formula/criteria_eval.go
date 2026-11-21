package formula

type criteriaReduceKind uint8

const (
	criteriaReduceCount criteriaReduceKind = iota
	criteriaReduceSum
	criteriaReduceAverage
)

type criteriaEvalPair struct {
	Range    EvalValue
	Criteria EvalValue
}

type criteriaEvalRequest struct {
	ScanRange            EvalValue
	ResultRange          *EvalValue
	Pairs                []criteriaEvalPair
	Reduce               criteriaReduceKind
	CollapseSingleResult bool
}

type criteriaValueSource struct {
	scalar         *Value
	grid           Grid
	rows           int
	cols           int
	legacyRows     [][]Value
	legacyRowCount int
	legacyColCount int
}

type criteriaPreparedPair struct {
	rangeSource    criteriaValueSource
	criteriaSource criteriaValueSource
}

type criteriaAccumulator struct {
	reduce criteriaReduceKind
	sum    float64
	count  int
}

func criteriaSingleIfFuncSpec(eval EvalFunc) FuncSpec {
	return FuncSpec{
		Kind: FnKindReduction,
		Args: []ArgSpec{
			{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
			{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
		},
		Return: ReturnModePassThrough,
		Eval:   eval,
	}
}

func criteriaSingleIfAggregateFuncSpec(eval EvalFunc) FuncSpec {
	return FuncSpec{
		Kind: FnKindReduction,
		Args: []ArgSpec{
			{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
			{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
		},
		Return: ReturnModePassThrough,
		Eval:   eval,
	}
}

func criteriaPairsFuncSpec(eval EvalFunc) FuncSpec {
	return FuncSpec{
		Kind: FnKindReduction,
		VarArg: func(i int) ArgSpec {
			if i%2 == 0 {
				return ArgSpec{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough}
			}
			return ArgSpec{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough}
		},
		Return: ReturnModePassThrough,
		Eval:   eval,
	}
}

func criteriaPairsAggregateFuncSpec(eval EvalFunc) FuncSpec {
	return FuncSpec{
		Kind: FnKindReduction,
		Args: []ArgSpec{
			{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
		},
		VarArg: func(i int) ArgSpec {
			if i == 0 {
				return ArgSpec{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough}
			}
			if i%2 == 1 {
				return ArgSpec{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough}
			}
			return ArgSpec{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough}
		},
		Return: ReturnModePassThrough,
		Eval:   eval,
	}
}

func evalCOUNTIFCriteria(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) != 2 {
		return evalError(ErrValVALUE), nil
	}
	if args[0].Kind == EvalKindError {
		return args[0], nil
	}
	return evalCriteriaRequest(criteriaEvalRequest{
		ScanRange: args[0],
		Pairs: []criteriaEvalPair{{
			Range:    args[0],
			Criteria: args[1],
		}},
		Reduce:               criteriaReduceCount,
		CollapseSingleResult: true,
	}), nil
}

func evalSUMIFCriteria(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) < 2 || len(args) > 3 {
		return evalError(ErrValVALUE), nil
	}
	if args[0].Kind == EvalKindError {
		return args[0], nil
	}
	req := criteriaEvalRequest{
		ScanRange: args[0],
		Pairs: []criteriaEvalPair{{
			Range:    args[0],
			Criteria: args[1],
		}},
		Reduce:               criteriaReduceSum,
		CollapseSingleResult: true,
	}
	if len(args) == 3 {
		req.ResultRange = ptrEvalValue(args[2])
	}
	return evalCriteriaRequest(req), nil
}

func evalAVERAGEIFCriteria(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) < 2 || len(args) > 3 {
		return evalError(ErrValVALUE), nil
	}
	if args[0].Kind == EvalKindError {
		return args[0], nil
	}
	req := criteriaEvalRequest{
		ScanRange: args[0],
		Pairs: []criteriaEvalPair{{
			Range:    args[0],
			Criteria: args[1],
		}},
		Reduce:               criteriaReduceAverage,
		CollapseSingleResult: true,
	}
	if len(args) == 3 {
		req.ResultRange = ptrEvalValue(args[2])
	}
	return evalCriteriaRequest(req), nil
}

func evalCOUNTIFSCriteria(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) < 2 || len(args)%2 != 0 {
		return evalError(ErrValVALUE), nil
	}
	pairs := make([]criteriaEvalPair, 0, len(args)/2)
	for i := 0; i < len(args); i += 2 {
		pairs = append(pairs, criteriaEvalPair{
			Range:    args[i],
			Criteria: args[i+1],
		})
	}
	return evalCriteriaRequest(criteriaEvalRequest{
		ScanRange: args[0],
		Pairs:     pairs,
		Reduce:    criteriaReduceCount,
	}), nil
}

func evalSUMIFSCriteria(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) < 3 || (len(args)-1)%2 != 0 {
		return evalError(ErrValVALUE), nil
	}
	pairs := make([]criteriaEvalPair, 0, (len(args)-1)/2)
	for i := 1; i < len(args); i += 2 {
		pairs = append(pairs, criteriaEvalPair{
			Range:    args[i],
			Criteria: args[i+1],
		})
	}
	return evalCriteriaRequest(criteriaEvalRequest{
		ScanRange:   args[0],
		ResultRange: ptrEvalValue(args[0]),
		Pairs:       pairs,
		Reduce:      criteriaReduceSum,
	}), nil
}

func evalAVERAGEIFSCriteria(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) < 3 || (len(args)-1)%2 != 0 {
		return evalError(ErrValVALUE), nil
	}
	pairs := make([]criteriaEvalPair, 0, (len(args)-1)/2)
	for i := 1; i < len(args); i += 2 {
		pairs = append(pairs, criteriaEvalPair{
			Range:    args[i],
			Criteria: args[i+1],
		})
	}
	return evalCriteriaRequest(criteriaEvalRequest{
		ScanRange:   args[0],
		ResultRange: ptrEvalValue(args[0]),
		Pairs:       pairs,
		Reduce:      criteriaReduceAverage,
	}), nil
}

func evalCriteriaRequest(req criteriaEvalRequest) EvalValue {
	if len(req.Pairs) == 0 {
		return evalError(ErrValVALUE)
	}

	scanSource := newCriteriaValueSource(req.ScanRange)
	resultSource := scanSource
	if req.ResultRange != nil {
		resultSource = newCriteriaValueSource(*req.ResultRange)
	}

	prepared := make([]criteriaPreparedPair, len(req.Pairs))
	broadcastRows := 0
	broadcastCols := 0
	hasBroadcast := false
	for i, pair := range req.Pairs {
		prepared[i] = criteriaPreparedPair{
			rangeSource:    newCriteriaValueSource(pair.Range),
			criteriaSource: newCriteriaValueSource(pair.Criteria),
		}
		if hasBroadcast || prepared[i].criteriaSource.isScalar() {
			continue
		}
		broadcastRows, broadcastCols = prepared[i].criteriaSource.dims()
		hasBroadcast = true
	}

	if !hasBroadcast {
		return evalScalar(evalCriteriaScalar(req.Reduce, scanSource, resultSource, prepared, 0, 0))
	}

	rows := make([][]Value, broadcastRows)
	for r := 0; r < broadcastRows; r++ {
		row := make([]Value, broadcastCols)
		for c := 0; c < broadcastCols; c++ {
			row[c] = evalCriteriaScalar(req.Reduce, scanSource, resultSource, prepared, r, c)
		}
		rows[r] = row
	}
	if req.CollapseSingleResult && broadcastRows == 1 && broadcastCols == 1 {
		return evalScalar(rows[0][0])
	}
	return evalArray(rows, SpillBounded)
}

func evalCriteriaScalar(
	reduce criteriaReduceKind,
	scanSource criteriaValueSource,
	resultSource criteriaValueSource,
	pairs []criteriaPreparedPair,
	criteriaRow int,
	criteriaCol int,
) Value {
	rows, cols := scanSource.dims()
	acc := criteriaAccumulator{reduce: reduce}
	iterRows, iterCols := criteriaMaterializedBounds(rows, cols, scanSource, resultSource, pairs)

	for r := 0; r < iterRows; r++ {
		for c := 0; c < iterCols; c++ {
			if !criteriaMatchesAll(pairs, r, c, criteriaRow, criteriaCol) {
				continue
			}
			if errVal := acc.include(resultSource.alignedCell(r, c)); errVal != nil {
				return *errVal
			}
		}
	}

	if tailCells := rows*cols - iterRows*iterCols; tailCells > 0 &&
		criteriaMatchesAllEmpty(pairs, criteriaRow, criteriaCol) {
		if errVal := acc.includeRepeated(resultSource.tailCell(), tailCells); errVal != nil {
			return *errVal
		}
	}

	return acc.result()
}

func criteriaMaterializedBounds(
	logicalRows int,
	logicalCols int,
	scanSource criteriaValueSource,
	resultSource criteriaValueSource,
	pairs []criteriaPreparedPair,
) (rows, cols int) {
	rows, cols = scanSource.materializedDims()
	expand := func(source criteriaValueSource) {
		r, c := source.materializedDims()
		if r > rows {
			rows = r
		}
		if c > cols {
			cols = c
		}
	}
	expand(resultSource)
	for _, pair := range pairs {
		expand(pair.rangeSource)
	}
	if rows > logicalRows {
		rows = logicalRows
	}
	if cols > logicalCols {
		cols = logicalCols
	}
	return rows, cols
}

func criteriaMatchesAll(
	pairs []criteriaPreparedPair,
	rangeRow int,
	rangeCol int,
	criteriaRow int,
	criteriaCol int,
) bool {
	for _, pair := range pairs {
		if !MatchesCriteria(
			pair.rangeSource.alignedCell(rangeRow, rangeCol),
			pair.criteriaSource.broadcastCell(criteriaRow, criteriaCol),
		) {
			return false
		}
	}
	return true
}

func criteriaMatchesAllEmpty(
	pairs []criteriaPreparedPair,
	criteriaRow int,
	criteriaCol int,
) bool {
	for _, pair := range pairs {
		if !MatchesCriteria(EmptyVal(), pair.criteriaSource.broadcastCell(criteriaRow, criteriaCol)) {
			return false
		}
	}
	return true
}

func (a *criteriaAccumulator) include(v Value) *Value {
	switch a.reduce {
	case criteriaReduceCount:
		a.count++
	case criteriaReduceSum:
		if v.Type == ValueError {
			return &v
		}
		if n, err := CoerceNum(v); err == nil {
			a.sum += n
		}
	case criteriaReduceAverage:
		if v.Type == ValueError {
			return &v
		}
		if v.Type == ValueEmpty {
			return nil
		}
		if n, err := CoerceNum(v); err == nil {
			a.sum += n
			a.count++
		}
	}
	return nil
}

func (a *criteriaAccumulator) includeRepeated(v Value, n int) *Value {
	if n <= 0 {
		return nil
	}
	switch a.reduce {
	case criteriaReduceCount:
		a.count += n
	case criteriaReduceSum:
		if v.Type == ValueError {
			return &v
		}
		if num, err := CoerceNum(v); err == nil {
			a.sum += num * float64(n)
		}
	case criteriaReduceAverage:
		if v.Type == ValueError {
			return &v
		}
		if v.Type == ValueEmpty {
			return nil
		}
		if num, err := CoerceNum(v); err == nil {
			a.sum += num * float64(n)
			a.count += n
		}
	}
	return nil
}

func (a criteriaAccumulator) result() Value {
	switch a.reduce {
	case criteriaReduceCount:
		return NumberVal(float64(a.count))
	case criteriaReduceSum:
		return NumberVal(a.sum)
	case criteriaReduceAverage:
		if a.count == 0 {
			return ErrorVal(ErrValDIV0)
		}
		return NumberVal(a.sum / float64(a.count))
	default:
		return ErrorVal(ErrValVALUE)
	}
}

func newCriteriaValueSource(v EvalValue) criteriaValueSource {
	switch v.Kind {
	case EvalKindError:
		scalar := ErrorVal(v.Err)
		return criteriaValueSource{scalar: &scalar}
	case EvalScalar:
		scalar := v.Scalar
		return criteriaValueSource{scalar: &scalar}
	case EvalArray:
		if v.Array == nil {
			return criteriaValueSource{}
		}
		source := criteriaValueSource{
			grid: v.Array.Grid,
			rows: v.Array.Rows,
			cols: v.Array.Cols,
		}
		source.attachLegacyRows()
		return source
	case EvalRef:
		rows, cols := 0, 0
		if v.Ref != nil {
			rows = v.Ref.ToRow - v.Ref.FromRow + 1
			cols = v.Ref.ToCol - v.Ref.FromCol + 1
		}
		source := criteriaValueSource{
			grid: criteriaRefGrid(v.Ref),
			rows: rows,
			cols: cols,
		}
		source.attachLegacyRows()
		return source
	default:
		scalar := EmptyVal()
		return criteriaValueSource{scalar: &scalar}
	}
}

func (s *criteriaValueSource) attachLegacyRows() {
	if s == nil || s.grid == nil {
		return
	}
	legacy, ok := s.grid.(interface {
		legacyRows() ([][]Value, int, int)
	})
	if !ok {
		return
	}
	s.legacyRows, s.legacyRowCount, s.legacyColCount = legacy.legacyRows()
}

func criteriaRefGrid(ref *RefValue) Grid {
	if ref == nil {
		return emptyRefGrid{}
	}
	if ref.Materialized != nil {
		return ref.Materialized
	}
	return emptyRefGrid{
		rows: ref.ToRow - ref.FromRow + 1,
		cols: ref.ToCol - ref.FromCol + 1,
	}
}

func (s criteriaValueSource) isScalar() bool {
	return s.scalar != nil
}

func (s criteriaValueSource) dims() (rows, cols int) {
	if s.scalar != nil {
		return 1, 1
	}
	if s.rows != 0 || s.cols != 0 {
		return s.rows, s.cols
	}
	if s.grid == nil {
		return 0, 0
	}
	return s.grid.Rows(), s.grid.Cols()
}

func (s criteriaValueSource) materializedDims() (rows, cols int) {
	if s.scalar != nil {
		return 1, 1
	}
	if s.legacyRows != nil {
		return s.legacyRowCount, s.legacyColCount
	}
	if s.grid == nil {
		return 0, 0
	}
	return s.grid.Rows(), s.grid.Cols()
}

func (s criteriaValueSource) alignedCell(row, col int) Value {
	if s.scalar != nil {
		return *s.scalar
	}
	rows, cols := s.dims()
	if row < 0 || col < 0 || row >= rows || col >= cols {
		return EmptyVal()
	}
	return s.materializedCell(row, col, EmptyVal())
}

func (s criteriaValueSource) broadcastCell(row, col int) Value {
	if s.scalar != nil {
		return *s.scalar
	}
	rows, cols := s.dims()
	if row < 0 || col < 0 || row >= rows || col >= cols {
		return ErrorVal(ErrValNA)
	}
	return s.materializedCell(row, col, EmptyVal())
}

func (s criteriaValueSource) tailCell() Value {
	if s.scalar != nil {
		return *s.scalar
	}
	return EmptyVal()
}

func (s criteriaValueSource) materializedCell(row, col int, fallback Value) Value {
	if s.legacyRows != nil {
		if row < 0 || col < 0 || row >= s.legacyRowCount || col >= s.legacyColCount {
			return fallback
		}
		if row < len(s.legacyRows) && col < len(s.legacyRows[row]) {
			return s.legacyRows[row][col]
		}
		return EmptyVal()
	}
	return EvalValueToValue(s.evalCell(row, col, fallback))
}

func (s criteriaValueSource) evalCell(row, col int, fallback Value) EvalValue {
	if s.grid == nil {
		if s.scalar != nil && row == 0 && col == 0 {
			return evalScalar(*s.scalar)
		}
		return evalScalar(fallback)
	}
	if row >= s.grid.Rows() || col >= s.grid.Cols() {
		return evalScalar(fallback)
	}
	return s.grid.Cell(row, col)
}

func ptrEvalValue(v EvalValue) *EvalValue {
	cp := v
	return &cp
}
