package formula

import (
	"math"
	"sort"
	"strconv"
)

func directRangeReducerFuncSpec(eval EvalFunc) FuncSpec {
	return FuncSpec{
		Kind: FnKindReduction,
		VarArg: func(i int) ArgSpec {
			return ArgSpec{
				Load:  ArgLoadDirectRange,
				Adapt: ArgAdaptPassThrough,
			}
		},
		Return: ReturnModeScalar,
		Eval:   eval,
	}
}

func evalSUMDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	sum := 0.0
	if err := iterateReducerEvalArgs(
		args,
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			n, e := CoerceNum(v)
			if e != nil {
				return e
			}
			sum += n
			return nil
		},
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			if v.Type == ValueNumber {
				sum += v.Num
			}
			return nil
		},
	); err != nil {
		return evalScalar(*err), nil
	}
	return EvalValue{Kind: EvalScalar, Scalar: NumberVal(sum)}, nil
}

func evalAVERAGEDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	sum := 0.0
	count := 0
	if err := iterateReducerEvalArgs(
		args,
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			n, e := CoerceNum(v)
			if e != nil {
				return e
			}
			sum += n
			count++
			return nil
		},
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			if v.Type == ValueNumber {
				sum += v.Num
				count++
			}
			return nil
		},
	); err != nil {
		return evalScalar(*err), nil
	}
	if count == 0 {
		return evalError(ErrValDIV0), nil
	}
	return EvalValue{Kind: EvalScalar, Scalar: NumberVal(sum / float64(count))}, nil
}

func evalCOUNTDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	count := 0
	for _, arg := range args {
		switch arg.Kind {
		case EvalKindError:
			continue
		case EvalScalar:
			switch arg.Scalar.Type {
			case ValueNumber:
				count++
			case ValueBool:
				if !arg.Scalar.FromCell {
					count++
				}
			case ValueString:
				if _, err := strconv.ParseFloat(arg.Scalar.Str, 64); err == nil {
					count++
				}
			}
		case EvalArray:
			if arg.Array == nil || arg.Array.Grid == nil {
				continue
			}
			ignoreRefDerived := arg.Array.Origin == nil
			arg.Array.Grid.Iterate(func(r, c int, cell EvalValue) bool {
				if countEvalNumberCell(cell, ignoreRefDerived) {
					count++
				}
				return true
			})
		case EvalRef:
			if err := iterateReducerGrid(reducerRefGrid(arg.Ref), func(v Value) *Value {
				if v.Type == ValueNumber {
					count++
				}
				return nil
			}); err != nil {
				return evalScalar(*err), nil
			}
		}
	}
	return EvalValue{Kind: EvalScalar, Scalar: NumberVal(float64(count))}, nil
}

func evalCOUNTADirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	count := 0
	if err := iterateReducerEvalArgs(
		args,
		func(v Value) *Value {
			if v.Type != ValueEmpty {
				count++
			}
			return nil
		},
		func(v Value) *Value {
			if v.Type != ValueEmpty {
				count++
			}
			return nil
		},
	); err != nil {
		return evalScalar(*err), nil
	}
	return EvalValue{Kind: EvalScalar, Scalar: NumberVal(float64(count))}, nil
}

func countEvalNumberCell(cell EvalValue, ignoreRefDerived bool) bool {
	if cell.Kind == EvalScalar {
		return cell.Scalar.Type == ValueNumber
	}
	if ignoreRefDerived && cell.Kind == EvalRef {
		return false
	}
	v := EvalValueToValue(cell)
	return v.Type == ValueNumber && !(ignoreRefDerived && valueHasRefMetadata(v))
}

func evalMINDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	min := 0.0
	found := false
	if err := iterateReducerEvalArgs(
		args,
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			n, e := CoerceNum(v)
			if e != nil {
				return e
			}
			if !found || n < min {
				min = n
				found = true
			}
			return nil
		},
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			if v.Type == ValueNumber {
				if !found || v.Num < min {
					min = v.Num
					found = true
				}
			}
			return nil
		},
	); err != nil {
		return evalScalar(*err), nil
	}
	if !found {
		return EvalValue{Kind: EvalScalar, Scalar: NumberVal(0)}, nil
	}
	return EvalValue{Kind: EvalScalar, Scalar: NumberVal(min)}, nil
}

func evalMAXDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	max := 0.0
	found := false
	if err := iterateReducerEvalArgs(
		args,
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			n, e := CoerceNum(v)
			if e != nil {
				return e
			}
			if !found || n > max {
				max = n
				found = true
			}
			return nil
		},
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			if v.Type == ValueNumber {
				if !found || v.Num > max {
					max = v.Num
					found = true
				}
			}
			return nil
		},
	); err != nil {
		return evalScalar(*err), nil
	}
	if !found {
		return EvalValue{Kind: EvalScalar, Scalar: NumberVal(0)}, nil
	}
	return EvalValue{Kind: EvalScalar, Scalar: NumberVal(max)}, nil
}

func evalSUMSQDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	sum := 0.0
	if err := iterateReducerEvalArgs(
		args,
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			n, e := CoerceNum(v)
			if e != nil {
				return e
			}
			sum += n * n
			return nil
		},
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			if v.Type == ValueNumber {
				sum += v.Num * v.Num
			}
			return nil
		},
	); err != nil {
		return evalScalar(*err), nil
	}
	return EvalValue{Kind: EvalScalar, Scalar: NumberVal(sum)}, nil
}

func evalDEVSQDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) == 0 {
		return evalError(ErrValVALUE), nil
	}
	nums, errVal := collectReducerNumericEvalArgs(args, true)
	if errVal != nil {
		return evalScalar(*errVal), nil
	}
	if len(nums) == 0 {
		return EvalValue{Kind: EvalScalar, Scalar: NumberVal(0)}, nil
	}
	sum := 0.0
	for _, n := range nums {
		sum += n
	}
	mean := sum / float64(len(nums))
	ssq := 0.0
	for _, n := range nums {
		d := n - mean
		ssq += d * d
	}
	return EvalValue{Kind: EvalScalar, Scalar: NumberVal(ssq)}, nil
}

func evalAVEDEVDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) == 0 {
		return evalError(ErrValVALUE), nil
	}
	nums, mean, _, errVal := meanAndSumSqReducerEvalArgs(args, true)
	if errVal != nil {
		return evalScalar(*errVal), nil
	}
	if len(nums) == 0 {
		return evalError(ErrValNUM), nil
	}
	absDevSum := 0.0
	for _, n := range nums {
		absDevSum += math.Abs(n - mean)
	}
	return EvalValue{Kind: EvalScalar, Scalar: NumberVal(absDevSum / float64(len(nums)))}, nil
}

func evalSTDEVDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	return evalReducerVarianceLikeDirectRange(args, true, true, meanAndSumSqReducerDirectArgs), nil
}

func evalSTDEVPDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	return evalReducerVarianceLikeDirectRange(args, false, true, meanAndSumSqReducerDirectArgs), nil
}

func evalSTDEVADirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	return evalReducerVarianceLikeDirectRange(args, true, true, meanAndSumSqReducerAEvalArgs), nil
}

func evalSTDEVPADirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	return evalReducerVarianceLikeDirectRange(args, false, true, meanAndSumSqReducerAEvalArgs), nil
}

func evalVARDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	return evalReducerVarianceLikeDirectRange(args, true, false, meanAndSumSqReducerDirectArgs), nil
}

func evalVARPDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	return evalReducerVarianceLikeDirectRange(args, false, false, meanAndSumSqReducerDirectArgs), nil
}

func evalVARADirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	return evalReducerVarianceLikeDirectRange(args, true, false, meanAndSumSqReducerAEvalArgs), nil
}

func evalVARPADirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	return evalReducerVarianceLikeDirectRange(args, false, false, meanAndSumSqReducerAEvalArgs), nil
}

func evalReducerVarianceLikeDirectRange(
	args []EvalValue,
	sample bool,
	sqrtResult bool,
	meanAndSumSq func([]EvalValue) ([]float64, float64, float64, *Value),
) EvalValue {
	if len(args) == 0 {
		return evalError(ErrValVALUE)
	}
	nums, _, ssq, errVal := meanAndSumSq(args)
	if errVal != nil {
		return evalScalar(*errVal)
	}
	minCount := 1
	denom := float64(len(nums))
	if sample {
		minCount = 2
		denom = float64(len(nums) - 1)
	}
	if len(nums) < minCount {
		return evalError(ErrValDIV0)
	}
	result := ssq / denom
	if sqrtResult {
		result = math.Sqrt(result)
	}
	return EvalValue{Kind: EvalScalar, Scalar: NumberVal(result)}
}

func meanAndSumSqReducerDirectArgs(args []EvalValue) ([]float64, float64, float64, *Value) {
	return meanAndSumSqReducerEvalArgs(args, true)
}

func meanAndSumSqReducerEvalArgs(args []EvalValue, ignoreNonNumericDirectCell bool) ([]float64, float64, float64, *Value) {
	nums, errVal := collectReducerNumericEvalArgs(args, ignoreNonNumericDirectCell)
	if errVal != nil {
		return nil, 0, 0, errVal
	}
	mean, ssq := meanAndSumSqReducerNumbers(nums)
	return nums, mean, ssq, nil
}

func meanAndSumSqReducerAEvalArgs(args []EvalValue) ([]float64, float64, float64, *Value) {
	nums, errVal := collectReducerNumericAEvalArgs(args)
	if errVal != nil {
		return nil, 0, 0, errVal
	}
	mean, ssq := meanAndSumSqReducerNumbers(nums)
	return nums, mean, ssq, nil
}

func meanAndSumSqReducerNumbers(nums []float64) (float64, float64) {
	if len(nums) == 0 {
		return 0, 0
	}
	sum := 0.0
	for _, n := range nums {
		sum += n
	}
	mean := sum / float64(len(nums))
	ssq := 0.0
	for _, n := range nums {
		d := n - mean
		ssq += d * d
	}
	return mean, ssq
}

func collectReducerNumericEvalArgs(args []EvalValue, ignoreNonNumericDirectCell bool) ([]float64, *Value) {
	var nums []float64
	err := iterateReducerEvalArgs(
		args,
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			if ignoreNonNumericDirectCell && v.FromCell &&
				(v.Type == ValueString || v.Type == ValueBool || v.Type == ValueEmpty) {
				return nil
			}
			n, e := CoerceNum(v)
			if e != nil {
				return e
			}
			nums = append(nums, n)
			return nil
		},
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			if v.Type == ValueNumber {
				nums = append(nums, v.Num)
			}
			return nil
		},
	)
	if err != nil {
		return nil, err
	}
	return nums, nil
}

func collectReducerNumericAEvalArgs(args []EvalValue) ([]float64, *Value) {
	var nums []float64
	err := iterateReducerEvalArgs(
		args,
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			if v.Type == ValueEmpty {
				return nil
			}
			n, e := CoerceNum(v)
			if e != nil {
				return e
			}
			nums = append(nums, n)
			return nil
		},
		func(v Value) *Value {
			if v.Type == ValueError {
				return &v
			}
			switch v.Type {
			case ValueNumber:
				nums = append(nums, v.Num)
			case ValueBool:
				if v.Bool {
					nums = append(nums, 1)
				} else {
					nums = append(nums, 0)
				}
			case ValueString:
				nums = append(nums, 0)
			}
			return nil
		},
	)
	if err != nil {
		return nil, err
	}
	return nums, nil
}

func iterateReducerEvalArgs(
	args []EvalValue,
	visitScalar func(Value) *Value,
	visitCollection func(Value) *Value,
) *Value {
	for _, arg := range args {
		switch arg.Kind {
		case EvalKindError:
			if visitScalar == nil {
				continue
			}
			if err := visitScalar(ErrorVal(arg.Err)); err != nil {
				return err
			}
		case EvalScalar:
			if visitScalar == nil {
				continue
			}
			if err := visitScalar(arg.Scalar); err != nil {
				return err
			}
		case EvalArray:
			if arg.Array == nil {
				continue
			}
			if err := iterateReducerGrid(arg.Array.Grid, visitCollection); err != nil {
				return err
			}
		case EvalRef:
			if err := iterateReducerGrid(reducerRefGrid(arg.Ref), visitCollection); err != nil {
				return err
			}
		}
	}
	return nil
}

func iterateReducerGrid(grid Grid, visit func(Value) *Value) *Value {
	if grid == nil || visit == nil {
		return nil
	}
	if legacy, ok := grid.(interface {
		legacyRows() ([][]Value, int, int)
	}); ok {
		rows, rowCount, colCount := legacy.legacyRows()
		for r := 0; r < rowCount; r++ {
			for c := 0; c < colCount; c++ {
				cell := EmptyVal()
				if r < len(rows) && c < len(rows[r]) {
					cell = rows[r][c]
				}
				if err := visit(cell); err != nil {
					return err
				}
			}
		}
		return nil
	}
	var err *Value
	grid.Iterate(func(r, c int, cell EvalValue) bool {
		err = visit(EvalValueToValue(cell))
		return err == nil
	})
	return err
}

func reducerRefGrid(ref *RefValue) Grid {
	if ref == nil {
		return nil
	}
	if ref.Materialized != nil {
		return ref.Materialized
	}
	return emptyRefGrid{
		rows: ref.ToRow - ref.FromRow + 1,
		cols: ref.ToCol - ref.FromCol + 1,
	}
}

func sumproductFuncSpec(eval EvalFunc) FuncSpec {
	return FuncSpec{
		Kind: FnKindReduction,
		VarArg: func(i int) ArgSpec {
			return ArgSpec{Load: ArgLoadArray, Adapt: ArgAdaptPassThrough}
		},
		Return: ReturnModeScalar,
		Eval:   eval,
	}
}

func evalSUMPRODUCTDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) == 0 {
		return evalError(ErrValVALUE), nil
	}
	sources := make([]criteriaValueSource, len(args))
	rows, cols := 0, 0
	for i, arg := range args {
		if arg.Kind == EvalKindError {
			return evalError(arg.Err), nil
		}
		sources[i] = newCriteriaValueSource(arg)
		argRows, argCols := sources[i].dims()
		if i == 0 {
			rows, cols = argRows, argCols
			continue
		}
		if argRows != rows || argCols != cols {
			return evalError(ErrValVALUE), nil
		}
	}

	sum := 0.0
	for r := 0; r < rows; r++ {
		for c := 0; c < cols; c++ {
			product := 1.0
			for i, source := range sources {
				cellEval := source.evalCell(r, c, EmptyVal())
				cell := EvalValueToValue(cellEval)
				if cell.Type == ValueError {
					return evalScalar(cell), nil
				}
				if anonymousRefDerivedEvalArrayCell(args[i], cellEval) {
					product = 0
					continue
				}
				if cell.Type == ValueString || cell.Type == ValueBool {
					product = 0
					continue
				}
				n, errVal := CoerceNum(cell)
				if errVal != nil {
					n = 0
				}
				product *= n
			}
			sum += product
		}
	}
	return evalScalar(NumberVal(sum)), nil
}

func anonymousRefDerivedEvalArrayCell(container EvalValue, cell EvalValue) bool {
	if container.Kind != EvalArray || container.Array == nil || container.Array.Origin != nil {
		return false
	}
	if cell.Kind == EvalRef {
		return true
	}
	return valueHasRefMetadata(EvalValueToValue(cell))
}

func criteriaExtremaFuncSpec(eval EvalFunc) FuncSpec {
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

func evalMAXIFSCriteria(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	return evalCriteriaExtremaEval(criteriaExtremaMax, args), nil
}

func evalMINIFSCriteria(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	return evalCriteriaExtremaEval(criteriaExtremaMin, args), nil
}

func evalCriteriaExtremaEval(kind criteriaExtremaKind, args []EvalValue) EvalValue {
	if len(args) < 3 || (len(args)-1)%2 != 0 {
		return evalError(ErrValVALUE)
	}
	if args[0].Kind != EvalArray && args[0].Kind != EvalRef {
		return evalError(ErrValVALUE)
	}
	scanSource := newCriteriaValueSource(args[0])
	prepared := make([]criteriaPreparedPair, 0, (len(args)-1)/2)
	broadcastRows := 0
	broadcastCols := 0
	hasBroadcast := false

	for k := 1; k < len(args); k += 2 {
		pair := criteriaPreparedPair{
			rangeSource:    newCriteriaValueSource(args[k]),
			criteriaSource: newCriteriaCriterionSource(args[k+1]),
		}
		prepared = append(prepared, pair)
		if hasBroadcast || pair.criteriaSource.isScalar() {
			continue
		}
		broadcastRows, broadcastCols = pair.criteriaSource.dims()
		hasBroadcast = true
	}

	if !hasBroadcast {
		return evalScalar(evalCriteriaExtremaScalar(kind, scanSource, prepared, 0, 0))
	}

	rows := make([][]Value, broadcastRows)
	for r := 0; r < broadcastRows; r++ {
		row := make([]Value, broadcastCols)
		for c := 0; c < broadcastCols; c++ {
			row[c] = evalCriteriaExtremaScalar(kind, scanSource, prepared, r, c)
		}
		rows[r] = row
	}
	return evalArray(rows, SpillBounded)
}

func newCriteriaCriterionSource(v EvalValue) criteriaValueSource {
	legacy := EvalValueToValue(v)
	if legacy.Type == ValueArray {
		return newCriteriaValueSource(v)
	}
	scalar := legacy
	return criteriaValueSource{scalar: &scalar}
}

func frequencyFuncSpec(eval EvalFunc) FuncSpec {
	return FuncSpec{
		Kind: FnKindArrayNative,
		Args: []ArgSpec{
			{Load: ArgLoadArray, Adapt: ArgAdaptPassThrough},
			{Load: ArgLoadArray, Adapt: ArgAdaptPassThrough},
		},
		Return: ReturnModeArray,
		Eval:   eval,
	}
}

func evalFREQUENCYDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) != 2 {
		return evalError(ErrValVALUE), nil
	}
	data, errVal := collectFrequencyNumbers(args[0])
	if errVal != nil {
		return evalScalar(*errVal), nil
	}
	bins, errVal := collectFrequencyNumbers(args[1])
	if errVal != nil {
		return evalScalar(*errVal), nil
	}
	return evalArray(frequencyRows(data, bins), SpillBounded), nil
}

func collectFrequencyNumbers(v EvalValue) ([]float64, *Value) {
	var nums []float64
	visit := func(cell EvalValue) *Value {
		legacy := EvalValueToValue(cell)
		if legacy.Type == ValueError {
			return &legacy
		}
		if legacy.Type == ValueNumber {
			nums = append(nums, legacy.Num)
		}
		return nil
	}
	switch v.Kind {
	case EvalKindError:
		errVal := ErrorVal(v.Err)
		return nil, &errVal
	case EvalScalar:
		if errVal := visit(v); errVal != nil {
			return nil, errVal
		}
	case EvalArray:
		if v.Array == nil || v.Array.Grid == nil {
			return nums, nil
		}
		var errVal *Value
		v.Array.Grid.Iterate(func(r, c int, cell EvalValue) bool {
			errVal = visit(cell)
			return errVal == nil
		})
		if errVal != nil {
			return nil, errVal
		}
	case EvalRef:
		var errVal *Value
		if grid := reducerRefGrid(v.Ref); grid != nil {
			grid.Iterate(func(r, c int, cell EvalValue) bool {
				errVal = visit(cell)
				return errVal == nil
			})
		}
		if errVal != nil {
			return nil, errVal
		}
	}
	return nums, nil
}

func frequencyRows(data []float64, bins []float64) [][]Value {
	if len(bins) == 0 {
		return [][]Value{{NumberVal(float64(len(data)))}}
	}
	sort.Float64s(bins)
	counts := make([]float64, len(bins)+1)
	for _, v := range data {
		placed := false
		for i, b := range bins {
			if v <= b {
				counts[i]++
				placed = true
				break
			}
		}
		if !placed {
			counts[len(bins)]++
		}
	}
	rows := make([][]Value, len(counts))
	for i, c := range counts {
		rows[i] = []Value{NumberVal(c)}
	}
	return rows
}
