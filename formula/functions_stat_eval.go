package formula

import "strconv"

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
		return ValueToEvalValue(*err), nil
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
		return ValueToEvalValue(*err), nil
	}
	if count == 0 {
		return ValueToEvalValue(ErrorVal(ErrValDIV0)), nil
	}
	return EvalValue{Kind: EvalScalar, Scalar: NumberVal(sum / float64(count))}, nil
}

func evalCOUNTDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	count := 0
	if err := iterateReducerEvalArgs(
		args,
		func(v Value) *Value {
			switch v.Type {
			case ValueNumber:
				count++
			case ValueBool:
				if !v.FromCell {
					count++
				}
			case ValueString:
				if _, err := strconv.ParseFloat(v.Str, 64); err == nil {
					count++
				}
			}
			return nil
		},
		func(v Value) *Value {
			if v.Type == ValueNumber {
				count++
			}
			return nil
		},
	); err != nil {
		return ValueToEvalValue(*err), nil
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
		return ValueToEvalValue(*err), nil
	}
	return EvalValue{Kind: EvalScalar, Scalar: NumberVal(float64(count))}, nil
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
		return ValueToEvalValue(*err), nil
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
		return ValueToEvalValue(*err), nil
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
		return ValueToEvalValue(*err), nil
	}
	return EvalValue{Kind: EvalScalar, Scalar: NumberVal(sum)}, nil
}

func evalDEVSQDirectRange(args []EvalValue, _ *EvalContext) (EvalValue, error) {
	if len(args) == 0 {
		return ValueToEvalValue(ErrorVal(ErrValVALUE)), nil
	}
	nums, errVal := collectReducerNumericEvalArgs(args, true)
	if errVal != nil {
		return ValueToEvalValue(*errVal), nil
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
