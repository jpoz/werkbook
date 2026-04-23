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
