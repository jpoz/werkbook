package formula

import "testing"

func TestUnaryInfoFuncSpecs(t *testing.T) {
	names := []string{
		"ISBLANK",
		"ISERR",
		"ISERROR",
		"ISNA",
		"ISNUMBER",
		"ISTEXT",
		"N",
	}

	for _, name := range names {
		t.Run(name, func(t *testing.T) {
			spec, ok := funcSpecForName(name)
			if !ok {
				t.Fatalf("FuncSpec missing for %s", name)
			}
			if spec.Kind != FnKindScalarLifted {
				t.Fatalf("Kind = %v, want %v", spec.Kind, FnKindScalarLifted)
			}
			if spec.Return != ReturnModePassThrough {
				t.Fatalf("Return = %v, want %v", spec.Return, ReturnModePassThrough)
			}
			if len(spec.Args) != 1 {
				t.Fatalf("len(Args) = %d, want 1", len(spec.Args))
			}
			if spec.Args[0].Load != ArgLoadPassthrough {
				t.Fatalf("Args[0].Load = %v, want %v", spec.Args[0].Load, ArgLoadPassthrough)
			}
			if spec.Args[0].Adapt != ArgAdaptLegacyIntersectRef {
				t.Fatalf("Args[0].Adapt = %v, want %v", spec.Args[0].Adapt, ArgAdaptLegacyIntersectRef)
			}
			if _, ok := funcMetaForName(name); ok {
				t.Fatalf("legacy metadata still registered for %s", name)
			}
			if inheritedArrayEvalForFuncArg(name, 0) {
				t.Fatalf("inheritedArrayEvalForFuncArg(%q, 0) = true, want false", name)
			}
			if !functionUsesElementwiseContract(name) {
				t.Fatalf("functionUsesElementwiseContract(%q) = false, want true", name)
			}
			if functionNeedsLegacyElementwisePreIntersect(name) {
				t.Fatalf("functionNeedsLegacyElementwisePreIntersect(%q) = true, want false", name)
			}
			if !functionCanReturnArrayFromArrayArgs(name) {
				t.Fatalf("functionCanReturnArrayFromArrayArgs(%q) = false, want true", name)
			}
		})
	}
}

func TestDirectRangeReducerFuncSpecs(t *testing.T) {
	names := []string{
		"SUM",
		"AVERAGE",
		"COUNT",
		"COUNTA",
		"MAX",
		"MIN",
		"SUMSQ",
		"DEVSQ",
	}

	for _, name := range names {
		t.Run(name, func(t *testing.T) {
			spec, ok := funcSpecForName(name)
			if !ok {
				t.Fatalf("FuncSpec missing for %s", name)
			}
			if spec.Kind != FnKindReduction {
				t.Fatalf("Kind = %v, want %v", spec.Kind, FnKindReduction)
			}
			if spec.Return != ReturnModeScalar {
				t.Fatalf("Return = %v, want %v", spec.Return, ReturnModeScalar)
			}
			if spec.Eval == nil {
				t.Fatalf("Eval = nil for %s", name)
			}
			arg0, ok := funcArgSpec(spec, 0)
			if !ok {
				t.Fatalf("missing arg contract for %s", name)
			}
			if arg0.Load != ArgLoadDirectRange {
				t.Fatalf("Args[0].Load = %v, want %v", arg0.Load, ArgLoadDirectRange)
			}
			if arg0.Adapt != ArgAdaptPassThrough {
				t.Fatalf("Args[0].Adapt = %v, want %v", arg0.Adapt, ArgAdaptPassThrough)
			}
		})
	}
}

func TestUnaryInfoFuncSpecParity(t *testing.T) {
	scalarCtx := &EvalContext{
		CurrentCol:     5,
		CurrentRow:     2,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}

	rangeArg := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("x")},
			{NumberVal(7)},
			{EmptyVal()},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 1,
			FromRow: 1,
			ToCol:   1,
			ToRow:   3,
		},
	}
	anonArrayArg := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(7)},
			{StringVal("x")},
		},
	}

	tests := []struct {
		caseName string
		name     string
		args     []Value
		ctx      *EvalContext
		want     Value
	}{
		{caseName: "blank_scalar", name: "ISBLANK", args: []Value{EmptyVal()}, ctx: scalarCtx, want: BoolVal(true)},
		{caseName: "iserr_scalar", name: "ISERR", args: []Value{ErrorVal(ErrValDIV0)}, ctx: scalarCtx, want: BoolVal(true)},
		{caseName: "iserror_scalar", name: "ISERROR", args: []Value{ErrorVal(ErrValNA)}, ctx: scalarCtx, want: BoolVal(true)},
		{caseName: "isna_scalar", name: "ISNA", args: []Value{ErrorVal(ErrValNA)}, ctx: scalarCtx, want: BoolVal(true)},
		{caseName: "direct_ref_scalarization", name: "ISNUMBER", args: []Value{rangeArg}, ctx: scalarCtx, want: BoolVal(true)},
		{caseName: "anonymous_array_broadcast", name: "ISNUMBER", args: []Value{anonArrayArg}, ctx: scalarCtx, want: Value{Type: ValueArray, Array: [][]Value{
			{BoolVal(true)},
			{BoolVal(false)},
		}}},
		{caseName: "istext_scalar", name: "ISTEXT", args: []Value{StringVal("hello")}, ctx: scalarCtx, want: BoolVal(true)},
		{caseName: "n_scalar", name: "N", args: []Value{BoolVal(true)}, ctx: scalarCtx, want: NumberVal(1)},
	}

	for _, tt := range tests {
		t.Run(tt.name+"_"+tt.caseName, func(t *testing.T) {
			fn := registry[normalizeFuncName(tt.name)]
			if fn == nil {
				t.Fatalf("function %s not registered", tt.name)
			}
			spec, ok := funcSpecForName(tt.name)
			if !ok {
				t.Fatalf("FuncSpec missing for %s", tt.name)
			}

			legacy, err := callLegacyScalarLifted(fn, tt.args, tt.ctx)
			if err != nil {
				t.Fatalf("legacy call: %v", err)
			}
			got, err := callFuncWithSpec(tt.name, fn, spec, tt.args, tt.ctx)
			if err != nil {
				t.Fatalf("contract call: %v", err)
			}

			assertLookupValueEqual(t, got, legacy)
			assertLookupValueEqual(t, got, tt.want)
		})
	}
}

func TestDirectRangeReducerFuncSpecParity(t *testing.T) {
	fullColumnRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(10)},
			{BoolVal(true)},
			{NumberVal(30)},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 1,
			FromRow: 1,
			ToCol:   1,
			ToRow:   maxRows,
		},
	}
	fullRowRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(6), EmptyVal(), NumberVal(12)},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 1,
			FromRow: 5,
			ToCol:   maxCols,
			ToRow:   5,
		},
	}
	boundedRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(2)},
			{NumberVal(4)},
			{EmptyVal()},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 2,
			FromRow: 7,
			ToCol:   2,
			ToRow:   9,
		},
	}
	anonArray := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(5), StringVal("x")},
		},
	}
	countaRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(7)},
			{EmptyVal()},
			{ErrorVal(ErrValDIV0)},
			{StringVal("x")},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 4,
			FromRow: 3,
			ToCol:   4,
			ToRow:   6,
		},
	}
	errorRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(9)},
			{ErrorVal(ErrValNA)},
			{NumberVal(1)},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 5,
			FromRow: 1,
			ToCol:   5,
			ToRow:   3,
		},
	}
	directCellText := Value{Type: ValueString, Str: "x", FromCell: true}

	tests := []struct {
		caseName string
		name     string
		args     []Value
		want     Value
	}{
		{
			caseName: "bounded_range",
			name:     "SUM",
			args:     []Value{boundedRange},
			want:     NumberVal(6),
		},
		{
			caseName: "full_column_trimmed_ref",
			name:     "SUM",
			args:     []Value{fullColumnRange},
			want:     NumberVal(40),
		},
		{
			caseName: "full_row_trimmed_ref",
			name:     "AVERAGE",
			args:     []Value{fullRowRange},
			want:     NumberVal(9),
		},
		{
			caseName: "scalar_and_array_inputs",
			name:     "AVERAGE",
			args:     []Value{NumberVal(10), anonArray},
			want:     NumberVal(7.5),
		},
		{
			caseName: "count_scalar_and_range_inputs",
			name:     "COUNT",
			args:     []Value{BoolVal(true), fullColumnRange, anonArray},
			want:     NumberVal(4),
		},
		{
			caseName: "counta_bounded_range_counts_errors",
			name:     "COUNTA",
			args:     []Value{countaRange},
			want:     NumberVal(3),
		},
		{
			caseName: "max_full_column_trimmed_ref",
			name:     "MAX",
			args:     []Value{fullColumnRange},
			want:     NumberVal(30),
		},
		{
			caseName: "max_direct_cell_text_errors",
			name:     "MAX",
			args:     []Value{directCellText, NumberVal(4)},
			want:     ErrorVal(ErrValVALUE),
		},
		{
			caseName: "min_full_row_trimmed_ref",
			name:     "MIN",
			args:     []Value{fullRowRange},
			want:     NumberVal(6),
		},
		{
			caseName: "min_error_propagation",
			name:     "MIN",
			args:     []Value{errorRange},
			want:     ErrorVal(ErrValNA),
		},
		{
			caseName: "sumsq_scalar_and_array_inputs",
			name:     "SUMSQ",
			args:     []Value{NumberVal(3), anonArray},
			want:     NumberVal(34),
		},
		{
			caseName: "sumsq_error_propagation",
			name:     "SUMSQ",
			args:     []Value{ErrorVal(ErrValDIV0), NumberVal(2)},
			want:     ErrorVal(ErrValDIV0),
		},
		{
			caseName: "devsq_scalar_and_array_inputs",
			name:     "DEVSQ",
			args:     []Value{NumberVal(1), anonArray},
			want:     NumberVal(8),
		},
		{
			caseName: "devsq_direct_cell_text_ignored",
			name:     "DEVSQ",
			args:     []Value{directCellText, NumberVal(4)},
			want:     NumberVal(0),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name+"_"+tt.caseName, func(t *testing.T) {
			fn := registry[normalizeFuncName(tt.name)]
			if fn == nil {
				t.Fatalf("function %s not registered", tt.name)
			}
			spec, ok := funcSpecForName(tt.name)
			if !ok {
				t.Fatalf("FuncSpec missing for %s", tt.name)
			}

			legacy, err := fn(tt.args, nil)
			if err != nil {
				t.Fatalf("legacy call: %v", err)
			}
			got, err := callFuncWithSpec(tt.name, fn, spec, tt.args, nil)
			if err != nil {
				t.Fatalf("contract call: %v", err)
			}

			assertLookupValueEqual(t, got, legacy)
			assertLookupValueEqual(t, got, tt.want)

			for i, arg := range tt.args {
				if arg.Type == ValueArray && arg.RangeOrigin != nil {
					loaded := loadEvalFuncArg(ArgSpec{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough}, arg, nil)
					if loaded.Kind != EvalRef || loaded.Ref == nil {
						t.Fatalf("loaded arg %d = %#v, want EvalRef", i, loaded)
					}
				}
			}
		})
	}
}

func callLegacyScalarLifted(fn Func, args []Value, ctx *EvalContext) (Value, error) {
	adapted := make([]Value, len(args))
	copy(adapted, args)
	if ctx != nil && !ctx.IsArrayFormula {
		for i := range adapted {
			if adapted[i].Type == ValueArray && adapted[i].RangeOrigin != nil {
				adapted[i] = implicitIntersect(adapted[i], ctx)
			}
		}
	}
	if hasArrayArg(adapted) {
		return callElementWise(adapted, ctx, fn)
	}
	return fn(adapted, ctx)
}
