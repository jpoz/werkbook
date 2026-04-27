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
		"AVEDEV",
		"SUM",
		"AVERAGE",
		"COUNT",
		"COUNTA",
		"MAX",
		"MIN",
		"STDEV",
		"STDEV.S",
		"STDEVA",
		"STDEVP",
		"STDEV.P",
		"STDEVPA",
		"SUMSQ",
		"DEVSQ",
		"VAR",
		"VAR.S",
		"VARA",
		"VARP",
		"VAR.P",
		"VARPA",
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

func TestCriteriaFuncSpecs(t *testing.T) {
	tests := []struct {
		name string
		want []ArgSpec
	}{
		{
			name: "COUNTIF",
			want: []ArgSpec{
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			},
		},
		{
			name: "SUMIF",
			want: []ArgSpec{
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
			},
		},
		{
			name: "AVERAGEIF",
			want: []ArgSpec{
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
			},
		},
		{
			name: "COUNTIFS",
			want: []ArgSpec{
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			},
		},
		{
			name: "SUMIFS",
			want: []ArgSpec{
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			},
		},
		{
			name: "AVERAGEIFS",
			want: []ArgSpec{
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			spec, ok := funcSpecForName(tt.name)
			if !ok {
				t.Fatalf("FuncSpec missing for %s", tt.name)
			}
			if spec.Kind != FnKindReduction {
				t.Fatalf("Kind = %v, want %v", spec.Kind, FnKindReduction)
			}
			if spec.Return != ReturnModePassThrough {
				t.Fatalf("Return = %v, want %v", spec.Return, ReturnModePassThrough)
			}
			if spec.Eval == nil {
				t.Fatalf("Eval = nil for %s", tt.name)
			}
			if !functionCanReturnArrayFromArrayArgs(tt.name) {
				t.Fatalf("functionCanReturnArrayFromArrayArgs(%q) = false, want true", tt.name)
			}
			for i, want := range tt.want {
				got, ok := funcArgSpec(spec, i)
				if !ok {
					t.Fatalf("missing arg contract for %s arg %d", tt.name, i)
				}
				if got != want {
					t.Fatalf("arg %d = %+v, want %+v", i, got, want)
				}
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
	statsFullColumnRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(2)},
			{BoolVal(true)},
			{NumberVal(4)},
			{EmptyVal()},
			{NumberVal(6)},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 6,
			FromRow: 1,
			ToCol:   6,
			ToRow:   maxRows,
		},
	}
	statsFullRowRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(2), BoolVal(true), NumberVal(4)},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 1,
			FromRow: 8,
			ToCol:   maxCols,
			ToRow:   8,
		},
	}
	aStatsFullColumnRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(2)},
			{EmptyVal()},
			{BoolVal(true)},
			{EmptyVal()},
			{StringVal("x")},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 8,
			FromRow: 1,
			ToCol:   8,
			ToRow:   maxRows,
		},
	}
	aStatsFullRowRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(2), EmptyVal(), BoolVal(true), EmptyVal(), StringVal("x")},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 1,
			FromRow: 9,
			ToCol:   maxCols,
			ToRow:   9,
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
	statsAnonArray := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(4), StringVal("x"), NumberVal(6)},
		},
	}
	aStatsAnonArray := Value{
		Type: ValueArray,
		Array: [][]Value{
			{BoolVal(true), StringVal("x")},
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
	textOnlyRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("x")},
			{EmptyVal()},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 7,
			FromRow: 2,
			ToCol:   7,
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
		{
			caseName: "avedev_bounded_range",
			name:     "AVEDEV",
			args:     []Value{boundedRange},
			want:     NumberVal(1),
		},
		{
			caseName: "avedev_full_column_trimmed_ref",
			name:     "AVEDEV",
			args:     []Value{statsFullColumnRange},
			want:     NumberVal(4.0 / 3.0),
		},
		{
			caseName: "avedev_empty_range_returns_num",
			name:     "AVEDEV",
			args:     []Value{textOnlyRange},
			want:     ErrorVal(ErrValNUM),
		},
		{
			caseName: "stdev_full_column_trimmed_ref",
			name:     "STDEV",
			args:     []Value{statsFullColumnRange},
			want:     NumberVal(2),
		},
		{
			caseName: "stdev_low_cardinality_div0",
			name:     "STDEV.S",
			args:     []Value{NumberVal(5)},
			want:     ErrorVal(ErrValDIV0),
		},
		{
			caseName: "stdev_error_propagation",
			name:     "STDEV",
			args:     []Value{errorRange},
			want:     ErrorVal(ErrValNA),
		},
		{
			caseName: "stdevp_full_row_trimmed_ref",
			name:     "STDEV.P",
			args:     []Value{statsFullRowRange},
			want:     NumberVal(1),
		},
		{
			caseName: "var_scalar_and_array_inputs",
			name:     "VAR.S",
			args:     []Value{NumberVal(2), statsAnonArray},
			want:     NumberVal(4),
		},
		{
			caseName: "var_error_propagation",
			name:     "VAR",
			args:     []Value{errorRange},
			want:     ErrorVal(ErrValNA),
		},
		{
			caseName: "varp_single_value_returns_zero",
			name:     "VARP",
			args:     []Value{NumberVal(5)},
			want:     NumberVal(0),
		},
		{
			caseName: "varp_full_row_trimmed_ref",
			name:     "VAR.P",
			args:     []Value{statsFullRowRange},
			want:     NumberVal(1),
		},
		{
			caseName: "vara_full_column_trimmed_ref",
			name:     "VARA",
			args:     []Value{aStatsFullColumnRange},
			want:     NumberVal(1),
		},
		{
			caseName: "vara_direct_cell_text_errors",
			name:     "VARA",
			args:     []Value{directCellText, NumberVal(2)},
			want:     ErrorVal(ErrValVALUE),
		},
		{
			caseName: "varpa_full_row_trimmed_ref",
			name:     "VARPA",
			args:     []Value{aStatsFullRowRange},
			want:     NumberVal(2.0 / 3.0),
		},
		{
			caseName: "stdeva_scalar_and_array_inputs",
			name:     "STDEVA",
			args:     []Value{NumberVal(2), aStatsAnonArray},
			want:     NumberVal(1),
		},
		{
			caseName: "stdeva_text_only_range_low_cardinality",
			name:     "STDEVA",
			args:     []Value{textOnlyRange},
			want:     ErrorVal(ErrValDIV0),
		},
		{
			caseName: "stdevpa_error_propagation",
			name:     "STDEVPA",
			args:     []Value{errorRange},
			want:     ErrorVal(ErrValNA),
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

func TestCriteriaFuncSpecParity(t *testing.T) {
	boundedTextRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("apple")},
			{StringVal("pear")},
			{StringVal("apple")},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 1,
			FromRow: 1,
			ToCol:   1,
			ToRow:   3,
		},
	}
	boundedValueRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(10)},
			{NumberVal(20)},
			{NumberVal(30)},
			{NumberVal(40)},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 2,
			FromRow: 1,
			ToCol:   2,
			ToRow:   4,
		},
	}
	fullColumnCriteriaRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("East")},
			{StringVal("West")},
			{StringVal("East")},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 3,
			FromRow: 1,
			ToCol:   3,
			ToRow:   maxRows,
		},
	}
	fullColumnSumRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(10)},
			{NumberVal(20)},
			{NumberVal(30)},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 4,
			FromRow: 1,
			ToCol:   4,
			ToRow:   maxRows,
		},
	}
	boundedRegionRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("East")},
			{StringVal("West")},
			{StringVal("East")},
			{StringVal("West")},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 5,
			FromRow: 1,
			ToCol:   5,
			ToRow:   4,
		},
	}
	boundedStatusRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("Open")},
			{StringVal("Open")},
			{StringVal("Closed")},
			{StringVal("Open")},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 6,
			FromRow: 1,
			ToCol:   6,
			ToRow:   4,
		},
	}
	fullRowAverageRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(10), NumberVal(20), ErrorVal(ErrValDIV0)},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 1,
			FromRow: 9,
			ToCol:   maxCols,
			ToRow:   9,
		},
	}
	fullRowRegionRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("East"), StringVal("East"), StringVal("East")},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 1,
			FromRow: 10,
			ToCol:   maxCols,
			ToRow:   10,
		},
	}
	fullRowStatusRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("Open"), StringVal("Open"), StringVal("Closed")},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 1,
			FromRow: 11,
			ToCol:   maxCols,
			ToRow:   11,
		},
	}

	tests := []struct {
		caseName string
		name     string
		args     []Value
		want     Value
	}{
		{
			caseName: "bounded_range_scalar_criteria",
			name:     "COUNTIF",
			args:     []Value{boundedTextRange, StringVal("apple")},
			want:     NumberVal(2),
		},
		{
			caseName: "array_criteria_broadcast",
			name:     "COUNTIF",
			args: []Value{
				boundedTextRange,
				{Type: ValueArray, Array: [][]Value{{StringVal("apple"), StringVal("pear")}}},
			},
			want: Value{Type: ValueArray, Array: [][]Value{{NumberVal(2), NumberVal(1)}}},
		},
		{
			caseName: "full_column_trimmed_ref",
			name:     "SUMIF",
			args:     []Value{fullColumnCriteriaRange, StringVal("East"), fullColumnSumRange},
			want:     NumberVal(40),
		},
		{
			caseName: "array_criteria_broadcast",
			name:     "SUMIF",
			args: []Value{
				fullColumnCriteriaRange,
				{Type: ValueArray, Array: [][]Value{{StringVal("East"), StringVal("West")}}},
				fullColumnSumRange,
			},
			want: Value{Type: ValueArray, Array: [][]Value{{NumberVal(40), NumberVal(20)}}},
		},
		{
			caseName: "bounded_range_with_sum_range",
			name:     "AVERAGEIF",
			args: []Value{
				boundedRegionRange,
				StringVal("East"),
				boundedValueRange,
			},
			want: NumberVal(20),
		},
		{
			caseName: "mixed_scalar_and_array_criteria",
			name:     "COUNTIFS",
			args: []Value{
				boundedRegionRange,
				{Type: ValueArray, Array: [][]Value{{StringVal("East"), StringVal("West")}}},
				boundedStatusRange,
				StringVal("Open"),
			},
			want: Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}}},
		},
		{
			caseName: "mixed_scalar_and_array_criteria",
			name:     "SUMIFS",
			args: []Value{
				boundedValueRange,
				boundedRegionRange,
				{Type: ValueArray, Array: [][]Value{{StringVal("East"), StringVal("West")}}},
				boundedStatusRange,
				StringVal("Open"),
			},
			want: Value{Type: ValueArray, Array: [][]Value{{NumberVal(10), NumberVal(60)}}},
		},
		{
			caseName: "full_row_trimmed_ref_matched_error",
			name:     "AVERAGEIFS",
			args: []Value{
				fullRowAverageRange,
				fullRowRegionRange,
				StringVal("East"),
				fullRowStatusRange,
				StringVal("Closed"),
			},
			want: ErrorVal(ErrValDIV0),
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
				argSpec, ok := funcArgSpec(spec, i)
				if !ok {
					continue
				}
				if argSpec.Load != ArgLoadDirectRange || arg.Type != ValueArray || arg.RangeOrigin == nil {
					continue
				}
				loaded := loadEvalFuncArg(argSpec, arg, nil)
				if loaded.Kind != EvalRef || loaded.Ref == nil {
					t.Fatalf("loaded arg %d = %#v, want EvalRef", i, loaded)
				}
			}
		})
	}
}

func TestLookupFamilyFuncSpecs(t *testing.T) {
	tests := []struct {
		name     string
		wantArgs []ArgSpec
	}{
		{
			name: "VLOOKUP",
			wantArgs: []ArgSpec{
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptScalarizeAny},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			},
		},
		{
			name: "HLOOKUP",
			wantArgs: []ArgSpec{
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptScalarizeAny},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			},
		},
		{
			name: "MATCH",
			wantArgs: []ArgSpec{
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptScalarizeAny},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			},
		},
		{
			name: "LOOKUP",
			wantArgs: []ArgSpec{
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptScalarizeAny},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			},
		},
		{
			name: "XLOOKUP",
			wantArgs: []ArgSpec{
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptScalarizeAny},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			},
		},
		{
			name: "XMATCH",
			wantArgs: []ArgSpec{
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptScalarizeAny},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			spec, ok := funcSpecForName(tt.name)
			if !ok {
				t.Fatalf("FuncSpec missing for %s", tt.name)
			}
			if spec.Kind != FnKindLookupArrayLift {
				t.Fatalf("Kind = %v, want %v", spec.Kind, FnKindLookupArrayLift)
			}
			if spec.Return != ReturnModePassThrough {
				t.Fatalf("Return = %v, want %v", spec.Return, ReturnModePassThrough)
			}
			if spec.Eval == nil {
				t.Fatalf("Eval missing for %s", tt.name)
			}
			for i, want := range tt.wantArgs {
				got, ok := funcArgSpec(spec, i)
				if !ok {
					t.Fatalf("missing arg contract for %s arg %d", tt.name, i)
				}
				if got != want {
					t.Fatalf("%s arg %d = %+v, want %+v", tt.name, i, got, want)
				}
			}
			// Trailing varargs should also pass through so optional
			// match_mode/search_mode/if_not_found scalars are not intersected.
			extra, ok := funcArgSpec(spec, len(tt.wantArgs))
			if !ok {
				t.Fatalf("missing VarArg contract for %s", tt.name)
			}
			if extra.Load != ArgLoadPassthrough || extra.Adapt != ArgAdaptPassThrough {
				t.Fatalf("VarArg for %s = %+v, want passthrough/passthrough", tt.name, extra)
			}
			// The legacy pre-intersect gate must skip spec-registered callees.
			if functionNeedsLegacyElementwisePreIntersect(tt.name) {
				t.Fatalf("functionNeedsLegacyElementwisePreIntersect(%q) = true, want false", tt.name)
			}
		})
	}
}

func TestLookupFamilyFuncSpecParity(t *testing.T) {
	// Test context simulates a formula at Sheet1!B2 in a non-array cell.
	scalarCtxB2 := &EvalContext{
		CurrentCol:     2,
		CurrentRow:     2,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}
	// Test context simulates a formula at Sheet1!B7 (for the SUM(XMATCH(...))
	// divergence where the anchor is on row 7).
	scalarCtxB7 := &EvalContext{
		CurrentCol:     2,
		CurrentRow:     7,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: false,
	}
	// Array-formula / CSE context: LegacyIntersectRef must NOT intersect.
	arrayCtxB2 := &EvalContext{
		CurrentCol:     2,
		CurrentRow:     2,
		CurrentSheet:   "Sheet1",
		IsArrayFormula: true,
	}

	// Shared lookup table: keys {A, B, C, D, E} → labels.
	lookupKeys := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("A")},
			{StringVal("B")},
			{StringVal("C")},
			{StringVal("D")},
			{StringVal("E")},
		},
		RangeOrigin: &RangeAddr{
			Sheet: "Sheet1", FromCol: 10, FromRow: 2, ToCol: 10, ToRow: 6,
		},
	}
	lookupLabels := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("A-label")},
			{StringVal("B-label")},
			{StringVal("C-label")},
			{StringVal("D-label")},
			{StringVal("E-label")},
		},
		RangeOrigin: &RangeAddr{
			Sheet: "Sheet1", FromCol: 11, FromRow: 2, ToCol: 11, ToRow: 6,
		},
	}

	// 10-row range-backed array mimicking a full column of categories in the
	// spill_functions_10k fixture — periodic A,B,C,D,E starting at row 2.
	catRows := make([][]Value, 10)
	cats := []string{"A", "B", "C", "D", "E"}
	for i := 0; i < 10; i++ {
		catRows[i] = []Value{StringVal(cats[i%5])}
	}
	categoryColumn := Value{
		Type:  ValueArray,
		Array: catRows,
		RangeOrigin: &RangeAddr{
			Sheet: "Sheet1", FromCol: 3, FromRow: 2, ToCol: 3, ToRow: 11,
		},
	}

	// Anonymous (non-range) array: LegacyIntersectRef should pass through.
	anonArrayArg0 := Value{
		Type: ValueArray,
		Array: [][]Value{
			{StringVal("C")},
		},
	}

	type scenario struct {
		name      string
		ctx       *EvalContext
		fnName    string
		args      []Value
		want      Value
		wantMatch string // describes what we expect re: legacy-vs-contract behaviour
	}

	scenarios := []scenario{
		{
			// Covers spill_functions_10k/05_xlookup_xmatch_10k.xlsx B2.
			// Non-array context with a range-backed array arg 0; the adapter
			// intersects to the current-row cell (B2 → row 2 → "A") and
			// XLOOKUP returns the scalar "A-label".
			name:      "xlookup_range_arg0_intersects_at_row_2",
			ctx:       scalarCtxB2,
			fnName:    "XLOOKUP",
			args:      []Value{categoryColumn, lookupKeys, lookupLabels},
			want:      StringVal("A-label"),
			wantMatch: "differs_from_legacy",
		},
		{
			// Same column, different evaluation row — row 7 intersects to the
			// 6th element in catRows (i=5) which is "A" again (periodic mod 5).
			name:      "xlookup_range_arg0_intersects_at_row_7",
			ctx:       scalarCtxB7,
			fnName:    "XLOOKUP",
			args:      []Value{categoryColumn, lookupKeys, lookupLabels},
			want:      StringVal("A-label"),
			wantMatch: "differs_from_legacy",
		},
		{
			// Covers spill_functions_10k/05_xlookup_xmatch_10k.xlsx B7.
			// Outer SUM wraps the XMATCH in the fixture, but at the XMATCH call
			// site the relevant ctx is still scalar (B7) — the adapter
			// intersects "A" out of the category column and XMATCH returns 1.
			name:      "xmatch_range_arg0_intersects_to_position_1",
			ctx:       scalarCtxB7,
			fnName:    "XMATCH",
			args:      []Value{categoryColumn, lookupKeys},
			want:      NumberVal(1),
			wantMatch: "differs_from_legacy",
		},
		{
			name:      "xlookup_scalar_arg0_no_intersection_needed",
			ctx:       scalarCtxB2,
			fnName:    "XLOOKUP",
			args:      []Value{StringVal("C"), lookupKeys, lookupLabels},
			want:      StringVal("C-label"),
			wantMatch: "same_as_legacy",
		},
		{
			name:      "xmatch_scalar_arg0_no_intersection_needed",
			ctx:       scalarCtxB2,
			fnName:    "XMATCH",
			args:      []Value{StringVal("D"), lookupKeys},
			want:      NumberVal(4),
			wantMatch: "same_as_legacy",
		},
		{
			// Anonymous arrays (no RangeOrigin) are now scalarized to their
			// top-left cell via ArgAdaptScalarizeAny — matching Excel's
			// legacy implicit intersection for the inline-array case. The
			// single-cell {{"C"}} collapses to "C" and XLOOKUP finds it.
			name:      "xlookup_anonymous_array_arg0_scalarizes_top_left",
			ctx:       scalarCtxB2,
			fnName:    "XLOOKUP",
			args:      []Value{anonArrayArg0, lookupKeys, lookupLabels},
			want:      StringVal("C-label"),
			wantMatch: "differs_from_legacy",
		},
		{
			// Inside an explicit array formula (CSE) LegacyIntersectRef does
			// not intersect, so the whole lookup_value array reaches the
			// FnKindLookupArrayLift dispatch, which fans it out per-element
			// and assembles an array of results — one label per category.
			name:   "xlookup_array_formula_ctx_fans_out_per_element",
			ctx:    arrayCtxB2,
			fnName: "XLOOKUP",
			args:   []Value{categoryColumn, lookupKeys, lookupLabels},
			want: Value{
				Type: ValueArray,
				Array: [][]Value{
					{StringVal("A-label")},
					{StringVal("B-label")},
					{StringVal("C-label")},
					{StringVal("D-label")},
					{StringVal("E-label")},
					{StringVal("A-label")},
					{StringVal("B-label")},
					{StringVal("C-label")},
					{StringVal("D-label")},
					{StringVal("E-label")},
				},
			},
			wantMatch: "differs_from_legacy",
		},
	}

	for _, tt := range scenarios {
		t.Run(tt.name, func(t *testing.T) {
			fn := registry[normalizeFuncName(tt.fnName)]
			if fn == nil {
				t.Fatalf("function %s not registered", tt.fnName)
			}
			spec, ok := funcSpecForName(tt.fnName)
			if !ok {
				t.Fatalf("FuncSpec missing for %s", tt.fnName)
			}

			legacyArgs := make([]Value, len(tt.args))
			copy(legacyArgs, tt.args)
			legacy, err := fn(legacyArgs, tt.ctx)
			if err != nil {
				t.Fatalf("legacy call: %v", err)
			}

			gotArgs := make([]Value, len(tt.args))
			copy(gotArgs, tt.args)
			got, err := callFuncWithSpec(tt.fnName, fn, spec, gotArgs, tt.ctx)
			if err != nil {
				t.Fatalf("contract call: %v", err)
			}

			assertLookupValueEqual(t, got, tt.want)

			switch tt.wantMatch {
			case "same_as_legacy":
				assertLookupValueEqual(t, got, legacy)
			case "differs_from_legacy":
				if legacy.Type == got.Type && legacy.Type == ValueError && legacy.Err == got.Err {
					t.Fatalf("expected contract to fix divergence but legacy=%v matched contract=%v", legacy, got)
				}
			default:
				t.Fatalf("unknown wantMatch %q", tt.wantMatch)
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
