package formula

import "testing"

func TestCriteriaFuncSpecHelpers(t *testing.T) {
	tests := []struct {
		name string
		spec FuncSpec
		want []ArgSpec
	}{
		{
			name: "COUNTIF",
			spec: criteriaSingleIfFuncSpec(evalCOUNTIFCriteria),
			want: []ArgSpec{
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			},
		},
		{
			name: "SUMIF",
			spec: criteriaSingleIfAggregateFuncSpec(evalSUMIFCriteria),
			want: []ArgSpec{
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
			},
		},
		{
			name: "COUNTIFS",
			spec: criteriaPairsFuncSpec(evalCOUNTIFSCriteria),
			want: []ArgSpec{
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadDirectRange, Adapt: ArgAdaptPassThrough},
				{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough},
			},
		},
		{
			name: "SUMIFS",
			spec: criteriaPairsAggregateFuncSpec(evalSUMIFSCriteria),
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
			if tt.spec.Kind != FnKindReduction {
				t.Fatalf("Kind = %v, want %v", tt.spec.Kind, FnKindReduction)
			}
			if tt.spec.Return != ReturnModePassThrough {
				t.Fatalf("Return = %v, want %v", tt.spec.Return, ReturnModePassThrough)
			}
			if tt.spec.Eval == nil {
				t.Fatal("Eval = nil")
			}
			for i, want := range tt.want {
				got, ok := funcArgSpec(tt.spec, i)
				if !ok {
					t.Fatalf("missing arg spec %d", i)
				}
				if got != want {
					t.Fatalf("arg %d = %+v, want %+v", i, got, want)
				}
			}
		})
	}
}

func TestEvalCOUNTIFCriteriaScalarAndArray(t *testing.T) {
	rangeArg := criteriaBoundedRef("Sheet1", 1, 1, [][]Value{
		{StringVal("apple")},
		{StringVal("pear")},
		{StringVal("apple")},
	})

	got, err := evalCOUNTIFCriteria([]EvalValue{
		rangeArg,
		ValueToEvalValue(StringVal("apple")),
	}, nil)
	if err != nil {
		t.Fatalf("scalar criteria: %v", err)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), NumberVal(2))

	got, err = evalCOUNTIFCriteria([]EvalValue{
		rangeArg,
		ValueToEvalValue(Value{Type: ValueArray, Array: [][]Value{
			{StringVal("apple"), StringVal("pear")},
		}}),
	}, nil)
	if err != nil {
		t.Fatalf("array criteria: %v", err)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(2), NumberVal(1)},
	}})
}

func TestEvalAVERAGEIFCriteriaArray(t *testing.T) {
	rangeArg := criteriaBoundedRef("Sheet1", 2, 3, [][]Value{
		{NumberVal(10)},
		{NumberVal(20)},
		{NumberVal(30)},
		{NumberVal(40)},
	})

	got, err := evalAVERAGEIFCriteria([]EvalValue{
		rangeArg,
		ValueToEvalValue(Value{Type: ValueArray, Array: [][]Value{
			{StringVal(">15"), StringVal(">30")},
		}}),
	}, nil)
	if err != nil {
		t.Fatalf("AVERAGEIF array criteria: %v", err)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(30), NumberVal(40)},
	}})
}

func TestEvalCOUNTIFSCriteriaMixedScalarAndArrayCriteria(t *testing.T) {
	regionRange := criteriaBoundedRef("Sheet1", 1, 1, [][]Value{
		{StringVal("East")},
		{StringVal("West")},
		{StringVal("East")},
		{StringVal("West")},
	})
	statusRange := criteriaBoundedRef("Sheet1", 2, 1, [][]Value{
		{StringVal("Open")},
		{StringVal("Open")},
		{StringVal("Closed")},
		{StringVal("Open")},
	})

	got, err := evalCOUNTIFSCriteria([]EvalValue{
		regionRange,
		ValueToEvalValue(Value{Type: ValueArray, Array: [][]Value{
			{StringVal("East"), StringVal("West")},
		}}),
		statusRange,
		ValueToEvalValue(StringVal("Open")),
	}, nil)
	if err != nil {
		t.Fatalf("COUNTIFS: %v", err)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(1), NumberVal(2)},
	}})
}

func TestEvalSUMIFCriteriaErrorPropagation(t *testing.T) {
	rangeArg := criteriaBoundedRef("Sheet1", 1, 1, [][]Value{
		{StringVal("East")},
		{StringVal("West")},
	})
	sumRange := criteriaBoundedRef("Sheet1", 2, 1, [][]Value{
		{ErrorVal(ErrValNAME)},
		{NumberVal(10)},
	})

	got, err := evalSUMIFCriteria([]EvalValue{
		rangeArg,
		ValueToEvalValue(StringVal("East")),
		sumRange,
	}, nil)
	if err != nil {
		t.Fatalf("SUMIF matched error: %v", err)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), ErrorVal(ErrValNAME))

	got, err = evalSUMIFCriteria([]EvalValue{
		rangeArg,
		ValueToEvalValue(StringVal("West")),
		sumRange,
	}, nil)
	if err != nil {
		t.Fatalf("SUMIF unmatched error: %v", err)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), NumberVal(10))
}

func TestEvalSUMIFSCriteriaFullColumnSparseRef(t *testing.T) {
	resolver := &sparseResolver{
		cells: map[CellAddr]Value{
			{Sheet: "Sheet1", Col: 1, Row: 1}: NumberVal(10),
			{Sheet: "Sheet1", Col: 1, Row: 3}: NumberVal(30),
			{Sheet: "Sheet1", Col: 2, Row: 1}: StringVal("East"),
			{Sheet: "Sheet1", Col: 2, Row: 3}: StringVal("East"),
		},
	}
	sumRange := criteriaFullColumnRef("Sheet1", 1, nil, resolver)
	critRange := criteriaFullColumnRef("Sheet1", 2, nil, resolver)

	got, err := evalSUMIFSCriteria([]EvalValue{
		sumRange,
		critRange,
		ValueToEvalValue(StringVal("East")),
	}, nil)
	if err != nil {
		t.Fatalf("SUMIFS full column: %v", err)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), NumberVal(40))
}

func TestEvalAVERAGEIFSCriteriaFullRowArrayCriteriaAndError(t *testing.T) {
	resolver := &panicRangeResolver{}
	avgRange := criteriaFullRowRef("Sheet1", 5, [][]Value{{
		NumberVal(10), NumberVal(20), ErrorVal(ErrValDIV0),
	}}, resolver)
	regionRange := criteriaFullRowRef("Sheet1", 6, [][]Value{{
		StringVal("East"), StringVal("East"), StringVal("East"),
	}}, resolver)
	statusRange := criteriaFullRowRef("Sheet1", 7, [][]Value{{
		StringVal("Open"), StringVal("Open"), StringVal("Closed"),
	}}, resolver)

	got, err := evalAVERAGEIFSCriteria([]EvalValue{
		avgRange,
		regionRange,
		ValueToEvalValue(StringVal("East")),
		statusRange,
		ValueToEvalValue(Value{Type: ValueArray, Array: [][]Value{
			{StringVal("Open"), StringVal("Closed")},
		}}),
	}, nil)
	if err != nil {
		t.Fatalf("AVERAGEIFS full row: %v", err)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(15), ErrorVal(ErrValDIV0)},
	}})
	if resolver.rangeCalls != 0 {
		t.Fatalf("resolver range calls = %d, want 0", resolver.rangeCalls)
	}
}

func criteriaBoundedRef(sheet string, fromCol, fromRow int, rows [][]Value) EvalValue {
	toCol := fromCol + materializedArrayCols(rows) - 1
	toRow := fromRow + len(rows) - 1
	return valueToEvalValueWithResolver(Value{
		Type:  ValueArray,
		Array: rows,
		RangeOrigin: &RangeAddr{
			Sheet:   sheet,
			FromCol: fromCol,
			FromRow: fromRow,
			ToCol:   toCol,
			ToRow:   toRow,
		},
	}, nil)
}

func criteriaFullColumnRef(sheet string, col int, rows [][]Value, resolver CellResolver) EvalValue {
	return valueToEvalValueWithResolver(Value{
		Type:  ValueArray,
		Array: rows,
		RangeOrigin: &RangeAddr{
			Sheet:   sheet,
			FromCol: col,
			FromRow: 1,
			ToCol:   col,
			ToRow:   maxRows,
		},
	}, resolver)
}

func criteriaFullRowRef(sheet string, row int, rows [][]Value, resolver CellResolver) EvalValue {
	return valueToEvalValueWithResolver(Value{
		Type:  ValueArray,
		Array: rows,
		RangeOrigin: &RangeAddr{
			Sheet:   sheet,
			FromCol: 1,
			FromRow: row,
			ToCol:   maxCols,
			ToRow:   row,
		},
	}, resolver)
}
