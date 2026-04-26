package formula

import "testing"

func TestIndexFuncSpec(t *testing.T) {
	spec, ok := funcSpecForName("INDEX")
	if !ok {
		t.Fatal("FuncSpec missing for INDEX")
	}
	if spec.Kind != FnKindLookup {
		t.Fatalf("Kind = %v, want %v", spec.Kind, FnKindLookup)
	}
	if spec.Return != ReturnModePassThrough {
		t.Fatalf("Return = %v, want %v", spec.Return, ReturnModePassThrough)
	}
	if spec.Eval == nil {
		t.Fatal("Eval = nil for INDEX")
	}
	for i := 0; i < 3; i++ {
		arg, ok := funcArgSpec(spec, i)
		if !ok {
			t.Fatalf("missing arg contract for INDEX arg %d", i)
		}
		want := ArgSpec{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough}
		if arg != want {
			t.Fatalf("arg %d = %+v, want %+v", i, arg, want)
		}
	}
}

func TestIndexFuncSpecParity(t *testing.T) {
	boundedRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(10), NumberVal(20)},
			{NumberVal(30), NumberVal(40)},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 1,
			FromRow: 1,
			ToCol:   2,
			ToRow:   2,
		},
	}
	singleColumn := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(10)},
			{NumberVal(20)},
			{NumberVal(30)},
			{NumberVal(40)},
			{NumberVal(50)},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 3,
			FromRow: 1,
			ToCol:   3,
			ToRow:   5,
		},
	}
	fullColumn := Value{
		Type:  ValueArray,
		Array: [][]Value{{EmptyVal()}},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 7,
			FromRow: 1,
			ToCol:   7,
			ToRow:   maxRows,
		},
	}

	tests := []struct {
		name string
		args []Value
		want Value
	}{
		{
			name: "bounded_scalar",
			args: []Value{boundedRange, NumberVal(2), NumberVal(2)},
			want: NumberVal(40),
		},
		{
			name: "direct_range_row0_no_spill",
			args: []Value{boundedRange, NumberVal(0), NumberVal(1)},
			want: Value{
				Type: ValueArray,
				Array: [][]Value{
					{NumberVal(10)},
					{NumberVal(30)},
				},
				RangeOrigin: &RangeAddr{
					Sheet:   "Sheet1",
					FromCol: 1,
					FromRow: 1,
					ToCol:   1,
					ToRow:   2,
				},
				NoSpill: true,
			},
		},
		{
			name: "anonymous_array_row0_spills",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}, {NumberVal(3), NumberVal(4)}}},
				NumberVal(0),
				NumberVal(1),
			},
			want: Value{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(3)}}},
		},
		{
			name: "selector_array_preserves_shape",
			args: []Value{
				singleColumn,
				Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(3), NumberVal(5)}}},
			},
			want: Value{Type: ValueArray, Array: [][]Value{{NumberVal(10), NumberVal(30), NumberVal(50)}}},
		},
		{
			name: "full_column_two_arg_empty",
			args: []Value{fullColumn, NumberVal(2)},
			want: EmptyVal(),
		},
		{
			name: "full_column_row0_no_spill",
			args: []Value{fullColumn, NumberVal(0), NumberVal(1)},
			want: Value{
				Type:  ValueArray,
				Array: [][]Value{{EmptyVal()}},
				RangeOrigin: &RangeAddr{
					Sheet:   "Sheet1",
					FromCol: 7,
					FromRow: 1,
					ToCol:   7,
					ToRow:   maxRows,
				},
				NoSpill: true,
			},
		},
		{
			name: "single_row_two_arg_zero",
			args: []Value{
				{Type: ValueArray, Array: [][]Value{{StringVal("OUT"), StringVal("IN")}}},
				NumberVal(0),
			},
			want: Value{Type: ValueArray, Array: [][]Value{{StringVal("OUT"), StringVal("IN")}}},
		},
	}

	spec, ok := funcSpecForName("INDEX")
	if !ok {
		t.Fatal("FuncSpec missing for INDEX")
	}
	fn := registry[normalizeFuncName("INDEX")]
	if fn == nil {
		t.Fatal("INDEX not registered")
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			legacy, err := fnINDEX(tt.args)
			if err != nil {
				t.Fatalf("legacy call: %v", err)
			}
			got, err := callFuncWithSpec("INDEX", fn, spec, tt.args, nil)
			if err != nil {
				t.Fatalf("contract call: %v", err)
			}

			assertLookupValueEqual(t, got, legacy)
			assertLookupValueEqual(t, got, tt.want)

			loaded := loadEvalFuncArg(ArgSpec{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough}, tt.args[0], nil)
			if tt.args[0].RangeOrigin != nil {
				if loaded.Kind != EvalRef || loaded.Ref == nil {
					t.Fatalf("loaded arg0 = %#v, want EvalRef", loaded)
				}
			}
		})
	}
}

func TestSelectorFuncSpecs(t *testing.T) {
	for _, name := range []string{"TAKE", "DROP", "CHOOSECOLS", "CHOOSEROWS"} {
		t.Run(name, func(t *testing.T) {
			spec, ok := funcSpecForName(name)
			if !ok {
				t.Fatalf("FuncSpec missing for %s", name)
			}
			if spec.Kind != FnKindLookup {
				t.Fatalf("Kind = %v, want %v", spec.Kind, FnKindLookup)
			}
			if spec.Return != ReturnModePassThrough {
				t.Fatalf("Return = %v, want %v", spec.Return, ReturnModePassThrough)
			}
			if spec.Eval == nil {
				t.Fatalf("Eval = nil for %s", name)
			}
			for i := 0; i < 4; i++ {
				arg, ok := funcArgSpec(spec, i)
				if !ok {
					t.Fatalf("missing arg contract for %s arg %d", name, i)
				}
				want := ArgSpec{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough}
				if arg != want {
					t.Fatalf("arg %d = %+v, want %+v", i, arg, want)
				}
			}
		})
	}
}

func TestSelectorFuncSpecParity(t *testing.T) {
	type selectorDirectFunc func([]Value) (Value, error)

	boundedRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(10), NumberVal(20), NumberVal(30)},
			{NumberVal(40), NumberVal(50), NumberVal(60)},
			{NumberVal(70), NumberVal(80), NumberVal(90)},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 1,
			FromRow: 1,
			ToCol:   3,
			ToRow:   3,
		},
	}
	trimmedCol := trimmedRangeValue([][]Value{{NumberVal(10)}}, 7, 1, 7, 3)
	fullColumn := Value{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(10)}},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 7,
			FromRow: 1,
			ToCol:   7,
			ToRow:   maxRows,
		},
	}

	tests := []struct {
		name        string
		fnName      string
		direct      selectorDirectFunc
		args        []Value
		want        Value
		wantRange   *RangeAddr
		wantNoSpill bool
	}{
		{
			name:   "take_trimmed_range_preserves_window_origin",
			fnName: "TAKE",
			direct: fnTAKE,
			args:   []Value{trimmedCol, NumberVal(2)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10)},
				{EmptyVal()},
			}},
			wantRange: &RangeAddr{
				FromCol: 7,
				FromRow: 1,
				ToCol:   7,
				ToRow:   2,
			},
		},
		{
			name:   "take_full_column_small_window_stays_bounded",
			fnName: "TAKE",
			direct: fnTAKE,
			args:   []Value{fullColumn, NumberVal(2)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10)},
				{EmptyVal()},
			}},
			wantRange: &RangeAddr{
				Sheet:   "Sheet1",
				FromCol: 7,
				FromRow: 1,
				ToCol:   7,
				ToRow:   2,
			},
		},
		{
			name:   "drop_bounded_range_preserves_window_origin",
			fnName: "DROP",
			direct: fnDROP,
			args:   []Value{boundedRange, NumberVal(1)},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(40), NumberVal(50), NumberVal(60)},
				{NumberVal(70), NumberVal(80), NumberVal(90)},
			}},
			wantRange: &RangeAddr{
				Sheet:   "Sheet1",
				FromCol: 1,
				FromRow: 2,
				ToCol:   3,
				ToRow:   3,
			},
		},
		{
			name:   "choosecols_selector_array_contiguous_preserves_origin",
			fnName: "CHOOSECOLS",
			direct: fnCHOOSECOLS,
			args: []Value{
				boundedRange,
				Value{Type: ValueArray, Array: [][]Value{{NumberVal(1), NumberVal(2)}}},
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10), NumberVal(20)},
				{NumberVal(40), NumberVal(50)},
				{NumberVal(70), NumberVal(80)},
			}},
			wantRange: &RangeAddr{
				Sheet:   "Sheet1",
				FromCol: 1,
				FromRow: 1,
				ToCol:   2,
				ToRow:   3,
			},
		},
		{
			name:   "choosecols_selector_array_noncontiguous_materializes",
			fnName: "CHOOSECOLS",
			direct: fnCHOOSECOLS,
			args: []Value{
				boundedRange,
				Value{Type: ValueArray, Array: [][]Value{{NumberVal(3), NumberVal(1)}}},
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(30), NumberVal(10)},
				{NumberVal(60), NumberVal(40)},
				{NumberVal(90), NumberVal(70)},
			}},
		},
		{
			name:   "chooserows_selector_array_contiguous_preserves_origin",
			fnName: "CHOOSEROWS",
			direct: fnCHOOSEROWS,
			args: []Value{
				boundedRange,
				Value{Type: ValueArray, Array: [][]Value{{NumberVal(2)}, {NumberVal(3)}}},
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(40), NumberVal(50), NumberVal(60)},
				{NumberVal(70), NumberVal(80), NumberVal(90)},
			}},
			wantRange: &RangeAddr{
				Sheet:   "Sheet1",
				FromCol: 1,
				FromRow: 2,
				ToCol:   3,
				ToRow:   3,
			},
		},
		{
			name:   "chooserows_full_column_selector_array_stays_unbounded",
			fnName: "CHOOSEROWS",
			direct: fnCHOOSEROWS,
			args: []Value{
				fullColumn,
				Value{Type: ValueArray, Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}}},
			},
			want: Value{Type: ValueArray, Array: [][]Value{
				{NumberVal(10)},
				{EmptyVal()},
			}},
			wantRange: &RangeAddr{
				Sheet:   "Sheet1",
				FromCol: 7,
				FromRow: 1,
				ToCol:   7,
				ToRow:   2,
			},
		},
		{
			name:   "take_scalar_stays_scalar",
			fnName: "TAKE",
			direct: fnTAKE,
			args:   []Value{NumberVal(42), NumberVal(1)},
			want:   NumberVal(42),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			spec, ok := funcSpecForName(tt.fnName)
			if !ok {
				t.Fatalf("FuncSpec missing for %s", tt.fnName)
			}
			fn := registry[normalizeFuncName(tt.fnName)]
			if fn == nil {
				t.Fatalf("%s not registered", tt.fnName)
			}

			legacy, err := tt.direct(tt.args)
			if err != nil {
				t.Fatalf("legacy call: %v", err)
			}
			got, err := callFuncWithSpec(tt.fnName, fn, spec, tt.args, nil)
			if err != nil {
				t.Fatalf("contract call: %v", err)
			}

			assertLookupValueEqual(t, got, legacy)
			assertLookupValueEqual(t, got, tt.want)
			assertSelectorMetadata(t, legacy, tt.wantRange, tt.wantNoSpill)
			assertSelectorMetadata(t, got, tt.wantRange, tt.wantNoSpill)

			if tt.args[0].RangeOrigin != nil {
				loaded := loadEvalFuncArg(ArgSpec{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough}, tt.args[0], nil)
				if loaded.Kind != EvalRef || loaded.Ref == nil {
					t.Fatalf("loaded arg0 = %#v, want EvalRef", loaded)
				}
			}
		})
	}
}

func TestSelectorEvalKeepsDirectRefWindowMaterializedBounds(t *testing.T) {
	fullColumn := Value{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(10)}},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 7,
			FromRow: 1,
			ToCol:   7,
			ToRow:   maxRows,
		},
	}

	got, err := evalINDEXSelector([]EvalValue{
		ValueToEvalValue(fullColumn),
		ValueToEvalValue(NumberVal(0)),
		ValueToEvalValue(NumberVal(1)),
	}, nil)
	if err != nil {
		t.Fatalf("evalINDEXSelector: %v", err)
	}
	if got.Kind != EvalArray || got.Array == nil {
		t.Fatalf("got = %#v, want EvalArray", got)
	}
	if got.Array.Rows != maxRows || got.Array.Cols != 1 {
		t.Fatalf("array dims = %dx%d, want %dx1", got.Array.Rows, got.Array.Cols, maxRows)
	}
	if got.Array.SpillClass != SpillScalarOnly {
		t.Fatalf("SpillClass = %v, want %v", got.Array.SpillClass, SpillScalarOnly)
	}
	if got.Array.Origin == nil || got.Array.Origin.Range == nil {
		t.Fatal("Origin.Range = nil, want full-column origin")
	}
	wantRange := RangeAddr{
		Sheet:   "Sheet1",
		FromCol: 7,
		FromRow: 1,
		ToCol:   7,
		ToRow:   maxRows,
	}
	if *got.Array.Origin.Range != wantRange {
		t.Fatalf("Origin.Range = %+v, want %+v", *got.Array.Origin.Range, wantRange)
	}
	if got.Array.Grid == nil {
		t.Fatal("Grid = nil, want trimmed materialized grid")
	}
	if got.Array.Grid.Rows() != 1 || got.Array.Grid.Cols() != 1 {
		t.Fatalf("grid dims = %dx%d, want 1x1", got.Array.Grid.Rows(), got.Array.Grid.Cols())
	}
	assertLookupValueEqual(t, EvalValueToValue(got.Array.Grid.Cell(0, 0)), NumberVal(10))
}

func TestSelectorEvalIndexScalarPreservesSingleCellOrigin(t *testing.T) {
	boundedRange := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(10), NumberVal(20)},
			{NumberVal(30), NumberVal(40)},
		},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 1,
			FromRow: 1,
			ToCol:   2,
			ToRow:   2,
		},
	}

	got, err := evalINDEXSelector([]EvalValue{
		ValueToEvalValue(boundedRange),
		ValueToEvalValue(NumberVal(2)),
		ValueToEvalValue(NumberVal(2)),
	}, nil)
	if err != nil {
		t.Fatalf("evalINDEXSelector: %v", err)
	}
	if got.Kind != EvalRef || got.Ref == nil {
		t.Fatalf("got = %#v, want EvalRef", got)
	}
	if got.Ref.FromCol != 2 || got.Ref.FromRow != 2 || got.Ref.ToCol != 2 || got.Ref.ToRow != 2 {
		t.Fatalf("ref bounds = %+v, want single cell B2", got.Ref.Bounds())
	}

	legacy := EvalValueToValue(got)
	assertLookupValueEqual(t, legacy, NumberVal(40))
	if legacy.CellOrigin == nil {
		t.Fatal("CellOrigin = nil, want Sheet1!B2")
	}
	wantCell := CellAddr{Sheet: "Sheet1", Col: 2, Row: 2}
	if *legacy.CellOrigin != wantCell {
		t.Fatalf("CellOrigin = %+v, want %+v", *legacy.CellOrigin, wantCell)
	}
}

func assertSelectorMetadata(t *testing.T, got Value, wantRange *RangeAddr, wantNoSpill bool) {
	t.Helper()

	if got.Type != ValueArray {
		if wantRange != nil {
			t.Fatalf("RangeOrigin = nil, want %+v", *wantRange)
		}
		return
	}

	if wantRange == nil {
		if got.RangeOrigin != nil {
			t.Fatalf("RangeOrigin = %+v, want nil", *got.RangeOrigin)
		}
	} else {
		if got.RangeOrigin == nil {
			t.Fatalf("RangeOrigin = nil, want %+v", *wantRange)
		}
		if *got.RangeOrigin != *wantRange {
			t.Fatalf("RangeOrigin = %+v, want %+v", *got.RangeOrigin, *wantRange)
		}
	}

	if got.NoSpill != wantNoSpill {
		t.Fatalf("NoSpill = %v, want %v", got.NoSpill, wantNoSpill)
	}
}
