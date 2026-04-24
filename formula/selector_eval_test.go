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
