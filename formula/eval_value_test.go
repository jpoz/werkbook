package formula

import "testing"

func TestValueEvalValueRoundTrip(t *testing.T) {
	tests := []struct {
		name string
		in   Value
	}{
		{name: "number", in: NumberVal(42)},
		{name: "string", in: StringVal("x")},
		{name: "bool", in: BoolVal(true)},
		{name: "error", in: ErrorVal(ErrValREF)},
		{name: "array", in: Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(1), NumberVal(2)},
			{NumberVal(3), NumberVal(4)},
		}}},
		{name: "range_array", in: Value{Type: ValueArray, Array: [][]Value{
			{NumberVal(10)},
			{NumberVal(20)},
		}, RangeOrigin: &RangeAddr{Sheet: "Sheet1", FromCol: 2, FromRow: 3, ToCol: 2, ToRow: 4}}},
		{name: "no_spill_array", in: Value{Type: ValueArray, Array: [][]Value{{NumberVal(1)}}, NoSpill: true}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ev := ValueToEvalValue(tt.in)
			got := EvalValueToValue(ev)
			assertLookupValueEqual(t, got, tt.in)
		})
	}
}

func TestValueEvalValueRoundTripSingleRef(t *testing.T) {
	in := Value{Type: ValueRef, Num: float64(7 + 9*100_000), Str: "Sheet2"}

	got := EvalValueToValue(ValueToEvalValue(in))
	if got.Type != ValueRef || got.Num != in.Num || got.Str != in.Str {
		t.Fatalf("round-trip = %#v, want %#v", got, in)
	}
}

func TestValueToEvalValueRangeBecomesRef(t *testing.T) {
	in := Value{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(1)}, {NumberVal(2)}},
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 1,
			FromRow: 2,
			ToCol:   1,
			ToRow:   3,
		},
	}

	got := ValueToEvalValue(in)
	if got.Kind != EvalRef {
		t.Fatalf("Kind = %v, want EvalRef", got.Kind)
	}
	if got.Ref == nil {
		t.Fatal("Ref = nil")
	}
	if bounds := got.Ref.Bounds(); bounds != *in.RangeOrigin {
		t.Fatalf("Bounds = %+v, want %+v", bounds, *in.RangeOrigin)
	}
}

func TestValueToEvalValueRangeWithSheetEndPreservesBounds(t *testing.T) {
	in := Value{
		Type:  ValueArray,
		Array: [][]Value{{NumberVal(1)}},
		RangeOrigin: &RangeAddr{
			Sheet:    "Sheet1",
			SheetEnd: "Sheet3",
			FromCol:  2,
			FromRow:  4,
			ToCol:    5,
			ToRow:    9,
		},
	}

	got := ValueToEvalValue(in)
	if got.Kind != EvalRef || got.Ref == nil {
		t.Fatalf("ValueToEvalValue = %#v, want EvalRef", got)
	}
	if bounds := got.Ref.Bounds(); bounds != *in.RangeOrigin {
		t.Fatalf("Bounds = %+v, want %+v", bounds, *in.RangeOrigin)
	}
}

func TestEvalValueToValueSingleCell3DRefPreservesRangeBoundary(t *testing.T) {
	addr := RangeAddr{
		Sheet:    "Sheet1",
		SheetEnd: "Sheet3",
		FromCol:  2,
		FromRow:  4,
		ToCol:    2,
		ToRow:    4,
	}

	got := EvalValueToValue(newEvalRangeRef(addr, [][]Value{{NumberVal(9)}}, nil, nil))
	if got.Type != ValueArray {
		t.Fatalf("EvalValueToValue(single-cell 3D ref) type = %v, want ValueArray", got.Type)
	}
	if got.RangeOrigin == nil || *got.RangeOrigin != addr {
		t.Fatalf("RangeOrigin = %+v, want %+v", got.RangeOrigin, addr)
	}

	roundTrip := ValueToEvalValue(got)
	if roundTrip.Kind != EvalRef || roundTrip.Ref == nil {
		t.Fatalf("ValueToEvalValue(roundTrip) = %#v, want EvalRef", roundTrip)
	}
	if bounds := roundTrip.Ref.Bounds(); bounds != addr {
		t.Fatalf("Bounds = %+v, want %+v", bounds, addr)
	}
}

func TestValueToEvalValueAnonymousArrayBecomesArray(t *testing.T) {
	in := Value{Type: ValueArray, Array: [][]Value{
		{StringVal("a"), StringVal("b")},
	}}

	got := ValueToEvalValue(in)
	if got.Kind != EvalArray {
		t.Fatalf("Kind = %v, want EvalArray", got.Kind)
	}
	if got.Array == nil {
		t.Fatal("Array = nil")
	}
	if got.Array.Rows != 1 || got.Array.Cols != 2 {
		t.Fatalf("dims = %dx%d, want 1x2", got.Array.Rows, got.Array.Cols)
	}
}

func TestValueToEvalValueFullColumnRangeKeepsTrimmedRefGrid(t *testing.T) {
	in := Value{
		Type: ValueArray,
		Array: [][]Value{
			{NumberVal(10)},
			{NumberVal(20)},
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

	got := valueToEvalValueWithResolver(in, &panicRangeResolver{})
	if got.Kind != EvalRef || got.Ref == nil {
		t.Fatalf("got = %#v, want EvalRef", got)
	}
	if got.Ref.Materialized == nil {
		t.Fatal("Materialized = nil")
	}
	if rows := got.Ref.Materialized.Rows(); rows != 3 {
		t.Fatalf("Rows = %d, want 3 trimmed rows", rows)
	}
	if cols := got.Ref.Materialized.Cols(); cols != 1 {
		t.Fatalf("Cols = %d, want 1", cols)
	}
}

func TestValueToEvalValueResolverFallbackLoadsRange(t *testing.T) {
	resolver := &mockResolver{
		cells: map[CellAddr]Value{
			{Sheet: "Sheet1", Col: 2, Row: 5}: NumberVal(11),
			{Sheet: "Sheet1", Col: 2, Row: 6}: NumberVal(13),
		},
	}
	in := Value{
		Type: ValueArray,
		RangeOrigin: &RangeAddr{
			Sheet:   "Sheet1",
			FromCol: 2,
			FromRow: 5,
			ToCol:   2,
			ToRow:   6,
		},
	}

	got := valueToEvalValueWithResolver(in, resolver)
	if got.Kind != EvalRef || got.Ref == nil {
		t.Fatalf("got = %#v, want EvalRef", got)
	}
	if got.Ref.Materialized == nil {
		t.Fatal("Materialized = nil")
	}
	if rows := got.Ref.Materialized.Rows(); rows != 2 {
		t.Fatalf("Rows = %d, want 2", rows)
	}
	if cols := got.Ref.Materialized.Cols(); cols != 1 {
		t.Fatalf("Cols = %d, want 1", cols)
	}
	if cell := EvalValueToValue(got.Ref.Materialized.Cell(1, 0)); cell.Type != ValueNumber || cell.Num != 13 {
		t.Fatalf("Cell(1,0) = %#v, want 13", cell)
	}
}

func TestSpillClassFromLegacy(t *testing.T) {
	tests := []struct {
		name string
		in   Value
		want SpillClass
	}{
		{
			name: "bounded_trimmed_range",
			in: Value{
				Type:  ValueArray,
				Array: [][]Value{{NumberVal(10)}},
				RangeOrigin: &RangeAddr{
					Sheet:   "Sheet1",
					FromCol: 1,
					FromRow: 2,
					ToCol:   1,
					ToRow:   4,
				},
			},
			want: SpillBounded,
		},
		{
			name: "full_axis_trimmed_range",
			in: Value{
				Type:  ValueArray,
				Array: [][]Value{{NumberVal(10)}},
				RangeOrigin: &RangeAddr{
					Sheet:   "Sheet1",
					FromCol: 1,
					FromRow: 1,
					ToCol:   1,
					ToRow:   maxRows,
				},
			},
			want: SpillUnbounded,
		},
		{
			name: "no_spill_array",
			in: Value{
				Type:    ValueArray,
				Array:   [][]Value{{NumberVal(10)}},
				NoSpill: true,
			},
			want: SpillScalarOnly,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := spillClassFromLegacy(tt.in); got != tt.want {
				t.Fatalf("spillClassFromLegacy = %v, want %v", got, tt.want)
			}
		})
	}
}
