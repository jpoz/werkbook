package formula

import "testing"

func v2RangeValue(rows [][]Value, addr RangeAddr) Value {
	return Value{
		Type:        ValueArray,
		Array:       rows,
		RangeOrigin: &addr,
	}
}

func TestV2ValueToEvalValueNoSpillRangeOriginStaysArray(t *testing.T) {
	addr := RangeAddr{
		Sheet:   "Sheet1",
		FromCol: 2,
		FromRow: 3,
		ToCol:   2,
		ToRow:   5,
	}
	in := v2RangeValue([][]Value{{NumberVal(10)}}, addr)
	in.NoSpill = true

	got := ValueToEvalValue(in)
	if got.Kind != EvalArray || got.Array == nil {
		t.Fatalf("ValueToEvalValue = %#v, want EvalArray", got)
	}
	if got.Array.SpillClass != SpillScalarOnly {
		t.Fatalf("SpillClass = %v, want %v", got.Array.SpillClass, SpillScalarOnly)
	}
	if got.Array.Origin == nil || got.Array.Origin.Range == nil {
		t.Fatal("Origin.Range = nil, want original range")
	}
	if *got.Array.Origin.Range != addr {
		t.Fatalf("Origin.Range = %+v, want %+v", *got.Array.Origin.Range, addr)
	}
	if got.Array.Rows != 3 || got.Array.Cols != 1 {
		t.Fatalf("array dims = %dx%d, want 3x1", got.Array.Rows, got.Array.Cols)
	}
}

func TestV2ValueToEvalValueCellOriginBecomesSingleCellRef(t *testing.T) {
	cell := CellAddr{Sheet: "Sheet1", Col: 3, Row: 4}
	in := NumberVal(42)
	in.CellOrigin = &cell

	got := ValueToEvalValue(in)
	if got.Kind != EvalRef || got.Ref == nil {
		t.Fatalf("ValueToEvalValue = %#v, want EvalRef", got)
	}
	if bounds := got.Ref.Bounds(); bounds != (RangeAddr{Sheet: "Sheet1", FromCol: 3, FromRow: 4, ToCol: 3, ToRow: 4}) {
		t.Fatalf("Bounds = %+v, want Sheet1!C4", bounds)
	}
	if got.Ref.Materialized == nil {
		t.Fatal("Materialized = nil, want single-cell grid")
	}
	assertLookupValueEqual(t, EvalValueToValue(got.Ref.Materialized.Cell(0, 0)), NumberVal(42))

	legacy := EvalValueToValue(got)
	assertLookupValueEqual(t, legacy, in)
	if legacy.CellOrigin == nil || *legacy.CellOrigin != cell {
		t.Fatalf("CellOrigin = %+v, want %+v", legacy.CellOrigin, cell)
	}
}

func TestV2LookupArrayLiftUsesLookupArgShapeOnly(t *testing.T) {
	lookupValues := Value{
		Type: ValueArray,
		Array: [][]Value{{
			StringVal("west"),
			StringVal("east"),
		}},
	}
	lookupRange := v2RangeValue([][]Value{
		{StringVal("east")},
		{StringVal("west")},
	}, RangeAddr{Sheet: "Sheet1", FromCol: 1, FromRow: 1, ToCol: 1, ToRow: 2})
	returnRange := v2RangeValue([][]Value{
		{NumberVal(10)},
		{NumberVal(20)},
	}, RangeAddr{Sheet: "Sheet1", FromCol: 2, FromRow: 1, ToCol: 2, ToRow: 2})

	got, err := evalXLOOKUP([]EvalValue{
		ValueToEvalValue(lookupValues),
		ValueToEvalValue(lookupRange),
		ValueToEvalValue(returnRange),
	}, nil)
	if err != nil {
		t.Fatalf("evalXLOOKUP: %v", err)
	}
	if got.Kind != EvalArray || got.Array == nil {
		t.Fatalf("evalXLOOKUP = %#v, want EvalArray", got)
	}
	if got.Array.Rows != 1 || got.Array.Cols != 2 {
		t.Fatalf("result dims = %dx%d, want lookup_value shape 1x2", got.Array.Rows, got.Array.Cols)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), Value{Type: ValueArray, Array: [][]Value{{
		NumberVal(20),
		NumberVal(10),
	}}})
}

func TestV2XLOOKUPReturnArrayPreservesRowAndColumnShape(t *testing.T) {
	verticalLookup := v2RangeValue([][]Value{
		{StringVal("alpha")},
		{StringVal("beta")},
	}, RangeAddr{Sheet: "Sheet1", FromCol: 1, FromRow: 1, ToCol: 1, ToRow: 2})
	twoColumnReturn := v2RangeValue([][]Value{
		{NumberVal(10), NumberVal(100)},
		{NumberVal(20), NumberVal(200)},
	}, RangeAddr{Sheet: "Sheet1", FromCol: 2, FromRow: 1, ToCol: 3, ToRow: 2})

	got, err := evalXLOOKUP([]EvalValue{
		ValueToEvalValue(StringVal("beta")),
		ValueToEvalValue(verticalLookup),
		ValueToEvalValue(twoColumnReturn),
	}, nil)
	if err != nil {
		t.Fatalf("evalXLOOKUP vertical return: %v", err)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), Value{Type: ValueArray, Array: [][]Value{{
		NumberVal(20),
		NumberVal(200),
	}}})

	horizontalLookup := v2RangeValue([][]Value{{
		StringVal("q1"),
		StringVal("q2"),
	}}, RangeAddr{Sheet: "Sheet1", FromCol: 1, FromRow: 1, ToCol: 2, ToRow: 1})
	twoRowReturn := v2RangeValue([][]Value{
		{NumberVal(10), NumberVal(20)},
		{NumberVal(100), NumberVal(200)},
	}, RangeAddr{Sheet: "Sheet1", FromCol: 1, FromRow: 2, ToCol: 2, ToRow: 3})

	got, err = evalXLOOKUP([]EvalValue{
		ValueToEvalValue(StringVal("q2")),
		ValueToEvalValue(horizontalLookup),
		ValueToEvalValue(twoRowReturn),
	}, nil)
	if err != nil {
		t.Fatalf("evalXLOOKUP horizontal return: %v", err)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(20)},
		{NumberVal(200)},
	}})
}

func TestV2MATCHFullColumnRefUsesMaterializedBounds(t *testing.T) {
	fullColumn := v2RangeValue([][]Value{
		{StringVal("alpha")},
		{StringVal("beta")},
		{StringVal("gamma")},
	}, RangeAddr{Sheet: "Sheet1", FromCol: 4, FromRow: 1, ToCol: 4, ToRow: maxRows})
	lookupArg := ValueToEvalValue(fullColumn)
	if lookupArg.Kind != EvalRef || lookupArg.Ref == nil {
		t.Fatalf("lookup arg = %#v, want EvalRef", lookupArg)
	}
	if bounds := lookupArg.Ref.Bounds(); bounds.ToRow != maxRows {
		t.Fatalf("ref bounds = %+v, want full-column bounds", bounds)
	}

	got, err := evalMATCH([]EvalValue{
		ValueToEvalValue(StringVal("gamma")),
		lookupArg,
		ValueToEvalValue(NumberVal(0)),
	}, nil)
	if err != nil {
		t.Fatalf("evalMATCH: %v", err)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), NumberVal(3))
}

func TestV2FrequencyEvalFullColumnRefUsesMaterializedRows(t *testing.T) {
	data := v2RangeValue([][]Value{
		{NumberVal(1)},
		{StringVal("ignored")},
		{NumberVal(8)},
	}, RangeAddr{Sheet: "Sheet1", FromCol: 1, FromRow: 1, ToCol: 1, ToRow: maxRows})
	bins := v2RangeValue([][]Value{
		{NumberVal(5)},
	}, RangeAddr{Sheet: "Sheet1", FromCol: 2, FromRow: 1, ToCol: 2, ToRow: maxRows})

	got, err := evalFREQUENCYDirectRange([]EvalValue{
		ValueToEvalValue(data),
		ValueToEvalValue(bins),
	}, nil)
	if err != nil {
		t.Fatalf("evalFREQUENCYDirectRange: %v", err)
	}
	if got.Kind != EvalArray || got.Array == nil {
		t.Fatalf("evalFREQUENCYDirectRange = %#v, want EvalArray", got)
	}
	if got.Array.SpillClass != SpillBounded {
		t.Fatalf("SpillClass = %v, want %v", got.Array.SpillClass, SpillBounded)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), Value{Type: ValueArray, Array: [][]Value{
		{NumberVal(1)},
		{NumberVal(1)},
	}})
}

func TestV2ShapeEvalPreservesTrimmedRangeBlankTail(t *testing.T) {
	trimmedColumn := v2RangeValue([][]Value{
		{NumberVal(10)},
	}, RangeAddr{Sheet: "Sheet1", FromCol: 1, FromRow: 1, ToCol: 1, ToRow: 3})

	got, err := evalTRANSPOSE([]EvalValue{ValueToEvalValue(trimmedColumn)}, nil)
	if err != nil {
		t.Fatalf("evalTRANSPOSE: %v", err)
	}
	if got.Kind != EvalArray || got.Array == nil {
		t.Fatalf("evalTRANSPOSE = %#v, want EvalArray", got)
	}
	if got.Array.Rows != 1 || got.Array.Cols != 3 {
		t.Fatalf("result dims = %dx%d, want 1x3", got.Array.Rows, got.Array.Cols)
	}
	assertLookupValueEqual(t, EvalValueToValue(got), Value{Type: ValueArray, Array: [][]Value{{
		NumberVal(10),
		EmptyVal(),
		EmptyVal(),
	}}})
}
