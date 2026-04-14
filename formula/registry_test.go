package formula

import "testing"

func TestInheritedArrayMetadataParity(t *testing.T) {
	tests := []struct {
		name      string
		kind      FnKind
		inherited map[int]bool
	}{
		{name: "IFERROR", kind: FnKindScalarLifted, inherited: map[int]bool{0: true, 1: true}},
		{name: "IFNA", kind: FnKindScalarLifted, inherited: map[int]bool{0: true, 1: true}},
		{name: "ABS", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "INT", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "SIGN", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "SQRT", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "LN", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "LOG10", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "EXP", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "FACT", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "SIN", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "COS", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "TAN", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "ASIN", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "ACOS", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "ATAN", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "SINH", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "COSH", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "TANH", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "DATEVALUE", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "ISNUMBER", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "ISTEXT", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "ISBLANK", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "ISERROR", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "ISERR", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "ISNA", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "NOT", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "N", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
		{name: "TYPE", kind: FnKindScalarLifted, inherited: map[int]bool{0: true}},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			meta, ok := funcMetaForName(tt.name)
			if !ok {
				t.Fatalf("metadata missing for %s", tt.name)
			}
			if meta.Kind != tt.kind {
				t.Fatalf("kind = %v, want %v", meta.Kind, tt.kind)
			}
			for i := 0; i < 3; i++ {
				want := tt.inherited[i]
				if got := inheritedArrayEvalForFuncArg(tt.name, i); got != want {
					t.Fatalf("inheritedArrayEvalForFuncArg(%q, %d) = %v, want %v", tt.name, i, got, want)
				}
			}
		})
	}
}

func TestInheritedArrayMetadataDoesNotExpand(t *testing.T) {
	cases := []struct {
		name     string
		argIndex int
	}{
		{name: "SUM", argIndex: 0},
		{name: "SUM", argIndex: 1},
		{name: "FILTER", argIndex: 0},
		{name: "FILTER", argIndex: 1},
		{name: "XLOOKUP", argIndex: 0},
		{name: "XLOOKUP", argIndex: 1},
		{name: "ACOSH", argIndex: 0},
		{name: "ACOT", argIndex: 0},
		{name: "ACOTH", argIndex: 0},
		{name: "ASINH", argIndex: 0},
		{name: "ODD", argIndex: 0},
		{name: "ISEVEN", argIndex: 0},
		{name: "ISODD", argIndex: 0},
	}

	for _, tt := range cases {
		t.Run(tt.name, func(t *testing.T) {
			if got := inheritedArrayEvalForFuncArg(tt.name, tt.argIndex); got {
				t.Fatalf("inheritedArrayEvalForFuncArg(%q, %d) = true, want false", tt.name, tt.argIndex)
			}
		})
	}
}

func TestRegisterClearsStaleMetadataOnOverride(t *testing.T) {
	const name = "TEST.REGISTER.CLEAR"

	RegisterWithMeta(name, NoCtx(func(args []Value) (Value, error) {
		return NumberVal(1), nil
	}), FuncMeta{
		Kind:               FnKindScalarLifted,
		InheritedArrayArgs: map[int]bool{0: true},
	})

	if got := inheritedArrayEvalForFuncArg(name, 0); !got {
		t.Fatalf("expected inherited array metadata for %s before override", name)
	}

	Register(name, NoCtx(func(args []Value) (Value, error) {
		return NumberVal(2), nil
	}))

	if _, ok := funcMetaForName(name); ok {
		t.Fatalf("expected metadata for %s to be cleared after plain Register override", name)
	}
	if got := inheritedArrayEvalForFuncArg(name, 0); got {
		t.Fatalf("inheritedArrayEvalForFuncArg(%q, 0) = true, want false after override", name)
	}
}
