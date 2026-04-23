package formula

// ArgLoadMode describes how a function argument should be loaded by the
// contract system. Only a subset is used in the first migration slice; the
// rest exist so later Part 2 slices do not need to rename the contract layer.
type ArgLoadMode uint8

const (
	ArgLoadScalar ArgLoadMode = iota
	ArgLoadArray
	ArgLoadRef
	ArgLoadDirectRange
	ArgLoadPassthrough
)

// ArgAdaptMode describes the runtime scalarization/adaptation policy for one
// loaded argument.
type ArgAdaptMode uint8

const (
	ArgAdaptLegacyIntersectRef ArgAdaptMode = iota
	ArgAdaptExplicitIntersect
	ArgAdaptTopLeftAnonymousArray
	ArgAdaptBroadcast
	ArgAdaptPassThrough
	ArgAdaptRequireScalar
)

// ReturnMode classifies the runtime result family a function can produce.
// Most migrated functions still return the legacy Value shape directly.
type ReturnMode uint8

const (
	ReturnModePassThrough ReturnMode = iota
	ReturnModeScalar
	ReturnModeArray
	ReturnModeRef
)

// EvalFunc is the internal contract-evaluator hook used by migrated v2
// families. It consumes EvalValue arguments while the public Value-based Func
// signature remains intact for non-migrated callers.
type EvalFunc func(args []EvalValue, ctx *EvalContext) (EvalValue, error)

// ArgSpec defines one positional argument contract.
type ArgSpec struct {
	Load  ArgLoadMode
	Adapt ArgAdaptMode
}

// FuncSpec is the contract-system registration record described in the v2
// architecture plan. The first migration slice uses it in shadow mode while
// legacy registry maps remain available as fallback for non-migrated families.
type FuncSpec struct {
	Kind   FnKind
	Args   []ArgSpec
	VarArg func(i int) ArgSpec
	Return ReturnMode
	Eval   EvalFunc
}

var funcSpecByName = map[string]FuncSpec{}

// RegisterWithSpec registers a function with a contract spec only.
func RegisterWithSpec(name string, fn Func, spec FuncSpec) {
	Register(name, fn)
	funcSpecByName[normalizeFuncName(name)] = cloneFuncSpec(spec)
}

// RegisterWithMetaAndSpec keeps legacy metadata while also attaching a
// FuncSpec. This supports shadow-mode migration slices where compiler/eval
// logic prefers the spec but the old metadata remains available as fallback.
func RegisterWithMetaAndSpec(name string, fn Func, meta FuncMeta, spec FuncSpec) {
	Register(name, fn)
	funcMetaByName[normalizeFuncName(name)] = cloneFuncMeta(meta)
	funcSpecByName[normalizeFuncName(name)] = cloneFuncSpec(spec)
}

// RegisterScalarLiftedUnarySpec registers a unary scalar-lifted function on
// the new contract path. Direct references still use legacy implicit
// intersection in scalar cells; anonymous arrays keep broadcast semantics.
func RegisterScalarLiftedUnarySpec(name string, fn Func) {
	RegisterWithSpec(name, fn, FuncSpec{
		Kind: FnKindScalarLifted,
		Args: []ArgSpec{{
			Load:  ArgLoadPassthrough,
			Adapt: ArgAdaptLegacyIntersectRef,
		}},
		Return: ReturnModePassThrough,
	})
}

func cloneFuncSpec(spec FuncSpec) FuncSpec {
	if len(spec.Args) == 0 {
		return spec
	}
	cloned := make([]ArgSpec, len(spec.Args))
	copy(cloned, spec.Args)
	spec.Args = cloned
	return spec
}

func funcSpecForName(name string) (FuncSpec, bool) {
	spec, ok := funcSpecByName[normalizeFuncName(name)]
	if !ok {
		return FuncSpec{}, false
	}
	return cloneFuncSpec(spec), true
}

func functionUsesElementwiseContract(name string) bool {
	if spec, ok := funcSpecForName(name); ok {
		return spec.Kind == FnKindScalarLifted
	}
	return elementWiseCallFuncs[normalizeFuncName(name)]
}

func functionNeedsLegacyElementwisePreIntersect(name string) bool {
	if _, ok := funcSpecForName(name); ok {
		return false
	}
	return elementWiseCallFuncs[normalizeFuncName(name)]
}

func callFuncWithSpec(name string, fn Func, spec FuncSpec, args []Value, ctx *EvalContext) (Value, error) {
	if spec.Eval != nil {
		result, err := spec.Eval(loadEvalFuncArgs(spec, args, ctx), ctx)
		if err != nil {
			return Value{}, err
		}
		return EvalValueToValue(result), nil
	}
	adapted := adaptFuncArgs(spec, args, ctx)
	switch spec.Kind {
	case FnKindScalarLifted:
		if hasArrayArg(adapted) {
			return callElementWise(adapted, ctx, fn)
		}
	}
	return fn(adapted, ctx)
}

func adaptFuncArgs(spec FuncSpec, args []Value, ctx *EvalContext) []Value {
	if len(args) == 0 {
		return args
	}
	out := make([]Value, len(args))
	copy(out, args)
	for i := range out {
		argSpec, ok := funcArgSpec(spec, i)
		if !ok {
			continue
		}
		out[i] = adaptFuncArg(argSpec, out[i], ctx)
	}
	return out
}

func loadEvalFuncArgs(spec FuncSpec, args []Value, ctx *EvalContext) []EvalValue {
	if len(args) == 0 {
		return nil
	}
	out := make([]EvalValue, len(args))
	for i := range args {
		argSpec, ok := funcArgSpec(spec, i)
		if !ok {
			argSpec = ArgSpec{Load: ArgLoadPassthrough, Adapt: ArgAdaptPassThrough}
		}
		out[i] = loadEvalFuncArg(argSpec, args[i], ctx)
	}
	return out
}

func funcArgSpec(spec FuncSpec, index int) (ArgSpec, bool) {
	if index >= 0 && index < len(spec.Args) {
		return spec.Args[index], true
	}
	if spec.VarArg != nil {
		return spec.VarArg(index), true
	}
	return ArgSpec{}, false
}

func loadEvalFuncArg(spec ArgSpec, arg Value, ctx *EvalContext) EvalValue {
	arg = adaptFuncArg(spec, arg, ctx)
	var resolver CellResolver
	if ctx != nil {
		resolver = ctx.Resolver
	}
	// Later migration slices will tighten this by load mode. For the direct
	// range slice, every mode shares the same EvalValue envelope.
	return valueToEvalValueWithResolver(arg, resolver)
}

func adaptFuncArg(spec ArgSpec, arg Value, ctx *EvalContext) Value {
	switch spec.Adapt {
	case ArgAdaptLegacyIntersectRef:
		if ctx != nil && !ctx.IsArrayFormula && !ctx.InheritedArray &&
			arg.Type == ValueArray && arg.RangeOrigin != nil {
			return implicitIntersect(arg, ctx)
		}
	case ArgAdaptExplicitIntersect:
		if ctx != nil && !ctx.IsArrayFormula {
			return explicitIntersect(arg, ctx)
		}
	case ArgAdaptTopLeftAnonymousArray:
		if arg.Type == ValueArray && arg.RangeOrigin == nil {
			return topLeftAnonymousArray(arg)
		}
	case ArgAdaptRequireScalar:
		if arg.Type == ValueArray {
			return ErrorVal(ErrValVALUE)
		}
	}
	return arg
}

func topLeftAnonymousArray(v Value) Value {
	if v.Type != ValueArray {
		return v
	}
	if len(v.Array) == 0 || len(v.Array[0]) == 0 {
		return ErrorVal(ErrValVALUE)
	}
	return v.Array[0][0]
}
