package formula

// ArgLoadMode describes how a function argument should be loaded by the
// contract system.
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
	// ArgAdaptScalarizeAny combines LegacyIntersectRef (for range-backed
	// arrays) with TopLeftAnonymousArray (for anonymous arrays). Used by
	// lookup functions whose first argument is semantically a scalar but
	// which currently accept arrays from either provenance — XLOOKUP and
	// XMATCH in non-array context match Excel by intersecting the direct
	// range (row/col-aligned cell) or collapsing the inline array to its
	// top-left element.
	ArgAdaptScalarizeAny
)

// ReturnMode classifies the runtime result family a function can produce.
// Most functions still return the public Value shape directly.
type ReturnMode uint8

const (
	ReturnModePassThrough ReturnMode = iota
	ReturnModeScalar
	ReturnModeArray
	ReturnModeRef
)

// EvalFunc is the internal contract-evaluator hook. It consumes EvalValue
// arguments while the public Value-based Func signature remains intact for
// callers that still use it directly.
type EvalFunc func(args []EvalValue, ctx *EvalContext) (EvalValue, error)

// ArgSpec defines one positional argument contract.
type ArgSpec struct {
	Load         ArgLoadMode
	Adapt        ArgAdaptMode
	InheritArray bool
}

// FuncSpec is the contract-system registration record.
type FuncSpec struct {
	Kind      FnKind
	Args      []ArgSpec
	VarArg    func(i int) ArgSpec
	Return    ReturnMode
	ArrayLift bool
	Eval      EvalFunc
}

var funcSpecByName = map[string]FuncSpec{}

// RegisterWithSpec registers a function with a contract spec only.
func RegisterWithSpec(name string, fn Func, spec FuncSpec) {
	Register(name, fn)
	funcSpecByName[normalizeFuncName(name)] = cloneFuncSpec(spec)
}

// RegisterWithMetaAndSpec preserves the older registration API while storing
// the merged behavior in the function contract.
func RegisterWithMetaAndSpec(name string, fn Func, meta FuncMeta, spec FuncSpec) {
	RegisterWithSpec(name, fn, mergeMetaIntoSpec(meta, spec))
}

// RegisterScalarLiftedUnarySpec registers a unary scalar-lifted function on
// the new contract path. Direct references still use legacy implicit
// intersection in scalar cells; anonymous arrays keep broadcast semantics.
func RegisterScalarLiftedUnarySpec(name string, fn Func) {
	RegisterWithSpec(name, fn, scalarLiftedFuncSpec(1, true))
}

func scalarLiftedFuncSpec(argCount int, arrayLift bool, inheritedArgs ...int) FuncSpec {
	if argCount < 0 {
		argCount = 0
	}
	for _, idx := range inheritedArgs {
		if idx+1 > argCount {
			argCount = idx + 1
		}
	}
	args := make([]ArgSpec, argCount)
	adapt := ArgAdaptPassThrough
	if arrayLift {
		adapt = ArgAdaptLegacyIntersectRef
	}
	for i := range args {
		args[i] = ArgSpec{
			Load:  ArgLoadPassthrough,
			Adapt: adapt,
		}
	}
	for _, idx := range inheritedArgs {
		if idx >= 0 && idx < len(args) {
			args[idx].InheritArray = true
		}
	}
	return FuncSpec{
		Kind:      FnKindScalarLifted,
		Args:      args,
		Return:    ReturnModePassThrough,
		ArrayLift: arrayLift,
	}
}

func funcSpecFromMeta(name string, meta FuncMeta) FuncSpec {
	argCount := 0
	inherited := make([]int, 0, len(meta.InheritedArrayArgs))
	for idx, ok := range meta.InheritedArrayArgs {
		if idx+1 > argCount {
			argCount = idx + 1
		}
		if ok {
			inherited = append(inherited, idx)
		}
	}
	if meta.Kind == FnKindScalarLifted {
		if argCount == 0 {
			argCount = 1
		}
		return scalarLiftedFuncSpec(argCount, elementWiseCallFuncs[normalizeFuncName(name)], inherited...)
	}
	args := make([]ArgSpec, argCount)
	for _, idx := range inherited {
		args[idx].InheritArray = true
	}
	return FuncSpec{
		Kind:   meta.Kind,
		Args:   args,
		Return: ReturnModePassThrough,
	}
}

func mergeMetaIntoSpec(meta FuncMeta, spec FuncSpec) FuncSpec {
	if spec.Kind == FnKindUnknown {
		spec.Kind = meta.Kind
	}
	for idx, ok := range meta.InheritedArrayArgs {
		if !ok || idx < 0 {
			continue
		}
		for len(spec.Args) <= idx {
			spec.Args = append(spec.Args, ArgSpec{
				Load:  ArgLoadPassthrough,
				Adapt: ArgAdaptPassThrough,
			})
		}
		spec.Args[idx].InheritArray = true
	}
	return spec
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
		if spec.ArrayLift && hasArrayArg(adapted) {
			return callElementWise(adapted, ctx, fn)
		}
	case FnKindLookupArrayLift:
		// In array context, ArgAdaptScalarizeAny is a no-op on arg 0 so
		// the lookup_value array still needs to be fanned out per-element.
		// In scalar context it already collapsed, so this branch is skipped.
		if len(adapted) > 0 && adapted[0].Type == ValueArray {
			return liftLookupArg0(fn, adapted, ctx)
		}
	}
	return fn(adapted, ctx)
}

// liftLookupArg0 broadcasts the lookup_value argument over its elements,
// calling the underlying lookup function once per element and assembling an
// anonymous ValueArray of the results. Used for VLOOKUP/HLOOKUP/MATCH/LOOKUP/
// XLOOKUP/XMATCH when the first argument arrives as an array in array context.
func liftLookupArg0(fn Func, args []Value, ctx *EvalContext) (Value, error) {
	lookup := args[0]
	rows, cols := arrayOpBounds(lookup)
	if rows == 0 || cols == 0 {
		return ErrorVal(ErrValVALUE), nil
	}
	result := make([][]Value, rows)
	scalarArgs := make([]Value, len(args))
	copy(scalarArgs, args)
	for i := 0; i < rows; i++ {
		result[i] = make([]Value, cols)
		for j := 0; j < cols; j++ {
			scalarArgs[0] = arrayElementDirect(lookup, rows, cols, i, j)
			cell, err := fn(scalarArgs, ctx)
			if err != nil {
				return Value{}, err
			}
			result[i][j] = cell
		}
	}
	if rows == 1 && cols == 1 {
		return result[0][0], nil
	}
	return Value{Type: ValueArray, Array: result}, nil
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
	case ArgAdaptScalarizeAny:
		if ctx != nil && !ctx.IsArrayFormula && !ctx.InheritedArray &&
			arg.Type == ValueArray {
			if arg.RangeOrigin != nil {
				return implicitIntersect(arg, ctx)
			}
			return topLeftAnonymousArray(arg)
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
