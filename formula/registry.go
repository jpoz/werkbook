package formula

import (
	"fmt"
	"strings"
)

// Func is the standard signature for all registered formula functions.
type Func func(args []Value, ctx *EvalContext) (Value, error)

// FnKind classifies a registered function for metadata-driven compiler logic.
type FnKind uint8

const (
	FnKindUnknown FnKind = iota
	FnKindScalarLifted
	FnKindReduction
	FnKindArrayNative
	FnKindLookup
	FnKindStateful
)

// FuncMeta stores registration-time metadata for a formula function.
type FuncMeta struct {
	Kind               FnKind
	InheritedArrayArgs map[int]bool
}

var (
	registry       = map[string]Func{}
	funcMetaByName = map[string]FuncMeta{}
	nameToID       = map[string]int{}
	idToName       []string
)

// Register adds a formula function to the global registry.
// It is safe to call from init(). Duplicate names overwrite silently,
// allowing external packages (e.g. werkbook-pro) to override or extend.
func Register(name string, fn Func) {
	upper := strings.ToUpper(name)
	if _, exists := nameToID[upper]; !exists {
		id := len(idToName)
		idToName = append(idToName, upper)
		nameToID[upper] = id
	}
	delete(funcMetaByName, upper)
	registry[upper] = fn
}

// RegisterWithMeta adds a formula function and stores registration metadata
// keyed by the function name.
func RegisterWithMeta(name string, fn Func, meta FuncMeta) {
	Register(name, fn)
	funcMetaByName[strings.ToUpper(name)] = cloneFuncMeta(meta)
}

// RegisterScalarLifted registers a scalar function that should inherit array
// context for its first argument when the compiler is already in an inherited
// array evaluation path.
func RegisterScalarLifted(name string, fn Func) {
	RegisterWithMeta(name, fn, FuncMeta{
		Kind:               FnKindScalarLifted,
		InheritedArrayArgs: map[int]bool{0: true},
	})
}

// LookupFunc returns the function ID for use by the compiler.
// Returns -1 if the function is not registered.
func LookupFunc(name string) int {
	id, ok := nameToID[strings.ToUpper(name)]
	if !ok {
		return -1
	}
	return id
}

// CallFunc dispatches a function call by ID at eval time.
func CallFunc(funcID int, args []Value, ctx *EvalContext) (Value, error) {
	if funcID < 0 || funcID >= len(idToName) {
		return Value{}, fmt.Errorf("unknown function ID %d", funcID)
	}
	name := idToName[funcID]
	fn := registry[name]
	if fn == nil {
		return Value{}, fmt.Errorf("unimplemented function: %s", name)
	}
	if elementWiseCallFuncs[name] && hasArrayArg(args) {
		return callElementWise(args, ctx, fn)
	}
	return fn(args, ctx)
}

// RegisteredFunctions returns the names of all registered functions.
func RegisteredFunctions() []string {
	out := make([]string, len(idToName))
	copy(out, idToName)
	return out
}

func cloneFuncMeta(meta FuncMeta) FuncMeta {
	if len(meta.InheritedArrayArgs) == 0 {
		return meta
	}
	cloned := make(map[int]bool, len(meta.InheritedArrayArgs))
	for idx, ok := range meta.InheritedArrayArgs {
		cloned[idx] = ok
	}
	meta.InheritedArrayArgs = cloned
	return meta
}

func funcMetaForName(name string) (FuncMeta, bool) {
	meta, ok := funcMetaByName[strings.ToUpper(name)]
	if !ok {
		return FuncMeta{}, false
	}
	return cloneFuncMeta(meta), true
}

// NoCtx wraps a function that doesn't need EvalContext into a Func.
func NoCtx(fn func([]Value) (Value, error)) Func {
	return func(args []Value, _ *EvalContext) (Value, error) {
		return fn(args)
	}
}

// elementWiseCallFuncs lists scalar functions that should broadcast over array
// arguments when they are evaluated in array context. Aggregate, lookup, and
// array-returning functions are intentionally excluded.
var elementWiseCallFuncs = map[string]bool{
	// Logic and info.
	"ERROR.TYPE": true,
	"IFERROR":    true,
	"IFNA":       true,
	"ISBLANK":    true,
	"ISERR":      true,
	"ISERROR":    true,
	"ISEVEN":     true,
	"ISLOGICAL":  true,
	"ISNA":       true,
	"ISNONTEXT":  true,
	"ISNUMBER":   true,
	"ISODD":      true,
	"ISTEXT":     true,
	"N":          true,
	"NOT":        true,

	// Text.
	"CHAR":        true,
	"CLEAN":       true,
	"CODE":        true,
	"DOLLAR":      true,
	"ENCODEURL":   true,
	"EXACT":       true,
	"FIND":        true,
	"FIXED":       true,
	"LEFT":        true,
	"LEN":         true,
	"LOWER":       true,
	"MID":         true,
	"NUMBERVALUE": true,
	"PROPER":      true,
	"REPLACE":     true,
	"REPT":        true,
	"RIGHT":       true,
	"ROMAN":       true,
	"SEARCH":      true,
	"SUBSTITUTE":  true,
	"T":           true,
	"TEXT":        true,
	"TEXTAFTER":   true,
	"TEXTBEFORE":  true,
	"TRIM":        true,
	"UNICHAR":     true,
	"UNICODE":     true,
	"UPPER":       true,
	"VALUE":       true,

	// Date and time.
	"DATE":       true,
	"DATEDIF":    true,
	"DATEVALUE":  true,
	"DAY":        true,
	"DAYS":       true,
	"DAYS360":    true,
	"EDATE":      true,
	"EOMONTH":    true,
	"HOUR":       true,
	"ISOWEEKNUM": true,
	"MINUTE":     true,
	"MONTH":      true,
	"SECOND":     true,
	"TIME":       true,
	"TIMEVALUE":  true,
	"WEEKDAY":    true,
	"WEEKNUM":    true,
	"YEAR":       true,
	"YEARFRAC":   true,

	// Scalar math.
	"ABS":             true,
	"ACOS":            true,
	"ACOSH":           true,
	"ACOT":            true,
	"ACOTH":           true,
	"ARABIC":          true,
	"ASIN":            true,
	"ASINH":           true,
	"ATAN":            true,
	"ATAN2":           true,
	"ATANH":           true,
	"BASE":            true,
	"BITAND":          true,
	"BITLSHIFT":       true,
	"BITOR":           true,
	"BITRSHIFT":       true,
	"BITXOR":          true,
	"CEILING":         true,
	"CEILING.MATH":    true,
	"CEILING.PRECISE": true,
	"COMBIN":          true,
	"COMBINA":         true,
	"COS":             true,
	"COSH":            true,
	"COT":             true,
	"COTH":            true,
	"CSC":             true,
	"CSCH":            true,
	"DECIMAL":         true,
	"DEGREES":         true,
	"ERF":             true,
	"ERF.PRECISE":     true,
	"ERFC":            true,
	"ERFC.PRECISE":    true,
	"EVEN":            true,
	"EXP":             true,
	"FACT":            true,
	"FACTDOUBLE":      true,
	"FLOOR":           true,
	"FLOOR.MATH":      true,
	"FLOOR.PRECISE":   true,
	"GAMMA":           true,
	"INT":             true,
	"ISO.CEILING":     true,
	"LN":              true,
	"LOG":             true,
	"LOG10":           true,
	"MOD":             true,
	"MROUND":          true,
	"ODD":             true,
	"PERMUT":          true,
	"POWER":           true,
	"QUOTIENT":        true,
	"RADIANS":         true,
	"ROUND":           true,
	"ROUNDDOWN":       true,
	"ROUNDUP":         true,
	"SEC":             true,
	"SECH":            true,
	"SIGN":            true,
	"SIN":             true,
	"SINH":            true,
	"SQRT":            true,
	"SQRTPI":          true,
	"TAN":             true,
	"TANH":            true,
	"TRUNC":           true,

	// Scalar statistics and distributions.
	"BINOM.DIST":       true,
	"BINOM.DIST.RANGE": true,
	"BINOM.INV":        true,
	"BETA.DIST":        true,
	"BETA.INV":         true,
	"CHISQ.DIST":       true,
	"CHISQ.DIST.RT":    true,
	"CHISQ.INV":        true,
	"CHISQ.INV.RT":     true,
	"CONFIDENCE.NORM":  true,
	"CONFIDENCE.T":     true,
	"EXPON.DIST":       true,
	"F.DIST":           true,
	"F.DIST.RT":        true,
	"F.INV":            true,
	"F.INV.RT":         true,
	"FISHER":           true,
	"FISHERINV":        true,
	"GAMMALN":          true,
	"GAMMALN.PRECISE":  true,
	"GAMMA.DIST":       true,
	"GAMMA.INV":        true,
	"GAUSS":            true,
	"HYPGEOM.DIST":     true,
	"LOGNORM.DIST":     true,
	"LOGNORM.INV":      true,
	"NEGBINOM.DIST":    true,
	"NORM.DIST":        true,
	"NORM.INV":         true,
	"NORM.S.DIST":      true,
	"NORM.S.INV":       true,
	"PERMUTATIONA":     true,
	"PHI":              true,
	"POISSON.DIST":     true,
	"STANDARDIZE":      true,
	"T.DIST":           true,
	"T.DIST.2T":        true,
	"T.DIST.RT":        true,
	"T.INV":            true,
	"T.INV.2T":         true,
	"WEIBULL.DIST":     true,
}

func hasArrayArg(args []Value) bool {
	for _, arg := range args {
		if arg.Type == ValueArray {
			return true
		}
	}
	return false
}

func arrayDimsAll(args []Value) (rows, cols int) {
	rows, cols = 1, 1
	for _, arg := range args {
		if arg.Type != ValueArray {
			continue
		}
		r, c := len(arg.Array), 0
		if r > 0 {
			c = len(arg.Array[0])
		}
		if r > rows {
			rows = r
		}
		if c > cols {
			cols = c
		}
	}
	return rows, cols
}

func callElementWise(args []Value, ctx *EvalContext, fn Func) (Value, error) {
	rows, cols := arrayDimsAll(args)
	// Precompute per-arg bounds to avoid O(rows) scans per element access.
	type argBounds struct{ rows, cols int }
	bounds := make([]argBounds, len(args))
	for k, arg := range args {
		r, c := arrayOpBoundsOrScalar(arg)
		bounds[k] = argBounds{r, c}
	}
	result := make([][]Value, rows)
	scalarArgs := make([]Value, len(args))
	for i := 0; i < rows; i++ {
		result[i] = make([]Value, cols)
		for j := 0; j < cols; j++ {
			for k, arg := range args {
				scalarArgs[k] = arrayElementDirect(arg, bounds[k].rows, bounds[k].cols, i, j)
			}
			cell, err := fn(scalarArgs, ctx)
			if err != nil {
				return Value{}, err
			}
			result[i][j] = cell
		}
	}
	return Value{Type: ValueArray, Array: result}, nil
}

// arrayForcingFuncs lists functions that evaluate their arguments in array
// context, suppressing implicit intersection. In legacy (non-dynamic-array)
// these functions treat expressions like range*range element-wise
// even when the formula is not entered as CSE (Ctrl+Shift+Enter).
var arrayForcingFuncs = map[string]bool{
	"SUMPRODUCT": true,
	"MMULT":      true,
	"TREND":      true,
	"GROWTH":     true,
	"LINEST":     true,
	"LOGEST":     true,
	"FREQUENCY":  true,
	"TRANSPOSE":  true,
	// Functions that accept range arguments and must not have them
	// implicitly intersected to a single cell.
	"SUMIF":      true,
	"SUMIFS":     true,
	"COUNTIF":    true,
	"COUNTIFS":   true,
	"AVERAGEIF":  true,
	"AVERAGEIFS": true,
	"MAXIFS":     true,
	"MINIFS":     true,
	"MATCH":      true,
	"INDEX":      true,
	"LOOKUP":     true,
	"VLOOKUP":    true,
	"HLOOKUP":    true,
}

// FuncArgEvalMode describes how the compiler should evaluate an argument for a
// particular function in a legacy non-array formula context.
type FuncArgEvalMode int

const (
	FuncArgEvalDefault FuncArgEvalMode = iota
	FuncArgEvalDirectRange
	FuncArgEvalArray
)

// directRangeAllArgFuncs preserve plain direct range references like A:A or
// 1:1 for every argument position, but still allow expressions like A:A*B:B to
// follow legacy implicit-intersection behavior unless the function is fully
// array-forcing.
var directRangeAllArgFuncs = map[string]bool{
	"AVERAGE":   true,
	"AVERAGEA":  true,
	"AVEDEV":    true,
	"COLUMNS":   true,
	"COUNT":     true,
	"COUNTA":    true,
	"DEVSQ":     true,
	"GEOMEAN":   true,
	"HARMEAN":   true,
	"KURT":      true,
	"MAX":       true,
	"MAXA":      true,
	"MEDIAN":    true,
	"MIN":       true,
	"MINA":      true,
	"MODE":      true,
	"MODE.MULT": true,
	"MODE.SNGL": true,
	"PRODUCT":   true,
	"ROWS":      true,
	"SKEW":      true,
	"SKEW.P":    true,
	"STDEV":     true,
	"STDEV.P":   true,
	"STDEV.S":   true,
	"STDEVA":    true,
	"STDEVP":    true,
	"STDEVPA":   true,
	"SUM":       true,
	"SUMSQ":     true,
	"VAR":       true,
	"VAR.P":     true,
	"VAR.S":     true,
	"VARA":      true,
	"VARP":      true,
	"VARPA":     true,
}

// directRangeArgFuncs preserve direct range references only for the argument
// positions that are range-like for a given function. This avoids accidentally
// suppressing implicit intersection for scalar parameters such as the first
// STANDARDIZE argument.
var directRangeArgFuncs = map[string]map[int]bool{
	"COUNTBLANK":      {0: true},
	"CORREL":          {0: true, 1: true},
	"COVAR":           {0: true, 1: true},
	"COVARIANCE.P":    {0: true, 1: true},
	"COVARIANCE.S":    {0: true, 1: true},
	"FORECAST":        {1: true, 2: true},
	"FORECAST.LINEAR": {1: true, 2: true},
	"INTERCEPT":       {0: true, 1: true},
	"LARGE":           {0: true},
	"PEARSON":         {0: true, 1: true},
	"PERCENTILE":      {0: true},
	"PERCENTILE.EXC":  {0: true},
	"PERCENTILE.INC":  {0: true},
	"PERCENTRANK":     {0: true},
	"PERCENTRANK.EXC": {0: true},
	"PERCENTRANK.INC": {0: true},
	"QUARTILE":        {0: true},
	"QUARTILE.EXC":    {0: true},
	"QUARTILE.INC":    {0: true},
	"RANK":            {1: true},
	"RANK.AVG":        {1: true},
	"RANK.EQ":         {1: true},
	"RSQ":             {0: true, 1: true},
	"SLOPE":           {0: true, 1: true},
	"SMALL":           {0: true},
	"STEYX":           {0: true, 1: true},
	"TRIMMEAN":        {0: true},
	"XLOOKUP":         {1: true, 2: true},
	"XMATCH":          {1: true},
}

// directRangeArgStartFuncs preserve direct ranges from the given argument
// index onward. This is used for functions that accept a trailing list of
// references after one or more scalar control arguments.
var directRangeArgStartFuncs = map[string]int{
	"AGGREGATE": 2,
	"SUBTOTAL":  1,
}

// arrayArgFuncs evaluate the listed argument positions in array context,
// suppressing implicit intersection for the whole argument expression. This is
// required for functions like FILTER whose include argument is commonly a
// boolean range expression rather than a plain range reference.
var arrayArgFuncs = map[string]map[int]bool{
	"FILTER": {0: true, 1: true},
}

// inheritedArrayArgFuncs is a temporary compatibility fallback for functions
// that have not yet been migrated to registration-time metadata.
var inheritedArrayArgFuncs = map[string]map[int]bool{
	"IFERROR": {0: true, 1: true},
	"IFNA":    {0: true, 1: true},

	// Scalar math/info functions that lift to arrays via LiftUnary.
	// They must inherit array context so that expressions like
	// SUMPRODUCT(ABS(range1-range2)) evaluate the subtraction
	// element-wise instead of implicitly intersecting.
	"ABS":   {0: true},
	"INT":   {0: true},
	"SIGN":  {0: true},
	"SQRT":  {0: true},
	"LN":    {0: true},
	"LOG10": {0: true},
	"EXP":   {0: true},
	"FACT":  {0: true},
	"SIN":   {0: true},
	"COS":   {0: true},
	"TAN":   {0: true},
	"ASIN":  {0: true},
	"ACOS":  {0: true},
	"ATAN":  {0: true},
	"SINH":  {0: true},
	"COSH":  {0: true},
	"TANH":  {0: true},

	// Date/text functions that lift to arrays via LiftUnary.
	"DATEVALUE": {0: true},

	// Info functions that also lift to arrays.
	"ISNUMBER": {0: true},
	"ISTEXT":   {0: true},
	"ISBLANK":  {0: true},
	"ISERROR":  {0: true},
	"ISERR":    {0: true},
	"ISNA":     {0: true},
	"NOT":      {0: true},
	"N":        {0: true},
	"TYPE":     {0: true},
}

// arrayFirstArgFuncs evaluate the first argument in array context because it is
// semantically an array input.
var arrayFirstArgFuncs = map[string]bool{
	"ARRAYTOTEXT": true,
	"CHOOSECOLS":  true,
	"CHOOSEROWS":  true,
	"DROP":        true,
	"EXPAND":      true,
	"SORT":        true,
	"TAKE":        true,
	"TOCOL":       true,
	"TOROW":       true,
	"UNIQUE":      true,
	"WRAPCOLS":    true,
	"WRAPROWS":    true,
}

// arrayAllArgFuncs evaluate every argument in array context.
var arrayAllArgFuncs = map[string]bool{
	"HSTACK": true,
	"VSTACK": true,
}

// IsArrayFunc reports whether the named function forces array evaluation of
// its arguments. The compiler uses this to emit OpEnterArrayCtx / OpLeaveArrayCtx
// around the function's argument expressions.
func IsArrayFunc(name string) bool {
	return arrayForcingFuncs[strings.ToUpper(name)]
}

// ArgEvalModeForFuncArg reports how the compiler should evaluate the given
// argument position for a function in legacy non-array formula contexts.
func ArgEvalModeForFuncArg(name string, argIndex int) FuncArgEvalMode {
	upper := strings.ToUpper(name)
	if arrayAllArgFuncs[upper] {
		return FuncArgEvalArray
	}
	if arrayFirstArgFuncs[upper] && argIndex == 0 {
		return FuncArgEvalArray
	}
	if positions, ok := arrayArgFuncs[upper]; ok && positions[argIndex] {
		return FuncArgEvalArray
	}
	if upper == "SORTBY" && (argIndex == 0 || argIndex%2 == 1) {
		return FuncArgEvalArray
	}
	if directRangeAllArgFuncs[upper] {
		return FuncArgEvalDirectRange
	}
	if start, ok := directRangeArgStartFuncs[upper]; ok && argIndex >= start {
		return FuncArgEvalDirectRange
	}
	if positions, ok := directRangeArgFuncs[upper]; ok && positions[argIndex] {
		return FuncArgEvalDirectRange
	}
	return FuncArgEvalDefault
}

func inheritedArrayEvalForFuncArg(name string, argIndex int) bool {
	if meta, ok := funcMetaForName(name); ok && meta.InheritedArrayArgs != nil {
		return meta.InheritedArrayArgs[argIndex]
	}
	positions, ok := inheritedArrayArgFuncs[strings.ToUpper(name)]
	return ok && positions[argIndex]
}
