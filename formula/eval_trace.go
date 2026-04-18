package formula

// EvalTracer receives intermediate values during formula evaluation.
// The VM checks ctx.Tracer != nil before calling — zero cost when unused.
type EvalTracer interface {
	// OnBinaryOp is called after a binaryArith or binaryCompare produces a result.
	OnBinaryOp(step int, op OpCode, a, b, result Value)

	// OnCallFunc is called after a function call returns.
	OnCallFunc(step int, funcName string, args []Value, result Value)
}
