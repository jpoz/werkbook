package formula

import (
	"fmt"
	"strings"
)

// Function IDs are assigned dynamically by the registry (registry.go).
// The compiler uses LookupFunc to resolve names to IDs at compile time,
// and the VM uses CallFunc to dispatch by ID at eval time.

// Compile walks the AST rooted at node and emits bytecode.
func Compile(source string, node Node) (*CompiledFormula, error) {
	compiled, err := compileWithMode(source, node, false)
	if err != nil {
		return nil, err
	}
	if formulaNeedsTopLevelArray(node) {
		topLevelArray, err := compileWithMode(source, node, true)
		if err != nil {
			return nil, err
		}
		compiled.TopLevelArray = topLevelArray
	}
	return compiled, nil
}

func compileWithMode(source string, node Node, topLevelArrayCtx bool) (*CompiledFormula, error) {
	c := &compiler{
		numIdx: make(map[float64]uint32),
		strIdx: make(map[string]uint32),
		refIdx: make(map[CellAddr]uint32),
		rngIdx: make(map[RangeAddr]uint32),
	}
	if topLevelArrayCtx {
		c.dynamicArrayDepth = 1
		c.emit(OpEnterArrayCtx, 0)
		if err := c.compileNodeCtx(node, true); err != nil {
			return nil, err
		}
		c.emit(OpLeaveArrayCtx, 0)
	} else {
		if err := c.compileNode(node); err != nil {
			return nil, err
		}
	}
	return &CompiledFormula{
		Source:      source,
		Code:        c.code,
		Consts:      c.consts,
		Refs:        c.refs,
		Ranges:      c.ranges,
		SubFormulas: c.subFormulas,
	}, nil
}

type compiler struct {
	code        []Instruction
	consts      []Value
	refs        []CellAddr
	ranges      []RangeAddr
	subFormulas []*CompiledFormula

	numIdx map[float64]uint32
	strIdx map[string]uint32
	refIdx map[CellAddr]uint32
	rngIdx map[RangeAddr]uint32

	// dynamicArrayDepth tracks nesting inside dynamic-array-native functions
	// (FILTER, SORT, UNIQUE, HSTACK, MAP, etc.). Inside those, IFERROR/IFNA
	// lift over arrays like any element-wise function. Outside (top level or
	// under legacy array-forcing like SUMPRODUCT), IFERROR/IFNA follow Excel's
	// scalar implicit-intersection semantics.
	dynamicArrayDepth int

	// legacyArrayDepth tracks compile-time nesting inside any array-forcing
	// caller (SUMPRODUCT, SUMIF family, INDEX, array-evaluated arg slots,
	// CSE probes, DynamicRangeRef). It is incremented only at "real" array
	// ctx entries — NOT at the suspend-for-IF pair that emits a
	// Leave/Enter pair around a scalar caller's arguments. OpCall uses this
	// counter to set the inheritedArrayCtx flag so element-wise functions
	// keep broadcasting when they appear inside SUMPRODUCT(IF(…), …) even
	// though runtime arrayCtxDepth drops to 0 inside IF's arms.
	legacyArrayDepth int
}

func (c *compiler) emit(op OpCode, operand uint32) {
	c.code = append(c.code, Instruction{Op: op, Operand: operand})
}

// enterLegacyArrayCtx emits OpEnterArrayCtx and bumps the compile-time
// legacyArrayDepth counter. Use this for "real" array-forcing entries
// (SUMPRODUCT and friends, array-consuming arg slots, CSE probes) so nested
// OpCall instructions know they're inside a legacy array context even when a
// sibling IF suspends the runtime array-ctx counter for its own arm.
func (c *compiler) enterLegacyArrayCtx() {
	c.emit(OpEnterArrayCtx, 0)
	c.legacyArrayDepth++
}

// leaveLegacyArrayCtx emits OpLeaveArrayCtx and decrements the counter.
// Must be paired one-to-one with enterLegacyArrayCtx — raw Leave/Enter pairs
// used only for the IF-suspend dance stay off this accounting because they
// cancel at the outer level.
func (c *compiler) leaveLegacyArrayCtx() {
	c.emit(OpLeaveArrayCtx, 0)
	c.legacyArrayDepth--
}

// addConst returns the index for a constant, deduplicating numbers and strings.
func (c *compiler) addNumConst(v float64) uint32 {
	if idx, ok := c.numIdx[v]; ok {
		return idx
	}
	idx := uint32(len(c.consts))
	c.consts = append(c.consts, NumberVal(v))
	c.numIdx[v] = idx
	return idx
}

func (c *compiler) addStrConst(s string) uint32 {
	if idx, ok := c.strIdx[s]; ok {
		return idx
	}
	idx := uint32(len(c.consts))
	c.consts = append(c.consts, StringVal(s))
	c.strIdx[s] = idx
	return idx
}

func (c *compiler) addRef(addr CellAddr) uint32 {
	if idx, ok := c.refIdx[addr]; ok {
		return idx
	}
	idx := uint32(len(c.refs))
	c.refs = append(c.refs, addr)
	c.refIdx[addr] = idx
	return idx
}

func (c *compiler) addRange(addr RangeAddr) uint32 {
	if idx, ok := c.rngIdx[addr]; ok {
		return idx
	}
	idx := uint32(len(c.ranges))
	c.ranges = append(c.ranges, addr)
	c.rngIdx[addr] = idx
	return idx
}

func (c *compiler) compileNode(node Node) error {
	return c.compileNodeCtx(node, false)
}

func (c *compiler) compileNodeCtx(node Node, inArrayCtx bool) error {
	switch n := node.(type) {
	case *NumberLit:
		idx := c.addNumConst(n.Value)
		c.emit(OpPushNum, idx)

	case *StringLit:
		idx := c.addStrConst(n.Value)
		c.emit(OpPushStr, idx)

	case *BoolLit:
		v := uint32(0)
		if n.Value {
			v = 1
		}
		c.emit(OpPushBool, v)

	case *EmptyArg:
		c.emit(OpPushEmpty, 0)

	case *ErrorLit:
		c.emit(OpPushError, uint32(errorCodeFromAST(n.Code)))

	case *CellRef:
		if n.DotNotation {
			// Dot-notation (Sheet1.A1) is a LibreOffice extension; returns #NAME?
			c.emit(OpPushError, uint32(ErrValNAME))
			return nil
		}
		if n.Name != "" {
			// Bare identifier that was not resolved as a defined name nor
			// consumed by LET/LAMBDA desugaring.
			c.emit(OpPushError, uint32(ErrValNAME))
			return nil
		}
		if n.Col < 1 || n.Col > maxCols {
			// Column outside valid range [A, XFD] (includes overflow-wrapped values)
			// that wasn't consumed by LET/LAMBDA desugaring — invalid ref.
			c.emit(OpPushError, uint32(ErrValNAME))
			return nil
		}
		addr := CellAddr{Sheet: n.Sheet, SheetEnd: n.SheetEnd, Col: n.Col, Row: n.Row}
		if n.SheetEnd != "" {
			// 3D reference (Sheet2:Sheet5!A1): treat as a range across sheets.
			rng := RangeAddr{
				Sheet: n.Sheet, SheetEnd: n.SheetEnd,
				FromCol: n.Col, FromRow: n.Row,
				ToCol: n.Col, ToRow: n.Row,
			}
			idx := c.addRange(rng)
			c.emit(OpLoad3DRange, idx)
		} else {
			idx := c.addRef(addr)
			c.emit(OpLoadCell, idx)
		}

	case *RangeRef:
		if n.From.DotNotation || n.To.DotNotation {
			// Dot-notation range (Sheet1.A1:Sheet1.A5) is a LibreOffice extension; returns #NAME?
			c.emit(OpPushError, uint32(ErrValNAME))
			return nil
		}
		if n.From.Col < 1 || n.From.Col > maxCols || n.To.Col < 1 || n.To.Col > maxCols {
			// Column outside valid range [A, XFD] (includes overflow-wrapped values)
			// that wasn't consumed by LET/LAMBDA desugaring — invalid ref.
			c.emit(OpPushError, uint32(ErrValNAME))
			return nil
		}
		sheet := n.From.Sheet
		sheetEnd := n.From.SheetEnd
		addr := RangeAddr{
			Sheet:    sheet,
			SheetEnd: sheetEnd,
			FromCol:  n.From.Col, FromRow: n.From.Row,
			ToCol: n.To.Col, ToRow: n.To.Row,
		}
		idx := c.addRange(addr)
		if sheetEnd != "" {
			c.emit(OpLoad3DRange, idx)
		} else {
			c.emit(OpLoadRange, idx)
		}

	case *IntersectRef:
		if !isReferenceNode(n.Left) || !isReferenceNode(n.Right) {
			c.emit(OpPushError, uint32(ErrValVALUE))
			return nil
		}
		if err := c.compileIntersectOperand(n.Left, inArrayCtx); err != nil {
			return err
		}
		if err := c.compileIntersectOperand(n.Right, inArrayCtx); err != nil {
			return err
		}
		c.emit(OpIntersect, 0)

	case *DynamicRangeRef:
		// Both endpoints must yield single-cell references at runtime.
		// CellRef operands compile to OpLoadCellRef so their address is
		// available without resolving the cell value; other ref-returning
		// expressions are compiled in array context so full-column/row
		// args (e.g. A:A inside INDEX(A:A,n)) don't implicit-intersect
		// before the function can use them as references.
		c.emit(OpEnterArrayCtx, 0)
		if err := c.compileIntersectOperand(n.From, true); err != nil {
			return err
		}
		if err := c.compileIntersectOperand(n.To, true); err != nil {
			return err
		}
		c.emit(OpLeaveArrayCtx, 0)
		c.emit(OpBuildRange, 0)

	case *UnionRef:
		if len(n.Areas) == 0 {
			c.emit(OpPushError, uint32(ErrValVALUE))
			return nil
		}
		for _, area := range n.Areas {
			if err := c.compileNodeCtx(area, inArrayCtx); err != nil {
				return err
			}
		}
		c.emit(OpUnion, uint32(len(n.Areas)))

	case *UnaryExpr:
		if err := c.compileNodeCtx(n.Operand, inArrayCtx); err != nil {
			return err
		}
		switch n.Op {
		case "-":
			c.emit(OpNeg, 0)
		case "+":
			// no-op
		case "@":
			c.emit(OpImplicitIntersect, 0)
		default:
			return fmt.Errorf("unknown unary operator %q", n.Op)
		}

	case *BinaryExpr:
		if err := c.compileNodeCtx(n.Left, inArrayCtx); err != nil {
			return err
		}
		if err := c.compileNodeCtx(n.Right, inArrayCtx); err != nil {
			return err
		}
		op, err := binaryOpCode(n.Op)
		if err != nil {
			return err
		}
		c.emit(op, 0)

	case *PostfixExpr:
		if err := c.compileNodeCtx(n.Operand, inArrayCtx); err != nil {
			return err
		}
		switch n.Op {
		case "%":
			c.emit(OpPercent, 0)
		default:
			return fmt.Errorf("unknown postfix operator %q", n.Op)
		}

	case *FuncCall:
		if err := c.compileFuncCall(n, inArrayCtx); err != nil {
			return err
		}

	case *ArrayLit:
		rows := len(n.Rows)
		cols := 0
		if rows > 0 {
			cols = len(n.Rows[0])
		}
		if rows > 65535 || cols > 65535 {
			return fmt.Errorf("array too large: %dx%d (max 65535x65535)", rows, cols)
		}
		for _, row := range n.Rows {
			for _, elem := range row {
				if err := c.compileNodeCtx(elem, inArrayCtx); err != nil {
					return err
				}
			}
		}
		operand := uint32(rows)<<16 | uint32(cols)
		c.emit(OpMakeArray, operand)

	case *MapExpr:
		// Compile array expressions — they push values onto the stack
		for _, arr := range n.Arrays {
			// Evaluate arrays in array context to prevent implicit intersection
			c.emit(OpEnterArrayCtx, 0)
			if err := c.compileNodeCtx(arr, true); err != nil {
				return err
			}
			c.emit(OpLeaveArrayCtx, 0)
		}
		// Compile the lambda body as a sub-formula
		subCompiler := &compiler{
			numIdx: make(map[float64]uint32),
			strIdx: make(map[string]uint32),
			refIdx: make(map[CellAddr]uint32),
			rngIdx: make(map[RangeAddr]uint32),
		}
		if err := subCompiler.compileNode(n.Body); err != nil {
			return err
		}
		subFormula := &CompiledFormula{
			Code:        subCompiler.code,
			Consts:      subCompiler.consts,
			Refs:        subCompiler.refs,
			Ranges:      subCompiler.ranges,
			SubFormulas: subCompiler.subFormulas,
		}
		subIdx := len(c.subFormulas)
		c.subFormulas = append(c.subFormulas, subFormula)
		// OpMap: subFormulaIdx << 8 | numArrays
		c.emit(OpMap, uint32(subIdx)<<8|uint32(len(n.Arrays)))

	case *ReduceExpr:
		// Push initial value
		if err := c.compileNodeCtx(n.InitialValue, inArrayCtx); err != nil {
			return err
		}
		// Push array in array context
		c.emit(OpEnterArrayCtx, 0)
		if err := c.compileNodeCtx(n.Array, true); err != nil {
			return err
		}
		c.emit(OpLeaveArrayCtx, 0)
		// Compile the lambda body as a sub-formula
		subCompiler := &compiler{
			numIdx: make(map[float64]uint32),
			strIdx: make(map[string]uint32),
			refIdx: make(map[CellAddr]uint32),
			rngIdx: make(map[RangeAddr]uint32),
		}
		if err := subCompiler.compileNode(n.Body); err != nil {
			return err
		}
		subFormula := &CompiledFormula{
			Code:        subCompiler.code,
			Consts:      subCompiler.consts,
			Refs:        subCompiler.refs,
			Ranges:      subCompiler.ranges,
			SubFormulas: subCompiler.subFormulas,
		}
		subIdx := len(c.subFormulas)
		c.subFormulas = append(c.subFormulas, subFormula)
		c.emit(OpReduce, uint32(subIdx))

	case *ScanExpr:
		// Push initial value
		if err := c.compileNodeCtx(n.InitialValue, inArrayCtx); err != nil {
			return err
		}
		// Push array in array context
		c.emit(OpEnterArrayCtx, 0)
		if err := c.compileNodeCtx(n.Array, true); err != nil {
			return err
		}
		c.emit(OpLeaveArrayCtx, 0)
		// Compile the lambda body as a sub-formula
		subCompiler := &compiler{
			numIdx: make(map[float64]uint32),
			strIdx: make(map[string]uint32),
			refIdx: make(map[CellAddr]uint32),
			rngIdx: make(map[RangeAddr]uint32),
		}
		if err := subCompiler.compileNode(n.Body); err != nil {
			return err
		}
		subFormula := &CompiledFormula{
			Code:        subCompiler.code,
			Consts:      subCompiler.consts,
			Refs:        subCompiler.refs,
			Ranges:      subCompiler.ranges,
			SubFormulas: subCompiler.subFormulas,
		}
		subIdx := len(c.subFormulas)
		c.subFormulas = append(c.subFormulas, subFormula)
		c.emit(OpScan, uint32(subIdx))

	case *ByRowExpr:
		// Push array in array context
		c.emit(OpEnterArrayCtx, 0)
		if err := c.compileNodeCtx(n.Array, true); err != nil {
			return err
		}
		c.emit(OpLeaveArrayCtx, 0)
		// Compile lambda body as sub-formula
		subCompiler := &compiler{
			numIdx: make(map[float64]uint32),
			strIdx: make(map[string]uint32),
			refIdx: make(map[CellAddr]uint32),
			rngIdx: make(map[RangeAddr]uint32),
		}
		if err := subCompiler.compileNode(n.Body); err != nil {
			return err
		}
		subFormula := &CompiledFormula{
			Code:        subCompiler.code,
			Consts:      subCompiler.consts,
			Refs:        subCompiler.refs,
			Ranges:      subCompiler.ranges,
			SubFormulas: subCompiler.subFormulas,
		}
		subIdx := len(c.subFormulas)
		c.subFormulas = append(c.subFormulas, subFormula)
		c.emit(OpByRow, uint32(subIdx))

	case *ByColExpr:
		// Push array in array context
		c.emit(OpEnterArrayCtx, 0)
		if err := c.compileNodeCtx(n.Array, true); err != nil {
			return err
		}
		c.emit(OpLeaveArrayCtx, 0)
		// Compile lambda body as sub-formula
		subCompiler := &compiler{
			numIdx: make(map[float64]uint32),
			strIdx: make(map[string]uint32),
			refIdx: make(map[CellAddr]uint32),
			rngIdx: make(map[RangeAddr]uint32),
		}
		if err := subCompiler.compileNode(n.Body); err != nil {
			return err
		}
		subFormula := &CompiledFormula{
			Code:        subCompiler.code,
			Consts:      subCompiler.consts,
			Refs:        subCompiler.refs,
			Ranges:      subCompiler.ranges,
			SubFormulas: subCompiler.subFormulas,
		}
		subIdx := len(c.subFormulas)
		c.subFormulas = append(c.subFormulas, subFormula)
		c.emit(OpByCol, uint32(subIdx))

	case *MakeArrayExpr:
		// Compile rows and cols expressions
		if err := c.compileNodeCtx(n.Rows, inArrayCtx); err != nil {
			return err
		}
		if err := c.compileNodeCtx(n.Cols, inArrayCtx); err != nil {
			return err
		}
		// Compile lambda body as sub-formula
		subCompiler := &compiler{
			numIdx: make(map[float64]uint32),
			strIdx: make(map[string]uint32),
			refIdx: make(map[CellAddr]uint32),
			rngIdx: make(map[RangeAddr]uint32),
		}
		if err := subCompiler.compileNode(n.Body); err != nil {
			return err
		}
		subFormula := &CompiledFormula{
			Code:        subCompiler.code,
			Consts:      subCompiler.consts,
			Refs:        subCompiler.refs,
			Ranges:      subCompiler.ranges,
			SubFormulas: subCompiler.subFormulas,
		}
		subIdx := len(c.subFormulas)
		c.subFormulas = append(c.subFormulas, subFormula)
		c.emit(OpMakeArrayLambda, uint32(subIdx))

	case *ParamRef:
		c.emit(OpLoadParam, uint32(n.Slot))

	default:
		return fmt.Errorf("unsupported AST node type %T", node)
	}
	return nil
}

func formulaNeedsTopLevelArray(node Node) bool {
	switch n := node.(type) {
	case *RangeRef:
		return true
	case *ArrayLit:
		return true
	case *UnaryExpr:
		return formulaNeedsTopLevelArray(n.Operand)
	case *BinaryExpr:
		return formulaNeedsTopLevelArray(n.Left) || formulaNeedsTopLevelArray(n.Right)
	case *FuncCall:
		return funcCallNeedsTopLevelArray(n)
	case *MapExpr:
		return true
	case *ScanExpr:
		return true
	case *ByRowExpr:
		return true
	case *ByColExpr:
		return true
	case *MakeArrayExpr:
		return true
	}
	return false
}

func funcCallNeedsTopLevelArray(call *FuncCall) bool {
	name := normalizeFuncName(call.Name)
	switch name {
	case "OFFSET", "INDIRECT":
		return true
	}
	if _, ok := dynamicArrayFunctions[name]; ok {
		return true
	}
	if indices, ok := criteriaTopLevelArrayArgIndexes(name, len(call.Args)); ok {
		for _, idx := range indices {
			if formulaNeedsTopLevelArray(call.Args[idx]) {
				return true
			}
		}
		return false
	}
	if !functionCanReturnArrayFromArrayArgs(name) {
		return false
	}
	for _, arg := range call.Args {
		if formulaNeedsTopLevelArray(arg) {
			return true
		}
	}
	return false
}

func criteriaTopLevelArrayArgIndexes(name string, argc int) ([]int, bool) {
	switch name {
	case "COUNTIF", "SUMIF", "AVERAGEIF":
		if argc >= 2 {
			return []int{1}, true
		}
		return nil, true
	case "COUNTIFS":
		idxs := make([]int, 0, argc/2)
		for i := 1; i < argc; i += 2 {
			idxs = append(idxs, i)
		}
		return idxs, true
	case "SUMIFS", "AVERAGEIFS":
		idxs := make([]int, 0, argc/2)
		for i := 2; i < argc; i += 2 {
			idxs = append(idxs, i)
		}
		return idxs, true
	default:
		return nil, false
	}
}

func functionCanReturnArrayFromArrayArgs(name string) bool {
	if name == "IF" {
		return true
	}
	if name == "INDEX" {
		return true
	}
	switch name {
	case "COUNTIF", "SUMIF", "AVERAGEIF", "COUNTIFS", "SUMIFS", "AVERAGEIFS":
		return true
	}
	if functionUsesElementwiseContract(name) {
		return true
	}
	return false
}

func (c *compiler) compileFuncCall(n *FuncCall, inArrayCtx bool) error {
	name := strings.ToUpper(n.Name)
	// The _xludf. prefix means "user-defined function" in the
	// saved formula format. These are not real functions and
	// must always produce #NAME?. We emit a runtime error instead of
	// a compile error so that wrapping functions like IFERROR can
	// catch the #NAME? value.
	if strings.HasPrefix(name, "_XLUDF.") {
		c.emit(OpPushError, uint32(ErrValNAME))
		return nil
	}
	// Strip OOXML prefixes (_xlfn., _xlfn._xlws.) so formulas read
	// from XLSX files compile correctly even if the prefix wasn't
	// removed earlier.
	name = strings.TrimPrefix(name, "_XLFN._XLWS.")
	name = strings.TrimPrefix(name, "_XLFN.")
	// ISOMITTED checks whether its argument was an omitted LAMBDA parameter.
	// Because LAMBDA invocation substitutes the raw AST node (including
	// *EmptyArg for empty argument slots), this is a purely syntactic check
	// at compile time.
	if name == "ISOMITTED" {
		if len(n.Args) != 1 {
			c.emit(OpPushError, uint32(ErrValVALUE))
			return nil
		}
		if _, isEmpty := n.Args[0].(*EmptyArg); isEmpty {
			c.emit(OpPushBool, 1)
			return nil
		}
		c.emit(OpPushBool, 0)
		return nil
	}
	funcID := LookupFunc(name)
	if funcID < 0 {
		// Unknown function names produce #NAME? at runtime in Excel so that
		// wrappers like IFERROR/ISERROR/ERROR.TYPE can catch the error.
		c.emit(OpPushError, uint32(ErrValNAME))
		return nil
	}
	argc := len(n.Args)
	if argc > 255 {
		return fmt.Errorf("function %q has %d arguments (max 255)", n.Name, argc)
	}

	// IFERROR and IFNA evaluate their arguments in Excel's legacy scalar
	// context when called inside a legacy array-forcing function (SUMPRODUCT,
	// SUMIF, MATCH, …): range references implicit-intersect at the formula
	// cell's row/column before the function runs. Evidence:
	// testdata/error_propagation/{06,11,12}_* all resolve SUMPRODUCT(IFERROR(
	// range, …), …) to #VALUE! because the implicit-intersected IFERROR
	// returns a scalar that SUMPRODUCT can't line up against the sibling
	// range. Dynamic-array-native wrappers (FILTER, SORT, UNIQUE, …) keep
	// array semantics — verified against Excel in spill fixture
	// 33_iferror_datevalue_full_column.xlsx. CSE array formulas opt out via
	// the opcode's ctx.IsArrayFormula guard.
	if name == "IFERROR" || name == "IFNA" {
		if c.dynamicArrayDepth == 0 {
			wasArrayCtx := inArrayCtx
			if wasArrayCtx {
				c.emit(OpLeaveArrayCtx, 0)
			}
			for i, arg := range n.Args {
				if err := c.compileNodeCtx(arg, false); err != nil {
					return err
				}
				// Arg 0 collapses fully (legacy implicit intersection over
				// range-backed and anonymous arrays alike) because Excel's
				// IFERROR oracle for patterns like
				//   SUM(IFERROR(MAP(...with-errors...), 0))
				// shows MAP's anonymous output collapsed to its top-left
				// cell. Arg 1 only intersects range-backed arrays so that
				// anonymous fallbacks like SEQUENCE(5) keep their shape in
				//   ROWS(IFERROR(FILTER(...empty...), SEQUENCE(5)))
				// — otherwise the fallback array would collapse to its
				// anchor scalar before reaching ROWS/SUM/etc.
				if i == 0 {
					c.emit(OpImplicitIntersect, 0)
				} else {
					c.emit(OpImplicitIntersectRefOnly, 0)
				}
			}
			if wasArrayCtx {
				c.emit(OpEnterArrayCtx, 0)
			}
			c.emit(OpCall, callOperand(funcID, argc, c.legacyArrayDepth > 0 || wasArrayCtx))
			return nil
		}
		// Fall through to the element-wise lifting path for dynamic-array
		// contexts.
	}

	inheritedArrayCtx := inArrayCtx
	// Array-forcing behavior should apply to the direct SUMPRODUCT/INDEX/etc.
	// argument expressions, but it must not leak into nested non-array
	// function arguments. Suspending the inherited array context here matches
	// Excel's legacy implicit-intersection behavior for formulas like:
	//   SUMPRODUCT(mask*IF(range-range*scalar>0, range-range*scalar, 0))
	// Element-wise functions (ROUND, IFERROR, etc.) must NOT have their
	// inherited array context suspended, because their arguments need array
	// semantics for the arithmetic inside them to broadcast correctly.
	// Without this, expressions like ROUND(G*scalar, 0) inside FILTER
	// conditions lose array context, causing G to be implicitly intersected
	// to a single cell.
	suspendInheritedArrayCtx := inArrayCtx && !IsArrayFunc(name) && !functionUsesElementwiseContract(name)
	if suspendInheritedArrayCtx {
		c.emit(OpLeaveArrayCtx, 0)
		// The deferred restore intentionally covers the early-return paths
		// below (AREAS, COLUMN/ROW, ISREF, OFFSET) — if we suspended the
		// inherited array context we must always re-enter it before returning.
		defer c.emit(OpEnterArrayCtx, 0)
		inArrayCtx = false
	}

	if name == "AREAS" && argc == 1 {
		if union, ok := n.Args[0].(*UnionRef); ok {
			idx := c.addNumConst(float64(len(union.Areas)))
			c.emit(OpPushNum, idx)
			return nil
		}
	}
	// COLUMN and ROW need the cell reference coordinates, not the resolved
	// cell value.  When the single argument is a direct cell reference, push
	// a ValueRef (address only) so the function can extract col/row.
	if (name == "AREAS" || name == "COLUMN" || name == "ROW" || name == "ISFORMULA" || name == "FORMULATEXT" || name == "ANCHORARRAY" || name == "ISREF") && argc == 1 {
		if cr, ok := n.Args[0].(*CellRef); ok && !cr.DotNotation && cr.SheetEnd == "" && cr.Col <= maxCols {
			idx := c.addRef(CellAddr{Sheet: cr.Sheet, Col: cr.Col, Row: cr.Row})
			c.emit(OpLoadCellRef, idx)
			c.emit(OpCall, callOperand(funcID, argc, c.legacyArrayDepth > 0 || inheritedArrayCtx))
			return nil
		}
		// For ISREF, range references are also references.
		if name == "ISREF" {
			if rr, ok := n.Args[0].(*RangeRef); ok && rr.From.Col >= 1 && rr.From.Col <= maxCols && rr.To.Col >= 1 && rr.To.Col <= maxCols {
				idx := c.addRef(CellAddr{Sheet: rr.From.Sheet, Col: rr.From.Col, Row: rr.From.Row})
				c.emit(OpLoadCellRef, idx)
				c.emit(OpCall, callOperand(funcID, argc, c.legacyArrayDepth > 0 || inheritedArrayCtx))
				return nil
			}
			// INDIRECT (and OFFSET) are ref-returning functions.
			// When wrapped in ISREF, compile the inner call normally
			// and then use OpRefResultToBool: non-error → TRUE,
			// error → FALSE.
			if fc, ok := n.Args[0].(*FuncCall); ok {
				inner := strings.ToUpper(fc.Name)
				inner = strings.TrimPrefix(inner, "_XLFN._XLWS.")
				inner = strings.TrimPrefix(inner, "_XLFN.")
				if inner == "INDIRECT" || inner == "OFFSET" {
					if err := c.compileNodeCtx(fc, inArrayCtx); err != nil {
						return err
					}
					c.emit(OpRefResultToBool, 0)
					return nil
				}
			}
		}
	}
	// OFFSET needs its first argument as a raw reference (ValueRef
	// for single cells, ValueArray+RangeOrigin for ranges) so it can
	// compute the offset address.  Remaining arguments are compiled
	// normally.
	if name == "OFFSET" && argc >= 1 {
		first := n.Args[0]
		switch ref := first.(type) {
		case *CellRef:
			if ref.Col <= maxCols {
				idx := c.addRef(CellAddr{Sheet: ref.Sheet, Col: ref.Col, Row: ref.Row})
				c.emit(OpLoadCellRef, idx)
			} else {
				c.emit(OpPushError, uint32(ErrValNAME))
			}
		default:
			// Range references already produce ValueArray with RangeOrigin.
			if isDirectRangeRefNode(first) {
				c.enterLegacyArrayCtx()
				if err := c.compileNodeCtx(first, true); err != nil {
					return err
				}
				c.leaveLegacyArrayCtx()
			} else {
				if err := c.compileNodeCtx(first, inArrayCtx); err != nil {
					return err
				}
			}
		}
		for _, arg := range n.Args[1:] {
			if err := c.compileNodeCtx(arg, inArrayCtx); err != nil {
				return err
			}
		}
		c.emit(OpCall, callOperand(funcID, argc, c.legacyArrayDepth > 0 || inheritedArrayCtx))
		return nil
	}

	arrayCtx := IsArrayFunc(name)
	if arrayCtx {
		c.enterLegacyArrayCtx()
	}
	if IsDynamicArrayFunc(name) {
		c.dynamicArrayDepth++
		defer func() { c.dynamicArrayDepth-- }()
	}
	for i, arg := range n.Args {
		forceInheritedArrayArg := suspendInheritedArrayCtx && inheritedArrayCtx && inheritedArrayEvalForFuncArg(name, i)
		if !arrayCtx {
			switch ArgEvalModeForFuncArg(name, i) {
			case FuncArgEvalArray:
				c.enterLegacyArrayCtx()
				if err := c.compileNodeCtx(arg, true); err != nil {
					return err
				}
				c.leaveLegacyArrayCtx()
				continue
			case FuncArgEvalDirectRange:
				if isDirectRangeRefNode(arg) {
					c.enterLegacyArrayCtx()
					if err := c.compileNodeCtx(arg, true); err != nil {
						return err
					}
					c.leaveLegacyArrayCtx()
					continue
				}
			}
			if forceInheritedArrayArg {
				c.enterLegacyArrayCtx()
				if err := c.compileNodeCtx(arg, true); err != nil {
					return err
				}
				c.leaveLegacyArrayCtx()
				continue
			}
		}
		if err := c.compileNodeCtx(arg, inArrayCtx || arrayCtx); err != nil {
			return err
		}
	}
	if arrayCtx {
		c.leaveLegacyArrayCtx()
	}
	c.emit(OpCall, callOperand(funcID, argc, c.legacyArrayDepth > 0 || inheritedArrayCtx || arrayCtx))
	return nil
}

// callOperand packs funcID, argc, and the inheritedArrayCtx flag into a single
// OpCall operand. The flag tells the runtime to skip the legacy
// implicit-intersection gate: set it whenever this call site was compiled
// inside a (possibly suspended) array-forcing context.
func callOperand(funcID, argc int, inheritedArrayCtx bool) uint32 {
	operand := uint32(funcID)<<8 | uint32(argc)
	if inheritedArrayCtx {
		operand |= callFlagInheritedArrayCtx
	}
	return operand
}

func binaryOpCode(op string) (OpCode, error) {
	switch op {
	case "+":
		return OpAdd, nil
	case "-":
		return OpSub, nil
	case "*":
		return OpMul, nil
	case "/":
		return OpDiv, nil
	case "^":
		return OpPow, nil
	case "&":
		return OpConcat, nil
	case "=":
		return OpEq, nil
	case "<>":
		return OpNe, nil
	case "<":
		return OpLt, nil
	case "<=":
		return OpLe, nil
	case ">":
		return OpGt, nil
	case ">=":
		return OpGe, nil
	default:
		return 0, fmt.Errorf("unknown binary operator %q", op)
	}
}

func isDirectRangeRefNode(node Node) bool {
	switch node.(type) {
	case *RangeRef, *IntersectRef, *UnionRef:
		return true
	}
	return false
}

// compileIntersectOperand compiles an operand of the intersection operator.
// CellRef operands are loaded as ValueRef so rangeIntersect can recover their
// address; otherwise OpLoadCell would push a value with no origin and the
// intersection of two single-cell refs would fall through to #VALUE! instead
// of computing the rectangle (and #NULL! when they don't overlap).
func (c *compiler) compileIntersectOperand(n Node, inArrayCtx bool) error {
	if cr, ok := n.(*CellRef); ok && !cr.DotNotation && cr.Name == "" && cr.SheetEnd == "" && cr.Col >= 1 && cr.Col <= maxCols {
		idx := c.addRef(CellAddr{Sheet: cr.Sheet, Col: cr.Col, Row: cr.Row})
		c.emit(OpLoadCellRef, idx)
		return nil
	}
	return c.compileNodeCtx(n, inArrayCtx)
}

// isReferenceNode reports whether a node can legitimately appear as an
// operand of the intersection operator. This mirrors Excel's rule: only
// reference-producing expressions (cells, ranges, nested intersections, and
// certain reference-returning function calls) are permitted.
func isReferenceNode(n Node) bool {
	switch v := n.(type) {
	case *CellRef, *RangeRef, *IntersectRef, *UnionRef:
		return true
	case *FuncCall:
		switch strings.ToUpper(v.Name) {
		case "OFFSET", "INDIRECT", "INDEX", "CHOOSE", "IF", "IFS", "SWITCH", "ANCHORARRAY":
			return true
		}
	}
	return false
}
