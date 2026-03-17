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
	c := &compiler{
		numIdx: make(map[float64]uint32),
		strIdx: make(map[string]uint32),
		refIdx: make(map[CellAddr]uint32),
		rngIdx: make(map[RangeAddr]uint32),
	}
	if err := c.compileNode(node); err != nil {
		return nil, err
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
}

func (c *compiler) emit(op OpCode, operand uint32) {
	c.code = append(c.code, Instruction{Op: op, Operand: operand})
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

	case *UnionRef:
		return fmt.Errorf("union references are only supported inside AREAS")

	case *UnaryExpr:
		if err := c.compileNode(n.Operand); err != nil {
			return err
		}
		switch n.Op {
		case "-":
			c.emit(OpNeg, 0)
		case "+":
			// no-op
		default:
			return fmt.Errorf("unknown unary operator %q", n.Op)
		}

	case *BinaryExpr:
		if err := c.compileNode(n.Left); err != nil {
			return err
		}
		if err := c.compileNode(n.Right); err != nil {
			return err
		}
		op, err := binaryOpCode(n.Op)
		if err != nil {
			return err
		}
		c.emit(op, 0)

	case *PostfixExpr:
		if err := c.compileNode(n.Operand); err != nil {
			return err
		}
		switch n.Op {
		case "%":
			c.emit(OpPercent, 0)
		default:
			return fmt.Errorf("unknown postfix operator %q", n.Op)
		}

	case *FuncCall:
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
		funcID := LookupFunc(name)
		if funcID < 0 {
			return fmt.Errorf("unknown function %q", n.Name)
		}
		argc := len(n.Args)
		if argc > 255 {
			return fmt.Errorf("function %q has %d arguments (max 255)", n.Name, argc)
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
			if cr, ok := n.Args[0].(*CellRef); ok && !cr.DotNotation && cr.SheetEnd == "" {
				idx := c.addRef(CellAddr{Sheet: cr.Sheet, Col: cr.Col, Row: cr.Row})
				c.emit(OpLoadCellRef, idx)
				c.emit(OpCall, uint32(funcID)<<8|uint32(argc))
				return nil
			}
			// For ISREF, range references are also references.
			if name == "ISREF" {
				if rr, ok := n.Args[0].(*RangeRef); ok {
					idx := c.addRef(CellAddr{Sheet: rr.From.Sheet, Col: rr.From.Col, Row: rr.From.Row})
					c.emit(OpLoadCellRef, idx)
					c.emit(OpCall, uint32(funcID)<<8|uint32(argc))
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
						if err := c.compileNode(fc); err != nil {
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
				idx := c.addRef(CellAddr{Sheet: ref.Sheet, Col: ref.Col, Row: ref.Row})
				c.emit(OpLoadCellRef, idx)
			default:
				// Range references already produce ValueArray with RangeOrigin.
				if err := c.compileNode(first); err != nil {
					return err
				}
			}
			for _, arg := range n.Args[1:] {
				if err := c.compileNode(arg); err != nil {
					return err
				}
			}
			operand := uint32(funcID)<<8 | uint32(argc)
			c.emit(OpCall, operand)
			break
		}
		arrayCtx := IsArrayFunc(name)
		if arrayCtx {
			c.emit(OpEnterArrayCtx, 0)
		}
		for i, arg := range n.Args {
			if !arrayCtx {
				switch ArgEvalModeForFuncArg(name, i) {
				case FuncArgEvalArray:
					c.emit(OpEnterArrayCtx, 0)
					if err := c.compileNode(arg); err != nil {
						return err
					}
					c.emit(OpLeaveArrayCtx, 0)
					continue
				case FuncArgEvalDirectRange:
					if isDirectRangeRefNode(arg) {
						c.emit(OpEnterArrayCtx, 0)
						if err := c.compileNode(arg); err != nil {
							return err
						}
						c.emit(OpLeaveArrayCtx, 0)
						continue
					}
				}
			}
			if err := c.compileNode(arg); err != nil {
				return err
			}
		}
		if arrayCtx {
			c.emit(OpLeaveArrayCtx, 0)
		}
		operand := uint32(funcID)<<8 | uint32(argc)
		c.emit(OpCall, operand)

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
				if err := c.compileNode(elem); err != nil {
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
			if err := c.compileNode(arr); err != nil {
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
		if err := c.compileNode(n.InitialValue); err != nil {
			return err
		}
		// Push array in array context
		c.emit(OpEnterArrayCtx, 0)
		if err := c.compileNode(n.Array); err != nil {
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
		if err := c.compileNode(n.InitialValue); err != nil {
			return err
		}
		// Push array in array context
		c.emit(OpEnterArrayCtx, 0)
		if err := c.compileNode(n.Array); err != nil {
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
		if err := c.compileNode(n.Array); err != nil {
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
		if err := c.compileNode(n.Array); err != nil {
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
		if err := c.compileNode(n.Rows); err != nil {
			return err
		}
		if err := c.compileNode(n.Cols); err != nil {
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
	_, ok := node.(*RangeRef)
	return ok
}
