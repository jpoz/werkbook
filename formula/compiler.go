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
		Source: source,
		Code:   c.code,
		Consts: c.consts,
		Refs:   c.refs,
		Ranges: c.ranges,
	}, nil
}

type compiler struct {
	code   []Instruction
	consts []Value
	refs   []CellAddr
	ranges []RangeAddr

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

	case *ErrorLit:
		c.emit(OpPushError, uint32(errorCodeFromAST(n.Code)))

	case *CellRef:
		if n.DotNotation {
			// Dot-notation (Sheet1.A1) is a LibreOffice extension; Excel returns #NAME?
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
			// Dot-notation range (Sheet1.A1:Sheet1.A5) is a LibreOffice extension; Excel returns #NAME?
			c.emit(OpPushError, uint32(ErrValNAME))
			return nil
		}
		sheet := n.From.Sheet
		sheetEnd := n.From.SheetEnd
		addr := RangeAddr{
			Sheet:   sheet,
			SheetEnd: sheetEnd,
			FromCol: n.From.Col, FromRow: n.From.Row,
			ToCol: n.To.Col, ToRow: n.To.Row,
		}
		idx := c.addRange(addr)
		if sheetEnd != "" {
			c.emit(OpLoad3DRange, idx)
		} else {
			c.emit(OpLoadRange, idx)
		}

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
		// COLUMN and ROW need the cell reference coordinates, not the resolved
		// cell value.  When the single argument is a direct cell reference, push
		// a ValueRef (address only) so the function can extract col/row.
		if (name == "COLUMN" || name == "ROW") && argc == 1 {
			if cr, ok := n.Args[0].(*CellRef); ok && !cr.DotNotation {
				idx := c.addRef(CellAddr{Sheet: cr.Sheet, Col: cr.Col, Row: cr.Row})
				c.emit(OpLoadCellRef, idx)
				c.emit(OpCall, uint32(funcID)<<8|uint32(argc))
				return nil
			}
		}
		arrayCtx := IsArrayFunc(name)
		if arrayCtx {
			c.emit(OpEnterArrayCtx, 0)
		}
		for _, arg := range n.Args {
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
