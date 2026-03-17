package formula

import (
	"fmt"
	"strings"
)

// Node is the interface implemented by all AST nodes.
type Node interface {
	// String returns an S-expression representation of the node for debugging and test output.
	String() string
	nodeMarker()
}

// ErrorCode represents a formula error value.
type ErrorCode string

const (
	ErrDIV0        ErrorCode = "#DIV/0!"
	ErrNA          ErrorCode = "#N/A"
	ErrNAME        ErrorCode = "#NAME?"
	ErrNULL        ErrorCode = "#NULL!"
	ErrNUM         ErrorCode = "#NUM!"
	ErrREF         ErrorCode = "#REF!"
	ErrVALUE       ErrorCode = "#VALUE!"
	ErrSPILL       ErrorCode = "#SPILL!"
	ErrCALC        ErrorCode = "#CALC!"
	ErrGETTINGDATA ErrorCode = "#GETTING_DATA"
)

// NumberLit represents a numeric literal.
type NumberLit struct {
	Value float64
	Raw   string // original text from the formula
}

func (n *NumberLit) nodeMarker() {}
func (n *NumberLit) String() string {
	return n.Raw
}

// StringLit represents a string literal.
type StringLit struct {
	Value string
}

func (n *StringLit) nodeMarker() {}
func (n *StringLit) String() string {
	return fmt.Sprintf("%q", n.Value)
}

// BoolLit represents a boolean literal (TRUE or FALSE).
type BoolLit struct {
	Value bool
}

func (n *BoolLit) nodeMarker() {}
func (n *BoolLit) String() string {
	if n.Value {
		return "TRUE"
	}
	return "FALSE"
}

// ErrorLit represents an error literal like #N/A or #DIV/0!.
type ErrorLit struct {
	Code ErrorCode
}

func (n *ErrorLit) nodeMarker() {}
func (n *ErrorLit) String() string {
	return string(n.Code)
}

// EmptyArg represents an omitted function argument (e.g. the missing 3rd arg in ADDRESS(1,1,,"Data")).
type EmptyArg struct{}

func (n *EmptyArg) nodeMarker() {}
func (n *EmptyArg) String() string {
	return "<empty>"
}

// CellRef represents a cell reference, possibly sheet-qualified.
type CellRef struct {
	Sheet       string // empty if not sheet-qualified
	SheetEnd    string // non-empty for 3D references (Sheet2:Sheet5!A1); the end sheet name
	Col         int    // 1-based column number
	Row         int    // 1-based row number
	AbsCol      bool   // true if column is absolute ($A)
	AbsRow      bool   // true if row is absolute ($1)
	DotNotation bool   // true if parsed with LibreOffice dot separator (Sheet1.A1); returns #NAME? in standard mode
}

func (n *CellRef) nodeMarker() {}
func (n *CellRef) String() string {
	var b strings.Builder
	if n.Sheet != "" {
		if needsQuoting(n.Sheet) {
			b.WriteByte('\'')
			b.WriteString(strings.ReplaceAll(n.Sheet, "'", "''"))
			b.WriteByte('\'')
		} else {
			b.WriteString(n.Sheet)
		}
		b.WriteByte('!')
	}
	if n.AbsCol {
		b.WriteByte('$')
	}
	b.WriteString(colNumberToLetters(n.Col))
	if n.AbsRow {
		b.WriteByte('$')
	}
	fmt.Fprintf(&b, "%d", n.Row)
	return b.String()
}

// RangeRef represents a range reference like A1:B5.
type RangeRef struct {
	From *CellRef
	To   *CellRef
}

func (n *RangeRef) nodeMarker() {}
func (n *RangeRef) String() string {
	return fmt.Sprintf("(: %s %s)", n.From, n.To)
}

// UnionRef represents a parenthesized multi-area reference list like
// (A1:B2,C3,D4:E5). The parser only produces this node for AREAS.
type UnionRef struct {
	Areas []Node
}

func (n *UnionRef) nodeMarker() {}
func (n *UnionRef) String() string {
	if len(n.Areas) == 0 {
		return "(union)"
	}
	parts := make([]string, len(n.Areas))
	for i, area := range n.Areas {
		parts[i] = area.String()
	}
	return fmt.Sprintf("(union %s)", strings.Join(parts, " "))
}

// UnaryExpr represents a prefix unary operation like -A1 or +5.
type UnaryExpr struct {
	Op      string
	Operand Node
}

func (n *UnaryExpr) nodeMarker() {}
func (n *UnaryExpr) String() string {
	return fmt.Sprintf("(%s %s)", n.Op, n.Operand)
}

// BinaryExpr represents an infix binary operation.
type BinaryExpr struct {
	Op    string
	Left  Node
	Right Node
}

func (n *BinaryExpr) nodeMarker() {}
func (n *BinaryExpr) String() string {
	return fmt.Sprintf("(%s %s %s)", n.Op, n.Left, n.Right)
}

// PostfixExpr represents a postfix operation like 50%.
type PostfixExpr struct {
	Op      string
	Operand Node
}

func (n *PostfixExpr) nodeMarker() {}
func (n *PostfixExpr) String() string {
	return fmt.Sprintf("(%s %s)", n.Op, n.Operand)
}

// FuncCall represents a function invocation like SUM(A1:A10).
type FuncCall struct {
	Name string // function name without the trailing '('
	Args []Node
}

func (n *FuncCall) nodeMarker() {}
func (n *FuncCall) String() string {
	if len(n.Args) == 0 {
		return fmt.Sprintf("(%s)", n.Name)
	}
	args := make([]string, len(n.Args))
	for i, a := range n.Args {
		args[i] = a.String()
	}
	return fmt.Sprintf("(%s %s)", n.Name, strings.Join(args, " "))
}

// ArrayLit represents an array literal like {1,2;3,4}.
type ArrayLit struct {
	Rows [][]Node
}

func (n *ArrayLit) nodeMarker() {}
func (n *ArrayLit) String() string {
	var b strings.Builder
	b.WriteByte('{')
	for i, row := range n.Rows {
		if i > 0 {
			b.WriteByte(';')
		}
		for j, elem := range row {
			if j > 0 {
				b.WriteByte(',')
			}
			b.WriteString(elem.String())
		}
	}
	b.WriteByte('}')
	return b.String()
}

// ParamRef represents a reference to a lambda parameter inside a MAP/REDUCE/SCAN body.
type ParamRef struct {
	Slot int    // parameter index (0-based)
	Name string // parameter name for debugging
}

func (n *ParamRef) nodeMarker() {}
func (n *ParamRef) String() string {
	return fmt.Sprintf("$param(%d:%s)", n.Slot, n.Name)
}

// MapExpr represents a MAP(arrays..., LAMBDA(params..., body)) expression.
type MapExpr struct {
	Arrays     []Node   // array expressions
	ParamNames []string // lambda parameter names (uppercase)
	Body       Node     // lambda body with param refs replaced by ParamRef nodes
}

func (n *MapExpr) nodeMarker() {}
func (n *MapExpr) String() string {
	return fmt.Sprintf("(MAP arrays=%d params=%v)", len(n.Arrays), n.ParamNames)
}

// ReduceExpr represents a REDUCE(initial, array, LAMBDA(acc, val, body)) expression.
type ReduceExpr struct {
	InitialValue Node     // initial accumulator value (may be *EmptyArg if omitted)
	Array        Node     // array expression
	ParamNames   []string // [accumulator_name, value_name]
	Body         Node     // lambda body with param refs replaced by ParamRef nodes
}

func (n *ReduceExpr) nodeMarker() {}
func (n *ReduceExpr) String() string {
	return fmt.Sprintf("(REDUCE params=%v)", n.ParamNames)
}

// ScanExpr represents a SCAN(initial, array, LAMBDA(acc, val, body)) expression.
type ScanExpr struct {
	InitialValue Node     // initial accumulator value (may be *EmptyArg if omitted)
	Array        Node     // array expression
	ParamNames   []string // [accumulator_name, value_name]
	Body         Node     // lambda body with param refs replaced by ParamRef nodes
}

func (n *ScanExpr) nodeMarker() {}
func (n *ScanExpr) String() string {
	return fmt.Sprintf("(SCAN params=%v)", n.ParamNames)
}

// ByRowExpr represents a BYROW(array, LAMBDA(row, body)) expression.
type ByRowExpr struct {
	Array      Node     // array expression
	ParamNames []string // single param name for the row
	Body       Node     // lambda body with param refs replaced by ParamRef nodes
}

func (n *ByRowExpr) nodeMarker() {}
func (n *ByRowExpr) String() string {
	return fmt.Sprintf("(BYROW params=%v)", n.ParamNames)
}

// ByColExpr represents a BYCOL(array, LAMBDA(col, body)) expression.
type ByColExpr struct {
	Array      Node     // array expression
	ParamNames []string // single param name for the column
	Body       Node     // lambda body with param refs replaced by ParamRef nodes
}

func (n *ByColExpr) nodeMarker() {}
func (n *ByColExpr) String() string {
	return fmt.Sprintf("(BYCOL params=%v)", n.ParamNames)
}

// MakeArrayExpr represents a MAKEARRAY(rows, cols, LAMBDA(r, c, body)) expression.
type MakeArrayExpr struct {
	Rows       Node     // rows expression
	Cols       Node     // cols expression
	ParamNames []string // [row_name, col_name]
	Body       Node     // lambda body
}

func (n *MakeArrayExpr) nodeMarker() {}
func (n *MakeArrayExpr) String() string {
	return fmt.Sprintf("(MAKEARRAY params=%v)", n.ParamNames)
}

// needsQuoting returns true if a sheet name contains characters that require quoting.
func needsQuoting(name string) bool {
	for _, c := range name {
		if !((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_') {
			return true
		}
	}
	return false
}

// colNumberToLetters converts a 1-based column number to column letters (e.g. 1→"A", 27→"AA").
func colNumberToLetters(col int) string {
	var buf [3]byte
	i := len(buf)
	for col > 0 {
		col--
		i--
		buf[i] = byte('A' + col%26)
		col /= 26
	}
	return string(buf[i:])
}
