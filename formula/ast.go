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

// ErrorCode represents an Excel error value.
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

// CellRef represents a cell reference, possibly sheet-qualified.
type CellRef struct {
	Sheet       string // empty if not sheet-qualified
	SheetEnd    string // non-empty for 3D references (Sheet2:Sheet5!A1); the end sheet name
	Col         int    // 1-based column number
	Row         int    // 1-based row number
	AbsCol      bool   // true if column is absolute ($A)
	AbsRow      bool   // true if row is absolute ($1)
	DotNotation bool   // true if parsed with LibreOffice dot separator (Sheet1.A1); Excel returns #NAME?
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
