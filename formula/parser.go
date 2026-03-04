package formula

import (
	"fmt"
	"strconv"
	"strings"
)

// Parser is a Pratt parser that transforms a token stream into an AST.
type Parser struct {
	tokens []Token
	pos    int
}

// Parse tokenizes and parses a formula string into an AST.
// The formula should not include the leading '=' sign.
func Parse(formula string) (Node, error) {
	tokens, err := Tokenize(formula)
	if err != nil {
		return nil, err
	}
	p := &Parser{tokens: tokens}
	node, err := p.parseExpression(0)
	if err != nil {
		return nil, err
	}
	if p.peek().Type != TokEOF {
		return nil, fmt.Errorf("unexpected token %s at position %d", p.peek(), p.peek().Pos)
	}
	return node, nil
}

func (p *Parser) peek() Token {
	if p.pos >= len(p.tokens) {
		return Token{Type: TokEOF}
	}
	return p.tokens[p.pos]
}

func (p *Parser) advance() Token {
	tok := p.peek()
	if p.pos < len(p.tokens) {
		p.pos++
	}
	return tok
}

func (p *Parser) expect(typ TokenType) (Token, error) {
	tok := p.advance()
	if tok.Type != typ {
		return tok, fmt.Errorf("expected %s but got %s at position %d", typ, tok.Type, tok.Pos)
	}
	return tok, nil
}

// Binding power definitions for infix operators.
// Left BP is compared against minBP; Right BP is passed to the recursive call.
// Left-associative: rightBP = leftBP + 1. Right-associative: rightBP = leftBP.
type bindingPower struct {
	left  int
	right int
}

var infixBP = map[string]bindingPower{
	"=":  {2, 3},
	"<>": {2, 3},
	"<":  {2, 3},
	">":  {2, 3},
	"<=": {2, 3},
	">=": {2, 3},
	"&":  {4, 5},
	"+":  {6, 7},
	"-":  {6, 7},
	"*":  {8, 9},
	"/":  {8, 9},
	"^":  {10, 10}, // right-associative
}

const (
	colonLeftBP  = 14
	colonRightBP = 15
	prefixRBP    = 11 // unary - and + bind tighter than ^ (Excel convention: -2^2 = 4)

	maxExcelRow = maxExcelRows // maximum row number in Excel
	maxExcelCol = maxExcelCols // maximum column number in Excel (XFD)
)

// parseExpression is the core Pratt parsing loop.
func (p *Parser) parseExpression(minBP int) (Node, error) {
	left, err := p.parseNud()
	if err != nil {
		return nil, err
	}

	// Greedy postfix % — consumed immediately, not in the BP table.
	for p.peek().Type == TokPercent {
		p.advance()
		left = &PostfixExpr{Op: "%", Operand: left}
	}

	for {
		tok := p.peek()

		if tok.Type == TokOp {
			bp, ok := infixBP[tok.Value]
			if !ok || bp.left < minBP {
				break
			}
			p.advance()
			right, err := p.parseExpression(bp.right)
			if err != nil {
				return nil, err
			}
			left = &BinaryExpr{Op: tok.Value, Left: left, Right: right}
			for p.peek().Type == TokPercent {
				p.advance()
				left = &PostfixExpr{Op: "%", Operand: left}
			}
			continue
		}

		if tok.Type == TokColon {
			if colonLeftBP < minBP {
				break
			}
			p.advance()
			right, err := p.parseExpression(colonRightBP)
			if err != nil {
				return nil, err
			}

			// Convert row-only references: both sides must be NumberLit
			// for a row range like 5:6 → A5:XFD6.
			fromRef, fromOK := left.(*CellRef)
			toRef, toOK := right.(*CellRef)
			if !fromOK || !toOK {
				fromNum, fromIsNum := left.(*NumberLit)
				toNum, toIsNum := right.(*NumberLit)
				if fromIsNum && toIsNum {
					fromRow := int(fromNum.Value)
					toRow := int(toNum.Value)
					if fromRow < 1 || toRow < 1 || float64(fromRow) != fromNum.Value || float64(toRow) != toNum.Value {
						return nil, fmt.Errorf("invalid row range %s:%s", fromNum.Raw, toNum.Raw)
					}
					fromRef = &CellRef{Col: 0, Row: fromRow}
					toRef = &CellRef{Col: 0, Row: toRow}
				} else if !fromOK {
					return nil, fmt.Errorf("left side of ':' must be a cell reference, got %s", left)
				} else {
					return nil, fmt.Errorf("right side of ':' must be a cell reference, got %s", right)
				}
			}

			// Expand column-only references (Row==0) into full-column ranges.
			// F:F becomes F1:F1048576.
			if fromRef.Row == 0 {
				fromRef.Row = 1
			}
			if toRef.Row == 0 {
				toRef.Row = maxExcelRow
			}
			// Expand row-only references (Col==0) into full-row ranges.
			// 5:6 becomes A5:XFD6.
			if fromRef.Col == 0 {
				fromRef.Col = 1
			}
			if toRef.Col == 0 {
				toRef.Col = maxExcelCol
			}
			left = &RangeRef{From: fromRef, To: toRef}
			for p.peek().Type == TokPercent {
				p.advance()
				left = &PostfixExpr{Op: "%", Operand: left}
			}
			continue
		}

		break
	}

	return left, nil
}

// parseNud handles prefix parselets (atoms and prefix operators).
func (p *Parser) parseNud() (Node, error) {
	tok := p.peek()

	switch tok.Type {
	case TokNumber:
		p.advance()
		val, err := strconv.ParseFloat(tok.Value, 64)
		if err != nil {
			return nil, fmt.Errorf("invalid number %q at position %d: %w", tok.Value, tok.Pos, err)
		}
		return &NumberLit{Value: val, Raw: tok.Value}, nil

	case TokString:
		p.advance()
		return &StringLit{Value: tok.Value}, nil

	case TokBool:
		p.advance()
		return &BoolLit{Value: strings.ToUpper(tok.Value) == "TRUE"}, nil

	case TokError:
		p.advance()
		return &ErrorLit{Code: ErrorCode(strings.ToUpper(tok.Value))}, nil

	case TokCellRef:
		p.advance()
		ref, err := parseCellRefToken(tok.Value)
		if err != nil {
			return nil, fmt.Errorf("invalid cell reference %q at position %d: %w", tok.Value, tok.Pos, err)
		}
		return ref, nil

	case TokFunc:
		return p.parseFunc()

	case TokLParen:
		p.advance()
		expr, err := p.parseExpression(0)
		if err != nil {
			return nil, err
		}
		if _, err := p.expect(TokRParen); err != nil {
			return nil, fmt.Errorf("unmatched '(' at position %d", tok.Pos)
		}
		return expr, nil

	case TokArrayOpen:
		return p.parseArray()

	case TokOp:
		if tok.Value == "-" || tok.Value == "+" {
			p.advance()
			operand, err := p.parseExpression(prefixRBP)
			if err != nil {
				return nil, err
			}
			return &UnaryExpr{Op: tok.Value, Operand: operand}, nil
		}
		return nil, fmt.Errorf("unexpected operator %q at position %d", tok.Value, tok.Pos)

	case TokEOF:
		return nil, fmt.Errorf("unexpected end of formula")

	default:
		return nil, fmt.Errorf("unexpected token %s at position %d", tok, tok.Pos)
	}
}

// parseFunc parses a function call: NAME( arg, arg, ... )
func (p *Parser) parseFunc() (Node, error) {
	tok := p.advance()
	name := strings.TrimSuffix(tok.Value, "(")

	// Zero-arg function: immediately followed by ).
	if p.peek().Type == TokRParen {
		p.advance()
		return &FuncCall{Name: name}, nil
	}

	var args []Node
	arg, err := p.parseExpression(0)
	if err != nil {
		return nil, err
	}
	args = append(args, arg)

	for p.peek().Type == TokComma {
		p.advance()
		arg, err := p.parseExpression(0)
		if err != nil {
			return nil, err
		}
		args = append(args, arg)
	}

	if _, err := p.expect(TokRParen); err != nil {
		return nil, fmt.Errorf("expected ')' to close function %s at position %d", name, tok.Pos)
	}

	return &FuncCall{Name: name, Args: args}, nil
}

// parseArray parses an array literal: { expr, expr ; expr, expr }
func (p *Parser) parseArray() (Node, error) {
	p.advance() // consume {

	var rows [][]Node
	var currentRow []Node

	elem, err := p.parseExpression(0)
	if err != nil {
		return nil, err
	}
	currentRow = append(currentRow, elem)

loop:
	for {
		tok := p.peek()
		switch tok.Type {
		case TokComma:
			p.advance()
			elem, err := p.parseExpression(0)
			if err != nil {
				return nil, err
			}
			currentRow = append(currentRow, elem)
		case TokSemicolon:
			p.advance()
			rows = append(rows, currentRow)
			currentRow = nil
			elem, err := p.parseExpression(0)
			if err != nil {
				return nil, err
			}
			currentRow = append(currentRow, elem)
		default:
			break loop
		}
	}
	rows = append(rows, currentRow)

	if _, err := p.expect(TokArrayClose); err != nil {
		return nil, fmt.Errorf("expected '}' to close array literal")
	}

	return &ArrayLit{Rows: rows}, nil
}
