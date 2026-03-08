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
	"^":  {10, 11}, // left-associative (matches Excel)
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
				} else {
					// If one side is already an error (e.g. from a cross-sheet
					// range that produced #VALUE!), propagate it instead of
					// returning a parse error. This allows COUNT(#VALUE!) to
					// return 0 instead of failing the entire formula.
					if _, leftIsErr := left.(*ErrorLit); leftIsErr {
						for p.peek().Type == TokPercent {
							p.advance()
							left = &PostfixExpr{Op: "%", Operand: left}
						}
						continue
					}
					if !fromOK {
						return nil, fmt.Errorf("left side of ':' must be a cell reference, got %s", left)
					}
					return nil, fmt.Errorf("right side of ':' must be a cell reference, got %s", right)
				}
			}

			// If the right side has an explicit sheet qualifier that differs
			// from the left side, this is a cross-sheet range which is invalid
			// in Excel (e.g. S1:S3!A1 or Sheet1!A1:Sheet2!B5). Return #VALUE!.
			// Note: Sheet1!A1:B5 is valid — B5 has no sheet and inherits Sheet1.
			if toRef.Sheet != "" && toRef.Sheet != fromRef.Sheet {
				left = &ErrorLit{Code: ErrVALUE}
				for p.peek().Type == TokPercent {
					p.advance()
					left = &PostfixExpr{Op: "%", Operand: left}
				}
				continue
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

	// Handle first argument: may be empty (e.g. SUM(,1))
	if p.peek().Type == TokComma {
		args = append(args, &EmptyArg{})
	} else {
		arg, err := p.parseExpression(0)
		if err != nil {
			return nil, err
		}
		args = append(args, arg)
	}

	for p.peek().Type == TokComma {
		p.advance()
		// Handle empty argument (e.g. ADDRESS(1,1,,"Data"))
		if p.peek().Type == TokComma || p.peek().Type == TokRParen {
			args = append(args, &EmptyArg{})
		} else {
			arg, err := p.parseExpression(0)
			if err != nil {
				return nil, err
			}
			args = append(args, arg)
		}
	}

	if _, err := p.expect(TokRParen); err != nil {
		return nil, fmt.Errorf("expected ')' to close function %s at position %d", name, tok.Pos)
	}

	call := &FuncCall{Name: name, Args: args}
	if isLambdaFuncName(name) && p.peek().Type == TokLParen {
		callArgs, err := p.parseCallArgs()
		if err != nil {
			return nil, err
		}
		return desugarLambdaInvocation(args, callArgs)
	}

	return call, nil
}

func (p *Parser) parseCallArgs() ([]Node, error) {
	if _, err := p.expect(TokLParen); err != nil {
		return nil, err
	}
	if p.peek().Type == TokRParen {
		p.advance()
		return nil, nil
	}

	var args []Node
	if p.peek().Type == TokComma {
		args = append(args, &EmptyArg{})
	} else {
		arg, err := p.parseExpression(0)
		if err != nil {
			return nil, err
		}
		args = append(args, arg)
	}

	for p.peek().Type == TokComma {
		p.advance()
		if p.peek().Type == TokComma || p.peek().Type == TokRParen {
			args = append(args, &EmptyArg{})
			continue
		}
		arg, err := p.parseExpression(0)
		if err != nil {
			return nil, err
		}
		args = append(args, arg)
	}

	if _, err := p.expect(TokRParen); err != nil {
		return nil, fmt.Errorf("expected ')' to close lambda invocation")
	}
	return args, nil
}

func isLambdaFuncName(name string) bool {
	upper := strings.ToUpper(name)
	return upper == "LAMBDA" || upper == "_XLFN.LAMBDA"
}

func desugarLambdaInvocation(lambdaArgs, callArgs []Node) (Node, error) {
	if len(lambdaArgs) == 0 {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	body := lambdaArgs[len(lambdaArgs)-1]
	params := lambdaArgs[:len(lambdaArgs)-1]
	if len(callArgs) != len(params) {
		return &ErrorLit{Code: ErrVALUE}, nil
	}

	subst := make(map[string]Node, len(params))
	for i, param := range params {
		name, ok := lambdaParamName(param)
		if !ok {
			return &ErrorLit{Code: ErrVALUE}, nil
		}
		subst[name] = callArgs[i]
	}

	return substituteLambdaNames(body, subst), nil
}

func lambdaParamName(n Node) (string, bool) {
	ref, ok := n.(*CellRef)
	if !ok || ref.Row != 0 || ref.Sheet != "" || ref.SheetEnd != "" || ref.AbsCol || ref.AbsRow || ref.DotNotation {
		return "", false
	}
	return strings.ToUpper(colNumberToLetters(ref.Col)), true
}

func substituteLambdaNames(n Node, subst map[string]Node) Node {
	switch v := n.(type) {
	case *CellRef:
		if v.Row == 0 && v.Sheet == "" && v.SheetEnd == "" && !v.AbsCol && !v.AbsRow && !v.DotNotation {
			if repl, ok := subst[strings.ToUpper(colNumberToLetters(v.Col))]; ok {
				return cloneNode(repl)
			}
		}
		return cloneNode(v)
	case *RangeRef:
		return &RangeRef{
			From: substituteLambdaNames(v.From, subst).(*CellRef),
			To:   substituteLambdaNames(v.To, subst).(*CellRef),
		}
	case *UnaryExpr:
		return &UnaryExpr{Op: v.Op, Operand: substituteLambdaNames(v.Operand, subst)}
	case *BinaryExpr:
		return &BinaryExpr{
			Op:    v.Op,
			Left:  substituteLambdaNames(v.Left, subst),
			Right: substituteLambdaNames(v.Right, subst),
		}
	case *PostfixExpr:
		return &PostfixExpr{Op: v.Op, Operand: substituteLambdaNames(v.Operand, subst)}
	case *FuncCall:
		args := make([]Node, len(v.Args))
		for i, arg := range v.Args {
			args[i] = substituteLambdaNames(arg, subst)
		}
		return &FuncCall{Name: v.Name, Args: args}
	case *ArrayLit:
		rows := make([][]Node, len(v.Rows))
		for i, row := range v.Rows {
			rows[i] = make([]Node, len(row))
			for j, elem := range row {
				rows[i][j] = substituteLambdaNames(elem, subst)
			}
		}
		return &ArrayLit{Rows: rows}
	default:
		return cloneNode(v)
	}
}

func cloneNode(n Node) Node {
	switch v := n.(type) {
	case *NumberLit:
		return &NumberLit{Value: v.Value, Raw: v.Raw}
	case *StringLit:
		return &StringLit{Value: v.Value}
	case *BoolLit:
		return &BoolLit{Value: v.Value}
	case *ErrorLit:
		return &ErrorLit{Code: v.Code}
	case *EmptyArg:
		return &EmptyArg{}
	case *CellRef:
		clone := *v
		return &clone
	case *RangeRef:
		return &RangeRef{
			From: cloneNode(v.From).(*CellRef),
			To:   cloneNode(v.To).(*CellRef),
		}
	case *UnaryExpr:
		return &UnaryExpr{Op: v.Op, Operand: cloneNode(v.Operand)}
	case *BinaryExpr:
		return &BinaryExpr{Op: v.Op, Left: cloneNode(v.Left), Right: cloneNode(v.Right)}
	case *PostfixExpr:
		return &PostfixExpr{Op: v.Op, Operand: cloneNode(v.Operand)}
	case *FuncCall:
		args := make([]Node, len(v.Args))
		for i, arg := range v.Args {
			args[i] = cloneNode(arg)
		}
		return &FuncCall{Name: v.Name, Args: args}
	case *ArrayLit:
		rows := make([][]Node, len(v.Rows))
		for i, row := range v.Rows {
			rows[i] = make([]Node, len(row))
			for j, elem := range row {
				rows[i][j] = cloneNode(elem)
			}
		}
		return &ArrayLit{Rows: rows}
	default:
		return v
	}
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
